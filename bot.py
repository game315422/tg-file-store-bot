import os
import sys
import base64
import secrets
import asyncio
import math
import sqlite3
import html
from datetime import datetime, date
import openpyxl

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    BotCommand,
    InputMediaPhoto,
    InputMediaVideo,
    InputMediaDocument,
)
from telegram.constants import ChatMemberStatus
from telegram.error import BadRequest
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes,
)

# ==================== 1. 配置与常量 ====================
BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    sys.exit("错误: 未配置 BOT_TOKEN 环境变量")

ADMIN_USER_ID = int(os.getenv("ADMIN_USER_ID", 0))
DB_CHANNEL_ID = int(os.getenv("DB_CHANNEL_ID", 0))

FORCE_SUB_CHANNEL_ID = int(os.getenv("FORCE_SUB_CHANNEL_ID", -1004348034007))
FORCE_SUB_INVITE_LINK = os.getenv("FORCE_SUB_INVITE_LINK", "https://t.me/wjtestb")

BOT_USERNAME = "wjbottest_bot"
CODE_PREFIX = "wjbottest_"
PAGE_SIZE = 10  # 每组文件数上限
DATA_DIR = "/app/data"
os.makedirs(DATA_DIR, exist_ok=True)
DB_FILE = os.path.join(DATA_DIR, "records.db")

FILTER_KEYWORDS = ["取件完成", "取件码", "防失联", "交流群", "已全部发送", "文件总数", "此代码已"]


# ==================== 2. 数据库操作层 (DAO) ====================
class DatabaseManager:
    @staticmethod
    def get_connection():
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        return conn

    @classmethod
    def init_db(cls):
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.executescript("""
                CREATE TABLE IF NOT EXISTS code_mapping (
                    code TEXT PRIMARY KEY,
                    start_id INTEGER,
                    end_id INTEGER,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS upload_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER,
                    username TEXT,
                    code TEXT UNIQUE,
                    note TEXT,
                    file_count INTEGER,
                    created_date TEXT,
                    created_time TEXT
                );
                CREATE TABLE IF NOT EXISTS batch_files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT,
                    file_id TEXT,
                    file_type TEXT,
                    file_order INTEGER
                );
                CREATE INDEX IF NOT EXISTS idx_batch_code ON batch_files(code);
            """)
            cursor.execute("PRAGMA table_info(upload_logs)")
            columns = [row[1] for row in cursor.fetchall()]
            if "note" not in columns:
                cursor.execute("ALTER TABLE upload_logs ADD COLUMN note TEXT")

    @classmethod
    def save_batch_files(cls, code: str, items: list[dict]):
        with cls.get_connection() as conn:
            records = [(code, item["file_id"], item["file_type"], idx) for idx, item in enumerate(items)]
            conn.executemany(
                "INSERT INTO batch_files (code, file_id, file_type, file_order) VALUES (?, ?, ?, ?)",
                records
            )

    @classmethod
    def get_batch_page(cls, code: str, page: int) -> list[tuple[str, str]]:
        offset = (page - 1) * PAGE_SIZE
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT file_id, file_type FROM batch_files WHERE code = ? ORDER BY file_order ASC LIMIT ? OFFSET ?",
                (code, PAGE_SIZE, offset)
            )
            return [(row["file_id"], row["file_type"]) for row in cursor.fetchall()]

    @classmethod
    def get_total_count(cls, code: str) -> int:
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) as cnt FROM batch_files WHERE code = ?", (code,))
            row = cursor.fetchone()
            return row["cnt"] if row else 0

    @classmethod
    def get_code_by_range(cls, start_id: int, end_id: int) -> str | None:
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT code FROM code_mapping WHERE start_id = ? AND end_id = ?", (start_id, end_id))
            row = cursor.fetchone()
            return row["code"] if row else None

    @classmethod
    def save_code_mapping(cls, code: str, start_id: int, end_id: int):
        with cls.get_connection() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO code_mapping (code, start_id, end_id) VALUES (?, ?, ?)",
                (code, start_id, end_id)
            )

    @classmethod
    def query_code_mapping(cls, code: str) -> tuple[int, int] | None:
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT start_id, end_id FROM code_mapping WHERE code = ?", (code,))
            row = cursor.fetchone()
            return (row["start_id"], row["end_id"]) if row else None

    @classmethod
    def log_upload(cls, user_id: int, username: str, code: str, note: str, file_count: int):
        now = datetime.now()
        with cls.get_connection() as conn:
            conn.execute("""
                INSERT OR IGNORE INTO upload_logs (user_id, username, code, note, file_count, created_date, created_time)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (user_id, username or "未知", code, note, file_count, now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")))

    @classmethod
    def get_logs_by_date(cls, target_date: str) -> list:
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, user_id, username, code, note, file_count, created_time
                FROM upload_logs WHERE created_date = ? ORDER BY id ASC
            """, (target_date,))
            return cursor.fetchall()


# ==================== 3. 业务工具类 ====================
class MediaHelper:
    @staticmethod
    def clean_caption(caption: str | None) -> str:
        if not caption:
            return ""
        if any(kw in caption for kw in FILTER_KEYWORDS):
            return ""
        return caption.strip()

    @staticmethod
    def parse_payload(message):
        if message.photo:
            return message.photo[-1].file_id, "P", "图片"
        elif message.video:
            return message.video.file_id, "V", message.video.file_name or "视频"
        elif message.document:
            return message.document.file_id, "D", message.document.file_name or "文档"
        return None, None, None

    @staticmethod
    def format_summary(types_list: list[str]) -> tuple[str, str]:
        counts = {"P": types_list.count("P"), "V": types_list.count("V"), "D": types_list.count("D")}
        tag_parts = [f"{cnt}{k}" for k, cnt in counts.items() if cnt > 0]
        detail_tag = "".join(tag_parts) if tag_parts else f"{len(types_list)}F"
        return f"{len(types_list)} 个 ({detail_tag})", detail_tag


class CodeEncoder:
    @staticmethod
    def encode(start_id: int, end_id: int, detail_tag: str = "") -> str:
        existing = DatabaseManager.get_code_by_range(start_id, end_id)
        if existing:
            return existing

        raw = f"{secrets.token_hex(3)}_{start_id if start_id == end_id else f'{start_id}-{end_id}'}_{secrets.token_hex(3)}"
        encoded = base64.urlsafe_b64encode(raw.encode()).decode().rstrip("=")
        code = f"{CODE_PREFIX}{f'{detail_tag}_' if detail_tag else ''}{encoded}"
        
        DatabaseManager.save_code_mapping(code, start_id, end_id)
        return code

    @staticmethod
    def decode(code: str) -> tuple[int, int] | None:
        db_res = DatabaseManager.query_code_mapping(code)
        if db_res:
            return db_res

        try:
            raw_code = code[len(CODE_PREFIX):] if code.startswith(CODE_PREFIX) else code
            if "_" in raw_code:
                parts_check = raw_code.split("_", 1)
                if any(ch in parts_check[0] for ch in "PVDFA"):
                    raw_code = parts_check[1]

            raw_code += "=" * (-len(raw_code) % 4)
            raw = base64.urlsafe_b64decode(raw_code.encode()).decode()
            parts = raw.split("_")
            target = parts[1] if len(parts) == 3 else parts[0]

            if "-" in target:
                s, e = target.split("-")
                res = (int(s), int(e))
            else:
                res = (int(target), int(target))

            DatabaseManager.save_code_mapping(code, res[0], res[1])
            return res
        except Exception:
            return None


# ==================== 4. 媒体统一打包发送函数 ====================
async def send_media_batch(bot, chat_id: int, items: list[dict | tuple], caption: str = "") -> list[int]:
    """
    以相册/媒体组（2-10个一组）形式发送媒体，返回消息 ID 列表。
    - 图片和视频混合打包为相册网格
    - 单个媒体自动安全降级（避免低于2个报错）
    - 文档独立成组发送
    """
    sent_ids = []
    norm_items = []
    for it in items:
        if isinstance(it, dict):
            norm_items.append({"file_id": it["file_id"], "file_type": it["file_type"]})
        else:
            norm_items.append({"file_id": it[0], "file_type": it[1]})

    pv_items = [it for it in norm_items if it["file_type"] in ("P", "V")]
    doc_items = [it for it in norm_items if it["file_type"] == "D"]

    # 1. 处理图片与视频相册
    if pv_items:
        if len(pv_items) >= 2:
            media_group = []
            for idx, it in enumerate(pv_items):
                cap = caption if idx == 0 else ""
                if it["file_type"] == "P":
                    media_group.append(InputMediaPhoto(media=it["file_id"], caption=cap, parse_mode="HTML" if cap else None))
                else:
                    media_group.append(InputMediaVideo(media=it["file_id"], caption=cap, parse_mode="HTML" if cap else None))
            try:
                msgs = await bot.send_media_group(chat_id=chat_id, media=media_group)
                sent_ids.extend([m.message_id for m in msgs])
            except Exception as e:
                print(f"相册发送失败降级: {e}")
                for idx, it in enumerate(pv_items):
                    cap = caption if idx == 0 else ""
                    if it["file_type"] == "P":
                        m = await bot.send_photo(chat_id=chat_id, photo=it["file_id"], caption=cap, parse_mode="HTML" if cap else None)
                    else:
                        m = await bot.send_video(chat_id=chat_id, video=it["file_id"], caption=cap, parse_mode="HTML" if cap else None)
                    sent_ids.append(m.message_id)
                    await asyncio.sleep(0.3)
        else:
            # 只有 1 个文件时 Telegram 不允许调用 send_media_group
            it = pv_items[0]
            if it["file_type"] == "P":
                m = await bot.send_photo(chat_id=chat_id, photo=it["file_id"], caption=caption, parse_mode="HTML" if caption else None)
            else:
                m = await bot.send_video(chat_id=chat_id, video=it["file_id"], caption=caption, parse_mode="HTML" if caption else None)
            sent_ids.append(m.message_id)

    # 2. 处理文档（如果有）
    if doc_items:
        doc_caption = caption if not pv_items else ""
        if len(doc_items) >= 2:
            media_group = []
            for idx, it in enumerate(doc_items):
                cap = doc_caption if idx == 0 else ""
                media_group.append(InputMediaDocument(media=it["file_id"], caption=cap, parse_mode="HTML" if cap else None))
            try:
                msgs = await bot.send_media_group(chat_id=chat_id, media=media_group)
                sent_ids.extend([m.message_id for m in msgs])
            except Exception:
                for idx, it in enumerate(doc_items):
                    cap = doc_caption if idx == 0 else ""
                    m = await bot.send_document(chat_id=chat_id, document=it["file_id"], caption=cap, parse_mode="HTML" if cap else None)
                    sent_ids.append(m.message_id)
                    await asyncio.sleep(0.3)
        else:
            m = await bot.send_document(chat_id=chat_id, document=doc_items[0]["file_id"], caption=doc_caption, parse_mode="HTML" if doc_caption else None)
            sent_ids.append(m.message_id)

    return sent_ids


# ==================== 5. 核心交互处理器 ====================
async def is_subscribed(bot, user_id: int) -> bool:
    if not FORCE_SUB_CHANNEL_ID:
        return True
    try:
        member = await bot.get_chat_member(chat_id=FORCE_SUB_CHANNEL_ID, user_id=user_id)
        return member.status in {ChatMemberStatus.MEMBER, ChatMemberStatus.ADMINISTRATOR, ChatMemberStatus.OWNER}
    except Exception:
        return False


def build_panel(code: str, current_group: int, total_groups: int, total_files: int):
    status_text = (
        f"📦 <b>资源查看控制台</b>\n"
        f"📊 <b>当前组别</b>：第 <code>{current_group}</code> / <code>{total_groups}</code> 组\n"
        f"📁 <b>文件总数</b>：共 {total_files} 个 (每组 {PAGE_SIZE} 个)"
    )

    # 提取短 code 避免 callback_data 超过 64 字节
    short_code = code[len(CODE_PREFIX):] if code.startswith(CODE_PREFIX) else code

    keyboard = []
    nav_row = [
        InlineKeyboardButton(
            "◀️ 上一组" if current_group > 1 else "已在首页 ⏹", 
            callback_data=f"pg:{current_group - 1}:{short_code}" if current_group > 1 else "ignore"
        ),
        InlineKeyboardButton(
            "下一组 ▶️" if current_group < total_groups else "已在末页 ⏹", 
            callback_data=f"pg:{current_group + 1}:{short_code}" if current_group < total_groups else "ignore"
        )
    ]
    keyboard.append(nav_row)

    num_row = []
    for g in range(1, total_groups + 1):
        cb_data = "ignore" if g == current_group else f"pg:{g}:{short_code}"
        num_row.append(InlineKeyboardButton(f"·{g}·" if g == current_group else f"{g}", callback_data=cb_data))
        if len(num_row) == 5:
            keyboard.append(num_row)
            num_row = []
    if num_row:
        keyboard.append(num_row)

    share_link = f"https://t.me/share/url?url=https://t.me/{BOT_USERNAME}?start={code}"
    keyboard.append([
        InlineKeyboardButton("⭐ 收藏此资源", url=share_link),
        InlineKeyboardButton("📢 官方防失联", url=FORCE_SUB_INVITE_LINK)
    ])

    return status_text, InlineKeyboardMarkup(keyboard)


async def deliver_group_page(bot, chat_id: int, code: str, current_group: int, old_msg_id: int | None = None):
    if old_msg_id:
        try:
            await bot.delete_message(chat_id=chat_id, message_id=old_msg_id)
        except Exception:
            pass

    total_files = DatabaseManager.get_total_count(code)
    if total_files == 0:
        payload = CodeEncoder.decode(code)
        if not payload:
            await bot.send_message(chat_id=chat_id, text="⚠️ 资源不存在或提取码无效。")
            return
        total_files = payload[1] - payload[0] + 1

    total_groups = max(1, math.ceil(total_files / PAGE_SIZE))
    page_items = DatabaseManager.get_batch_page(code, current_group)

    if page_items:
        # 使用统一封装的相册发送
        await send_media_batch(bot, chat_id, page_items)
    else:
        # 兼容旧单发逻辑
        payload = CodeEncoder.decode(code)
        if payload:
            start_id, end_id = payload
            g_start = start_id + (current_group - 1) * PAGE_SIZE
            g_end = min(g_start + PAGE_SIZE - 1, end_id)
            for mid in range(g_start, g_end + 1):
                try:
                    await bot.copy_message(chat_id=chat_id, from_chat_id=DB_CHANNEL_ID, message_id=mid)
                    await asyncio.sleep(0.3)
                except Exception:
                    pass

    # 控制面板紧跟在资源下方
    status_text, markup = build_panel(code, current_group, total_groups, total_files)
    await bot.send_message(chat_id=chat_id, text=status_text, reply_markup=markup, parse_mode="HTML")


# ==================== 6. 命令与消息绑定 ====================
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 <b>媒体文件存取机器人</b>\n\n"
        "• 直接发送图片/视频/文档，输入 <code>/done 备注</code> 归档生成取件码\n"
        "• 输入 <code>/cancel</code> 清空当前暂存列表\n"
        "• 发送提取码即可浏览翻页。",
        parse_mode="HTML"
    )

async def cancel_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("pending_items", None)
    context.user_data.pop("last_status_msg_id", None)
    await update.message.reply_text("🧹 已清空当前待上传队列。")

async def save_media(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    file_id, f_type, default_note = MediaHelper.parse_payload(msg)
    if not f_type:
        return

    if "pending_items" not in context.user_data:
        context.user_data["pending_items"] = []

    context.user_data["pending_items"].append({
        "file_id": file_id,
        "file_type": f_type,
        "caption": MediaHelper.clean_caption(msg.caption),
        "default_note": default_note
    })

    count = len(context.user_data["pending_items"])
    tip_text = f"📥 <b>已暂存 {count} 个文件</b>\n发送完毕后请输入 <code>/done 备注</code> 结算。"

    status_id = context.user_data.get("last_status_msg_id")
    try:
        if status_id:
            await context.bot.edit_message_text(chat_id=msg.chat_id, message_id=status_id, text=tip_text, parse_mode="HTML")
            return
    except BadRequest:
        pass

    sent = await msg.reply_text(tip_text, parse_mode="HTML")
    context.user_data["last_status_msg_id"] = sent.message_id


async def done_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    pending = context.user_data.get("pending_items", [])
    if not pending:
        await update.message.reply_text("未收到任何待处理的文件。")
        return

    if context.args:
        raw_arg = " ".join(context.args).strip()
        note = raw_arg.strip("[]【】") or "无"
    else:
        first_cap = next((it["caption"] for it in pending if it["caption"]), "")
        note = first_cap or pending[0]["default_note"] or "无"

    total_count = len(pending)
    summary, detail_tag = MediaHelper.format_summary([it["file_type"] for it in pending])
    total_groups = math.ceil(total_count / PAGE_SIZE)

    status_notify = await update.message.reply_text(f"⏳ 正在将 {total_count} 个文件打包转存至仓库...")

    saved_mids = []
    caption_banner = f"📝 备注：{html.escape(note)}"

    # 10 个一组打包发送到存储频道
    for g_idx in range(total_groups):
        group_items = pending[g_idx * PAGE_SIZE : (g_idx + 1) * PAGE_SIZE]
        cap = caption_banner if g_idx == 0 else ""
        try:
            mids = await send_media_batch(context.bot, DB_CHANNEL_ID, group_items, caption=cap)
            saved_mids.extend(mids)
        except Exception as e:
            print(f"转存存储频道出错: {e}")
        await asyncio.sleep(1.2)

    if not saved_mids:
        await status_notify.edit_text("❌ 转存失败，请检查后台日志。")
        return

    code = CodeEncoder.encode(min(saved_mids), max(saved_mids), detail_tag)
    DatabaseManager.save_batch_files(code, pending)
    DatabaseManager.log_upload(
        user_id=update.effective_user.id,
        username=update.effective_user.username or update.effective_user.first_name,
        code=code, note=note, file_count=total_count
    )

    context.user_data.clear()
    await status_notify.delete()

    await update.message.reply_text(
        f"📝 <b>备注</b>：{html.escape(note)}\n"
        f"📦 <b>文件</b>：{summary}\n"
        f"🤖 <b>取件</b>：@{BOT_USERNAME}\n"
        f"🔑 <b>取件码</b>：<code>{code}</code>",
        parse_mode="HTML"
    )


async def handle_start_or_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text == "/start":
        await help_command(update, context)
        return

    code = text.split(maxsplit=1)[1].strip() if text.startswith("/start ") else text
    if not CodeEncoder.decode(code):
        return

    user_id = update.effective_user.id
    if not await is_subscribed(context.bot, user_id):
        keyboard = [
            [InlineKeyboardButton("👥 点击加入官方群", url=FORCE_SUB_INVITE_LINK)],
            [InlineKeyboardButton("🔄 我已加入，继续提取", callback_data=f"verify_{code}")]
        ]
        await update.message.reply_text("⚠️ <b>访问受限</b>\n请先加入官方群组后方可提取！", reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="HTML")
        return

    await deliver_group_page(context.bot, user_id, code, 1)


async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "ignore":
        return

    user_id = update.effective_user.id
    
    # 修复：使用冒号分割，不受取件码中下划线的影响
    if data.startswith("pg:"):
        _, group_str, short_code = data.split(":", 2)
        code = f"{CODE_PREFIX}{short_code}" if not short_code.startswith(CODE_PREFIX) else short_code
        await deliver_group_page(context.bot, user_id, code, int(group_str), query.message.message_id)
        
    elif data.startswith("verify_"):
        code = data.replace("verify_", "")
        if not await is_subscribed(context.bot, user_id):
            await query.answer("❌ 尚未检测到您加入群组！", show_alert=True)
            return
        await query.delete_message()
        await deliver_group_page(context.bot, user_id, code, 1)


async def export_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_USER_ID:
        return
    today = date.today().strftime("%Y-%m-%d")
    rows = DatabaseManager.get_logs_by_date(today)
    if not rows:
        await update.message.reply_text(f"今日 ({today}) 暂无上传记录。")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = today
    ws.append(["序号", "用户ID", "用户名", "取件码", "备注", "包含文件数", "创建时间"])
    for r in rows:
        ws.append(list(r))

    file_path = os.path.join(DATA_DIR, f"{today}_报表.xlsx")
    wb.save(file_path)
    await update.message.reply_document(document=open(file_path, "rb"), caption=f"📊 今日 ({today}) 报表")


async def post_init(application: Application):
    await application.bot.set_my_commands([
        BotCommand("start", "启动机器人"),
        BotCommand("done", "打包生成取件码"),
        BotCommand("cancel", "清空暂存"),
        BotCommand("help", "帮助说明"),
        BotCommand("export", "导出报表(管)")
    ])


# ==================== 7. 主程序入口 ====================
def main():
    DatabaseManager.init_db()
    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()

    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("done", done_command))
    app.add_handler(CommandHandler("cancel", cancel_command))
    app.add_handler(CommandHandler("export", export_command))
    app.add_handler(CommandHandler("start", handle_start_or_code))

    media_filter = (filters.VIDEO | filters.Document.ALL | filters.PHOTO | filters.AUDIO) & filters.ChatType.PRIVATE
    app.add_handler(MessageHandler(media_filter, save_media))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND & filters.ChatType.PRIVATE, handle_start_or_code))
    app.add_handler(CallbackQueryHandler(handle_callback))

    print("Bot 修复版运行中...")
    app.run_polling()

if __name__ == "__main__":
    main()
