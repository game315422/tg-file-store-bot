import os
import sys
import base64
import secrets
import asyncio
import math
import sqlite3
import html
from datetime import datetime, date, time
from zoneinfo import ZoneInfo
import openpyxl

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
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

# ----------------- 环境变量读取 -----------------
BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    sys.exit("错误: 未配置 BOT_TOKEN 环境变量")

ADMIN_USER_ID = int(os.getenv("ADMIN_USER_ID", 0))
DB_CHANNEL_ID = int(os.getenv("DB_CHANNEL_ID", 0))
FORCE_SUB_CHANNEL_ID = int(os.getenv("FORCE_SUB_CHANNEL_ID", 0))
FORCE_SUB_INVITE_LINK = os.getenv("FORCE_SUB_INVITE_LINK", "")

BOT_USERNAME = "wjbottest_bot"
CODE_PREFIX = "wjbottest_"
PAGE_SIZE = 9  # 九宫格每组上限 9 个
DATA_DIR = "/app/data"
os.makedirs(DATA_DIR, exist_ok=True)
DB_FILE = os.path.join(DATA_DIR, "records.db")

# ----------------- 数据库初始化 -----------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS code_mapping (
            code TEXT PRIMARY KEY,
            start_id INTEGER,
            end_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS upload_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username TEXT,
            code TEXT UNIQUE,
            note TEXT,
            file_count INTEGER,
            created_date TEXT,
            created_time TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS batch_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT,
            file_id TEXT,
            file_type TEXT,
            file_order INTEGER
        )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_batch_code ON batch_files(code)")

    cursor.execute("PRAGMA table_info(upload_logs)")
    columns = [row[1] for row in cursor.fetchall()]
    if "note" not in columns:
        cursor.execute("ALTER TABLE upload_logs ADD COLUMN note TEXT")

    conn.commit()
    conn.close()

def save_batch_file_details(code: str, items: list[dict]):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    records = [(code, item["file_id"], item["file_type"], idx) for idx, item in enumerate(items)]
    cursor.executemany("INSERT INTO batch_files (code, file_id, file_type, file_order) VALUES (?, ?, ?, ?)", records)
    conn.commit()
    conn.close()

def get_batch_files_page(code: str, page: int, page_size: int = 9) -> list[tuple[str, str]]:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    offset = (page - 1) * page_size
    cursor.execute(
        "SELECT file_id, file_type FROM batch_files WHERE code = ? ORDER BY file_order ASC LIMIT ? OFFSET ?",
        (code, page_size, offset)
    )
    rows = cursor.fetchall()
    conn.close()
    return rows

def get_batch_total_count(code: str) -> int:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM batch_files WHERE code = ?", (code,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else 0

def get_code_by_range(start_id: int, end_id: int) -> str | None:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT code FROM code_mapping WHERE start_id = ? AND end_id = ?", (start_id, end_id))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else None

def save_code_mapping(code: str, start_id: int, end_id: int):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR REPLACE INTO code_mapping (code, start_id, end_id)
        VALUES (?, ?, ?)
    """, (code, start_id, end_id))
    conn.commit()
    conn.close()

def query_code_mapping(code: str) -> tuple[int, int] | None:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT start_id, end_id FROM code_mapping WHERE code = ?", (code,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return row[0], row[1]
    return None

def log_upload(user_id: int, username: str, code: str, note: str, file_count: int):
    now = datetime.now()
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR IGNORE INTO upload_logs (user_id, username, code, note, file_count, created_date, created_time)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        user_id,
        username or "未知",
        code,
        note,
        file_count,
        now.strftime("%Y-%m-%d"),
        now.strftime("%H:%M:%S")
    ))
    conn.commit()
    conn.close()

def export_excel(target_date: str) -> str | None:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, user_id, username, code, note, file_count, created_time
        FROM upload_logs
        WHERE created_date = ?
        ORDER BY id ASC
    """, (target_date,))
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = target_date

    headers = ["序号", "用户ID", "用户名", "取件码", "备注", "包含文件数", "创建时间"]
    ws.append(headers)

    for row in rows:
        ws.append(list(row))

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    file_path = os.path.join(DATA_DIR, f"{target_date}_提取码报表.xlsx")
    wb.save(file_path)
    return file_path

# ----------------- 媒体解析与严格类型判定 -----------------
FILTER_KEYWORDS = ["取件完成", "取件码", "防失联", "交流群", "已全部发送", "文件总数", "此代码已"]

def clean_caption(caption: str | None) -> str:
    if not caption:
        return ""
    for kw in FILTER_KEYWORDS:
        if kw in caption:
            return ""
    return caption.strip()

def get_media_payload(message):
    # 严格保持底层所属真实媒体类型，避免跨类作为相册被 Telegram 拒绝
    if message.photo:
        return message.photo[-1].file_id, "P", "图片"
    elif message.video:
        note_name = message.video.file_name or "视频"
        return message.video.file_id, "V", note_name
    elif message.document:
        note_name = message.document.file_name or "文档"
        return message.document.file_id, "D", note_name
    return None, None, None

def format_summary_text(types_list: list[str]) -> tuple[str, str]:
    p_count = types_list.count("P")
    v_count = types_list.count("V")
    d_count = types_list.count("D")

    tag_parts = []
    if p_count > 0:
        tag_parts.append(f"{p_count}P")
    if v_count > 0:
        tag_parts.append(f"{v_count}V")
    if d_count > 0:
        tag_parts.append(f"{d_count}D")

    detail_tag = "".join(tag_parts) if tag_parts else f"{len(types_list)}F"
    display_str = f"{len(types_list)} 个 ({detail_tag})"
    return display_str, detail_tag

def build_card_message(note: str, display_summary: str, code: str) -> str:
    escaped_note = html.escape(note)
    return (
        f"📝 <b>备注</b>：{escaped_note}\n"
        f"📦 <b>文件</b>：{display_summary}\n"
        f"🤖 <b>取件</b>：@{BOT_USERNAME}\n"
        f"🔑 <b>取件码</b>：<code>{code}</code>"
    )

# ----------------- 编解码逻辑 -----------------
def encode_payload(start_id: int, end_id: int, detail_tag: str = "") -> str:
    existing_code = get_code_by_range(start_id, end_id)
    if existing_code:
        return existing_code

    salt_prefix = secrets.token_hex(4)
    salt_suffix = secrets.token_hex(4)
    range_str = f"{start_id}-{end_id}" if start_id != end_id else str(start_id)
    raw = f"{salt_prefix}_{range_str}_{salt_suffix}"
    encoded = base64.urlsafe_b64encode(raw.encode()).decode().rstrip("=")

    mid_tag = f"{detail_tag}_" if detail_tag else ""
    code = f"{CODE_PREFIX}{mid_tag}{encoded}"
    
    save_code_mapping(code, start_id, end_id)
    return code

def decode_payload(code: str) -> tuple[int, int] | None:
    db_res = query_code_mapping(code)
    if db_res:
        return db_res

    try:
        raw_code = code
        if raw_code.startswith(CODE_PREFIX):
            raw_code = raw_code[len(CODE_PREFIX):]

        if "_" in raw_code:
            parts_check = raw_code.split("_", 1)
            if any(ch in parts_check[0] for ch in ["P", "V", "D", "A", "F"]):
                raw_code = parts_check[1]

        raw_code += "=" * (-len(raw_code) % 4)
        raw = base64.urlsafe_b64decode(raw_code.encode()).decode()
        parts = raw.split("_")
        
        if len(parts) == 3:
            target = parts[1]
        elif len(parts) == 1:
            target = parts[0]
        else:
            return None

        if "-" in target:
            s, e = target.split("-")
            res = (int(s), int(e))
        else:
            res = (int(target), int(target))

        save_code_mapping(code, res[0], res[1])
        return res
    except Exception:
        return None

# ----------------- 强制关注核验 -----------------
async def is_user_subscribed(bot, user_id: int) -> bool:
    if not FORCE_SUB_CHANNEL_ID:
        return True
    try:
        member = await bot.get_chat_member(chat_id=FORCE_SUB_CHANNEL_ID, user_id=user_id)
        return member.status in [
            ChatMemberStatus.MEMBER,
            ChatMemberStatus.ADMINISTRATOR,
            ChatMemberStatus.OWNER
        ]
    except Exception:
        return False

# ----------------- 命令与帮助 -----------------
def get_help_text() -> str:
    return (
        "🤖 <b>媒体文件存取机器人</b>\n\n"
        "• 发送图片/视频后，输入 <code>/done 你的备注</code> 归档并生成取件码\n"
        "• 输入 <code>/cancel</code> 清空当前暂存列表\n"
        "• 发送提取码即可在九宫格相册和控制台中浏览翻页。"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(get_help_text(), parse_mode="HTML")

async def cancel_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["pending_items"] = []
    context.user_data.pop("last_status_msg_id", None)
    await update.message.reply_text("🧹 已清空当前待上传队列。")

# ----------------- 接收与暂存文件 -----------------
async def save_media(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    file_id, file_type, default_note = get_media_payload(msg)

    if not file_type:
        return

    cleaned_caption = clean_caption(msg.caption)

    if "pending_items" not in context.user_data:
        context.user_data["pending_items"] = []

    context.user_data["pending_items"].append({
        "file_id": file_id,
        "file_type": file_type,
        "caption": cleaned_caption,
        "default_note": default_note or "",
        "message_id": msg.message_id
    })

    count = len(context.user_data["pending_items"])
    tip_text = (
        f"📥 <b>当前已暂存 {count} 个文件</b>\n"
        f"发送完成后，请输入 <code>/done 你的备注</code>（例如：<code>/done 杂集</code>）结算。"
    )

    status_msg_id = context.user_data.get("last_status_msg_id")
    updated = False
    if status_msg_id:
        try:
            await context.bot.edit_message_text(
                chat_id=msg.chat_id,
                message_id=status_msg_id,
                text=tip_text,
                parse_mode="HTML"
            )
            updated = True
        except BadRequest:
            updated = False

    if not updated:
        sent_msg = await msg.reply_text(tip_text, parse_mode="HTML")
        context.user_data["last_status_msg_id"] = sent_msg.message_id

# ----------------- 严格构建同构相册 -----------------
def create_safe_media_group(items: list[dict], caption_text: str = ""):
    media_list = []
    for idx, it in enumerate(items):
        cur_caption = caption_text if idx == 0 else None
        fid = it["file_id"]
        ft = it["file_type"]

        if ft == "P":
            media_list.append(InputMediaPhoto(media=fid, caption=cur_caption, parse_mode="HTML"))
        elif ft == "V":
            media_list.append(InputMediaVideo(media=fid, caption=cur_caption, parse_mode="HTML"))
        else:
            media_list.append(InputMediaDocument(media=fid, caption=cur_caption, parse_mode="HTML"))
    return media_list

# ----------------- 存入仓库频道（确保仓库也是九宫格） -----------------
async def dispatch_items_as_album(bot, target_chat_id: int, items: list[dict], caption_text: str = "") -> list[int]:
    if not items:
        return []

    # 若刚好只有 1 个文件，直接单独发送，避免触发 send_media_group 至少 2 个的限制
    if len(items) == 1:
        it = items[0]
        try:
            if it["file_type"] == "P":
                msg = await bot.send_photo(chat_id=target_chat_id, photo=it["file_id"], caption=caption_text, parse_mode="HTML")
            elif it["file_type"] == "V":
                msg = await bot.send_video(chat_id=target_chat_id, video=it["file_id"], caption=caption_text, parse_mode="HTML")
            else:
                msg = await bot.send_document(chat_id=target_chat_id, document=it["file_id"], caption=caption_text, parse_mode="HTML")
            return [msg.message_id]
        except Exception:
            return []

    # 包含图片和视频时正常组装为相册
    try:
        media_group = create_safe_media_group(items, caption_text=caption_text)
        sent_messages = await bot.send_media_group(chat_id=target_chat_id, media=media_group)
        return [m.message_id for m in sent_messages]
    except Exception as e:
        print(f"相册发送失败: {e}，尝试使用批量复制兜底")
        # 若用户传入的是原消息，尝试使用官方原样批量复制保证成册
        msg_ids = [it["message_id"] for it in items if "message_id" in it]
        if msg_ids and len(msg_ids) == len(items):
            try:
                copied = await bot.copy_messages(
                    chat_id=target_chat_id,
                    from_chat_id=items[0].get("from_chat_id", target_chat_id),
                    message_ids=msg_ids
                )
                return [m.message_id for m in copied]
            except Exception:
                pass
    return []

# ----------------- 结算入库 -----------------
async def done_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    pending = context.user_data.get("pending_items", [])
    if not pending:
        await update.message.reply_text("未收到任何待处理的文件，请先发送媒体文件。")
        return

    if context.args:
        raw_arg = " ".join(context.args).strip()
        if (raw_arg.startswith("[") and raw_arg.endswith("]")) or (raw_arg.startswith("【") and raw_arg.endswith("】")):
            raw_arg = raw_arg[1:-1].strip()
        note = raw_arg if raw_arg else "无"
    else:
        first_caption = next((it["caption"] for it in pending if it["caption"]), "")
        note = first_caption if first_caption else (pending[0]["default_note"] or "无")

    total_count = len(pending)
    types = [it["file_type"] for it in pending]
    display_summary, detail_tag = format_summary_text(types)
    total_groups = math.ceil(total_count / PAGE_SIZE)

    # 补全来源 chat_id
    for it in pending:
        it["from_chat_id"] = update.effective_chat.id

    status_notify = await update.message.reply_text(f"⏳ 正在将 {total_count} 个文件（共 {total_groups} 组九宫格）存入仓库...")

    saved_channel_mids = []
    caption_banner = f"📝 备注：{html.escape(note)}"

    for g_idx in range(total_groups):
        start_idx = g_idx * PAGE_SIZE
        end_idx = min(start_idx + PAGE_SIZE, total_count)
        group_items = pending[start_idx:end_idx]

        mids = await dispatch_items_as_album(context.bot, DB_CHANNEL_ID, group_items, caption_text=caption_banner)
        if not mids:
            await status_notify.edit_text(f"❌ 第 {g_idx+1}/{total_groups} 组存入失败，请确认 Bot 在存储频道的发帖权限！")
            return

        saved_channel_mids.extend(mids)
        await asyncio.sleep(1.2)

    start_id = min(saved_channel_mids)
    end_id = max(saved_channel_mids)
    code = encode_payload(start_id, end_id, detail_tag)

    save_batch_file_details(code, pending)

    log_upload(
        user_id=update.effective_user.id,
        username=update.effective_user.username or update.effective_user.first_name,
        code=code,
        note=note,
        file_count=len(pending)
    )

    context.user_data["pending_items"] = []
    context.user_data.pop("last_status_msg_id", None)
    await status_notify.delete()

    response_text = build_card_message(note, display_summary, code)
    await update.message.reply_text(response_text, parse_mode="HTML")

# ----------------- 生成按键控制台 -----------------
def build_delivery_panel(code: str, current_group: int, total_groups: int, total_files: int):
    status_text = (
        f"📦 <b>资源查看控制台</b>\n"
        f"📊 <b>当前组别</b>：第 <code>{current_group}</code> / <code>{total_groups}</code> 组\n"
        f"📁 <b>文件总数</b>：共 {total_files} 个 (每组 9 个)"
    )

    keyboard = []

    # 1. 翻页按键
    nav_row = []
    if current_group > 1:
        nav_row.append(InlineKeyboardButton("◀️ 上一组", callback_data=f"page_{code}_{current_group - 1}"))
    else:
        nav_row.append(InlineKeyboardButton("已在首页 ⏹", callback_data="ignore"))

    if current_group < total_groups:
        nav_row.append(InlineKeyboardButton("下一组 ▶️", callback_data=f"page_{code}_{current_group + 1}"))
    else:
        nav_row.append(InlineKeyboardButton("已在末页 ⏹", callback_data="ignore"))
    keyboard.append(nav_row)

    # 2. 数字选组按键
    num_row = []
    for g in range(1, total_groups + 1):
        if g == current_group:
            num_row.append(InlineKeyboardButton(f"·{g}·", callback_data="ignore"))
        else:
            num_row.append(InlineKeyboardButton(f"{g}", callback_data=f"page_{code}_{g}"))
        
        if len(num_row) == 5:
            keyboard.append(num_row)
            num_row = []
    if num_row:
        keyboard.append(num_row)

    # 3. 底部功能键
    share_link = f"https://t.me/share/url?url=https://t.me/{BOT_USERNAME}?start={code}"
    keyboard.append([
        InlineKeyboardButton("⭐ 收藏此资源", url=share_link),
        InlineKeyboardButton("📢 官方防失联", url=FORCE_SUB_INVITE_LINK if FORCE_SUB_INVITE_LINK else "https://t.me")
    ])

    return status_text, InlineKeyboardMarkup(keyboard)

# ----------------- 极速提取（相册九宫格 + 控制台置底） -----------------
async def deliver_group_page(bot, chat_id: int, code: str, current_group: int, old_panel_msg_id: int | None = None):
    # 删除旧控制台，确保翻页后新控制台紧随在媒体相册最底部
    if old_panel_msg_id:
        try:
            await bot.delete_message(chat_id=chat_id, message_id=old_panel_msg_id)
        except Exception:
            pass

    total_files = get_batch_total_count(code)
    start_id, end_id = None, None

    if total_files == 0:
        payload = decode_payload(code)
        if not payload:
            await bot.send_message(chat_id=chat_id, text="⚠️ 资源不存在或提取码无效。")
            return
        start_id, end_id = payload
        total_files = end_id - start_id + 1

    total_groups = max(1, math.ceil(total_files / PAGE_SIZE))
    page_items = get_batch_files_page(code, current_group, PAGE_SIZE)

    # 1. 优先使用本地索引极速发出九宫格相册
    if page_items:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT note FROM upload_logs WHERE code = ?", (code,))
        row = cursor.fetchone()
        conn.close()
        note_str = row[0] if row and row[0] else ""

        dict_items = [{"file_id": it[0], "file_type": it[1]} for it in page_items]
        caption_text = f"📝 <b>备注</b>：{html.escape(note_str)} (第 {current_group}/{total_groups} 组)" if note_str else ""
        await dispatch_items_as_album(bot, chat_id, dict_items, caption_text=caption_text)
    else:
        # 兼容老数据：使用 copy_messages 原生批量复制相册
        if not start_id or not end_id:
            payload = decode_payload(code)
            if payload:
                start_id, end_id = payload

        if start_id and end_id:
            g_start = start_id + (current_group - 1) * PAGE_SIZE
            g_end = min(g_start + PAGE_SIZE - 1, end_id)
            message_ids = list(range(g_start, g_end + 1))
            try:
                await bot.copy_messages(chat_id=chat_id, from_chat_id=DB_CHANNEL_ID, message_ids=message_ids)
            except Exception:
                for mid in message_ids:
                    try:
                        await bot.copy_message(chat_id=chat_id, from_chat_id=DB_CHANNEL_ID, message_id=mid)
                        await asyncio.sleep(0.3)
                    except Exception:
                        pass

    # 2. 控制台发送到最新媒体正下方
    status_text, markup = build_delivery_panel(code, current_group, total_groups, total_files)
    await bot.send_message(
        chat_id=chat_id,
        text=status_text,
        reply_markup=markup,
        parse_mode="HTML"
    )

async def handle_start_or_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()

    if text == "/start":
        await update.message.reply_text(get_help_text(), parse_mode="HTML")
        return

    for kw in FILTER_KEYWORDS:
        if kw in text and not text.startswith("/start") and not text.startswith(CODE_PREFIX):
            return

    if text.startswith("/start "):
        parts = text.split(maxsplit=1)
        code = parts[1].strip()
    else:
        code = text

    payload = decode_payload(code)
    if not payload:
        return

    if not await is_user_subscribed(context.bot, user_id):
        keyboard = [
            [InlineKeyboardButton("👥 点击加入官方群", url=FORCE_SUB_INVITE_LINK)],
            [InlineKeyboardButton("🔄 我已加入，继续提取", callback_data=f"verify_{code}")]
        ]
        await update.message.reply_text(
            "⚠️ <b>访问受限</b>\n\n请先加入我们的官方群组，才能提取资源！",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="HTML"
        )
        return

    await deliver_group_page(context.bot, user_id, code=code, current_group=1)

async def handle_callback_query(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = update.effective_user.id

    if data == "ignore":
        return

    if data.startswith("page_"):
        parts = data.rsplit("_", 1)
        target_group = int(parts[1])
        code = parts[0].replace("page_", "")

        await deliver_group_page(
            context.bot, 
            user_id, 
            code=code, 
            current_group=target_group, 
            old_panel_msg_id=query.message.message_id
        )

    elif data.startswith("verify_"):
        code = data.replace("verify_", "")
        if not await is_user_subscribed(context.bot, user_id):
            await query.answer("❌ 仍未检测到加入群组，请先加入群再点击！", show_alert=True)
            return

        await query.delete_message()
        await deliver_group_page(context.bot, user_id, code=code, current_group=1)

# ----------------- 报表管理 -----------------
async def export_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_USER_ID:
        return

    today_str = date.today().strftime("%Y-%m-%d")
    file_path = export_excel(today_str)
    if not file_path:
        await update.message.reply_text(f"今日 ({today_str}) 暂无任何新上传记录。")
        return

    await update.message.reply_document(
        document=open(file_path, "rb"),
        caption=f"📊 今日 ({today_str}) 提取码报表"
    )

async def daily_report_job(context: ContextTypes.DEFAULT_TYPE):
    today_str = date.today().strftime("%Y-%m-%d")
    file_path = export_excel(today_str)
    if file_path and ADMIN_USER_ID:
        try:
            await context.bot.send_document(
                chat_id=ADMIN_USER_ID,
                document=open(file_path, "rb"),
                caption=f"📋 <b>每日自动汇总</b>\n日期：{today_str}\n今日提取码报表已归档完毕！",
                parse_mode="HTML"
            )
        except Exception as e:
            print(f"每日报表推送失败: {e}")

# ----------------- 主程序入口 -----------------
def main():
    init_db()
    
    app = Application.builder().token(BOT_TOKEN).build()

    shanghai_tz = ZoneInfo("Asia/Shanghai")
    app.job_queue.run_daily(daily_report_job, time=time(hour=23, minute=59, second=0, tzinfo=shanghai_tz))

    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("done", done_command))
    app.add_handler(CommandHandler("cancel", cancel_command))
    app.add_handler(CommandHandler("export", export_command))
    app.add_handler(CommandHandler("start", handle_start_or_code))

    media_filter = (filters.VIDEO | filters.Document.ALL | filters.PHOTO | filters.AUDIO) & filters.ChatType.PRIVATE
    app.add_handler(MessageHandler(media_filter, save_media))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND & filters.ChatType.PRIVATE, handle_start_or_code))
    app.add_handler(CallbackQueryHandler(handle_callback_query))

    print("Bot 运行中 (仓库相册九宫格化 + 控制台置底)...")
    app.run_polling()

if __name__ == "__main__":
    main()
