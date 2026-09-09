import os
import sys
import base64
import secrets
import asyncio
import math
import sqlite3
from datetime import datetime, date, time
from zoneinfo import ZoneInfo
import openpyxl

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.constants import ChatMemberStatus
from telegram.error import BadRequest, TelegramError
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
# 此处可填写频道 ID 或群组 ID（群组 ID 一般以 -100 开头）
FORCE_SUB_CHANNEL_ID = int(os.getenv("FORCE_SUB_CHANNEL_ID", 0))
FORCE_SUB_INVITE_LINK = os.getenv("FORCE_SUB_INVITE_LINK", "")

CODE_PREFIX = "wjtqu"
PAGE_SIZE = 10
DATA_DIR = "/app/data"
os.makedirs(DATA_DIR, exist_ok=True)
DB_FILE = os.path.join(DATA_DIR, "records.db")

# ----------------- 数据库存储与去重 -----------------
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
        CREATE TABLE IF NOT EXISTS file_dedup (
            unique_id TEXT PRIMARY KEY,
            channel_msg_id INTEGER,
            code TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS upload_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username TEXT,
            code TEXT UNIQUE,
            link TEXT,
            file_count INTEGER,
            created_date TEXT,
            created_time TEXT
        )
    """)
    conn.commit()
    conn.close()

def get_existing_file(unique_id: str) -> tuple[int, str] | None:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT channel_msg_id, code FROM file_dedup WHERE unique_id = ?", (unique_id,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return row[0], row[1]
    return None

def save_file_fingerprint(unique_id: str, channel_msg_id: int, code: str):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR IGNORE INTO file_dedup (unique_id, channel_msg_id, code)
        VALUES (?, ?, ?)
    """, (unique_id, channel_msg_id, code))
    conn.commit()
    conn.close()

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

def log_upload(user_id: int, username: str, code: str, link: str, file_count: int):
    now = datetime.now()
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR IGNORE INTO upload_logs (user_id, username, code, link, file_count, created_date, created_time)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        user_id,
        username or "未知",
        code,
        link,
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
        SELECT id, user_id, username, code, link, file_count, created_time
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

    headers = ["序号", "用户ID", "用户名", "提取码", "直达链接", "包含文件数", "创建时间"]
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

# ----------------- 获取 Telegram 唯一文件指纹 -----------------
def extract_unique_id(message) -> str | None:
    if message.video:
        return message.video.file_unique_id
    elif message.document:
        return message.document.file_unique_id
    elif message.photo:
        return message.photo[-1].file_unique_id
    elif message.audio:
        return message.audio.file_unique_id
    return None

# ----------------- 编解码逻辑 -----------------
def encode_payload(start_id: int, end_id: int) -> str:
    existing_code = get_code_by_range(start_id, end_id)
    if existing_code:
        return existing_code

    salt_prefix = secrets.token_hex(6)
    salt_suffix = secrets.token_hex(4)
    range_str = f"{start_id}-{end_id}" if start_id != end_id else str(start_id)
    raw = f"{salt_prefix}_{range_str}_{salt_suffix}"
    encoded = base64.urlsafe_b64encode(raw.encode()).decode().rstrip("=")
    code = f"{CODE_PREFIX}{encoded}"
    
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

# ----------------- 强制加群/关注核验 -----------------
async def is_user_subscribed(bot, user_id: int) -> bool:
    if not FORCE_SUB_CHANNEL_ID:
        return True
    try:
        member = await bot.get_chat_member(chat_id=FORCE_SUB_CHANNEL_ID, user_id=user_id)
        # 判断状态：普通成员、管理员或群主均算加入
        return member.status in [
            ChatMemberStatus.MEMBER,
            ChatMemberStatus.ADMINISTRATOR,
            ChatMemberStatus.OWNER
        ]
    except Exception:
        return False

# ----------------- 帮助说明（已精简，去除管理员指令与去重说明） -----------------
def get_help_text() -> str:
    return (
        "🤖 <b>媒体文件存取机器人 使用指南</b>\n\n"
        "📌 <b>基本命令：</b>\n"
        "• /start - 启动机器人或提取资源\n"
        "• /help - 查看本帮助说明\n"
        "• /batch - 进入批量打包模式\n"
        "• /done - 结束打包并生成提取码\n\n"
        "📤 <b>保存文件：</b>\n"
        "1. 单文件：直接私聊发媒体文件给 Bot。\n"
        "2. 批量：发送 /batch -> 陆续发文件 -> 发送 /done 结算。\n\n"
        "📥 <b>提取文件：</b>\n"
        "• 直接将提取码发送给机器人即可接收文件。\n"
        "• 文件超过 10 个时，点击底部【📥 查看下一组】继续收取。"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(get_help_text(), parse_mode="HTML")

# ----------------- 上传与打包 -----------------
async def batch_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["batch_mode"] = True
    context.user_data["collected_ids"] = []
    await update.message.reply_text("📦 已开启批量打包模式！请发送需要打包的文件，发送完毕后输入 /done 结算。")

async def save_media(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    unique_id = extract_unique_id(msg)

    # 1. 批量打包模式
    if context.user_data.get("batch_mode"):
        forwarded = await context.bot.copy_message(
            chat_id=DB_CHANNEL_ID,
            from_chat_id=msg.chat_id,
            message_id=msg.message_id
        )
        context.user_data["collected_ids"].append(forwarded.message_id)
        await msg.reply_text(f"已暂存第 {len(context.user_data['collected_ids'])} 个文件...")
        return

    # 2. 单文件去重判断
    if unique_id:
        existing = get_existing_file(unique_id)
        if existing:
            _, code = existing
            bot_me = await context.bot.get_me()
            link = f"https://t.me/{bot_me.username}?start={code}"
            await msg.reply_text(
                f"ℹ️ <b>该文件已存在，已为您匹配已有提取码！</b>\n\n"
                f"提取代码：\n<code>{code}</code>\n\n"
                f"直达链接：\n{link}",
                parse_mode="HTML"
            )
            return

    # 3. 保存新文件
    forwarded = await context.bot.copy_message(
        chat_id=DB_CHANNEL_ID,
        from_chat_id=msg.chat_id,
        message_id=msg.message_id
    )

    code = encode_payload(forwarded.message_id, forwarded.message_id)
    bot_me = await context.bot.get_me()
    link = f"https://t.me/{bot_me.username}?start={code}"

    if unique_id:
        save_file_fingerprint(unique_id, forwarded.message_id, code)

    log_upload(
        user_id=msg.from_user.id,
        username=msg.from_user.username or msg.from_user.first_name,
        code=code,
        link=link,
        file_count=1
    )

    await msg.reply_text(
        f"✅ <b>单文件已保存！</b>\n\n"
        f"提取代码：\n<code>{code}</code>\n\n"
        f"直达链接：\n{link}",
        parse_mode="HTML"
    )

async def done_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get("batch_mode"):
        await update.message.reply_text("未处于打包模式，请先发送 /batch 开启。")
        return

    ids = context.user_data.get("collected_ids", [])
    if not ids:
        context.user_data["batch_mode"] = False
        await update.message.reply_text("未收到任何文件，打包已退出。")
        return

    start_id = min(ids)
    end_id = max(ids)
    code = encode_payload(start_id, end_id)
    bot_me = await context.bot.get_me()
    link = f"https://t.me/{bot_me.username}?start={code}"

    log_upload(
        user_id=update.effective_user.id,
        username=update.effective_user.username or update.effective_user.first_name,
        code=code,
        link=link,
        file_count=len(ids)
    )

    context.user_data["batch_mode"] = False
    context.user_data["collected_ids"] = []

    await update.message.reply_text(
        f"🎉 <b>打包成功！</b>共包含 {len(ids)} 个文件。\n\n"
        f"提取代码：\n<code>{code}</code>\n\n"
        f"直达链接：\n{link}",
        parse_mode="HTML"
    )

# ----------------- 报表管理（静默供管理员使用） -----------------
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

# ----------------- 分组提取逻辑 -----------------
async def deliver_group_page(bot, chat_id: int, start_id: int, end_id: int, current_group: int, code: str):
    total_files = end_id - start_id + 1
    total_groups = math.ceil(total_files / PAGE_SIZE)
    
    page_start_id = start_id + (current_group - 1) * PAGE_SIZE
    page_end_id = min(page_start_id + PAGE_SIZE - 1, end_id)

    success_count = 0
    for mid in range(page_start_id, page_end_id + 1):
        try:
            await bot.copy_message(chat_id=chat_id, from_chat_id=DB_CHANNEL_ID, message_id=mid)
            success_count += 1
            await asyncio.sleep(0.3)
        except BadRequest as e:
            if "Message to copy not found" in str(e):
                continue
            print(f"复制消息 ID {mid} 异常: {e}")
        except TelegramError as e:
            print(f"Telegram 网络或限制异常: {e}")
            continue
        except Exception as e:
            print(f"未知错误: {e}")
            continue

    if success_count == 0 and total_files > 0 and current_group == 1:
        await bot.send_message(
            chat_id=chat_id,
            text="⚠️ <b>提示</b>：该提取码对应的资源文件可能已被删除或失效。",
            parse_mode="HTML"
        )
        return

    keyboard = []
    if current_group < total_groups:
        next_group = current_group + 1
        keyboard.append([InlineKeyboardButton("📥 查看下一组", callback_data=f"page_{code}_{next_group}")])

    share_link = f"https://t.me/share/url?url=https://t.me/{(await bot.get_me()).username}?start={code}"
    keyboard.append([
        InlineKeyboardButton("⭐ 收藏", url=share_link),
        InlineKeyboardButton("📢 防失联", url=FORCE_SUB_INVITE_LINK if FORCE_SUB_INVITE_LINK else "https://t.me")
    ])

    status_text = f"✅ <b>已发送第{current_group}组，共{total_groups}组。</b>"
    await bot.send_message(
        chat_id=chat_id,
        text=status_text,
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="HTML"
    )

async def handle_start_or_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()

    if text == "/start":
        await update.message.reply_text(get_help_text(), parse_mode="HTML")
        return

    if text.startswith("/start "):
        parts = text.split(maxsplit=1)
        code = parts[1].strip()
    else:
        code = text

    payload = decode_payload(code)
    if not payload:
        await update.message.reply_text("❌ 代码无效或格式错误！如有疑问可发送 /help 查看说明。")
        return

    # 强制加群/关注检查
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

    start_id, end_id = payload
    await update.message.reply_text("🚀 验证通过，开始发送第 1 组文件...")
    await deliver_group_page(context.bot, user_id, start_id, end_id, current_group=1, code=code)

async def handle_callback_query(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = update.effective_user.id

    if data.startswith("page_"):
        parts = data.rsplit("_", 1)
        target_group = int(parts[1])
        code = parts[0].replace("page_", "")

        payload = decode_payload(code)
        if not payload:
            await query.answer("❌ 提取码无效或已失效！", show_alert=True)
            return

        start_id, end_id = payload
        await query.message.reply_text(f"🚀 正在发送第 {target_group} 组文件，请稍候...")
        await deliver_group_page(context.bot, user_id, start_id, end_id, current_group=target_group, code=code)

    elif data.startswith("verify_"):
        code = data.replace("verify_", "")
        if not await is_user_subscribed(context.bot, user_id):
            await query.answer("❌ 仍未检测到加入群组，请先加入群再点击！", show_alert=True)
            return

        payload = decode_payload(code)
        if not payload:
            await query.edit_message_text("❌ 代码已过期或失效。")
            return

        await query.edit_message_text("✅ 验证成功，开始发送第 1 组文件...")
        start_id, end_id = payload
        await deliver_group_page(context.bot, user_id, start_id, end_id, current_group=1, code=code)

# ----------------- 主程序入口 -----------------
def main():
    init_db()
    
    app = Application.builder().token(BOT_TOKEN).build()

    # 定时任务：每天 23:59:00（北京时间）自动导出并发送报表给管理员
    shanghai_tz = ZoneInfo("Asia/Shanghai")
    app.job_queue.run_daily(daily_report_job, time=time(hour=23, minute=59, second=0, tzinfo=shanghai_tz))

    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("batch", batch_command))
    app.add_handler(CommandHandler("done", done_command))
    app.add_handler(CommandHandler("export", export_command))
    app.add_handler(CommandHandler("start", handle_start_or_code))

    media_filter = (filters.VIDEO | filters.Document.ALL | filters.PHOTO | filters.AUDIO) & filters.ChatType.PRIVATE
    app.add_handler(MessageHandler(media_filter, save_media))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND & filters.ChatType.PRIVATE, handle_start_or_code))
    app.add_handler(CallbackQueryHandler(handle_callback_query))

    print("Bot 在 Docker 中已启动，系统运行正常...")
    app.run_polling()

if __name__ == "__main__":
    main()