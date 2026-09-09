# Telegram File Store Bot (媒体文件存取与分发机器人)

基于 `python-telegram-bot` (v20+) 与 Docker Compose 构建的高性能 Telegram 私有文件存取与批量分发机器人。支持长串混淆提取码生成、批量文件打包、10个/组分页查看、强制加群/关注频道验证、文件级指纹去重以及每日自动导出 Excel 报表。

---

## ✨ 核心特性

- 📦 **单文件 / 批量打包上传**：
  - 支持视频、文档、图片、音频等全格式媒体文件。
  - 支持 `/batch` 批量收录模式与 `/done` 结算，自动打包区间文件。
- 📑 **10个/组 交互式分页发货**：
  - 提取多文件时按 10 个自动分组，防刷屏并避免触发 Telegram 频率限制（Flood Wait）。
  - 提供 `📥 查看下一组` 动态交互按钮，发完自动隐藏。
  - 内置 `⭐ 收藏` 与 `📢 防失联` 快捷跳转。
- 🔍 **文件指纹去重 (Deduplication)**：
  - 基于 Telegram 官方 `file_unique_id` 指纹机制。
  - 重复上传同一文件自动复用已有提取码，避免私有存储频道冗余。
  - 每日报表自动去重，同一提取码仅统计 1 份。
- 🛡️ **双重解码与持久化防丢**： 
  - 混淆 Base64 提取码与 SQLite 本地持久化映射双重校验，容器重启后提取码永久有效。 
  - 容错机制：自动捕获并跳过频道中被删除的历史消息（`Message to copy not found`），防止进程崩溃。 
- 👥 **强制加群 / 频道关注验证**： 
  - 未加入指定群组或频道的用户无法提取，提取前自动弹出引导与“我已加入”实时核验。 
- 📊 **定时任务与每日 Excel 汇总**： 
  - 每天 23:59 (Asia/Shanghai) 自动聚合当日生成记录，生成 `.xlsx` 表格并静默私发给管理员。 
  - 支持管理员专属 `/export` 指令随时手动导出。 

---

## 🚀 快速部署指南 

### 1. 克隆代码仓库 

```bash
mkdir -p /root/tg-file-store-bot && cd /root/tg-file-store-bot
git clone https://github.com/game315422/tg-file-store-bot.git
```

### 2. 配置环境变量 (.env) 

在项目根目录下新建 `.env` 文件： 

```bash
cat << 'EOF' > .env
# Telegram Bot Token (从 @BotFather 获取)
BOT_TOKEN=1234567890:ABCdefGHIjklMNOpqrSTUvwxYZ

# 管理员 Telegram 数字 ID (用于接收每日 Excel 报表和执行 /export)
ADMIN_USER_ID=123456789

# 文件存储私有频道的 ID (Bot 需作为管理员加入，格式一般为 -100xxxxxxxxxx)
DB_CHANNEL_ID=-1001234567890

# 强制加入的官方群组/频道 ID (Bot 需作为管理员加入)
FORCE_SUB_CHANNEL_ID=-1009876543210

# 官方群组/频道的邀请链接 (用户未加群时弹出引导)
FORCE_SUB_INVITE_LINK=https://t.me/your_group_link
EOF
```

### 3. 一键构建并启动 (Docker Compose) 

```bash
docker compose up -d --build
```

查看实时运行日志： 

```bash
docker compose logs -f
```

当日志显示 `Bot 在 Docker 中已启动，系统运行正常...` 即表示运行成功。 

---

## 📖 使用说明 

### 普通用户 

| 操作 / 指令 | 说明 |
| :--- | :--- |
| `/start` | 查看机器人使用说明与功能引导 |
| `/help` | 查看详细帮助菜单 |
| **发送任意媒体文件** | 自动转存到私有数据库，并返回专属提取码与直达链接 |
| `/batch` | 开启批量打包模式，之后连续发送的所有媒体将被暂存 |
| `/done` | 结束批量打包，生成打包提取码与链接 |
| **发送提取码** | 校验群组/频道成员资格后，按 10 个一组分页获取文件 |

### 管理员专属 

| 指令 | 说明 |
| :--- | :--- |
| `/export` | 仅管理员可用，随时手动导出当日的 Excel 上传去重报表 |
| *每日 23:59:00* | 系统自动将今日产生的 `.xlsx` 报表推送到管理员私聊 |

---

## 📂 项目结构 

```text
.
├── Dockerfile              # Docker 镜像构建配置 (Python 3.11-slim)
├── docker-compose.yml      # 容器编排配置 (包含数据目录持久化挂载)
├── requirements.txt        # Python 依赖清单 (python-telegram-bot[job-queue], openpyxl)
├── bot.py                  # 机器人完整业务逻辑代码
├── .gitignore              # Git 忽略配置 (保护 .env 密钥与本地数据库)
├── data/                   # 挂载目录 (存放 records.db 与每日导出的 Excel 表格)
└── README.md               # 项目使用与部署文档
```

---

## 🛠️ 运维与更新 

- **查看容器状态**：`docker compose ps` 
- **重启机器人**：`docker compose restart` 
- **停止运行**：`docker compose down` 
- **拉取更新代码并重构**： 
  ```bash
  git pull
  docker compose down
  docker compose up -d --build
  ```
