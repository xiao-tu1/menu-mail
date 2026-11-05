# 每日菜单图片 · 定时发送

按周录入菜单，程序在固定时间取出对应餐次、生成图片，发到 QQ 邮箱。

两种跑法：

- **GitHub Actions（推荐）**：不依赖你的电脑开机，免费；
- **本地 Windows 任务**：电脑开着才准点，适合调试。

## 目录

```
menu-bot/
├─ send_meal.py                主程序
├─ menu.xlsx                   周菜单（当前是示例）
├─ menu_template.xlsx          模板
├─ config.example.json         本地配置模板
├─ requirements.txt            Pillow / openpyxl
├─ .github/workflows/send-menu.yml   GitHub Actions 定时任务
├─ install_tasks.ps1           本地 Windows 定时任务（可选）
├─ output/                     生成的图片
└─ logs/send.log               日志
```

## 一、菜单格式

`menu.xlsx` 支持两种：

**长表（推荐）**

| 日期 | 餐次 | 菜品 |
|---|---|---|
| 2026-11-05 | 早餐 | 小米粥、鸡蛋、鲜肉包 |
| 2026-11-05 | 午餐 | 红烧肉、青菜、米饭 |

**宽表**

| 日期 | 早餐 | 午餐 | 晚餐 |
|---|---|---|---|
| 2026-11-05 | 小米粥、鸡蛋 | 红烧肉、青菜 | 土豆牛肉、米饭 |

菜品用顿号、逗号、分号、换行分隔均可；日期支持 `2026-11-05`、`11月5日`、`11/5`、`周一`。

## 二、GitHub Actions 部署（推荐）

### 1. 建仓库

在 GitHub 新建一个**私有仓库**（Private），名字随意，例如 `menu-mail`。

### 2. 上传文件

把本目录里的这些文件传到仓库根目录：

```
send_meal.py
menu.xlsx
menu_template.xlsx
requirements.txt
config.example.json
.github/workflows/send-menu.yml
```

不要上传 `config.json`（里面有授权码，`.gitignore` 已排除）。

### 3. 配置 Secrets

仓库 → **Settings → Secrets and variables → Actions → New repository secret**，添加三个：

| Name | Value |
|---|---|
| `QQ_USER` | 你的 QQ 邮箱，如 `123456@qq.com` |
| `QQ_AUTH_CODE` | QQ 邮箱 SMTP 授权码（不是 QQ 密码） |
| `QQ_TO` | 收件邮箱，可多个，用英文逗号分隔 |

授权码获取：QQ 邮箱 → 设置 → 账户 → 开启「POP3/IMAP/SMTP 服务」→ 生成授权码。

### 4. 先手动跑一次

仓库 → **Actions** → 左侧「每日菜单邮件」→ **Run workflow**：

- meal 选 `lunch`
- send_date 选 `today`

跑完看日志和邮箱，确认能收到。

### 5. 定时规则

工作流里已经写好（UTC 时间，北京时间 = UTC+8）：

| UTC cron | 北京时间 | 内容 |
|---|---|---|
| `15 13 * * *` | 21:15 | 次日早餐 |
| `0 3 * * *` | 11:00 | 当天午餐 |
| `0 8 * * *` | 16:00 | 当天晚餐 |

### 6. 两个注意点

- GitHub 的定时任务**可能延迟几分钟**（高峰期常见 5–15 分钟），不能做到秒级准点；
- 仓库连续 **60 天没有任何提交**，GitHub 会自动暂停定时任务。每周更新菜单时提交一次即可保持。

## 三、本地运行（调试用）

1. 复制 `config.example.json` 为 `config.json`，填邮箱和授权码；
2. 测试：

```bash
python send_meal.py --meal lunch --no-send   # 只出图
python send_meal.py --meal lunch             # 真发邮件
```

3. 注册 Windows 定时任务（可选）：

```powershell
powershell -ExecutionPolicy Bypass -File install_tasks.ps1
```

电脑关机/睡眠时不会触发；要准点请用 GitHub Actions。

## 四、常见问题

**Actions 里报邮箱信息不完整**：三个 Secrets 没配全，或名字拼错。

**收不到邮件**：先看 Actions 日志；QQ 邮箱必须用授权码，且 `QQ_TO` 要写完整邮箱地址。

**图片里没有菜**：检查 `menu.xlsx` 的日期是否是今天（早餐是明天），餐次是否写对。

**想换发送渠道**：改 `send_email()` 即可；微信自动化稳定性不如邮箱，建议保留邮箱为主通道。
