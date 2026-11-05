# 每日菜单图片 · 定时发送

按周录入菜单，程序在固定时间取出对应餐次、生成图片，发到 QQ 邮箱。

- **运行在 GitHub Actions**：不依赖电脑开机；
- 菜单用 **`menu.csv` 纯文本**，每周在手机 GitHub App 里改几行就行。

## 目录

```
menu-bot/
├─ send_meal.py                 主程序
├─ menu.csv                    周菜单（推荐，纯文本可编辑）
├─ menu.xlsx                   食堂周食谱 Excel（备用，程序也支持）
├─ config.example.json         本地调试配置模板
├─ requirements.txt
├─ .github/workflows/send-menu.yml
├─ install_tasks.ps1           本地 Windows 定时任务（可选）
├─ output/                     生成图片
└─ logs/send.log
```

## 一、每周菜单怎么改（手机就能改）

打开仓库里的 `menu.csv`，点右上角铅笔编辑，按四列填写：

```csv
日期,餐次,分类,菜品
2026-11-09,早餐,小菜,凉拌黄瓜、炝拌豆芽
2026-11-09,早餐,主食/小吃,馒头、花卷、鸡蛋
2026-11-09,早餐,饮品粥,小米粥、牛奶
2026-11-09,午餐,特色/菜品,红烧肉、清炒时蔬
2026-11-09,午餐,主食/小吃,米饭、馒头
2026-11-09,晚餐,菜品,土豆牛肉、蒜蓉西兰花
```

规则：

- **日期**：`2026-11-09` 这种格式；
- **餐次**：只能写 `早餐` / `午餐` / `晚餐`；
- **分类**：随意写，会作为小标题显示（小菜、主食/小吃、饮品粥……）；
- **菜品**：同一行用顿号 `、` 分隔，程序会拆成一条条。

改完点 **Commit changes** 即可，下一次定时自动用新菜单。

> 也可以继续上传食堂的 Excel 周食谱，程序会自动识别「餐次 / 种类 / 星期」版式；但 CSV 在手机上更好改。

## 二、定时

| 北京时间 | 内容 |
|---|---|
| 21:15 | 次日早餐 |
| 11:00 | 当天午餐 |
| 16:00 | 当天晚餐 |

GitHub 的 cron 用 UTC（已换算好），可能有几分钟延迟。

## 三、配置（已配好，仅记录）

仓库 → Settings → Secrets and variables → Actions：

| Name | 说明 |
|---|---|
| `QQ_USER` | 发件 QQ 邮箱 |
| `QQ_AUTH_CODE` | QQ 邮箱 SMTP 授权码 |
| `QQ_TO` | 收件邮箱，多个用英文逗号分隔 |

## 四、手动测试

Actions →「每日菜单邮件」→ Run workflow：

- `meal` 选餐次；
- `send_date` 选 today / tomorrow；
- `dry_run` 选 `true` 只生成不发送，选 `false` 真发。

## 五、本地调试（可选）

1. 复制 `config.example.json` 为 `config.json`，填邮箱和授权码；
2. 运行：

```bash
python send_meal.py --meal lunch --no-send   # 只出图
python send_meal.py --meal supper            # 真发
```

## 六、注意

- GitHub 定时可能延迟几分钟；仓库连续 60 天无提交会暂停定时，改一次菜单即可恢复；
- 菜单日期覆盖不到当天时，会显示「今日暂无菜单」；
- 图片高度会根据菜品数量自动加长。
