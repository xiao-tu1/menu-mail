# -*- coding: utf-8 -*-
"""
每日餐单图片定时发送工具

用法：
    python send_meal.py --meal lunch
    python send_meal.py --meal breakfast --date tomorrow
    python send_meal.py --meal supper --no-send

菜单文件：menu.xlsx（可用 menu_template.xlsx 改名）
配置文件：config.json（参考 config.example.json）
"""
import argparse
import datetime as dt
import json
import logging
import os
import smtplib
import sys
from email.header import Header
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from PIL import Image, ImageDraw, ImageFont
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(BASE_DIR, "logs")
OUT_DIR = os.path.join(BASE_DIR, "output")

MEAL_NAMES = {
    "breakfast": "早餐",
    "lunch": "午餐",
    "supper": "晚餐",
}
MEAL_ALIASES = {
    "早餐": "breakfast", "早饭": "breakfast", "breakfast": "breakfast",
    "午餐": "lunch", "午饭": "lunch", "中餐": "lunch", "lunch": "lunch",
    "晚餐": "supper", "晚饭": "supper", "下午饭": "supper", "supper": "supper", "dinner": "supper",
}

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
]


def get_font(size, bold=False):
    paths = FONT_CANDIDATES[:2] if bold else FONT_CANDIDATES
    for p in paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


def setup_logging():
    os.makedirs(LOG_DIR, exist_ok=True)
    logging.basicConfig(
        filename=os.path.join(LOG_DIR, "send.log"),
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        encoding="utf-8",
    )


def load_config():
    path = os.path.join(BASE_DIR, "config.json")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    else:
        # GitHub Actions 没有 config.json 时，用环境变量拼一份
        cfg = {
            "menu_file": "menu.xlsx",
            "canteen": "学员食堂",
            "sender_name": "菜单助手",
            "email": {"smtp_host": "smtp.qq.com", "smtp_port": 465, "use_ssl": True},
        }
    em = cfg.setdefault("email", {})
    em["user"] = os.environ.get("QQ_USER", em.get("user", ""))
    em["auth_code"] = os.environ.get("QQ_AUTH_CODE", em.get("auth_code", ""))
    if os.environ.get("QQ_TO"):
        em["to"] = [x.strip() for x in os.environ["QQ_TO"].split(",") if x.strip()]
    if not em.get("user") or not em.get("auth_code") or not em.get("to"):
        raise SystemExit("邮箱信息不完整：请检查 config.json 或 QQ_USER / QQ_AUTH_CODE / QQ_TO 环境变量。")
    return cfg


def parse_date(value, today):
    """把 Excel 里的日期转成 date；支持 datetime、2026-11-05、11月5日、周一。"""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m-%d", "%m/%d"):
        try:
            d = dt.datetime.strptime(s, fmt).date()
            if d.year == 1900:
                d = d.replace(year=today.year)
            return d
        except ValueError:
            pass
    if "月" in s and "日" in s:
        try:
            md = s.replace("日", "").split("月")
            return dt.date(today.year, int(md[0]), int(md[1]))
        except Exception:
            pass
    weekdays = {"周一": 0, "周二": 1, "周三": 2, "周四": 3, "周五": 4, "周六": 5, "周日": 6, "周天": 6}
    if s in weekdays:
        monday = today - dt.timedelta(days=today.weekday())
        return monday + dt.timedelta(days=weekdays[s])
    return None


def read_menu(menu_path, target_date, meal):
    """读取 Excel，返回 (日期, 餐次, 菜品列表)。支持长表和宽表。"""
    wb = openpyxl.load_workbook(menu_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise SystemExit("菜单文件是空的。")

    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    # 长表：日期 | 餐次 | 菜品
    if any("餐次" in h for h in header):
        i_date = next(i for i, h in enumerate(header) if "日期" in h)
        i_meal = next(i for i, h in enumerate(header) if "餐次" in h)
        i_dish = next(i for i, h in enumerate(header) if "菜" in h or "内容" in h)
        for row in rows[1:]:
            if not row or row[i_date] is None:
                continue
            d = parse_date(row[i_date], target_date)
            m = MEAL_ALIASES.get(str(row[i_meal]).strip())
            if d == target_date and m == meal:
                text = str(row[i_dish] or "").strip()
                return d, m, split_dishes(text)
        return target_date, meal, []

    # 宽表：日期 | 早餐 | 午餐 | 晚餐
    meal_col = None
    for i, h in enumerate(header):
        if MEAL_ALIASES.get(h) == meal:
            meal_col = i
            break
    if meal_col is None:
        raise SystemExit(f"菜单表头里找不到「{MEAL_NAMES[meal]}」列。")
    i_date = 0
    for row in rows[1:]:
        if not row or row[i_date] is None:
            continue
        d = parse_date(row[i_date], target_date)
        if d == target_date:
            text = str(row[meal_col] or "").strip()
            return d, meal, split_dishes(text)
    return target_date, meal, []


def split_dishes(text):
    if not text:
        return []
    for sep in ("、", "，", ",", ";", "；", "\n", "|"):
        text = text.replace(sep, "\n")
    return [x.strip() for x in text.splitlines() if x.strip()]


def wrap_text(draw, text, font, max_width):
    lines, cur = [], ""
    for ch in text:
        if draw.textlength(cur + ch, font=font) <= max_width:
            cur += ch
        else:
            lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)
    return lines


def render_menu_image(date, meal, dishes, canteen, out_path):
    W, H = 900, 1200
    img = Image.new("RGB", (W, H), "#FBF7EE")
    draw = ImageDraw.Draw(img)

    # 顶部色带
    draw.rectangle([0, 0, W, 16], fill="#B8860B")

    f_title = get_font(60, bold=True)
    f_date = get_font(34)
    f_meal = get_font(46, bold=True)
    f_dish = get_font(34)
    f_foot = get_font(24)

    title = "今日菜单"
    tw = draw.textlength(title, font=f_title)
    draw.text(((W - tw) / 2, 70), title, font=f_title, fill="#3A2E12")

    date_text = f"{date.year} 年 {date.month} 月 {date.day} 日"
    dw = draw.textlength(date_text, font=f_date)
    draw.text(((W - dw) / 2, 160), date_text, font=f_date, fill="#6B5B33")

    meal_text = MEAL_NAMES[meal]
    mw = draw.textlength(meal_text, font=f_meal)
    draw.rounded_rectangle([(W - mw) / 2 - 40, 230, (W + mw) / 2 + 40, 320], radius=20, fill="#B8860B")
    draw.text(((W - mw) / 2, 245), meal_text, font=f_meal, fill="white")

    y = 390
    if canteen:
        draw.text((90, y), canteen, font=f_dish, fill="#6B5B33")
        y += 70

    if not dishes:
        draw.text((90, y), "今日暂无菜单", font=f_dish, fill="#A33B28")
    else:
        for item in dishes:
            draw.ellipse([90, y + 16, 104, y + 30], fill="#B8860B")
            for line in wrap_text(draw, item, f_dish, W - 220):
                draw.text((130, y), line, font=f_dish, fill="#2B2B2B")
                y += 54
            y += 14

    foot = "祝用餐愉快"
    fw = draw.textlength(foot, font=f_foot)
    draw.text(((W - fw) / 2, H - 80), foot, font=f_foot, fill="#9A8A60")

    img.save(out_path, "PNG")
    return out_path


def send_email(cfg, subject, body, image_path):
    em = cfg["email"]
    msg = MIMEMultipart()
    msg["From"] = em["user"]
    msg["To"] = ", ".join(em["to"])
    msg["Subject"] = Header(subject, "utf-8")

    msg.attach(MIMEText(body, "plain", "utf-8"))
    with open(image_path, "rb") as f:
        img = MIMEImage(f.read(), _subtype="png")
        img.add_header("Content-Disposition", "attachment",
                       filename=("utf-8", "", os.path.basename(image_path)))
        msg.attach(img)

    if em.get("use_ssl", True):
        server = smtplib.SMTP_SSL(em["smtp_host"], em["smtp_port"], timeout=30)
    else:
        server = smtplib.SMTP(em["smtp_host"], em["smtp_port"], timeout=30)
        server.starttls()
    server.login(em["user"], em["auth_code"])
    server.sendmail(em["user"], em["to"], msg.as_string())
    server.quit()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--meal", required=True, choices=["breakfast", "lunch", "supper"])
    parser.add_argument("--date", default="today", help="today / tomorrow / YYYY-MM-DD")
    parser.add_argument("--no-send", action="store_true", help="只生成图片，不发邮件")
    args = parser.parse_args()

    setup_logging()
    cfg = load_config()
    today = dt.date.today()
    if args.date == "today":
        target = today
    elif args.date == "tomorrow":
        target = today + dt.timedelta(days=1)
    else:
        target = dt.datetime.strptime(args.date, "%Y-%m-%d").date()

    menu_path = os.path.join(BASE_DIR, cfg.get("menu_file", "menu.xlsx"))
    date, meal, dishes = read_menu(menu_path, target, args.meal)

    os.makedirs(OUT_DIR, exist_ok=True)
    out_path = os.path.join(OUT_DIR, f"{date}_{args.meal}.png")
    render_menu_image(date, meal, dishes, cfg.get("canteen", ""), out_path)
    logging.info("生成图片 %s，菜品 %s", out_path, dishes)

    subject = f"{date.month}月{date.day}日 {MEAL_NAMES[args.meal]}菜单"
    body = f"{date} {MEAL_NAMES[args.meal]}菜单见附件。\n\n" + ("\n".join("· " + d for d in dishes) if dishes else "暂无菜单")

    if args.no_send:
        print("已生成（未发送）：", out_path)
        return

    send_email(cfg, subject, body, out_path)
    logging.info("已发送邮件：%s", subject)
    print("已发送：", subject)


if __name__ == "__main__":
    main()
