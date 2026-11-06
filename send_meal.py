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
import csv
import datetime as dt
import json
import logging
import os
import re
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


def load_config(require_email=True):
    path = os.path.join(BASE_DIR, "config.json")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    else:
        # GitHub Actions 没有 config.json 时，用环境变量拼一份
        cfg = {
            "menu_file": "menu.csv",
            "canteen": "学员食堂",
            "sender_name": "菜单助手",
            "email": {"smtp_host": "smtp.qq.com", "smtp_port": 465, "use_ssl": True},
        }
    em = cfg.setdefault("email", {})
    em["user"] = os.environ.get("QQ_USER", em.get("user", ""))
    em["auth_code"] = os.environ.get("QQ_AUTH_CODE", em.get("auth_code", ""))
    if os.environ.get("QQ_TO"):
        em["to"] = [x.strip() for x in os.environ["QQ_TO"].split(",") if x.strip()]
    if require_email and (not em.get("user") or not em.get("auth_code") or not em.get("to")):
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


def read_canteen_grid(rows, target_date, meal):
    """识别食堂周食谱版式：第 2 行是「星期X / 2026年11月2日」，A 列餐次、B 列种类。"""
    day_cols = {}
    header_idx = None
    for i, row in enumerate(rows[:8]):
        for j, cell in enumerate(row):
            s = str(cell or "")
            m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", s)
            if m and ("星期" in s or "周" in s):
                day_cols[j] = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                header_idx = i
    if not day_cols:
        return None

    if target_date not in day_cols.values():
        return target_date, meal, []
    col = next(j for j, d in day_cols.items() if d == target_date)

    meal_cn = MEAL_NAMES[meal]
    cur_meal, cur_cat = "", ""
    sections = []
    for row in rows[(header_idx or 0) + 1:]:
        a = str(row[0]).strip() if row and row[0] else ""
        b = str(row[1]).strip().replace("\n", "/") if len(row) > 1 and row[1] else ""
        if a:
            cur_meal = a
        if b:
            cur_cat = b
        if meal_cn not in cur_meal:
            continue
        val = row[col] if len(row) > col else None
        if not val:
            continue
        items = [x.strip() for x in re.split(r"[、,，\n;；]", str(val)) if x.strip()]
        if not items:
            continue
        if sections and sections[-1][0] == cur_cat:
            sections[-1][1].extend(items)
        else:
            sections.append([cur_cat, items])
    return target_date, meal, sections


def read_csv_menu(path, target_date, meal):
    """读取 menu.csv：日期,餐次,分类,菜品。"""
    sections = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            d = parse_date(row.get("日期", ""), target_date)
            if d != target_date:
                continue
            m = MEAL_ALIASES.get(str(row.get("餐次", "")).strip())
            if m != meal:
                continue
            cat = str(row.get("分类", "")).strip().replace("\n", "/")
            items = split_dishes(str(row.get("菜品", "")))
            if not items:
                continue
            if sections and sections[-1][0] == cat:
                sections[-1][1].extend(items)
            else:
                sections.append([cat, items])
    return target_date, meal, sections


def read_menu(menu_path, target_date, meal):
    """读取菜单，返回 (日期, 餐次, [(分类, 菜品列表)])。

    支持：menu.csv、食堂周食谱、长表（日期/餐次/菜品）、宽表（日期/早餐/午餐/晚餐）。
    """
    if str(menu_path).lower().endswith(".csv"):
        return read_csv_menu(menu_path, target_date, meal)

    wb = openpyxl.load_workbook(menu_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise SystemExit("菜单文件是空的。")

    grid = read_canteen_grid(rows, target_date, meal)
    if grid is not None:
        return grid

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
                return d, m, [(None, split_dishes(text))]
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
            return d, meal, [(None, split_dishes(text))]
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


def render_menu_image(date, meal, sections, canteen, out_path):
    W = 900
    f_title = get_font(56, bold=True)
    f_date = get_font(32)
    f_meal = get_font(44, bold=True)
    f_cat = get_font(30, bold=True)
    f_dish = get_font(32)
    f_foot = get_font(22)

    # 先量一遍高度，菜品多时自动加长
    tmp = Image.new("RGB", (10, 10))
    td = ImageDraw.Draw(tmp)
    ops = []
    content_h = 0
    for cat, items in sections:
        if cat:
            ops.append(("cat", cat))
            content_h += 56
        for item in items:
            lines = wrap_text(td, item, f_dish, W - 260)
            ops.append(("item", lines))
            content_h += 48 * len(lines) + 12

    header_h = 390 + (70 if canteen else 0)
    H = max(1200, int(header_h + content_h + 140))

    img = Image.new("RGB", (W, H), "#FBF7EE")
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, W, 16], fill="#B8860B")

    title = "今日菜单"
    tw = draw.textlength(title, font=f_title)
    draw.text(((W - tw) / 2, 70), title, font=f_title, fill="#3A2E12")

    date_text = f"{date.year} 年 {date.month} 月 {date.day} 日"
    dw = draw.textlength(date_text, font=f_date)
    draw.text(((W - dw) / 2, 150), date_text, font=f_date, fill="#6B5B33")

    meal_text = MEAL_NAMES[meal]
    mw = draw.textlength(meal_text, font=f_meal)
    draw.rounded_rectangle([(W - mw) / 2 - 40, 215, (W + mw) / 2 + 40, 300], radius=20, fill="#B8860B")
    draw.text(((W - mw) / 2, 228), meal_text, font=f_meal, fill="white")

    y = 350
    if canteen:
        draw.text((90, y), canteen, font=f_date, fill="#6B5B33")
        y += 70
    y += 20

    if not ops:
        draw.text((90, y), "今日暂无菜单", font=f_dish, fill="#A33B28")
    else:
        for kind, value in ops:
            if kind == "cat":
                draw.rectangle([90, y + 6, 98, y + 38], fill="#B8860B")
                draw.text((114, y), value, font=f_cat, fill="#8A6D1F")
                y += 56
            else:
                draw.ellipse([110, y + 14, 122, y + 26], fill="#B8860B")
                for line in value:
                    draw.text((150, y), line, font=f_dish, fill="#2B2B2B")
                    y += 48
                y += 12

    foot = "祝用餐愉快"
    fw = draw.textlength(foot, font=f_foot)
    draw.text(((W - fw) / 2, H - 70), foot, font=f_foot, fill="#9A8A60")

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
    cfg = load_config(require_email=not args.no_send)
    today = dt.date.today()
    if args.date == "today":
        target = today
    elif args.date == "tomorrow":
        target = today + dt.timedelta(days=1)
    else:
        target = dt.datetime.strptime(args.date, "%Y-%m-%d").date()

    menu_path = os.path.join(BASE_DIR, cfg.get("menu_file", "menu.csv"))
    if not os.path.exists(menu_path):
        for alt in ("menu.csv", "menu.xlsx"):
            p = os.path.join(BASE_DIR, alt)
            if os.path.exists(p):
                menu_path = p
                break
    date, meal, sections = read_menu(menu_path, target, args.meal)

    os.makedirs(OUT_DIR, exist_ok=True)
    out_path = os.path.join(OUT_DIR, f"{date}_{args.meal}.png")
    render_menu_image(date, meal, sections, cfg.get("canteen", ""), out_path)
    logging.info("生成图片 %s，菜单 %s", out_path, sections)

    body_lines = []
    for cat, items in sections:
        if items:
            prefix = f"【{cat}】" if cat else "· "
            body_lines.append(prefix + "、".join(items))
    subject = f"{date.month}月{date.day}日 {MEAL_NAMES[args.meal]}菜单"
    body = f"{date} {MEAL_NAMES[args.meal]}菜单见附件。\n\n" + ("\n".join(body_lines) if body_lines else "暂无菜单")

    if args.no_send:
        print("已生成（未发送）：", out_path)
        return

    send_email(cfg, subject, body, out_path)
    logging.info("已发送邮件：%s", subject)
    print("已发送：", subject)


if __name__ == "__main__":
    main()
