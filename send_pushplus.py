# -*- coding: utf-8 -*-
"""把当天菜单图片推到微信公众号（PushPlus）"""
import argparse
import datetime as dt
import json
import os
import urllib.request

import send_meal

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--meal", required=True, choices=["breakfast", "lunch", "supper"])
    parser.add_argument("--date", default="today", help="today / tomorrow / YYYY-MM-DD")
    args = parser.parse_args()

    today = dt.date.today()
    if args.date == "today":
        target = today
    elif args.date == "tomorrow":
        target = today + dt.timedelta(days=1)
    else:
        target = dt.datetime.strptime(args.date, "%Y-%m-%d").date()

    menu_path = os.path.join(BASE_DIR, "menu.csv")
    date, meal, sections = send_meal.read_menu(menu_path, target, args.meal)

    filename = f"{date}_{args.meal}.png"
    image_path = os.path.join(BASE_DIR, "output", filename)
    if not os.path.exists(image_path):
        raise SystemExit("图片还没生成：" + image_path)

    repo = os.environ.get("REPO", "xiao-tu1/menu-mail")
    image_url = f"https://cdn.jsdelivr.net/gh/{repo}@main/output/{filename}"

    title = f"{date.month}月{date.day}日 {send_meal.MEAL_NAMES[args.meal]}菜单"
    lines = []
    for cat, items in sections:
        if items:
            prefix = f"{cat}：" if cat else ""
            lines.append(prefix + "、".join(items))
    html = (
        f"<h2>{title}</h2>"
        f"<p><img src='{image_url}' style='max-width:100%;border-radius:8px'></p>"
        + ("<hr>" + "<br>".join(lines) if lines else "")
    )

    token = os.environ["PUSHPLUS_TOKEN"]
    payload = json.dumps({
        "token": token,
        "title": title,
        "content": html,
        "template": "html",
    }).encode("utf-8")
    req = urllib.request.Request(
        "https://www.pushplus.plus/send",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        body = r.read().decode("utf-8", "replace")
    print(body)
    try:
        if json.loads(body).get("code") != 200:
            raise SystemExit("PushPlus 返回失败")
    except json.JSONDecodeError:
        raise SystemExit("PushPlus 返回不是 JSON")


if __name__ == "__main__":
    main()
