# -*- coding: utf-8 -*-
"""把当天菜单以文字形式推到微信公众号（PushPlus）"""
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

    title = f"{date.month}月{date.day}日 {send_meal.MEAL_NAMES[args.meal]}菜单"
    parts = [f"<h3>{title}</h3>"]
    for cat, items in sections:
        if items:
            prefix = f"<b>{cat}</b>：" if cat else ""
            parts.append(f"<p>{prefix}{'、'.join(items)}</p>")
    if len(parts) == 1:
        parts.append("<p>今日暂无菜单</p>")
    html = "".join(parts)

    payload = {
        "token": os.environ["PUSHPLUS_TOKEN"],
        "title": title,
        "content": html,
        "template": "html",
    }
    topic = os.environ.get("PUSHPLUS_TOPIC", "").strip()
    if topic:
        payload["topic"] = topic

    req = urllib.request.Request(
        "https://www.pushplus.plus/send",
        data=json.dumps(payload).encode("utf-8"),
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
