#!/usr/bin/env python3
import argparse
import ast
import sys
import time

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# Config
SAMPLE_SIZE = 1000
RANDOM_SEED = 42
MIN_TAGALOG = 0.15
MIN_ENGLISH = 0.5
API_URL = "https://api.twitterapi.io/twitter/tweets"
BATCH_SIZE = 100

# Colours
CLR_HEADER_BG = "1DA1F2"
CLR_HEADER_FG = "FFFFFF"
CLR_ROW_A = "E8F5E9"
CLR_ROW_B = "F5F5F5"
CLR_ERROR_BG = "FFEBEE"
CLR_ANNOT_BG = "FFFDE7"


def parse_split(raw: str):
    try:
        vals = ast.literal_eval(raw)
        return float(vals[0]), float(vals[1]), float(vals[2])
    except Exception:
        return None, None, None


def fetch_batch(ids: list[str], api_key: str) -> dict:
    headers = {"X-API-Key": api_key}
    params = {"tweet_ids": ",".join(ids)}

    try:
        resp = requests.get(API_URL, headers=headers, params=params, timeout=30)
    except requests.RequestException as exc:
        print(f"[WARN] Request failed: {exc}", flush=True)
        return {}

    if resp.status_code == 401:
        sys.exit("[ERROR] Invalid API key.")

    if resp.status_code == 429:
        print("[Rate limited] Waiting 5 s ...", flush=True)
        time.sleep(5)
        return fetch_batch(ids, api_key)

    if not resp.ok:
        print(f"[WARN] HTTP {resp.status_code}: {resp.text[:200]}", flush=True)
        return {}

    out = {}
    for tweet in resp.json().get("tweets", []):
        tweet_id = tweet.get("id")
        if tweet_id:
            out[str(tweet_id)] = tweet
    return out


def thin_border():
    side = Side(style="thin", color="CCCCCC")
    return Border(left=side, right=side, top=side, bottom=side)


def header_font():
    return Font(bold=True, color=CLR_HEADER_FG, name="Arial", size=10)


def apply_header(ws, headers: list[str]):
    fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
    for col, label in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=label)
        cell.font = header_font()
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()
    ws.row_dimensions[1].height = 30


def main():
    parser = argparse.ArgumentParser(
        description="Fetch code-switched tweets via twitterapi.io and export to Excel."
    )
    parser.add_argument("--csv", required=True, help="Path to tweets_split_id.csv")
    parser.add_argument("--api-key", required=True, help="Your twitterapi.io API key")
    parser.add_argument("--output", default="codeswitched_tweets_othermaj.xlsx")
    args = parser.parse_args()

    # 1. Load, parse, filter, and sample.
    print(f"[1/4] Loading {args.csv} ...")
    df = pd.read_csv(args.csv, dtype={"id": str})
    df[["tagalog", "english", "other"]] = df["split"].apply(
        lambda value: pd.Series(parse_split(value))
    )
    filtered = df[df["other"] >= 0.50]
    sample = filtered.sample(SAMPLE_SIZE, random_state=RANDOM_SEED).reset_index(drop=True)
    print(f" {len(filtered):,} rows match filter -> sampling {SAMPLE_SIZE}")

    # 2. Fetch tweets.
    print(f"\n[2/4] Fetching {SAMPLE_SIZE} tweets from twitterapi.io ...")
    ids = sample["id"].tolist()
    all_tweets = {}
    n_batches = (len(ids) + BATCH_SIZE - 1) // BATCH_SIZE

    for i in range(0, len(ids), BATCH_SIZE):
        batch = ids[i : i + BATCH_SIZE]
        batch_num = i // BATCH_SIZE + 1
        print(f" Batch {batch_num}/{n_batches} ({len(batch)} IDs) ...", flush=True)
        all_tweets.update(fetch_batch(batch, args.api_key))
        time.sleep(0.2)

    found = len(all_tweets)
    missing = SAMPLE_SIZE - found
    print(f" Fetched: {found} | Unavailable/deleted: {missing}")

    # 3. Assemble rows.
    print("\n[3/4] Assembling data ...")
    rows = []
    for _, csv_row in sample.iterrows():
        tweet_id = csv_row["id"]
        tweet = all_tweets.get(str(tweet_id))
        error = None if tweet else "not returned by API"
        author = (tweet or {}).get("author", {})

        rows.append(
            {
                "tweet_id": tweet_id,
                "status": "unavailable" if error else "ok",
                "error_detail": error or "",
                "text": (tweet or {}).get("text", ""),
                "created_at": str((tweet or {}).get("createdAt", ""))[:10],
                "api_lang": (tweet or {}).get("lang", ""),
                "author_id": author.get("id", ""),
                "author_name": author.get("name", ""),
                "username": author.get("userName", ""),
                "tagalog_pct": round(csv_row["tagalog"] * 100, 1),
                "english_pct": round(csv_row["english"] * 100, 1),
                "other_pct": round(csv_row["other"] * 100, 1),
                "likes": (tweet or {}).get("likeCount", ""),
                "retweets": (tweet or {}).get("retweetCount", ""),
                "replies": (tweet or {}).get("replyCount", ""),
                "quotes": (tweet or {}).get("quoteCount", ""),
                "views": (tweet or {}).get("viewCount", ""),
                "cs_words": "",
                "pos_tags": "",
                "switch_type": "",
                "notes": "",
            }
        )

    # 4. Write the Excel workbook.
    print(f"\n[4/4] Writing {args.output} ...")
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Tweets"
    ws1.freeze_panes = "A2"

    headers = [
        "Tweet ID", "Status", "Error Detail", "Text", "Date", "API Lang",
        "Author ID", "Author Name", "Username", "Tagalog %", "English %", "Other %",
        "Likes", "Retweets", "Replies", "Quotes", "Views", "Code-Switched Words",
        "POS Tags", "Switch Type", "Notes",
    ]
    apply_header(ws1, headers)

    col_widths = [20, 12, 28, 80, 12, 10, 20, 20, 18, 11, 11, 11, 8, 10, 8, 8, 10, 30, 20, 18, 25]
    for i, width in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = width

    fill_a = PatternFill("solid", fgColor=CLR_ROW_A)
    fill_b = PatternFill("solid", fgColor=CLR_ROW_B)
    fill_err = PatternFill("solid", fgColor=CLR_ERROR_BG)
    fill_ann = PatternFill("solid", fgColor=CLR_ANNOT_BG)

    for row_index, row in enumerate(rows, 2):
        values = [
            row["tweet_id"], row["status"], row["error_detail"], row["text"],
            row["created_at"], row["api_lang"], row["author_id"], row["author_name"],
            row["username"], row["tagalog_pct"], row["english_pct"], row["other_pct"],
            row["likes"], row["retweets"], row["replies"], row["quotes"], row["views"],
            row["cs_words"], row["pos_tags"], row["switch_type"], row["notes"],
        ]
        is_error = row["status"] == "unavailable"
        row_fill = fill_err if is_error else (fill_a if row_index % 2 == 0 else fill_b)

        for column_index, value in enumerate(values, 1):
            cell = ws1.cell(row=row_index, column=column_index, value=value)
            cell.font = Font(name="Arial", size=9)
            cell.border = thin_border()
            cell.alignment = Alignment(vertical="top", wrap_text=(column_index == 4))
            cell.fill = fill_ann if column_index >= 18 else row_fill
            if column_index in (10, 11, 12):
                cell.number_format = "0.0"

        ws1.row_dimensions[row_index].height = 60 if not is_error else 20

    ws1.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    ws2 = wb.create_sheet("Summary")
    ws2.column_dimensions["A"].width = 32
    ws2.column_dimensions["B"].width = 40

    ok_rows = [row for row in rows if row["status"] == "ok"]
    avg = lambda key: f"{sum(row[key] for row in ok_rows) / len(ok_rows):.1f}%" if ok_rows else "n/a"

    summary_data = [
        ("Metric", "Value"),
        ("Source", "twitterapi.io"),
        ("Total sampled", SAMPLE_SIZE),
        ("Successfully fetched", found),
        ("Unavailable/deleted", missing),
        ("Filter: min Tagalog %", f"{int(MIN_TAGALOG * 100)}%"),
        ("Filter: min English %", f"{int(MIN_ENGLISH * 100)}%"),
        ("", ""),
        ("Avg Tagalog %", avg("tagalog_pct")),
        ("Avg English %", avg("english_pct")),
        ("Avg Other %", avg("other_pct")),
        ("", ""),
        ("-- Annotation Guide --", ""),
        ("Code-Switched Words", "Comma-separated words that switch language mid-tweet"),
        ("POS Tags", "NOUN / VERB / ADJ / ADV per switched word, in same order"),
        ("Switch Type", "inter-sentential / intra-sentential / tag-switching"),
        ("Notes", "Free-form linguistic observations"),
    ]

    header_fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
    for row_index, (label, value) in enumerate(summary_data, 1):
        cell_a = ws2.cell(row=row_index, column=1, value=label)
        cell_b = ws2.cell(row=row_index, column=2, value=value)
        for cell in (cell_a, cell_b):
            cell.font = Font(name="Arial", size=10, bold=(row_index == 1))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = thin_border()
            if row_index == 1:
                cell.fill = header_fill
                cell.font = header_font()
        ws2.row_dimensions[row_index].height = 20

    wb.save(args.output)
    print(f" Saved -> {args.output}")
    print("\nDone. Open the workbook and use the yellow annotation columns to tag code-switched words.")


if __name__ == "__main__":
    main()
