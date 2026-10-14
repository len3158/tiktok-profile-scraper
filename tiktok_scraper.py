"""
TikTok Profile Scraper
=======================================
Scrapes the last N videos from a TikTok profile and exports metrics to Excel.

Metrics captured per video:
  - Rank, Video ID, Post Date, Title/Description
  - Link, Views, Likes, Comments, Shares, Saves
  - Engagement Rate (likes+comments+saves / views)

Usage:
  python tiktok_scraper.py

Requirements:
  pip install playwright openpyxl
  playwright install chromium

HOW IT WORKS:
  Uses your existing logged-in Chrome profile so TikTok sees you as a real user.
  Navigates to each video page, extracts stats from TikTok's embedded SSR JSON,
  and writes everything to a formatted Excel file.

CONFIGURATION (edit below):
"""

import json
import re
import time
import random
import os
import sys
from datetime import datetime
from pathlib import Path

# ──────────────────────────────────────────────
#  CONFIGURATION — edit these before running
# ──────────────────────────────────────────────

TIKTOK_HANDLE = os.environ.get("TIKTOK_HANDLE", "@yourusername")
PROFILE_URL   = f"https://www.tiktok.com/{TIKTOK_HANDLE}"
NUM_VIDEOS    = int(os.environ.get("NUM_VIDEOS", 90))  # How many of your most recent videos to scrape

# Cooldown between each video page visit (seconds). Min 3 recommended.
COOLDOWN_MIN  = 3.5
COOLDOWN_MAX  = 6.0

# Output Excel file path (relative to this script)
OUTPUT_FILE   = "tiktok_metrics.xlsx"

# ── Chrome user data directory (where your logged-in session lives) ──
# This lets Playwright use your real TikTok cookies — no login needed.
#
# macOS default:
CHROME_USER_DATA_DIR = os.path.expanduser("~/Library/Application Support/Google/Chrome")
# Windows default (uncomment if on Windows):
# CHROME_USER_DATA_DIR = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")
# Linux default (uncomment if on Linux):
# CHROME_USER_DATA_DIR = os.path.expanduser("~/.config/google-chrome")

# Which Chrome profile to use ("Default" is the main profile)
CHROME_PROFILE = "Default"

# ──────────────────────────────────────────────
#  SCRAPER
# ──────────────────────────────────────────────

def parse_count(val):
    """Convert TikTok count values (int, str like '2909', '18000') to int."""
    if val is None:
        return 0
    if isinstance(val, int):
        return val
    if isinstance(val, float):
        return int(val)
    s = str(val).strip().replace(",", "")
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0

def format_count(n):
    """Format large numbers with K/M suffixes for display."""
    n = parse_count(n)
    if n >= 1_000_000:
        return f"{n/1_000_000:.1f}M"
    if n >= 1_000:
        return f"{n/1_000:.1f}K"
    return str(n)

def extract_video_stats(page_content: str, fallback_link: str = "") -> dict:
    """Extract stats from TikTok's embedded __UNIVERSAL_DATA_FOR_REHYDRATION__ JSON."""
    match = re.search(
        r'<script[^>]+id="__UNIVERSAL_DATA_FOR_REHYDRATION__"[^>]*>(.*?)</script>',
        page_content, re.DOTALL
    )
    if not match:
        return {}

    try:
        data = json.loads(match.group(1))
        scope = data.get("__DEFAULT_SCOPE__", {})
        detail = scope.get("webapp.video-detail", {})
        item = detail.get("itemInfo", {}).get("itemStruct", {})
        stats = item.get("stats", {})

        create_ts = item.get("createTime", 0)
        post_date = ""
        if create_ts:
            post_date = datetime.fromtimestamp(int(create_ts)).strftime("%Y-%m-%d")

        desc = item.get("desc", "")
        video_id = item.get("id", "")

        return {
            "videoId":   video_id,
            "title":     desc[:200] if desc else "",
            "postDate":  post_date,
            "views":     parse_count(stats.get("playCount", 0)),
            "likes":     parse_count(stats.get("diggCount", 0)),
            "comments":  parse_count(stats.get("commentCount", 0)),
            "shares":    parse_count(stats.get("shareCount", 0)),
            "saves":     parse_count(stats.get("collectCount", 0)),
        }
    except (json.JSONDecodeError, KeyError, TypeError) as e:
        print(f"  ⚠ Parse error: {e}")
        return {}


def scroll_profile_to_load_videos(page, target_count: int):
    """Scroll the profile page until at least target_count video items appear in DOM."""
    print(f"  Scrolling to load {target_count} videos...")
    prev_count = 0
    stall_count = 0

    for _ in range(60):  # max 60 scroll attempts
        count = page.evaluate(
            "document.querySelectorAll('[data-e2e=\"user-post-item\"]').length"
        )
        print(f"    {count} videos loaded...", end="\r")

        if count >= target_count:
            print(f"\n  ✓ {count} videos loaded")
            return count

        if count == prev_count:
            stall_count += 1
            if stall_count >= 5:
                print(f"\n  ✓ Reached end of feed at {count} videos")
                return count
        else:
            stall_count = 0

        prev_count = count
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(random.uniform(2.5, 3.5))

    return page.evaluate(
        "document.querySelectorAll('[data-e2e=\"user-post-item\"]').length"
    )


def get_video_links_from_profile(page, num_videos: int) -> list[dict]:
    """Extract video links + view counts from the profile grid."""
    items_js = """
    () => {
        const items = document.querySelectorAll('[data-e2e="user-post-item"]');
        const results = [];
        items.forEach((item, i) => {
            const link = item.querySelector('a[href*="/video/"]');
            const href = link ? link.href : '';
            const videoId = href.match(/\\/video\\/(\\d+)/)?.[1] || '';
            const viewEl = item.querySelector('[data-e2e="video-views"]');
            const views = viewEl ? viewEl.textContent.trim() : '';
            if (href) results.push({ rank: i + 1, videoId, link: href, viewsRaw: views });
        });
        return results;
    }
    """
    return page.evaluate(items_js)[:num_videos]


def scrape_tiktok_profile():
    from playwright.sync_api import sync_playwright

    print(f"\n{'='*55}")
    print(f"  TikTok Scraper — {TIKTOK_HANDLE}")
    print(f"  Target: {NUM_VIDEOS} most recent videos")
    print(f"{'='*55}\n")

    # Validate Chrome profile path
    profile_path = Path(CHROME_USER_DATA_DIR)
    if not profile_path.exists():
        print(f"❌ Chrome user data directory not found:\n   {CHROME_USER_DATA_DIR}")
        print("\n  Fix: update CHROME_USER_DATA_DIR at the top of this script.")
        sys.exit(1)

    all_results = []

    with sync_playwright() as p:
        print("▶ Launching Chrome with your existing profile...")
        print(f"  Profile dir : {CHROME_USER_DATA_DIR}")
        print(f"  Profile name: {CHROME_PROFILE}\n")

        # Launch using your real Chrome with existing cookies/session
        browser = p.chromium.launch_persistent_context(
            user_data_dir=CHROME_USER_DATA_DIR,
            channel="chrome",           # use system Chrome, not Chromium
            headless=False,             # keep visible so TikTok doesn't flag headless
            args=[
                f"--profile-directory={CHROME_PROFILE}",
                "--disable-blink-features=AutomationControlled",
                "--no-first-run",
                "--no-default-browser-check",
            ],
            ignore_default_args=["--enable-automation"],
        )

        page = browser.new_page()

        # ── Step 1: Load profile and scroll to get all video links ──
        print(f"▶ Loading profile: {PROFILE_URL}")
        page.goto(PROFILE_URL, wait_until="domcontentloaded", timeout=30_000)
        time.sleep(3)

        loaded = scroll_profile_to_load_videos(page, NUM_VIDEOS)
        video_list = get_video_links_from_profile(page, NUM_VIDEOS)

        print(f"\n✓ Found {len(video_list)} video links to scrape\n")
        if not video_list:
            print("❌ No videos found. Make sure you're logged in to TikTok in Chrome.")
            browser.close()
            sys.exit(1)

        # ── Step 2: Visit each video page and extract full stats ──
        print(f"▶ Visiting each video page ({COOLDOWN_MIN}–{COOLDOWN_MAX}s cooldown)...\n")

        for i, video in enumerate(video_list):
            rank = i + 1
            link = video["link"]
            video_id = video.get("videoId", "")
            prefix = f"  [{rank:02d}/{len(video_list)}]"

            print(f"{prefix} {link}")

            try:
                page.goto(link, wait_until="domcontentloaded", timeout=25_000)
                time.sleep(1.5)  # let SSR JSON settle

                content = page.content()
                stats = extract_video_stats(content, link)

                if stats:
                    result = {
                        "rank":       rank,
                        "videoId":    stats.get("videoId") or video_id,
                        "postDate":   stats.get("postDate", ""),
                        "title":      stats.get("title", ""),
                        "link":       link,
                        "views":      stats["views"],
                        "likes":      stats["likes"],
                        "comments":   stats["comments"],
                        "shares":     stats["shares"],
                        "saves":      stats["saves"],
                    }
                    views = result["views"]
                    eng = round((result["likes"] + result["comments"] + result["saves"]) / views * 100, 2) if views > 0 else 0
                    result["engagementRate"] = eng

                    all_results.append(result)
                    print(f"         👁 {format_count(result['views'])}  ❤ {format_count(result['likes'])}  💬 {format_count(result['comments'])}  📤 {format_count(result['shares'])}  🔖 {format_count(result['saves'])}  📈 {eng}%")
                else:
                    # Fallback: use view count from profile DOM if we have it
                    print(f"         ⚠ Could not parse stats — saving link only")
                    all_results.append({
                        "rank": rank, "videoId": video_id, "postDate": "",
                        "title": "", "link": link, "views": 0,
                        "likes": 0, "comments": 0, "shares": 0,
                        "saves": 0, "engagementRate": 0,
                    })

            except Exception as e:
                print(f"         ❌ Error: {e}")
                all_results.append({
                    "rank": rank, "videoId": video_id, "postDate": "",
                    "title": "", "link": link, "views": 0,
                    "likes": 0, "comments": 0, "shares": 0,
                    "saves": 0, "engagementRate": 0,
                })

            # Cooldown (skip after last video)
            if i < len(video_list) - 1:
                delay = random.uniform(COOLDOWN_MIN, COOLDOWN_MAX)
                time.sleep(delay)

        browser.close()

    return all_results


# ──────────────────────────────────────────────
#  EXCEL EXPORT
# ──────────────────────────────────────────────

def export_to_excel(results: list[dict], output_path: str):
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, GradientFill
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule

    wb = Workbook()
    ws = wb.active
    ws.title = "TikTok Metrics"

    # ── Colours ──
    DARK_BG    = "1A1A2E"  # deep navy (TikTok-ish dark)
    ACCENT     = "FF0050"  # TikTok red/pink
    ACCENT2    = "00F2EA"  # TikTok teal
    HEADER_FG  = "FFFFFF"
    ROW_A      = "F7F7FB"
    ROW_B      = "FFFFFF"
    BORDER_CLR = "DEDEDE"

    thin = Side(style="thin", color=BORDER_CLR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Column definitions ──
    columns = [
        ("Rank",             8),
        ("Video ID",        20),
        ("Post Date",       13),
        ("Title / Description", 55),
        ("Link",            14),
        ("Views",           13),
        ("Likes",           12),
        ("Comments",        12),
        ("Shares",          12),
        ("Saves",           12),
        ("Engagement %",    14),
    ]

    # ── Title row ──
    ws.merge_cells("A1:K1")
    title_cell = ws["A1"]
    title_cell.value = f"TikTok Performance Report — {TIKTOK_HANDLE}  |  Top {len(results)} Videos  |  Generated {datetime.now().strftime('%B %d, %Y')}"
    title_cell.font = Font(name="Arial", size=13, bold=True, color=HEADER_FG)
    title_cell.fill = PatternFill("solid", fgColor=DARK_BG)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── Header row ──
    for col_idx, (header, width) in enumerate(columns, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = Font(name="Arial", size=10, bold=True, color=HEADER_FG)
        cell.fill = PatternFill("solid", fgColor=ACCENT)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[2].height = 22

    # ── Data rows ──
    for row_idx, r in enumerate(results, start=3):
        fill_color = ROW_A if row_idx % 2 == 0 else ROW_B
        row_fill = PatternFill("solid", fgColor=fill_color)

        values = [
            r["rank"],
            r["videoId"],
            r["postDate"],
            r["title"],
            r["link"],
            r["views"],
            r["likes"],
            r["comments"],
            r["shares"],
            r["saves"],
            r["engagementRate"],
        ]

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Arial", size=9)
            cell.fill = row_fill
            cell.border = border

            # Alignment
            if col_idx == 4:  # Title
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            elif col_idx == 5:  # Link — make it a hyperlink
                cell.value = "Open ↗"
                cell.hyperlink = val
                cell.font = Font(name="Arial", size=9, color="0563C1", underline="single")
                cell.alignment = Alignment(horizontal="center")
            elif col_idx in (6, 7, 8, 9, 10):  # Numeric stats
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "#,##0"
            elif col_idx == 11:  # Engagement %
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0.00%"
                cell.value = r["engagementRate"] / 100  # store as decimal for % format
            elif col_idx == 1:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")

        ws.row_dimensions[row_idx].height = 30

    # ── Summary stats row ──
    last_data_row = 2 + len(results)
    summary_row = last_data_row + 2
    ws.merge_cells(f"A{summary_row}:C{summary_row}")
    s = ws.cell(row=summary_row, column=1, value="TOTALS / AVERAGES")
    s.font = Font(name="Arial", size=10, bold=True, color=HEADER_FG)
    s.fill = PatternFill("solid", fgColor=DARK_BG)
    s.alignment = Alignment(horizontal="center", vertical="center")

    summary_labels = ["", "", "", "Total →", "",
                      "Views", "Likes", "Comments", "Shares", "Saves", "Avg Eng%"]
    summary_formulas = [
        None, None, None, None, None,
        f"=SUM(F3:F{last_data_row})",
        f"=SUM(G3:G{last_data_row})",
        f"=SUM(H3:H{last_data_row})",
        f"=SUM(I3:I{last_data_row})",
        f"=SUM(J3:J{last_data_row})",
        f"=AVERAGE(K3:K{last_data_row})",
    ]

    for col_idx, (label, formula) in enumerate(zip(summary_labels, summary_formulas), start=1):
        if formula:
            cell = ws.cell(row=summary_row, column=col_idx, value=formula)
            cell.font = Font(name="Arial", size=10, bold=True)
            cell.fill = PatternFill("solid", fgColor="E8F4FD")
            cell.border = border
            cell.alignment = Alignment(horizontal="right")
            if col_idx == 11:
                cell.number_format = "0.00%"
            else:
                cell.number_format = "#,##0"

    # ── Conditional formatting: color scale on Views column ──
    data_range_views = f"F3:F{last_data_row}"
    ws.conditional_formatting.add(data_range_views, ColorScaleRule(
        start_type="min", start_color="FFCCCC",
        mid_type="percentile", mid_value=50, mid_color="FFE680",
        end_type="max", end_color="00CC66"
    ))

    # ── Freeze panes below header ──
    ws.freeze_panes = "A3"

    # ── Auto-filter on header row ──
    ws.auto_filter.ref = f"A2:K{last_data_row}"

    # ── Sheet tab colour ──
    ws.sheet_properties.tabColor = ACCENT

    wb.save(output_path)
    print(f"\n✅ Excel saved → {output_path}")


# ──────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────

if __name__ == "__main__":
    results = scrape_tiktok_profile()

    if results:
        output_path = OUTPUT_FILE
        print(f"\n▶ Building Excel report ({len(results)} videos)...")
        export_to_excel(results, output_path)
        print(f"\n{'='*55}")
        print(f"  Done! Open: {os.path.abspath(output_path)}")
        print(f"{'='*55}\n")
    else:
        print("\n❌ No data collected. Check your Chrome profile path and TikTok login.")
