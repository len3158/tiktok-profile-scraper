# TikTok Profile Scraper

Scrapes the last N videos from a TikTok profile and exports per-video metrics
(views, likes, comments, shares, saves, engagement rate) to a nicely
formatted Excel report, using your existing logged-in Chrome session so
TikTok treats requests as a real user.

## Requirements

```bash
pip install -r requirements.txt
playwright install chromium
```

You also need Google Chrome installed and logged in to TikTok (the script
reuses your Chrome profile's cookies — no separate login step needed).

## Configuration

Set these via environment variables, or edit the defaults at the top of
`tiktok_scraper.py` directly:

| Variable | Description | Default |
|---|---|---|
| `TIKTOK_HANDLE` | TikTok handle to scrape, including `@` | `@yourusername` |
| `NUM_VIDEOS` | How many of the most recent videos to scrape | `90` |

The script also expects `CHROME_USER_DATA_DIR` (top of the script) to point
at your Chrome profile directory — the macOS default is pre-filled; Windows
and Linux defaults are included as comments.

## Usage

```bash
export TIKTOK_HANDLE="@yourusername"
export NUM_VIDEOS=50

python tiktok_scraper.py
```

The script will:

1. Launch Chrome with your existing profile (visibly, not headless — TikTok
   is more likely to flag headless sessions).
2. Load the profile and scroll to collect video links.
3. Visit each video page, extract stats from TikTok's embedded JSON, and
   wait a randomized cooldown (3.5–6s) between requests.
4. Export everything to `tiktok_metrics.xlsx`, including totals/averages and
   a color-scaled "Views" column.

## Notes

- Keep the cooldown in place (or increase it) to avoid rate limiting/flags
  from TikTok.
- Relies on TikTok's current page structure (`__UNIVERSAL_DATA_FOR_REHYDRATION__`
  embedded JSON and `data-e2e` attributes) — may need updating if TikTok
  changes its frontend.
- Respect TikTok's Terms of Service when scraping.

## License

MIT — see [LICENSE](LICENSE).
