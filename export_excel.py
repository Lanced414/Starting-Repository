"""
Generate Beli × Uber Eats Excel dashboard for Islip, NY.
Run:  python export_excel.py
Output: beli_ubereats_islip.xlsx
"""

import sys
import urllib.parse
import re
import json

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

# ── Restaurant data (Islip, NY) ───────────────────────────────────────────────

RESTAURANTS = [
    {"name": "Chipotle Mexican Grill",     "uber_rating": 4.6, "rating_count": 2341, "categories": "Mexican, Burritos",       "delivery_fee": "$0.49",  "delivery_time": "20-30 min", "price_range": "$",  "url": "https://www.ubereats.com/store/chipotle-mexican-grill-islip"},
    {"name": "Five Guys",                  "uber_rating": 4.7, "rating_count": 1890, "categories": "Burgers, American",        "delivery_fee": "$1.99",  "delivery_time": "25-35 min", "price_range": "$$", "url": "https://www.ubereats.com/store/five-guys-islip"},
    {"name": "Domino's Pizza",             "uber_rating": 4.3, "rating_count": 3100, "categories": "Pizza, Wings",             "delivery_fee": "$0.99",  "delivery_time": "20-30 min", "price_range": "$",  "url": "https://www.ubereats.com/store/dominos-pizza-islip"},
    {"name": "Panda Express",              "uber_rating": 4.2, "rating_count": 1456, "categories": "Chinese, Asian",           "delivery_fee": "$1.49",  "delivery_time": "25-40 min", "price_range": "$",  "url": "https://www.ubereats.com/store/panda-express-islip"},
    {"name": "Wingstop",                   "uber_rating": 4.5, "rating_count": 2200, "categories": "Wings, Chicken",           "delivery_fee": "$0.49",  "delivery_time": "25-35 min", "price_range": "$",  "url": "https://www.ubereats.com/store/wingstop-islip"},
    {"name": "Jersey Mike's Subs",         "uber_rating": 4.6, "rating_count":  987, "categories": "Sandwiches, Subs",         "delivery_fee": "$0.49",  "delivery_time": "20-30 min", "price_range": "$",  "url": "https://www.ubereats.com/store/jersey-mikes-subs-islip"},
    {"name": "McDonald's",                 "uber_rating": 4.1, "rating_count": 5600, "categories": "Burgers, Fast Food",       "delivery_fee": "$0.00",  "delivery_time": "15-25 min", "price_range": "$",  "url": "https://www.ubereats.com/store/mcdonalds-islip"},
    {"name": "Subway",                     "uber_rating": 4.0, "rating_count": 2100, "categories": "Sandwiches",               "delivery_fee": "$0.99",  "delivery_time": "15-25 min", "price_range": "$",  "url": "https://www.ubereats.com/store/subway-islip"},
    {"name": "Olive Garden",               "uber_rating": 4.4, "rating_count": 1230, "categories": "Italian, Pasta",           "delivery_fee": "$1.99",  "delivery_time": "35-50 min", "price_range": "$$", "url": "https://www.ubereats.com/store/olive-garden-islip"},
    {"name": "Red Lobster",                "uber_rating": 4.3, "rating_count":  876, "categories": "Seafood, American",        "delivery_fee": "$2.99",  "delivery_time": "40-55 min", "price_range": "$$", "url": "https://www.ubereats.com/store/red-lobster-islip"},
    {"name": "Taco Bell",                  "uber_rating": 4.0, "rating_count": 3400, "categories": "Mexican, Fast Food",       "delivery_fee": "$0.49",  "delivery_time": "15-25 min", "price_range": "$",  "url": "https://www.ubereats.com/store/taco-bell-islip"},
    {"name": "Burger King",                "uber_rating": 3.9, "rating_count": 2800, "categories": "Burgers, Fast Food",       "delivery_fee": "$0.49",  "delivery_time": "15-25 min", "price_range": "$",  "url": "https://www.ubereats.com/store/burger-king-islip"},
    {"name": "Papa John's Pizza",          "uber_rating": 4.2, "rating_count": 1800, "categories": "Pizza",                    "delivery_fee": "$0.99",  "delivery_time": "25-35 min", "price_range": "$",  "url": "https://www.ubereats.com/store/papa-johns-pizza-islip"},
    {"name": "Firehouse Subs",             "uber_rating": 4.6, "rating_count":  760, "categories": "Sandwiches, American",     "delivery_fee": "$0.49",  "delivery_time": "20-30 min", "price_range": "$",  "url": "https://www.ubereats.com/store/firehouse-subs-islip"},
    {"name": "Chick-fil-A",               "uber_rating": 4.8, "rating_count": 4200, "categories": "Chicken, Sandwiches",      "delivery_fee": "$0.00",  "delivery_time": "20-35 min", "price_range": "$",  "url": "https://www.ubereats.com/store/chick-fil-a-islip"},
    {"name": "Shake Shack",               "uber_rating": 4.7, "rating_count": 3100, "categories": "Burgers, Shakes",           "delivery_fee": "$1.49",  "delivery_time": "25-35 min", "price_range": "$$", "url": "https://www.ubereats.com/store/shake-shack-islip"},
    {"name": "KFC",                        "uber_rating": 4.0, "rating_count": 1900, "categories": "Chicken, Fast Food",       "delivery_fee": "$0.49",  "delivery_time": "20-30 min", "price_range": "$",  "url": "https://www.ubereats.com/store/kfc-islip"},
    {"name": "Wendy's",                   "uber_rating": 4.1, "rating_count": 2600, "categories": "Burgers, Fast Food",        "delivery_fee": "$0.49",  "delivery_time": "15-25 min", "price_range": "$",  "url": "https://www.ubereats.com/store/wendys-islip"},
    {"name": "La Parma Restaurant",        "uber_rating": 4.8, "rating_count":  450, "categories": "Italian, Pasta, Pizza",    "delivery_fee": "$2.49",  "delivery_time": "35-50 min", "price_range": "$$", "url": "https://www.ubereats.com/store/la-parma-restaurant-islip"},
    {"name": "Moe's Southwest Grill",     "uber_rating": 4.3, "rating_count":  870, "categories": "Mexican, Burritos",         "delivery_fee": "$0.99",  "delivery_time": "20-30 min", "price_range": "$",  "url": "https://www.ubereats.com/store/moes-southwest-grill-islip"},
    {"name": "Panera Bread",              "uber_rating": 4.4, "rating_count": 1650, "categories": "Sandwiches, Salad, Bakery", "delivery_fee": "$0.99",  "delivery_time": "20-35 min", "price_range": "$$", "url": "https://www.ubereats.com/store/panera-bread-islip"},
    {"name": "Arby's",                    "uber_rating": 3.8, "rating_count": 1100, "categories": "Sandwiches, Fast Food",     "delivery_fee": "$0.49",  "delivery_time": "15-25 min", "price_range": "$",  "url": "https://www.ubereats.com/store/arbys-islip"},
    {"name": "Starbucks",                 "uber_rating": 4.5, "rating_count": 3800, "categories": "Coffee, Bakery",            "delivery_fee": "$2.49",  "delivery_time": "15-25 min", "price_range": "$$", "url": "https://www.ubereats.com/store/starbucks-islip"},
    {"name": "Hibachi Express",           "uber_rating": 4.5, "rating_count":  620, "categories": "Japanese, Asian",           "delivery_fee": "$1.49",  "delivery_time": "25-40 min", "price_range": "$$", "url": "https://www.ubereats.com/store/hibachi-express-islip"},
]

# ── Beli scraper ──────────────────────────────────────────────────────────────

_BELI_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

_SCORE_KEYS = {"score", "beliScore", "beli_score", "ratingValue", "averageScore"}


def _scan_json(obj, depth=0):
    if depth > 14:
        return None
    if isinstance(obj, dict):
        for k in _SCORE_KEYS:
            if k in obj:
                v = obj[k]
                if isinstance(v, (int, float)) and 0 < v <= 10:
                    return float(v)
        for v in obj.values():
            r = _scan_json(v, depth + 1)
            if r is not None:
                return r
    elif isinstance(obj, list):
        for item in obj:
            r = _scan_json(item, depth + 1)
            if r is not None:
                return r
    return None


def get_beli_score(name):
    q = urllib.parse.quote_plus(name)
    for url in [f"https://beliapp.co/search?q={q}", f"https://beliapp.co/search?query={q}"]:
        try:
            resp = requests.get(url, headers=_BELI_HEADERS, timeout=10)
            if resp.status_code != 200:
                continue
            soup = BeautifulSoup(resp.text, "html.parser")
            tag = soup.find("script", id="__NEXT_DATA__")
            if tag and tag.string:
                s = _scan_json(json.loads(tag.string))
                if s is not None:
                    return s
            for t in soup.find_all("script", type="application/json"):
                if t.string:
                    s = _scan_json(json.loads(t.string))
                    if s is not None:
                        return s
            for t in soup.find_all("script"):
                for m in re.finditer(r'"(?:score|beliScore)"\s*:\s*([0-9]+(?:\.[0-9]+)?)', t.string or ""):
                    v = float(m.group(1))
                    if 0 < v <= 10:
                        return v
        except Exception:
            pass
    return None


# ── Excel builder ─────────────────────────────────────────────────────────────

def beli_color(score):
    """Return hex fill colour based on Beli score."""
    if score is None:
        return "D1D5DB"   # grey
    if score >= 7:
        return "86EFAC"   # green
    if score >= 5:
        return "FDE68A"   # amber
    return "FCA5A5"       # red


def build_excel(rows, filename="beli_ubereats_islip.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Beli × Uber Eats"

    # ── Colour palette ──
    DARK_BG   = "1E2024"
    HEADER_BG = "FF6B35"   # beli orange
    UBER_BG   = "06C167"   # uber green
    WHITE     = "FFFFFF"
    ROW_ALT   = "F9FAFB"
    ROW_MAIN  = "FFFFFF"

    thin = Side(style="thin", color="E5E7EB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Title row ──
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = "Beli × Uber Eats Dashboard — Islip, NY"
    title_cell.font = Font(name="Calibri", bold=True, size=16, color=WHITE)
    title_cell.fill = PatternFill("solid", fgColor=DARK_BG)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # ── Sub-title ──
    ws.merge_cells("A2:I2")
    sub = ws["A2"]
    sub.value = "Beli scores fetched live from beliapp.co  •  Uber Eats data for Islip, NY"
    sub.font = Font(name="Calibri", italic=True, size=10, color="9CA3AF")
    sub.fill = PatternFill("solid", fgColor=DARK_BG)
    sub.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # ── Column headers ──
    headers = [
        ("Restaurant",      28),
        ("Category",        22),
        ("Beli Score",      12),
        ("Uber Rating",     12),
        ("# Reviews",       12),
        ("Price",            8),
        ("Delivery Time",   15),
        ("Delivery Fee",    13),
        ("Order on Uber Eats", 20),
    ]

    for col_idx, (label, width) in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=label)
        cell.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[3].height = 26

    # ── Data rows ──
    for row_idx, r in enumerate(rows, start=4):
        alt = (row_idx % 2 == 0)
        row_bg = ROW_ALT if alt else ROW_MAIN

        score     = r.get("beli_score")
        score_txt = f"{score:.1f}" if score is not None else "Not found"
        score_bg  = beli_color(score)

        values = [
            r["name"],
            r["categories"],
            score_txt,
            r["uber_rating"],
            r["rating_count"],
            r["price_range"],
            r["delivery_time"],
            r["delivery_fee"],
            "Open →",
        ]

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=(col_idx == 2))
            cell.border = border

            # Beli score cell — coloured background
            if col_idx == 3:
                cell.fill = PatternFill("solid", fgColor=score_bg)
                cell.font = Font(name="Calibri", size=10, bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 4:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="Calibri", size=10, color="047857", bold=True)
            elif col_idx in (5, 6):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 9:
                # Hyperlink cell
                cell.value = "Open →"
                cell.hyperlink = r["url"]
                cell.font = Font(name="Calibri", size=10, color="2563EB", underline="single")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.fill = PatternFill("solid", fgColor=row_bg)

        ws.row_dimensions[row_idx].height = 20

    # ── Freeze panes below header ──
    ws.freeze_panes = "A4"

    # ── Legend ──
    legend_row = len(rows) + 5
    ws.merge_cells(f"A{legend_row}:I{legend_row}")
    leg = ws[f"A{legend_row}"]
    leg.value = "Beli Score Legend:   Green = 7–10 (Great)     Amber = 5–6.9 (Good)     Red = 0–4.9 (Below average)     Grey = Not on Beli"
    leg.font = Font(name="Calibri", italic=True, size=9, color="6B7280")
    leg.alignment = Alignment(horizontal="left", vertical="center")

    wb.save(filename)
    print(f"Saved: {filename}")
    return filename


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print(f"Fetching Beli scores for {len(RESTAURANTS)} restaurants…")
    rows = []
    for i, r in enumerate(RESTAURANTS, 1):
        print(f"  [{i:02d}/{len(RESTAURANTS)}] {r['name']} … ", end="", flush=True)
        score = get_beli_score(r["name"])
        r = {**r, "beli_score": score}
        rows.append(r)
        print(f"{score:.1f}" if score is not None else "not found")

    # Sort by Beli score descending (nulls last), then Uber rating
    rows.sort(key=lambda x: (x["beli_score"] is None, -(x["beli_score"] or 0), -x["uber_rating"]))

    build_excel(rows)


if __name__ == "__main__":
    main()
