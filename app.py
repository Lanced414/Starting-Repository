"""
Beli × Uber Eats Dashboard
────────────────────────────
Backend Flask app that:
  1. Tries to fetch restaurants from the Uber Eats unofficial API
     (requires a real browser session / cookies to bypass bot protection).
  2. Falls back to realistic sample data for Islip, NY when the API is blocked.
  3. Scrapes beliapp.co to find Beli scores for each restaurant.
"""

import json
import time
import re
import urllib.parse

from flask import Flask, render_template, jsonify, request
import requests
from bs4 import BeautifulSoup

app = Flask(__name__)

DEFAULT_COORDS  = {"lat": 40.7298, "lng": -73.2048}
DEFAULT_ADDRESS = "Islip, NY 11751"

_UBER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "x-csrf-token": "x",
    "Content-Type": "application/json",
    "Referer": "https://www.ubereats.com/",
    "Origin": "https://www.ubereats.com",
}

_BELI_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

# Simple in-memory cache for Beli scores: name.lower() -> float | None
_beli_cache: dict = {}

# ── Demo data (used when Uber Eats API is blocked) ────────────────────────────

_DEMO_RESTAURANTS = [
    {
        "id": "demo-001", "name": "Chipotle Mexican Grill",
        "uber_rating": 4.6, "rating_count": 2341,
        "categories": ["Mexican", "Burritos"],
        "image": "", "url": "https://www.ubereats.com/store/chipotle-mexican-grill-islip",
        "delivery_fee": "$0.49", "delivery_time": "20-30 min", "price_range": "$",
    },
    {
        "id": "demo-002", "name": "Five Guys",
        "uber_rating": 4.7, "rating_count": 1890,
        "categories": ["Burgers", "American"],
        "image": "", "url": "https://www.ubereats.com/store/five-guys-islip",
        "delivery_fee": "$1.99", "delivery_time": "25-35 min", "price_range": "$$",
    },
    {
        "id": "demo-003", "name": "Domino's Pizza",
        "uber_rating": 4.3, "rating_count": 3100,
        "categories": ["Pizza", "Wings"],
        "image": "", "url": "https://www.ubereats.com/store/dominos-pizza-islip",
        "delivery_fee": "$0.99", "delivery_time": "20-30 min", "price_range": "$",
    },
    {
        "id": "demo-004", "name": "Panda Express",
        "uber_rating": 4.2, "rating_count": 1456,
        "categories": ["Chinese", "Asian"],
        "image": "", "url": "https://www.ubereats.com/store/panda-express-islip",
        "delivery_fee": "$1.49", "delivery_time": "25-40 min", "price_range": "$",
    },
    {
        "id": "demo-005", "name": "Wingstop",
        "uber_rating": 4.5, "rating_count": 2200,
        "categories": ["Wings", "Chicken"],
        "image": "", "url": "https://www.ubereats.com/store/wingstop-islip",
        "delivery_fee": "$0.49", "delivery_time": "25-35 min", "price_range": "$",
    },
    {
        "id": "demo-006", "name": "Jersey Mike's Subs",
        "uber_rating": 4.6, "rating_count": 987,
        "categories": ["Sandwiches", "Subs"],
        "image": "", "url": "https://www.ubereats.com/store/jersey-mikes-subs-islip",
        "delivery_fee": "$0.49", "delivery_time": "20-30 min", "price_range": "$",
    },
    {
        "id": "demo-007", "name": "McDonald's",
        "uber_rating": 4.1, "rating_count": 5600,
        "categories": ["Burgers", "Fast Food", "Breakfast"],
        "image": "", "url": "https://www.ubereats.com/store/mcdonalds-islip",
        "delivery_fee": "$0.00", "delivery_time": "15-25 min", "price_range": "$",
    },
    {
        "id": "demo-008", "name": "Subway",
        "uber_rating": 4.0, "rating_count": 2100,
        "categories": ["Sandwiches"],
        "image": "", "url": "https://www.ubereats.com/store/subway-islip",
        "delivery_fee": "$0.99", "delivery_time": "15-25 min", "price_range": "$",
    },
    {
        "id": "demo-009", "name": "Olive Garden Italian Restaurant",
        "uber_rating": 4.4, "rating_count": 1230,
        "categories": ["Italian", "Pasta"],
        "image": "", "url": "https://www.ubereats.com/store/olive-garden-italian-restaurant-islip",
        "delivery_fee": "$1.99", "delivery_time": "35-50 min", "price_range": "$$",
    },
    {
        "id": "demo-010", "name": "Red Lobster",
        "uber_rating": 4.3, "rating_count": 876,
        "categories": ["Seafood", "American"],
        "image": "", "url": "https://www.ubereats.com/store/red-lobster-islip",
        "delivery_fee": "$2.99", "delivery_time": "40-55 min", "price_range": "$$",
    },
    {
        "id": "demo-011", "name": "Taco Bell",
        "uber_rating": 4.0, "rating_count": 3400,
        "categories": ["Mexican", "Fast Food"],
        "image": "", "url": "https://www.ubereats.com/store/taco-bell-islip",
        "delivery_fee": "$0.49", "delivery_time": "15-25 min", "price_range": "$",
    },
    {
        "id": "demo-012", "name": "Burger King",
        "uber_rating": 3.9, "rating_count": 2800,
        "categories": ["Burgers", "Fast Food"],
        "image": "", "url": "https://www.ubereats.com/store/burger-king-islip",
        "delivery_fee": "$0.49", "delivery_time": "15-25 min", "price_range": "$",
    },
    {
        "id": "demo-013", "name": "Papa John's Pizza",
        "uber_rating": 4.2, "rating_count": 1800,
        "categories": ["Pizza"],
        "image": "", "url": "https://www.ubereats.com/store/papa-johns-pizza-islip",
        "delivery_fee": "$0.99", "delivery_time": "25-35 min", "price_range": "$",
    },
    {
        "id": "demo-014", "name": "Firehouse Subs",
        "uber_rating": 4.6, "rating_count": 760,
        "categories": ["Sandwiches", "American"],
        "image": "", "url": "https://www.ubereats.com/store/firehouse-subs-islip",
        "delivery_fee": "$0.49", "delivery_time": "20-30 min", "price_range": "$",
    },
    {
        "id": "demo-015", "name": "Chick-fil-A",
        "uber_rating": 4.8, "rating_count": 4200,
        "categories": ["Chicken", "Sandwiches"],
        "image": "", "url": "https://www.ubereats.com/store/chick-fil-a-islip",
        "delivery_fee": "$0.00", "delivery_time": "20-35 min", "price_range": "$",
    },
    {
        "id": "demo-016", "name": "Shake Shack",
        "uber_rating": 4.7, "rating_count": 3100,
        "categories": ["Burgers", "Shakes"],
        "image": "", "url": "https://www.ubereats.com/store/shake-shack-islip",
        "delivery_fee": "$1.49", "delivery_time": "25-35 min", "price_range": "$$",
    },
    {
        "id": "demo-017", "name": "KFC",
        "uber_rating": 4.0, "rating_count": 1900,
        "categories": ["Chicken", "Fast Food"],
        "image": "", "url": "https://www.ubereats.com/store/kfc-islip",
        "delivery_fee": "$0.49", "delivery_time": "20-30 min", "price_range": "$",
    },
    {
        "id": "demo-018", "name": "Wendy's",
        "uber_rating": 4.1, "rating_count": 2600,
        "categories": ["Burgers", "Fast Food"],
        "image": "", "url": "https://www.ubereats.com/store/wendys-islip",
        "delivery_fee": "$0.49", "delivery_time": "15-25 min", "price_range": "$",
    },
    {
        "id": "demo-019", "name": "La Parma Restaurant",
        "uber_rating": 4.8, "rating_count": 450,
        "categories": ["Italian", "Pasta", "Pizza"],
        "image": "", "url": "https://www.ubereats.com/store/la-parma-restaurant-islip",
        "delivery_fee": "$2.49", "delivery_time": "35-50 min", "price_range": "$$",
    },
    {
        "id": "demo-020", "name": "Moe's Southwest Grill",
        "uber_rating": 4.3, "rating_count": 870,
        "categories": ["Mexican", "Burritos"],
        "image": "", "url": "https://www.ubereats.com/store/moes-southwest-grill-islip",
        "delivery_fee": "$0.99", "delivery_time": "20-30 min", "price_range": "$",
    },
    {
        "id": "demo-021", "name": "Panera Bread",
        "uber_rating": 4.4, "rating_count": 1650,
        "categories": ["Sandwiches", "Salad", "Bakery"],
        "image": "", "url": "https://www.ubereats.com/store/panera-bread-islip",
        "delivery_fee": "$0.99", "delivery_time": "20-35 min", "price_range": "$$",
    },
    {
        "id": "demo-022", "name": "Arby's",
        "uber_rating": 3.8, "rating_count": 1100,
        "categories": ["Sandwiches", "Fast Food"],
        "image": "", "url": "https://www.ubereats.com/store/arbys-islip",
        "delivery_fee": "$0.49", "delivery_time": "15-25 min", "price_range": "$",
    },
    {
        "id": "demo-023", "name": "Starbucks",
        "uber_rating": 4.5, "rating_count": 3800,
        "categories": ["Coffee", "Bakery"],
        "image": "", "url": "https://www.ubereats.com/store/starbucks-islip",
        "delivery_fee": "$2.49", "delivery_time": "15-25 min", "price_range": "$$",
    },
    {
        "id": "demo-024", "name": "Hibachi Express",
        "uber_rating": 4.5, "rating_count": 620,
        "categories": ["Japanese", "Asian"],
        "image": "", "url": "https://www.ubereats.com/store/hibachi-express-islip",
        "delivery_fee": "$1.49", "delivery_time": "25-40 min", "price_range": "$$",
    },
]


# ── Uber Eats ─────────────────────────────────────────────────────────────────


def _uber_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(_UBER_HEADERS)
    try:
        s.get("https://www.ubereats.com", timeout=8)
    except Exception:
        pass
    return s


def fetch_uber_restaurants(lat: float, lng: float, address: str) -> tuple[list[dict], bool]:
    """Returns (restaurants, is_demo).  Falls back to demo data on API block."""
    session = _uber_session()
    payload = {
        "cacheKey": "",
        "feedSessionId": f"feed-{int(time.time())}",
        "userQuery": "",
        "location": {
            "address": address,
            "latitude": lat,
            "longitude": lng,
            "referenceId": "",
            "type": "manual",
        },
        "feed": {
            "storeUrl": None,
            "hasMore": False,
            "offset": 0,
            "pageInfo": {"offset": 0, "pageSize": 80},
        },
        "sortAndFilters": [],
    }

    try:
        resp = session.post(
            "https://www.ubereats.com/_p/api/getFeedV1?localeCode=en-US",
            json=payload,
            timeout=15,
        )
        resp.raise_for_status()
        restaurants = _parse_feed(resp.json())
        if restaurants:
            return restaurants, False
    except Exception as exc:
        print(f"[uber] API failed ({exc}), using demo data")

    return list(_DEMO_RESTAURANTS), True


def _parse_feed(data: dict) -> list[dict]:
    restaurants = []
    d = data.get("data", data)
    items = (
        d.get("feedItems")
        or d.get("feed", {}).get("feedItems")
        or d.get("stores")
        or []
    )

    for item in items:
        store = item.get("store") or item
        title = store.get("title") or store.get("name", "")
        if not title:
            continue

        rating = store.get("rating") or {}
        fare   = store.get("fareInfo") or {}
        eta    = store.get("etaRange") or {}
        cats   = store.get("categories") or []

        restaurants.append({
            "id":            store.get("storeUuid") or store.get("uuid", ""),
            "name":          title,
            "uber_rating":   float(rating.get("ratingValue") or 0),
            "rating_count":  int(rating.get("reviewCount") or 0),
            "categories":    cats if isinstance(cats, list) else [cats],
            "image":         store.get("heroImageUrl") or store.get("imageUrl") or "",
            "url":           "https://www.ubereats.com/store/" + (store.get("storeUrl") or ""),
            "delivery_fee":  fare.get("displayString") or "",
            "delivery_time": eta.get("text") or "",
            "price_range":   store.get("priceRange") or "",
        })

    return restaurants


# ── Beli ──────────────────────────────────────────────────────────────────────


def get_beli_score(restaurant_name: str) -> float | None:
    key = restaurant_name.lower().strip()
    if key in _beli_cache:
        return _beli_cache[key]
    score = _search_beli(restaurant_name)
    _beli_cache[key] = score
    return score


def _search_beli(name: str) -> float | None:
    q = urllib.parse.quote_plus(name)
    candidates = [
        f"https://beliapp.co/search?q={q}",
        f"https://beliapp.co/search?query={q}",
        f"https://app.beliapp.co/search?q={q}",
    ]
    for url in candidates:
        try:
            resp = requests.get(url, headers=_BELI_HEADERS, timeout=10, allow_redirects=True)
            if resp.status_code == 200:
                score = _parse_beli_html(resp.text)
                if score is not None:
                    return score
        except Exception as exc:
            print(f"[beli] {url}: {exc}")
    return None


def _parse_beli_html(html: str) -> float | None:
    soup = BeautifulSoup(html, "html.parser")

    # Next.js hydration data
    next_tag = soup.find("script", id="__NEXT_DATA__")
    if next_tag and next_tag.string:
        try:
            score = _scan_json(json.loads(next_tag.string))
            if score is not None:
                return score
        except Exception:
            pass

    # Embedded JSON blobs
    for tag in soup.find_all("script", type="application/json"):
        if tag.string:
            try:
                score = _scan_json(json.loads(tag.string))
                if score is not None:
                    return score
            except Exception:
                pass

    # Inline JS containing score values
    for tag in soup.find_all("script"):
        src = tag.string or ""
        for m in re.finditer(
            r'"(?:score|beliScore|beli_score)"\s*:\s*([0-9]+(?:\.[0-9]+)?)', src
        ):
            val = float(m.group(1))
            if 0 < val <= 10:
                return val

    return None


_SCORE_KEYS = {"score", "beliScore", "beli_score", "ratingValue", "averageScore"}


def _scan_json(obj, depth: int = 0) -> float | None:
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


# ── Routes ────────────────────────────────────────────────────────────────────


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/restaurants")
def api_restaurants():
    lat     = float(request.args.get("lat",     DEFAULT_COORDS["lat"]))
    lng     = float(request.args.get("lng",     DEFAULT_COORDS["lng"]))
    address = request.args.get("address", DEFAULT_ADDRESS)

    restaurants, is_demo = fetch_uber_restaurants(lat, lng, address)
    return jsonify({
        "status":      "success",
        "is_demo":     is_demo,
        "count":       len(restaurants),
        "restaurants": restaurants,
    })


@app.route("/api/beli")
def api_beli():
    name = request.args.get("name", "").strip()
    if not name:
        return jsonify({"error": "name is required"}), 400

    score = get_beli_score(name)
    return jsonify({"name": name, "beli_score": score, "found": score is not None})


if __name__ == "__main__":
    app.run(debug=True, port=5000)
