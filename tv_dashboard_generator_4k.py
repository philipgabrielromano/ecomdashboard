#!/usr/bin/env python3
"""Generate a responsive browser-rendered dashboard from Upright Labs / Lister sales data."""

from __future__ import annotations

import argparse
import base64
import html
import json
import os
import re
import statistics
import subprocess
import sys
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

try:
    import requests
except ImportError as exc:  # pragma: no cover
    raise SystemExit("This script requires the 'requests' package. Install it with: pip install requests") from exc


BASE_URL = "https://app.uprightlabs.com/api"
TIMEOUT = 60
DEFAULT_OUTPUT = "docs/index.html"
DEFAULT_JSON_OUTPUT = "docs/latest.json"
DEFAULT_PNG_OUTPUT = "docs/dashboard.png"
SORT_CHOICES = ("price", "subtotal", "paid_at", "title")
DEFAULT_LOGO_PATH = "logo.png"

GENERIC_BRAND_TERMS = {
    "nwt", "new", "with", "tags", "tag", "vintage", "lot", "set", "pair", "size",
    "mens", "womens", "women", "men", "kids", "the", "and", "authentic", "rare",
    "style", "large", "small", "bulk", "mixed", "used", "tested", "working", "nwot"
}


class ApiError(RuntimeError):
    pass


def log(message: str, enabled: bool = True) -> None:
    if enabled:
        print(message, file=sys.stderr)


def load_dotenv(path: str = ".env") -> None:
    env_path = Path(path)
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#") or "=" not in raw:
            continue
        key, value = raw.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def isoformat_z(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")


def default_date_window() -> tuple[str, str]:
    end = datetime.now(timezone.utc)
    start = end - timedelta(days=3)
    return isoformat_z(start), isoformat_z(end)


def default_title() -> str:
    return "eCommerce Dashboard"


def default_png_path(html_path: str) -> str:
    return str(Path(html_path).with_name("dashboard.png"))


def default_json_path(html_path: str) -> str:
    return str(Path(html_path).with_name("latest.json"))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a responsive browser-rendered dashboard from Upright Labs sold-item data."
    )
    parser.add_argument("--token", help="API token. If omitted, uses UPRIGHTLABS_API_TOKEN from .env or the environment.")
    parser.add_argument("--start", help="ISO8601 start datetime. Defaults to .env value or last three days.")
    parser.add_argument("--end", help="ISO8601 end datetime. Defaults to .env value or current run time.")
    parser.add_argument("--top", type=int, help="Number of featured items to include. Defaults to .env value or 3.")
    parser.add_argument("--sort", choices=SORT_CHOICES, help="Sort mode for featured items. Defaults to .env value or price.")
    parser.add_argument("--channel", help="Optional exact channel filter, e.g. Shopgoodwill")
    parser.add_argument("--category", help="Optional substring category filter.")
    parser.add_argument("--supplier", help="Optional substring supplier filter.")
    parser.add_argument("--tag", action="append", default=[], help="Optional tag filter. Repeatable.")
    parser.add_argument("--max-products", type=int, default=None, help="Optional safety cap on product enrichment calls.")
    parser.add_argument("--output", help=f"Output HTML path. Defaults to {DEFAULT_OUTPUT}")
    parser.add_argument("--png-output", help=f"Optional PNG output path. Defaults to {DEFAULT_PNG_OUTPUT}")
    parser.add_argument("--json-output", help=f"Optional JSON output path. Defaults to {DEFAULT_JSON_OUTPUT}")
    parser.add_argument("--title", help="Dashboard title.")
    parser.add_argument("--organization-name", help="Optional organization/display name shown in the header.")
    parser.add_argument("--image-mode", choices=("remote", "none"), help="Use API image URLs directly, or disable images.")
    parser.add_argument("--logo-path", default=DEFAULT_LOGO_PATH, help="Path to logo image for embedding.")
    parser.add_argument("--refresh-seconds", type=int, default=300, help="Auto-refresh interval in seconds.")
    parser.add_argument("--no-png", action="store_true", help="Skip PNG generation and only write the HTML/JSON outputs.")
    parser.add_argument("--verbose", action="store_true", help="Print progress information.")
    return parser.parse_args()


def require_token(cli_token: Optional[str]) -> str:
    token = cli_token or os.getenv("UPRIGHTLABS_API_TOKEN")
    if not token:
        raise SystemExit("Missing API token. Use --token or set UPRIGHTLABS_API_TOKEN.")
    return token


def validate_iso8601(value: str) -> str:
    try:
        if value.endswith("Z"):
            datetime.fromisoformat(value.replace("Z", "+00:00"))
        else:
            datetime.fromisoformat(value)
    except ValueError as exc:
        raise SystemExit(f"Invalid ISO8601 datetime: {value}") from exc
    return value


def api_get(session: requests.Session, path: str, params: Optional[Dict[str, Any]] = None) -> Any:
    url = f"{BASE_URL}{path}"
    response = session.get(url, params=params, timeout=TIMEOUT)
    if not response.ok:
        snippet = response.text[:500]
        raise ApiError(f"API request failed [{response.status_code}] {url}\n{snippet}")
    ctype = response.headers.get("content-type", "")
    if "application/json" not in ctype and not response.text.lstrip().startswith(("{", "[")):
        raise ApiError(f"Expected JSON from {url}, got content-type {ctype!r}")
    return response.json()


def build_session(token: str) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "X-Authorization": token,
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "uprightlabs-responsive-dashboard/1.0",
    })
    return session


def safe_float(value: Any) -> float:
    if value in (None, "", "null"):
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def parse_report_datetime(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    for fmt in ("%m/%d/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%SZ"):
        try:
            dt = datetime.strptime(value, fmt)
            return dt.replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def fmt_currency(value: Any) -> str:
    return f"${safe_float(value):,.2f}"


def fmt_short_date(value: Optional[datetime]) -> str:
    if not value:
        return "N/A"
    return value.astimezone(timezone.utc).strftime("%m/%d")


def fmt_generated_at(value: datetime) -> str:
    return value.astimezone(timezone.utc).strftime("%m/%d/%Y %I:%M %p UTC")


def fmt_report_span(start_iso: str, end_iso: str) -> str:
    start = parse_report_datetime(start_iso) or datetime.fromisoformat(start_iso.replace("Z", "+00:00"))
    end = parse_report_datetime(end_iso) or datetime.fromisoformat(end_iso.replace("Z", "+00:00"))
    return f"{start.astimezone(timezone.utc).strftime('%m/%d/%Y')} - {end.astimezone(timezone.utc).strftime('%m/%d/%Y')}"


def logo_data_uri(path: str = DEFAULT_LOGO_PATH) -> str:
    logo_path = Path(path)
    if not logo_path.exists():
        return ""
    raw = logo_path.read_bytes()
    encoded = base64.b64encode(raw).decode("ascii")
    suffix = logo_path.suffix.lower()
    mime = "image/png"
    if suffix in (".jpg", ".jpeg"):
        mime = "image/jpeg"
    elif suffix == ".svg":
        mime = "image/svg+xml"
    return f"data:{mime};base64,{encoded}"


def default_json_payload_path(html_path: str) -> str:
    return str(Path(html_path).with_name("latest.json"))


def html_to_png(html_path: str, png_path: str, width: int = 1920, height: int = 1080) -> None:
    script = f'''
const {{ chromium }} = require("playwright");
const path = require("path");
(async () => {{
  const browser = await chromium.launch({{ headless: true }});
  const page = await browser.newPage({{
    viewport: {{ width: {width}, height: {height} }},
    deviceScaleFactor: 1
  }});
  await page.goto("file://" + path.resolve({json.dumps(html_path)}), {{ waitUntil: "networkidle" }});
  await page.emulateMedia({{ media: "screen" }});
  await page.screenshot({{
    path: path.resolve({json.dumps(png_path)}),
    fullPage: true
  }});
  await browser.close();
}})().catch(err => {{ console.error(err); process.exit(1); }});
'''
    subprocess.run(["node", "-e", script], check=True)


def fetch_order_items(session: requests.Session, start: str, end: str, verbose: bool = False) -> List[Dict[str, Any]]:
    params = {"time_start": start, "time_end": end}
    log(f"[1/6] Requesting sold items report for {start} -> {end}", verbose)
    payload = api_get(session, "/reports/order_items", params=params)
    items = payload.get("data", [])
    if not isinstance(items, list):
        raise ApiError("Unexpected /reports/order_items response shape.")
    return items


def normalize_tag_list(value: Any) -> List[str]:
    if isinstance(value, list):
        return [str(v) for v in value]
    if value in (None, ""):
        return []
    return [str(value)]


def is_jewelry_category(value: Any) -> bool:
    text = str(value or "").lower()
    return "jewelry" in text


def filter_order_items(items: Iterable[Dict[str, Any]], args: argparse.Namespace) -> List[Dict[str, Any]]:
    filtered: List[Dict[str, Any]] = []
    want_tags = {t.lower() for t in args.tag}
    for item in items:
        if item.get("order_cancelled_at"):
            continue
        if is_jewelry_category(item.get("product_category")):
            continue
        if args.channel and str(item.get("channel", "")).lower() != args.channel.lower():
            continue
        if args.category and args.category.lower() not in str(item.get("product_category", "")).lower():
            continue
        if args.supplier and args.supplier.lower() not in str(item.get("supplier", "")).lower():
            continue
        item_tags = {t.lower() for t in normalize_tag_list(item.get("tags"))}
        if want_tags and not want_tags.issubset(item_tags):
            continue
        filtered.append(item)
    return filtered


def sort_items(items: List[Dict[str, Any]], sort_mode: str) -> List[Dict[str, Any]]:
    def key(item: Dict[str, Any]) -> Any:
        if sort_mode == "price":
            return (safe_float(item.get("order_item_price")), safe_float(item.get("order_item_subtotal")))
        if sort_mode == "subtotal":
            return (safe_float(item.get("order_item_subtotal")), safe_float(item.get("order_total")))
        if sort_mode == "paid_at":
            return parse_report_datetime(item.get("order_paid_at")) or datetime.min.replace(tzinfo=timezone.utc)
        return str(item.get("product_title", "")).lower()

    reverse = sort_mode != "title"
    return sorted(items, key=key, reverse=reverse)


def fetch_product(session: requests.Session, product_id: Any, verbose: bool = False) -> Dict[str, Any]:
    log(f"      ↳ fetching product details for product_id={product_id}", verbose)
    return api_get(session, f"/v4/products/{product_id}")


def extract_listing_url(product: Dict[str, Any]) -> str:
    for bucket in ("active_listings", "canon_listings", "product_listings"):
        for listing in product.get(bucket, []) or []:
            url = listing.get("listing_url")
            if url:
                return url
    return ""


def select_primary_image(product: Dict[str, Any], image_mode: str) -> str:
    if image_mode == "none":
        return ""
    images = product.get("images") or []
    if not images:
        return ""
    ranked = sorted(images, key=lambda x: (x.get("rank") is None, x.get("rank", 9999)))
    return str(ranked[0].get("url", ""))


def enrich_items(
    session: requests.Session,
    items: List[Dict[str, Any]],
    args: argparse.Namespace,
    product_cache: Optional[Dict[str, Dict[str, Any]]] = None,
    progress_label: str = "items",
) -> List[Dict[str, Any]]:
    enriched: List[Dict[str, Any]] = []
    product_cache = product_cache if product_cache is not None else {}
    fetched_count = 0
    total = len(items)
    for index, item in enumerate(items, start=1):
        product_id = item.get("upright_product_id") or item.get("product_id")
        if args.verbose:
            log(f"      processing {progress_label} {index}/{total}", True)
        if not product_id:
            product = {}
        else:
            key = str(product_id)
            if key not in product_cache:
                if args.max_products is not None and fetched_count >= args.max_products:
                    product_cache[key] = {}
                else:
                    product_cache[key] = fetch_product(session, product_id, verbose=args.verbose)
                    fetched_count += 1
            product = product_cache[key]
        image_url = select_primary_image(product, args.image_mode)
        paid_at = parse_report_datetime(item.get("order_paid_at"))
        enriched.append({
            "report": item,
            "product": product,
            "product_id": product.get("id") or product_id or "",
            "title": product.get("title") or item.get("product_title") or "Untitled item",
            "sku": product.get("sku") or item.get("product_sku") or "",
            "category": (product.get("category") or {}).get("path_cache") or item.get("product_category") or "",
            "supplier": (product.get("supplier") or {}).get("name") or item.get("supplier") or "",
            "inventory_location": (product.get("inventory_location") or {}).get("name") or item.get("inventory_location") or "",
            "image_url": image_url,
            "listing_url": extract_listing_url(product),
            "paid_at": paid_at,
            "item_price": safe_float(item.get("order_item_price")),
            "item_subtotal": safe_float(item.get("order_item_subtotal")),
            "quantity": int(item.get("quantity") or 0),
            "channel": item.get("channel") or "",
            "tags": product.get("tags") or item.get("tags") or [],
            "bid_count": ((product.get("active_listings") or [{}])[0] or {}).get("bid_count"),
            "view_count": ((product.get("active_listings") or [{}])[0] or {}).get("view_count"),
            "order_total": safe_float(item.get("order_total")),
            "shipping_total": safe_float(item.get("order_shipping_total")),
            "handling_total": safe_float(item.get("order_handling_total")),
            "buyer_id": item.get("channel_buyer_id") or "",
        })
    return enriched


def category_bucket(category: str) -> str:
    value = (category or "").lower()
    if any(term in value for term in ("clothing", "women's clothing", "men's clothing", "apparel", "shoes", "purses", "fashion")):
        return "apparel"
    if any(term in value for term in ("collectible", "collectibles", "toys", "memorabilia", "books", "media", "games", "art", "dolls", "comics", "cards")):
        return "collectibles"
    return "other"


def infer_brand(title: str) -> str:
    if not title:
        return "Unknown"

    cleaned = re.sub(r"[^A-Za-z0-9&'\- ]+", " ", title).strip()
    words = [w for w in cleaned.split() if w]
    if not words:
        return "Unknown"

    known_brands = [
        "Dooney & Bourke",
        "Abercrombie & Fitch",
        "Polo Ralph Lauren",
        "The North Face",
        "Michael Kors",
        "Kate Spade",
        "Under Armour",
        "Banana Republic",
        "Nike Air",
        "Nike Men's",
        "Coach New York",
    ]
    lowered_title = " ".join(words).lower()
    for brand in known_brands:
        if lowered_title.startswith(brand.lower()):
            return brand

    filtered = []
    for word in words[:6]:
        token = re.sub(r"[^A-Za-z0-9&'\-]+", "", word).strip()
        if not token:
            continue
        if token.lower() in GENERIC_BRAND_TERMS:
            continue
        filtered.append(token)

    if not filtered:
        return "Unknown"

    return filtered[0]


def normalize_collectible_brand(brand: str, title: str) -> str:
    brand = (brand or "").strip()
    title_lower = (title or "").lower()
    brand_lower = brand.lower()

    generic_terms = {
        "art glass", "glass", "bulk", "sports card", "mixed sports", "card", "cards", "doll",
        "barbie doll", "toy lot", "lot", "figurine", "collectible", "collectibles", "vintage lot",
        "new", "set", "vintage"
    }
    if brand_lower in generic_terms or brand_lower in GENERIC_BRAND_TERMS:
        return ""

    if "lego" in title_lower:
        return "LEGO"
    if "barbie" in title_lower:
        return "Barbie"
    if "american girl" in title_lower:
        return "American Girl"
    if "star wars" in title_lower:
        return "Star Wars"
    if "hot wheels" in title_lower:
        return "Hot Wheels"
    if "pokemon" in title_lower:
        return "Pokemon"
    if "funko" in title_lower:
        return "Funko"
    if "marvel" in title_lower:
        return "Marvel"
    if "dc comics" in title_lower or title_lower.startswith("dc "):
        return "DC Comics"

    reject_contains = [
        "glass", "bulk", "sports", "card", "doll", "mixed", "lot", "figurine", "ornament", "plush",
        "new", "set", "vintage"
    ]
    if any(term in brand_lower for term in reject_contains):
        return ""

    return brand


def compute_brand_summary(items: List[Dict[str, Any]]) -> Dict[str, List[tuple[str, int]]]:
    buckets = {"apparel": Counter(), "collectibles": Counter()}
    for item in items:
        bucket = category_bucket(str(item.get("product_category") or ""))
        if bucket not in buckets:
            continue
        title = str(item.get("product_title") or "")
        brand = infer_brand(title)
        if bucket == "collectibles":
            brand = normalize_collectible_brand(brand, title)
        if brand and brand.lower() != "unknown" and brand.lower() not in GENERIC_BRAND_TERMS:
            buckets[bucket][brand] += 1
    return {
        "apparel": buckets["apparel"].most_common(6),
        "collectibles": buckets["collectibles"].most_common(6),
    }


def compute_summary(items: List[Dict[str, Any]], featured: List[Dict[str, Any]]) -> Dict[str, Any]:
    revenue = sum(safe_float(i.get("order_item_subtotal")) for i in items)
    prices = [safe_float(i.get("order_item_price")) for i in items]
    quantities = [int(i.get("quantity") or 0) for i in items]
    suppliers = Counter(str(i.get("supplier") or "Unknown supplier") for i in items)
    top_item = max(featured, key=lambda x: x.get("item_price", 0.0), default=None)
    brand_summary = compute_brand_summary(items)
    return {
        "total_items": len(items),
        "total_units": sum(quantities),
        "revenue": revenue,
        "avg_price": statistics.mean(prices) if prices else 0.0,
        "median_price": statistics.median(prices) if prices else 0.0,
        "max_price": max(prices) if prices else 0.0,
        "top_suppliers": suppliers.most_common(6),
        "brand_summary": brand_summary,
        "highest_item": top_item,
    }


def esc(value: Any) -> str:
    return html.escape(str(value or ""))


def truncate_text(value: str, limit: int = 78) -> str:
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    if len(value) <= limit:
        return value
    return value[: limit - 1].rsplit(" ", 1)[0].strip() + "…"


def build_list(items: List[tuple[str, Any]], empty: str = "No data") -> str:
    if not items:
        return f'<li class="muted">{esc(empty)}</li>'
    return "".join(
        f'<li><span>{esc(name)}</span><strong>{esc(value)}</strong></li>'
        for name, value in items
    )


def build_kpi_cards(summary: Dict[str, Any]) -> str:
    highest = summary.get("highest_item")
    top_supplier = summary.get("top_suppliers", [("N/A", 0)])[0][0] if summary.get("top_suppliers") else "N/A"
    cards = [
        ("Sold Items", summary["total_items"]),
        ("Revenue", fmt_currency(summary["revenue"])),
        ("Avg Price", fmt_currency(summary["avg_price"])),
        ("Top Supplier", top_supplier),
        ("Highest Featured", fmt_currency(highest["item_price"]) if highest else "$0.00"),
    ]
    return "".join(
        f'''
        <div class="kpi-card">
          <div class="kpi-label">{esc(label)}</div>
          <div class="kpi-value">{esc(value)}</div>
        </div>
        '''
        for label, value in cards
    )


def build_featured_cards(featured: List[Dict[str, Any]]) -> str:
    if not featured:
        return '<div class="empty-state">No featured items available for this run.</div>'

    blocks: List[str] = []
    for index, item in enumerate(featured, start=1):
        image_html = (
            f'<img src="{esc(item["image_url"])}" alt="{esc(item["title"])}" />'
            if item.get("image_url")
            else '<div class="image-placeholder">No image available</div>'
        )

        detail_bits = []
        if item.get("paid_at"):
            detail_bits.append(f'Sold {fmt_short_date(item["paid_at"])}')
        if item.get("channel"):
            detail_bits.append(item["channel"])
        if item.get("bid_count") not in (None, ""):
            detail_bits.append(f'Bids {item["bid_count"]}')
        if item.get("view_count") not in (None, ""):
            detail_bits.append(f'Views {item["view_count"]}')
        details_line = " • ".join(detail_bits) if detail_bits else "No detail data"

        listing_link = (
            f'<a class="button" href="{esc(item["listing_url"])}" target="_blank" rel="noopener">View Listing</a>'
            if item.get("listing_url")
            else '<span class="button disabled">No Listing URL</span>'
        )

        blocks.append(
            f'''
            <article class="item-card">
              <div class="item-rank">#{index}</div>
              <div class="item-media">{image_html}</div>
              <div class="item-content">
                <div class="supplier-ribbon">{esc((item.get("supplier") or "Unknown Supplier").upper())}</div>
                <div class="item-body">
                  <h3>{esc(truncate_text(item["title"]))}</h3>
                  <div class="price-row">
                    <span class="price">{fmt_currency(item["item_price"])}</span>
                  </div>
                  <div class="details-row">{esc(details_line)}</div>
                  <div class="actions">{listing_link}</div>
                </div>
              </div>
            </article>
            '''
        )
    return "\n".join(blocks)


def render_html(
    args: argparse.Namespace,
    summary: Dict[str, Any],
    featured: List[Dict[str, Any]],
    generated_at: datetime,
) -> str:
    report_span = fmt_report_span(args.start, args.end)
    logo_uri = logo_data_uri(args.logo_path)
    logo_html = f'<div class="logo-shell"><img src="{logo_uri}" alt="Goodwill logo" /></div>' if logo_uri else ""
    org_line = f'<div class="org-name">{esc(args.organization_name)}</div>' if args.organization_name else ""
    top_suppliers = summary.get("top_suppliers", [])
    apparel_brands = summary.get("brand_summary", {}).get("apparel", [])
    collectible_brands = summary.get("brand_summary", {}).get("collectibles", [])

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="refresh" content="{int(args.refresh_seconds)}" />
  <title>{esc(args.title)}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Lato:wght@400;700;900&display=swap');

    :root {{
      --goodwill-blue: #00539F;
      --goodwill-dark: #212721;
      --goodwill-gray: #52585A;
      --goodwill-white: #FFFFFF;
      --goodwill-primary: #4F87C6;
      --goodwill-light: #B2D235;
      --bg-panel: rgba(255,255,255,0.97);
      --shadow: 0 16px 40px rgba(0, 0, 0, 0.18);
      --radius: 24px;
      --space: clamp(12px, 1.2vw, 24px);
    }}

    * {{ box-sizing: border-box; }}

    html, body {{
      margin: 0;
      min-height: 100%;
      font-family: 'Lato', Arial, Helvetica, sans-serif;
      background:
        radial-gradient(circle at top right, rgba(178,210,53,0.18), transparent 20%),
        linear-gradient(135deg, #00539F 0%, #004885 60%, #003867 100%);
      color: var(--goodwill-dark);
    }}

    body {{
      min-height: 100vh;
    }}

    .page {{
      width: 100%;
      max-width: 1800px;
      margin: 0 auto;
      padding: clamp(16px, 2vw, 36px);
      display: grid;
      gap: var(--space);
    }}

    .hero {{
      background: rgba(255,255,255,0.10);
      border: 1px solid rgba(255,255,255,0.18);
      border-radius: 28px;
      box-shadow: var(--shadow);
      color: white;
      padding: clamp(18px, 2vw, 30px);
      display: grid;
      grid-template-columns: auto 1fr auto;
      gap: var(--space);
      align-items: center;
    }}

    .logo-shell {{
      background: rgba(255,255,255,0.96);
      border-radius: 22px;
      padding: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
      min-width: 92px;
      min-height: 92px;
    }}

    .logo-shell img {{
      max-width: 110px;
      width: 100%;
      height: auto;
      display: block;
    }}

    .hero-copy h1 {{
      margin: 0;
      font-size: clamp(2rem, 4vw, 4rem);
      line-height: 1.02;
      font-weight: 900;
    }}

    .hero-subline {{
      margin-top: 8px;
      font-size: clamp(1rem, 1.5vw, 1.5rem);
      font-weight: 700;
      color: rgba(255,255,255,0.9);
    }}

    .org-name {{
      margin-top: 6px;
      font-size: clamp(0.95rem, 1.2vw, 1.2rem);
      color: rgba(255,255,255,0.88);
      font-weight: 700;
    }}

    .report-box {{
      background: rgba(255,255,255,0.14);
      border-left: 6px solid var(--goodwill-light);
      border-radius: 18px;
      padding: 14px 16px;
      min-width: min(34vw, 360px);
    }}

    .report-label {{
      text-transform: uppercase;
      letter-spacing: .08em;
      font-size: 0.8rem;
      color: rgba(255,255,255,.84);
      margin-bottom: 6px;
    }}

    .report-value {{
      font-size: clamp(1.1rem, 2vw, 1.8rem);
      font-weight: 900;
      color: white;
      line-height: 1.2;
    }}

    .kpi-row {{
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: var(--space);
    }}

    .kpi-card {{
      background: rgba(255,255,255,0.98);
      border-top: 8px solid var(--goodwill-light);
      border-radius: 20px;
      box-shadow: var(--shadow);
      padding: 16px 18px 14px;
      min-width: 0;
    }}

    .kpi-label {{
      color: var(--goodwill-gray);
      text-transform: uppercase;
      letter-spacing: .08em;
      font-size: 0.78rem;
      font-weight: 700;
    }}

    .kpi-value {{
      margin-top: 10px;
      color: var(--goodwill-blue);
      font-size: clamp(1.25rem, 2vw, 2.2rem);
      line-height: 1.05;
      font-weight: 900;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}

    .main {{
      display: grid;
      grid-template-columns: 1.8fr 1fr;
      gap: var(--space);
      align-items: start;
    }}

    .panel {{
      background: var(--bg-panel);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      overflow: hidden;
    }}

    .panel-header {{
      padding: 16px 20px 12px;
      border-bottom: 2px solid rgba(178,210,53,0.55);
      color: var(--goodwill-blue);
      font-size: clamp(1.2rem, 1.8vw, 2rem);
      font-weight: 900;
      background: linear-gradient(180deg, rgba(79,135,198,0.08), rgba(79,135,198,0.02));
    }}

    .featured-grid {{
      padding: var(--space);
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
      gap: var(--space);
    }}

    .item-card {{
      background: white;
      border: 2px solid rgba(0,83,159,0.10);
      border-top: 8px solid var(--goodwill-light);
      border-radius: 20px;
      overflow: hidden;
      display: grid;
      grid-template-columns: 64px minmax(220px, 32%) 1fr;
      min-height: 100%;
    }}

    .item-rank {{
      background: var(--goodwill-blue);
      color: white;
      font-weight: 900;
      font-size: clamp(1.1rem, 1.8vw, 1.8rem);
      display: flex;
      align-items: center;
      justify-content: center;
    }}

    .item-media {{
      background: #f6fbdc;
      border-left: 1px solid rgba(0,83,159,0.08);
      border-right: 1px solid rgba(0,83,159,0.08);
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 12px;
      min-height: 220px;
    }}

    .item-media img {{
      width: 100%;
      height: 100%;
      max-height: 240px;
      object-fit: contain;
      border-radius: 14px;
      background: white;
    }}

    .image-placeholder {{
      color: var(--goodwill-gray);
      font-size: 1rem;
      font-weight: 700;
      text-align: center;
    }}

    .item-content {{
      display: grid;
      grid-template-rows: auto 1fr;
      min-width: 0;
    }}

    .supplier-ribbon {{
      background: var(--goodwill-light);
      color: var(--goodwill-dark);
      font-size: clamp(0.9rem, 1.2vw, 1.15rem);
      font-weight: 900;
      letter-spacing: .03em;
      padding: 10px 14px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}

    .item-body {{
      padding: 14px 16px 16px;
      display: grid;
      grid-template-rows: 3.2em auto auto auto;
      gap: 8px;
      min-width: 0;
    }}

    .item-body h3 {{
      margin: 0;
      font-size: clamp(1.05rem, 1.4vw, 1.4rem);
      line-height: 1.12;
      font-weight: 800;
      color: var(--goodwill-dark);
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
      min-height: 2.25em;
      max-height: 2.25em;
    }}

    .price-row {{
      display: flex;
      align-items: baseline;
      gap: 10px;
      flex-wrap: wrap;
    }}

    .price {{
      font-size: clamp(1.35rem, 2vw, 2rem);
      line-height: 1;
      font-weight: 900;
      color: var(--goodwill-blue);
    }}

    .details-row {{
      font-size: clamp(0.84rem, 1vw, 1rem);
      line-height: 1.2;
      color: var(--goodwill-gray);
      font-weight: 700;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}

    .actions {{
      display: flex;
      align-items: end;
      gap: 10px;
      margin-top: 4px;
    }}

    .button {{
      display: inline-block;
      text-decoration: none;
      background: var(--goodwill-blue);
      color: white;
      padding: 9px 12px;
      border-radius: 10px;
      font-size: 0.9rem;
      font-weight: 800;
      border: 2px solid var(--goodwill-light);
    }}

    .button.disabled {{
      background: #cfd7da;
      border-color: transparent;
      color: #4b5458;
    }}

    .empty-state {{
      padding: 40px;
      color: var(--goodwill-gray);
      font-size: 1.2rem;
      font-weight: 700;
      text-align: center;
    }}

    .side-stack {{
      display: grid;
      gap: var(--space);
    }}

    .list-panel-body {{
      padding: 14px 20px 18px;
    }}

    .list-panel ul {{
      list-style: none;
      padding: 0;
      margin: 0;
    }}

    .list-panel li {{
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 14px;
      align-items: baseline;
      padding: 12px 0;
      border-bottom: 1px solid rgba(0,83,159,0.10);
      font-size: clamp(0.95rem, 1.2vw, 1.15rem);
    }}

    .list-panel li:last-child {{
      border-bottom: none;
    }}

    .list-panel span {{
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }}

    .list-panel strong {{
      color: var(--goodwill-blue);
      font-weight: 900;
    }}

    .muted {{
      color: var(--goodwill-gray);
    }}

    .footer {{
      display: grid;
      grid-template-columns: 1fr auto auto;
      gap: 18px;
      align-items: center;
      color: rgba(255,255,255,0.92);
      font-size: 0.95rem;
      padding: 0 4px 12px;
    }}

    .footer .left {{
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }}

    .footer .right {{
      font-weight: 700;
    }}

    @media (max-width: 1180px) {{
      .hero {{
        grid-template-columns: 1fr;
      }}
      .report-box {{
        min-width: 0;
      }}
      .kpi-row {{
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }}
      .main {{
        grid-template-columns: 1fr;
      }}
    }}

    @media (max-width: 860px) {{
      .kpi-row {{
        grid-template-columns: 1fr;
      }}
      .item-card {{
        grid-template-columns: 1fr;
      }}
      .item-rank {{
        min-height: 54px;
      }}
      .item-media {{
        min-height: 240px;
      }}
      .footer {{
        grid-template-columns: 1fr;
        text-align: center;
      }}
    }}
  </style>
</head>
<body>
  <main class="page">
    <section class="hero">
      {logo_html}
      <div class="hero-copy">
        <h1>{esc(args.title)}</h1>
        <div class="hero-subline">Top suppliers, top brands, and highest-value items from the last 3 days</div>
        {org_line}
      </div>
      <div class="report-box">
        <div class="report-label">Reporting Window</div>
        <div class="report-value">{esc(report_span)}</div>
      </div>
    </section>

    <section class="kpi-row">
      {build_kpi_cards(summary)}
    </section>

    <section class="main">
      <section class="panel">
        <div class="panel-header">Featured Items</div>
        <div class="featured-grid">
          {build_featured_cards(featured)}
        </div>
      </section>

      <section class="side-stack">
        <section class="panel list-panel">
          <div class="panel-header">Top Suppliers</div>
          <div class="list-panel-body">
            <ul>{build_list(top_suppliers, "No supplier data")}</ul>
          </div>
        </section>

        <section class="panel list-panel">
          <div class="panel-header">Top Apparel Brands</div>
          <div class="list-panel-body">
            <ul>{build_list(apparel_brands, "No apparel brand data")}</ul>
          </div>
        </section>

        <section class="panel list-panel">
          <div class="panel-header">Top Collectibles Brands</div>
          <div class="list-panel-body">
            <ul>{build_list(collectible_brands, "No collectibles brand data")}</ul>
          </div>
        </section>
      </section>
    </section>

    <footer class="footer">
      <div class="left">Generated by responsive dashboard using Upright Labs data.</div>
      <div class="right">Generated: {esc(fmt_generated_at(generated_at))}</div>
      <div class="right">Auto-refresh: {int(args.refresh_seconds)}s</div>
    </footer>
  </main>
</body>
</html>
"""


def write_metadata_json(
    json_path: str,
    args: argparse.Namespace,
    summary: Dict[str, Any],
    featured: List[Dict[str, Any]],
    generated_at: datetime,
) -> None:
    payload = {
        "title": args.title,
        "organization_name": args.organization_name,
        "start": args.start,
        "end": args.end,
        "report_span": fmt_report_span(args.start, args.end),
        "generated_at": isoformat_z(generated_at),
        "total_items": summary.get("total_items", 0),
        "total_units": summary.get("total_units", 0),
        "revenue": round(float(summary.get("revenue", 0.0)), 2),
        "avg_price": round(float(summary.get("avg_price", 0.0)), 2),
        "top_suppliers": summary.get("top_suppliers", []),
        "top_apparel_brands": summary.get("brand_summary", {}).get("apparel", []),
        "top_collectibles_brands": summary.get("brand_summary", {}).get("collectibles", []),
        "featured": [
            {
                "rank": index,
                "title": item.get("title"),
                "supplier": item.get("supplier"),
                "price": round(float(item.get("item_price", 0.0)), 2),
                "paid_at": item.get("paid_at").isoformat() if item.get("paid_at") else None,
                "channel": item.get("channel"),
                "bid_count": item.get("bid_count"),
                "view_count": item.get("view_count"),
                "listing_url": item.get("listing_url"),
                "image_url": item.get("image_url"),
            }
            for index, item in enumerate(featured, start=1)
        ],
    }

    path = Path(json_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def main() -> None:
    load_dotenv()
    args = parse_args()

    default_start, default_end = default_date_window()
    args.start = validate_iso8601(args.start or os.getenv("TV_DASHBOARD_START", default_start))
    args.end = validate_iso8601(args.end or os.getenv("TV_DASHBOARD_END", default_end))
    args.top = args.top if args.top is not None else int(os.getenv("TV_DASHBOARD_TOP", "3"))
    args.sort = args.sort or os.getenv("TV_DASHBOARD_SORT", "price")
    if args.sort not in SORT_CHOICES:
        raise SystemExit(f"Invalid sort mode: {args.sort}")
    args.channel = args.channel or os.getenv("TV_DASHBOARD_CHANNEL")
    args.category = args.category or os.getenv("TV_DASHBOARD_CATEGORY")
    args.supplier = args.supplier or os.getenv("TV_DASHBOARD_SUPPLIER")
    if not args.tag:
        env_tags = os.getenv("TV_DASHBOARD_TAGS", "").strip()
        args.tag = [tag.strip() for tag in env_tags.split(",") if tag.strip()] if env_tags else []
    args.output = args.output or os.getenv("TV_DASHBOARD_OUTPUT", DEFAULT_OUTPUT)
    args.png_output = args.png_output or os.getenv("TV_DASHBOARD_PNG_OUTPUT", default_png_path(args.output))
    args.json_output = args.json_output or os.getenv("TV_DASHBOARD_JSON_OUTPUT", default_json_path(args.output))
    args.title = args.title or os.getenv("TV_DASHBOARD_TITLE", default_title())
    args.organization_name = args.organization_name or os.getenv(
        "TV_DASHBOARD_ORGANIZATION",
        "Goodwill Industries of Greater Cleveland and East Central Ohio, Inc.",
    )
    args.image_mode = args.image_mode or os.getenv("TV_DASHBOARD_IMAGE_MODE", "remote")

    if args.top <= 0:
        raise SystemExit("--top must be greater than 0")

    verbose = True if args.verbose else True
    generated_at = datetime.now(timezone.utc)

    log("Starting responsive dashboard generation...", verbose)
    log(f"  Title: {args.title}", verbose)
    log(f"  HTML Output: {args.output}", verbose)
    log(f"  PNG Output: {args.png_output}", verbose)
    log(f"  JSON Output: {args.json_output}", verbose)
    log(f"  Date window: {args.start} -> {args.end}", verbose)
    log(f"  Sort mode: {args.sort}", verbose)
    log(f"  Top featured items requested: {args.top}", verbose)
    log("  Business rules: jewelry excluded, supplier emphasized, responsive browser rendering enabled", verbose)

    token = require_token(args.token)
    log("Token loaded successfully.", verbose)
    session = build_session(token)

    items = fetch_order_items(session, args.start, args.end, verbose=verbose)
    log(f"[2/6] Retrieved {len(items)} sold item rows from the report", verbose)

    filtered = filter_order_items(items, args)
    excluded = len(items) - len(filtered)
    log(f"[3/6] Applied filters and business rules. {len(filtered)} items remain; {excluded} excluded", verbose)

    ranked = sort_items(filtered, args.sort)
    featured_raw = ranked[: args.top]
    log(f"[4/6] Ranked items by '{args.sort}' and selected {len(featured_raw)} featured items", verbose)

    shared_cache: Dict[str, Dict[str, Any]] = {}
    log(f"[5/6] Enriching {len(featured_raw)} featured items with product details and images...", verbose)
    featured = enrich_items(session, featured_raw, args, product_cache=shared_cache, progress_label="featured item")
    log("      featured enrichment complete", verbose)

    summary = compute_summary(filtered, featured)
    log("[6/6] Rendering responsive HTML dashboard...", verbose)
    output_html = render_html(args, summary, featured, generated_at)

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(output_html, encoding="utf-8")

    json_path = Path(args.json_output)
    json_path.parent.mkdir(parents=True, exist_ok=True)
    write_metadata_json(str(json_path), args, summary, featured, generated_at)

    if not args.no_png:
        png_path = Path(args.png_output)
        png_path.parent.mkdir(parents=True, exist_ok=True)
        log("      generating optional PNG snapshot...", verbose)
        html_to_png(str(output_path), str(png_path), width=1920, height=1080)
        log(f"Completed. PNG written to {png_path}", verbose)

    log(f"Completed. HTML written to {output_path}", verbose)
    log(f"Completed. JSON written to {json_path}", verbose)


if __name__ == "__main__":
    try:
        main()
    except ApiError as exc:
        raise SystemExit(str(exc))
