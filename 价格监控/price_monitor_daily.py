import json
import re
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "price_monitor_input_43_updated.xlsx" if (ROOT / "price_monitor_input_43_updated.xlsx").exists() else ROOT / "price_monitor_input_43.xlsx"
DASHBOARD = ROOT / "price_monitor_dashboard.html"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131 Safari/537.36",
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
}


def parse_price(raw):
    raw = raw.replace(".", "").replace(",", ".")
    try:
        value = float(raw)
    except ValueError:
        return None
    return round(value, 2) if 20 < value < 10000 else None


def extract_price(html, retailer):
    text = re.sub(r"<script[\s\S]*?</script>|<style[\s\S]*?</style>|<[^>]+>", " ", html, flags=re.I)
    text = re.sub(r"\s+", " ", text.replace("&euro;", "€"))
    patterns = (
        [r"(\d{1,3}(?:\.\d{3})*,\d{2})\s*€\s*(?:CONSEGNA|IVA inclusa)", r"(\d{1,3}(?:\.\d{3})*,\d{2})\s*€"]
        if retailer == "MediaWorld"
        else [r"€\s*(\d{1,3}(?:\.\d{3})*,\d{2})\s*Iva inclusa", r"€\s*(\d{1,3}(?:\.\d{3})*,\d{2})"]
    )
    for pattern in patterns:
        for match in re.finditer(pattern, text, re.I):
            price = parse_price(match.group(1))
            if price is not None:
                return price
    return None


def split_item(item):
    normalized = str(item).replace("_", " ").strip()
    match = re.match(r"^([A-Za-z]+)\s+(\d{2})(.*)$", normalized)
    if not match:
        return "", "", normalized
    return match.group(2), match.group(1).upper(), f"{match.group(2)}{match.group(3)}"


def load_products():
    sheet = load_workbook(INPUT, data_only=True)["Input"]
    headers = [str(value or "").strip().lower() for value in next(sheet.iter_rows(values_only=True))]
    item_col = headers.index("item name")
    ue_col = next(i for i, value in enumerate(headers) if "unieuro" in value)
    mw_col = next(i for i, value in enumerate(headers) if "mediaworld" in value)
    products = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row[item_col]:
            continue
        products.append({
            "item": str(row[item_col]),
            "Unieuro": str(row[ue_col]) if row[ue_col] else None,
            "MediaWorld": str(row[mw_col]) if row[mw_col] else None,
        })
    return products


def fetch_price(url, retailer):
    if not url:
        return "/"
    try:
        request = Request(url, headers=HEADERS)
        with urlopen(request, timeout=30) as response:
            return extract_price(response.read().decode("utf-8", errors="ignore"), retailer)
    except (HTTPError, URLError, TimeoutError):
        return None


def load_history(html, fallback_date):
    match = re.search(r"const DATA = (\[.*?\]);\s*\nconst \$", html, flags=re.S)
    if not match:
        match = re.search(r"const DATA = (\[.*?\])\.map", html, flags=re.S)
    history = json.loads(match.group(1)) if match else []
    existing_date_match = re.search(r"最新数据：([0-9]{4}-[0-9]{2}-[0-9]{2})", html)
    existing_date = existing_date_match.group(1) if existing_date_match else fallback_date
    for row in history:
        row.setdefault("date", existing_date)
    return history


def main():
    today = datetime.now(ZoneInfo("Asia/Shanghai")).date().isoformat()
    html = DASHBOARD.read_text(encoding="utf-8")
    history = load_history(html, today)
    products = load_products()
    current = []
    valid_keys = set()
    tasks = []
    for product in products:
        size, brand, model = split_item(product["item"])
        for channel in ("Unieuro", "MediaWorld"):
            if not product[channel]:
                continue
            valid_keys.add((size, brand, model, channel))
            tasks.append((size, brand, model, channel, product[channel]))

    def scrape(task):
        size, brand, model, channel, url = task
        return {"size": size, "brand": brand, "model": model, "channel": channel, "price": fetch_price(url, channel), "date": today}

    with ThreadPoolExecutor(max_workers=12) as pool:
        current = list(pool.map(scrape, tasks))
    keys = {(r["size"], r["brand"], r["model"], r["channel"]) for r in current}
    history = [r for r in history if (r.get("size"), r.get("brand"), r.get("model"), r.get("channel")) in valid_keys and ((r.get("size"), r.get("brand"), r.get("model"), r.get("channel")) not in keys or r.get("date") != today)]
    payload = json.dumps(history + current, ensure_ascii=False, separators=(",", ":"))
    html = re.sub(r"const DATA = .*?;\s*\nconst \$", f"const DATA = {payload};\nconst $", html, count=1, flags=re.S)
    html = re.sub(r"最新数据：[0-9]{4}-[0-9]{2}-[0-9]{2}", f"最新数据：{today}", html)
    DASHBOARD.write_text(html, encoding="utf-8")
    print(f"updated {DASHBOARD} with {len(current)} records for {today}")


if __name__ == "__main__":
    main()
