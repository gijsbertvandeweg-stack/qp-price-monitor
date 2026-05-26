"""
QP Price Monitor — Wekelijkse prijsscraper
Haalt per product de actuele prijs op via de URL in het Excel-bestand
en voegt nieuwe rijen toe aan Sheet1.
"""
import json
import logging
import re
import time
import random
from datetime import date
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# ── Configuratie ───────────────────────────────────────────────────────────────
EXCEL_PATH = Path(__file__).parent / "data.xlsx"
LOG_PATH   = Path(__file__).parent / "scraper.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

# Minimale / maximale wachttijd tussen requests (seconden)
DELAY_MIN, DELAY_MAX = 2.5, 6.0

# Sites die een Playwright-browser nodig hebben (bot-detectie blokkeert requests)
PLAYWRIGHT_SITES = [
    "jumbo.com", "aldi.nl", "dirk.nl", "vomar.nl",
    "dekamarkt.nl", "hoogvliet.com", "poiesz-supermarkten.nl",
    "plus.nl",
]

# Sites waarvoor we eerst de homepage bezoeken om cookies te krijgen
SESSION_SITES = {
    "ah.nl": "https://www.ah.nl/",
}

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "nl-NL,nl;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Sec-Fetch-Mode": "navigate",
}

# ── Prijsparsing ───────────────────────────────────────────────────────────────
def _to_float(raw) -> float | None:
    if raw is None:
        return None
    s = str(raw).replace("€", "").replace("\xa0", "").replace(" ", "")
    s = re.sub(r"[^\d,\.]", "", s.split("\n")[0])
    s = s.replace(",", ".")
    # "3.89" of "389" (gesplitste opmaak)
    m = re.match(r"^(\d+)\.(\d{2})$", s)
    if m:
        val = float(f"{m.group(1)}.{m.group(2)}")
        return val if 0.01 < val < 500 else None
    return None


def _price_from_text(text: str) -> float | None:
    """Parset prijstekst, inclusief gesplitste opmaak zoals '4\\n59' of '3.\\n89'."""
    # Formaat: euros + optionele punt + newline + cents (b.v. "4\n59", "3.\n89")
    m = re.search(r"(\d+)\.?\n(\d{2})(?!\d)", text)
    if m:
        val = float(f"{m.group(1)}.{m.group(2)}")
        return val if 0.01 < val < 500 else None

    # Standaard formaat: "3,89" of "3.89"
    cleaned = re.sub(r"[€\s\xa0]", "", text.replace("\n", ""))
    m = re.search(r"(\d+)[,\.](\d{2})(?!\d)", cleaned)
    if m:
        val = float(f"{m.group(1)}.{m.group(2)}")
        return val if 0.01 < val < 500 else None
    return None


def _extract_from_soup(soup: BeautifulSoup) -> tuple[float | None, str]:
    # 1. JSON-LD (Product offers)
    for sc in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(sc.string or "")
            if isinstance(data, list):
                data = data[0]
            if data.get("@type") == "Product":
                offers = data.get("offers", {})
                if isinstance(offers, list):
                    offers = offers[0]
                price = _to_float(offers.get("price") or offers.get("lowPrice"))
                if price:
                    return price, "Succes (JSON-LD)"
        except Exception:
            pass

    # 2. Schema.org microdata
    el = soup.find(itemprop="price")
    if el:
        price = _to_float(el.get("content") or el.get_text())
        if price:
            return price, "Succes (Schema.org)"

    # 3. Aldi: {"priceValue":2.49,...} in script tag (escaped quotes)
    for sc in soup.find_all("script"):
        m = re.search(r'priceValue.{1,6}?([\d]+\.[\d]{1,2})', sc.string or "")
        if m:
            price = _to_float(m.group(1))
            if price:
                return price, "Succes (JSON-LD)"

    # 4. HTML price elements
    for el in soup.find_all(class_=re.compile(r"price|prijs|Price|Prijs", re.I)):
        text = el.get_text(" ", strip=True)
        if "per kg" in text.lower() or "per kilo" in text.lower():
            continue
        price = _price_from_text(text)
        if price:
            return price, "Succes (HTML price)"

    return None, "Geen prijs gevonden"


# ── Requests-gebaseerd scrapen ────────────────────────────────────────────────
_sessions: dict[str, requests.Session] = {}


def _get_session(url: str) -> requests.Session:
    domain = re.search(r"https?://(?:www\.)?([^/]+)", url)
    key = domain.group(1) if domain else "default"

    if key not in _sessions:
        s = requests.Session()
        s.headers.update(BROWSER_HEADERS)
        # Bezoek homepage als dat voor deze site nodig is
        for site, homepage in SESSION_SITES.items():
            if site in url:
                try:
                    s.get(homepage, timeout=15)
                    log.debug(f"Homepage bezocht voor {site}")
                except Exception:
                    pass
                break
        _sessions[key] = s
    return _sessions[key]


def scrape_requests(url: str) -> tuple[float | None, str]:
    session = _get_session(url)
    try:
        resp = session.get(url, timeout=20, allow_redirects=True)
        if resp.status_code == 404:
            return None, "Niet beschikbaar"
        if resp.status_code != 200:
            return None, f"Mislukt (HTTP {resp.status_code})"
        soup = BeautifulSoup(resp.text, "lxml")
        return _extract_from_soup(soup)
    except requests.Timeout:
        return None, "Mislukt (timeout)"
    except Exception as e:
        return None, f"Mislukt ({str(e)[:80]})"


# ── Playwright-gebaseerd scrapen ──────────────────────────────────────────────
def scrape_playwright(url: str) -> tuple[float | None, str]:
    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(
                extra_http_headers={"Accept-Language": "nl-NL,nl;q=0.9"},
                viewport={"width": 1280, "height": 800},
                user_agent=BROWSER_HEADERS["User-Agent"],
            )
            page = ctx.new_page()
            # Blokkeer media voor snelheid
            page.route("**/*.{png,jpg,jpeg,gif,webp,svg,woff,woff2,ttf,mp4}", lambda r: r.abort())

            try:
                page.goto(url, wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(4000)
            except Exception:
                pass

            # Probeer JSON-LD/Schema in de HTML
            html = page.content()
            soup = BeautifulSoup(html, "lxml")
            price, status = _extract_from_soup(soup)
            if price:
                if status == "Succes (HTML price)":
                    status = "Succes (Visueel HTML)"
                browser.close()
                return price, status

            # Fallback: lees price-elementen via Playwright locator
            price = _playwright_price_locator(page)
            browser.close()

            if price:
                return price, "Succes (Visueel HTML)"

            # Check "niet beschikbaar" in paginatekst
            body_text = page.evaluate("document.body.innerText") if not html else html
            not_available_patterns = [
                "niet beschikbaar", "niet leverbaar", "uitverkocht",
                "product niet beschikbaar", "temporarily unavailable",
            ]
            body_lower = str(body_text).lower()
            if any(p in body_lower for p in not_available_patterns):
                return None, "Niet beschikbaar"

            return None, "Geen prijs gevonden"

    except Exception as e:
        return None, f"Mislukt (Playwright: {str(e)[:80]})"


def _playwright_price_locator(page) -> float | None:
    """Leest prijselementen via CSS-selectors en pikt de eerste redelijke prijs."""
    selectors = [
        "[class*=price]",
        "[class*=prijs]",
        "[class*=Price]",
        "[itemprop=price]",
        "[data-testid*=price]",
        "[data-test*=price]",
    ]
    for sel in selectors:
        try:
            els = page.locator(sel).all()
            for el in els:
                try:
                    txt = el.inner_text(timeout=800).strip()
                    if not txt or len(txt) > 100:
                        continue
                    if "per kg" in txt.lower() or "per kilo" in txt.lower():
                        continue
                    price = _price_from_text(txt)
                    if price and price > 0.10:
                        return price
                except Exception:
                    pass
        except Exception:
            pass
    return None


# ── Hoofd scrape-logica ────────────────────────────────────────────────────────
def fetch_price(url: str) -> tuple[float | None, str]:
    needs_playwright = any(site in url for site in PLAYWRIGHT_SITES)

    if not needs_playwright:
        price, status = scrape_requests(url)
        if price:
            return price, status
        # Fallback naar Playwright als requests niets oplevert
        log.info("  → Fallback naar Playwright")
        return scrape_playwright(url)

    return scrape_playwright(url)


# ── Hoofdfunctie ───────────────────────────────────────────────────────────────
def run_scraper():
    log.info("=" * 65)
    log.info("QP Price Monitor — scraper gestart")

    if not EXCEL_PATH.exists():
        log.error(f"Excel niet gevonden: {EXCEL_PATH}")
        return

    df = pd.read_excel(EXCEL_PATH, sheet_name="Sheet1")
    df["datum"] = pd.to_datetime(df["datum"])

    today = date.today()

    # Sla over als vandaag al bestaat
    df = df.dropna(subset=["datum"])
    if (df["datum"].dt.date == today).any():
        log.warning(f"Data voor {today} bestaat al — scraper gestopt.")
        return

    # Unieke product × retailer combinaties (meest recente URL per combinatie)
    cols = ["Leverancier", "Artikel_nummer", "URL", "product_naam", "retailer", "Concurrentie"]
    unique = (
        df.sort_values("datum")
        .drop_duplicates(subset=["URL"], keep="last")[cols]
        .dropna(subset=["URL"])
        .reset_index(drop=True)
    )
    today_dt = pd.Timestamp(today)

    log.info(f"Datum: {today}  |  Producten: {len(unique)}")

    new_rows = []
    for i, row in unique.iterrows():
        url = str(row["URL"])
        label = f"{row['retailer']} — {str(row['product_naam'])[:45]}"
        log.info(f"[{i+1}/{len(unique)}] {label}")

        price, status = fetch_price(url)

        if price:
            log.info(f"  → €{price:.2f} | {status}")
        else:
            log.warning(f"  → {status}")

        new_rows.append([
            row["Leverancier"],
            row["Artikel_nummer"],
            url,
            row["product_naam"],
            price,
            row["retailer"],
            status,
            today,
            row.get("Concurrentie"),
        ])

        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    # Schrijf nieuwe rijen naar Excel
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Sheet1"]
    for r in new_rows:
        ws.append(r)
    wb.save(EXCEL_PATH)

    success = sum(1 for r in new_rows if r[4] is not None)
    log.info(f"Klaar: {success}/{len(new_rows)} prijzen opgehaald — Excel opgeslagen")
    log.info("=" * 65)


if __name__ == "__main__":
    run_scraper()
