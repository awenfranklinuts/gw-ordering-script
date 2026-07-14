from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from collections import defaultdict
from datetime import date, datetime, timedelta
from urllib.parse import quote
import argparse
import json
import os
import re
import sys
import time

def _base_dir():
    """Directory this app stores its own data in (chrome profile, demand json).

    When frozen by PyInstaller (onefile), __file__ resolves inside the throwaway
    temp folder (sys._MEIPASS) that's re-extracted on every launch — anything
    written there (like the Chrome login profile) would vanish the moment the
    app closes, forcing a fresh login every single run. Using the exe's own
    folder instead keeps that data next to the app and persists across runs."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


SCRAPER_PROFILE_DIR = os.path.join(_base_dir(), "chrome_scraper_profile")

LOGIN_WAIT_TIMEOUT = 300  # seconds to wait for the user to log in before giving up
LOGIN_POLL_INTERVAL = 3


def _default_emit(msg):
    """Fallback progress reporter for standalone CLI use. Never used when the GUI
    calls run() directly with its own on_progress callback — and must not be, since
    a frozen windowed (console=False) exe has no real stdout and print() would raise."""
    try:
        print(msg)
    except Exception:
        pass


def _looks_like_login_page(driver):
    url = driver.current_url.lower()
    return "identity.maropost.com" in url or "/cpanel/login" in url


LOGIN_REDIRECT_SETTLE_TIMEOUT = 2  # seconds to allow for a delayed/client-side redirect to the login page
LOGIN_REDIRECT_SETTLE_POLL = 0.5


def ensure_logged_in(driver, sales_url, emit=_default_emit):
    """Detect whether the scraper profile is actually logged in to Neto. If the sales
    orders page redirected to the Maropost login screen, pause here and wait for the
    user to log in in the visible browser window, polling until it succeeds or times out.

    The redirect to identity.maropost.com isn't always an immediate server response by
    the time Selenium's page-load wait resolves — it can land on an interim page that
    then redirects a moment later. So before concluding "already logged in", briefly
    poll to give a delayed redirect a chance to happen; otherwise we can race past the
    login page entirely and silently scrape zero rows off of it.
    """
    settle_deadline = time.time() + LOGIN_REDIRECT_SETTLE_TIMEOUT
    while time.time() < settle_deadline and not _looks_like_login_page(driver):
        time.sleep(LOGIN_REDIRECT_SETTLE_POLL)

    if not _looks_like_login_page(driver):
        return

    emit("Not logged in to Neto yet.")
    emit("Please log in using the Chrome window that just opened — waiting for you to finish...")

    waited = 0
    while waited < LOGIN_WAIT_TIMEOUT:
        time.sleep(LOGIN_POLL_INTERVAL)
        waited += LOGIN_POLL_INTERVAL
        if not _looks_like_login_page(driver):
            emit("Login detected — continuing.")
            driver.get(sales_url)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            return
        emit(f"Still waiting for login... ({waited}s elapsed)")

    raise RuntimeError("Timed out waiting for login. Please log in and click Fetch Stock from Neto again.")


def get_last_tuesday():
    """Tuesday of the week *before* the current one (weeks run Mon–Sun), not just the most
    recent past Tuesday. E.g. run on Thursday 2/7 -> current week is Mon 29/6-Sun 5/7, so
    this returns Tuesday of the prior week: 23/6. Stable for every day within a given week."""
    today = date.today()
    monday_this_week = today - timedelta(days=today.weekday())
    tuesday_last_week = monday_this_week + timedelta(days=1) - timedelta(days=7)
    return tuesday_last_week


def build_sales_orders_url(from_date, to_date=None):
    from_date_param = quote(from_date.strftime("%d/%m/%Y") + " 12:00am")
    to_date_param = quote(to_date.strftime("%d/%m/%Y") + " 11:59pm") if to_date else ""
    return (
        "https://www.pcmarket.com.au/_cpanel/orders?"
        "_note_credit_card_warning=0"
        f"&_ftr_dp_fmdate={from_date_param}"
        f"&_ftr_dp_todate={to_date_param}&_ftr_id=&_ftr_cus=&_ftr_sku="
        "&_ftr_da_fmdate=&_ftr_da_todate=&_ftr_di_fmdate=&_ftr_di_todate="
        "&_ftr_du_fmdate=&_ftr_du_todate=&_ftr_dc_fmdate=&_ftr_dc_todate="
        "&_ftr_dr_fmdate=&_ftr_dr_todate=&_ftr_dd_fmdate=&_ftr_dd_todate="
        "&_ftrc_alp=2&_ftrc_type=4&_ftrc_wh=3&_ftrc_status=12&_ftrc_pt=20"
        "&_ftrc_label=1&_ftrc_shgp=18&_ftr_sh="
        "&_ftrc_ebayst=9&_ftrc_ebaytype=5&_ftr_ebayusr=&_ftr_ebayrec="
        "&_ftr_ebay=&_ftr_ebaytid=&_ftr_cmbaddr=&_ftr_noln=&_ftr_po="
        "&_ftr_salch=&_ftr_bpay=&_ftr_model=&_ftr_snp=&_ftr_sei="
        "&_ftr_allol=n&_ftr_tolc_fm=&_ftr_tolc_to=&_ftr_sales=&_ftr_mgr="
        "&_ftr_opt=&_ftr_shexp=&_ftr_shpty=&_ftr_pk=&_ftr_pack=&_ftr_bko="
        "&_ftr_paid=&_ftr_addr=&_ftr_state=&_ftr_email="
        "&_ftr_sup=Games+Workshop"
        "&_ftr_ucc=&_ftr_pbsz=&_ftr_bin=&_ftr_shpk=&_ftr_par=&_ftr_quote="
        "&_ftr_spkzone=&_ftrc_pkzone=5&_ftr_upc=&_ftr_exprted=&_ftr_lnkid="
        "&_ftr_bp=&_ftr_cproc=&_ftr_last=&_ftr_shpickup=&_ftr_international="
        "&_ftr_tax=&_ftr_misc3=&_ftr_misc4=&_ftr_misc2=&_ftr_misc1="
        "&_sb_sortby=date_placed&_sb_orderby=2&_ftr_allst="
        "&_sb_pgnum=1&_sb_limit=500"
    )


def _parse_money(text):
    """Parse '$420.20 AUD' -> 420.20. Returns None if no number found."""
    match = re.search(r'-?[\d,]+(?:\.\d+)?', text or "")
    if not match:
        return None
    return float(match.group().replace(",", ""))


def scrape_order_headers(driver):
    """Build order_id -> {status, order_total, amount_owed} from the order header
    card above each order's line items. The sticky column-label bar also uses
    table.order-header-table, so require the a.oid order-id link to identify real
    order cards."""
    headers = {}
    for table in driver.find_elements(By.CSS_SELECTOR, "div.order-header table.order-header-table"):
        try:
            order_id = table.find_element(By.CSS_SELECTOR, "a.oid").text.strip()
        except Exception:
            continue

        status = ""
        try:
            status = table.find_element(By.CSS_SELECTOR, "td.col-status a").text.strip()
        except Exception:
            pass

        order_total = amount_owed = None
        try:
            order_total = _parse_money(table.find_element(By.CSS_SELECTOR, "td.col-order-total").text)
        except Exception:
            pass
        try:
            amount_owed = _parse_money(table.find_element(By.CSS_SELECTOR, "td.col-amount-owed").text)
        except Exception:
            pass

        headers[order_id] = {
            "status": status,
            "order_total": order_total,
            "amount_owed": amount_owed,
        }
    return headers


def order_qualifies(header):
    """Whether an order's line quantities should count toward demand.

    Every status counts EXCEPT:
    - Cancelled orders — always excluded.
    - "New" orders with no payment made: a New order only counts once some
      payment has been received (amount owed < order total). A New order with
      nothing paid (owed == total, or unparseable amounts) is excluded."""
    status = (header.get("status") or "").strip().lower()
    if status == "cancelled":
        return False
    if status == "new":
        total = header.get("order_total")
        owed = header.get("amount_owed")
        if total is None or owed is None:
            return False
        return owed < total
    return True


def scrape_order_lines(driver, order_headers=None, emit=_default_emit):
    """Scrape line items. If order_headers is provided, lines belonging to
    non-qualifying orders (see order_qualifies) are skipped; lines whose order id
    is missing from order_headers are kept, since only unpaid New orders are
    excluded and an unknown header can't prove that."""
    rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-order-id][data-qty]")
    order_lines = []
    skipped = 0
    for row in rows:
        order_id = row.get_attribute("data-order-id")

        if order_headers is not None:
            header = order_headers.get(order_id)
            if header is not None and not order_qualifies(header):
                skipped += 1
                continue
        qty = int(row.get_attribute("data-qty") or 0)

        sku = ""
        name = ""
        try:
            product_cell = row.find_element(By.CSS_SELECTOR, "td.col-product-name")
            text = product_cell.text
            # SKUs aren't always all-digit (e.g. "[ZJG01493]"), so match any
            # alphanumeric/hyphenated code in the leading brackets.
            match = re.match(r'\[([A-Za-z0-9\-]+)\]\s*(.*)', text)
            if match:
                sku = match.group(1)
                name = match.group(2).split("MARKETPLACEMAXIMIZER")[0].strip()
        except Exception:
            pass

        # The "Stock" cell shows one or two figures, identified by tooltip title
        # rather than tag/position (more resilient to markup tweaks):
        #   - "Total Stock On Hand (taking into account this orderline)" -> stock
        #     free to sell to OTHER customers once this order is fulfilled. 0 means
        #     we have no reserve/spare stock left to sell.
        #   - "Stock in PC Market" (in parens) -> physical qty currently on hand,
        #     regardless of this order. E.g. "Stock: 0 (1)" means we physically have
        #     1 unit, but it's already spoken for — pending dispatch to this existing
        #     order — so there's 0 left over to sell elsewhere.
        # This second figure is only rendered by Neto when it *differs* from the
        # first (i.e. something's reserved elsewhere) — e.g. "Stock: 1" with no
        # parenthesized number means nothing is pending, so on-hand == available.
        stock_available_to_sell = 0
        stock_on_hand = None
        try:
            available_cell = row.find_element(
                By.CSS_SELECTOR,
                'td.col-proc [data-original-title="Total Stock On Hand (taking into account this orderline)"]',
            )
            match = re.search(r'-?\d+', available_cell.text)
            if match:
                stock_available_to_sell = int(match.group())
        except Exception:
            pass
        try:
            on_hand_cell = row.find_element(
                By.CSS_SELECTOR, 'td.col-proc [data-original-title="Stock in PC Market"]'
            )
            match = re.search(r'-?\d+', on_hand_cell.text)
            if match:
                stock_on_hand = int(match.group())
        except Exception:
            pass
        if stock_on_hand is None:
            stock_on_hand = stock_available_to_sell

        if sku:
            order_lines.append({
                "order_id": order_id,
                "sku": sku,
                "product_name": name,
                "qty": qty,
                "stock": stock_available_to_sell,  # kept for backward compatibility
                "stock_available_to_sell": stock_available_to_sell,
                "stock_on_hand": stock_on_hand,
            })

    if order_headers is not None and skipped:
        emit(f"Excluded {skipped} order lines from Cancelled / unpaid New orders")
    return order_lines


def aggregate_by_sku(order_lines):
    grouped = defaultdict(lambda: {
        "product_name": "", "total_qty_needed": 0,
        "stock_available_to_sell": 0, "stock_on_hand": 0, "orders": [],
    })
    for line in order_lines:
        entry = grouped[line["sku"]]
        entry["product_name"] = line["product_name"]
        entry["total_qty_needed"] += line["qty"]
        entry["stock_available_to_sell"] = line["stock_available_to_sell"]
        entry["stock_on_hand"] = line["stock_on_hand"]
        entry["orders"].append({"order_id": line["order_id"], "qty": line["qty"]})

    result = []
    for sku, data in sorted(grouped.items()):
        result.append({
            "sku": sku,
            "product_name": data["product_name"],
            "total_qty_needed": data["total_qty_needed"],
            "stock": data["stock_available_to_sell"],  # kept for backward compatibility
            "stock_available_to_sell": data["stock_available_to_sell"],
            "stock_on_hand": data["stock_on_hand"],
            "order_count": len(data["orders"]),
        })
    return result


def create_driver(profile_dir=None):
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={profile_dir or SCRAPER_PROFILE_DIR}")
    chrome_options.add_argument("--profile-directory=Default")
    chrome_options.add_argument("--no-sandbox")

    # No executable_path/Service is passed here on purpose: Selenium Manager (built
    # into Selenium 4.6+) auto-detects the installed Chrome and downloads a matching
    # chromedriver itself. That's what lets this run on any device with Chrome
    # installed, with no manual chromedriver setup or version-matching required.
    driver = webdriver.Chrome(options=chrome_options)
    return driver


def parse_args():
    parser = argparse.ArgumentParser(description="Scrape Neto sales orders demand data.")
    parser.add_argument(
        "--from-date",
        dest="from_date",
        default=None,
        help="Date Placed From, in DD/MM/YYYY format (default: last Tuesday).",
    )
    parser.add_argument(
        "--to-date",
        dest="to_date",
        default=None,
        help="Date Placed Till, in DD/MM/YYYY format (default: no upper bound).",
    )
    return parser.parse_args()


def run(from_date=None, to_date=None, on_progress=None):
    """Run the full Neto scrape and return the per-SKU demand summary.

    to_date, if given, caps the Date Placed filter's upper end (inclusive, end of
    day) — otherwise Neto returns orders up to now with no upper bound.

    on_progress, if given, is called with each progress line instead of print() —
    this is what lets the GUI import this module directly and call run() in a
    background thread (see gw_order_tool.py's _run_neto_scraper), rather than
    launching this file as a subprocess via sys.executable. That subprocess
    approach broke once packaged into a frozen exe, since sys.executable then
    points at the exe itself rather than a Python interpreter. Calling run()
    in-process also means print() is never hit in a frozen windowed (console=False)
    build, where sys.stdout is None and print() would raise.
    """
    emit = on_progress or _default_emit

    last_tue = get_last_tuesday()
    resolved_from_date = from_date or last_tue

    emit("Starting Chrome with scraper profile...")
    emit("You can keep your normal Chrome open.")

    base_dir = _base_dir()
    profile_dir = os.path.join(base_dir, "chrome_scraper_profile")
    if not os.path.exists(profile_dir):
        emit("First run — please log in to Neto when the browser opens.")
        emit("Your session will be saved for future runs.")

    driver = create_driver(profile_dir)

    sales_url = build_sales_orders_url(resolved_from_date, to_date)
    date_note = "last Tuesday" if resolved_from_date == last_tue else "custom date"
    range_desc = f"{resolved_from_date.strftime('%d/%m/%Y')} ({date_note})"
    if to_date:
        range_desc += f" to {to_date.strftime('%d/%m/%Y')}"
    emit(f"Filtering sales orders from: {range_desc}")

    emit("Loading Neto sales orders page...")
    driver.get(sales_url)
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    ensure_logged_in(driver, sales_url, emit)

    if "app.maropost.com" in driver.current_url:
        emit("Redirected to Maropost dashboard, navigating back...")
        driver.get(sales_url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    order_headers = scrape_order_headers(driver)
    excluded = sum(1 for h in order_headers.values() if not order_qualifies(h))
    emit(f"Found {len(order_headers)} orders on page; {excluded} excluded (Cancelled / New with no payment)")

    order_lines = scrape_order_lines(driver, order_headers, emit)
    emit(f"Scraped {len(order_lines)} order lines")

    sku_summary = aggregate_by_sku(order_lines)
    emit(f"Aggregated into {len(sku_summary)} unique SKUs")

    output_path = os.path.join(base_dir, "sales_order_demand.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(sku_summary, f, indent=2)
    emit(f"Saved demand data to {output_path}")

    emit("Done — leaving the browser open for review.")

    # Deliberately not calling driver.quit() here — leaving Chrome open lets the
    # user glance over the scraped orders themselves if something looks off.
    return sku_summary


def main():
    """Standalone CLI entry point — useful for testing this file on its own
    (e.g. `python neto_scraper.py --from-date 01/01/2026`). The GUI does not use
    this; it imports run() directly instead."""
    args = parse_args()

    last_tue = get_last_tuesday()
    if args.from_date:
        try:
            from_date = datetime.strptime(args.from_date, "%d/%m/%Y").date()
        except ValueError:
            print(f"Invalid --from-date '{args.from_date}' (expected DD/MM/YYYY) — using last Tuesday instead.")
            from_date = last_tue
    else:
        from_date = last_tue

    to_date = None
    if args.to_date:
        try:
            to_date = datetime.strptime(args.to_date, "%d/%m/%Y").date()
        except ValueError:
            print(f"Invalid --to-date '{args.to_date}' (expected DD/MM/YYYY) — ignoring, no upper bound.")
            to_date = None

    run(from_date=from_date, to_date=to_date, on_progress=print)


if __name__ == "__main__":
    try:
        main()
    except RuntimeError as e:
        print(f"Error: {e}")
        sys.exit(1)
