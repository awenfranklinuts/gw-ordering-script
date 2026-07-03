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

SCRAPER_PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chrome_scraper_profile")

LOGIN_WAIT_TIMEOUT = 300  # seconds to wait for the user to log in before giving up
LOGIN_POLL_INTERVAL = 3


def _looks_like_login_page(driver):
    url = driver.current_url.lower()
    return "identity.maropost.com" in url or "/cpanel/login" in url


LOGIN_REDIRECT_SETTLE_TIMEOUT = 2  # seconds to allow for a delayed/client-side redirect to the login page
LOGIN_REDIRECT_SETTLE_POLL = 0.5


def ensure_logged_in(driver, sales_url):
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

    print("Not logged in to Neto yet.")
    print("Please log in using the Chrome window that just opened — waiting for you to finish...")

    waited = 0
    while waited < LOGIN_WAIT_TIMEOUT:
        time.sleep(LOGIN_POLL_INTERVAL)
        waited += LOGIN_POLL_INTERVAL
        if not _looks_like_login_page(driver):
            print("Login detected — continuing.")
            driver.get(sales_url)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            return
        print(f"Still waiting for login... ({waited}s elapsed)")

    raise RuntimeError("Timed out waiting for login. Please log in and click Fetch Stock from Neto again.")


def get_last_tuesday():
    """Tuesday of the week *before* the current one (weeks run Mon–Sun), not just the most
    recent past Tuesday. E.g. run on Thursday 2/7 -> current week is Mon 29/6-Sun 5/7, so
    this returns Tuesday of the prior week: 23/6. Stable for every day within a given week."""
    today = date.today()
    monday_this_week = today - timedelta(days=today.weekday())
    tuesday_last_week = monday_this_week + timedelta(days=1) - timedelta(days=7)
    return tuesday_last_week


def build_sales_orders_url(from_date):
    from_date_param = quote(from_date.strftime("%d/%m/%Y") + " 12:00am")
    return (
        "https://www.pcmarket.com.au/_cpanel/orders?"
        "_note_credit_card_warning=0"
        f"&_ftr_dp_fmdate={from_date_param}"
        "&_ftr_dp_todate=&_ftr_id=&_ftr_cus=&_ftr_sku="
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


def scrape_order_lines(driver, order_headers=None):
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
        print(f"Excluded {skipped} order lines from Cancelled / unpaid New orders")
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


def create_driver():
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={SCRAPER_PROFILE_DIR}")
    chrome_options.add_argument("--profile-directory=Default")
    chrome_options.add_argument("--no-sandbox")

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
    return parser.parse_args()


def main():
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

    print("Starting Chrome with scraper profile...")
    print("You can keep your normal Chrome open.")

    if not os.path.exists(SCRAPER_PROFILE_DIR):
        print("First run — please log in to Neto when the browser opens.")
        print("Your session will be saved for future runs.")

    driver = create_driver()

    sales_url = build_sales_orders_url(from_date)
    date_note = "last Tuesday" if from_date == last_tue else "custom date"
    print(f"Filtering sales orders from: {from_date.strftime('%d/%m/%Y')} ({date_note})")

    print("Loading Neto sales orders page...")
    driver.get(sales_url)
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    ensure_logged_in(driver, sales_url)

    if "app.maropost.com" in driver.current_url:
        print("Redirected to Maropost dashboard, navigating back...")
        driver.get(sales_url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    order_headers = scrape_order_headers(driver)
    excluded = sum(1 for h in order_headers.values() if not order_qualifies(h))
    print(f"Found {len(order_headers)} orders on page; {excluded} excluded (Cancelled / New with no payment)")

    order_lines = scrape_order_lines(driver, order_headers)
    print(f"Scraped {len(order_lines)} order lines")

    sku_summary = aggregate_by_sku(order_lines)
    print(f"Aggregated into {len(sku_summary)} unique SKUs")

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sales_order_demand.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(sku_summary, f, indent=2)
    print(f"Saved demand data to {output_path}")

    print("Done — leaving the browser open for review.")

    # Deliberately not calling driver.quit() and not blocking on input() here. This
    # script is launched as a subprocess by the GUI tool with the real Terminal's
    # stdin inherited, so input() would hang forever waiting for a keypress in a
    # window nobody's looking at — the GUI would never see the process finish, and
    # the fetched data would never get applied to the table. The data we care about
    # (sales_order_demand.json) is already written by this point, so it's safe to
    # just exit and leave Chrome open for manual review.


if __name__ == "__main__":
    try:
        main()
    except RuntimeError as e:
        print(f"Error: {e}")
        sys.exit(1)
