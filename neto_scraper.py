from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from collections import defaultdict
from datetime import date, timedelta
from urllib.parse import quote
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


def ensure_logged_in(driver, sales_url):
    """Detect whether the scraper profile is actually logged in to Neto. If the sales
    orders page redirected to the Maropost login screen, pause here and wait for the
    user to log in in the visible browser window, polling until it succeeds or times out.
    """
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
    today = date.today()
    days_since_tuesday = (today.weekday() - 1) % 7
    if days_since_tuesday == 0:
        days_since_tuesday = 7
    return today - timedelta(days=days_since_tuesday)


def build_sales_orders_url():
    last_tue = get_last_tuesday()
    from_date = quote(last_tue.strftime("%d/%m/%Y") + " 12:00am")
    return (
        "https://www.pcmarket.com.au/_cpanel/orders?"
        "_note_credit_card_warning=0"
        f"&_ftr_dp_fmdate={from_date}"
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


def scrape_order_lines(driver):
    rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-order-id][data-qty]")
    order_lines = []
    for row in rows:
        order_id = row.get_attribute("data-order-id")
        qty = int(row.get_attribute("data-qty") or 0)

        sku = ""
        name = ""
        try:
            product_cell = row.find_element(By.CSS_SELECTOR, "td.col-product-name")
            text = product_cell.text
            match = re.match(r'\[(\d+)\]\s*(.*)', text)
            if match:
                sku = match.group(1)
                name = match.group(2).split("MARKETPLACEMAXIMIZER")[0].strip()
        except Exception:
            pass

        stock = 0
        try:
            stock_cell = row.find_element(By.CSS_SELECTOR, "td.col-proc .ntooltip")
            stock_text = stock_cell.text.strip()
            stock = int(re.match(r'-?\d+', stock_text).group())
        except Exception:
            pass

        if sku:
            order_lines.append({
                "order_id": order_id,
                "sku": sku,
                "product_name": name,
                "qty": qty,
                "stock": stock,
            })

    return order_lines


def aggregate_by_sku(order_lines):
    grouped = defaultdict(lambda: {"product_name": "", "total_qty_needed": 0, "stock": 0, "orders": []})
    for line in order_lines:
        entry = grouped[line["sku"]]
        entry["product_name"] = line["product_name"]
        entry["total_qty_needed"] += line["qty"]
        entry["stock"] = line["stock"]
        entry["orders"].append({"order_id": line["order_id"], "qty": line["qty"]})

    result = []
    for sku, data in sorted(grouped.items()):
        result.append({
            "sku": sku,
            "product_name": data["product_name"],
            "total_qty_needed": data["total_qty_needed"],
            "stock": data["stock"],
            "order_count": len(data["orders"]),
        })
    return result


def create_driver():
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={SCRAPER_PROFILE_DIR}")
    chrome_options.add_argument("--profile-directory=Default")
    chrome_options.add_argument("--no-sandbox")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    chromedriver_path = os.path.join(script_dir, "chromedriver-win64", "chromedriver.exe")
    service = Service(executable_path=chromedriver_path)

    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def main():
    print("Starting Chrome with scraper profile...")
    print("You can keep your normal Chrome open.")

    if not os.path.exists(SCRAPER_PROFILE_DIR):
        print("First run — please log in to Neto when the browser opens.")
        print("Your session will be saved for future runs.")

    driver = create_driver()

    sales_url = build_sales_orders_url()
    last_tue = get_last_tuesday()
    print(f"Filtering sales orders from: {last_tue.strftime('%d/%m/%Y')} (last Tuesday)")

    driver.get(sales_url)
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    ensure_logged_in(driver, sales_url)

    if "app.maropost.com" in driver.current_url:
        print("Redirected to Maropost dashboard, navigating back...")
        driver.get(sales_url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    order_lines = scrape_order_lines(driver)
    print(f"Scraped {len(order_lines)} order lines")

    sku_summary = aggregate_by_sku(order_lines)
    print(f"Aggregated into {len(sku_summary)} unique SKUs")

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sales_order_demand.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(sku_summary, f, indent=2)
    print(f"Saved demand data to {output_path}")

    purchase_orders_url = (
        "https://www.pcmarket.com.au/_cpanel/report_purchase_orders?"
        "_ftr_sku=&_ftr_model=&_ftr_restockzero=n&_ftr_restock=y"
        "&_sb_sortby=SKU&_sb_orderby=1&_ftr_mid=GW&_ftrc_mid=70"
        "&_ftr_status=Y&_sb_pgnum=1&_sb_limit=500&_ftr_showcols="
        "&rei40445895=1&rei3236647=1&rei52237633=1&rei10007942=1"
        "&rei9623367=1&rei298557=1&rei17153192=1&rei347193=1&rei298448=1"
        "&rei305639=1&rei298429=1&rei21133135=1&rei426327=1&rei348377=1"
        "&rei434590=1&rei34238381=1&rei298386=1&rei298388=1&rei31158283=1"
        "&rei36160453=1&rei45839514=1&rei48675267=1&rei298510=1&rei298516=1"
        "&rei264470=1&rei298517=1&rei298523=1&rei464733=1&rei12598709=1"
        "&rei15305832=1&rei357515=1&rei26136630=1&rei49598445=1&rei53000102=1"
        "&rei53000103=1&rei53000101=1&rei305697=1&rei299196=1&rei299195=1"
        "&rei302750=1&rei272907=2&rei298495=1&rei41155147=1&rei14104963=1"
        "&rei20428286=1&rei25423350=1&rei403410=1&rei52618791=1&rei52618792=1"
        "&rei4748299=1&rei50658987=1&rei35119528=1&rei19725281=1&rei52618794=1"
        "&rei31700662=1&rei298719=1&rei305607=1&rei48085182=1&rei18199992=1"
        "&rei27209883=1&rei305478=1&rei50713400=1&rei348326=1&rei49000758=1"
        "&rei10706808=1&rei370147=1&rei23260845=1&rei33875820=1&rei15844211=1"
        "&rei35822987=1&rei53000108=1&rei35822989=1&rei298314=6&rei302769=1"
        "&rei303179=1&rei302765=1&rei303128=1&rei303130=1&rei302803=1"
        "&rei302823=1&rei298322=1&rei302777=1&rei302820=1&rei303160=1"
        "&rei303161=1&rei466970=1&rei302865=1&rei303243=1&rei303262=1"
        "&rei303297=1&rei305615=1&rei305616=1&rei305619=1&rei305620=1"
        "&rei305621=1&rei303296=1&rei304900=1&rei52237636=1&itm_total=98"
    )
    driver.execute_script(f"window.open('{purchase_orders_url}', '_blank');")
    driver.switch_to.window(driver.window_handles[1])
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print(f"Tab 2 (Purchase Orders): {driver.title}")
    driver.switch_to.window(driver.window_handles[0])
    print(f"Tab 1 (Sales Orders): {driver.title}")
    driver.switch_to.window(driver.window_handles[0])

    print("Browser is open with 2 tabs.")

    try:
        print("Press Enter to close the browser...")
        input()
    except EOFError:
        pass

    driver.quit()


if __name__ == "__main__":
    try:
        main()
    except RuntimeError as e:
        print(f"Error: {e}")
        sys.exit(1)
