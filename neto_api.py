"""Neto / Maropost Commerce Cloud API client — a drop-in alternative to neto_scraper.py.

Exposes the same run(from_date, to_date, on_progress) signature and writes the same
per-SKU demand JSON, so gw_order_tool.py can swap between this and the Selenium
scraper by changing one import.

Why this exists: the scraper drives a real Chrome window with a persistent login
profile, a bundled chromedriver, and CSS selectors against Neto's control panel
HTML — all of which break when Neto tweaks its markup, when Chrome auto-updates
past the bundled driver, or when the user's session expires. The API returns the
same data as structured JSON with a static key and no browser at all.

Deliberately stdlib-only (urllib, not requests) so this adds no new dependency to
the PyInstaller build, and imports no selenium so the API path works even on a
machine with no Chrome installed.

Credentials (never commit these):
  1. Environment variables NETO_API_KEY (and NETO_USERNAME for a user-based key), or
  2. neto_config.json next to this file / the exe:
         {"api_key": "...", "username": "...", "store_url": "https://www.pcmarket.com.au"}

"username" is the Neto staff username the key belongs to. It's required for a
user-based key (Staff User Manager) and ignored for the global key — sending it
when it isn't needed is harmless, omitting it when it is needed returns the
unhelpfully generic "Invalid API Key".

Get the key from Neto: Setup & Tools > API Settings. Prefer a *per-staff-user* key
created in Staff User Manager, scoped by permission group to read Orders and
Products only — the global key on the API Settings page is likely already in use
by the Shippit webhook integration, and regenerating it would break that.
"""

import argparse
import json
import os
import sys
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import date, datetime, timedelta

# Supplier codes (the item PrimarySupplier field) for the products this tool
# orders, replacing the scraper's &_ftr_sup=Games+Workshop URL filter.
#
# These are codes, not display names — "Games Workshop" is the Brand; PrimarySupplier
# holds "GW"/"GWD". Both are needed, as the scraper's results contained a mix.
# Re-derive them with `--diagnose --compare <baseline>` if they ever change.
#
# Applied client-side, NOT as a GetOrder filter. GetOrder does accept a Supplier
# filter, but it returns zero sales orders for any value — including the codes that
# demonstrably appear on the ordered items — which suggests it only applies to
# purchase orders, where the supplier belongs to the order itself rather than to
# the products on its lines. So we fetch every order in the date range and drop
# non-GW lines after resolving each SKU's supplier via GetItem, which is a call we
# already make for stock levels anyway.
SUPPLIERS = ["GW", "GWD"]

# Orders per GetOrder request. Neto's docs don't publish a hard ceiling, but large
# Limits with OrderLine in the OutputSelector get slow and can time out; 200 is a
# comfortable middle ground. Stock lookups are chunked separately (see SKU_CHUNK).
ORDER_PAGE_LIMIT = 200

# Safety stop for the pagination loop, which otherwise ends only when a page adds
# no new orders. At ORDER_PAGE_LIMIT per page this allows far more orders than any
# realistic weekly window, so hitting it means something is wrong (e.g. Page being
# ignored and the same rows returned forever) rather than a genuinely huge result.
MAX_PAGES = 50

# SKUs per GetItem request. Keeps the POST body and response to a sane size when a
# week's orders span several hundred distinct products.
SKU_CHUNK = 100

# Which warehouse's on-hand figure counts as "stock on hand". None = sum every
# warehouse. Set this to an integer WarehouseID to mirror the scraper's
# &_ftrc_wh=3 URL filter once you've confirmed which ID that code refers to —
# see the "unverified" notes in the module docstring of the CLI --compare mode.
WAREHOUSE_ID = None

# Hours to add to an API DatePlaced to get store-local time.
#
# The API returns and filters DatePlaced in UTC, while the control panel — and
# therefore the scraper, and therefore what staff mean by "orders placed Tuesday" —
# shows Brisbane time. Ignoring this silently drops orders placed between local
# midnight and 10am on the first day of the window, which is how a Tuesday 07:26
# order came back dated Monday 21:26 and vanished from the results.
#
# Australia/Brisbane is UTC+10 year-round and observes no daylight saving, so a
# fixed offset is exact here. A store in Sydney or Melbourne would need a real
# timezone (zoneinfo) instead, since their offset shifts with DST.
STORE_UTC_OFFSET_HOURS = 10

HTTP_TIMEOUT = 60  # seconds per API call

DEFAULT_STORE_URL = "https://www.pcmarket.com.au"
CONFIG_FILENAME = "neto_config.json"
OUTPUT_FILENAME = "sales_order_demand.json"


class NetoAPIError(RuntimeError):
    """Raised for anything the user can act on: missing key, auth failure, API
    returning Ack=Error. Subclasses RuntimeError so gw_order_tool.py's existing
    `except RuntimeError` branch — which shows a friendly dialog rather than a
    crash — catches these exactly like the scraper's login-timeout error."""


def _base_dir():
    """Directory this app stores its own data in (config, demand json).

    Same logic as neto_scraper._base_dir(): when frozen by PyInstaller (onefile),
    __file__ resolves inside the throwaway sys._MEIPASS temp folder that's
    re-extracted every launch, so config written there would vanish. Use the exe's
    own folder instead so neto_config.json persists next to the app."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _default_emit(msg):
    """Fallback progress reporter for standalone CLI use. Never used when the GUI
    calls run() with its own on_progress callback — and must not be, since a frozen
    windowed (console=False) exe has no real stdout and print() would raise."""
    try:
        print(msg)
    except Exception:
        pass


def load_credentials(api_key=None, username=None, store_url=None):
    """Resolve credentials, with explicit arguments taking priority.

    The arguments exist for the CLI's --key/--username overrides: working out which
    key style a store accepts means trying several combinations, and editing
    neto_config.json between every attempt is both tedious and a good way to leave
    a half-edited file behind."""
    return _load_credentials(api_key, username, store_url)


def _load_credentials(arg_key=None, arg_username=None, arg_store_url=None):
    """Resolve (store_url, api_key, username) from the environment or neto_config.json.

    Environment wins so CI or a quick shell test can override the file without
    editing it. Raises NetoAPIError with setup instructions rather than returning
    None, so a missing key surfaces as a readable dialog instead of a 401 later."""
    # `is not None` rather than a plain `or`: an explicit empty username is a
    # meaningful value ("send no NETOAPI_USERNAME header, i.e. global-key style"),
    # and `or` would treat it as unset and silently fall back to the config file —
    # making it impossible to test the global style while a username is configured.
    api_key = arg_key if arg_key is not None else os.environ.get("NETO_API_KEY")
    username = arg_username if arg_username is not None else os.environ.get("NETO_USERNAME")
    store_url = arg_store_url if arg_store_url is not None else os.environ.get("NETO_STORE_URL")

    config_path = os.path.join(_base_dir(), CONFIG_FILENAME)
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
        except (OSError, ValueError) as e:
            raise NetoAPIError(f"Couldn't read {CONFIG_FILENAME}: {e}")
        api_key = api_key if api_key is not None else config.get("api_key")
        username = username if username is not None else config.get("username")
        store_url = store_url if store_url is not None else config.get("store_url")

    if not api_key:
        raise NetoAPIError(
            "No Neto API key found.\n\n"
            f"Create {CONFIG_FILENAME} next to this app containing:\n"
            '  {"api_key": "YOUR_KEY", "username": "YOUR_STAFF_USERNAME",\n'
            '   "store_url": "https://www.pcmarket.com.au"}\n\n'
            "or set the NETO_API_KEY environment variable. Get a key from Neto under "
            "Setup & Tools > API Settings (ideally a per-user key from Staff User Manager)."
        )

    # Whitespace from copy/pasting the key out of the control panel is invisible in
    # a text editor but makes Neto reject the key outright.
    return (store_url or DEFAULT_STORE_URL).rstrip("/"), api_key.strip(), (username or "").strip()


def call_api(action, payload, credentials=None):
    """POST one action to the Neto API and return the decoded JSON response.

    Every Neto call is a POST to the same /do/WS/NetoAPI endpoint; the action is
    carried in the NETOAPI_ACTION header rather than the path or body."""
    store_url, api_key, username = credentials or load_credentials()
    endpoint = f"{store_url}/do/WS/NetoAPI"

    headers = {
        "NETOAPI_ACTION": action,
        "NETOAPI_KEY": api_key,
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    # NETOAPI_USERNAME is required alongside the key for user-based (Staff User
    # Manager) keys and ignored for the global key. Only sent when configured,
    # since an empty header value is worse than no header at all.
    if username:
        headers["NETOAPI_USERNAME"] = username

    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=HTTP_TIMEOUT) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise NetoAPIError(
                f"Neto rejected the API key ({e.code}). Check the key is correct and that "
                "its permission group allows reading Orders and Products."
            )
        raise NetoAPIError(f"Neto API returned HTTP {e.code} for {action}: {e.reason}")
    except urllib.error.URLError as e:
        raise NetoAPIError(f"Couldn't reach {endpoint}: {e.reason}")

    try:
        data = json.loads(body)
    except ValueError:
        # Neto serves an HTML error/login page on some misconfigurations rather
        # than JSON — surface a hint instead of a bare JSONDecodeError.
        raise NetoAPIError(
            f"Neto returned a non-JSON response for {action}. "
            f"Check store_url is correct. First 200 chars: {body[:200]!r}"
        )

    # Neto reports application-level failures with HTTP 200 + Ack=Error, so the
    # HTTPError branch above is not enough on its own.
    if str(data.get("Ack", "")).lower() == "error":
        messages = data.get("Messages", {})
        raw = json.dumps(messages)
        if "invalid api key" in raw.lower():
            # Neto uses this one message for several distinct causes, so spell them
            # out rather than making the user guess from three words.
            raise NetoAPIError(
                "Neto rejected the API key.\n\n"
                "Most likely causes, in order:\n"
                f"  1. The key is a user-based (Staff User Manager) key but no username was "
                f"configured — add \"username\" to {CONFIG_FILENAME}.\n"
                f"  2. The key was copied short. The API Settings field is narrower than the "
                f"key, so selecting the visible text truncates it — click into the field and "
                f"use Cmd+A, then Cmd+C. Your configured key is {len(api_key)} characters.\n"
                "  3. The staff user's permission group doesn't grant API access.\n"
                f"  4. store_url is wrong (currently {store_url}).\n\n"
                "Run `python3 neto_api.py --check` to retest after fixing."
            )
        raise NetoAPIError(f"Neto API error on {action}: {raw}")

    return data


def _as_list(value):
    """Neto collapses single-element collections to a bare object in JSON (an order
    with one line gives OrderLine={...}, not [{...}]), so every repeated field has
    to be normalised before iterating or single-line orders raise/ silently skip."""
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def _to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _to_int(value):
    """Quantities arrive as decimal strings ("2.000") — int() rejects those, so
    round-trip through float first."""
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def _local_window(from_date, to_date=None):
    """Local-time window bounds as naive datetimes: [from 00:00, to 23:59:59]."""
    start = datetime.combine(from_date, datetime.min.time())
    end = datetime.combine(to_date, datetime.max.time()) if to_date else None
    return start, end


def _to_api_datetime(local_dt):
    """Convert a store-local datetime to the UTC string the API filters on."""
    return (local_dt - timedelta(hours=STORE_UTC_OFFSET_HOURS)).strftime("%Y-%m-%d %H:%M:%S")


def _from_api_datetime(value):
    """Parse an API UTC timestamp into store-local time. None if unparseable."""
    try:
        return datetime.strptime(str(value), "%Y-%m-%d %H:%M:%S") + timedelta(
            hours=STORE_UTC_OFFSET_HOURS)
    except (TypeError, ValueError):
        return None


def get_last_tuesday():
    """Tuesday of the week *before* the current one (weeks run Mon-Sun), not just
    the most recent past Tuesday. Mirrors neto_scraper.get_last_tuesday() exactly so
    both paths default to the same date."""
    today = date.today()
    monday_this_week = today - timedelta(days=today.weekday())
    return monday_this_week + timedelta(days=1) - timedelta(days=7)


def fetch_orders(from_date, to_date=None, credentials=None, emit=_default_emit):
    """Page through GetOrder for the supplier's orders in the date range.

    Replaces the scraper's hand-built control-panel URL: DatePlacedFrom/To stand in
    for the _ftr_dp_* params, Supplier for _ftr_sup, and Page/Limit for _sb_pgnum /
    _sb_limit."""
    output_selectors = [
        "OrderID",
        "OrderStatus",
        "GrandTotal",
        "OrderPayment",
        "DatePlaced",
        "OrderLine",
        "OrderLine.ProductName",
        "OrderLine.Quantity",
    ]

    # The window the user means is local; the API filters in UTC.
    local_start, local_end = _local_window(from_date, to_date)
    date_from = _to_api_datetime(local_start)
    date_to = _to_api_datetime(local_end) if local_end else None

    orders = []
    seen_ids = set()
    page = 0
    while True:
        # Date only — see the SUPPLIERS comment for why the supplier is filtered
        # client-side instead of here.
        order_filter = {
            "DatePlacedFrom": date_from,
            "Page": page,
            "Limit": ORDER_PAGE_LIMIT,
            "OutputSelector": output_selectors,
        }
        if date_to:
            order_filter["DatePlacedTo"] = date_to

        response = call_api("GetOrder", {"Filter": order_filter}, credentials)
        batch = _as_list(response.get("Order"))
        if not batch:
            break

        # Dedupe by OrderID: Neto's docs don't state whether Page is 0- or
        # 1-indexed, so if it turns out to be 1-based our page 0 and page 1 return
        # the same rows. Deduping makes the loop correct either way instead of
        # silently doubling every quantity in the first batch.
        new_in_batch = 0
        for order in batch:
            order_id = order.get("OrderID")
            if order_id in seen_ids:
                continue
            seen_ids.add(order_id)
            orders.append(order)
            new_in_batch += 1

        emit(f"Fetched {len(orders)} orders...")

        # Stop only on a batch that yields nothing new — NOT on "fewer rows than
        # Limit". Neto can return fewer than the requested Limit while more pages
        # still exist, so treating a short page as the last one silently truncates
        # the result, and a truncated order set means under-ordering stock.
        if new_in_batch == 0:
            break
        page += 1
        if page > MAX_PAGES:
            emit(f"Warning: stopped at {MAX_PAGES} pages — results may be incomplete.")
            break

    return orders


def amount_owed(order):
    """GrandTotal minus everything paid against the order.

    The control panel shows this as its own "Amount Owed" column; the API has no
    equivalent field, so it's derived from the OrderPayment records. Returns None
    when GrandTotal is unparseable, which order_qualifies() treats as "can't prove
    payment" — matching the scraper's behaviour on an unreadable money cell."""
    total = _to_float(order.get("GrandTotal"))
    if total is None:
        return None
    paid = sum(_to_float(p.get("Amount")) or 0.0 for p in _as_list(order.get("OrderPayment")))
    return total - paid


def order_qualifies(order):
    """Whether an order's line quantities should count toward demand.

    Same rule as neto_scraper.order_qualifies, restated against API fields: every
    status counts EXCEPT Cancelled (always excluded) and New orders with no payment
    received (a New order only counts once amount owed < order total)."""
    status = str(order.get("OrderStatus") or "").strip().lower()
    if status == "cancelled":
        return False
    if status == "new":
        total = _to_float(order.get("GrandTotal"))
        owed = amount_owed(order)
        if total is None or owed is None:
            return False
        return owed < total
    return True


def build_order_lines(orders, emit=_default_emit):
    """Flatten qualifying orders into per-line dicts.

    Note there's no SKU regex here: the scraper had to pull "[ZJG01493] Name" apart
    out of a rendered table cell and strip a MARKETPLACEMAXIMIZER suffix, whereas
    the API returns SKU and ProductName as separate fields already."""
    order_lines = []
    skipped_orders = 0
    for order in orders:
        if not order_qualifies(order):
            skipped_orders += 1
            continue
        for line in _as_list(order.get("OrderLine")):
            sku = (line.get("SKU") or "").strip()
            if not sku:
                continue
            order_lines.append({
                "order_id": order.get("OrderID"),
                "sku": sku,
                "product_name": (line.get("ProductName") or "").strip(),
                "qty": _to_int(line.get("Quantity")),
            })

    if skipped_orders:
        emit(f"Excluded {skipped_orders} Cancelled / unpaid New orders")
    return order_lines


def fetch_item_info(skus, credentials=None, emit=_default_emit):
    """Look up stock and supplier for the given SKUs -> {sku: {available, on_hand, supplier}}.

    Order lines carry no stock figures, so this is a second call. Mapping to the
    scraper's two numbers:
      - stock_available_to_sell  <- AvailableSellQuantity
      - stock_on_hand            <- sum of WarehouseQuantity (or WAREHOUSE_ID only)

    Caveat worth verifying before trusting this: the scraper's "available" figure
    came from a tooltip reading "Total Stock On Hand (taking into account this
    orderline)" — i.e. computed per order line — while AvailableSellQuantity is a
    single per-item figure. Since aggregate_by_sku only ever kept the last line's
    value per SKU anyway, the API figure is arguably more correct, but it will not
    always match the old output number-for-number."""
    info = {}
    sku_list = sorted(set(skus))
    for start in range(0, len(sku_list), SKU_CHUNK):
        chunk = sku_list[start:start + SKU_CHUNK]
        response = call_api("GetItem", {
            "Filter": {
                "SKU": chunk,
                "OutputSelector": [
                    "SKU", "AvailableSellQuantity", "WarehouseQuantity", "PrimarySupplier",
                ],
            }
        }, credentials)

        for item in _as_list(response.get("Item")):
            sku = (item.get("SKU") or "").strip()
            if not sku:
                continue
            warehouse_rows = _as_list(item.get("WarehouseQuantity"))
            if WAREHOUSE_ID is not None:
                warehouse_rows = [
                    w for w in warehouse_rows if _to_int(w.get("WarehouseID")) == WAREHOUSE_ID
                ]
            info[sku] = {
                "available": _to_int(item.get("AvailableSellQuantity")),
                "on_hand": sum(_to_int(w.get("Quantity")) for w in warehouse_rows),
                "supplier": (item.get("PrimarySupplier") or "").strip(),
            }

        emit(f"Looked up {len(info)}/{len(sku_list)} products...")

    missing = [s for s in sku_list if s not in info]
    if missing:
        # SKUs on orders with no matching item record (deleted/renamed products).
        # Reported rather than silently defaulted, since it usually means the SKU
        # changed in Neto and the order pad row will look wrong.
        emit(f"Warning: no item record found for {len(missing)} SKUs, e.g. {', '.join(missing[:5])}")

    return info


def fetch_stock_for_skus(skus, on_progress=None):
    """Public helper: current stock for arbitrary SKUs -> {sku: (available, on_hand)}.

    run() only reports SKUs that actually sold, so the GUI has no stock figure for
    the (usually far larger) set of products with no orders this week. This looks up
    any SKU on demand, which is what lets the Sellable Stock column be populated for
    every row rather than only the ones with sales.

    Returns the pair shape gw_order_tool's neto_stock_lookup already uses, so
    results can be merged straight in. SKUs with no item record in Neto are simply
    absent from the result rather than defaulted to zero — a missing product and one
    genuinely out of stock are different things, and showing "0" for the former
    would invite ordering against a SKU that no longer exists."""
    emit = on_progress or _default_emit
    info = fetch_item_info(skus, None, emit)
    return {sku: (rec["available"], rec["on_hand"]) for sku, rec in info.items()}


def filter_to_suppliers(order_lines, item_info, emit=_default_emit):
    """Keep only lines whose product belongs to one of SUPPLIERS.

    This is the client-side replacement for the GetOrder Supplier filter. Lines
    whose SKU has no item record are dropped: without a supplier we can't prove the
    product is one this tool orders, and wrongly including a non-GW line would put
    a bogus row on the order pad."""
    wanted = {s.strip().lower() for s in SUPPLIERS}
    kept, dropped_supplier, dropped_unknown = [], 0, 0
    for line in order_lines:
        record = item_info.get(line["sku"])
        if record is None:
            dropped_unknown += 1
            continue
        if record["supplier"].lower() not in wanted:
            dropped_supplier += 1
            continue
        kept.append(line)

    emit(f"Kept {len(kept)} {'/'.join(SUPPLIERS)} lines "
         f"({dropped_supplier} other suppliers, {dropped_unknown} unknown SKUs)")
    return kept


def aggregate_by_sku(order_lines, item_info):
    """Group lines into the per-SKU summary the GUI consumes.

    Output keys are identical to neto_scraper.aggregate_by_sku's, including the
    redundant "stock" key kept there for backward compatibility, so
    _load_and_apply_neto_stock_data needs no changes."""
    grouped = defaultdict(lambda: {"product_name": "", "total_qty_needed": 0, "order_count": 0})
    for line in order_lines:
        entry = grouped[line["sku"]]
        entry["product_name"] = line["product_name"]
        entry["total_qty_needed"] += line["qty"]
        entry["order_count"] += 1

    result = []
    for sku, data in sorted(grouped.items()):
        levels = item_info.get(sku, {"available": 0, "on_hand": 0})
        result.append({
            "sku": sku,
            "product_name": data["product_name"],
            "total_qty_needed": data["total_qty_needed"],
            "stock": levels["available"],  # kept for backward compatibility
            "stock_available_to_sell": levels["available"],
            "stock_on_hand": levels["on_hand"],
            "order_count": data["order_count"],
        })
    return result


def collect(from_date=None, to_date=None, on_progress=None):
    """Do the whole fetch and return the intermediate stages as well as the summary.

    run() only needs the summary, but --compare needs the pre-supplier-filter lines
    to explain why a SKU is missing: "never appeared in any order in this window"
    and "appeared but belongs to another supplier" are very different problems, and
    they're indistinguishable from the final output alone."""
    emit = on_progress or _default_emit

    last_tue = get_last_tuesday()
    resolved_from_date = from_date or last_tue
    credentials = load_credentials()

    date_note = "last Tuesday" if resolved_from_date == last_tue else "custom date"
    range_desc = f"{resolved_from_date.strftime('%d/%m/%Y')} ({date_note})"
    if to_date:
        range_desc += f" to {to_date.strftime('%d/%m/%Y')}"
    emit(f"Fetching orders placed from: {range_desc}")

    orders = fetch_orders(resolved_from_date, to_date, credentials, emit)
    excluded = sum(1 for o in orders if not order_qualifies(o))
    emit(f"Found {len(orders)} orders; {excluded} excluded (Cancelled / New with no payment)")

    all_lines = build_order_lines(orders, emit)
    emit(f"Collected {len(all_lines)} order lines across all suppliers")

    # Every SKU in the window has to be looked up, not just the GW ones, because
    # the supplier is exactly what we're trying to learn.
    item_info = fetch_item_info({line["sku"] for line in all_lines}, credentials, emit)
    kept_lines = filter_to_suppliers(all_lines, item_info, emit)

    summary = aggregate_by_sku(kept_lines, item_info)
    emit(f"Aggregated into {len(summary)} unique SKUs")

    return {
        "summary": summary,
        "all_lines": all_lines,
        "item_info": item_info,
        "credentials": credentials,
        "orders": orders,
        "from_date": resolved_from_date,
        "to_date": to_date,
    }


def run(from_date=None, to_date=None, on_progress=None, output_filename=OUTPUT_FILENAME):
    """Fetch Neto demand via the API and return the per-SKU summary.

    Signature matches neto_scraper.run() so gw_order_tool.py can call either.
    output_filename is the one addition, letting the CLI write somewhere else while
    testing so a comparison run can't clobber the real sales_order_demand.json."""
    emit = on_progress or _default_emit
    sku_summary = collect(from_date, to_date, emit)["summary"]

    output_path = os.path.join(_base_dir(), output_filename)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(sku_summary, f, indent=2)
    emit(f"Saved demand data to {output_path}")

    emit("Done.")
    return sku_summary


def _explain_missing(missing_skus, details):
    """Say why each baseline SKU didn't survive the API run.

    A raw list of missing SKUs isn't actionable — the fix is completely different
    depending on whether the product has a supplier code we're not asking for, or
    simply had no qualifying order in the window. This buckets them so the next
    step is obvious."""
    if not details:
        print("  (run with --compare during a live fetch to explain these)")
        return

    item_info = details["item_info"]
    ordered_skus = {line["sku"] for line in details["all_lines"]}

    buckets = defaultdict(list)
    unlooked = [s for s in missing_skus if s not in ordered_skus]

    for sku in missing_skus:
        if sku not in ordered_skus:
            continue
        record = item_info.get(sku)
        if record is None:
            buckets["no item record in Neto"].append(sku)
        elif record["supplier"].lower() not in {s.lower() for s in SUPPLIERS}:
            buckets[f"supplier is {record['supplier']!r}, not in SUPPLIERS"].append(sku)
        else:
            buckets["ordered + correct supplier (unexpected!)"].append(sku)

    print("\n  Why they're missing:")
    for reason, skus in sorted(buckets.items()):
        print(f"    {len(skus):>3} x {reason}: {', '.join(skus[:8])}")

    if unlooked:
        print(f"    {len(unlooked):>3} x no order line in this window at all: "
              f"{', '.join(unlooked[:8])}")
        if details.get("to_date"):
            # Both runs covered an identical closed window, so drift can't explain
            # this — the API genuinely isn't seeing orders the scraper saw.
            print("\n  The window was pinned with --to-date, so this is NOT clock drift.\n"
                  "  The API is missing orders the scraper found. Investigate with:\n"
                  f"    python3 neto_api.py --investigate {','.join(unlooked[:12])}")
        else:
            print("\n  Neither run pins an end date, so a baseline captured earlier covers a\n"
                  "  different window. Re-run both with the same --to-date before digging in.")


def compare_with(summary, baseline_path, details=None):
    """Diff this run against a summary produced by the Selenium scraper.

    The point of this is to prove the API path is equivalent before switching the
    GUI over — a silent disagreement on quantities would mean under- or
    over-ordering real stock, so the two must be reconciled by hand first."""
    try:
        with open(baseline_path, "r", encoding="utf-8") as f:
            baseline = json.load(f)
    except (OSError, ValueError) as e:
        print(f"Couldn't read baseline {baseline_path}: {e}")
        return

    new_by_sku = {row["sku"]: row for row in summary}
    old_by_sku = {row["sku"]: row for row in baseline}

    only_new = sorted(set(new_by_sku) - set(old_by_sku))
    only_old = sorted(set(old_by_sku) - set(new_by_sku))
    shared = sorted(set(new_by_sku) & set(old_by_sku))

    print(f"\n--- Comparison against {os.path.basename(baseline_path)} ---")
    print(f"SKUs: {len(new_by_sku)} via API, {len(old_by_sku)} in baseline, {len(shared)} shared")

    if only_new:
        print(f"\nOnly in API result ({len(only_new)}): {', '.join(only_new[:20])}")
    if only_old:
        print(f"\nOnly in baseline ({len(only_old)}): {', '.join(only_old[:20])}")
        _explain_missing(only_old, details)

    qty_diffs, stock_diffs = [], []
    for sku in shared:
        new_row, old_row = new_by_sku[sku], old_by_sku[sku]
        if new_row["total_qty_needed"] != old_row["total_qty_needed"]:
            qty_diffs.append((sku, old_row["total_qty_needed"], new_row["total_qty_needed"]))
        if new_row["stock_available_to_sell"] != old_row.get("stock_available_to_sell"):
            stock_diffs.append(
                (sku, old_row.get("stock_available_to_sell"), new_row["stock_available_to_sell"])
            )

    print(f"\nQuantity differences: {len(qty_diffs)}")
    for sku, old, new in qty_diffs[:25]:
        print(f"  {sku}: baseline {old} -> API {new}")

    print(f"\nStock differences: {len(stock_diffs)}")
    for sku, old, new in stock_diffs[:25]:
        print(f"  {sku}: baseline {old} -> API {new}")

    if not qty_diffs and not stock_diffs and not only_new and not only_old:
        print("\nIdentical — the API path matches the scraper.")


def check_credentials(override=None):
    """Isolate auth problems from everything else.

    Neto has two mutually exclusive key styles — a global key (sent alone) and a
    user-based key (sent with NETOAPI_USERNAME) — and pairing a username with the
    global key fails just as surely as omitting it for a user-based one. The error
    text doesn't say which mistake you made, so rather than reason about it, this
    tries both shapes and reports which the store actually accepts."""
    store_url, api_key, username = load_credentials(*(override or (None, None, None)))

    print(f"Store URL: {store_url}")
    print(f"API key:   {len(api_key)} characters, ending {api_key[-4:]!r}")
    print(f"Username:  {username or '(not set)'}")
    if len(api_key) < 32:
        # The API Settings input is narrower than the key it holds, so selecting
        # the visible text silently drops the tail. Worth calling out before the
        # user goes hunting through permission groups for a non-existent problem.
        print(f"  ^ warning: Neto keys are usually 32+ characters. {len(api_key)} looks "
              f"truncated — click into the field and use Cmd+A, not a drag-select.")

    probe = {"Filter": {"Page": 0, "Limit": 1, "OutputSelector": ["SKU"]}}
    attempts = [("key alone (global key style)", (store_url, api_key, ""))]
    if username:
        attempts.append((f"key + username {username!r} (user-based key style)",
                         (store_url, api_key, username)))
    else:
        print("\nNo username configured, so only the global-key style can be tested.")

    working = None
    for label, creds in attempts:
        print(f"\nTrying {label}...")
        try:
            call_api("GetItem", probe, creds)
            print("  -> accepted")
            working = creds
            break
        except NetoAPIError as e:
            print(f"  -> rejected: {str(e).splitlines()[0]}")

    if not working:
        print(
            "\nNeither style worked. In order of likelihood:\n"
            "  1. The key is truncated — recopy it with Cmd+A from inside the field.\n"
            "  2. You're using the global key with a username, or a user-based key\n"
            "     without one. Check which key you actually copied.\n"
            "  3. For a user-based key, 'username' must be the staff user's login\n"
            "     username, which is not always their email address.\n"
            "  4. The staff user's permission group doesn't grant API access.\n"
            f"  5. store_url is wrong (currently {store_url})."
        )
        raise NetoAPIError("Credential check failed.")

    print("\nTesting GetOrder (Limit 1)...")
    response = call_api(
        "GetOrder",
        {"Filter": {"Page": 0, "Limit": 1, "OutputSelector": ["OrderID", "OrderStatus"]}},
        working,
    )
    print(f"  -> accepted, Ack={response.get('Ack')}, "
          f"{len(_as_list(response.get('Order')))} order returned.")

    used_username = working[2]
    print("\nCredentials work. Both Products and Orders are readable.")
    if bool(used_username) != bool(username):
        # The config disagrees with what actually worked, so say exactly what to
        # change rather than leaving a passing check that the real run won't match.
        if used_username:
            print(f'Update {CONFIG_FILENAME}: set "username" to {used_username!r}.')
        else:
            print(f'Update {CONFIG_FILENAME}: remove the "username" field — this is a '
                  f'global key and must be sent without one.')


def diagnose(from_date=None, baseline_path=None):
    """Work out why a filtered GetOrder returns nothing, by removing one filter at
    a time until orders appear.

    A GetOrder that returns 0 rows is indistinguishable from one that's correctly
    filtered down to nothing, so guessing at the cause is expensive. This narrows
    it to a single culprit, then reads the store's actual supplier spelling off
    products the scraper already found — the Supplier filter matches an exact
    string, so 'Games Workshop' vs 'Games Workshop AU' silently yields zero."""
    credentials = load_credentials()
    resolved_from = from_date or get_last_tuesday()
    date_from = datetime.combine(resolved_from, datetime.min.time()).strftime("%Y-%m-%d %H:%M:%S")

    def count_orders(label, order_filter):
        order_filter = dict(order_filter)
        order_filter.setdefault("OutputSelector", ["OrderID", "OrderStatus", "DatePlaced"])
        try:
            response = call_api("GetOrder", {"Filter": order_filter}, credentials)
        except NetoAPIError as e:
            print(f"  {label}: FAILED - {str(e).splitlines()[0]}")
            return []
        found = _as_list(response.get("Order"))
        print(f"  {label}: {len(found)} orders")
        return found

    print("=== 1. Which filter is returning nothing? ===")
    no_filter = count_orders("no filters at all      ", {"Page": 0, "Limit": 5})
    date_only = count_orders("date only              ", {"Page": 0, "Limit": 5,
                                                         "DatePlacedFrom": date_from})
    supplier_only = count_orders("supplier only          ", {"Page": 0, "Limit": 5,
                                                             "Supplier": list(SUPPLIERS)})
    both = count_orders("date + supplier (current)", {"Page": 0, "Limit": 5,
                                                      "DatePlacedFrom": date_from,
                                                      "Supplier": list(SUPPLIERS)})

    sample = no_filter or date_only
    if sample:
        print(f"\n  Sample DatePlaced values: {[o.get('DatePlaced') for o in sample[:3]]}")
        print(f"  We are asking for DatePlacedFrom >= {date_from}")

    print("\n=== 2. What does this store actually call the supplier? ===")
    discovered = []
    if not baseline_path or not os.path.exists(baseline_path):
        print("  (no baseline JSON given, skipping — pass --compare to enable)")
    else:
        with open(baseline_path, "r", encoding="utf-8") as f:
            skus = [row["sku"] for row in json.load(f)][:20]
        response = call_api("GetItem", {
            "Filter": {"SKU": skus, "OutputSelector": ["SKU", "PrimarySupplier", "Brand"]}
        }, credentials)
        items = _as_list(response.get("Item"))
        print(f"  Looked up {len(skus)} known-good SKUs, got {len(items)} items back.")
        discovered = sorted({(i.get("PrimarySupplier") or "").strip() for i in items} - {""})
        brands = sorted({(i.get("Brand") or "").strip() for i in items} - {""})
        print(f"  PrimarySupplier values: {discovered or '(none set)'}")
        print(f"  Brand values:           {brands or '(none set)'}")

        if discovered and set(discovered) - set(SUPPLIERS):
            print(f"\n  >>> Mismatch. neto_api.SUPPLIERS is {SUPPLIERS} but these products use "
                  f"{discovered}.\n      Update the SUPPLIERS constant at the top of neto_api.py.")
        elif discovered:
            print(f"\n  SUPPLIERS={SUPPLIERS} covers everything found here.")

        # Re-run the real query with the discovered codes so the fix is confirmed
        # here rather than after another round-trip of editing and re-running.
        if discovered and not both:
            print(f"\n  Retrying date + supplier using discovered codes {discovered}:")
            count_orders("  ->                     ", {"Page": 0, "Limit": 5,
                                                       "DatePlacedFrom": date_from,
                                                       "Supplier": discovered})

    print("\n=== Verdict ===")
    if both:
        print("  Current filters work. If the full run still returns nothing, the problem is\n"
              "  in pagination or the date range, not the filters.")
    elif not date_only and not no_filter:
        print("  GetOrder returns nothing even with only a date filter. The key's permission\n"
              "  group most likely doesn't grant order read access, even though GetItem works.")
    elif not date_only:
        print("  Orders exist, but the date filter excludes them all. Compare the DatePlaced\n"
              "  values above against the date being requested.")
    elif not supplier_only:
        # Deliberately checked after date_only: an unfiltered GetOrder returning 0
        # is normal (Neto wants at least one criterion) and must not be read as a
        # permissions failure when the date-filtered call plainly works.
        print("  Orders exist and the date filter is fine — the Supplier filter is what's\n"
              "  excluding everything. Use the PrimarySupplier codes printed above.")
    else:
        print("  Each filter works alone but not combined — likely genuinely no matching\n"
              "  orders in this window. Try an earlier --from-date.")


def investigate_skus(skus, from_date=None, to_date=None):
    """Ask Neto directly which orders contain these SKUs, ignoring the date filter.

    When a SKU is missing from a date-filtered fetch there are only two
    possibilities: the order isn't in the window (so the date filter is behaving
    differently than we think), or the order IS in the window but our fetch didn't
    return it (so pagination or the fetch itself is at fault). GetOrder's SKU filter
    settles which, because it finds the orders regardless of how we're paging."""
    credentials = load_credentials()
    resolved_from = from_date or get_last_tuesday()
    window_from, window_to = _local_window(resolved_from, to_date)

    print(f"Window under test (store-local): {window_from} .. "
          f"{window_to or '(no upper bound)'}")
    print(f"Sent to the API as UTC:          {_to_api_datetime(window_from)} .. "
          f"{_to_api_datetime(window_to) if window_to else '(none)'}")
    print(f"Investigating {len(skus)} SKUs. Times below are store-local.\n")

    # Constrain by date as well as SKU. An unconstrained SKU query returns the
    # product's entire order history oldest-first, which for a long-running store
    # fills the Limit with orders from years ago and never reaches the window —
    # making it look like there are no recent orders when we simply never saw them.
    wanted = set(skus)
    date_filter = {"DatePlacedFrom": _to_api_datetime(window_from)}
    if window_to:
        date_filter["DatePlacedTo"] = _to_api_datetime(window_to)

    # GrandTotal and OrderPayment must be requested explicitly: order_qualifies()
    # reads both, and without them every New order looks unpayable and gets
    # reported as excluded — an artefact of the query rather than a real finding.
    selectors = ["OrderID", "OrderStatus", "DatePlaced", "GrandTotal", "OrderPayment",
                 "OrderLine", "OrderLine.Quantity"]

    limit = 200
    response = call_api("GetOrder", {
        "Filter": dict(date_filter, SKU=list(skus), Page=0, Limit=limit,
                       OutputSelector=selectors)
    }, credentials)
    orders = _as_list(response.get("Order"))

    print(f"GetOrder(SKU + date window) returned {len(orders)} orders.")
    if len(orders) == limit:
        print(f"  (warning: exactly {limit} = the Limit, so this may be truncated)")

    def show(order):
        hits = [ln.get("SKU") for ln in _as_list(order.get("OrderLine"))
                if ln.get("SKU") in wanted]
        owed = amount_owed(order)
        verdict = "counts" if order_qualifies(order) else "excluded"
        local = _from_api_datetime(order.get("DatePlaced"))
        print(f"  {str(order.get('OrderID')):<14} {str(local):<20} "
              f"{str(order.get('OrderStatus')):<15} total={order.get('GrandTotal')!s:<9} "
              f"owed={owed!s:<9} {verdict:<9} {','.join(hits[:4])}")

    seen_in_window = set()
    for order in sorted(orders, key=lambda o: str(o.get("DatePlaced"))):
        show(order)
        seen_in_window.update(
            ln.get("SKU") for ln in _as_list(order.get("OrderLine"))
            if ln.get("SKU") in wanted
        )

    absent = sorted(wanted - seen_in_window)
    if absent:
        # Look further back to distinguish "this SKU sells, just not this week"
        # from "the SKU filter isn't matching at all". A 60-day lookback is recent
        # enough that the Limit won't fill with ancient history the way an
        # unbounded query does.
        lookback_from = window_from - timedelta(days=60)
        print(f"\n  {len(absent)} of the {len(wanted)} SKUs had no in-window order. "
              f"Looking back to {lookback_from.date()}:")
        response = call_api("GetOrder", {
            "Filter": {
                "SKU": absent,
                "DatePlacedFrom": _to_api_datetime(lookback_from),
                "Page": 0, "Limit": limit, "OutputSelector": selectors,
            }
        }, credentials)
        recent = _as_list(response.get("Order"))
        if len(recent) == limit:
            print(f"    (warning: exactly {limit} rows = the Limit, possibly truncated)")
        if not recent:
            print("    No orders at all in the last 60 days for these SKUs.")
        else:
            for order in sorted(recent, key=lambda o: str(o.get("DatePlaced")))[-15:]:
                show(order)

    print("\n  >>> Compare the dates above against the window. Orders inside the window\n"
          "      that the main fetch missed mean a fetch bug; orders only outside it mean\n"
          "      the scraper's control-panel filter was matching a wider range than its\n"
          "      URL date params imply.")


def _parse_date(value, label):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%d/%m/%Y").date()
    except ValueError:
        raise NetoAPIError(f"Invalid --{label} '{value}' (expected DD/MM/YYYY).")


def main():
    """Standalone CLI for testing this against the live store, e.g.
        python neto_api.py --from-date 21/07/2026 --compare sales_order_demand.json
    The GUI does not use this; it would import run() directly."""
    parser = argparse.ArgumentParser(description="Fetch Neto sales order demand via the API.")
    parser.add_argument("--from-date", dest="from_date", default=None,
                        help="Date Placed From, DD/MM/YYYY (default: last Tuesday).")
    parser.add_argument("--to-date", dest="to_date", default=None,
                        help="Date Placed Till, DD/MM/YYYY (default: no upper bound).")
    parser.add_argument("--output", default="sales_order_demand_api.json",
                        help="Output filename (default: sales_order_demand_api.json, so a "
                             "test run can't overwrite the scraper's output).")
    parser.add_argument("--compare", default=None, metavar="PATH",
                        help="Diff the result against an existing scraper-produced JSON.")
    parser.add_argument("--check", action="store_true",
                        help="Test credentials only, then exit. Use this to isolate auth "
                             "problems before debugging filters.")
    parser.add_argument("--investigate", default=None, metavar="SKUS",
                        help="Comma-separated SKUs: find every order containing them, "
                             "ignoring the date filter, to see whether the fetch should "
                             "have picked them up.")
    parser.add_argument("--diagnose", action="store_true",
                        help="Find out why a filtered fetch returns 0 orders, by testing "
                             "each filter in isolation. Combine with --compare to also "
                             "read the store's real supplier name off known SKUs.")
    parser.add_argument("--key", default=None,
                        help="Override the API key for this run, without editing "
                             f"{CONFIG_FILENAME}. Useful with --check.")
    parser.add_argument("--username", default=None,
                        help="Override the username for this run. Pass --username '' to "
                             "force global-key style with no username header.")
    parser.add_argument("--store-url", dest="store_url", default=None,
                        help="Override the store URL for this run.")
    args = parser.parse_args()

    if args.check:
        check_credentials((args.key, args.username, args.store_url))
        return

    if args.investigate:
        investigate_skus(
            [s.strip() for s in args.investigate.split(",") if s.strip()],
            _parse_date(args.from_date, "from-date"),
            _parse_date(args.to_date, "to-date"),
        )
        return

    if args.diagnose:
        baseline = args.compare
        if baseline and not os.path.isabs(baseline):
            baseline = os.path.join(_base_dir(), baseline)
        diagnose(_parse_date(args.from_date, "from-date"), baseline)
        return

    from_date = _parse_date(args.from_date, "from-date")
    to_date = _parse_date(args.to_date, "to-date")

    details = collect(from_date, to_date, print)
    summary = details["summary"]

    output_path = os.path.join(_base_dir(), args.output)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2)
    print(f"Saved demand data to {output_path}")

    if args.compare:
        baseline = args.compare
        if not os.path.isabs(baseline):
            baseline = os.path.join(_base_dir(), baseline)
        compare_with(summary, baseline, details)


if __name__ == "__main__":
    try:
        main()
    except NetoAPIError as e:
        print(f"Error: {e}")
        sys.exit(1)
