import os
import json
import logging
import requests
from datetime import datetime, timezone

import gspread
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%SZ",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DASHBOARD_UUID = "6b6cac6b-360d-405f-9596-39e10cd58ce2"
DASHBOARD_URL = f"https://bi.pondthreadsoms.com/api/public/dashboard/{DASHBOARD_UUID}"
DASHCARD_URL = DASHBOARD_URL + "/dashcard/{dashcard_id}/card/{card_id}"

SHEET_NAME = "Stimulus SLA Tracker"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Scalar cards: card_id -> column header
SCALAR_CARDS = {
    844: "new_unit_items_2w",
    846: "shipped_unit_items_2w",
    848: "orders_on_production",
    849: "unit_items_on_production",
    850: "pct_orders_past_due",
    851: "current_month_sla_pct",
}

# Series cards: card_id -> tab name
SERIES_CARDS = {
    845: "SKU_Pending",
    852: "Shipped_Per_Day",
    853: "New_vs_Shipped_Per_Day",
    854: "SLA_By_Month",
    855: "Fulfillment_Time_Per_Week",
    877: "New_vs_Shipped_Per_Month",
}

SNAPSHOT_TAB = "Daily_Snapshot"
SNAPSHOT_COLUMNS = [
    "run_date",
    "new_unit_items_2w",
    "shipped_unit_items_2w",
    "orders_on_production",
    "unit_items_on_production",
    "pct_orders_past_due",
    "current_month_sla_pct",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_dashcard_map() -> dict[int, int]:
    """Fetch the public dashboard and return {card_id: dashcard_id}."""
    log.info("Fetching dashboard metadata from %s", DASHBOARD_URL)
    resp = requests.get(DASHBOARD_URL, timeout=30)
    resp.raise_for_status()
    dashboard = resp.json()

    mapping: dict[int, int] = {}
    for dashcard in dashboard.get("dashcards", []):
        card = dashcard.get("card") or {}
        card_id = card.get("id")
        dashcard_id = dashcard.get("id")
        if card_id and dashcard_id:
            mapping[card_id] = dashcard_id
            log.info("  card_id=%s -> dashcard_id=%s", card_id, dashcard_id)

    log.info("Found %d dashcard mappings", len(mapping))
    return mapping


def fetch_card(card_id: int, dashcard_id: int) -> dict | None:
    url = DASHCARD_URL.format(dashcard_id=dashcard_id, card_id=card_id)
    log.info("Fetching card %s (dashcard %s)", card_id, dashcard_id)
    try:
        resp = requests.post(url, json={"parameters": []}, timeout=30)
        resp.raise_for_status()
        return resp.json()
    except Exception as exc:
        log.error("Failed to fetch card %s: %s", card_id, exc)
        return None


def extract_scalar(data: dict) -> object:
    """Pull a single value out of a Metabase scalar card response."""
    try:
        rows = data["data"]["rows"]
        if rows:
            log.info("  rows[0] = %s", rows[0])
            return rows[0][0]
    except (KeyError, IndexError, TypeError):
        pass
    return None


def extract_series(data: dict) -> tuple[list[str], list[list]]:
    """Return (headers, rows) from a Metabase series card response."""
    try:
        cols = [c.get("display_name") or c.get("name", f"col{i}")
                for i, c in enumerate(data["data"]["cols"])]
        rows = data["data"]["rows"]
        return cols, rows
    except (KeyError, TypeError):
        return [], []


def get_or_create_tab(sheet, tab_name: str, headers: list[str]) -> gspread.Worksheet:
    """Return the worksheet, creating it with a header row if needed."""
    try:
        ws = sheet.worksheet(tab_name)
        log.info("Tab '%s' already exists", tab_name)
        return ws
    except gspread.WorksheetNotFound:
        log.info("Creating tab '%s'", tab_name)
        ws = sheet.add_worksheet(title=tab_name, rows=1000, cols=len(headers) + 2)
        ws.append_row(headers)
        return ws


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    run_date = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    log.info("Run date: %s", run_date)

    # --- Google auth ---
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise EnvironmentError("GOOGLE_CREDENTIALS_JSON environment variable is not set")

    creds_info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    gc = gspread.authorize(creds)
    log.info("Authenticated with Google")

    sheet = gc.open(SHEET_NAME)
    log.info("Opened sheet '%s'", SHEET_NAME)

    # --- Build card_id -> dashcard_id map ---
    dashcard_map = get_dashcard_map()

    # --- Scalar cards ---
    scalar_values: dict[str, object] = {}
    for card_id, col_name in SCALAR_CARDS.items():
        dashcard_id = dashcard_map.get(card_id)
        if dashcard_id is None:
            log.warning("card_id=%s not found in dashboard, skipping", card_id)
            scalar_values[col_name] = None
            continue
        data = fetch_card(card_id, dashcard_id)
        scalar_values[col_name] = extract_scalar(data) if data else None
        log.info("Scalar card %s (%s) = %s", card_id, col_name, scalar_values[col_name])

    # Write snapshot row
    snapshot_ws = get_or_create_tab(sheet, SNAPSHOT_TAB, SNAPSHOT_COLUMNS)
    snapshot_row = [run_date] + [scalar_values.get(col) for col in SNAPSHOT_COLUMNS[1:]]
    snapshot_ws.append_row(snapshot_row)
    log.info("Appended snapshot row: %s", snapshot_row)

    # --- Series cards ---
    for card_id, tab_name in SERIES_CARDS.items():
        dashcard_id = dashcard_map.get(card_id)
        if dashcard_id is None:
            log.warning("card_id=%s not found in dashboard, skipping tab '%s'", card_id, tab_name)
            continue
        data = fetch_card(card_id, dashcard_id)
        if data is None:
            log.warning("Skipping series card %s (%s) due to fetch error", card_id, tab_name)
            continue

        headers, rows = extract_series(data)
        if not headers:
            log.warning("No data returned for card %s (%s)", card_id, tab_name)
            continue

        full_headers = ["run_date"] + headers
        ws = get_or_create_tab(sheet, tab_name, full_headers)

        for row in rows:
            ws.append_row([run_date] + list(row))

        log.info("Appended %d rows to tab '%s'", len(rows), tab_name)

    log.info("Done.")


if __name__ == "__main__":
    main()
