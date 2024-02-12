from datetime import datetime, timedelta
import time


STORE_MAP = {
    "6400": "Hyundai AKJD", # AK ~ AS
    "6401": "Galleria", # AT ~ BB
    "6402": "Hyundai COEX", # BC ~ BK
    "6403": "Lotte Avenuel", # BL ~ BT
    "6404": "SSG Kangnam", # BU ~ CC
    "6405": "SSG Time square", # CD ~ CL
    "6407": "SSG Centum", # CM ~ CU
    "6408": "Maison Cheongdam", # CV ~ DD
    "6409": "Lotte Avenuel WT", # DE ~ DM
    "6410": "SSG Daegu", # DN ~ DV
    "6411": "Hyundai Pangyo" # DW ~ EE
}

PASS, SUCCESS, JINNY_THINK, JOY_CONFIRM, NO_6M_SALES = "pass", "success", "jinny_think", "joy_confirm", "no_6M_sales"


def get_current_time(date_format=None, day_delta=None):
    today = datetime.utcnow() + timedelta(hours=9)
    today += timedelta(days=day_delta) if day_delta is not None else timedelta()

    return today if date_format is None else today.strftime(date_format)

def get_column_number(column_id):
    return sum((ord(char) - ord('A') + 1) * 26 ** i for i, char in enumerate(reversed(column_id.upper()))) - 1

def get_last_data_idx(ws):
    start_row = end_row = 8
    for row in ws.iter_rows(min_row=start_row, values_only=True):
        end_row += 1 if row[1] is not None else 0
        if row[1] is None:
            break
    return end_row - 1

def convert_to_zero(value):
    try:
        return float(value) if value not in (None, '-') else 0.0
    except ValueError:
        return 0.0