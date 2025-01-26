import os
import json
import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
API_KEY_PATH = os.path.join(BASE_DIR, 'api_key.json')
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')

with open(API_KEY_PATH) as f:
    config = json.load(f)
    API_KEY = config['api_key']

with open(CONFIG_PATH, 'r') as f:
    config = json.load(f)

METRIC_GROUPS = config['metric_groups']
COLORS = config['colors']

DARK_GREEN = COLORS['DARK_GREEN']
MED_GREEN = COLORS['MED_GREEN']
LIGHT_GREEN = COLORS['LIGHT_GREEN']
DARK_RED = COLORS['DARK_RED']
MED_RED = COLORS['MED_RED']
LIGHT_RED = COLORS['LIGHT_RED']
YELLOW = COLORS['YELLOW']


def calculate_percentiles(values):
    return {
        6: np.nanpercentile(values, 6),
        12: np.nanpercentile(values, 12),
        25: np.nanpercentile(values, 25),
        75: np.nanpercentile(values, 75),
        88: np.nanpercentile(values, 88),
        94: np.nanpercentile(values, 94)
    }

def format_metrics(range_obj, values, metric_name):
    good_high = ['EPS', 'Rev 3YCAGR', 'GP Marg', 'EBITDA Marg', 'Net Marg', 'Op Marg', 'FCF Marg',
                    'Cash Ratio', 'Cash/Debt', 'WC Turn', 'Asset Turn', 'ROA', 'ROE', 'ROIC']
    good_low = ['TEV/Rev', 'D/E', 'Debt/EBITDA']

    for cell, value in zip(range_obj, values):
        if pd.notna(value) and not np.isinf(value):
            if metric_name == 'CurrRatio':
                if value >= 3.0:
                    cell.color = DARK_GREEN
                elif value >= 2.0:
                    cell.color = MED_GREEN
                elif value >= 1.2:
                    cell.color = LIGHT_GREEN
                elif value >= 0.8:
                    cell.color = None
                elif value >= 0.5:
                    cell.color = LIGHT_RED
                else:
                    cell.color = DARK_RED

            elif metric_name == 'QuickRatio':
                if value >= 2.0:
                    cell.color = DARK_GREEN
                elif value >= 1.5:
                    cell.color = MED_GREEN
                elif value >= 1.0:
                    cell.color = LIGHT_GREEN
                elif value >= 0.5:
                    cell.color = LIGHT_RED
                else:
                    cell.color = DARK_RED
            
            elif metric_name in good_high:
                percentiles = calculate_percentiles(values)
                if percentiles[25] is not None:
                    if value >= percentiles[94]:
                        cell.color = DARK_GREEN
                    elif value >= percentiles[88]:
                        cell.color = MED_GREEN
                    elif value >= percentiles[75]:
                        cell.color = LIGHT_GREEN
                    elif value <= percentiles[6]:
                        cell.color = DARK_RED
                    elif value <= percentiles[12]:
                        cell.color = MED_RED
                    elif value <= percentiles[25]:
                        cell.color = LIGHT_RED

            elif metric_name in good_low:
                percentiles = calculate_percentiles(values)
                if percentiles[25] is not None:
                    if value <= percentiles[6]:
                        cell.color = DARK_GREEN
                    elif value <= percentiles[12]:
                        cell.color = MED_GREEN
                    elif value <= percentiles[25]:
                        cell.color = LIGHT_GREEN
                    elif value >= percentiles[94]:
                        cell.color = DARK_RED
                    elif value >= percentiles[88]:
                        cell.color = MED_RED
                    elif value >= percentiles[75]:
                        cell.color = LIGHT_RED