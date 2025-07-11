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
    # Metrics where higher values are better
    good_high = ['EPS', 'Rev 3YCAGR', 'GP Marg', 'EBITDA Marg', 'Net Marg', 'Op Marg', 'FCF Marg',
                'Cash Ratio', 'Cash/Debt', 'WC Turn', 'Asset Turn', 'ROA', 'ROE', 'ROIC', 'Net Cash',
                'NI to CFO', 'Recv Turn', 'Inv Turn', 'Int Cov']
    
    # Metrics where lower values are better
    good_low = ['TEV/Rev', 'D/E', 'Debt/EBITDA', 'R&D/Rev', 'SG&A/Rev', 'SBC/Rev', 
               'DSO', 'DIO', 'DPO', 'Cash Cycle']

    for cell, value in zip(range_obj, values):
        if pd.notna(value) and not np.isinf(value):
            if metric_name == 'Curr Ratio':
                if value >= 3.0:
                    cell.color = DARK_GREEN
                elif value >= 2.0:
                    cell.color = MED_GREEN
                elif value >= 1.2:
                    cell.color = LIGHT_GREEN
                elif value >= 0.8:
                    continue
                elif value >= 0.5:
                    cell.color = LIGHT_RED
                else:
                    cell.color = DARK_RED

            elif metric_name == 'Quick Ratio':
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

            elif metric_name == 'Ins Buys':
                if value >= 10.0:
                    cell.color = DARK_GREEN
                elif value >= 6.0:
                    cell.color = MED_GREEN
                elif value >= 3.0:
                    cell.color = LIGHT_GREEN

            elif metric_name == 'BB Yield':
                if value >= 0.05:
                    cell.color = DARK_GREEN
                elif value >= 0.02:
                    cell.color = MED_GREEN
                elif value >= 0.01:
                    cell.color = LIGHT_GREEN
                elif value >= 0.00:
                    continue
                elif value >= -0.02:
                    cell.color = LIGHT_RED
                elif value >= -0.04:
                    cell.color = MED_RED
                else:
                    cell.color = DARK_RED

            # Special handling for Cash Cycle (negative is better)
            elif metric_name == 'Cash Cycle':
                percentiles = calculate_percentiles(values)
                if percentiles[25] is not None:
                    if value <= percentiles[6]:  # Most negative (best)
                        cell.color = DARK_GREEN
                    elif value <= percentiles[12]:
                        cell.color = MED_GREEN
                    elif value <= percentiles[25]:
                        cell.color = LIGHT_GREEN
                    elif value >= percentiles[94]:  # Most positive (worst)
                        cell.color = DARK_RED
                    elif value >= percentiles[88]:
                        cell.color = MED_RED
                    elif value >= percentiles[75]:
                        cell.color = LIGHT_RED

            # Special handling for NI to CFO (should be close to 1.0, higher is better)
            elif metric_name == 'NI to CFO':
                if value >= 1.5:
                    cell.color = DARK_GREEN
                elif value >= 1.2:
                    cell.color = MED_GREEN
                elif value >= 1.0:
                    cell.color = LIGHT_GREEN
                elif value >= 0.8:
                    continue  # Neutral
                elif value >= 0.6:
                    cell.color = LIGHT_RED
                elif value >= 0.4:
                    cell.color = MED_RED
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