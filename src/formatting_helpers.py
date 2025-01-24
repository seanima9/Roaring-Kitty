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


def get_metric_group(metric_name):
    for group in METRIC_GROUPS:
        if metric_name in group['metrics']:
            return group['name']
    raise ValueError(f"Metric '{metric_name}' not found in any group")


def format_metrics(range_obj, values, metric_name):
    metrics_for_percentiles = [
    'Rev CAGR', 'GP CAGR', 'Net Inc CAGR', 'EBITDA CAGR', 
    'CFO CAGR', 'FCF CAGR','Rev \u0394', 'GP \u0394',
    'Net Inc \u0394', 'EBITDA \u0394', 'CFO \u0394', 'FCF \u0394',
    'Cash Ratio', 'Cash/Debt', 'WC Turn', 'GP Marg',
    'EBITDA Marg', 'Net Marg', 'Op Marg', 'FCF Marg', 
    'ROA', 'ROE', 'ROIC'
    ]
    
    for cell, value in zip(range_obj, values):
        if pd.notna(value) and not np.isinf(value):
            if metric_name == 'Current Ratio':
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
            
            elif metric_name in metrics_for_percentiles:
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