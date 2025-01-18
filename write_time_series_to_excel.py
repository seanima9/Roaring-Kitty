import sys
import os
import numpy as np
import pandas as pd
import nasdaqdatalink as ndl
import xlwings as xw
import json

BASE_DIR = os.path.dirname(__file__)
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

ndl.ApiConfig.api_key = API_KEY


########################################### Get Data ###########################################


def grab_sf1_time_series_data(ticker):
    """
    Grab fundamental data for a given ticker from the Sharadar SF1 dataset.
    """
    metrics = {}
    data = ndl.get_table('SHARADAR/SF1', ticker=ticker, paginate=True)

    data = data[data['dimension'] == 'ART']  # As Reported, Trailing Twelve Months (TTM)
    ltm = data.iloc[0:1]
    data = data[data['fiscalperiod'].str.contains('Q4')].copy()

    if data['fiscalperiod'].duplicated().any():
        data = data.sort_values('calendardate', ascending=False).drop_duplicates('fiscalperiod').sort_index()

    data['calendardate'] = pd.to_datetime(data['calendardate'])
    data['year'] = data['calendardate'].dt.year
    ltm['calendardate'] = pd.to_datetime(ltm['calendardate'])
    ltm['year'] = 'LTM'

    data = data.sort_values('year').reset_index(drop=True)
    data = pd.concat([data, ltm])

    # Valuation Metrics
    metrics['TEV'] = data['ev'] / 1_000_000
    metrics['Market Cap'] = data['marketcap'] / 1_000_000
    metrics['TEV / EBITDA'] = data['ev'] / data['ebitda']
    metrics['TEV / Revenue'] = data['ev'] / data['revenue']
    metrics['TEV / FCF'] = data['ev'] / data['fcf']
    metrics['P/E'] = data['pe']
    metrics['P/B'] = data['pb']

    # Income Statement
    metrics['Revenue'] = data['revenue'] / 1_000_000
    metrics['Revenue %'] = (data['revenue'] / 1_000_000).pct_change()
    metrics['Gross Profit'] = data['gp'] / 1_000_000
    metrics['GP %'] = (data['gp'] / 1_000_000).pct_change()
    metrics['Net Income'] = data['netinc'] / 1_000_000
    metrics['Net Income %'] = (data['netinc'] / 1_000_000).pct_change()
    metrics['Operating Income'] = data['opinc'] / 1_000_000
    metrics['Total Operating Expense'] = data['opex'] / 1_000_000
    metrics['EBITDA'] = data['ebitda'] / 1_000_000
    metrics['EBITDA %'] = (data['ebitda'] / 1_000_000).pct_change()
    metrics['EPS'] = data['eps']
    metrics['Interest Expense'] = data['intexp'] / 1_000_000

    # Cash Flow
    metrics['CFO'] = data['ncfo'] / 1_000_000
    metrics['CFO %'] = (data['ncfo'] / 1_000_000).pct_change()
    metrics['FCF'] = data['fcf'] / 1_000_000
    metrics['FCF %'] = (data['fcf'] / 1_000_000).pct_change()

    # Balance Sheet
    metrics['Equity'] = data['equity'] / 1_000_000
    metrics['Debt'] = data['debt'] / 1_000_000
    metrics['Total Assets'] = data['assets'] / 1_000_000
    metrics['Total Liabilities'] = data['liabilities'] / 1_000_000
    metrics['Net Working Capital'] = (data['assetsc'] - data['liabilitiesc']) / 1_000_000
    metrics['Cash Ratio'] = data['cashneq'] / data['liabilitiesc']
    metrics['Tangible Book Value'] = (data['assets'] - data['intangibles'] - data['liabilities']) / 1_000_000
    metrics['TBV Per Share'] = metrics['Tangible Book Value'] / data['shareswa']
    metrics['Leverage Ratio'] = data['assets'] / data['equity']
    metrics['Interest-Bearing Debt'] = (data['debtc'] + data['debtnc']) / 1_000_000
    metrics['Debt to Capital'] = data['debt'] / (data['debt'] + data['equity'])
    metrics['Cash to Debt'] = data['cashneq'] / data['debt']
    metrics['Net Debt to Total Assets'] = (data['debt'] - data['cashneq']) / data['assets']
    metrics['Working Capital Turnover'] = data['revenue'] / (data['assetsc'] - data['liabilitiesc'])

    # Margins
    metrics['GP Margin'] = data['grossmargin']
    metrics['EBITDA Margin'] = data['ebitdamargin']
    metrics['Net Margin'] = data['netmargin']
    metrics['Operating Margin'] = data['opinc'] / data['revenue']
    metrics['Free Cash Flow Margin'] = data['fcf'] / data['revenue']

    # Ratios
    metrics['Current Ratio'] = data['currentratio']
    metrics['Quick Ratio'] = (data['assetsc'] - data['inventory']) / data['liabilitiesc']
    metrics['Payout Ratio'] = data['payoutratio']
    metrics['D/E'] = data['debt'] / data['equity']
    metrics['Debt to EBITDA'] = data['debt'] / data['ebitda']
    metrics['Asset Turnover'] = data['assetturnover']
    metrics['Int Coverage'] = data['ebit'] / data['intexp']

    # Return Metrics
    metrics['ROA'] = data['roa']
    metrics['ROE'] = data['roe']
    metrics['ROIC'] = data['roic']

    # Create DataFrame with years as index
    metrics_df = pd.DataFrame(metrics)
    metrics_df = metrics_df.replace([np.inf, -np.inf], np.nan)
    metrics_df.index = data['year']
    
    return metrics_df.round(2)


########################################### Excel Formatting ###########################################


def calculate_percentiles(values):
    return {
        10: np.nanpercentile(values, 10),
        25: np.nanpercentile(values, 25),
        50: np.nanpercentile(values, 50),
        75: np.nanpercentile(values, 75),
        90: np.nanpercentile(values, 90)
    }


def get_metric_group(metric_name):
    for group in METRIC_GROUPS:
        if metric_name in group['metrics']:
            return group['name']
    raise ValueError(f"Metric '{metric_name}' not found in any group")


def format_growth_metrics(range_obj, values):
    percentiles = calculate_percentiles(values)
    
    if percentiles[50] is None:
        return
        
    for cell, value in zip(range_obj, values):
        if pd.notna(value) and not np.isinf(value):
            if value >= percentiles[75]:
                cell.color = DARK_GREEN
            elif value >= percentiles[50]:
                cell.color = MED_GREEN
            elif value >= percentiles[25]:
                cell.color = LIGHT_GREEN
            elif value >= percentiles[10]:
                cell.color = LIGHT_RED
            else:
                cell.color = DARK_RED
        else:
            cell.color = None


def format_margin_metrics(range_obj, values):
    percentiles = calculate_percentiles(values)
    
    if percentiles[50] is None:
        return
        
    for cell, value in zip(range_obj, values):
        if pd.notna(value) and not np.isinf(value):
            if value >= percentiles[75]:
                cell.color = DARK_GREEN
            elif value >= percentiles[50]:
                cell.color = MED_GREEN
            elif value >= percentiles[25]:
                cell.color = LIGHT_GREEN
            elif value >= percentiles[10]:
                cell.color = LIGHT_RED
            else:
                cell.color = DARK_RED
        else:
            cell.color = None


def format_cash_flow_metrics(range_obj, values):
    percentiles = calculate_percentiles(values)
    
    if percentiles[50] is None:
        return
        
    for cell, value in zip(range_obj, values):
        if pd.notna(value) and not np.isinf(value):
            if value >= percentiles[75]:
                cell.color = DARK_GREEN
            elif value >= percentiles[50]:
                cell.color = MED_GREEN
            elif value >= percentiles[25]:
                cell.color = LIGHT_GREEN
            elif value >= percentiles[10]:
                cell.color = LIGHT_RED
            else:
                cell.color = DARK_RED
        else:
            cell.color = None


def format_ratio_metrics(range_obj, values, metric_name):
    for cell, value in zip(range_obj, values):
        if pd.notna(value) and not np.isinf(value):
            if metric_name in ['Current Ratio', 'Quick Ratio']:
                if value >= 1.5:
                    cell.color = DARK_GREEN
                elif value >= 1.2:
                    cell.color = MED_GREEN
                elif value >= 1.0:
                    cell.color = YELLOW
                elif value >= 0.8:
                    cell.color = LIGHT_RED
                else:
                    cell.color = DARK_RED

            elif metric_name in ['D/E', 'Debt to EBITDA']:
                percentiles = calculate_percentiles(values)
                if percentiles[50] is not None:
                    if value <= percentiles[25]:
                        cell.color = DARK_GREEN
                    elif value <= percentiles[50]:
                        cell.color = MED_GREEN
                    elif value <= percentiles[75]:
                        cell.color = LIGHT_GREEN
                    elif value <= percentiles[90]:
                        cell.color = LIGHT_RED
                    else:
                        cell.color = DARK_RED
        else:
            cell.color = None


def format_return_metrics(range_obj, values):
    percentiles = calculate_percentiles(values)
    
    # If no valid percentiles, skip formatting
    if percentiles[50] is None:
        return
        
    for cell, value in zip(range_obj, values):
        if pd.notna(value) and not np.isinf(value):
            if value >= percentiles[75]:
                cell.color = DARK_GREEN
            elif value >= percentiles[50]:
                cell.color = MED_GREEN
            elif value >= percentiles[25]:
                cell.color = LIGHT_GREEN
            elif value >= percentiles[10]:
                cell.color = LIGHT_RED
            else:
                cell.color = DARK_RED
        else:
            cell.color = None


def apply_conditional_formatting(sheet, metrics_df, start_row, start_col):
    # Work with transposed data to match Excel layout
    transposed_metrics = metrics_df.transpose()
    
    for row_idx, metric_name in enumerate(transposed_metrics.index):
        metric_group = get_metric_group(metric_name)
        
        if metric_group == 'Valuation Metrics':
            continue
            
        row_values = transposed_metrics.loc[metric_name].values
        
        # Calculate the range for this row's data cells
        current_row = start_row + 1 + row_idx  # +1 to skip header row
        data_range = sheet.range((current_row, start_col + 2),  # +2 to skip category and metric name columns
                               (current_row, start_col + 2 + len(row_values) - 1))
        
        # Apply formatting based on metric group
        if metric_group == 'Income Statement' and '%' in metric_name:
            format_growth_metrics(data_range, row_values)
        elif metric_group == 'Cash Flow' and '%' in metric_name:
            format_cash_flow_metrics(data_range, row_values)
        elif metric_group == 'Margins':
            format_margin_metrics(data_range, row_values)
        elif metric_group == 'Ratios':
            format_ratio_metrics(data_range, row_values, metric_name)
        elif metric_group == 'Returns':
            format_return_metrics(data_range, row_values)


########################################### Write to Excel ###########################################


def write_to_excel(sheet, metrics, start_row=3, start_col=12):
    years = sorted([idx for idx in metrics.index if idx != 'LTM'])
    if len(years) > 20:
        years = years[-20:]
        metrics = metrics.loc[years + ['LTM']]

    transposed_metrics = metrics.transpose()

    headers = ["Category", "Metric"] + [str(year) for year in years] + ["LTM"]
    for col_num, header in enumerate(headers, start=start_col):
        sheet.cells(start_row, col_num).value = header

    current_row = start_row + 1

    for i, group in enumerate(METRIC_GROUPS):
        group_start_row = current_row
        first_metric = True
        
        for metric_name, description in group['metrics'].items():
            if metric_name in transposed_metrics.index:
                # Write category name (only for first metric in group)
                if first_metric:
                    category_cell = sheet.cells(current_row, start_col)
                    category_cell.value = group['name']
                    first_metric = False
                
                # Write metric name
                metric_cell = sheet.cells(current_row, start_col + 1)
                metric_cell.value = metric_name
                
                # Add tooltip as comment
                if metric_cell.api.Comment is not None:
                    metric_cell.api.Comment.Delete()
                metric_cell.api.AddComment(description)
                metric_cell.api.Comment.Visible = False
                
                # Write values
                for col_num, value in enumerate(transposed_metrics.loc[metric_name], start=start_col + 2):
                    cell = sheet.cells(current_row, col_num)
                    cell.value = value
                    
                    # Format cells
                    if '%' in metric_name:
                        cell.api.NumberFormat = "0.0%"
                    elif isinstance(value, (int, float)) and abs(value) >= 1000:
                        cell.api.NumberFormat = "#,##0.00"
                    
                current_row += 1

        if i < len(METRIC_GROUPS) - 1:
            border_range = sheet.range(
                sheet.cells(group_start_row, start_col),
                sheet.cells(current_row - 1, start_col + len(years) + 2)
            )
            border_range.api.Borders(9).Weight = 2  # Bottom border weight = 2 (thick)

    # Create Excel table
    table_range = sheet.range(
        sheet.cells(start_row, start_col),
        sheet.cells(current_row - 1, start_col + len(years) + 2)
    )
    
    # Convert range to table and apply style
    table = sheet.api.ListObjects.Add(1, table_range.api, None, 1)
    table.TableStyle = "TableStyleLight1"

    # Set minimum width for category and metric columns
    sheet.api.Columns(start_col).ColumnWidth = 20
    sheet.api.Columns(start_col + 1).ColumnWidth = 25
    
    # Auto-fit remaining columns
    for col in range(start_col + 2, start_col + len(years) + 2):
        sheet.api.Columns(col).AutoFit()

    sheet = apply_conditional_formatting(sheet, metrics, start_row, start_col)

    return sheet


########################################### Main ###########################################


def main():
    spreadsheet_path = sys.argv[1]
    ticker = sys.argv[2]
    metrics = grab_sf1_time_series_data(ticker)

    wb = xw.books.active
    sheet = wb.sheets.active
    sheet = write_to_excel(sheet, metrics)
    wb.save(spreadsheet_path)


main()