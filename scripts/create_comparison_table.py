import sys
import os
import numpy as np
import pandas as pd
import nasdaqdatalink as ndl
import xlwings as xw
import json

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

ndl.ApiConfig.api_key = API_KEY


########################################### Get Data ###########################################


def grab_sf1_data(tickers):
    def calculate_yoy_changes(values):
        """Calculate average year-over-year percentage change."""
        yoy_changes = []
        for i in range(1, len(values)):
            if values[i-1] != 0:
                yoy_change = (values[i] - values[i-1]) / abs(values[i-1]) * 100
                yoy_changes.append(yoy_change)
        
        return np.mean(yoy_changes) if yoy_changes else np.nan
    
    all_metrics = {}
    
    for ticker in tickers:
        metrics = {}
        data = ndl.get_table('SHARADAR/SF1', ticker=ticker, paginate=True)
        
        # Filter for TTM data and last 5 years
        data = data[data['dimension'] == 'ART']
        ltm = data.iloc[0:1]
        annual_data = data[data['fiscalperiod'].str.contains('Q4')].copy()
        annual_data = annual_data.sort_values('calendardate', ascending=False).head(5)
        
        if annual_data['fiscalperiod'].duplicated().any():
            annual_data = annual_data.sort_values('calendardate', 
                                                ascending=False).drop_duplicates('fiscalperiod').sort_index()
        
        # Prepare data
        annual_data['calendardate'] = pd.to_datetime(annual_data['calendardate'])
        annual_data['year'] = annual_data['calendardate'].dt.year
        annual_data = annual_data.sort_values('year')
        
        # Valuation Metrics
        metrics['TEV'] = ltm['ev'] / 1_000_000
        metrics['Market Cap'] = ltm['marketcap'] / 1_000_000
        metrics['TEV / EBITDA'] = ltm['ev'] / ltm['ebitda']
        metrics['TEV / Revenue'] = ltm['ev'] / ltm['revenue']
        metrics['TEV / FCF'] = ltm['ev'] / ltm['fcf']
        metrics['P/E'] = ltm['pe']
        metrics['P/B'] = ltm['pb']

        # Income Statement
        metrics['Revenue'] = ltm['revenue'] / 1_000_000
        metrics['Revenue YoY %'] = calculate_yoy_changes(annual_data['revenue'].values)
        metrics['Gross Profit'] = ltm['gp'] / 1_000_000
        metrics['Gross Profit YoY %'] = calculate_yoy_changes(annual_data['gp'].values)
        metrics['Net Income'] = ltm['netinc'] / 1_000_000
        metrics['Net Income YoY %'] = calculate_yoy_changes(annual_data['netinc'].values)
        metrics['Operating Income'] = ltm['opinc'] / 1_000_000
        metrics['Total Operating Expense'] = ltm['opex'] / 1_000_000
        metrics['EBITDA'] = ltm['ebitda'] / 1_000_000
        metrics['EBITDA YoY %'] = calculate_yoy_changes(annual_data['ebitda'].values)
        metrics['EPS'] = ltm['eps']
        metrics['Interest Expense'] = ltm['intexp'] / 1_000_000

        # Cash Flow
        metrics['CFO'] = ltm['ncfo'] / 1_000_000
        metrics['CFO YoY %'] = calculate_yoy_changes(annual_data['ncfo'].values)
        metrics['FCF'] = ltm['fcf'] / 1_000_000
        metrics['FCF YoY %'] = calculate_yoy_changes(annual_data['fcf'].values)

        # Balance Sheet
        metrics['Equity'] = ltm['equity'] / 1_000_000
        metrics['Debt'] = ltm['debt'] / 1_000_000
        metrics['Total Assets'] = ltm['assets'] / 1_000_000
        metrics['Total Liabilities'] = ltm['liabilities'] / 1_000_000
        metrics['Net Working Capital'] = (ltm['assetsc'] - ltm['liabilitiesc']) / 1_000_000
        metrics['Cash Ratio'] = ltm['cashneq'] / ltm['liabilitiesc']
        metrics['Tangible Book Value'] = (ltm['assets'] - ltm['intangibles'] - ltm['liabilities']) / 1_000_000
        metrics['Leverage Ratio'] = ltm['assets'] / ltm['equity']
        metrics['Interest-Bearing Debt'] = (ltm['debtc'] + ltm['debtnc']) / 1_000_000
        metrics['Debt to Capital'] = ltm['debt'] / (ltm['debt'] + ltm['equity'])
        metrics['Cash to Debt'] = ltm['cashneq'] / ltm['debt']
        metrics['Net Debt to Total Assets'] = (ltm['debt'] - ltm['cashneq']) / ltm['assets']
        metrics['Working Capital Turnover'] = ltm['revenue'] / (ltm['assetsc'] - ltm['liabilitiesc'])

        # Margins
        metrics['GP Margin'] = ltm['grossmargin']
        metrics['EBITDA Margin'] = ltm['ebitdamargin']
        metrics['Net Margin'] = ltm['netmargin']
        metrics['Operating Margin'] = ltm['opinc'] / ltm['revenue']
        metrics['Free Cash Flow Margin'] = ltm['fcf'] / ltm['revenue']

        # Ratios
        metrics['Current Ratio'] = ltm['currentratio']
        metrics['Quick Ratio'] = (ltm['assetsc'] - ltm['inventory']) / ltm['liabilitiesc']
        metrics['Payout Ratio'] = ltm['payoutratio']
        metrics['D/E'] = ltm['debt'] / ltm['equity']
        metrics['Debt to EBITDA'] = ltm['debt'] / ltm['ebitda']
        metrics['Asset Turnover'] = ltm['assetturnover']
        metrics['Int Coverage'] = ltm['ebit'] / ltm['intexp']

        # Return Metrics
        metrics['ROA'] = ltm['roa']
        metrics['ROE'] = ltm['roe']
        metrics['ROIC'] = ltm['roic']
        
        # Extract scalar values from Series
        metrics = {k: float(v.iloc[0]) if isinstance(v, pd.Series) else v for k, v in metrics.items()}
        all_metrics[ticker] = metrics
    
    metrics_df = pd.DataFrame(all_metrics).T
    metrics_df = metrics_df.replace([np.inf, -np.inf], np.nan)
    
    return metrics_df.round(2)


########################################### Excel Formatting ###########################################


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
        'Revenue YoY %', 'Gross Profit YoY %', 'Net Income YoY %', 'EBITDA YoY %', 
        'CFO YoY %', 'FCF YoY %', 'Cash Ratio', 'Cash to Debt', 
        'Working Capital Turnover', 'GP Margin', 'EBITDA Margin', 'Net Margin', 
        'Operating Margin', 'Free Cash Flow Margin', 'ROA', 'ROE', 'ROIC'
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

def apply_conditional_formatting(sheet, metrics_df, start_row, start_col):
    for col_idx, metric_name in enumerate(metrics_df.columns):
        metric_group = get_metric_group(metric_name)
        
        if metric_group == 'Valuation Metrics':
            continue
            
        current_col = start_col + 2 + col_idx
        data_range = sheet.range(
            sheet.cells(start_row + 1, current_col), # Skip header row
            sheet.cells(start_row + len(metrics_df.index), current_col)
        )
        
        # Get values for this metric
        metric_values = metrics_df[metric_name].values
        
        # Apply formatting
        format_metrics(data_range, metric_values, metric_name)


########################################### Write to Excel ###########################################


def write_to_excel(sheet, metrics_df, companies_dict, start_row=3, start_col=5):
    """
    Write metrics DataFrame to Excel with companies as rows and metrics as columns.
    
    Args:
        sheet: xlwings worksheet object
        metrics_df: DataFrame with companies as index and metrics as columns
        companies_dict: Dictionary mapping sectors to lists of company tickers
        start_row: Starting row in Excel (default=3)
        start_col: Starting column in Excel (default=5)
    """
    # Write headers
    sheet.cells(start_row, start_col).value = "Sector"
    sheet.cells(start_row, start_col + 1).value = "Company"
    
    # Write metric headers
    for col_num, metric in enumerate(metrics_df.columns, start=start_col + 2):
        cell = sheet.cells(start_row, col_num)
        cell.value = metric

    # Write data
    current_row = start_row + 1
    
    for sector, companies in companies_dict.items():
        # Write sector name only for the first company
        sheet.cells(current_row, start_col).value = sector
        
        for i, company in enumerate(companies):
            # Write company ticker
            sheet.cells(current_row, start_col + 1).value = company
            
            # Write metric values
            if company in metrics_df.index:
                for col_num, value in enumerate(metrics_df.loc[company], start=start_col + 2):
                    cell = sheet.cells(current_row, col_num)
                    cell.value = value
                    
                    # Format numbers
                    metric_name = metrics_df.columns[col_num - (start_col + 2)]
                    if '%' in metric_name or metric_name in ['ROA', 'ROE', 'ROIC', 'GP Margin', 'EBITDA Margin', 'Net Margin', 'Operating Margin', 'Free Cash Flow Margin']:
                        cell.api.NumberFormat = "0.0%"
                    elif isinstance(value, (int, float)) and abs(value) >= 1000:
                        cell.api.NumberFormat = "#,##0.00"
            
            current_row += 1

    # Create Excel Table
    table_range = sheet.range(
        sheet.cells(start_row, start_col),
        sheet.cells(current_row - 1, start_col + len(metrics_df.columns) + 1)
    )
    
    # Convert range to Excel Table with specific style
    table = sheet.api.ListObjects.Add(1, table_range.api.Address, 0, 1)
    table.TableStyle = "TableStyleLight8"

    # Set column widths
    sheet.api.Columns(start_col).ColumnWidth = 20  # Category column
    sheet.api.Columns(start_col + 1).ColumnWidth = 15  # Company column
    
    # Auto-fit metric columns
    for col in range(start_col + 2, start_col + len(metrics_df.columns) + 2):
        sheet.api.Columns(col).AutoFit()

    # Apply conditional formatting
    apply_conditional_formatting(sheet, metrics_df, start_row, start_col)

    return sheet
    

########################################### Main ###########################################


def api_test():
    tickers = ['NVDA', 'AVGO', 'AMD']
    metrics = grab_sf1_data(tickers)
    print(metrics)

def main():
    spreadsheet_path = sys.argv[1]
    company_string = sys.argv[2]
    
    items = company_string.split(',')
    
    companies_dict = {}
    current_sector = None
    tickers = []
    for item in items:
        if any(c.islower() for c in item): # Sector name
            current_sector = item
            companies_dict[current_sector] = []
        else:
            tickers.append(item)
            if current_sector:
                companies_dict[current_sector].append(item)
    
    metrics = grab_sf1_data(tickers)
    
    wb = xw.books.active
    sheet = wb.sheets.active
    sheet = write_to_excel(sheet, metrics, companies_dict)
    wb.save(spreadsheet_path)


main()