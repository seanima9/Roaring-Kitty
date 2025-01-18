import sys
import os
import numpy as np
import pandas as pd
import nasdaqdatalink as ndl
import xlwings as xw
import json

config_path = os.path.join(os.path.dirname(__file__), 'config.json')
with open(config_path) as f:
    config = json.load(f)
    api_key = config['api_key']

ndl.ApiConfig.api_key = api_key

METRIC_GROUPS = [
    {
        'name': 'Valuation Metrics',
        'metrics': {
            'TEV': 'Total Enterprise Value in millions',
            'Market Cap': 'Market Capitalization in millions',
            'TEV / EBITDA': 'Enterprise Value to EBITDA ratio - measures company value relative to earnings',
            'TEV / Revenue': 'Enterprise Value to Revenue ratio - measures company value relative to sales',
            'TEV / FCF': 'Enterprise Value to Free Cash Flow ratio - measures company value relative to cash flow',
            'P/E': 'Price to Earnings ratio - measures stock price relative to earnings per share',
            'P/B': 'Price to Book ratio - measures stock price relative to book value per share'
        }
    },
    {
        'name': 'Income Statement',
        'metrics': {
            'Revenue': 'Total sales in millions',
            'Revenue %': 'Year-over-year revenue growth',
            'Gross Profit': 'Revenue minus cost of goods sold in millions',
            'GP %': 'Year-over-year gross profit growth',
            'Net Income': 'Total earnings in millions',
            'Net Income %': 'Year-over-year net income growth',
            'Operating Income': 'Profit from core business operations in millions',
            'Total Operating Expense': 'All operating costs excluding non-operating items in millions',
            'EBITDA': 'Earnings Before Interest, Taxes, Depreciation & Amortization in millions',
            'EBITDA %': 'Year-over-year EBITDA growth',
            'EPS': 'Earnings Per Share'
        }
    },
    {
        'name': 'Cash Flow',
        'metrics': {
            'CFO': 'Cash Flow from Operations in millions',
            'CFO %': 'Year-over-year operating cash flow growth',
            'FCF': 'Free Cash Flow in millions',
            'FCF %': 'Year-over-year free cash flow growth'
        }
    },
    {
        'name': 'Balance Sheet',
        'metrics': {
            'Equity': 'Total Shareholder Equity in millions',
            'Debt': 'Total Debt in millions',
            'Total Assets': 'Total Assets in millions',
            'Total Liabilities': 'Total Liabilities in millions',
            'Net Working Capital': 'Current Assets minus Current Liabilities in millions',
            'Cash Ratio': 'Cash and Cash Equivalents divided by Current Liabilities',
            'Tangible Book Value': 'Total Assets minus Intangibles and Liabilities in millions',
            'TBV Per Share': 'Tangible Book Value divided by Shares Outstanding',
            'Leverage Ratio': 'Total Assets divided by Total Equity',
            'Interest-Bearing Debt': 'Short-Term and Long-Term Debt in millions',
            'Debt to Capital': 'Total Debt divided by Total Debt plus Equity',
            'Cash to Debt': 'Cash and Cash Equivalents divided by Total Debt',
            'Net Debt to Total Assets': 'Net Debt divided by Total Assets',
            'Working Capital Turnover': 'Revenue divided by Net Working Capital'
        }
    },
    {
        'name': 'Margins',
        'metrics': {
            'GP Margin': 'Gross Profit as a percentage of Revenue',
            'EBITDA Margin': 'EBITDA as a percentage of Revenue',
            'Net Margin': 'Net Income as a percentage of Revenue',
            'Operating Margin': 'Operating Income as a percentage of Revenue',
            'Free Cash Flow Margin': 'Free Cash Flow as a percentage of Revenue'
        }
    },
    {
        'name': 'Ratios',
        'metrics': {
            'Current Ratio': 'Current Assets divided by Current Liabilities - measures short-term liquidity',
            'Quick Ratio': 'Current Assets minus Inventory divided by Current Liabilities - measures immediate liquidity',
            'Payout Ratio': 'Dividends divided by Net Income - measures dividend sustainability',
            'D/E': 'Debt to Equity ratio - measures financial leverage',
            'Debt to EBITDA': 'Total Debt divided by EBITDA - measures debt leverage relative to earnings',
            'Asset Turnover': 'Revenue divided by Average Total Assets - measures asset efficiency',
            'Int Coverage': 'EBIT divided by Interest Expense - measures ability to pay interest'
        }
    },
    {
        'name': 'Returns',
        'metrics': {
            'ROA': 'Return on Assets - measures profitability relative to total assets',
            'ROE': 'Return on Equity - measures profitability relative to shareholder equity',
            'ROIC': 'Return on Invested Capital - measures profitability relative to invested capital'
        }
    }
]

# Color constants for Excel conditional formatting
DARK_GREEN = (51, 153, 51)      # Strongest positive indication  
MED_GREEN = (102, 187, 102)     # Medium positive
LIGHT_GREEN = (144, 213, 144)   # Mild positive

DARK_RED = (204, 51, 51)        # Serious concern
MED_RED = (230, 102, 102)       # Moderate concern
LIGHT_RED = (247, 153, 153)     # Mild concern

YELLOW = (255, 217, 102)        # Used for borderline cases


########################################### Get Data ###########################################


def grab_data(ticker):
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
    metrics_df.index = data['year']
    
    return metrics_df.round(2)


########################################### Excel Formatting ###########################################


def calculate_percentiles(values):
    """Calculate percentiles for a numpy array of values"""
    return {
        10: np.percentile(values, 10),
        25: np.percentile(values, 25),
        50: np.percentile(values, 50),
        75: np.percentile(values, 75),
        90: np.percentile(values, 90)
    }

def get_metric_group(metric_name):
    """Helper function to identify which group a metric belongs to"""
    for group in METRIC_GROUPS:
        if metric_name in group['metrics']:
            return group['name']
    raise ValueError(f"Metric '{metric_name}' not found in any group")

def format_growth_metrics(range_obj, values):
    """Color by percentiles - higher is better"""
    percentiles = calculate_percentiles(values)
    for cell, value in zip(range_obj, values):
        if pd.notna(value):  # Only format if value exists
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

def format_margin_metrics(range_obj, values):
    """Format margin metrics - higher is better"""
    percentiles = calculate_percentiles(values)
    for cell, value in zip(range_obj, values):
        if pd.notna(value):
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

def format_cash_flow_metrics(range_obj, values):
    """Format cash flow metrics - higher is better"""
    percentiles = calculate_percentiles(values)
    for cell, value in zip(range_obj, values):
        if pd.notna(value):
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

def format_ratio_metrics(range_obj, values, metric_name):
    """Format ratio metrics - uses absolute thresholds for some metrics"""
    for cell, value in zip(range_obj, values):
        if pd.notna(value):
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
            elif metric_name == 'Int Coverage':
                if value >= 4:
                    cell.color = DARK_GREEN
                elif value >= 3:
                    cell.color = MED_GREEN
                elif value >= 2:
                    cell.color = YELLOW
                elif value >= 1:
                    cell.color = LIGHT_RED
                else:
                    cell.color = DARK_RED
            elif metric_name in ['D/E', 'Debt to EBITDA']:
                # Lower is better for these metrics
                percentiles = calculate_percentiles(values)
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

def format_return_metrics(range_obj, values):
    """Format return metrics - higher is better"""
    percentiles = calculate_percentiles(values)
    for cell, value in zip(range_obj, values):
        if pd.notna(value):
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

def apply_conditional_formatting(sheet, metrics_df, start_row, start_col):
    """
    Apply conditional formatting to financial metrics.
    Expects metrics_df to be in the same orientation as the Excel sheet
    (metrics as rows, years as columns)
    """
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


def app():
    spreadsheet_path = sys.argv[1]
    ticker = sys.argv[2]
    metrics = grab_data(ticker)

    wb = xw.books.active
    sheet = wb.sheets.active
    sheet = write_to_excel(sheet, metrics)
    wb.save(spreadsheet_path)


app()