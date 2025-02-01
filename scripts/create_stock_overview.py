import sys
import os
import numpy as np
import pandas as pd
import nasdaqdatalink as ndl
import xlwings as xw
import json

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.formatting_helpers import format_metrics

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


def calculate_rolling_cagr(values):
    cagr = pd.Series(index=values.index)
    
    for i in range(3, len(values)):
        end_value = values.iloc[i]
        start_value = values.iloc[i-3]
        cagr.iloc[i] = (end_value / start_value) ** (1/3) - 1
                
    return cagr

def grab_time_series_data(ticker):
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

    sf2_data = ndl.get_table('SHARADAR/SF2', ticker=ticker, paginate=True)
    sf2_data['transactiondate'] = pd.to_datetime(sf2_data['transactiondate'])
    insider_buys = sf2_data[sf2_data['transactioncode'] == 'P']

    def count_insider_buys(end_date):
        if end_date == data['calendardate'].max():
            end_date = pd.to_datetime('today')
        start_date = end_date - pd.DateOffset(months=12)

        count = insider_buys[
            (insider_buys['transactiondate'] > start_date) & 
            (insider_buys['transactiondate'] <= end_date)
        ].shape[0]
        return count
    
    sep_data = ndl.get_table('SHARADAR/SEP', ticker=ticker, paginate=True)
    sep_data['date'] = pd.to_datetime(sep_data['date'])
    latest_share_price = sep_data.sort_values('date').iloc[-1]['close']
    current_shares_outstanding = data['sharesbas'].iloc[-1]
    current_market_cap = latest_share_price * current_shares_outstanding

    # Valuation Metrics
    metrics['TEV'] = data['ev'] / 1_000_000
    metrics['Mkt Cap'] = data['marketcap'] / 1_000_000
    metrics['SP'] = latest_share_price
    metrics['TEV/EBITDA'] = data['ev'] / data['ebitda']
    metrics['TEV/Rev'] = data['ev'] / data['revenue']
    metrics['TEV/FCF'] = data['ev'] / data['fcf']
    metrics['P/E'] = data['pe']
    metrics['P/B'] = data['pb']
    metrics['EPS'] = data['eps']

    ltm_debt = data['debt'].iloc[-1]
    ltm_cash = data['cashneq'].iloc[-1]
    ltm_ebitda = data['ebitda'].iloc[-1]
    ltm_revenue = data['revenue'].iloc[-1]
    ltm_fcf = data['fcf'].iloc[-1]
    ltm_netinc = data['netinc'].iloc[-1]
    ltm_equity = data['equity'].iloc[-1]

    new_ev = current_market_cap + ltm_debt - ltm_cash

    metrics['TEV'].iat[-1] = new_ev / 1_000_000
    metrics['Mkt Cap'].iat[-1] = current_market_cap / 1_000_000
    metrics['TEV/EBITDA'].iat[-1] = new_ev / ltm_ebitda
    metrics['TEV/Rev'].iat[-1] = new_ev / ltm_revenue
    metrics['TEV/FCF'].iat[-1] = new_ev / ltm_fcf
    metrics['P/E'].iat[-1] = current_market_cap / ltm_netinc
    metrics['P/B'].iat[-1] = current_market_cap / ltm_equity

    # Income Statement
    metrics['Rev'] = data['revenue'] / 1_000_000
    metrics['Rev 3YCAGR'] = calculate_rolling_cagr(data['revenue'])
    metrics['GP'] = data['gp'] / 1_000_000
    metrics['Net Inc'] = data['netinc'] / 1_000_000
    metrics['Op Inc'] = data['opinc'] / 1_000_000
    metrics['EBITDA'] = data['ebitda'] / 1_000_000

    # Cash Flow
    metrics['CFO'] = data['ncfo'] / 1_000_000
    metrics['FCF'] = data['fcf'] / 1_000_000
    metrics['Op Exp'] = data['opex'] / 1_000_000
    metrics['CapEx'] = data['capex'] / 1_000_000
    metrics['Int Exp'] = data['intexp'] / 1_000_000

    # Margins
    metrics['GP Marg'] = data['grossmargin']
    metrics['EBITDA Marg'] = data['ebitdamargin']
    metrics['Net Marg'] = data['netmargin']
    metrics['Op Marg'] = data['opinc'] / data['revenue']
    metrics['FCF Marg'] = data['fcf'] / data['revenue']

    # Shareholder Yield
    metrics['BB Yield'] = (data['sharesbas'].shift(1) - data['sharesbas']) / data['sharesbas'].shift(1)
    metrics['Ins Buys'] = data['calendardate'].apply(count_insider_buys)

    # Balance Sheet
    metrics['Equity'] = data['equity'] / 1_000_000
    metrics['Debt'] = data['debt'] / 1_000_000
    metrics['Assets'] = data['assets'] / 1_000_000
    metrics['Liab'] = data['liabilities'] / 1_000_000
    metrics['TBV'] = (data['assets'] - data['intangibles'] - data['liabilities']) / 1_000_000

    # Solvency
    metrics['D/E'] = data['debt'] / data['equity']
    metrics['Debt/EBITDA'] = data['debt'] / data['ebitda']
    metrics['Cash Ratio'] = data['cashneq'] / data['liabilitiesc']
    metrics['Cash/Debt'] = data['cashneq'] / data['debt']
    metrics['Int Cov'] = data['ebit'] / data['intexp']

    # Liquidity
    metrics['Curr Ratio'] = data['currentratio']
    metrics['Quick Ratio'] = (data['assetsc'] - data['inventory']) / data['liabilitiesc']

    # Efficiency
    metrics['WC Turn'] = data['revenue'] / (data['assetsc'] - data['liabilitiesc'])
    metrics['Asset Turn'] = data['assetturnover']

    # Profitability
    metrics['ROA'] = data['roa']
    metrics['ROE'] = data['roe']
    metrics['ROIC'] = data['roic']

    metrics_df = pd.DataFrame(metrics)
    metrics_df = metrics_df.replace([np.inf, -np.inf], np.nan)
    metrics_df.index = data['year']
    
    return metrics_df.round(2)


def apply_conditional_formatting(sheet, metrics_df, start_row, start_col):
    # Work with transposed data to match Excel layout
    transposed_metrics = metrics_df.transpose()
    
    for row_idx, metric_name in enumerate(transposed_metrics.index):    
        row_values = transposed_metrics.loc[metric_name].values
        
        current_row = start_row + 1 + row_idx  # +1 to skip header row
        data_range = sheet.range((current_row, start_col + 2),  # +2 to skip category and metric name columns
                               (current_row, start_col + 2 + len(row_values) - 1))
        
        format_metrics(data_range, row_values, metric_name)


def write_to_excel(sheet, metrics, start_row=4, start_col=5):
    years = sorted([idx for idx in metrics.index if idx != 'LTM'])
    max_years_for_data = 15
    if len(years) > max_years_for_data:
        years = years[-max_years_for_data:]
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
                
                metric_cell = sheet.cells(current_row, start_col + 1)
                metric_cell.value = metric_name
                
                if metric_cell.api.Comment is not None:
                    metric_cell.api.Comment.Delete()
                metric_cell.api.AddComment(description)
                metric_cell.api.Comment.Visible = False
                
                for col_num, value in enumerate(transposed_metrics.loc[metric_name], start=start_col + 2):
                    cell = sheet.cells(current_row, col_num)
                    cell.value = value
                    
                    if 'CAGR' in metric_name or 'Yield' in metric_name:
                        cell.api.NumberFormat = "0.0%"
                    elif isinstance(value, (int, float)) and abs(value) >= 1000:
                        cell.api.NumberFormat = "#,##0.00"
                    
                current_row += 1

        if i < len(METRIC_GROUPS) - 1:
            border_range = sheet.range(
                sheet.cells(group_start_row, start_col),
                sheet.cells(current_row - 1, start_col + len(years) + 2)
            )
            border_range.api.Borders(9).Weight = 2

    table_range = sheet.range(
    sheet.cells(start_row, start_col),
    sheet.cells(current_row - 1, start_col + len(years) + 2)
    )

    for i, row in enumerate(table_range.rows):
        row_number = start_row + i

        if row_number == start_row:
            row.color = (180, 180, 180)  # Dark grey
        elif row_number == start_row + 1:
            pass
        elif (row_number - start_row) % 2 == 0:
            row.color = (217, 217, 217)  # Light grey
    
    for col in range(start_col, start_col + len(years) + 2):
        sheet.api.Columns(col).AutoFit()

    apply_conditional_formatting(sheet, metrics, start_row, start_col)


def api_test():
    data = grab_time_series_data('NVDA')
    print(data)

def main():
    spreadsheet_path = sys.argv[1]
    ticker = sys.argv[2]
    metrics = grab_time_series_data(ticker)

    wb = xw.books.active
    sheet = wb.sheets.active
    write_to_excel(sheet, metrics, start_row=4, start_col=5)
    header_cell = sheet.cells(1, 5)
    header_cell.value = f"{ticker} Overview"
    header_cell.api.Font.Size = 36
    header_cell.api.Font.Bold = True


main()