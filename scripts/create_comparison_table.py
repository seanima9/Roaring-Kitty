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


def calculate_cagr(values):
    start_value = values[0]
    end_value = values[-1]
    n_years = len(values) - 1
    
    return ((end_value / start_value) ** (1/n_years) - 1)
    
def grab_sf1_data(tickers):  
    all_metrics = {}
    for ticker in tickers:
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
        data = pd.concat([data, ltm]).tail(4)
        
        # Valuation Metrics
        metrics['TEV'] = ltm['ev'] / 1_000_000
        metrics['TEV/EBITDA'] = ltm['ev'] / ltm['ebitda']
        metrics['TEV/Rev'] = ltm['ev'] / ltm['revenue']
        metrics['TEV/FCF'] = ltm['ev'] / ltm['fcf']
        metrics['P/B'] = ltm['pb']
        metrics['EPS'] = ltm['eps']

        # Income Statement
        metrics['Rev'] = ltm['revenue'] / 1_000_000
        metrics['Rev 3YCAGR'] = calculate_cagr(data['revenue'].values)

        # Margins
        metrics['GP Marg'] = ltm['grossmargin']
        metrics['EBITDA Marg'] = ltm['ebitdamargin']
        metrics['Net Marg'] = ltm['netmargin']
        metrics['Op Marg'] = ltm['opinc'] / ltm['revenue']
        metrics['FCF Marg'] = ltm['fcf'] / ltm['revenue']

        # Balance Sheet
        metrics['Assets'] = ltm['assets'] / 1_000_000
        metrics['Liab'] = ltm['liabilities'] / 1_000_000

        # Solvency
        metrics['D/E'] = ltm['debt'] / ltm['equity']
        metrics['Debt/EBITDA'] = ltm['debt'] / ltm['ebitda']
        metrics['Cash Ratio'] = ltm['cashneq'] / ltm['liabilitiesc']
        metrics['Cash/Debt'] = ltm['cashneq'] / ltm['debt']
        metrics['Int Cov'] = ltm['ebit'] / ltm['intexp']

        # Liquidity
        metrics['Curr Ratio'] = ltm['currentratio']
        metrics['Quick Ratio'] = (ltm['assetsc'] - ltm['inventory']) / ltm['liabilitiesc']

        # Efficiency
        metrics['WC Turn'] = ltm['revenue'] / (ltm['assetsc'] - ltm['liabilitiesc'])
        metrics['Asset Turn'] = ltm['assetturnover']

        # Profitability
        metrics['ROA'] = ltm['roa']
        metrics['ROE'] = ltm['roe']
        metrics['ROIC'] = ltm['roic']

        metrics = {k: float(v.iloc[0]) if isinstance(v, pd.Series) else v for k, v in metrics.items()}
        all_metrics[ticker] = metrics
    
    metrics_df = pd.DataFrame(all_metrics).T
    metrics_df = metrics_df.replace([np.inf, -np.inf], np.nan)
    
    return metrics_df.round(2)


def apply_conditional_formatting(sheet, metrics_df, start_row, start_col):
    for col_idx, metric_name in enumerate(metrics_df.columns):            
        current_col = start_col + 2 + col_idx
        data_range = sheet.range(
            sheet.cells(start_row + 1, current_col), # Skip header row
            sheet.cells(start_row + len(metrics_df.index), current_col)
        )
        
        metric_values = metrics_df[metric_name].values
        
        format_metrics(data_range, metric_values, metric_name)


def write_to_excel(sheet, metrics_df, companies_dict, start_row=4, start_col=5):
    sheet.cells(start_row, start_col).value = "Ticker"
    sheet.cells(start_row, start_col + 1).value = "Sector"
    
    for col_num, metric in enumerate(metrics_df.columns, start=start_col + 2):
        cell = sheet.cells(start_row, col_num)
        cell.value = metric

        description = None
        for group in METRIC_GROUPS:
            if metric in group['metrics']:
                description = group['metrics'][metric]
                break
        
        if not description:
            raise ValueError(f"Could not find description for metric {metric}")
        if cell.api.Comment is not None:
            cell.api.Comment.Delete()
        cell.api.AddComment(description)
        cell.api.Comment.Visible = False

    current_row = start_row + 1
    
    for sector, companies in companies_dict.items():
        for company in companies:
            sheet.cells(current_row, start_col).value = company
            sheet.cells(current_row, start_col + 1).value = sector
            
            if company in metrics_df.index:
                for col_num, value in enumerate(metrics_df.loc[company], start=start_col + 2):
                    cell = sheet.cells(current_row, col_num)
                    cell.value = value

                    metric_name = metrics_df.columns[col_num - (start_col + 2)]
                    if 'CAGR' in metric_name:
                        cell.api.NumberFormat = "0.0%"
                    elif isinstance(value, (int, float)) and abs(value) >= 1000:
                        cell.api.NumberFormat = "#,##0.00"

            current_row += 1

    table_range = sheet.range(
        sheet.cells(start_row, start_col),
        sheet.cells(current_row - 1, start_col + len(metrics_df.columns) + 1)
    )

    table = sheet.api.ListObjects.Add(1, table_range.api.Address, 0, 1)
    table.TableStyle = "TableStyleLight1"
    table_range.rows[0].color = (180, 180, 180)  # Dark grey
    
    for col in range(start_col, start_col + len(metrics_df.columns) + 2):
        sheet.api.Columns(col).AutoFit()
    
    sheet.api.Application.ActiveWindow.SplitRow = start_row
    sheet.api.Application.ActiveWindow.SplitColumn = start_col
    sheet.api.Application.ActiveWindow.FreezePanes = True

    apply_conditional_formatting(sheet, metrics_df, start_row, start_col)


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
    write_to_excel(sheet, metrics, companies_dict, start_row=4, start_col=5)
    header_cell = sheet.cells(1, 5)
    header_cell.value = "RK Tracker"
    header_cell.api.Font.Size = 36
    header_cell.api.Font.Bold = True


main()