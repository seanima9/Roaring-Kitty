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
    
def grab_data(tickers):  
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

        if len(data) >= 2:
            previous_shares = data['sharesbas'].iloc[-2]
            current_shares = data['sharesbas'].iloc[-1]
            bb_yield = (previous_shares - current_shares) / previous_shares
        else:
            bb_yield = np.nan


        sf2_data = ndl.get_table('SHARADAR/SF2', ticker=ticker, paginate=True)
        sf2_data['transactiondate'] = pd.to_datetime(sf2_data['transactiondate'])
        insider_buys = sf2_data[sf2_data['transactioncode'] == 'P']

        end_date = pd.to_datetime(sf2_data['transactiondate'].max())
        start_date = end_date - pd.DateOffset(months=12)

        insider_buys_count = insider_buys[
            (insider_buys['transactiondate'] > start_date) & 
            (insider_buys['transactiondate'] <= end_date)
        ].shape[0]

        sep_data = ndl.get_table('SHARADAR/SEP', ticker=ticker, paginate=True)
        sep_data['date'] = pd.to_datetime(sep_data['date'])
        latest_share_price = sep_data.sort_values('date').iloc[-1]['close']
        current_shares_outstanding = data['sharesbas'].iloc[-1] * data['sharefactor'].iloc[-1]
        current_market_cap = latest_share_price * current_shares_outstanding

        fx_conv_ltm = ltm['fxusd'].iloc[0]
        new_ev = current_market_cap + ((ltm['debt'].iloc[0] - ltm['cashneq'].iloc[0]) / fx_conv_ltm)

        if np.isnan(data['ebitda'].iloc[-1]): # LTM EBITDA is nan for Chinese stocks
            ltm_ebitda = data['ebitda'].iloc[-2]
        else:
            ltm_ebitda = data['ebitda'].iloc[-1]
            
        # Valuation Metrics
        metrics['TEV'] = new_ev / 1_000_000
        metrics['SP'] = latest_share_price
        metrics['TEV/EBITDA'] = new_ev / (ltm_ebitda / fx_conv_ltm)
        metrics['TEV/Rev'] = new_ev / (ltm['revenue'].iloc[0] / fx_conv_ltm)
        metrics['TEV/FCF'] = new_ev / (ltm['fcf'].iloc[0] / fx_conv_ltm)
        metrics['P/E'] = current_market_cap / (ltm['netinc'].iloc[0] / fx_conv_ltm)
        metrics['P/B'] = current_market_cap / (ltm['equity'].iloc[0] / fx_conv_ltm)
        metrics['EPS'] = ltm['eps'] / fx_conv_ltm

        # Income Statement
        metrics['Rev'] = (ltm['revenue'] / fx_conv_ltm) / 1_000_000
        metrics['Rev 3YCAGR'] = calculate_cagr((data['revenue'] / data['fxusd']).values)

        # Margins
        metrics['GP Marg'] = ltm['grossmargin']
        metrics['EBITDA Marg'] = ltm['ebitdamargin']
        metrics['Net Marg'] = ltm['netmargin']
        metrics['Op Marg'] = ltm['opinc'] / ltm['revenue']
        metrics['FCF Marg'] = ltm['fcf'] / ltm['revenue']

        # Shareholder Yield
        metrics['Div Yield'] = data['divyield']
        metrics['BB Yield'] = bb_yield
        metrics['Ins Buys'] = insider_buys_count

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
    sheet.range((1, 1), (1, sheet.api.Columns.Count)).color = (185, 216, 72)
    sheet.range((2, 1), (3, sheet.api.Columns.Count)).color = (0, 201, 192)
    
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
                    
                    # Apply formatting based on metric type
                    if 'Marg' in metric_name or 'Yield' in metric_name or 'CAGR' in metric_name:
                        cell.api.NumberFormat = "0%"
                    elif 'Ratio' in metric_name or '/' in metric_name or \
                    metric_name in ['ROA', 'ROE', 'ROIC', 'WC Turn', 'Asset Turn', 'EPS', 'SP']:
                        cell.api.NumberFormat = "0.00"
                    elif isinstance(value, (int, float)):
                        cell.api.NumberFormat = "#,##0"

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
    tickers = ['NVDA', 'TSM', 'BABA']
    metrics = grab_data(tickers)
    print(metrics)

def main():
    _ = sys.argv[1] # spreadsheet path
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
    
    metrics = grab_data(tickers)
    
    wb = xw.books.active
    sheet = wb.sheets.active
    write_to_excel(sheet, metrics, companies_dict, start_row=4, start_col=5)
    header_cell = sheet.cells(1, 5)
    header_cell.value = "RK Tracker"
    header_cell.api.Font.Size = 20


main()