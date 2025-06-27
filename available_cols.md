# Sharadar Financial Data Columns Reference

## Metadata & Identifiers
- **ticker** - Stock ticker symbol
- **dimension** - Data dimension (ART, ARQ, ARY, MRT, MRQ, MRY)
- **calendardate** - Calendar date of the data
- **datekey** - Date key for the record
- **reportperiod** - Reporting period end date
- **fiscalperiod** - Fiscal period (Q1, Q2, Q3, Q4, FY)
- **lastupdated** - Last updated timestamp
- **sharefactor** - Share adjustment factor for stock splits/dividends

## Income Statement
- **revenue** - Total revenue
- **revenueusd** - Total revenue in USD
- **cor** - Cost of revenue (Cost of goods sold)
- **gp** - Gross profit
- **grossmargin** - Gross profit margin (%)
- **opex** - Operating expenses
- **sgna** - Selling, general & administrative expenses
- **rnd** - Research & development expenses
- **depamor** - Depreciation & amortization
- **opinc** - Operating income
- **intexp** - Interest expense
- **ebt** - Earnings before taxes
- **taxexp** - Tax expense
- **netinc** - Net income
- **netinccmn** - Net income available to common shareholders
- **netinccmnusd** - Net income available to common (USD)
- **netincdis** - Net income from discontinued operations
- **netincnci** - Net income attributable to non-controlling interests
- **ebit** - Earnings before interest and taxes
- **ebitda** - Earnings before interest, taxes, depreciation & amortization
- **ebitdausd** - EBITDA in USD
- **ebitusd** - EBIT in USD
- **ebitdamargin** - EBITDA margin (%)
- **netmargin** - Net profit margin (%)
- **consolinc** - Consolidated income
- **sbcomp** - Stock-based compensation

## Balance Sheet

### Assets
- **assets** - Total assets
- **assetsavg** - Average total assets
- **assetsc** - Current assets
- **assetsnc** - Non-current assets
- **cashneq** - Cash and cash equivalents
- **cashnequsd** - Cash and cash equivalents (USD)
- **receivables** - Accounts receivable
- **inventory** - Inventory
- **investments** - Total investments
- **investmentsc** - Current investments
- **investmentsnc** - Non-current investments
- **intangibles** - Intangible assets
- **ppnenet** - Property, plant & equipment (net)
- **tangibles** - Tangible assets
- **taxassets** - Tax assets

### Liabilities
- **liabilities** - Total liabilities
- **liabilitiesc** - Current liabilities
- **liabilitiesnc** - Non-current liabilities
- **debt** - Total debt
- **debtc** - Current debt
- **debtnc** - Non-current debt
- **debtusd** - Total debt in USD
- **payables** - Accounts payable
- **deferredrev** - Deferred revenue
- **taxliabilities** - Tax liabilities
- **deposits** - Customer deposits

### Equity
- **equity** - Total shareholders' equity
- **equityavg** - Average shareholders' equity
- **equityusd** - Total shareholders' equity (USD)
- **retearn** - Retained earnings
- **accoci** - Accumulated other comprehensive income

## Cash Flow Statement

### Operating Cash Flow
- **ncfo** - Net cash flow from operations
- **ncfbus** - Net cash flow from business activities

### Investing Cash Flow
- **ncfi** - Net cash flow from investing activities
- **ncfinv** - Net cash flow from investments
- **capex** - Capital expenditures

### Financing Cash Flow
- **ncff** - Net cash flow from financing activities
- **ncfdebt** - Net cash flow from debt activities
- **ncfdiv** - Net cash flow from dividends
- **ncfcommon** - Net cash flow from common equity

### Total Cash Flow
- **ncf** - Net cash flow (total)
- **ncfx** - Effect of exchange rate changes on cash
- **fcf** - Free cash flow

## Market Data & Valuation Metrics
- **price** - Share price
- **marketcap** - Market capitalization
- **ev** - Enterprise value
- **shareswa** - Weighted average shares outstanding
- **shareswadil** - Weighted average shares outstanding (diluted)
- **sharesbas** - Basic shares outstanding

## Per Share Metrics
- **eps** - Earnings per share (basic)
- **epsdil** - Earnings per share (diluted)
- **epsusd** - Earnings per share in USD
- **dps** - Dividends per share
- **bvps** - Book value per share
- **tbvps** - Tangible book value per share
- **fcfps** - Free cash flow per share
- **sps** - Sales per share

## Financial Ratios

### Profitability Ratios
- **roa** - Return on assets
- **roe** - Return on equity
- **roic** - Return on invested capital
- **ros** - Return on sales

### Liquidity Ratios
- **currentratio** - Current ratio
- **workingcapital** - Working capital

### Leverage Ratios
- **de** - Debt-to-equity ratio

### Valuation Ratios
- **pe** - Price-to-earnings ratio
- **pe1** - Forward price-to-earnings ratio
- **pb** - Price-to-book ratio
- **ps** - Price-to-sales ratio
- **ps1** - Forward price-to-sales ratio
- **evebit** - Enterprise value to EBIT
- **evebitda** - Enterprise value to EBITDA

### Other Ratios
- **assetturnover** - Asset turnover ratio
- **payoutratio** - Dividend payout ratio
- **divyield** - Dividend yield

## Calculated Fields
- **invcap** - Invested capital
- **invcapavg** - Average invested capital
- **fxusd** - Foreign exchange rate to USD
- **prefdivis** - Preferred dividends