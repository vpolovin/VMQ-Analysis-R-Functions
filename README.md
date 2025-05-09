# README: Stock Investor Pro Analysis Script
An R script that ranks and exports the top 50 stocks based on value and momentum indicators. 

## **📌 Overview**

This R script ranks stocks based on value and momentum indicators by performing a multi-factor investment analysis using data exported from AAII's Stock Investor Pro. It ranks stocks based on value, quality, profitability, and momentum metrics—then generates top stock lists and calculates their performance values and margins of safety. These outputs are based on my investing strategy that combines ​the research from ​James O'Shaughnessy's book is titled What Works on Wall Street and Joel Greenblatt's book that introduces the "Magic Formula" is titled The Little Book That Beats the Market.

## **🧠 Key Features**

- Loads and parses SIPro data files.

- Computes percentiles for valuation and quality metrics.

- Implements composite scores (e.g., CVF2, EQC, Magic Formula).

- Filters stocks based on quality and safety thresholds.

- Estimates performance value (PV) and Margin of Safety (MOS) using 10-year discounted valuation models.

- Outputs top-ranked lists as .xlsx files and R dump files.

## **📁 Inputs**

- Two .TXT files must be placed in the chosen directory:

- One containing raw SIPro data (no headers).

- One containing the key (column headers).

- The key file must contain header mappings used to label the data columns correctly.

  📌 Filenames must contain "DATA" and one of them must contain "Key" for the script to recognize them.

## **🧮 Metrics and Calculations**

**Value Factors (used in CVF2)**

- B/P

- E/P

- S/P

- EBITDA/EV

- CFPS/P

- Shareholder Yield

**Quality Factors (used in EQC)**

- Accruals to Assets

- Pct. Chg. in NOA

- Accruals to Avg. Assets

- CapEx to Deprec/Amort.

**Additional Filtering Metrics**

- ROE

- Asset Turnover

- Coverage Ratio

- External Financing

- Debt Change

- RSI (Momentum)

- Op. Income Growth

- F Score

**Magic Formula**

- ROC + EY (Earnings Yield)

**Composite Ranking**

- MQSP Rank: Mean of Momentum, Quality, Safety, and Profitability scores.

**Performance Value (PV)**

- Calculated as the 10-year discounted present value of the expected stock price based on ROE growth and average PE.

**🛠️ Dependencies**

- tcltk — for folder selection UI.

- XLConnect — to write .xlsx files.

- FinCal, stringi — for financial calculations.

  📦 Install packages (if not already installed): install.packages(c("tcltk", "XLConnect", "FinCal", "stringi"))

## **⚙️ How to Use**

1. Run Analyze(): Prompts you to select a folder with SIPro TXT files.

2. Checks and Loads: Validates that exactly 2 data files exist and identifies the key file.

3. Processes Data: Computes various financial ratios, assigns percentiles, and generates composite ranks.

4. Saves Outputs:

- .xlsx files for top 50 stocks under multiple strategies.
  
- .R script dump of selected objects for reloading later.

**📤 Output Files (Saved in Selected Folder)**

- MM.DD.YYYY VMQ.VEN.50.xlsx

- MM.DD.YYYY VMQ.DEC.50.xlsx

- MM.DD.YYYY MAGIC.VEN.50.xlsx

- MM.DD.YYYY Analysis.R

Each file contains the top 50 stock picks for a given strategy, quality checks, valuation metrics, and margin of safety.

## **⚠️ Notes & Limitations**

- Filtering logic may exclude some stocks based on sector or exchange (e.g., OTC, Financials, Utilities).

- The script assumes specific column names from Stock Investor Pro. These must match exactly.

- No visualization or plotting is currently included—this could be a future enhancement.
