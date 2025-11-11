# Stock Portfolio Analysis - VBA Excel

Comparative analysis of 5 international stocks vs MSCI World (2018-2023)

## About

This project analyzes the performance of 5 major stocks over 5 years to identify the best investment opportunity. All calculations are automated using VBA macros in Excel.

## Stocks Analyzed

- **NVIDIA** (semiconductors/AI)
- **LVMH** (luxury)
- **Alibaba** (e-commerce)
- **Novo Nordisk** (pharmaceutical)
- **United Health** (healthcare)

Benchmark: **MSCI World** index

## What I Did

The project includes 11 automated VBA macros that calculate:

- Daily returns with color coding (green/red)
- 5-year cumulative and annualized performance
- Annual performance for each year (2019-2023)
- Volatility (annualized)
- Sharpe Ratio (risk-adjusted return)
- Beta (market sensitivity)
- Best and worst trading days

## Key Results

| Stock | 5-Year Return | Annualized | Sharpe | Beta |
|-------|--------------|------------|--------|------|
| NVIDIA | +1,395% | +72% | 1.38 | 1.25 |
| Novo Nordisk | +427% | +39% | 1.38 | 0.85 |
| LVMH | +198% | +24% | 0.82 | 0.75 |
| United Health | +128% | +18% | 0.61 | 0.68 |
| MSCI World | +74% | +12% | 0.59 | 1.00 |
| Alibaba | -43% | -11% | -0.23 | 0.92 |

**Winner**: NVIDIA outperformed all other stocks, driven by the AI boom.

## Technical Details

- **Language**: VBA (Visual Basic for Applications)
- **Platform**: Microsoft Excel
- **Data**: Daily adjusted closing prices (1,259 trading days)
- **Period**: Dec 2018 - Dec 2023

## How to Use

1. Download the Excel file
2. Enable macros when opening
3. Go to Developer > Macros
4. Run any macro from qst1 to qst11
5. View results in the comparison table

## Skills Demonstrated

- VBA programming (loops, functions, error handling)
- Financial analysis (performance, risk, ratios)
- Data processing and automation
- Statistical calculations

## Formulas

**Daily Return**: (Pt / Pt-1) - 1

**Annualized Performance**: (1 + Total Return)^(1/5) - 1

**Volatility**: Daily Std Dev × √252

**Sharpe Ratio**: (Return - Risk-free Rate) / Volatility

**Beta**: Covariance(Stock, Market) / Variance(Market)


## Author

Emilie Chen  


