# Stock-financials

Automated generation of key financial indicators for stock analysis

### Current Issues

- Quarterly earnings dates are different for each stock, so the order of the quarters needs to be implemented to correctly retrieve their information
- `yfinance` API gives quarterly results up to 4 quarters back, but not further than this
