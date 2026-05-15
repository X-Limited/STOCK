# STOCK

Enterprise vending analytics and stock management platform.

## Features

- Multicore XLSX processing for macOS
- Automated vending analytics
- Risk scoring
- KPI calculations
- Forecasting-ready architecture
- Power BI exports
- Executive reports
- Parallel processing pipeline

## Structure

```text
STOCK/
├── src/
│   └── vending/
├── data/
├── outputs/
├── prompts/
├── requirements.txt
├── CHANGELOG.md
└── README.md
```

## Installation

```bash
pip install -r requirements.txt
```

## Run

```bash
python src/vending/vending_pipeline_multicore_mac.py
```
