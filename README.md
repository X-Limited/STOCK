# STOCK

[![CI](https://github.com/X-Limited/STOCK/actions/workflows/ci.yml/badge.svg)](https://github.com/X-Limited/STOCK/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Code style: ruff](https://img.shields.io/badge/code%20style-ruff-000000.svg)](https://github.com/astral-sh/ruff)

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

## License

[MIT](LICENSE) © 2026 X-Limited
