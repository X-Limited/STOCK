# CHANGELOG

## 2026-05-15

### Added

- Initial multicore macOS vending analytics pipeline
- Parallel XLSX processing
- KPI calculations
- Risk scoring
- Excel exports
- Executive DOCX reports
- Histogram chart export
- File picker support for macOS
- GitHub-ready project structure

### Fixed

- Fixed macOS multiprocessing stability issues
- Replaced unstable ProcessPool with ThreadPoolExecutor
- Added robust case-insensitive column detection
- Added automatic filtering of location/master XLSX files
- Added safer stock file validation
- Added debug logging for detected columns
- Fixed Machine column autodetection failures
- Improved stability for Python 3.13 on macOS
