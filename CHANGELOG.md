# CHANGELOG

## v0.4.0 — Parquet Safe Edition
### Added
- Automatic object -> string conversion before pyarrow export
- Stable Parquet export workflow for mixed vending schemas
- Improved enterprise-safe raw data export handling

### Fixed
- Fixed pyarrow ArrowInvalid crashes
- Fixed mixed datatype failures in MachineExternalID and similar columns
- Improved compatibility with Excel/SAP/vending mixed exports

---

## v0.3.0 — Excel Limit Safe Edition
### Added
- CSV sample export
- Raw parquet export workflow
- Excel summary-only export strategy

### Fixed
- Fixed Excel 1,048,576 row limit issue
- Removed oversized raw merged sheet export
- Improved handling for datasets over 1.5M rows

---

## v0.2.0 — Stable macOS Enterprise Pipeline
### Added
- Robust stock/location XLSX separation
- Automatic location/master file filtering
- Debug logging for schema validation
- Improved schema autodetection

### Fixed
- Replaced unstable ProcessPoolExecutor with ThreadPoolExecutor
- Fixed macOS multiprocessing stability problems
- Fixed Python 3.13 multiprocessing compatibility
- Fixed Machine column autodetection issues
- Added safer stock file validation

---

## v0.1.0 — Initial Enterprise Release
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
