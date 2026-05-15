"""
STOCK ENTERPRISE VENDING ANALYTICS PIPELINE
MULTICORE MAC STABLE VERSION
"""

import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
import numpy as np
from concurrent.futures import ThreadPoolExecutor
import multiprocessing as mp
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
import warnings
import os

warnings.filterwarnings("ignore")

CPU_COUNT = max(mp.cpu_count() - 1, 1)


def detect_column(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def classify(fill):
    if pd.isna(fill):
        return "UNKNOWN"
    if fill == 0:
        return "EMPTY"
    if fill <= 20:
        return "CRITICAL"
    if fill <= 40:
        return "LOW"
    if fill <= 70:
        return "NORMAL"
    return "FULL"


def load_stock_file(file):
    try:
        df = pd.read_excel(file, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        df["source_file"] = os.path.basename(file)
        return df
    except Exception:
        return pd.DataFrame()


def main():

    root = tk.Tk()
    root.withdraw()

    stock_files = filedialog.askopenfilenames(
        title="Select STOCK XLSX files",
        filetypes=[("Excel files", "*.xlsx")]
    )

    location_file = filedialog.askopenfilename(
        title="Select LOCATION XLSX file",
        filetypes=[("Excel files", "*.xlsx")]
    )

    OUTPUT_DIR = Path.home() / "Desktop" / "STOCK_OUTPUT"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    locations = pd.read_excel(location_file)
    locations.columns = [str(c).strip() for c in locations.columns]

    vm_col = "CZ automatu" if "CZ automatu" in locations.columns else locations.columns[0]
    locations["vm_id"] = locations[vm_col].astype(str)

    with ThreadPoolExecutor(max_workers=CPU_COUNT) as executor:
        frames = list(executor.map(load_stock_file, stock_files))

    stock = pd.concat([f for f in frames if len(f) > 0], ignore_index=True)

    machine_col = detect_column(stock, ["MachineNumber", "VM"])
    capacity_col = detect_column(stock, ["Product/Component capacity", "Capacity"])
    fill_col = detect_column(stock, ["Product/Component Fill quantity", "Fill quantity"])

    stock["vm_id"] = stock[machine_col].astype(str)
    stock["capacity"] = pd.to_numeric(stock[capacity_col], errors="coerce")
    stock["fill_qty"] = pd.to_numeric(stock[fill_col], errors="coerce")

    stock["fill_percent"] = np.where(
        stock["capacity"] > 0,
        (stock["fill_qty"] / stock["capacity"]) * 100,
        np.nan
    )

    stock["status"] = stock["fill_percent"].apply(classify)

    merged = stock.merge(locations, on="vm_id", how="left")

    machine_summary = merged.groupby("vm_id").agg(
        avg_fill=("fill_percent", "mean")
    ).reset_index()

    excel_path = OUTPUT_DIR / "STOCK_VENDING_ANALYTICS.xlsx"

    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        merged.to_excel(writer, sheet_name="Merged_Data", index=False)
        machine_summary.to_excel(writer, sheet_name="Machine_Summary", index=False)

    plt.figure(figsize=(12, 6))
    plt.hist(machine_summary["avg_fill"].dropna(), bins=20)
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / "fill_distribution.png")
    plt.close()

    doc = Document()
    doc.add_heading("STOCK EXECUTIVE REPORT", level=1)
    doc.add_paragraph(f"Machines: {merged['vm_id'].nunique():,}")
    doc.add_paragraph(f"Rows: {len(merged):,}")
    doc.add_paragraph(f"Threads: {CPU_COUNT}")
    doc.add_picture(str(OUTPUT_DIR / "fill_distribution.png"), width=Inches(6))
    doc.save(OUTPUT_DIR / "EXECUTIVE_REPORT.docx")

    print("PIPELINE FINISHED")


if __name__ == "__main__":
    mp.freeze_support()
    main()
