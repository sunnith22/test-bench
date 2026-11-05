#!/usr/bin/env python3
"""
Multi-Engine Variant Test Bench Data Processor & Graph Generator

Scans directory hierarchy:
  <root>/<variant>/<year>_W<week>_<ref>_<cell>_<engine>_<variant>_<serial>_<purpose>.xlsx

Fetches and aggregates all engine data for any selected variant (e.g., 202, 203, 319, 328),
and generates multi-engine comparison graphs (e.g. RPM vs Temp for all engines in Variant 202).
"""

import os
import re
import json
import glob
import math
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt

# -----------------------------------------------------------------------------
# Filename Parsing Regex
# Matches: 2025_W26_54_TC-1_TNGA_202_NE35384_Regular+Dis assy.xlsx
# -----------------------------------------------------------------------------
FILENAME_REGEX = re.compile(
    r"^(?P<year>\d{4})_W(?P<week>\d+?)_(?P<ref>\d+?)_(?P<cell>[^_]+)_(?P<engine_type>[^_]+)_(?P<variant>[^_]+)_(?P<serial>[^_]+)_(?P<purpose>.+?)(?:\s*\(\d+\))?\.xlsx$",
    re.IGNORECASE
)

# Standard Column Mapping for the 3-row header format
COLUMN_KEYS = [
    "rpm", "acc", "k", "power", "corrected_power", "dyno_torque", "corrected_torque",
    "bsfc", "fuel_consumption", "water_temp_in", "water_temp_out", "oil_temp",
    "oil_press", "afr", "in_air_press", "in_air_temp", "ex_gas_temp", "back_press",
    "td_temp", "tw_temp", "fuel_press", "crank_press", "atm_press"
]

METRIC_LABELS = {
    "ex_gas_temp": {"name": "Exhaust Gas Temperature", "unit": "℃", "key": "ex_gas_temp"},
    "water_temp_out": {"name": "Coolant Outlet Temperature", "unit": "℃", "key": "water_temp_out"},
    "water_temp_in": {"name": "Coolant Inlet Temperature", "unit": "℃", "key": "water_temp_in"},
    "oil_temp": {"name": "Engine Oil Sump Temperature", "unit": "℃", "key": "oil_temp"},
    "in_air_temp": {"name": "Intake Air Temperature", "unit": "℃", "key": "in_air_temp"},
    "power": {"name": "Dyno Power", "unit": "kW", "key": "power"},
    "corrected_power": {"name": "Corrected Power", "unit": "kW", "key": "corrected_power"},
    "dyno_torque": {"name": "Dyno Torque", "unit": "N·m", "key": "dyno_torque"},
    "corrected_torque": {"name": "Corrected Torque", "unit": "N·m", "key": "corrected_torque"},
    "bsfc": {"name": "Fuel Consumption Ratio (BSFC)", "unit": "g/kW·h", "key": "bsfc"},
    "fuel_consumption": {"name": "Fuel Flow Rate", "unit": "g/s", "key": "fuel_consumption"},
    "oil_press": {"name": "Oil Pressure", "unit": "kPa", "key": "oil_press"},
    "afr": {"name": "Air Fuel Ratio (AFR)", "unit": "-", "key": "afr"},
    "back_press": {"name": "Exhaust Back Pressure", "unit": "kPa", "key": "back_press"},
    "fuel_press": {"name": "Fuel Rail Pressure", "unit": "kPa", "key": "fuel_press"},
    "crank_press": {"name": "Crankcase Blow-by Pressure", "unit": "kPa", "key": "crank_press"},
    "atm_press": {"name": "Atmospheric Pressure", "unit": "kPa", "key": "atm_press"}
}

def parse_engine_excel(filepath):
    """
    Parses a single test bench Excel file, extracting test metadata and tabular RPM channels.
    """
    fname = os.path.basename(filepath)
    match = FILENAME_REGEX.match(fname)
    meta = {}
    if match:
        meta = match.groupdict()
    else:
        # Fallback if non-standard name
        meta = {
            "year": "Unknown", "week": "W00", "ref": "00", "cell": "TC-1",
            "engine_type": "TNGA", "variant": "Unknown", "serial": fname, "purpose": "Test"
        }

    meta["filename"] = fname
    meta["filepath"] = filepath

    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # Rows 1 to 3 are headers; data starts at row 4
    data_rows = []
    metadata_block = {}

    for row_idx in range(4, ws.max_row + 1):
        c1_val = ws.cell(row=row_idx, column=1).value
        # If we hit an empty row or metadata label row
        if c1_val is None or str(c1_val).strip() == "":
            continue
        try:
            rpm_val = float(c1_val)
        except ValueError:
            # Metadata rows like Date:, Testing ref no:, etc.
            label_val = str(c1_val).strip()
            val2 = ws.cell(row=row_idx, column=2).value or ws.cell(row=row_idx, column=3).value
            if label_val and val2:
                metadata_block[label_val.rstrip(":")] = str(val2).strip()
            continue

        row_dict = {}
        for col_idx, key in enumerate(COLUMN_KEYS, start=1):
            val = ws.cell(row=row_idx, column=col_idx).value
            try:
                row_dict[key] = float(val) if val is not None else 0.0
            except (ValueError, TypeError):
                row_dict[key] = 0.0
        data_rows.append(row_dict)

    wb.close()
    meta["metadata_block"] = metadata_block
    meta["data"] = data_rows
    meta["num_points"] = len(data_rows)
    return meta

def scan_variants_directory(root_dir):
    """
    Scans root directory for variant subfolders and indexes all engine Excel files.
    """
    variants = {}
    if not os.path.exists(root_dir):
        return variants

    # Find directories
    for item in os.listdir(root_dir):
        item_path = os.path.join(root_dir, item)
        if os.path.isdir(item_path):
            variant_name = item
            excel_files = glob.glob(os.path.join(item_path, "*.xlsx"))
            # Skip temporary excel lock files (~$...)
            excel_files = [f for f in excel_files if not os.path.basename(f).startswith("~$")]
            if excel_files:
                engines_data = []
                for f in sorted(excel_files):
                    try:
                        p_meta = parse_engine_excel(f)
                        engines_data.append(p_meta)
                    except Exception as e:
                        print(f"Warning: Failed to parse {f}: {e}")
                if engines_data:
                    variants[variant_name] = engines_data
    return variants

def get_variant_aggregate(variant_engines, metric_key="ex_gas_temp"):
    """
    Computes RPM-aligned series for all engines of a variant, plus mean and std deviation.
    """
    if not variant_engines:
        return None

    # Get common RPM list
    rpms = variant_engines[0]["data"]
    rpm_list = [r["rpm"] for r in rpms]

    engine_curves = []
    matrix = []  # shape: (n_engines, n_rpms)

    for eng in variant_engines:
        rpm_map = {r["rpm"]: r.get(metric_key, 0.0) for r in eng["data"]}
        curve = [rpm_map.get(rpm, np.nan) for rpm in rpm_list]
        engine_curves.append({
            "serial": eng["serial"],
            "ref": eng["ref"],
            "week": eng["week"],
            "cell": eng["cell"],
            "date": eng["metadata_block"].get("Date", eng["year"]),
            "purpose": eng["purpose"],
            "label": f"{eng['serial']} ({eng['week']} Ref {eng['ref']})",
            "values": curve
        })
        matrix.append(curve)

    arr = np.array(matrix, dtype=float)
    mean_curve = np.nanmean(arr, axis=0).round(1).tolist()
    min_curve = np.nanmin(arr, axis=0).round(1).tolist()
    max_curve = np.nanmax(arr, axis=0).round(1).tolist()
    std_curve = np.nanstd(arr, axis=0).round(2).tolist()

    return {
        "rpms": rpm_list,
        "engine_curves": engine_curves,
        "mean_curve": mean_curve,
        "min_curve": min_curve,
        "max_curve": max_curve,
        "std_curve": std_curve
    }

def generate_multi_engine_plots(variant_name, variant_engines, output_dir):
    """
    Generates high-resolution multi-engine comparison charts for the variant.
    """
    os.makedirs(output_dir, exist_ok=True)
    plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')

    # Color palette
    color_palette = [
        "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
        "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
        "#e41a1c", "#377eb8", "#4daf4a", "#984ea3"
    ]

    # -------------------------------------------------------------
    # 1. RPM vs Exhaust Gas Temp (All Engines in Variant)
    # -------------------------------------------------------------
    agg_egt = get_variant_aggregate(variant_engines, "ex_gas_temp")
    if agg_egt:
        fig, ax = plt.subplots(figsize=(13, 7.5), dpi=300)
        rpms = agg_egt["rpms"]

        # Plot individual engines
        for idx, eng in enumerate(agg_egt["engine_curves"]):
            c = color_palette[idx % len(color_palette)]
            ax.plot(rpms, eng["values"], label=eng["label"], color=c, linewidth=2.0, alpha=0.85, marker='o', markersize=3)

        # Plot Mean Curve and Tolerance Band
        ax.plot(rpms, agg_egt["mean_curve"], label=f"Variant {variant_name} Mean", color="#000000", linewidth=3.5, linestyle="--")
        ax.fill_between(rpms, agg_egt["min_curve"], agg_egt["max_curve"], color="#7f8c8d", alpha=0.18, label="Engine-to-Engine Range (Min-Max)")

        ax.set_title(f"Engine Exhaust Gas Temperature Comparison: RPM vs Ex-gas Temp (Variant {variant_name})\nAll Tested Engines Overlay ({len(agg_egt['engine_curves'])} Dyno Runs)",
                     fontsize=14, fontweight='bold', pad=14)
        ax.set_xlabel("Engine Speed (RPM)", fontsize=12, fontweight='bold', labelpad=8)
        ax.set_ylabel("Exhaust Gas Temperature (℃)", fontsize=12, fontweight='bold', labelpad=8)
        ax.legend(bbox_to_anchor=(1.02, 1), loc="upper left", frameon=True, fontsize=9.5)
        plt.tight_layout()

        p1_path = os.path.join(output_dir, f"rpm_vs_temp_variant_{variant_name}_all_engines.png")
        plt.savefig(p1_path, dpi=300)
        plt.close()
        print(f"Saved: {p1_path}")

    # -------------------------------------------------------------
    # 2. 4-Panel Thermal Overview (Exhaust, Coolant Out, Oil Temp, In-Air)
    # -------------------------------------------------------------
    temp_metrics = [
        ("ex_gas_temp", "Exhaust Gas Temp (℃)", "#C0392B"),
        ("water_temp_out", "Coolant Outlet Temp (℃)", "#2980B9"),
        ("oil_temp", "Oil Sump Temp (℃)", "#D35400"),
        ("in_air_temp", "Intake Air Temp (℃)", "#27AE60")
    ]

    fig, axes = plt.subplots(2, 2, figsize=(16, 12), dpi=250)
    axes = axes.flatten()

    for idx, (m_key, title, base_color) in enumerate(temp_metrics):
        ax = axes[idx]
        agg_m = get_variant_aggregate(variant_engines, m_key)
        if not agg_m:
            continue
        rpms = agg_m["rpms"]

        for e_idx, eng in enumerate(agg_m["engine_curves"]):
            c = color_palette[e_idx % len(color_palette)]
            ax.plot(rpms, eng["values"], color=c, alpha=0.7, linewidth=1.6, label=eng["label"] if idx == 0 else None)

        ax.plot(rpms, agg_m["mean_curve"], color="#000000", linewidth=3.0, linestyle="--", label="Mean" if idx == 0 else None)
        ax.fill_between(rpms, agg_m["min_curve"], agg_m["max_curve"], color=base_color, alpha=0.15)
        ax.set_title(f"{title} vs RPM (Variant {variant_name})", fontsize=12, fontweight='bold', pad=8)
        ax.set_xlabel("Engine Speed (RPM)", fontsize=10)
        ax.set_ylabel(title.split("(")[-1].replace(")", ""), fontsize=10, fontweight='bold')
        ax.grid(True, linestyle=":", alpha=0.6)

    axes[0].legend(loc="upper left", fontsize=7.5, framealpha=0.9)
    plt.suptitle(f"Thermal Profile Suite vs RPM Across All Engines (Variant {variant_name})\nCooling, Lubrication, Air Induction & Exhaust Thermal Response",
                 fontsize=15, fontweight='bold', y=0.995)
    plt.tight_layout()
    p2_path = os.path.join(output_dir, f"rpm_vs_all_temperatures_variant_{variant_name}.png")
    plt.savefig(p2_path, dpi=250)
    plt.close()
    print(f"Saved: {p2_path}")

    # -------------------------------------------------------------
    # 3. RPM vs Power & Torque (All Engines)
    # -------------------------------------------------------------
    agg_p = get_variant_aggregate(variant_engines, "corrected_power")
    agg_t = get_variant_aggregate(variant_engines, "corrected_torque")
    if agg_p and agg_t:
        fig, ax1 = plt.subplots(figsize=(13, 7.5), dpi=300)
        ax2 = ax1.twinx()
        rpms = agg_p["rpms"]

        for idx, eng in enumerate(agg_p["engine_curves"]):
            ax1.plot(rpms, eng["values"], color="#D9381E", alpha=0.45, linewidth=1.5)
        ax1.plot(rpms, agg_p["mean_curve"], color="#900C3F", linewidth=3.5, label="Mean Corrected Power (kW)")

        for idx, eng in enumerate(agg_t["engine_curves"]):
            ax2.plot(rpms, eng["values"], color="#1F77B4", alpha=0.45, linewidth=1.5)
        ax2.plot(rpms, agg_t["mean_curve"], color="#002060", linewidth=3.5, label="Mean Corrected Torque (N·m)")

        ax1.set_xlabel("Engine Speed (RPM)", fontsize=12, fontweight='bold')
        ax1.set_ylabel("Corrected Power (kW)", fontsize=12, fontweight='bold', color="#D9381E")
        ax2.set_ylabel("Corrected Torque (N·m)", fontsize=12, fontweight='bold', color="#1F77B4")
        ax1.tick_params(axis='y', labelcolor="#D9381E")
        ax2.tick_params(axis='y', labelcolor="#1F77B4")

        lines = ax1.get_lines()[-1:] + ax2.get_lines()[-1:]
        ax1.legend(lines, [l.get_label() for l in lines], loc="center left", fontsize=10, frameon=True)

        plt.title(f"Engine Dynamometer Power & Torque Spread (Variant {variant_name})\nOverlay of {len(variant_engines)} Engine Full-Load Calibration Runs",
                  fontsize=14, fontweight='bold', pad=14)
        plt.tight_layout()
        p3_path = os.path.join(output_dir, f"rpm_vs_power_torque_variant_{variant_name}.png")
        plt.savefig(p3_path, dpi=300)
        plt.close()
        print(f"Saved: {p3_path}")

def generate_comparison_report(variant_dict, output_excel_path):
    """
    Builds the 'comparision_report.xlsx' workbook aggregating all variants and engines.
    """
    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    # 1. Summary Sheet
    ws1 = wb.active
    ws1.title = "Variant_Comparison_Summary"
    ws1.views.sheetView[0].showGridLines = True

    summary_headers = [
        "Variant", "Engine Serial", "Week", "Test Ref", "Cell", "Test Date", "Testing Purpose",
        "Peak Power (kW)", "Peak Torque (N·m)", "Best BSFC (g/kW·h)", "Max EGT (℃)",
        "Water Out Max (℃)", "Oil Temp Max (℃)", "Oil Press Max (kPa)"
    ]
    for c_i, h in enumerate(summary_headers, start=1):
        cell = ws1.cell(row=1, column=c_i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    row_count = 2
    for var_name, eng_list in variant_dict.items():
        for eng in eng_list:
            d = eng["data"]
            pk_p = max(r.get("corrected_power", 0.0) for r in d)
            pk_t = max(r.get("corrected_torque", 0.0) for r in d)
            min_b = min(r.get("bsfc", 999.0) for r in d)
            max_egt = max(r.get("ex_gas_temp", 0.0) for r in d)
            max_wout = max(r.get("water_temp_out", 0.0) for r in d)
            max_oil = max(r.get("oil_temp", 0.0) for r in d)
            max_op = max(r.get("oil_press", 0.0) for r in d)

            vals = [
                var_name, eng["serial"], eng["week"], eng["ref"], eng["cell"],
                eng["metadata_block"].get("Date", eng["year"]), eng["purpose"],
                pk_p, pk_t, min_b, max_egt, max_wout, max_oil, max_op
            ]
            for c_i, val in enumerate(vals, start=1):
                cell = ws1.cell(row=row_count, column=c_i, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="right" if isinstance(val, (int, float)) else "left")
            row_count += 1

    for col in ws1.columns:
        max_len = max(len(str(c.value or '')) for c in col)
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # 2. RPM Matrix Sheet for Variant 202 Temperatures
    if "202" in variant_dict:
        ws2 = wb.create_sheet(title="Variant_202_RPM_vs_Temp")
        ws2.views.sheetView[0].showGridLines = True
        agg_egt = get_variant_aggregate(variant_dict["202"], "ex_gas_temp")

        ws2.cell(row=1, column=1, value="Engine Speed (RPM)").fill = header_fill
        ws2.cell(row=1, column=1).font = header_font
        for c_i, eng_curve in enumerate(agg_egt["engine_curves"], start=2):
            cell = ws2.cell(row=1, column=c_i, value=f"{eng_curve['serial']} ({eng_curve['week']}) [℃]")
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        ws2.cell(row=1, column=len(agg_egt["engine_curves"]) + 2, value="Mean EGT (℃)").fill = header_fill
        ws2.cell(row=1, column=len(agg_egt["engine_curves"]) + 2).font = header_font
        ws2.cell(row=1, column=len(agg_egt["engine_curves"]) + 3, value="Min EGT (℃)").fill = header_fill
        ws2.cell(row=1, column=len(agg_egt["engine_curves"]) + 3).font = header_font
        ws2.cell(row=1, column=len(agg_egt["engine_curves"]) + 4, value="Max EGT (℃)").fill = header_fill
        ws2.cell(row=1, column=len(agg_egt["engine_curves"]) + 4).font = header_font

        for r_i, rpm in enumerate(agg_egt["rpms"], start=2):
            ws2.cell(row=r_i, column=1, value=rpm).border = thin_border
            for e_i, eng_curve in enumerate(agg_egt["engine_curves"], start=2):
                ws2.cell(row=r_i, column=e_i, value=eng_curve["values"][r_i - 2]).border = thin_border
            ws2.cell(row=r_i, column=len(agg_egt["engine_curves"]) + 2, value=agg_egt["mean_curve"][r_i - 2]).border = thin_border
            ws2.cell(row=r_i, column=len(agg_egt["engine_curves"]) + 3, value=agg_egt["min_curve"][r_i - 2]).border = thin_border
            ws2.cell(row=r_i, column=len(agg_egt["engine_curves"]) + 4, value=agg_egt["max_curve"][r_i - 2]).border = thin_border

        for col in ws2.columns:
            max_len = max(len(str(c.value or '')) for c in col[:10])
            col_letter = get_column_letter(col[0].column)
            ws2.column_dimensions[col_letter].width = max(max_len + 3, 14)

    wb.save(output_excel_path)
    print(f"Generated comparison report: {output_excel_path}")

# -----------------------------------------------------------------------------
# Main Execution: Scan and Process
# -----------------------------------------------------------------------------
def run_pipeline(data_root_dir, output_dir):
    print(f"Scanning directory: {data_root_dir}")
    variants = scan_variants_directory(data_root_dir)
    print(f"Found variants: {list(variants.keys())}")
    for var, engs in variants.items():
        print(f"  Variant {var}: {len(engs)} engines parsed.")

    # Generate charts for Variant 202 (or all variants)
    for var_name, eng_list in variants.items():
        generate_multi_engine_plots(var_name, eng_list, output_dir)

    # Generate comparison report Excel
    report_path = os.path.join(output_dir, "comparision_report.xlsx")
    generate_comparison_report(variants, report_path)

    # Export clean JSON for Web Dashboard
    dashboard_export = {}
    for var_name, eng_list in variants.items():
        dashboard_export[var_name] = []
        for eng in eng_list:
            dashboard_export[var_name].append({
                "serial": eng["serial"],
                "week": eng["week"],
                "ref": eng["ref"],
                "cell": eng["cell"],
                "date": eng["metadata_block"].get("Date", eng["year"]),
                "purpose": eng["purpose"],
                "filename": eng["filename"],
                "data": eng["data"]
            })

    json_path = os.path.join(output_dir, "multi_engine_data.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(dashboard_export, f, indent=2)
    print(f"Saved dashboard dataset: {json_path}")
    return dashboard_export

if __name__ == "__main__":
    import sys
    data_dir = sys.argv[1] if len(sys.argv) > 1 else r"data"
    out_dir = sys.argv[2] if len(sys.argv) > 2 else r"."
    run_pipeline(data_dir, out_dir)
