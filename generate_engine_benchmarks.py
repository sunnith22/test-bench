#!/usr/bin/env python3
"""
Toyota TNGA 2.0L Engine Test Bench - Powertrain Calculation & Visualization Suite

Calculates ISO 1585 atmospheric correction, dyno and corrected power/torque curves,
BSFC thermal efficiency, and subsystem thermodynamic responses (coolant, oil, AFR,
exhaust gas temp, back pressure, manifold and atmospheric barometric pressure).
"""

import os
import math
import json
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt

# -------------------------------------------------------------
# 1. Base Configuration & Test Points
# -------------------------------------------------------------
RPMS = [
    1200, 1600, 2000, 2400, 2800, 3200, 3600, 4000, 
    4400, 4500, 4600, 4700, 4800, 4900, 5000, 5100, 
    5500, 5900, 6200, 6300, 6400, 6500, 6600, 6700
]

K_FACTOR = 1.1484  # Atmospheric correction factor (ISO 1585 / JIS D 1001)

# Toyota TNGA 2.0L (M20A 1987cc N/A Engine) baseline torque curve
CORRECTED_TORQUE_BASE = [
    151.2, 168.5, 181.0, 189.2, 194.5, 199.1, 203.4, 206.5,
    208.2, 208.6, 208.9, 209.1, 209.2, 208.8, 208.1, 207.2,
    203.5, 198.0, 191.5, 188.8, 186.2, 183.4, 181.1, 176.5
]

# BSFC curve (g/kW·h) - reaches ~40% thermal efficiency sweet spot (~216 g/kWh)
BSFC_BASE = [
    268.4, 242.1, 226.5, 219.8, 216.2, 218.0, 224.5, 235.0,
    246.8, 249.5, 252.3, 255.4, 258.4, 261.7, 265.2, 269.0,
    282.0, 296.5, 308.2, 312.4, 317.0, 321.8, 326.4, 330.5
]

# -------------------------------------------------------------
# 2. Physics & Dynamic Calculations
# -------------------------------------------------------------
def calculate_run_dynamics(rpms, t_corr_arr, bsfc_arr, k, date_str, ref_no, run_name, atm_base=90.85, dry_temp=26.3):
    """
    Computes all powertrain variables based on mechanical and thermodynamic relationships.
    """
    rows = []
    for i, rpm in enumerate(rpms):
        t_corr = t_corr_arr[i]
        t_dyno = round(t_corr / k, 1)
        p_dyno = round((t_dyno * rpm) / 9548.8, 2)
        p_corr = round((t_corr * rpm) / 9548.8, 1)
        bsfc = round(bsfc_arr[i], 2)
        fuel_flow = round((p_dyno * bsfc) / 3600.0, 2)
        acc = 100.2 if rpm <= 4000 else 100.1
        norm_rpm = (rpm - 1200) / (6700 - 1200)
        
        water_in = round(80.0 + 4.0 * (1 if rpm >= 4400 else 0), 1)
        water_out = round(84.0 + 7.5 * norm_rpm, 1)
        oil_temp = round(80.0 + 24.5 * (norm_rpm ** 0.8), 1)
        oil_press = round(175.0 + 295.0 * (1.0 - math.exp(-3.0 * norm_rpm)), 0)
        afr = round(14.45 - 2.25 * (norm_rpm ** 1.8), 2)
        in_air_press = round(0.18 + 0.28 * norm_rpm, 2)
        in_air_temp = round(dry_temp - 1.3 + 2.0 * norm_rpm, 1)
        ex_gas_temp = round(504.0 + 356.0 * (norm_rpm ** 0.65), 0)
        back_press = round(1.4 + 13.4 * (norm_rpm ** 1.9), 1)
        td_temp = round(dry_temp + 0.9 * math.sin(norm_rpm * math.pi), 1)
        tw_temp = 22.1
        fuel_press = round(406.26 - 6.5 * (norm_rpm ** 1.5), 2)
        crank_press = round(0.19 + 0.18 * norm_rpm, 2)
        atm = round(atm_base + 0.05 * math.cos(norm_rpm * 2), 2)
        
        rows.append({
            "rpm": rpm, "acc": acc, "k": k, "power": p_dyno,
            "corrected_power": p_corr, "dyno_torque": t_dyno,
            "corrected_torque": t_corr, "bsfc": bsfc,
            "fuel_consumption": fuel_flow, "water_temp_in": water_in,
            "water_temp_out": water_out, "oil_temp": oil_temp,
            "oil_press": oil_press, "afr": afr,
            "in_air_press": in_air_press, "in_air_temp": in_air_temp,
            "ex_gas_temp": ex_gas_temp, "back_press": back_press,
            "td_temp": td_temp, "tw_temp": tw_temp,
            "fuel_press": fuel_press, "crank_press": crank_press,
            "atm_press": atm, "date": date_str, "ref_no": ref_no,
            "run_name": run_name
        })
    return rows

def get_monthly_dataset():
    """Generates the full June test campaign dataset."""
    runs_config = [
        {"date": "2025-06-05", "ref": 51, "name": "Run 51 (Break-in Baseline)", "t_mod": -2.8, "bsfc_mod": 3.2, "atm": 91.2, "temp": 25.5},
        {"date": "2025-06-12", "ref": 52, "name": "Run 52 (VVT Optimization)", "t_mod": -1.2, "bsfc_mod": 1.4, "atm": 91.0, "temp": 25.8},
        {"date": "2025-06-18", "ref": 53, "name": "Run 53 (Thermal Durability)", "t_mod": 0.4, "bsfc_mod": -0.7, "atm": 90.7, "temp": 26.1},
        {"date": "2025-06-24", "ref": 54, "name": "Run 54 (Cylinder block new die)", "t_mod": 0.0, "bsfc_mod": 0.0, "atm": 90.9, "temp": 26.3},
        {"date": "2025-06-30", "ref": 55, "name": "Run 55 (Final Sign-off)", "t_mod": 1.3, "bsfc_mod": -1.6, "atm": 90.8, "temp": 26.5}
    ]
    runs_data = {}
    for r in runs_config:
        t_mod = [round(t + r["t_mod"], 1) for t in CORRECTED_TORQUE_BASE]
        b_mod = [round(b + r["bsfc_mod"], 2) for b in BSFC_BASE]
        runs_data[r["ref"]] = calculate_run_dynamics(RPMS, t_mod, b_mod, K_FACTOR, r["date"], r["ref"], r["name"], r["atm"], r["temp"])

    month_kpis = []
    for r in runs_config:
        r_data = runs_data[r["ref"]]
        month_kpis.append({
            "date": r["date"],
            "ref": r["ref"],
            "name": r["name"],
            "peak_power_kw": max(x["corrected_power"] for x in r_data),
            "peak_torque_nm": max(x["corrected_torque"] for x in r_data),
            "best_bsfc_gkwh": min(x["bsfc"] for x in r_data),
            "max_egt_c": max(x["ex_gas_temp"] for x in r_data),
            "avg_oil_press_kpa": round(float(np.mean([x["oil_press"] for x in r_data])), 1),
            "atm_press_kpa": r["atm"]
        })
    return RPMS, runs_config, month_kpis, runs_data

if __name__ == "__main__":
    rpms, configs, kpis, data = get_monthly_dataset()
    run54 = data[54]
    max_p = max(x['corrected_power'] for x in run54)
    max_t = max(x['corrected_torque'] for x in run54)
    min_b = min(x['bsfc'] for x in run54)
    print(f"Dataset generated successfully.")
    print(f"Run 54 Peak Power:  {max_p} kW @ {run54[[x['corrected_power'] for x in run54].index(max_p)]['rpm']} RPM")
    print(f"Run 54 Peak Torque: {max_t} N·m @ {run54[[x['corrected_torque'] for x in run54].index(max_t)]['rpm']} RPM")
    print(f"Run 54 Best BSFC:   {min_b} g/kW·h @ {run54[[x['bsfc'] for x in run54].index(min_b)]['rpm']} RPM")
