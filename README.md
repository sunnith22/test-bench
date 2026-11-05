# Toyota TNGA Multi-Engine Dyno Bench - Variant Analytics & Dashboard

An interactive dynamometer analytics platform and automated multi-engine variant processing engine for **Toyota TNGA Engine Variants (202, 203, 319, 328)**.

---

## 🏎️ Overview & Architecture

Engine dyno data is organized in variant folders (e.g. `202`, `203`, `319`, `328`), where each folder contains Excel test runs (`.xlsx`) named following Toyota test bench conventions:

```text
data/ (or Pavithra/)
├── 202/
│   ├── 2025_W26_54_TC-1_TNGA_202_NE35384_Regular+Dis assy.xlsx
│   ├── 2025_W32_64_TC-2_TNGA_202_NE56301_Regular+Dis assy.xlsx
│   ├── 2025_W33_825_TC-2_TNGA_202_NE56301_Special Request+Dis assy.xlsx
│   ├── 2025_W38_80_TC-2_TNGA_202_NE78826_Regular+Dis assy.xlsx
│   ├── 2025_W41_86_TC-2_TNGA_202_NE88226_Regular+Dis assy.xlsx
│   ├── 2025_W50_103_TC-2_TNGA_202_NF23111_Regular+Dis assy.xlsx
│   ├── 2026_W11_31_TC-2_TNGA_202_NF73447_Special Request+Shipment.xlsx
│   ├── 2026_W16_43_TC-1_TNGA_202_NF92780_Regular+Dis assy.xlsx
│   ├── 2026_W24_60_TC-1_TNGA_202_NG24149_Regular+Dis assy.xlsx
│   └── 2026_W35_85_TC-1_TNGA_202_NG63293_Regular+Dis assy.xlsx
├── 203/
├── 319/
└── 328/
```

When analyzing a variant (e.g. **Variant 202**), the system automatically scans the `202/` folder, parses every engine file, extracts the 23 recorded channels against RPM, and superimposes **all engines on a single interactive comparison graph** along with the **Variant Mean curve** and **Min-Max dispersion envelope**.

---

## 📊 Key Features

- **Automated Directory Scanner (`multi_engine_analyzer.py`):**
  - Scans any root directory (local folder or OneDrive corporate folder like `Pavithra`).
  - Matches filenames via regex to extract: Year, Week, Test Ref, Cell, Engine Type, Variant, Serial Number, and Purpose.
  - Parses 3-row header structures and dyno calibration tables across all RPM test points (1,200 to 6,700 RPM).
- **Multi-Engine RPM vs Temperature Suite:**
  - **Exhaust Gas Temperature (Ex-gas Temp, ℃)** across all engines.
  - **Coolant Outlet & Inlet Temperatures (Water Temp out/in, ℃)**.
  - **Engine Oil Sump Temperature (Oil Temp, ℃)**.
  - **Intake Air Temperature (In Air Temp, ℃)**.
- **Powertrain Performance & Fluid Response:**
  - Corrected Power (kW) & Dyno Power (kW)
  - Corrected Torque (N·m) & Dyno Torque (N·m)
  - Fuel Consumption Ratio / BSFC (g/kW·h) & Flow Rate (g/s)
  - Oil Pressure, AFR, Back Pressure, Fuel Pressure, Crankcase Pressure, Atmospheric Pressure.
- **Statistical Aggregation & Reporting:**
  - Generates `comparision_report.xlsx` aggregating cross-engine benchmarks, rankings, and RPM matrix.
  - Generates high-resolution comparison plots (`rpm_vs_temp_variant_202_all_engines.png`).
- **Interactive Web Dashboard (`engine_test_dashboard.html`):**
  - **Variant Selector:** Switch between 202, 203, 319, 328.
  - **Metric Selector:** Select any temperature or performance channel against RPM.
  - **Engine Filter Chips:** Toggle individual engines on/off to compare specific serials.
  - **Mean & Tolerance Band Toggles:** View average curve and ±2σ spread.
  - **Searchable Table & CSV Export:** Real-time search across test runs.

---

## 🛠️ Usage

### 1. View the Interactive Dashboard
Open `engine_test_dashboard.html` in any browser.

### 2. Run the Multi-Engine Analyzer
```bash
# Process default 'data' directory:
python multi_engine_analyzer.py data .

# Or process a custom folder (e.g., OneDrive path):
python multi_engine_analyzer.py "C:\path\to\Pavithra" .
```

---

## 📁 Repository Structure

```text
├── engine_test_dashboard.html      # Interactive multi-engine web dashboard
├── multi_engine_analyzer.py        # Automated directory scanner & multi-engine aggregator
├── generate_engine_benchmarks.py   # Baseline powertrain calculation model
├── requirements.txt                # Python dependencies
├── .gitignore                      # Excludes raw data & media (code-only repo)
└── README.md                       # Documentation & usage guide
```
