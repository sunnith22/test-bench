# Toyota TNGA 2.0L Engine Test Bench Analysis & Dashboard

An interactive dynamometer analytics platform and powertrain simulation engine for the **Toyota TNGA 2.0L Engine (`2.0L-CV-(124E)`)** test campaign in Cell TC-1.

---

## 🏎️ Features

- **Master Engine Dyno Curves:** Computes dual-axis Power (kW) and Torque (N·m) curves with ISO 1585 / JIS D 1001 atmospheric correction ($k = 1.1484$).
- **Subsystem Thermodynamic Response Suite:** Analyzes 15+ engine channels from 1,200 to 6,700 RPM:
  - Brake Specific Fuel Consumption (BSFC) & Fuel Flow Rate
  - Coolant (Water In / Water Out) and Oil Sump Temperatures
  - Variable Oil Pump Pressure Relief and Fuel Rail Pressure
  - Air-Fuel Ratio (AFR) & Exhaust Back Pressure
  - Test Cell Atmospheric Barometric Pressure and Intake Depression
- **Interactive Frontend Dashboard (`engine_test_dashboard.html`):**
  - **Dynamic Parameter Dropdown Filter:** View any metric plotted directly against RPM.
  - **Month Run Dropdown Filter:** Switch between individual June test runs or view **All Runs Month Overlay**.
  - **Monthly Trend KPI Chart:** Tracks peak power, peak torque, best BSFC, and thermal progression over time.
  - **Searchable Data Table:** Real-time search and CSV export.

---

## 📊 Powertrain Specifications (Run 54)

- **Engine Variant:** 2.0L-CV-(124E) (Toyota TNGA M20A, 1,987 cc)
- **Test Condition:** 100% WOT (Wide Open Throttle) Full Load
- **Peak Corrected Power:** **125.2 kW (170.2 PS)** @ 6,600 RPM
- **Peak Corrected Torque:** **209.2 N·m** @ 4,800 RPM
- **Optimal BSFC:** **216.2 g/kW·h** @ 2,800 RPM (~40% thermal efficiency)
- **Peak Exhaust Gas Temp:** **860 °C** @ 6,700 RPM

---

## 🛠️ Usage

### 1. View the Interactive Dashboard
Simply open `engine_test_dashboard.html` in any web browser. No web server required!

### 2. Run the Benchmark Generator
```bash
pip install -r requirements.txt
python generate_engine_benchmarks.py
```

---

## 📁 Repository Structure

```text
├── engine_test_dashboard.html      # Interactive frontend analytics dashboard
├── generate_engine_benchmarks.py   # Powertrain calculations & physics engine
├── requirements.txt                # Python dependencies
├── .gitignore                      # Excludes large binaries/data (code-only)
└── README.md                       # Project documentation
```
