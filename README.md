

# 📊 Investment & Trading Decision Toolkit

A comprehensive toolkit built in **Excel**, **VBA**, and **Python**, designed to assist traders and investors with macroeconomic analysis, rate expectations, COT positioning, market flows, and cross-country comparisons. Some dashboards integrate real-time data via **LSEG Workspace**.

---

## 📁 Repository Contents

| File                                  | Description                                                                                                                                       |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| **`BSM_Model_TTE.xlsx`**              | Black-Scholes-Merton (BSM) pricing model with scenario testing for options valuation.                                                             |
| **`COT_Report.xlsm`**                 | Commitment of Traders (COT) Excel dashboard showing net positioning by trader category. Uses VBA for dynamic chart updates.                       |
| **`Flows.py`**                        | Python script for building directional network graphs of FX strength based on pairwise flows. Auto-embeds charts into `Weekly.xlsx`.              |
| **`Fondamental Comparison.xlsm`**     | Macroeconomic comparison dashboard between countries with charts for inflation, GDP, current account, etc.                                        |
| **`Interest Rates Probability.xlsm`** | Table and charts estimating rate hike probabilities across central banks using market-implied data.                                               |
| **`Macro Snapshot.xlsm`**             | Quick-view dashboard showing high-level macro indicators by country with historical charts.                                                       |
| **`Weekly.xlsx`**                     | Multi-sheet Excel workbook with embedded FX network charts for each trading day (`MONDAY` to `FRIDAY`) and summary views (`WEEKLY`, `YESTERDAY`). |

---

## 🔧 Tools & Features

* **Excel Dashboards**

  * Visual indicators of macro data and rate expectations.
  * VBA-powered update buttons.
  * Historical and comparative analysis.

* **Python Automation**

  * Generates currency strength networks from performance tables.
  * Auto-injects visual graphs into Excel.

* **LSEG Workspace Integration**

  * Connects Excel dashboards to real-time or refreshed market data (e.g., rates, inflation, GDP, etc.).

* **VBA Automation**

  * One-click chart updates and data refresh mechanisms.

---

## 💻 Requirements

### Python (for Flows.py)

Install dependencies:

```bash
pip install pandas matplotlib networkx openpyxl numpy
```

### Excel

* Macros must be **enabled**.
* **LSEG Workspace Add-In** must be installed (for real-time data connections).
* Some files reference active or saved workspace series.

---

## 🚀 Getting Started

1. **Clone the repository**

   ```bash
   git clone https://github.com/your-username/your-repo-name.git
   ```

2. **Run the FX flow visualizer**

   ```bash
   python Flows.py
   ```

3. **Open Excel Dashboards**

   * Enable content and macros.
   * Use update buttons or refresh via LSEG if connected.

---

## 📌 Use Cases

* 🧠 Macro trader tracking economic divergence.
* 📅 Prepping for central bank meetings.
* 💱 Comparing FX relative strength through flows.
* 📈 Spotting trends in COT positioning.

