# M&A Analysis: Cloudberry Clean Energy / Orrön Energy

Automated financial model and downside stress-testing framework for a strategic consolidation scenario in the Nordic renewable market (hydro/wind).

## Transaction Overview

* **Acquirer:** Cloudberry Clean Energy ASA (NOK / EUR)
* **Target:** Orrön Energy AB (SEK / EUR)
* **Target Implied Enterprise Value (EV):** ~€148.5M
* **Implied EV/EBITDA Multiple:** ~11.8x
* **Pro Forma 2026E EPS Accretion:** +3.8% (Base Case)
* **Financing Structure:** 95% Equity (107M new shares issued), 5% Debt (€4.2M new debt)

## Executive Rationale

1. **Portfolio Balancing:** Combines Cloudberry’s run-of-river hydro assets (NO3/NO4 price zones) with Orrön’s wind development portfolio to hedge spot price volatility and capture area price differentials.
2. **Grid & Balancing Synergies:** Estimated €5.0M post-tax annual operating synergies realized through shared grid connections and optimized PPA structuring.
3. **Capital Structure Preservation:** Minimizes debt leverage to protect credit metrics against elevated long-term interest rates.

## Valuation & Deal Summary

| Metric | Standalone (Orrön) | Pro Forma Combined | Delta / Impact |
| :--- | :--- | :--- | :--- |
| **2026E Revenue (€M)** | €37.8M | €180.0M+ | Enlarged Platform |
| **EBITDA Margin (%)** | 35.0% | 38.2% | +320 bps (Synergies) |
| **2026E Net Income (€M)** | €8.4M | €42.3M | Consolidated |
| **2026E EPS (€)** | €0.08 | €0.11 | **+3.8% Accretion** |

---

## Downside Stress Tests & Breaking Points

Rather than relying on baseline projections, the model evaluates transaction failure thresholds across commodity price shocks and cost of capital shifts.

### 1. Valuation Sensitivity (Implied EV vs. WACC & Terminal Growth)

* **Base Case WACC:** 8.5% | **Terminal Growth:** 2.0%
* **Breaking Point:** Valuation collapses below purchase consideration if WACC exceeds 11.0% under zero terminal growth.

| WACC \ PGR | 0.0% | 1.0% | 2.0% (Base) | 3.0% | 4.0% |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **6.5%** | €182.4M | €198.1M | €218.0M | €244.5M | €281.2M |
| **8.5% (Base)** | €132.1M | €139.8M | **€148.5M** | €160.1M | €175.2M |
| **9.5%** | €116.2M | €121.9M | €128.6M | €136.7M | €146.6M |
| **11.0%** | €98.1M | €102.0M | €106.5M | €111.8M | €118.0M |
| **12.5%** | €84.5M | €87.3M | €90.5M | €94.2M | €98.5M |

### 2. Accretion / Dilution Sensitivity (Synergies vs. Revenue Growth)

* **Base Case Synergies:** €5.0M | **Base Case Target Growth:** 5.0%
* **Breaking Point:** Deal turns dilutive if target organic revenue drops by >10% with zero operational synergies realized.

| Synergies \ Rev Growth | -10.0% | -5.0% | 0.0% | 5.0% (Base) | 10.0% |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **€0.0M** | -2.1% | -0.8% | +0.4% | +1.7% | +3.0% |
| **€2.5M** | -0.5% | +0.8% | +2.0% | +3.3% | +4.6% |
| **€5.0M (Base)** | +1.1% | +2.4% | +3.6% | **+4.9%** | +6.2% |
| **€7.5M** | +2.7% | +4.0% | +5.2% | +6.5% | +7.8% |
| **€10.0M** | +4.3% | +5.6% | +6.8% | +8.1% | +9.4% |

---

## Technical Architecture

The financial model is programmatically constructed using Python (`xlsxwriter`) to ensure formula consistency, dynamic link verification, and automated formatting compliance with investment banking standards.

### Model Features
* **Dynamic Linking:** Hardcoded inputs (Blue) separated from calculations (Black) and inter-tab references (Green).
* **Circuit Breakers:** Pre-programmed error traps handling edge cases (WACC <= PGR, Shares = 0).
* **Automated Audit Suite:** Integrated `pytest` suite validating three-statement integrity, zero `#REF!` errors, and sign consistency.

### Running the Generator & Tests

```bash
# Install dependencies
pip install -r requirements.txt

# Run model audit tests
pytest tests/

# Generate Excel model artifact
python build_model.py
```

The resulting audit-ready workbook is exported to `output/Cloudberry_Orron_MA_Model_v3.xlsx`.

***
*Disclaimer: This is a hypothetical case study based on public information. It is not investment advice or an actual solicitation of any transaction.*
