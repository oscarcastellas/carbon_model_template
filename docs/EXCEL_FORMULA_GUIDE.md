# Excel Formula-Based Model Guide

## Overview

The Excel export now generates a **fully formula-based model** where **every calculation is linked via Excel formulas** to the input assumptions. This ensures complete transparency and auditability for external reviewers.

## Excel File Structure

### Sheet 1: Inputs & Assumptions
**Purpose:** Central location for all model inputs

**Contents:**
- **Base Financial Assumptions:**
  - WACC (Weighted Average Cost of Capital)
  - Rubicon Investment Total
  - Investment Tenor (Years)
  - Initial Streaming Percentage
  - Target IRR
  - Target Streaming Percentage

- **Monte Carlo Assumptions** (if applicable):
  - Price Growth Base (Mean)
  - Price Growth Std Dev
  - Volume Multiplier Base (Mean)
  - Volume Std Dev

**Key Features:**
- All input values are clearly labeled
- Input cells are highlighted in yellow (`#FFF2CC`) for easy identification
- These cells are referenced by formulas throughout the model

---

### Sheet 2: Valuation Schedule
**Purpose:** Detailed 20-year cash flow analysis with all calculations as formulas

**Columns:**
1. **Year** - Year number (1-20)
2. **Carbon Credits Gross** - Input data (base credits)
3. **Rubicon Share of Credits** - **FORMULA:** `=B{row}*'Inputs & Assumptions'!$B$6`
4. **Base Carbon Price** - Input data (base price)
5. **Rubicon Revenue** - **FORMULA:** `=C{row}*D{row}`
6. **Project Implementation Costs** - Input data (costs)
7. **Rubicon Investment Drawdown** - **FORMULA:** `=IF(A{row}<=Tenor, -Investment/Tenor, 0)`
8. **Rubicon Net Cash Flow** - **FORMULA:** `=E{row}+G{row}`
9. **Discount Factor** - **FORMULA:** `=1/((1+WACC)^(A{row}-1))`
10. **Present Value** - **FORMULA:** `=H{row}*I{row}`
11. **Cumulative Cash Flow** - **FORMULA:** `=K{prev_row}+H{row}` (or `=H{row}` for Year 1)
12. **Cumulative PV** - **FORMULA:** `=L{prev_row}+J{row}` (or `=J{row}` for Year 1)

**Key Features:**
- All calculated cells are highlighted in green (`#E2EFDA`) to show they contain formulas
- Input data cells (Credits, Price, Costs) are in white - these can be modified
- Totals row at bottom uses `SUM()` formulas
- All formulas reference the Inputs & Assumptions sheet for key parameters

---

### Sheet 3: Summary & Results
**Purpose:** Key financial metrics with formulas linking to Valuation Schedule

**Key Metrics:**
- **Net Present Value (NPV)** - **FORMULA:** `=SUM('Valuation Schedule'!J3:J22)`
- **Internal Rate of Return (IRR)** - **FORMULA:** `=IRR('Valuation Schedule'!H3:H22)`
- **Payback Period** - **FORMULA:** `=MATCH(0,'Valuation Schedule'!K3:K22,1)+1`

**Target Metrics:**
- **Target IRR** - **FORMULA:** `='Inputs & Assumptions'!$B$7`
- **Target Streaming Percentage** - **FORMULA:** `='Inputs & Assumptions'!$B$8`
- **Actual IRR Achieved** - Shows calculated value

**Monte Carlo Summary** (if applicable):
- MC Mean IRR, P10 IRR, P90 IRR
- MC Mean NPV, P10 NPV, P90 NPV
- Standard deviations

**Key Features:**
- All metrics are calculated via formulas, not hardcoded values
- Easy to trace back to source data
- Clear separation between calculated and target metrics

---

### Sheet 4: Sensitivity Analysis
**Purpose:** 2D sensitivity table showing IRR under different scenarios

**Structure:**
- Rows: Credit Volume Multipliers (e.g., 0.90x, 1.00x, 1.10x)
- Columns: Carbon Price Multipliers (e.g., 0.80x, 1.00x, 1.20x)
- Values: IRR for each combination

**Key Features:**
- Clear row and column headers
- Formatted as percentages
- Shows how IRR changes with different assumptions

---

### Sheet 5: Monte Carlo Results
**Purpose:** Full Monte Carlo simulation results and analysis

**Contents:**
1. **Summary Statistics:**
   - Mean IRR, P10 IRR, P90 IRR
   - Mean NPV, P10 NPV, P90 NPV
   - Standard deviations
   - Total and valid simulations

2. **Full Simulation Results:**
   - Two columns: IRR and NPV for each of 5,000 simulations
   - Allows for custom analysis in Excel

3. **IRR Histogram Chart:**
   - Visual distribution of IRR outcomes
   - Shows risk profile of the investment

**Key Features:**
- Complete transparency of all simulation results
- Easy to perform additional analysis in Excel
- Visual representation of risk distribution

---

## Formula Color Coding

The model uses color coding to help reviewers understand the structure:

- **Yellow Background (`#FFF2CC`):** Input assumptions (can be modified)
- **Green Background (`#E2EFDA`):** Calculated cells with formulas
- **Blue Background (`#D9E1F2`):** Summary/total cells
- **White Background:** Input data (can be modified)

---

## How to Use for Review

1. **Start with Inputs & Assumptions Sheet:**
   - Review all input values
   - Modify any assumptions to see impact

2. **Check Valuation Schedule:**
   - Verify formulas are correct
   - Trace calculations back to inputs
   - Modify input data (Credits, Price, Costs) to see recalculations

3. **Review Summary & Results:**
   - Verify NPV, IRR formulas reference correct cells
   - Check that metrics match expectations

4. **Analyze Sensitivity:**
   - Review sensitivity table for risk assessment
   - Understand impact of different scenarios

5. **Review Monte Carlo (if applicable):**
   - Check simulation results
   - Review histogram for risk distribution
   - Perform additional analysis as needed

---

## Benefits of Formula-Based Model

1. **Full Transparency:** Every calculation can be traced and verified
2. **Easy Auditing:** Reviewers can see exactly how results are calculated
3. **Flexibility:** Change inputs and see immediate impact
4. **No Black Box:** All calculations are visible and understandable
5. **Professional:** Meets standards for financial model review

---

## Technical Notes

- All formulas use absolute references (`$B$6`) for key assumptions to ensure consistency
- Relative references are used for row-by-row calculations
- Excel functions used: `SUM()`, `IRR()`, `IF()`, `MATCH()`
- Sheet references use format: `'Sheet Name'!CellReference`

---

## Support

For questions or issues with the formula-based Excel model, refer to:
- `ARCHITECTURE.md` - System architecture
- `HOW_TO_USE.md` - Usage guide
- Source code in `reporting/excel_exporter.py`

