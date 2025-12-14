# Sheet Creation Audit: Old vs New Version

## Executive Summary

This document compares how the **old version** (using `xlsxwriter`, creating sheets from scratch) and the **new version** (using `openpyxl`, populating a template) create the first 3 sheets:
1. **Inputs & Assumptions**
2. **Valuation Schedule**
3. **Summary & Results**

---

## 1. Inputs & Assumptions Sheet

### Old Version (`export/excel.py` - xlsxwriter)
- **Library**: `xlsxwriter`
- **Creation Method**: Creates sheet from scratch
- **Row Layout**:
  - Row 0: Title "Carbon Model - Inputs & Assumptions"
  - Row 2: Subtitle "Base Financial Assumptions"
  - Row 3: WACC (cell B3)
  - Row 4: Rubicon Investment Total (cell B4)
  - Row 5: Investment Tenor (cell B5)
  - Row 6: Initial Streaming Percentage (cell B6)
  - Row 7: Target IRR (cell B7)
  - Row 8: Target Streaming Percentage (cell B8)
- **Named Ranges**: Creates named ranges for formula references
- **Formatting**: Uses predefined format dictionary with colors, borders, fonts

### New Version (`export/template_based_export.py` - openpyxl)
- **Library**: `openpyxl`
- **Creation Method**: Populates existing template sheet
- **Row Layout**:
  - Row 1: Title "Carbon Model - Inputs & Assumptions"
  - Row 3: Subtitle "Base Financial Assumptions"
  - Row 4: WACC (cell B4) ⚠️ **DIFFERENT**
  - Row 5: Rubicon Investment Total (cell B5) ⚠️ **DIFFERENT**
  - Row 6: Investment Tenor (cell B6) ⚠️ **DIFFERENT**
  - Row 7: Initial Streaming Percentage (cell B7) ⚠️ **DIFFERENT**
  - Row 8: Target IRR (cell B8) ⚠️ **DIFFERENT**
  - Row 9: Target Streaming Percentage (cell B9) ⚠️ **DIFFERENT**
- **Search Method**: Searches for labels and populates values (less precise)
- **Formatting**: Applied via `ProfessionalFormatter` class

### Key Differences
1. **Row Offsets**: New version has all inputs shifted down by 1 row (B3→B4, B4→B5, etc.)
2. **Precision**: Old version writes to exact cells; new version searches for labels
3. **Named Ranges**: Old version creates named ranges; new version uses direct cell references

---

## 2. Valuation Schedule Sheet

### Old Version (`export/excel.py` - xlsxwriter)
- **Title**: "Valuation Schedule - 20 Year Cash Flow" (Row 0)
- **Header Row**: Row 2 (0-indexed) = Excel Row 3
- **Years**: 20 years (2025-2044), columns B-U
- **Total Column**: Column V (column 22, 0-indexed)
- **Line Items** (11 total):
  1. Carbon Credits Gross (Row 3, Excel Row 4)
  2. Rubicon Share of Credits (Row 4, Excel Row 5) - Formula
  3. Base Carbon Price (Row 5, Excel Row 6)
  4. Rubicon Revenue (Row 6, Excel Row 7) - Formula
  5. Project Implementation Costs (Row 7, Excel Row 8)
  6. Rubicon Investment Drawdown (Row 8, Excel Row 9) - Formula
  7. Rubicon Net Cash Flow (Row 9, Excel Row 10) - Formula
  8. Discount Factor (Row 10, Excel Row 11) - Formula
  9. Present Value (Row 11, Excel Row 12) - Formula
  10. Cumulative Cash Flow (Row 12, Excel Row 13) - Formula
  11. Cumulative PV (Row 13, Excel Row 14) - Formula

- **Input References**:
  - WACC: `'Inputs & Assumptions'!$B$3`
  - Investment: `'Inputs & Assumptions'!$B$4`
  - Tenor: `'Inputs & Assumptions'!$B$5`
  - Streaming: `'Inputs & Assumptions'!$B$6`

- **Formula Examples**:
  - Rubicon Share: `=B4*'Inputs & Assumptions'!$B$6`
  - Revenue: `=B5*B6`
  - Investment: `=IF(1<='Inputs & Assumptions'!$B$5,-'Inputs & Assumptions'!$B$4/'Inputs & Assumptions'!$B$5,0)`
  - Net CF: `=B7+B9`
  - Discount: `=1/((1+'Inputs & Assumptions'!$B$3)^(1-1))`
  - PV: `=B10*B11`
  - Cum CF: `=B10` (first year), `=A13+B10` (subsequent years)
  - Cum PV: `=B12` (first year), `=A14+B12` (subsequent years)

### New Version (`export/template_based_export.py` - openpyxl)
- **Title**: "Valuation Schedule - 20 Year Cash Flow" (Row 1)
- **Header Row**: Row 3 (1-indexed)
- **Years**: 21 years (2025-2045), columns B-V ⚠️ **DIFFERENT**
- **Total Column**: Column W (column 23, 1-indexed) ⚠️ **DIFFERENT**
- **Line Items**: Same 11 line items, same order
- **Row Positions**: Same Excel row numbers (4-14)

- **Input References**:
  - WACC: `'Inputs & Assumptions'!$B$4` ⚠️ **DIFFERENT**
  - Investment: `'Inputs & Assumptions'!$B$5` ⚠️ **DIFFERENT**
  - Tenor: `'Inputs & Assumptions'!$B$6` ⚠️ **DIFFERENT**
  - Streaming: `'Inputs & Assumptions'!$B$7` ⚠️ **DIFFERENT**

- **Formula Examples**:
  - Rubicon Share: `=B4*'Inputs & Assumptions'!$B$7` ⚠️ **DIFFERENT REFERENCE**
  - Revenue: `=B5*B6` ✓ Same
  - Investment: `=IF(1<='Inputs & Assumptions'!$B$6,-'Inputs & Assumptions'!$B$5/'Inputs & Assumptions'!$B$6,0)` ⚠️ **DIFFERENT REFERENCES**
  - Net CF: `=B7+B9` ✓ Same
  - Discount: `=1/((1+'Inputs & Assumptions'!$B$4)^(1-1))` ⚠️ **DIFFERENT REFERENCE**
  - PV: `=B10*B11` ✓ Same
  - Cum CF: `=B10` (first year), `=A13+B10` (subsequent years) ✓ Same
  - Cum PV: `=B12` (first year), `=A14+B12` (subsequent years) ✓ Same

### Key Differences
1. **Number of Years**: Old = 20 years (2025-2044), New = 21 years (2025-2045) ⚠️ **CRITICAL**
2. **Input Cell References**: All shifted down by 1 row (B3→B4, B4→B5, etc.) ⚠️ **CRITICAL**
3. **Total Column**: Old = Column V, New = Column W
4. **Formula Structure**: Same logic, but references different input cells

---

## 3. Summary & Results Sheet

### Old Version (`export/excel.py` - xlsxwriter)
- **Title**: "Summary & Results" (Row 0)
- **Structure**:
  - Row 2: "Key Financial Metrics" (subtitle)
  - Row 3: NPV label
  - Row 4: IRR label
  - Row 5: Payback Period label (if provided)
  - Row 7: "Target Metrics" (subtitle)
  - Row 8: Target IRR
  - Row 9: Target Streaming Percentage
  - Row 10: Actual IRR Achieved
  - Row 12: "Monte Carlo Simulation Summary" (if provided)
  - Row 19: "Risk Assessment" (if provided)
  - Row 37: "Breakeven Analysis" (if provided)

- **Key Formulas**:
  - **NPV**: `=SUM('Valuation Schedule'!B12:U12)` (Row 12, columns B-U = 20 years)
  - **IRR**: `=IRR('Valuation Schedule'!B10:U10)` (Row 10, columns B-U = 20 years)
  - **Payback**: `=MATCH(0,'Valuation Schedule'!B13:U13,1)` (Row 13, columns B-U = 20 years)
  - **Target IRR**: `='Inputs & Assumptions'!$B$7`
  - **Target Streaming**: `='Inputs & Assumptions'!$B$8`

- **Formatting**: Uses predefined format dictionary with colors, borders, fonts

### New Version (`export/template_based_export.py` - openpyxl)
- **Title**: "Summary & Results" (Row 1)
- **Structure**: Same sections, same order
- **Row Positions**: Same relative positions

- **Key Formulas**:
  - **NPV**: `=SUM('Valuation Schedule'!B12:V12)` ⚠️ **DIFFERENT** (Row 12, columns B-V = 21 years)
  - **IRR**: `=IRR('Valuation Schedule'!B10:V10)` ⚠️ **DIFFERENT** (Row 10, columns B-V = 21 years)
  - **Payback**: `=MATCH(0,'Valuation Schedule'!B13:V13,1)+1` ⚠️ **DIFFERENT** (Row 13, columns B-V = 21 years, +1 added)
  - **Target IRR**: `='Inputs & Assumptions'!$B$8` ⚠️ **DIFFERENT**
  - **Target Streaming**: `='Inputs & Assumptions'!$B$9` ⚠️ **DIFFERENT**

- **Formatting**: Applied via `ProfessionalFormatter` class

### Key Differences
1. **Formula Ranges**: Old = B-U (20 years), New = B-V (21 years) ⚠️ **CRITICAL**
2. **Payback Formula**: New version adds `+1` to the MATCH result ⚠️ **CRITICAL**
3. **Input References**: Target IRR and Streaming shifted down by 1 row (B7→B8, B8→B9) ⚠️ **CRITICAL**

---

## Critical Issues Identified

### 1. **Year Count Mismatch**
- **Old**: 20 years (2025-2044)
- **New**: 21 years (2025-2045)
- **Impact**: All formulas referencing year ranges are incorrect
- **Fix Required**: Either:
  - Change new version to 20 years, OR
  - Update all formulas to use 21 years consistently

### 2. **Input Sheet Row Offset**
- **Old**: Inputs start at Row 3 (B3, B4, B5, B6, B7, B8)
- **New**: Inputs start at Row 4 (B4, B5, B6, B7, B8, B9)
- **Impact**: All formulas in Valuation Schedule and Summary reference wrong cells
- **Fix Required**: Either:
  - Change Inputs sheet to match old layout (B3-B8), OR
  - Update all formula references to new layout (B4-B9)

### 3. **Formula Range References**
- **Old**: Uses columns B-U (20 columns)
- **New**: Uses columns B-V (21 columns)
- **Impact**: Summary formulas may include extra year or miss last year
- **Fix Required**: Ensure consistency with chosen year count

### 4. **Payback Formula**
- **Old**: `=MATCH(0,'Valuation Schedule'!B13:U13,1)`
- **New**: `=MATCH(0,'Valuation Schedule'!B13:V13,1)+1`
- **Impact**: New version adds +1, which may be incorrect
- **Fix Required**: Verify if +1 is needed or if old formula was correct

---

## Recommendations

### Option 1: Match Old Version Exactly (Recommended)
1. **Change Inputs Sheet**: Move all inputs up by 1 row (B4→B3, B5→B4, etc.)
2. **Change Year Count**: Use 20 years (2025-2044) instead of 21
3. **Update Formula Ranges**: Change B-V to B-U in all Summary formulas
4. **Fix Payback Formula**: Remove `+1` if old version was correct

### Option 2: Keep New Version But Fix References
1. **Keep Inputs Sheet**: Maintain current layout (B4-B9)
2. **Keep Year Count**: Use 21 years (2025-2045)
3. **Verify Formula Ranges**: Ensure all formulas use B-V consistently
4. **Verify Payback Formula**: Test if `+1` is correct for 21-year model

### Option 3: Hybrid Approach
1. **Match Old Input Layout**: Use B3-B8 for inputs (matches old version)
2. **Use 20 Years**: Match old version exactly
3. **Update All References**: Ensure all formulas reference correct cells

---

## Testing Checklist

After fixing, verify:
- [ ] Inputs sheet has correct values in correct cells
- [ ] Valuation Schedule has 20 (or 21) years with correct formulas
- [ ] All formulas in Valuation Schedule reference correct Inputs cells
- [ ] Summary NPV formula sums correct range (B12:U12 or B12:V12)
- [ ] Summary IRR formula uses correct range (B10:U10 or B10:V10)
- [ ] Summary Payback formula works correctly
- [ ] All formulas calculate correctly (no circular references)
- [ ] Formatting matches old version's professional appearance

---

## Conclusion

The new version has **structural differences** that cause formula reference errors:
1. Inputs sheet rows are offset by +1
2. Year count changed from 20 to 21
3. Formula ranges need to match the chosen year count
4. Payback formula may have incorrect +1

**Recommended Action**: Match the old version exactly (20 years, inputs at B3-B8) to ensure compatibility and correctness.

