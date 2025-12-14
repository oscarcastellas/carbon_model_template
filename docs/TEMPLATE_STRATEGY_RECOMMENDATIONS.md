# Template Strategy Recommendations

## Executive Summary

This document provides strategic recommendations for creating a **generic, reusable Excel template** that can be used with the GUI application to analyze any carbon credit project. The template should be company-agnostic and work seamlessly with the existing workflow.

---

## Current Workflow

```
GUI Application
    ↓
Upload Raw Data (Excel/CSV)
    ↓
Data Analysis (Python backend)
    ↓
Excel Export (populate template)
    ↓
Final Excel File (ready for review)
```

---

## Recommended Approach: **Hybrid Template System**

### Option 1: Master Template with Smart Population (RECOMMENDED) ⭐

**Concept**: Create a single, generic master template that gets copied and populated by the GUI.

#### Advantages:
- ✅ **Single Source of Truth**: One template file to maintain
- ✅ **Consistent Structure**: All outputs have identical layout
- ✅ **Fast Generation**: Copy template, populate data, done
- ✅ **Easy Updates**: Update template once, all future exports benefit
- ✅ **Preserves Formatting**: All professional formatting stays intact
- ✅ **Formula Preservation**: All formulas remain in template

#### Structure:
```
templates/
  └── master_template.xlsx (or .xlsm if VBA needed)
      ├── Inputs & Assumptions (with generic labels)
      ├── Valuation Schedule (with formulas, empty data)
      ├── Summary & Results (with formulas, empty data)
      ├── Analysis (separator)
      ├── Deal Valuation (interactive)
      ├── Monte Carlo Results (interactive)
      ├── Sensitivity Analysis (interactive)
      └── Breakeven Analysis (interactive)
```

#### Generic Naming Convention:
- **"Rubicon"** → **"Investor"** or **"Company"**
- **"Rubicon Investment"** → **"Total Investment"** or **"Investor Investment"**
- **"Rubicon Share"** → **"Investor Share"** or **"Company Share"**
- **"Rubicon Revenue"** → **"Investor Revenue"** or **"Company Revenue"**
- **"Rubicon Net Cash Flow"** → **"Investor Net Cash Flow"** or **"Net Cash Flow to Investor"**

#### Implementation Steps:
1. **Create Master Template** (one-time):
   - Use old version's structure (20 years, correct cell references)
   - Replace all "Rubicon" with "Investor"
   - Add all formulas (empty data cells)
   - Apply professional formatting
   - Save as `templates/master_template.xlsx`

2. **GUI Export Process**:
   - Copy master template to output location
   - Populate Inputs & Assumptions sheet
   - Populate Valuation Schedule data cells (keep formulas)
   - Populate Summary & Results (formulas auto-calculate)
   - Save final file

3. **Benefits**:
   - Template is static (never modified)
   - Each export is independent
   - Easy to version control template
   - Can be shared/distributed separately

---

### Option 2: Generate from Scratch Each Time

**Concept**: Generate Excel file from scratch using xlsxwriter (like old version).

#### Advantages:
- ✅ **No Template File**: No need to maintain template
- ✅ **Dynamic Structure**: Can adapt to different project lengths
- ✅ **Full Control**: Every cell created programmatically

#### Disadvantages:
- ❌ **Slower**: Takes longer to generate
- ❌ **More Code**: More complex export logic
- ❌ **Harder to Update**: Changes require code changes
- ❌ **Formatting Risk**: Formatting might not be consistent

#### When to Use:
- If template structure needs to vary significantly per project
- If you want maximum flexibility in output structure

---

### Option 3: Hybrid - Template Base + Dynamic Generation

**Concept**: Use template for structure, generate formulas dynamically.

#### Advantages:
- ✅ **Best of Both**: Template consistency + dynamic formulas
- ✅ **Flexible**: Can adapt to different project parameters

#### Disadvantages:
- ❌ **Complex**: More moving parts
- ❌ **Maintenance**: Need to maintain both template and code

---

## Recommended Choice: **Option 1 - Master Template**

### Why This is Best:

1. **Simplicity**: One template file, straightforward workflow
2. **Consistency**: Every output looks identical (professional)
3. **Maintainability**: Update template once, all exports improve
4. **Performance**: Fast (just copy + populate)
5. **User Experience**: Template can be reviewed/approved before use
6. **Distribution**: Template can be shared independently

---

## Generic Naming Strategy

### Recommended Generic Terms:

| Old Term | New Generic Term | Rationale |
|----------|------------------|-----------|
| "Rubicon Investment Total" | "Total Investment" or "Investor Investment" | Clear, generic |
| "Rubicon Share of Credits" | "Investor Share of Credits" | Generic but clear |
| "Rubicon Revenue" | "Investor Revenue" | Generic but clear |
| "Rubicon Investment Drawdown" | "Investment Drawdown" | Can drop company name |
| "Rubicon Net Cash Flow" | "Investor Net Cash Flow" | Clear ownership |

### Alternative: Company Name Placeholder

Instead of "Investor", use a placeholder that gets replaced:
- `{COMPANY_NAME}` → Replaced with actual company name during export
- Allows customization per client/project
- Still generic in template

---

## Template Structure Recommendations

### 1. Inputs & Assumptions Sheet

**Layout** (matching old version):
```
Row 0: Title "Carbon Credit Investment Model - Inputs & Assumptions"
Row 2: "Base Financial Assumptions"
Row 3: WACC (B3)
Row 4: Total Investment (B4) [was: Rubicon Investment Total]
Row 5: Investment Tenor (B5)
Row 6: Initial Streaming Percentage (B6)
Row 7: Target IRR (B7)
Row 8: Target Streaming Percentage (B8)
```

**Generic Labels**:
- "Total Investment" instead of "Rubicon Investment Total"
- Keep "Streaming Percentage" (already generic)
- Keep "Target IRR" (already generic)

### 2. Valuation Schedule Sheet

**Line Items** (generic names):
1. Carbon Credits Gross ✓ (already generic)
2. **Investor Share of Credits** (was: Rubicon Share of Credits)
3. Base Carbon Price ✓ (already generic)
4. **Investor Revenue** (was: Rubicon Revenue)
5. Project Implementation Costs ✓ (already generic)
6. **Investment Drawdown** (was: Rubicon Investment Drawdown)
7. **Investor Net Cash Flow** (was: Rubicon Net Cash Flow)
8. Discount Factor ✓ (already generic)
9. Present Value ✓ (already generic)
10. Cumulative Cash Flow ✓ (already generic)
11. Cumulative PV ✓ (already generic)

### 3. Summary & Results Sheet

**Labels**:
- All labels already generic (NPV, IRR, Payback Period)
- No changes needed

---

## Template Creation Process

### Step 1: Create Base Template
1. Use old version's ExcelExporter (xlsxwriter) as reference
2. Generate template with:
   - 20 years (2025-2044)
   - All formulas in place
   - Generic labels (no "Rubicon")
   - Professional formatting
   - Empty data cells (formulas reference empty cells initially)

### Step 2: Template Validation
1. Open template in Excel
2. Verify all formulas work (even with empty data)
3. Check formatting is professional
4. Ensure no hardcoded values (all references)

### Step 3: Integration with GUI
1. GUI copies template to output location
2. GUI populates data cells (not formula cells)
3. Formulas automatically calculate
4. Save final file

---

## Technical Implementation Strategy

### Current State:
- **Old Version**: Uses `xlsxwriter`, creates from scratch
- **New Version**: Uses `openpyxl`, populates template
- **Issue**: Cell references don't match

### Recommended Fix:

1. **Create Master Template** (using old version's structure):
   - Use `xlsxwriter` to create template once
   - 20 years (2025-2044)
   - Inputs at B3-B8
   - All formulas correct
   - Generic labels
   - Save as `templates/master_template.xlsx`

2. **Update Template-Based Exporter**:
   - Use `openpyxl` to copy template
   - Populate data cells only
   - Don't modify formula cells
   - Match old version's cell references exactly

3. **Benefits**:
   - Template is static (never changes)
   - Export code is simple (just populate)
   - Formulas always correct (in template)
   - Easy to maintain

---

## File Naming Recommendations

### Template File:
- `master_template.xlsx` (if no VBA)
- `master_template.xlsm` (if VBA needed)

### Output Files:
- `{project_name}_analysis.xlsx`
- `{client_name}_carbon_model.xlsx`
- `{date}_carbon_analysis.xlsx`

---

## Version Control Strategy

### Template Versioning:
- Store template in `templates/` directory
- Version in filename: `master_template_v1.0.xlsx`
- Or use git tags for template versions

### Template Updates:
- Create new version when structure changes
- Keep old versions for backward compatibility
- Document changes in template changelog

---

## Distribution Strategy

### For End Users:
1. **Template as Separate Download**:
   - Users can download template independently
   - Can review structure before using GUI
   - Can customize if needed

2. **Template Bundled with GUI**:
   - Template included in GUI package
   - Automatic updates when GUI updates
   - No separate download needed

### Recommended: **Bundled Approach**
- Simpler for users
- Ensures template matches GUI version
- Single download/install

---

## Quality Assurance Checklist

### Template Validation:
- [ ] All formulas work with empty data
- [ ] No hardcoded values (all references)
- [ ] Generic labels (no company-specific terms)
- [ ] Professional formatting applied
- [ ] Correct number of years (20)
- [ ] Correct cell references (B3-B8 for inputs)
- [ ] All sheets present and formatted
- [ ] Charts/visualizations work (if included)

### Export Validation:
- [ ] Template copies correctly
- [ ] Data populates in correct cells
- [ ] Formulas calculate correctly
- [ ] Formatting preserved
- [ ] No circular references
- [ ] All sheets populated correctly

---

## Migration Plan

### Phase 1: Create Generic Template
1. Use old version's ExcelExporter code
2. Replace "Rubicon" with "Investor"
3. Generate template file
4. Validate formulas and formatting

### Phase 2: Update Export Code
1. Update `template_based_export.py` to:
   - Copy master template
   - Populate data cells (match old cell references)
   - Preserve all formulas
   - Apply formatting if needed

### Phase 3: Testing
1. Test with sample data
2. Verify all formulas work
3. Check formatting is professional
4. Validate all sheets populate correctly

### Phase 4: Documentation
1. Update user guide
2. Document template structure
3. Create template customization guide (if needed)

---

## Risk Mitigation

### Potential Issues:

1. **Template Gets Corrupted**:
   - **Mitigation**: Store template in git, version control
   - **Mitigation**: Validate template on startup

2. **Cell References Change**:
   - **Mitigation**: Use named ranges (if possible)
   - **Mitigation**: Document all cell references
   - **Mitigation**: Unit tests for cell references

3. **Formatting Lost**:
   - **Mitigation**: Apply formatting after population
   - **Mitigation**: Use ProfessionalFormatter class

4. **Formula Errors**:
   - **Mitigation**: Validate formulas in template
   - **Mitigation**: Test with sample data

---

## Success Criteria

### Template is Successful If:
- ✅ Works with any carbon credit project data
- ✅ All formulas calculate correctly
- ✅ Professional appearance (investor-ready)
- ✅ No company-specific references
- ✅ Easy to maintain and update
- ✅ Fast export process (< 5 seconds)
- ✅ Compatible with existing GUI workflow

---

## Next Steps

1. **Review This Document**: Confirm approach aligns with goals
2. **Create Generic Template**: Use old version's structure, generic labels
3. **Update Export Code**: Match old version's cell references
4. **Test Thoroughly**: Validate with multiple datasets
5. **Document**: Update user guides and technical docs

---

## Questions to Consider

1. **Company Name**: Should template support custom company names, or always use "Investor"?
2. **Year Count**: Stick with 20 years, or make it configurable?
3. **Template Format**: .xlsx or .xlsm (if VBA needed)?
4. **Distribution**: Bundled with GUI or separate download?
5. **Customization**: Should users be able to customize template?

---

## Conclusion

**Recommended Approach**: **Option 1 - Master Template with Smart Population**

This approach provides:
- ✅ Simplicity and maintainability
- ✅ Consistency across all exports
- ✅ Professional appearance
- ✅ Fast generation
- ✅ Easy updates

The template should:
- Use generic labels ("Investor" instead of "Rubicon")
- Match old version's structure (20 years, correct cell references)
- Include all formulas pre-configured
- Have professional formatting
- Work seamlessly with GUI export process

