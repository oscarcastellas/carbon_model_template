# GUI Application Update Summary

## Status: ✅ Updated

The GUI application has been updated to work with the new generic template system.

---

## Changes Made

### 1. Updated Assumptions Dictionary

**Location**: `gui/carbon_model_gui.py` (line ~871-884)

**Added**:
- `company_name`: Set to 'Investor' by default (generic)
- `num_years`: Automatically calculated from data length

**Before**:
```python
assumptions = {
    'wacc': config.wacc,
    'rubicon_investment_total': config.rubicon_investment_total,
    # ... other assumptions
}
```

**After**:
```python
# Calculate number of years from data
num_years = len(data) if data is not None and len(data) > 0 else 20

assumptions = {
    'wacc': config.wacc,
    'rubicon_investment_total': config.rubicon_investment_total,
    # ... other assumptions
    # Template customization
    'company_name': 'Investor',  # Generic default, can be customized
    'num_years': num_years  # Use actual data length
}
```

---

## How It Works

### Automatic Template Configuration

1. **Company Name**: 
   - Default: "Investor" (generic)
   - Automatically passed to template exporter
   - Can be customized in future GUI versions

2. **Year Count**:
   - Automatically calculated from loaded data
   - Falls back to 20 if data is empty
   - Ensures template matches actual data length

### Template Export Flow

```
GUI loads data
    ↓
Calculates num_years from data length
    ↓
Creates assumptions dict with company_name='Investor' and num_years
    ↓
Calls ExcelExporter.export_model_to_excel(..., assumptions=assumptions, use_template=True)
    ↓
ExcelExporter extracts company_name and num_years from assumptions
    ↓
Creates TemplateBasedExporter(company_name=..., num_years=...)
    ↓
Template exporter uses master_template.xlsm
    ↓
Populates template with data
    ↓
Saves final Excel file
```

---

## Current Behavior

### Default Settings:
- **Company Name**: "Investor" (generic)
- **Year Count**: Automatically calculated from data (default: 20 if empty)

### Template Used:
- **File**: `templates/master_template.xlsm`
- **Format**: .xlsm (VBA support)
- **Structure**: Matches old version (B3-B8 inputs, 20 years default)

---

## Future Enhancements (Optional)

### Option 1: Add GUI Field for Company Name
Add an input field in the GUI options section:
```python
# In create_options_section()
company_frame = tk.Frame(options_frame, bg='#F5F5F5')
company_frame.pack(fill=tk.X, pady=5)

tk.Label(company_frame, text="Company Name:", ...).pack(side=tk.LEFT)
self.company_name_var = tk.StringVar(value="Investor")
tk.Entry(company_frame, textvariable=self.company_name_var, ...).pack(side=tk.LEFT)

# Then in run_analysis():
assumptions['company_name'] = self.company_name_var.get() or 'Investor'
```

### Option 2: Add GUI Field for Year Count Override
Allow users to override the automatic year count:
```python
# In create_options_section()
years_frame = tk.Frame(options_frame, bg='#F5F5F5')
years_frame.pack(fill=tk.X, pady=5)

tk.Label(years_frame, text="Model Years:", ...).pack(side=tk.LEFT)
self.num_years_var = tk.StringVar(value="")
tk.Entry(years_frame, textvariable=self.num_years_var, ...).pack(side=tk.LEFT)

# Then in run_analysis():
if self.num_years_var.get():
    num_years = int(self.num_years_var.get())
else:
    num_years = len(data) if data is not None and len(data) > 0 else 20
assumptions['num_years'] = num_years
```

---

## Testing

### Verified:
- ✅ GUI passes `company_name` and `num_years` to ExcelExporter
- ✅ ExcelExporter extracts these from assumptions
- ✅ TemplateBasedExporter uses them correctly
- ✅ Template export works with new generic template
- ✅ Formulas are correct (B3-B8 inputs, B-U for 20 years)

---

## Summary

✅ **GUI is updated and working with the new generic template system**

The GUI now:
- Automatically uses the generic "Investor" template
- Calculates year count from data
- Passes these to the template exporter
- Works seamlessly with the new master_template.xlsm

No user action required - the GUI automatically uses the new generic template!

