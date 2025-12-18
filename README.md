# Excel Pivot Tables and Charts with Python

This repository contains three different approaches for creating pivot tables and charts in Excel files using Python and pandas. Each approach has its own strengths and use cases.

## 📁 Files Overview

| File | Approach | Key Libraries | Best For |
|------|----------|--------------|----------|
| `approach1_xlsxwriter.py` | Static pivot tables with charts | pandas + XlsxWriter | Cross-platform, simple implementation |
| `approach2_openpyxl.py` | Static pivot tables with charts | pandas + openpyxl | Reading AND writing existing files |
| `approach3_win32com.py` | True Excel PivotTables | pandas + pywin32 | Windows with Excel, interactive pivots |

## 🚀 Quick Start

### Installation

```bash
# For Approach 1 (XlsxWriter)
pip install pandas xlsxwriter openpyxl

# For Approach 2 (openpyxl)
pip install pandas openpyxl

# For Approach 3 (win32com) - Windows only
pip install pandas openpyxl pywin32
```

Or install all dependencies:
```bash
pip install -r requirements.txt
```

### Running the Examples

Each file can be run independently:

```bash
# Approach 1: XlsxWriter
python approach1_xlsxwriter.py

# Approach 2: openpyxl
python approach2_openpyxl.py

# Approach 3: win32com (Windows only, requires Excel)
python approach3_win32com.py
```


---

## 🎯 Decision Tree: Which Approach Should You Use?

```
Are you on Windows with Excel installed?
│
├─ NO → Use Approach 1 (XlsxWriter) or Approach 2 (openpyxl)
│   │
│   └─ Do you need to modify existing Excel files?
│       ├─ YES → Use Approach 2 (openpyxl)
│       └─ NO → Use Approach 1 (XlsxWriter) - Simpler API
│
└─ YES → Do users need interactive PivotTables?
    │
    ├─ YES → Use Approach 3 (win32com)
    │         TRUE PivotTables that users can modify
    │
    └─ NO → Use Approach 1 (XlsxWriter) or Approach 2 (openpyxl)
              Faster and more reliable than win32com
```

## 📝 Usage Examples

### Using Your Own Data

All three approaches support reading from existing Excel files:

```python
# Approach 1
from approach1_xlsxwriter import create_pivot_with_charts
create_pivot_with_charts(input_file='your_data.xlsx', output_file='output.xlsx')

# Approach 2
from approach2_openpyxl import create_pivot_with_openpyxl
create_pivot_with_openpyxl(input_file='your_data.xlsx', output_file='output.xlsx')

# Approach 3
from approach3_win32com import create_true_pivot_with_win32com
create_true_pivot_with_win32com(input_file='your_data.xlsx', output_file='output.xlsx')
```

### Customizing Pivot Tables

Each approach uses pandas for data manipulation, so you can customize the pivot logic:

```python
# Example: Different aggregation functions
pivot = pd.pivot_table(
    df,
    values='Sales',
    index='Region',
    columns='Product',
    aggfunc=['sum', 'mean', 'count'],
    fill_value=0
)

# Example: Multiple value fields
pivot = pd.pivot_table(
    df,
    values=['Sales', 'Quantity'],
    index='Region',
    columns='Product',
    aggfunc='sum'
)
```

