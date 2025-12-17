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

## 📊 Approach Comparison

### Approach 1: pandas + XlsxWriter

**File:** `approach1_xlsxwriter.py`

**What it creates:**
- Pivoted data as regular Excel tables (using pandas pivot_table)
- Static charts that reference the pivoted data
- Column charts, bar charts, and pie charts

**Pros:**
- ✅ Simple and intuitive API
- ✅ Cross-platform (Windows, Mac, Linux)
- ✅ Excellent chart creation capabilities
- ✅ Good performance
- ✅ Clean, readable code

**Cons:**
- ❌ Cannot create true Excel PivotTable objects
- ❌ Charts are static (reference fixed data ranges)
- ❌ Cannot read existing Excel files (write-only)

**When to use:**
- You need cross-platform compatibility
- You want simple, maintainable code
- Static reports are sufficient (users don't need to modify pivots)
- You're creating new Excel files from scratch

**Example output:**
- `output_xlsxwriter.xlsx` with multiple sheets and charts

---

### Approach 2: pandas + openpyxl

**File:** `approach2_openpyxl.py`

**What it creates:**
- Pivoted data as regular Excel tables
- Bar charts and pie charts
- Can read and modify existing Excel files

**Pros:**
- ✅ Can read AND write Excel files
- ✅ Modify existing workbooks without losing formatting
- ✅ Full control over Excel formatting
- ✅ Cross-platform
- ✅ Active development and community support

**Cons:**
- ❌ Creating true PivotTables is extremely complex with openpyxl
- ❌ Chart API is less intuitive than XlsxWriter
- ❌ Requires more code for the same results
- ❌ Documentation for charts can be sparse

**When to use:**
- You need to update existing Excel files
- You want to read data from Excel and write back to it
- You need cross-platform compatibility
- You need fine control over Excel cell formatting

**Example output:**
- `output_openpyxl.xlsx` with pivoted data and charts
- `updated_openpyxl.xlsx` when updating existing files

---

### Approach 3: pandas + win32com (pywin32)

**File:** `approach3_win32com.py`

**What it creates:**
- **TRUE Excel PivotTable objects** (fully interactive)
- Charts linked to PivotTables
- Full Excel automation capabilities

**Pros:**
- ✅ Creates actual PivotTable objects (not just pivoted data)
- ✅ PivotTables are fully interactive in Excel
- ✅ Users can drag/drop fields, filter, refresh data
- ✅ Charts update automatically when PivotTable refreshes
- ✅ Access to ALL Excel features and functionality
- ✅ Can automate any Excel task

**Cons:**
- ❌ **Windows ONLY** (requires Excel installed)
- ❌ More complex API (COM automation)
- ❌ Slower than other approaches
- ❌ Can have stability issues with large operations
- ❌ Requires Excel to be installed on the machine

**When to use:**
- You're on Windows with Excel installed
- Users need TRUE interactive PivotTables
- Users need to modify/refresh PivotTables
- You need advanced Excel automation
- You want PivotTables that behave exactly like manually created ones

**Example output:**
- `output_win32com.xlsx` with interactive PivotTables

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

## 📈 About PySpark

**Note:** PySpark is mentioned in the task but is primarily for big data processing, not direct Excel manipulation.

**When to use PySpark:**
- Processing very large datasets (GBs to TBs)
- Data doesn't fit in memory
- Distributed computing needed

**How to integrate with Excel:**
```python
# Process large data with PySpark
spark_df = spark.read.csv('large_data.csv')
processed_df = spark_df.groupBy('Region').sum('Sales')

# Convert to pandas for Excel export
pandas_df = processed_df.toPandas()

# Then use any of the three approaches above
# create_pivot_with_charts(pandas_df)
```

## 🔧 Common Issues and Solutions

### Issue 1: Module not found
```bash
# Solution: Install required packages
pip install pandas xlsxwriter openpyxl pywin32
```

### Issue 2: win32com not working on Windows
```bash
# Solution: Run this after installing pywin32
python -m win32com.client.makepy.py
```

### Issue 3: Charts not appearing
- Check that data ranges are correct
- Ensure pivot tables have data (not empty)
- Verify chart is positioned within Excel limits

### Issue 4: Permission errors when saving files
- Close the output Excel file if it's open
- Check file permissions
- Use a different output filename

## 📚 Key Pandas Pivot Table Methods

```python
# Basic pivot
pd.pivot_table(df, values='Sales', index='Region', aggfunc='sum')

# Multiple aggregations
pd.pivot_table(df, values='Sales', index='Region', aggfunc=['sum', 'mean', 'count'])

# Multiple dimensions
pd.pivot_table(df, values='Sales', index='Region', columns='Product', aggfunc='sum')

# Multiple value fields
pd.pivot_table(df, values=['Sales', 'Quantity'], index='Region', aggfunc='sum')

# Custom aggregations
pd.pivot_table(df, values='Sales', index='Region', aggfunc=lambda x: x.max() - x.min())
```

## 🎨 Chart Types Available

All approaches support various chart types:

- **Column/Bar Charts** - Good for comparing values across categories
- **Line Charts** - Good for trends over time
- **Pie Charts** - Good for showing proportions
- **Area Charts** - Good for cumulative totals
- **Scatter Charts** - Good for correlations
- **Combo Charts** - Combining multiple chart types

## 📖 Additional Resources

- [pandas pivot_table documentation](https://pandas.pydata.org/docs/reference/api/pandas.pivot_table.html)
- [XlsxWriter documentation](https://xlsxwriter.readthedocs.io/)
- [openpyxl documentation](https://openpyxl.readthedocs.io/)
- [pywin32 documentation](https://github.com/mhammond/pywin32)

## 💡 Best Practices

1. **Always close file handles** - Use context managers (`with` statements)
2. **Test with small data first** - Especially with win32com
3. **Handle errors gracefully** - Excel operations can fail for many reasons
4. **Keep raw data** - Always preserve original data in a separate sheet
5. **Document your pivots** - Add sheets explaining what was calculated
6. **Use meaningful names** - For sheets, charts, and pivot tables

## 🤝 Contributing

Feel free to extend these examples with:
- Additional chart types
- More complex pivot configurations
- Error handling improvements
- Performance optimizations

## 📄 License

These examples are provided as-is for educational and commercial use.

---

**Summary:** Choose XlsxWriter for simplicity, openpyxl for reading/writing existing files, and win32com when you need TRUE interactive Excel PivotTables on Windows.
