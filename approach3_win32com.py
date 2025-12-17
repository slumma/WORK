import pandas as pd
import os
from pathlib import Path

try:
    import win32com.client as win32
    from win32com.client import constants as c
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    print("pywin32 not installed: pip install pywin32")

def create_sample_data():
    """Create sample sales data for demonstration
    """
    data = {
        'Date': pd.date_range('2024-01-01', periods=100, freq='D'),
        'Region': ['North', 'South', 'East', 'West'] * 25,
        'Product': ['Product A', 'Product B', 'Product C', 'Product D'] * 25,
        'Sales': [100, 150, 200, 175, 120, 180, 210, 190] * 12 + [100, 150, 200, 175],
        'Quantity': [10, 15, 20, 18, 12, 16, 22, 19] * 12 + [10, 15, 20, 18]
    }
    return pd.DataFrame(data)

def create_true_pivot_with_win32com(input_file=None, output_file='output_win32com.xlsx'):
    """
    Excel PivotTables and charts
    Args:
        input_file: Path to input Excel file (if None, uses sample data)
        output_file: Path to output Excel file
    """
    
    if not WIN32COM_AVAILABLE:
        print("ERROR: This function requires pywin32 and Excel to be installed.")
        return
    
    # read or create data
    if input_file:
        df = pd.read_excel(input_file)
    else:
        df = create_sample_data()
        print("Using sample data (no input file provided)")
    
    # save data to a temporary Excel file using pandas
    temp_file = 'temp_data_for_pivot.xlsx'
    df.to_excel(temp_file, sheet_name='Raw Data', index=False)
    
    # Convert to absolute path
    abs_output_path = str(Path(output_file).resolve())
    abs_temp_path = str(Path(temp_file).resolve())
    
    # Start Excel application
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # Set to True if you want to see Excel working
    excel.DisplayAlerts = False
    
    try:
        # Open the workbook
        wb = excel.Workbooks.Open(abs_temp_path)
        ws_data = wb.Worksheets('Raw Data')
        
        # Get data range
        last_row = ws_data.Cells(ws_data.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
        last_col = ws_data.Cells(1, ws_data.Columns.Count).End(-4159).Column  # xlToLeft = -4159
        data_range = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(last_row, last_col))
        
        # Create a new sheet for the pivot table
        ws_pivot1 = wb.Worksheets.Add()
        ws_pivot1.Name = "Pivot - Sales by Region"
        
        # Create PivotCache
        pivot_cache = wb.PivotCaches().Create(
            SourceType=1,  # xlDatabase
            SourceData=data_range
        )
        
        # Create PivotTable
        pivot_table1 = pivot_cache.CreatePivotTable(
            TableDestination=ws_pivot1.Range("A3"),
            TableName="SalesByRegion"
        )
        
        # Configure PivotTable fields
        # Add Region to Row
        pivot_table1.PivotFields("Region").Orientation = 1  # xlRowField
        pivot_table1.PivotFields("Region").Position = 1
        
        # Add Product to Column
        pivot_table1.PivotFields("Product").Orientation = 2  # xlColumnField
        pivot_table1.PivotFields("Product").Position = 1
        
        # Add Sales to Values
        pivot_table1.AddDataField(
            pivot_table1.PivotFields("Sales"),
            "Sum of Sales",
            -4157  # xlSum
        )
        
        # Apply style
        pivot_table1.TableStyle2 = "PivotStyleMedium9"
        
        # Create a chart from the PivotTable
        chart1 = ws_pivot1.Shapes.AddChart2(251, 51).Chart  # xlColumnClustered
        chart1.SetSourceData(pivot_table1.TableRange2)
        chart1.ChartType = 51  # xlColumnClustered
        chart1.HasTitle = True
        chart1.ChartTitle.Text = "Sales by Region and Product"
        
        # Position the chart
        chart_shape = chart1.Parent
        chart_shape.Top = ws_pivot1.Range("G3").Top
        chart_shape.Left = ws_pivot1.Range("G3").Left
        chart_shape.Width = 400
        chart_shape.Height = 300
        
        # Create second PivotTable for Product Statistics
        ws_pivot2 = wb.Worksheets.Add()
        ws_pivot2.Name = "Pivot - Product Stats"
        
        pivot_cache2 = wb.PivotCaches().Create(
            SourceType=1,
            SourceData=data_range
        )
        
        pivot_table2 = pivot_cache2.CreatePivotTable(
            TableDestination=ws_pivot2.Range("A3"),
            TableName="ProductStats"
        )
        
        # Configure second PivotTable
        pivot_table2.PivotFields("Product").Orientation = 1  # xlRowField
        
        # Add multiple value fields
        pivot_table2.AddDataField(
            pivot_table2.PivotFields("Quantity"),
            "Sum of Quantity",
            -4157  # xlSum
        )
        
        pivot_table2.AddDataField(
            pivot_table2.PivotFields("Quantity"),
            "Average Quantity",
            -4106  # xlAverage
        )
        
        pivot_table2.TableStyle2 = "PivotStyleMedium2"
        
        # Create a pie chart for product distribution
        chart2 = ws_pivot2.Shapes.AddChart2(201, 5).Chart  # xlPie
        chart2.SetSourceData(pivot_table2.TableRange2)
        chart2.ChartType = 5  # xlPie
        chart2.HasTitle = True
        chart2.ChartTitle.Text = "Quantity Distribution by Product"
        
        # Position the chart
        chart_shape2 = chart2.Parent
        chart_shape2.Top = ws_pivot2.Range("F3").Top
        chart_shape2.Left = ws_pivot2.Range("F3").Left
        chart_shape2.Width = 350
        chart_shape2.Height = 300
        
        # Add data labels to pie chart
        chart2.SeriesCollection(1).HasDataLabels = True
        chart2.SeriesCollection(1).DataLabels().ShowPercentage = True
        
        # Create a summary sheet with instructions
        ws_summary = wb.Worksheets.Add()
        ws_summary.Name = "Summary"
        wb.Worksheets("Summary").Move(Before=wb.Worksheets(1))
        
        ws_summary.Range("A1").Value = "Excel PivotTable Summary"
        ws_summary.Range("A1").Font.Size = 16
        ws_summary.Range("A1").Font.Bold = True
        
        ws_summary.Range("A3").Value = "This workbook contains TRUE Excel PivotTables created with win32com"
        ws_summary.Range("A5").Value = "Features:"
        ws_summary.Range("A6").Value = "✓ Interactive PivotTables that can be modified in Excel"
        ws_summary.Range("A7").Value = "✓ Charts linked to PivotTables (update when pivot refreshes)"
        ws_summary.Range("A8").Value = "✓ Full Excel functionality preserved"
        ws_summary.Range("A9").Value = "✓ Users can drag/drop fields, filter, and refresh data"
        
        ws_summary.Range("A11").Value = "Sheets in this workbook:"
        ws_summary.Range("A12").Value = "• Raw Data: Original data source"
        ws_summary.Range("A13").Value = "• Pivot - Sales by Region: PivotTable with column chart"
        ws_summary.Range("A14").Value = "• Pivot - Product Stats: PivotTable with pie chart"
        
        # Auto-fit columns
        ws_summary.Columns("A:A").ColumnWidth = 60
        
        # Save as new file
        if os.path.exists(abs_output_path):
            os.remove(abs_output_path)
        
        wb.SaveAs(abs_output_path)
        wb.Close()
        
    except Exception as e:
        print(f"ERROR: {e}")
        raise
    
    finally:
        # Clean up
        excel.Quit()
        if os.path.exists(temp_file):
            os.remove(temp_file)

def add_pivot_to_existing_file(input_file, output_file='updated_win32com.xlsx'):
    """
    Add PivotTable to an existing Excel file
    
    Args:
        input_file: Path to existing Excel file
        output_file: Path to save updated file
    """
    
    if not WIN32COM_AVAILABLE:
        print("ERROR: This function requires pywin32 and Excel to be installed.")
        return
    
    abs_input_path = str(Path(input_file).resolve())
    abs_output_path = str(Path(output_file).resolve())
    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(abs_input_path)
        ws_data = wb.Worksheets(1)  # First sheet
        
        # Get data range
        last_row = ws_data.Cells(ws_data.Rows.Count, 1).End(-4162).Row
        last_col = ws_data.Cells(1, ws_data.Columns.Count).End(-4159).Column
        data_range = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(last_row, last_col))
        
        # Create new sheet for pivot
        ws_pivot = wb.Worksheets.Add()
        ws_pivot.Name = "New PivotTable"
        
        # Create PivotTable
        pivot_cache = wb.PivotCaches().Create(
            SourceType=1,
            SourceData=data_range
        )
        
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=ws_pivot.Range("A3"),
            TableName="NewPivot"
        )
        
        # Configure with first available fields
        # This is a generic example - adjust based on your data
        fields = [pivot_table.PivotFields(i) for i in range(1, min(4, last_col + 1))]
        
        if len(fields) >= 2:
            fields[0].Orientation = 1  # Row field
            fields[1].Orientation = 4  # Data field
        
        wb.SaveAs(abs_output_path)
        wb.Close()
        
        print(f"PivotTable added to existing file: {output_file}")
        
    except Exception as e:
        print(f"ERROR: {e}")
        raise
    
    finally:
        excel.Quit()

# main
if __name__ == "__main__":
    if not WIN32COM_AVAILABLE:
        print("ERROR: pywin32 is not installed or import failed")

    else:
        # Create from sample data -> for example
        create_true_pivot_with_win32com()  #add input file here as argument 
        
        # Add pivot to existing file 
        # add_pivot_to_existing_file('your_input_file.xlsx', 'updated_win32com.xlsx')
        
        # Create from existing Excel file
        # create_true_pivot_with_win32com(input_file='your_input_file.xlsx', output_file='output_win32com.xlsx')

