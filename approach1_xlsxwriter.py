"""
Approach 1: Using pandas + XlsxWriter
Creates static pivot tables (as regular data) with charts
Best for: Simple, cross-platform solution with good chart capabilities
"""

import pandas as pd
import xlsxwriter

def create_sample_data():
    """Create sample sales data for demonstration"""
    data = {
        'Date': pd.date_range('2024-01-01', periods=100, freq='D'),
        'Region': ['North', 'South', 'East', 'West'] * 25,
        'Product': ['Product A', 'Product B', 'Product C', 'Product D'] * 25,
        'Sales': [100, 150, 200, 175, 120, 180, 210, 190] * 12 + [100, 150, 200, 175],
        'Quantity': [10, 15, 20, 18, 12, 16, 22, 19] * 12 + [10, 15, 20, 18]
    }
    return pd.DataFrame(data)

def create_pivot_with_charts(input_file=None, output_file='output_xlsxwriter.xlsx'):
    """
    Create pivot tables and charts using XlsxWriter
    
    Args:
        input_file: Path to input Excel file (if None, uses sample data)
        output_file: Path to output Excel file
    """
    
    # Read or create data
    if input_file:
        df = pd.read_excel(input_file)
    else:
        df = create_sample_data()
        print("Using sample data (no input file provided)")
    
    # Create pivot tables using pandas
    pivot1 = pd.pivot_table(
        df, 
        values='Sales', 
        index='Region', 
        columns='Product', 
        aggfunc='sum',
        fill_value=0
    )
    
    pivot2 = pd.pivot_table(
        df,
        values='Quantity',
        index='Product',
        aggfunc=['sum', 'mean', 'count']
    )
    
    # Create Excel file with XlsxWriter
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Write original data
        df.to_excel(writer, sheet_name='Raw Data', index=False)
        
        # Write pivot table 1 (Sales by Region and Product)
        pivot1.to_excel(writer, sheet_name='Pivot - Sales by Region')
        worksheet1 = writer.sheets['Pivot - Sales by Region']
        
        # Create a column chart for pivot1
        chart1 = workbook.add_chart({'type': 'column'})
        
        # Configure the chart from the pivot data
        # Note: Row/col numbers start at 0, and we need to account for headers
        num_products = len(pivot1.columns)
        num_regions = len(pivot1.index)
        
        for col in range(num_products):
            chart1.add_series({
                'name': ['Pivot - Sales by Region', 0, col + 1],
                'categories': ['Pivot - Sales by Region', 1, 0, num_regions, 0],
                'values': ['Pivot - Sales by Region', 1, col + 1, num_regions, col + 1],
            })
        
        chart1.set_title({'name': 'Sales by Region and Product'})
        chart1.set_x_axis({'name': 'Region'})
        chart1.set_y_axis({'name': 'Total Sales'})
        chart1.set_style(11)
        
        worksheet1.insert_chart('H2', chart1, {'x_scale': 1.5, 'y_scale': 1.5})
        
        # Write pivot table 2 (Product Statistics)
        pivot2.to_excel(writer, sheet_name='Pivot - Product Stats')
        worksheet2 = writer.sheets['Pivot - Product Stats']
        
        # Create a bar chart for pivot2 (showing sum of quantities)
        chart2 = workbook.add_chart({'type': 'bar'})
        
        num_products2 = len(pivot2.index)
        
        chart2.add_series({
            'name': 'Total Quantity',
            'categories': ['Pivot - Product Stats', 1, 0, num_products2, 0],
            'values': ['Pivot - Product Stats', 1, 1, num_products2, 1],
        })
        
        chart2.set_title({'name': 'Total Quantity by Product'})
        chart2.set_x_axis({'name': 'Total Quantity'})
        chart2.set_y_axis({'name': 'Product'})
        chart2.set_style(12)
        
        worksheet2.insert_chart('F2', chart2, {'x_scale': 1.5, 'y_scale': 1.5})
        
        # Create a pie chart for the same data
        chart3 = workbook.add_chart({'type': 'pie'})
        
        chart3.add_series({
            'name': 'Quantity Distribution',
            'categories': ['Pivot - Product Stats', 1, 0, num_products2, 0],
            'values': ['Pivot - Product Stats', 1, 1, num_products2, 1],
            'data_labels': {'percentage': True},
        })
        
        chart3.set_title({'name': 'Quantity Distribution by Product'})
        chart3.set_style(10)
        
        worksheet2.insert_chart('F20', chart3, {'x_scale': 1.5, 'y_scale': 1.5})
    
    print(f"✓ Excel file created successfully: {output_file}")
    print(f"  - Sheets: Raw Data, Pivot - Sales by Region, Pivot - Product Stats")
    print(f"  - Charts: Column chart, Bar chart, Pie chart")

# Example usage
if __name__ == "__main__":
    # Method 1: Use sample data
    create_pivot_with_charts()
    
    # Method 2: Use existing Excel file (uncomment to use)
    # create_pivot_with_charts(input_file='your_input_file.xlsx', output_file='output_xlsxwriter.xlsx')
    
    print("\n--- XlsxWriter Approach ---")
    print("Pros:")
    print("  ✓ Simple and intuitive API")
    print("  ✓ Cross-platform (works on Windows, Mac, Linux)")
    print("  ✓ Excellent chart creation capabilities")
    print("  ✓ Good performance")
    print("\nCons:")
    print("  ✗ Cannot create true Excel PivotTable objects")
    print("  ✗ Charts are static (reference fixed data ranges)")
    print("  ✗ Cannot read existing Excel files (write-only)")
