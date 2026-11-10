import json
import os
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def flatten_json(y, separator="_"):
    """
    Recursively flattens nested dictionaries and lists inside a JSON object.
    """
    out = {}

    def flatten(x, name=""):
        # If it's a nested dictionary
        if type(x) is dict:
            for a in x:
                flatten(x[a], name + a + separator)
        # If it's a nested list
        elif type(x) is list:
            i = 0
            for a in x:
                flatten(a, name + str(i) + separator)
                i += 1
        # Base case: primitive data types
        else:
            out[name[:-1]] = x

    flatten(y)
    return out

def convert_json_to_excel(json_filepath, excel_filepath, sheet_name="Data Report"):
    """
    Reads a JSON file (single object or array), flattens it, 
    and exports it as a beautifully styled Excel workbook.
    """
    print(f"🔄 Loading: {json_filepath}...")
    with open(json_filepath, 'user', encoding='utf-8') as f:
        data = json.load(f)
    
    # Ensure data is a list of objects for tabular format
    if isinstance(data, dict):
        data = [data]
        
    # Flatten every record in the array
    print("🧠 Flattening complex nested structures...")
    flattened_data = [flatten_json(record) for record in data]
    
    # Create DataFrame
    df = pd.DataFrame(flattened_data)
    
    # Clean up header names (capitalize words, replace separators with spaces)
    df.columns = [col.replace('_', ' ').title() for col in df.columns]
    
    print(f"📊 Exporting styled Excel sheet to: {excel_filepath}...")
    
    # Write to Excel using openpyxl engine for styling
    with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Access the openpyxl objects to apply professional styling
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Define Styles
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # Classic Corporate Navy
        data_font = Font(name="Segoe UI", size=10)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        
        # Format Headers
        for col_num in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            worksheet.row_dimensions[1].height = 28
            
        # Format Data Rows & Auto-fit Columns
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.font = data_font
                cell.border = thin_border
                # Left-align text, right-align numbers
                if isinstance(cell.value, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
        
        # Auto-adjust column widths dynamically based on content length
        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            # Add padding so data isn't cramped
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    print("✅ Done! Professional spreadsheet generated.")

# Example execution usage
if __name__ == "__main__":
    # Mock complex nested JSON data for testing
    sample_json = [
        {
            "id": 101,
            "product": "Enterprise Cloud Server",
            "pricing": {"monthly": 89.99, "yearly": 950.00},
            "specs": {"ram": "32GB", "storage": ["1TB NVMe", "2TB HDD"]},
            "status": "active"
        },
        {
            "id": 102,
            "product": "Business VPN Gateway",
            "pricing": {"monthly": 24.50, "yearly": 240.00},
            "specs": {"ram": "8GB", "storage": ["256GB SSD"]},
            "status": "pending"
        }
    ]
    
    # Write mock data to a temporary file
    with open("sample_input.json", "w") as f:
        json.dump(sample_json, f, indent=4)
        
    # Run the converter script
    convert_json_to_excel("sample_input.json", "Professional_Data_Report.xlsx")
