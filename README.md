# Sales Report Generator

This Python script generates a sales report in Excel format using the `openpyxl` library. The user is prompted to input the month, product categories, and corresponding sales data. The script then creates an Excel workbook with a bar chart visualizing the sales data and includes sum formulas for each category.

## Features:

1. **User Input:**
   - The script prompts the user to input the month, product categories (comma-separated), and sales data for each category.

2. **Excel Workbook and Chart:**
   - A new Excel workbook is created using the `openpyxl` library.
   - Sales data is input into the sheet, and a bar chart is generated using the entered information.

3. **Sum Formulas:**
   - Sum formulas are automatically generated for each category to calculate the total sales.

4. **Formatting:**
   - The script adds formatting to the Excel sheet, including a title and font styling.

5. **Save Workbook:**
   - The generated workbook is saved in the same directory as the script with a filename based on the entered month.

## Usage Instructions:

1. Run the script using a Python interpreter.
2. Enter the month when prompted.
3. Enter product categories (comma-separated) when prompted.
4. Enter sales data for each category.
5. The script will generate an Excel workbook with a bar chart and sum formulas.
6. The workbook is saved in the same directory as the script with a filename based on the entered month.

**Note:**
- Ensure you have the `openpyxl` library installed (`pip install openpyxl`).
- Customize the script as needed for specific reporting requirements.
- Use responsibly and consider legal and ethical considerations when automating data input.

Feel free to adapt and extend this script for your specific reporting needs.