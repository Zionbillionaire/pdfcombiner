import os
import win32com.client as win32

def excel_to_pdf(input_dir, output_dir):
    # Ensure output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Create a client to interact with Microsoft Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False

    # Process each Excel file in the directory
    for file in os.listdir(input_dir):
        if file.endswith(".xlsx"):
            file_path = os.path.join(input_dir, file)
            output_path = os.path.join(output_dir, file.replace('.xlsx', '.pdf'))
            
            # Open the workbook
            workbook = excel.Workbooks.Open(file_path)
            
            # Adjust all worksheets to fit content on one page
            for sheet in workbook.Sheets:
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1
            
            # Save as PDF
            workbook.ExportAsFixedFormat(0, output_path)
            
            # Close the workbook
            workbook.Close()

    # Quit Excel
    excel.Quit()

input_dir = 'input-path'
output_dir = 'output-path'

excel_to_pdf(input_dir, output_dir)
