import win32com.client
import os

def excel_to_pdf(input_folder, input_file, output_folder, output_file_name):
    """
    Works for converting XLSX files in PDF format
    
    Args:
        input_folder (str): Takes name of folder where xlsx exists
        input_file (str): Takes name of the input file.
        output_folder (str): Takes name of the output folder.
        Output_file (str): Takes name of the output file.
    """
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(BASE_DIR, input_folder, input_file)
    output_dir = os.path.join(BASE_DIR, output_folder)
    output_path = os.path.join(output_dir, output_file_name)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input path does not exist: {input_path}")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        wb = excel.Workbooks.Open(input_path)
        ws = wb.Worksheets[0]
        ws.Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, output_path)  # 0 = PDF
        wb.Close(False)
        excel.Quit()

        print("Conversion Successful!")
    except Exception as e:
        print(f"Failed with error: {e}")

excel_to_pdf("input_files", "SuperStore.xlsx", "outputs", "Super_store.pdf")
