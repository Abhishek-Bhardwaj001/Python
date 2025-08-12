from docx2pdf import convert
import os

def convert_docx_to_pdf(input_folder:str,input_file:str | None=None, output_folder:str|None=None):
    """
    The convert_docx_to_pdf is a wrapper function wrapped around the class docx2pdf for converting docx files into a pdf

    Usage:
    # Useful when working with data validation of content extracted from Word documents
    # Useful for file formats conversion of word to pdf for sharing, archiving and security from accidental edits.

    Args:
    1.Input folder (required | Data Type: str): Name of the folder where your docx is present.
    2.input file (optional| Data Type: str): Name of the docx file.
    3.Output_folder (optional| Data Type: str): Name of the Output Folder

    Working:
    Converts 1 file if the file name is provided.
    Converts all the file if only the input folder is provided.
    Generates output in the input folder if no output folder is provided
    """
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    if input_file:
        input_path = os.path.join(BASE_DIR,input_folder,input_file)
    else:
        input_path = os.path.join(BASE_DIR,input_folder)

    if output_folder:
        output_path = os.path.join(BASE_DIR, output_folder)
    else:
        output_path=input_path

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input path does not exist: {input_path}")
    
    if output_folder and not os.path.exists(output_path):
        os.makedirs(output_path, exist_ok=True)

    try:
        convert(input_path, output_path)
        print(f"Converted succesfully:{input_path}-> {output_path}")
    except Exception as e:
        print(f"Error during conversion: {e}")




convert_docx_to_pdf("input_files",output_folder="outputs")

