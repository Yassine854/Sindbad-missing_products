import os
import zipfile

from fastapi import UploadFile, HTTPException
from fastapi.responses import JSONResponse
import openpyxl
import xlrd

# Constants
# (Assuming existing constants remain unchanged)

async def handle_upload(file: UploadFile):
    if not (file.filename.endswith('.xls') or file.filename.endswith('.xlsx')):
        return JSONResponse(status_code=400, content={"message": "Invalid file type"})

    temp_file_path = f'/tmp/{file.filename}'
    with open(temp_file_path, 'wb') as buffer:
        shutil.copyfileobj(file.file, buffer)

    headers_index_map = {}
    missing_products = {}

    try:
        if file.filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(temp_file_path)
            sheet = wb.active
            headers = [cell.value for cell in sheet[2]]  # Row 2 as header
            headers_index_map = {header: index + 1 for index, header in enumerate(headers)}
            # Normalize and process...
        elif file.filename.endswith('.xls'):
            wb = xlrd.open_workbook(temp_file_path)
            sheet = wb.sheet_by_index(0)
            headers = sheet.row_values(1)  # Row 2 as header
            headers_index_map = {header: index for index, header in enumerate(headers)}
            # Normalize and process...

        # Check for required columns
        required_columns = ["Site", "Code", "Désignation Article"]
        for column in required_columns:
            if column not in headers_index_map:
                return JSONResponse(status_code=400, content={"message": f'Missing column: {column}'})

        # Implement the logic for dealing with rows and processing data
        # Normalization, filtering, computing missing products...

        # Output results to separate Excel files per cam_sites (Stubbed output)
        output_files = []
        zip_file_path = '/tmp/all_cams.zip'
        with zipfile.ZipFile(zip_file_path, 'w') as zipf:
            for cam_site in missing_products.keys():
                output_wb = openpyxl.Workbook()
                output_path = f'/tmp/{cam_site}.xlsx'
                # Populate Workbook...
                output_wb.save(output_path)
                zipf.write(output_path, os.path.basename(output_path))
                output_files.append(output_path)

    finally:
        os.remove(temp_file_path)  # Ensure cleanup of temp file
        for output_file in output_files:
            os.remove(output_file)  # Clean up generated files after zipping

    return zip_file_path  # Return path to zip file containing results