from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import pandas as pd
import os
import shutil

app = FastAPI()

OUTPUT_DIR = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

cam_sites = ["CAM01","CAM02","CAM03","CAM04","CAM05","CAM06",
             "CAM36","CAM37","CAM38","CAM48","CAM49"]

ignored_codes = [ "CASHOUMACHLAV", "CASNAPBOROVA200M", "CASNAPBOROVA250M" ]  # shorten here (keep full list)

@app.get("/")
def home():
    return {"message": "Excel Processor API is running"}

@app.post("/process/")
async def process_file(file: UploadFile = File(...)):
    
    file_path = f"temp_{file.filename}"
    
    # Save uploaded file
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    df = pd.read_excel(file_path, header=1)
    df.columns = df.columns.str.strip()

    df["Site"] = df["Site"].str.strip().str.upper()
    df["Désignation Article"] = df["Désignation Article"].str.strip().str.upper()

    sfx_df = df[df["Site"] == "SFX"].copy()

    sfx_df["Stock Disponible"] = pd.to_numeric(
        sfx_df["Stock Disponible"], errors='coerce'
    )

    sfx_filtered = sfx_df[sfx_df["Stock Disponible"] >= 4]

    sfx_products = set(zip(
        sfx_filtered["Code"],
        sfx_filtered["Désignation Article"]
    ))

    generated_files = []

    for cam in cam_sites:
        cam_products = set(zip(
            df[df["Site"] == cam]["Code"],
            df[df["Site"] == cam]["Désignation Article"]
        ))

        missing = sfx_products - cam_products
        missing = {p for p in missing if p[0] not in ignored_codes}

        designations = [p[1] for p in missing]

        result_df = pd.DataFrame({"Article": designations})

        output_file = os.path.join(OUTPUT_DIR, f"{cam}.xlsx")
        result_df.to_excel(output_file, index=False)

        generated_files.append(output_file)

    os.remove(file_path)

    return {
        "message": "Processing complete",
        "files": generated_files
    }