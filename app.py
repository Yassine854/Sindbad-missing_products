from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import os
import shutil
import zipfile
from openpyxl import load_workbook, Workbook
from openpyxl.utils.exceptions import InvalidFileException

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

cam_sites = [
    "CAM01", "CAM02", "CAM03", "CAM04", "CAM05", "CAM06",
    "CAM36", "CAM37", "CAM38", "CAM48", "CAM49"
]

ignored_codes = set([
    "CASHOUMACHLAV","CASNAPBOROVA200M","CASNAPBOROVA250M","CASNAPBORREC250M",
    "CASNAPCAR140M","CASNAPCARCIR160M","CASNAPOVACIR2M","CASNAPRECTCIR20M",
    "CASNAPRECTCIR25M","CASNAPRONCIR140M","CASNAPRONCIR160M","CASTABCUIS",
    "DEKBARQALUM1025L","DEKBARQALUM5069L","FIMBARGL080","GEATBARGALU3CP",
    "GEATBARGALUCAK","GEATBARGALUGM","GEATBARGALUMM","GEATBARGALUPM",
    "GEATBOLALUM87","LBP03BARALUMCKE","LBP03BARALUPOGM","LBP04BAR1POC",
    "LBP04BARALUP2","LBP04BARPIZMD","LBP04PLASPEGRI","LBP05BARALUPOPM",
    "LBPBARCAR50CCCV","LBPBARQALUM504","LBPLT03BARQALCK","LBPPALTALUMRECT",
    "UNVBOWSAL69OZ","UNVECOTRKF1500","UNVECOTRKF500","DIVBARPETL101",
    "FIMBARGL332","LBPBARCAR900GRCV","LBPBARQPET850L","TERBARQPET1000",
    "TERBARQPET2000","UNVBOWSAL43OZ","GEATCVBARGALUGM","GEATCVBARGALUMM",
    "ECOPOCHFRITBL","ECOPOCHMBL155","ECOPOCHMBL255","ECOPOCHMBL355",
    "SITPOIGAT120KF","SITPOIGAT15","SITPOIGAT20","DIVBOUGANN00",
    "DIVBOUGANN02","DIVBOUGANN03","DIVBOUGANN04","DIVBOUGANN05",
    "DIVBOUGANN06","DIVBOUGANN07","DIVBOUGANN08","DIVBOUGANN09",
    "DIVBOUTENIGH","HYP50CVPETTRA35CL","INN50COUVPET","VAR50COVPET25",
    "LBPCOUTBOIS","LBPCUILLBOIS","LBPFOURBOIS","DIVCOUTPLAS100",
    "LBP20COUTPLASNO","WAFCOUPLNOI","DIVCUISOPLAS100","DIVPALEGLACOL",
    "GOLTCUIAGLAGLD","INDCUIACAFTR","LBP20CUILPLSTRG","LBP20CUILPLSTRO",
    "LBPAGITCUILCAFTR","LBPCUIACAF","LBPCUILASOUP100","WAFCUILUXTRP",
    "WAFCUISOPNO","LBP50COUTTRANSP","LBPCOUTPLAST100","DIVFOUPLAS100",
    "LBPFOUCPLAST100","LBPFOUCPLASTBER","WAFFOURLUXTRP","LBPSURBLSJETB",
    "HYGTOQMSCHEF","ALPGANVINYMEDM","BPGANSPCOT","INNGANHDPEL",
    "LBPCOUPBOIS12","LBPPICKBUTTE12","LBPPICKEVANT","CASHOUSMICOND",
    "LBPPOCHBAG","LBPPOCHBAGIMP","CASSETAPTIS3P","EMBTAGITAT",
    "ALPBROCHETT20","ALP100CUREDENTSB","ALP500CUREDENTSB","LBPLBCD500P",
    "DIVPAIFELX100","GOLPAISIMP","GOLSPAIJUS","PAP20DEN075",
    "PAP250DEN075","PAP250DEN085","MINVERPHYR90P","WAFVERPYCRI250",
    "WAFVERPYCRI45","LBP25TASECAPPUC","VAR50GOBCART12","VAR50GOBCART18",
    "VAR50GOBCARTFR25","VAR50GOBCARTFR35","VAR50GOBCARTKF25","VAR50GOBCARTVR22",
    "INN50GOBPET10OZ","INN50GOBPET12OZ","INN50GOBPET14OZ","INN50GOBPET16OZ",
    "ADMPAPHYGIMP4","LBPPAPHGY4CH","MBAPAPHYGCOM","MBAPAPYHG300G",
    "EMESSMAI150F","GEATESSTOUT400","GEATESSTOUT600","LBPESSTTLILATTEX",
    "LBPESSTTSCOM400","LBPPAPESSTCHAR2","MBAESSTTJAB400","MBAESSTTJAB750",
    "ADMPAPSERVREST","GEATPAPSERV3030","MBASERV3030PP","MBASERVROURA",
    "ECORPAPCUI500F","ECOSOPLAIMP32","TTALUMLAF8M","TTALUMSTR08M",
    "DEKALUMFOIDA","TTALUMLAF100M","TTALUMSTR100M","INNPAPALU400",
    "TTALUMSTR12M","TTALUMSTR16M","TTALUMLAF430","TTALUMLAF70M",
    "LBPLOALU12FI12G","INNETIRBLAF08M","TTETIRBLAF08M","INNETIRBLAF100M",
    "LBPLOCUI6FI8G","TTETIRBLAF12M","TTETIRBLAF16M","HYPFILETIR200M",
    "INNETIRBV300M","ECOPAPCUI5MLF","ECOPAPCUI6M","BAPSASCONGSB1L",
    "BAPSASCONGSB2L","HYPSACCUISFOR","BAPSACPB100LB","BAPSACPB50L",
    "BAPSACPB70","BAPSACPB70B","BAPSACPB70N","BAPSACPBHD30L"
])

def _norm(v):
    if v is None:
        return ""
    return str(v).strip().upper()

def _to_float(v):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload/")
async def upload(file: UploadFile = File(...)):
    temp_path = os.path.join(BASE_DIR, f"temp_{file.filename}")

    with open(temp_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        try:
            wb = load_workbook(temp_path, data_only=True)
        except (InvalidFileException, zipfile.BadZipFile):
            return JSONResponse(
                status_code=400,
                content={"error": "Invalid file. Please upload a valid .xlsx Excel file."}
            )
        ws = wb.active
        if ws is None:
            return JSONResponse(
                status_code=400,
                content={"error": "The uploaded Excel file has no active sheet."}
            )

        # pandas header=1 => header is Excel row 2 (1-indexed)
        header_row_idx = 2
        headers = {}
        for col_idx, cell in enumerate(ws[header_row_idx], start=1):
            name = str(cell.value).strip() if cell.value is not None else ""
            if name:
                headers[name] = col_idx

        required = ["Site", "Code", "Désignation Article"]
        missing_cols = [c for c in required if c not in headers]
        if missing_cols:
            return JSONResponse(
                status_code=400,
                content={
                    "error": "Missing required columns in Excel header row (row 2).",
                    "missing": missing_cols,
                    "found": list(headers.keys())
                }
            )

        site_col = headers["Site"]
        code_col = headers["Code"]
        des_col = headers["Désignation Article"]
        stock_col = headers.get("Stock Disponible")

        # Build products sets
        sfx_products = set()
        cam_products_map = {cam: set() for cam in cam_sites}

        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            site = _norm(row[site_col - 1])
            if not site:
                continue

            code = _norm(row[code_col - 1])
            des = _norm(row[des_col - 1])

            if not code or not des:
                continue

            if site == "SFX":
                if stock_col is not None:
                    stock_val = _to_float(row[stock_col - 1])
                    if stock_val is None or stock_val < 4:
                        continue
                sfx_products.add((code, des))
            elif site in cam_products_map:
                cam_products_map[site].add((code, des))

        results = []

        # Ensure output dir clean-ish
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        for cam in cam_sites:
            missing = sfx_products - cam_products_map[cam]
            missing = {p for p in missing if p[0] not in ignored_codes}

            designations = sorted([p[1] for p in missing])

            out_file = os.path.join(OUTPUT_DIR, f"{cam}.xlsx")

            out_wb = Workbook()
            out_ws = out_wb.active
            out_ws.title = "Sheet1"

            out_ws["A1"] = f"CAM Site: {cam}"
            out_ws["A2"] = "Article"

            for i, des in enumerate(designations, start=3):
                out_ws.cell(row=i, column=1, value=des)

            out_wb.save(out_file)

            results.append({
                "cam": cam,
                "count": len(designations),
                "file": f"/download/{cam}.xlsx"
            })

        zip_path = os.path.join(OUTPUT_DIR, "all_cams.zip")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for cam in cam_sites:
                z.write(os.path.join(OUTPUT_DIR, f"{cam}.xlsx"), f"{cam}.xlsx")

        return {"results": results, "zip": "/download/all"}

    finally:
        # Always cleanup temp upload
        if os.path.exists(temp_path):
            os.remove(temp_path)

_ZIP_PATH = os.path.join(OUTPUT_DIR, "all_cams.zip")
_CAM_FILE_MAP: dict[str, str] = {f"{cam}.xlsx": os.path.join(OUTPUT_DIR, f"{cam}.xlsx") for cam in cam_sites}

@app.get("/download/{file}")
def download(file: str):
    if file == "all":
        if not os.path.isfile(_ZIP_PATH):
            return JSONResponse(status_code=404, content={"error": "No processed data available. Please upload a file first."})
        return FileResponse(_ZIP_PATH, filename="all_cams.zip")

    if file not in _CAM_FILE_MAP:
        return JSONResponse(status_code=404, content={"error": "File not found."})

    # Path is looked up from a pre-built map of hardcoded paths; user input is only used as the key.
    path = _CAM_FILE_MAP[file]
    if not os.path.isfile(path):
        return JSONResponse(status_code=404, content={"error": "File not found. Please upload a file first."})
    return FileResponse(path, filename=file)
