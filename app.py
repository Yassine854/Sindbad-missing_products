from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates
import pandas as pd
import os
import shutil
import zipfile

app = FastAPI()

templates = Jinja2Templates(directory="templates")

OUTPUT_DIR = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

cam_sites = [
    "CAM01","CAM02","CAM03","CAM04","CAM05","CAM06",
    "CAM36","CAM37","CAM38","CAM48","CAM49"
]

ignored_codes = [
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
]

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/upload/")
async def upload(file: UploadFile = File(...)):

    temp_path = f"temp_{file.filename}"

    with open(temp_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    df = pd.read_excel(temp_path, header=1)
    df.columns = df.columns.str.strip()

    df["Site"] = df["Site"].str.strip().str.upper()
    df["Désignation Article"] = df["Désignation Article"].str.strip().str.upper()

    sfx = df[df["Site"] == "SFX"].copy()
    sfx["Stock Disponible"] = pd.to_numeric(sfx["Stock Disponible"], errors="coerce")
    sfx = sfx[sfx["Stock Disponible"] >= 4]

    sfx_products = set(zip(sfx["Code"], sfx["Désignation Article"]))

    results = []

    for cam in cam_sites:

        cam_products = set(zip(
            df[df["Site"] == cam]["Code"],
            df[df["Site"] == cam]["Désignation Article"]
        ))

        missing = sfx_products - cam_products
        missing = {p for p in missing if p[0] not in ignored_codes}

        designations = [p[1] for p in missing]

        out_file = os.path.join(OUTPUT_DIR, f"{cam}.xlsx")
        pd.DataFrame({"Article": designations}).to_excel(out_file, index=False)

        results.append({
            "cam": cam,
            "count": len(designations),
            "file": f"/download/{cam}.xlsx"
        })

    # ZIP
    zip_path = os.path.join(OUTPUT_DIR, "all_cams.zip")

    with zipfile.ZipFile(zip_path, "w") as z:
        for cam in cam_sites:
            z.write(os.path.join(OUTPUT_DIR, f"{cam}.xlsx"), f"{cam}.xlsx")

    os.remove(temp_path)

    return {
        "results": results,
        "zip": "/download/all"
    }


@app.get("/download/{file}")
def download(file: str):

    if file == "all":
        return FileResponse(
            os.path.join(OUTPUT_DIR, "all_cams.zip"),
            filename="all_cams.zip"
        )

    return FileResponse(
        os.path.join(OUTPUT_DIR, file),
        filename=file
    )