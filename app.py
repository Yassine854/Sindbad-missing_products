# main.py
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

import os
import shutil
import zipfile
from datetime import datetime
import re

from openpyxl import load_workbook, Workbook
from openpyxl.utils.exceptions import InvalidFileException
import xlrd

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

# Mount static files for logo
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(STATIC_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

LOGO_PATH = os.path.join(STATIC_DIR, "sindbad.jpg")

# Depot to CAM mapping
DEPOT_CAMS = {
    "SFX": ["CAM01", "CAM02", "CAM03", "CAM04", "CAM05", "CAM06", "CAM07", "CAM36", "CAM37", "CAM38", "CAM48", "CAM49", "CAM58", "CAM59", "CAM65"],
    "SSE": ["CAM08", "CAM09", "CAM10", "CAM11", "CAM12", "CAM13", "CAM14", "CAM15", "CAM46", "CAM47", "CAM63"],
    "TUN": ["CAM16", "CAM17", "CAM18", "CAM19", "CAM20", "CAM21", "CAM22", "CAM23", "CAM24", "CAM25", "CAM26", "CAM27", "CAM28", "CAM29", "CAM30", "CAM31", "CAM32", "CAM33", "CAM34", "CAM35", "CAM64"],
    "BJSSE": ["CAM39"],
    "MAH": ["CAM40", "CAM41", "CAM42", "CAM43", "CAM44", "CAM45", "CAM57"],
    "NAB": ["CAM50", "CAM51", "CAM52", "CAM53", "CAM54", "CAM55"],
}

# All CAMs (for backwards compatibility)
all_cam_sites = []
for cams in DEPOT_CAMS.values():
    all_cam_sites.extend(cams)

ignored_codes = [
    "CASHOUMACHLAV", "CASNAPBOROVA200M", "CASNAPBOROVA250M", "CASNAPBORREC250M",
    "CASNAPCAR140M", "CASNAPCARCIR160M", "CASNAPOVACIR2M", "CASNAPRECTCIR20M",
    "CASNAPRECTCIR25M", "CASNAPRONCIR140M", "CASNAPRONCIR160M", "CASTABCUIS",
    "DEKBARQALUM1025L", "DEKBARQALUM5069L", "FIMBARGL080", "GEATBARGALU3CP",
    "GEATBARGALUCAK", "GEATBARGALUGM", "GEATBARGALUMM", "GEATBARGALUPM",
    "GEATBOLALUM87", "LBP03BARALUMCKE", "LBP03BARALUPOGM", "LBP04BAR1POC",
    "LBP04BARALUP2", "LBP04BARPIZMD", "LBP04PLASPEGRI", "LBP05BARALUPOPM",
    "LBPBARCAR50CCCV", "LBPBARQALUM504", "LBPLT03BARQALCK", "LBPPALTALUMRECT",
    "UNVBOWSAL69OZ", "UNVECOTRKF1500", "UNVECOTRKF500", "DIVBARPETL101",
    "FIMBARGL332", "LBPBARCAR900GRCV", "LBPBARQPET850L", "TERBARQPET1000",
    "TERBARQPET2000", "UNVBOWSAL43OZ", "GEATCVBARGALUGM", "GEATCVBARGALUMM",
    "ECOPOCHFRITBL", "ECOPOCHMBL155", "ECOPOCHMBL255", "ECOPOCHMBL355",
    "SITPOIGAT120KF", "SITPOIGAT15", "SITPOIGAT20", "DIVBOUGANN00",
    "DIVBOUGANN02", "DIVBOUGANN03", "DIVBOUGANN04", "DIVBOUGANN05",
    "DIVBOUGANN06", "DIVBOUGANN07", "DIVBOUGANN08", "DIVBOUGANN09",
    "DIVBOUTENIGH", "HYP50CVPETTRA35CL", "INN50COUVPET", "VAR50COVPET25",
    "LBPCOUTBOIS", "LBPCUILLBOIS", "LBPFOURBOIS", "DIVCOUTPLAS100",
    "LBP20COUTPLASNO", "WAFCOUPLNOI", "DIVCUISOPLAS100", "DIVPALEGLACOL",
    "GOLTCUIAGLAGLD", "INDCUIACAFTR", "LBP20CUILPLSTRG", "LBP20CUILPLSTRO",
    "LBPAGITCUILCAFTR", "LBPCUIACAF", "LBPCUILASOUP100", "WAFCUILUXTRP",
    "WAFCUISOPNO", "LBP50COUTTRANSP", "LBPCOUTPLAST100", "DIVFOUPLAS100",
    "LBPFOUCPLAST100", "LBPFOUCPLASTBER", "WAFFOURLUXTRP", "LBPSURBLSJETB",
    "HYGTOQMSCHEF", "ALPGANVINYMEDM", "BPGANSPCOT", "INNGANHDPEL",
    "LBPCOUPBOIS12", "LBPPICKBUTTE12", "LBPPICKEVANT", "CASHOUSMICOND",
    "LBPPOCHBAG", "LBPPOCHBAGIMP", "CASSETAPTIS3P", "EMBTAGITAT",
    "ALPBROCHETT20", "ALP100CUREDENTSB", "ALP500CUREDENTSB", "LBPLBCD500P",
    "DIVPAIFELX100", "GOLPAISIMP", "GOLSPAIJUS", "PAP20DEN075",
    "PAP250DEN075", "PAP250DEN085", "MINVERPHYR90P", "WAFVERPYCRI250",
    "WAFVERPYCRI45", "LBP25TASECAPPUC", "VAR50GOBCART12", "VAR50GOBCART18",
    "VAR50GOBCARTFR25", "VAR50GOBCARTFR35", "VAR50GOBCARTKF25", "VAR50GOBCARTVR22",
    "INN50GOBPET10OZ", "INN50GOBPET12OZ", "INN50GOBPET14OZ", "INN50GOBPET16OZ",
    "ADMPAPHYGIMP4", "LBPPAPHGY4CH", "MBAPAPHYGCOM", "MBAPAPYHG300G",
    "EMESSMAI150F", "GEATESSTOUT400", "GEATESSTOUT600", "LBPESSTTLILATTEX",
    "LBPESSTTSCOM400", "LBPPAPESSTCHAR2", "MBAESSTTJAB400", "MBAESSTTJAB750",
    "ADMPAPSERVREST", "GEATPAPSERV3030", "MBASERV3030PP", "MBASERVROURA",
    "ECORPAPCUI500F", "ECOSOPLAIMP32", "TTALUMLAF8M", "TTALUMSTR08M",
    "DEKALUMFOIDA", "TTALUMLAF100M", "TTALUMSTR100M", "INNPAPALU400",
    "TTALUMSTR12M", "TTALUMSTR16M", "TTALUMLAF430", "TTALUMLAF70M",
    "LBPLOALU12FI12G", "INNETIRBLAF08M", "TTETIRBLAF08M", "INNETIRBLAF100M",
    "LBPLOCUI6FI8G", "TTETIRBLAF12M", "TTETIRBLAF16M", "HYPFILETIR200M",
    "INNETIRBV300M", "ECOPAPCUI5MLF", "ECOPAPCUI6M", "BAPSASCONGSB1L",
    "BAPSASCONGSB2L", "HYPSACCUISFOR", "BAPSACPB100LB", "BAPSACPB50L",
    "BAPSACPB70", "BAPSACPB70B", "BAPSACPB70N", "BAPSACPBHD30L"
]

IGNORED_CODES_SET = set(ignored_codes)
PRODUCTS_PER_COLUMN = 50


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


def _headers_from_row(values):
    headers = {}
    for i, val in enumerate(values, start=1):
        name = str(val).strip() if val is not None else ""
        if name:
            headers[name] = i
    return headers


def _extract_rows_xlsx(path: str):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    if ws is None:
        raise ValueError("Le fichier Excel n'a pas de feuille active.")

    header_row_idx = 2
    header_values = [c.value for c in ws[header_row_idx]]
    headers = _headers_from_row(header_values)

    def iter_data_rows():
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            yield row

    return headers, iter_data_rows()


def _extract_rows_xls(path: str):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)

    header_row_idx0 = 1
    header_values = sheet.row_values(header_row_idx0)
    headers = _headers_from_row(header_values)

    def iter_data_rows():
        for r in range(header_row_idx0 + 1, sheet.nrows):
            yield sheet.row_values(r)

    return headers, iter_data_rows()

def _read_stock_time(path: str, ext: str) -> str:
    """Read date from C1 and time from F1"""
    try:
        date_val = None
        time_val = None
        
        if ext == ".xlsx":
            wb = load_workbook(path, data_only=True)
            ws = wb.active
            if ws is None:
                return ""
            
            date_val = ws.cell(row=1, column=3).value
            time_val = ws.cell(row=1, column=6).value
            
        elif ext == ".xls":
            book = xlrd.open_workbook(path)
            sheet = book.sheet_by_index(0)
            if sheet.nrows < 1:
                return ""
            
            if sheet.ncols > 2:
                date_val = sheet.cell_value(0, 2)
            if sheet.ncols > 5:
                time_val = sheet.cell_value(0, 5)
        else:
            return ""
        
        print(f"[DEBUG] C1 (date): {repr(date_val)} | Type: {type(date_val)}")
        print(f"[DEBUG] F1 (time): {repr(time_val)} | Type: {type(time_val)}")
        
        date_str = ""
        time_str = ""
        
        if date_val is not None:
            if isinstance(date_val, datetime):
                date_str = date_val.strftime("%d/%m/%Y")
            elif isinstance(date_val, (int, float)):
                try:
                    from datetime import timedelta
                    excel_epoch = datetime(1899, 12, 30)
                    dt = excel_epoch + timedelta(days=int(date_val))
                    date_str = dt.strftime("%d/%m/%Y")
                    print(f"[DEBUG] Converted date from {date_val}: {date_str}")
                except Exception as e:
                    print(f"[DEBUG] Date conversion error: {e}")
            else:
                val_str = str(date_val).strip()
                patterns = [
                    r'(\d{2}/\d{2}/\d{4})',
                    r'(\d{2}-\d{2}-\d{4})',
                    r'(\d{4}/\d{2}/\d{2})',
                    r'(\d{4}-\d{2}-\d{2})',
                ]
                for pattern in patterns:
                    match = re.search(pattern, val_str)
                    if match:
                        date_str = match.group(1).replace('-', '/')
                        break
        
        if time_val is not None:
            if isinstance(time_val, datetime):
                time_str = time_val.strftime("%H:%M:%S")
            elif isinstance(time_val, (int, float)):
                try:
                    total_seconds = int(time_val * 24 * 60 * 60)
                    hours = total_seconds // 3600
                    minutes = (total_seconds % 3600) // 60
                    seconds = total_seconds % 60
                    time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    print(f"[DEBUG] Converted time from {time_val}: {time_str}")
                except Exception as e:
                    print(f"[DEBUG] Time conversion error: {e}")
            else:
                val_str = str(time_val).strip()
                patterns = [
                    r'(\d{2}:\d{2}:\d{2})',
                    r'(\d{1,2}:\d{2}:\d{2})',
                    r'(\d{2}:\d{2})',
                ]
                for pattern in patterns:
                    match = re.search(pattern, val_str)
                    if match:
                        time_str = match.group(1)
                        if len(time_str.split(':')) == 2:
                            time_str += ":00"
                        break
        
        print(f"[DEBUG] Extracted date: '{date_str}'")
        print(f"[DEBUG] Extracted time: '{time_str}'")
        
        if date_str and time_str:
            try:
                for date_format in ["%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y", "%Y-%m-%d"]:
                    try:
                        dt = datetime.strptime(f"{date_str} {time_str}", f"{date_format} %H:%M:%S")
                        result = dt.strftime("%d-%m-%Y %I:%M %p")
                        print(f"[DEBUG] Final formatted: '{result}'")
                        return result
                    except ValueError:
                        continue
                return f"{date_str} {time_str}"
            except Exception as e:
                print(f"[DEBUG] Format error: {e}")
                return f"{date_str} {time_str}"
        elif date_str:
            return date_str
        elif time_str:
            return time_str
        
        print("[DEBUG] No date/time found")
        return ""
        
    except Exception as e:
        print(f"[ERROR] Reading stock time: {e}")
        import traceback
        traceback.print_exc()
        return ""

def write_pdf_report(out_file: str, cam: str, rows: list, stock_time: str = "", depot: str = "") -> None:
    """
    Create PDF with professional layout in LANDSCAPE mode:
    - Logo at top left
    - Site and stock time info at top right
    - Multiple columns with smaller font
    """
    logo_path_absolute = os.path.abspath(LOGO_PATH)
    
    doc = SimpleDocTemplate(
        out_file, 
        pagesize=landscape(A4),
        topMargin=0.8*cm, 
        bottomMargin=0.8*cm,
        leftMargin=0.8*cm,
        rightMargin=0.8*cm
    )
    story = []
    styles = getSampleStyleSheet()
    
    print(f"[DEBUG] LOGO_PATH: {LOGO_PATH}")
    print(f"[DEBUG] Absolute LOGO_PATH: {logo_path_absolute}")
    print(f"[DEBUG] Logo exists: {os.path.exists(logo_path_absolute)}")
    if os.path.exists(logo_path_absolute):
        print(f"[DEBUG] Logo file size: {os.path.getsize(logo_path_absolute)} bytes")
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#dc2626'),
        spaceAfter=5,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#ea580c'),
        alignment=TA_CENTER,
        spaceAfter=10,
        fontName='Helvetica-Bold'
    )
    
    print(f"PDF Generation - Stock time for {cam}: '{stock_time}'")
    
    title_text = f"<b>{cam}</b>"
    
    if stock_time and stock_time.strip():
        subtitle_text = f"<b>Date de mise à jour du stock:</b> {stock_time}"
    else:
        subtitle_text = f"<b>Date de mise à jour du stock:</b> Non disponible"
    
    logo = None
    if os.path.exists(logo_path_absolute):
        try:
            logo = Image(logo_path_absolute, width=3*cm, height=1.5*cm)
            print(f"[DEBUG] Logo Image object created successfully")
        except Exception as e:
            print(f"[ERROR] Error loading logo: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"[ERROR] Logo file not found at: {logo_path_absolute}")
    
    title_para = Paragraph(title_text, title_style)
    subtitle_para = Paragraph(subtitle_text, subtitle_style)
    
    if logo:
        try:
            header_table_data = [[logo, [title_para, subtitle_para]]]
            header_table = Table(header_table_data, colWidths=[3.5*cm, 23.5*cm])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (0, 0), 'TOP'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('LEFTPADDING', (0, 0), (0, 0), 0),
                ('RIGHTPADDING', (0, 0), (0, 0), 10),
            ]))
            story.append(header_table)
            print(f"[DEBUG] Header table with logo added to PDF")
        except Exception as e:
            print(f"[ERROR] Error creating header table with logo: {e}")
            import traceback
            traceback.print_exc()
            story.append(title_para)
            story.append(subtitle_para)
    else:
        print(f"[DEBUG] No logo, using text-only header")
        story.append(title_para)
        story.append(subtitle_para)
    
    story.append(Spacer(1, 0.4*cm))
    
    if not rows:
        story.append(Paragraph("Aucun produit manquant", styles['Normal']))
        doc.build(story)
        return
    
    num_products = len(rows)
    num_data_columns = (num_products + PRODUCTS_PER_COLUMN - 1) // PRODUCTS_PER_COLUMN
    num_data_columns = min(num_data_columns, 4)
    
    header_row = []
    for i in range(num_data_columns):
        header_row.extend([depot, 'Article'])
        if i < num_data_columns - 1:
            header_row.append('')
    
    table_data = [header_row]
    
    for row_idx in range(PRODUCTS_PER_COLUMN):
        data_row = []
        for col_idx in range(num_data_columns):
            product_idx = col_idx * PRODUCTS_PER_COLUMN + row_idx
            if product_idx < len(rows):
                sfx_qty, designation = rows[product_idx]
                
                if sfx_qty != "":
                    try:
                        qty_float = float(sfx_qty)
                        if qty_float == int(qty_float):
                            sfx_qty_str = str(int(qty_float))
                        else:
                            sfx_qty_str = str(qty_float)
                    except (ValueError, TypeError):
                        sfx_qty_str = str(sfx_qty)
                else:
                    sfx_qty_str = ""
                
                if len(designation) > 35:
                    designation = designation[:32] + "..."
                data_row.extend([sfx_qty_str, designation])
            else:
                data_row.extend(["", ""])
            
            if col_idx < num_data_columns - 1:
                data_row.append('')
        
        table_data.append(data_row)
    
    available_width = 25 * cm
    separator_width = 0.2 * cm
    total_separator_width = (num_data_columns - 1) * separator_width
    data_width = available_width - total_separator_width
    
    col_set_width = data_width / num_data_columns
    sfx_col_width = col_set_width * 0.2
    article_col_width = col_set_width * 0.8
    
    col_widths = []
    for i in range(num_data_columns):
        col_widths.extend([sfx_col_width, article_col_width])
        if i < num_data_columns - 1:
            col_widths.append(separator_width)
    
    table = Table(table_data, colWidths=col_widths)
    
    table_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dc2626')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('TOPPADDING', (0, 0), (-1, 0), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 3),
        ('TOPPADDING', (0, 1), (-1, -1), 3),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#fbbf24')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#fef3c7')]),
    ]
    
    for i in range(num_data_columns - 1):
        sep_col = 2 + i * 3
        table_style.append(('BACKGROUND', (sep_col, 0), (sep_col, -1), colors.white))
        table_style.append(('GRID', (sep_col, 0), (sep_col, -1), 0, colors.white))
    
    table.setStyle(TableStyle(table_style))
    story.append(table)
    
    try:
        doc.build(story)
        print(f"[DEBUG] PDF created successfully: {out_file}")
    except Exception as e:
        print(f"[ERROR] Error building PDF: {e}")
        import traceback
        traceback.print_exc()

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(
    request=request,
    name="index.html",
    context={"depots": list(DEPOT_CAMS.keys())}
)


@app.post("/upload/")
async def upload(file: UploadFile = File(...), depot: str = Form("SFX")):
    filename = file.filename or ""
    ext = os.path.splitext(filename.lower())[1]
    temp_path = os.path.join(BASE_DIR, f"temp_{filename}")

    with open(temp_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        # Validate depot
        if depot not in DEPOT_CAMS:
            return JSONResponse(
                status_code=400,
                content={"error": f"Dépôt invalide: {depot}"}
            )
        
        cam_sites = DEPOT_CAMS[depot]
        print(f"[DEBUG] Processing depot: {depot} with CAMs: {cam_sites}")
        
        try:
            if ext == ".xlsx":
                headers, rows_iter = _extract_rows_xlsx(temp_path)
            elif ext == ".xls":
                headers, rows_iter = _extract_rows_xls(temp_path)
            else:
                return JSONResponse(
                    status_code=400,
                    content={"error": "Type de fichier non supporté. Téléchargez un fichier .xls ou .xlsx"}
                )
        except (InvalidFileException, zipfile.BadZipFile, xlrd.biffh.XLRDError) as e:
            return JSONResponse(status_code=400, content={"error": "Fichier Excel invalide.", "details": str(e)})

        stock_time = _read_stock_time(temp_path, ext)

        required = ["Site", "Code", "Désignation Article"]
        missing_cols = [c for c in required if c not in headers]
        if missing_cols:
            return JSONResponse(
                status_code=400,
                content={
                    "error": "Colonnes requises manquantes dans la ligne d'en-tête (ligne 2).",
                    "missing": missing_cols,
                    "found": list(headers.keys())
                }
            )

        site_col = headers["Site"]
        code_col = headers["Code"]
        des_col = headers["Désignation Article"]
        stock_col = headers.get("Stock Disponible")

        def get_cell(row, col_1based):
            i = col_1based - 1
            return row[i] if 0 <= i < len(row) else None

        depot_products = {}  # Products at the depot (e.g., SFX)
        cam_products_map = {cam: set() for cam in cam_sites}

        for row in rows_iter:
            site = _norm(get_cell(row, site_col))
            if not site:
                continue

            code = _norm(get_cell(row, code_col))
            des = _norm(get_cell(row, des_col))
            if not code or not des:
                continue

            # Check if this is the depot we're processing
            if site == depot:
                if stock_col is not None:
                    stock_val = _to_float(get_cell(row, stock_col))
                    if stock_val is None or stock_val < 4:
                        continue
                    depot_qty = stock_val
                else:
                    depot_qty = ""
                depot_products[(code, des)] = depot_qty
            elif site in cam_products_map:
                cam_products_map[site].add((code, des))

        results = []
        depot_set = set(depot_products.keys())

        for cam in cam_sites:
            missing = depot_set - cam_products_map[cam]
            missing = {p for p in missing if p[0] not in IGNORED_CODES_SET}
            missing = {p for p in missing if not p[1].startswith("PALETTES")}

            rows = sorted([(depot_products[p], p[1]) for p in missing], key=lambda item: item[1])

            out_file = os.path.join(OUTPUT_DIR, f"{depot}_{cam}.pdf")
            write_pdf_report(out_file, cam, rows, stock_time, depot)

            results.append({"cam": cam, "count": len(rows), "file": f"/download/{depot}_{cam}.pdf"})

        zip_path = os.path.join(OUTPUT_DIR, f"{depot}_tous_les_sites.zip")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for cam in cam_sites:
                z.write(os.path.join(OUTPUT_DIR, f"{depot}_{cam}.pdf"), f"{depot}_{cam}.pdf")

        return {"results": results, "zip": f"/download/{depot}_all", "depot": depot}

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": "Erreur serveur", "details": str(e)})

    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


@app.get("/download/{file}")
def download(file: str):
    if "_all" in file:
        depot = file.replace("_all", "")
        zip_path = os.path.join(OUTPUT_DIR, f"{depot}_tous_les_sites.zip")
        if not os.path.isfile(zip_path):
            return JSONResponse(status_code=404, content={"error": "Aucun fichier ZIP traité. Téléchargez d'abord un fichier."})
        return FileResponse(zip_path, filename=f"{depot}_tous_les_sites.zip")

    filename = file if file.endswith(".pdf") else f"{file}.pdf"
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.isfile(path):
        return JSONResponse(status_code=404, content={"error": "Fichier introuvable. Téléchargez d'abord un fichier."})
    return FileResponse(path, filename=filename)