from fastapi import FastAPI, Request, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse,FileResponse, JSONResponse,RedirectResponse,StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
from typing import List
from io import BytesIO
import os, re, csv, uuid,shutil,tempfile
import pandas as pd
from ezstdf.StdfReader import StdfReader

app = FastAPI()
app.add_middleware(#CORS (Cross-Origin Resource Sharing):
    CORSMiddleware,
    allow_origins=["*"],# ✅ Allow all origins (only for dev)
    allow_credentials=True,
    allow_methods=["*"],# ✅ Allow all methods ex. GET, POST, PUT
    allow_headers=["*"],# ✅ Allow all headers ex. Content-Type, Authorization
)
templates = Jinja2Templates(directory="templates")
UPLOAD_DIR = "output"
os.makedirs(UPLOAD_DIR,exist_ok=True)

TEMP_DIR = "temp"
os.makedirs(TEMP_DIR, exist_ok=True)



@app.get("/", include_in_schema=False)
async def root():
    try:
        return RedirectResponse(url="/docs")
    except Exception as e:
        return {"error": str(e)}

#app.post use for export data


@app.post("/upload-stdf/")
async def upload_stdf(files: List[UploadFile] = File(...)):
    results = []

    for file in files:
        contents = await file.read()
        safe_filename = file.filename.replace("\\", "/").split("/")[-1]

        # สร้างไฟล์ชั่วคราวเพื่อให้ StdfReader อ่านได้
        with tempfile.NamedTemporaryFile(delete=False, suffix=".stdf") as temp_file:
            temp_file.write(contents)
            temp_file_path = temp_file.name

        try:
            stdf = StdfReader()
            stdf.parse_file(temp_file_path)

            # แปลงเป็น Excel ลงใน memory
            output_stream = BytesIO()
            stdf.to_excel(output_stream)
            output_stream.seek(0)

            # สร้างชื่อไฟล์ใหม่
            safe_filename = file.filename.replace("\\", "/").split("/")[-1] 
            output_filename = os.path.splitext(safe_filename)[0] + ".xlsx"


            # ส่งกลับเป็นไฟล์ให้ดาวน์โหลด
            return StreamingResponse(
                output_stream,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f"attachment; filename={output_filename}"}
            )
        finally:
            os.remove(temp_file_path)  # ลบไฟล์ชั่วคราว

    return {"detail": "No files processed"}



@app.post("/process-self-converted-datalog/")
async def process_selfconverted_datalog_excel(file: UploadFile = File(...)):
    contents = await file.read()
    xls = pd.ExcelFile(BytesIO(contents), engine="openpyxl")
    if "mpr" not in xls.sheet_names:
        return {"error": "Sheet 'mpr' not found in the uploaded file."}
    df = xls.parse("mpr")
    # ค้นหา column TEST_NUM
    test_num_col = next((col for col in df.columns if col.strip().upper() == "TEST_NUM"), None)
    if not test_num_col:
        return {"error": "Column 'TEST_NUM' not found in sheet 'mpr'."}
    # ค้นหา column TEST_TXT
    test_txt_col = next((col for col in df.columns if col.strip().upper() == "TEST_TXT"), None)
    if not test_txt_col:
        return {"error": "Column 'TEST_TXT' not found in sheet 'mpr'."}
    #ค้นหา column RTN_RSLT
    Rslt_col = next((col for col in df.columns if col.strip().upper() == "RTN_RSLT"), None)
    if not Rslt_col:
        return {"error": "Column 'TEST_RSLT' not found in sheet 'mpr'."}
    # เตรียมข้อมูล suite/test name test stage check พร้อมเช็ค Test number
    test_data = []
    seen = set()
    for _, row in df.iterrows():
        test_num = row.get(test_num_col)
        test_txt = row.get(test_txt_col)
        Rslt_data = row.get(Rslt_col)
        if pd.isna(test_num) or pd.isna(test_txt) or pd.isna(Rslt_data):
            continue
        # แยกข้อความด้วย :
        if ":" in test_txt:
            suite, test = map(str.strip, test_txt.split(":", 1))
            test = test.split("@", 1)[0].strip()

        else:
            suite, test = "", test_txt.strip()
        if str(Rslt_data) == "":
            YN_check = "N"
        else: YN_check = "Y"
        # กรองเฉพาะ test_num และ test_name
        key = (test_num, test)
        if key not in seen:
            seen.add(key)
            test_data.append({
                "test_number": test_num,
                "suite_name": suite,
                "test_name": test,
                "YN_check" : YN_check
            })
    return {
        "test_data": test_data
    }





@app.post("/upload-folder/")
async def upload_folder(files: List[UploadFile] = File(...)):
    uploaded_mfh = []
    base_dir = "/tmp/uploads" 
    os.makedirs(base_dir, exist_ok=True)

    for file in files:
        relative_path = file.filename.replace("\\", "/")
        full_path = os.path.join(base_dir, relative_path)
        os.makedirs(os.path.dirname(full_path), exist_ok=True)

        with open(full_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        if relative_path.endswith(".mfh"):
            uploaded_mfh.append(relative_path)

    return {"mfh_files": uploaded_mfh}




@app.get("/process-testtable/")
async def process_mfh_file(filename: str):
    base_dir = "/tmp/uploads"
    mfh_path = os.path.join(base_dir, filename)

    if not os.path.exists(mfh_path):
        raise HTTPException(status_code=404, detail="MFH file not found")

    uploaded_files = {}
    for root, _, files in os.walk(base_dir):
        for f in files:
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, base_dir).replace("\\", "/")
            uploaded_files[rel_path.lower()] = full_path

    results = []
    try:
        with open(mfh_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        for path in lines:
            path = path.strip().replace("\\", "/")
            if not path or "," in path:
                continue
            if path.lower().startswith("testerfile "):
                path = path[len("testerfile "):].strip()

            matched_path = uploaded_files.get(path.lower())

            if not matched_path:
                for key, full_path in uploaded_files.items():
                    if key.endswith(path.lower()):
                        matched_path = full_path
                        break

            if matched_path and matched_path.endswith(".csv"):
                try:
                    data, file_type = try_read_as_excel_then_csv(matched_path, encoding="utf-8")
                    results.append({
                        "path": path,
                        "status": "ok",
                        "type": file_type,
                        "content": data
                    })
                except Exception as e:
                    results.append({
                        "path": path,
                        "status": "error",
                        "error": str(e)
                    })
            else:
                results.append({
                    "path": path,
                    "status": "not found"
                })
    finally:
        try:
            os.remove(mfh_path)
        except:
            pass
    return JSONResponse(content={"files": results})

@app.post("/process-EY/")
async def process_EY_file(file: UploadFile = File(...)):
    ext = os.path.splitext(file.filename)[1].lower()

    if ext not in ['.xls', '.xlsx']:
        return {"error": "Only Excel files are supported."}

    # อ่านชีทเดียวจาก stream
    df = pd.read_excel(file.file, sheet_name=0, engine='openpyxl')
    df.columns = [str(col).strip() for col in df.columns]

    col_map = {}
    for col in df.columns:
        col_upper = col.upper()
        if 'PARAMETER NUMBER' in col_upper:
            col_map['test_number'] = col
        elif 'PARAMETER NAME' in col_upper:
            col_map['test_name'] = col
        elif 'COUNT' in col_upper:
            col_map['count'] = col
        elif 'PRODUCT' in col_upper:
            col_map['product'] = col
        elif 'STAGE' in col_upper:
            col_map['stage'] = col
            

    required_keys = ['test_number', 'test_name', 'count', 'product', 'stage']
    missing_keys = [k for k in required_keys if k not in col_map]

    if missing_keys:
        return {"error": f"Missing required columns: {', '.join(missing_keys)}"}
    
    if not all(k in col_map for k in required_keys):
        return {"error": "Missing required columns."}

    all_data = []
    for _, row in df.iterrows():
        try:
            test_num = int(row.get(col_map['test_number']))
            if test_num == 0:
                continue
        except:
            continue

        test_txt = row.get(col_map['test_name'])
        if pd.isna(test_txt):
            continue

        test_txt = str(test_txt)
        if ":" in test_txt:
            suite, test = map(str.strip, test_txt.split(":", 1))
            test = test.split("@", 1)[0].strip()
        else:
            suite, test = "", test_txt.strip()

        try:
            YN_check = "Y" if float(row.get(col_map['count'])) > 0 else "N"
        except:
            YN_check = "N"

        all_data.append({
            "test_number": test_num,
            "suite_name": suite,
            "test_name": test,
            "YN_check": YN_check,
            "product": row.get(col_map['product']),
            "stage": row.get(col_map['stage'])
        })

    # Group by stage
    grouped_data = {}
    for entry in all_data:
        stage = entry['stage']
        grouped_data.setdefault(stage, []).append(entry)

    # Flatten
    final_result = []
    for stage_entries in grouped_data.values():
        final_result.extend(stage_entries)

    return {"data": final_result}
 
