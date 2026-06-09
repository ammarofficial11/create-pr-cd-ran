from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from pydantic import BaseModel
from fastapi.responses import HTMLResponse

import pandas as pd
import sys
import os
import shutil
import threading
import json
from datetime import datetime

from src.run_pipeline import run_pipeline

LAST_BOM_FILE = None
LAST_EPMS_FILE = None
# =====================================================
# GLOBALS
# =====================================================

JOB_STATUS = "idle"


# =====================================================
# REQUEST MODEL
# =====================================================

class GenerateRequest(BaseModel):
    project: str


# =====================================================
# APP
# =====================================================

app = FastAPI(
    title="RAN PR Automation API",
    version="2.0"
)


# =====================================================
# HEALTH CHECK
# =====================================================

@app.get("/")
def health():

    return {
        "status": "healthy",
        "service": "RAN PR Automation"
    }


# =====================================================
# PROJECT LIST
# =====================================================

@app.get("/projects")
def get_projects():

    excel_file = "config/GENERAL ITEM FOR ALL DU PROJECT Overall.xlsx"

    if not os.path.exists(excel_file):
        return []

    all_sheets = pd.read_excel(
        excel_file,
        sheet_name=None
    )

    projects = set()

    for sheet_name, df in all_sheets.items():

        columns = list(df.columns)

        # Column E onwards
        for col in columns[4:]:

            project = str(col).strip()

            if (
                project
                and project.lower() != "nan"
                and not project.startswith("Unnamed")
            ):
                projects.add(project)

    return sorted(list(projects))


# =====================================================
# TEST ENDPOINT
# =====================================================

@app.get("/hello/{name}")
def hello(name: str):

    return {
        "message": f"Hello {name}"
    }


# =====================================================
# VERSION
# =====================================================

@app.get("/version")
def version():

    return {
        "system": "RAN PR Automation",
        "status": "connected"
    }


# =====================================================
# PYTHON INFO
# =====================================================

@app.get("/python-info")
def python_info():

    return {
        "python_version": sys.version
    }


# =====================================================
# UPLOAD BOM
# =====================================================

@app.post("/upload-bom")
async def upload_bom(file: UploadFile = File(...)):
    global LAST_BOM_FILE

    LAST_BOM_FILE = file.filename

    os.makedirs("input", exist_ok=True)

    destination = "input/BOM.xlsx"

    with open(destination, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {
        "status": "success",
        "saved_as": destination,
        "original_filename": file.filename,
        "size_bytes": os.path.getsize(destination)
    }


# =====================================================
# UPLOAD EPMS
# =====================================================

@app.post("/upload-epms")
async def upload_epms(file: UploadFile = File(...)):
    global LAST_EPMS_FILE

    LAST_EPMS_FILE = file.filename
    os.makedirs("input", exist_ok=True)

    destination = "input/EPMS.xlsx"

    with open(destination, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {
        "status": "success",
        "saved_as": destination,
        "original_filename": file.filename,
        "size_bytes": os.path.getsize(destination)
    }


# =====================================================
# FILE CHECK
# =====================================================

@app.get("/check-files")
def check_files():

    return {
        "bom_exists": os.path.exists("input/BOM.xlsx"),
        "epms_exists": os.path.exists("input/EPMS.xlsx"),
        "bom_size": os.path.getsize("input/BOM.xlsx")
            if os.path.exists("input/BOM.xlsx") else 0,
        "epms_size": os.path.getsize("input/EPMS.xlsx")
            if os.path.exists("input/EPMS.xlsx") else 0
    }


# =====================================================
# BACKGROUND JOB
# =====================================================

def run_job(project):

    global JOB_STATUS
    global LAST_BOM_FILE
    global LAST_EPMS_FILE

    JOB_STATUS = "running"

    run_pipeline(
        selected_project=project
    )

    os.makedirs(
        "output",
        exist_ok=True
    )

    with open(
        "output/job_info.json",
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            {
                "project": project,
                "generated_time":
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                "bom_file":
                    LAST_BOM_FILE,
                "epms_file":
                    LAST_EPMS_FILE
            },
            f,
            indent=4
        )

    JOB_STATUS = "completed"

# =====================================================
# GENERATE PR
# =====================================================

@app.post("/generate-pr")
def generate_pr(request: GenerateRequest):

    thread = threading.Thread(
        target=run_job,
        args=(request.project,)
    )

    thread.start()

    return {
        "status": "started",
        "project": request.project
    }

@app.get("/job-info")
def job_info():

    file_path = "output/job_info.json"

    if not os.path.exists(file_path):

        return {
            "status": "no_job"
        }

    with open(
        file_path,
        "r",
        encoding="utf-8"
    ) as f:

        return json.load(f)

# =====================================================
# JOB STATUS
# =====================================================

@app.get("/job-status")
def job_status():

    return {
        "status": JOB_STATUS
    }


# =====================================================
# SYSTEM STATUS
# =====================================================

@app.get("/status")
def status():

    return {
        "bom_uploaded":
            os.path.exists("input/BOM.xlsx"),

        "epms_uploaded":
            os.path.exists("input/EPMS.xlsx"),

        "ecc_exists":
            os.path.exists(
                "output/ECC_PR_Output.xlsx"
            ),

        "ecc_general_exists":
            os.path.exists(
                "output/ECC_PR_Output_With_GeneralItems.xlsx"
            )
    }


# =====================================================
# DOWNLOAD STANDARD OUTPUT
# =====================================================

@app.get("/download-pr")
def download_pr():

    file_path = "output/ECC_PR_Output.xlsx"

    if not os.path.exists(file_path):

        return {
            "status": "error",
            "message": "Output file not found"
        }

    return FileResponse(
        path=file_path,
        filename="ECC_PR_Output.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =====================================================
# DOWNLOAD GENERAL ITEM OUTPUT
# =====================================================

@app.get("/download-pr-general")
def download_pr_general():

    file_path = "output/ECC_PR_Output_With_GeneralItems.xlsx"

    if not os.path.exists(file_path):

        return {
            "status": "error",
            "message": "Output file not found"
        }

    return FileResponse(
        path=file_path,
        filename="ECC_PR_Output_With_GeneralItems.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.get("/ui", response_class=HTMLResponse)
def ui():

    with open("web/index.html") as f:
        return f.read()

@app.get("/output-info")
def output_info():

    import datetime

    result = {}

    files = {
        "bom_ti":
        "output/ECC_PR_Output.xlsx",

        "general":
        "output/ECC_PR_Output_With_GeneralItems.xlsx"
    }

    for key, path in files.items():

        if os.path.exists(path):

            modified_time = os.path.getmtime(path)

            result[key] = datetime.datetime.fromtimestamp(
                modified_time
            ).strftime("%Y-%m-%d %H:%M:%S")

        else:

            result[key] = None

    return result