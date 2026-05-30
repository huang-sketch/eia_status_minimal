import json
import os
import shutil
import subprocess
import sys
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
from uuid import uuid4

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles

from noise_table_preprocessor import export_noise_table_debug
from word_processor import load_docx_chunks


BASE_DIR = Path(__file__).resolve().parent
WEB_JOBS_DIR = BASE_DIR / "runs" / "web_jobs"
STATIC_DIR = BASE_DIR / "static"

REPORT_FILENAME = "监测报告.docx"
PLAN_FILENAME = "监测方案.docx"

NOISE_MONITORING_COLUMNS = [
    "监测点编号",
    "监测点名称",
    "监测点位置",
    "day1_day",
    "day1_night",
    "day2_day",
    "day2_night",
    "avg_day",
    "avg_night",
    "traffic_flow_day1_day",
    "traffic_flow_day1_night",
    "traffic_flow_day2_day",
    "traffic_flow_day2_night",
]

NOISE_COMPLIANCE_COLUMNS = [
    "监测点编号",
    "监测点名称",
    "监测点位置",
    "avg_day",
    "avg_night",
    "standard_day",
    "standard_night",
    "exceed_day",
    "exceed_night",
    "needs_review",
    "warning",
]

RESULT_GROUP_DEFINITIONS = {
    "monitoring_extraction": {
        "title": "监测数据提取",
        "files": [
            "monitoring_records.json",
            "standard_config.json",
            "debug_tables/detected_factors.json",
            "debug_tables/noise_monitor_points_table.json",
            "debug_tables/surface_water_monitor_points_table.json",
            "debug_tables/surface_water_monitor_results_table.json",
        ],
    },
    "compliance": {
        "title": "达标判定",
        "files": [
            "compliance_results.json",
            "debug_tables/surface_water_compliance_table.json",
            "debug_tables/noise_compliance_summary.json",
            "debug_tables/noise_sensitive_points_result_table.json",
            "debug_tables/traffic_noise_attenuation_table.json",
        ],
    },
    "text_output": {
        "title": "文本输出",
        "files": [
            "surface_water_section.docx",
            "noise_section.docx",
            "eia_outputs.zip",
        ],
    },
}

app = FastAPI(title="环评现状分析自动化系统")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.post("/api/jobs")
async def create_job(
    background_tasks: BackgroundTasks,
    report_name: str = Form(...),
    admin_division: str = Form(...),
    run_surface_water: bool = Form(True),
    run_noise: bool = Form(True),
    enable_llm_text_polish: bool = Form(False),
    monitoring_report: UploadFile = File(...),
    monitoring_plan: UploadFile = File(...),
) -> JSONResponse:
    if not run_surface_water and not run_noise:
        raise HTTPException(status_code=400, detail="请至少选择一个生成内容")
    validate_docx_upload(monitoring_report, "监测报告")
    validate_docx_upload(monitoring_plan, "监测方案")

    job_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}"
    job_dir = WEB_JOBS_DIR / job_id
    input_dir = job_dir / "input"
    output_dir = job_dir / "output"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    await save_upload(monitoring_report, input_dir / REPORT_FILENAME)
    await save_upload(monitoring_plan, input_dir / PLAN_FILENAME)
    write_json(
        input_dir / "project_meta.json",
        {
            "job_id": job_id,
            "report_name": report_name,
            "admin_division": admin_division,
            "run_surface_water": run_surface_water,
            "run_noise": run_noise,
            "enable_llm_text_polish": enable_llm_text_polish,
            "created_at": datetime.now().isoformat(timespec="seconds"),
        },
    )

    update_status(
        job_dir,
        {
            "job_id": job_id,
            "status": "pending",
            "current_step": "等待执行",
            "error": None,
            "result_groups": {},
        },
    )
    background_tasks.add_task(
        run_job,
        job_id,
        run_surface_water,
        run_noise,
        enable_llm_text_polish,
    )
    return JSONResponse({"job_id": job_id, "status": "pending"})


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str) -> JSONResponse:
    job_dir = resolve_job_dir(job_id)
    status = read_status(job_dir)
    status["result_groups"] = build_result_groups(job_id)
    return JSONResponse(status)


@app.get("/api/jobs/{job_id}/preview")
def get_job_preview(job_id: str) -> JSONResponse:
    job_dir = resolve_job_dir(job_id)
    output_dir = job_dir / "output"
    return JSONResponse(build_preview_payload(output_dir))


@app.get("/api/jobs/{job_id}/logs")
def get_logs(job_id: str) -> PlainTextResponse:
    job_dir = resolve_job_dir(job_id)
    log_path = job_dir / "job.log"
    if not log_path.exists():
        return PlainTextResponse("")
    return PlainTextResponse(log_path.read_text(encoding="utf-8", errors="replace"))


@app.get("/api/jobs/{job_id}/download/{file_path:path}")
def download_file(job_id: str, file_path: str) -> FileResponse:
    job_dir = resolve_job_dir(job_id)
    output_dir = job_dir / "output"
    target = (output_dir / file_path).resolve()
    if not is_relative_to(target, output_dir.resolve()) or not target.is_file():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(target, filename=target.name)


def run_job(
    job_id: str,
    run_surface_water: bool,
    run_noise: bool,
    enable_llm_text_polish: bool,
) -> None:
    job_dir = WEB_JOBS_DIR / job_id
    input_dir = job_dir / "input"
    output_dir = job_dir / "output"
    log_path = job_dir / "job.log"

    try:
        update_status(job_dir, {"status": "running", "current_step": "初始化任务", "error": None})
        append_log(log_path, f"job_id: {job_id}")
        append_log(log_path, f"input_dir: {input_dir}")
        append_log(log_path, f"output_dir: {output_dir}")

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["EIA_INPUT_DIR"] = str(input_dir)
        env["EIA_OUTPUT_DIR"] = str(output_dir)
        env["ENABLE_LLM_TEXT_POLISH"] = "true" if enable_llm_text_polish else "false"

        if run_surface_water:
            update_status(job_dir, {"current_step": "地表水：解析监测方案和监测报告，执行达标判定"})
            run_script("surface_water_pipeline.py", env, log_path)
            update_status(job_dir, {"current_step": "地表水：生成 Word 章节"})
            run_script("surface_water_section_generator.py", env, log_path)

        if run_noise:
            update_status(job_dir, {"current_step": "声环境：预处理噪声表格"})
            prepare_noise_debug_tables(input_dir, output_dir, log_path)
            update_status(job_dir, {"current_step": "声环境：生成 Word 章节"})
            run_script("noise_section_generator.py", env, log_path)

        update_status(job_dir, {"current_step": "打包输出文件"})
        create_output_zip(output_dir)
        update_status(
            job_dir,
            {
                "status": "success",
                "current_step": "完成",
                "error": None,
                "result_groups": build_result_groups(job_id),
            },
        )
        append_log(log_path, "job finished successfully")
    except Exception as exc:
        append_log(log_path, "job failed")
        append_log(log_path, traceback.format_exc())
        update_status(
            job_dir,
            {
                "status": "failed",
                "current_step": "失败",
                "error": str(exc),
                "result_groups": build_result_groups(job_id),
            },
        )


def run_script(script_name: str, env: Dict[str, str], log_path: Path) -> None:
    append_log(log_path, f"> python {script_name}")
    completed = subprocess.run(
        [sys.executable, script_name],
        cwd=BASE_DIR,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if completed.stdout:
        append_log(log_path, completed.stdout.rstrip())
    if completed.stderr:
        append_log(log_path, completed.stderr.rstrip())
    if completed.returncode != 0:
        raise RuntimeError(f"{script_name} 执行失败，退出码 {completed.returncode}")


def prepare_noise_debug_tables(input_dir: Path, output_dir: Path, log_path: Path) -> None:
    debug_dir = output_dir / "debug_tables"
    debug_dir.mkdir(parents=True, exist_ok=True)
    report_path = input_dir / REPORT_FILENAME
    chunks = load_docx_chunks(report_path)
    table_index = 0
    for chunk in chunks:
        exported = export_noise_table_debug(chunk, table_index, debug_dir)
        if exported:
            append_log(
                log_path,
                f"noise table exported: {exported['flattened'].relative_to(output_dir)}",
            )
            table_index += 1
    if table_index == 0:
        raise RuntimeError("未识别到噪声监测表，无法生成声环境章节")


def create_output_zip(output_dir: Path) -> None:
    zip_path = output_dir / "eia_outputs.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in output_dir.rglob("*"):
            if path.is_file() and path != zip_path:
                archive.write(path, path.relative_to(output_dir))


def build_result_groups(job_id: str) -> Dict[str, Dict[str, Any]]:
    job_dir = resolve_job_dir(job_id)
    output_dir = job_dir / "output"
    groups: Dict[str, Dict[str, Any]] = {}
    for group_key, definition in RESULT_GROUP_DEFINITIONS.items():
        files: List[Dict[str, str]] = []
        for relative in definition["files"]:
            path = output_dir / relative
            if path.exists():
                files.append({"name": path.name, "path": relative})
        groups[group_key] = {"title": definition["title"], "files": files}
    return groups


def build_preview_payload(output_dir: Path) -> Dict[str, Any]:
    return {
        "monitoring_extraction": {
            "surface_water_results": read_table_preview(
                output_dir,
                "debug_tables/surface_water_monitor_results_table.json",
            ),
            "noise_sensitive_results": project_table_columns(
                read_table_preview(
                    output_dir,
                    "debug_tables/noise_sensitive_points_result_table.json",
                ),
                NOISE_MONITORING_COLUMNS,
            ),
            "noise_attenuation_results": project_table_columns(
                read_table_preview(
                    output_dir,
                    "debug_tables/traffic_noise_attenuation_table.json",
                ),
                NOISE_MONITORING_COLUMNS,
            ),
        },
        "compliance": {
            "surface_water_compliance": read_table_preview(
                output_dir,
                "debug_tables/surface_water_compliance_table.json",
            ),
            "noise_sensitive": project_table_columns(
                read_table_preview(
                    output_dir,
                    "debug_tables/noise_sensitive_points_result_table.json",
                ),
                NOISE_COMPLIANCE_COLUMNS,
            ),
            "noise_attenuation": project_table_columns(
                read_table_preview(
                    output_dir,
                    "debug_tables/traffic_noise_attenuation_table.json",
                ),
                NOISE_COMPLIANCE_COLUMNS,
            ),
            "noise_summary": read_optional_json(
                output_dir / "debug_tables/noise_compliance_summary.json"
            ),
        },
    }


def project_table_columns(table: Dict[str, Any], columns: List[str]) -> Dict[str, Any]:
    source_headers = set(table.get("headers", []))
    rows = table.get("rows", [])
    headers = [
        header for header in columns
        if header in source_headers or any(isinstance(row, dict) and header in row for row in rows)
    ]
    projected_rows = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        projected_rows.append({header: row.get(header, "") for header in headers})
    projected = dict(table)
    projected["headers"] = headers
    projected["rows"] = projected_rows
    projected["row_count"] = len(projected_rows)
    return projected


def read_table_preview(output_dir: Path, relative_path: str) -> Dict[str, Any]:
    path = output_dir / relative_path
    if not path.exists():
        return {
            "title": "",
            "headers": [],
            "rows": [],
            "source_path": relative_path,
            "exists": False,
        }

    payload = read_optional_json(path)
    if isinstance(payload, list):
        headers = infer_headers(payload)
        rows = payload
        title = path.stem
    elif isinstance(payload, dict):
        rows = payload.get("rows") or payload.get("records") or []
        headers = payload.get("headers") or payload.get("flattened_headers") or infer_headers(rows)
        title = payload.get("title") or payload.get("table_title") or path.stem
    else:
        headers = []
        rows = []
        title = path.stem

    return {
        "title": title,
        "headers": headers,
        "rows": rows,
        "source_path": relative_path,
        "exists": True,
        "row_count": len(rows) if isinstance(rows, list) else 0,
    }


def infer_headers(rows: Any) -> List[str]:
    if not isinstance(rows, list):
        return []
    headers: List[str] = []
    seen = set()
    for row in rows[:20]:
        if not isinstance(row, dict):
            continue
        for key in row.keys():
            if key in seen or str(key).startswith("_"):
                continue
            seen.add(key)
            headers.append(key)
    return headers


def read_optional_json(path: Path) -> Any:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def validate_docx_upload(upload: UploadFile, label: str) -> None:
    filename = upload.filename or ""
    if not filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail=f"{label}必须是 .docx 文件")


async def save_upload(upload: UploadFile, path: Path) -> None:
    with path.open("wb") as file:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            file.write(chunk)


def resolve_job_dir(job_id: str) -> Path:
    if not job_id or any(part in job_id for part in ("..", "/", "\\")):
        raise HTTPException(status_code=404, detail="任务不存在")
    job_dir = (WEB_JOBS_DIR / job_id).resolve()
    if not is_relative_to(job_dir, WEB_JOBS_DIR.resolve()) or not job_dir.exists():
        raise HTTPException(status_code=404, detail="任务不存在")
    return job_dir


def read_status(job_dir: Path) -> Dict[str, Any]:
    status_path = job_dir / "status.json"
    if not status_path.exists():
        raise HTTPException(status_code=404, detail="任务状态不存在")
    return json.loads(status_path.read_text(encoding="utf-8"))


def update_status(job_dir: Path, patch: Dict[str, Any]) -> None:
    status_path = job_dir / "status.json"
    current: Dict[str, Any] = {}
    if status_path.exists():
        current = json.loads(status_path.read_text(encoding="utf-8"))
    current.update(patch)
    current["updated_at"] = datetime.now().isoformat(timespec="seconds")
    write_json(status_path, current)


def append_log(path: Path, message: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as file:
        file.write(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


if __name__ == "__main__":
    import uvicorn

    WEB_JOBS_DIR.mkdir(parents=True, exist_ok=True)
    port = int(os.getenv("EIA_WEB_PORT", "8001"))
    uvicorn.run("web_app:app", host="127.0.0.1", port=port, reload=False)
