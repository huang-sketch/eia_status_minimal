import json
import os
import shutil
import subprocess
import sys
import time
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
from docx_numbering import reset_numbering
from section_docx_combiner import (
    COMBINED_SECTION_FILENAME,
    NOISE_SECTION_FILENAME,
    PROJECT_AREA_OVERVIEW_FILENAME,
    SURFACE_WATER_SECTION_FILENAME,
    build_combined_section_docx,
)
from section_docx_finalize import finalize_and_rebuild_section_docx
from word_processor import load_docx_chunks
from table_schema_mapper import summarize_schema_status


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
            "extraction/eia_result.json",
            "extraction/records.json",
            "extraction/meta.json",
            "debug_tables/extraction_summary.json",
            "debug_tables/detected_factors.json",
            "debug_tables/project_area_overview_sources.json",
            "debug_tables/project_area_overview_llm_input.json",
            "debug_tables/project_area_overview_llm_output.json",
            "debug_tables/project_area_overview_status.json",
            "debug_tables/noise_monitor_points_table.json",
            "debug_tables/surface_water_monitor_points_table.json",
            "debug_tables/surface_water_monitor_results_table.json",
            "debug_tables/table_schema_detection.json",
            "debug_tables/table_schema_llm_input.json",
            "debug_tables/table_schema_llm_output.json",
            "debug_tables/table_schema_validation.json",
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
            COMBINED_SECTION_FILENAME,
            "eia_outputs.zip",
        ],
    },
}

app = FastAPI(title="环评现状分析自动化系统")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/api/health")
def health() -> JSONResponse:
    return JSONResponse({"status": "ok"})


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/jobs/{job_id}")
def job_page(job_id: str) -> FileResponse:
    if not job_id or any(part in job_id for part in ("..", "/", "\\")):
        raise HTTPException(status_code=404, detail="任务不存在")
    return FileResponse(STATIC_DIR / "job.html")


@app.post("/api/jobs")
async def create_job(
    background_tasks: BackgroundTasks,
    report_name: str = Form(...),
    admin_division: str = Form(...),
    run_surface_water: bool = Form(True),
    run_noise: bool = Form(True),
    enable_llm_text_polish: bool = Form(False),
    enable_llm_extraction: bool = Form(True),
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
            "enable_llm_extraction": enable_llm_extraction,
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
            "enable_llm_text_polish": enable_llm_text_polish,
            "enable_llm_extraction": enable_llm_extraction,
            "llm_text_polish": build_llm_text_polish_status(output_dir, enable_llm_text_polish),
            "schema_fallback": summarize_schema_status(output_dir, enable_llm_extraction),
        },
    )
    background_tasks.add_task(
        run_job,
        job_id,
        run_surface_water,
        run_noise,
        enable_llm_text_polish,
        enable_llm_extraction,
    )
    return JSONResponse({"job_id": job_id, "status": "pending"})


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str) -> JSONResponse:
    job_dir = resolve_job_dir(job_id)
    status = read_status(job_dir)
    status["result_groups"] = build_result_groups(job_id)
    status["llm_text_polish"] = build_llm_text_polish_status(
        job_dir / "output",
        bool(status.get("enable_llm_text_polish")),
    )
    status["schema_fallback"] = summarize_schema_status(
        job_dir / "output",
        bool(status.get("enable_llm_extraction", True)),
    )
    meta_path = job_dir / "input" / "project_meta.json"
    if meta_path.exists():
        status["project_meta"] = json.loads(meta_path.read_text(encoding="utf-8"))
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
    enable_llm_extraction: bool,
) -> None:
    job_dir = WEB_JOBS_DIR / job_id
    input_dir = job_dir / "input"
    output_dir = job_dir / "output"
    log_path = job_dir / "job.log"

    try:
        update_status(job_dir, {"status": "running", "current_step": "初始化任务", "error": None})
        update_status(
            job_dir,
            {
                "enable_llm_text_polish": enable_llm_text_polish,
                "llm_text_polish": build_llm_text_polish_status(output_dir, enable_llm_text_polish),
                "schema_fallback": summarize_schema_status(output_dir, enable_llm_extraction),
            },
        )
        append_log(log_path, f"job_id: {job_id}")
        append_log(log_path, f"input_dir: {input_dir}")
        append_log(log_path, f"output_dir: {output_dir}")

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["EIA_INPUT_DIR"] = str(input_dir)
        env["EIA_OUTPUT_DIR"] = str(output_dir)
        env["ENABLE_LLM_TEXT_POLISH"] = "true" if enable_llm_text_polish else "false"
        env["ENABLE_LLM_EXTRACTION"] = "true" if enable_llm_extraction else "false"
        env["ENABLE_SCHEMA_FALLBACK"] = "true" if enable_llm_extraction else "false"
        env.setdefault("EIA_MAX_CHUNKS_PER_RUN", "100")
        reset_numbering(output_dir)
        cleanup_stale_section_docx(output_dir, run_noise, run_surface_water)

        update_status(job_dir, {"current_step": "生成项目区域环境概况"})
        run_optional_script("project_area_overview_generator.py", env, log_path)

        if run_noise:
            update_status(job_dir, {"current_step": "声环境：预处理噪声表格"})
            noise_prepare_started = time.perf_counter()
            prepare_noise_debug_tables(input_dir, output_dir, log_path)
            append_log(log_path, f"step duration: 声环境：预处理噪声表格 {time.perf_counter() - noise_prepare_started:.2f}s")
            update_status(job_dir, {"current_step": "声环境：生成 Word 章节"})
            run_script("noise_section_generator.py", env, log_path)
            update_status(
                job_dir,
                {
                    "llm_text_polish": build_llm_text_polish_status(output_dir, enable_llm_text_polish),
                    "schema_fallback": summarize_schema_status(output_dir, enable_llm_extraction),
                },
            )
            schema_status = summarize_schema_status(output_dir, enable_llm_extraction)
            append_log(log_path, f"schema fallback status: {schema_status['label']}; elapsed {schema_status.get('elapsed_ms', 0)}ms")

        if run_surface_water:
            update_status(job_dir, {"current_step": "地表水：CLI 监测数据抽取"})
            run_script("monitoring_extraction.py", env, log_path)
            update_status(job_dir, {"current_step": "地表水：解析监测方案和监测报告，执行达标判定"})
            run_script("surface_water_pipeline.py", env, log_path)
            update_status(job_dir, {"schema_fallback": summarize_schema_status(output_dir, enable_llm_extraction)})
            schema_status = summarize_schema_status(output_dir, enable_llm_extraction)
            append_log(log_path, f"schema fallback status: {schema_status['label']}; elapsed {schema_status.get('elapsed_ms', 0)}ms")
            if has_surface_water_monitoring_data(output_dir):
                update_status(job_dir, {"current_step": "地表水：生成 Word 章节"})
                run_script("surface_water_section_generator.py", env, log_path)
            else:
                cleanup_surface_water_section(output_dir)
                append_log(log_path, "surface water section skipped: no surface water monitoring data")

        update_status(job_dir, {"current_step": "刷新章节编号并重建 Word"})
        finalize_started = time.perf_counter()
        numbering_plan = finalize_and_rebuild_section_docx(output_dir)
        append_log(log_path, f"section numbering plan: {numbering_plan}")
        append_log(log_path, f"step duration: 刷新章节编号并重建 Word {time.perf_counter() - finalize_started:.2f}s")

        update_status(job_dir, {"current_step": "生成合并版 Word"})
        combine_started = time.perf_counter()
        combined_path = build_combined_section_docx(output_dir)
        append_log(log_path, f"combined section generated: {combined_path.relative_to(output_dir)}")
        append_log(log_path, f"step duration: 生成合并版 Word {time.perf_counter() - combine_started:.2f}s")

        update_status(job_dir, {"current_step": "打包输出文件"})
        zip_started = time.perf_counter()
        create_output_zip(output_dir)
        append_log(log_path, f"step duration: 打包输出文件 {time.perf_counter() - zip_started:.2f}s")
        update_status(
            job_dir,
            {
                "status": "success",
                "current_step": "完成",
                "error": None,
                "result_groups": build_result_groups(job_id),
                "llm_text_polish": build_llm_text_polish_status(output_dir, enable_llm_text_polish),
                "schema_fallback": summarize_schema_status(output_dir, enable_llm_extraction),
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
                "llm_text_polish": build_llm_text_polish_status(output_dir, enable_llm_text_polish),
                "schema_fallback": summarize_schema_status(output_dir, enable_llm_extraction),
            },
        )


def cleanup_stale_section_docx(output_dir: Path, run_noise: bool, run_surface_water: bool) -> None:
    output_dir = Path(output_dir)
    if not run_noise:
        noise_path = output_dir / NOISE_SECTION_FILENAME
        if noise_path.exists():
            noise_path.unlink()
    if not run_surface_water:
        cleanup_surface_water_section(output_dir)


def cleanup_surface_water_section(output_dir: Path) -> None:
    surface_path = Path(output_dir) / SURFACE_WATER_SECTION_FILENAME
    if surface_path.exists():
        surface_path.unlink()


def has_surface_water_monitoring_data(output_dir: Path) -> bool:
    for relative_path in ("monitoring_records.json", "compliance_results.json"):
        path = Path(output_dir) / relative_path
        if not path.exists():
            continue
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if isinstance(payload, list) and any(is_surface_water_record(item) for item in payload):
            return True
    return False


def is_surface_water_record(item: Any) -> bool:
    if not isinstance(item, dict):
        return False
    monitor_type = str(item.get("monitor_type") or "")
    source_type = str(item.get("source_type") or "")
    point_code = str(item.get("point_code") or "")
    if monitor_type == "surface_water":
        return True
    if point_code.upper().startswith("WJ"):
        return True
    return "地表水" in source_type


def run_script(script_name: str, env: Dict[str, str], log_path: Path) -> None:
    append_log(log_path, f"> python {script_name}")
    started = time.perf_counter()
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
    append_log(log_path, f"step duration: {script_name} {time.perf_counter() - started:.2f}s")
    if completed.returncode != 0:
        raise RuntimeError(f"{script_name} 执行失败，退出码 {completed.returncode}")


def run_optional_script(script_name: str, env: Dict[str, str], log_path: Path) -> bool:
    try:
        run_script(script_name, env, log_path)
        return True
    except Exception as exc:
        append_log(log_path, f"optional step skipped: {script_name}: {exc}")
        append_log(log_path, traceback.format_exc())
        return False


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
    excluded_word_sections = {
        NOISE_SECTION_FILENAME,
        PROJECT_AREA_OVERVIEW_FILENAME,
        SURFACE_WATER_SECTION_FILENAME,
    }
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in output_dir.rglob("*"):
            if not path.is_file() or path == zip_path:
                continue
            if path.name in excluded_word_sections:
                continue
            if path.suffix.lower() == ".docx" and path.name != COMBINED_SECTION_FILENAME:
                continue
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
        "llm_text_polish": build_llm_text_polish_status(output_dir, None),
        "schema_fallback": summarize_schema_status(output_dir, True),
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


def _is_rule_text_fallback(validation: Dict[str, Any]) -> bool:
    if validation.get("fallback") == "rule_texts" and validation.get("valid") is True:
        return True
    if validation.get("used_llm") is not True or validation.get("valid") is not True:
        return False
    if validation.get("llm_applied") is False:
        return True
    warnings = validation.get("warnings") if isinstance(validation.get("warnings"), list) else []
    markers = ("网络连接异常", "网络异常", "已使用规则文本", "failed after all retries")
    return any(any(marker in str(item) for marker in markers) for item in warnings)


def _polish_validation_label(validation: Dict[str, Any]) -> tuple[str, str, bool]:
    used_llm = validation.get("used_llm") is True
    valid = validation.get("valid") is True
    if _is_rule_text_fallback(validation):
        if validation.get("error_type") == "network" or _warnings_indicate_network(validation):
            return "fallback", "网络异常，已使用规则文本", True
        return "fallback", "已使用规则文本", True
    if not used_llm:
        state = "not_enabled" if validation.get("valid") else "failed"
        return state, "未启用" if state == "not_enabled" else "未调用", False
    if valid:
        return "success", "调用成功", True
    return "failed", "润色失败（已使用规则文本）", False


def _warnings_indicate_network(validation: Dict[str, Any]) -> bool:
    if validation.get("error_type") == "network":
        return True
    warnings = validation.get("warnings") if isinstance(validation.get("warnings"), list) else []
    markers = ("网络连接异常", "网络异常", "10054", "urlopen error", "RemoteDisconnected")
    return any(any(marker in str(item) for marker in markers) for item in warnings)


def _collect_polish_warnings(name: str, validation: Dict[str, Any], state: str) -> List[str]:
    if state in {"fallback", "success"}:
        return []
    items = validation.get("warnings") if isinstance(validation.get("warnings"), list) else []
    return [f"{name}: {item}" for item in items if item]


def _build_multi_polish_label(
    succeeded: List[str],
    fallback: List[str],
    failed: List[str],
) -> str:
    parts: List[str] = []
    if succeeded:
        parts.append("、".join(succeeded) + "调用成功")
    if fallback:
        parts.append("、".join(fallback) + "网络异常，已使用规则文本")
    if failed:
        parts.append("、".join(failed) + "润色失败（已使用规则文本）")
    return "；".join(parts)


def build_llm_text_polish_status(output_dir: Path, enabled: Any = None) -> Dict[str, Any]:
    validations = {
        "地表水": read_optional_json(output_dir / "debug_tables/surface_water_llm_text_validation.json"),
        "声环境": read_optional_json(output_dir / "debug_tables/noise_llm_text_validation.json"),
    }
    available = {
        name: validation
        for name, validation in validations.items()
        if isinstance(validation, dict) and validation
    }
    if len(available) > 1:
        warnings: List[str] = []
        failed: List[str] = []
        succeeded: List[str] = []
        fallback: List[str] = []
        used_any = False
        overall_valid = True
        for name, validation in available.items():
            used_llm = validation.get("used_llm") is True
            valid = validation.get("valid") is True
            used_any = used_any or used_llm
            overall_valid = overall_valid and valid
            state, _, _ = _polish_validation_label(validation)
            if state == "success":
                succeeded.append(name)
            elif state == "fallback":
                fallback.append(name)
            elif validation.get("valid") is not True:
                failed.append(name)
            warnings.extend(_collect_polish_warnings(name, validation, state))
        if failed:
            return {
                "enabled": enabled if enabled is not None else used_any,
                "state": "failed",
                "label": _build_multi_polish_label(succeeded, fallback, failed),
                "used_llm": used_any,
                "valid": False,
                "warnings": warnings,
            }
        if succeeded and fallback:
            return {
                "enabled": enabled if enabled is not None else used_any,
                "state": "fallback",
                "label": _build_multi_polish_label(succeeded, fallback, failed),
                "used_llm": used_any,
                "valid": overall_valid,
                "warnings": warnings,
            }
        if fallback:
            return {
                "enabled": enabled if enabled is not None else used_any,
                "state": "fallback",
                "label": _build_multi_polish_label(succeeded, fallback, failed),
                "used_llm": used_any,
                "valid": overall_valid,
                "warnings": warnings,
            }
        if succeeded:
            return {
                "enabled": enabled if enabled is not None else used_any,
                "state": "success",
                "label": _build_multi_polish_label(succeeded, fallback, failed),
                "used_llm": used_any,
                "valid": True,
                "warnings": warnings,
            }

    validation = available.get("地表水") or available.get("声环境") or {}
    if not isinstance(validation, dict) or not validation:
        if enabled is False:
            return {
                "enabled": False,
                "state": "not_enabled",
                "label": "未启用",
                "warnings": [],
            }
        return {
            "enabled": bool(enabled),
            "state": "pending" if enabled else "unknown",
            "label": "等待结果" if enabled else "暂无结果",
            "warnings": [],
        }

    used_llm = validation.get("used_llm") is True
    valid = validation.get("valid") is True
    warnings = validation.get("warnings") if isinstance(validation.get("warnings"), list) else []
    state, label, _ = _polish_validation_label(validation)
    if state == "fallback":
        warnings = []
    return {
        "enabled": enabled if enabled is not None else used_llm,
        "state": state,
        "label": label,
        "used_llm": used_llm,
        "valid": valid,
        "warnings": [str(item) for item in warnings],
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
    port = int(os.getenv("EIA_WEB_PORT", "8010"))
    # 任务运行时会频繁写入 runs/web_jobs，若开启 reload 会导致服务反复重启、前端 Failed to fetch
    reload = os.getenv("EIA_WEB_RELOAD", "false").lower() == "true"
    uvicorn_kwargs: Dict[str, Any] = {
        "host": "127.0.0.1",
        "port": port,
        "reload": reload,
    }
    if reload:
        uvicorn_kwargs["reload_excludes"] = [
            "runs/*",
            "runs/**",
            "**/__pycache__/*",
            "**/*.pyc",
        ]
    uvicorn.run("web_app:app", **uvicorn_kwargs)
