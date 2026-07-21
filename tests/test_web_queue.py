import asyncio
import json
from datetime import datetime, timedelta
from io import BytesIO

import pytest
from fastapi import HTTPException, UploadFile

import web_app


@pytest.fixture(autouse=True)
def isolated_queue(tmp_path, monkeypatch):
    monkeypatch.setattr(web_app, "WEB_JOBS_DIR", tmp_path / "jobs")
    monkeypatch.setattr(web_app, "QUEUE_POLL_SECONDS", 0.01)
    web_app.WEB_JOBS_DIR.mkdir(parents=True)
    web_app.QUEUED_JOBS.clear()
    web_app.CURRENT_JOB_ID = None
    if web_app.RUN_LOCK.locked():
        web_app.RUN_LOCK.release()
    yield
    web_app.QUEUED_JOBS.clear()
    web_app.CURRENT_JOB_ID = None
    if web_app.RUN_LOCK.locked():
        web_app.RUN_LOCK.release()


def create_job_files(job_id, status, meta=None):
    job_dir = web_app.WEB_JOBS_DIR / job_id
    (job_dir / "input").mkdir(parents=True)
    web_app.write_json(job_dir / "status.json", {"job_id": job_id, **status})
    if meta is not None:
        web_app.write_json(job_dir / "input" / "project_meta.json", meta)
    return job_dir


def test_enqueue_job_updates_fifo_positions():
    first = create_job_files("job-1", {"status": "queued"})
    second = create_job_files("job-2", {"status": "queued"})

    assert web_app.enqueue_job("job-1") == 1
    assert web_app.enqueue_job("job-2") == 2

    assert web_app.read_status_file(first)["queue_position"] == 1
    assert web_app.read_status_file(second)["queue_position"] == 2


def test_recovery_orders_queued_jobs_and_marks_running_interrupted(monkeypatch):
    common_meta = {
        "run_surface_water": True,
        "run_noise": False,
        "enable_llm_text_polish": False,
        "enable_llm_extraction": False,
    }
    create_job_files(
        "job-later",
        {"status": "queued", "queued_at": "2026-01-02T10:00:00"},
        common_meta,
    )
    create_job_files(
        "job-first",
        {"status": "queued", "queued_at": "2026-01-01T10:00:00"},
        common_meta,
    )
    interrupted = create_job_files(
        "job-running",
        {"status": "running", "updated_at": "2026-01-01T09:00:00"},
        common_meta,
    )
    started = []

    class FakeThread:
        def __init__(self, *, target, args, name, daemon):
            self.target = target
            self.args = args
            self.name = name
            self.daemon = daemon

        def start(self):
            started.append(self.args[0])

    monkeypatch.setattr(web_app.threading, "Thread", FakeThread)

    web_app.recover_persisted_jobs()

    assert web_app.QUEUED_JOBS == ["job-first", "job-later"]
    assert started == ["job-first", "job-later"]
    interrupted_status = web_app.read_status_file(interrupted)
    assert interrupted_status["status"] == "failed"
    assert interrupted_status["current_step"] == "\u670d\u52a1\u4e2d\u65ad"
    assert interrupted_status["queue_position"] == 0


def test_recovery_rejects_queued_job_without_metadata():
    job_dir = create_job_files(
        "job-missing-meta",
        {"status": "queued", "queued_at": "2026-01-01T10:00:00"},
    )

    web_app.recover_persisted_jobs()

    status = web_app.read_status_file(job_dir)
    assert status["status"] == "failed"
    assert status["current_step"] == "\u6062\u590d\u5931\u8d25"


def test_cleanup_keeps_active_job_and_removes_expired_job(monkeypatch):
    active = create_job_files("active", {"status": "queued"})
    old_time = (datetime.now() - timedelta(days=30)).isoformat(timespec="seconds")
    expired = create_job_files(
        "expired",
        {"status": "success", "updated_at": old_time},
    )
    monkeypatch.setattr(web_app, "JOB_RETENTION_COUNT", 0)
    monkeypatch.setattr(web_app, "JOB_RETENTION_DAYS", 1)
    web_app.enqueue_job("active")

    web_app.cleanup_old_jobs()

    assert active.exists()
    assert not expired.exists()


def test_status_write_is_atomic_and_leaves_no_temp_files(tmp_path):
    job_dir = tmp_path / "atomic-job"
    job_dir.mkdir()

    web_app.update_status(job_dir, {"status": "queued"})
    web_app.update_status(job_dir, {"status": "running"})

    payload = json.loads((job_dir / "status.json").read_text(encoding="utf-8"))
    assert payload["status"] == "running"
    assert list(job_dir.glob(".*.tmp")) == []


def test_upload_limit_removes_partial_file(tmp_path, monkeypatch):
    monkeypatch.setattr(web_app, "MAX_UPLOAD_BYTES", 4)
    upload = UploadFile(filename="report.docx", file=BytesIO(b"12345"))
    target = tmp_path / "report.docx"

    with pytest.raises(HTTPException) as exc_info:
        asyncio.run(web_app.save_upload(upload, target))

    assert exc_info.value.status_code == 413
    assert not target.exists()


def test_single_worker_guard(monkeypatch):
    monkeypatch.setenv("WEB_CONCURRENCY", "2")
    with pytest.raises(RuntimeError, match="requires one server worker"):
        web_app.validate_single_worker_mode()
