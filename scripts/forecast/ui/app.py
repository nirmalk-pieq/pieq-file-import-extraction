"""Local FastAPI UI for running the commission statement forecast."""

from __future__ import annotations

import asyncio
import json
import os
import re
import signal
import subprocess
import threading
import time
import uuid as uuid_lib
from collections import deque
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any, Deque, Dict, List, Optional
from urllib import parse

from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

UI_DIR = Path(__file__).resolve().parent
FORECAST_DIR = UI_DIR.parent
REPO_ROOT = FORECAST_DIR.parent.parent
FORECAST_SCRIPT = FORECAST_DIR / "review_to_reconcile_file_forecast.py"
PAYABLES_SCRIPT = UI_DIR / "run_payables_comparison.py"
FORECAST_OUT = REPO_ROOT / "forecast-out"
PAYABLES_OUT = FORECAST_OUT / "payables_comparison"
STATIC_DIR = UI_DIR / "static"

UUID_RE = re.compile(
    r"(?i)\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b"
)
REVIEW_STAGE = 3
STAGE_LABELS = {
    1: "Cleared",
    2: "Processing",
    3: "Review",
    4: "Waiting",
}

STATEMENTS_SQL = """
SELECT
    uuid::text AS uuid,
    name,
    agency_id,
    stage,
    status,
    created_at,
    updated_at
FROM commission_statements
WHERE agency_id = %s
  AND (%s::int IS NULL OR stage = %s)
ORDER BY created_at DESC NULLS LAST, name
LIMIT %s
"""

@asynccontextmanager
async def lifespan(_: FastAPI):
    loop = asyncio.get_running_loop()
    state.loop = loop
    payables_state.loop = loop
    FORECAST_OUT.mkdir(parents=True, exist_ok=True)
    PAYABLES_OUT.mkdir(parents=True, exist_ok=True)
    yield


app = FastAPI(title="Forecast", docs_url=None, redoc_url=None, lifespan=lifespan)


class RunRequest(BaseModel):
    env: str = Field(..., pattern="^(prod|preprod)$")
    mode: str = Field(..., pattern="^(files|review)$")
    uuids: Optional[List[str]] = None
    agency_id: Optional[int] = None


PAYABLES_UPLOAD_SUFFIXES = {".csv", ".xlsx", ".xls"}


class RunState:
    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.process: Optional[subprocess.Popen] = None
        self.run_id: Optional[str] = None
        self.started_at: Optional[float] = None
        self.finished_at: Optional[float] = None
        self.exit_code: Optional[int] = None
        self.running = False
        self.command: List[str] = []
        self.lines: Deque[str] = deque(maxlen=8000)
        self.subscribers: List[asyncio.Queue] = []
        self.loop: Optional[asyncio.AbstractEventLoop] = None

    def snapshot(self) -> Dict[str, Any]:
        with self.lock:
            return {
                "running": self.running,
                "run_id": self.run_id,
                "started_at": self.started_at,
                "finished_at": self.finished_at,
                "exit_code": self.exit_code,
                "command": list(self.command),
                "line_count": len(self.lines),
            }


state = RunState()
payables_state = RunState()


def normalize_uuids(values: Optional[List[str]]) -> List[str]:
    if not values:
        raise ValueError("Select one or more statement files")
    seen = set()
    ordered: List[str] = []
    for value in values:
        text = str(value or "").strip()
        if not text:
            continue
        match = UUID_RE.search(text)
        if not match:
            raise ValueError(f"Not a commission statement UUID: {text}")
        key = match.group(0).lower()
        if key in seen:
            continue
        seen.add(key)
        ordered.append(str(uuid_lib.UUID(match.group(0))))
    if not ordered:
        raise ValueError("Select one or more statement files")
    return ordered


def build_command(req: RunRequest) -> List[str]:
    cmd = [resolve_python(), str(FORECAST_SCRIPT), req.env]
    if req.mode == "review":
        if req.agency_id is None:
            raise ValueError("Agency ID is required for Review mode")
        cmd.extend(["review", "--agency_id", str(req.agency_id)])
    else:
        cmd.extend(normalize_uuids(req.uuids))
    return cmd


def read_config(environment: str) -> Dict[str, str]:
    if environment not in {"prod", "preprod"}:
        raise ValueError("Environment must be prod or preprod")
    env_path = FORECAST_DIR / f"config_{environment}"
    if not env_path.is_file():
        raise FileNotFoundError(f"Config file not found: {env_path}")
    values: Dict[str, str] = {}
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            # Shell env wins, matching forecast script behavior.
            values[key] = os.environ.get(key, value).strip() if key in os.environ else value
    return values


def connect_from_config(cfg: Dict[str, str]):
    password = (cfg.get("PIEQ_DB_PASSWORD") or "").strip()
    if not password:
        raise RuntimeError("Missing PIEQ_DB_PASSWORD in config")
    explicit = (cfg.get("PIEQ_DB_CONNECTION_STRING") or "").strip()
    if explicit and "://" in explicit:
        url = explicit
    else:
        host = explicit or (cfg.get("PIEQ_DB_HOST") or "").strip()
        if not host:
            raise RuntimeError("Set PIEQ_DB_HOST or PIEQ_DB_CONNECTION_STRING")
        user = (cfg.get("PIEQ_DB_USERNAME") or "").strip()
        database = (cfg.get("PIEQ_DB_NAME") or "").strip()
        port = (cfg.get("PIEQ_DB_PORT") or "5432").strip() or "5432"
        if not user or not database:
            raise RuntimeError("Missing PIEQ_DB_USERNAME or PIEQ_DB_NAME in config")
        url = f"postgresql://{parse.quote(user, safe='')}@{host}:{port}/{database}"

    conninfo = url[len("jdbc:") :] if url.startswith("jdbc:") else url
    try:
        import psycopg

        return psycopg.connect(
            conninfo,
            password=password,
            autocommit=True,
            keepalives=1,
            keepalives_idle=30,
            keepalives_interval=10,
            keepalives_count=5,
        )
    except ImportError:
        import psycopg2

        conn = psycopg2.connect(conninfo, password=password)
        conn.set_session(readonly=True, autocommit=True)
        return conn


def fetch_statements(environment: str, agency_id: int, review_only: bool, limit: int) -> List[Dict[str, Any]]:
    cfg = read_config(environment)
    conn = connect_from_config(cfg)
    stage = REVIEW_STAGE if review_only else None
    try:
        with conn.cursor() as cur:
            cur.execute(STATEMENTS_SQL, (agency_id, stage, stage, limit))
            columns = [
                getattr(col, "name", None) or col[0] for col in cur.description
            ]
            rows = [dict(zip(columns, row)) for row in cur.fetchall()]
    finally:
        try:
            conn.close()
        except Exception:  # noqa: BLE001
            pass

    files = []
    for row in rows:
        stage_val = row.get("stage")
        created = row.get("created_at")
        updated = row.get("updated_at")
        files.append(
            {
                "uuid": row.get("uuid"),
                "name": row.get("name") or "(unnamed)",
                "agency_id": row.get("agency_id"),
                "stage": stage_val,
                "stage_label": STAGE_LABELS.get(stage_val, str(stage_val)),
                "status": row.get("status"),
                "created_at": created.isoformat() if hasattr(created, "isoformat") else created,
                "updated_at": updated.isoformat() if hasattr(updated, "isoformat") else updated,
            }
        )
    return files


def broadcast(run_state: RunState, event: str, data: Any) -> None:
    payload = {"event": event, "data": data}
    with run_state.lock:
        if event == "log" and isinstance(data, str):
            run_state.lines.append(data)
        subscribers = list(run_state.subscribers)
        loop = run_state.loop
    if not subscribers or loop is None:
        return

    def _push() -> None:
        for queue in subscribers:
            try:
                queue.put_nowait(payload)
            except asyncio.QueueFull:
                pass

    try:
        loop.call_soon_threadsafe(_push)
    except RuntimeError:
        pass


def reader_thread(run_state: RunState, proc: subprocess.Popen, run_id: str, label: str = "ui") -> None:
    assert proc.stdout is not None
    try:
        for raw in iter(proc.stdout.readline, ""):
            if not raw:
                break
            line = raw.rstrip("\n")
            broadcast(run_state, "log", line)
        proc.wait()
        code = proc.returncode
    except Exception as exc:  # noqa: BLE001
        broadcast(run_state, "log", f"[{label}] reader error: {exc}")
        code = proc.poll()
        if code is None:
            code = 1
    finally:
        with run_state.lock:
            if run_state.run_id == run_id:
                run_state.running = False
                run_state.exit_code = code
                run_state.finished_at = time.time()
                run_state.process = None
        broadcast(run_state, "done", {"exit_code": code, "run_id": run_id})


def resolve_python() -> str:
    override = (os.environ.get("FORECAST_PYTHON") or "").strip()
    if override:
        return override
    candidates = [
        REPO_ROOT / ".venv" / "bin" / "python",
        REPO_ROOT / ".venv" / "Scripts" / "python.exe",
        REPO_ROOT / ".venv" / "Scripts" / "python",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    # Windows installs usually expose `python`, not `python3`.
    return "python" if os.name == "nt" else "python3"


def subprocess_extra_kwargs() -> Dict[str, Any]:
    """Create a new process group so Stop can terminate the child tree on both OSes."""
    if os.name == "nt":
        # CREATE_NEW_PROCESS_GROUP lets us signal the child on Windows.
        return {"creationflags": subprocess.CREATE_NEW_PROCESS_GROUP}
    return {"start_new_session": True}


def terminate_process(proc: subprocess.Popen) -> None:
    if proc.poll() is not None:
        return
    try:
        if os.name == "nt":
            proc.terminate()
        else:
            os.killpg(proc.pid, signal.SIGTERM)
    except (ProcessLookupError, OSError, AttributeError):
        try:
            proc.terminate()
        except (ProcessLookupError, OSError):
            pass


def kill_process(proc: subprocess.Popen) -> None:
    if proc.poll() is not None:
        return
    try:
        if os.name == "nt":
            proc.kill()
        else:
            os.killpg(proc.pid, signal.SIGKILL)
    except (ProcessLookupError, OSError, AttributeError):
        try:
            proc.kill()
        except (ProcessLookupError, OSError):
            pass


def stop_managed_process(proc: subprocess.Popen, soft_timeout: float = 8.0) -> int:
    terminate_process(proc)
    try:
        return int(proc.wait(timeout=soft_timeout))
    except subprocess.TimeoutExpired:
        kill_process(proc)
        return int(proc.wait(timeout=3))


def sanitize_upload_name(name: str, fallback: str) -> str:
    base = Path(str(name or "").strip()).name
    if not base or base in {".", ".."}:
        base = fallback
    safe = re.sub(r"[^\w.\- ()]+", "_", base).strip(" ._")
    return safe or fallback


async def save_upload(upload: UploadFile, dest_dir: Path, label: str) -> Path:
    if upload is None or not upload.filename:
        raise HTTPException(status_code=400, detail=f"Select an {label} file")
    filename = sanitize_upload_name(upload.filename, f"{label.lower()}.csv")
    suffix = Path(filename).suffix.lower()
    if suffix not in PAYABLES_UPLOAD_SUFFIXES:
        raise HTTPException(
            status_code=400,
            detail=f"{label} must be a CSV or Excel file (.csv, .xlsx, .xls)",
        )
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / filename
    # Avoid clobbering if both uploads somehow share a name.
    if dest.exists():
        dest = dest_dir / f"{dest.stem}_{label.lower()}{dest.suffix}"
    try:
        data = await upload.read()
    finally:
        await upload.close()
    if not data:
        raise HTTPException(status_code=400, detail=f"{label} file is empty")
    dest.write_bytes(data)
    return dest


def make_sse_stream(run_state: RunState) -> StreamingResponse:
    queue: asyncio.Queue = asyncio.Queue(maxsize=2000)
    with run_state.lock:
        run_state.subscribers.append(queue)
        backlog = list(run_state.lines)
        snap = {
            "running": run_state.running,
            "run_id": run_state.run_id,
            "exit_code": run_state.exit_code,
            "command": list(run_state.command),
        }

    async def event_gen():
        try:
            yield f"event: status\ndata: {json.dumps(snap)}\n\n"
            for line in backlog:
                yield f"event: log\ndata: {json.dumps(line)}\n\n"
            while True:
                try:
                    item = await asyncio.wait_for(queue.get(), timeout=15.0)
                except asyncio.TimeoutError:
                    yield "event: ping\ndata: {}\n\n"
                    continue
                event = item.get("event", "message")
                yield f"event: {event}\ndata: {json.dumps(item.get('data'))}\n\n"
                if event == "done":
                    break
        finally:
            with run_state.lock:
                if queue in run_state.subscribers:
                    run_state.subscribers.remove(queue)

    return StreamingResponse(
        event_gen(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


def collect_payables_files(folder: Path) -> List[Dict[str, str]]:
    files: List[Dict[str, str]] = []
    preferred = (
        "Grouped_Policy_Combined.xlsx",
        "Grouped_Policy_Combined.csv",
        "Validation_Output.xlsx",
        "Validation_Output.csv",
        "Passed_Output.xlsx",
        "Passed_Output.csv",
        "Grouped_Policy_Passed.xlsx",
        "Grouped_Policy_Failed.xlsx",
    )
    for name in preferred:
        path = folder / name
        if path.is_file():
            files.append(
                {
                    "name": name,
                    "path": path.relative_to(FORECAST_OUT).as_posix(),
                }
            )
    segregated = folder / "Segregated_Errors"
    if segregated.is_dir():
        for path in sorted(segregated.iterdir()):
            if path.is_file() and path.suffix.lower() in {".csv", ".xlsx"}:
                files.append(
                    {
                        "name": f"Segregated_Errors/{path.name}",
                        "path": path.relative_to(FORECAST_OUT).as_posix(),
                    }
                )
    return files


@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    index_path = STATIC_DIR / "index.html"
    if not index_path.is_file():
        raise HTTPException(status_code=500, detail="UI index.html is missing")
    return HTMLResponse(index_path.read_text(encoding="utf-8"))


@app.get("/api/status")
async def status() -> Dict[str, Any]:
    snap = state.snapshot()
    with state.lock:
        recent = list(state.lines)[-80:]
    snap["recent_lines"] = recent
    return snap


@app.get("/api/statements")
async def list_statements(
    env: str = Query(..., pattern="^(prod|preprod)$"),
    agency_id: int = Query(..., ge=1),
    review_only: bool = Query(True),
    limit: int = Query(200, ge=1, le=500),
) -> Dict[str, Any]:
    try:
        files = await asyncio.to_thread(fetch_statements, env, agency_id, review_only, limit)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Could not load statements: {exc}") from exc
    return {
        "env": env,
        "agency_id": agency_id,
        "review_only": review_only,
        "count": len(files),
        "files": files,
    }


@app.post("/api/run")
async def start_run(req: RunRequest) -> Dict[str, Any]:
    try:
        command = build_command(req)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    with state.lock:
        if state.running and state.process and state.process.poll() is None:
            raise HTTPException(status_code=409, detail="A forecast is already running")
        run_id = str(uuid_lib.uuid4())
        state.run_id = run_id
        state.started_at = time.time()
        state.finished_at = None
        state.exit_code = None
        state.running = True
        state.command = command
        state.lines.clear()

    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"

    try:
        proc = subprocess.Popen(
            command,
            cwd=str(REPO_ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            env=env,
            **subprocess_extra_kwargs(),
        )
    except OSError as exc:
        with state.lock:
            state.running = False
            state.exit_code = 1
            state.finished_at = time.time()
            state.process = None
        raise HTTPException(status_code=500, detail=f"Failed to start forecast: {exc}") from exc

    with state.lock:
        state.process = proc

    broadcast(state, "log", f"[ui] Starting: {' '.join(command)}")
    broadcast(state, "log", f"[ui] Working directory: {REPO_ROOT}")
    threading.Thread(
        target=reader_thread, args=(state, proc, run_id, "ui"), daemon=True
    ).start()
    return {"ok": True, "run_id": run_id, "command": command}


@app.post("/api/run/stop")
async def stop_run() -> Dict[str, Any]:
    with state.lock:
        proc = state.process
        run_id = state.run_id
        running = state.running
    if not running or proc is None or proc.poll() is not None:
        return {"ok": True, "stopped": False, "message": "No active forecast"}
    broadcast(state, "log", "[ui] Stop requested — terminating forecast process")
    exit_code = stop_managed_process(proc)
    with state.lock:
        if state.run_id == run_id:
            state.running = False
            state.exit_code = exit_code
            state.finished_at = time.time()
            state.process = None
    broadcast(
        state,
        "done",
        {"exit_code": exit_code, "run_id": run_id, "stopped": True},
    )
    return {"ok": True, "stopped": True, "exit_code": exit_code}


@app.get("/api/run/stream")
async def stream_run() -> StreamingResponse:
    return make_sse_stream(state)


@app.get("/api/payables/status")
async def payables_status() -> Dict[str, Any]:
    snap = payables_state.snapshot()
    with payables_state.lock:
        recent = list(payables_state.lines)[-80:]
    snap["recent_lines"] = recent
    return snap


@app.post("/api/payables/run")
async def start_payables_run(
    apl_file: UploadFile = File(...),
    icm_file: UploadFile = File(...),
    mode: str = Form("icm"),
    tolerance: float = Form(0.02),
) -> Dict[str, Any]:
    mode_norm = str(mode or "icm").strip().lower()
    if mode_norm not in {"apl", "icm"}:
        raise HTTPException(status_code=400, detail="Mode must be apl or icm")
    try:
        tolerance_val = float(tolerance)
    except (TypeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail="Tolerance must be a number") from exc
    if tolerance_val < 0:
        raise HTTPException(status_code=400, detail="Tolerance must be non-negative")

    if not PAYABLES_SCRIPT.is_file():
        raise HTTPException(status_code=500, detail="Payables runner script is missing")

    with payables_state.lock:
        if payables_state.running and payables_state.process and payables_state.process.poll() is None:
            raise HTTPException(status_code=409, detail="A payables comparison is already running")
        run_id = str(uuid_lib.uuid4())

    out_dir = PAYABLES_OUT / run_id
    input_dir = out_dir / "input"
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        apl_path = await save_upload(apl_file, input_dir, "APL")
        icm_path = await save_upload(icm_file, input_dir, "ICM")
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Could not save uploaded files: {exc}") from exc

    command = [
        resolve_python(),
        str(PAYABLES_SCRIPT),
        "--apl",
        str(apl_path),
        "--parallel",
        str(icm_path),
        "--out-dir",
        str(out_dir),
        "--mode",
        mode_norm,
        "--tolerance",
        str(tolerance_val),
    ]

    with payables_state.lock:
        payables_state.run_id = run_id
        payables_state.started_at = time.time()
        payables_state.finished_at = None
        payables_state.exit_code = None
        payables_state.running = True
        payables_state.command = command
        payables_state.lines.clear()

    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"

    try:
        proc = subprocess.Popen(
            command,
            cwd=str(REPO_ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            env=env,
            **subprocess_extra_kwargs(),
        )
    except OSError as exc:
        with payables_state.lock:
            payables_state.running = False
            payables_state.exit_code = 1
            payables_state.finished_at = time.time()
            payables_state.process = None
        raise HTTPException(status_code=500, detail=f"Failed to start payables comparison: {exc}") from exc

    with payables_state.lock:
        payables_state.process = proc

    broadcast(payables_state, "log", f"[payables] APL file: {apl_path.name}")
    broadcast(payables_state, "log", f"[payables] ICM file: {icm_path.name}")
    broadcast(payables_state, "log", f"[payables] Starting comparison (mode={mode_norm}, tolerance={tolerance_val})")
    broadcast(payables_state, "log", f"[payables] Output folder: {out_dir}")
    threading.Thread(
        target=reader_thread,
        args=(payables_state, proc, run_id, "payables"),
        daemon=True,
    ).start()
    return {
        "ok": True,
        "run_id": run_id,
        "command": command,
        "out_dir": str(out_dir.relative_to(FORECAST_OUT)),
        "apl_file": apl_path.name,
        "icm_file": icm_path.name,
    }


@app.post("/api/payables/stop")
async def stop_payables_run() -> Dict[str, Any]:
    with payables_state.lock:
        proc = payables_state.process
        run_id = payables_state.run_id
        running = payables_state.running
    if not running or proc is None or proc.poll() is not None:
        return {"ok": True, "stopped": False, "message": "No active payables comparison"}
    broadcast(payables_state, "log", "[payables] Stop requested — terminating comparison")
    exit_code = stop_managed_process(proc)
    with payables_state.lock:
        if payables_state.run_id == run_id:
            payables_state.running = False
            payables_state.exit_code = exit_code
            payables_state.finished_at = time.time()
            payables_state.process = None
    broadcast(
        payables_state,
        "done",
        {"exit_code": exit_code, "run_id": run_id, "stopped": True},
    )
    return {"ok": True, "stopped": True, "exit_code": exit_code}


@app.get("/api/payables/stream")
async def stream_payables_run() -> StreamingResponse:
    return make_sse_stream(payables_state)


@app.get("/api/payables/outputs")
async def list_payables_outputs(limit: int = Query(40, ge=1, le=200)) -> Dict[str, Any]:
    PAYABLES_OUT.mkdir(parents=True, exist_ok=True)
    candidates = []
    for entry in PAYABLES_OUT.iterdir():
        if not entry.is_dir() or entry.name.startswith("."):
            continue
        candidates.append((entry.stat().st_mtime, entry))
    candidates.sort(key=lambda item: item[0], reverse=True)

    runs: List[Dict[str, Any]] = []
    for mtime, entry in candidates[:limit]:
        runs.append(
            {
                "folder": entry.name,
                "mtime": mtime,
                "files": collect_payables_files(entry),
            }
        )
    return {"root": str(PAYABLES_OUT), "runs": runs}


def safe_under_forecast_out(rel_path: str) -> Path:
    if not rel_path or rel_path.startswith("/") or ".." in Path(rel_path).parts:
        raise HTTPException(status_code=400, detail="Invalid path")
    target = (FORECAST_OUT / rel_path).resolve()
    root = FORECAST_OUT.resolve()
    if root not in target.parents and target != root:
        raise HTTPException(status_code=400, detail="Path outside forecast-out")
    if not target.is_file():
        raise HTTPException(status_code=404, detail="File not found")
    return target


def collect_xlsx(folder: Path) -> List[Dict[str, str]]:
    files = []
    for name in (
        "forecast.xlsx",
        "Policies_To_Update.xlsx",
        "Policies_To_Update_Comparison.xlsx",
    ):
        path = folder / name
        if path.is_file():
            rel = path.relative_to(FORECAST_OUT).as_posix()
            files.append({"name": name, "path": rel})
    # Also surface consolidated Policies_To_Update at forecast-out root via parent listing.
    return files


@app.get("/api/outputs")
async def list_outputs(limit: int = Query(40, ge=1, le=200)) -> Dict[str, Any]:
    FORECAST_OUT.mkdir(parents=True, exist_ok=True)
    runs: List[Dict[str, Any]] = []
    root_files = []
    for name in ("Policies_To_Update.xlsx", "Policies_To_Update_Comparison.xlsx"):
        path = FORECAST_OUT / name
        if path.is_file():
            root_files.append(
                {
                    "name": name,
                    "path": name,
                    "mtime": path.stat().st_mtime,
                    "size": path.stat().st_size,
                }
            )

    candidates = []
    for entry in FORECAST_OUT.iterdir():
        if not entry.is_dir():
            continue
        if entry.name.startswith(".") or entry.name in {"Before Update", "payables_comparison", "upload"}:
            continue
        mtime = entry.stat().st_mtime
        candidates.append((mtime, entry))
    candidates.sort(key=lambda item: item[0], reverse=True)

    for mtime, entry in candidates[:limit]:
        summary = None
        summary_path = entry / "summary.json"
        if summary_path.is_file():
            try:
                summary = json.loads(summary_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                summary = None
        xlsx = collect_xlsx(entry)
        # Include nested validation xlsx lightly
        validation = entry / "apl_icm_validation"
        if validation.is_dir():
            for path in sorted(validation.glob("*.xlsx")):
                xlsx.append(
                    {
                        "name": f"apl_icm_validation/{path.name}",
                        "path": path.relative_to(FORECAST_OUT).as_posix(),
                    }
                )
        runs.append(
            {
                "folder": entry.name,
                "mtime": mtime,
                "summary": summary,
                "files": xlsx,
            }
        )

    return {"root": str(FORECAST_OUT), "root_files": root_files, "runs": runs}


@app.get("/api/outputs/file")
async def download_file(path: str = Query(...)) -> FileResponse:
    target = safe_under_forecast_out(path)
    return FileResponse(
        path=target,
        filename=target.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if target.suffix.lower() == ".xlsx"
        else "application/octet-stream",
    )


app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
