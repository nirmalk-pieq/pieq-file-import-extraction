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

from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

UI_DIR = Path(__file__).resolve().parent
FORECAST_DIR = UI_DIR.parent
REPO_ROOT = FORECAST_DIR.parent.parent
FORECAST_SCRIPT = FORECAST_DIR / "review_to_reconcile_file_forecast.py"
FORECAST_OUT = REPO_ROOT / "forecast-out"
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
    state.loop = asyncio.get_running_loop()
    FORECAST_OUT.mkdir(parents=True, exist_ok=True)
    yield


app = FastAPI(title="Forecast", docs_url=None, redoc_url=None, lifespan=lifespan)


class RunRequest(BaseModel):
    env: str = Field(..., pattern="^(prod|preprod)$")
    mode: str = Field(..., pattern="^(files|review)$")
    uuids: Optional[List[str]] = None
    agency_id: Optional[int] = None


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
    python = os.environ.get("FORECAST_PYTHON") or "python3"
    venv_python = REPO_ROOT / ".venv" / "bin" / "python"
    if venv_python.is_file():
        python = str(venv_python)
    cmd = [python, str(FORECAST_SCRIPT), req.env]
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


def broadcast(event: str, data: Any) -> None:
    payload = {"event": event, "data": data}
    with state.lock:
        if event == "log" and isinstance(data, str):
            state.lines.append(data)
        subscribers = list(state.subscribers)
        loop = state.loop
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


def reader_thread(proc: subprocess.Popen, run_id: str) -> None:
    assert proc.stdout is not None
    try:
        for raw in iter(proc.stdout.readline, ""):
            if not raw:
                break
            line = raw.rstrip("\n")
            broadcast("log", line)
        proc.wait()
        code = proc.returncode
    except Exception as exc:  # noqa: BLE001
        broadcast("log", f"[ui] reader error: {exc}")
        code = proc.poll()
        if code is None:
            code = 1
    finally:
        with state.lock:
            if state.run_id == run_id:
                state.running = False
                state.exit_code = code
                state.finished_at = time.time()
                state.process = None
        broadcast("done", {"exit_code": code, "run_id": run_id})


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
            start_new_session=True,
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

    broadcast("log", f"[ui] Starting: {' '.join(command)}")
    broadcast("log", f"[ui] Working directory: {REPO_ROOT}")
    threading.Thread(target=reader_thread, args=(proc, run_id), daemon=True).start()
    return {"ok": True, "run_id": run_id, "command": command}


@app.post("/api/run/stop")
async def stop_run() -> Dict[str, Any]:
    with state.lock:
        proc = state.process
        run_id = state.run_id
        running = state.running
    if not running or proc is None or proc.poll() is not None:
        return {"ok": True, "stopped": False, "message": "No active forecast"}
    try:
        os.killpg(proc.pid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    broadcast("log", "[ui] Stop requested — terminating forecast process")
    try:
        proc.wait(timeout=8)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(proc.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        proc.wait(timeout=3)
    with state.lock:
        if state.run_id == run_id:
            state.running = False
            state.exit_code = proc.returncode
            state.finished_at = time.time()
            state.process = None
    broadcast("done", {"exit_code": proc.returncode, "run_id": run_id, "stopped": True})
    return {"ok": True, "stopped": True, "exit_code": proc.returncode}


@app.get("/api/run/stream")
async def stream_run() -> StreamingResponse:
    queue: asyncio.Queue = asyncio.Queue(maxsize=2000)
    with state.lock:
        state.subscribers.append(queue)
        backlog = list(state.lines)
        snap = {
            "running": state.running,
            "run_id": state.run_id,
            "exit_code": state.exit_code,
            "command": list(state.command),
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
            with state.lock:
                if queue in state.subscribers:
                    state.subscribers.remove(queue)

    return StreamingResponse(
        event_gen(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


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
        if entry.name.startswith(".") or entry.name == "Before Update":
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
