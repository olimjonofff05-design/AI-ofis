"""AI Ofis serveri. Ishga tushirish: python -m uvicorn server.app:app --port 8000"""
import asyncio
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")

from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect  # noqa: E402
from fastapi.responses import FileResponse  # noqa: E402
from pydantic import BaseModel  # noqa: E402

from . import agents  # noqa: E402
from .bus import bus  # noqa: E402
from .llm import get_llm  # noqa: E402

app = FastAPI(title="AI Ofis")
agents.load()


class TaskIn(BaseModel):
    text: str


class ApproveIn(BaseModel):
    ok: bool


@app.get("/")
def index():
    return FileResponse(ROOT / "office-ui" / "index.html")


@app.get("/api/status")
def status():
    llm = get_llm()
    return {"live": True, "ai": llm.enabled, "model": llm.model if llm.enabled else None,
            "connected": sorted(agents.CONNECTED)}


@app.get("/api/tasks")
def tasks():
    return agents.public()


@app.post("/api/tasks")
async def new_task(body: TaskIn):
    if not body.text.strip():
        raise HTTPException(400, "Vazifa bo'sh")
    return await agents.create(body.text)


@app.post("/api/tasks/{tid}/approve")
def approve(tid: str, body: ApproveIn):
    if not agents.approve(tid, body.ok):
        raise HTTPException(404, "Bu vazifa ruxsat kutmayapti")
    return {"ok": True}


@app.get("/api/files/{name}")
def file(name: str):
    path = (agents.OUT / name).resolve()
    if path.parent != agents.OUT.resolve() or not path.is_file() or path.name == "tasks.json":
        raise HTTPException(404, "Fayl topilmadi")
    return FileResponse(path, filename=path.name)


@app.websocket("/ws")
async def ws(sock: WebSocket):
    await sock.accept()
    q = bus.subscribe()
    try:
        await sock.send_json({"type": "snapshot", "tasks": agents.public()})
        while True:
            ev = await q.get()
            await sock.send_json(ev)
    except (WebSocketDisconnect, RuntimeError, asyncio.CancelledError):
        pass
    finally:
        bus.unsubscribe(q)
