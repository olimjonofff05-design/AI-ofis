"""Agentlar: Rais vazifani yo'naltiradi, Hujjatchi bajaradi, Tekshiruvchi tekshiradi."""
import asyncio
import json
import re
import time
import uuid
from pathlib import Path

from . import office
from .bus import bus
from .llm import get_llm

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs"
OUT.mkdir(exist_ok=True)
DB = OUT / "tasks.json"

NAMES = {"rais": "Rais", "hujjatchi": "Hujjatchi", "tadqiqotchi": "Tadqiqotchi", "rassom": "Rassom",
         "montajchi": "Montajchi", "jarvis": "Jarvis", "qoriqchi": "Qo'riqchi", "tekshiruvchi": "Tekshiruvchi"}
CONNECTED = {"rais", "hujjatchi", "tekshiruvchi"}
STAGE = {"tadqiqotchi": 2, "qoriqchi": 2, "jarvis": 3, "rassom": 4, "montajchi": 4}
KEYWORDS = [
    ("montajchi", r"video|montaj|reels|klip|subtitr|rolik|shorts"),
    ("rassom", r"rasm|logo|dizayn|poster|banner|foto|surat"),
    ("qoriqchi", r"github|repo|xavfsiz|kalit|token|zaif|parol"),
    ("jarvis", r"kompyuter|papka|ilova|brauzer|\boch\b|tartib|o'chir|skrinshot"),
    ("hujjatchi", r"prezent|presen|slayd|slide|word|doc|hujjat|matn|xat|excel|jadval|hisobot|pdf|tarjima|ariza|pptx|xlsx"),
    ("tadqiqotchi", r"izla|qidir|tadqiq|tahlil|narx|raqobat|bozor|trend"),
]
DANGER = re.compile(r"o'chir|yubor|nashr|to'la|push|delete", re.I)
RAIS_SYS = (
    "Sen AI Ofisning Raisisan. Foydalanuvchi vazifasini bitta agentga topshirasan. Agentlar: "
    "hujjatchi (prezentatsiya, Word, Excel, matn, xat, hisobot, tarjima), tadqiqotchi (internetdan qidiruv, tahlil), "
    "rassom (rasm, logo, poster), montajchi (video), jarvis (kompyuterni boshqarish, fayl va papkalar), "
    "qoriqchi (GitHub va xavfsizlik). FAQAT JSON qaytar: "
    '{"agent": "nomi", "reason": "o\'zbekcha, 10 so\'zgacha sabab"}'
)

TASKS: dict[str, dict] = {}
APPROVALS: dict[str, asyncio.Future] = {}
LOCKS = {a: asyncio.Lock() for a in NAMES}


def now_ms() -> int:
    return int(time.time() * 1000)


def load() -> None:
    try:
        for t in json.loads(DB.read_text(encoding="utf-8")):
            if t.get("st") not in ("tayyor", "bekor"):
                t.update(st="bekor", note="Server qayta ishga tushdi, vazifa to'xtadi")
            TASKS[t["id"]] = t
    except (FileNotFoundError, ValueError):
        pass


def save() -> None:
    data = sorted(TASKS.values(), key=lambda t: t["t0"], reverse=True)[:100]
    DB.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")


def public() -> list[dict]:
    return sorted(TASKS.values(), key=lambda t: t["t0"], reverse=True)


async def update(t: dict, **kw) -> None:
    t.update(kw)
    save()
    await bus.emit({"type": "task", "task": t})


async def create(text: str) -> dict:
    t = {"id": uuid.uuid4().hex[:8], "text": text.strip()[:500], "ag": None, "st": "navbatda", "p": 0,
         "note": "Rais o'ylayapti", "t0": now_ms(), "t1": None, "file": None, "help": None}
    TASKS[t["id"]] = t
    await update(t)
    asyncio.create_task(run(t))
    return t


async def route(text: str) -> tuple[str, str]:
    llm = get_llm()
    if llm.enabled:
        try:
            r = await llm.json(RAIS_SYS, text, 300)
            if r.get("agent") in NAMES and r["agent"] not in ("rais", "tekshiruvchi"):
                return r["agent"], str(r.get("reason", ""))[:120]
        except Exception:
            pass
    for ag, rx in KEYWORDS:
        if re.search(rx, text, re.I):
            return ag, ""
    return "hujjatchi", ""


async def hujjatchi(t: dict, feedback: str | None = None) -> Path:
    fmt = office.detect_format(t["text"])
    llm = get_llm()
    await update(t, p=15, note="Reja tuzyapti")
    if llm.enabled:
        spec = await llm.json(office.PROMPTS[fmt], office.prompt(t["text"], feedback), 8000)
    else:
        await asyncio.sleep(2)
        spec = office.offline_spec(fmt, t["text"])
    await update(t, p=65, note="Faylni yig'yapti")
    path = OUT / f"{office.slug(spec.get('title') or t['text'])}-{t['id']}.{fmt}"
    await asyncio.to_thread(office.render, fmt, spec, path)
    await update(t, p=85, note="")
    return path


HANDLERS = {"hujjatchi": hujjatchi}


async def run(t: dict) -> None:
    await asyncio.sleep(1.2)
    ag, why = await route(t["text"])
    await update(t, ag=ag, note=why)
    if ag not in CONNECTED:
        await asyncio.sleep(1.5)
        await update(t, st="bekor", t1=now_ms(),
                     note=f"{NAMES[ag]} hali ulanmagan: {STAGE.get(ag, 2)}-bosqichda ishga tushadi")
        return
    async with LOCKS[ag]:
        if DANGER.search(t["text"]):
            fut = asyncio.get_running_loop().create_future()
            APPROVALS[t["id"]] = fut
            await update(t, st="ruxsat", note="Ruxsatingizni kutyapti")
            ok = await fut
            APPROVALS.pop(t["id"], None)
            if not ok:
                await update(t, st="bekor", t1=now_ms(), note="Siz rad etdingiz, hech narsa o'zgarmadi")
                return
        await update(t, st="bajarilyapti", p=5, note="")
        feedback, ok, reason, path = None, False, "", None
        try:
            for attempt in range(3):
                path = await HANDLERS[ag](t, feedback)
                await update(t, st="tekshiruvda", p=92, note="Tekshiruvchi ko'ryapti")
                await asyncio.sleep(1.5)
                ok, reason = await asyncio.to_thread(office.check, path, t["text"])
                if ok or attempt == 2:
                    break
                feedback = reason
                await update(t, st="bajarilyapti", p=50, note=f"Qayta ishlayapti: {reason}",
                             help={"to": "tekshiruvchi", "q": "Nimasi xato?!", "a": reason, "n": attempt + 1})
        except Exception as e:
            await update(t, st="bekor", t1=now_ms(), note=f"Xato: {str(e)[:160]}")
            return
        await update(t, st="tayyor", p=100, t1=now_ms(), file=path.name,
                     note=reason if ok else f"Ogohlantirish: {reason}")


def approve(tid: str, ok: bool) -> bool:
    fut = APPROVALS.get(tid)
    if fut and not fut.done():
        fut.set_result(ok)
        return True
    return False
