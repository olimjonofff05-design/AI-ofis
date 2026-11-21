"""Hodisalar oqimi: agentlar holatini ofis sahifasiga WebSocket orqali yuboradi."""
import asyncio
import time


class Bus:
    def __init__(self):
        self.clients: set[asyncio.Queue] = set()

    def subscribe(self) -> asyncio.Queue:
        q: asyncio.Queue = asyncio.Queue(maxsize=500)
        self.clients.add(q)
        return q

    def unsubscribe(self, q: asyncio.Queue) -> None:
        self.clients.discard(q)

    async def emit(self, event: dict) -> None:
        event = {**event, "ts": time.time()}
        for q in list(self.clients):
            if q.full():
                continue
            q.put_nowait(event)


bus = Bus()
