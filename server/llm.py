"""Claude bilan ishlash. API kalit bo'lmasa, oflayn rejim ishlaydi."""
import json
import os
import re

DEFAULT_MODEL = "claude-sonnet-5-5"


def parse_json(text: str) -> dict:
    m = re.search(r"```(?:json)?\s*(.*?)```", text, re.S)
    if m:
        text = m.group(1)
    i, j = text.find("{"), text.rfind("}")
    if i < 0 or j <= i:
        raise ValueError("AI javobida JSON topilmadi")
    return json.loads(text[i : j + 1])


class LLM:
    def __init__(self):
        self.key = os.getenv("ANTHROPIC_API_KEY", "").strip()
        self.model = os.getenv("ANTHROPIC_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL
        self.enabled = self.key.startswith("sk-") and "..." not in self.key
        self.client = None
        if self.enabled:
            from anthropic import AsyncAnthropic

            self.client = AsyncAnthropic(api_key=self.key)

    async def json(self, system: str, prompt: str, max_tokens: int = 6000) -> dict:
        msg = await self.client.messages.create(
            model=self.model,
            max_tokens=max_tokens,
            system=system,
            messages=[{"role": "user", "content": prompt}],
        )
        text = "".join(getattr(b, "text", "") for b in msg.content)
        return parse_json(text)


_llm: LLM | None = None


def get_llm() -> LLM:
    global _llm
    if _llm is None:
        _llm = LLM()
    return _llm
