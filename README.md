# AI Ofis

8 ta AI agentdan iborat ishchi jamoa va ularning ishini jonli ko'rsatadigan "ofis" ekrani.

| Agent | Vazifasi | Holati |
|---|---|---|
| Rais | Vazifani tushunadi va mos agentga topshiradi | Ishlaydi |
| Hujjatchi | Prezentatsiya (.pptx), Word (.docx), Excel (.xlsx) | Ishlaydi |
| Tekshiruvchi | Faylni ochib tekshiradi, xato bo'lsa qaytaradi (2 martagacha) | Ishlaydi |
| Tadqiqotchi | Internetdan qidiruv va tahlil | 2-bosqich |
| Qo'riqchi | GitHub repolar xavfsizligi | 2-bosqich |
| Jarvis | Kompyuterni ovoz va matn bilan boshqarish | 3-bosqich |
| Rassom | Rasm generatsiyasi va tahriri | 4-bosqich |
| Montajchi | Video montaj, namuna uslubida video | 4-bosqich |

## Ishga tushirish (Windows)

1. [python.org](https://www.python.org/downloads/) dan **Python 3.11+** o'rnating ("Add python.exe to PATH" belgisini qo'ying).
2. Repoda **Code → Download ZIP**, arxivni oching.
3. **`start.bat`** ni ikki marta bosing. Birinchi safar kutubxonalar o'rnatiladi (1–3 daqiqa).
4. Ochilgan bloknotdagi `.env` fayliga Claude API kalitingizni yozing va saqlang:
   `ANTHROPIC_API_KEY=sk-ant-...` (kalit: https://console.anthropic.com → API Keys).
5. Oynani yopib, `start.bat` ni qayta ishga tushiring. Brauzerda **http://localhost:8000** ochiladi.

Kalitsiz ham ishlaydi: ofis "Jonli rejim · AI kalitsiz" deb yozadi va Hujjatchi shablon fayl yasaydi.

## Qanday ishlaydi

```
Siz (ofis sahifasi) → Rais (Claude) → Hujjatchi (Claude + python-pptx/docx/openpyxl)
                    → Tekshiruvchi (faylni ochadi, slaydlar sonini tekshiradi) → outputs/ papka
Har bir qadam WebSocket orqali ofis ekranida jonli ko'rinadi.
```

- Tayyor fayllar `outputs/` papkasida va "Loyihalar" sahifasida.
- "O'chir", "yubor", "nashr" kabi so'zli vazifalar oldin ruxsatingizni so'raydi.
- `office-ui/index.html` ni to'g'ridan-to'g'ri ochsangiz, namuna (simulyatsiya) rejimi ishlaydi.

## Tuzilma

```
server/app.py      FastAPI: /api/tasks, /api/files, /ws
server/agents.py   Rais, Hujjatchi, Tekshiruvchi mantig'i
server/office.py   .pptx/.docx/.xlsx yasash va tekshirish
server/llm.py      Claude API
office-ui/         jonli ofis sahifasi
```
