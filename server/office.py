"""Hujjatchi asboblari: AI yozgan rejadan (JSON) haqiqiy .pptx, .docx, .xlsx fayl yasaydi."""
import re
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.util import Emu, Inches, Pt as PPt

# ---------- format ----------

def detect_format(text: str) -> str:
    t = text.lower()
    if re.search(r"prezent|presen|slayd|slide|pptx|taqdimot", t):
        return "pptx"
    if re.search(r"excel|xlsx|jadval|byudjet|hisob-kitob|ro'yxat tuz", t):
        return "xlsx"
    return "docx"


def requested_slides(text: str) -> int | None:
    m = re.search(r"(\d+)\s*(?:ta\s*)?(?:slayd|slide)", text.lower())
    return int(m.group(1)) if m else None


def slug(text: str) -> str:
    s = re.sub(r"[^a-z0-9]+", "-", text.lower().replace("'", ""))[:40].strip("-")
    return s or "hujjat"


# ---------- AI ko'rsatmalari ----------

COMMON = (
    "Sen AI Ofisdagi Hujjatchisan. Foydalanuvchi vazifasini bajarib, FAQAT bitta JSON obyekt qaytar, "
    "boshqa hech narsa yozma. Matnni foydalanuvchi yozgan tilda yoz (odatda o'zbek, lotin yozuvi). "
    "Faqat ishonchli faktlarni yoz; aniq bilmagan raqamni o'ylab topma. Mavzu berilmagan bo'lsa, qiziqarli mavzuni o'zing tanla."
)
PROMPTS = {
    "pptx": COMMON
    + ' JSON tuzilishi: {"title": str, "palette": {"bg": "6 xonali hex", "fg": "hex", "accent": "hex"}, '
    '"slides": [{"kind": "title"|"content", "title": str, "subtitle": str (faqat title uchun), '
    '"bullets": [str] (content uchun 3-5 ta, har biri 12 so\'zdan qisqa), '
    '"stat": {"value": str, "label": str} yoki null, "notes": str}]}. '
    "Birinchi slayd kind=title. Slaydlar soni foydalanuvchi so'raganiga teng bo'lsin (sarlavha slaydi ham hisobga kiradi); "
    "so'ramagan bo'lsa 6 ta. Ranglar mavzuga mos va zamonaviy, bg to'q rang bo'lsin.",
    "docx": COMMON
    + ' JSON tuzilishi: {"title": str, "sections": [{"heading": str, "paragraphs": [str], "bullets": [str]}]}. '
    "Rasmiy hujjat (ariza, xat, buyruq) so'ralsa, uning odatiy tuzilishiga amal qil.",
    "xlsx": COMMON
    + ' JSON tuzilishi: {"title": str, "sheet": str, "columns": [str], "rows": [[qiymat]], "totals": true|false}. '
    "Raqamli ustunlarga son yoz (matn emas). Namuna ma'lumot bo'lsa, sarlavhada 'namuna' deb belgilab qo'y.",
}


def prompt(text: str, feedback: str | None) -> str:
    p = f"Vazifa: {text}"
    if feedback:
        p += f"\n\nTekshiruvchi oldingi variantni qaytardi. Sababi: {feedback}. Shuni tuzatib qayta yoz."
    return p


# ---------- oflayn (API kalitsiz) ----------

def offline_spec(fmt: str, text: str) -> dict:
    note = "Oflayn rejim: AI ulanmagan, bu shablon. .env fayliga API kalit qo'shsangiz, mazmunni AI yozadi."
    if fmt == "pptx":
        n = requested_slides(text) or 4
        slides = [{"kind": "title", "title": text[:60], "subtitle": "AI Ofis · Hujjatchi", "notes": note}]
        for i in range(1, n):
            slides.append({"kind": "content", "title": f"{i}-bo'lim", "bullets": ["Asosiy fikr", "Muhim dalil", "Xulosa"],
                           "stat": {"value": f"{i * 25}%", "label": "namuna raqam"} if i % 2 else None, "notes": note})
        return {"title": text[:60], "palette": {"bg": "0E2A3B", "fg": "FFFFFF", "accent": "2DE2E6"}, "slides": slides}
    if fmt == "xlsx":
        return {"title": f"{text[:60]} (namuna)", "sheet": "Jadval", "columns": ["Nomi", "Miqdor", "Narx"],
                "rows": [["Birinchi", 2, 15000], ["Ikkinchi", 5, 8000], ["Uchinchi", 1, 42000]], "totals": True}
    return {"title": text[:80], "sections": [{"heading": "Eslatma", "paragraphs": [note], "bullets": []},
                                             {"heading": "Vazifa", "paragraphs": [text], "bullets": []}]}


# ---------- PPTX ----------

def _rgb(h: str, default: str) -> PRGB:
    h = (h or default).lstrip("#")
    return PRGB.from_string(h if re.fullmatch(r"[0-9A-Fa-f]{6}", h) else default)


def _text(slide, x, y, w, h, text, size, color, bold=False, name=None, italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size, r.font.bold, r.font.italic = PPt(size), bold, italic
    r.font.color.rgb = color
    return tb


P_NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
TRANSITION = (
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' + P_NS + '>'
    '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
    'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p159">'
    '<p:transition spd="slow" p14:dur="1400"><p159:morph option="byObject"/></p:transition></mc:Choice>'
    '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>'
)


def _timing(ids: list[int]) -> str:
    n = [2]

    def nid():
        n[0] += 1
        return n[0]

    effs = []
    for k, sp in enumerate(ids):
        c = nid()
        effs.append(
            f'<p:par><p:cTn id="{c}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" '
            f'nodeType="{"afterEffect" if k == 0 else "withEffect"}"><p:stCondLst><p:cond delay="{k * 350}"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{sp}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{nid()}" dur="600"/>'
            f'<p:tgtEl><p:spTgt spid="{sp}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>'
        )
    g, h = nid(), nid()
    bld = "".join(f'<p:bldP spid="{s}" grpId="0" animBg="1"/>' for s in ids)
    return (
        f'<p:timing {P_NS}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        f'<p:par><p:cTn id="{g}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="{h}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{"".join(effs)}</p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
        f'</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>{bld}</p:bldLst></p:timing>'
    )


def render_pptx(spec: dict, path: Path) -> None:
    pal = spec.get("palette") or {}
    bg, fg, acc = _rgb(pal.get("bg"), "0E2A3B"), _rgb(pal.get("fg"), "FFFFFF"), _rgb(pal.get("accent"), "2DE2E6")
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    W, H = prs.slide_width, prs.slide_height
    slides = spec.get("slides") or []
    for i, s in enumerate(slides):
        sl = prs.slides.add_slide(blank)
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = bg
        anim = []
        orb_pos = [(Inches(9.6), Inches(1.2), Inches(3.2)), (Inches(11.6), Inches(0.5), Inches(1.1)), (Inches(0.6), Inches(6.2), Inches(0.7))]
        ox, oy, od = orb_pos[0] if s.get("kind") == "title" else orb_pos[1 + i % 2]
        orb = sl.shapes.add_shape(MSO_SHAPE.OVAL, ox, oy, od, od)
        orb.name = "!!orb"
        orb.fill.solid()
        orb.fill.fore_color.rgb = acc
        orb.line.fill.background()
        if s.get("kind") == "title":
            t = _text(sl, Inches(0.9), Inches(2.4), Inches(8.4), Inches(1.6), s.get("title", ""), 48, fg, True, "!!title")
            st = _text(sl, Inches(0.9), Inches(4.1), Inches(8.4), Inches(1.0), s.get("subtitle", ""), 22, acc)
            anim += [t.shape_id, st.shape_id]
        else:
            t = _text(sl, Inches(0.9), Inches(0.6), Inches(10.4), Inches(1.0), s.get("title", ""), 36, fg, True, "!!title")
            anim.append(t.shape_id)
            stat = s.get("stat")
            bw = Inches(7.4) if stat else Inches(11.4)
            body = sl.shapes.add_textbox(Inches(0.9), Inches(1.9), bw, Inches(4.8))
            tf = body.text_frame
            tf.word_wrap = True
            for k, b in enumerate(s.get("bullets") or []):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.space_after = PPt(14)
                r = p.add_run()
                r.text = f"•  {b}"
                r.font.size = PPt(22)
                r.font.color.rgb = fg
            anim.append(body.shape_id)
            if stat:
                card = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.7), Inches(2.0), Inches(3.9), Inches(2.9))
                card.fill.solid()
                card.fill.fore_color.rgb = acc
                card.line.fill.background()
                tf = card.text_frame
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                r = p.add_run()
                r.text = str(stat.get("value", ""))
                r.font.size, r.font.bold, r.font.color.rgb = PPt(48), True, bg
                p2 = tf.add_paragraph()
                p2.alignment = PP_ALIGN.CENTER
                r2 = p2.add_run()
                r2.text = str(stat.get("label", ""))
                r2.font.size, r2.font.color.rgb = PPt(16), bg
                anim.append(card.shape_id)
        if s.get("notes"):
            sl.notes_slide.notes_text_frame.text = s["notes"]
        el = sl._element
        el.append(parse_xml(TRANSITION))
        el.append(parse_xml(_timing(anim)))
    prs.save(path)


# ---------- DOCX ----------

def render_docx(spec: dict, path: Path) -> None:
    d = Document()
    st = d.styles["Normal"]
    st.font.name, st.font.size = "Calibri", Pt(12)
    d.add_heading(spec.get("title", "Hujjat"), 0)
    for sec in spec.get("sections") or []:
        if sec.get("heading"):
            d.add_heading(sec["heading"], 1)
        for p in sec.get("paragraphs") or []:
            d.add_paragraph(p)
        for b in sec.get("bullets") or []:
            d.add_paragraph(b, style="List Bullet")
    d.save(path)


# ---------- XLSX ----------

def render_xlsx(spec: dict, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = (spec.get("sheet") or "Jadval")[:31]
    cols = spec.get("columns") or []
    rows = spec.get("rows") or []
    ws.append([spec.get("title", "")])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=2, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2B2622")
        cell.alignment = Alignment(horizontal="center")
    for r in rows:
        ws.append(r)
    last = 2 + len(rows)
    if spec.get("totals") and rows:
        ws.append(["Jami"])
        ws.cell(row=last + 1, column=1).font = Font(bold=True)
        for c in range(2, len(cols) + 1):
            if all(isinstance(r[c - 1], (int, float)) for r in rows if len(r) >= c):
                L = get_column_letter(c)
                ws.cell(row=last + 1, column=c, value=f"=SUM({L}3:{L}{last})").font = Font(bold=True)
    for c in range(1, len(cols) + 1):
        width = max([len(str(cols[c - 1]))] + [len(str(r[c - 1])) for r in rows if len(r) >= c]) + 4
        ws.column_dimensions[get_column_letter(c)].width = min(width, 50)
    ws.freeze_panes = "A3"
    wb.save(path)


RENDER = {"pptx": render_pptx, "docx": render_docx, "xlsx": render_xlsx}


def render(fmt: str, spec: dict, path: Path) -> None:
    RENDER[fmt](spec, path)


# ---------- Tekshiruvchi ----------

def check(path: Path, text: str) -> tuple[bool, str]:
    try:
        ext = path.suffix.lower()
        if ext == ".pptx":
            n = len(Presentation(path).slides)
            want = requested_slides(text)
            if n == 0:
                return False, "prezentatsiyada slayd yo'q"
            if want and n != want:
                return False, f"{want} ta slayd so'ralgan, lekin {n} ta chiqdi"
            return True, f"{n} ta slayd, fayl ochiladi"
        if ext == ".docx":
            paras = [p for p in Document(path).paragraphs if p.text.strip()]
            if len(paras) < 3:
                return False, "hujjat juda qisqa"
            return True, f"{len(paras)} ta xatboshi, fayl ochiladi"
        if ext == ".xlsx":
            ws = load_workbook(path).active
            if ws.max_row < 3:
                return False, "jadvalda ma'lumot yo'q"
            return True, f"{ws.max_row - 2} ta qator, fayl ochiladi"
        return False, "noma'lum format"
    except Exception as e:  # fayl ochilmasa
        return False, f"fayl ochilmadi: {e}"
