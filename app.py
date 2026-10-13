import io
import re
import cv2
import numpy as np
import streamlit as st
import pytesseract
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

st.set_page_config(page_title="Сверочный лист → слайды", layout="wide")
st.title("Сверочный лист → презентация")
st.caption("Фото таблицы → проверка строк → готовый PPTX")

# ---------- Выравнивание наклона ----------

def deskew(pil_img: Image.Image) -> Image.Image:
    img = cv2.cvtColor(np.array(pil_img.convert("RGB")), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150)
    lines_p = cv2.HoughLinesP(edges, 1, np.pi / 360, 150,
                              minLineLength=gray.shape[1] // 4, maxLineGap=50)
    angles = []
    if lines_p is not None:
        for x1, y1, x2, y2 in lines_p.reshape(-1, 4):
            if abs(x2 - x1) > abs(y2 - y1):
                angles.append(np.degrees(np.arctan2(y2 - y1, x2 - x1)))
    angle = float(np.median(angles)) if angles else 0.0
    h, w = gray.shape
    M = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    fixed = cv2.warpAffine(img, M, (w, h),
                           flags=cv2.INTER_CUBIC, borderValue=(255, 255, 255))
    return Image.fromarray(cv2.cvtColor(fixed, cv2.COLOR_BGR2GRAY))

# ---------- Поиск сетки таблицы ----------

def find_grid_cells(pil_img: Image.Image):
    """Находит линии сетки и возвращает ячейки построчно или None."""
    img = np.array(pil_img.convert("L"))
    h, w = img.shape
    thr = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    kx = max(30, w // 12)
    ky = max(20, h // 20)
    horiz = cv2.morphologyEx(thr, cv2.MORPH_OPEN,
                             cv2.getStructuringElement(cv2.MORPH_RECT, (kx, 1)))
    vert = cv2.morphologyEx(thr, cv2.MORPH_OPEN,
                            cv2.getStructuringElement(cv2.MORPH_RECT, (1, ky)))
    vproj = (vert > 0).sum(axis=0)
    xs = [x for x in range(w) if vproj[x] > h * 0.25]
    vpos = []
    for x in xs:
        if not vpos or x - vpos[-1][-1] > 10:
            vpos.append([x])
        else:
            vpos[-1].append(x)
    vpos = [int(np.mean(g)) for g in vpos]
    hproj = (horiz > 0).sum(axis=1)
    ys = [y for y in range(h) if hproj[y] > w * 0.5]
    hpos = []
    for y in ys:
        if not hpos or y - hpos[-1][-1] > 10:
            hpos.append([y])
        else:
            hpos[-1].append(y)
    hpos = [int(np.mean(g)) for g in hpos]
    if len(vpos) < 4 or len(hpos) < 3:
        return None
    cells = []
    for r in range(len(hpos) - 1):
        y1, y2 = hpos[r], hpos[r + 1]
        if y2 - y1 < 18:
            continue
        row = []
        for c in range(len(vpos) - 1):
            x1, x2 = vpos[c], vpos[c + 1]
            if x2 - x1 < 15:
                continue
            row.append((x1 + 2, y1 + 2, x2 - 2, y2 - 2))
        cells.append(row)
    return cells if cells else None

def cell_ocr(pil_img: Image.Image, box) -> str:
    x1, y1, x2, y2 = box
    crop = pil_img.crop((x1, y1, x2, y2))
    crop = crop.resize((crop.width * 3, crop.height * 3))
    txt = pytesseract.image_to_string(crop, lang="rus", config="--psm 7").strip()
    if not txt:
        txt = pytesseract.image_to_string(crop, lang="rus", config="--psm 6").strip()
    return re.sub(r"[|_]+", "", txt).strip()

# ---------- Классификация колонок ----------

place_re = re.compile(
    r'^\s*(?:[МмMмГгРрP№TтТ]\s*[-–—.]?\s*\d{1,2}\s*[-–—.]?\s*\d{1,2}\s*[-–—.]?\s*\d{1,2}\s*[-–—.]?'
    r'(?:\d{1,4}|[А-Яа-яЁёA-Za-z]{1,2})'
    r'|\d{1,2}\s*[-–—.]\s*\d{1,2}\s*[-–—.]\s*\d{1,2}\s*[-–—.]\s*(?:\d{1,4}|[А-Яа-яЁё]{1,2}))\s*$'
)
lot_re = re.compile(r'^\s*[АВСДЕFавесдa-fA-F]\s*$')
qty_re = re.compile(r'^\s*\d{1,3}\s*$')
series_re = re.compile(r'^\s*([А-ЯЁA-Zа-яёa-z]{0,3}\d{3,}[А-ЯЁA-Zа-яё]?[./]\d{1,3}[\w./-]*)\s*$')
rest_re = re.compile(r"\d+\s*\*\s*$\d+$\s*\+\s*\d+")

def detect_roles_by_header(table):
    header = [t.lower() for t in table[0]]
    roles = {}
    for j, t in enumerate(header):
        if "место" in t: roles["place"] = j
        elif "кол" in t: roles["qty"] = j
        elif "лот" in t: roles["lot"] = j
        elif "наимен" in t: roles["name"] = j
        elif "сери" in t: roles["series"] = j
        elif "остат" in t: roles["rest"] = j
    need = {"place", "qty", "lot", "name"}
    return roles if need <= roles.keys() else None

def detect_roles_heuristic(table):
    ncols = min(len(r) for r in table)
    body = table[1:]
    stats = []
    for j in range(ncols):
        vals = [r[j] for r in body if j < len(r) and r[j].strip()]
        if not vals:
            stats.append(None)
            continue
        stats.append(dict(
            place=sum(bool(place_re.fullmatch(v)) for v in vals) / len(vals),
            lot=sum(bool(lot_re.fullmatch(v)) for v in vals) / len(vals),
            num=sum(bool(qty_re.fullmatch(v)) for v in vals) / len(vals),
            series=sum(bool(series_re.fullmatch(v)) for v in vals) / len(vals),
            rest=sum(bool(rest_re.search(v)) for v in vals) / len(vals),
            avg_len=sum(len(v) for v in vals) / len(vals),
        ))
    roles = {}
    best = max(range(ncols), key=lambda j: (stats[j] or {}).get("lot", -1))
    if stats[best] and stats[best]["lot"] >= 0.5: roles["lot"] = best
    best = max(range(ncols), key=lambda j: (stats[j] or {}).get("place", -1))
    if stats[best] and stats[best]["place"] >= 0.5: roles["place"] = best
    cands = [j for j in range(ncols) if j not in roles.values()]
    best = max(cands, key=lambda j: (stats[j] or {}).get("series", -1))
    if stats[best] and stats[best]["series"] >= 0.5: roles["series"] = best
    for j in cands:
        if stats[j] and stats[j]["rest"] >= 0.5:
            roles["rest"] = j
            break
    nums = [j for j in range(ncols)
            if j not in roles.values() and stats[j] and stats[j]["num"] >= 0.5]
    if len(nums) >= 2:
        roles["num"], roles["qty"] = nums[0], nums[1]
    elif len(nums) == 1:
        roles["qty"] = nums[0]
    cands = [j for j in range(ncols) if j not in roles.values()]
    if cands:
        best = max(cands, key=lambda j: (stats[j] or {}).get("avg_len", -1))
        if stats[best] and stats[best]["avg_len"] >= 8: roles["name"] = best
    need = {"place", "qty", "lot", "name"}
    return roles if need <= roles.keys() else None

def normalize_place(place: str) -> str:
    place = place.strip()
    if not place:
        return ""
    first = place[0]
    if first in "№МмMm": place = "М" + place[1:]
    elif first in "Гг": place = "Г" + place[1:]
    elif first in "РрPp": place = "Р" + place[1:]
    return place.replace(" ", "")

def assemble_rows(table, roles):
    rows = []
    for row in table[1:]:
        def get(role):
            j = roles.get(role)
            return row[j].strip() if j is not None and j < len(row) else ""
        place = normalize_place(get("place"))
        qty = re.sub(r"\D", "", get("qty"))
        lot = get("lot")[:1].upper()
        name = get("name").strip(" ,;")
        series = get("series").replace(" ", "")
        # из наименования — только первые три слова
        name = " ".join(name.split()[:3])
        valid = bool(place_re.fullmatch(place)) or (qty and lot)
        if name and valid:
            rows.append({"Наименование": name, "Место": place, "Серия": series,
                         "Кол-во": qty, "Лоток": lot})
    return rows

# ---------- Запасной путь: сплошной OCR (если сетка не найдена) ----------

def ocr_lines(pil_img: Image.Image) -> list:
    big = pil_img.resize((pil_img.width * 2, pil_img.height * 2))
    data = pytesseract.image_to_data(big, lang="rus", config="--psm 6",
                                     output_type=pytesseract.Output.DATAFRAME)
    data = data[(data.conf > 0) & data.text.notna()]
    data["text"] = data.text.astype(str).str.strip()
    data["text"] = data["text"].str.replace(r"[|_]+", "", regex=True)
    data = data[data.text != ""]
    lines = []
    for _, g in data.groupby(["block_num", "par_num", "line_num"]):
        g = g.sort_values("left")
        lines.append((int(g.top.min()), " ".join(g.text)))
    lines.sort(key=lambda t: t[0])
    return [t for _, t in lines]

main_re = re.compile(
    r"^\s*(?:(\d{1,2})\s+)?"
    r"([МмMмГгРрP№][-–—.]?\d{1,2}[-–—.]?\d{1,2}[-–—.]?\d{1,2}[-–—.]?(?:\d{1,4}|[А-Яа-яЁёA-Za-z]{1,2}))\s+"
    r"(\d{1,3})\s*"
    r"([АВЕСДA-Fa-fавесд])\b\s*"
    r"(.*)$"
)
nopl_re = re.compile(r"^\s*(\d{1,2})\s+(\d{1,3})\s+([АВЕСДA-Fa-fавесд])\b\s+(.*)$")

def parse_fallback(lines):
    rows = []
    for line in lines:
        tail = rest_re.sub("", line).strip()
        m = main_re.match(tail)
        if not m:
            m2 = nopl_re.match(tail)
            if not m2:
                continue
            _, qty, lot, tail = m2.groups()
            place = ""
        else:
            _, place, qty, lot, tail = m.groups()
        sm = series_re.search(tail)
        series = sm.group(1) if sm else ""
        name = tail[:sm.start()].strip(" ,;") if sm else tail.strip(" ,;")
        name = " ".join(name.split()[:3])
        if name:
            rows.append({"Наименование": name, "Место": place, "Серия": series,
                         "Кол-во": qty, "Лоток": lot.upper()})
    return rows

# ---------- Сборка презентации ----------

def build_pptx(rows) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def add_text(slide, x, y, w, h, text, size, bold=False):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(size)
        p.font.bold = bold

    for r in rows:
        slide = prs.slides.add_slide(blank)
        add_text(slide, 0.7, 0.5, 11, 1.8, r["Наименование"], 40, bold=True)
        add_text(slide, 0.7, 2.5, 12, 0.9, f"Место отбора: {r['Место']}", 32)
        add_text(slide, 0.7, 3.6, 12, 0.9, f"Серия: {r['Серия']}", 28)
        add_text(slide, 0.7, 4.7, 8, 0.9, f"Кол-во: {r['Кол-во']}", 28)
        add_text(slide, 10.8, 4.8, 2, 2, r["Лоток"], 96, bold=True)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ---------- Интерфейс ----------

photo = st.camera_input("Сфотографируйте сверочный лист")
if photo is None:
    photo = st.file_uploader("…или загрузите фото", type=["jpg", "jpeg", "png"])

if photo and "rows" not in st.session_state:
    img = deskew(Image.open(photo))
    cells = find_grid_cells(img)
    rows, raw_text, used_cells = [], [], False

    if cells:
       dbg = np.array(img.convert("RGB"))
    for row in cells:
        for x1, y1, x2, y2 in row:
            cv2.rectangle(dbg, (x1, y1), (x2, y2), (255, 0, 0), 2)
    with st.expander("Как программа видит сетку (для отладки)"):
        st.image(dbg)
        if roles is None:
            table = [[cell_ocr(img, b) for b in cells[0]] for _ in [0]]  # placeholder
            full = [ [cell_ocr(img, b) for b in row] for row in cells[1:] ]
            table = [header] + full
            roles = detect_roles_heuristic(table)
            raw_text = [" | ".join(r) for r in table]
        if roles is not None:
            need = [roles[k] for k in ("place", "qty", "lot", "name", "series") if k in roles]
            bar = st.progress(0.0)
            for i, row_cells in enumerate(cells[1:], 1):
                vals = {j: cell_ocr(img, row_cells[j])
                        for j in need if j < len(row_cells)}
                table_row = [vals.get(j, "") for j in range(max(need) + 1)]
                table.append(table_row)
                bar.progress(i / max(1, len(cells) - 1))
            rows = assemble_rows(table, roles)
    if not rows:
        lines = ocr_lines(img)
        raw_text = lines
        rows = parse_fallback(lines)

    if not rows:
        st.error("Строки не распознаны. Снимите лист ровнее, сверху и при хорошем свете.")
        with st.expander("Что увидел OCR (для отладки)"):
            st.text("\n".join(raw_text) or "(пусто)")
    else:
        st.session_state.rows = rows
        msg = f"Распознано строк: {len(rows)}"
        st.success(msg + (" · распознавание по ячейкам таблицы" if used_cells
                          else " · по сплошному тексту (сетка не найдена)"))
        with st.expander("Что увидел OCR (для отладки)"):
            st.text("\n".join(raw_text) or "(пусто)")

if "rows" in st.session_state:
    st.subheader("Проверьте строки — правьте прямо в таблице")
    edited = st.data_editor(st.session_state.rows,
                            num_rows="dynamic", use_container_width=True)
    st.session_state.rows = edited

    if st.button("📄 Собрать презентацию", type="primary"):
        pptx = build_pptx(edited)
        st.download_button("Скачать slides_otbor.pptx", pptx,
                           file_name="slides_otbor.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    if st.button("Загрузить другое фото (сброс)"):
        del st.session_state.rows
        st.rerun()
