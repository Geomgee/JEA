# PDFtoBOM.py
from __future__ import annotations

import os, math, csv, sqlite3, re
import fitz  # PyMuPDF
import openpyxl
from openpyxl.utils import get_column_letter
from typing import Dict, List, Tuple, Optional

# =========================
# Config / Constants
# =========================
DEBUG_DUMP_ALL = False     # True -> 디버그 확인용

#수치 ----------------------------------------------------------
TARGET_DWG_COORD = (46.5, 755.5)  # DWG No. 기준 좌표(포인트 단위 좌표계)

TAG_TOL_MM = 3           # (기존) 수직 인접 판단에 쓰던 기본 mm 값, x0 정렬 허용치(PT)
NEARBY_RADIUS_PT = 10     # (기존) 주변 검색 반경(pt)
MM_TO_PT = 72 / 25.4      # 1mm ≈ 2.83465pt

#주석 부분 style ---------------------------------------------------
COLOR_DEFAULT = (92/255.0, 209/255.0, 229/255.0) # Q'TY == 1
COLOR_MULTI = (1.0, 0.0, 0.0) # Q'TY != 1
border = 1 # 사각형 두께
rec_size = 5.0 # 사각형 사이즈 (rec_size * 2pt)

CENTER_X_TOL_MM = 4       # x는 같거나 ±4mm
CENTER_Y_TOL_MM = 1       # y는 같거나 ±0.5~1mm 정도 → 1mm로 설정 (필요시 0.5로 낮추세요)

DRAW_COMBINED_MODE = "unionY"

# 접미어 리스트 : PART NAME
VALID_PREFIXES = ["PSV","PRV","MFM","PG","PT","BV","CV","NV","SV","XV","LF","GD","FD","YS"]
PART_NAMES = {
    "BV":"BALL VALVE","CV":"CHECK VALVE","NV":"NEEDLE VALVE","SV":"SHUT OFF VALVE","XV":"SOLENOID VALVE",
    "PSV":"PRESSURE SAFETY VALVE","PRV":"PRESSURE REGULATOR VALVE","LF":"LINE FILTER","MFM":"MASS FLOW METER",
    "PG":"PRESSURE GAUGE","PT":"PRESSURE TRANSMITTER","GD":"GAS DETECTOR","FD":"FIRE DETECTOR","YS":"스트레이너"
}

_ALLOWED_CHUNK_RE = re.compile(r"^[A-Za-z0-9 ]+$")  # 특문 제외, 공백 허용

# ---- GUI (파일 선택 / 알림) ----
try:
    from tkinter import Tk, filedialog, messagebox
    TK_OK = True
except Exception:
    TK_OK = False

# =========================
# Small utilities
# =========================
def dbg(msg: str) -> None:
    if DEBUG_DUMP_ALL:
        print(f"[DEBUG] {msg}")

def show_alert(msg: str, title="알림"):
    if TK_OK:
        root = Tk(); root.withdraw()
        try:
            messagebox.showinfo(title, msg)
        finally:
            root.destroy()
    else:
        print(f"[{title}] {msg}")

def ask_open_pdf() -> str:
    if not TK_OK:
        return input("입력 PDF 경로를 적어주세요: ").strip()
    root = Tk(); root.withdraw()
    try:
        path = filedialog.askopenfilename(title="PDF 선택", filetypes=[("PDF files","*.pdf")])
    finally:
        root.destroy()
    return path

def ask_save_xlsx(default_dir: str, default_name: str) -> str:
    if not TK_OK:
        return input(f"엑셀 저장 경로(기본 {os.path.join(default_dir,default_name)}): ").strip() \
            or os.path.join(default_dir,default_name)
    root = Tk(); root.withdraw()
    try:
        path = filedialog.asksaveasfilename(
            title="엑셀 저장 위치",
            defaultextension=".xlsx",
            initialdir=default_dir,
            initialfile=default_name,
            filetypes=[("Excel Workbook","*.xlsx")],
        )
    finally:
        root.destroy()
    return path

def _is_simple_chunk(s: str) -> bool:
    return bool(_ALLOWED_CHUNK_RE.match(s or ""))

def _has_digit(s: str) -> bool:
    return any(ch.isdigit() for ch in (s or ""))

def _is_numeric_chunk_only(s: str) -> bool:
    """공백 제거 후 전부 숫자인 경우만 True (예: '001')"""
    return (s or "").replace(" ", "").isdigit()

def _center_of_rect(x0,y0,x1,y1):
    return ((x0+x1)/2.0, (y0+y1)/2.0)

# =========================
# Text / PDF helpers
# =========================
def annot_text(annot) -> str:
    try:
        info = annot.info or {}
    except Exception:
        info = {}
    return (info.get("content") or info.get("subject") or "").strip()

def get_annots_list(page) -> List:
    it = page.annots()
    return list(it) if it else []

# ---- DWG No. 근접 추정 (현 방식 유지) ----
def nearest_dwg_from_annots(annots, target: Tuple[float, float]) -> Optional[str]:
    best_txt, best_d = None, float("inf")
    for a in annots:
        txt = annot_text(a)
        if not txt:
            continue
        if "-" in txt and "FROM." not in txt and "TO." not in txt:
            r = a.rect
            cx, cy = (r.x0 + r.x1)/2, (r.y0 + r.y1)/2
            d = math.hypot(cx - target[0], cy - target[1])
            if d < best_d:
                best_d, best_txt = d, txt
    return best_txt

def nearest_dwg_from_words(page, target: Tuple[float, float]) -> Optional[str]:
    words = page.get_text("words")  # (x0,y0,x1,y1,word,block,line,word_no)
    if not words:
        return None
    best_txt, best_d = None, float("inf")
    for x0,y0,x1,y1,w,*_ in words:
        w = (w or "").strip()
        if "-" in w and "FROM." not in w and "TO." not in w:
            cx, cy = (x0+x1)/2, (y0+y1)/2
            d = math.hypot(cx - target[0], cy - target[1])
            if d < best_d:
                best_d, best_txt = d, w
    return best_txt

# =========================
# Tag validation / joining
# =========================
_HANGUL_RANGE = ('\uac00', '\ud7a3')
PREFIX_ALT = "|".join(map(re.escape, VALID_PREFIXES))
TAG_RE = re.compile(rf'^(?:{PREFIX_ALT})[-\sA-Za-z0-9_]*\d[-\sA-Za-z0-9_]*$')

def _has_hangul(s: str) -> bool:
    return any(_HANGUL_RANGE[0] <= ch <= _HANGUL_RANGE[1] for ch in s)

def _prefix_of(tag: str) -> Optional[str]:
    return next((p for p in VALID_PREFIXES if tag.startswith(p)), None)

def is_valid_tag(tag: str) -> bool:
    # 정상 모드에서만 강한 검증. 디버그 모드면 True로 통과
    if DEBUG_DUMP_ALL:
        return True
    if not tag or '/' in tag or '"' in tag:
        return False
    pf = _prefix_of(tag)
    if not pf:
        return False
    rest = tag[len(pf):].strip()
    if not rest or _has_hangul(rest):
        return False
    return bool(TAG_RE.match(tag))

# =========================
# Collectors (통합)
# =========================
def _iter_objects(source, from_annots: bool):
    """
    공통 이터레이터:
      - annots: (index, text, rect(x0,y0,x1,y1))
      - words:  (index, text, rect(x0,y0,x1,y1))
    """
    if from_annots:
        for i, a in enumerate(source):
            yield i, annot_text(a).strip(), (a.rect.x0, a.rect.y0, a.rect.x1, a.rect.y1)
    else:
        for i, o in enumerate(source):
            yield i, (o["w"] or "").strip(), (o["x0"], o["y0"], o["x1"], o["y1"])

def _objects_for_words(page):
    words = page.get_text("words")
    if not words:
        return []
    return [{"x0":x0,"y0":y0,"x1":x1,"y1":y1,"w":(w or "").strip()}
            for x0,y0,x1,y1,w,*_ in words if (w or "").strip()]

def _starts_with_valid_prefix(s: str) -> bool:
    """다른 접두사로 시작한다면 붙이지 않기"""
    s = (s or "").strip()
    return any(s.startswith(p) for p in VALID_PREFIXES)

# center 좌표로 접두어-접미어 결합
def _attach_by_center_if_prefix(objects, idx, text, rect,
                                x_tol_mm: float, y_tol_mm: float,
                                from_annots: bool):
    """
    접두어만 있을 때, 'center 거리'로 접미어(숫자 포함, 특수문자 제외, 공백 허용)를 결합.
    - 조건:
      * 후보 텍스트는 다른 접두사로 시작하지 않아야 함
      * 특수문자 없음(영문/숫자/공백만), 공백은 출력 시 제거
      * 공백 제거 후 숫자 1개 이상 포함
      * 중심점 거리: |Δx| ≤ x_tol_mm, |Δy| ≤ y_tol_mm (mm 기준)
    - 반환: (결합된 텍스트, 붙인 인덱스 or None, 붙인 원본 텍스트 or None)
    - 주의: 붙인 인덱스는 used로 마킹하지 말 것(공용 사용 허용)
    """
    if text not in VALID_PREFIXES or not objects:
        return text, None, None, None

    x_tol_pt = x_tol_mm * MM_TO_PT
    y_tol_pt = y_tol_mm * MM_TO_PT

    pcx, pcy = _center_of_rect(*rect)
    best = None  # (distance, j, raw_t2)

    for j, t2, r2 in _iter_objects(objects, from_annots):
        if j == idx:
            continue
        t2s = (t2 or "").strip()
        if _starts_with_valid_prefix(t2s):
            continue
        if not _is_simple_chunk(t2s):
            continue
        if not any(ch.isdigit() for ch in t2s.replace(" ", "")):
            continue

        scx, scy = _center_of_rect(*r2)
        if abs(scx - pcx) <= x_tol_pt and abs(scy - pcy) <= y_tol_pt:
            d = math.hypot(scx - pcx, scy - pcy)
            if best is None or d < best[0]:
                best = (d, j, t2s, r2)

    if best:
        _, j, t2s, r2 = best
        return text + t2s.replace(" ", ""), j, t2s, r2

    return text, None, None, None

# (기존) 아래/주변 결합 로직은 폴백으로 유지 가능
def _attach_nearby_numeric(objects, idx, base_rect, radius_pt: float, from_annots: bool) -> Optional[Tuple[str, int, Tuple[float,float,float,float]]]:
    if not objects:
        return None

    bx0, by0, bx1, by1 = base_rect
    bcx, bcy = _center_of_rect(bx0,by0,bx1,by1)
    tol_pt = TAG_TOL_MM * MM_TO_PT

    def ok_candidate(t0: str) -> bool:
        if _starts_with_valid_prefix(t0):
            return False
        if not _is_simple_chunk(t0):
            return False
        return any(ch.isdigit() for ch in (t0 or "").replace(" ", ""))

    #1) 같은 열 & 아래쪽 우선
    best1 = None
    for j, t, r in _iter_objects(objects, from_annots):
        if j == idx:
            continue
        t0 = (t or "").strip()
        if not ok_candidate(t0):
            continue
        if abs(r[0] - bx0) <= TAG_TOL_MM and 0 < (r[1] - by1) <= tol_pt:
            d = r[1] - by1
            if best1 is None or d < best1[0]:
                best1 = (d, j, t0)
    if best1:
        _, j, t0 = best1
        return t0.replace(" ", ""), j

    #2) 반경 내 최단 거리
    best2 = None
    for j, t, r in _iter_objects(objects, from_annots):
        if j == idx:
            continue
        t0 = (t or "").strip()
        if not ok_candidate(t0):
            continue
        cx, cy = _center_of_rect(*r)
        d = math.hypot(cx - bcx, cy - bcy)
        if d <= radius_pt and (best2 is None or d < best2[0]):
            best2 = (d, j, t0)
    if best2:
        _, j, t0 = best2
        return t0.replace(" ", ""), j

    return None

def collect_tags_generic(source, dwg_no: str, from_annots: bool,
                         tag_tol_mm: float = TAG_TOL_MM, page=None,
                         hits: Optional[List[Tuple[int, Tuple[float,float,float,float], str, str]]] = None,
                         page_index: Optional[int] = None) -> Dict[str, int]:
    """
    주석(annots) 또는 단어(words)에서 태그를 수집하는 통합 수집기.
    1) 접두사 단독 발견 → **center 결합**(_attach_by_center_if_prefix)
    2) 여전히 '문자만'이면 → (폴백) 기존 주변 결합(_attach_nearby_numeric)
    3) 완성된 문자열이 is_valid_tag 통과 시 카운트
    """
    tol = tag_tol_mm * MM_TO_PT
    counts: Dict[str, int] = {}
    used = set()

    objects = source if from_annots else _objects_for_words(source)

    for i, txt, rect in _iter_objects(source if from_annots else objects, from_annots):
        if i in used or not txt:
            continue

        draw_rect = rect  # 기본: 자기 자신의 위치에 그림
        suffix_rect = None

        # 1) center 결합 (x±5mm, y±1mm)
        if txt in VALID_PREFIXES:
            new_txt, joined_idx, _joined_raw, joined_rect = _attach_by_center_if_prefix(
                objects if not from_annots else source,
                i, txt, rect,
                CENTER_X_TOL_MM, CENTER_Y_TOL_MM,
                from_annots
            )
            txt = new_txt
            suffix_rect = joined_rect
            # suffix 인덱스는 공용 사용 허용 → used로 마킹하지 않음
        """
    
        # 2) (폴백) 여전히 숫자 없음 → 기존 주변 결합
        pf = _prefix_of(txt)
        if pf:
            rest = txt[len(pf):]
            if rest and not any(ch.isdigit() for ch in rest):
                extra = _attach_nearby_numeric(
                    objects if not from_annots else source,
                    i, rect, NEARBY_RADIUS_PT, from_annots
                )
                if extra:
                    extra_txt, _extra_idx = extra
                    txt = txt + extra_txt
                    suffix_rect = extra_rect
        """

        # 2.5) 결합이 있었다면 평균점에 1개만 그림
        if suffix_rect is not None:
            pcx, pcy = _center_of_rect(*rect)
            scx, scy = _center_of_rect(*suffix_rect)

            # 안전하게 한 번 더 허용 오차 검증
            if (abs(scx - pcx) <= CENTER_X_TOL_MM * MM_TO_PT and
                    abs(scy - pcy) <= CENTER_Y_TOL_MM * MM_TO_PT):
                mx, my = (pcx + scx) / 2.0, (pcy + scy) / 2.0
                half = rec_size
                draw_rect = (mx - half, my - half, mx + half, my + half)
                # else: 오차 밖이면 draw_rect는 접두어 rect 그대로 둠

        # 3) 유효성 체크 및 집계
        if is_valid_tag(txt):
            key = f"{dwg_no}|{txt}"
            counts[key] = counts.get(key, 0) + 1
            if hits is not None:
                pid = page.number if page is not None else (page_index if page_index is not None else 0)
                hits.append((pid, draw_rect, txt, dwg_no))
            dbg(f"{'ANNOTS' if from_annots else 'WORDS'} HIT: {key}")
    return counts

# =========================
# Pipeline
# =========================
def extract_annotations(pdf_path: str,
                        target_coord: Tuple[float, float] = TARGET_DWG_COORD,
                        tag_tol_mm: float = TAG_TOL_MM) -> Dict[str, int]:
    doc = fitz.open(pdf_path)
    tag_counts: Dict[str, int] = {}

    for page in doc:
        annots = get_annots_list(page)
        dwg_no = nearest_dwg_from_annots(annots, target_coord) if annots else None
        if not dwg_no:
            dwg_no = nearest_dwg_from_words(page, target_coord)
        if not dwg_no:
            dbg("DWG NO not found on page; skipping")
            continue

        page_counts = collect_tags_generic(annots, dwg_no, from_annots=True, tag_tol_mm=tag_tol_mm, page=page) if annots else {}
        if not page_counts:
            page_counts = collect_tags_generic(page, dwg_no, from_annots=False, tag_tol_mm=tag_tol_mm, page=page)

        for k, v in page_counts.items():
            tag_counts[k] = tag_counts.get(k, 0) + v

    dbg(f"TOTAL TAGS: {len(tag_counts)}")
    return tag_counts

# =========================
# Rows / Output
# =========================
def build_filtered_rows(tag_counts: Dict[str, int]) -> List[Tuple[str, str, str, int]]:
    """정렬된 (DWG, TAG, PART, QTY) 튜플 리스트 생성 (출력 시 공백 제거)"""
    rows: List[Tuple[str, str, str, int]] = []
    for key, cnt in tag_counts.items():
        dwg, tag = key.split("|", 1)
        if not is_valid_tag(tag):
            continue
        pf = _prefix_of(tag)
        part = PART_NAMES.get(pf or "", "")
        # 출력 시 공백 제거
        tag_out = re.sub(r"\s+", "", tag)
        rows.append((dwg, tag_out, part, cnt))
    rows.sort(key=lambda r: (r[0], r[1]))
    return rows

def write_to_excel(tag_counts: Dict[str, int], output_path: str) -> None:
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Annotations"
    headers = ["DWG NO.", "TAG NO.", "PART NAME", "Q'TY"]
    for col, h in enumerate(headers, start=2):
        ws[f"{get_column_letter(col)}2"] = h
    rows = build_filtered_rows(tag_counts)
    for i, (dwg, tag, part, cnt) in enumerate(rows, start=3):
        ws[f"B{i}"] = dwg; ws[f"C{i}"] = tag; ws[f"D{i}"] = part; ws[f"E{i}"] = cnt
    ws.auto_filter.ref = f"B2:E{2 + len(rows)}"

    # ===== 열 너비 자동 맞춤 (openpyxl 방식) =====
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        # 글자수 * 1.2 정도 = 적당한 여유
        ws.column_dimensions[col_letter].width = max_length * 1.2
    wb.save(output_path)


# 주석 pdf 생성
def write_to_annots(pdf_path: str,
                    out_path: Optional[str] = None,
                    target_coord: Tuple[float, float] = TARGET_DWG_COORD,
                    existing_counts: Optional[Dict[str,int]] = None) -> str:
    """
    문서를 스캔해 TAG 좌표(hit)를 수집하고,
    최종 Q'TY로 색상을 결정해 사각형 주석을 추가한 뒤 저장.
    - Q'TY == 1 → COLOR_DEFAULT (하늘색)
    - Q'TY != 1 → COLOR_MULTI   (빨강)
    """
    with fitz.open(pdf_path) as doc:
        hits: List[Tuple[int, Tuple[float,float,float,float], str, str]] = []
        counts: Dict[str,int] = {} if existing_counts is None else dict(existing_counts)

        # 1) 좌표 수집(+필요시 집계)
        for p_idx in range(len(doc)):
            page = doc[p_idx]
            annots = get_annots_list(page)
            dwg_no = nearest_dwg_from_annots(annots, target_coord) if annots else None
            if not dwg_no:
                dwg_no = nearest_dwg_from_words(page, target_coord)
            if not dwg_no:
                continue

            # annots 우선, 없으면 words
            c1 = collect_tags_generic(annots, dwg_no, from_annots=True, page=None, hits=hits, page_index=p_idx) if annots else {}
            if not c1:
                c2 = collect_tags_generic(page, dwg_no, from_annots=False, page=None, hits=hits, page_index=p_idx)
            else:
                c2 = {}

            # 기존 counts가 없을 때만 현 스캔으로 누적
            if existing_counts is None:
                for k, v in {**c1, **c2}.items():
                    counts[k] = counts.get(k, 0) + v

        # 2) 색상 결정 후 실제 주석 그리기
        for p_idx, rect, tag_text, dwg_no in hits:
            key = f"{dwg_no}|{tag_text}"
            qty = counts.get(key, 0)
            color = COLOR_MULTI if qty != 1 else COLOR_DEFAULT
            page = doc[p_idx]
            a = page.add_rect_annot(fitz.Rect(*rect))
            a.set_colors(stroke=color)
            a.set_border(width=border)
            a.set_info(content=tag_text)
            a.update()

        # 3) 저장
        if out_path is None:
            base_dir = os.path.dirname(pdf_path) or os.getcwd()
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_path = os.path.join(base_dir, f"{base_name}_with_annots.pdf")
        doc.save(out_path)

    return out_path



# =========================
# (Optional outputs) CSV / SQLite
# =========================
def write_to_csv(tag_counts: Dict[str, int], csv_path: str) -> None:
    rows = build_filtered_rows(tag_counts)
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["DWG NO.", "TAG NO.", "PART NAME", "Q'TY"])
        w.writerows(rows)

def init_db(db_path: str):
    folder = os.path.dirname(db_path)
    if folder and not os.path.exists(folder):
        os.makedirs(folder, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
    except Exception:
        pass
    cur = conn.cursor()
    cur.execute("""
                CREATE TABLE IF NOT EXISTS tags (
                                                    dwg_no    TEXT NOT NULL,
                                                    tag_no    TEXT NOT NULL,
                                                    part_name TEXT,
                                                    qty       INTEGER NOT NULL,
                                                    PRIMARY KEY (dwg_no, tag_no)
                    )
                """)
    conn.commit()
    return conn

def upsert_tag_counts(conn, tag_counts: Dict[str, int]) -> None:
    cur = conn.cursor()
    use_upsert = tuple(int(x) for x in sqlite3.sqlite_version.split(".")) >= (3,24,0)
    for key, cnt in tag_counts.items():
        dwg, tag = key.split("|", 1)
        pf = _prefix_of(tag)
        part = PART_NAMES.get(pf or "", "")
        if use_upsert:
            cur.execute("""
                        INSERT INTO tags (dwg_no, tag_no, part_name, qty)
                        VALUES (?, ?, ?, ?)
                            ON CONFLICT(dwg_no, tag_no) DO UPDATE SET
                            qty = tags.qty + excluded.qty,
                                                               part_name = excluded.part_name
                        """, (dwg, tag, part, cnt))
        else:
            cur.execute("SELECT qty FROM tags WHERE dwg_no=? AND tag_no=?", (dwg, tag))
            row = cur.fetchone()
            if row:
                cur.execute(
                    "UPDATE tags SET qty = qty + ?, part_name = ? WHERE dwg_no=? AND tag_no=?",
                    (cnt, part, dwg, tag)
                )
            else:
                cur.execute(
                    "INSERT INTO tags (dwg_no, tag_no, part_name, qty) VALUES (?,?,?,?)",
                    (dwg, tag, part, cnt)
                )
    conn.commit()



# =========================
# Main
# =========================
def main():
    # 1) PDF 선택
    pdf_path = ask_open_pdf()
    if not pdf_path:
        show_alert("작업이 취소되었습니다.")
        return
    if not os.path.isfile(pdf_path):
        show_alert(f"입력 PDF를 찾을 수 없습니다:\n{pdf_path}")
        return

    base_dir = os.path.dirname(pdf_path) or os.getcwd()
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    xlsx_default = f"{base_name}_BOM.xlsx"

    # 2) 엑셀 저장 위치
    xlsx_path = ask_save_xlsx(base_dir, xlsx_default)
    if not xlsx_path:
        show_alert("엑셀 저장 위치가 지정되지 않아 작업을 종료합니다.")
        return

    # 3) 추출 & 저장
    tags = extract_annotations(pdf_path, target_coord=TARGET_DWG_COORD, tag_tol_mm=TAG_TOL_MM)

    # 4) 결과 안내
    if not tags:
        show_alert(
            "추출된 항목이 없습니다 (0건).\n\n"
            "PDF가 SHX 글꼴로 출력되어 텍스트가 벡터로 깨졌을 수 있어요.\n"
            "CAD에서 PDFSHX=1로 설정해 다시 출력 후 재시도해 주세요."
        )
    else:
        write_to_excel(tags, xlsx_path)
        write_to_annots(pdf_path, existing_counts=tags)
        show_alert(f"BOM 추출이 완료되었습니다.\nTAG가 없는 경우, 출력이 되지 않으니 꼭! 확인해주세요.\n\n엑셀 위치: {xlsx_path}", title="완료")

if __name__ == "__main__":
    main()
