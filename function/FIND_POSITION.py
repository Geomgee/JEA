import fitz
from tkinter import Tk, filedialog, simpledialog

def find_annotation_positions(pdf_path: str, query: str,
                              contains: bool = True, ignore_case: bool = True):
    if ignore_case:
        norm = lambda s: (s or "").strip().lower()
        q = norm(query)
    else:
        norm = lambda s: (s or "").strip()
        q = norm(query)

    doc = fitz.open(pdf_path)
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        annots = page.annots()
        if not annots:
            continue
        for annot in annots:
            try:
                info = annot.info or {}
            except Exception:
                info = {}
            text = (info.get("content") or info.get("subject") or "").strip()
            t_norm = norm(text)

            if contains and q not in t_norm:
                continue
            if not contains and q != t_norm:
                continue

            rect = annot.rect
            cx = (rect.x0 + rect.x1) / 2
            cy = (rect.y0 + rect.y1) / 2
            print(
                f"page={page_idx}, "
                f"rect=({rect.x0:.1f}, {rect.y0:.1f}, {rect.x1:.1f}, {rect.y1:.1f}), "
                f"center=({cx:.1f}, {cy:.1f}), "
                f"text='{text}'"
            )

if __name__ == "__main__":
    # Tkinter GUI 초기화
    root = Tk()
    root.withdraw()

    # PDF 파일 선택
    pdf_path = filedialog.askopenfilename(
        title="PDF 선택",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        print("PDF 파일을 선택하지 않았습니다.")
        exit()

    # 찾을 주석 값 입력
    value = simpledialog.askstring("입력", "찾을 주석 값:")
    if not value:
        print("검색어를 입력하지 않았습니다.")
        exit()

    # 좌표 검색
    find_annotation_positions(pdf_path, value, contains=True, ignore_case=True)
