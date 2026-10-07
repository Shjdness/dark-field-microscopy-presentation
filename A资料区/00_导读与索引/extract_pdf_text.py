from pathlib import Path

from pypdf import PdfReader


root = Path(__file__).resolve().parents[2]
out_dir = Path(__file__).resolve().parent / "全文文本"
out_dir.mkdir(exist_ok=True)

for pdf in sorted(root.rglob("*.pdf")):
    reader = PdfReader(str(pdf))
    pages = []
    for index, page in enumerate(reader.pages, start=1):
        pages.append(f"\n\n===== PDF PAGE {index} =====\n\n")
        pages.append(page.extract_text() or "")
    (out_dir / f"{pdf.stem}.txt").write_text("".join(pages), encoding="utf-8")
