"""
วิธีใช้:  python debug_pdf.py ชื่อไฟล์.pdf
แสดงข้อมูลข้อความ keyword และเส้น/กล่องที่ตรวจจับได้ในแต่ละหน้า
"""
import sys
from pathlib import Path
import fitz

KEYWORDS = ["ลงชื่อ", "ขอแสดงความเคารพ", "ตราประทับ", "นิติบุคคล", "ประทับตรา"]

def debug(pdf_path: str) -> None:
    doc = fitz.open(pdf_path)
    print(f"\n📄 ไฟล์: {pdf_path}  ({doc.page_count} หน้า)")

    for page_idx, page in enumerate(doc):
        print(f"\n{'='*60}")
        print(f"  หน้า {page_idx + 1}  (ขนาด {page.rect.width:.0f} x {page.rect.height:.0f} pt)")
        print(f"{'='*60}")

        # --- ค้นหา keyword ---
        for kw in KEYWORDS:
            hits = page.search_for(kw)
            if hits:
                for h in hits:
                    print(f"  [TEXT] \"{kw}\"  →  x0={h.x0:.1f} y0={h.y0:.1f} x1={h.x1:.1f} y1={h.y1:.1f}")

        # --- แสดง drawing ทั้งหมดในหน้า ---
        drawings = page.get_drawings()
        print(f"\n  Drawing elements: {len(drawings)} รายการ")
        for i, draw in enumerate(drawings):
            r = draw["rect"]
            shape = "วงกลม?" if abs(r.width - r.height) < 15 and r.width > 30 else "เส้น/กล่อง" if r.width > r.height * 2 else "อื่นๆ"
            print(f"    [{i:02d}] {shape:10s}  w={r.width:.1f} h={r.height:.1f}  "
                  f"x0={r.x0:.1f} y0={r.y0:.1f} x1={r.x1:.1f} y1={r.y1:.1f}")

    doc.close()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("วิธีใช้: python debug_pdf.py ชื่อไฟล์.pdf")
        sys.exit(1)
    debug(sys.argv[1])
