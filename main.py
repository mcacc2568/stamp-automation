"""
PDF Batch Stamp Processor
=========================
ประมวลผลไฟล์ PDF แบบกลุ่ม โดย:
  1. Flatten ทุกหน้าเป็นภาพ 300 DPI (ล็อกลายเซ็นเดิม)
  2. ประทับตรา PNG ที่มุมขวาล่างของหน้าสุดท้าย
  3. บันทึกลงโฟลเดอร์ stamped_output/

วิธีติดตั้ง:
  pip install -r requirements.txt

วิธีใช้งาน:
  python main.py
"""

import io
import json
import platform
import random
import subprocess
import sys
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    HAS_GUI = True
except ImportError:
    HAS_GUI = False

if getattr(sys, "frozen", False) and not HAS_GUI:
    import ctypes
    ctypes.windll.user32.MessageBoxW(
0,
        "ไม่สามารถโหลด tkinter ได้\nกรุณาติดต่อผู้ดูแล",
        "ข้อผิดพลาด",
        0x10,
    )
    sys.exit(1)

try:
    import fitz  # PyMuPDF
except ImportError:
    print("❌  ไม่พบ PyMuPDF  —  กรุณาติดตั้งก่อน:")
    print("     pip install -r requirements.txt")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("❌  ไม่พบ Pillow  —  กรุณาติดตั้งก่อน:")
    print("     pip install -r requirements.txt")
    sys.exit(1)


# ---------------------------------------------------------------------------
# ขนาดและระยะขอบของตรายาง (หน่วย: point  —  1 pt = 1/72 นิ้ว)
# ---------------------------------------------------------------------------
STAMP_W: int = 150   # ความกว้างสูงสุด
STAMP_H: int = 150   # ความสูงสูงสุด
MARGIN: int = 20     # ระยะห่างจากขอบขวาและล่าง
DPI: int = 300       # ความละเอียดในการ flatten

MAX_FILE_BYTES: int = int(3.8 * 1024 * 1024)  # 3.8 MB (เผื่อ margin ก่อนถึง 4 MB)

# BASE_DIR ทำงานได้ทั้งตอนรันเป็น script และตอน frozen เป็น .exe
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

# config อยู่ใน home dir ไม่ผูกกับที่อยู่ของ .exe
CONFIG_PATH = Path.home() / ".stamp_automation_config.json"


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_config(data: dict) -> None:
    CONFIG_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

# ---------------------------------------------------------------------------
# ค่าพารามิเตอร์สำหรับ Human-like Stamping
# ---------------------------------------------------------------------------
JITTER_PX: float = 20.0    # สุ่มขยับ ±20 pt รอบจุดอ้างอิง (ลดจาก 60 ไม่ให้ขึ้นทับจำนวนเงิน)
ROTATION_MAX: float = 15.0  # สุ่มเอียง ±15 องศา
OPACITY_MIN: float = 0.90   # สุ่มความเข้ม 90%–100%
STAMP_Y_OFFSET: int = 30    # เลื่อนตราลงมาจาก base position อีก 30 pt

SIGNATURE_W: int = 120  # ความกว้างลายเซ็น (pt)
SIGNATURE_H: int = 60   # ความสูงลายเซ็น (pt)


# ---------------------------------------------------------------------------
# GUI: form หลัก — เลือกทุกอย่างในหน้าต่างเดียว
# ---------------------------------------------------------------------------
def show_main_form(cfg: dict) -> tuple[Path, Path, Path, Path | None]:
    result: dict = {}
    default_stamps_dir = Path(cfg["stamps_dir"]) if cfg.get("stamps_dir") else BASE_DIR / "stamps"

    root = tk.Tk()
    root.title("ระบบประทับตรา PDF")
    root.resizable(False, False)

    def on_close():
        sys.exit(0)
    root.protocol("WM_DELETE_WINDOW", on_close)

    pad = {"padx": 12, "pady": 6}

    # --- Header ---
    tk.Label(root, text="ระบบประทับตรา PDF", font=("", 14, "bold")).grid(
        row=0, column=0, columnspan=3, pady=(18, 10)
    )

    # --- โฟลเดอร์ PDF ---
    folder_var = tk.StringVar(value=cfg.get("source_dir", ""))
    pdf_count_var = tk.StringVar(value="")
    if folder_var.get():
        n = len(sorted(Path(folder_var.get()).glob("*.pdf")))
        pdf_count_var.set(f"พบ {n} ไฟล์ PDF" if n else "⚠️  ไม่พบไฟล์ PDF")

    def browse_source():
        path = filedialog.askdirectory(
            title="เลือกโฟลเดอร์ที่มีไฟล์ PDF",
            initialdir=folder_var.get() or str(Path.home() / "Desktop"),
        )
        if path:
            folder_var.set(path)
            n = len(sorted(Path(path).glob("*.pdf")))
            pdf_count_var.set(f"พบ {n} ไฟล์ PDF" if n else "⚠️  ไม่พบไฟล์ PDF")

    tk.Label(root, text="โฟลเดอร์ PDF:", anchor="e").grid(row=1, column=0, sticky="e", **pad)
    tk.Entry(root, textvariable=folder_var, width=38, state="readonly").grid(row=1, column=1, **pad)
    tk.Button(root, text="เลือก…", command=browse_source, width=7).grid(row=1, column=2, **pad)
    tk.Label(root, textvariable=pdf_count_var, fg="gray", anchor="w").grid(
        row=2, column=1, sticky="w", padx=12, pady=(0, 4)
    )

    # --- โฟลเดอร์ตรายาง ---
    stamps_dir_var = tk.StringVar(value=str(default_stamps_dir))
    stamp_var = tk.StringVar()
    stamp_combo: list = []  # mutable ref เพื่อให้ refresh ได้

    def refresh_stamps(directory: str) -> None:
        stamps = list_stamps(Path(directory))
        names = [s.name for s in stamps]
        if stamp_combo:
            stamp_combo[0]["values"] = names
            stamp_var.set(names[0] if names else "")

    def browse_stamps_dir():
        path = filedialog.askdirectory(
            title="เลือกโฟลเดอร์ที่เก็บไฟล์ตรายาง (.png)",
            initialdir=stamps_dir_var.get() or str(Path.home() / "Desktop"),
        )
        if path:
            stamps_dir_var.set(path)
            refresh_stamps(path)

    tk.Label(root, text="โฟลเดอร์ตรายาง:", anchor="e").grid(row=3, column=0, sticky="e", **pad)
    tk.Entry(root, textvariable=stamps_dir_var, width=38, state="readonly").grid(row=3, column=1, **pad)
    tk.Button(root, text="เลือก…", command=browse_stamps_dir, width=7).grid(row=3, column=2, **pad)

    # --- ตรายาง dropdown ---
    initial_stamps = list_stamps(default_stamps_dir)
    stamp_var.set(initial_stamps[0].name if initial_stamps else "")
    combo = ttk.Combobox(
        root, textvariable=stamp_var,
        values=[s.name for s in initial_stamps],
        state="readonly", width=36,
    )
    combo.grid(row=4, column=1, **pad)
    stamp_combo.append(combo)
    tk.Label(root, text="ตรายาง:", anchor="e").grid(row=4, column=0, sticky="e", **pad)

    # --- ลายเซ็น ---
    sig_var = tk.StringVar(value=cfg.get("signature_path", ""))

    def browse_signature():
        path = filedialog.askopenfilename(
            title="เลือกไฟล์ลายเซ็น (.png)",
            filetypes=[("PNG", "*.png")],
            initialdir=str(Path(sig_var.get()).parent) if sig_var.get() else str(Path.home() / "Desktop"),
        )
        if path:
            sig_var.set(path)

    tk.Label(root, text="ลายเซ็น:", anchor="e").grid(row=5, column=0, sticky="e", **pad)
    tk.Entry(root, textvariable=sig_var, width=38, state="readonly").grid(row=5, column=1, **pad)
    tk.Button(root, text="เลือก…", command=browse_signature, width=7).grid(row=5, column=2, **pad)
    tk.Label(root, text="(ไม่บังคับ)", fg="gray", anchor="w").grid(row=6, column=1, sticky="w", padx=12, pady=(0, 4))

    # --- บันทึกที่ ---
    output_var = tk.StringVar(value=cfg.get("output_dir", str(Path.home() / "Desktop")))

    def browse_output():
        path = filedialog.askdirectory(
            title="เลือกโฟลเดอร์สำหรับบันทึกไฟล์ที่ประทับตราแล้ว",
            initialdir=output_var.get() or str(Path.home() / "Desktop"),
        )
        if path:
            output_var.set(path)

    tk.Label(root, text="บันทึกที่:", anchor="e").grid(row=7, column=0, sticky="e", **pad)
    tk.Entry(root, textvariable=output_var, width=38, state="readonly").grid(row=7, column=1, **pad)
    tk.Button(root, text="เลือก…", command=browse_output, width=7).grid(row=7, column=2, **pad)

    # --- ปุ่มเริ่ม ---
    def on_submit():
        if not folder_var.get():
            messagebox.showwarning("ยังไม่ครบ", "กรุณาเลือกโฟลเดอร์ PDF")
            return
        pdf_files = sorted(Path(folder_var.get()).glob("*.pdf"))
        if not pdf_files:
            messagebox.showerror("ไม่พบไฟล์", "ไม่พบไฟล์ PDF ในโฟลเดอร์ที่เลือก")
            return
        if not stamp_var.get():
            messagebox.showwarning("ยังไม่ครบ", "ไม่พบตรายางในโฟลเดอร์ที่เลือก\nกรุณาเลือกโฟลเดอร์ที่มีไฟล์ .png")
            return
        if not output_var.get():
            messagebox.showwarning("ยังไม่ครบ", "กรุณาเลือกโฟลเดอร์บันทึก")
            return

        stamps = list_stamps(Path(stamps_dir_var.get()))
        stamp_path = next((s for s in stamps if s.name == stamp_var.get()), None)
        if stamp_path is None:
            messagebox.showerror("ไม่พบตรายาง", f"ไม่พบไฟล์ {stamp_var.get()}")
            return

        save_config({
            "stamps_dir": stamps_dir_var.get(),
            "source_dir": folder_var.get(),
            "output_dir": output_var.get(),
            "signature_path": sig_var.get(),
        })

        result["folder"]    = Path(folder_var.get())
        result["stamp"]     = stamp_path
        result["output"]    = Path(output_var.get())
        result["signature"] = Path(sig_var.get()) if sig_var.get() else None
        root.destroy()

    tk.Button(
        root, text="  เริ่มประทับตรา  ", command=on_submit,
        bg="#1a73e8", fg="white", font=("", 11, "bold"),
        relief="flat", padx=16, pady=10, cursor="hand2",
    ).grid(row=8, column=0, columnspan=3, pady=(14, 20))

    root.lift()
    root.focus_force()
    root.mainloop()

    if not result:
        sys.exit(0)
    return result["folder"], result["stamp"], result["output"], result.get("signature")


# ---------------------------------------------------------------------------
# 1. รับ path โฟลเดอร์ PDF
# ---------------------------------------------------------------------------
def get_pdf_folder() -> Path:
    print("\n" + "=" * 60)
    print("  ระบบประทับตรา PDF")
    print("=" * 60)

    if HAS_GUI:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        raw = filedialog.askdirectory(
            title="เลือกโฟลเดอร์ที่มีไฟล์ PDF",
            initialdir=Path.home() / "Desktop",
        )
        root.destroy()
        if not raw:
            sys.exit(0)
        folder = Path(raw)
    else:
        raw = input("\nกรอก path โฟลเดอร์ที่มีไฟล์ PDF (Enter = โฟลเดอร์ปัจจุบัน): ").strip()
        folder = Path(raw) if raw else Path.cwd()
        if not folder.exists() or not folder.is_dir():
            print(f"❌  ไม่พบโฟลเดอร์: {folder}")
            sys.exit(1)

    pdf_files = sorted(folder.glob("*.pdf"))
    if not pdf_files:
        msg = f"ไม่พบไฟล์ .pdf ในโฟลเดอร์:\n{folder}"
        if HAS_GUI:
            messagebox.showerror("ไม่พบไฟล์ PDF", msg)
        else:
            print(f"❌  {msg}")
        sys.exit(1)

    print(f"\n✅  พบไฟล์ PDF จำนวน {len(pdf_files)} ไฟล์")
    print(f"     โฟลเดอร์: {folder}")
    return folder


# ---------------------------------------------------------------------------
# helper: เปิดโฟลเดอร์ใน Finder / Explorer
# ---------------------------------------------------------------------------
def open_output_folder(path: Path) -> None:
    try:
        if platform.system() == "Darwin":
            subprocess.run(["open", str(path)], check=False)
        elif platform.system() == "Windows":
            subprocess.run(["explorer", str(path)], check=False)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 1b. เลือกโฟลเดอร์ปลายทาง
# ---------------------------------------------------------------------------
def get_output_folder(source: Path) -> Path:
    if HAS_GUI:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        raw = filedialog.askdirectory(
            title="เลือกโฟลเดอร์สำหรับบันทึกไฟล์ที่ประทับตราแล้ว",
            initialdir=Path.home() / "Desktop",
        )
        root.destroy()
        if not raw:
            sys.exit(0)
        return Path(raw)
    else:
        raw = input("\nกรอก path โฟลเดอร์ปลายทาง (Enter = stamped_output/): ").strip()
        return Path(raw) if raw else source / "stamped_output"


# ---------------------------------------------------------------------------
# 2. สแกนหาตรายาง PNG
# ---------------------------------------------------------------------------
def list_stamps(stamps_dir: Path) -> list[Path]:
    if not stamps_dir.exists() or not stamps_dir.is_dir():
        return []
    return sorted(stamps_dir.glob("*.png"))


# ---------------------------------------------------------------------------
# 3. ให้ผู้ใช้เลือกตรายาง
# ---------------------------------------------------------------------------
def select_stamp(stamps: list[Path]) -> Path:
    if len(stamps) == 1:
        print(f"\n✅  ใช้ตรายาง: {stamps[0].name}")
        return stamps[0]

    print("\n--- เลือกตรายางที่ต้องการประทับ ---")
    for i, s in enumerate(stamps, start=1):
        print(f"  {i}. {s.name}")

    while True:
        raw = input(f"\nกด 1–{len(stamps)} เพื่อเลือก: ").strip()
        if raw.isdigit():
            choice = int(raw)
            if 1 <= choice <= len(stamps):
                selected = stamps[choice - 1]
                print(f"✅  เลือกตรายาง: {selected.name}")
                return selected
        print(f"⚠️   กรุณากรอกตัวเลข 1 ถึง {len(stamps)}")




# ---------------------------------------------------------------------------
# 5a. เตรียมภาพตรายางพร้อม rotation + opacity (ใช้ Pillow)
# ---------------------------------------------------------------------------
EMBED_PX: int = 300  # ความละเอียด embed สูงสุด (pixel) — ลดขนาดไฟล์โดยไม่ flatten

def prepare_stamp_bytes(stamp_path: Path, angle: float, opacity: float) -> bytes:
    img = Image.open(stamp_path).convert("RGBA")

    # จำกัดขนาดก่อน rotate เพื่อลดขนาดไฟล์ที่ฝังลง PDF
    if max(img.size) > EMBED_PX:
        img.thumbnail((EMBED_PX, EMBED_PX), Image.LANCZOS)

    rotated = img.rotate(angle, expand=True, resample=Image.BICUBIC)

    r, g, b, a = rotated.split()
    a = a.point(lambda x: int(x * opacity))
    rotated.putalpha(a)

    buf = io.BytesIO()
    rotated.save(buf, format="PNG")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 5b. ประทับตราที่มุมขวาล่างของหน้า พร้อม Human-like Stamping
# ---------------------------------------------------------------------------
def apply_stamp(page: fitz.Page, stamp_path: Path, stamp_bytes: bytes | None = None) -> None:
    offset_x = random.uniform(-JITTER_PX, JITTER_PX)
    offset_y = random.uniform(-JITTER_PX, JITTER_PX)

    if stamp_bytes is None:
        angle   = random.uniform(-ROTATION_MAX, ROTATION_MAX)
        opacity = random.uniform(OPACITY_MIN, 1.0)
        stamp_bytes = prepare_stamp_bytes(stamp_path, angle, opacity)

    w = page.rect.width
    h = page.rect.height
    base_x = w - STAMP_W - MARGIN
    base_y = h - STAMP_H - MARGIN + STAMP_Y_OFFSET
    x0 = min(max(base_x + offset_x, MARGIN), w - STAMP_W)
    y0 = min(max(base_y + offset_y, MARGIN), h - STAMP_H)
    rect = fitz.Rect(x0, y0, x0 + STAMP_W, y0 + STAMP_H)
    page.insert_image(rect, stream=stamp_bytes, keep_proportion=True)


# ---------------------------------------------------------------------------
# 5c. วางลายเซ็นซ้ายของตราประทับ พร้อม human-like jitter + rotation + opacity
# ---------------------------------------------------------------------------
SIG_ROTATION_MAX: float = 5.0  # ลายเซ็นเอียงน้อยกว่าตรา ±5 องศา

def apply_signature(page: fitz.Page, sig_path: Path) -> None:
    angle   = random.uniform(-SIG_ROTATION_MAX, SIG_ROTATION_MAX)
    opacity = random.uniform(OPACITY_MIN, 1.0)
    sig_bytes = prepare_stamp_bytes(sig_path, angle, opacity)

    offset_x = random.uniform(-JITTER_PX, JITTER_PX)
    offset_y = random.uniform(-JITTER_PX, JITTER_PX)

    w = page.rect.width
    h = page.rect.height
    stamp_base_x = w - STAMP_W - MARGIN
    stamp_base_y = h - STAMP_H - MARGIN + STAMP_Y_OFFSET
    base_x = stamp_base_x - SIGNATURE_W - 10
    base_y = stamp_base_y + (STAMP_H - SIGNATURE_H) // 2
    sig_x0 = min(max(base_x + offset_x, MARGIN), w - SIGNATURE_W)
    sig_y0 = min(max(base_y + offset_y, MARGIN), h - SIGNATURE_H)
    rect = fitz.Rect(sig_x0, sig_y0, sig_x0 + SIGNATURE_W, sig_y0 + SIGNATURE_H)
    page.insert_image(rect, stream=sig_bytes, keep_proportion=True)




# ---------------------------------------------------------------------------
# 6. ประมวลผลไฟล์เดียว
# ---------------------------------------------------------------------------
def process_pdf(pdf_path: Path, stamp_path: Path, output_dir: Path, sig_path: Path | None = None) -> None:
    doc = fitz.open(str(pdf_path))

    if doc.needs_pass:
        doc.close()
        raise PermissionError("ไฟล์ถูกเข้ารหัส / ต้องการรหัสผ่าน")
    if doc.page_count == 0:
        doc.close()
        raise ValueError("ไฟล์ไม่มีหน้าเอกสาร (0 หน้า)")

    for page in doc:
        if sig_path:
            apply_signature(page, sig_path)
        apply_stamp(page, stamp_path)

    best = doc.tobytes(deflate=True, garbage=4)
    doc.close()

    size_mb = len(best) / 1024 / 1024
    print(f"  ⚙️   {size_mb:.2f} MB", end="  ")

    out_path = output_dir / (pdf_path.stem + "_stamped.pdf")
    out_path.write_bytes(best)


# ---------------------------------------------------------------------------
# 7. ประมวลผลทั้งโฟลเดอร์
# ---------------------------------------------------------------------------
def batch_process(folder: Path, stamp_path: Path, output_dir: Path, sig_path: Path | None = None) -> None:
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        print(f"\n❌  สร้างโฟลเดอร์ปลายทางไม่ได้: {e}")
        sys.exit(1)

    pdf_files = sorted(folder.glob("*.pdf"))
    total = len(pdf_files)
    succeeded: list[str] = []
    failed: list[tuple[str, str]] = []

    print(f"\n{'=' * 60}")
    print(f"  เริ่มประมวลผล {total} ไฟล์  →  ตรายาง: {stamp_path.name}")
    print(f"  บันทึกผลลัพธ์ที่: {output_dir}")
    print(f"{'=' * 60}\n")

    for idx, pdf_path in enumerate(pdf_files, start=1):
        label = f"[{idx}/{total}] {pdf_path.name}"
        print(f"  ⏳  {label}", end="", flush=True)

        try:
            process_pdf(pdf_path, stamp_path, output_dir, sig_path)
            succeeded.append(pdf_path.name)
            print(f"\r  ✅  {label}")
        except PermissionError as e:
            failed.append((pdf_path.name, str(e)))
            print(f"\r  🔒  {label}  —  {e}")
        except fitz.FileDataError as e:
            failed.append((pdf_path.name, f"ไฟล์เสียหาย: {e}"))
            print(f"\r  💥  {label}  —  ไฟล์เสียหาย")
        except Exception as e:
            failed.append((pdf_path.name, str(e)))
            print(f"\r  ❌  {label}  —  {e}")

    # รายงานสรุป
    print(f"\n{'=' * 60}")
    print(f"  สรุปผลการประมวลผล")
    print(f"{'=' * 60}")
    print(f"  ✅  สำเร็จ  : {len(succeeded)} ไฟล์")
    print(f"  ❌  ข้ามไป  : {len(failed)} ไฟล์")

    if failed:
        print(f"\n  รายการที่ข้ามไป:")
        for name, reason in failed:
            print(f"    • {name}  →  {reason}")

    print(f"\n  ไฟล์ผลลัพธ์อยู่ที่: {output_dir}")
    print(f"{'=' * 60}\n")

    open_output_folder(output_dir)
    if HAS_GUI:
        if failed:
            messagebox.showwarning(
                "เสร็จแล้ว (มีบางไฟล์ข้ามไป)",
                f"สำเร็จ {len(succeeded)} ไฟล์\nข้ามไป {len(failed)} ไฟล์\n\nดูรายละเอียดในหน้าต่าง Terminal",
            )
        else:
            messagebox.showinfo(
                "เสร็จแล้ว ✅",
                f"ประทับตราเสร็จ {len(succeeded)} ไฟล์\n\nเปิดโฟลเดอร์ผลลัพธ์ให้แล้ว",
            )


# ---------------------------------------------------------------------------
# 8. ยืนยันก่อนเริ่ม
# ---------------------------------------------------------------------------
def confirm_processing(folder: Path, stamp_path: Path, output_dir: Path) -> None:
    pdf_count = len(sorted(folder.glob("*.pdf")))

    if HAS_GUI:
        msg = (
            f"โฟลเดอร์ต้นทาง  :  {folder.name}\n"
            f"ไฟล์ PDF         :  {pdf_count} ไฟล์\n"
            f"ตรายาง           :  {stamp_path.name}\n"
            f"บันทึกที่           :  {output_dir.name}\n\n"
            f"ต้องการเริ่มประทับตราหรือไม่?"
        )
        if not messagebox.askyesno("ยืนยันการประทับตรา", msg):
            sys.exit(0)
    else:
        print(f"\n{'─' * 40}")
        print(f"  โฟลเดอร์ต้นทาง : {folder.name}")
        print(f"  ไฟล์ PDF       : {pdf_count} ไฟล์")
        print(f"  ตรายาง         : {stamp_path.name}")
        print(f"  บันทึกที่       : {output_dir}")
        print(f"{'─' * 40}")
        input("กด Enter เพื่อเริ่ม  (Ctrl+C = ยกเลิก): ")


# ---------------------------------------------------------------------------
# 9. จุดเริ่มต้นโปรแกรม
# ---------------------------------------------------------------------------
def main() -> None:
    cfg = load_config()
    if HAS_GUI:
        folder, stamp_path, output_dir, sig_path = show_main_form(cfg)
    else:
        default_stamps_dir = Path(cfg["stamps_dir"]) if cfg.get("stamps_dir") else BASE_DIR / "stamps"
        stamps = list_stamps(default_stamps_dir)
        if not stamps:
            print(f"❌  ไม่พบไฟล์ .png ในโฟลเดอร์ตรายาง: {default_stamps_dir}")
            sys.exit(1)
        folder     = get_pdf_folder()
        stamp_path = select_stamp(stamps)
        output_dir = get_output_folder(folder)
        sig_raw = cfg.get("signature_path", "")
        sig_path = Path(sig_raw) if sig_raw and Path(sig_raw).exists() else None
        confirm_processing(folder, stamp_path, output_dir)
    batch_process(folder, stamp_path, output_dir, sig_path)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        msg = traceback.format_exc()
        if HAS_GUI:
            messagebox.showerror("เกิดข้อผิดพลาด", msg)
        else:
            print(msg)
        sys.exit(1)
