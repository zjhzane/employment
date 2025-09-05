from pathlib import Path
from PIL import Image
import pypdfium2 as pdfium
from reportlab.pdfgen import canvas

# ========= 基本配置（只改这里） =========
DESKTOP = Path.home() / "Desktop"

INPUT_DIR  = DESKTOP / "output_docs"   # 待处理PDF与stamp所在的文件夹
OUTPUT_DIR = DESKTOP / "毕业生转递单"     # 输出文件夹（自动创建）
STAMP_NAME = "stamp.png"               # 章图片文件名（存放在 INPUT_DIR 里）

# 截取高度（从顶部开始，0~1）
TOP_RATIO = 0.42

# 渲染清晰度（越大越清晰，文件越大）
DPI = 400

# 章的位置与大小（相对于“截取后的图片”宽高的百分比）
STAMP_X_RATIO = 0.75  # 章左上角相对宽度的比例（0~1）
STAMP_Y_RATIO = 0.60  # 章左上角相对高度的比例（0~1）
STAMP_W_RATIO = 0.15  # 章宽度相对整张图宽度的比例（0~1）
STAMP_OPACITY = 1     # 透明度 0~1（1 为不透明）
# ======================================


def render_first_page_to_image(pdf_path: Path, dpi: int) -> Image.Image:
    """用 pypdfium2 渲染PDF第一页为 PIL.Image"""
    if not pdf_path.exists():
        raise FileNotFoundError(f"找不到 PDF：{pdf_path}")
    pdf = pdfium.PdfDocument(str(pdf_path))
    page = pdf.get_page(0)
    scale = dpi / 72.0  # 72pt = 1 inch
    img = page.render(scale=scale).to_pil()
    page.close()
    pdf.close()
    return img


def crop_top(img: Image.Image, top_ratio: float) -> Image.Image:
    """从顶部按比例裁剪"""
    w, h = img.size
    crop_h = max(1, int(h * top_ratio))
    return img.crop((0, 0, w, crop_h))


def paste_stamp(base: Image.Image, stamp_path: Path,
                x_ratio: float, y_ratio: float, w_ratio: float, opacity: float = 1.0) -> Image.Image:
    """按比例在裁剪图上叠加章（透明PNG最佳）"""
    if not stamp_path.exists():
        print(f"[WARN] 未找到章图片：{stamp_path}，跳过盖章。")
        return base
    base = base.convert("RGBA")
    stamp = Image.open(stamp_path).convert("RGBA")

    bw, bh = base.size
    target_w = max(1, int(bw * w_ratio))
    scale = target_w / stamp.width
    target_h = max(1, int(stamp.height * scale))
    stamp = stamp.resize((target_w, target_h), Image.LANCZOS)

    if opacity < 1.0:
        alpha = stamp.split()[-1]
        alpha = alpha.point(lambda p: int(p * opacity))
        stamp.putalpha(alpha)

    x = max(0, min(bw - target_w, int(bw * x_ratio)))
    y = max(0, min(bh - target_h, int(bh * y_ratio)))

    base.alpha_composite(stamp, dest=(x, y))
    return base


def save_as_single_page_pdf(img: Image.Image, out_pdf: Path):
    """把图片嵌入为同尺寸单页PDF"""
    w, h = img.size
    c = canvas.Canvas(out_pdf.as_posix(), pagesize=(w, h))
    tmp_png = out_pdf.with_suffix(".tmp.png")
    # 铺白底以避免透明通道导致的查看器差异
    if img.mode == "RGBA":
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        bg.save(tmp_png.as_posix(), format="PNG")
    else:
        img.save(tmp_png.as_posix(), format="PNG")
    c.drawImage(tmp_png.as_posix(), 0, 0, width=w, height=h)
    c.showPage()
    c.save()
    try:
        tmp_png.unlink()
    except Exception:
        pass


def process_one_pdf(pdf_path: Path, stamp_path: Path, out_dir: Path):
    """处理单个PDF：第一页→裁剪→盖章→输出PDF"""
    try:
        img = render_first_page_to_image(pdf_path, DPI)
        cropped = crop_top(img, TOP_RATIO)
        stamped = paste_stamp(
            cropped, stamp_path,
            x_ratio=STAMP_X_RATIO, y_ratio=STAMP_Y_RATIO,
            w_ratio=STAMP_W_RATIO, opacity=STAMP_OPACITY
        )

        out_dir.mkdir(parents=True, exist_ok=True)
        out_pdf = out_dir / f"{pdf_path.stem}_截图_盖章.pdf"
        save_as_single_page_pdf(stamped, out_pdf)
        print(f"[OK] {pdf_path.name} → {out_pdf}")
    except Exception as e:
        print(f"[FAIL] 处理失败：{pdf_path.name}，原因：{e}")


if __name__ == "__main__":
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"找不到输入文件夹：{INPUT_DIR}")

    stamp_path = INPUT_DIR / STAMP_NAME
    if not stamp_path.exists():
        print(f"[WARN] 未找到章图片：{stamp_path}（将不盖章，仅裁剪输出）")

    pdf_list = sorted([p for p in INPUT_DIR.glob("*.pdf") if p.is_file()])
    if not pdf_list:
        raise FileNotFoundError(f"输入文件夹内没有 PDF：{INPUT_DIR}")

    print(f"[INFO] 找到 {len(pdf_list)} 个PDF，开始处理…")
    for pdf in pdf_list:
        process_one_pdf(pdf, stamp_path, OUTPUT_DIR)

    print(f"[DONE] 全部完成。输出目录：{OUTPUT_DIR}")
