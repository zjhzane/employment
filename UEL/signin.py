from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional

def render_pdf_page_to_image(pdf_path: Path, page_index: int = 0, dpi: int = 400):
    """用 pypdfium2 把指定页渲染成 PIL.Image。"""
    import pypdfium2 as pdfium

    if not pdf_path.exists():
        raise FileNotFoundError(f"找不到 PDF：{pdf_path}")

    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        page = pdf.get_page(page_index)
        scale = dpi / 72.0  # 72pt = 1 inch
        img = page.render(scale=scale).to_pil()
        page.close()
        return img
    finally:
        pdf.close()


def paste_stamp(
    base_img,
    stamp_path: Path,
    *,
    x_ratio: float,
    y_ratio: float,
    w_ratio: float,
    opacity: float = 1.0,
):
    """在 base_img 上叠加 stamp（按比例定位/缩放）。"""
    from PIL import Image

    if not stamp_path.exists():
        raise FileNotFoundError(f"找不到章图片：{stamp_path}")

    base_img = base_img.convert("RGBA")
    stamp = Image.open(stamp_path).convert("RGBA")

    bw, bh = base_img.size
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

    base_img.alpha_composite(stamp, dest=(x, y))
    return base_img


def save_as_single_page_pdf(img, out_pdf: Path):
    """把 PIL.Image 嵌入为同尺寸单页 PDF。"""
    from PIL import Image
    from reportlab.pdfgen import canvas

    out_pdf.parent.mkdir(parents=True, exist_ok=True)
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


def find_stamp_image(search_dirs, pattern: str = "*141217*.png") -> Optional[Path]:
    """在多个目录里按通配符找章图，找不到返回 None。"""
    for d in search_dirs:
        if not d.exists():
            continue
        matches = sorted(d.glob(pattern))
        if matches:
            return matches[0]
    return None


def find_pdf_by_prefix(search_dirs, prefix: str = "score", suffix: str = ".pdf") -> Optional[Path]:
    """在多个目录里找文件名以 prefix 开头且以后缀结尾的文件。"""
    for d in search_dirs:
        if not d.exists():
            continue
        matches = sorted(
            p for p in d.iterdir() if p.is_file() and p.name.startswith(prefix) and p.suffix.lower() == suffix
        )
        if matches:
            return matches[0]
    return None


def add_stamp_to_transcript_pdf(
    *,
    pdf_path: Path,
    stamp_path: Path,
    out_pdf: Path,
    page_index: int = 0,
    dpi: int = 400,
    stamp_x_ratio: float = 0.75,
    stamp_y_ratio: float = 0.60,
    stamp_w_ratio: float = 0.15,
    stamp_opacity: float = 1.0,
):
    img = render_pdf_page_to_image(pdf_path, page_index=page_index, dpi=dpi)

    stamped = paste_stamp(
        img,
        stamp_path,
        x_ratio=stamp_x_ratio,
        y_ratio=stamp_y_ratio,
        w_ratio=stamp_w_ratio,
        opacity=stamp_opacity,
    )
    save_as_single_page_pdf(stamped, out_pdf)


def main():
    desktop = Path.home() / "Desktop"
    day3_dir = Path(__file__).resolve().parent

    parser = argparse.ArgumentParser(description="给成绩单 PDF 添加印章（叠加图片并导出新 PDF）")
    parser.add_argument(
        "--pdf",
        type=str,
        default="",
        help="成绩单 PDF 路径；不填则在 day3 目录和桌面自动匹配以 score 开头的 .pdf",
    )
    parser.add_argument(
        "--stamp",
        type=str,
        default="",
        help="章图片 PNG 路径；不填则在 day3 目录和桌面自动匹配 *141217*.png",
    )
    parser.add_argument(
        "--out",
        type=str,
        default="",
        help="输出 PDF 路径；不填则输出到桌面：transcript.pdf",
    )
    parser.add_argument("--page", type=int, default=0, help="要盖章的页码（0=第一页）")
    parser.add_argument("--dpi", type=int, default=400, help="渲染清晰度，越大越清晰但文件更大")
    parser.add_argument("--stamp_x_ratio", type=float, default=0.7, help="章左上角相对宽度比例")
    parser.add_argument("--stamp_y_ratio", type=float, default=0.7, help="章左上角相对高度比例")
    parser.add_argument("--stamp_w_ratio", type=float, default=0.2, help="章宽度相对整图宽度比例")
    parser.add_argument("--stamp_opacity", type=float, default=1.0, help="章透明度 0~1")

    args = parser.parse_args()

    if args.pdf:
        pdf_path = Path(args.pdf)
    else:
        pdf_path = find_pdf_by_prefix(search_dirs=[day3_dir, desktop], prefix="score", suffix=".pdf")
        if not pdf_path:
            raise FileNotFoundError("在 day3 或桌面找不到以 score 开头的 PDF 文件")

    out_pdf = Path(args.out) if args.out else desktop / "transcript.pdf"

    if args.stamp:
        stamp_path = Path(args.stamp)
    else:
        stamp_path = find_stamp_image(search_dirs=[day3_dir, desktop], pattern="*141217*.png")
        if not stamp_path:
            raise FileNotFoundError("在 day3 或桌面找不到章图（匹配 *141217*.png）")

    add_stamp_to_transcript_pdf(
        pdf_path=pdf_path,
        stamp_path=stamp_path,
        out_pdf=out_pdf,
        page_index=args.page,
        dpi=args.dpi,
        stamp_x_ratio=args.stamp_x_ratio,
        stamp_y_ratio=args.stamp_y_ratio,
        stamp_w_ratio=args.stamp_w_ratio,
        stamp_opacity=args.stamp_opacity,
    )
    print(f"[OK] 盖章完成：{out_pdf}")


if __name__ == "__main__":
    main()

