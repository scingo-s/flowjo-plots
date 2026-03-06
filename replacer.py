"""FlowJo PDF からプロット画像を切り出し、Excel の画像を差し替えるモジュール."""

import io
import zipfile
from copy import deepcopy

import fitz
from PIL import Image


# PDF グリッド座標 (PDF points)
Y_STARTS = [75.0, 121.0, 167.0, 214.0, 260.0, 306.0, 352.0,
            399.0, 445.0, 491.0, 537.0, 584.0, 630.0, 676.0]
COL_RANGES = [
    (128, 180),   # FSC/SSC
    (180, 230),   # Pop1/Pop2
    (220, 272),   # CAR
    (272, 316),   # CD3/CD56
    (316, 360),   # CD8a/CD56
    (360, 404),   # CD8b/CD56
    (404, 449),   # CD8a/CD8b
    (449, 500),   # CD16/CD56
]
ROW_HEIGHT = 46.0
DPI = 500
NUM_ROWS = 14
NUM_COLS = 8


def extract_plots_from_pdf(pdf_bytes: bytes) -> list[bytes]:
    """PDF から 112枚のプロット画像を切り出す.

    Returns:
        list of PNG bytes, 順序: row0-col0, row0-col1, ..., row0-col7, row1-col0, ...
        (image1.png=row0-col0, image2.png=row0-col1, ..., image8.png=row0-col7,
         image9.png=row1-col0, ...)
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    scale = DPI / 72
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat)
    img = Image.open(io.BytesIO(pix.tobytes("png")))

    plots = []
    for row in range(NUM_ROWS):
        for col in range(NUM_COLS):
            x1 = int(COL_RANGES[col][0] * scale)
            x2 = int(COL_RANGES[col][1] * scale)
            y1 = int(Y_STARTS[row] * scale)
            y2 = int((Y_STARTS[row] + ROW_HEIGHT) * scale)
            crop = img.crop((x1, y1, x2, y2))

            buf = io.BytesIO()
            crop.save(buf, format="PNG")
            plots.append(buf.getvalue())

    doc.close()
    return plots


def _build_image_name_order() -> list[str]:
    """Excel 内の image1.png ~ image112.png の順序リストを生成.

    Excel の画像番号はサンプル順: image1~8 = sample1, image9~16 = sample2, ...
    """
    return [f"image{i}.png" for i in range(1, NUM_ROWS * NUM_COLS + 1)]


def replace_images_in_xlsx(xlsx_bytes: bytes, plot_images: list[bytes]) -> bytes:
    """xlsx 内の xl/media/image*.png を差し替える.

    Args:
        xlsx_bytes: 元の Excel ファイル (bytes)
        plot_images: 112枚の PNG 画像 (bytes), extract_plots_from_pdf() の出力順

    Returns:
        差し替え済みの xlsx (bytes)
    """
    if len(plot_images) != NUM_ROWS * NUM_COLS:
        raise ValueError(
            f"Expected {NUM_ROWS * NUM_COLS} images, got {len(plot_images)}"
        )

    image_names = _build_image_name_order()
    image_map = {
        f"xl/media/{name}": data
        for name, data in zip(image_names, plot_images)
    }

    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zin:
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in image_map:
                    zout.writestr(item, image_map[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))

    return output.getvalue()


def process(pdf_bytes: bytes, xlsx_bytes: bytes) -> bytes:
    """PDF からプロットを切り出し、Excel の画像を差し替える.

    Args:
        pdf_bytes: FlowJo バッチ出力 PDF
        xlsx_bytes: Summary Excel ファイル

    Returns:
        画像差し替え済みの xlsx (bytes)
    """
    plots = extract_plots_from_pdf(pdf_bytes)
    return replace_images_in_xlsx(xlsx_bytes, plots)
