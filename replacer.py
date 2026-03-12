"""FlowJo PDF からプロット画像を切り出し、Excel に IMAGE 関数として挿入するモジュール."""

from __future__ import annotations

import io
import re
import zipfile

import fitz
import numpy as np
from lxml import etree
from PIL import Image

Image.MAX_IMAGE_PIXELS = 200_000_000  # 高DPI レンダリング用

DETECT_DPI = 500   # グリッド検出用
RENDER_DPI = 1000  # 画像切り出し用（高解像度）
NUM_COLS = 8  # プロット列数（固定: FSC/SSC, Pop1/2, CAR, CD3/56, etc.）

# Excel Plot シートのセル配置
EXCEL_COLS = ["E", "F", "H", "J", "L", "M", "N", "O"]
EXCEL_ROW_START = 5
EXCEL_ROW_STEP = 5

# 名前空間
NS_SHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_RICHDATA = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
NS_RICHDATA2 = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
NS_RICHVALREL = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Rich Data relationship types
REL_TYPE_METADATA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
REL_TYPE_RICHVALUE = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue"
REL_TYPE_RICHVALREL = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel"
REL_TYPE_RICHVALSTRUCT = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure"
REL_TYPE_RICHVALTYPES = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes"
REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

GUID_RICHVALUE = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"

# 列文字 → 列番号 (1-indexed)
_COL_MAP = {}
for _i, _c in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
    _COL_MAP[_c] = _i + 1


def _col_to_num(col_letter: str) -> int:
    return _COL_MAP[col_letter]


def _find_zero_gaps(proj: np.ndarray, min_width: int = 15) -> list[tuple[int, int, int]]:
    """投影配列からゼロコンテンツのギャップを検出する.

    Returns:
        (start, end, width) のリスト.
    """
    in_gap = proj == 0
    gaps = []
    gap_start = None
    for i in range(len(proj)):
        if in_gap[i] and gap_start is None:
            gap_start = i
        elif not in_gap[i] and gap_start is not None:
            width = i - gap_start
            if width >= min_width:
                gaps.append((gap_start, i, width))
            gap_start = None
    if gap_start is not None:
        width = len(proj) - gap_start
        if width >= min_width:
            gaps.append((gap_start, len(proj), width))
    return gaps


def _detect_grid(img: Image.Image) -> tuple[list[tuple[int, int]], list[tuple[int, int]]]:
    """画像からプロットグリッドの行・列バンドを自動検出する.

    ゼロコンテンツギャップ方式: コンテンツストリップ内の完全白列/行ギャップ(>=15px)
    を検出し、左右マージンとテキストラベル領域を除外して8プロット列を特定する.

    Returns:
        (row_bands, col_bands) - 各バンドは (start, end) のピクセル座標タプル.
        col_bands はラベル列を除いたプロット列のみ.
    """
    arr = np.array(img.convert("L"))
    h, w = arr.shape
    content = arr < 240  # non-white pixels

    # コンテンツ行範囲を検出
    row_proj = content.sum(axis=1)
    row_active = np.where(row_proj > w * 0.01)[0]
    if len(row_active) == 0:
        return [], []
    r1, r2 = int(row_active[0]), int(row_active[-1])

    col_active = np.where(content.sum(axis=0) > h * 0.01)[0]
    if len(col_active) == 0:
        return [], []
    c1, c2 = int(col_active[0]), int(col_active[-1])

    # --- 列バンド検出 ---
    # コンテンツ行ストリップ内の列投影でゼロギャップを検出
    strip = content[r1:r2 + 1, :]
    col_proj = strip.sum(axis=0)
    col_gaps = _find_zero_gaps(col_proj, min_width=15)

    # 左右マージン（画像端に接するギャップ）を除外
    internal_col_gaps = [(s, e, gw) for s, e, gw in col_gaps if s > 0 and e < w]

    # 最大の内部ギャップ（テキストラベル-プロット間）を除外、残りがプロット区切り
    col_dividers = []
    if len(internal_col_gaps) > 1:
        sorted_by_width = sorted(internal_col_gaps, key=lambda x: x[2], reverse=True)
        text_gap = sorted_by_width[0]
        col_dividers = sorted(sorted_by_width[1:], key=lambda x: x[0])
        # プロット列領域を定義
        first_plot_start = text_gap[1]  # テキストギャップの右端
        col_bands = []
        prev_start = first_plot_start
        for gs, ge, _ in col_dividers:
            col_bands.append((prev_start, gs))
            prev_start = ge
        col_bands.append((prev_start, c2))
    elif len(internal_col_gaps) == 1:
        # テキストギャップのみ → プロット領域は1つ
        col_bands = [(internal_col_gaps[0][1], c2)]
    else:
        col_bands = [(c1, c2)]

    # --- 行バンド検出 ---
    # プロット列範囲内の行投影でゼロギャップを検出
    plot_left = col_bands[0][0] if col_bands else c1
    plot_right = col_bands[-1][1] if col_bands else c2
    col_strip = content[:, plot_left:plot_right + 1]
    row_proj2 = col_strip.sum(axis=1)
    row_gaps = _find_zero_gaps(row_proj2, min_width=15)

    # マージンギャップ（コンテンツ外）を除外
    internal_row_gaps = [(s, e, gw) for s, e, gw in row_gaps
                         if s > r1 and e < r2]

    # 行領域を構築
    row_dividers = sorted(internal_row_gaps, key=lambda x: x[0])
    row_bands = []
    prev_start = r1
    for gs, ge, _ in row_dividers:
        row_bands.append((prev_start, gs))
        prev_start = ge
    row_bands.append((prev_start, r2))

    return row_bands, col_bands


def _extract_page_plots(page, detect_scale, render_scale, ratio):
    """1ページからプロット画像を切り出す."""
    det_mat = fitz.Matrix(detect_scale, detect_scale)
    det_pix = page.get_pixmap(matrix=det_mat)
    det_img = Image.open(io.BytesIO(det_pix.tobytes("png")))

    row_bands, col_bands = _detect_grid(det_img)

    ren_mat = fitz.Matrix(render_scale, render_scale)
    ren_pix = page.get_pixmap(matrix=ren_mat)
    hi_img = Image.open(io.BytesIO(ren_pix.tobytes("png")))

    plots = []
    for row_s, row_e in row_bands:
        for col_s, col_e in col_bands:
            crop = hi_img.crop((
                int(col_s * ratio), int(row_s * ratio),
                int(col_e * ratio), int(row_e * ratio),
            ))
            buf = io.BytesIO()
            crop.save(buf, format="PNG")
            plots.append(buf.getvalue())

    return plots, len(row_bands), len(col_bands)


def extract_plots_from_pdf(pdf_bytes: bytes) -> tuple[list[bytes], int, int]:
    """PDF からプロット画像を自動検出・切り出す.

    1ページPDF: フルグリッド検出 (行×列).
    マルチページPDF: 各ページを1サンプル (1行×N列) として処理し結合.

    Returns:
        (plots, num_rows, num_cols) - 画像バイトのリスト、検出された行数・列数.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    detect_scale = DETECT_DPI / 72
    render_scale = RENDER_DPI / 72
    ratio = RENDER_DPI / DETECT_DPI

    if doc.page_count == 1:
        plots, num_rows, num_cols = _extract_page_plots(
            doc[0], detect_scale, render_scale, ratio)
    else:
        # マルチページ: 各ページ = 1サンプル (1行)
        all_plots = []
        num_cols = 0
        for page_idx in range(doc.page_count):
            page_plots, _, page_cols = _extract_page_plots(
                doc[page_idx], detect_scale, render_scale, ratio)
            all_plots.extend(page_plots)
            if page_cols > num_cols:
                num_cols = page_cols
        plots = all_plots
        num_rows = doc.page_count

    doc.close()
    return plots, num_rows, num_cols


def _find_plot_sheet(zf: zipfile.ZipFile) -> tuple[str, str]:
    """Plot シートのシート名と XML ファイルパスを返す."""
    ns = {"x": NS_SHEET, "r": NS_REL}

    wb_root = etree.fromstring(zf.read("xl/workbook.xml"))
    rels_root = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    # シート名→rId
    target_name = None
    target_rid = None
    for s in wb_root.findall(".//x:sheet", ns):
        name = s.get("name")
        if "Plot" in name and "報告" not in name:
            target_name = name
            target_rid = s.get(f"{{{NS_REL}}}id")
            break

    if target_rid is None:
        # fallback to first sheet
        s = wb_root.findall(".//x:sheet", ns)[0]
        target_name = s.get("name")
        target_rid = s.get(f"{{{NS_REL}}}id")

    # rId→ファイルパス
    for rel in rels_root:
        if rel.get("Id") == target_rid:
            return target_name, "xl/" + rel.get("Target")

    raise ValueError("Plot sheet not found")



def insert_images_to_xlsx(xlsx_bytes: bytes, plot_images: list[bytes],
                          num_rows: int = 0, num_cols: int = NUM_COLS) -> bytes:
    """Excel の Plot シートに IMAGE 関数 (Rich Data) として画像を挿入する."""
    n = len(plot_images)
    if num_rows == 0:
        num_rows = n // num_cols

    # 画像セルの位置を生成
    cells = []
    for row_i in range(num_rows):
        excel_row = EXCEL_ROW_START + row_i * EXCEL_ROW_STEP
        for col_i in range(num_cols):
            idx = row_i * num_cols + col_i
            if idx >= n:
                break
            col_letter = EXCEL_COLS[col_i] if col_i < len(EXCEL_COLS) else None
            if col_letter is None:
                break
            cells.append((col_letter, excel_row, idx))

    # --- Rich Data XML を文字列テンプレートで生成 (lxml 非依存) ---
    _XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

    # 1. richValueRel.xml
    rv_rel_xml = _XML_DECL
    rv_rel_xml += '<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    for i in range(n):
        rv_rel_xml += f'<rel r:id="rId{i + 1}"/>'
    rv_rel_xml += '</richValueRels>'

    # 2. richValueRel.xml.rels
    rv_rels_xml = _XML_DECL
    rv_rels_xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    for i in range(n):
        rv_rels_xml += (f'<Relationship Id="rId{i + 1}" '
                        f'Type="{REL_TYPE_IMAGE}" '
                        f'Target="../media/image{i + 1}.png"/>')
    rv_rels_xml += '</Relationships>'

    # 3. rdrichvalue.xml
    rv_data_xml = _XML_DECL
    rv_data_xml += f'<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="{n}">'
    for i in range(n):
        rv_data_xml += f'<rv s="0"><v>{i}</v><v>5</v></rv>'
    rv_data_xml += '</rvData>'

    # 4. rdrichvaluestructure.xml
    rv_struct_xml = _XML_DECL
    rv_struct_xml += '<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">'
    rv_struct_xml += '<s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s>'
    rv_struct_xml += '</rvStructures>'

    # 5. rdRichValueTypes.xml
    rv_types_xml = _XML_DECL
    rv_types_xml += ('<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"'
                     ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
                     ' mc:Ignorable="x"'
                     ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    rv_types_xml += '<global><keyFlags>'
    for kname in ["_Self", "_DisplayString", "_Flags", "_Format",
                   "_SubLabel", "_Attribution", "_Icon", "_Display",
                   "_CanonicalPropertyNames", "_ClassificationId"]:
        if kname == "_Self":
            rv_types_xml += f'<key name="{kname}"><flag name="ExcludeFromFile" value="1"/><flag name="ExcludeFromCalcComparison" value="1"/></key>'
        else:
            rv_types_xml += f'<key name="{kname}"><flag name="ExcludeFromCalcComparison" value="1"/></key>'
    rv_types_xml += '</keyFlags></global></rvTypesInfo>'

    # 6. metadata.xml
    meta_xml = _XML_DECL
    meta_xml += '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">'
    meta_xml += '<metadataTypes count="1">'
    meta_xml += '<metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>'
    meta_xml += '</metadataTypes>'
    meta_xml += f'<futureMetadata name="XLRICHVALUE" count="{n}">'
    for i in range(n):
        meta_xml += f'<bk><extLst><ext uri="{GUID_RICHVALUE}"><xlrd:rvb i="{i}"/></ext></extLst></bk>'
    meta_xml += '</futureMetadata>'
    meta_xml += f'<valueMetadata count="{n}">'
    for i in range(n):
        meta_xml += f'<bk><rc t="1" v="{i}"/></bk>'
    meta_xml += '</valueMetadata>'
    meta_xml += '</metadata>'

    # --- xlsx ZIP を書き換え ---
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zin, \
         zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:

        # 既存ファイルのうち、上書き対象でないものをコピー
        skip = {
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
            "xl/metadata.xml",
            "xl/richData/richValueRel.xml",
            "xl/richData/_rels/richValueRel.xml.rels",
            "xl/richData/rdrichvalue.xml",
            "xl/richData/rdrichvaluestructure.xml",
            "xl/richData/rdRichValueTypes.xml",
        }

        _, sheet_path = _find_plot_sheet(zin)
        skip.add(sheet_path)

        # 既存 media ファイルも除外 (上書き)
        existing_media = {name for name in zin.namelist() if name.startswith("xl/media/")}
        skip.update(existing_media)

        for item in zin.namelist():
            if item not in skip:
                zout.writestr(item, zin.read(item))

        # シート XML を更新 (vm 属性を追加) - 正規表現で最小限の書き換え
        sheet_xml = zin.read(sheet_path).decode("utf-8")

        for col_letter, excel_row, idx in cells:
            cell_ref = f"{col_letter}{excel_row}"
            vm_val = idx + 1

            # セルが存在するか確認
            cell_pat = re.compile(
                rf'(<c\b[^>]*?\br="{re.escape(cell_ref)}"[^>]*?)(/>|>(.*?)</c>)',
                re.DOTALL,
            )
            m = cell_pat.search(sheet_xml)
            if m:
                tag_open = m.group(1)
                # 既存の t= を除去して vm, t を追加
                tag_open = re.sub(r'\s+t="[^"]*"', "", tag_open)
                tag_open = re.sub(r'\s+vm="[^"]*"', "", tag_open)
                tag_open += f' vm="{vm_val}" t="e"'
                sheet_xml = (sheet_xml[:m.start()]
                             + tag_open + '><v>#VALUE!</v></c>'
                             + sheet_xml[m.end():])

        zout.writestr(sheet_path, sheet_xml.encode("utf-8"))

        # Rich Data ファイルを書き出し
        zout.writestr("xl/richData/richValueRel.xml", rv_rel_xml.encode("utf-8"))
        zout.writestr("xl/richData/_rels/richValueRel.xml.rels", rv_rels_xml.encode("utf-8"))
        zout.writestr("xl/richData/rdrichvalue.xml", rv_data_xml.encode("utf-8"))
        zout.writestr("xl/richData/rdrichvaluestructure.xml", rv_struct_xml.encode("utf-8"))
        zout.writestr("xl/richData/rdRichValueTypes.xml", rv_types_xml.encode("utf-8"))
        zout.writestr("xl/metadata.xml", meta_xml.encode("utf-8"))

        # 画像ファイルを追加
        for i, img_data in enumerate(plot_images):
            zout.writestr(f"xl/media/image{i + 1}.png", img_data)

        # [Content_Types].xml を更新 (文字列操作)
        ct_xml = zin.read("[Content_Types].xml").decode("utf-8")
        if 'Extension="png"' not in ct_xml:
            ct_xml = ct_xml.replace("</Types>",
                '<Default Extension="png" ContentType="image/png"/></Types>')
        ct_overrides = {
            "/xl/metadata.xml":
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
            "/xl/richData/richValueRel.xml":
                "application/vnd.ms-excel.richvaluerel+xml",
            "/xl/richData/rdrichvalue.xml":
                "application/vnd.ms-excel.rdrichvalue+xml",
            "/xl/richData/rdrichvaluestructure.xml":
                "application/vnd.ms-excel.rdrichvaluestructure+xml",
            "/xl/richData/rdRichValueTypes.xml":
                "application/vnd.ms-excel.rdrichvaluetypes+xml",
        }
        for part, ctype in ct_overrides.items():
            if part not in ct_xml:
                ct_xml = ct_xml.replace("</Types>",
                    f'<Override PartName="{part}" ContentType="{ctype}"/></Types>')
        zout.writestr("[Content_Types].xml", ct_xml.encode("utf-8"))

        # workbook.xml.rels を更新 (文字列操作)
        wb_rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', wb_rels_xml)]
        max_rid = max(rids) if rids else 0

        rel_entries = [
            ("metadata.xml", REL_TYPE_METADATA),
            ("richData/richValueRel.xml", REL_TYPE_RICHVALREL),
            ("richData/rdrichvalue.xml", REL_TYPE_RICHVALUE),
            ("richData/rdrichvaluestructure.xml", REL_TYPE_RICHVALSTRUCT),
            ("richData/rdRichValueTypes.xml", REL_TYPE_RICHVALTYPES),
        ]
        for target, rel_type in rel_entries:
            if f'Target="{target}"' not in wb_rels_xml:
                max_rid += 1
                wb_rels_xml = wb_rels_xml.replace("</Relationships>",
                    f'<Relationship Id="rId{max_rid}" Type="{rel_type}" Target="{target}"/></Relationships>')
        zout.writestr("xl/_rels/workbook.xml.rels", wb_rels_xml.encode("utf-8"))

    return output.getvalue()


def validate_pdf(pdf_bytes: bytes) -> list[str]:
    """PDF の基本的な検証を行い、情報/警告リストを返す."""
    warnings = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if doc.page_count > 1:
        warnings.append(f"info:{doc.page_count} ページ検出 → {doc.page_count} サンプルとして処理します。")
    doc.close()
    return warnings


def validate_extracted_images(plots: list[bytes]) -> dict:
    """抽出した画像の品質を簡易チェックする."""
    blank_count = 0
    for img_bytes in plots:
        img = Image.open(io.BytesIO(img_bytes))
        extrema = img.convert("L").getextrema()
        if extrema[1] - extrema[0] < 10:  # ほぼ単色 → 空白と判定
            blank_count += 1
    return {
        "total": len(plots),
        "blank_count": blank_count,
    }


def create_preview_grid(plots: list[bytes], cols: int = NUM_COLS,
                        rows: int = 0,
                        thumb_size: tuple[int, int] = (120, 100)) -> bytes:
    """プロット画像のグリッドプレビュー (PNG) を生成する."""
    if rows == 0:
        rows = (len(plots) + cols - 1) // cols
    tw, th = thumb_size
    pad = 4
    grid_w = cols * (tw + pad) + pad
    grid_h = rows * (th + pad) + pad
    grid = Image.new("RGB", (grid_w, grid_h), "white")

    for i, img_bytes in enumerate(plots):
        r, c = divmod(i, cols)
        img = Image.open(io.BytesIO(img_bytes))
        img.thumbnail(thumb_size)
        x = pad + c * (tw + pad)
        y = pad + r * (th + pad)
        grid.paste(img, (x, y))

    buf = io.BytesIO()
    grid.save(buf, format="PNG")
    return buf.getvalue()


def validate_excel(xlsx_bytes: bytes) -> list[str]:
    """Excel ファイルの構造を検証し、警告リストを返す."""
    warnings = []
    try:
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zf:
            name, _ = _find_plot_sheet(zf)
            if "Plot" not in name:
                warnings.append(f"'Plot' シートが見つからず、代わりに '{name}' を使用します。")
    except Exception as e:
        warnings.append(f"Excel の構造読み取りに失敗: {e}")
    return warnings


def process(pdf_bytes: bytes, xlsx_bytes: bytes) -> bytes:
    """PDF からプロットを切り出し、Excel に画像を挿入する."""
    plots, num_rows, num_cols = extract_plots_from_pdf(pdf_bytes)
    return insert_images_to_xlsx(xlsx_bytes, plots, num_rows, num_cols)
