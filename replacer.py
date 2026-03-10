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


def _detect_grid(img: Image.Image) -> tuple[list[tuple[int, int]], list[tuple[int, int]]]:
    """画像からプロットグリッドの行・列バンドを自動検出する.

    Returns:
        (row_bands, col_bands) - 各バンドは (start, end) のピクセル座標タプル.
        col_bands はラベル列を除いたプロット列のみ.
    """
    arr = np.array(img.convert("L"))
    h, w = arr.shape
    content = arr < 240  # non-white pixels

    # 行バンド検出: 各行の非白ピクセル数
    row_proj = content.sum(axis=1)
    row_threshold = w * 0.02
    content_rows = np.where(row_proj > row_threshold)[0]

    row_bands = []
    if len(content_rows) > 0:
        start = int(content_rows[0])
        prev = start
        for r in content_rows[1:]:
            if r - prev > 15:
                row_bands.append((start, int(prev)))
                start = int(r)
            prev = int(r)
        row_bands.append((start, int(prev)))

    # 列バンド検出: 各列の非白ピクセル数
    col_proj = content.sum(axis=0)
    col_threshold = h * 0.02
    content_cols = np.where(col_proj > col_threshold)[0]

    col_bands_all = []
    if len(content_cols) > 0:
        start = int(content_cols[0])
        prev = start
        for c in content_cols[1:]:
            if c - prev > 15:
                col_bands_all.append((start, int(prev)))
                start = int(c)
            prev = int(c)
        col_bands_all.append((start, int(prev)))

    # 狭いバンド（軸ラベル等）を除外: 最大幅の50%未満のバンドを除去
    col_bands = col_bands_all
    if len(col_bands_all) > 2:
        widths = [e - s for s, e in col_bands_all]
        max_width = max(widths)
        col_bands = [(s, e) for s, e in col_bands_all
                     if (e - s) >= max_width * 0.5]

    return row_bands, col_bands


def extract_plots_from_pdf(pdf_bytes: bytes) -> tuple[list[bytes], int, int]:
    """PDF からプロット画像を自動検出・切り出す.

    Returns:
        (plots, num_rows, num_cols) - 画像バイトのリスト、検出された行数・列数.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    # 低DPIでグリッド検出
    detect_scale = DETECT_DPI / 72
    det_mat = fitz.Matrix(detect_scale, detect_scale)
    det_pix = page.get_pixmap(matrix=det_mat)
    det_img = Image.open(io.BytesIO(det_pix.tobytes("png")))

    row_bands, col_bands = _detect_grid(det_img)
    num_rows = len(row_bands)
    num_cols = len(col_bands)

    # 高DPIで切り出し（座標をスケール）
    ratio = RENDER_DPI / DETECT_DPI
    render_scale = RENDER_DPI / 72
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


def _get_max_rid(rels_root: etree._Element) -> int:
    """rels ファイル内の最大 rId 番号を返す."""
    max_id = 0
    for rel in rels_root:
        rid = rel.get("Id", "")
        m = re.search(r"(\d+)", rid)
        if m:
            max_id = max(max_id, int(m.group(1)))
    return max_id


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

    # --- Rich Data XML を生成 ---

    # 1. richValueRel.xml
    rv_rel = etree.Element(f"{{{NS_RICHVALREL}}}richValueRels",
                           nsmap={None: NS_RICHVALREL, "r": NS_REL})
    for i in range(n):
        el = etree.SubElement(rv_rel, f"{{{NS_RICHVALREL}}}rel")
        el.set(f"{{{NS_REL}}}id", f"rId{i + 1}")

    # 2. richValueRel.xml.rels
    rv_rels = etree.Element("Relationships", xmlns=NS_PKG_REL)
    for i in range(n):
        etree.SubElement(rv_rels, "Relationship",
                         Id=f"rId{i + 1}",
                         Type=REL_TYPE_IMAGE,
                         Target=f"../media/image{i + 1}.png")

    # 3. rdrichvalue.xml
    rv_data = etree.Element(f"{{{NS_RICHDATA}}}rvData", count=str(n),
                            nsmap={None: NS_RICHDATA})
    for i in range(n):
        rv = etree.SubElement(rv_data, f"{{{NS_RICHDATA}}}rv", s="0")
        v1 = etree.SubElement(rv, f"{{{NS_RICHDATA}}}v")
        v1.text = str(i)
        v2 = etree.SubElement(rv, f"{{{NS_RICHDATA}}}v")
        v2.text = "5"

    # 4. rdrichvaluestructure.xml
    rv_struct = etree.Element(f"{{{NS_RICHDATA}}}rvStructures", count="1",
                              nsmap={None: NS_RICHDATA})
    s_el = etree.SubElement(rv_struct, f"{{{NS_RICHDATA}}}s", t="_localImage")
    etree.SubElement(s_el, f"{{{NS_RICHDATA}}}k", n="_rvRel:LocalImageIdentifier", t="i")
    etree.SubElement(s_el, f"{{{NS_RICHDATA}}}k", n="CalcOrigin", t="i")

    # 5. rdRichValueTypes.xml
    rv_types = etree.Element(f"{{{NS_RICHDATA2}}}rvTypesInfo",
                             nsmap={None: NS_RICHDATA2, "mc": NS_MC, "x": NS_SHEET})
    rv_types.set(f"{{{NS_MC}}}Ignorable", "x")
    g = etree.SubElement(rv_types, f"{{{NS_RICHDATA2}}}global")
    kf = etree.SubElement(g, f"{{{NS_RICHDATA2}}}keyFlags")
    for kname in ["_Self", "_DisplayString", "_Flags", "_Format",
                   "_SubLabel", "_Attribution", "_Icon", "_Display",
                   "_CanonicalPropertyNames", "_ClassificationId"]:
        k = etree.SubElement(kf, f"{{{NS_RICHDATA2}}}key", name=kname)
        if kname == "_Self":
            etree.SubElement(k, f"{{{NS_RICHDATA2}}}flag",
                             name="ExcludeFromFile", value="1")
        etree.SubElement(k, f"{{{NS_RICHDATA2}}}flag",
                         name="ExcludeFromCalcComparison", value="1")

    # 6. metadata.xml
    meta = etree.Element(f"{{{NS_SHEET}}}metadata",
                         nsmap={None: NS_SHEET, "xlrd": NS_RICHDATA})

    mt = etree.SubElement(meta, f"{{{NS_SHEET}}}metadataTypes", count="1")
    etree.SubElement(mt, f"{{{NS_SHEET}}}metadataType",
                     name="XLRICHVALUE", minSupportedVersion="120000",
                     copy="1", pasteAll="1", pasteValues="1", merge="1",
                     splitFirst="1", rowColShift="1", clearFormats="1",
                     clearComments="1", assign="1", coerce="1")

    fm = etree.SubElement(meta, f"{{{NS_SHEET}}}futureMetadata",
                          name="XLRICHVALUE", count=str(n))
    for i in range(n):
        bk = etree.SubElement(fm, f"{{{NS_SHEET}}}bk")
        ext_lst = etree.SubElement(bk, f"{{{NS_SHEET}}}extLst")
        ext = etree.SubElement(ext_lst, f"{{{NS_SHEET}}}ext", uri=GUID_RICHVALUE)
        etree.SubElement(ext, f"{{{NS_RICHDATA}}}rvb", i=str(i))

    vm = etree.SubElement(meta, f"{{{NS_SHEET}}}valueMetadata", count=str(n))
    for i in range(n):
        bk = etree.SubElement(vm, f"{{{NS_SHEET}}}bk")
        etree.SubElement(bk, f"{{{NS_SHEET}}}rc", t="1", v=str(i))

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
        def _to_xml(root_el):
            return etree.tostring(root_el, xml_declaration=True,
                                  encoding="UTF-8", standalone=True)

        zout.writestr("xl/richData/richValueRel.xml", _to_xml(rv_rel))
        zout.writestr("xl/richData/_rels/richValueRel.xml.rels", _to_xml(rv_rels))
        zout.writestr("xl/richData/rdrichvalue.xml", _to_xml(rv_data))
        zout.writestr("xl/richData/rdrichvaluestructure.xml", _to_xml(rv_struct))
        zout.writestr("xl/richData/rdRichValueTypes.xml", _to_xml(rv_types))
        zout.writestr("xl/metadata.xml", _to_xml(meta))

        # 画像ファイルを追加
        for i, img_data in enumerate(plot_images):
            zout.writestr(f"xl/media/image{i + 1}.png", img_data)

        # [Content_Types].xml を更新
        ct_root = etree.fromstring(zin.read("[Content_Types].xml"))
        ct_ns = NS_CONTENT

        # png Default がなければ追加
        has_png = any(
            el.get("Extension") == "png"
            for el in ct_root.findall(f"{{{ct_ns}}}Default")
        )
        if not has_png:
            etree.SubElement(ct_root, f"{{{ct_ns}}}Default",
                             Extension="png", ContentType="image/png")

        # Override エントリ追加
        overrides = {
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
        existing_parts = {
            el.get("PartName")
            for el in ct_root.findall(f"{{{ct_ns}}}Override")
        }
        for part, ctype in overrides.items():
            if part not in existing_parts:
                etree.SubElement(ct_root, f"{{{ct_ns}}}Override",
                                 PartName=part, ContentType=ctype)

        zout.writestr("[Content_Types].xml", _to_xml(ct_root))

        # workbook.xml.rels を更新
        wb_rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        max_rid = _get_max_rid(wb_rels)
        existing_targets = {rel.get("Target") for rel in wb_rels}

        rel_entries = [
            ("metadata.xml", REL_TYPE_METADATA),
            ("richData/richValueRel.xml", REL_TYPE_RICHVALREL),
            ("richData/rdrichvalue.xml", REL_TYPE_RICHVALUE),
            ("richData/rdrichvaluestructure.xml", REL_TYPE_RICHVALSTRUCT),
            ("richData/rdRichValueTypes.xml", REL_TYPE_RICHVALTYPES),
        ]
        for target, rel_type in rel_entries:
            if target not in existing_targets:
                max_rid += 1
                etree.SubElement(wb_rels, "Relationship",
                                 Id=f"rId{max_rid}",
                                 Type=rel_type,
                                 Target=target)

        zout.writestr("xl/_rels/workbook.xml.rels", _to_xml(wb_rels))

    return output.getvalue()


def validate_pdf(pdf_bytes: bytes) -> list[str]:
    """PDF の基本的な検証を行い、警告リストを返す."""
    warnings = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if doc.page_count != 1:
        warnings.append(f"PDF が {doc.page_count} ページあります（期待値: 1ページ）。先頭ページのみ処理します。")
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
