"""FlowJo PDF からプロット画像を切り出し、Excel に IMAGE 関数として挿入するモジュール."""

import io
import re
import zipfile

import fitz
from lxml import etree
from PIL import Image


DPI = 500
NUM_ROWS = 14
NUM_COLS = 8

# 自動検出済みのプロット境界 (500 DPI ピクセル座標)
COL_STARTS = [962, 1270, 1582, 1891, 2199, 2508, 2815, 3123]
COL_ENDS = [1233, 1541, 1849, 2157, 2465, 2775, 3081, 3389]
ROW_STARTS = [520, 841, 1163, 1484, 1805, 2126, 2447, 2768, 3089, 3410, 3732, 4053, 4374, 4695]
ROW_ENDS = [783, 1104, 1426, 1747, 2068, 2389, 2710, 3031, 3352, 3674, 3995, 4316, 4637, 4958]

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
REL_TYPE_RICHVALUE = "http://schemas.microsoft.com/office/2022/10/relationships/rdRichValue"
REL_TYPE_RICHVALREL = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel"
REL_TYPE_RICHVALSTRUCT = "http://schemas.microsoft.com/office/2022/10/relationships/rdRichValueStructure"
REL_TYPE_RICHVALTYPES = "http://schemas.microsoft.com/office/2022/10/relationships/rdRichValueTypes"
REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

GUID_RICHVALUE = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"

# 列文字 → 列番号 (1-indexed)
_COL_MAP = {}
for _i, _c in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
    _COL_MAP[_c] = _i + 1


def _col_to_num(col_letter: str) -> int:
    return _COL_MAP[col_letter]


def extract_plots_from_pdf(pdf_bytes: bytes) -> list[bytes]:
    """PDF から 112枚のプロット画像を均一サイズで切り出す."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    scale = DPI / 72
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat)
    img = Image.open(io.BytesIO(pix.tobytes("png")))

    plots = []
    for row in range(NUM_ROWS):
        for col in range(NUM_COLS):
            crop = img.crop((
                COL_STARTS[col], ROW_STARTS[row],
                COL_ENDS[col], ROW_ENDS[row],
            ))

            buf = io.BytesIO()
            crop.save(buf, format="PNG")
            plots.append(buf.getvalue())

    doc.close()
    return plots


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


def insert_images_to_xlsx(xlsx_bytes: bytes, plot_images: list[bytes]) -> bytes:
    """Excel の Plot シートに IMAGE 関数 (Rich Data) として画像を挿入する."""
    n = len(plot_images)

    # 画像セルの位置を生成
    cells = []
    for row_i in range(NUM_ROWS):
        excel_row = EXCEL_ROW_START + row_i * EXCEL_ROW_STEP
        for col_i in range(NUM_COLS):
            idx = row_i * NUM_COLS + col_i
            if idx >= n:
                break
            cells.append((EXCEL_COLS[col_i], excel_row, idx))

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
    rv_data = etree.Element(f"{{{NS_RICHDATA}}}rvData", count=str(n))
    for i in range(n):
        rv = etree.SubElement(rv_data, f"{{{NS_RICHDATA}}}rv", s="0")
        v1 = etree.SubElement(rv, f"{{{NS_RICHDATA}}}v")
        v1.text = str(i)
        v2 = etree.SubElement(rv, f"{{{NS_RICHDATA}}}v")
        v2.text = "5"

    # 4. rdrichvaluestructure.xml
    rv_struct = etree.Element(f"{{{NS_RICHDATA}}}rvStructures", count="1")
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

        # シート XML を更新 (vm 属性を追加)
        sheet_xml = zin.read(sheet_path)
        sheet_root = etree.fromstring(sheet_xml)
        ns = {"x": NS_SHEET}

        # セル→vm マッピング
        vm_map = {}
        for col_letter, excel_row, idx in cells:
            cell_ref = f"{col_letter}{excel_row}"
            vm_map[cell_ref] = idx + 1  # 1-indexed

        # 既存の行を走査してセルに vm を付与、なければ作成
        sheet_data = sheet_root.find(f"{{{NS_SHEET}}}sheetData")
        existing_rows = {int(r.get("r")): r for r in sheet_data.findall(f"{{{NS_SHEET}}}row")}

        for col_letter, excel_row, idx in cells:
            cell_ref = f"{col_letter}{excel_row}"
            col_num = _col_to_num(col_letter)

            row_el = existing_rows.get(excel_row)
            if row_el is None:
                row_el = etree.SubElement(sheet_data, f"{{{NS_SHEET}}}row", r=str(excel_row))
                existing_rows[excel_row] = row_el

            # セル要素を探す
            cell_el = None
            for c in row_el.findall(f"{{{NS_SHEET}}}c"):
                if c.get("r") == cell_ref:
                    cell_el = c
                    break

            if cell_el is None:
                cell_el = etree.SubElement(row_el, f"{{{NS_SHEET}}}c", r=cell_ref)

            cell_el.set("vm", str(idx + 1))
            # IMAGE 関数のセルには t="e" (error) と v 要素が必要
            cell_el.set("t", "e")
            v_el = cell_el.find(f"{{{NS_SHEET}}}v")
            if v_el is None:
                v_el = etree.SubElement(cell_el, f"{{{NS_SHEET}}}v")
            v_el.text = "#VALUE!"

        zout.writestr(sheet_path,
                      etree.tostring(sheet_root, xml_declaration=True,
                                     encoding="UTF-8", standalone=True))

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
    page = doc[0]
    if page.rect.width < page.rect.height:
        warnings.append("PDF が縦長です。FlowJo バッチ出力は通常横長です。正しいファイルか確認してください。")
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
                        rows: int = NUM_ROWS,
                        thumb_size: tuple[int, int] = (120, 100)) -> bytes:
    """プロット画像のグリッドプレビュー (PNG) を生成する."""
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
    plots = extract_plots_from_pdf(pdf_bytes)
    return insert_images_to_xlsx(xlsx_bytes, plots)
