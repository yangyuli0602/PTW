# -*- coding: utf-8 -*-
"""往已生成的 xlsx 里插一张 PNG。

为什么要手工改 zip：内置运行时没有 Pillow，openpyxl 的 `add_image` 一调就 ImportError。
而"汇率来源截图"是审计底稿的必需件——复核人要看到汇率是从哪份公告、哪一天取的，
光写个网址不够（网址会失效、公告会下架）。

做法：拆开 xlsx（它就是个 zip），补三样东西——
  xl/media/image1.png          图片本体
  xl/drawings/drawing1.xml     图片的锚点与尺寸
  xl/worksheets/_rels/sheetN.xml.rels   把 sheet 指向 drawing
再改两处——sheet XML 里加 <drawing r:id="…"/>，[Content_Types].xml 里登记 png 与 drawing。

**失败就放弃插图，绝不让文件坏掉**：调用方拿到 None 就用原 bytes。
一张图不值得让整份底稿打不开。
"""

import io
import re
import zipfile

_NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_DRAWING = ("application/vnd.openxmlformats-officedocument."
               "drawing+xml")


def png_size(data):
    """从 PNG 头读宽高（IHDR 第 16-24 字节），不需要 PIL。"""
    if len(data) < 24 or data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("不是 PNG")
    return (int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big"))


def _sheet_path(zf, title):
    """工作表名 → xl/worksheets/sheetN.xml 的路径。"""
    wb = zf.read("xl/workbook.xml").decode("utf-8", errors="replace")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="replace")
    rid = None
    for m in re.finditer(r"<sheet[^>]*?/>", wb):
        tag = m.group(0)
        nm = re.search(r'name="([^"]*)"', tag)
        ri = re.search(r'r:id="([^"]*)"', tag) or re.search(r'id="([^"]*)"', tag)
        if nm and ri and nm.group(1) == title:
            rid = ri.group(1)
            break
    if not rid:
        return None
    for m in re.finditer(r"<Relationship[^>]*?/>", rels):
        tag = m.group(0)
        if 'Id="%s"' % rid in tag:
            t = re.search(r'Target="([^"]*)"', tag)
            if t:
                p = t.group(1)
                return p[1:] if p.startswith("/") else "xl/" + p
    return None


def _append_to_existing(src, names, sheet_path, sheet_xml, sheet_rid,
                        png_bytes, cx, cy, anchor_row, anchor_col):
    """复用工作表已有的 drawing：追加一个 oneCellAnchor + 图片关系 + media。"""
    # sheet 的 rels → drawing 路径
    srel = sheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
    if srel not in names:
        return None
    rels = src.read(srel).decode("utf-8")
    m = re.search(r'<Relationship[^>]*Id="%s"[^>]*/>' % sheet_rid, rels)
    if not m:
        return None
    t = re.search(r'Target="([^"]*)"', m.group(0))
    if not t:
        return None
    target = t.group(1)
    draw_path = target[1:] if target.startswith("/") else \
        ("xl/" + target.replace("../", ""))
    if draw_path not in names:
        return None
    draw_xml = src.read(draw_path).decode("utf-8")
    drel_path = draw_path.replace("drawings/", "drawings/_rels/") + ".rels"
    if drel_path in names:
        drels = src.read(drel_path).decode("utf-8")
    else:
        drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="%s"></Relationships>' % _NS_PKG_REL)
    used = [int(x) for x in re.findall(r'Id="rId(\d+)"', drels)] or [0]
    rid = "rId%d" % (max(used) + 1)
    # media 编号：避开已占用的 imageN
    n = 1
    while "xl/media/image%d.png" % n in names:
        n += 1
    media_path = "xl/media/image%d.png" % n
    drels = drels.replace("</Relationships>",
        '<Relationship Id="%s" Type="%s/image" Target="../media/image%d.png"/>'
        "</Relationships>" % (rid, _NS_R, n))

    frm_col = max(0, anchor_col - 1)
    frm_row = max(0, anchor_row - 1)
    anchor = (
        '<xdr:oneCellAnchor>'
        '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
        '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        '<xdr:ext cx="%d" cy="%d"/>'
        '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="%d" name="汇率来源截图"/>'
        '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
        '<xdr:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
        '</xdr:pic><xdr:clientData/></xdr:oneCellAnchor>'
        % (frm_col, frm_row, cx, cy, max(used) + 10, rid, cx, cy))
    if "</xdr:wsDr>" not in draw_xml:
        return None
    draw_xml = draw_xml.replace("</xdr:wsDr>", anchor + "</xdr:wsDr>", 1)

    ct = src.read("[Content_Types].xml").decode("utf-8")
    if 'Extension="png"' not in ct:
        ct = ct.replace("<Default", '<Default Extension="png" ContentType="image/png"/><Default', 1)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for nm in names:
            if nm == draw_path:
                dst.writestr(nm, draw_xml)
            elif nm == drel_path:
                dst.writestr(nm, drels)
            elif nm == "[Content_Types].xml":
                dst.writestr(nm, ct)
            else:
                dst.writestr(nm, src.read(nm))
        if drel_path not in names:
            dst.writestr(drel_path, drels)
        dst.writestr(media_path, png_bytes)
    zipfile.ZipFile(io.BytesIO(out.getvalue())).namelist()
    return out.getvalue()


def add_png(xlsx_bytes, sheet_title, png_bytes, anchor_row=18, anchor_col=2,
            max_width_px=820):
    """把 PNG 插到指定工作表。

    anchor_row / anchor_col 从 1 开始（与 Excel 一致）。
    返回新的 bytes；任何一步失败返回 None（调用方沿用原文件）。
    """
    try:
        w, h = png_size(png_bytes)
        scale = min(1.0, float(max_width_px) / w) if w else 1.0
        cx = int(w * scale * 9525)          # EMU = px × 9525
        cy = int(h * scale * 9525)

        src = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
        names = src.namelist()
        sheet_path = _sheet_path(src, sheet_title)
        if not sheet_path or sheet_path not in names:
            return None
        sheet_xml = src.read(sheet_path).decode("utf-8")

        # 同一份文件被插两次会报坏图：按图片名判断是否已插过
        if 'name="汇率来源截图"' in sheet_xml:
            return None

        # 模板的 C.02a 页本来就可能带 drawing（占位图 / 水印）。
        # 这种情况不能放弃，也不能新建第二个 drawing（Excel 只认一个），
        # 要往**已有**的 drawing 里追加一个图片锚点。
        _exist = re.search(r'<drawing[^>]*r:id="(rId\d+)"\s*/>', sheet_xml)
        if _exist:
            return _append_to_existing(src, names, sheet_path, sheet_xml,
                                       _exist.group(1), png_bytes, cx, cy,
                                       anchor_row, anchor_col)

        # 找一个没被占用的 drawing / media 编号
        idx = 1
        while "xl/drawings/drawing%d.xml" % idx in names:
            idx += 1
        draw_path = "xl/drawings/drawing%d.xml" % idx
        draw_rel = "xl/drawings/_rels/drawing%d.xml.rels" % idx
        media_path = "xl/media/image%d.png" % idx

        # sheet 的 rels：可能已存在（有超链接等）
        sheet_rel_path = sheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if sheet_rel_path in names:
            rels_xml = src.read(sheet_rel_path).decode("utf-8")
        else:
            rels_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        '<Relationships xmlns="%s"></Relationships>' % _NS_PKG_REL)
        used = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)] or [0]
        sheet_rid = "rId%d" % (max(used) + 1)
        rels_xml = rels_xml.replace("</Relationships>",
            '<Relationship Id="%s" Type="%s/drawing" Target="../drawings/drawing%d.xml"/>'
            "</Relationships>" % (sheet_rid, _NS_R, idx))

        # 挂 drawing。两个坑：
        # 1) openpyxl 把 xmlns:r 声明挂在 <legacyDrawing> 元素上（局部声明），不是根元素。
        #    写在它后面的 <drawing r:id> 继承不到，会报 unbound prefix / 文件损坏。
        #    → 自己带上 xmlns:r。
        # 2) CT_Worksheet 的元素顺序要求 drawing 在 legacyDrawing 之前，
        #    否则 Excel 会提示"无法读取的内容"。→ 插在 legacyDrawing 前面。
        _tag = '<drawing xmlns:r="%s" r:id="%s"/>' % (_NS_R, sheet_rid)
        if "<legacyDrawing" in sheet_xml:
            sheet_xml = sheet_xml.replace("<legacyDrawing", _tag + "<legacyDrawing", 1)
        else:
            sheet_xml = sheet_xml.replace("</worksheet>", _tag + "</worksheet>")

        # drawing XML
        frm_col = max(0, anchor_col - 1)
        frm_row = max(0, anchor_row - 1)
        drawing = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<xdr:wsDr xmlns:xdr="%s" xmlns:a="%s" xmlns:r="%s">'
            '<xdr:oneCellAnchor>'
            '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:ext cx="%d" cy="%d"/>'
            '<xdr:pic>'
            '<xdr:nvPicPr><xdr:cNvPr id="2" name="汇率来源截图"/>'
            '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
            '<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
            '</xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>'
            % (_NS_XDR, _NS_A, _NS_R, frm_col, frm_row, cx, cy, cx, cy))
        drawing_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="%s">'
            '<Relationship Id="rId1" Type="%s/image" Target="../media/image%d.png"/>'
            "</Relationships>" % (_NS_PKG_REL, _NS_R, idx))

        # [Content_Types].xml：登记 png 与 drawing
        ct = src.read("[Content_Types].xml").decode("utf-8")
        if 'Extension="png"' not in ct:
            ct = ct.replace("<Default", '<Default Extension="png" ContentType="image/png"/><Default', 1)
        ct = ct.replace("</Types>",
            '<Override PartName="/%s" ContentType="%s"/></Types>' % (draw_path, _CT_DRAWING))

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
            for n in names:
                if n == sheet_path:
                    dst.writestr(n, sheet_xml)
                elif n == "[Content_Types].xml":
                    dst.writestr(n, ct)
                elif n == sheet_rel_path:
                    dst.writestr(n, rels_xml)
                else:
                    dst.writestr(n, src.read(n))
            if sheet_rel_path not in names:
                dst.writestr(sheet_rel_path, rels_xml)
            dst.writestr(draw_path, drawing)
            dst.writestr(draw_rel, drawing_rels)
            dst.writestr(media_path, png_bytes)
        # 自检：能被 zip 正常打开
        zipfile.ZipFile(io.BytesIO(out.getvalue())).namelist()
        return out.getvalue()
    except Exception:
        return None
