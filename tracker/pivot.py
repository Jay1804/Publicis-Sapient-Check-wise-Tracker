"""Inject a native Excel PivotTable into the tracker workbook.

Pivot: rows = Check Unique Name, columns = Check Severity, values = Count of Check Severity,
placed on the "Summary" sheet and sourced from the 'Publicis Sapient' sheet. xlsxwriter cannot
create pivots, so the pivot parts are added to the package. The cache refreshes on open
(refreshOnLoad); the static numbers written by pipeline.to_excel_bytes show until then.
"""
import io
import re
import zipfile
from xml.sax.saxutils import quoteattr

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def add_pivot(xlsx, columns, summary, nrows):
    row_names = [r for r in summary.index if r != "Grand Total"]
    col_names = [c for c in summary.columns if c != "Grand Total"]
    r_fld, c_fld = columns.index("Check Unique Name"), columns.index("Check Severity")

    def shared(names):
        return f'<sharedItems count="{len(names)}">' + "".join(f"<s v={quoteattr(str(n))}/>" for n in names) + "</sharedItems>"

    fields = []
    for i, name in enumerate(columns):
        si = shared(row_names) if i == r_fld else shared(col_names) if i == c_fld else "<sharedItems/>"
        fields.append(f"<cacheField name={quoteattr(name)} numFmtId=\"0\">{si}</cacheField>")
    cache_xml = (
        HDR + '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'xmlns:r="{REL}" refreshOnLoad="1" saveData="0" createdVersion="8" refreshedVersion="8" '
        f'minRefreshableVersion="3" recordCount="{nrows}"><cacheSource type="worksheet">'
        f'<worksheetSource ref="A1:{_col_letter(len(columns))}{nrows + 1}" sheet="Publicis Sapient"/></cacheSource>'
        f'<cacheFields count="{len(columns)}">{"".join(fields)}</cacheFields></pivotCacheDefinition>')

    def idx_items(n):
        return "".join(f'<item x="{i}"/>' for i in range(n)) + '<item t="default"/>'

    pfs = []
    for i in range(len(columns)):
        if i == r_fld:
            pfs.append(f'<pivotField axis="axisRow" showAll="0"><items count="{len(row_names) + 1}">{idx_items(len(row_names))}</items></pivotField>')
        elif i == c_fld:
            pfs.append(f'<pivotField axis="axisCol" dataField="1" showAll="0"><items count="{len(col_names) + 1}">{idx_items(len(col_names))}</items></pivotField>')
        else:
            pfs.append('<pivotField showAll="0"/>')

    def line(n):
        return "".join("<i><x/></i>" if k == 0 else f'<i><x v="{k}"/></i>' for k in range(n)) + '<i t="grand"><x/></i>'

    last_row = len(row_names) + 3
    pt_xml = (
        HDR + '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" '
        'applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" '
        'updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="0" preserveFormatting="1" itemPrintTitles="1" createdVersion="8" '
        'indent="0" outline="1" outlineData="1" multipleFieldFilters="0" rowHeaderCaption="Check Unique Name">'
        f'<location ref="A1:{_col_letter(len(col_names) + 2)}{last_row}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>'
        f'<pivotFields count="{len(columns)}">{"".join(pfs)}</pivotFields>'
        f'<rowFields count="1"><field x="{r_fld}"/></rowFields><rowItems count="{len(row_names) + 1}">{line(len(row_names))}</rowItems>'
        f'<colFields count="1"><field x="{c_fld}"/></colFields><colItems count="{len(col_names) + 1}">{line(len(col_names))}</colItems>'
        f'<dataFields count="1"><dataField name="Count of Check Severity" fld="{c_fld}" subtotal="count" baseField="0" baseItem="0"/></dataFields>'
        '<pivotTableStyleInfo name="PivotStyleMedium2" showRowHeaders="1" showColHeaders="1" showRowStripes="1" '
        'showColStripes="0" showLastColumn="1"/></pivotTableDefinition>')

    zin = zipfile.ZipFile(io.BytesIO(xlsx))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            data = zin.read(n)
            if n == "[Content_Types].xml":
                data = data.decode("utf8").replace(
                    "</Types>",
                    '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
                    '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>').encode("utf8")
            elif n == "xl/_rels/workbook.xml.rels":
                data = data.decode("utf8").replace(
                    "</Relationships>",
                    f'<Relationship Id="rIdPivot1" Type="{REL}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>').encode("utf8")
            elif n == "xl/workbook.xml":
                t = data.decode("utf8")
                tag = '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivot1"/></pivotCaches>'
                m = re.search(r"<calcPr[^>]*/>", t)
                t = t[:m.end()] + tag + t[m.end():] if m else t.replace("</workbook>", tag + "</workbook>")
                if "xmlns:r=" not in t:
                    t = t.replace("<workbook ", f'<workbook xmlns:r="{REL}" ', 1)
                data = t.encode("utf8")
            zout.writestr(n, data)
        zout.writestr("xl/worksheets/_rels/sheet1.xml.rels",
                      HDR + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>')
        zout.writestr("xl/pivotCache/pivotCacheDefinition1.xml", cache_xml)
        zout.writestr("xl/pivotTables/pivotTable1.xml", pt_xml)
        zout.writestr("xl/pivotTables/_rels/pivotTable1.xml.rels",
                      HDR + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>')
    return out.getvalue()
