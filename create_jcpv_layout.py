from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Helpers ──────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    hex_color = hex_color.lstrip('#')
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for shd in tcPr.findall(qn('w:shd')):
        tcPr.remove(shd)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_margins(cell, top=0, left=0, bottom=0, right=0):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top', top), ('left', left), ('bottom', bottom), ('right', right)]:
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'), str(val))
        m.set(qn('w:type'), 'dxa')
        tcMar.append(m)
    tcPr.append(tcMar)

def set_row_height(row, height_twips, exact=True):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(height_twips))
    trHeight.set(qn('w:hRule'), 'exact' if exact else 'atLeast')
    trPr.append(trHeight)

def no_borders(table):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{name}')
        b.set(qn('w:val'), 'none')
        b.set(qn('w:sz'), '0')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), 'auto')
        tblBorders.append(b)
    tblPr.append(tblBorders)

def set_table_width(table, twips):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(twips))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

def no_space(para):
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)

def gold_line(cell):
    """Add a bottom-border gold line to a paragraph inside a cell."""
    p = cell.add_paragraph()
    no_space(p)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single')
    bot.set(qn('w:sz'), '4')
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), 'B89B6A')
    pBdr.append(bot)
    pPr.append(pBdr)
    return p

# ── Constants ─────────────────────────────────────────────────────────────────

PAGE_W = 11906   # A4 width in twips
DARK   = '1E1E1E'
GOLD   = 'B89B6A'
NAVY   = '1B3A5C'
WHITE  = 'FFFFFF'
LIGHT  = 'F5F5F0'
DARK2  = '2C2C2C'

DARK_RGB  = RGBColor(0x1E, 0x1E, 0x1E)
GOLD_RGB  = RGBColor(0xB8, 0x9B, 0x6A)
NAVY_RGB  = RGBColor(0x1B, 0x3A, 0x5C)
WHITE_RGB = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_RGB  = RGBColor(0x66, 0x66, 0x66)

# ── Document ──────────────────────────────────────────────────────────────────

doc = Document()
sec = doc.sections[0]
sec.page_height = Cm(29.7)
sec.page_width  = Cm(21)
sec.top_margin    = Cm(0)
sec.bottom_margin = Cm(0)
sec.left_margin   = Cm(0)
sec.right_margin  = Cm(0)

# Remove default paragraph spacing
style = doc.styles['Normal']
style.paragraph_format.space_before = Pt(0)
style.paragraph_format.space_after  = Pt(0)

# ── 1. HEADER: fundo escuro com logo ─────────────────────────────────────────

ht = doc.add_table(rows=1, cols=1)
no_borders(ht)
set_table_width(ht, PAGE_W)
hc = ht.rows[0].cells[0]
set_cell_bg(hc, DARK)
set_cell_margins(hc, top=520, left=720, bottom=420, right=720)
set_row_height(ht.rows[0], 2100, exact=False)

# Ícone balança
p = hc.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
no_space(p)
r = p.add_run('⚖')
r.font.size = Pt(30)
r.font.color.rgb = GOLD_RGB
r.font.name = 'Segoe UI'

# JCPV
p2 = hc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
no_space(p2)
p2.paragraph_format.space_before = Pt(4)
r2 = p2.add_run('J C P V')
r2.font.size = Pt(28)
r2.font.bold = True
r2.font.color.rgb = GOLD_RGB
r2.font.name = 'Garamond'

# ADVOCACIA
p3 = hc.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
no_space(p3)
p3.paragraph_format.space_before = Pt(2)
r3 = p3.add_run('A  D  V  O  C  A  C  I  A')
r3.font.size = Pt(9)
r3.font.color.rgb = GOLD_RGB
r3.font.name = 'Garamond'
r3.font.bold = False

# Slogan abaixo do nome
p4 = hc.add_paragraph()
p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
no_space(p4)
p4.paragraph_format.space_before = Pt(8)
r4 = p4.add_run('A defesa técnica do seu patrimônio.')
r4.font.size = Pt(9)
r4.font.color.rgb = RGBColor(0xC8, 0xB8, 0x99)
r4.font.name = 'Garamond'
r4.font.italic = True

# ── 2. BARRA DOURADA ─────────────────────────────────────────────────────────

def gold_bar(width=PAGE_W, height=90):
    t = doc.add_table(rows=1, cols=1)
    no_borders(t)
    set_table_width(t, width)
    c = t.rows[0].cells[0]
    set_cell_bg(c, GOLD)
    set_row_height(t.rows[0], height, exact=True)
    c.paragraphs[0].text = ''

gold_bar()

# ── 3. BARRA DE CREDENCIAIS ──────────────────────────────────────────────────

ct = doc.add_table(rows=1, cols=3)
no_borders(ct)
set_table_width(ct, PAGE_W)

cred = [
    ('⚖  OAB/PR 118.596',    WD_ALIGN_PARAGRAPH.LEFT),
    ('🔒  Sigilo Absoluto',   WD_ALIGN_PARAGRAPH.CENTER),
    ('🌐  Atuação Nacional',  WD_ALIGN_PARAGRAPH.RIGHT),
]
for i, (txt, align) in enumerate(cred):
    c = ct.rows[0].cells[i]
    set_cell_bg(c, DARK2)
    set_cell_margins(c, top=180, left=500, bottom=180, right=500)
    p = c.paragraphs[0]
    p.alignment = align
    no_space(p)
    r = p.add_run(txt)
    r.font.size = Pt(8)
    r.font.color.rgb = GOLD_RGB
    r.font.name = 'Garamond'

gold_bar(height=40)

# ── 4. CORPO DO DOCUMENTO ────────────────────────────────────────────────────

body_t = doc.add_table(rows=1, cols=1)
no_borders(body_t)
set_table_width(body_t, PAGE_W)
bc = body_t.rows[0].cells[0]
set_cell_bg(bc, WHITE)
set_cell_margins(bc, top=560, left=1100, bottom=560, right=1100)
set_row_height(body_t.rows[0], 10200, exact=False)

# Destinatário / Exmo.
p_dest = bc.paragraphs[0]
no_space(p_dest)
p_dest.paragraph_format.space_after = Pt(6)
r = p_dest.add_run('Exmo.(a) Sr.(a) ')
r.font.size = Pt(11)
r.font.name = 'Garamond'
r.font.color.rgb = DARK_RGB
r2 = p_dest.add_run('__________________________________________')
r2.font.size = Pt(11)
r2.font.name = 'Garamond'
r2.font.color.rgb = GRAY_RGB

p_loc = bc.add_paragraph()
no_space(p_loc)
p_loc.paragraph_format.space_after = Pt(18)
r = p_loc.add_run('Curitiba/PR, _____ de __________________ de 20_____')
r.font.size = Pt(10)
r.font.name = 'Garamond'
r.font.color.rgb = GRAY_RGB

# Linha divisória dourada
gold_line(bc)

# Título do documento
p_title = bc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.paragraph_format.space_before = Pt(18)
p_title.paragraph_format.space_after = Pt(6)
r = p_title.add_run('TÍTULO DO DOCUMENTO / PETIÇÃO')
r.font.size = Pt(14)
r.font.bold = True
r.font.name = 'Garamond'
r.font.color.rgb = DARK_RGB

p_ref = bc.add_paragraph()
p_ref.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_ref.paragraph_format.space_after = Pt(20)
r = p_ref.add_run('Processo nº: ______________________ | Vara: ______________________')
r.font.size = Pt(9)
r.font.name = 'Garamond'
r.font.color.rgb = GOLD_RGB

gold_line(bc)

# Corpo do texto
p_body = bc.add_paragraph()
p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_body.paragraph_format.space_before = Pt(16)
p_body.paragraph_format.space_after = Pt(10)
p_body.paragraph_format.first_line_indent = Cm(1.25)
r = p_body.add_run(
    'Vem respeitosamente à presença de Vossa Excelência, por meio de seu advogado '
    'regularmente constituído, nos termos da procuração que se acosta, expor e requerer '
    'o quanto segue:'
)
r.font.size = Pt(11)
r.font.name = 'Garamond'
r.font.color.rgb = DARK_RGB

# Parágrafos de texto
for _ in range(4):
    p_text = bc.add_paragraph()
    p_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_text.paragraph_format.space_before = Pt(0)
    p_text.paragraph_format.space_after = Pt(10)
    p_text.paragraph_format.first_line_indent = Cm(1.25)
    r = p_text.add_run(
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Atuação especializada na '
        'Lei do Superendividamento. Analisamos contratos, evitamos abusos bancários e '
        'reestruturamos o passivo com absoluto sigilo e rigor técnico jurídico.'
    )
    r.font.size = Pt(11)
    r.font.name = 'Garamond'
    r.font.color.rgb = DARK_RGB

# Requerimento
p_req_title = bc.add_paragraph()
p_req_title.paragraph_format.space_before = Pt(14)
p_req_title.paragraph_format.space_after = Pt(6)
r = p_req_title.add_run('DOS REQUERIMENTOS')
r.font.size = Pt(11)
r.font.bold = True
r.font.name = 'Garamond'
r.font.color.rgb = GOLD_RGB

p_req = bc.add_paragraph()
p_req.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_req.paragraph_format.space_before = Pt(0)
p_req.paragraph_format.space_after = Pt(16)
p_req.paragraph_format.first_line_indent = Cm(1.25)
r = p_req.add_run(
    'Diante do exposto, requer a Vossa Excelência o deferimento do presente pedido, '
    'em consonância com os princípios da dignidade da pessoa humana e da proteção ao '
    'mínimo existencial.'
)
r.font.size = Pt(11)
r.font.name = 'Garamond'
r.font.color.rgb = DARK_RGB

p_termos = bc.add_paragraph()
p_termos.paragraph_format.space_after = Pt(30)
r = p_termos.add_run('Termos em que pede deferimento.')
r.font.size = Pt(11)
r.font.name = 'Garamond'
r.font.color.rgb = DARK_RGB

# Assinatura
p_sig_loc = bc.add_paragraph()
p_sig_loc.paragraph_format.space_after = Pt(30)
r = p_sig_loc.add_run('Curitiba/PR, _____ de __________________ de 20_____.')
r.font.size = Pt(11)
r.font.name = 'Garamond'
r.font.color.rgb = DARK_RGB

p_sig_line = bc.add_paragraph()
p_sig_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sig_line.paragraph_format.space_after = Pt(4)
r = p_sig_line.add_run('_' * 45)
r.font.color.rgb = GOLD_RGB
r.font.name = 'Garamond'

p_sig_name = bc.add_paragraph()
p_sig_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sig_name.paragraph_format.space_after = Pt(2)
r = p_sig_name.add_run('Advogado(a) Responsável')
r.font.size = Pt(10)
r.font.bold = True
r.font.name = 'Garamond'
r.font.color.rgb = DARK_RGB

p_sig_oab = bc.add_paragraph()
p_sig_oab.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_sig_oab.add_run('OAB/PR 118.596 | JCPV Advocacia')
r.font.size = Pt(9)
r.font.name = 'Garamond'
r.font.color.rgb = GOLD_RGB

# ── 5. BARRA DOURADA + RODAPÉ ESCURO ─────────────────────────────────────────

gold_bar(height=50)

ft = doc.add_table(rows=1, cols=1)
no_borders(ft)
set_table_width(ft, PAGE_W)
fc = ft.rows[0].cells[0]
set_cell_bg(fc, DARK)
set_cell_margins(fc, top=280, left=720, bottom=280, right=720)
set_row_height(ft.rows[0], 900, exact=False)

p_f1 = fc.paragraphs[0]
p_f1.alignment = WD_ALIGN_PARAGRAPH.CENTER
no_space(p_f1)
p_f1.paragraph_format.space_after = Pt(3)
r = p_f1.add_run('JCPV  ADVOCACIA')
r.font.size = Pt(11)
r.font.bold = True
r.font.name = 'Garamond'
r.font.color.rgb = GOLD_RGB

p_f2 = fc.add_paragraph()
p_f2.alignment = WD_ALIGN_PARAGRAPH.CENTER
no_space(p_f2)
p_f2.paragraph_format.space_after = Pt(3)
r = p_f2.add_run('www.jcpvadvocacia.com.br')
r.font.size = Pt(8)
r.font.name = 'Garamond'
r.font.color.rgb = RGBColor(0xC8, 0xB8, 0x99)

p_f3 = fc.add_paragraph()
p_f3.alignment = WD_ALIGN_PARAGRAPH.CENTER
no_space(p_f3)
r = p_f3.add_run('Alta Performance Jurídica  ·  Sigilo Absoluto  ·  Segurança Jurídica  ·  Atuação Nacional')
r.font.size = Pt(7.5)
r.font.name = 'Garamond'
r.font.color.rgb = GOLD_RGB

# ── Salvar ────────────────────────────────────────────────────────────────────

out = '/home/user/Apppainel-prestes/JCPV_Advocacia_Folha_Layout.docx'
doc.save(out)
print(f'Documento criado: {out}')
