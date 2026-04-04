"""
Recria a logo JCPV Advocacia em PNG transparente:
- Balança da justiça (triângulos + haste horizontal + fio)
- Dois quadrados geométricos sobrepostos no centro/topo
- Texto JCPV e ADVOCACIA abaixo
- Cor dourada #B89B6A em fundo transparente
"""
from PIL import Image, ImageDraw, ImageFont
import os

W, H = 500, 500
img = Image.new('RGBA', (W, H), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

GOLD = (184, 155, 106, 255)
LW = 3   # line width

# ── Geometria da balança ──────────────────────────────────────────────────────
cx = W // 2          # centro horizontal = 250
top_y = 80           # topo da haste vertical
pivot_y = 245        # ponto de pivô (onde a haste horizontal encosta)
arm_half = 150       # metade do comprimento do braço horizontal
cup_w = 60           # largura dos pratos
cup_h = 28           # altura dos pratos
cup_y_left  = pivot_y + 55   # altura do prato esquerdo
cup_y_right = pivot_y + 55   # altura do prato direito

# Fio vertical central (de cima ao pivô)
draw.line([(cx, top_y + 60), (cx, pivot_y)], fill=GOLD, width=LW)

# Braço horizontal
draw.line([(cx - arm_half, pivot_y), (cx + arm_half, pivot_y)],
          fill=GOLD, width=LW)

# ── Fios dos pratos ───────────────────────────────────────────────────────────
# Esquerdo: três fios saindo das extremidades + centro do prato
lx = cx - arm_half
rx = cx + arm_half

# Prato esquerdo — triângulo
draw.line([(lx, pivot_y), (lx - cup_w//2, cup_y_left)],  fill=GOLD, width=LW)
draw.line([(lx, pivot_y), (lx + cup_w//2, cup_y_left)],  fill=GOLD, width=LW)
# base do prato (linha horizontal)
draw.line([(lx - cup_w//2, cup_y_left),
           (lx + cup_w//2, cup_y_left)], fill=GOLD, width=LW)
# linhas internas do prato (detalhe)
for i in range(1, 4):
    yy = cup_y_left + i * 5
    xl = lx - cup_w//2 + i*3
    xr = lx + cup_w//2 - i*3
    if xl < xr:
        draw.line([(xl, yy), (xr, yy)], fill=GOLD, width=LW)

# Prato direito — triângulo
draw.line([(rx, pivot_y), (rx - cup_w//2, cup_y_right)], fill=GOLD, width=LW)
draw.line([(rx, pivot_y), (rx + cup_w//2, cup_y_right)], fill=GOLD, width=LW)
draw.line([(rx - cup_w//2, cup_y_right),
           (rx + cup_w//2, cup_y_right)], fill=GOLD, width=LW)
for i in range(1, 4):
    yy = cup_y_right + i * 5
    xl = rx - cup_w//2 + i*3
    xr2 = rx + cup_w//2 - i*3
    if xl < xr2:
        draw.line([(xl, yy), (xr2, yy)], fill=GOLD, width=LW)

# ── Quadrados geométricos (logo JCPV) ────────────────────────────────────────
# Dois quadrados sobrepostos escalonados, centralizados na haste vertical
# Quadrado inferior: menor, centrado em cx, com base em pivot_y-20
sq1_size = 62
sq1_x = cx - sq1_size // 2
sq1_y = pivot_y - sq1_size - 10
draw.rectangle(
    [sq1_x, sq1_y, sq1_x + sq1_size, sq1_y + sq1_size],
    outline=GOLD, width=LW
)

# Quadrado superior: um pouco maior, deslocado para cima e para a direita
sq2_size = 50
sq2_x = cx - sq2_size // 2 + 18
sq2_y = sq1_y - sq2_size + 14
draw.rectangle(
    [sq2_x, sq2_y, sq2_x + sq2_size, sq2_y + sq2_size],
    outline=GOLD, width=LW
)

# ── Texto JCPV ───────────────────────────────────────────────────────────────
text_y = cup_y_left + 26

try:
    font_jcpv = ImageFont.truetype('/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf', 42)
    font_adv  = ImageFont.truetype('/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf', 18)
except:
    font_jcpv = ImageFont.load_default()
    font_adv  = ImageFont.load_default()

# JCPV
bbox = draw.textbbox((0, 0), 'JCPV', font=font_jcpv)
tw = bbox[2] - bbox[0]
draw.text(((W - tw) // 2, text_y + 6), 'JCPV', font=font_jcpv, fill=GOLD)

# ADVOCACIA
bbox2 = draw.textbbox((0, 0), 'A D V O C A C I A', font=font_adv)
tw2 = bbox2[2] - bbox2[0]
draw.text(((W - tw2) // 2, text_y + 56), 'A D V O C A C I A',
          font=font_adv, fill=GOLD)

# Salva
out = '/home/user/Apppainel-prestes/jcpv_logo.png'
img.save(out, 'PNG')
print(f'Logo salva: {out}  ({os.path.getsize(out)} bytes)')
