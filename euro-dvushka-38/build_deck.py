#!/usr/bin/env python3
"""Сборка презентации «Евродвушка 38,6 м²»: схемы (SVG) + слайды (HTML)."""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from geometry import (WALLS, EXISTING, DEMOLISH, NEW, AFTER, ROOMS_BEFORE, ROOMS_AFTER,
                      furniture, C, H)
from draw import (svg_wrap, walls_svg, floors_svg, furn_svg, rooms_svg, dims_svg, rect, line,
                  text, p, iso_svg, scene, wall_boxes, INK, DEM, NEWC)

ROOT = sys.argv[1]
SCHEMES = sys.argv[2] if len(sys.argv) > 2 else None

# ---------------- оформление ----------------
BG = '#F5F1EA'
BG2 = '#EAE3D6'
DARK = '#2B2824'
BODY_C = '#4A443D'
WOOD = '#8A5A2B'
SAGE = '#5E7356'
HEAD = "'Jost', Arial, sans-serif"
BODY = "'Manrope', Arial, sans-serif"
TOTAL_SLIDES = 0

slides = []


def add(sid, html):
    slides.append((sid, html))


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def eyebrow(t, color=WOOD):
    return (f'<p style="font-size:24px;letter-spacing:3px;text-transform:uppercase;color:{color};'
            f'font-weight:600;font-family:{BODY}">{esc(t)}</p>')


def h2(t, size=60, color=DARK):
    return (f'<h2 style="font-family:{HEAD};font-size:{size}px;font-weight:400;line-height:1.1;'
            f'color:{color}">{esc(t)}</h2>')


def para(t, size=26, color=BODY_C, extra=''):
    return f'<p style="font-size:{size}px;line-height:1.4;color:{color}{extra}">{t}</p>'


def footer(n, left=1010, width=790, color='#6B645B'):
    return (f'<p style="position:absolute;left:{left}px;bottom:48px;width:{width}px;font-size:24px;'
            f'color:{color}">{n:02d} · Евродвушка 38,6 м²</p>')


def section(sid, inner, notes, bg=BG, color=DARK, layout='display:flex;flex-direction:column;gap:32px;padding:128px'):
    return (f'<section id="{sid}" data-transition="fade" style="background:{bg};color:{color};'
            f'font-family:{BODY};{layout}">{inner}<aside>{esc(notes)}</aside></section>')


def svg_inline(svg, left, top, w, h):
    return svg.replace('<svg ', f'<svg style="position:absolute;left:{left}px;top:{top}px;width:{w}px;height:{h}px" ', 1)


def legend_row(sw, label, shape='square'):
    radius = '50%' if shape == 'round' else '6px'
    return (f'<div style="display:flex;flex-direction:row;gap:16px;align-items:center">'
            f'<div style="width:28px;height:28px;background:{sw};border-radius:{radius};border:1px solid #7C7469"></div>'
            f'<p style="font-size:24px;line-height:1.3;color:{BODY_C}">{label}</p></div>')


def right_col(inner, top=96, height=None):
    return (f'<div style="position:absolute;left:1010px;top:{top}px;width:790px;display:flex;'
            f'flex-direction:column;gap:22px">{inner}</div>')


def plan_slide(sid, n, svg, title, eb, inner, notes, bg=BG):
    body = svg_inline(svg, 72, 84, 880, 919)
    body += right_col(eyebrow(eb) + h2(title, 54) + inner)
    body += footer(n)
    add(sid, section(sid, body, notes, bg=bg, layout='padding:128px'))


def save_svg(name, svg):
    if SCHEMES:
        os.makedirs(SCHEMES, exist_ok=True)
        with open(os.path.join(SCHEMES, name + '.svg'), 'w', encoding='utf-8') as f:
            f.write(svg)


# ================= СХЕМЫ =================
F_DAY = furniture('day')


def furn_plan(items, faint=False, cats=None, highlight=None):
    """Мебель на плане: верхние шкафы — пунктиром."""
    out = []
    fills = dict(furn='#EFE6D6', store='#E9D7BC', kit='#E4E9E3', san='#E3EBEE')
    for it in items:
        x0, y0, x1, y1, z0, z1, col, label, cat = it
        if cat == 'deco':
            continue
        if cats and cat not in cats:
            continue
        upper = z0 >= 1.4
        if highlight is not None:
            f = '#E8C79A' if cat == 'store' else '#F3EFE8'
            st = WOOD if cat == 'store' else '#C9C1B5'
            sw = 2.5 if cat == 'store' else 1
        elif faint:
            f, st, sw = '#F3EFE8', '#C9C1B5', 1
        else:
            f, st, sw = fills.get(cat, '#EEE'), '#7C7469', 1.2
        if upper:
            out.append(rect(x0, y0, x1, y1, 'none', st, sw, ' stroke-dasharray="6 4"'))
        else:
            out.append(rect(x0, y0, x1, y1, f, st, sw))
    return ''.join(out)


def base_plan(walls=AFTER, rooms=ROOMS_AFTER, furn=True, faint_furn=False, room_labels=True,
              area=True, demolish=(), new=(), dims=True, wall_style='normal', extra='', furn_kw=None):
    body = floors_svg()
    if furn:
        body += furn_plan(F_DAY, faint=faint_furn, **(furn_kw or {}))
    body += walls_svg(walls, wall_style, demolish, new)
    if room_labels:
        body += rooms_svg(rooms, area=area)
    if dims:
        body += dims_svg()
    body += extra
    return body


# ---------- 1. Обмерный план (существующее положение) ----------
exist_extra = ''
for nm, (x, y) in {'Ст.': (6.57, 4.42)}.items():
    exist_extra += f'<circle cx="{p(x)}" cy="{p(y)}" r="12" fill="#2F7FA0"/>'
    exist_extra += text(x - 0.35, y + 0.07, 'стояки', 16, '#2F7FA0')
svg_exist = svg_wrap(base_plan(EXISTING, ROOMS_BEFORE, furn=False) + exist_extra,
                     'Обмерный план существующей квартиры')
save_svg('01_obmernyi_plan', svg_exist)

# ---------- 2. Демонтаж ----------
svg_dem = svg_wrap(base_plan(EXISTING, ROOMS_BEFORE, furn=False, demolish=DEMOLISH, area=False) +
                   text(2.94, 3.95, '1', 20, DEM, 700) + text(2.6, 4.85, '2', 20, DEM, 700) +
                   text(3.9, 4.0, '3', 20, DEM, 700),
                   'План демонтажа перегородок')
save_svg('02_plan_demontazha', svg_dem)

# ---------- 3. Монтаж ----------
mont_extra = (text(1.44, 3.32, 'Перегородка ГКЛ 100 мм', 18, NEWC, 700) +
              text(2.45, 3.82, 'дверь 800', 16, NEWC, 600))
svg_mont = svg_wrap(base_plan(AFTER, ROOMS_AFTER, furn=False, new=NEW, area=False, extra=mont_extra),
                    'План монтажа новых перегородок')
save_svg('03_plan_montazha', svg_mont)

# ---------- 4. План после перепланировки ----------
svg_new = svg_wrap(base_plan(AFTER, ROOMS_AFTER, furn=False), 'План после перепланировки')
save_svg('04_plan_posle', svg_new)

# ---------- 5. Мебель (нумерация) ----------
FURN_NUMS = [
    (1, (1.0, 1.55), 'Кровать 160×200 с подъёмным механизмом и ящиком'),
    (2, (1.0, 3.13), 'Шкаф-купе в спальне, глубина 55 см, до потолка'),
    (3, (1.0, 3.72), 'Закрытый стеллаж-шкаф 40 см в гостиной зоне'),
    (4, (0.47, 5.6), 'Диван-кровать «еврокнижка» 200×95 с ящиком'),
    (5, (2.66, 6.1), 'ТВ на пилоне + подвесная тумба'),
    (6, (6.45, 2.85), 'Кухня Г-образная: 2,5 + 1,8 м, шкафы до потолка'),
    (7, (6.45, 1.87), 'Встроенный холодильник + антресоль'),
    (8, (3.95, 2.45), 'Обеденный стол Ø90, 3–4 стула'),
    (9, (3.3, 6.1), 'Шкаф в прихожей с зеркальной дверью'),
    (10, (4.85, 6.3), 'Обувница-скамья + вешалка'),
    (11, (5.45, 4.54), 'Стиральная машина под столешницей'),
    (12, (5.94, 6.44), 'Ванна 160×70 со шторкой'),
    (13, (4.35, 0.35), 'Рабочее место на утеплённой лоджии'),
    (14, (3.3, 0.94), 'Шкаф для хранения на лоджии'),
]
num_svg = ''
for n, (x, y), _ in FURN_NUMS:
    num_svg += f'<circle cx="{p(x)}" cy="{p(y)}" r="15" fill="{INK}"/>' + text(x, y + 0.065, str(n), 16, '#FBFBF8', 700)
svg_furn = svg_wrap(base_plan(AFTER, ROOMS_AFTER, room_labels=False, extra=num_svg),
                    'План расстановки мебели')
save_svg('05_plan_mebeli', svg_furn)

# ---------- 6. Хранение ----------
STORE_NUMS = [(1, (1.0, 1.55)), (2, (1.0, 3.13)), (3, (1.0, 3.72)), (4, (0.47, 5.6)),
              (5, (2.66, 6.1)), (6, (6.57, 3.0)), (7, (3.3, 6.1)), (8, (4.85, 6.3)),
              (9, (6.07, 4.5)), (10, (6.65, 5.17)), (11, (3.3, 0.94))]
st_svg = ''
for n, (x, y) in STORE_NUMS:
    st_svg += f'<circle cx="{p(x)}" cy="{p(y)}" r="15" fill="{WOOD}"/>' + text(x, y + 0.065, str(n), 16, '#FBFBF8', 700)
svg_store = svg_wrap(base_plan(AFTER, ROOMS_AFTER, room_labels=False, furn_kw=dict(highlight=True),
                               extra=st_svg), 'Схема скрытых систем хранения')
save_svg('06_skhema_khraneniya', svg_store)

# ---------- 7. Водоснабжение ----------
BLUE, RED, BROWN = '#2F6FB0', '#C8402F', '#7A4E2D'


def pl(pts, color, sw=4, dash=''):
    d = f' stroke-dasharray="{dash}"' if dash else ''
    return (f'<polyline points="{" ".join(f"{p(x)},{p(y)}" for x, y in pts)}" fill="none" '
            f'stroke="{color}" stroke-width="{sw}" stroke-linejoin="round" stroke-linecap="round"{d}/>')


def mark(x, y, label, color=INK, dx=0.0, dy=-0.2, anchor='middle', size=17):
    return (f'<circle cx="{p(x)}" cy="{p(y)}" r="9" fill="#FBFBF8" stroke="{color}" stroke-width="3"/>' +
            text(x + dx, y + dy, label, size, color, 700, anchor))


risers = (f'<circle cx="{p(6.5)}" cy="{p(4.36)}" r="8" fill="{BLUE}"/>'
          f'<circle cx="{p(6.66)}" cy="{p(4.36)}" r="8" fill="{RED}"/>'
          f'<circle cx="{p(6.58)}" cy="{p(4.52)}" r="10" fill="{BROWN}"/>')
water = ''
# кухня: от стояков через перегородку КВ
water += pl([(6.5, 4.36), (6.5, 4.05), (5.75, 4.05), (5.75, 3.85)], BLUE)
water += pl([(6.66, 4.36), (6.66, 3.98), (5.85, 3.98), (5.85, 3.85)], RED)
water += pl([(5.75, 4.05), (5.2, 4.05)], BLUE)
# ванная
water += pl([(6.5, 4.36), (6.3, 4.36), (6.3, 4.3), (5.45, 4.3)], BLUE)
water += pl([(6.66, 4.36), (6.66, 4.4), (6.1, 4.4), (6.1, 4.32)], RED)
water += pl([(6.5, 4.36), (6.5, 4.62), (6.62, 4.62), (6.62, 5.15)], BLUE)
water += pl([(6.62, 5.15), (6.62, 6.0), (6.0, 6.0), (6.0, 6.05)], BLUE)
water += pl([(6.7, 4.62), (6.7, 6.04), (6.05, 6.04)], RED)
water += risers
water += mark(5.8, 3.82, 'М', BLUE, 0.0, -0.2)
water += mark(5.2, 4.05, 'ПММ', BLUE, 0.0, -0.2)
water += mark(5.45, 4.3, 'СМ', BLUE, 0.0, 0.42)
water += mark(6.1, 4.32, 'Р', BLUE, -0.22, 0.42)
water += mark(6.62, 5.15, 'У', BLUE, -0.3, 0.07)
water += mark(6.0, 6.03, 'В', BLUE, -0.3, 0.07)
water += text(6.3, 4.95, 'коллектор', 15, BLUE, 600)
svg_water = svg_wrap(base_plan(AFTER, ROOMS_AFTER, faint_furn=True, area=False, wall_style='faint') + water,
                     'Схема водоснабжения')
save_svg('07_skhema_vodosnabzheniya', svg_water)

# ---------- 8. Канализация ----------
sew = ''
sew += pl([(5.75, 3.85), (5.75, 4.0), (6.45, 4.0), (6.45, 4.52), (6.58, 4.52)], BROWN, 5)          # мойка Ø50
sew += pl([(6.62, 5.18), (6.62, 4.6)], BROWN, 8)                         # унитаз Ø110
sew += pl([(5.45, 4.6), (6.4, 4.6), (6.4, 4.55)], BROWN, 5)              # СМ + раковина Ø50
sew += pl([(6.55, 6.45), (6.72, 6.45), (6.72, 4.62)], BROWN, 5)          # ванна Ø50
sew += f'<circle cx="{p(6.58)}" cy="{p(4.52)}" r="13" fill="{BROWN}"/>'
sew += mark(5.75, 3.85, 'М', BROWN, 0.0, -0.2)
sew += mark(6.62, 5.18, 'У', BROWN, -0.3, 0.07)
sew += mark(5.45, 4.6, 'СМ', BROWN, 0.0, 0.4)
sew += mark(6.1, 4.6, 'Р', BROWN, 0.0, 0.4)
sew += mark(6.55, 6.45, 'В', BROWN, -0.3, 0.07)
sew += text(5.5, 3.95, 'Ø50, i=3%', 16, BROWN, 700, 'end')
sew += text(6.25, 5.0, 'Ø110', 16, BROWN, 700, 'end')
sew += text(6.4, 6.3, 'Ø50, i=2%', 16, BROWN, 700, 'end')
sew += text(6.25, 4.25, 'К1 Ø110', 16, BROWN, 700, 'end')
svg_sewer = svg_wrap(base_plan(AFTER, ROOMS_AFTER, faint_furn=True, area=False, wall_style='faint') + sew,
                     'Схема канализации')
save_svg('08_skhema_kanalizatsii', svg_sewer)

# ---------- 9. Розетки / силовые ----------
ORANGE, POWER, IP, LOW = '#C26A2E', '#8E2F22', '#2F7FA0', '#3E7F74'


def sock(x, y, kind='r', n=1, lab=''):
    s = ''
    if kind == 'r':
        s += f'<circle cx="{p(x)}" cy="{p(y)}" r="10" fill="{ORANGE}" stroke="#FBFBF8" stroke-width="2"/>'
    elif kind == 'w':
        s += f'<circle cx="{p(x)}" cy="{p(y)}" r="10" fill="{IP}" stroke="#FBFBF8" stroke-width="2"/>'
    elif kind == 'p':
        s += f'<rect x="{float(p(x)) - 10:.1f}" y="{float(p(y)) - 10:.1f}" width="20" height="20" fill="{POWER}" stroke="#FBFBF8" stroke-width="2"/>'
    elif kind == 'l':
        cx, cy = float(p(x)), float(p(y))
        s += f'<polygon points="{cx:.1f},{cy - 11:.1f} {cx + 11:.1f},{cy + 9:.1f} {cx - 11:.1f},{cy + 9:.1f}" fill="{LOW}" stroke="#FBFBF8" stroke-width="2"/>'
    elif kind == 'ac':
        s += f'<rect x="{float(p(x)) - 22:.1f}" y="{float(p(y)) - 11:.1f}" width="44" height="22" rx="4" fill="{POWER}"/>'
        s += text(x, y + 0.055, 'AC', 15, '#FBFBF8', 700)
    if n > 1:
        s += text(x + 0.17, y - 0.1, f'×{n}', 15, INK, 700, 'start')
    if lab:
        s += text(x + 0.17, y + 0.2, lab, 14, INK, 600, 'start')
    return s


SOCKETS = [
    # кухня
    (6.68, 2.35, 'r', 2, ''), (6.68, 3.35, 'r', 2, ''), (6.3, 4.06, 'r', 2, ''),
    (6.68, 2.85, 'p', 1, ''), (5.2, 4.06, 'r', 1, ''), (6.68, 1.87, 'r', 1, ''), (3.06, 2.45, 'r', 1, ''),
    (3.35, 1.64, 'ac', 1, ''), (6.0, 1.64, 'r', 1, ''),
    # гостиная зона
    (0.07, 6.5, 'r', 2, ''), (0.07, 4.0, 'r', 1, ''), (2.82, 6.2, 'r', 2, ''), (2.82, 5.85, 'l', 1, ''),
    # спальня
    (0.07, 0.55, 'r', 2, ''), (0.07, 2.58, 'r', 2, ''), (2.82, 0.9, 'ac', 1, ''), (2.82, 0.4, 'r', 1, ''),
    # прихожая
    (4.96, 6.0, 'r', 1, ''), (4.96, 4.55, 'r', 1, ''), (3.8, 6.73, 'l', 1, ''),
    # ванная
    (5.45, 4.3, 'w', 1, ''), (5.85, 4.3, 'w', 1, ''), (5.19, 5.9, 'w', 1, ''),
    # лоджия
    (3.07, 0.3, 'r', 2, ''), (3.07, 0.05, 'l', 1, ''),
]
sock_svg = ''.join(sock(*s) for s in SOCKETS)
# щит
sock_svg += (f'<rect x="{p(3.62)}" y="{p(6.66)}" width="{p(0.32)}" height="{p(0.13)}" fill="{INK}"/>' +
             text(3.78, 6.6, 'ЩК', 16, INK, 700))
svg_sock = svg_wrap(base_plan(AFTER, ROOMS_AFTER, faint_furn=True, area=False, wall_style='faint') + sock_svg,
                    'Схема розеток и силовых линий')
save_svg('09_skhema_rozetok', svg_sock)

# ---------- 10. Освещение ----------
YEL = '#D99A1E'


def spot(x, y):
    return (f'<circle cx="{p(x)}" cy="{p(y)}" r="8" fill="#FFF6DE" stroke="{INK}" stroke-width="2"/>' +
            line(x - 0.06, y, x + 0.06, y, INK, 1.5) + line(x, y - 0.06, x, y + 0.06, INK, 1.5))


def pend(x, y):
    return f'<circle cx="{p(x)}" cy="{p(y)}" r="15" fill="{YEL}" stroke="{INK}" stroke-width="2"/><circle cx="{p(x)}" cy="{p(y)}" r="5" fill="{INK}"/>'


def ceil(x, y):
    return (f'<circle cx="{p(x)}" cy="{p(y)}" r="17" fill="#FFF6DE" stroke="{INK}" stroke-width="2.5"/>' +
            line(x - 0.12, y - 0.12, x + 0.12, y + 0.12, INK, 2) + line(x - 0.12, y + 0.12, x + 0.12, y - 0.12, INK, 2))


def sconce(x, y):
    return f'<circle cx="{p(x)}" cy="{p(y)}" r="10" fill="{YEL}" stroke="{INK}" stroke-width="2"/>'


def strip(pts):
    return pl(pts, YEL, 6, '10 6')


def sw(x, y, lab):
    return (f'<circle cx="{p(x)}" cy="{p(y)}" r="11" fill="{INK}"/>' + text(x, y + 0.06, lab, 15, '#FBFBF8', 700))


def link(a, b):
    return line(a[0], a[1], b[0], b[1], '#7C7469', 1.5, ' stroke-dasharray="4 5"')


light = ''
# кухня
k_spots = [(5.85, 2.6), (5.85, 3.25), (5.2, 3.3)]
light += strip([(6.36, 2.22), (6.36, 3.72), (5.0, 3.72)])
for s in k_spots:
    light += spot(*s)
light += pend(3.95, 2.45)
# гостиная зона
lv = [(2.3, 4.45), (2.3, 5.2), (2.3, 5.95)]
for s in lv:
    light += spot(*s)
light += sconce(0.2, 6.45)
# прихожая
hl = [(4.05, 4.7), (4.4, 5.2), (4.2, 6.4)]
for s in hl:
    light += spot(*s)
light += strip([(3.07, 5.5), (3.07, 6.7)])
# спальня
light += ceil(1.44, 2.62)
light += sconce(0.1, 0.5) + sconce(0.1, 2.62)
light += strip([(0.1, 2.92), (1.9, 2.92)])
# ванная
bl = [(5.45, 5.15), (6.35, 5.8), (5.55, 6.45)]
for s in bl:
    light += spot(*s)
light += sconce(6.08, 4.3)
light += f'<circle cx="{p(6.57)}" cy="{p(4.42)}" r="10" fill="#FBFBF8" stroke="{INK}" stroke-width="2"/>' + text(6.57, 4.47, 'В', 13, INK, 700)
# лоджия
light += ceil(5.4, 0.62)
# выключатели и связи
SW = {'1': (4.95, 6.6), '2': (3.06, 3.25), '3': (2.82, 5.5), '4': (2.75, 3.25), '5': (0.12, 0.25), '6': (4.95, 4.9), '7': (5.95, 1.65)}
light += link(SW['1'], hl[2]) + link(hl[2], hl[1]) + link(hl[1], hl[0])
light += link(SW['2'], (3.95, 2.45)) + link(SW['2'], k_spots[2]) + link(k_spots[2], k_spots[1]) + link(k_spots[1], k_spots[0])
light += link(SW['3'], lv[2]) + link(lv[2], lv[1]) + link(lv[1], lv[0])
light += link(SW['4'], (1.44, 2.62)) + link(SW['5'], (1.44, 2.62))
light += link(SW['6'], bl[0]) + link(bl[0], bl[1]) + link(bl[0], bl[2])
light += link(SW['7'], (5.4, 0.62))
for k, (x, y) in SW.items():
    light += sw(x, y, k)
svg_light = svg_wrap(base_plan(AFTER, ROOMS_AFTER, faint_furn=True, area=False, wall_style='faint') + light,
                     'Схема освещения и выключателей')
save_svg('10_skhema_osveshcheniya', svg_light)

# ---------- 3D ----------
REG_ALL = (-0.2, -0.4, 6.95, 6.99)
iso_over = iso_svg(scene('day'), 'pp', REG_ALL, '3D: общий вид квартиры', size=(1000, 820), cut=1.0)
iso_live_day = iso_svg(scene('day'), 'pp', (-0.2, 3.4, 3.0, 6.99), '3D: гостиная зона, день', size=(800, 620), cut=0.12, back_pad=0.3)
iso_live_guest = iso_svg(scene('guest'), 'pp', (-0.2, 3.4, 3.0, 6.99), '3D: гостиная зона, диван разложен', size=(800, 620), cut=0.12, back_pad=0.3)
iso_kitchen = iso_svg(scene('day'), 'mp', (2.88, 1.27, 6.95, 4.24), '3D: кухня', size=(1000, 760), cut=0.12, back_pad=0.6, furn_cut=1.2)
iso_bed = iso_svg(scene('day'), 'pm', (-0.2, -0.4, 3.0, 3.52), '3D: спальня', size=(1000, 760), cut=0.12, back_pad=0.3)
iso_bath = iso_svg(scene('day'), 'mp', (5.03, 4.12, 6.95, 6.99), '3D: ванная', size=(800, 700), cut=0.12, back_pad=0.3)
iso_hall = iso_svg(scene('day'), 'pm', (2.88, 4.12, 5.13, 6.99), '3D: прихожая', size=(800, 700), cut=0.12, back_pad=0.3)
for nm, s in [('11_3d_obshchii_vid', iso_over), ('12_3d_gostinaya_den', iso_live_day),
              ('13_3d_gostinaya_gosti', iso_live_guest), ('14_3d_kukhnya', iso_kitchen),
              ('15_3d_spalnya', iso_bed), ('16_3d_vannaya', iso_bath), ('17_3d_prikhozhaya', iso_hall)]:
    save_svg(nm, s)


# ---------- x-embed: вращающаяся 3D-модель ----------
def embed_js():
    boxes = []
    for r, col in [((-0.2, -0.4, 6.95, 6.99), None)]:
        pass
    from geometry import FLOORS
    pal = []

    def ci(c):
        if c not in pal:
            pal.append(c)
        return pal.index(c)

    for r, col in FLOORS:
        boxes.append([*r, -0.06, 0, ci(col)])
    for b in wall_boxes(AFTER, zcut=lambda w: 1.1):
        x0, y0, x1, y1, z0, z1, col, kind = b
        boxes.append([x0, y0, x1, y1, z0, min(z1, 1.1), ci(col)])
    for it in furniture('day'):
        x0, y0, x1, y1, z0, z1, col, _, cat = it
        boxes.append([x0, y0, x1, y1, z0, z1, ci(col)])
    data = json.dumps([[round(v, 2) if isinstance(v, float) else v for v in b] for b in boxes if b[5] > b[4]],
                      separators=(',', ':'))
    js = ("const B=" + data + ",P=" + json.dumps(pal) + ";"
          "const c=document.getElementById('c'),g=c.getContext('2d');let w,h;"
          "function rs(){const d=devicePixelRatio||1;w=c.width=innerWidth*d;h=c.height=innerHeight*d;}rs();onresize=rs;"
          "const F=[];function sub(a,b){const n=Math.max(1,Math.ceil((b-a)/0.45)),r=[];for(let i=0;i<n;i++)r.push([a+(b-a)*i/n,a+(b-a)*(i+1)/n]);return r}"
          "const RGB=P.map(s=>[1,3,5].map(i=>parseInt(s.substr(i,2),16)));"
          "for(const b of B){const[x0,y0,z0,x1,y1,z1,k]=b;"
          "for(const[a,bb]of sub(x0,x1))for(const[cc,d]of sub(y0,y1))for(const[e,f]of sub(z0,z1)){"
          "if(f===z1)F.push({p:[[a,cc,f],[bb,cc,f],[bb,d,f],[a,d,f]],n:[0,0,1],k});"
          "if(a===x0)F.push({p:[[a,cc,e],[a,d,e],[a,d,f],[a,cc,f]],n:[-1,0,0],k});"
          "if(bb===x1)F.push({p:[[bb,cc,e],[bb,d,e],[bb,d,f],[bb,cc,f]],n:[1,0,0],k});"
          "if(cc===y0)F.push({p:[[a,cc,e],[bb,cc,e],[bb,cc,f],[a,cc,f]],n:[0,-1,0],k});"
          "if(d===y1)F.push({p:[[a,d,e],[bb,d,e],[bb,d,f],[a,d,f]],n:[0,1,0],k});}}"
          "const cx=3.35,cy=3.3,el=0.62,ce=Math.cos(el),se=Math.sin(el);"
          "function draw(t){const a=t/9000,ca=Math.cos(a),sa=Math.sin(a),S=Math.min(w/9.4,h/7.6);"
          "g.fillStyle='#F5F1EA';g.fillRect(0,0,w,h);const L=[];"
          "for(const f of F){const[nx,ny,nz]=f.n,rx=nx*ca-ny*sa,ry=nx*sa+ny*ca;"
          "if(ry*ce-nz*se>=0)continue;let dd=0;"
          "const q=f.p.map(([x,y,z])=>{const X=(x-cx)*ca-(y-cy)*sa,Y=(x-cx)*sa+(y-cy)*ca;dd+=Y*ce-z*se;"
          "return[w/2+X*S,h*0.56-(Y*se+z*ce)*S]});"
          "const s=0.66+0.34*Math.max(0,-0.45*rx-0.55*ry+0.7*nz);L.push([dd,q,f.k,s]);}"
          "L.sort((a,b)=>b[0]-a[0]);"
          "for(const[dd,q,k,s]of L){const c=RGB[k];"
          "g.fillStyle='rgb('+(c[0]*s|0)+','+(c[1]*s|0)+','+(c[2]*s|0)+')';"
          "g.strokeStyle='rgba(60,50,40,0.18)';g.lineWidth=1;g.beginPath();g.moveTo(q[0][0],q[0][1]);"
          "for(let i=1;i<4;i++)g.lineTo(q[i][0],q[i][1]);g.closePath();g.fill();g.stroke();}"
          "requestAnimationFrame(draw)}requestAnimationFrame(draw);")
    return ('<!doctype html><html><head><meta charset="utf-8"></head>'
            '<body style="margin:0;background:#F5F1EA;overflow:hidden">'
            '<canvas id="c" style="width:100vw;height:100vh;display:block"></canvas><script>' + js + '</script></body></html>')


EMBED = embed_js()

# ================= СЛАЙДЫ =================
N = 0


def nxt():
    global N
    N += 1
    return N


# 1. Обложка
n = nxt()
inner = svg_inline(iso_over, 860, 150, 980, 804)
inner += (f'<div style="position:absolute;left:128px;top:150px;width:760px;display:flex;flex-direction:column;gap:28px">'
          + eyebrow('Дизайн-концепция · перепланировка')
          + f'<h1 style="font-family:{HEAD};font-size:92px;font-weight:400;line-height:1.05;color:{DARK}">Евродвушка из однокомнатной квартиры</h1>'
          + para('38,6 м² · современный минимализм со скандинавскими акцентами · эконом-бюджет', 30)
          + '<div style="display:flex;flex-direction:row;gap:48px">'
          + ''.join(f'<div style="display:flex;flex-direction:column;gap:4px"><p style="font-family:{HEAD};font-size:64px;color:{WOOD};line-height:1.1">{a}</p><p style="font-size:24px;color:{BODY_C}">{b}</p></div>'
                    for a, b in [('19 м²', 'кухня-гостиная'), ('9,7 м²', 'спальня с окном'), ('11', 'зон скрытого хранения')])
          + '</div></div>')
inner += f'<p style="position:absolute;left:128px;bottom:64px;width:900px;font-size:24px;color:#6B645B">Схемы, 3D-виды и инженерия · октябрь 2026</p>'
add('cover', section('cover', inner,
                     'Обложка. Квартира 38,6 м²: однокомнатная превращается в евродвушку: кухня-гостиная плюс отдельная спальня с окном. Справа общий 3D-вид после перепланировки.',
                     layout='padding:128px'))

# 2. Вводные
n = nxt()
rows = [
    ('Площадь', '38,6 м², 1 комната + кухня 9,57 + лоджия', 'Евродвушка: кухня-гостиная ≈19 м² + спальня 9,7 м²'),
    ('Кто живёт', 'Пара, возможно с ребёнком', 'Спальня с местом под кроватку, гостевое спальное место'),
    ('Стиль', 'Минимализм + скандинавский', 'Гладкие фасады без ручек, светлое дерево, текстиль'),
    ('Бюджет', 'Эконом', 'Минимум сноса, мокрые зоны на месте, ГКЛ, ЛДСП'),
    ('Окна', 'Солнечная сторона, дорога', 'Блэкаут, 2 кондиционера, шумозащитные стеклопакеты'),
    ('Обязательно', 'Диван-кровать, кровать, хранение', 'Еврокнижка, кровать с подъёмником, 11 зон хранения'),
    ('Цвет', 'Светлые нейтральные + дерево', 'Тёплый белый, грейж, дуб, шалфей'),
]
tbl = (f'<table style="font-size:25px;color:{BODY_C};width:1664px">'
       f'<tr><th style="width:18%;text-align:left">Параметр</th><th style="width:36%;text-align:left">Вводные</th><th style="width:46%;text-align:left">Решение в проекте</th></tr>'
       + ''.join(f'<tr><td><b>{a}</b></td><td>{b}</td><td>{c}</td></tr>' for a, b, c in rows) + '</table>')
add('brief', section('brief', eyebrow('01 · Задача') + h2('Вводные и ответ проекта') + tbl + footer(n, 128, 900),
                     'Свод вводных заказчика и того, как каждое требование закрыто в проекте.', layout='display:flex;flex-direction:column;gap:28px;padding:128px 128px 160px'))

# 3. Анализ плана
n = nxt()
cards = [
    ('Сильные стороны', 'Длинная гостиная 19,3 м² с окном на фасад. Кухня 9,57 м² с выходом на лоджию. Ванная у стояков рядом с кухней.'),
    ('Ограничения', 'Одно окно в гостиной, ширина комнаты ≈2,9 м. Делить её вдоль нельзя, только поперёк: спальне нужно окно.'),
    ('Ключевая идея', 'Спальня в светлой части у окна. Глубина комнаты, кухня и прихожая объединяются в одну кухню-гостиную.'),
]
row = '<div style="display:flex;flex-direction:row;gap:32px">' + ''.join(
    f'<div style="flex:1;display:flex;flex-direction:column;gap:16px;background:#FBFAF6;padding:40px;border-radius:16px;border:1px solid #DDD5C8">'
    f'<h3 style="font-family:{HEAD};font-size:40px;font-weight:500;color:{DARK}">{a}</h3>{para(b, 26)}</div>' for a, b in cards) + '</div>'
add('analysis', section('analysis', eyebrow('02 · Анализ плана') + h2('Что даёт ваша планировка') + row +
                        para('Размеры сняты с плана застройщика без размерной сетки. Перед рабочей документацией нужен обмер с точностью до 1 см.', 24, '#6B645B') + footer(n, 128, 900),
                        'Анализ исходного плана. Главная идея: спальня занимает светлую часть у окна, глубина комнаты уходит в общую зону вместе с кухней.',
                        layout='display:flex;flex-direction:column;gap:40px;padding:128px 128px 160px'))

# 4. Концепция и палитра
n = nxt()
pal = [('#F4F1EA', 'Тёплый белый', 'стены, фасады'), ('#DDD5C7', 'Грейж', 'текстиль, акцент'),
       ('#D9BC93', 'Дуб светлый', 'пол SPC'), ('#C79E6E', 'Дуб натуральный', 'столешницы, рейки'),
       ('#9DAE92', 'Шалфей', 'текстиль, декор'), ('#4B4845', 'Графит', 'светильники, фурнитура')]
sw_row = '<div style="display:flex;flex-direction:row;gap:24px">' + ''.join(
    f'<div style="flex:1;display:flex;flex-direction:column;gap:10px"><div style="height:180px;background:{c};border-radius:12px;border:1px solid #D2CABD"></div>'
    f'<p style="font-size:26px;font-weight:600;color:{DARK}">{nm}</p><p style="font-size:24px;color:#6B645B">{use} · {c}</p></div>' for c, nm, use in pal) + '</div>'
princ = '<div style="display:flex;flex-direction:row;gap:32px">' + ''.join(
    f'<div style="flex:1;display:flex;flex-direction:column;gap:10px"><h3 style="font-family:{HEAD};font-size:34px;font-weight:500;color:{WOOD}">{a}</h3>{para(b, 25)}</div>'
    for a, b in [('Один пол на всю квартиру', 'SPC «светлый дуб» без порогов, кроме ванной, зрительно расширяет пространство.'),
                 ('Мебель в цвет стен', 'Шкафы до потолка с белыми фасадами без ручек растворяются в стене.'),
                 ('Тепло дерева и ткани', 'Дуб в столешницах и мелочах, лён и шалфей в текстиле: скандинавский уют без декора ради декора.')]) + '</div>'
add('concept', section('concept', eyebrow('03 · Концепция') + h2('Светлая база, тёплое дерево, ничего лишнего') + sw_row + princ + footer(n, 128, 900),
                       'Палитра и принципы стиля. Шесть цветов на всю квартиру, один пол, мебель в цвет стен.',
                       bg=BG, layout='display:flex;flex-direction:column;gap:40px;padding:110px 128px 150px'))

# 5. Обмерный план
n = nxt()
expl_before = [('Гостиная', '19,30'), ('Кухня', '9,57'), ('Прихожая', '5,13'), ('Ванная', '4,23'), ('Лоджия (к=0,5)', '1,92')]
tb = (f'<table style="font-size:25px;color:{BODY_C};width:790px"><tr><th style="width:70%;text-align:left">Помещение</th><th style="width:30%;text-align:right">м²</th></tr>'
      + ''.join(f'<tr><td>{a}</td><td style="text-align:right">{b}</td></tr>' for a, b in expl_before) + '</table>')
plan_slide('existing', n, svg_exist, 'Обмерный план: как есть', '04 · Существующее положение',
           tb + para('Стена между гостиной и прихожей у входа толстая, вероятно несущий пилон, её не трогаем. Тонкие перегородки проверяем по техпаспорту БТИ.', 24),
           'Существующее положение по плану заказчика. Площади по документам, размеры ориентировочные.')

# 6. Демонтаж
n = nxt()
lst = ''.join(legend_row(c, t) for c, t in [
    ('#F6DCD7', '<b>1</b> · часть перегородки гостиная/кухня, ≈0,8 м'),
    ('#F6DCD7', '<b>2</b> · перегородка гостиная/прихожая с дверью, ≈1,1 м'),
    ('#F6DCD7', '<b>3</b> · перегородка кухня/прихожая с проёмом, ≈1,85 м'),
    ('#3A3631', 'несущие стены и пилон не трогаем'),
])
plan_slide('demolition', n, svg_dem, 'План демонтажа', '05 · Перепланировка', lst +
           para('Сносим только ненесущие перегородки, около 3,7 погонных метров. Мокрые зоны остаются на месте, поэтому согласование идёт как простая перепланировка.', 24) +
           para('<b>До работ:</b> подтвердить по техпаспорту, что стены 1–3 ненесущие. Если на кухне газ, нужна раздвижная перегородка.', 24, WOOD),
           'Демонтаж трёх участков ненесущих перегородок. Несущий пилон у входа остаётся.')

# 7. Монтаж
n = nxt()
plan_slide('mounting', n, svg_mont, 'План монтажа перегородок', '06 · Перепланировка',
           ''.join(legend_row(c, t) for c, t in [('#D6E5DE', 'новая перегородка ГКЛ, 100 мм'), ('#9A9286', 'существующие перегородки')]) +
           para('<b>Перегородка спальни</b>: металлокаркас 75 мм, минвата 50 мм, ГКЛ в 2 слоя с каждой стороны. Звукоизоляция около 45 дБ, длина 2,88 м, дверь 800 мм.', 24) +
           para('В каркас заранее ставим закладные под шкаф-купе и стеллаж, проводку под розетки у кровати и выключатели.', 24),
           'Монтаж одной новой перегородки, которая отделяет спальню. Звукоизоляция обязательна.')

# 8. План после перепланировки
n = nxt()
expl_after = [('Спальня', '9,7', 'у окна'), ('Гостиная зона', '9,3', 'диван-кровать'), ('Кухня', '9,57', 'открыта'),
              ('Прихожая', '5,13', 'открыта'), ('Ванная', '4,23', 'без изменений'), ('Лоджия (к=0,5)', '1,92', 'утеплить')]
tb2 = (f'<table style="font-size:25px;color:{BODY_C};width:790px"><tr><th style="width:44%;text-align:left">Помещение</th><th style="width:20%;text-align:right">м²</th><th style="width:36%;text-align:left">Примечание</th></tr>'
       + ''.join(f'<tr><td>{a}</td><td style="text-align:right">{b}</td><td>{c}</td></tr>' for a, b, c in expl_after) + '</table>')
plan_slide('newplan', n, svg_new, 'План после перепланировки', '07 · Евродвушка', tb2 +
           para('Кухня, гостиная зона и прихожая образуют общее пространство около 24 м², единую кухню-гостиную. Свет приходит из кухонного окна.', 24),
           'Итоговая планировка: спальня 9,7 м² с окном и общая кухня-гостиная. Площади ориентировочные, уточняются после обмера.')

# 9. Мебель
n = nxt()
fl = '<div style="display:grid;grid-template-columns:52px 1fr;gap:10px 14px">' + ''.join(
    f'<p style="font-size:24px;font-weight:700;color:{DARK}">{k}</p><p style="font-size:24px;color:{BODY_C}">{t}</p>' for k, _, t in FURN_NUMS) + '</div>'
plan_slide('furniture', n, svg_furn, 'План мебели', '08 · Меблировка', fl,
           'Расстановка мебели с нумерацией. Пунктиром показаны верхние шкафы кухни.')

# 10. 3D общий вид
n = nxt()
inner = svg_inline(iso_over, 560, 110, 1220, 1000 * 0 + 1000)
inner = svg_inline(iso_over, 620, 96, 1180, 968)
inner += (f'<div style="position:absolute;left:128px;top:128px;width:520px;display:flex;flex-direction:column;gap:24px">'
          + eyebrow('09 · 3D') + h2('Общий вид после перепланировки', 56)
          + para('Стены ближнего плана срезаны на 1 м, чтобы видеть всю квартиру.', 25)
          + para('Слева спальня у окна и гостиная зона. В центре открытая прихожая и кухня. Справа утеплённая лоджия и ванная.', 25) + '</div>')
inner += footer(n, 128, 500)
add('iso-overview', section('iso-overview', inner, 'Аксонометрия всей квартиры: стены на переднем плане срезаны на 1 м.', layout='padding:128px'))

# 11. Гостиная: день / гости
n = nxt()
inner = (eyebrow('10 · 3D · Кухня-гостиная') + h2('Диван-кровать: днём гостиная, ночью гостевая спальня')
         + '<div style="display:flex;flex-direction:row;gap:48px">'
         + f'<div style="flex:1;display:flex;flex-direction:column;gap:12px">{svg_inline(iso_live_day, 0, 0, 0, 0).replace("position:absolute;left:0px;top:0px;width:0px;height:0px", "width:780px;height:604px")}'
         + para('<b>День.</b> Еврокнижка 200 см у стены, журнальный столик, стеллаж-шкаф вдоль перегородки спальни.', 24) + '</div>'
         + f'<div style="flex:1;display:flex;flex-direction:column;gap:12px">{svg_inline(iso_live_guest, 0, 0, 0, 0).replace("position:absolute;left:0px;top:0px;width:0px;height:0px", "width:780px;height:604px")}'
         + para('<b>Гости.</b> Спальное место 160×200, столик убирается. Бельё хранится в ящике дивана.', 24) + '</div></div>')
add('iso-living', section('iso-living', inner + footer(n, 128, 900), 'Два режима гостиной зоны: дневной и гостевой с разложенным диваном.',
                          layout='display:flex;flex-direction:column;gap:20px;padding:96px 128px 128px'))

# 12. Кухня
n = nxt()
inner = svg_inline(iso_kitchen, 96, 150, 1000, 760)
inner += right_col(eyebrow('11 · 3D · Кухня') + h2('Г-образная кухня у стояков', 54)
                   + para('Мойка и варочная панель остаются на прежних местах, поэтому не нужно переносить коммуникации и согласовывать мокрые зоны.', 25)
                   + '<ul style="font-size:25px;line-height:1.45;color:#4A443D">'
                   + '<li>Верхние шкафы до потолка: +30% хранения</li><li>Встроенный холодильник и антресоль</li>'
                   + '<li>Посудомойка 45 см под столешницей</li><li>Стол Ø90 у окна лоджии, подвес над ним</li>'
                   + '<li>Фартук: белая плитка 7,5×30, столешница «дуб»</li></ul>', top=150)
inner += footer(n)
add('iso-kitchen', section('iso-kitchen', inner, 'Кухня: вид от гостиной зоны на окно лоджии и кухонный гарнитур.', layout='padding:128px'))

# 13. Спальня
n = nxt()
inner = svg_inline(iso_bed, 96, 150, 1000, 760)
inner += right_col(eyebrow('12 · 3D · Спальня') + h2('Спальня 9,7 м² у окна', 54)
                   + '<ul style="font-size:25px;line-height:1.45;color:#4A443D">'
                   + '<li>Кровать 160×200 изголовьем к стене, подъёмный механизм и ящик</li>'
                   + '<li>Шкаф-купе 2 м до потолка, фасады цвета стен</li>'
                   + '<li>Подвесные тумбы и бра: пол остаётся свободным</li>'
                   + '<li>У окна место под детскую кроватку 60×120</li>'
                   + '<li>Блэкаут-шторы: окно на солнечную сторону и дорогу</li></ul>', top=150)
inner += footer(n)
add('iso-bedroom', section('iso-bedroom', inner, 'Спальня: вид со стороны окна на изголовье и шкаф-купе вдоль новой перегородки.', layout='padding:128px'))

# 14. Прихожая и ванная
n = nxt()
inner = (eyebrow('13 · 3D · Прихожая и ванная') + h2('Хранение у входа, ванная без переноса стояков')
         + '<div style="display:flex;flex-direction:row;gap:48px">'
         + f'<div style="flex:1;display:flex;flex-direction:column;gap:12px">{svg_inline(iso_hall, 0, 0, 0, 0).replace("position:absolute;left:0px;top:0px;width:0px;height:0px", "width:780px;height:600px")}'
         + para('<b>Прихожая.</b> Шкаф во всю высоту с зеркальной дверью, обувница-скамья, крючки.', 24) + '</div>'
         + f'<div style="flex:1;display:flex;flex-direction:column;gap:12px">{svg_inline(iso_bath, 0, 0, 0, 0).replace("position:absolute;left:0px;top:0px;width:0px;height:0px", "width:780px;height:600px")}'
         + para('<b>Ванная.</b> Стиральная машина под общей столешницей, подвесной унитаз, ванна 160 для ребёнка.', 24) + '</div></div>')
add('iso-hall-bath', section('iso-hall-bath', inner + footer(n, 128, 900), 'Прихожая и ванная в 3D.',
                             layout='display:flex;flex-direction:column;gap:20px;padding:96px 128px 128px'))

# 15. Облёт
n = nxt()
inner = (f'<x-embed style="position:absolute;left:600px;top:80px;width:1260px;height:920px">{esc(EMBED)}</x-embed>'
         if False else f'<x-embed style="position:absolute;left:600px;top:80px;width:1260px;height:920px">{EMBED}</x-embed>')
inner += (f'<div style="position:absolute;left:128px;top:128px;width:460px;display:flex;flex-direction:column;gap:24px">'
          + eyebrow('14 · 3D-облёт') + h2('Модель вращается', 56)
          + para('Живая 3D-модель квартиры: стены срезаны на 1,1 м, мебель в реальных габаритах.', 25)
          + para('В PDF и PPTX здесь будет статичный кадр. Вращение работает в веб-версии презентации.', 24, '#6B645B') + '</div>')
inner += footer(n, 128, 440)
add('flythrough', section('flythrough', inner, 'Живая вращающаяся 3D-модель квартиры (работает в веб-просмотре).', layout='padding:128px'))

# 16. Хранение
n = nxt()
st_items = ['Кровать с подъёмником: ≈0,6 м³', 'Шкаф-купе в спальне 2,0 м', 'Стеллаж-шкаф 1,95 м в гостиной',
            'Ящик дивана для гостевого белья', 'Подвесная ТВ-тумба', 'Кухня: шкафы до потолка',
            'Шкаф в прихожей 1,4 м до потолка', 'Обувница-скамья', 'Шкаф под раковиной',
            'Полка инсталляции + короб', 'Шкаф на утеплённой лоджии']
sl = '<div style="display:grid;grid-template-columns:52px 1fr;gap:8px 14px">' + ''.join(
    f'<p style="font-size:24px;font-weight:700;color:{WOOD}">{i + 1}</p><p style="font-size:24px;color:{BODY_C}">{t}</p>' for i, t in enumerate(st_items)) + '</div>'
plan_slide('storage', n, svg_store, 'Скрытые системы хранения', '15 · Хранение', sl +
           para('Все фасады гладкие, в цвет стен, без ручек (push-to-open). Хранение не бросается в глаза.', 24),
           'Одиннадцать зон скрытого хранения выделены на плане.')

# 17. Водоснабжение
n = nxt()
plan_slide('water', n, svg_water, 'Схема водоснабжения', '16 · Инженерия',
           ''.join(legend_row(c, t, 'round') for c, t in [(BLUE, 'ХВС, PP-R Ø20 / PEX 16'), (RED, 'ГВС, PP-R Ø20 / PEX 16')]) +
           f'<table style="font-size:24px;color:{BODY_C};width:790px"><tr><th style="width:22%;text-align:left">Точка</th><th style="width:52%;text-align:left">Прибор</th><th style="width:26%;text-align:right">Высота</th></tr>'
           + ''.join(f'<tr><td>{a}</td><td>{b}</td><td style="text-align:right">{c}</td></tr>' for a, b, c in [
               ('М', 'мойка, Х+Г', '500 мм'), ('ПММ', 'посудомойка, Х', '300 мм'), ('Р', 'раковина, Х+Г', '550 мм'),
               ('СМ', 'стиральная машина, Х', '600 мм'), ('У', 'инсталляция, Х', 'по паспорту'), ('В', 'смеситель ванны, Х+Г', '800 мм')]) + '</table>'
           + para('В коробе стояков: краны, фильтры, счётчики, редуктор давления, коллектор. Опционально система защиты от протечек.', 24),
           'Разводка водоснабжения от стояков в ванной. Мокрые точки остаются на прежних местах.')

# 18. Канализация
n = nxt()
plan_slide('sewer', n, svg_sewer, 'Схема канализации', '17 · Инженерия',
           legend_row(BROWN, 'Канализация ПП: Ø110 унитаз, Ø50 остальные', 'round') +
           '<ul style="font-size:25px;line-height:1.45;color:#4A443D">'
           '<li>Стояк К1 Ø110 остаётся на месте, к нему обеспечен доступ через люк в коробе</li>'
           '<li>Уклон Ø50 не меньше 3%, Ø110 не меньше 2%</li>'
           '<li>Мойка: трасса около 1,1 м через стену в короб, посудомойка через сифон мойки</li>'
           '<li>Унитаз подвесной: отвод Ø110 ≈0,6 м</li>'
           '<li>Ванна и стиральная машина через сифоны с обратным клапаном</li>'
           '<li>Ревизии на поворотах, шумопоглощающие трубы в коробе</li></ul>',
           'Канализация: все приборы на прежних местах, короткие трассы к стояку.')

# 19. Розетки
n = nxt()
plan_slide('sockets', n, svg_sock, 'Схема розеток и силовых линий', '18 · Электрика',
           ''.join(legend_row(c, t, sh) for c, t, sh in [
               (ORANGE, 'розетка 16 А (×2 = двойная)', 'round'), (POWER, 'силовая: варочная 32 А / кондиционер', 'square'),
               (IP, 'розетка IP44 с отдельным УЗО 10 мА', 'round'), (LOW, 'слаботочка: интернет / ТВ', 'square')]) +
           f'<table style="font-size:24px;color:{BODY_C};width:790px"><tr><th style="width:62%;text-align:left">Где</th><th style="width:38%;text-align:right">Высота</th></tr>'
           + ''.join(f'<tr><td>{a}</td><td style="text-align:right">{b}</td></tr>' for a, b in [
               ('Над столешницей', '1100 мм'), ('Встроенная техника', '300 мм'), ('У кровати', '700 мм'),
               ('Стандарт / за ТВ', '300 / 1100 мм'), ('Кондиционер', '2200 мм'), ('Стиральная машина', '1000 мм')]) + '</table>',
           'Розетки и силовые линии. Два кондиционера, потому что квартира на солнечной стороне.')

# 20. Освещение
n = nxt()
plan_slide('lighting', n, svg_light, 'Схема освещения', '19 · Свет',
           ''.join(legend_row(c, t, sh) for c, t, sh in [
               ('#FFF6DE', 'точечный светильник GX53 / LED', 'round'), (YEL, 'подвес, бра, торшер', 'round'),
               ('#E6B85C', 'LED-лента (кухня, шкафы, спальня)', 'square'), (INK, 'выключатель, номер группы', 'round')]) +
           para('<b>1</b> вход: свет прихожей + мастер-выключатель. <b>2</b> кухня: споты, подвес, лента. <b>3</b> гостиная зона. '
                '<b>4–5</b> спальня, проходные: у двери и у кровати. <b>6</b> ванная: свет + вентилятор с таймером. <b>7</b> лоджия.', 24) +
           para('Свет 3000 К во всей квартире, на кухне над рабочей зоной 4000 К. Подсветка шкафов с датчиком открытия.', 24),
           'Сценарии освещения: общий, рабочий, акцентный. Проходные выключатели в спальне.')

# 21. Щит
n = nxt()
groups = [
    ('QF0', 'Ввод + реле напряжения', '2P 40 А', '3×10'),
    ('QF1', 'Свет: кухня, гостиная, прихожая', '1P 10 А', '3×1,5'),
    ('QF2', 'Свет: спальня, ванная, лоджия', '1P 10 А', '3×1,5'),
    ('QD1', 'Варочная панель', 'дифф. 32 А 30 мА', '3×6'),
    ('QD2', 'Кухня: розетки столешницы', 'дифф. 16 А 30 мА', '3×2,5'),
    ('QD3', 'Духовка/СВЧ, посудомойка', 'дифф. 16 А 30 мА', '3×2,5'),
    ('QD4', 'Холодильник', 'дифф. 16 А 30 мА', '3×2,5'),
    ('QD5', 'Стиральная машина', 'дифф. 16 А 10 мА', '3×2,5'),
    ('QD6', 'Ванная: розетки, полотенцесушитель', 'дифф. 16 А 10 мА', '3×2,5'),
    ('QD7', 'Розетки: гостиная, прихожая, лоджия', 'дифф. 16 А 30 мА', '3×2,5'),
    ('QD8', 'Розетки: спальня', 'дифф. 16 А 30 мА', '3×2,5'),
    ('QF3', 'Кондиционеры (2 шт.)', '1P 16 А', '3×2,5'),
]
tbl = (f'<table style="font-size:24px;color:{BODY_C};width:1664px"><tr><th style="width:9%;text-align:left">Поз.</th>'
       f'<th style="width:45%;text-align:left">Группа</th><th style="width:28%;text-align:left">Защита</th><th style="width:18%;text-align:left">Кабель ВВГнг-LS</th></tr>'
       + ''.join(f'<tr><td>{a}</td><td>{b}</td><td>{c}</td><td>{d}</td></tr>' for a, b, c, d in groups) + '</table>')
add('panel', section('panel', eyebrow('20 · Электрика') + h2('Щит и группы') + tbl +
                     para('Выделенная мощность и номинал вводного автомата уточняются по техусловиям дома. Щит на 24 модуля с запасом в 3–4 места.', 24, '#6B645B') + footer(n, 128, 900),
                     'Однолинейная схема щита в табличной форме.', layout='display:flex;flex-direction:column;gap:20px;padding:96px 128px 150px'))

# 22. Отделка
n = nxt()
fin = [
    ('Кухня-гостиная', 'SPC 4–5 мм «светлый дуб», 43 класс', 'Моющаяся краска, тёплый белый', 'Натяжной матовый'),
    ('Спальня', 'SPC «светлый дуб»', 'Краска + стеновые рейки за изголовьем', 'Натяжной матовый'),
    ('Прихожая', 'SPC «светлый дуб», без порогов', 'Моющаяся краска', 'Натяжной матовый'),
    ('Ванная', 'Керамогранит 60×60 светлый', 'Плитка 30×60 белая + 1 стена «под дерево»', 'Натяжной, IP-споты'),
    ('Лоджия', 'Утеплённый пол + SPC', 'Вагонка/ГКЛ + краска', 'Покраска'),
]
tbl = (f'<table style="font-size:24px;color:{BODY_C};width:1664px"><tr><th style="width:16%;text-align:left">Помещение</th>'
       f'<th style="width:28%;text-align:left">Пол</th><th style="width:34%;text-align:left">Стены</th><th style="width:22%;text-align:left">Потолок</th></tr>'
       + ''.join(f'<tr><td><b>{a}</b></td><td>{b}</td><td>{c}</td><td>{d}</td></tr>' for a, b, c, d in fin) + '</table>')
add('finishes', section('finishes', eyebrow('21 · Материалы') + h2('Ведомость отделки') + tbl +
                        para('Двери: белые, скрытая коробка или эмаль, высота 2,0–2,1 м. Плинтус белый МДФ 80 мм в цвет стен.', 24) + footer(n, 128, 900),
                        'Ведомость отделки по помещениям.', layout='display:flex;flex-direction:column;gap:28px;padding:110px 128px 150px'))

# 23. Солнце и шум
n = nxt()
cards = [
    ('Солнце', ['Рулонные шторы «день-ночь» в кассете + льняные шторы', 'В спальне блэкаут', 'Солнцезащитный или мультифункциональный стеклопакет',
                'Два кондиционера: в кухне-гостиной и в спальне', 'Матовые поверхности, без бликов']),
    ('Шум дороги', ['Шумозащитные двухкамерные стеклопакеты с разной толщиной стёкол', 'Приточный клапан с шумоглушителем: проветривание без открытого окна',
                    'Пена и герметизация откосов', 'Плотный текстиль и ковёр', 'Звукоизоляция перегородки спальни']),
]
row = '<div style="display:flex;flex-direction:row;gap:40px">' + ''.join(
    f'<div style="flex:1;display:flex;flex-direction:column;gap:16px;background:#FBFAF6;padding:44px;border-radius:16px;border:1px solid #DDD5C8">'
    f'<h3 style="font-family:{HEAD};font-size:44px;font-weight:500;color:{WOOD}">{a}</h3>'
    f'<ul style="font-size:25px;line-height:1.45;color:{BODY_C}">' + ''.join(f'<li>{x}</li>' for x in b) + '</ul></div>' for a, b in cards) + '</div>'
add('sun-noise', section('sun-noise', eyebrow('22 · Особенности квартиры') + h2('Солнечная сторона и центральная дорога') + row + footer(n, 128, 900),
                         'Решения под ориентацию квартиры: защита от солнца и шума.', bg=BG2,
                         layout='display:flex;flex-direction:column;gap:36px;padding:110px 128px 150px'))

# 24. Бюджет
n = nxt()
bud = [('Черновые работы и материалы', '35%', 'Демонтаж, перегородка, электрика, сантехника, выравнивание'),
       ('Чистовые материалы', '20%', 'SPC, краска, плитка в ванной, натяжные потолки'),
       ('Кухня и встроенные шкафы', '25%', 'ЛДСП/МДФ, фурнитура Blum-аналог, столешница HPL'),
       ('Сантехника и техника', '12%', 'Инсталляция, ванна, смесители, кондиционеры'),
       ('Свет, текстиль, декор', '8%', 'Споты, подвесы, шторы, ковры')]
bars = ''.join(
    f'<div style="display:flex;flex-direction:row;gap:24px;align-items:center">'
    f'<p style="width:420px;font-size:25px;color:{DARK}">{a}</p>'
    f'<div style="width:{int(int(b[:-1]) * 16)}px;height:40px;background:{WOOD if i == 2 else "#C79E6E"};border-radius:6px"></div>'
    f'<p style="width:90px;font-size:28px;font-weight:700;color:{DARK}">{b}</p>'
    f'<p style="flex:1;font-size:24px;color:#6B645B">{c}</p></div>' for i, (a, b, c) in enumerate(bud))
save = ('<div style="display:flex;flex-direction:row;gap:40px">'
        f'<div style="flex:1;display:flex;flex-direction:column;gap:8px"><h3 style="font-family:{HEAD};font-size:34px;font-weight:500;color:{SAGE}">Экономим</h3>'
        f'{para("Мокрые зоны не переносим. Одна перегородка из ГКЛ. Натяжные потолки. Корпуса ЛДСП. Фасады — крашеный МДФ только на кухне.", 24)}</div>'
        f'<div style="flex:1;display:flex;flex-direction:column;gap:8px"><h3 style="font-family:{HEAD};font-size:34px;font-weight:500;color:{WOOD}">Не экономим</h3>'
        f'{para("Электрика и щит, гидроизоляция ванной, звукоизоляция спальни, стеклопакеты, механизм дивана и кровати.", 24)}</div></div>')
add('budget', section('budget', eyebrow('23 · Бюджет') + h2('Структура эконом-бюджета') + bars + save +
                      para('Итоговую сумму [__] посчитаю после выбора материалов и смет подрядчиков вашего города.', 24, '#6B645B') + footer(n, 128, 900),
                      'Ориентировочная структура бюджета в процентах. Абсолютные цифры зависят от города и подрядчика.',
                      layout='display:flex;flex-direction:column;gap:22px;padding:100px 128px 150px'))

# 25. Этапы работ
n = nxt()
steps = [('1', 'Обмер и проект', 'Обмер, рабочие чертежи, проект перепланировки'),
         ('2', 'Согласование', 'Проект + техпаспорт → жилинспекция / местный орган'),
         ('3', 'Демонтаж и стены', 'Снос 3 участков, перегородка ГКЛ, утепление лоджии'),
         ('4', 'Инженерия', 'Электрика, щит, водопровод, канализация, короба'),
         ('5', 'Черновая отделка', 'Штукатурка, стяжка, гидроизоляция, плитка'),
         ('6', 'Чистовая отделка', 'Покраска, SPC, потолки, двери'),
         ('7', 'Мебель и свет', 'Кухня, шкафы, светильники, текстиль'),
         ('8', 'Приёмка', 'Акт о завершении перепланировки, новый техпаспорт')]
grid = '<div style="display:grid;grid-template-columns:repeat(4, 1fr);gap:28px">' + ''.join(
    f'<div style="display:flex;flex-direction:column;gap:10px;background:#FBFAF6;padding:32px;border-radius:14px;border:1px solid #DDD5C8">'
    f'<p style="font-family:{HEAD};font-size:56px;color:{WOOD};line-height:1">{a}</p>'
    f'<h3 style="font-size:28px;font-weight:700;color:{DARK}">{b}</h3><p style="font-size:24px;line-height:1.35;color:{BODY_C}">{c}</p></div>'
    for a, b, c in steps) + '</div>'
add('works', section('works', eyebrow('24 · Порядок') + h2('Этапы ремонта и согласования') + grid +
                     para('Ориентировочно 2,5–3,5 месяца работ плюс время на согласование.', 24, '#6B645B') + footer(n, 128, 900),
                     'Последовательность работ от обмера до приёмки.', layout='display:flex;flex-direction:column;gap:32px;padding:110px 128px 150px'))

# 26. Следующие шаги
n = nxt()
inner = (eyebrow('25 · Дальше', '#D9BC93')
         + f'<h2 style="font-family:{HEAD};font-size:72px;font-weight:400;line-height:1.1;color:#F4F1EA">Что нужно от вас, чтобы перейти к рабочему проекту</h2>'
         + f'<ul style="font-size:30px;line-height:1.5;color:#E5DDD0">'
         + '<li>Техпаспорт БТИ или план с размерами и отметкой несущих стен</li>'
         + '<li>Тип дома (панель, монолит, кирпич), этаж и тип плиты: газ или электричество</li>'
         + '<li>Подтвердить: ванна или душ, нужна ли детская кроватка сейчас</li>'
         + '<li>Город, чтобы привязать бюджет к местным ценам</li></ul>'
         + para('После этого уточню размеры и обновлю все схемы.', 26, '#D9BC93'))
add('next', section('next', inner + footer(n, 128, 900, '#BFB6A8'), 'Финальный слайд: что требуется от заказчика.', bg=DARK, color='#F4F1EA',
                    layout='display:flex;flex-direction:column;gap:36px;padding:128px 128px 160px;justify-content:center'))

# ================= ЗАПИСЬ =================
os.makedirs(os.path.join(ROOT, 'project', 'slides'), exist_ok=True)
order = [sid for sid, _ in slides]
deck = {
    'v': 4,
    'createdOnFiles': {'v': 1, 'at': '2026-10-07T17:30:00Z'},
    'lists': 'css',
    'title': 'Евродвушка 38,6 м² — дизайн-проект',
    'order': order,
    'cover': 'cover',
    'sections': {
        's1': {'description': 'Задача, анализ плана и концепция', 'start': 'cover'},
        's2': {'description': 'Перепланировка: обмер, демонтаж, монтаж, итоговый план, мебель', 'start': 'existing'},
        's3': {'description': '3D-визуализация', 'start': 'iso-overview'},
        's4': {'description': 'Хранение и инженерные схемы', 'start': 'storage'},
        's5': {'description': 'Материалы, бюджет и этапы', 'start': 'finishes'},
    },
    'faces': {
        'jost': {'family': 'Jost', 'href': 'https://fonts.googleapis.com/css2?family=Jost:wght@300..600&display=swap'},
        'manrope': {'family': 'Manrope', 'href': 'https://fonts.googleapis.com/css2?family=Manrope:wght@400..700&display=swap'},
    },
    'designSystems': [],
}
with open(os.path.join(ROOT, 'project', 'deck.json'), 'w', encoding='utf-8') as f:
    json.dump(deck, f, ensure_ascii=False, indent=1)
for sid, html in slides:
    with open(os.path.join(ROOT, 'project', 'slides', sid + '.html'), 'w', encoding='utf-8') as f:
        f.write(html)
print(len(slides), 'slides;', 'embed KB:', len(EMBED.encode()) // 1024)
for sid, html in slides:
    print(sid, len(html.encode()) // 1024, 'KB')
