"""Отрисовка: обмерные/рабочие планы (SVG) и изометрия (SVG)."""
import math
from geometry import WALLS, OPENINGS, H, C, FLOORS, furniture

S = 100  # px на метр
VB = (-115, -135, 900, 940)
FONT = 'Arial, Helvetica, sans-serif'

INK = '#2B2824'
BEAR = '#3A3631'
PART = '#9A9286'
NEWC = '#4F7A6B'
DEM = '#C0392B'
MUTE = '#8A8278'


def p(v):
    return f'{v * S:.1f}'.rstrip('0').rstrip('.')


def horiz(w):
    x0, y0, x1, y1, _ = WALLS[w]
    return (x1 - x0) >= (y1 - y0)


def ops_for(w, include_old=True):
    return [o for o in OPENINGS if o['w'] == w]


def wall_pieces(w):
    """Куски стены без проёмов (в метрах)."""
    x0, y0, x1, y1, _ = WALLS[w]
    hz = horiz(w)
    a, b = (x0, x1) if hz else (y0, y1)
    cuts = sorted((o['a0'], o['a1']) for o in ops_for(w))
    pieces, cur = [], a
    for c0, c1 in cuts:
        if c0 > cur:
            pieces.append((cur, c0))
        cur = max(cur, c1)
    if cur < b:
        pieces.append((cur, b))
    out = []
    for s0, s1 in pieces:
        out.append((s0, y0, s1, y1) if hz else (x0, s0, x1, s1))
    return out


def rect(x0, y0, x1, y1, fill='none', stroke='none', sw=1, extra=''):
    return (f'<rect x="{p(x0)}" y="{p(y0)}" width="{p(x1 - x0)}" height="{p(y1 - y0)}" '
            f'fill="{fill}" stroke="{stroke}" stroke-width="{sw}"{extra}/>')


def line(x0, y0, x1, y1, stroke=INK, sw=1.5, extra=''):
    return f'<line x1="{p(x0)}" y1="{p(y0)}" x2="{p(x1)}" y2="{p(y1)}" stroke="{stroke}" stroke-width="{sw}"{extra}/>'


def text(x, y, s, size=20, fill=INK, weight=400, anchor='middle', extra=''):
    s = s.replace('&', '&amp;').replace('<', '&lt;')
    return (f'<text x="{p(x)}" y="{p(y)}" font-family="{FONT}" font-size="{size}" '
            f'font-weight="{weight}" fill="{fill}" text-anchor="{anchor}"{extra}>{s}</text>')


def hatch_defs():
    return ('<defs>'
            f'<pattern id="hd" width="10" height="10" patternUnits="userSpaceOnUse" patternTransform="rotate(45)">'
            f'<rect width="10" height="10" fill="#F6DCD7"/><line x1="0" y1="0" x2="0" y2="10" stroke="{DEM}" stroke-width="3"/></pattern>'
            f'<pattern id="hn" width="10" height="10" patternUnits="userSpaceOnUse" patternTransform="rotate(-45)">'
            f'<rect width="10" height="10" fill="#D6E5DE"/><line x1="0" y1="0" x2="0" y2="10" stroke="{NEWC}" stroke-width="3"/></pattern>'
            '</defs>')


def door_svg(o, color=INK):
    x0, y0, x1, y1, _ = WALLS[o['w']]
    hz = horiz(o['w'])
    a0, a1 = o['a0'], o['a1']
    wdt = a1 - a0
    side = o.get('side', 1)
    hinge = a1 if o.get('hinge') == 'a1' else a0
    other = a0 if hinge == a1 else a1
    out = []
    if hz:
        face = y1 if side > 0 else y0
        hx, hy = hinge, face
        lx, ly = hinge, face + side * wdt
        pts = []
        for i in range(13):
            t = i / 12 * math.pi / 2
            px = hinge + (other - hinge) * math.sin(t)
            py = face + side * wdt * math.cos(t)
            pts.append((px, py))
    else:
        face = x1 if side > 0 else x0
        hx, hy = face, hinge
        lx, ly = face + side * wdt, hinge
        pts = []
        for i in range(13):
            t = i / 12 * math.pi / 2
            py = hinge + (other - hinge) * math.sin(t)
            px = face + side * wdt * math.cos(t)
            pts.append((px, py))
    out.append(line(hx, hy, lx, ly, color, 3))
    out.append('<polyline points="' + ' '.join(f'{p(a)},{p(b)}' for a, b in pts) +
               f'" fill="none" stroke="{color}" stroke-width="1.2" stroke-dasharray="5 4"/>')
    return ''.join(out)


def window_svg(o):
    x0, y0, x1, y1, _ = WALLS[o['w']]
    out = []
    if horiz(o['w']):
        out.append(rect(o['a0'], y0, o['a1'], y1, '#FBFBF8', INK, 1.2))
        mid = (y0 + y1) / 2
        out.append(line(o['a0'], mid, o['a1'], mid, '#5B8FA3', 2.5))
    else:
        out.append(rect(x0, o['a0'], x1, o['a1'], '#FBFBF8', INK, 1.2))
        mid = (x0 + x1) / 2
        out.append(line(mid, o['a0'], mid, o['a1'], '#5B8FA3', 2.5))
    return ''.join(out)


def walls_svg(wall_ids, style='normal', demolish=(), new=()):
    """style: normal | faint."""
    out = []
    for w in wall_ids:
        kind = WALLS[w][4]
        for pc in wall_pieces(w):
            if w in demolish:
                out.append(rect(*pc, 'url(#hd)', DEM, 1.5, ' stroke-dasharray="6 4"'))
            elif w in new:
                out.append(rect(*pc, 'url(#hn)', NEWC, 1.5))
            elif style == 'faint':
                out.append(rect(*pc, '#D8D2C8' if kind == 'ext' else '#E6E1D8', '#B8B0A4', 1))
            else:
                out.append(rect(*pc, BEAR if kind == 'ext' else PART, BEAR if kind == 'ext' else '#6E675D', 1))
    for o in OPENINGS:
        if o['w'] not in wall_ids:
            continue
        if o['k'] in ('win', 'glaze'):
            out.append(window_svg(o))
        elif o['k'] == 'door' and o['w'] not in demolish:
            out.append(door_svg(o, MUTE if style == 'faint' else INK))
        elif o['k'] == 'door' and o['w'] in demolish:
            out.append(door_svg(o, DEM))
    return ''.join(out)


def dims_svg():
    out = []
    ext_col = '#6B645B'

    def hdim(xa, xb, y, label):
        out.append(line(xa, y, xb, y, ext_col, 1.2))
        for x in (xa, xb):
            out.append(line(x - 0.06, y + 0.06, x + 0.06, y - 0.06, ext_col, 2))
            out.append(line(x, y - 0.1, x, y + 0.1, ext_col, 1))
        out.append(text((xa + xb) / 2, y - 0.08, label, 18, ext_col))

    def vdim(ya, yb, x, label):
        out.append(line(x, ya, x, yb, ext_col, 1.2))
        for y in (ya, yb):
            out.append(line(x - 0.06, y + 0.06, x + 0.06, y - 0.06, ext_col, 2))
            out.append(line(x - 0.1, y, x + 0.1, y, ext_col, 1))
        cx, cy = x - 0.1, (ya + yb) / 2
        out.append(text(cx, cy, label, 18, ext_col, extra=f' transform="rotate(-90 {p(cx)} {p(cy)})"'))

    hdim(0, 2.88, -0.75, '2880')
    hdim(3.0, 6.02, -0.75, '3020')
    hdim(0, 6.75, -1.1, '6750')
    vdim(0, 6.79, -0.65, '6790')
    vdim(1.57, 4.12, 7.35, '2550')
    vdim(4.24, 6.79, 7.35, '2550')
    out_txt = ''.join(out)
    # подписи справа повернуть корректно
    return out_txt


def rooms_svg(rooms, big=True, area=True):
    out = []
    for name, ar, _, (cx, cy) in rooms:
        out.append(text(cx, cy, name, 22 if big else 18, INK, 600))
        if area:
            out.append(text(cx, cy + 0.3, ar + ' м²', 20 if big else 16, '#6B645B'))
    return ''.join(out)


def furn_svg(items, cats=None, fill_map=None, numbered=None, faint=False):
    out = []
    fills = dict(furn='#EFE6D6', store='#E9D7BC', kit='#E4E9E3', san='#E3EBEE', deco='#F1ECE3')
    if fill_map:
        fills.update(fill_map)
    for it in items:
        x0, y0, x1, y1, z0, z1, col, label, cat = it
        if cat == 'deco':
            continue
        if cats and cat not in cats:
            continue
        f = fills.get(cat, '#EEE')
        st = '#7C7469' if not faint else '#C9C1B5'
        out.append(rect(x0, y0, x1, y1, f if not faint else '#F3EFE8', st, 1.2))
    if numbered:
        for n, (x, y) in numbered:
            out.append(f'<circle cx="{p(x)}" cy="{p(y)}" r="14" fill="{INK}"/>')
            out.append(text(x, y + 0.065, str(n), 16, '#FBFBF8', 700))
    return ''.join(out)


def floors_svg(tint=True):
    out = []
    for (r, col) in FLOORS:
        out.append(rect(*r, '#F7F1E7' if col == C['floor'] else '#EDEDEA'))
    return ''.join(out)


def svg_wrap(body, label, vb=VB):
    x, y, w, h = vb
    return (f'<svg xmlns="http://www.w3.org/2000/svg" aria-label="{label}" width="{w}" height="{h}" '
            f'viewBox="{x} {y} {w} {h}">' + hatch_defs() +
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#FBFAF6"/>' + body + '</svg>')


# ============================ ИЗОМЕТРИЯ ============================

def _rot(view, x, y):
    if view == 'pp':
        return x, y
    if view == 'mp':
        return y, -x
    if view == 'pm':
        return -y, x
    if view == 'mm':
        return -x, -y
    raise ValueError(view)


def _shade(hexc, k):
    hexc = hexc.lstrip('#')
    r, g, b = (int(hexc[i:i + 2], 16) for i in (0, 2, 4))
    return '#%02x%02x%02x' % tuple(max(0, min(255, int(v * k))) for v in (r, g, b))


def wall_boxes(wall_ids, zcut=None):
    """3D-коробки стен с проёмами. zcut(w) -> высота обреза."""
    boxes = []
    for w in wall_ids:
        x0, y0, x1, y1, kind = WALLS[w]
        hz = horiz(w)
        top = zcut(w) if zcut else H
        for pc in wall_pieces(w):
            boxes.append((*pc, 0.0, top, C['wall'], 'wall'))
        for o in ops_for(w):
            a0, a1 = o['a0'], o['a1']
            r = (a0, y0, a1, y1) if hz else (x0, a0, x1, a1)
            if o['k'] in ('win', 'glaze'):
                sill = 0.85 if o['k'] == 'win' else 0.9
                head = 2.25 if o['k'] == 'win' else 2.4
                boxes.append((*r, 0.0, min(sill, top), C['wall'], 'wall'))
                if top > head:
                    boxes.append((*r, head, top, C['wall'], 'wall'))
                gl = ((a0, (y0 + y1) / 2 - 0.02, a1, (y0 + y1) / 2 + 0.02) if hz else
                      ((x0 + x1) / 2 - 0.02, a0, (x0 + x1) / 2 + 0.02, a1))
                if top > sill:
                    boxes.append((*gl, sill, min(head, top), C['glass'], 'glass'))
            elif o['k'] in ('door', 'gap'):
                if top > 2.1:
                    boxes.append((*r, 2.1, top, C['wall'], 'wall'))
    return boxes


def iso_svg(boxes, view, region, label, size=(1000, 760), back_pad=0.5, cut=1.0, pad=30, furn_cut=None):
    """boxes: (x0,y0,x1,y1,z0,z1,color,kind). region: (x0,y0,x1,y1) — обрезка."""
    rx0, ry0, rx1, ry1 = region
    rb = []
    for b in boxes:
        x0, y0, x1, y1, z0, z1, col, kind = b
        x0, y0 = max(x0, rx0), max(y0, ry0)
        x1, y1 = min(x1, rx1), min(y1, ry1)
        if x1 - x0 <= 1e-6 or y1 - y0 <= 1e-6:
            continue
        ax, ay = _rot(view, x0, y0)
        bx, by = _rot(view, x1, y1)
        rb.append([min(ax, bx), min(ay, by), z0, max(ax, bx), max(ay, by), z1, col, kind])
    # обрезка стен ближнего плана
    cx0 = min(b[0] for b in rb)
    cy0 = min(b[1] for b in rb)
    for b in rb:
        if b[7] in ('wall', 'glass'):
            back = b[3] <= cx0 + back_pad or b[4] <= cy0 + back_pad
            if not back:
                b[5] = min(b[5], cut)
                b[7] = 'wallcut' if b[7] == 'wall' else b[7]
        elif furn_cut is not None and b[7] == 'furn':
            back = b[3] <= cx0 + back_pad or b[4] <= cy0 + back_pad
            if not back and b[2] >= furn_cut:
                b[5] = b[2]
    rb = [b for b in rb if b[5] > b[2] + 1e-6]

    c30, s30 = math.cos(math.pi / 6), 0.5

    def proj(x, y, z):
        return ((x - y) * c30, (x + y) * s30 - z)

    def bbox2(b):
        pts = [proj(x, y, z) for x in (b[0], b[3]) for y in (b[1], b[4]) for z in (b[2], b[5])]
        xs, ys = [q[0] for q in pts], [q[1] for q in pts]
        return min(xs), min(ys), max(xs), max(ys)

    bb = [bbox2(b) for b in rb]
    eps = 1e-6

    def in_front(a, b):
        if a[0] >= b[3] - eps:
            return 1
        if b[0] >= a[3] - eps:
            return -1
        if a[1] >= b[4] - eps:
            return 1
        if b[1] >= a[4] - eps:
            return -1
        if a[2] >= b[5] - eps:
            return 1
        if b[2] >= a[5] - eps:
            return -1
        sa = a[0] + a[3] + a[1] + a[4] + a[2] + a[5]
        sb = b[0] + b[3] + b[1] + b[4] + b[2] + b[5]
        return 1 if sa > sb else -1

    n = len(rb)
    behind = [set() for _ in range(n)]  # behind[i] = те, кто должен рисоваться раньше i
    for i in range(n):
        for j in range(i + 1, n):
            A, B = bb[i], bb[j]
            if A[2] <= B[0] or B[2] <= A[0] or A[3] <= B[1] or B[3] <= A[1]:
                continue
            r = in_front(rb[i], rb[j])
            if r > 0:
                behind[i].add(j)
            else:
                behind[j].add(i)
    order, done = [], [False] * n
    remaining = set(range(n))
    while remaining:
        ready = [i for i in remaining if all(done[j] for j in behind[i])]
        if not ready:
            ready = [min(remaining, key=lambda i: rb[i][0] + rb[i][1] + rb[i][2])]
        ready.sort(key=lambda i: rb[i][0] + rb[i][1] + rb[i][2])
        for i in ready:
            done[i] = True
            remaining.discard(i)
            order.append(i)

    allx = [v for B in bb for v in (B[0], B[2])]
    ally = [v for B in bb for v in (B[1], B[3])]
    mx0, mx1, my0, my1 = min(allx), max(allx), min(ally), max(ally)
    W, Hh = size
    sc = min((W - 2 * pad) / (mx1 - mx0), (Hh - 2 * pad) / (my1 - my0))
    ox = (W - (mx1 - mx0) * sc) / 2 - mx0 * sc
    oy = (Hh - (my1 - my0) * sc) / 2 - my0 * sc

    def P(x, y, z):
        a, b = proj(x, y, z)
        return f'{a * sc + ox:.1f},{b * sc + oy:.1f}'

    out = []
    for i in order:
        x0, y0, z0, x1, y1, z1, col, kind = rb[i]
        top_col = col
        if kind == 'wallcut':
            top_col = C['wallcap']
        faces = [
            ([(x0, y0, z1), (x1, y0, z1), (x1, y1, z1), (x0, y1, z1)], top_col, 1.0),
            ([(x1, y0, z0), (x1, y1, z0), (x1, y1, z1), (x1, y0, z1)], col, 0.86),
            ([(x0, y1, z0), (x1, y1, z0), (x1, y1, z1), (x0, y1, z1)], col, 0.74),
        ]
        op = ' fill-opacity="0.45"' if kind == 'glass' else ''
        for pts, fc, k in faces:
            fill = _shade(fc, k)
            out.append(f'<polygon points="{" ".join(P(*q) for q in pts)}" fill="{fill}"{op} '
                       f'stroke="{_shade(fc, k * 0.82)}" stroke-width="0.6" stroke-linejoin="round"/>')
    return (f'<svg xmlns="http://www.w3.org/2000/svg" aria-label="{label}" width="{W}" height="{Hh}" '
            f'viewBox="0 0 {W} {Hh}">' + ''.join(out) + '</svg>')


def scene(mode='day', wall_ids=None):
    from geometry import AFTER
    wall_ids = wall_ids or AFTER
    boxes = []
    for r, col in FLOORS:
        boxes.append((*r, -0.06, 0.0, col, 'floor'))
    boxes += wall_boxes(wall_ids)
    for it in furniture(mode):
        x0, y0, x1, y1, z0, z1, col, _, cat = it
        kind = 'glass' if col == C['glass'] else 'furn'
        boxes.append((x0, y0, x1, y1, z0, z1, col, kind))
    return boxes
