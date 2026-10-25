#!/usr/bin/env python
"""Build a slide-master PPTX from a template-spec.json (theme, master background, named layouts).

Usage: python spec-to-master.py <template-spec.json> <assetsDir> <out.pptx>

Theme colours come from spec.observed.colours (dk1 = body, accent1 = title, accent2 = accent),
lt2 from the off-white decoration fill when the spec has one, the rest from spec.palette.
Fonts: major/minor latin and ea from spec.observed.fonts (title/body).
Each spec layout becomes a layout of the same name: background, decorations (pictures and
filled rectangles, behind the placeholders) and placeholders. Title/body/subtitle placeholders
take position and text style from spec.observed.layoutStyles[name] when present.
"""
import json
import os
import sys
from collections import Counter

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.oxml.shapes.picture import CT_Picture

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS = f'xmlns:p="{P}" xmlns:a="{A}"'
PROMPT = {'title': 'Click to edit title', 'ctrTitle': 'Click to edit title', 'subTitle': 'Click to edit subtitle',
          'body': 'Click to edit text', 'obj': 'Click to edit text', 'pic': 'Picture', 'ftr': '', 'dt': ''}
# spec placeholder type -> observed.layoutStyles role
ROLE = {'title': 'title', 'ctrTitle': 'title', 'body': 'body', 'obj': 'body', 'subTitle': 'subtitle'}


def hx(c):
    return c.lstrip('#').upper()


def scale_x(el, f):
    for xfrm in el.iter(qn('a:xfrm')):
        off, ext = xfrm.find(qn('a:off')), xfrm.find(qn('a:ext'))
        if off is not None:
            off.set('x', str(int(int(off.get('x')) * f)))
        if ext is not None:
            ext.set('cx', str(int(int(ext.get('cx')) * f)))


def next_id(tree):
    return max([int(e.get('id')) for e in tree.iter(qn('p:cNvPr'))] + [1]) + 1


def set_bg(el, fill_xml):
    csld = el.find(qn('p:cSld'))
    old = csld.find(qn('p:bg'))
    if old is not None:
        csld.remove(old)
    csld.insert(0, etree.fromstring(f'<p:bg {NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                                    f'<p:bgPr>{fill_xml}<a:effectLst/></p:bgPr></p:bg>'))


def style_xml(st, latin, no_bullet=False):
    """lstStyle level-1 paragraph/run defaults from a merged style dict."""
    if not st:
        return '<a:lstStyle/>'
    algn = f' algn="{st["align"]}"' if st.get('align') else ''
    if no_bullet:
        algn += ' marL="0" indent="0"'
    attrs = ''
    if st.get('size'):
        attrs += f' sz="{int(round(st["size"] * 100))}"'
    if st.get('bold') is not None:
        attrs += f' b="{1 if st["bold"] else 0}"'
    bu = '<a:buNone/>' if no_bullet else ''
    inner = f'<a:solidFill><a:srgbClr val="{hx(st["colour"])}"/></a:solidFill>' if st.get('colour') else ''
    face = st.get('latin') or latin
    if face:
        inner += f'<a:latin typeface="{face}"/><a:ea typeface="{face}"/>'
    return f'<a:lstStyle><a:lvl1pPr{algn}>{bu}<a:defRPr{attrs}>{inner}</a:defRPr></a:lvl1pPr></a:lstStyle>'


def add_placeholder(tree, ph_type, idx, box, st, latin, wrap=True):
    sid = next_id(tree)
    t = '' if ph_type == 'obj' else f' type="{ph_type}"'
    i = '' if ph_type in ('title', 'ctrTitle') else f' idx="{idx}"'
    if ph_type == 'sldNum':
        run = '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:t>&#8249;#&#8250;</a:t></a:fld>'
    else:
        txt = PROMPT.get(ph_type, '')
        run = f'<a:r><a:rPr lang="ko-KR" altLang="en-US"/><a:t>{txt}</a:t></a:r>' if txt else ''
    x, y, w, h = box
    nowrap = '' if wrap else ' wrap="none"'
    xml = (f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{sid}" name="{ph_type} {sid}"/>'
           f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph{t}{i}/></p:nvPr></p:nvSpPr>'
           f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm></p:spPr>'
           f'<p:txBody><a:bodyPr{nowrap}/>{style_xml(st, latin, ph_type == "subTitle")}<a:p>{run}</a:p></p:txBody></p:sp>')
    tree.append(etree.fromstring(xml))


def add_rect(tree, x, y, w, h, fill):
    sid = next_id(tree)
    tree.append(etree.fromstring(
        f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{sid}" name="Decoration {sid}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{hx(fill)}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln></p:spPr></p:sp>'))


def add_picture(layout, path, x, y, w, h):
    _, rid = layout.part.get_or_add_image_part(path)
    tree = layout.shapes._spTree
    pic = CT_Picture.new_pic(next_id(tree), os.path.basename(path), '', rid, x, y, w, h)
    pic.find(qn('p:nvPicPr')).find(qn('p:nvPr')).set('userDrawn', '1')
    tree.append(pic)


def merged_style(spec_ph, observed):
    """observed value where present, else the spec placeholder's own textStyle."""
    ts = spec_ph.get('textStyle') or {}
    st = {'align': ts.get('align'), 'size': ts.get('size'), 'bold': ts.get('bold'), 'colour': ts.get('color'), 'latin': None}
    for k in st:
        if observed and observed.get(k) is not None:
            st[k] = observed[k]
    return st


def main():
    if len(sys.argv) != 4:
        print('usage: spec-to-master.py <template-spec.json> <assetsDir> <out.pptx>', file=sys.stderr)
        sys.exit(2)
    spec_path, assets, out = sys.argv[1:]
    with open(spec_path, encoding='utf-8') as fh:
        spec = json.load(fh)
    obs = spec.get('observed') or {}
    oc, of = obs.get('colours') or {}, obs.get('fonts') or {}
    styles = obs.get('layoutStyles') or {}

    prs = Presentation()
    old_w = prs.slide_width
    prs.slide_width, prs.slide_height = spec['slideSize']['cx'], spec['slideSize']['cy']
    f = prs.slide_width / old_w
    master = prs.slide_master
    scale_x(master._element, f)
    for l in master.slide_layouts:
        scale_x(l._element, f)

    # ---- theme colours
    pal = {k: hx(v) for k, v in spec['palette'].items()}
    if oc.get('body'):
        pal['dk1'] = hx(oc['body'])
    if oc.get('title'):
        pal['accent1'] = hx(oc['title'])
    if oc.get('accent'):
        pal['accent2'] = hx(oc['accent'])
    pal['lt1'] = 'FFFFFF'
    # lt2: the most common off-white decoration fill (light, near-neutral) if any, else palette lt2
    fills = Counter(hx(d['fill']) for l in spec['layouts'] for d in l.get('decorations', [])
                    if d.get('kind') == 'shape' and d.get('fill'))
    offwhite = [c for c, _ in fills.most_common()
                if min(int(c[i:i + 2], 16) for i in (0, 2, 4)) >= 0xE0 and c != 'FFFFFF']
    lt2_src = 'palette'
    if offwhite:
        pal['lt2'], lt2_src = offwhite[0], 'decoration fill'
    print('colours:', pal, '(lt2 from %s)' % lt2_src)

    theme_part = master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(theme_part.blob)
    ns = {'a': A}
    theme.find('.//a:clrScheme', ns).set('name', 'SK hynix')
    for k, v in pal.items():
        slot = theme.find(f'.//a:clrScheme/a:{k}', ns)
        if slot is None:
            continue
        for c in list(slot):
            slot.remove(c)
        etree.SubElement(slot, f'{{{A}}}srgbClr', val=v)
    major = (of.get('title') or {}).get('latin') or spec['fonts']['major']['latin']
    minor = (of.get('body') or {}).get('latin') or spec['fonts']['minor']['latin']
    fs = theme.find('.//a:fontScheme', ns)
    fs.set('name', 'SK hynix')
    for tag, face in (('majorFont', major), ('minorFont', minor)):
        el = fs.find(f'a:{tag}', ns)
        el.find('a:latin', ns).set('typeface', face)
        el.find('a:ea', ns).set('typeface', face)
        for extra in el.findall('a:font', ns):
            if extra.get('script') == 'Hang':
                extra.set('typeface', face)
    print('fonts: major', major, '/ minor', minor)
    theme_part._blob = etree.tostring(theme, xml_declaration=True, encoding='UTF-8', standalone=True)

    # ---- master: white background, default placeholders kept
    set_bg(master._element, '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>')

    # ---- layouts: reuse defaults with the same name, rename spare ones, drop the rest
    defaults = list(master.slide_layouts)
    by_name = {l.name: l for l in defaults}
    wanted = [l['name'] for l in spec['layouts']]
    spare = [l for l in defaults if l.name not in wanted]
    chosen = {}
    for n in wanted:
        lay = by_name.get(n) or spare.pop(0)
        lay._element.find(qn('p:cSld')).set('name', n)
        chosen[n] = lay
    keep_ids = {id(l) for l in chosen.values()}
    for l in defaults:
        if id(l) not in keep_ids:
            master.slide_layouts.remove(l)

    for ls in spec['layouts']:
        name = ls['name']
        lay = chosen[name]
        el = lay._element
        el.set('showMasterSp', '1' if ls.get('showMasterSp', True) else '0')
        tree = lay.shapes._spTree
        for c in list(tree):
            if c.tag not in (qn('p:nvGrpSpPr'), qn('p:grpSpPr')):
                tree.remove(c)
        bg = ls.get('background') or {}
        if bg.get('type') == 'image':
            _, rid = lay.part.get_or_add_image_part(os.path.join(assets, bg['value']))
            set_bg(el, f'<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="{rid}"/><a:srcRect/>'
                       f'<a:stretch><a:fillRect/></a:stretch></a:blipFill>')
        elif bg.get('type') == 'color' and bg.get('value'):
            set_bg(el, f'<a:solidFill><a:srgbClr val="{hx(bg["value"])}"/></a:solidFill>')
        for d in ls.get('decorations', []):
            box = (d['x'], d['y'], d['w'], d['h'])
            if d.get('asset'):
                add_picture(lay, os.path.join(assets, d['asset']), *box)
            elif d.get('fill'):
                add_rect(tree, *box, d['fill'])
        ob = styles.get(name) or {}
        used = set()
        phs = list(ls.get('placeholders', []))
        # an observed subtitle on a layout with no subtitle placeholder becomes one
        if ob.get('subtitle') and not any(p['type'] == 'subTitle' for p in phs):
            s = ob['subtitle']
            free = max([p['idx'] for p in phs] + [0]) + 1
            phs.append({'type': 'subTitle', 'idx': free, 'x': s['x'], 'y': s['y'], 'w': s['w'], 'h': s['h'], 'textStyle': {}, 'nowrap': True})
        for ph in phs:
            role = ROLE.get(ph['type'])
            o = ob.get(role) if role and role not in used else None
            if role:
                used.add(role)
            x, y, w, h = ph['x'], ph['y'], ph['w'], ph['h']
            if o and o.get('x') is not None:
                # observed position; keep the wider of the two widths so one-line titles do not wrap
                x, y, h = o['x'], o['y'], o['h']
                w = max(o['w'], min(ph['w'], prs.slide_width - x))
            st = merged_style(ph, o) if role else merged_style(ph, None)
            if role and not st.get('align'):
                st['align'] = 'l'
            latin = (major if role == 'title' else minor) if role else None
            add_placeholder(tree, ph['type'], ph['idx'], (x, y, w, h), st, latin, not ph.get('nowrap'))
        print(f'layout "{name}": {len(ls.get("decorations", []))} decorations, {len(phs)} placeholders')

    # layout order as in the spec
    lst = master._element.find(qn('p:sldLayoutIdLst'))
    ids = {master.part.related_part(e.get(qn('r:id'))).slide_layout.name: e for e in lst}
    for e in list(lst):
        lst.remove(e)
    for n in wanted:
        lst.append(ids[n])

    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    prs.save(out)
    print('wrote', out)


if __name__ == '__main__':
    main()
