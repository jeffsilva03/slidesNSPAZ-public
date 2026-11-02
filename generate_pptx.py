#!/usr/bin/env python3
"""
Gera PPTX inserindo slides com fundo preto nas seções corretas do template.
Entrada JSON: [{"section_key": "cantico_entrada", "slides": [{"text":"...", "bold": true}, ...]}, ...]
"""
import sys, json, os, re, zipfile, uuid
from lxml import etree

P_NS   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A_NS   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS  = 'http://schemas.openxmlformats.org/package/2006/content-types'
P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

# Mapeamento: section_key PHP → nome(s) de seção no template (case-insensitive, parcial)
SECTION_MAP = {
    'entrada_sao_miguel':  ['ENTRADA DA IMAGEM DE SÃO MIGUEL', 'ENTRADA DA IMAGEM DE SAO MIGUEL'],
    'cantico_entrada':      ['ENTRADA'],
    'aspercao':             ['ASPERÇÃO', 'ASPERSÃO', 'ASPERCAO', 'ASPERSAO'],
    'ato_penitencial':      ['ATO PENITENCIAL'],
    'gloria':               ['HINO DE LOUVOR', 'GLORIA'],
    'primeira_leitura':     ['LEITURAS'],
    'salmo_responsorial':   ['LEITURAS'],
    'segunda_leitura':      ['LEITURAS'],
    'aclamacao_evangelho':  ['LEITURAS'],
    'evangelho':            ['LEITURAS'],
    'preces_oracao_dizimista': ['PRECES - ORAÇÃO DO DIZIMISTA', 'PRECES'],
    'ofertorio':            ['OFERT'],
    'prefacio_eucaristico': ['PREFÁCIO EUCARÍSTICO', 'PREFACIO EUCARISTICO', 'PREFACIO EUCAR'],
    'santo':                ['SANTO'],
    'cordeiro':             ['CORDEIRO'],
    'comunhao':             ['COMUNHAO I', 'COMUNHAO'],
    'translado':            ['TRANSLADO'],
    'pos_comunhao':         ['POS COMUNH', 'PÓS COMUNH'],
    'cantico_final':        ['FINAL'],
    'outro':                [],  # appended at end
}

# Seções que devem ter o texto EDITADO nos slides existentes do template.
# Formato: (título_keyword, modo)
# modo 'body'    → substitui placeholder de corpo (pode ter múltiplos slides + overflow)
# modo 'textbox' → substitui primeira caixa de texto sem placeholder (shape NO_PH)
# modo 'title'   → substitui placeholder de título
EDIT_IN_PLACE = {
    'salmo_responsorial':  ('SALMO',    'body'),
    'aclamacao_evangelho': ('ACLAMA',   'body'),
    'primeira_leitura':    ('PRIMEIRA', 'body'),
    'segunda_leitura':     ('SEGUNDA',  'textbox'),
    'evangelho':           ('EVANGELHO','title'),
    'preces_oracao_dizimista': ('PRECES', 'body'),
}

# Seções de rúbrica: só substitui a referência entre parênteses, mantém o label
RUBRIC_KEYS = frozenset({'primeira_leitura', 'segunda_leitura', 'evangelho'})
PSALM_SECTION = 'salmo_responsorial'
READING_ORDER = ('primeira_leitura', PSALM_SECTION, 'segunda_leitura',
                 'aclamacao_evangelho', 'evangelho')

# Prefácio Eucarístico: primeiros N slides de texto são fixos (nunca tocamos);
# o último slide de texto tem ícones com hiperlink — ao substituir texto, preservamos as shapes.
PREFACIO_KEY         = 'prefacio_eucaristico'
PREFACIO_FIXED_COUNT = 2  # slides de texto iniciais que jamais são modificados

def px(tag): return f'{{{P_NS}}}{tag}'
def p14(tag): return f'{{{P14_NS}}}{tag}'
def rx(attr): return f'{{{R_NS}}}{attr}'

def normalize(s):
    import unicodedata
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s.upper()

def find_template_section(sections, section_key):
    """Retorna o elemento section do template que corresponde ao section_key."""
    candidates = SECTION_MAP.get(section_key, [])
    for candidate in candidates:
        cand_norm = normalize(candidate)
        for sec in sections:
            sec_norm = normalize(sec.get('name', ''))
            if cand_norm in sec_norm or (len(sec_norm) >= 4 and sec_norm in cand_norm):
                return sec
    return None

def get_sections(pres_root):
    """Retorna lista de dicts {name, el, sld_ids} das seções do template."""
    sections = []
    for sec_el in pres_root.iter(p14('section')):
        ids = [int(el.get('id')) for el in sec_el.iter(p14('sldId'))]
        sections.append({'name': sec_el.get('name', ''), 'el': sec_el, 'sld_ids': ids})
    return sections

def get_global_slide_order(pres_root):
    """Retorna lista de IDs na ordem global (p:sldIdLst)."""
    sldIdLst = pres_root.find(px('sldIdLst'))
    if sldIdLst is None:
        return []
    return [int(el.get('id')) for el in sldIdLst]

def build_black_slide_xml(text, bold, cx, cy, layout_rel_target):
    """Gera XML completo de um slide com fundo preto."""
    b = '1' if bold else '0'
    paras = ''
    for line in text.split('\n'):
        e = (line.replace('&','&amp;').replace('<','&lt;')
                 .replace('>','&gt;').replace('"','&quot;'))
        paras += (
            f'<a:p xmlns:a="{A_NS}">'
            f'<a:pPr algn="ctr"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/></a:pPr>'
            f'<a:r><a:rPr lang="pt-BR" sz="7200" b="{b}" dirty="0">'
            f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            f'<a:latin typeface="Arial Black"/></a:rPr>'
            f'<a:t>{e}</a:t></a:r></a:p>'
        )
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p="{P_NS}">'
        f'<p:cSld>'
        f'<p:bg><p:bgPr>'
        f'<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
        f'<a:effectLst/>'
        f'</p:bgPr></p:bg>'
        f'<p:spTree>'
        f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f'<p:grpSpPr><a:xfrm>'
        f'<a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>'
        f'</a:xfrm></p:grpSpPr>'
        f'<p:sp>'
        f'<p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr wrap="square" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr>'
        f'<a:lstStyle/>{paras}'
        f'</p:txBody></p:sp>'
        f'</p:spTree></p:cSld>'
        f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        f'</p:sld>'
    )

def build_slide_rels(layout_target):
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_NS}">'
        f'<Relationship Id="rId1"'
        f' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"'
        f' Target="{layout_target}"/>'
        f'</Relationships>'
    )

def escape_xml(text):
    return (str(text or '').replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))

def get_image_size(path):
    try:
        with open(path, 'rb') as f:
            data = f.read(32)
            if data.startswith(b'\x89PNG\r\n\x1a\n') and len(data) >= 24:
                return int.from_bytes(data[16:20], 'big'), int.from_bytes(data[20:24], 'big')
            f.seek(0)
            if f.read(2) != b'\xff\xd8':
                return None
            while True:
                marker_start = f.read(1)
                if not marker_start:
                    return None
                if marker_start != b'\xff':
                    continue
                marker = f.read(1)
                while marker == b'\xff':
                    marker = f.read(1)
                if marker in (b'\xd8', b'\xd9'):
                    continue
                size_data = f.read(2)
                if len(size_data) != 2:
                    return None
                size = int.from_bytes(size_data, 'big')
                if marker in [bytes([m]) for m in range(0xC0, 0xCF) if m not in (0xC4, 0xC8, 0xCC)]:
                    f.read(1)
                    h = int.from_bytes(f.read(2), 'big')
                    w = int.from_bytes(f.read(2), 'big')
                    return w, h
                f.seek(size - 2, os.SEEK_CUR)
    except Exception:
        return None

def build_src_rect_for_cover(image_path, cx, cy):
    size = get_image_size(image_path)
    if not size:
        return ''
    iw, ih = size
    if iw <= 0 or ih <= 0:
        return ''
    slide_ratio = cx / cy
    image_ratio = iw / ih
    attrs = {}
    if image_ratio > slide_ratio:
        needed_w = ih * slide_ratio
        crop = max(0, min(49999, int(((iw - needed_w) / iw) * 50000)))
        attrs['l'] = attrs['r'] = str(crop)
    elif image_ratio < slide_ratio:
        needed_h = iw / slide_ratio
        crop = max(0, min(49999, int(((ih - needed_h) / ih) * 50000)))
        attrs['t'] = attrs['b'] = str(crop)
    if not attrs:
        return ''
    return '<a:srcRect ' + ' '.join(f'{k}="{v}"' for k, v in attrs.items()) + '/>'

def build_custom_cover_slide_xml(title, image_rid, cx, cy, src_rect=''):
    safe_title = escape_xml(title).upper()
    font_size = 5200 if len(safe_title) <= 22 else 4200 if len(safe_title) <= 34 else 3400
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p="{P_NS}">'
        f'<p:cSld><p:spTree>'
        f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'<p:pic><p:nvPicPr><p:cNvPr id="2" name="Capa da seção"/>'
        f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        f'<p:blipFill><a:blip r:embed="{image_rid}"/>{src_rect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        f'<p:sp><p:nvSpPr><p:cNvPr id="3" name="Sombra da capa"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="000000"><a:alpha val="45000"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>'
        f'<p:sp><p:nvSpPr><p:cNvPr id="4" name="Título da seção"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="571500" y="2571750"/><a:ext cx="8001000" cy="1714500"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square" anchor="ctr"><a:normAutofit fontScale="70000" lnSpcReduction="20000"/></a:bodyPr><a:lstStyle/>'
        f'<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr lang="pt-BR" sz="{font_size}" b="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="F2CC7E"/></a:solidFill><a:latin typeface="Arial Black"/></a:rPr><a:t>{safe_title}</a:t></a:r></a:p>'
        f'</p:txBody></p:sp>'
        f'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    )

def build_custom_cover_rels(layout_target, media_file):
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_NS}">'
        f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="{layout_target}"/>'
        f'<Relationship Id="rId2" Type="{REL_IMAGE}" Target="../media/{media_file}"/>'
        f'</Relationships>'
    )

def get_slide1_layout(zf):
    """Pega o layout target do slide1 para usar nos novos slides."""
    if 'ppt/slides/_rels/slide1.xml.rels' not in zf.namelist():
        return '../slideLayouts/slideLayout1.xml'
    rels_root = etree.fromstring(zf.read('ppt/slides/_rels/slide1.xml.rels'))
    for r in rels_root:
        if r.get('Type','').endswith('/slideLayout'):
            return r.get('Target')
    return '../slideLayouts/slideLayout1.xml'

def splitIntoSlides_py(text, chars_per_line=14, max_lines=4,
                       short_slide_lines=2, safe_merged_max_lines=5,
                       hard_max_lines=5):
    """Versão Python do splitIntoSlides do PHP."""
    def word_wrap(line):
        words = line.split()
        lines, cur = [], ''
        for w in words:
            test = (cur + ' ' + w).strip()
            if len(test) > chars_per_line and cur:
                lines.append(cur)
                cur = w
            else:
                cur = test
        if cur:
            lines.append(cur)
        return lines

    def line_count(slide):
        return len([line for line in slide.split('\n') if line.strip()])

    def merge_short(slides):
        merged = []
        i = 0
        while i < len(slides):
            current = slides[i]
            if i + 1 < len(slides):
                current_count = line_count(current)
                next_count = line_count(slides[i + 1])
                if (
                    0 < current_count <= short_slide_lines
                    and current_count + next_count <= safe_merged_max_lines
                ):
                    merged.append(current + '\n' + slides[i + 1])
                    i += 2
                    continue
            merged.append(current)
            i += 1
        return merged

    def split_phrase_units(line):
        line = line.strip()
        if not line:
            return []
        units = [part.strip() for part in re.split(r'(?<=[.!?;:])\s+', line)
                 if part.strip()]
        if not units:
            units = [line]
        expanded = []
        for unit in units:
            wrapped_unit = word_wrap(unit)
            if len(wrapped_unit) > hard_max_lines:
                comma_units = [part.strip() for part in re.split(r'(?<=,)\s+', unit)
                               if part.strip()]
                if len(comma_units) > 1:
                    expanded.extend(comma_units)
                    continue
            expanded.append(unit)
        return expanded

    def flush_slide(slides, cur_slide):
        if cur_slide:
            slides.append('\n'.join(cur_slide))
            cur_slide.clear()

    def append_wrapped_unit(para_slides, cur_slide, wrapped_unit):
        while len(wrapped_unit) > hard_max_lines:
            flush_slide(para_slides, cur_slide)
            para_slides.append('\n'.join(wrapped_unit[:hard_max_lines]))
            wrapped_unit = wrapped_unit[hard_max_lines:]
        if not wrapped_unit:
            return
        if cur_slide and len(cur_slide) + len(wrapped_unit) > hard_max_lines:
            flush_slide(para_slides, cur_slide)
        cur_slide.extend(wrapped_unit)

    paragraphs = []
    cur = []
    for line in text.split('\n'):
        t = line.strip()
        if not t:
            if cur:
                paragraphs.append(cur)
                cur = []
        else:
            cur.append(t)
    if cur:
        paragraphs.append(cur)

    slides = []
    for para in paragraphs:
        para_slides = []
        cur_s = []
        for line in para:
            for unit in split_phrase_units(line):
                append_wrapped_unit(para_slides, cur_s, word_wrap(unit))
        flush_slide(para_slides, cur_s)
        slides.extend(merge_short(para_slides))
    return slides if slides else [text]

def slide_has_body(zf, fname):
    """Retorna True se o slide tem um placeholder de corpo editável."""
    try:
        root = etree.fromstring(zf.read(f'ppt/slides/{fname}'))
        for sp in root.iter(f'{{{P_NS}}}sp'):
            ph = sp.find(f'.//{{{P_NS}}}ph')
            if ph is None:
                continue
            if ph.get('type', 'body') not in ('title', 'ctrTitle', 'dt', 'ftr', 'sldNum'):
                return True
    except Exception:
        pass
    return False

def get_shape_full_text(zf, fname, mode):
    """Retorna o texto completo do shape alvo (body/textbox/title) de um slide."""
    try:
        root = etree.fromstring(zf.read(f'ppt/slides/{fname}'))
        for sp in root.iter(f'{{{P_NS}}}sp'):
            ph = sp.find(f'.//{{{P_NS}}}ph')
            ph_type = ph.get('type', 'body') if ph is not None else None
            if mode == 'body':
                if ph is not None and ph_type not in ('title', 'ctrTitle', 'dt', 'ftr', 'sldNum'):
                    return ''.join(t.text or '' for t in sp.iter(f'{{{A_NS}}}t'))
            elif mode == 'textbox':
                if ph is None:
                    text = ''.join(t.text or '' for t in sp.iter(f'{{{A_NS}}}t'))
                    if text.strip():
                        return text
            elif mode == 'title':
                if ph is not None and ph_type in ('title', 'ctrTitle'):
                    return ''.join(t.text or '' for t in sp.iter(f'{{{A_NS}}}t'))
    except Exception:
        pass
    return ''

def format_rubric_text(current_full, ref):
    """Mantém o label (prefixo) e substitui apenas a referência, sempre com quebra de linha."""
    ref = ref.strip().strip('()')
    if '(' in current_full:
        prefix = current_full[:current_full.index('(')].rstrip()
    else:
        prefix = current_full.rstrip()
    return prefix + '\n(' + ref + ')'

def get_slide_title_text(zf, fname):
    """Retorna o texto do placeholder title de um slide."""
    try:
        root = etree.fromstring(zf.read(f'ppt/slides/{fname}'))
        for sp in root.iter(f'{{{P_NS}}}sp'):
            ph = sp.find(f'.//{{{P_NS}}}ph')
            if ph is not None and ph.get('type') in ('title', 'ctrTitle'):
                texts = [t.text for t in sp.iter(f'{{{A_NS}}}t') if t.text]
                return ''.join(texts).upper()
    except Exception:
        pass
    return ''

def ensure_no_bullet(pPr):
    """Remove bullet explícito/herdado e insere buNone na posição válida do DrawingML."""
    bullet_tags = {
        f'{{{A_NS}}}buClr', f'{{{A_NS}}}buClrTx',
        f'{{{A_NS}}}buSzTx', f'{{{A_NS}}}buSzPct', f'{{{A_NS}}}buSzPts',
        f'{{{A_NS}}}buFont', f'{{{A_NS}}}buFontTx',
        f'{{{A_NS}}}buNone', f'{{{A_NS}}}buAutoNum', f'{{{A_NS}}}buChar',
        f'{{{A_NS}}}buBlip',
    }
    for child in list(pPr):
        if child.tag in bullet_tags:
            pPr.remove(child)

    bu_none = etree.Element(f'{{{A_NS}}}buNone')
    insert_before = {
        f'{{{A_NS}}}tabLst',
        f'{{{A_NS}}}defRPr',
        f'{{{A_NS}}}extLst',
    }
    insert_at = len(pPr)
    for i, child in enumerate(list(pPr)):
        if child.tag in insert_before:
            insert_at = i
            break
    pPr.insert(insert_at, bu_none)
    return pPr

def replace_body_text(slide_xml_data, new_text):
    """
    Substitui o texto do ph=body de um slide mantendo toda a formatação original.
    """
    import copy as _copy

    root = etree.fromstring(slide_xml_data)

    target_sp = None
    for sp in root.iter(f'{{{P_NS}}}sp'):
        ph = sp.find(f'.//{{{P_NS}}}ph')
        if ph is None:
            continue
        ph_type = ph.get('type', 'body')
        if ph_type in ('title', 'ctrTitle', 'dt', 'ftr', 'sldNum'):
            continue
        target_sp = sp
        break

    if target_sp is None:
        return slide_xml_data

    txBody = target_sp.find(f'{{{P_NS}}}txBody')
    if txBody is None:
        return slide_xml_data

    # Habilita auto-fit para que mais linhas não sejam cortadas
    bodyPr = txBody.find(f'{{{A_NS}}}bodyPr')
    if bodyPr is not None:
        # Remove noAutofit, mantém outras configurações
        no_auto = bodyPr.find(f'{{{A_NS}}}noAutofit')
        if no_auto is not None:
            bodyPr.remove(no_auto)
        # Garante normAutofit
        if bodyPr.find(f'{{{A_NS}}}normAutofit') is None and bodyPr.find(f'{{{A_NS}}}spAutoFit') is None:
            etree.SubElement(bodyPr, f'{{{A_NS}}}normAutofit')

    # Pega o primeiro parágrafo como template de formatação (pPr + rPr)
    template_pPr = None
    template_rPr = None
    for para in txBody.findall(f'{{{A_NS}}}p'):
        pPr = para.find(f'{{{A_NS}}}pPr')
        if pPr is not None and template_pPr is None:
            template_pPr = _copy.deepcopy(pPr)
            ensure_no_bullet(template_pPr)
        for run in para.findall(f'{{{A_NS}}}r'):
            rPr = run.find(f'{{{A_NS}}}rPr')
            if rPr is not None and template_rPr is None:
                template_rPr = _copy.deepcopy(rPr)
        if template_pPr is not None and template_rPr is not None:
            break

    # Fallback pPr se o slide não tiver nenhum
    if template_pPr is None:
        template_pPr = etree.Element(f'{{{A_NS}}}pPr')
        template_pPr.set('algn', 'ctr')
        ensure_no_bullet(template_pPr)

    # Remove todos os parágrafos existentes
    for p in txBody.findall(f'{{{A_NS}}}p'):
        txBody.remove(p)

    # Cria novos parágrafos com o novo texto
    for line in new_text.split('\n'):
        p_el = etree.SubElement(txBody, f'{{{A_NS}}}p')
        p_el.append(_copy.deepcopy(template_pPr))
        r_el = etree.SubElement(p_el, f'{{{A_NS}}}r')
        if template_rPr is not None:
            r_el.append(_copy.deepcopy(template_rPr))
        t_el = etree.SubElement(r_el, f'{{{A_NS}}}t')
        t_el.text = line

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def _replace_shape_text(root, target_sp, new_text):
    """Helper: substitui o texto de um txBody preservando formatação do primeiro run."""
    import copy as _copy
    txBody = target_sp.find(f'{{{P_NS}}}txBody')
    if txBody is None:
        return None
    template_pPr, template_rPr = None, None
    for para in txBody.findall(f'{{{A_NS}}}p'):
        if template_pPr is None:
            pPr = para.find(f'{{{A_NS}}}pPr')
            if pPr is not None:
                template_pPr = _copy.deepcopy(pPr)
        for run in para.findall(f'{{{A_NS}}}r'):
            if template_rPr is None:
                rPr = run.find(f'{{{A_NS}}}rPr')
                if rPr is not None:
                    template_rPr = _copy.deepcopy(rPr)
        if template_pPr is not None and template_rPr is not None:
            break
    if template_pPr is None:
        template_pPr = etree.Element(f'{{{A_NS}}}pPr')
    for p in txBody.findall(f'{{{A_NS}}}p'):
        txBody.remove(p)
    for line in new_text.split('\n'):
        p_el = etree.SubElement(txBody, f'{{{A_NS}}}p')
        p_el.append(_copy.deepcopy(template_pPr))
        r_el = etree.SubElement(p_el, f'{{{A_NS}}}r')
        if template_rPr is not None:
            r_el.append(_copy.deepcopy(template_rPr))
        t_el = etree.SubElement(r_el, f'{{{A_NS}}}t')
        t_el.text = line
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def replace_textbox_text(slide_xml_data, new_text):
    """Substitui o texto da primeira caixa de texto (shape sem placeholder) do slide."""
    root = etree.fromstring(slide_xml_data)
    for sp in root.iter(f'{{{P_NS}}}sp'):
        if sp.find(f'.//{{{P_NS}}}ph') is not None:
            continue  # pula placeholders
        txBody = sp.find(f'{{{P_NS}}}txBody')
        if txBody is None:
            continue
        texts = [t.text for t in txBody.iter(f'{{{A_NS}}}t') if t.text]
        if any(t.strip() for t in texts):
            result = _replace_shape_text(root, sp, new_text)
            return result if result is not None else slide_xml_data
    return slide_xml_data

def replace_title_text(slide_xml_data, new_text):
    """Substitui o texto do placeholder de título do slide."""
    root = etree.fromstring(slide_xml_data)
    for sp in root.iter(f'{{{P_NS}}}sp'):
        ph = sp.find(f'.//{{{P_NS}}}ph')
        if ph is not None and ph.get('type') in ('title', 'ctrTitle'):
            result = _replace_shape_text(root, sp, new_text)
            return result if result is not None else slide_xml_data
    return slide_xml_data

def replace_main_text_preserve_shapes(slide_xml_data, new_text):
    """
    Substitui o texto do shape principal do slide (body placeholder ou primeiro textbox com texto)
    sem tocar em nenhuma outra shape — em especial shapes com hiperlinks (ícones).
    Usado no último slide de texto do Prefácio Eucarístico.
    """
    # Tenta primeiro via placeholder de corpo (igual a replace_body_text)
    result = replace_body_text(slide_xml_data, new_text)
    if result is not slide_xml_data:
        return result
    # Fallback: substitui o texto do primeiro textbox não-placeholder que tenha conteúdo
    root = etree.fromstring(slide_xml_data)
    for sp in root.iter(f'{{{P_NS}}}sp'):
        ph = sp.find(f'.//{{{P_NS}}}ph')
        if ph is not None:
            continue  # pula placeholders (título, rodapé etc.)
        # Verifica se tem hiperlink — se sim, pula (é um ícone)
        if sp.find(f'.//{{{A_NS}}}hlinkClick') is not None:
            continue
        txBody = sp.find(f'{{{P_NS}}}txBody')
        if txBody is None:
            continue
        bodyPr = txBody.find(f'{{{A_NS}}}bodyPr')
        if bodyPr is not None:
            no_auto = bodyPr.find(f'{{{A_NS}}}noAutofit')
            if no_auto is not None:
                bodyPr.remove(no_auto)
            if bodyPr.find(f'{{{A_NS}}}normAutofit') is None and bodyPr.find(f'{{{A_NS}}}spAutoFit') is None:
                etree.SubElement(bodyPr, f'{{{A_NS}}}normAutofit')
        existing = ''.join(t.text or '' for t in txBody.iter(f'{{{A_NS}}}t')).strip()
        if not existing:
            continue
        # Substitui o conteúdo deste textbox preservando formatação do primeiro parágrafo
        paras = txBody.findall(f'{{{A_NS}}}p')
        template_para = paras[0] if paras else None
        for p in list(paras):
            txBody.remove(p)
        for line in new_text.split('\n'):
            if template_para is not None:
                import copy as _copy
                new_p = _copy.deepcopy(template_para)
                for r in new_p.findall(f'{{{A_NS}}}r'):
                    new_p.remove(r)
                template_r = template_para.find(f'{{{A_NS}}}r')
                if template_r is not None:
                    new_r = _copy.deepcopy(template_r)
                    t_el = new_r.find(f'{{{A_NS}}}t')
                    if t_el is None:
                        t_el = etree.SubElement(new_r, f'{{{A_NS}}}t')
                    t_el.text = line
                    new_p.append(new_r)
                txBody.append(new_p)
            else:
                p_el = etree.SubElement(txBody, f'{{{A_NS}}}p')
                r_el = etree.SubElement(p_el, f'{{{A_NS}}}r')
                t_el = etree.SubElement(r_el, f'{{{A_NS}}}t')
                t_el.text = line
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return slide_xml_data

def slide_xml_text_content(slide_xml_data):
    try:
        root = etree.fromstring(slide_xml_data)
        return ' '.join(t.text or '' for t in root.iter(f'{{{A_NS}}}t')).strip()
    except Exception:
        return ''

def slide_text_content(zf, fname):
    try:
        return slide_xml_text_content(zf.read(f'ppt/slides/{fname}'))
    except Exception:
        return ''

def comparable_text(text):
    return re.sub(r'\s+', ' ', str(text or '')).strip()

def slide_has_hyperlinks(zf, fname):
    try:
        root = etree.fromstring(zf.read(f'ppt/slides/{fname}'))
        return root.find(f'.//{{{A_NS}}}hlinkClick') is not None
    except Exception:
        return False

def find_prefacio_final_slide(zf, tpl_sec, id_to_idx, id_to_rid, rid_to_file):
    """
    Encontra o slide de texto do Prefácio que carrega os hyperlinks.
    O modelo também pode ter um slide final vazio com link, então preferimos
    o último slide que tenha hyperlink e texto editável.
    """
    if not tpl_sec or not tpl_sec.get('sld_ids'):
        return None

    sorted_sids = sorted(tpl_sec['sld_ids'], key=lambda sid: id_to_idx.get(sid, 99999))

    def fname_for_sid(sid):
        rid = id_to_rid.get(sid)
        return rid_to_file.get(rid) if rid else None

    for sid in reversed(sorted_sids):
        fname = fname_for_sid(sid)
        if fname and slide_has_hyperlinks(zf, fname) and slide_text_content(zf, fname):
            return fname

    for sid in reversed(sorted_sids):
        fname = fname_for_sid(sid)
        if fname and slide_has_hyperlinks(zf, fname):
            return fname

    for sid in reversed(sorted_sids):
        fname = fname_for_sid(sid)
        if fname and slide_text_content(zf, fname):
            return fname

    return None

def copy_element(el):
    import copy
    return copy.deepcopy(el)

def get_slide_size(zf):
    root = etree.fromstring(zf.read('ppt/presentation.xml'))
    sldSz = root.find(px('sldSz'))
    if sldSz is not None:
        return int(sldSz.get('cx', 9144000)), int(sldSz.get('cy', 6858000))
    return 9144000, 6858000

def ensure_default_content_type(ct_root, ext):
    ext = ext.lower().lstrip('.')
    content_types = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
    }
    if ext not in content_types:
        return
    for el in ct_root:
        if el.tag == f'{{{CT_NS}}}Default' and el.get('Extension', '').lower() == ext:
            return
    default_el = etree.Element(f'{{{CT_NS}}}Default')
    default_el.set('Extension', ext)
    default_el.set('ContentType', content_types[ext])
    ct_root.insert(0, default_el)

def next_available_media_name(all_names, new_media_files, base, ext):
    base = re.sub(r'[^a-zA-Z0-9_-]+', '_', base).strip('_') or 'custom_section'
    used = set(all_names) | set(new_media_files)
    idx = 1
    while True:
        name = f'ppt/media/{base}_{idx}.{ext}'
        if name not in used:
            return name.split('/')[-1]
        idx += 1

def build_menu_link_shape(shape_id, rid, title, x, y, w, h, font_size):
    text = escape_xml(title).upper()
    return (
        f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">'
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="Seção extra {shape_id}">'
        f'<a:hlinkClick r:id="{rid}" action="ppaction://hlinksldjump"/></p:cNvPr>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:solidFill>'
        f'<a:ln w="9525"><a:solidFill><a:srgbClr val="F2CC7E"><a:alpha val="65000"/></a:srgbClr></a:solidFill></a:ln>'
        f'</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="68580" rIns="68580" tIns="22860" bIns="22860">'
        f'<a:normAutofit fontScale="65000" lnSpcReduction="20000"/></a:bodyPr><a:lstStyle/>'
        f'<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr lang="pt-BR" sz="{font_size}" b="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="F2CC7E"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>{text}</a:t></a:r></a:p>'
        f'</p:txBody></p:sp>'
    )

def add_menu_links_to_slide1(slide_xml_data, slide_rels_data, links, cx, cy):
    if not links:
        return slide_xml_data, slide_rels_data
    slide_root = etree.fromstring(slide_xml_data)
    rels_root = etree.fromstring(slide_rels_data)
    sp_tree = slide_root.find(f'.//{{{P_NS}}}spTree')
    if sp_tree is None:
        return slide_xml_data, slide_rels_data

    max_shape_id = 1
    for cNvPr in slide_root.iter(f'{{{P_NS}}}cNvPr'):
        try:
            max_shape_id = max(max_shape_id, int(cNvPr.get('id', '0')))
        except ValueError:
            pass
    max_rid = 0
    for rel in rels_root:
        m = re.search(r'\d+', rel.get('Id', ''))
        if m:
            max_rid = max(max_rid, int(m.group()))

    n = len(links)
    cols = 4 if n <= 16 else 5
    cols = min(cols, max(1, n))
    rows = (n + cols - 1) // cols
    margin_x = 457200
    gap_x = 137160
    available_w = cx - (margin_x * 2) - (gap_x * (cols - 1))
    w = max(914400, available_w // cols)
    available_h = 1257300 if rows <= 4 else 4572000
    top = cy - available_h - 228600 if rows <= 4 else 2057400
    gap_y = 91440
    h = max(152400, min(320040, (available_h - gap_y * (rows - 1)) // rows))
    font_size = max(900, min(1600, int(1700 - max(0, rows - 3) * 140)))

    for idx, link in enumerate(links):
        max_rid += 1
        max_shape_id += 1
        rid = f'rId{max_rid}'
        rel = etree.SubElement(rels_root, f'{{{PKG_NS}}}Relationship')
        rel.set('Id', rid)
        rel.set('Type', REL_SLIDE)
        rel.set('Target', link['target'])

        col = idx % cols
        row = idx // cols
        x = margin_x + col * (w + gap_x)
        y = top + row * (h + gap_y)
        shape_xml = build_menu_link_shape(max_shape_id, rid, link['title'], x, y, w, h, font_size)
        sp_tree.append(etree.fromstring(shape_xml))

    return (
        etree.tostring(slide_root, xml_declaration=True, encoding='UTF-8', standalone=True),
        etree.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True),
    )

def create_custom_section_element(title, sids):
    sec_el = etree.Element(p14('section'))
    sec_el.set('name', title or 'Seção extra')
    sec_el.set('id', '{' + str(uuid.uuid4()).upper() + '}')
    sec_sldIdLst = etree.SubElement(sec_el, p14('sldIdLst'))
    for sid in sids:
        el = etree.SubElement(sec_sldIdLst, p14('sldId'))
        el.set('id', str(sid))
    return sec_el

def sort_sections_by_slide_order(pres_root):
    section_lst = None
    for el in pres_root.iter(p14('sectionLst')):
        section_lst = el
        break
    if section_lst is None:
        return
    sldIdLst = pres_root.find(px('sldIdLst'))
    if sldIdLst is None:
        return
    order = {int(el.get('id')): i for i, el in enumerate(sldIdLst)}
    sections = list(section_lst.findall(p14('section')))
    def section_pos(sec):
        ids = [int(el.get('id')) for el in sec.iter(p14('sldId')) if el.get('id')]
        positions = [order.get(sid, 10**9) for sid in ids]
        return min(positions) if positions else 10**9
    for sec in sections:
        ids = sec.find(p14('sldIdLst'))
        if ids is not None:
            ids[:] = sorted(ids, key=lambda el: order.get(int(el.get('id')), 10**9))
        section_lst.remove(sec)
    for sec in sorted(sections, key=section_pos):
        section_lst.append(sec)

def find_anchor_insert_after(zf, sections, anchor_key, id_to_idx, id_to_rid, rid_to_file, sldIdLst):
    tpl_sec = find_template_section(sections, anchor_key)
    if not tpl_sec or not tpl_sec.get('sld_ids'):
        return len(list(sldIdLst)) - 1

    if anchor_key in EDIT_IN_PLACE:
        title_keyword, replace_mode = EDIT_IN_PLACE[anchor_key]
        title_keyword = title_keyword.upper()
        sorted_sids = sorted(tpl_sec['sld_ids'], key=lambda sid: id_to_idx.get(sid, 99999))

        if replace_mode == 'body':
            found_title = False
            last_idx = -1
            for sid in sorted_sids:
                rid = id_to_rid.get(sid)
                fname = rid_to_file.get(rid) if rid else None
                if not fname:
                    continue
                title = get_slide_title_text(zf, fname)
                if not found_title:
                    if title_keyword in title:
                        found_title = True
                        last_idx = id_to_idx.get(sid, last_idx)
                else:
                    if title and title_keyword not in title:
                        break
                    last_idx = id_to_idx.get(sid, last_idx)
            if last_idx >= 0:
                return last_idx

        for sid in sorted_sids:
            rid = id_to_rid.get(sid)
            fname = rid_to_file.get(rid) if rid else None
            if not fname:
                continue
            if title_keyword in get_slide_title_text(zf, fname):
                return id_to_idx.get(sid, len(list(sldIdLst)) - 1)

    return max((id_to_idx.get(sid, -1) for sid in tpl_sec['sld_ids']), default=len(list(sldIdLst)) - 1)

def find_slide_by_title_keyword(zf, tpl_sec, title_keyword, id_to_idx, id_to_rid, rid_to_file):
    if not tpl_sec:
        return None, -1
    for sid in sorted(tpl_sec.get('sld_ids', []), key=lambda s: id_to_idx.get(s, 99999)):
        rid = id_to_rid.get(sid)
        fname = rid_to_file.get(rid) if rid else None
        if fname and title_keyword.upper() in get_slide_title_text(zf, fname):
            return fname, id_to_idx.get(sid, -1)
    return None, -1

def find_reading_insert_after(zf, sections, section_key, id_to_idx, id_to_rid,
                              rid_to_file, sldIdLst):
    """Posiciona conteúdo sem campo correspondente entre as leituras do molde."""
    tpl_sec = find_template_section(sections, section_key)
    position = READING_ORDER.index(section_key)
    for key in reversed(READING_ORDER[:position + 1]):
        fname, _ = find_slide_by_title_keyword(
            zf, tpl_sec, EDIT_IN_PLACE[key][0], id_to_idx, id_to_rid, rid_to_file)
        if fname:
            return find_anchor_insert_after(
                zf, sections, key, id_to_idx, id_to_rid, rid_to_file, sldIdLst)
    for key in READING_ORDER[position + 1:]:
        fname, index = find_slide_by_title_keyword(
            zf, tpl_sec, EDIT_IN_PLACE[key][0], id_to_idx, id_to_rid, rid_to_file)
        if fname:
            return index - 1
    return min((id_to_idx[sid] for sid in tpl_sec['sld_ids'] if sid in id_to_idx),
               default=len(list(sldIdLst)) - 1)


def complete_group_slides(group):
    """Expande leituras e salmo quando não há campos do molde para editá-los."""
    slides = list(group.get('slides') or [])
    slides.extend(group.get('reading_slides') or [])
    for stanza in group.get('psalm_stanzas') or []:
        slides.extend(stanza)
        # O refrão abre o salmo e se repete depois de cada estrofe.
        slides.extend(group.get('slides') or [])
    return slides


def generate_with_template(template_path, output_path, section_groups):
    with zipfile.ZipFile(template_path, 'r') as src:
        all_names = src.namelist()
        cx, cy = get_slide_size(src)
        layout_target = get_slide1_layout(src)

        pres_xml_data  = src.read('ppt/presentation.xml')
        pres_rels_data = src.read('ppt/_rels/presentation.xml.rels')
        ct_data        = src.read('[Content_Types].xml')

        pres_root      = etree.fromstring(pres_xml_data)
        pres_rels_root = etree.fromstring(pres_rels_data)
        ct_root        = etree.fromstring(ct_data)

        sections      = get_sections(pres_root)
        global_ids    = get_global_slide_order(pres_root)
        sldIdLst      = pres_root.find(px('sldIdLst'))

        # Maior rId e sldId existentes
        max_rid = 0
        for r in pres_rels_root:
            m = re.search(r'\d+', r.get('Id',''))
            if m: max_rid = max(max_rid, int(m.group()))

        max_sld_id = max(global_ids) if global_ids else 255

        # Mapeia slide_id → índice em sldIdLst
        id_to_idx = {int(el.get('id')): i for i, el in enumerate(sldIdLst)}

        # Remove notesSlides e notesMasters do Content_Types
        for el in list(ct_root):
            ct = el.get('ContentType', '')
            if 'notesSlide' in ct or 'notesMaster' in ct:
                ct_root.remove(el)

        # Remove notesMaster da presentation.xml
        nm = pres_root.find(px('notesMasterIdLst'))
        if nm is not None:
            pres_root.remove(nm)

        # Remove apenas notesMaster. Metadados de importação podem ser exigidos
        # pelo PowerPoint para abrir modelos que passaram pelo Google Slides.
        for r in list(pres_rels_root):
            t = r.get('Type','')
            if t.endswith('/notesMaster'):
                pres_rels_root.remove(r)

        # ---- Build map: slide_id → filename ----
        pres_rels_xml  = etree.fromstring(pres_rels_data)
        rid_to_file = {}
        for r in pres_rels_xml:
            if r.get('Type','').endswith('/slide'):
                rid_to_file[r.get('Id')] = r.get('Target','').split('slides/')[-1]
        id_to_rid = {}
        for el in sldIdLst:
            # tenta com namespace e sem (compatibilidade)
            rid = el.get(f'{{{R_NS}}}id') or el.get('r:id')
            if rid is None:
                for attr, val in el.attrib.items():
                    if attr.endswith('}id') and val.startswith('rId'):
                        rid = val
                        break
            id_to_rid[int(el.get('id'))] = rid

        # ---- Edit-in-place para seções com estilização própria ----
        edit_in_place_files = {}  # fname → novo conteúdo
        groups_to_insert = []
        for group in section_groups:
            sec_key = group.get('section_key', 'outro')
            if sec_key == PREFACIO_KEY:
                tpl_sec = find_template_section(sections, sec_key)
                prefacio_group = dict(group)
                prefacio_slides = list(prefacio_group.get('slides', []))
                final_text = str(prefacio_group.get('prefacio_final_text') or '').strip()

                # Compatibilidade com JSON antigo: se não veio campo dedicado,
                # a última lâmina do Prefácio vira o slide hiperlinkado do modelo.
                if not final_text and prefacio_slides:
                    final_slide = prefacio_slides.pop()
                    final_text = str(final_slide.get('text', '')).strip()

                final_applied = False
                if final_text and tpl_sec:
                    final_fname = find_prefacio_final_slide(src, tpl_sec, id_to_idx, id_to_rid, rid_to_file)
                    if final_fname:
                        data = src.read(f'ppt/slides/{final_fname}')
                        updated = replace_main_text_preserve_shapes(data, final_text)
                        if comparable_text(final_text) in comparable_text(slide_xml_text_content(updated)):
                            edit_in_place_files[final_fname] = updated
                            final_applied = True

                if final_text and not final_applied:
                    prefacio_slides.append({'text': final_text, 'bold': False})

                prefacio_group['slides'] = prefacio_slides
                if prefacio_group.get('slides'):
                    groups_to_insert.append(prefacio_group)
                continue

            if sec_key not in EDIT_IN_PLACE:
                groups_to_insert.append(group)
                continue

            title_keyword, replace_mode = EDIT_IN_PLACE[sec_key]
            tpl_sec = find_template_section(sections, sec_key)
            if not tpl_sec:
                groups_to_insert.append(group)
                continue

            slide_texts = [s['text'] for s in group.get('slides', [])]
            reading_slides = list(group.get('reading_slides') or [])
            psalm_stanzas = list(group.get('psalm_stanzas') or [])
            if not slide_texts and not reading_slides and not psalm_stanzas:
                continue

            if replace_mode == 'body':
                # Edição multi-slide com overflow clonado
                body_slides = []
                found_title = False
                last_block_idx = -1
                for sid in tpl_sec['sld_ids']:
                    rid = id_to_rid.get(sid)
                    fname = rid_to_file.get(rid) if rid else None
                    if not fname:
                        continue
                    title = get_slide_title_text(src, fname)
                    if not found_title:
                        if title_keyword.upper() in title and slide_has_body(src, fname):
                            found_title = True
                            body_slides.append(fname)
                            if sid in id_to_idx:
                                last_block_idx = id_to_idx[sid]
                    else:
                        if title and title_keyword.upper() not in title:
                            break
                        if slide_has_body(src, fname):
                            body_slides.append(fname)
                        if sid in id_to_idx:
                            last_block_idx = id_to_idx[sid]

                if not body_slides:
                    groups_to_insert.append(group)
                    continue

                for i, fname in enumerate(body_slides):
                    txt = slide_texts[i] if i < len(slide_texts) else ''
                    if txt:
                        if sec_key in RUBRIC_KEYS:
                            current = get_shape_full_text(src, fname, 'body')
                            txt = format_rubric_text(current, txt)
                        data = src.read(f'ppt/slides/{fname}')
                        edit_in_place_files[fname] = replace_body_text(data, txt)

                overflow = slide_texts[len(body_slides):]
                clone_src = body_slides[-1]
                if psalm_stanzas and last_block_idx >= 0:
                    refrain_src = body_slides[0]
                    refrain_text = (
                        slide_texts[0] if slide_texts
                        else get_shape_full_text(src, refrain_src, 'body')
                    )
                    psalm_slides = []
                    for stanza in psalm_stanzas:
                        psalm_slides.extend(stanza)
                        if refrain_text.strip():
                            psalm_slides.append({
                                'text': refrain_text,
                                'bold': False,
                                'clone_from': refrain_src,
                            })
                    if psalm_slides:
                        groups_to_insert.append({
                            'section_key': sec_key,
                            '_force_insert_after': last_block_idx,
                            'slides': psalm_slides,
                        })
                elif overflow and last_block_idx >= 0:
                    groups_to_insert.append({
                        'section_key': sec_key,
                        '_force_insert_after': last_block_idx,
                        'slides': [{'text': t, 'bold': False, 'clone_from': clone_src}
                                   for t in overflow],
                    })

                if reading_slides and last_block_idx >= 0:
                    groups_to_insert.append({
                        'section_key': sec_key,
                        '_force_insert_after': last_block_idx,
                        'slides': reading_slides,
                    })

            else:
                # Edição de slide único (textbox ou title)
                txt = slide_texts[0] if slide_texts else ''
                target_idx = -1
                for sid in tpl_sec['sld_ids']:
                    rid = id_to_rid.get(sid)
                    fname = rid_to_file.get(rid) if rid else None
                    if not fname:
                        continue
                    title = get_slide_title_text(src, fname)
                    if title_keyword.upper() in title:
                        target_idx = id_to_idx.get(sid, -1)
                        if not txt:
                            break
                        data = src.read(f'ppt/slides/{fname}')
                        target_mode = replace_mode
                        # Alguns moldes usam corpo em vez de caixa livre na segunda leitura.
                        if (target_mode == 'textbox'
                                and not get_shape_full_text(src, fname, 'textbox').strip()
                                and slide_has_body(src, fname)):
                            target_mode = 'body'
                        # Rúbrica: mantém o label, substitui só a referência
                        final_txt = format_rubric_text(
                            get_shape_full_text(src, fname, target_mode), txt)
                        if target_mode == 'textbox':
                            edit_in_place_files[fname] = replace_textbox_text(data, final_txt)
                        elif target_mode == 'body':
                            edit_in_place_files[fname] = replace_body_text(data, final_txt)
                        else:  # 'title'
                            edit_in_place_files[fname] = replace_title_text(data, final_txt)
                        break

                if target_idx < 0 and sec_key == 'segunda_leitura':
                    clone_fname, _ = find_slide_by_title_keyword(
                        src, tpl_sec, 'PRIMEIRA', id_to_idx, id_to_rid, rid_to_file)
                    if not clone_fname:
                        groups_to_insert.append(group)
                        continue
                    anchor_idx = find_reading_insert_after(
                        src, sections, sec_key, id_to_idx, id_to_rid, rid_to_file, sldIdLst)
                    fallback_slides = []
                    if txt:
                        rubric_text = format_rubric_text('Segunda Leitura', txt)
                        rubric_slide = {'text': rubric_text, 'bold': False}
                        if clone_fname:
                            rubric_slide['clone_from'] = clone_fname
                            rubric_slide['title_text'] = 'Segunda leitura'
                        fallback_slides.append(rubric_slide)
                    fallback_slides.extend(reading_slides)
                    if fallback_slides and anchor_idx >= 0:
                        groups_to_insert.append({
                            'section_key': sec_key,
                            '_force_insert_after': anchor_idx,
                            'slides': fallback_slides,
                        })
                elif target_idx < 0:
                    groups_to_insert.append(group)
                elif reading_slides:
                    groups_to_insert.append({
                        'section_key': sec_key,
                        '_force_insert_after': target_idx,
                        'slides': reading_slides,
                    })

        # Prepara lista de novos slides com posição de inserção
        insert_plan = []  # list of dicts: insert_after, slides, section_el

        for group in groups_to_insert:
            sec_key    = group.get('section_key', 'outro')
            sec_slides = complete_group_slides(group)
            if not sec_slides:
                continue

            is_custom = bool(group.get('custom_section'))
            tpl_sec = find_template_section(sections, sec_key)
            sec_el  = tpl_sec['el'] if tpl_sec and not is_custom else None

            if '_force_insert_after' in group:
                # Overflow de edit-in-place: inserir após o último slide do bloco
                insert_after = group['_force_insert_after']
            elif is_custom:
                anchor_key = group.get('insert_after_key') or 'outro'
                insert_after = find_anchor_insert_after(src, sections, anchor_key, id_to_idx, id_to_rid, rid_to_file, sldIdLst)
            elif tpl_sec and sec_key in READING_ORDER:
                insert_after = find_reading_insert_after(
                    src, sections, sec_key, id_to_idx, id_to_rid, rid_to_file, sldIdLst)
            elif sec_key == PREFACIO_KEY and tpl_sec and tpl_sec['sld_ids']:
                # Prefácio Eucarístico: inserir após os slides fixos (capa + PREFACIO_FIXED_COUNT slides de texto)
                # Ordena os IDs da seção pela posição global para garantir a ordem correta
                sorted_sids = sorted(tpl_sec['sld_ids'],
                                     key=lambda sid: id_to_idx.get(sid, 99999))
                # Pula: sorted_sids[0]=capa, [1]=fixo1, [2]=fixo2 → inserir após índice PREFACIO_FIXED_COUNT
                target_pos = min(PREFACIO_FIXED_COUNT, len(sorted_sids) - 1)
                target_sid = sorted_sids[target_pos]
                insert_after = id_to_idx.get(target_sid, len(list(sldIdLst)) - 1)
            elif tpl_sec and tpl_sec['sld_ids']:
                # Inserir após a primeira capa (entre as duas capas da seção)
                first_idx = len(list(sldIdLst))  # fallback = final
                for sid in tpl_sec['sld_ids']:
                    if sid in id_to_idx and id_to_idx[sid] < first_idx:
                        first_idx = id_to_idx[sid]
                insert_after = first_idx
            else:
                insert_after = len(list(sldIdLst)) - 1  # final

            insert_plan.append({
                'insert_after': insert_after,
                'slides': sec_slides,
                'section_el': sec_el,
                'custom_section': is_custom,
                'title': group.get('title', '').strip() or 'Seção extra',
                'cover_image': group.get('cover_image', ''),
            })

        # Ordena pelo insert_after decrescente para inserir sem deslocar índices
        insert_plan.sort(key=lambda x: x['insert_after'], reverse=True)

        # Gera slides novos e monta plano de inserção no XML
        new_slide_files = {}  # nome → conteúdo
        new_media_files = {}
        file_overrides = {}
        menu_links = []
        rid_counter = max_rid
        sid_counter = max_sld_id

        for plan in insert_plan:
            insert_after = plan['insert_after']
            sec_el       = plan['section_el']
            new_elements = []  # (sldId_el, rid)
            first_custom_slide_target = None

            render_slides = plan['slides']
            cover_media_file = None
            cover_src_rect = ''
            if plan.get('custom_section'):
                cover_image = plan.get('cover_image', '')
                if not cover_image or not os.path.isfile(cover_image):
                    continue
                ext = os.path.splitext(cover_image)[1].lower().lstrip('.') or 'jpg'
                if ext not in ('jpg', 'jpeg', 'png'):
                    continue
                cover_media_file = next_available_media_name(all_names, new_media_files, 'custom_section_cover', ext)
                with open(cover_image, 'rb') as img:
                    new_media_files[f'ppt/media/{cover_media_file}'] = img.read()
                ensure_default_content_type(ct_root, ext)
                cover_src_rect = build_src_rect_for_cover(cover_image, cx, cy)
                cover_slide = {'kind': 'custom_cover', 'title': plan.get('title', 'Seção extra')}
                render_slides = [cover_slide] + plan['slides'] + [cover_slide]

            for s in render_slides:
                if s.get('kind') == 'liturgy_response':
                    continue
                rid_counter += 1
                sid_counter += 1
                n = len([k for k in new_slide_files if re.match(r'ppt/slides/slide\d+\.xml$', k)])
                # Usa o próximo número disponível
                existing = set(int(re.search(r'\d+', x).group())
                               for x in all_names if re.match(r'^ppt/slides/slide\d+\.xml$', x))
                used_new  = set(int(re.search(r'\d+', k).group())
                                for k in new_slide_files if re.match(r'^ppt/slides/slide\d+\.xml$', k))
                slide_n = max((existing | used_new), default=0) + 1

                slide_name = f'ppt/slides/slide{slide_n}.xml'
                rels_name  = f'ppt/slides/_rels/slide{slide_n}.xml.rels'

                clone_from = s.get('clone_from')
                if s.get('kind') == 'custom_cover':
                    new_slide_files[slide_name] = build_custom_cover_slide_xml(
                        s.get('title', plan.get('title', 'Seção extra')),
                        'rId2',
                        cx,
                        cy,
                        cover_src_rect,
                    ).encode('utf-8')
                    new_slide_files[rels_name] = build_custom_cover_rels(layout_target, cover_media_file).encode('utf-8')
                elif clone_from:
                    # Clona slide do template e substitui o texto, mantendo o estilo
                    clone_data = src.read(f'ppt/slides/{clone_from}')
                    if s.get('title_text'):
                        clone_data = replace_title_text(clone_data, s['title_text'])
                    new_slide_files[slide_name] = replace_body_text(clone_data, s['text'])
                    # Copia as rels do slide clonado, removendo notesSlide
                    clone_rels_name = f'ppt/slides/_rels/{clone_from}.rels'
                    if clone_rels_name in all_names:
                        clone_rels_data = src.read(clone_rels_name)
                        rels_root = etree.fromstring(clone_rels_data)
                        for r in list(rels_root):
                            if 'notesSlide' in r.get('Type', ''):
                                rels_root.remove(r)
                        new_slide_files[rels_name] = etree.tostring(
                            rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                    else:
                        new_slide_files[rels_name] = build_slide_rels(layout_target).encode('utf-8')
                else:
                    new_slide_files[slide_name] = build_black_slide_xml(
                        s['text'], s['bold'], cx, cy, layout_target
                    ).encode('utf-8')
                    new_slide_files[rels_name] = build_slide_rels(layout_target).encode('utf-8')

                rid = f'rId{rid_counter}'
                sid = sid_counter

                # Adiciona rel na presentation
                rel_el = etree.SubElement(pres_rels_root, f'{{{PKG_NS}}}Relationship')
                rel_el.set('Id', rid)
                rel_el.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
                rel_el.set('Target', f'slides/slide{slide_n}.xml')

                # Override no Content_Types
                ov_el = etree.SubElement(ct_root, f'{{{CT_NS}}}Override')
                ov_el.set('PartName', f'/ppt/slides/slide{slide_n}.xml')
                ov_el.set('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')

                # Cria elemento sldId para inserir na posição correta
                sldId_el = etree.Element(px('sldId'))
                sldId_el.set('id', str(sid))
                sldId_el.set(rx('id'), rid)

                new_elements.append((sldId_el, sid))
                if plan.get('custom_section') and first_custom_slide_target is None:
                    first_custom_slide_target = f'slide{slide_n}.xml'

            # Insere na posição correta no sldIdLst
            pos = insert_after + 1
            for offset, (sldId_el, sid) in enumerate(new_elements):
                sldIdLst.insert(pos + offset, sldId_el)

            # Atualiza posições de seções seguintes
            for other in insert_plan:
                if other is not plan and other['insert_after'] >= insert_after:
                    other['insert_after'] += len(new_elements)

            # Adiciona IDs à seção no template
            if sec_el is not None:
                sec_sldIdLst = sec_el.find(p14('sldIdLst'))
                if sec_sldIdLst is None:
                    sec_sldIdLst = etree.SubElement(sec_el, p14('sldIdLst'))
                for _, sid in new_elements:
                    el = etree.SubElement(sec_sldIdLst, p14('sldId'))
                    el.set('id', str(sid))
            elif plan.get('custom_section') and new_elements:
                sections_parent = sections[0]['el'].getparent() if sections else None
                if sections_parent is not None:
                    custom_sec = create_custom_section_element(plan.get('title', 'Seção extra'), [sid for _, sid in new_elements])
                    sections_parent.append(custom_sec)
                if first_custom_slide_target:
                    menu_links.append({'title': plan.get('title', 'Seção extra'), 'target': first_custom_slide_target, 'sid': new_elements[0][1]})

        if menu_links and 'ppt/slides/slide1.xml' in all_names:
            rels_name = 'ppt/slides/_rels/slide1.xml.rels'
            if rels_name in all_names:
                final_order = {int(el.get('id')): i for i, el in enumerate(sldIdLst)}
                menu_links.sort(key=lambda link: final_order.get(link.get('sid'), 10**9))
                slide1_xml, slide1_rels = add_menu_links_to_slide1(
                    src.read('ppt/slides/slide1.xml'),
                    src.read(rels_name),
                    menu_links,
                    cx,
                    cy,
                )
                file_overrides['ppt/slides/slide1.xml'] = slide1_xml
                file_overrides[rels_name] = slide1_rels
        sort_sections_by_slide_order(pres_root)

        # Escreve o ZIP de saída
        skip_exact = {'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'}
        skip_pat   = [r'^ppt/notesSlides/', r'^ppt/notesMasters/', r'^metadata']
        remove_rel_types = {'notesSlide', 'notesMaster'}

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as out:
            for name in all_names:
                if name in skip_exact: continue
                if any(re.match(p, name) for p in skip_pat): continue
                try:
                    data = file_overrides.get(name)
                    if data is None:
                        data = src.read(name)
                    # Remove rels de notesSlide de cada slide
                    if name not in file_overrides and re.match(r'^ppt/slides/_rels/slide\d+\.xml\.rels$', name):
                        root = etree.fromstring(data)
                        for r in list(root):
                            if any(x in r.get('Type','') for x in remove_rel_types):
                                root.remove(r)
                        data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                    # Aplica edição in-place se houver
                    slide_fname = name.replace('ppt/slides/', '')
                    if slide_fname in edit_in_place_files:
                        data = edit_in_place_files[slide_fname]
                    out.writestr(name, data)
                except Exception:
                    pass

            # Escreve novos slides
            for name, data in new_slide_files.items():
                out.writestr(name, data)
            for name, data in new_media_files.items():
                out.writestr(name, data)

            # Escreve XMLs atualizados
            out.writestr('ppt/presentation.xml',
                etree.tostring(pres_root, xml_declaration=True, encoding='UTF-8', standalone=True))
            out.writestr('ppt/_rels/presentation.xml.rels',
                etree.tostring(pres_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True))
            out.writestr('[Content_Types].xml',
                etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True))

def generate_no_template(output_path, section_groups):
    from pptx import Presentation
    from pptx.util import Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    prs = Presentation()
    prs.slide_width  = Emu(9144000)
    prs.slide_height = Emu(6858000)
    layout = prs.slide_layouts[6]

    for group in section_groups:
        if group.get('custom_section') and group.get('cover_image') and os.path.isfile(group.get('cover_image')):
            for cover_idx in [0]:
                slide = prs.slides.add_slide(layout)
                slide.shapes.add_picture(group['cover_image'], 0, 0, width=prs.slide_width, height=prs.slide_height)
                txBox = slide.shapes.add_textbox(0, 0, prs.slide_width, prs.slide_height)
                tf = txBox.text_frame
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                run = p.add_run()
                run.text = group.get('title', 'Seção extra').upper()
                run.font.name = 'Arial Black'
                run.font.size = Pt(40)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xF2, 0xCC, 0x7E)
        for s in complete_group_slides(group):
            slide = prs.slides.add_slide(layout)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
            cx, cy = prs.slide_width, prs.slide_height
            txBox = slide.shapes.add_textbox(0, 0, cx, cy)
            tf = txBox.text_frame
            tf.word_wrap = True
            bp = txBox.text_frame._txBody.find(f'{{{A_NS}}}bodyPr')
            if bp is None:
                from lxml import etree as _et
                bp = _et.SubElement(txBox.text_frame._txBody, f'{{{A_NS}}}bodyPr')
            bp.set('anchor', 'ctr')
            for idx, line in enumerate(s['text'].split('\n')):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.CENTER
                run = p.add_run()
                run.text = line
                run.font.name = 'Arial Black'
                run.font.size = Pt(54)
                run.font.bold = s['bold']
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        if group.get('custom_section') and group.get('cover_image') and os.path.isfile(group.get('cover_image')):
            slide = prs.slides.add_slide(layout)
            slide.shapes.add_picture(group['cover_image'], 0, 0, width=prs.slide_width, height=prs.slide_height)
            txBox = slide.shapes.add_textbox(0, 0, prs.slide_width, prs.slide_height)
            tf = txBox.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = group.get('title', 'Seção extra').upper()
            run.font.name = 'Arial Black'
            run.font.size = Pt(40)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xF2, 0xCC, 0x7E)
    prs.save(output_path)

def main():
    if len(sys.argv) < 3:
        print(json.dumps({"success": False, "error": "Argumentos insuficientes"}))
        sys.exit(1)

    slides_arg    = sys.argv[1]
    output_path   = sys.argv[2]
    template_path = sys.argv[3] if len(sys.argv) > 3 else None

    try:
        if os.path.isfile(slides_arg):
            with open(slides_arg, 'r', encoding='utf-8') as f:
                data = json.load(f)
        else:
            data = json.loads(slides_arg)
        # Suporta formato antigo (lista plana) e novo (lista de seções)
        if data and isinstance(data[0], dict) and 'section_key' in data[0]:
            section_groups = data
        else:
            section_groups = [{'section_key': 'outro', 'slides': data}]
    except Exception as e:
        print(json.dumps({"success": False, "error": "JSON inválido: " + str(e)}))
        sys.exit(1)

    try:
        if template_path and os.path.isfile(template_path):
            generate_with_template(template_path, output_path, section_groups)
        else:
            generate_no_template(output_path, section_groups)
        print(json.dumps({"success": True}))
    except Exception as e:
        import traceback
        print(json.dumps({"success": False, "error": str(e), "trace": traceback.format_exc()}))
        sys.exit(1)

if __name__ == '__main__':
    main()
