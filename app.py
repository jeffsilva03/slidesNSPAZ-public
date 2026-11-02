import os
import re
import sys
import json
import zlib
import html
import zipfile
import tempfile
import urllib.request
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory
from werkzeug.exceptions import RequestEntityTooLarge

import subprocess

try:
    from extract_text import extract_file_text as robust_extract_file_text
    from extract_text import extract_url_text as robust_extract_url_text
except Exception:
    robust_extract_file_text = None
    robust_extract_url_text = None

BASE_DIR = Path(__file__).parent

app = Flask(__name__, static_folder=str(BASE_DIR / 'assets' / 'images'), static_url_path='/assets/images')
MAX_TEMPLATE_UPLOAD_MB = int(os.environ.get('MAX_TEMPLATE_UPLOAD_MB', '300'))
app.config['MAX_CONTENT_LENGTH'] = MAX_TEMPLATE_UPLOAD_MB * 1024 * 1024

CHARS_PER_LINE = 14
MAX_LINES = 4
HARD_MAX_LINES = 5
SHORT_SLIDE_LINES = 2
SAFE_MERGED_MAX_LINES = 5
DOMINICAL_TEMPLATE_PATH = BASE_DIR / 'assets' / 'template.pptx'
MAOS_TEMPLATE_PATH = BASE_DIR / 'assets' / 'template_maos_ensanguentadas.pptx'
CERCO_TEMPLATE_PATH = BASE_DIR / 'assets' / 'template_cerco_jerico.pptx'
TEMPLATE_PATH = DOMINICAL_TEMPLATE_PATH
UPLOAD_DIR = BASE_DIR / 'uploads'
OUTPUT_DIR = BASE_DIR / 'output'

(BASE_DIR / 'assets').mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BASE_LITURGICAL_SECTIONS = {
    'cantico_entrada':    'Cântico de Entrada',
    'ato_penitencial':    'Ato Penitencial',
    'gloria':             'Glória',
    'primeira_leitura':   'Primeira Leitura',
    'salmo_responsorial': 'Salmo Responsorial',
    'segunda_leitura':    'Segunda Leitura',
    'aclamacao_evangelho':'Aclamação ao Evangelho',
    'evangelho':          'Evangelho',
    'preces_oracao_dizimista': 'Preces - Oração do Dizimista',
    'ofertorio':          'Ofertório',
    'prefacio_eucaristico': 'Prefácio Eucarístico',
    'santo':              'Santo',
    'cordeiro':           'Cordeiro de Deus',
    'comunhao':           'Comunhão',
    'pos_comunhao':       'Pós-Comunhão',
    'cantico_final':      'Cântico Final',
    'outro':              'Outro',
}

MAOS_LITURGICAL_SECTIONS = {
    'cantico_entrada':    'Cântico de Entrada',
    'aspercao':           'Asperção',
    'ato_penitencial':    'Ato Penitencial',
    'gloria':             'Glória',
    'primeira_leitura':   'Primeira Leitura',
    'salmo_responsorial': 'Salmo Responsorial',
    'segunda_leitura':    'Segunda Leitura',
    'aclamacao_evangelho':'Aclamação ao Evangelho',
    'evangelho':          'Evangelho',
    'preces_oracao_dizimista': 'Preces - Oração do Dizimista',
    'ofertorio':          'Ofertório',
    'prefacio_eucaristico': 'Prefácio Eucarístico',
    'santo':              'Santo',
    'cordeiro':           'Cordeiro de Deus',
    'comunhao':           'Comunhão',
    'translado':          'Translado',
    'pos_comunhao':       'Pós-Comunhão',
    'cantico_final':      'Cântico Final',
    'outro':              'Outro',
}

CERCO_LITURGICAL_SECTIONS = {
    'cantico_entrada': BASE_LITURGICAL_SECTIONS['cantico_entrada'],
    'entrada_sao_miguel': 'Entrada da Imagem de São Miguel',
    **{
        key: label
        for key, label in BASE_LITURGICAL_SECTIONS.items()
        if key != 'cantico_entrada'
    },
}

LITURGICAL_SECTION_SETS = {
    'dominical': BASE_LITURGICAL_SECTIONS,
    'maos': MAOS_LITURGICAL_SECTIONS,
    'jerico': CERCO_LITURGICAL_SECTIONS,
}
LITURGICAL_SECTIONS = BASE_LITURGICAL_SECTIONS

RUBRIC_SECTIONS = {'primeira_leitura', 'segunda_leitura', 'evangelho'}
READING_TEXT_SECTIONS = {'primeira_leitura', 'segunda_leitura', 'evangelho'}
PSALM_SECTION = 'salmo_responsorial'
IMAGE_EXTENSIONS = {'jpg', 'jpeg', 'png'}
PREFACIO_KEY = 'prefacio_eucaristico'

def normalize_mass_type(value):
    return value if value in {'maos', 'jerico'} else 'dominical'

def get_liturgical_sections(mass_type):
    return LITURGICAL_SECTION_SETS.get(
        normalize_mass_type(mass_type),
        LITURGICAL_SECTION_SETS['dominical'],
    )

def find_bundled_maos_template_path():
    for path in BASE_DIR.glob('SLIDE MODELO - M*OS ENSANGUENTADAS.pptx'):
        if not path.name.startswith('~') and path.exists():
            return path
    return None

def find_bundled_template_path(mass_type):
    if normalize_mass_type(mass_type) == 'maos':
        return find_bundled_maos_template_path()
    path = BASE_DIR / 'SLIDE MODELO - DOMINICAL.pptx'
    return path if path.exists() else None

def newer_template_path(primary, bundled):
    if primary and primary.exists() and bundled and bundled.exists():
        return bundled if bundled.stat().st_mtime > primary.stat().st_mtime else primary
    if primary and primary.exists():
        return primary
    if bundled and bundled.exists():
        return bundled
    return None

def get_template_path(mass_type):
    mass_type = normalize_mass_type(mass_type)
    if mass_type == 'maos':
        return MAOS_TEMPLATE_PATH
    if mass_type == 'jerico':
        return CERCO_TEMPLATE_PATH
    return DOMINICAL_TEMPLATE_PATH

def get_effective_template_path(mass_type):
    path = get_template_path(mass_type)
    return newer_template_path(path, find_bundled_template_path(mass_type))

def get_template_status(mass_type):
    path = get_template_path(mass_type)
    uploaded = path.exists()
    bundled = find_bundled_template_path(mass_type)
    effective = newer_template_path(path, bundled)
    source = 'uploaded' if effective == path else ('bundled' if effective else 'none')
    return {
        'success': True,
        'exists': effective is not None,
        'uploaded': uploaded,
        'source': source,
    }

# ================================================================
# TEXT CLEANING
# ================================================================

def limpar_texto(text):
    text = re.sub(r'[\U00010000-\U0010FFFF]', '', text)
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
    lines = []
    for line in text.split('\n'):
        line = line.strip()
        if line == '':
            lines.append('')
            continue
        if line.count('|') > 2:
            continue
        if line.count('/') > 4 and len(line) < 100:
            continue
        if re.match(r'^https?://', line):
            continue
        if re.search(r'(copyright|todos os direitos|cookie|privacidade)', line, re.I):
            continue
        if re.match(r'^[\d\s.,\-]+$', line) and len(line) < 15:
            continue
        lines.append(line)
    return '\n'.join(lines)

def fix_punctuation(text):
    text = re.sub(r'([.!?;:,])([A-Za-z\u00C0-\u024F])', r'\1 \2', text)
    text = re.sub(r'([A-Za-z\u00C0-\u024F.!?])(["\'\)\]])([A-Z\u00C0-\u00DC])', r'\1\2 \3', text)
    return text

def capitalize_sacred_words(text):
    main_words = [
        'deus', 'senhor', 'jesus', 'cristo', 'esp\u00edrito', 'espirito',
        'salvador', 'redentor', 'messias', 'alt\u00edssimo', 'altissimo',
        'onipotente', 'maria', 'virgem', 'evangelho',
        'aleluia', 'am\u00e9m', 'hosana', 'gl\u00f3ria', 'gloria',
    ]
    for w in main_words:
        text = re.sub(
            r'\b' + re.escape(w) + r'\b',
            lambda m: m.group(0)[0].upper() + m.group(0)[1:],
            text, flags=re.I | re.U
        )
    text = re.sub(r'\bnossa\s+senhora\b',       'Nossa Senhora',  text, flags=re.I | re.U)
    text = re.sub(r'\besp[i\u00ed]rito\s+santo\b', 'Esp\u00edrito Santo', text, flags=re.I | re.U)
    text = re.sub(r'\bjesus\s+cristo\b',        'Jesus Cristo',   text, flags=re.I | re.U)
    text = re.sub(r'\bm\u00e3e\s+de\s+deus\b', 'M\u00e3e de Deus', text, flags=re.I | re.U)
    text = re.sub(r'\bfilho\s+de\s+deus\b',    'Filho de Deus',  text, flags=re.I | re.U)
    return text

def capitalize(s):
    if not s:
        return s
    return s[0].upper() + s[1:]

def word_wrap_custom(text):
    words = text.split()
    if not words:
        return []
    lines = []
    cur = ''
    for word in words:
        test = cur + ' ' + word if cur else word
        if len(test) > CHARS_PER_LINE and len(cur) > 0:
            lines.append(cur)
            cur = word
        else:
            cur = test
    if cur:
        lines.append(cur)
    return lines

def slide_lines_count(slide):
    return len([line for line in slide.split('\n') if line.strip()])

def merge_short_slides(slides):
    merged = []
    i = 0
    while i < len(slides):
        current = slides[i]
        if i + 1 < len(slides):
            current_count = slide_lines_count(current)
            next_count = slide_lines_count(slides[i + 1])
            if (
                0 < current_count <= SHORT_SLIDE_LINES
                and current_count + next_count <= SAFE_MERGED_MAX_LINES
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
    units = [part.strip() for part in re.split(r'(?<=[.!?;:])\s+', line) if part.strip()]
    if not units:
        units = [line]

    expanded = []
    for unit in units:
        wrapped = word_wrap_custom(unit)
        if len(wrapped) > HARD_MAX_LINES:
            comma_units = [part.strip() for part in re.split(r'(?<=,)\s+', unit) if part.strip()]
            if len(comma_units) > 1:
                expanded.extend(comma_units)
                continue
        expanded.append(unit)
    return expanded

def flush_slide(slides, cur_slide):
    if cur_slide:
        slides.append('\n'.join(cur_slide))
        cur_slide.clear()

def append_wrapped_unit(para_slides, cur_slide, wrapped):
    while len(wrapped) > HARD_MAX_LINES:
        flush_slide(para_slides, cur_slide)
        para_slides.append('\n'.join(wrapped[:HARD_MAX_LINES]))
        wrapped = wrapped[HARD_MAX_LINES:]
    if not wrapped:
        return

    if cur_slide and len(cur_slide) + len(wrapped) > HARD_MAX_LINES:
        flush_slide(para_slides, cur_slide)
    cur_slide.extend(wrapped)

def split_into_slides(text):
    paragraphs = []
    cur = []
    for line in text.split('\n'):
        t = line.strip()
        if t == '':
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
        cur_slide = []
        for line in para:
            for unit in split_phrase_units(line):
                append_wrapped_unit(para_slides, cur_slide, word_wrap_custom(unit))
        flush_slide(para_slides, cur_slide)
        slides.extend(merge_short_slides(para_slides))

    final = []
    i = 0
    while i < len(slides):
        slide_words = slides[i].split()
        if len(slide_words) <= 1:
            current_lines = slide_lines_count(slides[i])
            if final and slide_lines_count(final[-1]) + current_lines <= HARD_MAX_LINES:
                final[-1] += '\n' + slides[i]
            elif i + 1 < len(slides) and current_lines + slide_lines_count(slides[i + 1]) <= HARD_MAX_LINES:
                slides[i + 1] = slides[i] + '\n' + slides[i + 1]
            else:
                final.append(slides[i])
        else:
            final.append(slides[i])
        i += 1
    return final

def preprocess_lyric_text(text):
    text = limpar_texto(text)
    text = fix_punctuation(text)
    return capitalize_sacred_words(text)

def strip_verse_numbers(text):
    """Remove a numeração dos versículos sem alterar referências bíblicas."""
    return re.sub(r'(?<![\w])\d{1,3}(?=[A-Za-z\u00C0-\u024F“"\'\[])', '', text)

def split_text_blocks(text):
    blocks = []
    cur = []
    for line in text.split('\n'):
        if line.strip() == '':
            if cur:
                blocks.append('\n'.join(cur).strip())
                cur = []
        else:
            cur.append(line.strip())
    if cur:
        blocks.append('\n'.join(cur).strip())
    return [block for block in blocks if block]

def build_projected_text_slides(text):
    text = strip_verse_numbers(preprocess_lyric_text(text))
    return [
        {'text': capitalize(part), 'bold': False}
        for part in split_into_slides(text)
        if part.strip()
    ]

def split_psalm_stanza_blocks(text):
    text = text.replace('\r\n', '\n').replace('\r', '\n').strip()
    if not text:
        return []

    blank_blocks = [
        block.strip()
        for block in re.split(r'\n\s*\n+', text)
        if block.strip()
    ]
    if len(blank_blocks) > 1:
        return blank_blocks

    lines = [line.strip() for line in text.split('\n') if line.strip()]
    blocks = []
    cur = []
    for line in lines:
        if re.match(r'^[-–—]\s+', line) and cur:
            blocks.append('\n'.join(cur).strip())
            cur = [line]
        else:
            cur.append(line)
    if cur:
        blocks.append('\n'.join(cur).strip())
    return [block for block in blocks if block]

def build_psalm_stanza_slides(text):
    text = strip_verse_numbers(preprocess_lyric_text(text))
    stanzas = []
    for block in split_psalm_stanza_blocks(text):
        slides = [
            {'text': capitalize(part), 'bold': False}
            for part in split_into_slides(block)
            if part.strip()
        ]
        if slides:
            stanzas.append(slides)
    return stanzas

# ================================================================
# FILE EXTRACTION
# ================================================================

def extract_from_file(filepath, ext):
    if robust_extract_file_text is not None:
        try:
            text = robust_extract_file_text(filepath, ext)
            if text and text.strip():
                return text
        except Exception:
            pass
    if ext == 'docx':
        return extract_docx(filepath)
    if ext == 'doc':
        return extract_doc(filepath)
    if ext == 'pdf':
        return extract_pdf(filepath)
    if ext == 'txt':
        return extract_txt(filepath)
    return ''

def extract_docx(filepath):
    try:
        with zipfile.ZipFile(filepath) as z:
            xml = z.read('word/document.xml').decode('utf-8', errors='replace')
    except Exception:
        return 'ERRO: Não foi possível abrir o DOCX'
    xml = re.sub(r'<w:br[^>]*/?>',  '\n', xml, flags=re.I)
    xml = re.sub(r'</w:p>',         '\n', xml, flags=re.I)
    text = re.sub(r'<[^>]+>', '', xml)
    text = html.unescape(text)
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    return limpar_texto('\n'.join(lines))

def extract_doc(filepath):
    try:
        with open(filepath, 'rb') as f:
            content = f.read()
    except Exception:
        return 'ERRO: Não foi possível ler o arquivo .doc'
    try:
        text = content.decode('windows-1252', errors='replace')
    except Exception:
        text = content.decode('latin-1', errors='replace')
    chunks = re.split(r'[^\x20-\x7e\x80-\xff\r\n\t]{5,}', text)
    best = ''
    for chunk in chunks:
        clean = chunk.replace('\r', '\n')
        clean = re.sub(r'[^\x20-\x7e\x80-\xff\n]', '', clean).strip()
        words = re.findall(r'[a-zA-Z\u00C0-\u00FF]{2,}', clean)
        if len(words) > 3 and len(clean) > len(best):
            best = clean
    best = re.sub(r'\n{3,}', '\n\n', best)
    return limpar_texto(best.strip() or 'ERRO: Não foi possível extrair texto do .doc')

def extract_pdf(filepath):
    try:
        with open(filepath, 'rb') as f:
            content = f.read()
    except Exception:
        return 'ERRO: Não foi possível ler o PDF'
    text = ''
    for stream_match in re.finditer(rb'stream\s*(.*?)\s*endstream', content, re.DOTALL):
        stream = stream_match.group(1)
        try:
            decoded = zlib.decompress(stream)
        except Exception:
            decoded = stream
        for t in re.findall(rb'\(([^)]*)\)', decoded):
            t = t.replace(b'\\(', b'(').replace(b'\\)', b')').replace(b'\\\\', b'\\')
            try:
                text += t.decode('latin-1')
            except Exception:
                pass
        for tj in re.findall(rb'\[(.*?)\]\s*TJ', decoded, re.DOTALL):
            for t in re.findall(rb'\(([^)]*)\)', tj):
                t = t.replace(b'\\(', b'(').replace(b'\\)', b')').replace(b'\\\\', b'\\')
                try:
                    text += t.decode('latin-1')
                except Exception:
                    pass
    if not text.strip():
        return 'ERRO: Não foi possível extrair texto do PDF'
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    return limpar_texto('\n'.join(lines))

def extract_txt(filepath):
    for enc in ['utf-8', 'iso-8859-1', 'windows-1252']:
        try:
            with open(filepath, 'r', encoding=enc) as f:
                return limpar_texto(f.read().strip())
        except Exception:
            continue
    return 'ERRO: Não foi possível ler o arquivo TXT'

def extract_from_url(url):
    if robust_extract_url_text is not None:
        try:
            text = robust_extract_url_text(url)
            if text and text.strip():
                return text
        except Exception:
            pass
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=15) as resp:
            html_bytes = resp.read()
            page = html_bytes.decode('utf-8', errors='replace')
    except Exception:
        return 'ERRO: Não foi possível acessar a URL'
    for tag in ['script','style','nav','header','footer','aside','form','noscript','iframe']:
        page = re.sub(f'<{tag}[^>]*>.*?</{tag}>', '', page, flags=re.S|re.I)
    page = re.sub(r'<!--.*?-->', '', page, flags=re.S)
    main = ''
    m = re.search(r'<(?:article|main)[^>]*>(.*?)</(?:article|main)>', page, re.S|re.I)
    if m:
        main = m.group(1)
    if not main:
        for p in [
            r'class="[^"]*(?:lyrics|letra|content|verse|cifra)[^"]*"[^>]*>(.*?)</div>',
            r'class="[^"]*(?:entry-content|post-content|article-body)[^"]*"[^>]*>(.*?)</div>',
        ]:
            m = re.search(p, page, re.S|re.I)
            if m:
                main = m.group(1)
                break
    if not main:
        m = re.search(r'<body[^>]*>(.*?)</body>', page, re.S|re.I)
        main = m.group(1) if m else page
    text = re.sub(r'<br\s*/?>', '\n', main, flags=re.I)
    text = re.sub(r'</p>',      '\n', text, flags=re.I)
    text = re.sub(r'</div>',    '\n', text, flags=re.I)
    text = re.sub(r'<[^>]+>', '', text)
    text = html.unescape(text)
    lines = []
    for line in text.split('\n'):
        cleaned = ' '.join(line.split()).strip()
        if len(cleaned) <= 1:
            continue
        if re.match(r'^(home|voltar|menu|buscar|login|sair|cadastrar|compartilhar|imprimir)$', cleaned, re.I):
            continue
        if re.match(r'^(tom|capo|cifra|acordes?)[\s:.]', cleaned, re.I):
            continue
        if re.search(r'(copyright|todos os direitos|cookie|privacidade)', cleaned, re.I):
            continue
        if re.match(r'^https?://', cleaned):
            continue
        lines.append(cleaned)
    result = '\n'.join(lines)
    if len(result) > 3000:
        short = []
        c = 0
        for l in lines:
            short.append(l)
            c += len(l)
            if c > 2000:
                break
        result = '\n'.join(short)
    return limpar_texto(result)

# ================================================================
# PPTX GENERATION
# ================================================================

def build_section_groups(sections):
    section_groups = []
    for sec in sections:
        key   = sec.get('key', 'outro')
        text  = sec.get('text', '').strip()
        reading_text = sec.get('reading_text', '').strip()
        slides = []
        prefacio_final_text = ''
        if text:
            if key in RUBRIC_SECTIONS:
                slides.append({'text': text, 'bold': False})
            elif key == PSALM_SECTION and reading_text:
                text = preprocess_lyric_text(text)
                slides.append({'text': capitalize(text), 'bold': False})
            else:
                text = preprocess_lyric_text(text)
                parts = []
                if key == PREFACIO_KEY:
                    blocks = split_text_blocks(text)
                    if len(blocks) > 1:
                        parts = split_into_slides('\n\n'.join(blocks[:-1]))
                        final_parts = split_into_slides(blocks[-1])
                        if final_parts:
                            prefacio_final_text = capitalize(final_parts.pop())
                            parts.extend(final_parts)
                    else:
                        parts = split_into_slides(text)
                        if parts:
                            prefacio_final_text = capitalize(parts.pop())
                else:
                    parts = split_into_slides(text)
                for part in parts:
                    slides.append({'text': capitalize(part), 'bold': False})
        group = {'section_key': key, 'slides': slides}
        if key in READING_TEXT_SECTIONS and reading_text:
            group['reading_slides'] = build_projected_text_slides(reading_text)
        elif key == PSALM_SECTION and reading_text:
            group['psalm_stanzas'] = build_psalm_stanza_slides(reading_text)
        if prefacio_final_text:
            group['prefacio_final_text'] = prefacio_final_text
        if sec.get('custom_section'):
            group.update({
                'custom_section': True,
                'title': sec.get('title', '').strip() or 'Seção extra',
                'insert_after_key': sec.get('insert_after_key', 'outro'),
                'cover_image': sec.get('cover_image', ''),
            })
        section_groups.append(group)
    return section_groups

def run_generate(output_path, sections, mass_type='dominical'):
    section_groups = build_section_groups(sections)
    if not section_groups:
        return False

    tmp = tempfile.NamedTemporaryFile(
        delete=False, suffix='.json', mode='w', encoding='utf-8', dir=UPLOAD_DIR
    )
    try:
        json.dump(section_groups, tmp, ensure_ascii=False)
        tmp.close()

        cmd = [sys.executable, str(BASE_DIR / 'generate_pptx.py'), tmp.name, output_path]
        template_path = get_effective_template_path(mass_type)
        if template_path:
            cmd.append(str(template_path))

        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        stdout = proc.stdout.strip()
        stderr = proc.stderr.strip()

        try:
            result = json.loads(stdout)
        except Exception:
            raise RuntimeError(f'Script output: {stdout}\nStderr: {stderr}')

        if not result.get('success'):
            err = result.get('error', 'Erro desconhecido')
            trace = result.get('trace', stderr)
            raise RuntimeError(f'{err}\n{trace}')

        return True
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass

# ================================================================
# ROUTES
# ================================================================

@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.after_request
def cors(response):
    response.headers['Access-Control-Allow-Origin']  = '*'
    response.headers['Access-Control-Allow-Methods'] = 'POST, GET, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response

@app.errorhandler(RequestEntityTooLarge)
def upload_too_large(error):
    return jsonify({
        'success': False,
        'error': f'Arquivo muito grande. Envie um .pptx de ate {MAX_TEMPLATE_UPLOAD_MB} MB.',
    }), 413

@app.route('/api', methods=['GET', 'POST', 'OPTIONS'])
def api():
    if request.method == 'OPTIONS':
        return '', 200
    action = request.args.get('action', '')
    mass_type = normalize_mass_type(request.args.get('mass_type', 'dominical'))
    if action == 'sections':
        return jsonify({'success': True, 'sections': get_liturgical_sections(mass_type)})
    if action == 'generate':
        return handle_generate()
    if action == 'download':
        return handle_download()
    if action == 'upload_template':
        return handle_upload_template(mass_type)
    if action == 'has_template':
        return jsonify(get_template_status(mass_type))
    if action == 'delete_template':
        get_template_path(mass_type).unlink(missing_ok=True)
        return jsonify({'success': True})
    if action == 'debug':
        return handle_debug()
    if action == 'debug_template':
        return handle_debug_template()
    return jsonify({'success': False, 'error': 'Ação inválida'})

def handle_generate():
    try:
        mass_type = normalize_mass_type(request.form.get('mass_type', 'dominical'))
        liturgical_sections = get_liturgical_sections(mass_type)
        sections = []
        errors   = []
        blocking_errors = []
        cleanup_files = []

        for key, label in liturgical_sections.items():
            reading_text = ''
            if key in READING_TEXT_SECTIONS or key == PSALM_SECTION:
                reading_text = request.form.get(f'reading_text_{key}', '').strip()
            text = request.form.get(f'text_{key}', '').strip()

            if text or reading_text:
                item = {'key': key, 'title': label, 'text': text}
                if reading_text:
                    item['reading_text'] = reading_text
                sections.append(item)

        custom_ids = request.form.getlist('custom_section_ids[]') or request.form.getlist('custom_section_ids')
        for raw_id in custom_ids:
            custom_id = re.sub(r'[^a-zA-Z0-9_\-]', '', raw_id or '')
            if not custom_id:
                continue

            title = request.form.get(f'custom_section_title_{custom_id}', '').strip()
            after_key = request.form.get(f'custom_section_after_{custom_id}', 'outro').strip() or 'outro'
            text = request.form.get(f'custom_section_text_{custom_id}', '').strip()
            cover = request.files.get(f'custom_section_cover_{custom_id}')
            has_cover = bool(cover and cover.filename)

            if not title and not text and not has_cover:
                continue
            if not title or not text or not has_cover:
                blocking_errors.append('Seção extra incompleta: preencha nome, capa e texto.')
                continue

            ext = Path(cover.filename).suffix.lower().lstrip('.')
            if ext not in IMAGE_EXTENSIONS:
                blocking_errors.append(f'{title}: capa .{ext} não suportada. Use .jpg, .jpeg ou .png.')
                continue

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f'.{ext}', dir=UPLOAD_DIR)
            cover.save(tmp.name)
            tmp.close()
            cleanup_files.append(tmp.name)
            sections.append({
                'key': f'custom_section_{custom_id}',
                'title': title,
                'text': text,
                'custom_section': True,
                'insert_after_key': after_key,
                'cover_image': tmp.name,
            })

        if blocking_errors:
            for path in cleanup_files:
                try:
                    os.unlink(path)
                except Exception:
                    pass
            return jsonify({'success': False, 'error': 'Corrija as seções extras antes de gerar.', 'errors': errors + blocking_errors})

        if not sections:
            return jsonify({'success': False, 'error': 'Nenhuma seção com conteúdo.', 'errors': errors})

        custom_name = request.form.get('custom_filename', '').strip()
        if custom_name:
            safe_name = re.sub(r'[^a-zA-Z0-9_\-\u00C0-\u024F ]', '', custom_name)
            safe_name = re.sub(r'\s+', '_', safe_name.strip())
            if not safe_name:
                safe_name = 'liturgia'
        else:
            safe_name = 'liturgia_' + datetime.now().strftime('%Y-%m-%d_%H%M%S')

        output_filename = safe_name + '.pptx'
        counter = 1
        while (OUTPUT_DIR / output_filename).exists():
            output_filename = f'{safe_name}_{counter}.pptx'
            counter += 1

        output_path = str(OUTPUT_DIR / output_filename)
        try:
            ok = run_generate(output_path, sections, mass_type)
        finally:
            for path in cleanup_files:
                try:
                    os.unlink(path)
                except Exception:
                    pass

        if ok:
            return jsonify({'success': True, 'filename': output_filename,
                            'sections_count': len(sections), 'errors': errors})
        return jsonify({'success': False, 'error': 'Erro ao gerar PPTX', 'errors': errors})

    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()})

def handle_download():
    filename = request.args.get('file', '')
    filename = Path(filename).name  # strip any path traversal
    filepath = OUTPUT_DIR / filename
    if not filename or not filepath.exists():
        return jsonify({'success': False, 'error': 'Arquivo não encontrado'}), 404
    return send_file(
        str(filepath),
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        as_attachment=True,
        download_name=filename,
    )

def handle_upload_template(mass_type):
    max_length = app.config.get('MAX_CONTENT_LENGTH')
    if max_length and request.content_length and request.content_length > max_length:
        return upload_too_large(None)
    if 'template' not in request.files:
        return jsonify({'success': False, 'error': 'Campo template não recebido pelo servidor'})
    f = request.files['template']
    if f.filename == '':
        return jsonify({'success': False, 'error': 'Nenhum arquivo enviado'})
    ext = Path(f.filename).suffix.lower()
    if ext != '.pptx':
        return jsonify({'success': False, 'error': 'Apenas arquivos .pptx são aceitos'})
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx', dir=UPLOAD_DIR)
    f.save(tmp.name)
    tmp.close()
    try:
        with zipfile.ZipFile(tmp.name) as z:
            pass  # just validate it's a valid zip
    except Exception:
        os.unlink(tmp.name)
        return jsonify({'success': False, 'error': 'Arquivo .pptx corrompido ou inválido'})
    os.replace(tmp.name, get_template_path(mass_type))
    return jsonify({'success': True})

def handle_debug():
    return jsonify({
        'python_version': sys.version,
        'template_exists': get_template_status(request.args.get('mass_type', 'dominical'))['exists'],
        'template_status': get_template_status(request.args.get('mass_type', 'dominical')),
        'output_dir_writable': os.access(OUTPUT_DIR, os.W_OK),
        'upload_dir_writable': os.access(UPLOAD_DIR, os.W_OK),
        'chars_per_line': CHARS_PER_LINE,
        'max_lines': MAX_LINES,
    })

def handle_debug_template():
    template_path = get_effective_template_path(request.args.get('mass_type', 'dominical'))
    if not template_path:
        return jsonify({'error': 'sem template'})
    try:
        with zipfile.ZipFile(template_path) as z:
            slide1_rels = z.read('ppt/slides/_rels/slide1.xml.rels').decode('utf-8', errors='replace')
            pres_xml    = z.read('ppt/presentation.xml').decode('utf-8', errors='replace')
        return jsonify({
            'slide1_rels':      slide1_rels,
            'pres_xml_snippet': pres_xml[:2000],
        })
    except Exception as e:
        return jsonify({'error': str(e)})

# ================================================================
# ENTRY POINT
# ================================================================

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
