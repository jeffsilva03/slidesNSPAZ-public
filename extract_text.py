#!/usr/bin/env python3
import html
import json
import os
import re
import sys
import urllib.request
import zipfile
import zlib
from pathlib import Path

try:
    from lxml import etree
    from lxml import html as lxml_html
except Exception:
    etree = None
    lxml_html = None

JUNK_RE = re.compile(
    r'(cookie|privacidade|termos de uso|pol[ií]tica|copyright|todos os direitos|'
    r'publicidade|an[uú]ncio|newsletter|compartilh|coment[aá]rios?|responder|'
    r'entrar|login|cadastrar|cadastro|baixar app|aceitar|relacionad[ao]s|'
    r'leia tamb[eé]m|veja tamb[eé]m|reportar|erro na letra|top m[uú]sicas|'
    r'exibi[cç][oõ]es?|exibic[oõ]es?|exibi.{0,4}es?|visualiza[cç][oõ]es?|visualiza.{0,6}es?|views?|acessos?|visitas?|'
    r'estat[ií]sticas?|n[uú]mero de exibi[cç][oõ]es?|'
    r'mais acessadas|facebook|instagram|twitter|whatsapp|telegram|youtube|spotify|'
    r'cifra club|letras\.mus|vagalume)',
    re.I,
)

NAV_EXACT = {
    'home', 'inicio', 'início', 'voltar', 'menu', 'buscar', 'pesquisar', 'login',
    'sair', 'cadastrar', 'compartilhar', 'imprimir', 'favoritar', 'salvar',
    'próxima', 'proxima', 'anterior', 'play', 'pause', 'ouvir', 'artista',
    'álbum', 'album', 'músicas', 'musicas', 'letras', 'cifras',
}

GOOD_ATTR_RE = re.compile(
    r'(lyrics?|letra|verse|verso|cifra-letra|cnt-letra|song|music|'
    r'content|post-content|entry-content|article-body|article|main|texto|poema)',
    re.I,
)

LYRIC_ATTR_RE = re.compile(
    r'(lyric-original|lyrics?|cnt-letra|cifra-letra|letra(?:__|[-_])?|'
    r'texto-letra|song-lyrics?|music-lyrics?|verse|verso)',
    re.I,
)

BLOCK_CLOSE_RE = re.compile(
    r'</(?:p|div|section|article|main|li|h[1-6]|blockquote|tr|td|pre)>',
    re.I,
)


def decode_bytes(data, content_type=''):
    if data.startswith(b'\xef\xbb\xbf'):
        return data.decode('utf-8-sig', errors='replace')
    m = re.search(r'charset=([\w.-]+)', content_type or '', re.I)
    encodings = []
    if m:
        encodings.append(m.group(1))
    head = data[:4000].decode('ascii', errors='ignore')
    m = re.search(r'<meta[^>]+charset=["\']?([\w.-]+)', head, re.I)
    if m:
        encodings.append(m.group(1))
    encodings += ['utf-8', 'windows-1252', 'iso-8859-1']
    for enc in encodings:
        try:
            return data.decode(enc, errors='replace')
        except Exception:
            pass
    return data.decode('utf-8', errors='replace')


def normalize_line(line):
    line = html.unescape(line)
    line = line.replace('\xa0', ' ').replace('\u200b', '')
    line = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', line)
    line = re.sub(r'\s+', ' ', line)
    return line.strip(' \t-–—•·')


def looks_like_chord_line(line):
    low = line.lower().strip()
    if re.match(r'^(tom|capo|cifra|acordes?|afina[cç][aã]o|intro|tablatura)\b', low):
        return True
    tokens = [t for t in re.split(r'\s+', line) if t]
    if len(tokens) < 2 or len(tokens) > 18:
        return False
    chord = re.compile(r'^[A-G](?:#|b)?m?(?:maj|min|dim|aug|sus|add)?\d*(?:/[A-G](?:#|b)?m?)?$')
    chordish = sum(1 for t in tokens if chord.match(t.strip('()[]{}|,.;:')))
    return chordish >= max(2, int(len(tokens) * 0.7))


def is_boilerplate_line(line):
    if not line or len(line) <= 1:
        return True
    low = line.lower().strip()
    if low in NAV_EXACT:
        return True
    if re.match(r'^https?://', low) or re.search(r'\S+@\S+\.\S+', low):
        return True
    if re.match(r'^[\d\s.,:;/-]+$', line) and len(line) < 18:
        return True
    if 'exibi' in low and 'letra' in low:
        return True
    if re.search(r'\b\d{1,3}(?:[.,]\d{3})+\b', line) and re.search(r'\b(exibi[cç][oõ]es?|exibic[oõ]es?|exibi.{0,4}es?|visualiza[cç][oõ]es?|visualiza.{0,6}es?|views?|acessos?|visitas?)\b', low):
        return True
    if re.search(r'\b(exibi[cç][oõ]es?|exibic[oõ]es?|exibi.{0,4}es?|visualiza[cç][oõ]es?|visualiza.{0,6}es?|views?|acessos?|visitas?|estat[ií]sticas?)\b', low):
        return True
    if line.count('|') > 2:
        return True
    if line.count('/') > 4 and len(line) < 120:
        return True
    if re.search(r'(function\s*\(|window\.|document\.|var\s+\w+|const\s+\w+|=>|\{|\})', line):
        return True
    if JUNK_RE.search(line):
        return True
    if looks_like_chord_line(line):
        return True
    return False


def clean_extracted_text(text, max_chars=6000):
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r'\n[ \t]+', '\n', text)
    text = re.sub(r'[ \t]+\n', '\n', text)
    raw_lines = text.split('\n')
    lines = []
    previous = ''
    blank = False
    for raw in raw_lines:
        line = normalize_line(raw)
        if not line:
            if lines and not blank:
                lines.append('')
                blank = True
            continue
        blank = False
        if is_boilerplate_line(line):
            continue
        sig = re.sub(r'\W+', '', line.lower())
        if sig and sig == previous:
            continue
        previous = sig
        lines.append(line)

    while lines and lines[0] == '':
        lines.pop(0)
    while lines and lines[-1] == '':
        lines.pop()

    result = '\n'.join(lines)
    result = re.sub(r'\n{3,}', '\n\n', result).strip()
    if len(result) > max_chars:
        kept, count = [], 0
        for line in result.split('\n'):
            kept.append(line)
            count += len(line) + 1
            if count >= max_chars:
                break
        result = '\n'.join(kept).strip()
    return result


def html_fragment_to_text(fragment):
    fragment = re.sub(r'<br\s*/?>', '\n', fragment, flags=re.I)
    fragment = BLOCK_CLOSE_RE.sub('\n', fragment)
    fragment = re.sub(r'<[^>]+>', ' ', fragment)
    return html.unescape(fragment)


def node_text(node):
    fragment = lxml_html.tostring(node, encoding='unicode', method='html')
    return html_fragment_to_text(fragment)


def extract_json_ld_text(page):
    texts = []
    for raw in re.findall(r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>', page, re.I | re.S):
        raw = html.unescape(raw).strip()
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            continue

        def walk(obj, parent_key=''):
            if isinstance(obj, dict):
                obj_type = str(obj.get('@type', '')).lower()
                for key, value in obj.items():
                    key_l = str(key).lower()
                    if key_l in ('lyrics', 'lyric') or (key_l == 'text' and ('lyric' in parent_key or 'creativework' in obj_type)):
                        if isinstance(value, str):
                            texts.append(value)
                        else:
                            walk(value, key_l)
                    elif key_l not in ('description', 'name', 'url', 'image', 'sameas'):
                        walk(value, key_l)
            elif isinstance(obj, list):
                for item in obj:
                    walk(item, parent_key)

        walk(data)

    candidates = [clean_extracted_text(text, max_chars=12000) for text in texts]
    candidates = [text for text in candidates if text and len(text) >= 40]
    if not candidates:
        return ''
    return max(candidates, key=len)


def extract_prioritized_lyrics(root):
    if lxml_html is None:
        return ''
    candidates = []
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue
        attrs = ' '.join([node.get('id', ''), node.get('class', '')])
        if not LYRIC_ATTR_RE.search(attrs):
            continue
        # Containers with lots of links are usually menus or "related lyrics" blocks.
        if len(node.findall('.//a')) > 4:
            continue
        cleaned = clean_extracted_text(node_text(node), max_chars=12000)
        lines = [line for line in cleaned.split('\n') if line.strip()]
        if len(cleaned) >= 30 and len(lines) >= 2:
            candidates.append((len(cleaned) + len(lines) * 25, cleaned))
    if not candidates:
        return ''
    candidates.sort(reverse=True, key=lambda item: item[0])
    return candidates[0][1]


def candidate_score(node, text):
    cleaned = clean_extracted_text(text, max_chars=12000)
    if len(cleaned) < 40:
        return -9999, ''
    lines = [line for line in cleaned.split('\n') if line.strip()]
    attrs = ' '.join([node.get('id', ''), node.get('class', ''), node.tag or ''])
    links = len(node.findall('.//a'))
    junk_hits = len(JUNK_RE.findall(text))
    avg_line = sum(len(line) for line in lines) / max(1, len(lines))
    score = min(len(cleaned), 5000) / 35
    score += min(len(lines), 80) * 4
    if GOOD_ATTR_RE.search(attrs):
        score += 120
    if node.tag in ('article', 'main'):
        score += 60
    score -= links * 12
    score -= junk_hits * 35
    if avg_line > 130:
        score -= 50
    return score, cleaned


def extract_from_html(page):
    json_ld_text = extract_json_ld_text(page)
    if json_ld_text:
        return json_ld_text

    if lxml_html is None:
        text = html_fragment_to_text(page)
        return clean_extracted_text(text)

    try:
        root = lxml_html.fromstring(page)
    except Exception:
        return clean_extracted_text(html_fragment_to_text(page))

    for bad in root.xpath('//script|//style|//nav|//header|//footer|//aside|//form|//noscript|//iframe|//svg|//button|//input|//select|//textarea'):
        parent = bad.getparent()
        if parent is not None:
            parent.remove(bad)

    lyric_text = extract_prioritized_lyrics(root)
    if lyric_text:
        return lyric_text

    candidates = []
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue
        attrs = ' '.join([node.get('id', ''), node.get('class', ''), node.tag])
        if node.tag in ('article', 'main', 'section') or GOOD_ATTR_RE.search(attrs):
            candidates.append(node)

    body = root.find('.//body')
    if body is not None:
        candidates.append(body)
    candidates.append(root)

    best_score, best_text = -9999, ''
    seen = set()
    for node in candidates:
        ident = id(node)
        if ident in seen:
            continue
        seen.add(ident)
        score, text = candidate_score(node, node_text(node))
        if score > best_score:
            best_score, best_text = score, text

    return best_text or clean_extracted_text(root.text_content())


def extract_url_text(url):
    req = urllib.request.Request(
        url,
        headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,text/plain,*/*',
        },
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        data = resp.read(2_000_000)
        content_type = resp.headers.get('Content-Type', '')
    page = decode_bytes(data, content_type)
    if 'text/plain' in content_type:
        return clean_extracted_text(page)
    return extract_from_html(page)


def extract_docx_text(filepath):
    with zipfile.ZipFile(filepath) as z:
        xml = z.read('word/document.xml').decode('utf-8', errors='replace')
    xml = re.sub(r'<w:tab[^>]*/?>', ' ', xml, flags=re.I)
    xml = re.sub(r'<w:br[^>]*/?>', '\n', xml, flags=re.I)
    xml = re.sub(r'</w:p>', '\n', xml, flags=re.I)
    text = re.sub(r'<[^>]+>', '', xml)
    return clean_extracted_text(text, max_chars=20000)


def extract_txt_text(filepath):
    data = Path(filepath).read_bytes()
    return clean_extracted_text(decode_bytes(data), max_chars=20000)


def extract_doc_text(filepath):
    data = Path(filepath).read_bytes()
    text = decode_bytes(data)
    chunks = re.split(r'[^\x20-\x7e\x80-\xff\r\n\t]{5,}', text)
    best = ''
    for chunk in chunks:
        clean = re.sub(r'[^\x20-\x7e\x80-\xff\n]', '', chunk.replace('\r', '\n')).strip()
        words = re.findall(r'[A-Za-zÀ-ÿ]{2,}', clean)
        if len(words) > 3 and len(clean) > len(best):
            best = clean
    return clean_extracted_text(best or text, max_chars=20000)


def decode_pdf_literal(raw):
    out = bytearray()
    i = 0
    while i < len(raw):
        ch = raw[i]
        if ch == 92 and i + 1 < len(raw):
            nxt = raw[i + 1]
            maps = {ord('n'): 10, ord('r'): 10, ord('t'): 9, ord('b'): 8, ord('f'): 12, ord('('): 40, ord(')'): 41, ord('\\'): 92}
            if nxt in maps:
                out.append(maps[nxt])
                i += 2
                continue
            if 48 <= nxt <= 55:
                m = re.match(rb'[0-7]{1,3}', raw[i + 1:i + 4])
                if m:
                    out.append(int(m.group(0), 8) & 0xFF)
                    i += 1 + len(m.group(0))
                    continue
        out.append(ch)
        i += 1
    return decode_bytes(bytes(out))


def extract_pdf_text(filepath):
    data = Path(filepath).read_bytes()
    pieces = []
    for stream_match in re.finditer(rb'stream\r?\n?(.*?)\r?\n?endstream', data, re.S):
        stream = stream_match.group(1).strip(b'\r\n')
        decoded = None
        for attempt in (stream, stream.strip()):
            try:
                decoded = zlib.decompress(attempt)
                break
            except Exception:
                pass
        if decoded is None:
            decoded = stream
        for obj in re.findall(rb'BT(.*?)ET', decoded, re.S):
            for token in re.finditer(rb'(\((?:\\.|[^\\)])*\)\s*(?:Tj|\'|"))|(\[(.*?)\]\s*TJ)|(\bT\*|\bTD|\bTd)', obj, re.S):
                if token.group(4):
                    pieces.append('\n')
                    continue
                if token.group(1):
                    m = re.match(rb'\((.*)\)\s*(?:Tj|\'|")', token.group(1), re.S)
                    if m:
                        pieces.append(decode_pdf_literal(m.group(1)))
                elif token.group(3):
                    line = ''.join(decode_pdf_literal(m.group(0)[1:-1]) for m in re.finditer(rb'\((?:\\.|[^\\)])*\)', token.group(3), re.S))
                    if line:
                        pieces.append(line)
            pieces.append('\n')
    text = ''.join(pieces)
    if len(text.strip()) < 10:
        # Last-resort fallback for simple PDFs.
        literals = [decode_pdf_literal(m.group(0)[1:-1]) for m in re.finditer(rb'\((?:\\.|[^\\)])*\)', data, re.S)]
        text = '\n'.join(literals)
    return clean_extracted_text(text, max_chars=20000)


def extract_file_text(filepath, ext):
    ext = (ext or Path(filepath).suffix.lstrip('.')).lower()
    if ext == 'docx':
        return extract_docx_text(filepath)
    if ext == 'doc':
        return extract_doc_text(filepath)
    if ext == 'pdf':
        return extract_pdf_text(filepath)
    if ext == 'txt':
        return extract_txt_text(filepath)
    return ''


def main():
    if len(sys.argv) < 3:
        print(json.dumps({'success': False, 'error': 'Argumentos insuficientes'}, ensure_ascii=False))
        return 1
    mode = sys.argv[1]
    try:
        if mode == 'url':
            text = extract_url_text(sys.argv[2])
        elif mode == 'file':
            ext = sys.argv[3] if len(sys.argv) > 3 else ''
            text = extract_file_text(sys.argv[2], ext)
        else:
            raise ValueError('Modo inválido')
        if not text.strip():
            raise ValueError('Nenhum texto útil encontrado')
        print(json.dumps({'success': True, 'text': text}, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(json.dumps({'success': False, 'error': str(exc)}, ensure_ascii=False))
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
