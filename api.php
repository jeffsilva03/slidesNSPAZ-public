<?php
error_reporting(0);
ini_set('display_errors', '0');

// Constantes de slide (calibrado: Arial Black 72pt)
define('CHARS_PER_LINE', 14);
define('MAX_LINES', 4);
define('HARD_MAX_LINES', 5);
define('SHORT_SLIDE_LINES', 2);
define('SAFE_MERGED_MAX_LINES', 5);
define('DOMINICAL_TEMPLATE_PATH', __DIR__ . '/assets/template.pptx');
define('MAOS_TEMPLATE_PATH', __DIR__ . '/assets/template_maos_ensanguentadas.pptx');
define('TEMPLATE_PATH', DOMINICAL_TEMPLATE_PATH);
define('PREFACIO_KEY', 'prefacio_eucaristico');

$UPLOAD_DIR = __DIR__ . '/uploads/';
$OUTPUT_DIR = __DIR__ . '/output/';
if (!is_dir(__DIR__ . '/assets/')) @mkdir(__DIR__ . '/assets/', 0755, true);
if (!is_dir($UPLOAD_DIR)) @mkdir($UPLOAD_DIR, 0755, true);
if (!is_dir($OUTPUT_DIR)) @mkdir($OUTPUT_DIR, 0755, true);

header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, GET, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') { http_response_code(200); exit; }

$BASE_LITURGICAL_SECTIONS = [
    'cantico_entrada'=>'Cântico de Entrada','ato_penitencial'=>'Ato Penitencial',
    'gloria'=>'Glória',
    'primeira_leitura'=>'Primeira Leitura',
    'salmo_responsorial'=>'Salmo Responsorial',
    'segunda_leitura'=>'Segunda Leitura',
    'aclamacao_evangelho'=>'Aclamação ao Evangelho',
    'evangelho'=>'Evangelho',
    'preces_oracao_dizimista'=>'Preces - Oração do Dizimista',
    'ofertorio'=>'Ofertório',
    'prefacio_eucaristico'=>'Prefácio Eucarístico',
    'santo'=>'Santo','cordeiro'=>'Cordeiro de Deus','comunhao'=>'Comunhão',
    'pos_comunhao'=>'Pós-Comunhão','cantico_final'=>'Cântico Final','outro'=>'Outro',
];

// Seções específicas da missa das Mãos Ensanguentadas
$MAOS_LITURGICAL_SECTIONS = [
    'cantico_entrada'=>'Cântico de Entrada',
    'aspercao'=>'Asperção',
    'ato_penitencial'=>'Ato Penitencial',
    'gloria'=>'Glória',
    'primeira_leitura'=>'Primeira Leitura',
    'salmo_responsorial'=>'Salmo Responsorial',
    'segunda_leitura'=>'Segunda Leitura',
    'aclamacao_evangelho'=>'Aclamação ao Evangelho',
    'evangelho'=>'Evangelho',
    'preces_oracao_dizimista'=>'Preces - Oração do Dizimista',
    'ofertorio'=>'Ofertório',
    'prefacio_eucaristico'=>'Prefácio Eucarístico',
    'santo'=>'Santo',
    'cordeiro'=>'Cordeiro de Deus',
    'comunhao'=>'Comunhão',
    'translado'=>'Translado',
    'pos_comunhao'=>'Pós-Comunhão',
    'cantico_final'=>'Cântico Final',
    'outro'=>'Outro',
];

$LITURGICAL_SECTION_SETS = [
    'dominical' => $BASE_LITURGICAL_SECTIONS,
    'maos' => $MAOS_LITURGICAL_SECTIONS,
];
$LITURGICAL_SECTIONS = $BASE_LITURGICAL_SECTIONS;

// Seções de rúbrica: texto literal, sem word-wrap nem processamento de letras
$RUBRIC_SECTIONS = ['primeira_leitura','segunda_leitura','evangelho'];
$READING_TEXT_SECTIONS = ['primeira_leitura','segunda_leitura','evangelho'];
$PSALM_SECTION = 'salmo_responsorial';
$IMAGE_EXTENSIONS = ['jpg','jpeg','png'];

function normalizeMassType($value) {
    return $value === 'maos' ? 'maos' : 'dominical';
}

function currentMassType() {
    return normalizeMassType($_GET['mass_type'] ?? $_POST['mass_type'] ?? 'dominical');
}

function getLiturgicalSections($massType) {
    global $LITURGICAL_SECTION_SETS;
    $massType = normalizeMassType($massType);
    return $LITURGICAL_SECTION_SETS[$massType] ?? $LITURGICAL_SECTION_SETS['dominical'];
}

function findBundledMaosTemplatePath() {
    $matches = glob(__DIR__ . '/SLIDE MODELO - M*OS ENSANGUENTADAS.pptx');
    if (!$matches) return null;
    foreach ($matches as $path) {
        if (basename($path)[0] !== '~' && file_exists($path)) return $path;
    }
    return null;
}

function findBundledTemplatePath($massType) {
    if (normalizeMassType($massType) === 'maos') {
        return findBundledMaosTemplatePath();
    }
    $path = __DIR__ . '/SLIDE MODELO - DOMINICAL.pptx';
    return file_exists($path) ? $path : null;
}

function getTemplatePath($massType) {
    return normalizeMassType($massType) === 'maos' ? MAOS_TEMPLATE_PATH : DOMINICAL_TEMPLATE_PATH;
}

function getEffectiveTemplatePath($massType) {
    $path = getTemplatePath($massType);
    $bundled = findBundledTemplatePath($massType);
    if (file_exists($path) && $bundled && file_exists($bundled)) {
        return filemtime($bundled) > filemtime($path) ? $bundled : $path;
    }
    if (file_exists($path)) return $path;
    return ($bundled && file_exists($bundled)) ? $bundled : null;
}

function getTemplateStatus($massType) {
    $path = getTemplatePath($massType);
    $uploaded = file_exists($path);
    $effective = getEffectiveTemplatePath($massType);
    $source = $effective === null ? 'none' : ($effective === $path ? 'uploaded' : 'bundled');
    return [
        'success' => true,
        'exists' => $effective !== null,
        'uploaded' => $uploaded,
        'source' => $source,
    ];
}

$action = $_GET['action'] ?? $_POST['action'] ?? '';
switch ($action) {
    case 'sections': echo json_encode(['success'=>true,'sections'=>getLiturgicalSections(currentMassType())],JSON_UNESCAPED_UNICODE); break;
    case 'generate': handleGenerate(); break;
    case 'download': handleDownload(); break;
    case 'upload_template': handleUploadTemplate(); break;
    case 'has_template': echo json_encode(getTemplateStatus(currentMassType()),JSON_UNESCAPED_UNICODE); break;
    case 'delete_template': @unlink(getTemplatePath(currentMassType())); echo json_encode(['success'=>true],JSON_UNESCAPED_UNICODE); break;
    case 'debug': handleDebug(); break;
    case 'debug_template': handleDebugTemplate(); break;
    default: echo json_encode(['success'=>false,'error'=>'Ação inválida'],JSON_UNESCAPED_UNICODE);
}

function handleUploadTemplate() {
    $uploadErrors = [
        UPLOAD_ERR_INI_SIZE   => 'Arquivo maior que upload_max_filesize no php.ini',
        UPLOAD_ERR_FORM_SIZE  => 'Arquivo maior que MAX_FILE_SIZE no formulário',
        UPLOAD_ERR_PARTIAL    => 'Upload incompleto',
        UPLOAD_ERR_NO_FILE    => 'Nenhum arquivo enviado',
        UPLOAD_ERR_NO_TMP_DIR => 'Pasta temporária não encontrada',
        UPLOAD_ERR_CANT_WRITE => 'Falha ao escrever no disco',
        UPLOAD_ERR_EXTENSION  => 'Upload bloqueado por extensão PHP',
    ];
    if (!isset($_FILES['template'])) {
        echo json_encode(['success'=>false,'error'=>'Campo template não recebido pelo servidor'],JSON_UNESCAPED_UNICODE); return;
    }
    $errCode = $_FILES['template']['error'];
    if ($errCode !== UPLOAD_ERR_OK) {
        $msg = $uploadErrors[$errCode] ?? "Erro de upload código $errCode";
        echo json_encode(['success'=>false,'error'=>$msg],JSON_UNESCAPED_UNICODE); return;
    }
    $file = $_FILES['template'];
    $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
    if ($ext !== 'pptx') {
        echo json_encode(['success'=>false,'error'=>'Apenas arquivos .pptx são aceitos'],JSON_UNESCAPED_UNICODE); return;
    }
    // Valida que é um ZIP válido
    $zip = new ZipArchive();
    $zipResult = $zip->open($file['tmp_name']);
    if ($zipResult !== true) {
        echo json_encode(['success'=>false,'error'=>'Arquivo .pptx corrompido ou inválido (código ZIP: '.$zipResult.')'],JSON_UNESCAPED_UNICODE); return;
    }
    $zip->close();
    $dest = getTemplatePath(currentMassType());
    if (!move_uploaded_file($file['tmp_name'], $dest)) {
        echo json_encode(['success'=>false,'error'=>'move_uploaded_file falhou — destino: '.$dest.' | gravável: '.(is_writable(dirname($dest))?'sim':'não')],JSON_UNESCAPED_UNICODE); return;
    }
    echo json_encode(['success'=>true],JSON_UNESCAPED_UNICODE);
}

/* ================================================================
   DEBUG TEMPLATE
   ================================================================ */
function handleDebugTemplate() {
    $templatePath = getEffectiveTemplatePath(currentMassType());
    if (!$templatePath) { echo json_encode(['error'=>'sem template']); return; }
    $zip = new ZipArchive();
    $zip->open($templatePath);
    $rels  = $zip->getFromName('ppt/_rels/presentation.xml.rels');
    $ct    = $zip->getFromName('[Content_Types].xml');
    $slide1rels = $zip->getFromName('ppt/slides/_rels/slide1.xml.rels');
    $zip->close();

    // Testa os regexes
    $relsClean = preg_replace('/<Relationship\s[^>]*relationships\/slide"[^>]*?\/>/is', '', $rels);
    $relsClean = preg_replace('/<Relationship\s[^>]*relationships\/slide"[^>]*?>.*?<\/Relationship>/is', '', $relsClean);
    $ctClean = preg_replace('/<Override\s[^>]*presentationml\.slide\+xml"[^>]*?\/>/is', '', $ct);

    $slide1xml = $zip->getFromName('ppt/slides/slide1.xml');
    $presXml   = $zip->getFromName('ppt/presentation.xml');

    header('Content-Type: application/json; charset=utf-8');
    echo json_encode([
        'slide1_rels'   => $slide1rels,
        'slide1_xml'    => $slide1xml,
        'presentation_rels_after_clean' => $relsClean,
        'content_types_after_clean'     => substr($ctClean, 0, 3000),
        'pres_xml_snippet'              => substr($presXml, 0, 2000),
    ], JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
}

/* ================================================================
   DEBUG
   ================================================================ */
function handleDebug() {
    global $OUTPUT_DIR, $UPLOAD_DIR;
    echo json_encode([
        'php_version'=>phpversion(),
        'zip'=>class_exists('ZipArchive'),
        'mbstring'=>extension_loaded('mbstring'),
        'output_dir'=>is_dir($GLOBALS['OUTPUT_DIR'])&&is_writable($GLOBALS['OUTPUT_DIR']),
        'upload_dir'=>is_dir($GLOBALS['UPLOAD_DIR'])&&is_writable($GLOBALS['UPLOAD_DIR']),
        'chars_per_line'=>CHARS_PER_LINE,
        'max_lines'=>MAX_LINES,
    ], JSON_PRETTY_PRINT|JSON_UNESCAPED_UNICODE);
}

/* ================================================================
   GENERATE
   ================================================================ */
function handleGenerate() {
    global $UPLOAD_DIR, $OUTPUT_DIR, $IMAGE_EXTENSIONS;
    try {
        $massType = currentMassType();
        $liturgicalSections = getLiturgicalSections($massType);
        $sections = [];
        $errors = [];
        $blockingErrors = [];
        $cleanupFiles = [];

        foreach ($liturgicalSections as $key => $label) {
            $inputType = $_POST["input_type_{$key}"] ?? '';
            global $READING_TEXT_SECTIONS, $PSALM_SECTION;
            $readingText = '';
            if (in_array($key, $READING_TEXT_SECTIONS, true) || $key === $PSALM_SECTION) {
                $readingText = trim($_POST["reading_text_{$key}"] ?? '');
            }
            if ((empty($inputType) || $inputType === 'none') && $readingText === '') continue;
            $text = '';

            if ($inputType === 'url') {
                $url = trim($_POST["url_{$key}"] ?? '');
                if (empty($url) && $readingText === '') continue;
                if (empty($url)) {
                    $text = '';
                } else {
                    $text = extractFromUrl($url);
                    if (trim($text) === '' || stripos(trim($text), 'ERRO:') === 0) {
                    $errors[] = "{$label}: não foi possível extrair texto útil do link";
                        $text = '';
                        if ($readingText === '') continue;
                    }
                }
            } elseif ($inputType === 'file') {
                if (!isset($_FILES["file_{$key}"]) || $_FILES["file_{$key}"]['error'] !== UPLOAD_ERR_OK) continue;
                $file = $_FILES["file_{$key}"];
                $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
                if (!in_array($ext, ['doc','docx','pdf','txt'])) { $errors[] = "{$label}: .{$ext} não suportado"; continue; }
                $filename = uniqid("{$key}_") . ".{$ext}";
                $filepath = $UPLOAD_DIR . $filename;
                if (!move_uploaded_file($file['tmp_name'], $filepath)) { $errors[] = "{$label}: erro ao salvar"; continue; }
                $text = extractFromFile($filepath, $ext);
                @unlink($filepath);
                if (trim($text) === '' || stripos(trim($text), 'ERRO:') === 0) {
                    $errors[] = "{$label}: não foi possível extrair texto útil do arquivo";
                    continue;
                }
            } elseif ($inputType === 'text') {
                $text = trim($_POST["text_{$key}"] ?? '');
            }
            if (!empty($text) || $readingText !== '') {
                $item = ['key'=>$key,'title'=>$label,'text'=>$text];
                if ($readingText !== '') $item['reading_text'] = $readingText;
                $sections[] = $item;
            }
        }

        $customIds = $_POST['custom_section_ids'] ?? [];
        if (!is_array($customIds)) $customIds = [$customIds];
        foreach ($customIds as $rawId) {
            $customId = preg_replace('/[^a-zA-Z0-9_\-]/', '', (string)$rawId);
            if ($customId === '') continue;

            $title = trim($_POST["custom_section_title_{$customId}"] ?? '');
            $afterKey = trim($_POST["custom_section_after_{$customId}"] ?? 'outro');
            if ($afterKey === '') $afterKey = 'outro';
            $text = trim($_POST["custom_section_text_{$customId}"] ?? '');
            $fileKey = "custom_section_cover_{$customId}";
            $hasCover = isset($_FILES[$fileKey]) && $_FILES[$fileKey]['error'] === UPLOAD_ERR_OK && !empty($_FILES[$fileKey]['name']);

            if ($title === '' && $text === '' && !$hasCover) continue;
            if ($title === '' || $text === '' || !$hasCover) {
                $blockingErrors[] = 'Seção extra incompleta: preencha nome, capa e texto.';
                continue;
            }

            $file = $_FILES[$fileKey];
            $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
            if (!in_array($ext, $IMAGE_EXTENSIONS, true)) {
                $blockingErrors[] = "{$title}: capa .{$ext} não suportada. Use .jpg, .jpeg ou .png.";
                continue;
            }

            $filepath = $UPLOAD_DIR . uniqid("custom_section_{$customId}_") . ".{$ext}";
            if (!move_uploaded_file($file['tmp_name'], $filepath)) {
                $blockingErrors[] = "{$title}: erro ao salvar a capa.";
                continue;
            }
            $cleanupFiles[] = $filepath;
            $sections[] = [
                'key'=>"custom_section_{$customId}",
                'title'=>$title,
                'text'=>$text,
                'custom_section'=>true,
                'insert_after_key'=>$afterKey,
                'cover_image'=>$filepath,
            ];
        }

        if (!empty($blockingErrors)) {
            foreach ($cleanupFiles as $fp) @unlink($fp);
            echo json_encode(['success'=>false,'error'=>'Corrija as seções extras antes de gerar.','errors'=>array_merge($errors,$blockingErrors)],JSON_UNESCAPED_UNICODE);
            return;
        }

        if (empty($sections)) {
            echo json_encode(['success'=>false,'error'=>'Nenhuma seção com conteúdo.','errors'=>$errors],JSON_UNESCAPED_UNICODE);
            return;
        }

        $customName = trim($_POST['custom_filename'] ?? '');
        if (!empty($customName)) {
            $safeName = preg_replace('/[^a-zA-Z0-9_\-àáâãéêíóôõúüçÀÁÂÃÉÊÍÓÔÕÚÜÇ ]/u', '', $customName);
            $safeName = preg_replace('/\s+/', '_', trim($safeName));
            if (empty($safeName)) $safeName = 'liturgia';
        } else {
            $safeName = 'liturgia_' . date('Y-m-d_His');
        }
        $outputFilename = $safeName . '.pptx';
        $counter = 1;
        while (file_exists($OUTPUT_DIR . $outputFilename)) {
            $outputFilename = $safeName . '_' . $counter . '.pptx';
            $counter++;
        }

        try {
            $ok = generatePptxPure($OUTPUT_DIR . $outputFilename, $sections, $massType);
        } finally {
            foreach ($cleanupFiles as $fp) @unlink($fp);
        }
        if ($ok) {
            echo json_encode(['success'=>true,'filename'=>$outputFilename,'sections_count'=>count($sections),'errors'=>$errors],JSON_UNESCAPED_UNICODE);
        } else {
            echo json_encode(['success'=>false,'error'=>'Erro ao gerar PPTX','errors'=>$errors],JSON_UNESCAPED_UNICODE);
        }
    } catch (\Throwable $e) {
        echo json_encode(['success'=>false,'error'=>'Erro PHP: '.$e->getMessage().' (linha '.$e->getLine().')'],JSON_UNESCAPED_UNICODE);
    }
}

/* ================================================================
   TEXT EXTRACTION (tradução do extract_text.py)
   ================================================================ */
function extractViaPython($mode, $value, $ext = '') {
    $scriptPath = __DIR__ . '/extract_text.py';
    if (!file_exists($scriptPath)) return null;

    $pythonCandidates = ['C:/Python314/python.exe', 'python'];
    $pythonBin = null;
    foreach ($pythonCandidates as $candidate) {
        if ($candidate === 'python' || file_exists($candidate)) {
            $pythonBin = $candidate;
            break;
        }
    }
    if ($pythonBin === null) return null;

    $cmd = escapeshellarg($pythonBin)
         . ' ' . escapeshellarg($scriptPath)
         . ' ' . escapeshellarg($mode)
         . ' ' . escapeshellarg($value);
    if ($mode === 'file') $cmd .= ' ' . escapeshellarg($ext);
    $cmd .= ' 2>&1';

    $output = @shell_exec($cmd);
    if ($output === null || trim($output) === '') return null;
    $result = @json_decode(trim($output), true);
    if (!is_array($result) || empty($result['success']) || empty($result['text'])) return null;
    return $result['text'];
}

function extractFromFile($filepath, $ext) {
    $viaPython = extractViaPython('file', $filepath, $ext);
    if ($viaPython !== null && trim($viaPython) !== '') return $viaPython;

    switch ($ext) {
        case 'docx': return extractDocx($filepath);
        case 'doc':  return extractDoc($filepath);
        case 'pdf':  return extractPdf($filepath);
        case 'txt':  return extractTxt($filepath);
        default: return '';
    }
}

function extractDocx($filepath) {
    $zip = new ZipArchive();
    if ($zip->open($filepath) !== true) return 'ERRO: Não foi possível abrir o DOCX';
    $xml = $zip->getFromName('word/document.xml');
    $zip->close();
    if (!$xml) return 'ERRO: document.xml não encontrado';
    $xml = preg_replace('/<w:br[^>]*\/?>/i', "\n", $xml);
    $xml = preg_replace('/<\/w:p>/i', "\n", $xml);
    $text = strip_tags($xml);
    $text = html_entity_decode($text, ENT_QUOTES, 'UTF-8');
    $lines = [];
    foreach (explode("\n", $text) as $line) {
        $line = trim($line);
        if ($line !== '') $lines[] = $line;
    }
    return limparTexto(implode("\n", $lines));
}

function extractDoc($filepath) {
    $content = @file_get_contents($filepath);
    if ($content === false) return 'ERRO: Não foi possível ler o arquivo .doc';
    $text = @mb_convert_encoding($content, 'UTF-8', 'Windows-1252');
    if ($text === false) $text = $content;
    $chunks = preg_split('/[^\x20-\x7e\x80-\xff\r\n\t]{5,}/s', $text);
    $best = '';
    foreach ($chunks as $chunk) {
        $clean = str_replace("\r", "\n", $chunk);
        $clean = preg_replace('/[^\x20-\x7e\x80-\xff\n]/', '', $clean);
        $clean = trim($clean);
        preg_match_all('/[a-zA-Z\x{00C0}-\x{00FF}]{2,}/u', $clean, $matches);
        if (count($matches[0]) > 3 && mb_strlen($clean) > mb_strlen($best)) {
            $best = $clean;
        }
    }
    $best = preg_replace('/\n{3,}/', "\n\n", $best);
    return limparTexto(trim($best) ?: 'ERRO: Não foi possível extrair texto do .doc');
}

function extractPdf($filepath) {
    $content = @file_get_contents($filepath);
    if ($content === false) return 'ERRO: Não foi possível ler o PDF';
    $text = '';
    preg_match_all('/stream\s*(.*?)\s*endstream/s', $content, $streams);
    foreach ($streams[1] as $stream) {
        $decoded = @gzuncompress($stream);
        if ($decoded === false) $decoded = $stream;
        preg_match_all('/\((.*?)\)/s', $decoded, $texts);
        foreach ($texts[1] as $t) {
            $t = str_replace(['\\(','\\)','\\\\'], ['(',')',  '\\'], $t);
            $text .= $t;
        }
        preg_match_all('/\[(.*?)\]\s*TJ/s', $decoded, $tjTexts);
        foreach ($tjTexts[1] as $tjContent) {
            preg_match_all('/\((.*?)\)/', $tjContent, $innerTexts);
            foreach ($innerTexts[1] as $t) {
                $t = str_replace(['\\(','\\)','\\\\'], ['(',')',  '\\'], $t);
                $text .= $t;
            }
        }
    }
    if (empty(trim($text))) return 'ERRO: Não foi possível extrair texto do PDF';
    $lines = [];
    foreach (explode("\n", $text) as $line) { $line = trim($line); if ($line !== '') $lines[] = $line; }
    return limparTexto(implode("\n", $lines));
}

function extractTxt($filepath) {
    foreach (['UTF-8','ISO-8859-1','Windows-1252'] as $enc) {
        $content = @file_get_contents($filepath);
        if ($content === false) continue;
        $converted = @mb_convert_encoding($content, 'UTF-8', $enc);
        if ($converted !== false) return limparTexto(trim($converted));
    }
    return 'ERRO: Não foi possível ler o arquivo TXT';
}

function extractFromUrl($url) {
    $viaPython = extractViaPython('url', $url);
    if ($viaPython !== null && trim($viaPython) !== '') return $viaPython;

    $ctx = @stream_context_create([
        'http'=>['header'=>"User-Agent: Mozilla/5.0\r\n",'timeout'=>15],
        'ssl'=>['verify_peer'=>false,'verify_peer_name'=>false],
    ]);
    $html = @file_get_contents($url, false, $ctx);
    if ($html === false) return 'ERRO: Não foi possível acessar a URL';
    foreach (['script','style','nav','header','footer','aside','form','noscript','iframe'] as $tag) {
        $html = preg_replace("/<{$tag}[^>]*>.*?<\/{$tag}>/si", '', $html);
    }
    $html = preg_replace('/<!--.*?-->/s', '', $html);
    $main = '';
    if (preg_match('/<(?:article|main)[^>]*>(.*?)<\/(?:article|main)>/si', $html, $m)) $main = $m[1];
    if (empty($main)) {
        foreach (['/class="[^"]*(?:lyrics|letra|content|verse|cifra)[^"]*"[^>]*>(.*?)<\/div>/si',
                  '/class="[^"]*(?:entry-content|post-content|article-body)[^"]*"[^>]*>(.*?)<\/div>/si'] as $p) {
            if (preg_match($p, $html, $m)) { $main = $m[1]; break; }
        }
    }
    if (empty($main)) { if (preg_match('/<body[^>]*>(.*?)<\/body>/si', $html, $m)) $main = $m[1]; else $main = $html; }
    $text = preg_replace('/<br\s*\/?>/i', "\n", $main);
    $text = preg_replace('/<\/p>/i', "\n", $text);
    $text = preg_replace('/<\/div>/i', "\n", $text);
    $text = strip_tags($text);
    $text = html_entity_decode($text, ENT_QUOTES, 'UTF-8');
    $lines = [];
    foreach (explode("\n", $text) as $line) {
        $cleaned = trim(preg_replace('/\s+/', ' ', $line));
        if (mb_strlen($cleaned) <= 1) continue;
        if (preg_match('/^(home|voltar|menu|buscar|login|sair|cadastrar|compartilhar|imprimir)$/i', $cleaned)) continue;
        if (preg_match('/^(tom|capo|cifra|acordes?)[\s:.]/i', $cleaned)) continue;
        if (preg_match('/(copyright|todos os direitos|cookie|privacidade)/i', $cleaned)) continue;
        if (preg_match('/^https?:\/\//', $cleaned)) continue;
        $lines[] = $cleaned;
    }
    $result = implode("\n", $lines);
    if (mb_strlen($result) > 3000) {
        $short = []; $c = 0;
        foreach ($lines as $l) { $short[] = $l; $c += mb_strlen($l); if ($c > 2000) break; }
        $result = implode("\n", $short);
    }
    return limparTexto($result);
}

function limparTexto($text) {
    $text = preg_replace('/[\x{10000}-\x{10FFFF}]/u', '', $text);
    $text = preg_replace('/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/', '', $text);
    $lines = [];
    foreach (explode("\n", $text) as $line) {
        $line = trim($line);
        if ($line === '') { $lines[] = ''; continue; }
        if (substr_count($line, '|') > 2) continue;
        if (substr_count($line, '/') > 4 && mb_strlen($line) < 100) continue;
        if (preg_match('/^https?:\/\//', $line)) continue;
        if (preg_match('/(copyright|todos os direitos|cookie|privacidade)/i', $line)) continue;
        if (preg_match('/^[\d\s\.\,\-]+$/', $line) && mb_strlen($line) < 15) continue;
        $lines[] = $line;
    }
    return implode("\n", $lines);
}

/* ================================================================
   SLIDE TEXT PROCESSING (tradução exata do JS)
   ================================================================ */
function fixPunctuation($text) {
    $text = preg_replace('/([.!?;:,])([A-Za-z\x{00C0}-\x{024F}])/u', '$1 $2', $text);
    $text = preg_replace('/([A-Za-z\x{00C0}-\x{024F}.!?])(["\'\)\]])([A-Z\x{00C0}-\x{00DC}])/u', '$1$2 $3', $text);
    return $text;
}

function capitalizeSacredWords($text) {
    $mainWords = ['deus','senhor','jesus','cristo','espírito','espirito',
        'salvador','redentor','messias','altíssimo','altissimo',
        'onipotente','maria','virgem','evangelho',
        'aleluia','amém','hosana','glória','gloria'];
    foreach ($mainWords as $w) {
        $text = preg_replace_callback('/\b' . preg_quote($w, '/') . '\b/iu', function($m) {
            return mb_strtoupper(mb_substr($m[0], 0, 1)) . mb_substr($m[0], 1);
        }, $text);
    }
    $text = preg_replace_callback('/\bnossa\s+senhora\b/iu', function($m) { return 'Nossa Senhora'; }, $text);
    $text = preg_replace_callback('/\besp[ií]rito\s+santo\b/iu', function($m) { return 'Espírito Santo'; }, $text);
    $text = preg_replace_callback('/\bjesus\s+cristo\b/iu', function($m) { return 'Jesus Cristo'; }, $text);
    $text = preg_replace_callback('/\bmãe\s+de\s+deus\b/iu', function($m) { return 'Mãe de Deus'; }, $text);
    $text = preg_replace_callback('/\bfilho\s+de\s+deus\b/iu', function($m) { return 'Filho de Deus'; }, $text);
    return $text;
}

function capitalize($str) {
    if (empty($str)) return $str;
    return mb_strtoupper(mb_substr($str, 0, 1)) . mb_substr($str, 1);
}

function wordWrap_custom($text) {
    $words = preg_split('/\s+/', $text);
    $words = array_values(array_filter($words, function($w) { return mb_strlen($w) > 0; }));
    if (empty($words)) return [];
    $lines = [];
    $cur = '';
    foreach ($words as $word) {
        $test = ($cur !== '') ? $cur . ' ' . $word : $word;
        if (mb_strlen($test) > CHARS_PER_LINE && mb_strlen($cur) > 0) {
            $lines[] = $cur;
            $cur = $word;
        } else {
            $cur = $test;
        }
    }
    if ($cur !== '') $lines[] = $cur;
    return $lines;
}

function slideLinesCount($slide) {
    $lines = preg_split('/\n/', $slide);
    $count = 0;
    foreach ($lines as $line) {
        if (trim($line) !== '') $count++;
    }
    return $count;
}

function mergeShortSlides($slides) {
    $merged = [];
    $i = 0;
    while ($i < count($slides)) {
        $current = $slides[$i];
        if ($i + 1 < count($slides)) {
            $currentCount = slideLinesCount($current);
            $nextCount = slideLinesCount($slides[$i + 1]);
            if (
                $currentCount > 0
                && $currentCount <= SHORT_SLIDE_LINES
                && ($currentCount + $nextCount) <= SAFE_MERGED_MAX_LINES
            ) {
                $merged[] = $current . "\n" . $slides[$i + 1];
                $i += 2;
                continue;
            }
        }
        $merged[] = $current;
        $i++;
    }
    return $merged;
}

function splitPhraseUnits($line) {
    $line = trim($line);
    if ($line === '') return [];
    $units = preg_split('/(?<=[.!?;:])\s+/u', $line, -1, PREG_SPLIT_NO_EMPTY);
    if (empty($units)) $units = [$line];

    $expanded = [];
    foreach ($units as $unit) {
        $unit = trim($unit);
        if ($unit === '') continue;
        $wrapped = wordWrap_custom($unit);
        if (count($wrapped) > HARD_MAX_LINES) {
            $commaUnits = preg_split('/(?<=,)\s+/u', $unit, -1, PREG_SPLIT_NO_EMPTY);
            if (count($commaUnits) > 1) {
                foreach ($commaUnits as $commaUnit) {
                    $commaUnit = trim($commaUnit);
                    if ($commaUnit !== '') $expanded[] = $commaUnit;
                }
                continue;
            }
        }
        $expanded[] = $unit;
    }
    return $expanded;
}

function flushSlide(&$slides, &$curSlide, &$curCount) {
    if (!empty($curSlide)) {
        $slides[] = implode("\n", $curSlide);
        $curSlide = [];
        $curCount = 0;
    }
}

function appendWrappedUnit(&$paraSlides, &$curSlide, &$curCount, $wrapped) {
    while (count($wrapped) > HARD_MAX_LINES) {
        flushSlide($paraSlides, $curSlide, $curCount);
        $paraSlides[] = implode("\n", array_slice($wrapped, 0, HARD_MAX_LINES));
        $wrapped = array_slice($wrapped, HARD_MAX_LINES);
    }
    if (empty($wrapped)) return;

    $lineCount = count($wrapped);
    if ($curCount > 0 && ($curCount + $lineCount) > HARD_MAX_LINES) {
        flushSlide($paraSlides, $curSlide, $curCount);
    }
    foreach ($wrapped as $line) $curSlide[] = $line;
    $curCount += $lineCount;
}

function splitIntoSlides($text) {
    // Cada bloco separado por linha em branco = um verso
    $paragraphs = [];
    $cur = [];
    foreach (explode("\n", $text) as $line) {
        $t = trim($line);
        if ($t === '') {
            if (!empty($cur)) { $paragraphs[] = $cur; $cur = []; }
        } else {
            $cur[] = $t;
        }
    }
    if (!empty($cur)) $paragraphs[] = $cur;

    $slides = [];
    foreach ($paragraphs as $para) {
        $paraSlides = [];
        $curSlide = [];
        $curCount = 0;
        foreach ($para as $line) {
            foreach (splitPhraseUnits($line) as $unit) {
                appendWrappedUnit($paraSlides, $curSlide, $curCount, wordWrap_custom($unit));
            }
        }
        flushSlide($paraSlides, $curSlide, $curCount);
        $slides = array_merge($slides, mergeShortSlides($paraSlides));
    }

    // Post-process: nunca deixar uma palavra sozinha num slide
    $final = [];
    for ($i = 0; $i < count($slides); $i++) {
        $slideWords = preg_split('/\s+/', trim($slides[$i]), -1, PREG_SPLIT_NO_EMPTY);
        if (count($slideWords) <= 1) {
            $currentLines = slideLinesCount($slides[$i]);
            if (!empty($final) && (slideLinesCount($final[count($final)-1]) + $currentLines) <= HARD_MAX_LINES) {
                $final[count($final)-1] .= "\n" . $slides[$i];
            } elseif ($i + 1 < count($slides) && ($currentLines + slideLinesCount($slides[$i + 1])) <= HARD_MAX_LINES) {
                $slides[$i+1] = $slides[$i] . "\n" . $slides[$i+1];
            } else {
                $final[] = $slides[$i];
            }
        } else {
            $final[] = $slides[$i];
        }
    }
    return $final;
}

function preprocessLyricText($text) {
    $text = limparTexto($text);
    $text = fixPunctuation($text);
    return capitalizeSacredWords($text);
}

function splitTextBlocks($text) {
    $blocks = [];
    $cur = [];
    foreach (explode("\n", $text) as $line) {
        if (trim($line) === '') {
            if (!empty($cur)) {
                $blocks[] = trim(implode("\n", $cur));
                $cur = [];
            }
        } else {
            $cur[] = trim($line);
        }
    }
    if (!empty($cur)) $blocks[] = trim(implode("\n", $cur));
    return array_values(array_filter($blocks, function($block) { return $block !== ''; }));
}

function buildProjectedTextSlides($text) {
    $text = preprocessLyricText($text);
    $slides = [];
    foreach (splitIntoSlides($text) as $part) {
        if (trim($part) !== '') $slides[] = ['text'=>capitalize($part), 'bold'=>false];
    }
    return $slides;
}

function splitPsalmStanzaBlocks($text) {
    $text = trim(str_replace(["\r\n", "\r"], "\n", $text));
    if ($text === '') return [];

    $blankBlocks = preg_split('/\n\s*\n+/u', $text, -1, PREG_SPLIT_NO_EMPTY);
    $blankBlocks = array_values(array_filter(array_map('trim', $blankBlocks), function($block) {
        return $block !== '';
    }));
    if (count($blankBlocks) > 1) return $blankBlocks;

    $lines = array_values(array_filter(array_map('trim', explode("\n", $text)), function($line) {
        return $line !== '';
    }));
    $blocks = [];
    $cur = [];
    foreach ($lines as $line) {
        if (preg_match('/^[-–—]\s+/u', $line) && !empty($cur)) {
            $blocks[] = trim(implode("\n", $cur));
            $cur = [$line];
        } else {
            $cur[] = $line;
        }
    }
    if (!empty($cur)) $blocks[] = trim(implode("\n", $cur));
    return array_values(array_filter($blocks, function($block) { return $block !== ''; }));
}

function buildPsalmStanzaSlides($text) {
    $text = preprocessLyricText($text);
    $stanzas = [];
    foreach (splitPsalmStanzaBlocks($text) as $block) {
        $slides = [];
        foreach (splitIntoSlides($block) as $part) {
            if (trim($part) !== '') $slides[] = ['text'=>capitalize($part), 'bold'=>false];
        }
        if (!empty($slides)) $stanzas[] = $slides;
    }
    return $stanzas;
}

/* ================================================================
   TEMPLATE EXTRACTION
   ================================================================ */
function extractTemplateData($templatePath) {
    $zip = new ZipArchive();
    if ($zip->open($templatePath) !== true) return null;

    $data = ['files' => [], 'cx' => 9144000, 'cy' => 6858000, 'layoutTarget' => '../slideLayouts/slideLayout1.xml'];

    // Slide size from presentation.xml
    $presXml = $zip->getFromName('ppt/presentation.xml');
    if ($presXml && preg_match('/sldSz\s[^>]*cx="(\d+)"[^>]*cy="(\d+)"/', $presXml, $m)) {
        $data['cx'] = (int)$m[1]; $data['cy'] = (int)$m[2];
    }

    // Layout used by first slide
    $slide1Rels = $zip->getFromName('ppt/slides/_rels/slide1.xml.rels');
    if ($slide1Rels && preg_match('/slideLayout[^"]*"\s+Target="([^"]+)"/', $slide1Rels, $m)) {
        $data['layoutTarget'] = $m[1];
    }

    // Copia apenas o necessário: masters, layouts, temas, mídia, fontes, props
    for ($i = 0; $i < $zip->numFiles; $i++) {
        $name = $zip->getNameIndex($i);
        if (preg_match('#^ppt/slides/#', $name)) continue;
        if (preg_match('#^ppt/notesSlides/#', $name)) continue;
        if (preg_match('#^ppt/notesMasters/#', $name)) continue;
        if ($name === 'metadata' || preg_match('#^metadata/#', $name)) continue;
        $data['files'][$name] = $zip->getFromIndex($i);
    }
    $zip->close();
    return $data;
}

/* ================================================================
   PPTX GENERATION (PHP puro - ZipArchive + XML)
   ================================================================ */
function generatePptxPure($outputPath, $sections, $massType = 'dominical') {
    // Monta formato de grupos por seção para o Python
    $sectionGroups = [];
    foreach ($sections as $sec) {
        $key   = $sec['key']   ?? 'outro';
        $title = $sec['title'] ?? '';
        $text  = trim($sec['text'] ?? '');
        $readingText = trim($sec['reading_text'] ?? '');

        $slides = [];
        $prefacioFinalText = '';

        if (!empty($text)) {
            global $RUBRIC_SECTIONS, $PSALM_SECTION;
            if (in_array($key, $RUBRIC_SECTIONS)) {
                // Rúbrica: passa o texto literal sem processamento
                $slides[] = ['text' => $text, 'bold' => false];
            } elseif ($key === $PSALM_SECTION && $readingText !== '') {
                $text = preprocessLyricText($text);
                $slides[] = ['text' => capitalize($text), 'bold' => false];
            } else {
                $text = preprocessLyricText($text);
                $parts = [];
                if ($key === PREFACIO_KEY) {
                    $blocks = splitTextBlocks($text);
                    if (count($blocks) > 1) {
                        $parts = splitIntoSlides(implode("\n\n", array_slice($blocks, 0, -1)));
                        $finalParts = splitIntoSlides($blocks[count($blocks) - 1]);
                        if (!empty($finalParts)) {
                            $prefacioFinalText = capitalize(array_pop($finalParts));
                            $parts = array_merge($parts, $finalParts);
                        }
                    } else {
                        $parts = splitIntoSlides($text);
                        if (!empty($parts)) $prefacioFinalText = capitalize(array_pop($parts));
                    }
                } else {
                    $parts = splitIntoSlides($text);
                }
                foreach ($parts as $part) {
                    $slides[] = ['text' => capitalize($part), 'bold' => false];
                }
            }
        }

        $group = ['section_key' => $key, 'slides' => $slides];
        global $READING_TEXT_SECTIONS, $PSALM_SECTION;
        if (in_array($key, $READING_TEXT_SECTIONS, true) && $readingText !== '') {
            $group['reading_slides'] = buildProjectedTextSlides($readingText);
        } elseif ($key === $PSALM_SECTION && $readingText !== '') {
            $group['psalm_stanzas'] = buildPsalmStanzaSlides($readingText);
        }
        if ($prefacioFinalText !== '') $group['prefacio_final_text'] = $prefacioFinalText;
        if (!empty($sec['custom_section'])) {
            $group['custom_section'] = true;
            $group['title'] = trim($sec['title'] ?? '') ?: 'Seção extra';
            $group['insert_after_key'] = $sec['insert_after_key'] ?? 'outro';
            $group['cover_image'] = $sec['cover_image'] ?? '';
        }
        $sectionGroups[] = $group;
    }

    if (empty($sectionGroups)) return false;
    return generateViaPython($outputPath, $sectionGroups, $massType);

    // código legado removido

    $zip->addFromString('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
    $zip->addFromString('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Liturgia Slides</Application><Slides>'.$sc.'</Slides></Properties>');
    $zip->addFromString('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Liturgia</dc:title><dc:creator>Liturgia Slides</dc:creator></cp:coreProperties>');
    $zip->addFromString('ppt/presProps.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>');
    $zip->addFromString('ppt/tableStyles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>');
    $zip->addFromString('ppt/viewProps.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr></p:viewPr>');
    $zip->addFromString('ppt/theme/theme1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>');
    $zip->addFromString('ppt/slideLayouts/slideLayout1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1"><p:cSld name="DEFAULT"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>');
    $zip->addFromString('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
    $zip->addFromString('ppt/slideMasters/slideMaster1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>');
    $zip->addFromString('ppt/slideMasters/_rels/slideMaster1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>');

    // Slides
    for ($i = 0; $i < $sc; $i++) {
        $n = $i + 1;
        $s = $allSlides[$i];
        $zip->addFromString("ppt/slides/slide{$n}.xml", buildSlideXml($s['text'], $s['bold']));
        $zip->addFromString("ppt/slides/_rels/slide{$n}.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
    }

    // Presentation rels
    $sldIds = '';
    $rels = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>';
    $overrides = '';
    for ($i = 0; $i < $sc; $i++) {
        $rid = 'rId' . ($i + 2);
        $n = $i + 1;
        $sldIds .= '<p:sldId id="' . (256 + $i) . '" r:id="' . $rid . '"/>';
        $rels .= '<Relationship Id="' . $rid . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' . $n . '.xml"/>';
        $overrides .= '<Override PartName="/ppt/slides/slide' . $n . '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
    }
    $nx = $sc + 2;
    $rels .= '<Relationship Id="rId'.$nx.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>';
    $rels .= '<Relationship Id="rId'.($nx+1).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>';
    $rels .= '<Relationship Id="rId'.($nx+2).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
    $rels .= '<Relationship Id="rId'.($nx+3).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>';

    $zip->addFromString('ppt/_rels/presentation.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.$rels.'</Relationships>');
    $zip->addFromString('ppt/presentation.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>'.$sldIds.'</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>');
    $zip->addFromString('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'.$overrides.'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');

    $zip->close();
    return file_exists($outputPath);
}

function generateFromTemplate($outputPath, $allSlides, $tpl) {
    $zip = new ZipArchive();
    if ($zip->open($outputPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) return false;

    $sc  = count($allSlides);
    $cx  = $tpl['cx'];
    $cy  = $tpl['cy'];

    // ---- Descobre os masters existentes no template ----
    // Varre presentation.xml.rels para pegar só slideMaster
    $masterRels = []; // [ rId => target ]
    $presRelsXml = $tpl['files']['ppt/_rels/presentation.xml.rels'] ?? '';
    preg_match_all(
        '/<Relationship\s[^>]*Id="([^"]+)"[^>]*Type="[^"]*\/slideMaster"[^>]*Target="([^"]+)"/i',
        $presRelsXml, $m, PREG_SET_ORDER
    );
    foreach ($m as $row) $masterRels[$row[1]] = $row[2];

    // Se não achou nenhum master, usa o primeiro disponível pelos arquivos copiados
    if (empty($masterRels)) {
        foreach ($tpl['files'] as $name => $_) {
            if (preg_match('#^ppt/slideMasters/(slideMaster\d+\.xml)$#', $name, $mm)) {
                $masterRels['rId1'] = 'slideMasters/' . $mm[1];
            }
        }
    }

    // ---- Descobre um layout para usar nos novos slides ----
    // Pega o layout que o slide1 do template usava
    $layoutTarget = $tpl['layoutTarget']; // ex: ../slideLayouts/slideLayout1.xml
    // Normaliza para path relativo a ppt/
    $layoutPath = 'ppt/' . preg_replace('#^\.\./|^/#', '', $layoutTarget);

    // ---- Copia apenas arquivos de estrutura do template ----
    $allowedPrefixes = [
        'ppt/slideMasters/',
        'ppt/slideLayouts/',
        'ppt/theme/',
        'ppt/media/',
        'ppt/fonts/',
        'docProps/',
    ];
    $allowedExact = ['_rels/.rels'];
    foreach ($tpl['files'] as $name => $content) {
        $allowed = in_array($name, $allowedExact);
        if (!$allowed) {
            foreach ($allowedPrefixes as $pfx) {
                if (strpos($name, $pfx) === 0) { $allowed = true; break; }
            }
        }
        if ($allowed) $zip->addFromString($name, $content);
    }

    // ---- Slides novos ----
    $sldIds = ''; $slideRels = ''; $slideOverrides = '';
    // layoutTarget relativo a ppt/slides/_rels/ → '../../slideLayouts/slideLayoutN.xml'
    // Converte layoutPath (ppt/slideLayouts/slideLayout1.xml) → relativo a slides/
    $layoutRelTarget = '../../' . substr($layoutPath, 4); // remove 'ppt/'
    for ($i = 0; $i < $sc; $i++) {
        $n   = $i + 1;
        $rid = 'rId' . ($i + 200);
        $zip->addFromString("ppt/slides/slide{$n}.xml",
            buildSlideXmlNoBackground($allSlides[$i]['text'], $allSlides[$i]['bold'], $cx, $cy));
        $zip->addFromString("ppt/slides/_rels/slide{$n}.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' .
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' .
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"' .
            ' Target="' . htmlspecialchars($layoutRelTarget, ENT_XML1) . '"/>' .
            '</Relationships>');
        $sldIds        .= '<p:sldId id="' . (256 + $i) . '" r:id="' . $rid . '"/>';
        $slideRels     .= '<Relationship Id="' . $rid . '"' .
                          ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"' .
                          ' Target="slides/slide' . $n . '.xml"/>';
        $slideOverrides .= '<Override PartName="/ppt/slides/slide' . $n . '.xml"' .
                           ' ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
    }

    // ---- presentation.xml — construído do zero ----
    $masterIdLst = '';
    $masterOverrides = '';
    $presRels = '';
    $rIdx = 1;
    foreach ($masterRels as $origRid => $target) {
        $rid = 'rId' . $rIdx++;
        $masterIdLst    .= '<p:sldMasterId id="' . (2147483647 + $rIdx) . '" r:id="' . $rid . '"/>';
        $presRels       .= '<Relationship Id="' . $rid . '"' .
                           ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"' .
                           ' Target="' . htmlspecialchars($target, ENT_XML1) . '"/>';
        $masterOverrides .= '<Override PartName="/ppt/' . $target . '"' .
                            ' ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
    }
    $rBase = $rIdx;
    $presRels .= $slideRels;
    $presRels .= '<Relationship Id="rId' . ($rBase)   . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"   Target="presProps.xml"/>';
    $presRels .= '<Relationship Id="rId' . ($rBase+1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"    Target="viewProps.xml"/>';
    $presRels .= '<Relationship Id="rId' . ($rBase+2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"  Target="tableStyles.xml"/>';

    $zip->addFromString('ppt/presentation.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' .
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' .
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' .
        ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">' .
        '<p:sldMasterIdLst>' . $masterIdLst . '</p:sldMasterIdLst>' .
        '<p:sldIdLst>' . $sldIds . '</p:sldIdLst>' .
        '<p:sldSz cx="' . $cx . '" cy="' . $cy . '"/>' .
        '<p:notesSz cx="6858000" cy="9144000"/>' .
        '</p:presentation>');

    $zip->addFromString('ppt/_rels/presentation.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' .
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' .
        $presRels . '</Relationships>');

    // ---- Layouts: overrides de todos os layouts copiados ----
    $layoutOverrides = '';
    foreach ($tpl['files'] as $name => $_) {
        if (preg_match('#^ppt/slideLayouts/slideLayout\d+\.xml$#', $name)) {
            $layoutOverrides .= '<Override PartName="/' . $name . '"' .
                                ' ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
        }
    }

    // ---- [Content_Types].xml — construído do zero ----
    $zip->addFromString('[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' .
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' .
        '<Default Extension="xml"  ContentType="application/xml"/>' .
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' .
        '<Default Extension="png"  ContentType="image/png"/>' .
        '<Default Extension="jpg"  ContentType="image/jpeg"/>' .
        '<Default Extension="jpeg" ContentType="image/jpeg"/>' .
        '<Default Extension="svg"  ContentType="image/svg+xml"/>' .
        '<Default Extension="fntdata" ContentType="application/x-fontdata"/>' .
        '<Override PartName="/ppt/presentation.xml"' .
        ' ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' .
        $masterOverrides .
        $layoutOverrides .
        $slideOverrides .
        '<Override PartName="/ppt/presProps.xml"    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' .
        '<Override PartName="/ppt/viewProps.xml"    ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>' .
        '<Override PartName="/ppt/tableStyles.xml"  ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' .
        '<Override PartName="/docProps/core.xml"    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' .
        '<Override PartName="/docProps/app.xml"     ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' .
        '</Types>');

    // presProps, viewProps, tableStyles limpos
    $zip->addFromString('ppt/presProps.xml',   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>');
    $zip->addFromString('ppt/viewProps.xml',   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr></p:viewPr>');
    $zip->addFromString('ppt/tableStyles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>');
    $zip->addFromString('docProps/app.xml',    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Liturgia Slides</Application><Slides>' . $sc . '</Slides></Properties>');
    $zip->addFromString('docProps/core.xml',   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Liturgia</dc:title><dc:creator>Liturgia Slides</dc:creator></cp:coreProperties>');
    $zip->addFromString('_rels/.rels',         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');

    $zip->close();
    return file_exists($outputPath);
}

function generateViaPython($outputPath, $allSlides, $massType = 'dominical') {
    $pythonBin  = 'C:/Python314/python.exe';
    $scriptPath = __DIR__ . '/generate_pptx.py';
    $slidesJson = json_encode($allSlides, JSON_UNESCAPED_UNICODE);
    $templatePath = getEffectiveTemplatePath($massType);
    $templateArg = $templatePath ? escapeshellarg($templatePath) : '';

    // Escreve o JSON em arquivo temporário para evitar problemas com aspas no CLI
    $tmpJson = tempnam(sys_get_temp_dir(), 'pptx_') . '.json';
    file_put_contents($tmpJson, $slidesJson);

    $cmd = escapeshellarg($pythonBin)
         . ' ' . escapeshellarg($scriptPath)
         . ' ' . escapeshellarg($tmpJson)
         . ' ' . escapeshellarg($outputPath)
         . ($templateArg ? ' ' . $templateArg : '')
         . ' 2>&1';

    $output = shell_exec($cmd);
    @unlink($tmpJson);

    if ($output === null) {
        // shell_exec bloqueado — tenta via proc_open
        $descriptors = [1 => ['pipe','w'], 2 => ['pipe','w']];
        $proc = proc_open($cmd, $descriptors, $pipes);
        if ($proc) {
            $output = stream_get_contents($pipes[1]) . stream_get_contents($pipes[2]);
            proc_close($proc);
        }
    }

    $result = @json_decode(trim((string)$output), true);
    if (!isset($result['success'])) {
        // Propaga o erro para aparecer na tela
        throw new \RuntimeException('Python output: ' . substr((string)$output, 0, 500));
    }
    return $result['success'] === true;
}

function buildSlideXmlNoBackground($text, $bold, $cx, $cy) {
    $b = $bold ? ' b="1"' : '';
    $paras = '';
    foreach (explode("\n", $text) as $p) {
        $e = htmlspecialchars($p, ENT_XML1, 'UTF-8');
        $paras .= '<a:p><a:pPr algn="ctr" indent="0" marL="0"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/></a:pPr><a:r><a:rPr lang="pt-BR" sz="7200" dirty="0"'.$b.'><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial Black" pitchFamily="34" charset="0"/><a:ea typeface="Arial Black" pitchFamily="34" charset="-122"/><a:cs typeface="Arial Black" pitchFamily="34" charset="-120"/></a:rPr><a:t>'.$e.'</a:t></a:r><a:endParaRPr lang="pt-BR" sz="7200" dirty="0"/></a:p>';
    }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'.$cx.'" cy="'.$cy.'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/>'.$paras.'</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
}

function buildSlideXml($text, $bold = false) {
    $mx = 0; $w = 9144000; $h = 6858000; $sz = 7200;
    $b = $bold ? ' b="1"' : '';
    $paras = '';
    foreach (explode("\n", $text) as $p) {
        $e = htmlspecialchars($p, ENT_XML1, 'UTF-8');
        $paras .= '<a:p><a:pPr algn="ctr" indent="0" marL="0"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/></a:pPr><a:r><a:rPr lang="pt-BR" sz="'.$sz.'" dirty="0"'.$b.'><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial Black" pitchFamily="34" charset="0"/><a:ea typeface="Arial Black" pitchFamily="34" charset="-122"/><a:cs typeface="Arial Black" pitchFamily="34" charset="-120"/></a:rPr><a:t>'.$e.'</a:t></a:r><a:endParaRPr lang="pt-BR" sz="'.$sz.'" dirty="0"/></a:p>';
    }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="'.$mx.'" y="'.$mx.'"/><a:ext cx="'.$w.'" cy="'.$h.'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/>'.$paras.'</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
}

/* ================================================================
   DOWNLOAD
   ================================================================ */
function handleDownload() {
    global $OUTPUT_DIR;
    $filename = basename($_GET['file'] ?? '');
    $filepath = $OUTPUT_DIR . $filename;
    if (empty($filename) || !file_exists($filepath)) {
        http_response_code(404);
        echo json_encode(['success'=>false,'error'=>'Arquivo não encontrado'],JSON_UNESCAPED_UNICODE);
        return;
    }
    header('Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation');
    header('Content-Disposition: attachment; filename="'.$filename.'"');
    header('Content-Length: '.filesize($filepath));
    header('Cache-Control: no-cache');
    readfile($filepath);
    exit;
}
