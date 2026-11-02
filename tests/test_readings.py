import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from pptx import Presentation

from app import app, build_section_groups
from generate_pptx import (
    P14_NS, PKG_NS, generate_no_template, generate_with_template,
    get_global_slide_order, get_sections,
)


SECTIONS = [
    {'key': 'primeira_leitura', 'text': 'Ex 22,20-26',
     'reading_text': '20Assim diz o Senhor: Acolhei o estrangeiro.'},
    {'key': 'salmo_responsorial', 'text': 'O Senhor é minha força.',
     'reading_text': '– Eu vos amo, ó Senhor!\n– Bendito seja o meu rochedo!'},
    {'key': 'segunda_leitura', 'text': '1Ts 1,5-10',
     'reading_text': '5Irmãos, acolhei a Palavra com alegria.'},
    {'key': 'evangelho', 'text': 'Mt 22,34-40',
     'reading_text': '34Naquele tempo, Jesus disse: Amai o próximo.'},
]


def slide_texts(path):
    return ['\n'.join(shape.text for shape in slide.shapes if shape.has_text_frame)
            for slide in Presentation(path).slides]


def make_template(path, titles, with_section=True):
    prs = Presentation()
    for title in titles:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = title
        slide.placeholders[1].text = 'Conteúdo original'
    if with_section:
        ns = '{' + P14_NS + '}'
        ext_lst = etree.SubElement(prs._element, '{%s}extLst' % prs._element.nsmap['p'])
        ext = etree.SubElement(ext_lst, '{%s}ext' % prs._element.nsmap['p'], uri='{TEST}')
        sections = etree.SubElement(ext, ns + 'sectionLst')
        section = etree.SubElement(sections, ns + 'section', name='LEITURAS', id='{TEST}')
        ids = etree.SubElement(section, ns + 'sldIdLst')
        for slide_id in prs.slides._sldIdLst:
            etree.SubElement(ids, ns + 'sldId', id=slide_id.get('id'))
    prs.save(path)


class ReadingGenerationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.output = self.root / 'output.pptx'
        self.groups = build_section_groups(SECTIONS)

    def expected_slides(self):
        expected = []
        for group in self.groups:
            expected.extend(s['text'] for s in group['slides'])
            expected.extend(s['text'] for s in group.get('reading_slides', []))
            for stanza in group.get('psalm_stanzas', []):
                expected.extend(s['text'] for s in stanza)
                expected.extend(s['text'] for s in group['slides'])
        return expected

    def test_no_template_keeps_all_readings_and_psalm_responses_in_order(self):
        generate_no_template(self.output, self.groups)
        self.assertEqual(slide_texts(self.output), self.expected_slides())

    def test_reading_without_reference_is_not_lost(self):
        groups = build_section_groups([{'key': 'evangelho', 'text': '',
                                       'reading_text': 'Jesus disse: Amai-vos.'}])
        generate_no_template(self.output, groups)
        self.assertTrue(slide_texts(self.output))
        self.assertIn('Amai-vos.', ' '.join(slide_texts(self.output)))

    def test_template_without_liturgy_section_keeps_complete_content(self):
        template = self.root / 'template.pptx'
        make_template(template, ['Capa'], with_section=False)
        generate_with_template(template, self.output, self.groups)
        self.assertEqual(slide_texts(self.output)[1:], self.expected_slides())

    def test_template_without_expected_fields_keeps_complete_content(self):
        template = self.root / 'template.pptx'
        make_template(template, ['Capa das leituras'])
        generate_with_template(template, self.output, self.groups)
        self.assertEqual(slide_texts(self.output)[1:], self.expected_slides())

    def test_template_with_expected_fields_keeps_readings_in_order(self):
        template = self.root / 'template.pptx'
        make_template(template, ['Primeira leitura', 'Salmo', 'Segunda leitura', 'Evangelho'])
        generate_with_template(template, self.output, self.groups)
        texts = slide_texts(self.output)
        joined = '\n'.join(texts)
        for group in self.groups:
            for slide in group.get('reading_slides', []):
                self.assertIn(slide['text'], texts)
        for section in SECTIONS:
            self.assertIn(section['text'], joined)
        positions = [joined.index(word) for word in ['Acolhei', 'Eu vos amo', 'Irmãos', 'Naquele tempo']]
        self.assertEqual(positions, sorted(positions))

    def test_partial_template_keeps_missing_gospel_after_other_readings(self):
        template = self.root / 'template.pptx'
        make_template(template, ['Primeira leitura', 'Salmo', 'Segunda leitura'])
        generate_with_template(template, self.output, self.groups)
        joined = '\n'.join(slide_texts(self.output))
        positions = [joined.index(word) for word in ['Acolhei', 'Eu vos amo', 'Irmãos', 'Naquele tempo']]
        self.assertEqual(positions, sorted(positions))

    def test_partial_template_keeps_missing_first_reading_before_psalm(self):
        template = self.root / 'template.pptx'
        make_template(template, ['Salmo', 'Segunda leitura', 'Evangelho'])
        generate_with_template(template, self.output, self.groups)
        joined = '\n'.join(slide_texts(self.output))
        positions = [joined.index(word) for word in ['Acolhei', 'Eu vos amo', 'Irmãos', 'Naquele tempo']]
        self.assertEqual(positions, sorted(positions))

    def test_missing_second_reading_is_inserted_after_psalm(self):
        template = self.root / 'template.pptx'
        make_template(template, ['Primeira leitura', 'Salmo', 'Evangelho'])
        generate_with_template(template, self.output, self.groups)
        joined = '\n'.join(slide_texts(self.output))
        self.assertIn('1Ts 1,5-10', joined)
        positions = [joined.index(word) for word in ['Acolhei', 'Eu vos amo', 'Irmãos', 'Naquele tempo']]
        self.assertEqual(positions, sorted(positions))

    def test_second_reading_in_free_textbox_keeps_reference(self):
        from pptx.util import Inches
        template = self.root / 'template.pptx'
        make_template(template, ['Segunda leitura'])
        prs = Presentation(template)
        body = prs.slides[0].placeholders[1]._element
        body.getparent().remove(body)
        textbox = prs.slides[0].shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
        textbox.text = 'Segunda leitura (referência antiga)'
        prs.save(template)
        generate_with_template(template, self.output, [self.groups[2]])
        joined = '\n'.join(slide_texts(self.output))
        self.assertIn('1Ts 1,5-10', joined)
        self.assertNotIn('referência antiga', joined)

    def test_ordinary_song_slides_are_preserved(self):
        slides = [{'text': 'Cântico de entrada', 'bold': True},
                  {'text': 'Cantemos ao Senhor', 'bold': False}]
        generate_no_template(self.output, [{'section_key': 'cantico_entrada', 'slides': slides}])
        self.assertEqual(slide_texts(self.output), [s['text'] for s in slides])

    def test_section_slide_ids_follow_presentation_order(self):
        template = self.root / 'template.pptx'
        make_template(template, ['Primeira leitura', 'Salmo', 'Segunda leitura', 'Evangelho'])
        generate_with_template(template, self.output, self.groups)
        with ZipFile(self.output) as package:
            root = etree.fromstring(package.read('ppt/presentation.xml'))
        order = get_global_slide_order(root)
        for section in get_sections(root):
            positions = [order.index(sid) for sid in section['sld_ids']]
            self.assertEqual(positions, sorted(positions))

    def test_import_metadata_relationship_is_preserved(self):
        template = self.root / 'template.pptx'
        make_template(template, ['Primeira leitura', 'Salmo', 'Evangelho'])
        with ZipFile(template) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        rels_name = 'ppt/_rels/presentation.xml.rels'
        rels = etree.fromstring(parts[rels_name])
        etree.SubElement(rels, '{%s}Relationship' % PKG_NS,
                         Id='rIdImportMetadata',
                         Type='http://customschemas.google.com/relationships/presentationmetadata',
                         Target='metadata.xml')
        parts[rels_name] = etree.tostring(rels)
        parts['ppt/metadata.xml'] = b'<metadata>import metadata</metadata>'
        with ZipFile(template, 'w') as package:
            for name, data in parts.items():
                package.writestr(name, data)
        generate_with_template(template, self.output, self.groups)
        with ZipFile(self.output) as package:
            rels = etree.fromstring(package.read(rels_name))
            matching = [r for r in rels if r.get('Id') == 'rIdImportMetadata']
            self.assertEqual(len(matching), 1)
            self.assertEqual(matching[0].get('Target'), 'metadata.xml')
            self.assertEqual(package.read('ppt/metadata.xml'), parts['ppt/metadata.xml'])

    def test_generate_api_download_contains_readings(self):
        from unittest.mock import patch
        form = {'custom_filename': 'regression_readings'}
        for section in SECTIONS:
            form['text_' + section['key']] = section['text']
            form['reading_text_' + section['key']] = section['reading_text']
        with patch('app.OUTPUT_DIR', self.root), patch('app.get_effective_template_path', return_value=None):
            client = app.test_client()
            response = client.post('/api?action=generate', data=form)
            result = response.get_json()
            self.assertTrue(result['success'], result)
            download = client.get('/api?action=download&file=' + result['filename'])
            self.assertEqual(download.status_code, 200)
            download.close()
            self.assertEqual(slide_texts(self.root / result['filename']), self.expected_slides())


if __name__ == '__main__':
    unittest.main()
