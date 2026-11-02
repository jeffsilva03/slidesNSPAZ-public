import unittest

from app import app


class StaticFileIsolationTests(unittest.TestCase):
    def test_ui_and_referenced_assets_remain_available(self):
        with app.test_client() as client:
            for path in ('/', '/assets/images/igrejas.png', '/assets/images/logoBranca.png'):
                with self.subTest(path=path):
                    response = client.get(path)
                    self.assertEqual(response.status_code, 200)
                    response.close()

    def test_root_files_and_work_directories_are_not_public(self):
        with app.test_client() as client:
            for path in ('/app.py', '/.git/config', '/.env', '/requirements.txt', '/uploads/private.txt', '/output/private.pptx', '/assets/template.pptx'):
                with self.subTest(path=path):
                    self.assertEqual(client.get(path).status_code, 404)
