import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from metadata_editor.engine import load, write_office


class WordMetadataTests(unittest.TestCase):
    def make_docx(self, folder, custom=None):
        path = Path(folder) / "sample.docx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>')
            archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
            archive.writestr("docProps/core.xml", '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>')
            archive.writestr("docProps/app.xml", '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>')
            archive.writestr("word/document.xml", '<document>body stays intact</document>')
            if custom:
                archive.writestr("docProps/custom.xml", custom)
        return path

    def test_new_custom_property_adds_package_registration(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.make_docx(folder)
            meta = load(str(path))
            write_office(meta, meta.values, {"项目": "测试"}, False)
            self.assertEqual(load(str(path)).custom["项目"], "测试")
            with zipfile.ZipFile(path) as archive:
                self.assertIn(b"custom-properties+xml", archive.read("[Content_Types].xml"))
                self.assertIn(b"custom-properties", archive.read("_rels/.rels"))
                self.assertEqual(archive.read("word/document.xml"), b"<document>body stays intact</document>")

    def test_untouched_typed_custom_property_keeps_type_and_id(self):
        custom = ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                  '<property name="Count" pid="19" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"><vt:i4>42</vt:i4></property>'
                  '</Properties>')
        with tempfile.TemporaryDirectory() as folder:
            path = self.make_docx(folder, custom)
            meta = load(str(path))
            write_office(meta, {**meta.values, "title": "报告"}, meta.custom, False)
            with zipfile.ZipFile(path) as archive:
                root = ET.fromstring(archive.read("docProps/custom.xml"))
            prop = next(iter(root))
            self.assertEqual(prop.get("pid"), "19")
            self.assertTrue(next(iter(prop)).tag.endswith("}i4"))
            self.assertEqual(next(iter(prop)).text, "42")


if __name__ == "__main__":
    unittest.main()
