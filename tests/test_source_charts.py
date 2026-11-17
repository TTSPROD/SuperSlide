"""Verify native chart replacement against the provided source PPTX files."""

import copy
import io
import math
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

from superslide.source_charts import C, inspect_chart, replace_chart


ROOT = Path(__file__).resolve().parents[1]
NS = {"c": C}
X = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def source_charts():
    for path in sorted((ROOT / "resources/severstal/templates").glob("*.pptx")):
        with ZipFile(path) as archive:
            for part in archive.namelist():
                if "/charts/chart" in part and part.endswith(".xml") and "/_rels/" not in part:
                    yield path.name + ":" + part, archive.read(part)


def data_for(metadata):
    return {"categories": [f"Период {i + 1}" for i in range(metadata["category_count"])],
            "series": [{"name": f"Показатель {i + 1}",
                        "values": [(-1 if j % 2 else 1) * (100 * (i + 1) + j + 0.25)
                                   for j in range(metadata["category_count"])]}
                       for i in range(metadata["series_count"])]}


def style_xml(xml_bytes):
    root = E.fromstring(xml_bytes)
    for series in root.xpath(".//c:ser", namespaces=NS):
        for child in list(series):
            if E.QName(child).namespace == C and E.QName(child).localname in {"tx", "cat", "val"}:
                series.remove(child)
    return E.tostring(root, method="c14n")


class SourceChartTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.sources = list(source_charts())
        cls.supported = [(name, xml, inspect_chart(xml)) for name, xml in cls.sources
                         if inspect_chart(xml)["supported"]]

    def test_all_supported_source_charts_replace_values_and_retain_every_style_node(self):
        self.assertGreaterEqual(len(self.supported), 70)
        kinds = set()
        for name, original, metadata in self.supported:
            with self.subTest(chart=name):
                data = data_for(metadata)
                replacement, workbook = replace_chart(original, data)
                after = inspect_chart(replacement)
                self.assertTrue(after["supported"])
                self.assertEqual(after["categories"], data["categories"])
                self.assertEqual([{k: s[k] for k in ("name", "values")} for s in after["series"]], data["series"])
                self.assertEqual([s["number_format"] for s in after["series"]],
                                 [s["number_format"] for s in metadata["series"]])
                self.assertEqual(style_xml(original), style_xml(replacement))
                self.assertTrue(workbook.startswith(b"PK"))
                kinds.update(metadata["chart_types"])
        self.assertTrue({"barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "ofPieChart"} <= kinds)

    def test_combo_chart_preserves_plots_and_embeds_a_workbook_matching_the_cache(self):
        name, original, metadata = next(item for item in self.supported if len(item[2]["chart_types"]) > 1)
        data = data_for(metadata)
        replacement, workbook = replace_chart(original, data)
        self.assertEqual(inspect_chart(replacement)["chart_types"], metadata["chart_types"])
        with ZipFile(io.BytesIO(workbook)) as archive:
            shared = E.fromstring(archive.read("xl/sharedStrings.xml"))
            strings = ["".join(item.itertext()) for item in shared.findall("s:si", X)]
            sheet = E.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            cells = {}
            for cell in sheet.findall(".//s:c", X):
                value = cell.find("s:v", X).text
                cells[cell.get("r")] = strings[int(value)] if cell.get("t") == "s" else float(value)
            for row, category in enumerate(data["categories"], 2):
                self.assertEqual(cells[f"A{row}"], category)
            for col, series in enumerate(data["series"], 1):
                letter = chr(ord("A") + col)
                self.assertEqual(cells[f"{letter}1"], series["name"])
                self.assertEqual([cells[f"{letter}{row}"] for row in range(2, len(data["categories"]) + 2)], series["values"])
            self.assertFalse(sheet.findall(".//s:f", X))

    def test_source_dimensions_and_numeric_contract_are_enforced(self):
        _, original, metadata = self.supported[0]
        good = data_for(metadata)
        invalid = []
        changed = copy.deepcopy(good); changed["categories"].append("Лишняя категория"); invalid.append(changed)
        changed = copy.deepcopy(good); changed["series"].append(changed["series"][0]); invalid.append(changed)
        changed = copy.deepcopy(good); changed["categories"][0] = "=1+1"; invalid.append(changed)
        changed = copy.deepcopy(good); changed["series"][0]["name"] = "{=1+1}"; invalid.append(changed)
        for value in (math.nan, math.inf, -math.inf, True, "12"):
            changed = copy.deepcopy(good); changed["series"][0]["values"][0] = value; invalid.append(changed)
        for data in invalid:
            with self.subTest(data=data):
                with self.assertRaises(ValueError):
                    replace_chart(original, data)

    def test_unsupported_source_charts_are_described_and_rejected_without_changes(self):
        unsupported = [(name, xml, inspect_chart(xml)) for name, xml in self.sources
                       if not inspect_chart(xml)["supported"]]
        self.assertTrue(unsupported)
        self.assertTrue(any("пользовательские подписи" in item[2]["reason"] for item in unsupported))
        for name, original, metadata in unsupported:
            with self.subTest(chart=name):
                self.assertTrue(metadata["reason"])
                with self.assertRaises(ValueError):
                    replace_chart(original, {"categories": ["A"], "series": [{"name": "Ряд", "values": [1]}]})

    def test_scatter_bubble_and_multilevel_categories_are_not_silently_converted(self):
        _, original, _ = self.supported[0]
        for kind in ("scatterChart", "bubbleChart"):
            root = E.fromstring(original)
            plot = next(child for child in root.find("c:chart/c:plotArea", NS)
                        if E.QName(child).localname.endswith("Chart"))
            plot.tag = f"{{{C}}}{kind}"
            metadata = inspect_chart(E.tostring(root))
            self.assertFalse(metadata["supported"])
            self.assertIn(kind, metadata["chart_types"])
        root = E.fromstring(original)
        category = root.find(".//c:cat", NS)
        E.SubElement(category, f"{{{C}}}multiLvlStrRef")
        self.assertFalse(inspect_chart(E.tostring(root))["supported"])


if __name__ == "__main__":
    unittest.main()
