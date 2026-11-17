"""Replace native category-chart data while retaining the source chart styling.

Only the series name/category/value nodes and the embedded XLSX are rewritten.
The source series and category counts stay fixed, retaining all plot assignments,
axes, point formatting, labels, legends and extension data.
"""

from __future__ import annotations

import math

from lxml import etree as E
from pptx.chart.data import CategoryChartData
from pptx.chart.xmlwriter import SeriesXmlRewriterFactory
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml import parse_xml

from .models import _xml_text


C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS = {"c": C}
SUPPORTED_PLOTS = {
    "areaChart", "area3DChart", "barChart", "bar3DChart", "lineChart", "line3DChart",
    "pieChart", "pie3DChart", "doughnutChart", "ofPieChart", "radarChart",
}


def _root(xml_bytes):
    if not isinstance(xml_bytes, bytes):
        raise ValueError("Ожидаются исходные байты XML диаграммы")
    try:
        return E.fromstring(xml_bytes, E.XMLParser(resolve_entities=False, no_network=True))
    except E.XMLSyntaxError as exc:
        raise ValueError("Некорректная XML-разметка диаграммы") from exc


def _cached_points(node):
    """Read a source cache in point-index order, retaining absent data as None."""
    if node is None:
        return []
    cache = next((child for child in node.iter()
                  if E.QName(child).namespace == C
                  and E.QName(child).localname in {"strCache", "numCache", "strLit", "numLit"}), None)
    if cache is None:
        return []
    points = {}
    for point in cache.findall(f"{{{C}}}pt"):
        value = point.find(f"{{{C}}}v")
        points[int(point.get("idx"))] = value.text if value is not None and value.text is not None else ""
    count = cache.find(f"{{{C}}}ptCount")
    size = int(count.get("val")) if count is not None else max(points, default=-1) + 1
    if not 0 <= size <= 10000:
        raise ValueError("Некорректное число точек в диаграмме")
    return [points.get(index) for index in range(size)]


def inspect_chart(xml_bytes: bytes) -> dict:
    """Return cached source data and whether this category chart can be filled.

    Combo charts retain the original order of plots and their series. Modern
    chartEx, scatter/bubble, date/multilevel axes and manually written data labels
    are reported as unsupported instead of silently changing their semantics.
    """
    root = _root(xml_bytes)
    meta = {"supported": False, "chart_types": [], "categories": [], "series": [],
            "series_count": 0, "category_count": 0, "requires_same_dimensions": True,
            "reason": ""}

    def unsupported(reason):
        meta["reason"] = reason
        return meta

    if E.QName(root).namespace != C or E.QName(root).localname != "chartSpace":
        return unsupported("Поддерживаются классические категориальные диаграммы DrawingML; chartEx сохраняется без изменений")
    area = root.find(f"{{{C}}}chart/{{{C}}}plotArea")
    if area is None:
        return unsupported("В диаграмме отсутствует plotArea")
    plots = [node for node in area if E.QName(node).namespace == C
             and E.QName(node).localname.endswith("Chart")]
    meta["chart_types"] = [E.QName(node).localname for node in plots]
    source_series = [series for plot in plots for series in plot.findall(f"{{{C}}}ser")]
    categories_by_series = []
    for index, series in enumerate(source_series):
        categories = _cached_points(series.find(f"{{{C}}}cat"))
        categories_by_series.append(categories)
        title = _cached_points(series.find(f"{{{C}}}tx"))
        literal_title = series.find(f"{{{C}}}tx/{{{C}}}v")
        name = title[0] if title else literal_title.text if literal_title is not None else f"Ряд {index + 1}"
        values = []
        for value in _cached_points(series.find(f"{{{C}}}val")):
            try:
                number = float(value) if value not in (None, "") else None
                values.append(number if number is None or math.isfinite(number) else None)
            except (TypeError, ValueError):
                values.append(None)
        number_format = series.find(f"{{{C}}}val/{{{C}}}numRef/{{{C}}}numCache/{{{C}}}formatCode")
        if number_format is None:
            number_format = series.find(f"{{{C}}}val/{{{C}}}numLit/{{{C}}}formatCode")
        meta["series"].append({"name": name or f"Ряд {index + 1}", "values": values,
                               "number_format": number_format.text if number_format is not None else "General"})
    meta["categories"] = categories_by_series[0] if categories_by_series else []
    meta["category_count"] = len(meta["categories"])
    meta["series_count"] = len(source_series)
    if not plots or any(kind not in SUPPORTED_PLOTS for kind in meta["chart_types"]):
        return unsupported("Тип диаграммы требует другого контракта данных; XY, bubble и stock сохраняются без изменений")
    if root.find(f"{{{C}}}pivotSource") is not None:
        return unsupported("Сводная диаграмма связана с исходной сводной таблицей")
    if area.find(f"{{{C}}}dateAx") is not None:
        return unsupported("Ось дат требует отдельного контракта дат")
    if root.xpath(".//c:multiLvlStrRef", namespaces=NS):
        return unsupported("Многоуровневые категории сохраняются без изменений")
    if root.xpath(".//c:dLbl/c:tx | .//*[local-name()='datalabelsRange']", namespaces=NS):
        return unsupported("В диаграмме есть пользовательские подписи данных, которые нельзя безопасно обновить по числам")
    if not source_series or not meta["categories"] or any(category is None for category in meta["categories"]):
        return unsupported("В источнике нет полного кэша категорий")
    if any(categories != meta["categories"] for categories in categories_by_series):
        return unsupported("Ряды используют разные наборы категорий")
    if any(len(series["values"]) != meta["category_count"] for series in meta["series"]):
        return unsupported("Длина исходного ряда отличается от числа категорий")
    meta["supported"] = True
    return meta


def _label(value, name):
    if not isinstance(value, str) or not value.strip() or len(value) > 1000:
        raise ValueError(f"{name}: нужен непустой текст длиной до 1000 символов")
    _xml_text(value)
    if value.startswith(("=", "{=")):
        raise ValueError(f"{name}: подпись должна быть текстом, а не формулой Excel")
    return value


def _finite_number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    try:
        return math.isfinite(value)
    except OverflowError:
        return False


def _validated_data(chart_data, metadata):
    if not isinstance(chart_data, dict) or set(chart_data) != {"categories", "series"}:
        raise ValueError("chart_data должен содержать только categories и series")
    categories, series = chart_data["categories"], chart_data["series"]
    if not isinstance(categories, list) or len(categories) != metadata["category_count"]:
        raise ValueError(f"Сохраните исходное число категорий: {metadata['category_count']}")
    categories = [_label(value, "Категория") for value in categories]
    if not isinstance(series, list) or len(series) != metadata["series_count"]:
        raise ValueError(f"Сохраните исходное число рядов: {metadata['series_count']}")
    normalized = []
    for item in series:
        if not isinstance(item, dict) or set(item) != {"name", "values"}:
            raise ValueError("Каждый ряд должен содержать только name и values")
        name = _label(item["name"], "Название ряда")
        values = item["values"]
        if not isinstance(values, list) or len(values) != len(categories):
            raise ValueError("Число значений каждого ряда должно совпадать с числом категорий")
        if any(not _finite_number(value) for value in values):
            raise ValueError("Значения диаграммы должны быть конечными числами")
        normalized.append((name, values))
    return categories, normalized


def replace_chart(xml_bytes: bytes, chart_data: dict) -> tuple[bytes, bytes]:
    """Replace supported source data and return chart XML plus an embedded XLSX.

    The caller retains the chart's original relationships and stores the XLSX at
    its existing embedded-workbook target. No presentation package is modified.
    """
    metadata = inspect_chart(xml_bytes)
    if not metadata["supported"]:
        raise ValueError("Изменение данных этой диаграммы не поддерживается: " + metadata["reason"])
    categories, series = _validated_data(chart_data, metadata)
    data = CategoryChartData()
    data.categories = categories
    for index, (name, values) in enumerate(series):
        data.add_series(name, values, number_format=metadata["series"][index]["number_format"] or "General")
    chart_space = parse_xml(xml_bytes)
    # All supported plots use the same category-series rewriter. Rewriting
    # existing series individually avoids count adjustment and plot dispatch;
    # python-pptx cannot classify the source ofPieChart plot itself.
    rewriter = SeriesXmlRewriterFactory(XL_CHART_TYPE.COLUMN_CLUSTERED, data)
    series_nodes = chart_space.xpath(".//c:plotArea/*/c:ser")
    for node, series_data in zip(series_nodes, data):
        rewriter._rewrite_ser_data(node, series_data, chart_space.date_1904)
    result = E.tostring(chart_space, xml_declaration=True, encoding="UTF-8", standalone=True)
    return result, data.xlsx_blob
