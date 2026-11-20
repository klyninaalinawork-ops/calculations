import json
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "public" / "data" / "calculations.json"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
METRICS = {
    "Випуск продукції": "volume",
    "Виробнича собівартість": "productionCost",
    "Повні витрати": "fullCost",
    "Прибуток (+) Збитки (-)": "profit",
    "Рентабельність": "margin",
    "Оптова вартість": "wholesalePrice",
    "Відпускна вартість": "retailPrice",
}


def column_name(cell_ref):
    match = re.match(r"[A-Z]+", cell_ref)
    return match.group(0) if match else ""


def number(value):
    if not value:
        return None
    try:
        return float(value.replace(",", "."))
    except ValueError:
        return None


def read_workbook(path):
    with zipfile.ZipFile(path) as workbook:
        shared_strings = []
        if "xl/SharedStrings.xml" in workbook.namelist():
            strings_root = ET.fromstring(workbook.read("xl/SharedStrings.xml"))
            shared_strings = [
                "".join(text.text or "" for text in item.findall(".//m:t", NS))
                for item in strings_root.findall("m:si", NS)
            ]

        sheet = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
        rows = []
        for row in sheet.findall(".//m:sheetData/m:row", NS):
            values = {}
            for cell in row.findall("m:c", NS):
                raw = cell.find("m:v", NS)
                value = raw.text if raw is not None else ""
                if cell.attrib.get("t") == "s" and value:
                    value = shared_strings[int(value)]
                elif cell.attrib.get("t") == "inlineStr":
                    value = "".join(text.text or "" for text in cell.findall(".//m:t", NS))
                if value:
                    values[column_name(cell.attrib.get("r", ""))] = value.strip()
            rows.append((int(row.attrib["r"]), values))

    starts = [index for index, (_, values) in enumerate(rows) if values.get("B") == "Калькуляція"]
    records = []
    for calculation_index, start in enumerate(starts):
        end = starts[calculation_index + 1] if calculation_index + 1 < len(starts) else len(rows)
        section = rows[start:end]
        product = section[1][1].get("B", "") if len(section) > 1 else ""
        if not product:
            continue

        values_by_label = {}
        cost_items = []
        period = section[5][1].get("B", "").split(" - ")[0].strip() if len(section) > 5 else ""
        entity = section[2][1].get("C", "") if len(section) > 2 else ""
        for row_number, values in section:
            label = values.get("B", "").strip()
            clean_label = label.rstrip(":").strip()
            amount = number(values.get("D", ""))
            if clean_label:
                values_by_label[clean_label] = amount
            if label.startswith("23 ") and amount is not None:
                cost_items.append({"label": label, "value": amount})

        metrics = {
            key: values_by_label.get(label)
            for label, key in METRICS.items()
        }
        if any(metrics[key] is None for key in ("volume", "fullCost", "profit")):
            raise ValueError(f"Не вдалося прочитати обов'язкові показники: {path.name} — {product}")

        records.append({
            "id": f"{path.stem}-{calculation_index + 1}",
            "client": path.stem.removeprefix("Калькуляція ").strip(),
            "entity": entity,
            "product": product,
            "period": period,
            "source": path.name,
            "metrics": metrics,
            "costItems": sorted(cost_items, key=lambda item: abs(item["value"]), reverse=True)[:8],
        })
    return records


def main():
    files = sorted(ROOT.glob("Калькуляція *.xlsx"))
    if not files:
        raise FileNotFoundError(f"Не знайдено Excel-файлів у {ROOT}")

    records = [record for path in files for record in read_workbook(path)]
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(
        json.dumps({"records": records}, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    print(f"Збережено {len(records)} калькуляцій із {len(files)} файлів у {OUTPUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
