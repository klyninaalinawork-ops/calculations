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


def extract_materials(section):
    raw_quantity = 0.0
    raw_cost = 0.0
    packaging_quantity = 0.0
    packaging_cost = 0.0
    current_category = ""
    current_item = ""
    pending_item = None
    items = []
    has_raw_quantity = False
    has_packaging_quantity = False

    for row_number, values in section:
        raw_label = values.get("B", "")
        label = raw_label.strip()
        unit = values.get("C", "").strip().lower()
        amount = number(values.get("D", ""))

        if label.startswith("23 "):
            current_category = label
            current_item = ""
            pending_item = None
        elif label == "Количество" and amount is not None:
            is_packaging = "упаков" in current_category.lower()
            if is_packaging:
                packaging_quantity += amount
                has_packaging_quantity = True
            elif unit == "кг":
                raw_quantity += amount
                has_raw_quantity = True
            if current_item:
                pending_item = {
                    "name": current_item,
                    "type": "packaging" if is_packaging else "raw",
                    "quantity": amount,
                    "unit": "шт" if is_packaging else "кг",
                    "price": None,
                    "cost": None,
                }
        elif label == "Цена" and amount is not None and pending_item:
            pending_item["price"] = amount
        elif label == "Сумма" and amount is not None:
            is_packaging = "упаков" in current_category.lower()
            if is_packaging:
                packaging_cost += amount
            elif is_material_category(current_category):
                raw_cost += amount
            if pending_item:
                pending_item["cost"] = amount
                items.append(pending_item)
                pending_item = None
        elif label and not label.startswith("24 ") and not label.startswith("Показник") and not label.startswith("23 "):
            current_item = label

    return {
        "items": items,
        "rawQuantityKg": raw_quantity if has_raw_quantity else None,
        "rawCost": raw_cost if raw_cost else None,
        "rawPricePerKg": raw_cost / raw_quantity if has_raw_quantity and raw_quantity else None,
        "packagingQuantity": packaging_quantity if has_packaging_quantity else None,
        "packagingUnit": "шт",
        "packagingCost": packaging_cost if packaging_cost else None,
        "packagingPrice": packaging_cost / packaging_quantity if has_packaging_quantity and packaging_quantity else None,
    }


def is_material_category(label):
    category = label.lower()
    excluded = (
        "упаков", "оплата праці", "заробіт", "єсв", "подат",
        "газ", "електро", "амортизац", "енерг", "паливо",
        "витрати на збут", "адміністратив",
    )
    return label.startswith("23 ") and not any(term in category for term in excluded)


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
            "materials": extract_materials(section),
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
