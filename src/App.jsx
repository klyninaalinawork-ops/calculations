import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const STATUSES = ["Нова", "В роботі", "На перевірці", "Погоджено"];
const STATUS_CLASS = {
  Нова: "status-new",
  "В роботі": "status-progress",
  "На перевірці": "status-review",
  Погоджено: "status-approved",
};
const DEFAULT_MANAGEMENT = { status: "Нова", note: "" };
const EMPTY_RECORDS = [];
const currency = new Intl.NumberFormat("uk-UA", {
  style: "currency",
  currency: "UAH",
  maximumFractionDigits: 0,
});
const amount = new Intl.NumberFormat("uk-UA", { maximumFractionDigits: 0 });
const quantity = new Intl.NumberFormat("uk-UA", { maximumFractionDigits: 1 });
const decimal = new Intl.NumberFormat("uk-UA", { maximumFractionDigits: 2 });

function readManagement() {
  const saved = localStorage.getItem("calculations-management");
  return saved ? JSON.parse(saved) : {};
}

function readManualRecords() {
  const saved = localStorage.getItem("calculations-manual");
  return saved ? JSON.parse(saved) : [];
}

function Glyph({ name, size = 18 }) {
  const paths = {
    grid: <><rect x="3" y="3" width="7" height="7" rx="1.5" /><rect x="14" y="3" width="7" height="7" rx="1.5" /><rect x="3" y="14" width="7" height="7" rx="1.5" /><rect x="14" y="14" width="7" height="7" rx="1.5" /></>,
    search: <><circle cx="10.8" cy="10.8" r="6.8" /><path d="m16 16 4 4" /></>,
    download: <><path d="M12 3v12m0 0 4-4m-4 4-4-4" /><path d="M5 17v4h14v-4" /></>,
    plus: <><path d="M12 5v14M5 12h14" /></>,
    box: <><path d="m12 3 9 5-9 5-9-5 9-5Z" /><path d="m3 8 9 5 9-5M3 8v9l9 5 9-5V8M12 13v9" /></>,
    trend: <><path d="m3 17 6-6 4 4 8-9" /><path d="M15 6h6v6" /></>,
    layers: <><path d="m12 3 9 5-9 5-9-5 9-5Z" /><path d="m3 12 9 5 9-5M3 16l9 5 9-5" /></>,
    close: <><path d="m18 6-12 12M6 6l12 12" /></>,
    arrow: <><path d="M7 17 17 7M7 7h10v10" /></>,
    note: <><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><path d="M14 2v6h6M8 13h8M8 17h8" /></>,
  };
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name]}</svg>;
}

function MetricCard({ label, value, caption, icon, tone }) {
  return (
    <article className="metric-card">
      <div className={`metric-icon ${tone}`}><Glyph name={icon} size={19} /></div>
      <div className="metric-label">{label}</div>
      <div className="metric-value">{value}</div>
      <div className="metric-caption">{caption}</div>
    </article>
  );
}

function formatMoney(value) {
  return currency.format(Math.round(value || 0));
}

function toCsvCell(value) {
  const text = String(value ?? "");
  const safe = /^[=+\-@]/.test(text) ? `'${text}` : text;
  return `"${safe.replaceAll('"', '""')}"`;
}

function downloadCsv(rows) {
  const headers = ["Контрагент", "Товар", "SKU", "Період", "Обсяг, т", "Склад сировини по SKU", "Сировина разом, кг", "Ціна сировини за кг, грн", "Вартість сировини, грн", "Упаковка по SKU", "Упаковка разом, шт", "Ціна упаковки за шт, грн", "Вартість упаковки, грн", "Повні витрати на 1 т, грн", "Прибуток на 1 т, грн", "Рентабельність, %", "Статус", "Нотатка"];
  const content = [
    headers,
    ...rows.map(({ record, management }) => [
      record.client,
      record.product,
      record.product,
      record.period,
      record.metrics.volume,
      (record.materials?.items ?? []).filter((item) => item.type === "raw").map((item) => `${item.name}: ${item.quantity} кг × ${item.price ?? ""} грн/кг = ${item.cost ?? ""} грн`).join(" | "),
      record.materials?.rawQuantityKg ?? "",
      record.materials?.rawPricePerKg ?? "",
      record.materials?.rawCost ?? "",
      (record.materials?.items ?? []).filter((item) => item.type === "packaging").map((item) => `${item.name}: ${item.quantity} шт × ${item.price ?? ""} грн/шт = ${item.cost ?? ""} грн`).join(" | "),
      record.materials?.packagingQuantity ?? "",
      record.materials?.packagingPrice ?? "",
      record.materials?.packagingCost ?? "",
      record.metrics.fullCost,
      record.metrics.profit,
      record.metrics.margin,
      management.status,
      management.note,
    ]),
  ].map((line) => line.map(toCsvCell).join(";")).join("\r\n");
  const link = document.createElement("a");
  link.href = URL.createObjectURL(new Blob(["\uFEFF", content], { type: "text/csv;charset=utf-8" }));
  link.download = "kalkulyatsiyi.csv";
  link.click();
  URL.revokeObjectURL(link.href);
}

function parseExcelNumber(value) {
  if (typeof value === "number") return value;
  if (typeof value !== "string" || !value.trim()) return null;
  const parsed = Number(value.trim().replace(/\s/g, "").replace(",", "."));
  return Number.isFinite(parsed) ? parsed : null;
}

function parseCalculationWorkbook(file, workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error(`У файлі ${file.name} немає аркуша.`);

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const starts = rows
    .map((row, index) => (String(row[1] ?? "").trim() === "Калькуляція" ? index : -1))
    .filter((index) => index >= 0);
  if (!starts.length) throw new Error(`У файлі ${file.name} не знайдено розділів «Калькуляція».`);

  const client = file.name
    .replace(/\.xlsx$/i, "")
    .replace(/^Калькуляція\s*/i, "")
    .trim();

  return starts.map((start, calculationIndex) => {
    const end = starts[calculationIndex + 1] ?? rows.length;
    const section = rows.slice(start, end);
    const product = String(section[1]?.[1] ?? "").trim();
    if (!product) throw new Error(`У файлі ${file.name} не вдалося прочитати назву товару.`);

    const valuesByLabel = {};
    const costItems = [];
    section.forEach((row) => {
      const label = String(row[1] ?? "").trim();
      const cleanLabel = label.replace(/:$/, "").trim();
      const value = parseExcelNumber(row[3]);
      if (cleanLabel) valuesByLabel[cleanLabel] = value;
      if (label.startsWith("23 ") && value !== null) {
        costItems.push({ label, value });
      }
    });

    const metrics = {
      volume: valuesByLabel["Випуск продукції"] ?? null,
      productionCost: valuesByLabel["Виробнича собівартість"] ?? null,
      fullCost: valuesByLabel["Повні витрати"] ?? null,
      profit: valuesByLabel["Прибуток (+) Збитки (-)"] ?? null,
      margin: valuesByLabel["Рентабельність"] ?? null,
      wholesalePrice: valuesByLabel["Оптова вартість"] ?? null,
      retailPrice: valuesByLabel["Відпускна вартість"] ?? null,
    };
    if ([metrics.volume, metrics.fullCost, metrics.profit].some((value) => value === null)) {
      throw new Error(`У файлі ${file.name} не вдалося прочитати обов'язкові показники: ${product}`);
    }

    const period = String(section[5]?.[1] ?? "").split(" - ")[0].trim();
    const entity = String(section[2]?.[2] ?? "").trim();
    const stem = file.name.replace(/\.xlsx$/i, "");
    let rawQuantityKg = 0;
    let rawCost = 0;
    let packagingQuantity = 0;
    let packagingCost = 0;
    let hasRawQuantity = false;
    let hasPackagingQuantity = false;
    let currentCategory = "";
    let currentItem = "";
    let pendingItem = null;
    const materialItems = [];
    section.forEach((row) => {
      const label = String(row[1] ?? "").trim();
      const unit = String(row[2] ?? "").trim().toLocaleLowerCase("uk");
      const value = parseExcelNumber(row[3]);
      if (label.startsWith("23 ")) {
        currentCategory = label;
        currentItem = "";
        pendingItem = null;
      } else if (label === "Количество" && value !== null) {
        const isPackaging = currentCategory.toLocaleLowerCase("uk").includes("упаков");
        if (currentItem) {
          pendingItem = {
            name: currentItem,
            type: isPackaging ? "packaging" : "raw",
            quantity: value,
            unit: isPackaging ? "шт" : "кг",
            price: null,
            cost: null,
          };
        }
        if (isPackaging) {
          packagingQuantity += value;
          hasPackagingQuantity = true;
        } else if (unit === "кг") {
          rawQuantityKg += value;
          hasRawQuantity = true;
        }
      } else if (label === "Цена" && value !== null && pendingItem) {
        pendingItem.price = value;
      } else if (label === "Сумма" && value !== null) {
        if (currentCategory.toLocaleLowerCase("uk").includes("упаков")) {
          packagingCost += value;
        } else if (/^23 /.test(currentCategory) && !/(упаков|оплата праці|єсв|газ|електро|амортизац|енерг|збут|адміністратив)/i.test(currentCategory)) {
          rawCost += value;
        }
        if (pendingItem) {
          pendingItem.cost = value;
          materialItems.push(pendingItem);
          pendingItem = null;
        }
      } else if (label && !label.startsWith("24 ") && !label.startsWith("Показник")) {
        const original = String(row[1] ?? "");
        if (/^\s/.test(original) || original.startsWith("   ")) currentItem = original.trim();
        else if (!label.startsWith("23 ")) currentItem = "";
      }
    });
    const materials = {
      items: materialItems,
      rawQuantityKg: hasRawQuantity ? rawQuantityKg : null,
      rawCost: rawCost || null,
      rawPricePerKg: hasRawQuantity && rawQuantityKg ? rawCost / rawQuantityKg : null,
      packagingQuantity: hasPackagingQuantity ? packagingQuantity : null,
      packagingUnit: "шт",
      packagingCost: packagingCost || null,
      packagingPrice: hasPackagingQuantity && packagingQuantity ? packagingCost / packagingQuantity : null,
    };

    return {
      id: `${stem}-${calculationIndex + 1}`,
      client,
      entity,
      product,
      period,
      source: file.name,
      metrics,
      materials,
      costItems: costItems
        .sort((a, b) => Math.abs(b.value) - Math.abs(a.value))
        .slice(0, 8),
    };
  });
}

function DetailsPanel({ item, onClose, updateManagement }) {
  if (!item) return null;
  const { record, status, note } = item;
  const metrics = record.metrics;
  return (
    <div className="overlay" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <aside className="details-panel" aria-label="Деталі калькуляції">
        <div className="panel-top">
          <div className="eyebrow">Деталі калькуляції</div>
          <button className="icon-button" onClick={onClose} aria-label="Закрити"><Glyph name="close" /></button>
        </div>
        <div className="panel-client"><span className="client-mark">{record.client.slice(0, 2).toUpperCase()}</span><div><strong>{record.client}</strong><span>{record.entity || "Контрагент"}</span></div></div>
        <h2 className="panel-product">{record.product}</h2>
        <div className="period-chip">{record.period || "Період не вказано"}</div>
        <div className="detail-grid">
          <div><span>Випуск за період</span><strong>{quantity.format(metrics.volume)} т</strong></div>
          <div><span>Повні витрати / 1 т</span><strong>{formatMoney(metrics.fullCost)}</strong></div>
          <div><span>Прибуток / 1 т</span><strong className={metrics.profit < 0 ? "negative" : "positive"}>{formatMoney(metrics.profit)}</strong></div>
          <div><span>Рентабельність</span><strong className={metrics.margin < 0 ? "negative" : "positive"}>{decimal.format(metrics.margin)}%</strong></div>
          <div><span>Оптова вартість / 1 т</span><strong>{formatMoney(metrics.wholesalePrice)}</strong></div>
          <div><span>Відпускна вартість / 1 т</span><strong>{formatMoney(metrics.retailPrice)}</strong></div>
        </div>
        <section className="panel-section">
          <div className="section-title">Основні статті витрат <span>на 1 т</span></div>
          {record.costItems.length ? record.costItems.map((cost) => (
            <div className="cost-line" key={cost.label}><span>{cost.label.replace(/^\d+\s*/, "")}</span><strong>{formatMoney(cost.value)}</strong></div>
          )          ) : <p className="empty-copy">Немає деталізації статей витрат.</p>}
        </section>
        <section className="panel-section">
          <div className="section-title">Сировина та упаковка <span>по SKU</span></div>
          {record.materials?.items?.length ? (
            <div className="materials-table" role="table" aria-label="Склад сировини та упаковки">
              <div className="materials-table-head" role="row">
                <span role="columnheader">Найменування</span>
                <span role="columnheader">Кількість</span>
                <span role="columnheader">Ціна</span>
                <span role="columnheader">Вартість</span>
              </div>
              {record.materials.items.map((item, index) => (
                <div className="materials-table-row" role="row" key={`${item.name}-${index}`}>
                  <span role="cell"><span className={`material-type ${item.type}`}>{item.type === "packaging" ? "Упак." : "Сир."}</span>{item.name}</span>
                  <strong role="cell">{decimal.format(item.quantity)} {item.unit}</strong>
                  <strong role="cell">{item.price != null ? `${decimal.format(item.price)} ₴/${item.unit}` : "—"}</strong>
                  <strong role="cell">{item.cost != null ? formatMoney(item.cost) : "—"}</strong>
                </div>
              ))}
            </div>
          ) : <p className="empty-copy">Немає деталізації сировини та упаковки.</p>}
          {record.materials?.rawQuantityKg != null ? (
            <div className="material-summary">
              <div><span>Сировина</span><strong>{decimal.format(record.materials.rawQuantityKg)} кг</strong></div>
              <div><span>Ціна за кг</span><strong>{formatMoney(record.materials.rawPricePerKg)}</strong></div>
              <div><span>Вартість сировини</span><strong>{formatMoney(record.materials.rawCost)}</strong></div>
            </div>
          ) : <p className="empty-copy">Немає кількості сировини в кг.</p>}
          {record.materials?.packagingQuantity != null ? (
            <div className="material-summary packaging-summary">
              <div><span>Упаковка</span><strong>{quantity.format(record.materials.packagingQuantity)} шт</strong></div>
              <div><span>Ціна за шт</span><strong>{formatMoney(record.materials.packagingPrice)}</strong></div>
              <div><span>Вартість упаковки</span><strong>{formatMoney(record.materials.packagingCost)}</strong></div>
            </div>
          ) : <p className="empty-copy">Немає кількості упаковки.</p>}
        </section>
        <section className="panel-section management-section">
          <label className="field-label" htmlFor="detail-status">Статус</label>
          <select id="detail-status" value={status} onChange={(event) => updateManagement(record.id, "status", event.target.value)}>
            {STATUSES.map((option) => <option key={option}>{option}</option>)}
          </select>
          <label className="field-label note-label" htmlFor="detail-note">Нотатка</label>
          <textarea id="detail-note" value={note} onChange={(event) => updateManagement(record.id, "note", event.target.value)} placeholder="Додайте коментар або наступний крок…" rows="3" />
          <div className="save-hint"><span className="save-dot" /> Зберігається автоматично у цьому браузері</div>
        </section>
        <div className="source-note"><Glyph name="note" size={15} /> Джерело: {record.source}</div>
      </aside>
    </div>
  );
}

function AddCalculationPanel({ onClose, onAdd }) {
  function submitCalculation(event) {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const text = (name) => String(formData.get(name) ?? "").trim();
    const numeric = (name) => Number(formData.get(name));
    onAdd({
      id: `manual-${crypto.randomUUID()}`,
      client: text("client"),
      entity: text("entity"),
      product: text("product"),
      period: text("period"),
      source: "Додано вручну",
      metrics: {
        volume: numeric("volume"),
        productionCost: numeric("productionCost"),
        fullCost: numeric("fullCost"),
        profit: numeric("profit"),
        margin: numeric("margin"),
        wholesalePrice: numeric("wholesalePrice"),
        retailPrice: numeric("retailPrice"),
      },
      costItems: [],
      materials: null,
    });
  }

  return (
    <div className="overlay" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <aside className="details-panel add-panel" aria-label="Додати калькуляцію">
        <div className="panel-top">
          <div><div className="eyebrow">НОВИЙ ЗАПИС</div><h2 className="add-title">Додати калькуляцію</h2></div>
          <button className="icon-button" onClick={onClose} aria-label="Закрити"><Glyph name="close" /></button>
        </div>
        <p className="add-description">Запис збережеться у цьому браузері та з’явиться в загальному переліку й експорті CSV.</p>
        <form className="add-form" onSubmit={submitCalculation}>
          <label className="form-field">Контрагент<input name="client" required autoFocus /></label>
          <label className="form-field"><span className="form-field-heading">Юридична особа <span>(необов’язково)</span></span><input name="entity" /></label>
          <label className="form-field form-field-wide">Товар<input name="product" required /></label>
          <label className="form-field form-field-wide"><span className="form-field-heading">Період <span>(необов’язково)</span></span><input name="period" placeholder="Наприклад, жовтень 2026" /></label>
          <div className="form-divider">Показники калькуляції</div>
          <label className="form-field">Випуск, т<input name="volume" type="number" min="0" step="any" required /></label>
          <label className="form-field">Виробнича собівартість, грн<input name="productionCost" type="number" min="0" step="any" required /></label>
          <label className="form-field">Повні витрати, грн<input name="fullCost" type="number" min="0" step="any" required /></label>
          <label className="form-field">Прибуток / збиток, грн<input name="profit" type="number" step="any" required /></label>
          <label className="form-field">Рентабельність, %<input name="margin" type="number" step="any" required /></label>
          <label className="form-field">Оптова вартість, грн<input name="wholesalePrice" type="number" min="0" step="any" required /></label>
          <label className="form-field">Відпускна вартість, грн<input name="retailPrice" type="number" min="0" step="any" required /></label>
          <div className="add-form-actions">
            <button type="button" className="secondary-button" onClick={onClose}>Скасувати</button>
            <button type="submit" className="primary-button"><Glyph name="plus" size={16} /> Додати калькуляцію</button>
          </div>
        </form>
      </aside>
    </div>
  );
}

function App() {
  const [sourceRecords, setSourceRecords] = useState(null);
  const [loadError, setLoadError] = useState("");
  const [management, setManagement] = useState(readManagement);
  const [manualRecords, setManualRecords] = useState(readManualRecords);
  const [importMessage, setImportMessage] = useState("");
  const [isImporting, setIsImporting] = useState(false);
  const [query, setQuery] = useState("");
  const [clientFilter, setClientFilter] = useState("Усі контрагенти");
  const [statusFilter, setStatusFilter] = useState("Усі статуси");
  const [page, setPage] = useState(1);
  const [selectedId, setSelectedId] = useState(null);
  const [isAdding, setIsAdding] = useState(false);
  const records = useMemo(() => sourceRecords ? [...sourceRecords, ...manualRecords] : EMPTY_RECORDS, [manualRecords, sourceRecords]);
  const pageSize = 10;

  useEffect(() => {
    fetch("/data/calculations.json")
      .then((response) => {
        if (!response.ok) throw new Error(`Не вдалося завантажити калькуляції (${response.status}).`);
        return response.json();
      })
      .then((payload) => {
        if (!Array.isArray(payload.records)) throw new Error("Файл даних має невірний формат.");
        setSourceRecords(payload.records);
      })
      .catch((error) => setLoadError(error.message));
  }, []);

  useEffect(() => {
    localStorage.setItem("calculations-management", JSON.stringify(management));
  }, [management]);

  useEffect(() => {
    localStorage.setItem("calculations-manual", JSON.stringify(manualRecords));
  }, [manualRecords]);

  const clients = useMemo(() => [...new Set(records.map((record) => record.client))].sort((a, b) => a.localeCompare(b, "uk")), [records]);
  const managedRecords = useMemo(() => records.map((record) => ({
    record,
    management: { ...DEFAULT_MANAGEMENT, ...management[record.id] },
  })), [management, records]);
  const filtered = useMemo(() => {
    const normalizedQuery = query.trim().toLocaleLowerCase("uk");
    return managedRecords.filter(({ record, management: itemManagement }) => {
      const matchesQuery = !normalizedQuery || `${record.product} ${record.client} ${record.entity}`.toLocaleLowerCase("uk").includes(normalizedQuery);
      return matchesQuery
        && (clientFilter === "Усі контрагенти" || record.client === clientFilter)
        && (statusFilter === "Усі статуси" || itemManagement.status === statusFilter);
    });
  }, [clientFilter, managedRecords, query, statusFilter]);
  const summaries = useMemo(() => clients.map((client) => {
    const rows = records.filter((record) => record.client === client);
    return {
      client,
      count: rows.length,
      volume: rows.reduce((total, record) => total + record.metrics.volume, 0),
      profit: rows.reduce((total, record) => total + record.metrics.profit * record.metrics.volume, 0),
    };
  }).sort((a, b) => b.volume - a.volume), [clients, records]);
  const totals = useMemo(() => records.reduce((sum, record) => ({
    volume: sum.volume + record.metrics.volume,
    cost: sum.cost + record.metrics.fullCost * record.metrics.volume,
    profit: sum.profit + record.metrics.profit * record.metrics.volume,
  }), { volume: 0, cost: 0, profit: 0 }), [records]);
  const pageCount = Math.max(1, Math.ceil(filtered.length / pageSize));
  const visibleRows = filtered.slice((page - 1) * pageSize, page * pageSize);
  const selected = selectedId ? managedRecords.find(({ record }) => record.id === selectedId) : null;
  const maxVolume = Math.max(...summaries.map((summary) => summary.volume), 1);
  const needsAttention = Object.values(management).filter((item) => item?.status === "На перевірці").length;

  if (loadError) return <div className="loading-state error-state"><strong>Не вдалося відкрити калькуляції</strong><span>{loadError}</span><span>Перевірте файл даних і перезапустіть локальний сервер.</span></div>;
  if (!sourceRecords) return <div className="loading-state">Завантажую калькуляції…</div>;

  function updateManagement(id, field, value) {
    setManagement((current) => ({
      ...current,
      [id]: { ...DEFAULT_MANAGEMENT, ...current[id], [field]: value },
    }));
  }

  function updateFilter(setter, value) {
    setter(value);
    setPage(1);
  }

  function addManualRecord(record) {
    setManualRecords((current) => [...current, record]);
    setQuery("");
    setClientFilter("Усі контрагенти");
    setStatusFilter("Усі статуси");
    setPage(Math.ceil((records.length + 1) / pageSize));
    setIsAdding(false);
    setSelectedId(record.id);
  }

  async function importCalculationFiles(event) {
    const files = Array.from(event.currentTarget.files ?? []);
    if (!files.length) return;

    setIsImporting(true);
    setImportMessage("");
    try {
      const imported = [];
      for (const file of files) {
        if (!/\.xlsx$/i.test(file.name)) {
          throw new Error(`Підтримуються лише файли .xlsx: ${file.name}`);
        }
        const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
        imported.push(...parseCalculationWorkbook(file, workbook));
      }

      setManualRecords((current) => {
        const importedById = new Map(imported.map((record) => [record.id, record]));
        const updated = current.map((record) => importedById.get(record.id) ?? record);
        const existingIds = new Set(updated.map((record) => record.id));
        return [...updated, ...imported.filter((record) => !existingIds.has(record.id))];
      });
      setQuery("");
      setClientFilter("Усі контрагенти");
      setStatusFilter("Усі статуси");
      setImportMessage(`Імпортовано ${imported.length} калькуляцій.`);
    } catch (error) {
      setImportMessage(error instanceof Error ? error.message : "Не вдалося імпортувати файл.");
    } finally {
      setIsImporting(false);
      event.currentTarget.value = "";
    }
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <a className="brand" href="#" aria-label="Калькуляції — головна"><span className="brand-symbol"><Glyph name="layers" size={21} /></span><span>calc<span className="brand-dot">.</span>desk</span></a>
        <div className="workspace-label">РОБОЧИЙ ПРОСТІР</div>
        <button className="nav-item active"><Glyph name="grid" size={17} /> Огляд <span className="nav-count">{records.length}</span></button>
        <button className="nav-item" onClick={() => document.getElementById("calculations")?.scrollIntoView({ behavior: "smooth" })}><Glyph name="box" size={17} /> Калькуляції</button>
        <div className="sidebar-divider" />
        <div className="sidebar-summary">
          <div className="workspace-label">КОНТРАГЕНТИ <span>{clients.length}</span></div>
          {summaries.map((summary) => (
            <button key={summary.client} className={`client-nav ${clientFilter === summary.client ? "chosen" : ""}`} onClick={() => { setClientFilter(clientFilter === summary.client ? "Усі контрагенти" : summary.client); setPage(1); }}>
              <span className="client-mark small">{summary.client.slice(0, 2).toUpperCase()}</span><span>{summary.client}</span><span className="client-total">{summary.count}</span>
            </button>
          ))}
        </div>
        <div className="sidebar-bottom"><span className="online-dot" /> Локальний робочий простір <small>Дані зберігаються на пристрої</small></div>
      </aside>

      <main className="main-content">
        <header className="topbar"><div className="breadcrumb">Робочий простір <span>/</span> <strong>Огляд</strong></div><div className="topbar-right"><span className="data-period">Період даних: <strong>Вересень 2026</strong></span><span className="avatar">К</span></div></header>
        <div className="page-content">
          <section className="welcome-row">
            <div><div className="eyebrow">ВІТАЄМО У CALC.DESK</div><h1>Огляд калькуляцій</h1><p>Уся собівартість, прибутковість і робота з контрагентами — в одному місці.</p></div>
            <div className="welcome-actions">
              <button className="primary-button" onClick={() => setIsAdding(true)} aria-label="Додати калькуляцію вручну" title="Додати калькуляцію вручну"><Glyph name="plus" size={17} /> Додати вручну</button>
              <label className={`secondary-button ${isImporting ? "disabled-button" : ""}`}>
                <input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" multiple onChange={importCalculationFiles} disabled={isImporting} hidden />
                <Glyph name="download" size={17} /> {isImporting ? "Імпортую…" : "Імпорт з Excel"}
              </label>
              <button className="secondary-button" onClick={() => downloadCsv(filtered)} aria-label="Експортувати CSV" title="Експортувати CSV"><Glyph name="download" size={17} /> Експортувати CSV</button>
            </div>
          </section>
          {importMessage && <div className={`import-message ${importMessage.startsWith("Імпортовано") ? "import-success" : "import-error"}`}>{importMessage}</div>}
          <section className="metrics-grid" aria-label="Ключові показники">
            <MetricCard label="Калькуляцій у базі" value={records.length} caption={`з ${clients.length} контрагентів`} icon="layers" tone="blue" />
            <MetricCard label="Випуск за період" value={`${quantity.format(totals.volume)} т`} caption="Сумарний обсяг виробництва" icon="box" tone="teal" />
            <MetricCard label="Повні витрати" value={formatMoney(totals.cost)} caption="Обсяг × витрати на 1 тонну" icon="grid" tone="amber" />
            <MetricCard label="Прибуток / збиток" value={formatMoney(totals.profit)} caption={`${needsAttention} на перевірці`} icon="trend" tone={totals.profit < 0 ? "red" : "green"} />
          </section>

          <section className="insights-grid">
            <article className="surface chart-card">
              <div className="card-heading"><div><h2>Випуск за контрагентами</h2><p>Загальний обсяг виробництва за період</p></div><span className="heading-icon"><Glyph name="trend" size={18} /></span></div>
              <div className="bars">
                {summaries.map((summary, index) => (
                  <button className="bar-row" key={summary.client} onClick={() => { setClientFilter(summary.client); setPage(1); document.getElementById("calculations")?.scrollIntoView({ behavior: "smooth" }); }}>
                    <span className="bar-name">{summary.client}</span><span className="bar-track"><span className={`bar-fill bar-${index % 5}`} style={{ width: `${Math.max(2, (summary.volume / maxVolume) * 100)}%` }} /></span><strong>{quantity.format(summary.volume)} т</strong>
                  </button>
                ))}
              </div>
            </article>
            <article className="surface portfolio-card">
              <div className="card-heading"><div><h2>Портфель калькуляцій</h2><p>Стан роботи по кожному контрагенту</p></div><button className="text-button" onClick={() => { setStatusFilter("Усі статуси"); document.getElementById("calculations")?.scrollIntoView({ behavior: "smooth" }); }}>Детальніше <Glyph name="arrow" size={14} /></button></div>
              <div className="portfolio-list">
                {summaries.map((summary) => {
                  const completed = summary.count - managedRecords.filter(({ record, management: item }) => record.client === summary.client && item.status !== "Погоджено").length;
                  return <button className="portfolio-row" key={summary.client} onClick={() => { setClientFilter(summary.client); setPage(1); document.getElementById("calculations")?.scrollIntoView({ behavior: "smooth" }); }}>
                    <span className="client-mark small">{summary.client.slice(0, 2).toUpperCase()}</span><span className="portfolio-name">{summary.client}<small>{summary.count} калькуляцій</small></span><span className="portfolio-progress"><span><i style={{ width: `${summary.count ? (completed / summary.count) * 100 : 0}%` }} /></span><small>{completed}/{summary.count}</small></span>
                  </button>;
                })}
              </div>
            </article>
          </section>

          <section className="surface calculations-card" id="calculations">
            <div className="table-heading"><div><div className="table-title"><h2>Усі калькуляції</h2><span className="result-count">{filtered.length}</span></div><p>Переглядайте показники та оновлюйте статуси роботи</p></div><button className="mobile-export" onClick={() => downloadCsv(filtered)} aria-label="Експортувати CSV"><Glyph name="download" /></button></div>
            <div className="filters">
              <label className="search-field"><Glyph name="search" size={17} /><input value={query} onChange={(event) => updateFilter(setQuery, event.target.value)} placeholder="Пошук за товаром або контрагентом" /></label>
              <select aria-label="Фільтр контрагентів" value={clientFilter} onChange={(event) => updateFilter(setClientFilter, event.target.value)}><option>Усі контрагенти</option>{clients.map((client) => <option key={client}>{client}</option>)}</select>
              <select aria-label="Фільтр статусу" value={statusFilter} onChange={(event) => updateFilter(setStatusFilter, event.target.value)}><option>Усі статуси</option>{STATUSES.map((status) => <option key={status}>{status}</option>)}</select>
            </div>
            <div className="table-scroll"><table><thead><tr><th>ТОВАР / SKU</th><th>КОНТРАГЕНТ</th><th>СИРОВИНА</th><th>УПАКОВКА</th><th>ВИПУСК, Т</th><th>ПОВНІ ВИТРАТИ / 1 Т</th><th>ПРИБУТОК / 1 Т</th><th>СТАТУС</th><th aria-label="Нотатка" /></tr></thead>
              <tbody>{visibleRows.map(({ record, management: itemManagement }) => (
                <tr key={record.id} onClick={() => setSelectedId(record.id)} tabIndex={0} onKeyDown={(event) => event.key === "Enter" && setSelectedId(record.id)}>
                  <td><div className="product-cell"><strong>{record.product}</strong><span>{record.period || record.source}</span></div></td>
                  <td><span className="table-client"><span className="client-mark tiny">{record.client.slice(0, 1)}</span>{record.client}</span></td>
                  <td className="number-cell">{record.materials?.rawQuantityKg != null ? <><strong>{decimal.format(record.materials.rawQuantityKg)} кг</strong><span className="material-price">{formatMoney(record.materials.rawCost)} · {formatMoney(record.materials.rawPricePerKg)} / кг</span></> : "—"}</td>
                  <td className="number-cell">{record.materials?.packagingQuantity != null ? <><strong>{quantity.format(record.materials.packagingQuantity)} шт</strong><span className="material-price">{formatMoney(record.materials.packagingCost)} · {formatMoney(record.materials.packagingPrice)} / шт</span></> : "—"}</td>
                  <td className="number-cell">{quantity.format(record.metrics.volume)}</td>
                  <td className="number-cell">{amount.format(record.metrics.fullCost)} ₴</td>
                  <td className={`number-cell ${record.metrics.profit < 0 ? "negative" : "positive"}`}>{record.metrics.profit < 0 ? "−" : "+"}{amount.format(Math.abs(record.metrics.profit))} ₴</td>
                  <td onClick={(event) => event.stopPropagation()}><select className={`status-select ${STATUS_CLASS[itemManagement.status] || "status-new"}`} value={itemManagement.status} onChange={(event) => updateManagement(record.id, "status", event.target.value)} aria-label={`Статус: ${record.product}`}>{STATUSES.map((status) => <option key={status}>{status}</option>)}</select></td>
                  <td>{itemManagement.note && <span className="note-indicator" title="Є нотатка"><Glyph name="note" size={15} /></span>}</td>
                </tr>
              ))}
              {!visibleRows.length && <tr><td className="empty-state" colSpan="9"><span className="empty-search"><Glyph name="search" size={21} /></span><strong>Нічого не знайдено</strong><span>Змініть пошуковий запит або фільтри.</span></td></tr>}
              </tbody></table></div>
            <div className="table-footer"><span>Показано {filtered.length ? (page - 1) * pageSize + 1 : 0}–{Math.min(page * pageSize, filtered.length)} з {filtered.length}</span><div className="pagination"><button disabled={page <= 1} onClick={() => setPage((current) => current - 1)}>Назад</button><span>Сторінка {page} з {pageCount}</span><button disabled={page >= pageCount} onClick={() => setPage((current) => current + 1)}>Далі</button></div></div>
          </section>
          <footer className="page-footer"><span>calc.desk <span>·</span> локальний дашборд калькуляцій</span><span>{records.length} записів <span>·</span> {clients.length} контрагентів</span></footer>
        </div>
      </main>
      <DetailsPanel item={selected} onClose={() => setSelectedId(null)} updateManagement={updateManagement} />
      {isAdding && <AddCalculationPanel onClose={() => setIsAdding(false)} onAdd={addManualRecord} />}
    </div>
  );
}

export default App;
