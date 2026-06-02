import * as XLSX from "xlsx";

const TARGET_SHEETS = ["ID003", "ID004", "ID005", "ID006"];
const EXPENSE_SHEET = "TOTALE SPESE Sandu Francesco";
const SPLITS = {
  ID003: { francesco: 0.20, sandu: 0.05, total: 0.25, airbnb_auto: true },
  ID004: { francesco: 0.15, sandu: 0.10, total: 0.25, airbnb_auto: true },
  ID005: { francesco: 0.05, sandu: 0.05, total: 0.10, airbnb_auto: false },
  ID006: { francesco: 0.10, sandu: 0.05, total: 0.15, airbnb_auto: false },
};

const norm = (value) => String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
const money = (value) => {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;
  let text = String(value).trim().replace(/[^\d,.-]/g, "");
  if (!text) return 0;
  if (text.includes(",") && text.includes(".")) text = text.replace(/\./g, "").replace(",", ".");
  else if (text.includes(",")) text = text.replace(",", ".");
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
};
const isoDate = (value) => {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return `${date.y}-${String(date.m).padStart(2, "0")}-${String(date.d).padStart(2, "0")}`;
  }
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  return text;
};
const parseDate = (value) => {
  const text = isoDate(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  const date = new Date(`${text}T00:00:00Z`);
  return Number.isNaN(date.getTime()) ? null : date;
};
const previousMonthEnd = () => {
  const now = new Date();
  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 0));
};

function sheetRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return {};
  const ref = sheet["!ref"];
  const out = {};
  if (!ref) return out;
  const range = XLSX.utils.decode_range(ref);
  for (let rowIdx = range.s.r; rowIdx <= range.e.r; rowIdx += 1) {
    const row = {};
    for (let colIdx = range.s.c; colIdx <= range.e.c; colIdx += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
      const cell = sheet[address];
      if (!cell || cell.v === undefined || cell.v === null || cell.v === "") continue;
      row[XLSX.utils.encode_col(colIdx)] = cell.v;
    }
    if (Object.keys(row).length) out[rowIdx + 1] = row;
  }
  return out;
}

function sheetByName(workbook, target) {
  const wanted = norm(target);
  const prefix = norm(target.slice(0, 5));
  return workbook.SheetNames.find((name) => norm(name) === wanted)
    || workbook.SheetNames.find((name) => norm(name).startsWith(prefix))
    || null;
}

function findCol(headers, candidates) {
  const normalized = Object.fromEntries(Object.entries(headers).map(([col, value]) => [col, norm(value)]));
  const wanted = candidates.map(norm);
  for (const candidate of wanted) {
    for (const [col, value] of Object.entries(normalized)) if (value === candidate) return col;
  }
  for (const candidate of wanted) {
    for (const [col, value] of Object.entries(normalized)) if (candidate && value.includes(candidate)) return col;
  }
  return null;
}

function sheetLabel(source, id) {
  let text = String(source || id).trim();
  if (text.startsWith(id)) text = text.slice(id.length).replace(/^[_ -]+/, "");
  text = text.replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  return text ? `${id} - ${text}` : id;
}

function channelName(value) {
  const key = norm(value);
  if (key.includes("airbnb")) return "Airbnb";
  if (key.includes("booking")) return "Booking";
  if (key.includes("dirett") || key.includes("direct") || key.includes("privat")) return "Dirette";
  return String(value || "").trim() || "Non indicato";
}

function parseApartmentSheet(id, rows, sourceSheet) {
  const label = sheetLabel(sourceSheet, id);
  const headerRows = Object.keys(rows).map(Number).filter((rn) => {
    const row = rows[rn];
    return findCol(row, ["Data IN", "Check in", "Check-in", "Arrivo"])
      && findCol(row, ["Data OUT", "Check out", "Check-out", "Partenza"])
      && findCol(row, ["Portale", "Canale", "Sito", "OTA", "Provenienza"]);
  });
  if (!headerRows.length) return { id, label, sourceSheet, bookings: [], summary: {}, warnings: ["Intestazione prenotazioni non trovata."] };

  const bookings = [];
  const split = SPLITS[id];
  headerRows.forEach((headerRow, idx) => {
    const nextHeader = headerRows[idx + 1] || Infinity;
    const headers = rows[headerRow];
    const colIn = findCol(headers, ["Data IN", "Check in", "Check-in", "Arrivo"]);
    const colOut = findCol(headers, ["Data OUT", "Check out", "Check-out", "Partenza"]);
    const colChannel = findCol(headers, ["Portale", "Canale", "Sito", "OTA", "Provenienza"]);
    const colGross = findCol(headers, ["Lordo", "Totale", "Importo", "Incasso", "Prezzo", "Pagamento"]);
    const colOta = findCol(headers, ["% OTA", "Commissioni OTA", "Commissione OTA", "Fee OTA", "OTA"]);
    const colCleaning = findCol(headers, ["Pulizie", "Pulizia", "Cleaning"]);
    const colNet = findCol(headers, ["Netto proprieta", "Netto proprietà", "Netto Prop", "Proprietà"]);
    const colFrancesco = id === "ID005" ? findCol(headers, ["Gestione PM - Francesco", "Gestione Francesco"]) : findCol(headers, ["Francesco"]);
    const colSandu = id === "ID005" ? null : findCol(headers, ["Sandu"]);

    Object.keys(rows).map(Number).sort((a, b) => a - b).forEach((rn) => {
      if (rn <= headerRow || rn >= nextHeader) return;
      const row = rows[rn];
      const channel = channelName(row[colChannel]);
      if (!["Airbnb", "Booking", "Dirette"].includes(channel)) return;
      const gross = money(row[colGross]);
      const ota = money(row[colOta]);
      const cleaning = money(row[colCleaning]);
      let netOwner = money(row[colNet]);
      if (!netOwner && gross) netOwner = gross - ota - cleaning;
      let francesco = money(row[colFrancesco]);
      let sandu = id === "ID005" && colFrancesco ? francesco / 2 : money(row[colSandu]);
      if (id === "ID005" && colFrancesco) francesco = sandu;
      if (!francesco && !sandu && netOwner && !(id === "ID005" && colFrancesco)) {
        francesco = netOwner * split.francesco;
        sandu = netOwner * split.sandu;
      }
      bookings.push({
        sheet: id,
        apartmentLabel: label,
        row: rn,
        checkIn: isoDate(row[colIn]),
        checkOut: isoDate(row[colOut]),
        channel,
        gross,
        ota,
        cleaning,
        netOwner,
        francesco,
        sandu,
        sanduDueByFrancesco: channel === "Airbnb" && split.airbnb_auto ? 0 : sandu,
      });
    });
  });
  return { id, label, sourceSheet, bookings, summary: {}, warnings: [] };
}

function parseExpenseGroups(rows) {
  const titleInfo = (value) => {
    const text = String(value || "").trim();
    const match = text.match(/(20\d{2})/);
    if (!text || !match) return null;
    const upper = text.toUpperCase();
    if (!upper.includes("ADS") && !upper.includes("SPESE")) return null;
    if (upper.includes("TOTALE SPESE")) return null;
    const kind = upper.includes("ADS") ? "ADS" : "APPARTAMENTO";
    const name = text.replace(/\s*-\s*20\d{2}\s*$/i, "").replace("SPESE", "").replace("ADS", "ADS").trim().replace(/^[- ]+|[- ]+$/g, "") || "ADS";
    return { title: text, year: Number(match[1]), kind, name };
  };
  const parseSide = (first, desc, totalCol, shareCol, noteCol) => {
    const groups = [];
    let current = null;
    Object.keys(rows).map(Number).sort((a, b) => a - b).forEach((rn) => {
      const row = rows[rn];
      const info = titleInfo(row[first]);
      if (info) {
        current = { ...info, items: [], total: 0, share: 0 };
        groups.push(current);
        return;
      }
      if (!current) return;
      const category = String(row[first] || "").trim();
      if (!category || norm(category) === "cosa") return;
      if (norm(category) === "totale") {
        current.total = money(row[totalCol]) || current.total;
        current.share = money(row[shareCol]) || current.share;
        current = null;
        return;
      }
      const item = { row: rn, category, description: String(row[desc] || "").trim(), total: money(row[totalCol]), share: money(row[shareCol]), note: String(row[noteCol] || "").trim() };
      if (item.description || item.total || item.share) {
        current.items.push(item);
        current.total += item.total;
        current.share += item.share;
      }
    });
    return groups;
  };
  const groups = [...parseSide("AH", "AI", "AJ", "AK", "AL"), ...parseSide("AN", "AO", "AP", "AQ", "AR")];
  const items = [];
  groups.forEach((group) => group.items.forEach((item) => items.push({ ...item, year: group.year, kind: group.kind, group: group.name, francescoShare: item.share, sanduShare: item.share })));
  return { sections: {}, groups, items, totals: {} };
}

function parseDueBlocks(rows, side) {
  const cols = side === "francesco_to_sandu"
    ? { head: "B", category: "B", description: "C", checkIn: "D", checkOut: "E", amount: "F", note: "G", paid: "H" }
    : { head: "T", category: "T", description: "U", checkIn: "V", checkOut: "W", amount: "X", note: "Y", paid: "Z" };
  const blocks = [];
  let current = null;
  Object.keys(rows).map(Number).sort((a, b) => a - b).forEach((rn) => {
    const row = rows[rn];
    const head = String(row[cols.head] || "").trim();
    const match = head.match(/(20\d{2})/);
    if (match && head.toUpperCase().includes("SPESE")) {
      current = { year: Number(match[1]), title: head, side, items: [], total: 0 };
      blocks.push(current);
      return;
    }
    if (!current) return;
    const category = String(row[cols.category] || "").trim();
    if (!category || norm(category) === "cosa") return;
    const amount = money(row[cols.amount]);
    if (norm(category) === "totale") {
      current.total = amount || current.total;
      current = null;
      return;
    }
    const item = { row: rn, year: current.year, side, category, description: String(row[cols.description] || "").trim(), checkIn: isoDate(row[cols.checkIn]), checkOut: isoDate(row[cols.checkOut]), amount, note: String(row[cols.note] || "").trim(), paid: String(row[cols.paid] || "").trim() };
    if (item.amount || item.description || item.note) {
      current.items.push(item);
      current.total += amount;
    }
  });
  return blocks;
}

function parsePayments(rows) {
  const payments = [];
  let current = null;
  Object.keys(rows).map(Number).sort((a, b) => a - b).forEach((rn) => {
    const label = String(rows[rn].AB || "").trim();
    const key = norm(label);
    const yearMatch = label.match(/(20\d{2})/);
    if (key.includes("pagamenti") && yearMatch) {
      const side = key.includes("francescosandu") ? "francesco_to_sandu" : key.includes("sandufrancesco") ? "sandu_to_francesco" : null;
      current = side ? { year: Number(yearMatch[1]), side, title: label } : null;
      return;
    }
    if (!current) return;
    if (key === "data" || key === "totale") {
      if (key === "totale") current = null;
      return;
    }
    const amount = money(rows[rn].AC);
    if (!amount) return;
    payments.push({
      row: rn,
      year: current.year,
      side: current.side,
      amount,
      label: current.title,
      date: isoDate(rows[rn].AB),
      note: String(rows[rn].AD || "").trim(),
    });
  });
  return payments;
}

function addBucket(years, year, side, group, amount) {
  const key = String(year);
  const entry = years[key] ||= { year, francescoToSandu: {}, sanduToFrancesco: {}, francescoToSanduTotal: 0, sanduToFrancescoTotal: 0, delta: 0 };
  const mapKey = side === "francesco_to_sandu" ? "francescoToSandu" : "sanduToFrancesco";
  const totalKey = side === "francesco_to_sandu" ? "francescoToSanduTotal" : "sanduToFrancescoTotal";
  entry[mapKey][group] = (entry[mapKey][group] || 0) + amount;
  entry[totalKey] += amount;
}

function finalizeYears(years) {
  const overall = { francescoToSandu: 0, sanduToFrancesco: 0, delta: 0 };
  Object.values(years).forEach((entry) => {
    entry.delta = entry.francescoToSanduTotal - entry.sanduToFrancescoTotal;
    overall.francescoToSandu += entry.francescoToSanduTotal;
    overall.sanduToFrancesco += entry.sanduToFrancescoTotal;
  });
  overall.delta = overall.francescoToSandu - overall.sanduToFrancesco;
  return { overall, list: Object.values(years).sort((a, b) => b.year - a.year) };
}

export function parseWorkbook(bytes, filename) {
  const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
  const warnings = [];
  const apartments = [];
  const bookings = [];
  for (const id of TARGET_SHEETS) {
    const source = sheetByName(workbook, id);
    const parsed = source ? parseApartmentSheet(id, sheetRows(workbook, source), source) : { id, label: id, sourceSheet: id, bookings: [], warnings: ["Foglio non trovato."] };
    apartments.push(parsed);
    bookings.push(...parsed.bookings);
  }
  const expenseSource = sheetByName(workbook, EXPENSE_SHEET);
  if (!expenseSource) warnings.push(`Foglio spese non trovato. Fogli presenti: ${workbook.SheetNames.join(", ")}`);
  const expenseRows = expenseSource ? sheetRows(workbook, expenseSource) : {};
  const expenses = parseExpenseGroups(expenseRows);
  if (expenseSource && !expenses.groups.length) warnings.push(`Foglio ${expenseSource} trovato, ma nessun gruppo spese letto nelle colonne AH:AR.`);
  const dueToSanduBlocks = parseDueBlocks(expenseRows, "francesco_to_sandu");
  const dueToFrancescoBlocks = parseDueBlocks(expenseRows, "sandu_to_francesco");
  const payments = parsePayments(expenseRows);

  const channels = {};
  bookings.forEach((booking) => {
    const bucket = channels[booking.channel] ||= { gross: 0, netOwner: 0, francesco: 0, sandu: 0, sanduDueByFrancesco: 0, count: 0 };
    bucket.count += 1;
    ["gross", "netOwner", "francesco", "sandu", "sanduDueByFrancesco"].forEach((key) => bucket[key] += Number(booking[key] || 0));
  });

  const cutoff = previousMonthEnd();
  const settled = {};
  const projected = {};
  bookings.forEach((booking) => {
    const date = parseDate(booking.checkIn);
    const amount = money(booking.sanduDueByFrancesco);
    if (!date || !amount) return;
    addBucket(date <= cutoff ? settled : projected, date.getUTCFullYear(), "francesco_to_sandu", "Utili appartamenti", amount);
  });
  dueToSanduBlocks.forEach((block) => block.items.forEach((item) => addBucket(settled, item.year, "francesco_to_sandu", item.category || "Sanremo", item.amount)));
  expenses.items.forEach((item) => {
    const group = norm(item.category) === "varie" ? "Varie" : item.kind === "ADS" ? "ADS" : "Spese appartamenti";
    addBucket(settled, item.year, "sandu_to_francesco", group, money(item.share));
  });
  dueToFrancescoBlocks.forEach((block) => block.items.forEach((item) => addBucket(settled, item.year, "sandu_to_francesco", "Tassa soggiorno Sanremo", item.amount)));
  payments.forEach((payment) => addBucket(settled, payment.year, payment.side, "Pagamenti avvenuti", -payment.amount));
  const settlementFinal = finalizeYears(settled);
  const projectionFinal = finalizeYears(projected);

  return {
    filename,
    uploadedAt: new Date().toISOString().replace(/\.\d{3}Z$/, "Z"),
    apartments,
    bookings,
    channels,
    expenses,
    settlement: {
      overall: settlementFinal.overall,
      cutoffDate: cutoff.toISOString().slice(0, 10),
      years: settlementFinal.list,
      projection: {
        overall: projectionFinal.overall,
        years: projectionFinal.list,
        note: "Proiezione su prenotazioni future: valori da confermare quando i soggiorni saranno conclusi o modificati.",
      },
      payments,
      dueToSanduBlocks,
      dueToFrancescoBlocks,
      netFrancescoPaysSandu: settlementFinal.overall.delta,
    },
    warnings,
    rules: SPLITS,
  };
}
