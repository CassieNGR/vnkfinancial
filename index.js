const path = require("path");
const fs = require("fs");
const express = require("express");
const cors = require("cors");
const multer = require("multer");
const db = require("./db");
const { sheetToRows, parseAgingRows, parseTxnRows, parseWorkbookBuffer, sheetToObjects, parseNGRRows, parseIFSRows, parseSalesDetailRows, parseQBSummaryRows, parseOpenSalesOrdersRows, parseOpenPurchaseOrdersRows, parseInventoryValuationRows, round2 } = require("./parse");

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 15 * 1024 * 1024 } });
app.use(cors());
app.use(express.json());

/* ---------------- one-time seed from July 2026 data ---------------- */
function seedIfEmpty() {
  const seeded = db.prepare("SELECT value FROM meta WHERE key='seeded'").get();
  if (seeded) return; // already seeded once (or deliberately reset) - never auto-reseed again
  const seedPath = path.join(__dirname, "seed", "seed.json");
  if (!fs.existsSync(seedPath)) { db.prepare("INSERT OR REPLACE INTO meta (key,value) VALUES ('seeded','true')").run(); return; }
  const seed = JSON.parse(fs.readFileSync(seedPath, "utf8"));
  const salesDetailSeed = seed._salesDetail || [];
  const ifsSeed = seed._ifsCommission || [];
  delete seed._ifsCommission;
  const qbSummarySeed = seed._qbSummary || {};
  delete seed._qbSummary;
  const inventorySeed = seed._inventoryValuation || {};
  delete seed._inventoryValuation;
  const openSoSeed = seed._openSalesOrders || {};
  delete seed._openSalesOrders;
  const openPoSeed = seed._openPurchaseOrders || {};
  delete seed._openPurchaseOrders;
  delete seed._salesDetail; // not a month key - handled separately below
  const insertSO = db.prepare("INSERT INTO sales_orders (date,entity,type,num,memo,amount) VALUES (?,?,?,?,?,?)");
  const insertPO = db.prepare("INSERT INTO purchase_orders (date,entity,type,num,memo,amount) VALUES (?,?,?,?,?,?)");
  const insertAR = db.prepare("INSERT INTO ar_items (date,name,total,current,d1_30,d31_60,d61_90,d90plus) VALUES (?,?,?,?,?,?,?,?)");
  const insertAP = db.prepare("INSERT INTO ap_items (date,name,total,current,d1_30,d31_60,d61_90,d90plus) VALUES (?,?,?,?,?,?,?,?)");
  const insertDetail = db.prepare("INSERT INTO sales_detail (date,customer,type,num,memo,item,qty,salesPrice,amount) VALUES (?,?,?,?,?,?,?,?,?)");
  const insertIFS = db.prepare("INSERT INTO ifs_commission (date,salesRep,customer,invoiceNum,factory,jobName,status,invoiceTotalAmount,invoicePaymentAmount,totalCommAmount,actualCommissionPct,factoryCommissionPct) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)");
  const insertQBSummary = db.prepare("INSERT INTO qb_sales_summary (asOfDate,customer,a,b,dollarChange,pctChange) VALUES (?,?,?,?,?,?)");
  const insertInventory = db.prepare("INSERT INTO inventory_valuation (asOfDate,item,onHand,avgCost,assetValue,pctOfTotalAsset,salesPrice,retailValue,pctOfTotalRetail) VALUES (?,?,?,?,?,?,?,?,?)");
  const insertOpenSO = db.prepare("INSERT INTO open_sales_orders (asOfDate,customer,type,date,orderNum,memo,amount,openBalance,prepaymentOpenBalance) VALUES (?,?,?,?,?,?,?,?,?)");
  const insertOpenPO = db.prepare("INSERT INTO open_purchase_orders (asOfDate,vendor,type,date,orderNum,memo,deliveryDate,amount,openBalance) VALUES (?,?,?,?,?,?,?,?,?)");
  const tx = db.transaction(() => {
    Object.values(seed).forEach((month) => {
      (month.so || []).forEach((t) => insertSO.run(t.date, t.entity, t.type, t.num, t.memo, t.amount));
      (month.po || []).forEach((t) => insertPO.run(t.date, t.entity, t.type, t.num, t.memo, t.amount));
      Object.entries(month.arSnapshots || {}).forEach(([date, snap]) => {
        snap.items.forEach((it) => insertAR.run(date, it.name, it.total, it.current, it.d1_30, it.d31_60, it.d61_90, it.d90plus));
      });
      Object.entries(month.apSnapshots || {}).forEach(([date, snap]) => {
        snap.items.forEach((it) => insertAP.run(date, it.name, it.total, it.current, it.d1_30, it.d31_60, it.d61_90, it.d90plus));
      });
    });
    salesDetailSeed.forEach((r) => insertDetail.run(r.date, r.customer, r.type, r.num, r.memo, r.item, r.qty, r.salesPrice, r.amount));
    ifsSeed.forEach((r) => insertIFS.run(r.date, r.salesRep, r.customer, r.invoiceNum, r.factory, r.jobName, r.status, r.invoiceTotalAmount, r.invoicePaymentAmount, r.totalCommAmount, r.actualCommissionPct, r.factoryCommissionPct));
    Object.entries(qbSummarySeed).forEach(([asOfDate, rows]) => rows.forEach((r) => insertQBSummary.run(asOfDate, r.customer, r.a, r.b, r.dollarChange, r.pctChange)));
    Object.entries(inventorySeed).forEach(([asOfDate, rows]) => rows.forEach((r) => insertInventory.run(asOfDate, r.item, r.onHand, r.avgCost, r.assetValue, r.pctOfTotalAsset, r.salesPrice, r.retailValue, r.pctOfTotalRetail)));
    Object.entries(openSoSeed).forEach(([asOfDate, rows]) => rows.forEach((r) => insertOpenSO.run(asOfDate, r.customer, r.type, r.date, r.orderNum, r.memo, r.amount, r.openBalance, r.prepaymentOpenBalance)));
    Object.entries(openPoSeed).forEach(([asOfDate, rows]) => rows.forEach((r) => insertOpenPO.run(asOfDate, r.vendor, r.type, r.date, r.orderNum, r.memo, r.deliveryDate, r.amount, r.openBalance)));
  });
  tx();
  const seedTime = Date.UTC(2026, 8, 13, 9, 0);
  const insHist = db.prepare("INSERT INTO upload_history (ts, filename, kind, detail) VALUES (?,?,?,?)");
  insHist.run(seedTime + 5 * 60000, "Detailed_Sales_Jan_to_Sept_13__2026.xlsx", "detail", `${salesDetailSeed.length} line items`);
  insHist.run(seedTime + 6 * 60000, "IFS_Commission_Reconcilliation_Report_Factory_Customer.csv", "ifs", `${ifsSeed.length} records`);
  insHist.run(seedTime + 4 * 60000, "Sales_Order_Jan_1_-_Sep_16__2026.xlsx", "so", "578 Sales Order records (other transaction types excluded)");
  insHist.run(seedTime + 3 * 60000, "Purchase_Order_Jan_1_-_Sep_16__2026.xlsx", "po", "539 Purchase Order records (other transaction types excluded)");
  insHist.run(seedTime + 2 * 60000, "Accounts_Receivable_Sept_13_2026.xlsx", "ar", "34 records for 2026-09-13");
  insHist.run(seedTime + 1 * 60000, "Accounts_Payable_Sept_13_2026.xlsx", "ap", "56 records for 2026-09-13");
  const sep17 = Date.UTC(2026, 8, 17, 9, 0);
  Object.entries(qbSummarySeed).forEach(([asOfDate, rows]) => insHist.run(sep17 + 4 * 60000, "Sales_by_Customer_Summary_Jan_1_-_Sep_17__2026.xlsx", "qbsummary", `${rows.length} customers as of ${asOfDate}`));
  Object.entries(inventorySeed).forEach(([asOfDate, rows]) => insHist.run(sep17 + 3 * 60000, "Inventory_Valuation_Summary_as_of_Sep_17__2026.xlsx", "inventory", `${rows.length} items as of ${asOfDate}`));
  Object.entries(openSoSeed).forEach(([asOfDate, rows]) => insHist.run(sep17 + 2 * 60000, "Open_Sales_Order_As_of_Sep_17__2026.xlsx", "open_so", `${rows.length} records as of ${asOfDate}`));
  Object.entries(openPoSeed).forEach(([asOfDate, rows]) => insHist.run(sep17 + 1 * 60000, "Open_Purchase_Order_as_of_Sep_17__2026.xlsx", "open_po", `${rows.length} records as of ${asOfDate}`));
  db.prepare("INSERT OR REPLACE INTO meta (key,value) VALUES ('seeded','true')").run();
  console.log("Seeded database with fresh January-September 13, 2026 dataset.");
}
seedIfEmpty();

/* ---------------- helpers ---------------- */
function monthLabel(monthKey) {
  const [y, m] = monthKey.split("-");
  return new Date(Number(y), Number(m) - 1, 1).toLocaleDateString("en-US", { month: "long", year: "numeric" });
}

/* ---------------- routes ---------------- */

// list months that have any data at all
app.get("/api/months", (req, res) => {
  const rows = db.prepare(`
    SELECT substr(date,1,7) AS mk FROM sales_orders
    UNION SELECT substr(date,1,7) FROM purchase_orders
    UNION SELECT substr(date,1,7) FROM ar_items
    UNION SELECT substr(date,1,7) FROM ap_items
  `).all();
  const keys = [...new Set(rows.map((r) => r.mk))].sort();
  res.json(keys.map((k) => ({ key: k, label: monthLabel(k) })));
});

// full month payload: same shape the frontend previously kept in window.storage
app.get("/api/month/:key", (req, res) => {
  const { key } = req.params;
  const so = db.prepare("SELECT date,entity,type,num,memo,amount FROM sales_orders WHERE substr(date,1,7)=? ORDER BY date").all(key);
  const po = db.prepare("SELECT date,entity,type,num,memo,amount FROM purchase_orders WHERE substr(date,1,7)=? ORDER BY date").all(key);
  const arRows = db.prepare("SELECT * FROM ar_items WHERE substr(date,1,7)=? ORDER BY date").all(key);
  const apRows = db.prepare("SELECT * FROM ap_items WHERE substr(date,1,7)=? ORDER BY date").all(key);

  const arSnapshots = {};
  arRows.forEach((r) => {
    (arSnapshots[r.date] = arSnapshots[r.date] || { count: 0, total: 0, items: [] }).items.push(r);
  });
  Object.values(arSnapshots).forEach((s) => { s.count = s.items.length; s.total = Math.round(s.items.reduce((a, b) => a + b.total, 0) * 100) / 100; });

  const apSnapshots = {};
  apRows.forEach((r) => {
    (apSnapshots[r.date] = apSnapshots[r.date] || { count: 0, total: 0, items: [] }).items.push(r);
  });
  Object.values(apSnapshots).forEach((s) => { s.count = s.items.length; s.total = Math.round(s.items.reduce((a, b) => a + b.total, 0) * 100) / 100; });

  res.json({ monthLabel: monthLabel(key), so, po, arSnapshots, apSnapshots });
});

// combined data across every month ever imported - powers period presets (This Fiscal Year, etc.)
app.get("/api/data/all", (req, res) => {
  const so = db.prepare("SELECT date,entity,type,num,memo,amount FROM sales_orders ORDER BY date").all();
  const po = db.prepare("SELECT date,entity,type,num,memo,amount FROM purchase_orders ORDER BY date").all();
  const arRows = db.prepare("SELECT * FROM ar_items ORDER BY date").all();
  const apRows = db.prepare("SELECT * FROM ap_items ORDER BY date").all();
  const ngr = db.prepare("SELECT date,salesRep,factory,customer,invoiceNum,totalAmount,paidRevenue FROM ngr_sales ORDER BY date").all();
  const ifs = db.prepare("SELECT date,salesRep,customer,invoiceNum,factory,jobName,status,invoiceTotalAmount,invoicePaymentAmount,totalCommAmount,actualCommissionPct,factoryCommissionPct FROM ifs_commission ORDER BY date").all();
  const salesDetail = db.prepare("SELECT date,customer,type,num,memo,item,qty,salesPrice,amount FROM sales_detail ORDER BY date").all();
  const qbSummaryRows = db.prepare("SELECT asOfDate,customer,a,b,dollarChange,pctChange FROM qb_sales_summary ORDER BY asOfDate").all();
  const qbSummary = {};
  qbSummaryRows.forEach((r) => { (qbSummary[r.asOfDate] = qbSummary[r.asOfDate] || []).push(r); });

  const openSoRows = db.prepare("SELECT asOfDate,customer,type,date,orderNum,memo,amount,openBalance,prepaymentOpenBalance FROM open_sales_orders ORDER BY asOfDate").all();
  const openSalesOrders = {};
  openSoRows.forEach((r) => { (openSalesOrders[r.asOfDate] = openSalesOrders[r.asOfDate] || { count: 0, totalAmount: 0, totalOpenBalance: 0, totalPrepayment: 0, items: [] }).items.push(r); });
  Object.values(openSalesOrders).forEach((s) => {
    s.count = s.items.length;
    s.totalAmount = round2(s.items.reduce((a, b) => a + b.amount, 0));
    s.totalOpenBalance = round2(s.items.reduce((a, b) => a + b.openBalance, 0));
    s.totalPrepayment = round2(s.items.reduce((a, b) => a + b.prepaymentOpenBalance, 0));
  });

  const openPoRows = db.prepare("SELECT asOfDate,vendor,type,date,orderNum,memo,deliveryDate,amount,openBalance FROM open_purchase_orders ORDER BY asOfDate").all();
  const openPurchaseOrders = {};
  openPoRows.forEach((r) => { (openPurchaseOrders[r.asOfDate] = openPurchaseOrders[r.asOfDate] || { count: 0, totalAmount: 0, totalOpenBalance: 0, items: [] }).items.push(r); });
  Object.values(openPurchaseOrders).forEach((s) => {
    s.count = s.items.length;
    s.totalAmount = round2(s.items.reduce((a, b) => a + b.amount, 0));
    s.totalOpenBalance = round2(s.items.reduce((a, b) => a + b.openBalance, 0));
  });

  const invRows = db.prepare("SELECT asOfDate,item,onHand,avgCost,assetValue,pctOfTotalAsset,salesPrice,retailValue,pctOfTotalRetail FROM inventory_valuation ORDER BY asOfDate").all();
  const inventoryValuation = {};
  invRows.forEach((r) => {
    r.flagged = r.onHand < 0 || r.assetValue < 0;
    (inventoryValuation[r.asOfDate] = inventoryValuation[r.asOfDate] || { count: 0, totalOnHand: 0, totalAssetValue: 0, totalRetailValue: 0, flaggedCount: 0, items: [] }).items.push(r);
  });
  Object.values(inventoryValuation).forEach((s) => {
    s.count = s.items.length;
    s.totalOnHand = round2(s.items.reduce((a, b) => a + b.onHand, 0));
    s.totalAssetValue = round2(s.items.reduce((a, b) => a + b.assetValue, 0));
    s.totalRetailValue = round2(s.items.reduce((a, b) => a + b.retailValue, 0));
    s.flaggedCount = s.items.filter((i) => i.flagged).length;
  });

  const arSnapshots = {};
  arRows.forEach((r) => { (arSnapshots[r.date] = arSnapshots[r.date] || { count: 0, total: 0, items: [] }).items.push(r); });
  Object.values(arSnapshots).forEach((s) => { s.count = s.items.length; s.total = Math.round(s.items.reduce((a, b) => a + b.total, 0) * 100) / 100; });

  const apSnapshots = {};
  apRows.forEach((r) => { (apSnapshots[r.date] = apSnapshots[r.date] || { count: 0, total: 0, items: [] }).items.push(r); });
  Object.values(apSnapshots).forEach((s) => { s.count = s.items.length; s.total = Math.round(s.items.reduce((a, b) => a + b.total, 0) * 100) / 100; });

  res.json({ so, po, arSnapshots, apSnapshots, ngr, ifs, salesDetail, qbSummary, openSalesOrders, openPurchaseOrders, inventoryValuation });
});

// Wipes every table completely - for starting over from scratch. This also
// marks the database as "seeded" so a later redeploy never silently brings
// the old seed data back; the tables just stay empty until you upload again.
app.post("/api/reset-all", (req, res) => {
  const tx = db.transaction(() => {
    db.prepare("DELETE FROM sales_orders").run();
    db.prepare("DELETE FROM purchase_orders").run();
    db.prepare("DELETE FROM ar_items").run();
    db.prepare("DELETE FROM ap_items").run();
    db.prepare("DELETE FROM ngr_sales").run();
    db.prepare("DELETE FROM ifs_commission").run();
    db.prepare("DELETE FROM sales_detail").run();
    db.prepare("DELETE FROM qb_sales_summary").run();
    db.prepare("DELETE FROM open_sales_orders").run();
    db.prepare("DELETE FROM open_purchase_orders").run();
    db.prepare("DELETE FROM inventory_valuation").run();
    db.prepare("DELETE FROM upload_history").run();
    db.prepare("INSERT OR REPLACE INTO meta (key,value) VALUES ('seeded','true')").run();
  });
  tx();
  res.json({ ok: true });
});

app.get("/api/history", (req, res) => {
  const rows = db.prepare("SELECT id, ts AS `when`, filename, kind, detail FROM upload_history ORDER BY ts DESC LIMIT 200").all();
  res.json(rows);
});

// Remove a single upload history entry. For AR/AP uploads this also removes the
// snapshot data itself (safe: snapshots are keyed by date, easy to target exactly).
// For SO/PO uploads this only removes the log entry - see /api/cleanup/invalid-dates
// below for removing bad transaction data from a wrong-file-type mistake.
app.delete("/api/history/:id", (req, res) => {
  const row = db.prepare("SELECT * FROM upload_history WHERE id=?").get(req.params.id);
  if (!row) return res.status(404).json({ error: "History entry not found." });
  if (row.kind === "ar" || row.kind === "ap") {
    const m = row.detail.match(/for (\d{4}-\d{2}-\d{2})/);
    if (m) {
      const table = row.kind === "ar" ? "ar_items" : "ap_items";
      db.prepare(`DELETE FROM ${table} WHERE date=?`).run(m[1]);
    }
  } else if (row.kind === "so" || row.kind === "po" || row.kind === "ngr" || row.kind === "ifs" || row.kind === "detail" || row.kind === "qbsummary" || row.kind === "open_so" || row.kind === "open_po" || row.kind === "inventory") {
    // Precise: only the rows this specific upload created, regardless of
    // date overlaps with other uploads. Older history rows (from before this
    // tracking existed) have upload_id = NULL on their rows, so nothing gets
    // touched for those - only the log entry is removed, same as before.
    const table = { so: "sales_orders", po: "purchase_orders", ngr: "ngr_sales", ifs: "ifs_commission", detail: "sales_detail", qbsummary: "qb_sales_summary", open_so: "open_sales_orders", open_po: "open_purchase_orders", inventory: "inventory_valuation" }[row.kind];
    db.prepare(`DELETE FROM ${table} WHERE upload_id=?`).run(row.id);
  }
  db.prepare("DELETE FROM upload_history WHERE id=?").run(req.params.id);
  res.json({ ok: true });
});

// One-time cleanup: removes any sales_orders/purchase_orders rows with an
// implausible date (a strong signal that an AR/AP file was accidentally
// uploaded as a Sales/Purchase Orders file, which misreads dollar amounts as
// Excel date serial numbers, landing in the early 1900s).
app.post("/api/cleanup/invalid-dates", (req, res) => {
  const bound = { min: "2020-01-01", max: "2035-12-31" };
  const soRemoved = db.prepare("DELETE FROM sales_orders WHERE date < ? OR date > ?").run(bound.min, bound.max).changes;
  const poRemoved = db.prepare("DELETE FROM purchase_orders WHERE date < ? OR date > ?").run(bound.min, bound.max).changes;
  res.json({ ok: true, soRemoved, poRemoved });
});

// upload a file: multipart form fields -> file, kind ('so'|'po'|'ar'|'ap'|'ngr'|'ifs'), date (required for ar/ap)
app.post("/api/upload", upload.single("file"), (req, res) => {
  try {
    const { kind, date } = req.body;
    if (!req.file) return res.status(400).json({ error: "No file received." });
    if (!["so", "po", "ar", "ap", "ngr", "ifs", "detail", "qbsummary", "open_so", "open_po", "inventory"].includes(kind)) return res.status(400).json({ error: "Invalid kind." });
    if ((kind === "ar" || kind === "ap" || kind === "qbsummary" || kind === "open_so" || kind === "open_po" || kind === "inventory") && !date) return res.status(400).json({ error: "A date is required for this file type." });

    const wb = parseWorkbookBuffer(req.file.buffer);

    if (kind === "inventory") {
      const rawRows = sheetToRows(wb);
      const records = parseInventoryValuationRows(rawRows);
      if (!records.length) {
        return res.status(400).json({ error: "No item rows found. Check this is a QuickBooks Inventory Valuation Summary export." });
      }
      const historyId = db.prepare("INSERT INTO upload_history (ts, filename, kind, detail) VALUES (?,?,?,?)").run(Date.now(), req.file.originalname, kind, `${records.length} items as of ${date}`).lastInsertRowid;
      const del = db.prepare("DELETE FROM inventory_valuation WHERE asOfDate=?");
      const ins = db.prepare("INSERT INTO inventory_valuation (asOfDate,item,onHand,avgCost,assetValue,pctOfTotalAsset,salesPrice,retailValue,pctOfTotalRetail,upload_id) VALUES (?,?,?,?,?,?,?,?,?,?)");
      const tx = db.transaction(() => {
        del.run(date);
        records.forEach((r) => ins.run(date, r.item, r.onHand, r.avgCost, r.assetValue, r.pctOfTotalAsset, r.salesPrice, r.retailValue, r.pctOfTotalRetail, historyId));
      });
      tx();
      const negativeCount = records.filter((r) => r.onHand < 0 || r.assetValue < 0).length;
      return res.json({ ok: true, count: records.length, date, negativeCount });
    }

    if (kind === "open_so" || kind === "open_po") {
      const rawRows = sheetToRows(wb);
      const records = kind === "open_so" ? parseOpenSalesOrdersRows(rawRows) : parseOpenPurchaseOrdersRows(rawRows);
      if (!records.length) {
        return res.status(400).json({ error: `No rows found. Check this is a QuickBooks Open ${kind === "open_so" ? "Sales" : "Purchase"} Orders export.` });
      }
      const table = kind === "open_so" ? "open_sales_orders" : "open_purchase_orders";
      const historyId = db.prepare("INSERT INTO upload_history (ts, filename, kind, detail) VALUES (?,?,?,?)").run(Date.now(), req.file.originalname, kind, `${records.length} records as of ${date}`).lastInsertRowid;
      const del = db.prepare(`DELETE FROM ${table} WHERE asOfDate=?`);
      const tx = db.transaction(() => {
        del.run(date);
        if (kind === "open_so") {
          const ins = db.prepare("INSERT INTO open_sales_orders (asOfDate,customer,type,date,orderNum,memo,amount,openBalance,prepaymentOpenBalance,upload_id) VALUES (?,?,?,?,?,?,?,?,?,?)");
          records.forEach((r) => ins.run(date, r.customer, r.type, r.date, r.orderNum, r.memo, r.amount, r.openBalance, r.prepaymentOpenBalance, historyId));
        } else {
          const ins = db.prepare("INSERT INTO open_purchase_orders (asOfDate,vendor,type,date,orderNum,memo,deliveryDate,amount,openBalance,upload_id) VALUES (?,?,?,?,?,?,?,?,?,?)");
          records.forEach((r) => ins.run(date, r.vendor, r.type, r.date, r.orderNum, r.memo, r.deliveryDate, r.amount, r.openBalance, historyId));
        }
      });
      tx();
      return res.json({ ok: true, count: records.length, date });
    }

    if (kind === "qbsummary") {
      const rows = sheetToRows(wb);
      const records = parseQBSummaryRows(rows);
      if (!records.length) {
        return res.status(400).json({ error: "No customer rows found. Check this is a QuickBooks Sales by Customer Summary export with a customer name column, not the total-only version." });
      }
      const historyId = db.prepare("INSERT INTO upload_history (ts, filename, kind, detail) VALUES (?,?,?,?)").run(Date.now(), req.file.originalname, kind, `${records.length} customers as of ${date}`).lastInsertRowid;
      const del = db.prepare("DELETE FROM qb_sales_summary WHERE asOfDate=?");
      const ins = db.prepare("INSERT INTO qb_sales_summary (asOfDate,customer,a,b,dollarChange,pctChange,upload_id) VALUES (?,?,?,?,?,?,?)");
      const tx = db.transaction(() => {
        del.run(date);
        records.forEach((r) => ins.run(date, r.customer, r.a, r.b, r.dollarChange, r.pctChange, historyId));
      });
      tx();
      return res.json({ ok: true, count: records.length, date });
    }

    if (kind === "detail") {
      const rows = sheetToRows(wb);
      const records = parseSalesDetailRows(rows);
      if (!records.length) {
        return res.status(400).json({ error: "No line-item sales rows found. Check this is a QuickBooks 'Sales by Customer Detail' export (with Type, Date, Num, Memo, Name, Item, Qty, Sales Price, Amount columns)." });
      }
      const badDates = records.filter((r) => r.date < "2015-01-01" || r.date > "2035-12-31");
      if (badDates.length > records.length * 0.3) {
        return res.status(400).json({ error: `This file's dates don't look right (found dates like ${badDates[0].date}). Double-check this is really a detailed sales export.` });
      }
      const goodRecords = records.filter((r) => r.date >= "2015-01-01" && r.date <= "2035-12-31");
      const detail = badDates.length
        ? `${goodRecords.length} line items (${badDates.length} skipped - invalid date in source file)`
        : `${goodRecords.length} line items`;
      const historyId = db.prepare("INSERT INTO upload_history (ts, filename, kind, detail) VALUES (?,?,?,?)").run(Date.now(), req.file.originalname, kind, detail).lastInsertRowid;
      const uploadDates = [...new Set(goodRecords.map((r) => r.date))];
      const del = db.prepare("DELETE FROM sales_detail WHERE date=?");
      const ins = db.prepare("INSERT INTO sales_detail (date,customer,type,num,memo,item,qty,salesPrice,amount,upload_id) VALUES (?,?,?,?,?,?,?,?,?,?)");
      const tx = db.transaction(() => {
        uploadDates.forEach((d) => del.run(d));
        goodRecords.forEach((r) => ins.run(r.date, r.customer, r.type, r.num, r.memo, r.item, r.qty, r.salesPrice, r.amount, historyId));
      });
      tx();
      return res.json({ ok: true, count: goodRecords.length, skipped: badDates.length, dates: uploadDates });
    }

    if (kind === "ngr" || kind === "ifs") {
      const objs = sheetToObjects(wb);
      const records = kind === "ngr" ? parseNGRRows(objs) : parseIFSRows(objs);
      if (!records.length) {
        return res.status(400).json({ error: `No ${kind === "ngr" ? "NGR Sales" : "IFS Commission"} rows found. Check this file has the expected columns (Sales Rep, Customer, Invoice Date, etc.) and a header row.` });
      }
      // This system only tracks 2026 onward for NGR/IFS - anything dated earlier
      // is out of scope (not a data error) and is excluded automatically.
      const outOfScope = records.filter((r) => r.date < "2026-01-01");
      const badDates = records.filter((r) => r.date >= "2026-01-01" && r.date > "2035-12-31");
      if (badDates.length > records.length * 0.3) {
        return res.status(400).json({ error: `This file's dates don't look right (found dates like ${badDates[0].date}). Double-check this is really a ${kind === "ngr" ? "NGR Sales" : "IFS Commission"} report.` });
      }
      const goodRecords = records.filter((r) => r.date >= "2026-01-01" && r.date <= "2035-12-31");
      const table = kind === "ngr" ? "ngr_sales" : "ifs_commission";
      const skippedNote = [];
      if (outOfScope.length) skippedNote.push(`${outOfScope.length} excluded - dated before 2026`);
      if (badDates.length) skippedNote.push(`${badDates.length} skipped - invalid date`);
      const detail = skippedNote.length
        ? `${goodRecords.length} records (${skippedNote.join("; ")})`
        : `${goodRecords.length} records`;
      const historyId = db.prepare("INSERT INTO upload_history (ts, filename, kind, detail) VALUES (?,?,?,?)").run(Date.now(), req.file.originalname, kind, detail).lastInsertRowid;
      const uploadDates = [...new Set(goodRecords.map((r) => r.date))];
      const del = db.prepare(`DELETE FROM ${table} WHERE date=?`);
      const tx = db.transaction(() => {
        uploadDates.forEach((d) => del.run(d));
        if (kind === "ngr") {
          const ins = db.prepare("INSERT INTO ngr_sales (date,salesRep,factory,customer,invoiceNum,totalAmount,paidRevenue,upload_id) VALUES (?,?,?,?,?,?,?,?)");
          goodRecords.forEach((r) => ins.run(r.date, r.salesRep, r.factory, r.customer, r.invoiceNum, r.totalAmount, r.paidRevenue, historyId));
        } else {
          const ins = db.prepare("INSERT INTO ifs_commission (date,salesRep,customer,invoiceNum,factory,jobName,status,invoiceTotalAmount,invoicePaymentAmount,totalCommAmount,actualCommissionPct,factoryCommissionPct,upload_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)");
          goodRecords.forEach((r) => ins.run(r.date, r.salesRep, r.customer, r.invoiceNum, r.factory, r.jobName, r.status, r.invoiceTotalAmount, r.invoicePaymentAmount, r.totalCommAmount, r.actualCommissionPct, r.factoryCommissionPct, historyId));
        }
      });
      tx();
      return res.json({ ok: true, count: goodRecords.length, skipped: badDates.length, excluded: outOfScope.length, dates: uploadDates });
    }

    const rows = sheetToRows(wb);

    if (kind === "so" || kind === "po") {
      const allTxns = parseTxnRows(rows);
      if (!allTxns.length) return res.status(400).json({ error: "No transactions found in that file." });
      // FR-01/FR-02: only rows whose QuickBooks Type is exactly "Sales Order" or
      // "Purchase Order" count as order activity. Invoices, Payments, Bills,
      // Checks, Credit Memos etc. must never be counted here, even though they
      // appear in the same export - counting them was inflating both the entry
      // count and the dollar totals.
      const wantedType = kind === "so" ? "Sales Order" : "Purchase Order";
      const excludedOtherTypes = allTxns.filter((t) => t.type !== wantedType).length;
      const txns = allTxns.filter((t) => t.type === wantedType);
      if (!txns.length) {
        return res.status(400).json({ error: `No rows with Type = "${wantedType}" were found in that file (it had ${allTxns.length} other transaction rows - Invoices, Payments, Bills, etc. - which don't count as ${kind === "so" ? "Sales Order" : "Purchase Order"} activity). Double-check this is the right file.` });
      }
      // Guard rail: an AR/AP aging file selected under the wrong file type will
      // still "parse" as transactions, but the dates come out wildly implausible
      // (dollar amounts misread as date serial numbers, landing in the early
      // 1900s). Reject those outright instead of importing garbage.
      const badDates = txns.filter((t) => t.date < "2020-01-01" || t.date > "2035-12-31");
      if (badDates.length > txns.length * 0.3) {
        return res.status(400).json({
          error: `This file's dates don't look right for a ${kind === "so" ? "Sales Orders" : "Purchase Orders"} import (found dates like ${badDates[0].date}). Double-check the File Type dropdown - this may actually be an AR/AP snapshot file.`,
        });
      }
      const goodTxns = txns.filter((t) => t.date >= "2020-01-01" && t.date <= "2035-12-31");
      const table = kind === "so" ? "sales_orders" : "purchase_orders";
      // Log the upload first, so every inserted row can be tagged with the
      // exact history entry that created it - this is what makes "Delete"
      // on this row precise later, even if it turns out to have been
      // uploaded under the wrong File Type (e.g. a PO file selected as SO).
      const detailNotes = [];
      if (excludedOtherTypes) detailNotes.push(`${excludedOtherTypes} other transaction types excluded`);
      if (badDates.length) detailNotes.push(`${badDates.length} skipped - invalid date`);
      const detail = detailNotes.length
        ? `${goodTxns.length} ${wantedType} records (${detailNotes.join("; ")})`
        : `${goodTxns.length} ${wantedType} records`;
      const historyId = db.prepare("INSERT INTO upload_history (ts, filename, kind, detail) VALUES (?,?,?,?)").run(Date.now(), req.file.originalname, kind, detail).lastInsertRowid;
      // Clear every date this file covers (using the full, unfiltered set of
      // transaction dates in the source file) before inserting the filtered
      // order-type rows. This matters: a date with zero "Sales Order"/"Purchase
      // Order" rows (only Invoices, Bills, etc. that day) must end up with zero
      // rows here too - if we only cleared dates that still have surviving
      // filtered rows, old mixed-type data from before this fix would linger
      // on any date that turns out to have no real order activity.
      const coveredDates = [...new Set(allTxns.map((t) => t.date))];
      const uploadDates = [...new Set(goodTxns.map((t) => t.date))];
      const del = db.prepare(`DELETE FROM ${table} WHERE date=?`);
      const ins = db.prepare(`INSERT INTO ${table} (date,entity,type,num,memo,amount,upload_id) VALUES (?,?,?,?,?,?,?)`);
      const tx = db.transaction(() => {
        coveredDates.forEach((d) => del.run(d));
        goodTxns.forEach((t) => ins.run(t.date, t.entity, t.type, t.num, t.memo, t.amount, historyId));
      });
      tx();
      return res.json({ ok: true, count: goodTxns.length, skipped: badDates.length, excludedOtherTypes, dates: uploadDates });
    } else {
      const { count, items } = parseAgingRows(rows);
      if (!count) return res.status(400).json({ error: "No aging rows found in that file." });
      const table = kind === "ar" ? "ar_items" : "ap_items";
      const del = db.prepare(`DELETE FROM ${table} WHERE date=?`);
      const ins = db.prepare(`INSERT INTO ${table} (date,name,total,current,d1_30,d31_60,d61_90,d90plus) VALUES (?,?,?,?,?,?,?,?)`);
      const tx = db.transaction(() => {
        del.run(date);
        items.forEach((it) => ins.run(date, it.name, it.total, it.current, it.d1_30, it.d31_60, it.d61_90, it.d90plus));
      });
      tx();
      const detail = `${count} records for ${date}`;
      db.prepare("INSERT INTO upload_history (ts, filename, kind, detail) VALUES (?,?,?,?)").run(Date.now(), req.file.originalname, kind, detail);
      return res.json({ ok: true, count, date });
    }
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Could not process that file." });
  }
});

/* ---------------- serve built frontend in production ---------------- */
const clientDist = path.join(__dirname, "..", "client", "dist");
if (fs.existsSync(clientDist)) {
  app.use(express.static(clientDist));
  app.get("*", (req, res) => {
    if (req.path.startsWith("/api/")) return res.status(404).json({ error: "Not found" });
    res.sendFile(path.join(clientDist, "index.html"));
  });
}

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`V&K dashboard API listening on port ${PORT}`));
