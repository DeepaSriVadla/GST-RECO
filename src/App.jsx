import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

/*
===========================================================
GST RECO
SMART GST RECONCILIATION
===========================================================

This application reads the uploaded files directly.

BOOKS:
Purchase Register

GSTR-2B:
B2B
B2BA
B2B-CDNR
B2B-CDNRA
IMPG
IMPGA
IMPGSEZ
IMPGSEZA

Important:
- Blank is NOT automatically treated as zero.
- Books + 2B is NOT a duplicate.
- Duplicate means duplicate within the same source.
- Matching is conservative.
- Ambiguous matches are not silently selected.
===========================================================
*/


/* =======================================================
   BASIC HELPERS
======================================================= */

function text(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}


function normalizedText(value) {
  return text(value)
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}


function normalizeInvoice(value) {
  return text(value)
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/[^A-Z0-9]/g, "");
}


function normalizeGSTIN(value) {
  const valueText = text(value)
    .toUpperCase()
    .replace(/\s+/g, "");

  if (!valueText) {
    return "";
  }

  const match = valueText.match(
    /\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]/
  );

  return match ? match[0] : valueText;
}


function numberValue(value) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return null;
  }

  if (typeof value === "number") {
    return Number.isFinite(value)
      ? value
      : null;
  }

  const cleaned = String(value)
    .replace(/₹/g, "")
    .replace(/,/g, "")
    .replace(/\(/g, "-")
    .replace(/\)/g, "")
    .trim();

  if (!cleaned) {
    return null;
  }

  const result = Number(cleaned);

  return Number.isFinite(result)
    ? result
    : null;
}


function amount(value) {
  if (
    value === null ||
    value === undefined
  ) {
    return "—";
  }

  return `₹${Number(value).toLocaleString(
    "en-IN",
    {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }
  )}`;
}


function dateValue(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  /*
    Always normalize dates to YYYY-MM-DD.

    IMPORTANT:
    Books Excel files commonly return a JavaScript Date object,
    while GSTR-2B commonly returns a text date such as 01/08/2026.
    Comparing those two representations directly caused genuine
    matches to be incorrectly shown as PARTIAL.
  */

  if (
    value instanceof Date &&
    !Number.isNaN(value.getTime())
  ) {
    return `${value.getFullYear()}-${String(
      value.getMonth() + 1
    ).padStart(2, "0")}-${String(
      value.getDate()
    ).padStart(2, "0")}`;
  }

  /* Excel serial date number. */
  if (typeof value === "number" && Number.isFinite(value)) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(
      excelEpoch.getTime() + value * 86400000
    );

    if (!Number.isNaN(date.getTime())) {
      return `${date.getUTCFullYear()}-${String(
        date.getUTCMonth() + 1
      ).padStart(2, "0")}-${String(
        date.getUTCDate()
      ).padStart(2, "0")}`;
    }
  }

  const raw = text(value).trim();

  if (!raw) {
    return "";
  }

  /* DD/MM/YYYY or DD-MM-YYYY */
  let match = raw.match(
    /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/
  );

  if (match) {
    const day = match[1].padStart(2, "0");
    const month = match[2].padStart(2, "0");
    const year = match[3];

    return `${year}-${month}-${day}`;
  }

  /* YYYY/MM/DD or YYYY-MM-DD */
  match = raw.match(
    /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/
  );

  if (match) {
    const year = match[1];
    const month = match[2].padStart(2, "0");
    const day = match[3].padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  /* ISO date/time strings such as 2026-08-01T00:00:00 */
  match = raw.match(
    /^(\d{4})-(\d{2})-(\d{2})/
  );

  if (match) {
    return `${match[1]}-${match[2]}-${match[3]}`;
  }

  /* Leave an unknown format untouched rather than guessing. */
  return raw;
}


function equalAmount(a, b, tolerance = 0.01) {
  if (
    a === null ||
    b === null ||
    a === undefined ||
    b === undefined
  ) {
    return false;
  }

  return (
    Math.abs(Number(a) - Number(b)) <=
    tolerance
  );
}


function difference(a, b) {
  if (
    a === null ||
    b === null ||
    a === undefined ||
    b === undefined
  ) {
    return null;
  }

  return Math.abs(
    Number(a) - Number(b)
  );
}


function addNumbers(values) {
  const valid = values.filter(
    (v) =>
      v !== null &&
      v !== undefined
  );

  if (!valid.length) {
    return null;
  }

  return valid.reduce(
    (sum, v) => sum + Number(v),
    0
  );
}


/* =======================================================
   COLUMN SEARCH
======================================================= */

function normalizedHeader(value) {
  return text(value)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}


function findColumn(headers, names) {
  const normalizedHeaders =
    headers.map(normalizedHeader);

  for (const name of names) {
    const target =
      normalizedHeader(name);

    const exact =
      normalizedHeaders.indexOf(
        target
      );

    if (exact !== -1) {
      return exact;
    }
  }

  return -1;
}


/* =======================================================
   BOOKS PARSER
======================================================= */

function findBooksHeader(rows) {
  for (
    let i = 0;
    i < Math.min(rows.length, 40);
    i++
  ) {
    const row = rows[i] || [];

    const joined = row
      .map(text)
      .join(" | ")
      .toLowerCase();

    if (
      joined.includes(
        "supplier invoice no."
      ) &&
      joined.includes(
        "supplier invoice date"
      ) &&
      joined.includes(
        "gstin/uin"
      )
    ) {
      return i;
    }
  }

  return -1;
}


function parseBooks(workbook) {
  const sheetName =
    workbook.SheetNames.find(
      (name) =>
        normalizedText(name) ===
        "PURCHASE REGISTER"
    );

  if (!sheetName) {
    throw new Error(
      "Purchase Register sheet was not found in the Books file."
    );
  }

  const sheet =
    workbook.Sheets[sheetName];

  const rows =
    XLSX.utils.sheet_to_json(
      sheet,
      {
        header: 1,
        defval: "",
        raw: true,
      }
    );

  const headerIndex =
    findBooksHeader(rows);

  if (headerIndex === -1) {
    throw new Error(
      "Could not identify the Purchase Register header row."
    );
  }

  const headers =
    rows[headerIndex].map(text);

  const dateCol =
    findColumn(headers, [
      "Date",
    ]);

  const particularsCol =
    findColumn(headers, [
      "Particulars",
    ]);

  const supplierCol =
    findColumn(headers, [
      "Supplier",
    ]);

  const chapterCol =
    findColumn(headers, [
      "Chapter Heading",
    ]);

  const dcNoCol =
    findColumn(headers, [
      "DC No.",
    ]);

  const dcDateCol =
    findColumn(headers, [
      "DC Dt.",
    ]);

  const poNoCol =
    findColumn(headers, [
      "PO No.",
    ]);

  const poDateCol =
    findColumn(headers, [
      "PO Dt.",
    ]);

  const voucherCol =
    findColumn(headers, [
      "Voucher Type",
    ]);

  const invoiceCol =
    findColumn(headers, [
      "Supplier Invoice No.",
    ]);

  const invoiceDateCol =
    findColumn(headers, [
      "Supplier Invoice Date",
    ]);

  const gstinCol =
    findColumn(headers, [
      "GSTIN/UIN",
      "GSTIN",
    ]);

  const boeCol =
    findColumn(headers, [
      "Bill of Entry No.",
    ]);

  const boeDateCol =
    findColumn(headers, [
      "Bill of Entry Date",
    ]);

  const portCol =
    findColumn(headers, [
      "Port Code",
    ]);

  const valueCol =
    findColumn(headers, [
      "Value",
    ]);

  const additionalCostCol =
    findColumn(headers, [
      "Addl. Cost",
    ]);

  const grossTotalCol =
    findColumn(headers, [
      "Gross Total",
    ]);

  const igstCol =
    findColumn(headers, [
      "Integrated Tax - IGST",
    ]);

  const cgstCol =
    findColumn(headers, [
      "Central Tax - CGST",
    ]);

  const sgstCol =
    findColumn(headers, [
      "State Tax -   SGST",
      "State Tax - SGST",
    ]);

  const cessCol =
    findColumn(headers, [
      "Cess",
      "Cess Amount",
      "Compensation Cess",
      "Cess Tax",
    ]);

  const records = [];

  for (
    let rowIndex =
      headerIndex + 1;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] || [];

    const hasData =
      row.some(
        (cell) =>
          text(cell) !== ""
      );

    if (!hasData) {
      continue;
    }

    const rowText =
      row
        .map(text)
        .join(" ")
        .toLowerCase();

    /*
      Ignore final total / grand total rows.
    */

    if (
      rowText.includes(
        "grand total"
      ) ||
      rowText === "total"
    ) {
      continue;
    }

    const invoiceNo =
      invoiceCol >= 0
        ? text(row[invoiceCol])
        : "";

    const gstin =
      gstinCol >= 0
        ? normalizeGSTIN(
            row[gstinCol]
          )
        : "";

    const supplier =
      supplierCol >= 0
        ? text(row[supplierCol])
        : "";

    const invoiceDate =
      invoiceDateCol >= 0
        ? dateValue(
            row[invoiceDateCol]
          )
        : "";

    const postingDate =
      dateCol >= 0
        ? dateValue(
            row[dateCol]
          )
        : "";

    /*
      Completely empty transaction rows are ignored.
    */

    if (
      !invoiceNo &&
      !gstin &&
      !supplier &&
      !invoiceDate &&
      !postingDate
    ) {
      continue;
    }

    const value =
      valueCol >= 0
        ? numberValue(
            row[valueCol]
          )
        : null;

    const grossTotal =
      grossTotalCol >= 0
        ? numberValue(
            row[grossTotalCol]
          )
        : null;

    const igst =
      igstCol >= 0
        ? numberValue(
            row[igstCol]
          )
        : null;

    const cgst =
      cgstCol >= 0
        ? numberValue(
            row[cgstCol]
          )
        : null;

    const sgst =
      sgstCol >= 0
        ? numberValue(
            row[sgstCol]
          )
        : null;

    const cess =
      cessCol >= 0
        ? numberValue(
            row[cessCol]
          )
        : null;

    const totalTax =
      addNumbers([
        igst,
        cgst,
        sgst,
        cess,
      ]);

    const raw = {};

    headers.forEach(
      (header, index) => {
        if (header) {
          raw[header] =
            row[index] ?? "";
        }
      }
    );

    records.push({
      id:
        `BOOK-${records.length + 1}`,

      source:
        "BOOKS",

      rowNumber:
        rowIndex + 1,

      date:
        postingDate,

      particulars:
        particularsCol >= 0
          ? text(
              row[particularsCol]
            )
          : "",

      supplier,

      chapterHeading:
        chapterCol >= 0
          ? text(
              row[chapterCol]
            )
          : "",

      dcNo:
        dcNoCol >= 0
          ? text(
              row[dcNoCol]
            )
          : "",

      dcDate:
        dcDateCol >= 0
          ? dateValue(
              row[dcDateCol]
            )
          : "",

      poNo:
        poNoCol >= 0
          ? text(
              row[poNoCol]
            )
          : "",

      poDate:
        poDateCol >= 0
          ? dateValue(
              row[poDateCol]
            )
          : "",

      voucherType:
        voucherCol >= 0
          ? text(
              row[voucherCol]
            )
          : "",

      invoiceNo,

      normalizedInvoice:
        normalizeInvoice(
          invoiceNo
        ),

      invoiceDate,

      gstin,

      billOfEntryNo:
        boeCol >= 0
          ? text(row[boeCol])
          : "",

      billOfEntryDate:
        boeDateCol >= 0
          ? dateValue(
              row[boeDateCol]
            )
          : "",

      portCode:
        portCol >= 0
          ? text(row[portCol])
          : "",

      taxableValue:
        value,

      grossTotal,

      igst,

      cgst,

      sgst,

      cess,

      totalTax,

      raw,
    });
  }

  return {
    sheetName,
    headers,
    records,
  };
}


/* =======================================================
   GSTR-2B HEADER HELPERS
======================================================= */

function buildHeaders(
  top,
  bottom
) {
  const length =
    Math.max(
      top.length,
      bottom.length
    );

  const headers = [];

  for (
    let i = 0;
    i < length;
    i++
  ) {
    const a =
      text(top[i]);

    const b =
      text(bottom[i]);

    /*
      In the GST portal workbook,
      top row contains grouped headings.
      Bottom row contains actual field names.
    */

    const grouped = [
      "Invoice Details",
      "Tax Amount",
      "Credit note/Debit note details",
      "Bill of Entry Details",
      "Bill of Entry details",
      "Document Details",
      "Tax amount",
      "Amount of tax (₹)",
      "Input tax distribution by ISD",
    ];

    if (
      grouped.some(
        (g) =>
          normalizedText(a) ===
          normalizedText(g)
      )
    ) {
      headers.push(
        b || a
      );
    } else if (
      !a &&
      b
    ) {
      headers.push(b);
    } else {
      headers.push(a);
    }
  }

  return headers;
}


function findB2BHeader(rows) {
  for (
    let i = 0;
    i < Math.min(rows.length - 1, 25);
    i++
  ) {
    const top =
      (rows[i] || [])
        .map(text)
        .join(" | ")
        .toLowerCase();

    const bottom =
      (rows[i + 1] || [])
        .map(text)
        .join(" | ")
        .toLowerCase();

    if (
      top.includes(
        "gstin of supplier"
      ) &&
      top.includes(
        "invoice details"
      ) &&
      bottom.includes(
        "invoice number"
      )
    ) {
      return i;
    }
  }

  return -1;
}


/* =======================================================
   B2B PARSER
======================================================= */

function parseB2B(
  workbook,
  sheetName
) {
  if (
    !workbook.SheetNames.includes(
      sheetName
    )
  ) {
    return [];
  }

  const sheet =
    workbook.Sheets[sheetName];

  const rows =
    XLSX.utils.sheet_to_json(
      sheet,
      {
        header: 1,
        defval: "",
        raw: true,
      }
    );

  const headerIndex =
    findB2BHeader(rows);

  if (headerIndex === -1) {
    return [];
  }

  const headers =
    buildHeaders(
      rows[headerIndex],
      rows[headerIndex + 1]
    );

  const gstinCol =
    findColumn(headers, [
      "GSTIN of supplier",
    ]);

  const supplierCol =
    findColumn(headers, [
      "Trade/Legal name",
    ]);

  const invoiceNoCol =
    findColumn(headers, [
      "Invoice number",
    ]);

  const invoiceTypeCol =
    findColumn(headers, [
      "Invoice type",
    ]);

  const invoiceDateCol =
    findColumn(headers, [
      "Invoice Date",
    ]);

  const invoiceValueCol =
    findColumn(headers, [
      "Invoice Value(₹)",
    ]);

  const placeSupplyCol =
    findColumn(headers, [
      "Place of supply",
    ]);

  const reverseChargeCol =
    findColumn(headers, [
      "Supply Attract Reverse Charge",
    ]);

  const taxableCol =
    findColumn(headers, [
      "Taxable Value (₹)",
    ]);

  const igstCol =
    findColumn(headers, [
      "Integrated Tax(₹)",
    ]);

  const cgstCol =
    findColumn(headers, [
      "Central Tax(₹)",
    ]);

  const sgstCol =
    findColumn(headers, [
      "State/UT Tax(₹)",
    ]);

  const cessCol =
    findColumn(headers, [
      "Cess(₹)",
    ]);

  const periodCol =
    findColumn(headers, [
      "GSTR-1/1A/IFF/GSTR-5 Period",
    ]);

  const filingDateCol =
    findColumn(headers, [
      "GSTR-1/1A/IFF/GSTR-5 Filing Date",
    ]);

  const itcCol =
    findColumn(headers, [
      "ITC Availability",
    ]);

  const reasonCol =
    findColumn(headers, [
      "Reason",
    ]);

  const taxRateCol =
    findColumn(headers, [
      "Applicable % of Tax Rate",
    ]);

  const sourceCol =
    findColumn(headers, [
      "Source",
    ]);

  const irnCol =
    findColumn(headers, [
      "IRN",
    ]);

  const irnDateCol =
    findColumn(headers, [
      "IRN Date",
    ]);

  const records = [];

  for (
    let rowIndex =
      headerIndex + 2;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] || [];

    if (
      !row.some(
        (cell) =>
          text(cell) !== ""
      )
    ) {
      continue;
    }

    const invoiceNo =
      invoiceNoCol >= 0
        ? text(
            row[invoiceNoCol]
          )
        : "";

    const gstin =
      gstinCol >= 0
        ? normalizeGSTIN(
            row[gstinCol]
          )
        : "";

    const supplier =
      supplierCol >= 0
        ? text(
            row[supplierCol]
          )
        : "";

    if (
      !invoiceNo &&
      !gstin &&
      !supplier
    ) {
      continue;
    }

    const igst =
      igstCol >= 0
        ? numberValue(
            row[igstCol]
          )
        : null;

    const cgst =
      cgstCol >= 0
        ? numberValue(
            row[cgstCol]
          )
        : null;

    const sgst =
      sgstCol >= 0
        ? numberValue(
            row[sgstCol]
          )
        : null;

    const cess =
      cessCol >= 0
        ? numberValue(
            row[cessCol]
          )
        : null;

    const totalTax =
      addNumbers([
        igst,
        cgst,
        sgst,
        cess,
      ]);

    const raw = {};

    headers.forEach(
      (header, index) => {
        if (header) {
          raw[header] =
            row[index] ?? "";
        }
      }
    );

    records.push({
      id:
        `2B-${sheetName}-${records.length + 1}`,

      source:
        "GSTR-2B",

      section:
        sheetName,

      rowNumber:
        rowIndex + 1,

      gstin,

      supplier,

      invoiceNo,

      normalizedInvoice:
        normalizeInvoice(
          invoiceNo
        ),

      invoiceType:
        invoiceTypeCol >= 0
          ? text(
              row[invoiceTypeCol]
            )
          : "",

      invoiceDate:
        invoiceDateCol >= 0
          ? dateValue(
              row[invoiceDateCol]
            )
          : "",

      invoiceValue:
        invoiceValueCol >= 0
          ? numberValue(
              row[invoiceValueCol]
            )
          : null,

      placeOfSupply:
        placeSupplyCol >= 0
          ? text(
              row[placeSupplyCol]
            )
          : "",

      reverseCharge:
        reverseChargeCol >= 0
          ? text(
              row[reverseChargeCol]
            )
          : "",

      taxableValue:
        taxableCol >= 0
          ? numberValue(
              row[taxableCol]
            )
          : null,

      igst,

      cgst,

      sgst,

      cess,

      totalTax,

      period:
        periodCol >= 0
          ? text(
              row[periodCol]
            )
          : "",

      filingDate:
        filingDateCol >= 0
          ? dateValue(
              row[filingDateCol]
            )
          : "",

      itcAvailability:
        itcCol >= 0
          ? text(row[itcCol])
          : "",

      reason:
        reasonCol >= 0
          ? text(row[reasonCol])
          : "",

      taxRate:
        taxRateCol >= 0
          ? text(
              row[taxRateCol]
            )
          : "",

      sourceDocument:
        sourceCol >= 0
          ? text(
              row[sourceCol]
            )
          : "",

      irn:
        irnCol >= 0
          ? text(row[irnCol])
          : "",

      irnDate:
        irnDateCol >= 0
          ? dateValue(
              row[irnDateCol]
            )
          : "",

      raw,
    });
  }

  return records;
}


/* =======================================================
   B2BA PARSER

   B2BA is special because:
   ORIGINAL DETAILS:
      Invoice number
      Invoice Date
      GSTIN

   REVISED DETAILS:
      Invoice number
      Invoice type
      Invoice Date
      Invoice Value
      IGST
      CGST
      SGST
      Cess

   Revised details must be used for the
   current amended transaction.
======================================================= */

function parseB2BA(workbook) {
  const sheetName = "B2BA";

  if (
    !workbook.SheetNames.includes(
      sheetName
    )
  ) {
    return [];
  }

  const sheet =
    workbook.Sheets[sheetName];

  const rows =
    XLSX.utils.sheet_to_json(
      sheet,
      {
        header: 1,
        defval: "",
        raw: true,
      }
    );

  let headerIndex = -1;

  for (
    let i = 0;
    i < Math.min(rows.length - 1, 20);
    i++
  ) {
    const top =
      (rows[i] || [])
        .map(text)
        .join(" | ")
        .toLowerCase();

    const bottom =
      (rows[i + 1] || [])
        .map(text)
        .join(" | ")
        .toLowerCase();

    if (
      top.includes(
        "original details"
      ) &&
      top.includes(
        "revised details"
      ) &&
      bottom.includes(
        "invoice number"
      ) &&
      bottom.includes(
        "gstin of supplier"
      )
    ) {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) {
    return [];
  }

  /*
    Based on the official workbook layout
    observed in the uploaded file.

    Original:
      0 Invoice number
      1 Invoice Date
      2 GSTIN
      3 Trade/Legal name

    Revised:
      4 Invoice number
      5 Invoice type
      6 Invoice Date
      7 Invoice Value

    11 IGST
    12 CGST
    13 SGST
    14 Cess
  */

  const records = [];

  for (
    let rowIndex =
      headerIndex + 2;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] || [];

    if (
      !row.some(
        (cell) =>
          text(cell) !== ""
      )
    ) {
      continue;
    }

    const originalInvoice =
      text(row[0]);

    const originalDate =
      dateValue(row[1]);

    const gstin =
      normalizeGSTIN(row[2]);

    const supplier =
      text(row[3]);

    const revisedInvoice =
      text(row[4]);

    const revisedType =
      text(row[5]);

    const revisedDate =
      dateValue(row[6]);

    const revisedInvoiceValue =
      numberValue(row[7]);

    const placeOfSupply =
      text(row[8]);

    const reverseCharge =
      text(row[9]);

    const taxableValue =
      numberValue(row[10]);

    const igst =
      numberValue(row[11]);

    const cgst =
      numberValue(row[12]);

    const sgst =
      numberValue(row[13]);

    const cess =
      numberValue(row[14]);

    const totalTax =
      addNumbers([
        igst,
        cgst,
        sgst,
        cess,
      ]);

    /*
      Additional revised fields.
    */

    const itcReduction =
      text(row[15]);

    const reductionIGST =
      numberValue(row[16]);

    const reductionCGST =
      numberValue(row[17]);

    const reductionSGST =
      numberValue(row[18]);

    const reductionCESS =
      numberValue(row[19]);

    const remarks =
      text(row[20]);

    const period =
      text(row[21]);

    const filingDate =
      dateValue(row[22]);

    const itcAvailability =
      text(row[23]);

    const reason =
      text(row[24]);

    const taxRate =
      text(row[25]);

    /*
      Empty B2BA rows are ignored.
    */

    if (
      !originalInvoice &&
      !revisedInvoice &&
      !gstin
    ) {
      continue;
    }

    const raw = {};

    rows[headerIndex].forEach(
      (header, index) => {
        const h =
          text(header);

        if (h) {
          raw[h] =
            row[index] ?? "";
        }
      }
    );

    records.push({
      id:
        `2B-B2BA-${records.length + 1}`,

      source:
        "GSTR-2B",

      section:
        "B2BA",

      rowNumber:
        rowIndex + 1,

      gstin,

      supplier,

      /*
        Use REVISED invoice as the active
        invoice for reconciliation.
      */

      invoiceNo:
        revisedInvoice ||
        originalInvoice,

      normalizedInvoice:
        normalizeInvoice(
          revisedInvoice ||
          originalInvoice
        ),

      originalInvoice,

      originalNormalizedInvoice:
        normalizeInvoice(
          originalInvoice
        ),

      originalInvoiceDate:
        originalDate,

      invoiceType:
        revisedType,

      invoiceDate:
        revisedDate ||
        originalDate,

      invoiceValue:
        revisedInvoiceValue,

      placeOfSupply,

      reverseCharge,

      taxableValue,

      igst,

      cgst,

      sgst,

      cess,

      totalTax,

      itcReduction,

      reductionIGST,

      reductionCGST,

      reductionSGST,

      reductionCESS,

      remarks,

      period,

      filingDate,

      itcAvailability,

      reason,

      taxRate,

      raw,
    });
  }

  return records;
}


/* =======================================================
   CREDIT / DEBIT NOTES
======================================================= */

function parseCDNR(
  workbook,
  sheetName
) {
  if (
    !workbook.SheetNames.includes(
      sheetName
    )
  ) {
    return [];
  }

  const sheet =
    workbook.Sheets[sheetName];

  const rows =
    XLSX.utils.sheet_to_json(
      sheet,
      {
        header: 1,
        defval: "",
        raw: true,
      }
    );

  let headerIndex = -1;

  for (
    let i = 0;
    i < Math.min(rows.length - 1, 20);
    i++
  ) {
    const top =
      (rows[i] || [])
        .map(text)
        .join(" | ")
        .toLowerCase();

    const bottom =
      (rows[i + 1] || [])
        .map(text)
        .join(" | ")
        .toLowerCase();

    if (
      top.includes(
        "gstin of supplier"
      ) &&
      top.includes(
        "credit note/debit note details"
      ) &&
      bottom.includes(
        "note number"
      )
    ) {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) {
    return [];
  }

  const records = [];

  for (
    let rowIndex =
      headerIndex + 2;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] || [];

    if (
      !row.some(
        (cell) =>
          text(cell) !== ""
      )
    ) {
      continue;
    }

    const gstin =
      normalizeGSTIN(row[0]);

    const supplier =
      text(row[1]);

    const noteNo =
      text(row[2]);

    const noteType =
      text(row[3]);

    const noteSupplyType =
      text(row[4]);

    const noteDate =
      dateValue(row[5]);

    const noteValue =
      numberValue(row[6]);

    const placeOfSupply =
      text(row[7]);

    const reverseCharge =
      text(row[8]);

    const taxableValue =
      numberValue(row[9]);

    const igst =
      numberValue(row[10]);

    const cgst =
      numberValue(row[11]);

    const sgst =
      numberValue(row[12]);

    const cess =
      numberValue(row[13]);

    const totalTax =
      addNumbers([
        igst,
        cgst,
        sgst,
        cess,
      ]);

    if (
      !noteNo &&
      !gstin &&
      !supplier
    ) {
      continue;
    }

    records.push({
      id:
        `CDNR-${sheetName}-${records.length + 1}`,

      source:
        "GSTR-2B",

      section:
        sheetName,

      rowNumber:
        rowIndex + 1,

      gstin,

      supplier,

      invoiceNo:
        noteNo,

      normalizedInvoice:
        normalizeInvoice(
          noteNo
        ),

      noteNo,

      noteType,

      noteSupplyType,

      noteDate,

      noteValue,

      placeOfSupply,

      reverseCharge,

      taxableValue,

      igst,

      cgst,

      sgst,

      cess,

      totalTax,

      raw: {
        "GSTIN of supplier":
          row[0] ?? "",
        "Trade/Legal name":
          row[1] ?? "",
        "Note number":
          row[2] ?? "",
        "Note type":
          row[3] ?? "",
        "Note Supply type":
          row[4] ?? "",
        "Note date":
          row[5] ?? "",
        "Note Value (₹)":
          row[6] ?? "",
        "Place of supply":
          row[7] ?? "",
        "Supply Attract Reverse Charge":
          row[8] ?? "",
        "Taxable Value (₹)":
          row[9] ?? "",
        "Integrated Tax(₹)":
          row[10] ?? "",
        "Central Tax(₹)":
          row[11] ?? "",
        "State/UT Tax(₹)":
          row[12] ?? "",
        "Cess(₹)":
          row[13] ?? "",
      },
    });
  }

  return records;
}


/* =======================================================
   IMPORT SHEETS
======================================================= */

function parseImport(
  workbook,
  sheetName
) {
  if (
    !workbook.SheetNames.includes(
      sheetName
    )
  ) {
    return [];
  }

  const sheet =
    workbook.Sheets[sheetName];

  const rows =
    XLSX.utils.sheet_to_json(
      sheet,
      {
        header: 1,
        defval: "",
        raw: true,
      }
    );

  let headerIndex = -1;

  for (
    let i = 0;
    i < Math.min(rows.length - 1, 20);
    i++
  ) {
    const top =
      (rows[i] || [])
        .map(text)
        .join(" | ")
        .toLowerCase();

    const bottom =
      (rows[i + 1] || [])
        .map(text)
        .join(" | ")
        .toLowerCase();

    if (
      top.includes(
        "bill of entry details"
      ) &&
      bottom.includes(
        "number"
      )
    ) {
      headerIndex = i;
      break;
    }
  }

  /*
    Some import sheets have a slightly different
    heading. If the normal heading is not found,
    search for the data pattern.
  */

  if (headerIndex === -1) {
    for (
      let i = 0;
      i < Math.min(rows.length - 1, 20);
      i++
    ) {
      const joined =
        (rows[i] || [])
          .map(text)
          .join(" | ")
          .toLowerCase();

      if (
        joined.includes(
          "icegate reference date"
        )
      ) {
        headerIndex = i;
        break;
      }
    }
  }

  if (headerIndex === -1) {
    return [];
  }

  const records = [];

  for (
    let rowIndex =
      headerIndex + 2;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const row =
      rows[rowIndex] || [];

    if (
      !row.some(
        (cell) =>
          text(cell) !== ""
      )
    ) {
      continue;
    }

    /*
      Import sheets use the first columns
      for ICEGATE reference / port /
      bill of entry information.
    */

    const referenceDate =
      dateValue(row[0]);

    const portCode =
      text(row[1]);

    const billOfEntryNo =
      text(row[2]);

    const billOfEntryDate =
      dateValue(row[3]);

    const taxableValue =
      numberValue(row[4]);

    const igst =
      numberValue(row[5]);

    const cess =
      numberValue(row[6]);

    if (
      !referenceDate &&
      !portCode &&
      !billOfEntryNo
    ) {
      continue;
    }

    records.push({
      id:
        `IMPORT-${sheetName}-${records.length + 1}`,

      source:
        "GSTR-2B",

      section:
        sheetName,

      rowNumber:
        rowIndex + 1,

      referenceDate,

      portCode,

      billOfEntryNo,

      billOfEntryDate,

      taxableValue,

      igst,

      cess,

      totalTax:
        addNumbers([
          igst,
          cess,
        ]),

      raw: {
        "Icegate Reference Date":
          row[0] ?? "",
        "Port Code":
          row[1] ?? "",
        "Bill of Entry Number":
          row[2] ?? "",
        "Bill of Entry Date":
          row[3] ?? "",
        "Taxable Value":
          row[4] ?? "",
        "Integrated Tax":
          row[5] ?? "",
        Cess:
          row[6] ?? "",
      },
    });
  }

  return records;
}


/* =======================================================
   COMPLETE 2B PARSER
======================================================= */

function parseTwoB(workbook) {
  let invoiceRecords = [];

  /*
    B2B
  */

  if (
    workbook.SheetNames.includes(
      "B2B"
    )
  ) {
    invoiceRecords.push(
      ...parseB2B(
        workbook,
        "B2B"
      )
    );
  }

  /*
    B2BA
  */

  const b2ba =
    parseB2BA(workbook);

  /*
    IMPORTANT:
    B2BA is an amendment.

    If an amended invoice has the same
    original invoice number as a B2B
    record, the original B2B record should
    not be blindly counted again.

    Remove the B2B original when a B2BA
    explicitly identifies it.
  */

  if (b2ba.length) {

    const amendedOriginalKeys =
      new Set(
        b2ba
          .filter(
            (r) =>
              r.gstin &&
              r.originalNormalizedInvoice
          )
          .map(
            (r) =>
              `${r.gstin}|${r.originalNormalizedInvoice}`
          )
      );

    invoiceRecords =
      invoiceRecords.filter(
        (record) => {

          if (
            record.section !==
            "B2B"
          ) {
            return true;
          }

          const key =
            `${record.gstin}|${record.normalizedInvoice}`;

          return !amendedOriginalKeys.has(
            key
          );
        }
      );

    invoiceRecords.push(
      ...b2ba
    );
  }

  /*
    Credit / Debit notes
  */

  const creditNotes = [
    "B2B-CDNR",
    "B2B-CDNRA",
  ].flatMap(
    (sheet) =>
      parseCDNR(
        workbook,
        sheet
      )
  );

  /*
    Imports
  */

  const imports = [
    "IMPG",
    "IMPGA",
    "IMPGSEZ",
    "IMPGSEZA",
  ].flatMap(
    (sheet) =>
      parseImport(
        workbook,
        sheet
      )
  );

  return {
    invoiceRecords,
    creditNotes,
    imports,
  };
}


/* =======================================================
   DUPLICATES
======================================================= */

function getDuplicates(
  records
) {
  const map = new Map();

  records.forEach(
    (record) => {

      if (
        !record.normalizedInvoice
      ) {
        return;
      }

      /*
        Prefer GSTIN + invoice.

        If GSTIN is unavailable,
        invoice number alone is used,
        but this is clearly displayed.
      */

      const key =
        record.gstin
          ? `${record.gstin}|${record.normalizedInvoice}`
          : `NO-GSTIN|${record.normalizedInvoice}`;

      if (!map.has(key)) {
        map.set(key, []);
      }

      map
        .get(key)
        .push(record);
    }
  );

  const result = [];

  map.forEach(
    (items) => {

      if (
        items.length > 1
      ) {

        items.forEach(
          (record) => {

            result.push({
              ...record,

              duplicateCount:
                items.length,
            });

          }
        );
      }
    }
  );

  return result;
}


/* =======================================================
   MATCHING ENGINE
======================================================= */

function reconcileBooks(
  books,
  twoB
) {
  const results = [];

  /*
    Build exact GSTIN + invoice index. Amount tolerance is checked after the key is found.
  */

  const exactIndex =
    new Map();

  twoB.forEach(
    (record, index) => {

      if (
        !record.normalizedInvoice
      ) {
        return;
      }

      const exactKey =
        `${record.gstin}|${record.normalizedInvoice}`;

      if (
        !exactIndex.has(
          exactKey
        )
      ) {
        exactIndex.set(
          exactKey,
          []
        );
      }

      exactIndex
        .get(exactKey)
        .push({
          record,
          index,
        });

    }
  );

  const usedTwoB =
    new Set();

  books.forEach(
    (book) => {

      /*
        No invoice number.
      */

      if (
        !book.normalizedInvoice
      ) {

        results.push({
          id:
            `RESULT-${results.length + 1}`,

          status:
            "REVIEW",

          books:
            book,

          twoB:
            null,

          reason:
            "Books record does not contain a supplier invoice number.",
        });

        return;
      }

      /*
        Main matching requires GSTIN.
        Invoice number alone is never used as a match.
      */

      if (!book.gstin) {
        results.push({
          id: `RESULT-${results.length + 1}`,
          status: "REVIEW",
          books: book,
          twoB: null,
          reason: "Books record does not contain a GSTIN. GSTIN is required for automatic matching.",
        });
        return;
      }

      /*
        Exact GSTIN + invoice.
      */

      const exactKey =
        `${book.gstin}|${book.normalizedInvoice}`;

      let candidates =
        exactIndex.get(
          exactKey
        ) || [];

      let matchingMethod =
        "GSTIN + Invoice Number + Amount ± ₹1";

      /*
        MAIN MATCHING RULE
        ------------------
        1. GSTIN must match
        2. Invoice number must match
        3. Invoice amount must be within +/- ₹1

        Invoice date, taxable value and individual tax
        components are NOT used to decide whether the
        invoice is a MATCHED record. They are displayed
        separately so the user can inspect every value.
      */

      const unusedCandidates =
        candidates.filter(
          (item) => !usedTwoB.has(item.index)
        );

      const amountMatchedCandidates =
        unusedCandidates.filter((item) => {
          const twoBValue = item.record.invoiceValue;
          return (
            book.grossTotal !== null &&
            twoBValue !== null &&
            equalAmount(book.grossTotal, twoBValue, 1)
          );
        });

      let candidate = null;

      if (amountMatchedCandidates.length === 1) {
        candidate = amountMatchedCandidates[0];
      } else if (amountMatchedCandidates.length > 1) {
        results.push({
          id: `RESULT-${results.length + 1}`,
          status: "REVIEW",
          books: book,
          twoB: null,
          reason:
            "Multiple GSTR-2B records have the same GSTIN and invoice number and more than one is within the ₹1 amount tolerance. The system did not automatically choose one.",
        });
        return;
      } else if (unusedCandidates.length === 1) {
        // Keep the same GSTIN + invoice record for a PARTIAL result
        // when the amount does not satisfy the +/- ₹1 rule.
        candidate = unusedCandidates[0];
      } else if (unusedCandidates.length > 1) {
        results.push({
          id: `RESULT-${results.length + 1}`,
          status: "REVIEW",
          books: book,
          twoB: null,
          reason:
            "Multiple GSTR-2B records have the same GSTIN and invoice number, and none could be uniquely selected using the ₹1 amount tolerance.",
        });
        return;
      }

      /*
        No GSTIN + invoice candidate.
      */

      if (!candidate) {
        results.push({
          id: `RESULT-${results.length + 1}`,
          status: "NOT IN 2B",
          books: book,
          twoB: null,
          reason:
            "No GSTR-2B record was found with the same GSTIN and invoice number.",
        });
        return;
      }

      const gstRecord = candidate.record;

      usedTwoB.add(candidate.index);

      /*
        -----------------------------------------------
        IDENTIFICATION COMPARISON
        -----------------------------------------------
      */

      const invoiceMatch =
        book.normalizedInvoice ===
        gstRecord.normalizedInvoice;

      const bothGSTINAvailable =
        Boolean(
          book.gstin &&
          gstRecord.gstin
        );

      const gstinMatch =
        bothGSTINAvailable &&
        book.gstin ===
          gstRecord.gstin;

      /*
        -----------------------------------------------
        INVOICE VALUE
        -----------------------------------------------

        A record is marked MATCHED only when the
        invoice value is also available on both sides
        and matches within ₹0.01.
      */

      const invoiceValueAvailable =
        book.grossTotal !== null &&
        gstRecord.invoiceValue !== null;

      const invoiceValueMatch =
        invoiceValueAvailable &&
        equalAmount(
          book.grossTotal,
          gstRecord.invoiceValue,
          1
        );

      const invoiceValueDifference =
        difference(
          book.grossTotal,
          gstRecord.invoiceValue
        );

      /*
        -----------------------------------------------
        TAXABLE VALUE
        -----------------------------------------------
      */

      const taxableAvailable =
        book.taxableValue !== null &&
        gstRecord.taxableValue !==
          null;

      const taxableMatch =
        taxableAvailable &&
        equalAmount(
          book.taxableValue,
          gstRecord.taxableValue
        );

      const taxableDifference =
        difference(
          book.taxableValue,
          gstRecord.taxableValue
        );

      /*
        -----------------------------------------------
        GST TAX
        -----------------------------------------------
      */

      const taxAvailable =
        book.totalTax !== null &&
        gstRecord.totalTax !==
          null;

      const taxMatch =
        taxAvailable &&
        equalAmount(
          book.totalTax,
          gstRecord.totalTax
        );

      const taxDifference =
        difference(
          book.totalTax,
          gstRecord.totalTax
        );

      /*
        -----------------------------------------------
        INDIVIDUAL GST COMPONENTS
        -----------------------------------------------
      */

      /*
        A blank tax component is treated as ₹0 only for
        component comparison. This is necessary because
        Books commonly leaves non-applicable tax columns
        blank while GSTR-2B records them as 0.
      */

      const taxComponentEqual =
        (booksValue, twoBValue) =>
          equalAmount(
            booksValue === null || booksValue === undefined
              ? 0
              : booksValue,
            twoBValue === null || twoBValue === undefined
              ? 0
              : twoBValue
          );

      const igstMatch =
        taxComponentEqual(
          book.igst,
          gstRecord.igst
        );

      const cgstMatch =
        taxComponentEqual(
          book.cgst,
          gstRecord.cgst
        );

      const sgstMatch =
        taxComponentEqual(
          book.sgst,
          gstRecord.sgst
        );

      const cessMatch =
        taxComponentEqual(
          book.cess,
          gstRecord.cess
        );

      /*
        -----------------------------------------------
        STATUS
        -----------------------------------------------
      */

      let status =
        "PARTIAL";

      let reason =
        "Invoice found, but one or more reconciliation fields require review.";

      /*
        MATCHED is based ONLY on the user's three main rules:
          1. GSTIN
          2. Invoice number
          3. Invoice value within +/- ₹1

        Invoice date, taxable value, IGST, CGST, SGST, Cess and total
        GST are NOT matching conditions. They remain visible
        in the detailed MATCHED table for inspection.
      */

      if (
        invoiceMatch &&
        gstinMatch &&
        invoiceValueAvailable &&
        invoiceValueMatch
      ) {

        status = "MATCHED";

        reason =
          "MATCHED: GSTIN and invoice number match, and invoice amount is within +/- ₹1. Invoice date and other tax values are displayed for verification but do not decide the match.";

      } else {

        if (!bothGSTINAvailable) {
          reason =
            "Invoice number was found, but GSTIN is unavailable in one or both records.";
        } else if (!gstinMatch) {
          reason =
            "Invoice number matched, but GSTIN differs.";
        } else if (!invoiceValueAvailable) {
          reason =
            "GSTIN and invoice number matched, but invoice amount is unavailable in one or both records.";
        } else if (!invoiceValueMatch) {
          reason =
            `GSTIN and invoice number matched, but invoice amount differs by more than ₹1 (difference: ₹${Math.abs(invoiceValueDifference ?? 0).toFixed(2)}).`;
        }
      }

      results.push({
        id:
          `RESULT-${results.length + 1}`,

        status,

        books:
          book,

        twoB:
          gstRecord,

        matchingMethod,

        invoiceMatch,

        gstinMatch,

        invoiceValueMatch,

        invoiceValueDifference,

        taxableMatch,

        taxMatch,

        igstMatch,

        cgstMatch,

        sgstMatch,

        cessMatch,

        taxableDifference,

        taxDifference,

        reason,
      });
    }
  );

  /*
    Any unused 2B record is not in Books.
  */

  twoB.forEach(
    (record, index) => {

      if (
        usedTwoB.has(index)
      ) {
        return;
      }

      results.push({
        id:
          `RESULT-${results.length + 1}`,

        status:
          "NOT IN BOOKS",

        books:
          null,

        twoB:
          record,

        taxableDifference:
          record.taxableValue,

        taxDifference:
          record.totalTax,

        reason:
          "This GSTR-2B invoice was not found in the Books Purchase Register.",
      });
    }
  );

  return results;
}


/* =======================================================
   STATUS BADGE
======================================================= */

function StatusBadge({
  status,
}) {
  const style = {
    MATCHED:
      "bg-green-100 text-green-700",

    PARTIAL:
      "bg-yellow-100 text-yellow-700",

    "NOT IN BOOKS":
      "bg-red-100 text-red-700",

    "NOT IN 2B":
      "bg-red-100 text-red-700",

    REVIEW:
      "bg-orange-100 text-orange-700",

    BOOKS:
      "bg-blue-100 text-blue-700",

    "GSTR-2B":
      "bg-purple-100 text-purple-700",

    DUPLICATE:
      "bg-orange-100 text-orange-700",
  };

  return (
    <span
      className={`inline-flex items-center px-3 py-1 rounded-full text-xs font-semibold ${
        style[status] ||
        "bg-gray-100 text-gray-700"
      }`}
    >
      {status}
    </span>
  );
}


/* =======================================================
   SIDEBAR
======================================================= */

const menu = [
  ["start", "START HERE"],
  ["books", "BOOKS INPUT"],
  ["2b", "2B INPUT"],
  ["recon", "RECON"],
  ["summary", "SUMMARY"],
  ["matched", "MATCHED"],
  ["partial", "PARTIAL"],
  ["duplicate", "DUPLICATE"],
  ["not-books", "NOT IN BOOKS"],
  ["not-2b", "NOT IN 2B"],
  ["rcm", "RCM"],
  ["imports", "IMPORTS"],
  ["credit-notes", "DR CREDIT NOTES"],
];


function Sidebar({
  active,
  setActive,
}) {
  return (
    <aside className="fixed left-0 top-0 w-64 h-screen bg-white border-r border-gray-200 z-50">

      <div className="h-full flex flex-col">

        <div className="p-6 border-b border-gray-200">

          <div className="text-2xl font-bold text-green-700">
            GST RECO
          </div>

          <div className="text-xs text-gray-500 mt-1">
            Smart GST Reconciliation
          </div>

        </div>

        <nav className="p-3 flex-1">

          {menu.map(
            ([id, label], index) => (

              <button
                key={id}
                onClick={() =>
                  setActive(id)
                }
                className={`w-full flex items-center gap-3 px-3 py-2.5 mb-1 rounded-lg text-left ${
                  active === id
                    ? "bg-green-100 text-green-700 font-semibold"
                    : "text-gray-700 hover:bg-gray-100"
                }`}
              >

                <span className="w-7 text-xs font-bold">
                  {index + 1}.
                </span>

                <span className="text-xs">
                  {label}
                </span>

              </button>

            )
          )}

        </nav>

      </div>

    </aside>
  );
}


/* =======================================================
   SUMMARY CARD
======================================================= */

function Card({
  title,
  value,
  subtitle,
  className = "",
}) {
  return (
    <div
      className={`bg-white border rounded-xl p-5 ${className}`}
    >

      <div className="text-xs text-gray-500 font-semibold">
        {title}
      </div>

      <div className="text-2xl font-bold mt-2">
        {value}
      </div>

      {subtitle && (
        <div className="text-xs text-gray-400 mt-1">
          {subtitle}
        </div>
      )}

    </div>
  );
}


/* =======================================================
   UPLOAD CARD
======================================================= */

function UploadCard({
  title,
  subtitle,
  file,
  onFile,
  color,
}) {
  const [dragging, setDragging] =
    useState(false);

  function handleDrop(event) {
    event.preventDefault();

    setDragging(false);

    const dropped =
      event.dataTransfer.files?.[0];

    if (dropped) {
      onFile(dropped);
    }
  }

  return (
    <div className="bg-white border rounded-xl p-6">

      <div className="text-center">

        <div
          className={`mx-auto w-12 h-12 rounded-xl flex items-center justify-center ${
            color === "purple"
              ? "bg-purple-50"
              : "bg-blue-50"
          }`}
        >
          📄
        </div>

        <h3 className="font-bold mt-3">
          {title}
        </h3>

        <p className="text-xs text-gray-500 mt-1">
          {subtitle}
        </p>

      </div>

      <div
        onDragOver={(event) => {
          event.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() =>
          setDragging(false)
        }
        onDrop={handleDrop}
        className={`mt-5 border-2 border-dashed rounded-xl p-8 text-center ${
          dragging
            ? "border-green-500 bg-green-50"
            : "border-gray-300 bg-gray-50"
        }`}
      >

        <div className="text-sm text-gray-700">
          Drag & Drop your file
        </div>

        <div className="text-xs text-gray-400 my-2">
          or
        </div>

        <label
          className={`inline-block cursor-pointer px-5 py-2 rounded-lg text-white text-sm font-semibold ${
            color === "purple"
              ? "bg-purple-600 hover:bg-purple-700"
              : "bg-blue-600 hover:bg-blue-700"
          }`}
        >
          Browse File

          <input
            type="file"
            accept=".xls,.xlsx,.csv"
            className="hidden"
            onChange={(event) => {

              const selected =
                event.target.files?.[0];

              if (selected) {
                onFile(selected);
              }

              event.target.value = "";
            }}
          />

        </label>

        {file && (
          <div className="mt-4 bg-green-50 border border-green-200 rounded-lg p-3">

            <div className="text-sm font-semibold text-green-700">
              ✓ {file.name}
            </div>

            <div className="text-xs text-green-600 mt-1">
              {file.entries.toLocaleString(
                "en-IN"
              )} records loaded
            </div>

          </div>
        )}

        <div className="text-xs text-gray-400 mt-4">
          .xls / .xlsx / .csv supported
        </div>

      </div>

    </div>
  );
}


/* =======================================================
   SEARCH + PAGINATION
======================================================= */

function SearchBox({
  value,
  setValue,
}) {
  return (
    <input
      value={value}
      onChange={(event) =>
        setValue(
          event.target.value
        )
      }
      placeholder="Search invoice, GSTIN, supplier..."
      className="w-full md:w-96 border border-gray-300 rounded-lg px-4 py-2 text-sm outline-none focus:ring-2 focus:ring-green-200"
    />
  );
}


function Pagination({
  page,
  setPage,
  total,
  pageSize,
  setPageSize,
}) {
  const actualPageSize =
    pageSize === "ALL"
      ? Math.max(total, 1)
      : Number(pageSize);

  const totalPages = Math.max(
    1,
    Math.ceil(total / actualPageSize)
  );

  const pageSizeOptions = [
    50,
    100,
    150,
    200,
    250,
    300,
    350,
    400,
    450,
    500,
    1000,
    "ALL",
  ];

  return (
    <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3 w-full">
      <div className="flex items-center gap-2">
        <span className="text-xs text-gray-500">Show</span>
        <select
          value={pageSize}
          onChange={(event) => {
            const value =
              event.target.value === "ALL"
                ? "ALL"
                : Number(event.target.value);
            setPageSize(value);
            setPage(1);
          }}
          className="border border-gray-300 rounded-lg px-3 py-1.5 text-xs bg-white focus:outline-none focus:ring-2 focus:ring-green-200"
        >
          {pageSizeOptions.map((size) => (
            <option key={size} value={size}>
              {size === "ALL" ? "All" : size}
            </option>
          ))}
        </select>
        <span className="text-xs text-gray-500">
          records per page
        </span>
      </div>

      <div className="flex items-center gap-2">
        <span className="text-xs text-gray-500 mr-2">
          Page {page} of {totalPages}
        </span>
        <button
          disabled={page <= 1}
          onClick={() =>
            setPage((p) => Math.max(1, p - 1))
          }
          className="px-3 py-1.5 border rounded-lg text-xs bg-white hover:bg-gray-50 disabled:opacity-40 disabled:cursor-not-allowed"
        >
          Previous
        </button>
        <button
          disabled={page >= totalPages}
          onClick={() =>
            setPage((p) => Math.min(totalPages, p + 1))
          }
          className="px-3 py-1.5 border rounded-lg text-xs bg-white hover:bg-gray-50 disabled:opacity-40 disabled:cursor-not-allowed"
        >
          Next
        </button>
      </div>
    </div>
  );
}

/* =======================================================
   DATA TABLE
======================================================= */

function DataTable({
  records,
  onSelect,
  search,
  setSearch,
}) {
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);

  const [gstinFilter, setGstinFilter] = useState("");
  const [supplierFilter, setSupplierFilter] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");

  const gstins = useMemo(() => {
    return [
      ...new Set(
        records
          .map((record) => record.gstin)
          .filter(Boolean)
      ),
    ].sort();
  }, [records]);

  const suppliers = useMemo(() => {
    return [
      ...new Set(
        records
          .map((record) => record.supplier)
          .filter(Boolean)
      ),
    ].sort();
  }, [records]);

  const filtered = useMemo(() => {
    const query = normalizedText(search);

    return records.filter((record) => {
      if (query) {
        const values = [
          record.gstin,
          record.invoiceNo,
          record.noteNo,
          record.supplier,
          record.particulars,
          record.section,
          record.invoiceDate,
        ];

        const matchesSearch = values.some((value) =>
          normalizedText(value).includes(query)
        );

        if (!matchesSearch) return false;
      }

      if (
        gstinFilter &&
        normalizedText(record.gstin) !== normalizedText(gstinFilter)
      ) {
        return false;
      }

      if (
        supplierFilter &&
        normalizedText(record.supplier) !== normalizedText(supplierFilter)
      ) {
        return false;
      }

      const recordDate = dateValue(
        record.invoiceDate ||
        record.noteDate ||
        record.billOfEntryDate ||
        ""
      );

      if (dateFrom && (!recordDate || recordDate < dateFrom)) {
        return false;
      }

      if (dateTo && (!recordDate || recordDate > dateTo)) {
        return false;
      }

      return true;
    });
  }, [
    records,
    search,
    gstinFilter,
    supplierFilter,
    dateFrom,
    dateTo,
  ]);

  const actualPageSize =
    pageSize === "ALL"
      ? Math.max(filtered.length, 1)
      : Number(pageSize);

  const totalPages = Math.max(
    1,
    Math.ceil(filtered.length / actualPageSize)
  );

  const safePage = Math.min(page, totalPages);

  const visible = filtered.slice(
    (safePage - 1) * actualPageSize,
    safePage * actualPageSize
  );

  function clearFilters() {
    setGstinFilter("");
    setSupplierFilter("");
    setDateFrom("");
    setDateTo("");
    setPage(1);
  }

  const filtersApplied =
    gstinFilter || supplierFilter || dateFrom || dateTo;

  return (
    <div>
      <div className="mb-4">
        <SearchBox
          value={search}
          setValue={(value) => {
            setSearch(value);
            setPage(1);
          }}
        />
      </div>

      <div className="bg-white border rounded-xl p-4 mb-4">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              GSTIN
            </label>
            <select
              value={gstinFilter}
              onChange={(event) => {
                setGstinFilter(event.target.value);
                setPage(1);
              }}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm bg-white"
            >
              <option value="">All GSTINs</option>
              {gstins.map((gstin) => {
                const supplier =
                  records.find((record) => record.gstin === gstin)?.supplier ||
                  "Supplier name unavailable";
                return (
                  <option key={gstin} value={gstin}>
                    {gstin} — {supplier}
                  </option>
                );
              })}
            </select>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              Supplier
            </label>
            <select
              value={supplierFilter}
              onChange={(event) => {
                setSupplierFilter(event.target.value);
                setPage(1);
              }}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm bg-white"
            >
              <option value="">All Suppliers</option>
              {suppliers.map((supplier) => (
                <option key={supplier} value={supplier}>
                  {supplier}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              Date From
            </label>
            <input
              type="date"
              value={dateFrom}
              onChange={(event) => {
                setDateFrom(event.target.value);
                setPage(1);
              }}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm"
            />
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              Date To
            </label>
            <input
              type="date"
              value={dateTo}
              onChange={(event) => {
                setDateTo(event.target.value);
                setPage(1);
              }}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm"
            />
          </div>

          <div className="flex items-end">
            {filtersApplied ? (
              <button
                onClick={clearFilters}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg text-sm bg-white hover:bg-gray-50"
              >
                Clear Filters
              </button>
            ) : (
              <div className="text-xs text-gray-400 pb-2">
                Use filters to narrow the records.
              </div>
            )}
          </div>
        </div>

        {filtersApplied && (
          <div className="mt-3 text-xs text-gray-500">
            Showing {filtered.length.toLocaleString("en-IN")} filtered records
          </div>
        )}
      </div>

      <div className="bg-white border rounded-xl overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b">
              <tr>
                <th className="text-left p-3 whitespace-nowrap">S.NO.</th>
                <th className="text-left p-3 whitespace-nowrap">GSTIN</th>
                <th className="text-left p-3 whitespace-nowrap">INVOICE</th>
                <th className="text-left p-3 whitespace-nowrap">SUPPLIER NAME</th>
                <th className="text-right p-3 whitespace-nowrap">TAXABLE VALUE</th>
                <th className="text-right p-3 whitespace-nowrap">IGST</th>
                <th className="text-right p-3 whitespace-nowrap">CGST</th>
                <th className="text-right p-3 whitespace-nowrap">SGST</th>
                <th className="text-right p-3 whitespace-nowrap">TOTAL TAX</th>
              </tr>
            </thead>
            <tbody>
              {visible.map((record, index) => {
                const serial =
                  (safePage - 1) * actualPageSize + index + 1;

                return (
                  <tr
                    key={record.id || serial}
                    className="border-b hover:bg-gray-50"
                  >
                    <td className="p-3 font-semibold text-gray-600">{serial}</td>
                    <td className="p-3 font-mono text-xs whitespace-nowrap">
                      {record.gstin || "—"}
                    </td>
                    <td className="p-3 whitespace-nowrap">
                      <button
                        onClick={() => onSelect(record)}
                        className="text-blue-600 hover:underline font-semibold"
                      >
                        {record.invoiceNo ||
                          record.noteNo ||
                          record.billOfEntryNo ||
                          "—"}
                      </button>
                    </td>
                    <td className="p-3 whitespace-nowrap">
                      {record.supplier || "—"}
                    </td>
                    <td className="p-3 whitespace-nowrap">
                      {record.invoiceDate ||
                        record.noteDate ||
                        record.billOfEntryDate ||
                        "—"}
                    </td>
                    <td className="p-3 text-right whitespace-nowrap">
                      {amount(record.taxableValue)}
                    </td>
                    <td className="p-3 text-right whitespace-nowrap">
                      {amount(record.igst)}
                    </td>
                    <td className="p-3 text-right whitespace-nowrap">
                      {amount(record.cgst)}
                    </td>
                    <td className="p-3 text-right whitespace-nowrap">
                      {amount(record.sgst)}
                    </td>
                    <td className="p-3 text-right whitespace-nowrap font-semibold">
                      {amount(record.totalTax)}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {!visible.length && (
            <div className="p-12 text-center text-gray-400">
              No records found.
            </div>
          )}
        </div>

        <div className="p-4 border-t">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div className="text-xs text-gray-500">
              Showing {filtered.length ? ((safePage - 1) * actualPageSize + 1).toLocaleString("en-IN") : 0}
              {" – "}
              {Math.min(safePage * actualPageSize, filtered.length).toLocaleString("en-IN")}
              {" of "}
              {filtered.length.toLocaleString("en-IN")}
              {" records"}
            </div>

            <Pagination
              page={safePage}
              setPage={setPage}
              total={filtered.length}
              pageSize={pageSize}
              setPageSize={setPageSize}
            />
          </div>
        </div>
      </div>
    </div>
  );
}

/* =======================================================
   RECON TABLE
======================================================= */

function ReconTable({
  records,
  onSelect,
  search,
  setSearch,
}) {
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);

  const [statusFilter, setStatusFilter] = useState("");
  const [gstinFilter, setGstinFilter] = useState("");
  const [supplierFilter, setSupplierFilter] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");

  const gstins = useMemo(() => {
    return [
      ...new Set(
        records
          .map(
            (result) =>
              result.books?.gstin || result.twoB?.gstin
          )
          .filter(Boolean)
      ),
    ].sort();
  }, [records]);

  const suppliers = useMemo(() => {
    return [
      ...new Set(
        records
          .map(
            (result) =>
              result.books?.supplier || result.twoB?.supplier
          )
          .filter(Boolean)
      ),
    ].sort();
  }, [records]);

  const filtered = useMemo(() => {
    const query = normalizedText(search);

    return records.filter((result) => {
      if (query) {
        const values = [
          result.books?.gstin,
          result.twoB?.gstin,
          result.books?.invoiceNo,
          result.twoB?.invoiceNo,
          result.books?.supplier,
          result.twoB?.supplier,
          result.status,
        ];

        const matchesSearch = values.some((value) =>
          normalizedText(value).includes(query)
        );

        if (!matchesSearch) return false;
      }

      if (statusFilter && result.status !== statusFilter) {
        return false;
      }

      const gstin =
        result.books?.gstin || result.twoB?.gstin || "";

      if (
        gstinFilter &&
        normalizedText(gstin) !== normalizedText(gstinFilter)
      ) {
        return false;
      }

      const supplier =
        result.books?.supplier || result.twoB?.supplier || "";

      if (
        supplierFilter &&
        normalizedText(supplier) !== normalizedText(supplierFilter)
      ) {
        return false;
      }

      const recordDate = dateValue(
        result.books?.invoiceDate ||
        result.twoB?.invoiceDate ||
        ""
      );

      if (dateFrom && (!recordDate || recordDate < dateFrom)) {
        return false;
      }

      if (dateTo && (!recordDate || recordDate > dateTo)) {
        return false;
      }

      return true;
    });
  }, [
    records,
    search,
    statusFilter,
    gstinFilter,
    supplierFilter,
    dateFrom,
    dateTo,
  ]);

  const actualPageSize =
    pageSize === "ALL"
      ? Math.max(filtered.length, 1)
      : Number(pageSize);

  const totalPages = Math.max(
    1,
    Math.ceil(filtered.length / actualPageSize)
  );

  const safePage = Math.min(page, totalPages);

  const visible = filtered.slice(
    (safePage - 1) * actualPageSize,
    safePage * actualPageSize
  );

  function clearFilters() {
    setStatusFilter("");
    setGstinFilter("");
    setSupplierFilter("");
    setDateFrom("");
    setDateTo("");
    setPage(1);
  }

  const filtersApplied =
    statusFilter ||
    gstinFilter ||
    supplierFilter ||
    dateFrom ||
    dateTo;

  return (
    <div>
      <div className="mb-4">
        <SearchBox
          value={search}
          setValue={(value) => {
            setSearch(value);
            setPage(1);
          }}
        />
      </div>

      <div className="bg-white border rounded-xl p-4 mb-4">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              Status
            </label>
            <select
              value={statusFilter}
              onChange={(event) => {
                setStatusFilter(event.target.value);
                setPage(1);
              }}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm bg-white"
            >
              <option value="">All Status</option>
              <option value="MATCHED">MATCHED</option>
              <option value="PARTIAL">PARTIAL</option>
              <option value="REVIEW">REVIEW</option>
              <option value="NOT IN BOOKS">NOT IN BOOKS</option>
              <option value="NOT IN 2B">NOT IN 2B</option>
            </select>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              GSTIN
            </label>
            <select
              value={gstinFilter}
              onChange={(event) => {
                setGstinFilter(event.target.value);
                setPage(1);
              }}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm bg-white"
            >
              <option value="">All GSTINs</option>
              {gstins.map((gstin) => {
                const matchingResult = records.find(
                  (result) =>
                    (result.books?.gstin || result.twoB?.gstin) === gstin
                );
                const supplier =
                  matchingResult?.books?.supplier ||
                  matchingResult?.twoB?.supplier ||
                  "Supplier name unavailable";
                return (
                  <option key={gstin} value={gstin}>
                    {gstin} — {supplier}
                  </option>
                );
              })}
            </select>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              Supplier
            </label>
            <select
              value={supplierFilter}
              onChange={(event) => {
                setSupplierFilter(event.target.value);
                setPage(1);
              }}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm bg-white"
            >
              <option value="">All Suppliers</option>
              {suppliers.map((supplier) => (
                <option key={supplier} value={supplier}>
                  {supplier}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              Date From
            </label>
            <input
              type="date"
              value={dateFrom}
              onChange={(event) => {
                setDateFrom(event.target.value);
                setPage(1);
              }}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm"
            />
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              Date To
            </label>
            <input
              type="date"
              value={dateTo}
              onChange={(event) => {
                setDateTo(event.target.value);
                setPage(1);
              }}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm"
            />
          </div>
        </div>

        {filtersApplied && (
          <div className="mt-4 flex items-center justify-between gap-4">
            <div className="text-xs text-gray-500">
              Showing {filtered.length.toLocaleString("en-IN")} filtered records
            </div>
            <button
              onClick={clearFilters}
              className="px-4 py-2 border border-gray-300 rounded-lg text-sm bg-white hover:bg-gray-50"
            >
              Clear Filters
            </button>
          </div>
        )}
      </div>

      <div className="bg-white border rounded-xl overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b">
              <tr>
                <th className="text-left p-3 whitespace-nowrap">S.NO.</th>
                <th className="text-left p-3 whitespace-nowrap">STATUS</th>
                <th className="text-left p-3 whitespace-nowrap">GSTIN</th>
                <th className="text-left p-3 whitespace-nowrap">SUPPLIER NAME</th>
                <th className="text-left p-3 whitespace-nowrap">DATE</th>
                <th className="text-left p-3 whitespace-nowrap">BOOKS INVOICE</th>
                <th className="text-left p-3 whitespace-nowrap">2B INVOICE</th>
                <th className="text-right p-3 whitespace-nowrap">BOOKS TAXABLE</th>
                <th className="text-right p-3 whitespace-nowrap">2B TAXABLE</th>
                <th className="text-right p-3 whitespace-nowrap">DIFFERENCE</th>
                <th className="text-right p-3 whitespace-nowrap">TAX DIFF.</th>
              </tr>
            </thead>
            <tbody>
              {visible.map((result, index) => {
                const serial =
                  (safePage - 1) * actualPageSize + index + 1;

                const gstin =
                  result.books?.gstin ||
                  result.twoB?.gstin ||
                  "—";

                return (
                  <tr
                    key={result.id}
                    className="border-b hover:bg-gray-50"
                  >
                    <td className="p-3 font-semibold text-gray-600">{serial}</td>
                    <td className="p-3">
                      <StatusBadge status={result.status} />
                    </td>
                    <td className="p-3 font-mono text-xs whitespace-nowrap">
                      {gstin}
                    </td>
                    <td className="p-3 whitespace-nowrap">
                      {result.books?.supplier || result.twoB?.supplier || "—"}
                    </td>

                    {/* DATE FIELD - DISPLAY ONLY */}
                    <td className="p-3 whitespace-nowrap">
                      <div className="text-xs leading-5">
                        <div className="font-semibold text-gray-800">Date</div>
                        <div>
                          <span className="font-semibold">Books:</span>{" "}
                          <span className="font-mono">
                            {result.books?.invoiceDate
                              ? dateValue(result.books.invoiceDate)
                              : "—"}
                          </span>
                        </div>
                        <div>
                          <span className="font-semibold">2B:</span>{" "}
                          <span className="font-mono">
                            {result.twoB?.invoiceDate
                              ? dateValue(result.twoB.invoiceDate)
                              : "—"}
                          </span>
                        </div>
                      </div>
                    </td>

                    <td className="p-3 whitespace-nowrap">
                      {result.books ? (
                        <button
                          onClick={() => onSelect(result)}
                          className="text-blue-600 hover:underline font-semibold"
                        >
                          {result.books.invoiceNo}
                        </button>
                      ) : (
                        <span className="text-red-500">Missing</span>
                      )}
                    </td>
                    <td className="p-3 whitespace-nowrap">
                      {result.twoB ? (
                        <button
                          onClick={() => onSelect(result)}
                          className="text-purple-600 hover:underline font-semibold"
                        >
                          {result.twoB.invoiceNo}
                        </button>
                      ) : (
                        <span className="text-red-500">Missing</span>
                      )}
                    </td>
                    <td className="p-3 text-right whitespace-nowrap">
                      {amount(result.books?.taxableValue)}
                    </td>
                    <td className="p-3 text-right whitespace-nowrap">
                      {amount(result.twoB?.taxableValue)}
                    </td>
                    <td className="p-3 text-right whitespace-nowrap">
                      {amount(result.taxableDifference)}
                    </td>
                    <td className="p-3 text-right whitespace-nowrap">
                      {amount(result.taxDifference)}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {!visible.length && (
            <div className="p-12 text-center text-gray-400">
              No reconciliation records found.
            </div>
          )}
        </div>

        <div className="p-4 border-t">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div className="text-xs text-gray-500">
              Showing {filtered.length ? ((safePage - 1) * actualPageSize + 1).toLocaleString("en-IN") : 0}
              {" – "}
              {Math.min(safePage * actualPageSize, filtered.length).toLocaleString("en-IN")}
              {" of "}
              {filtered.length.toLocaleString("en-IN")}
            </div>

            <Pagination
              page={safePage}
              setPage={setPage}
              total={filtered.length}
              pageSize={pageSize}
              setPageSize={setPageSize}
            />
          </div>
        </div>
      </div>
    </div>
  );
}

/* =======================================================
   MATCHED DETAILS TABLE

   MATCHED filtering/search/pagination is kept separate from
   the reconciliation matching logic. Date is ONLY a filter here;
   it is never used to decide whether an invoice is matched.
======================================================= */

function MatchedDetailsTable({
  records,
  onSelect,
}) {
  const rows = records || [];

  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);
  const [statusFilter, setStatusFilter] = useState("");
  const [gstinFilter, setGstinFilter] = useState("");
  const [supplierFilter, setSupplierFilter] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [selectedMatchChecks, setSelectedMatchChecks] = useState([]);

  const normalizeFilterDate = (value) => {
    if (!value) return "";
    const raw = String(value).trim();

    // Already ISO: YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

    // DD/MM/YYYY or DD-MM-YYYY
    let match = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if (match) {
      const [, d, m, y] = match;
      return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    }

    // YYYY/MM/DD or YYYY-MM-DD variants
    match = raw.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
    if (match) {
      const [, y, m, d] = match;
      return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    }

    // Excel/JS date values
    const parsed = new Date(raw);
    if (!Number.isNaN(parsed.getTime())) {
      return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
    }

    return "";
  };

  const gstins = useMemo(() => {
    return [
      ...new Set(
        rows
          .map((result) => result.books?.gstin || result.twoB?.gstin)
          .filter(Boolean)
      ),
    ].sort();
  }, [rows]);

  const suppliers = useMemo(() => {
    return [
      ...new Set(
        rows
          .map((result) => result.books?.supplier || result.twoB?.supplier)
          .filter(Boolean)
      ),
    ].sort((a, b) => String(a).localeCompare(String(b)));
  }, [rows]);

  const filtered = useMemo(() => {
    const query = normalizedText(search);

    return rows.filter((result) => {
      const book = result.books;
      const twoB = result.twoB;
      const gstin = book?.gstin || twoB?.gstin || "";
      const supplier = book?.supplier || twoB?.supplier || "";
      const bookDate = normalizeFilterDate(book?.invoiceDate);
      const twoBDate = normalizeFilterDate(twoB?.invoiceDate);

      if (query) {
        const values = [
          book?.gstin,
          twoB?.gstin,
          book?.supplier,
          twoB?.supplier,
          book?.invoiceNo,
          twoB?.invoiceNo,
          book?.invoiceDate,
          twoB?.invoiceDate,
          book?.grossTotal,
          twoB?.invoiceValue,
          book?.taxableValue,
          twoB?.taxableValue,
          book?.igst,
          twoB?.igst,
          book?.cgst,
          twoB?.cgst,
          book?.sgst,
          twoB?.sgst,
          result.status,
        ];

        if (!values.some((value) => normalizedText(value).includes(query))) {
          return false;
        }
      }

      if (statusFilter && result.status !== statusFilter) {
        return false;
      }

      if (
        gstinFilter &&
        normalizedText(gstin) !== normalizedText(gstinFilter)
      ) {
        return false;
      }

      if (
        supplierFilter &&
        normalizedText(supplier) !== normalizedText(supplierFilter)
      ) {
        return false;
      }

      // Date is only used for filtering/display, never for matching.
      // Use either Books date or 2B date so the record remains visible
      // when either source falls inside the selected date range.
      if (dateFrom || dateTo) {
        const dates = [bookDate, twoBDate].filter(Boolean);

        if (!dates.length) return false;

        const passesDateRange = dates.some((recordDate) => {
          if (dateFrom && recordDate < dateFrom) return false;
          if (dateTo && recordDate > dateTo) return false;
          return true;
        });

        if (!passesDateRange) return false;
      }

      /*
        MATCH CHECK FILTER
        ------------------
        These checks use the actual comparison results generated
        by reconcileBooks(). Nothing is hard-coded here.
      */
      const checkMap = {
        INVOICE: Boolean(result.invoiceMatch),
        GSTIN: Boolean(result.gstinMatch),
        INVOICE_GSTIN: Boolean(
          result.invoiceMatch && result.gstinMatch
        ),
        INVOICE_VALUE: Boolean(result.invoiceValueMatch),
        CORE_MATCH: Boolean(
          result.invoiceMatch &&
          result.gstinMatch &&
          result.invoiceValueMatch
        ),
        TAXABLE_VALUE: Boolean(result.taxableMatch),
        IGST: Boolean(result.igstMatch),
        CGST: Boolean(result.cgstMatch),
        SGST: Boolean(result.sgstMatch),
        CESS: Boolean(result.cessMatch),
        TOTAL_TAX: Boolean(result.taxMatch),
        ALL_AVAILABLE: Boolean(
          result.invoiceMatch &&
          result.gstinMatch &&
          result.invoiceValueMatch &&
          (book?.taxableValue == null || twoB?.taxableValue == null || result.taxableMatch) &&
          (book?.igst == null || twoB?.igst == null || result.igstMatch) &&
          (book?.cgst == null || twoB?.cgst == null || result.cgstMatch) &&
          (book?.sgst == null || twoB?.sgst == null || result.sgstMatch) &&
          (book?.cess == null || twoB?.cess == null || result.cessMatch) &&
          (book?.totalTax == null || twoB?.totalTax == null || result.taxMatch)
        ),
      };

      // Multiple selected checks work together (AND logic).
      // Example: Invoice + GSTIN + Taxable Value means that
      // all three selected checks must pass for the row to remain visible.
      if (
        selectedMatchChecks.length > 0 &&
        !selectedMatchChecks.every((check) => checkMap[check])
      ) {
        return false;
      }

      return true;
    });
  }, [
    rows,
    search,
    statusFilter,
    gstinFilter,
    supplierFilter,
    dateFrom,
    dateTo,
    selectedMatchChecks,
  ]);

  const actualPageSize =
    pageSize === "ALL" ? Math.max(filtered.length, 1) : Number(pageSize);

  const totalPages = Math.max(
    1,
    Math.ceil(filtered.length / actualPageSize)
  );

  const safePage = Math.min(page, totalPages);

  const visible = filtered.slice(
    (safePage - 1) * actualPageSize,
    safePage * actualPageSize
  );

  function resetPage() {
    setPage(1);
  }

  function clearFilters() {
    setSearch("");
    setStatusFilter("");
    setGstinFilter("");
    setSupplierFilter("");
    setDateFrom("");
    setDateTo("");
    setSelectedMatchChecks([]);
    setPage(1);
  }

  const filtersApplied = Boolean(
    search ||
    statusFilter ||
    gstinFilter ||
    supplierFilter ||
    dateFrom ||
    dateTo ||
    selectedMatchChecks.length > 0
  );

  return (
    <div>
      {/* Search */}
      <div className="mb-4">
        <SearchBox
          value={search}
          setValue={(value) => {
            setSearch(value);
            resetPage();
          }}
        />
      </div>

      {/* MATCHED filters */}
      <div className="bg-white border border-black rounded-xl p-3 mb-4">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-6 gap-3">
          {/* Status */}
          <div>
            <label className="block text-[11px] text-gray-600 mb-1">
              Status
            </label>
            <select
              value={statusFilter}
              onChange={(event) => {
                setStatusFilter(event.target.value);
                resetPage();
              }}
              className="w-full border border-gray-300 rounded-md px-2 py-2 text-xs bg-white"
            >
              <option value="">All Status</option>
              <option value="MATCHED">MATCHED</option>
              <option value="PARTIAL">PARTIAL</option>
              <option value="REVIEW">REVIEW</option>
              <option value="NOT IN BOOKS">NOT IN BOOKS</option>
              <option value="NOT IN 2B">NOT IN 2B</option>
            </select>
          </div>

          {/* GSTIN */}
          <div>
            <label className="block text-[11px] text-gray-600 mb-1">
              GSTIN
            </label>
            <select
              value={gstinFilter}
              onChange={(event) => {
                setGstinFilter(event.target.value);
                resetPage();
              }}
              className="w-full border border-gray-300 rounded-md px-2 py-2 text-xs bg-white"
            >
              <option value="">All GSTINs</option>
              {gstins.map((gstin) => {
                const matchingResult = rows.find(
                  (result) =>
                    (result.books?.gstin || result.twoB?.gstin) === gstin
                );
                const supplier =
                  matchingResult?.books?.supplier ||
                  matchingResult?.twoB?.supplier ||
                  "Supplier name unavailable";

                return (
                  <option key={gstin} value={gstin}>
                    {gstin} — {supplier}
                  </option>
                );
              })}
            </select>
          </div>

          {/* Supplier */}
          <div>
            <label className="block text-[11px] text-gray-600 mb-1">
              Supplier
            </label>
            <select
              value={supplierFilter}
              onChange={(event) => {
                setSupplierFilter(event.target.value);
                resetPage();
              }}
              className="w-full border border-gray-300 rounded-md px-2 py-2 text-xs bg-white"
            >
              <option value="">All Suppliers</option>
              {suppliers.map((supplier) => (
                <option key={supplier} value={supplier}>
                  {supplier}
                </option>
              ))}
            </select>
          </div>

          {/* Match Check - Multiple Checkboxes */}
          <div className="md:col-span-2 lg:col-span-6">
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 mb-2">
              <label className="block text-[11px] font-semibold text-gray-700">
                MATCH CHECK — Select one or more checks
              </label>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => {
                    setSelectedMatchChecks([
                      "INVOICE",
                      "GSTIN",
                      "INVOICE_GSTIN",
                      "INVOICE_VALUE",
                      "CORE_MATCH",
                      "TAXABLE_VALUE",
                      "IGST",
                      "CGST",
                      "SGST",
                      "CESS",
                      "TOTAL_TAX",
                      "ALL_AVAILABLE",
                    ]);
                    resetPage();
                  }}
                  className="px-2.5 py-1 text-[10px] font-semibold border border-gray-300 rounded-md bg-white hover:bg-gray-50"
                >
                  Select All
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setSelectedMatchChecks([]);
                    resetPage();
                  }}
                  className="px-2.5 py-1 text-[10px] font-semibold border border-gray-300 rounded-md bg-white hover:bg-gray-50"
                >
                  Clear
                </button>
                <span className="text-[10px] text-gray-500">
                  {selectedMatchChecks.length} selected
                </span>
              </div>
            </div>

            <div className="border border-gray-300 rounded-lg bg-gray-50 p-3">
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2">
                {[
                  ["INVOICE", "Invoice Match"],
                  ["GSTIN", "GSTIN Match"],
                  ["INVOICE_GSTIN", "Invoice + GSTIN"],
                  ["INVOICE_VALUE", "Invoice Value ± ₹1"],
                  ["CORE_MATCH", "Core Match — GSTIN + Invoice + Amount ± ₹1"],
                  ["TAXABLE_VALUE", "Taxable Value Match"],
                  ["IGST", "IGST Match"],
                  ["CGST", "CGST Match"],
                  ["SGST", "SGST Match"],
                  ["CESS", "Cess Match"],
                  ["TOTAL_TAX", "Total Tax Match"],
                  ["ALL_AVAILABLE", "All Available Checks Match"],
                ].map(([value, label]) => {
                  const checked = selectedMatchChecks.includes(value);

                  return (
                    <label
                      key={value}
                      className={`flex items-start gap-2 p-2 rounded-md border cursor-pointer transition ${
                        checked
                          ? "border-green-400 bg-green-50"
                          : "border-gray-200 bg-white hover:bg-gray-50"
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => {
                          setSelectedMatchChecks((current) =>
                            current.includes(value)
                              ? current.filter((item) => item !== value)
                              : [...current, value]
                          );
                          resetPage();
                        }}
                        className="mt-0.5 h-4 w-4 accent-green-600"
                      />
                      <span className="text-[11px] text-gray-700 leading-4">
                        {label}
                      </span>
                    </label>
                  );
                })}
              </div>

              <div className="mt-2 text-[10px] text-gray-500">
                Select multiple checks to show only invoices where <b>all selected checks match</b>.
                Leave everything unchecked to show all matched invoices.
              </div>
            </div>
          </div>

          {/* Date From */}
          <div>
            <label className="block text-[11px] text-gray-600 mb-1">
              Date From
            </label>
            <input
              type="date"
              value={dateFrom}
              onChange={(event) => {
                setDateFrom(event.target.value);
                resetPage();
              }}
              className="w-full border border-gray-300 rounded-md px-2 py-2 text-xs bg-white"
            />
          </div>

          {/* Date To */}
          <div>
            <label className="block text-[11px] text-gray-600 mb-1">
              Date To
            </label>
            <input
              type="date"
              value={dateTo}
              onChange={(event) => {
                setDateTo(event.target.value);
                resetPage();
              }}
              className="w-full border border-gray-300 rounded-md px-2 py-2 text-xs bg-white"
            />
          </div>
        </div>

        {filtersApplied && (
          <div className="flex justify-end mt-3">
            <button
              type="button"
              onClick={clearFilters}
              className="px-3 py-1.5 text-xs font-semibold border border-gray-300 rounded-md bg-white hover:bg-gray-50"
            >
              Clear Filters
            </button>
          </div>
        )}
      </div>

      <div className="bg-white border rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b bg-green-50">
          <div className="font-bold text-green-800">
            Perfectly Matched Invoices
          </div>
          <div className="text-xs text-green-700 mt-1">
            Matching is based on GSTIN + Invoice Number + Invoice Amount within +/- ₹1.
            Date and all other Books/GSTR-2B values are shown side-by-side for complete verification.
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-xs min-w-[2600px]">
            <thead className="bg-gray-50 border-b">
              <tr>
                <th rowSpan="2" className="p-3 text-left whitespace-nowrap">S.NO.</th>
                <th rowSpan="2" className="p-3 text-left whitespace-nowrap">STATUS</th>
                <th rowSpan="2" className="p-3 text-left whitespace-nowrap">SUPPLIER NAME</th>
                <th rowSpan="2" className="p-3 text-left whitespace-nowrap">GSTIN</th>
                <th colSpan="10" className="p-3 text-center border-l">BOOKS</th>
                <th colSpan="10" className="p-3 text-center border-l">GSTR-2B</th>
                <th colSpan="3" className="p-3 text-center border-l">MATCH CHECK</th>
              </tr>
              <tr>
                {[
                  "INVOICE NO.", "INVOICE DATE", "INVOICE VALUE", "TAXABLE VALUE",
                  "IGST", "CGST", "SGST", "CESS", "TOTAL TAX", "ROW NO."
                ].map((h) => (
                  <th key={`b-${h}`} className="p-3 text-right whitespace-nowrap border-l">
                    {h}
                  </th>
                ))}
                {[
                  "INVOICE NO.", "INVOICE DATE", "INVOICE VALUE", "TAXABLE VALUE",
                  "IGST", "CGST", "SGST", "CESS", "TOTAL TAX", "ROW NO."
                ].map((h) => (
                  <th key={`t-${h}`} className="p-3 text-right whitespace-nowrap border-l">
                    {h}
                  </th>
                ))}
                <th className="p-3 text-center whitespace-nowrap border-l">INVOICE</th>
                <th className="p-3 text-center whitespace-nowrap">GSTIN</th>
                <th className="p-3 text-center whitespace-nowrap">CORE MATCH</th>
              </tr>
            </thead>

            <tbody>
              {visible.map((result, index) => {
                const book = result.books;
                const twoB = result.twoB;
                const supplier = book?.supplier || twoB?.supplier || "—";
                const gstin = book?.gstin || twoB?.gstin || "—";
                const serial =
                  (safePage - 1) * actualPageSize + index + 1;

                return (
                  <tr
                    key={result.id}
                    className="border-b hover:bg-green-50/40"
                  >
                    <td className="p-3 font-semibold text-gray-600">{serial}</td>
                    <td className="p-3">
                      <StatusBadge status="MATCHED" />
                    </td>
                    <td className="p-3 font-semibold whitespace-nowrap">{supplier}</td>
                    <td className="p-3 font-mono whitespace-nowrap">{gstin}</td>

                    <td className="p-3 whitespace-nowrap border-l">
                      <button
                        onClick={() => onSelect(result)}
                        className="text-blue-600 hover:underline font-semibold"
                      >
                        {book?.invoiceNo || "—"}
                      </button>
                    </td>
                    <td className="p-3 whitespace-nowrap border-l">{book?.invoiceDate || "—"}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(book?.grossTotal)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(book?.taxableValue)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(book?.igst)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(book?.cgst)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(book?.sgst)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(book?.cess)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(book?.totalTax)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{book?.rowNumber || "—"}</td>

                    <td className="p-3 whitespace-nowrap border-l">
                      <button
                        onClick={() => onSelect(result)}
                        className="text-purple-600 hover:underline font-semibold"
                      >
                        {twoB?.invoiceNo || "—"}
                      </button>
                    </td>
                    <td className="p-3 whitespace-nowrap border-l">{twoB?.invoiceDate || "—"}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(twoB?.invoiceValue)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(twoB?.taxableValue)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(twoB?.igst)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(twoB?.cgst)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(twoB?.sgst)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(twoB?.cess)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{amount(twoB?.totalTax)}</td>
                    <td className="p-3 text-right whitespace-nowrap border-l">{twoB?.rowNumber || "—"}</td>

                    <td className={`p-3 text-center border-l font-bold whitespace-nowrap ${result.invoiceMatch ? "text-green-600" : "text-red-500"}`}>
                      {result.invoiceMatch ? "✓ MATCH" : "✗ NO MATCH"}
                    </td>
                    <td className={`p-3 text-center font-bold whitespace-nowrap ${result.gstinMatch ? "text-green-600" : "text-red-500"}`}>
                      {result.gstinMatch ? "✓ MATCH" : "✗ NO MATCH"}
                    </td>
                    <td className={`p-3 text-center font-bold whitespace-nowrap ${result.invoiceMatch && result.gstinMatch && result.invoiceValueMatch ? "text-green-600" : "text-red-500"}`}>
                      {result.invoiceMatch && result.gstinMatch && result.invoiceValueMatch
                        ? "✓ GSTIN + INV + AMOUNT ± ₹1"
                        : "✗ CORE MISMATCH"}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {!visible.length && (
            <div className="p-12 text-center text-gray-400">
              No matched records found.
            </div>
          )}
        </div>

        <div className="p-4 border-t">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div className="text-xs text-gray-500">
              Showing {filtered.length ? ((safePage - 1) * actualPageSize + 1).toLocaleString("en-IN") : 0}
              {" – "}
              {Math.min(safePage * actualPageSize, filtered.length).toLocaleString("en-IN")}
              {" of "}
              {filtered.length.toLocaleString("en-IN")}
            </div>

            <Pagination
              page={safePage}
              setPage={setPage}
              total={filtered.length}
              pageSize={pageSize}
              setPageSize={(value) => {
                setPageSize(value);
                setPage(1);
              }}
            />
          </div>
        </div>
      </div>
    </div>
  );
}

/* =======================================================
   DETAIL ROW
======================================================= */

function DetailRow({
  label,
  value,
}) {
  return (
    <div className="flex justify-between gap-5 py-2 border-b">

      <span className="text-xs text-gray-500">
        {label}
      </span>

      <span className="text-sm font-semibold text-right break-all">
        {value === "" ||
        value === null ||
        value === undefined
          ? "—"
          : value}
      </span>

    </div>
  );
}


/* =======================================================
   SOURCE DETAIL CARD
======================================================= */

function SourceCard({
  title,
  record,
}) {
  if (!record) {
    return (
      <div className="border rounded-xl p-5">

        <h3 className="font-bold">
          {title}
        </h3>

        <div className="text-sm text-red-500 mt-4">
          Record not available.
        </div>

      </div>
    );
  }

  return (
    <div className="border rounded-xl p-5">

      <h3 className="font-bold mb-4">
        {title}
      </h3>

      <DetailRow
        label="Source"
        value={
          record.source
        }
      />

      <DetailRow
        label="Section"
        value={
          record.section ||
          "Purchase Register"
        }
      />

      <DetailRow
        label="GSTIN"
        value={
          record.gstin
        }
      />

      <DetailRow
        label="Supplier Name"
        value={
          record.supplier
        }
      />

      <DetailRow
        label="Invoice Number"
        value={
          record.invoiceNo
        }
      />

      {record.originalInvoice && (
        <DetailRow
          label="Original Invoice"
          value={
            record.originalInvoice
          }
        />
      )}

      <DetailRow
        label="Invoice Date"
        value={
          record.invoiceDate
        }
      />

      <DetailRow
        label="Invoice Value"
        value={
          amount(
            record.invoiceValue
          )
        }
      />

      <DetailRow
        label="Taxable Value"
        value={
          amount(
            record.taxableValue
          )
        }
      />

      <DetailRow
        label="IGST"
        value={
          amount(
            record.igst
          )
        }
      />

      <DetailRow
        label="CGST"
        value={
          amount(
            record.cgst
          )
        }
      />

      <DetailRow
        label="SGST"
        value={
          amount(
            record.sgst
          )
        }
      />

      <DetailRow
        label="Cess"
        value={
          amount(
            record.cess
          )
        }
      />

      <DetailRow
        label="Total Tax"
        value={
          amount(
            record.totalTax
          )
        }
      />

      <DetailRow
        label="Reverse Charge"
        value={
          record.reverseCharge
        }
      />

      <DetailRow
        label="ITC Availability"
        value={
          record.itcAvailability
        }
      />

      <DetailRow
        label="Reason"
        value={
          record.reason
        }
      />

      <DetailRow
        label="Row Number"
        value={
          record.rowNumber
        }
      />

      <details className="mt-5">

        <summary className="cursor-pointer text-sm font-semibold text-blue-600">
          View all source fields
        </summary>

        <div className="mt-4 border rounded-lg overflow-x-auto">

          <table className="w-full text-xs">

            <tbody>

              {Object.entries(
                record.raw || {}
              ).map(
                ([key, value]) => (

                  <tr
                    key={key}
                    className="border-b"
                  >

                    <td className="p-2 bg-gray-50 font-semibold whitespace-nowrap">
                      {key}
                    </td>

                    <td className="p-2 whitespace-nowrap">
                      {text(value) ||
                        "—"}
                    </td>

                  </tr>

                )
              )}

            </tbody>

          </table>

        </div>

      </details>

    </div>
  );
}


/* =======================================================
   DETAILS MODAL
======================================================= */

function DetailsModal({
  result,
  close,
}) {
  return (
    <div
      className="fixed inset-0 bg-black/40 z-[100] flex items-center justify-center p-6"
      onClick={close}
    >

      <div
        className="bg-white rounded-2xl shadow-xl w-full max-w-7xl max-h-[92vh] overflow-y-auto"
        onClick={(event) =>
          event.stopPropagation()
        }
      >

        <div className="sticky top-0 z-10 bg-white border-b px-6 py-5 flex items-center justify-between">

          <div>

            <h2 className="text-xl font-bold">
              Reconciliation Details
            </h2>

            <div className="mt-3 flex flex-wrap items-center gap-3">
              <StatusBadge
                status={
                  result.status
                }
              />
              <div className="text-sm text-gray-600">
                <span className="font-semibold text-gray-800">Supplier:</span>{" "}
                {result.books?.supplier || result.twoB?.supplier || "—"}
              </div>
              <div className="text-sm text-gray-600 font-mono">
                <span className="font-sans font-semibold text-gray-800">GSTIN:</span>{" "}
                {result.books?.gstin || result.twoB?.gstin || "—"}
              </div>
            </div>

          </div>

          <button
            onClick={close}
            className="text-3xl text-gray-400 hover:text-gray-700"
          >
            ×
          </button>

        </div>

        <div className="p-6">

          <div className="grid lg:grid-cols-2 gap-6">

            <SourceCard
              title="BOOKS — PURCHASE REGISTER"
              record={
                result.books
              }
            />

            <SourceCard
              title="GSTR-2B"
              record={
                result.twoB
              }
            />

          </div>

          {result.books &&
            result.twoB && (
              <div className="mt-6 border rounded-xl p-5">

                <h3 className="font-bold mb-4">
                  Field Comparison
                </h3>

                <div className="grid md:grid-cols-4 gap-4">

                  <Comparison
                    title="GSTIN"
                    value={
                      result.gstinMatch
                    }
                  />

                  <Comparison
                    title="SUPPLIER NAME"
                    value={
                      Boolean(
                        result.books?.supplier &&
                        result.twoB?.supplier &&
                        normalizedText(result.books.supplier) === normalizedText(result.twoB.supplier)
                      )
                    }
                  />

                  <Comparison
                    title="INVOICE"
                    value={
                      result.invoiceMatch
                    }
                  />

                  <Comparison
                    title="TAXABLE VALUE"
                    value={
                      result.taxableMatch
                    }
                  />

                  <Comparison
                    title="TOTAL TAX"
                    value={
                      result.taxMatch
                    }
                  />

                  <Comparison
                    title="IGST"
                    value={
                      result.igstMatch
                    }
                  />

                  <Comparison
                    title="CGST"
                    value={
                      result.cgstMatch
                    }
                  />

                  <Comparison
                    title="SGST"
                    value={
                      result.sgstMatch
                    }
                  />

                </div>

                <div className="mt-5 bg-gray-50 rounded-lg p-4">

                  <div className="text-xs text-gray-500">
                    Matching Method
                  </div>

                  <div className="text-sm font-semibold mt-1">
                    {
                      result.matchingMethod ||
                      "—"
                    }
                  </div>

                </div>

              </div>
            )}

          <div className="mt-6 border rounded-xl p-5 bg-gray-50">

            <h3 className="font-bold">
              Reconciliation Note
            </h3>

            <p className="text-sm text-gray-600 mt-2">
              {
                result.reason ||
                "No additional note."
              }
            </p>

          </div>

        </div>

      </div>

    </div>
  );
}


/* =======================================================
   COMPARISON CARD
======================================================= */

function Comparison({
  title,
  value,
}) {
  let label =
    "NOT AVAILABLE";

  if (value === true) {
    label = "MATCH";
  }

  if (value === false) {
    label = "DIFFERENT";
  }

  return (
    <div className="border rounded-lg p-4">

      <div className="text-xs text-gray-500">
        {title}
      </div>

      <div
        className={`font-bold mt-1 ${
          value === true
            ? "text-green-600"
            : value === false
            ? "text-red-600"
            : "text-gray-400"
        }`}
      >
        {label}
      </div>

    </div>
  );
}


/* =======================================================
   EXPORT CSV
======================================================= */

function exportCSV(results, fileName = "GST_RECO_Reconciliation.xlsx") {
  const rows =
    results.map(
      (result, index) => ({
        "S.No.":
          index + 1,

        Status:
          result.status,

        GSTIN:
          result.books?.gstin ||
          result.twoB?.gstin ||
          "",

        "Supplier Name":
          result.books?.supplier ||
          result.twoB?.supplier ||
          "",

        "Books Invoice":
          result.books
            ?.invoiceNo ||
          "",

        "2B Invoice":
          result.twoB
            ?.invoiceNo ||
          "",

        "Books Supplier Name":
          result.books?.supplier || "",

        "2B Supplier Name":
          result.twoB?.supplier || "",

        "Books Invoice Date":
          result.books
            ?.invoiceDate ||
          "",

        "2B Invoice Date":
          result.twoB
            ?.invoiceDate ||
          "",

        "Books Taxable Value":
          result.books
            ?.taxableValue ??
          "",

        "Books Invoice Value":
          result.books?.grossTotal ?? "",

        "2B Invoice Value":
          result.twoB?.invoiceValue ?? "",

        "2B Taxable Value":
          result.twoB
            ?.taxableValue ??
          "",

        "Books IGST": result.books?.igst ?? "",
        "2B IGST": result.twoB?.igst ?? "",
        "Books CGST": result.books?.cgst ?? "",
        "2B CGST": result.twoB?.cgst ?? "",
        "Books SGST": result.books?.sgst ?? "",
        "2B SGST": result.twoB?.sgst ?? "",
        "Books Cess": result.books?.cess ?? "",
        "2B Cess": result.twoB?.cess ?? "",
        "Books Row Number": result.books?.rowNumber ?? "",
        "2B Row Number": result.twoB?.rowNumber ?? "",

        "Taxable Difference":
          result.taxableDifference ??
          "",

        "Books Total Tax":
          result.books
            ?.totalTax ??
          "",

        "2B Total Tax":
          result.twoB
            ?.totalTax ??
          "",

        "Tax Difference":
          result.taxDifference ??
          "",

        Reason:
          result.reason ||
          "",
      })
    );

  const worksheet =
    XLSX.utils.json_to_sheet(
      rows
    );

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "GST Reconciliation"
  );

  XLSX.writeFile(
    workbook,
    fileName
  );
}


/* =======================================================
   MAIN APP
======================================================= */

export default function App() {

  const [active, setActive] =
    useState("start");

  const [booksFile, setBooksFile] =
    useState(null);

  const [twoBFile, setTwoBFile] =
    useState(null);

  const [books, setBooks] =
    useState([]);

  const [twoB, setTwoB] =
    useState([]);

  const [creditNotes, setCreditNotes] =
    useState([]);

  const [imports, setImports] =
    useState([]);

  const [results, setResults] =
    useState([]);

  const [selected, setSelected] =
    useState(null);

  const [error, setError] =
    useState("");

  const [processing, setProcessing] =
    useState(false);

  const [booksSearch, setBooksSearch] =
    useState("");

  const [twoBSearch, setTwoBSearch] =
    useState("");

  const [reconSearch, setReconSearch] =
    useState("");

  const [summaryDateFrom, setSummaryDateFrom] =
    useState("");

  const [summaryDateTo, setSummaryDateTo] =
    useState("");


  /* =====================================================
     BOOKS UPLOAD
  ===================================================== */

  async function handleBooks(file) {

    try {

      setError("");

      if (!file) {
        return;
      }

      const buffer =
        await file.arrayBuffer();

      const workbook =
        XLSX.read(
          buffer,
          {
            type: "array",
            cellDates: true,
          }
        );

      const parsed =
        parseBooks(
          workbook
        );

      if (
        !parsed.records.length
      ) {
        throw new Error(
          "No transaction records were found in the Purchase Register."
        );
      }

      setBooks(
        parsed.records
      );

      setBooksFile({
        name:
          file.name,

        entries:
          parsed.records.length,
      });

      setResults([]);

    } catch (err) {

      console.error(
        "Books parsing error:",
        err
      );

      setBooks([]);

      setBooksFile(null);

      setError(
        `Books file error: ${
          err.message ||
          "Unable to read the Purchase Register."
        }`
      );
    }
  }


  /* =====================================================
     2B UPLOAD
  ===================================================== */

  async function handle2B(file) {

    try {

      setError("");

      if (!file) {
        return;
      }

      const buffer =
        await file.arrayBuffer();

      const workbook =
        XLSX.read(
          buffer,
          {
            type: "array",
            cellDates: true,
          }
        );

      const parsed =
        parseTwoB(
          workbook
        );

      if (
        !parsed.invoiceRecords.length &&
        !parsed.creditNotes.length &&
        !parsed.imports.length
      ) {
        throw new Error(
          "No usable GSTR-2B records were found."
        );
      }

      setTwoB(
        parsed.invoiceRecords
      );

      setCreditNotes(
        parsed.creditNotes
      );

      setImports(
        parsed.imports
      );

      setTwoBFile({
        name:
          file.name,

        entries:
          parsed.invoiceRecords.length,
      });

      setResults([]);

    } catch (err) {

      console.error(
        "2B parsing error:",
        err
      );

      setTwoB([]);

      setCreditNotes([]);

      setImports([]);

      setTwoBFile(null);

      setError(
        `GSTR-2B file error: ${
          err.message ||
          "Unable to read the GSTR-2B workbook."
        }`
      );
    }
  }


  /* =====================================================
     RUN RECONCILIATION
  ===================================================== */

  function runReconciliation() {

    if (
      !books.length
    ) {

      setError(
        "Please upload the Books Purchase Register."
      );

      return;
    }

    if (
      !twoB.length
    ) {

      setError(
        "No B2B/B2BA invoice records were found in GSTR-2B."
      );

      return;
    }

    try {

      setError("");

      setProcessing(true);

      /*
        Small timeout so the UI visibly enters
        processing state.
      */

      setTimeout(
        () => {

          try {

            const output =
              reconcileBooks(
                books,
                twoB
              );

            setResults(
              output
            );

            setActive(
              "recon"
            );

          } catch (err) {

            console.error(
              "Reconciliation error:",
              err
            );

            setError(
              `Reconciliation error: ${
                err.message
              }`
            );

          } finally {

            setProcessing(false);

          }

        },
        200
      );

    } catch (err) {

      setProcessing(false);

      setError(
        err.message
      );
    }
  }


  /* =====================================================
     RESET
  ===================================================== */

  function resetAll() {

    setBooksFile(null);

    setTwoBFile(null);

    setBooks([]);

    setTwoB([]);

    setCreditNotes([]);

    setImports([]);

    setResults([]);

    setSelected(null);

    setError("");

    setBooksSearch("");

    setTwoBSearch("");

    setReconSearch("");

    setSummaryDateFrom("");

    setSummaryDateTo("");

    setActive(
      "start"
    );
  }


  /* =====================================================
     DUPLICATES
  ===================================================== */

  const booksDuplicates =
    useMemo(
      () =>
        getDuplicates(
          books
        ),
      [books]
    );

  const twoBDuplicates =
    useMemo(
      () =>
        getDuplicates(
          twoB
        ),
      [twoB]
    );


  /* =====================================================
     RCM
  ===================================================== */

  const booksRCM =
    useMemo(
      () =>
        books.filter(
          (record) =>
            `${record.voucherType} ${record.particulars} ${record.supplier}`
              .toUpperCase()
              .includes("RCM")
        ),
      [books]
    );

  const twoBRCM =
    useMemo(
      () =>
        twoB.filter(
          (record) =>
            normalizedText(
              record.reverseCharge
            ) ===
            "YES"
        ),
      [twoB]
    );


  /* =====================================================
     SUMMARY
  ===================================================== */

  const summary =
    useMemo(
      () => {

        return {
          total:
            results.length,

          matched:
            results.filter(
              (r) =>
                r.status ===
                "MATCHED"
            ).length,

          partial:
            results.filter(
              (r) =>
                r.status ===
                "PARTIAL"
            ).length,

          review:
            results.filter(
              (r) =>
                r.status ===
                "REVIEW"
            ).length,

          notInBooks:
            results.filter(
              (r) =>
                r.status ===
                "NOT IN BOOKS"
            ).length,

          notIn2B:
            results.filter(
              (r) =>
                r.status ===
                "NOT IN 2B"
            ).length,
        };

      },
      [results]
    );


  /* =====================================================
     START
  ===================================================== */

  function renderStart() {

    return (
      <div className="max-w-6xl mx-auto">

        <div className="text-center mb-8">

          <h1 className="text-3xl font-bold text-gray-900">
            GST Reconciliation
          </h1>

          <p className="text-sm text-gray-500 mt-2">
            Upload the company Purchase Register and official GSTR-2B workbook.
          </p>

        </div>

        <div className="grid md:grid-cols-2 gap-6">

          <UploadCard
            title="BOOKS"
            subtitle="Company Purchase Register"
            file={
              booksFile
            }
            onFile={
              handleBooks
            }
            color="blue"
          />

          <UploadCard
            title="GSTR-2B"
            subtitle="Official GST Portal Workbook"
            file={
              twoBFile
            }
            onFile={
              handle2B
            }
            color="purple"
          />

        </div>

        {error && (
          <div className="mt-6 p-4 bg-red-50 border border-red-200 rounded-xl text-red-700 text-sm">
            {error}
          </div>
        )}

        <div className="text-center mt-8">

          <button
            onClick={
              runReconciliation
            }
            disabled={
              !books.length ||
              !twoB.length ||
              processing
            }
            className={`px-8 py-3 rounded-lg text-white font-semibold ${
              books.length &&
              twoB.length &&
              !processing
                ? "bg-green-600 hover:bg-green-700"
                : "bg-gray-300 cursor-not-allowed"
            }`}
          >
            {processing
              ? "Processing..."
              : "↔ Reconcile Now"}
          </button>

        </div>

        {books.length > 0 &&
          twoB.length > 0 && (
            <div className="mt-5 grid md:grid-cols-3 gap-4">

              <Card
                title="BOOKS"
                value={
                  books.length.toLocaleString(
                    "en-IN"
                  )
                }
                subtitle="Purchase Register records"
              />

              <Card
                title="GSTR-2B"
                value={
                  twoB.length.toLocaleString(
                    "en-IN"
                  )
                }
                subtitle="B2B + B2BA invoice records"
              />

              <Card
                title="ADDITIONAL 2B"
                value={
                  (
                    creditNotes.length +
                    imports.length
                  ).toLocaleString(
                    "en-IN"
                  )
                }
                subtitle="Credit/debit notes + imports"
              />

            </div>
          )}

      </div>
    );
  }


  /* =====================================================
     BOOKS INPUT
  ===================================================== */

  function renderBooks() {

    return (
      <div>

        <h1 className="text-2xl font-bold">
          Books Input
        </h1>

        <p className="text-sm text-gray-500 mt-1">
          Actual records loaded from the company's Purchase Register.
        </p>

        <div className="mt-6">

          <DataTable
            records={
              books
            }
            onSelect={(record) =>
              setSelected({
                status:
                  "BOOKS",
                books:
                  record,
                twoB:
                  null,
              })
            }
            search={
              booksSearch
            }
            setSearch={
              setBooksSearch
            }
          />

        </div>

      </div>
    );
  }


  /* =====================================================
     2B INPUT
  ===================================================== */

  function render2B() {

    return (
      <div>

        <h1 className="text-2xl font-bold">
          GSTR-2B Input
        </h1>

        <p className="text-sm text-gray-500 mt-1">
          B2B and B2BA invoice records from the official workbook.
        </p>

        <div className="mt-6">

          <DataTable
            records={
              twoB
            }
            onSelect={(record) =>
              setSelected({
                status:
                  "GSTR-2B",
                books:
                  null,
                twoB:
                  record,
              })
            }
            search={
              twoBSearch
            }
            setSearch={
              setTwoBSearch
            }
          />

        </div>

      </div>
    );
  }


  /* =====================================================
     RECON
  ===================================================== */

  function renderRecon(
    records =
      results
  ) {

    return (
      <div>

        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">

          <div>

            <h1 className="text-2xl font-bold">
              Reconciliation
            </h1>

            <p className="text-sm text-gray-500 mt-1">
              Actual Books vs GSTR-2B comparison.
            </p>

          </div>

          {records.length > 0 && (
            <button
              onClick={() => {
                const statusNames = [...new Set(records.map((r) => r.status))];
                const fileName =
                  statusNames.length === 1
                    ? `GST_RECO_${statusNames[0].replace(/\s+/g, "_")}.xlsx`
                    : "GST_RECO_Reconciliation.xlsx";
                exportCSV(records, fileName);
              }}
              className="px-4 py-2 bg-gray-900 text-white rounded-lg text-sm font-semibold"
            >
              Export Excel
            </button>
          )}

        </div>

        <div className="grid grid-cols-2 md:grid-cols-6 gap-4 mt-6">

          <Card
            title="TOTAL"
            value={
              records.length
            }
          />

          <Card
            title="MATCHED"
            value={
              records.filter(
                (r) =>
                  r.status ===
                  "MATCHED"
              ).length
            }
            className="bg-green-50 border-green-200"
          />

          <Card
            title="PARTIAL"
            value={
              records.filter(
                (r) =>
                  r.status ===
                  "PARTIAL"
              ).length
            }
            className="bg-yellow-50 border-yellow-200"
          />

          <Card
            title="REVIEW"
            value={
              records.filter(
                (r) =>
                  r.status ===
                  "REVIEW"
              ).length
            }
            className="bg-orange-50 border-orange-200"
          />

          <Card
            title="NOT IN BOOKS"
            value={
              records.filter(
                (r) =>
                  r.status ===
                  "NOT IN BOOKS"
              ).length
            }
            className="bg-red-50 border-red-200"
          />

          <Card
            title="NOT IN 2B"
            value={
              records.filter(
                (r) =>
                  r.status ===
                  "NOT IN 2B"
              ).length
            }
            className="bg-red-50 border-red-200"
          />

        </div>

        <div className="mt-6">

          <ReconTable
            records={
              records
            }
            onSelect={
              setSelected
            }
            search={
              reconSearch
            }
            setSearch={
              setReconSearch
            }
          />

        </div>

      </div>
    );
  }


  /* =====================================================
     SUMMARY
  ===================================================== */

  function renderSummary() {
    const inDateRange = (value) => {
      const d = dateValue(value);
      if (summaryDateFrom && (!d || d < summaryDateFrom)) return false;
      if (summaryDateTo && (!d || d > summaryDateTo)) return false;
      return true;
    };

    const filteredBooks = books.filter((r) =>
      inDateRange(r.invoiceDate || r.date)
    );

    const filteredTwoB = twoB.filter((r) =>
      inDateRange(r.invoiceDate)
    );

    const filteredResults = results.filter((r) =>
      inDateRange(r.books?.invoiceDate || r.twoB?.invoiceDate)
    );

    const filteredBooksDuplicates = booksDuplicates.filter((r) =>
      inDateRange(r.invoiceDate || r.date)
    );

    const filteredTwoBDuplicates = twoBDuplicates.filter((r) =>
      inDateRange(r.invoiceDate || r.date)
    );

    const filteredCreditNotes = creditNotes.filter((r) =>
      inDateRange(r.noteDate || r.invoiceDate || r.date)
    );

    const filteredImports = imports.filter((r) =>
      inDateRange(r.invoiceDate || r.billOfEntryDate || r.date)
    );

    const summaryFiltered = {
      matched: filteredResults.filter((r) => r.status === "MATCHED").length,
      partial: filteredResults.filter((r) => r.status === "PARTIAL").length,
      review: filteredResults.filter((r) => r.status === "REVIEW").length,
      notInBooks: filteredResults.filter((r) => r.status === "NOT IN BOOKS").length,
      notIn2B: filteredResults.filter((r) => r.status === "NOT IN 2B").length,
    };

    const filtersApplied = summaryDateFrom || summaryDateTo;

    function clearSummaryDates() {
      setSummaryDateFrom("");
      setSummaryDateTo("");
    }

    return (
      <div>
        <h1 className="text-2xl font-bold">Summary</h1>

        <p className="text-sm text-gray-500 mt-1">
          Overall reconciliation statistics. Use the date range to view statistics for a specific period.
        </p>

        {/* DATE FILTER — available on Summary also */}
        <div className="bg-white border border-black rounded-xl p-3 mt-5 mb-6">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div>
              <label className="block text-[11px] text-gray-600 mb-1">
                Date From
              </label>
              <input
                type="date"
                value={summaryDateFrom}
                onChange={(event) => setSummaryDateFrom(event.target.value)}
                className="w-full border border-gray-300 rounded-md px-2 py-2 text-xs bg-white"
              />
            </div>

            <div>
              <label className="block text-[11px] text-gray-600 mb-1">
                Date To
              </label>
              <input
                type="date"
                value={summaryDateTo}
                onChange={(event) => setSummaryDateTo(event.target.value)}
                className="w-full border border-gray-300 rounded-md px-2 py-2 text-xs bg-white"
              />
            </div>

            <div className="flex items-end">
              {filtersApplied ? (
                <button
                  type="button"
                  onClick={clearSummaryDates}
                  className="w-full px-3 py-2 text-xs font-semibold border border-gray-300 rounded-md bg-white hover:bg-gray-50"
                >
                  Clear Filters
                </button>
              ) : (
                <div className="text-xs text-gray-400 pb-2">
                  Showing all dates
                </div>
              )}
            </div>
          </div>
        </div>

        <div className="grid md:grid-cols-3 gap-4">
          <Card title="BOOKS RECORDS" value={filteredBooks.length} subtitle="Purchase Register" />
          <Card title="2B INVOICE RECORDS" value={filteredTwoB.length} subtitle="B2B + B2BA" />
          <Card title="RECONCILIATION RESULTS" value={filteredResults.length} />

          <Card title="MATCHED" value={summaryFiltered.matched} className="bg-green-50 border-green-200" />
          <Card title="PARTIAL" value={summaryFiltered.partial} className="bg-yellow-50 border-yellow-200" />
          <Card title="REVIEW" value={summaryFiltered.review} className="bg-orange-50 border-orange-200" />
          <Card title="NOT IN BOOKS" value={summaryFiltered.notInBooks} className="bg-red-50 border-red-200" />
          <Card title="NOT IN 2B" value={summaryFiltered.notIn2B} className="bg-red-50 border-red-200" />

          <Card title="BOOKS DUPLICATES" value={filteredBooksDuplicates.length} className="bg-orange-50 border-orange-200" />
          <Card title="2B DUPLICATES" value={filteredTwoBDuplicates.length} className="bg-orange-50 border-orange-200" />
          <Card title="CREDIT / DEBIT NOTES" value={filteredCreditNotes.length} />
          <Card title="IMPORT RECORDS" value={filteredImports.length} />
        </div>
      </div>
    );
  }

  /* =====================================================
     MATCHED
  ===================================================== */

  function renderMatched() {

    const matchedRecords = results.filter(
      (r) => r.status === "MATCHED"
    );

    return (
      <div>
        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 mb-6">
          <div>
            <h1 className="text-2xl font-bold">Matched</h1>
            <p className="text-sm text-gray-500 mt-1">
              Complete Books and GSTR-2B values for invoices matched by GSTIN + invoice number + invoice amount (±₹1).
            </p>
          </div>
          {matchedRecords.length > 0 && (
            <button
              onClick={() => exportCSV(matchedRecords, "GST_RECO_MATCHED.xlsx")}
              className="px-4 py-2 bg-gray-900 text-white rounded-lg text-sm font-semibold"
            >
              Export Matched Excel
            </button>
          )}
        </div>

        <div className="grid grid-cols-2 md:grid-cols-5 gap-4 mb-6">
          <Card title="MATCHED" value={matchedRecords.length} className="bg-green-50 border-green-200" />
          <Card title="GSTIN MATCHED" value={matchedRecords.length} className="bg-green-50 border-green-200" />
          <Card title="INVOICE MATCHED" value={matchedRecords.length} className="bg-green-50 border-green-200" />
          <Card title="VALUE MATCHED" value={matchedRecords.length} className="bg-green-50 border-green-200" />
          <Card title="AMOUNT ± ₹1" value={matchedRecords.length} className="bg-green-50 border-green-200" />
        </div>

        <MatchedDetailsTable
          records={matchedRecords}
          onSelect={setSelected}
        />
      </div>
    );
  }


  /* =====================================================
     PARTIAL
  ===================================================== */

  function renderPartial() {

    return renderRecon(
      results.filter(
        (r) =>
          r.status ===
          "PARTIAL"
      )
    );
  }


  /* =====================================================
     DUPLICATES
  ===================================================== */

  function renderDuplicates() {

    const duplicates = [
      ...booksDuplicates,
      ...twoBDuplicates,
    ];

    return (
      <div>

        <h1 className="text-2xl font-bold">
          Duplicate Records
        </h1>

        <p className="text-sm text-gray-500 mt-1">
          Duplicate records are checked separately within Books and GSTR-2B.
        </p>

        <div className="grid md:grid-cols-2 gap-4 mt-6">

          <Card
            title="BOOKS DUPLICATE ROWS"
            value={
              booksDuplicates.length
            }
            subtitle="Inside Purchase Register only"
          />

          <Card
            title="2B DUPLICATE ROWS"
            value={
              twoBDuplicates.length
            }
            subtitle="Inside GSTR-2B only"
          />

        </div>

        <div className="mt-6">

          <DataTable
            records={
              duplicates
            }
            onSelect={(record) =>
              setSelected({
                status:
                  "DUPLICATE",
                books:
                  record.source ===
                  "BOOKS"
                    ? record
                    : null,
                twoB:
                  record.source ===
                  "GSTR-2B"
                    ? record
                    : null,
              })
            }
            search=""
            setSearch={() => {}}
          />

        </div>

      </div>
    );
  }


  /* =====================================================
     NOT IN BOOKS
  ===================================================== */

  function renderNotInBooks() {

    return renderRecon(
      results.filter(
        (r) =>
          r.status ===
          "NOT IN BOOKS"
      )
    );
  }


  /* =====================================================
     NOT IN 2B
  ===================================================== */

  function renderNotIn2B() {

    return renderRecon(
      results.filter(
        (r) =>
          r.status ===
          "NOT IN 2B"
      )
    );
  }


  /* =====================================================
     RCM
  ===================================================== */

  function renderRCM() {

    return (
      <div>

        <h1 className="text-2xl font-bold">
          RCM
        </h1>

        <p className="text-sm text-gray-500 mt-1">
          Reverse Charge records identified from the uploaded data.
        </p>

        <div className="grid md:grid-cols-2 gap-4 mt-6">

          <Card
            title="BOOKS RCM"
            value={
              booksRCM.length
            }
          />

          <Card
            title="GSTR-2B RCM"
            value={
              twoBRCM.length
            }
          />

        </div>

        <div className="mt-6">

          <h2 className="font-bold mb-3">
            Books RCM
          </h2>

          <DataTable
            records={
              booksRCM
            }
            onSelect={(record) =>
              setSelected({
                status:
                  "BOOKS",
                books:
                  record,
                twoB:
                  null,
              })
            }
            search=""
            setSearch={() => {}}
          />

        </div>

        <div className="mt-8">

          <h2 className="font-bold mb-3">
            GSTR-2B RCM
          </h2>

          <DataTable
            records={
              twoBRCM
            }
            onSelect={(record) =>
              setSelected({
                status:
                  "GSTR-2B",
                books:
                  null,
                twoB:
                  record,
              })
            }
            search=""
            setSearch={() => {}}
          />

        </div>

      </div>
    );
  }


  /* =====================================================
     IMPORTS
  ===================================================== */

  function renderImports() {

    return (
      <div>

        <h1 className="text-2xl font-bold">
          Imports
        </h1>

        <p className="text-sm text-gray-500 mt-1">
          Import records from GSTR-2B.
        </p>

        <div className="mt-6">

          <DataTable
            records={
              imports
            }
            onSelect={(record) =>
              setSelected({
                status:
                  "GSTR-2B",
                books:
                  null,
                twoB:
                  record,
              })
            }
            search=""
            setSearch={() => {}}
          />

        </div>

      </div>
    );
  }


  /* =====================================================
     CREDIT NOTES
  ===================================================== */

  function renderCreditNotes() {

    return (
      <div>

        <h1 className="text-2xl font-bold">
          DR / CREDIT NOTES
        </h1>

        <p className="text-sm text-gray-500 mt-1">
          Debit and credit notes reported in GSTR-2B.
        </p>

        <div className="mt-6">

          <DataTable
            records={
              creditNotes
            }
            onSelect={(record) =>
              setSelected({
                status:
                  "GSTR-2B",
                books:
                  null,
                twoB:
                  record,
              })
            }
            search=""
            setSearch={() => {}}
          />

        </div>

      </div>
    );
  }


  /* =====================================================
     PAGE ROUTER
  ===================================================== */

  function page() {

    switch (
      active
    ) {

      case "start":
        return renderStart();

      case "books":
        return renderBooks();

      case "2b":
        return render2B();

      case "recon":
        return renderRecon();

      case "summary":
        return renderSummary();

      case "matched":
        return renderMatched();

      case "partial":
        return renderPartial();

      case "duplicate":
        return renderDuplicates();

      case "not-books":
        return renderNotInBooks();

      case "not-2b":
        return renderNotIn2B();

      case "rcm":
        return renderRCM();

      case "imports":
        return renderImports();

      case "credit-notes":
        return renderCreditNotes();

      default:
        return renderStart();
    }
  }


  /* =====================================================
     FINAL UI
  ===================================================== */

  return (
    <div className="min-h-screen bg-slate-50">

      {/* FIXED SIDEBAR */}

      <Sidebar
        active={
          active
        }
        setActive={
          setActive
        }
      />

      {/* MAIN AREA */}

      <main className="ml-64 min-h-screen">

        {/* TOP BAR */}

        <header className="h-16 bg-white border-b border-gray-200 flex items-center justify-between px-8 sticky top-0 z-30">

          <div>

            <div className="font-bold text-green-700">
              GST RECO
            </div>

            <div className="text-xs text-gray-400">
              Smart GST Reconciliation
            </div>

          </div>

          <button
            onClick={
              resetAll
            }
            className="px-4 py-2 border border-gray-300 rounded-lg bg-white hover:bg-gray-50 text-sm"
          >
            ↩ Upload New Files
          </button>

        </header>

        {/* CONTENT */}

        <div className="p-8">

          {error && (
            <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-xl text-red-700 text-sm">
              {error}
            </div>
          )}

          {page()}

        </div>

      </main>

      {/* DETAILS */}

      {selected && (
        <DetailsModal
          result={
            selected
          }
          close={() =>
            setSelected(
              null
            )
          }
        />
      )}

    </div>
  );
}