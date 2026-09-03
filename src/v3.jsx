import React, { useState, useMemo, useCallback } from "react";
import {
  Plus, Trash2, ChevronDown, ChevronRight, Copy, Download,
  FileJson, Users, Ticket, ListTree, Workflow, Check, X,
  GripVertical, AlertTriangle, RefreshCw, FileSpreadsheet,
  Search, ArrowLeft
} from "lucide-react";
import * as XLSX from "xlsx";

/* ------------------------------------------------------------------ */
/* utils                                                               */
/* ------------------------------------------------------------------ */

let uidCounter = 0;
const uid = () => `id_${Date.now().toString(36)}_${(uidCounter++).toString(36)}`;

const FIELD_TYPES = [
  { value: "dropdown", label: "Dropdown" },
  { value: "casecadeDropdown", label: "Cascading Dropdown" },
  { value: "date", label: "Date" },
  { value: "time", label: "Time" },
  { value: "text", label: "Text" },
];

const NOTIF_TYPES = ["inApp", "email"];

const newDropdownField = () => ({
  id: uid(), label: "", key: "", type: "dropdown", options: [],
});
const newCascadeField = () => ({
  id: uid(), label: "", key: "", type: "casecadeDropdown",
  childLabel: "", childKey: "", parentOptionType: "multiplechoice", options: [],
});
const newSimpleField = (type) => ({ id: uid(), label: "", key: "", type });

const newField = (type) => {
  if (type === "dropdown") return newDropdownField();
  if (type === "casecadeDropdown") return newCascadeField();
  return newSimpleField(type);
};

const newParentOption = () => ({ id: uid(), name: "", key: "", children: [] });
const newChildOption = () => ({ id: uid(), value: "", key: "" });
const newSimpleOption = () => ({ id: uid(), value: "", key: "" });

const newAutoEscalation = () => ({
  id: uid(), assignee: "", days: "", overrideStatus: "",
  notification: true, notificationType: [], ccAddress: [],
});

const newStatus = () => ({
  id: uid(), status: "", label: "", public: false,
  roles: [], assigneeRoles: [], notification: false, notificationType: [],
  notificationRole: [], mandatoryCustomFields: [], overrideStatus: "",
  customFieldsMetaData: [], autoEscalationConfig: [],
});

const newSubType = () => ({
  id: uid(), ticketKey: "", prefix: "", ticketType: "",
  deleted: false, isAutoEscalation: false,
  customFieldsMetaData: [], statusWorkFlow: [],
});

const emptyDoc = () => ({
  ticketKey: "", ticketType: "", tenantId: "", deleted: false,
  creators: [], viewers: [], assignee: [],
  ticketSubType: [],
});

/* deep-clean: drop empty strings / empty arrays / empty objects,
   keep booleans and numbers (incl. 0/false) as explicitly provided */
function clean(value) {
  if (Array.isArray(value)) {
    const arr = value.map(clean).filter((v) => v !== undefined);
    return arr.length ? arr : undefined;
  }
  if (value && typeof value === "object") {
    const out = {};
    Object.keys(value).forEach((k) => {
      const c = clean(value[k]);
      if (c !== undefined) out[k] = c;
    });
    return Object.keys(out).length ? out : undefined;
  }
  if (typeof value === "string") {
    return value.trim() === "" ? undefined : value;
  }
  return value; // booleans, numbers, null->kept only if explicitly not undefined
}

function fieldToJson(f) {
  if (f.type === "dropdown") {
    return {
      label: f.label, key: f.key, type: f.type,
      options: (f.options || []).map((o) => ({ value: o.value, key: o.key })),
    };
  }
  if (f.type === "casecadeDropdown") {
    return {
      label: f.label, key: f.key, type: f.type,
      childLabel: f.childLabel, childKey: f.childKey,
      options: (f.options || []).map((p) => ({
        value: p.name, label: p.name, type: f.parentOptionType, key: p.key,
        options: (p.children || []).map((c) => ({ value: c.value, key: c.key })),
      })),
    };
  }
  return { label: f.label, key: f.key, type: f.type };
}

function statusToJson(s) {
  return {
    status: s.status, label: s.label, public: s.public,
    roles: s.roles, assignee: (s.assigneeRoles || []).map((r) => ({ roleId: r })),
    notification: s.notification, notificationType: s.notificationType,
    notificationRole: s.notificationRole,
    mandatoryCustomFields: s.mandatoryCustomFields,
    overrideStatus: s.overrideStatus,
    customFieldsMetaData: (s.customFieldsMetaData || []).map(fieldToJson),
    autoEscalationConfig: (s.autoEscalationConfig || []).map((e) => ({
      assignee: e.assignee, days: e.days === "" ? undefined : Number(e.days),
      overrideStatus: e.overrideStatus, notification: e.notification,
      notificationType: e.notificationType, ccAddress: e.ccAddress,
    })),
  };
}

function subTypeToJson(st) {
  return {
    ticketKey: st.ticketKey, prefix: st.prefix, ticketType: st.ticketType,
    customFieldsMetaData: (st.customFieldsMetaData || []).map(fieldToJson),
    statusWorkFlow: (st.statusWorkFlow || []).map(statusToJson),
    deleted: st.deleted, isAutoEscalation: st.isAutoEscalation,
  };
}

function docToJson(doc) {
  const raw = {
    ticketKey: doc.ticketKey, ticketType: doc.ticketType, tenantId: doc.tenantId,
    creators: doc.creators.map((r) => ({ roleId: r })),
    viewers: doc.viewers.map((r) => ({ roleId: r })),
    assignee: doc.assignee.map((r) => ({ roleId: r })),
    deleted: doc.deleted,
    ticketSubType: doc.ticketSubType.map(subTypeToJson),
  };
  return clean(raw);
}

/* reverse mapping: raw ticket-type JSON -> editable state, so an
   existing schema (or one exported by this tool) can be re-opened */

function rolesFromJson(arr) {
  return (Array.isArray(arr) ? arr : []).map((r) => (r && r.roleId) || "").filter(Boolean);
}

function fieldFromJson(f) {
  if (!f) return newField("text");
  if (f.type === "dropdown") {
    return {
      id: uid(), label: f.label || "", key: f.key || "", type: "dropdown",
      options: (f.options || []).map((o) => ({ id: uid(), value: o.value || "", key: o.key != null ? String(o.key) : "" })),
    };
  }
  if (f.type === "casecadeDropdown" || f.type === "cascadeDropdown") {
    const firstParentType = f.options && f.options[0] && f.options[0].type;
    return {
      id: uid(), label: f.label || "", key: f.key || "", type: "casecadeDropdown",
      childLabel: f.childLabel || "", childKey: f.childKey || "",
      parentOptionType: firstParentType || "multiplechoice",
      options: (f.options || []).map((p) => ({
        id: uid(), name: p.value || p.label || "", key: p.key != null ? String(p.key) : "",
        children: (p.options || []).map((c) => ({ id: uid(), value: c.value || "", key: c.key != null ? String(c.key) : "" })),
      })),
    };
  }
  return { id: uid(), label: f.label || "", key: f.key || "", type: f.type || "text" };
}

function statusFromJson(s) {
  if (!s) return newStatus();
  return {
    id: uid(), status: s.status || "", label: s.label || "", public: !!s.public,
    roles: Array.isArray(s.roles) ? s.roles : [],
    assigneeRoles: rolesFromJson(s.assignee),
    notification: !!s.notification,
    notificationType: Array.isArray(s.notificationType) ? s.notificationType : [],
    notificationRole: Array.isArray(s.notificationRole) ? s.notificationRole : (Array.isArray(s.notificationRoles) ? s.notificationRoles : []),
    mandatoryCustomFields: Array.isArray(s.mandatoryCustomFields) ? s.mandatoryCustomFields : [],
    overrideStatus: s.overrideStatus || "",
    customFieldsMetaData: (s.customFieldsMetaData || []).map(fieldFromJson),
    autoEscalationConfig: (s.autoEscalationConfig || []).map((e) => ({
      id: uid(), assignee: e.assignee || "", days: e.days != null ? String(e.days) : "",
      overrideStatus: e.overrideStatus || "", notification: !!e.notification,
      notificationType: Array.isArray(e.notificationType) ? e.notificationType : [],
      ccAddress: Array.isArray(e.ccAddress) ? e.ccAddress : [],
    })),
  };
}

function subTypeFromJson(st) {
  if (!st) return newSubType();
  return {
    id: uid(), ticketKey: st.ticketKey || "", prefix: st.prefix || "", ticketType: st.ticketType || "",
    deleted: !!st.deleted, isAutoEscalation: !!st.isAutoEscalation,
    customFieldsMetaData: (st.customFieldsMetaData || []).map(fieldFromJson),
    statusWorkFlow: (st.statusWorkFlow || []).map(statusFromJson),
  };
}

function docFromRaw(src) {
  if (!src || typeof src !== "object") return emptyDoc();
  return {
    ticketKey: src.ticketKey || "", ticketType: src.ticketType || "", tenantId: src.tenantId || "",
    deleted: !!src.deleted,
    creators: rolesFromJson(src.creators),
    viewers: rolesFromJson(src.viewers),
    assignee: rolesFromJson(src.assignee),
    ticketSubType: (src.ticketSubType || []).map(subTypeFromJson),
  };
}

/* Backward-compatible single-doc conversion (used by the Excel importer,
   which always produces exactly one ticket-type object). Takes the first
   entry if handed an array. */
function jsonToDoc(raw) {
  return docFromRaw(Array.isArray(raw) ? raw[0] : raw);
}

/* A JSON import can be a single ticket-type object OR an array of many —
   this is the common shape for an export of "all my ticket types". Every
   entry becomes its own library item instead of only keeping the first. */
function parseTicketTypesJson(raw) {
  const arr = Array.isArray(raw) ? raw : [raw];
  return arr.map(docFromRaw).filter((d) => d.ticketType || d.ticketKey || d.ticketSubType.length);
}

/* Match key for a whole ticket-type document — used to tell whether an
   imported entry is updating something already in the library or is brand new. */
const docMatchKey = (d) => (d.ticketKey || d.ticketType || "").trim().toLowerCase();

function upsertLibraryEntries(library, incomingDocs) {
  const result = library.map((e) => ({ ...e }));
  let added = 0;
  let updated = 0;
  const touchedIds = [];
  incomingDocs.forEach((incoming) => {
    const key = docMatchKey(incoming);
    const idx = key ? result.findIndex((e) => docMatchKey(e.doc) === key) : -1;
    if (idx >= 0) {
      const { doc: merged } = mergeIncomingDoc(result[idx].doc, incoming);
      result[idx] = { ...result[idx], doc: merged };
      updated += 1;
      touchedIds.push(result[idx].id);
    } else {
      const entry = { id: uid(), doc: incoming };
      result.push(entry);
      added += 1;
      touchedIds.push(entry.id);
    }
  });
  return { library: result, added, updated, touchedIds };
}

/* ------------------------------------------------------------------ */
/* merge / upsert — importing a file adds new sub-types and updates    */
/* matching existing ones (by sub-type key, falling back to name)      */
/* instead of wiping out everything already in the editor.             */
/* ------------------------------------------------------------------ */

const subTypeMatchKey = (st) => (st.ticketKey || st.ticketType || "").trim().toLowerCase();

function upsertSubTypes(existingList, incomingList) {
  const result = existingList.map((s) => ({ ...s }));
  let added = 0;
  let updated = 0;
  incomingList.forEach((incoming) => {
    const key = subTypeMatchKey(incoming);
    const idx = key ? result.findIndex((s) => subTypeMatchKey(s) === key) : -1;
    if (idx >= 0) {
      result[idx] = { ...incoming, id: result[idx].id };
      updated += 1;
    } else {
      result.push(incoming);
      added += 1;
    }
  });
  return { subTypes: result, added, updated };
}

function mergeIncomingDoc(existing, incoming) {
  const { subTypes, added, updated } = upsertSubTypes(existing.ticketSubType, incoming.ticketSubType);
  const merged = {
    ticketKey: incoming.ticketKey || existing.ticketKey,
    ticketType: incoming.ticketType || existing.ticketType,
    tenantId: incoming.tenantId || existing.tenantId,
    deleted: existing.ticketSubType.length ? existing.deleted : incoming.deleted,
    creators: incoming.creators && incoming.creators.length ? incoming.creators : existing.creators,
    viewers: incoming.viewers && incoming.viewers.length ? incoming.viewers : existing.viewers,
    assignee: incoming.assignee && incoming.assignee.length ? incoming.assignee : existing.assignee,
    ticketSubType: subTypes,
  };
  return { doc: merged, added, updated };
}

/* ------------------------------------------------------------------ */
/* Excel (.xlsx) import — same 6-sheet convention as the standalone     */
/* excel_to_ticket_json.py script. Produces the same raw JSON shape    */
/* docToJson() outputs, then hands it to jsonToDoc() above so both     */
/* import paths (Excel + JSON) share one reverse-mapping code path.    */
/* ------------------------------------------------------------------ */

const xnorm = (v) => (v === null || v === undefined ? "" : String(v).trim());
const xIsYes = (v) => ["yes", "y", "true", "1"].includes(xnorm(v).toLowerCase());
const xSplitRoles = (v) => xnorm(v).split(",").map((s) => s.trim()).filter(Boolean);
const xToCamelKey = (label) => {
  const stripped = xnorm(label).replace(/\(.*?\)/g, "");
  const words = stripped.match(/[A-Za-z0-9]+/g) || [];
  if (!words.length) return "field";
  return words[0].toLowerCase() + words.slice(1).map((w) => w[0].toUpperCase() + w.slice(1).toLowerCase()).join("");
};
const xToStatusCode = (label) => xnorm(label).toLowerCase().replace(/[^a-z0-9]/g, "");
const xSheetKey = (name) => name.toLowerCase().replace(/[^a-z0-9]/g, "");

function xFindSheet(wb, wanted) {
  const wantedKey = xSheetKey(wanted);
  const match = wb.SheetNames.find((n) => xSheetKey(n) === wantedKey);
  return match ? wb.Sheets[match] : null;
}

function xRowsOf(sheet) {
  if (!sheet) return [];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
  return rows.filter((r) => r.some((c) => c !== null && xnorm(c) !== ""));
}
const xCell = (row, i) => (i < row.length ? row[i] : null);
const xFindCol = (header, needle) => header.findIndex((h) => h.includes(needle));

function xParseTicketDetails(sheet) {
  const rows = xRowsOf(sheet);
  if (rows.length < 2) return {};
  const header = rows[0].map((h) => xnorm(h).toLowerCase());
  const col = (n) => xFindCol(header, n);
  const r = rows[1];
  return {
    ticketKey: xnorm(xCell(r, col("ticket key"))),
    ticketType: xnorm(xCell(r, col("ticket type"))),
    tenantId: xnorm(xCell(r, col("tenant"))),
    creators: xSplitRoles(xCell(r, col("creator"))),
    viewers: xSplitRoles(xCell(r, col("viewer"))),
    assignee: xSplitRoles(xCell(r, col("assignee"))),
    deleted: xIsYes(xCell(r, col("delete"))),
  };
}

function xParseSubTypes(sheet) {
  const rows = xRowsOf(sheet);
  if (!rows.length) return { subtypes: {}, order: [] };
  const header = rows[0].map((h) => xnorm(h).toLowerCase());
  const col = (n) => xFindCol(header, n);
  const subtypes = {};
  const order = [];
  for (const r of rows.slice(1)) {
    const name = xnorm(xCell(r, col("ticket subtype")));
    if (!name) continue;
    const key = xnorm(xCell(r, col("subtype key"))) || name;
    subtypes[name] = {
      ticketKey: key,
      prefix: xnorm(xCell(r, col("prefix"))),
      isAutoEscalation: xIsYes(xCell(r, col("auto escalation"))),
      deleted: xIsYes(xCell(r, col("delete"))),
      ticketType: name,
      customFieldsMetaData: [],
      statusWorkFlow: [],
    };
    order.push(name);
  }
  return { subtypes, order };
}

function xParseTicketStructure(sheet, subtypes) {
  const rows = xRowsOf(sheet);
  let currentSubtype = null;
  let currentField = null;
  for (const r of rows.slice(1)) {
    const subtypeCell = xnorm(xCell(r, 0));
    const labelCell = xnorm(xCell(r, 1));
    const typeCell = xnorm(xCell(r, 2));
    const rolesCell = xnorm(xCell(r, 4));

    if (subtypeCell) currentSubtype = subtypeCell;

    if (labelCell) {
      const typeMatch = labelCell.match(/\((.*?)\)/);
      const explicitType = typeMatch ? typeMatch[1].trim().toLowerCase() : "";
      const cleanLabel = labelCell.replace(/\(.*?\)/g, "").trim();
      let ftype = "dropdown";
      if (explicitType.includes("drop")) ftype = "dropdown";
      else if (explicitType.includes("date")) ftype = "date";
      else if (explicitType.includes("time")) ftype = "time";
      else if (explicitType.includes("text")) ftype = "text";

      currentField = { label: cleanLabel, key: xToCamelKey(cleanLabel), type: ftype, editRoles: xSplitRoles(rolesCell), options: [] };
      if (typeCell && ftype === "dropdown") currentField.options.push({ value: typeCell, key: "" });
      if (currentSubtype && subtypes[currentSubtype]) subtypes[currentSubtype].customFieldsMetaData.push(currentField);
      continue;
    }
    if (currentField && typeCell) currentField.options.push({ value: typeCell, key: "" });
  }
  Object.values(subtypes).forEach((st) => {
    st.customFieldsMetaData.forEach((f) => {
      if (f.type === "dropdown" && !f.options.length) f.type = "text";
      f.options.forEach((o, i) => { if (!o.key) o.key = String(i + 1); });
    });
  });
}

function xParseStatusAccessRoles(sheet) {
  const rows = xRowsOf(sheet);
  if (!rows.length) return {};
  let headerIdx = rows.findIndex((r) => xnorm(xCell(r, 0)).toLowerCase().includes("status"));
  if (headerIdx === -1) return {};
  const out = {};
  for (const r of rows.slice(headerIdx + 1)) {
    const label = xnorm(xCell(r, 0));
    if (!label) continue;
    out[xToStatusCode(label)] = { label, roles: xSplitRoles(xCell(r, 1)) };
  }
  return out;
}

function xParseMandateRules(sheet) {
  const rows = xRowsOf(sheet);
  if (!rows.length) return {};
  let headerIdx = rows.findIndex((r) => xnorm(xCell(r, 0)).toLowerCase().includes("status"));
  if (headerIdx === -1) return {};
  const out = {};
  for (const r of rows.slice(headerIdx + 1)) {
    const label = xnorm(xCell(r, 0));
    if (!label) continue;
    const imagesCell = xnorm(xCell(r, 1));
    const commentsCell = xnorm(xCell(r, 2));
    const assigneeYnCell = xnorm(xCell(r, 3));
    const assigneeRoleCell = xnorm(xCell(r, 4));
    let commentsLabel = commentsCell && !xIsYes(commentsCell) ? commentsCell : "Comments";
    commentsLabel = commentsLabel.replace(/\(mandatory\)/i, "").trim();
    out[xToStatusCode(label)] = {
      mandatoryImages: xIsYes(imagesCell),
      mandatoryComments: xIsYes(commentsCell) || !!commentsCell,
      commentsLabel,
      mandatoryNewAssignee: xIsYes(assigneeYnCell),
      newAssigneeRole: assigneeRoleCell,
    };
  }
  return out;
}

function xParseEscalationDetails(sheet) {
  const rows = xRowsOf(sheet);
  const result = {};
  const warnings = [];
  let currentStatus = null;
  let inTable = false;
  for (const r of rows) {
    const a0 = xnorm(xCell(r, 0));
    const m = a0.match(/^status\s*:\s*(.+)$/i);
    if (m) {
      currentStatus = xToStatusCode(m[1]);
      result[currentStatus] = result[currentStatus] || {};
      inTable = false;
      continue;
    }
    if (a0.toLowerCase() === "ticket subtype") { inTable = true; continue; }
    if (!currentStatus || !inTable) continue;

    const subtype = a0;
    if (!subtype) continue;
    const issueType = xnorm(xCell(r, 1));
    const chain = [];
    let colIdx = 2;
    while (colIdx + 1 < r.length) {
      const days = xnorm(xCell(r, colIdx));
      const role = xnorm(xCell(r, colIdx + 1));
      if (days && role) chain.push({ assignee: role, days });
      else if (days || role) {
        warnings.push(`Escalation row for '${subtype}' (${issueType || "no issue type"}) under STATUS:${currentStatus} has a Days/Role pair with only one side filled — skipped that step.`);
      }
      colIdx += 2;
    }
    const bucket = result[currentStatus][subtype] || [];
    if (bucket.length && chain.length) {
      warnings.push(`Sub-type '${subtype}' has more than one Issue Type row under STATUS:${currentStatus} (latest: '${issueType}') — escalation steps were merged into one chain.`);
    }
    result[currentStatus][subtype] = bucket.concat(chain);
  }
  return { escalationByStatus: result, warnings };
}

function xBuildStatusWorkflow(subtypeName, globalStatusRoles, mandateRules, escalationByStatus) {
  const preferredOrder = ["open", "inprogress", "resolved", "closed"];
  const allCodes = new Set([...Object.keys(globalStatusRoles), ...Object.keys(mandateRules), ...Object.keys(escalationByStatus)]);
  const orderedCodes = preferredOrder.filter((c) => allCodes.has(c));
  [...allCodes].sort().forEach((c) => { if (!orderedCodes.includes(c)) orderedCodes.push(c); });

  return orderedCodes.map((code) => {
    const roleInfo = globalStatusRoles[code] || {};
    const mandate = mandateRules[code] || {};
    const escRows = (escalationByStatus[code] || {})[subtypeName] || [];

    const mandatoryFields = [];
    const customFields = [];
    if (mandate.mandatoryImages) mandatoryFields.push("images");
    if (mandate.mandatoryComments) {
      const key = xToCamelKey(mandate.commentsLabel || "Comments");
      mandatoryFields.push(key);
      customFields.push({ label: mandate.commentsLabel || "Comments", key, type: "text" });
    }
    if (mandate.mandatoryNewAssignee) mandatoryFields.push("newAssignee");

    const assigneeRoles = mandate.newAssigneeRole ? xSplitRoles(mandate.newAssigneeRole) : [];
    const autoEscalation = escRows.map((e) => ({
      assignee: e.assignee,
      days: /^\d+$/.test(xnorm(e.days)) ? Number(e.days) : e.days,
      notification: true,
      notificationType: ["inApp", "email"],
    }));

    return {
      status: code,
      label: roleInfo.label || code.charAt(0).toUpperCase() + code.slice(1),
      public: true,
      roles: roleInfo.roles || [],
      assignee: assigneeRoles.map((r) => ({ roleId: r })),
      notification: autoEscalation.length > 0,
      notificationType: autoEscalation.length ? ["inApp", "email"] : [],
      notificationRole: roleInfo.roles || [],
      mandatoryCustomFields: mandatoryFields,
      customFieldsMetaData: customFields,
      autoEscalationConfig: autoEscalation,
    };
  });
}

function xFieldToOutput(f) {
  const out = { label: f.label, key: f.key, type: f.type };
  if (f.type === "dropdown") out.options = f.options.map((o) => ({ value: o.value, key: o.key }));
  return out;
}

/** Parses a workbook (from XLSX.read) built with the 6-sheet convention
 * into the same raw JSON shape docToJson() produces, then converts it
 * to editable state via jsonToDoc(). Returns { doc, warnings, missingSheets }. */
function excelWorkbookToDoc(wb) {
  const sheetSpecs = [
    ["Ticket Details", "details"], ["Sub Types", "subtypes"], ["Ticket Structure", "structure"],
    ["Status Access Roles", "statusRoles"], ["Mandate rules for status change", "mandate"],
    ["Escalation details", "escalation"],
  ];
  const sheets = {};
  const missingSheets = [];
  sheetSpecs.forEach(([name, key]) => {
    const s = xFindSheet(wb, name);
    if (!s) missingSheets.push(name);
    sheets[key] = s;
  });

  const details = sheets.details ? xParseTicketDetails(sheets.details) : {};
  const { subtypes, order } = sheets.subtypes ? xParseSubTypes(sheets.subtypes) : { subtypes: {}, order: [] };
  if (sheets.structure) xParseTicketStructure(sheets.structure, subtypes);
  const globalStatusRoles = sheets.statusRoles ? xParseStatusAccessRoles(sheets.statusRoles) : {};
  const mandateRules = sheets.mandate ? xParseMandateRules(sheets.mandate) : {};
  const { escalationByStatus, warnings } = sheets.escalation ? xParseEscalationDetails(sheets.escalation) : { escalationByStatus: {}, warnings: [] };

  const ticketSubType = order.map((name) => {
    const st = subtypes[name];
    return {
      ticketKey: st.ticketKey, prefix: st.prefix, ticketType: st.ticketType,
      customFieldsMetaData: st.customFieldsMetaData.map(xFieldToOutput),
      statusWorkFlow: xBuildStatusWorkflow(name, globalStatusRoles, mandateRules, escalationByStatus),
      deleted: st.deleted, isAutoEscalation: st.isAutoEscalation,
    };
  });

  const rawDoc = {
    ticketKey: details.ticketKey || "", ticketType: details.ticketType || "", tenantId: details.tenantId || "",
    deleted: !!details.deleted,
    creators: (details.creators || []).map((r) => ({ roleId: r })),
    viewers: (details.viewers || []).map((r) => ({ roleId: r })),
    assignee: (details.assignee || []).map((r) => ({ roleId: r })),
    ticketSubType,
  };

  return { doc: jsonToDoc(rawDoc), warnings, missingSheets };
}

/* ------------------------------------------------------------------ */
/* tiny UI primitives — dark, navy/black theme                         */
/* ------------------------------------------------------------------ */

const Label = ({ children, mono }) => (
  <label className={`block text-[11px] font-semibold uppercase tracking-wider text-slate-500 mb-1.5 ${mono ? "font-mono normal-case tracking-normal" : ""}`}>
    {children}
  </label>
);

const Input = ({ value, onChange, placeholder, mono, ...rest }) => (
  <input
    value={value}
    onChange={(e) => onChange(e.target.value)}
    placeholder={placeholder}
    className={`w-full rounded-lg border border-slate-700 bg-[#0B101B] px-3 py-2 text-sm text-slate-100 placeholder:text-slate-600 outline-none transition focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20 ${mono ? "font-mono" : ""}`}
    {...rest}
  />
);

const Select = ({ value, onChange, options }) => (
  <select
    value={value}
    onChange={(e) => onChange(e.target.value)}
    className="w-full rounded-lg border border-slate-700 bg-[#0B101B] px-3 py-2 text-sm text-slate-100 outline-none transition focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20"
  >
    {options.map((o) => (
      <option key={o.value} value={o.value} className="bg-[#0B101B]">{o.label}</option>
    ))}
  </select>
);

const Toggle = ({ checked, onChange, label }) => (
  <button
    type="button"
    onClick={() => onChange(!checked)}
    className="flex items-center gap-2 text-sm text-slate-300"
  >
    <span className={`relative inline-flex h-5 w-9 items-center rounded-full transition ${checked ? "bg-teal-500" : "bg-slate-700"}`}>
      <span className="inline-block h-3.5 w-3.5 transform rounded-full bg-white shadow transition" style={{ transform: checked ? "translateX(18px)" : "translateX(4px)" }} />
    </span>
    {label}
  </button>
);

const IconBtn = ({ onClick, title, danger, children }) => (
  <button
    type="button"
    onClick={onClick}
    title={title}
    className={`inline-flex items-center justify-center h-7 w-7 rounded-md transition ${danger ? "text-rose-400 hover:bg-rose-500/10 hover:text-rose-300" : "text-slate-500 hover:bg-white/5 hover:text-slate-200"}`}
  >
    {children}
  </button>
);

const Btn = ({ onClick, children, variant = "primary", small }) => {
  const base = "inline-flex items-center gap-1.5 rounded-lg font-medium transition";
  const size = small ? "px-2.5 py-1.5 text-xs" : "px-3.5 py-2 text-sm";
  const styles = {
    primary: "bg-slate-100 text-slate-900 hover:bg-white",
    accent: "bg-teal-500 text-white hover:bg-teal-400",
    ghost: "bg-transparent text-slate-300 hover:bg-white/5 border border-slate-700",
    dashed: "border border-dashed border-slate-600 text-slate-400 hover:border-teal-400 hover:text-teal-300 hover:bg-teal-500/5",
  };
  return (
    <button type="button" onClick={onClick} className={`${base} ${size} ${styles[variant]}`}>
      {children}
    </button>
  );
};

const Chip = ({ text, onRemove, tone = "slate" }) => {
  const tones = {
    slate: "bg-white/5 text-slate-300",
    teal: "bg-teal-500/10 text-teal-300",
    amber: "bg-amber-500/10 text-amber-300",
  };
  return (
    <span className={`inline-flex items-center gap-1 rounded-md px-2 py-1 text-xs font-mono ${tones[tone]}`}>
      {text}
      <button type="button" onClick={onRemove} className="hover:text-rose-400"><X size={12} /></button>
    </span>
  );
};

const TagInput = ({ values, onChange, placeholder, tone }) => {
  const [draft, setDraft] = useState("");
  const commit = () => {
    const v = draft.trim();
    if (v && !values.includes(v)) onChange([...values, v]);
    setDraft("");
  };
  return (
    <div className="rounded-lg border border-slate-700 bg-[#0B101B] p-2">
      <div className="flex flex-wrap gap-1.5 mb-1.5">
        {values.map((v, i) => (
          <Chip key={v + i} text={v} tone={tone} onRemove={() => onChange(values.filter((_, idx) => idx !== i))} />
        ))}
      </div>
      <input
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onKeyDown={(e) => { if (e.key === "Enter" || e.key === ",") { e.preventDefault(); commit(); } }}
        onBlur={commit}
        placeholder={placeholder}
        className="w-full text-sm text-slate-100 bg-transparent outline-none placeholder:text-slate-600"
      />
    </div>
  );
};

const CheckChips = ({ options, values, onChange }) => (
  <div className="flex flex-wrap gap-1.5">
    {options.map((o) => {
      const active = values.includes(o);
      return (
        <button
          key={o}
          type="button"
          onClick={() => onChange(active ? values.filter((v) => v !== o) : [...values, o])}
          className={`rounded-md px-2.5 py-1 text-xs font-mono border transition ${active ? "border-teal-400 bg-teal-500 text-white" : "border-slate-700 text-slate-400 hover:border-slate-500"}`}
        >
          {o}
        </button>
      );
    })}
  </div>
);

const Section = ({ title, icon, children, onAdd, addLabel, right }) => (
  <div className="mb-6">
    <div className="flex items-center justify-between mb-3">
      <div className="flex items-center gap-2 text-slate-100 font-semibold text-sm">
        {icon}{title}
      </div>
      <div className="flex items-center gap-2">
        {right}
        {onAdd && <Btn small variant="dashed" onClick={onAdd}><Plus size={14} />{addLabel}</Btn>}
      </div>
    </div>
    {children}
  </div>
);

const Accordion = ({ title, subtitle, tone = "slate", defaultOpen, onDelete, children }) => {
  const [open, setOpen] = useState(!!defaultOpen);
  const tones = { slate: "border-slate-700/70", teal: "border-teal-500/30", amber: "border-amber-500/30" };
  return (
    <div className={`rounded-xl border ${tones[tone]} bg-[#121826] mb-3 overflow-hidden`}>
      <div className="flex items-center justify-between px-3 py-2.5 cursor-pointer select-none" onClick={() => setOpen(!open)}>
        <div className="flex items-center gap-2 min-w-0">
          {open ? <ChevronDown size={16} className="text-slate-500 shrink-0" /> : <ChevronRight size={16} className="text-slate-500 shrink-0" />}
          <div className="min-w-0">
            <div className="text-sm font-medium text-slate-100 truncate">{title}</div>
            {subtitle && <div className="text-xs text-slate-500 font-mono truncate">{subtitle}</div>}
          </div>
        </div>
        {onDelete && (
          <IconBtn danger title="Remove" onClick={(e) => { e.stopPropagation(); onDelete(); }}>
            <Trash2 size={14} />
          </IconBtn>
        )}
      </div>
      {open && <div className="px-3 pb-3 pt-1 border-t border-slate-800">{children}</div>}
    </div>
  );
};

/* ------------------------------------------------------------------ */
/* Options editors                                                     */
/* ------------------------------------------------------------------ */

function SimpleOptionsEditor({ options, onChange }) {
  const update = (id, patch) => onChange(options.map((o) => (o.id === id ? { ...o, ...patch } : o)));
  const remove = (id) => onChange(options.filter((o) => o.id !== id));
  const add = () => onChange([...options, newSimpleOption()]);
  return (
    <div>
      <div className="grid grid-cols-[1fr_1fr_28px] gap-2 mb-1.5">
        <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">Value (label shown)</span>
        <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">Key</span>
        <span />
      </div>
      <div className="space-y-2">
        {options.map((o) => (
          <div key={o.id} className="grid grid-cols-[1fr_1fr_28px] gap-2 items-center">
            <Input value={o.value} onChange={(v) => update(o.id, { value: v })} placeholder="e.g. Marshmallow-Bl" />
            <Input mono value={o.key} onChange={(v) => update(o.id, { key: v })} placeholder="e.g. 24" />
            <IconBtn danger title="Remove option" onClick={() => remove(o.id)}><Trash2 size={14} /></IconBtn>
          </div>
        ))}
      </div>
      <button type="button" onClick={add} className="mt-2 text-xs font-medium text-teal-300 hover:text-teal-200 inline-flex items-center gap-1">
        <Plus size={13} /> Add option
      </button>
    </div>
  );
}

function CascadeOptionsEditor({ field, onChange }) {
  const options = field.options || [];
  const updateParent = (id, patch) => onChange({ ...field, options: options.map((p) => (p.id === id ? { ...p, ...patch } : p)) });
  const removeParent = (id) => onChange({ ...field, options: options.filter((p) => p.id !== id) });
  const addParent = () => onChange({ ...field, options: [...options, newParentOption()] });

  const updateChild = (pid, cid, patch) => {
    onChange({
      ...field,
      options: options.map((p) => p.id === pid
        ? { ...p, children: p.children.map((c) => (c.id === cid ? { ...c, ...patch } : c)) }
        : p),
    });
  };
  const removeChild = (pid, cid) => onChange({
    ...field, options: options.map((p) => p.id === pid ? { ...p, children: p.children.filter((c) => c.id !== cid) } : p),
  });
  const addChild = (pid) => onChange({
    ...field, options: options.map((p) => p.id === pid ? { ...p, children: [...p.children, newChildOption()] } : p),
  });

  return (
    <div>
      <div className="flex items-center gap-3 mb-3">
        <span className="text-xs text-slate-500">Parent option style:</span>
        <div className="flex gap-1.5">
          {["multiplechoice", "dropdown"].map((t) => (
            <button
              key={t}
              type="button"
              onClick={() => onChange({ ...field, parentOptionType: t })}
              className={`px-2.5 py-1 rounded-md text-xs font-mono border ${field.parentOptionType === t ? "border-teal-400 bg-teal-500 text-white" : "border-slate-700 text-slate-400"}`}
            >
              {t}
            </button>
          ))}
        </div>
      </div>
      {options.map((p) => (
        <div key={p.id} className="rounded-lg border border-slate-800 bg-[#0D1320] p-2.5 mb-2.5">
          <div className="grid grid-cols-[1fr_120px_28px] gap-2 items-center mb-2">
            <Input value={p.name} onChange={(v) => updateParent(p.id, { name: v })} placeholder="Parent value, e.g. HC-House Hold-Bath" />
            <Input mono value={p.key} onChange={(v) => updateParent(p.id, { key: v })} placeholder="key" />
            <IconBtn danger title="Remove parent" onClick={() => removeParent(p.id)}><Trash2 size={14} /></IconBtn>
          </div>
          <div className="pl-3 border-l-2 border-teal-500/20 space-y-1.5">
            {p.children.map((c) => (
              <div key={c.id} className="grid grid-cols-[1fr_120px_28px] gap-2 items-center">
                <Input value={c.value} onChange={(v) => updateChild(p.id, c.id, { value: v })} placeholder="Child value, e.g. Buttercup" />
                <Input mono value={c.key} onChange={(v) => updateChild(p.id, c.id, { key: v })} placeholder="key" />
                <IconBtn danger title="Remove child" onClick={() => removeChild(p.id, c.id)}><Trash2 size={14} /></IconBtn>
              </div>
            ))}
            <button type="button" onClick={() => addChild(p.id)} className="text-xs font-medium text-teal-300 hover:text-teal-200 inline-flex items-center gap-1 mt-1">
              <Plus size={12} /> Add child option
            </button>
          </div>
        </div>
      ))}
      <button type="button" onClick={addParent} className="text-xs font-medium text-teal-300 hover:text-teal-200 inline-flex items-center gap-1">
        <Plus size={13} /> Add parent option
      </button>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Custom field editor (used both at sub-type level and status level)  */
/* ------------------------------------------------------------------ */

function CustomFieldEditor({ field, onChange, onDelete, restrictToDropdown }) {
  const typeOptions = restrictToDropdown
    ? FIELD_TYPES.filter((t) => t.value === "dropdown")
    : FIELD_TYPES;

  const changeType = (type) => {
    const base = { id: field.id, label: field.label, key: field.key };
    onChange(newField(type) && { ...newField(type), ...base, type });
  };

  return (
    <Accordion
      title={field.label || "(untitled field)"}
      subtitle={field.key ? `${field.key} · ${field.type}` : field.type}
      onDelete={onDelete}
    >
      <div className="grid grid-cols-2 gap-3 mb-3">
        <div>
          <Label>Field label</Label>
          <Input value={field.label} onChange={(v) => onChange({ ...field, label: v })} placeholder="e.g. Division Code" />
        </div>
        <div>
          <Label mono>Field key</Label>
          <Input mono value={field.key} onChange={(v) => onChange({ ...field, key: v })} placeholder="e.g. divisionCode" />
        </div>
      </div>
      <div className="mb-3">
        <Label>Field type</Label>
        <Select value={field.type} onChange={changeType} options={typeOptions} />
      </div>

      {field.type === "casecadeDropdown" && (
        <>
          <div className="grid grid-cols-2 gap-3 mb-3">
            <div>
              <Label>Child field label</Label>
              <Input value={field.childLabel} onChange={(v) => onChange({ ...field, childLabel: v })} placeholder="e.g. Plano" />
            </div>
            <div>
              <Label mono>Child field key</Label>
              <Input mono value={field.childKey} onChange={(v) => onChange({ ...field, childKey: v })} placeholder="e.g. planoCode" />
            </div>
          </div>
          <Label>Options</Label>
          <CascadeOptionsEditor field={field} onChange={onChange} />
        </>
      )}

      {field.type === "dropdown" && (
        <>
          <Label>Options</Label>
          <SimpleOptionsEditor options={field.options || []} onChange={(opts) => onChange({ ...field, options: opts })} />
        </>
      )}

      {(field.type === "date" || field.type === "time" || field.type === "text") && (
        <p className="text-xs text-slate-500 italic">No options needed for this field type.</p>
      )}
    </Accordion>
  );
}

function CustomFieldsList({ fields, onChange, restrictToDropdown }) {
  const update = (id, patch) => onChange(fields.map((f) => (f.id === id ? patch : f)));
  const remove = (id) => onChange(fields.filter((f) => f.id !== id));
  const add = () => onChange([...fields, newField(restrictToDropdown ? "dropdown" : "text")]);
  return (
    <div>
      {fields.map((f) => (
        <CustomFieldEditor
          key={f.id}
          field={f}
          restrictToDropdown={restrictToDropdown}
          onChange={(patch) => update(f.id, patch)}
          onDelete={() => remove(f.id)}
        />
      ))}
      <Btn small variant="dashed" onClick={add}><Plus size={14} />Add custom field</Btn>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Auto escalation                                                     */
/* ------------------------------------------------------------------ */

function AutoEscalationList({ rules, onChange }) {
  const update = (id, patch) => onChange(rules.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  const remove = (id) => onChange(rules.filter((r) => r.id !== id));
  const add = () => onChange([...rules, newAutoEscalation()]);
  return (
    <div>
      {rules.map((r) => (
        <div key={r.id} className="rounded-lg border border-amber-500/20 bg-amber-500/[0.04] p-3 mb-2.5">
          <div className="flex justify-between items-center mb-2">
            <div className="flex items-center gap-1.5 text-amber-400 text-xs font-semibold"><AlertTriangle size={13} /> Escalation rule</div>
            <IconBtn danger onClick={() => remove(r.id)}><Trash2 size={14} /></IconBtn>
          </div>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <div>
              <Label mono>Assignee role</Label>
              <Input mono value={r.assignee} onChange={(v) => update(r.id, { assignee: v })} placeholder="Dist_Mgr" />
            </div>
            <div>
              <Label>Days</Label>
              <Input value={r.days} onChange={(v) => update(r.id, { days: v.replace(/[^0-9]/g, "") })} placeholder="2" />
            </div>
            <div>
              <Label mono>Override status</Label>
              <Input mono value={r.overrideStatus} onChange={(v) => update(r.id, { overrideStatus: v })} placeholder="rejected" />
            </div>
          </div>
          <div className="grid grid-cols-2 gap-3 mb-2">
            <div>
              <Label>Notify</Label>
              <Toggle checked={r.notification} onChange={(v) => update(r.id, { notification: v })} label={r.notification ? "Enabled" : "Disabled"} />
            </div>
            <div>
              <Label>Notification channels</Label>
              <CheckChips options={NOTIF_TYPES} values={r.notificationType} onChange={(v) => update(r.id, { notificationType: v })} />
            </div>
          </div>
          <div>
            <Label>CC addresses</Label>
            <TagInput values={r.ccAddress} onChange={(v) => update(r.id, { ccAddress: v })} placeholder="type email, press enter" />
          </div>
        </div>
      ))}
      <button type="button" onClick={add} className="text-xs font-medium text-amber-300 hover:text-amber-200 inline-flex items-center gap-1">
        <Plus size={13} /> Add escalation rule
      </button>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Status workflow                                                     */
/* ------------------------------------------------------------------ */

function StatusEditor({ status, onChange, onDelete }) {
  return (
    <Accordion
      title={status.label || "(untitled status)"}
      subtitle={status.status || "status-code"}
      tone="teal"
      onDelete={onDelete}
    >
      <div className="grid grid-cols-2 gap-3 mb-3">
        <div>
          <Label mono>Status code</Label>
          <Input mono value={status.status} onChange={(v) => onChange({ ...status, status: v })} placeholder="e.g. open" />
        </div>
        <div>
          <Label>Display label</Label>
          <Input value={status.label} onChange={(v) => onChange({ ...status, label: v })} placeholder="e.g. Open" />
        </div>
      </div>

      <div className="flex flex-wrap items-center gap-6 mb-3">
        <Toggle checked={status.public} onChange={(v) => onChange({ ...status, public: v })} label="Visible to viewers" />
        <Toggle checked={status.notification} onChange={(v) => onChange({ ...status, notification: v })} label="Send notification" />
      </div>

      {status.notification && (
        <div className="mb-3">
          <Label>Notification channels</Label>
          <CheckChips options={NOTIF_TYPES} values={status.notificationType} onChange={(v) => onChange({ ...status, notificationType: v })} />
        </div>
      )}

      <div className="grid grid-cols-2 gap-3 mb-3">
        <div>
          <Label>Roles allowed to set this status</Label>
          <TagInput values={status.roles} onChange={(v) => onChange({ ...status, roles: v })} placeholder="e.g. corpadmin" tone="teal" />
        </div>
        <div>
          <Label>Notify roles</Label>
          <TagInput values={status.notificationRole} onChange={(v) => onChange({ ...status, notificationRole: v })} placeholder="e.g. Dist_Mgr" tone="teal" />
        </div>
      </div>

      <div className="grid grid-cols-2 gap-3 mb-3">
        <div>
          <Label>Assignee roles at this status</Label>
          <TagInput values={status.assigneeRoles} onChange={(v) => onChange({ ...status, assigneeRoles: v })} placeholder="e.g. corpadmin" tone="teal" />
        </div>
        <div>
          <Label mono>Override status on completion</Label>
          <Input mono value={status.overrideStatus} onChange={(v) => onChange({ ...status, overrideStatus: v })} placeholder="e.g. closed" />
        </div>
      </div>

      <div className="mb-3">
        <Label>Mandatory fields to move into this status</Label>
        <TagInput values={status.mandatoryCustomFields} onChange={(v) => onChange({ ...status, mandatoryCustomFields: v })} placeholder="field key, press enter" tone="amber" />
      </div>

      <div className="mb-3">
        <Label>Custom fields captured at this status</Label>
        <CustomFieldsList
          fields={status.customFieldsMetaData}
          restrictToDropdown
          onChange={(fs) => onChange({ ...status, customFieldsMetaData: fs })}
        />
      </div>

      <div>
        <Label>Auto-escalation</Label>
        <AutoEscalationList rules={status.autoEscalationConfig} onChange={(v) => onChange({ ...status, autoEscalationConfig: v })} />
      </div>
    </Accordion>
  );
}

function StatusPipelinePreview({ statuses }) {
  if (!statuses.length) return null;
  return (
    <div className="mb-4 overflow-x-auto">
      <div className="flex items-center gap-0 min-w-max py-1">
        {statuses.map((s, i) => (
          <React.Fragment key={s.id}>
            <div className="flex flex-col items-center px-1">
              <div className={`h-2.5 w-2.5 rounded-full ${s.status ? "bg-teal-400" : "bg-slate-700"}`} />
              <div className="mt-1.5 text-[11px] font-mono text-slate-400 whitespace-nowrap">{s.label || s.status || "…"}</div>
            </div>
            {i < statuses.length - 1 && <div className="h-px w-8 bg-slate-800 mb-4" />}
          </React.Fragment>
        ))}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Sub-type editor                                                     */
/* ------------------------------------------------------------------ */

function SubTypeEditor({ sub, onChange, onDelete }) {
  const [tab, setTab] = useState("basic");
  const tabs = [
    { id: "basic", label: "Basic" },
    { id: "fields", label: `Custom Fields (${sub.customFieldsMetaData.length})` },
    { id: "workflow", label: `Status Workflow (${sub.statusWorkFlow.length})` },
  ];
  return (
    <Accordion
      title={sub.ticketType || "(untitled sub-type)"}
      subtitle={sub.ticketKey || "sub-type key"}
      onDelete={onDelete}
    >
      <div className="flex gap-1 mb-4 border-b border-slate-800">
        {tabs.map((t) => (
          <button
            key={t.id}
            type="button"
            onClick={() => setTab(t.id)}
            className={`px-3 py-2 text-xs font-medium border-b-2 -mb-px transition ${tab === t.id ? "border-teal-400 text-teal-300" : "border-transparent text-slate-500 hover:text-slate-300"}`}
          >
            {t.label}
          </button>
        ))}
      </div>

      {tab === "basic" && (
        <div>
          <div className="grid grid-cols-2 gap-3 mb-3">
            <div>
              <Label mono>Sub-type key</Label>
              <Input mono value={sub.ticketKey} onChange={(v) => onChange({ ...sub, ticketKey: v })} placeholder="e.g. StockGap" />
            </div>
            <div>
              <Label mono>Prefix</Label>
              <Input mono value={sub.prefix} onChange={(v) => onChange({ ...sub, prefix: v })} placeholder="e.g. ${locationId}" />
            </div>
          </div>
          <div className="mb-3">
            <Label>Sub-type name</Label>
            <Input value={sub.ticketType} onChange={(v) => onChange({ ...sub, ticketType: v })} placeholder="e.g. DS Stock gaps in plotted plannos" />
          </div>
          <div className="flex gap-6">
            <Toggle checked={sub.deleted} onChange={(v) => onChange({ ...sub, deleted: v })} label="Deleted" />
            <Toggle checked={sub.isAutoEscalation} onChange={(v) => onChange({ ...sub, isAutoEscalation: v })} label="Auto-escalation enabled" />
          </div>
        </div>
      )}

      {tab === "fields" && (
        <CustomFieldsList fields={sub.customFieldsMetaData} onChange={(fs) => onChange({ ...sub, customFieldsMetaData: fs })} />
      )}

      {tab === "workflow" && (
        <div>
          <StatusPipelinePreview statuses={sub.statusWorkFlow} />
          {sub.statusWorkFlow.map((s) => (
            <StatusEditor
              key={s.id}
              status={s}
              onChange={(patch) => onChange({ ...sub, statusWorkFlow: sub.statusWorkFlow.map((x) => (x.id === s.id ? patch : x)) })}
              onDelete={() => onChange({ ...sub, statusWorkFlow: sub.statusWorkFlow.filter((x) => x.id !== s.id) })}
            />
          ))}
          <Btn small variant="dashed" onClick={() => onChange({ ...sub, statusWorkFlow: [...sub.statusWorkFlow, newStatus()] })}>
            <Plus size={14} />Add status
          </Btn>
        </div>
      )}
    </Accordion>
  );
}

/* ------------------------------------------------------------------ */
/* Top-level tabs: Details / Roles                                     */
/* ------------------------------------------------------------------ */

function DetailsTab({ doc, onChange }) {
  return (
    <div className="max-w-2xl">
      <div className="grid grid-cols-2 gap-4 mb-4">
        <div>
          <Label mono>Ticket key</Label>
          <Input mono value={doc.ticketKey} onChange={(v) => onChange({ ...doc, ticketKey: v })} placeholder="e.g. DSHHPLv1" />
        </div>
        <div>
          <Label mono>Tenant ID</Label>
          <Input mono value={doc.tenantId} onChange={(v) => onChange({ ...doc, tenantId: v })} placeholder="e.g. abc" />
        </div>
      </div>
      <div className="mb-4">
        <Label>Ticket type name</Label>
        <Input value={doc.ticketType} onChange={(v) => onChange({ ...doc, ticketType: v })} placeholder="e.g. DS Household Plano Stock Request-v1" />
      </div>
      <Toggle checked={doc.deleted} onChange={(v) => onChange({ ...doc, deleted: v })} label="Deleted" />
    </div>
  );
}

function RolesTab({ doc, onChange }) {
  return (
    <div className="max-w-2xl space-y-5">
      <div>
        <Label>Creators — roles allowed to raise this ticket</Label>
        <TagInput values={doc.creators} onChange={(v) => onChange({ ...doc, creators: v })} placeholder="e.g. merchandiseanalyst — press enter" tone="teal" />
      </div>
      <div>
        <Label>Viewers — roles allowed to view this ticket</Label>
        <TagInput values={doc.viewers} onChange={(v) => onChange({ ...doc, viewers: v })} placeholder="e.g. vendorRole — press enter" tone="teal" />
      </div>
      <div>
        <Label>Assignees — roles this ticket can be routed to</Label>
        <TagInput values={doc.assignee} onChange={(v) => onChange({ ...doc, assignee: v })} placeholder="e.g. Dist_Mgr — press enter" tone="teal" />
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* JSON tab                                                             */
/* ------------------------------------------------------------------ */

function JsonTab({ doc }) {
  const [copied, setCopied] = useState(false);
  const json = useMemo(() => JSON.stringify(docToJson(doc) || {}, null, 2), [doc]);

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(json);
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    } catch (e) { /* clipboard unavailable */ }
  };

  const download = () => {
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${doc.ticketKey || "ticket-type"}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <p className="text-sm text-slate-500">Fields left blank are automatically left out of the generated JSON.</p>
        <div className="flex gap-2">
          <Btn variant="ghost" small onClick={copy}>
            {copied ? <Check size={14} /> : <Copy size={14} />} {copied ? "Copied" : "Copy"}
          </Btn>
          <Btn variant="accent" small onClick={download}><Download size={14} />Download JSON</Btn>
        </div>
      </div>
      <pre className="rounded-xl bg-[#080B12] border border-slate-800 text-slate-200 text-xs leading-relaxed p-4 overflow-auto font-mono max-h-[70vh]">
        {json}
      </pre>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Library view — search across every imported ticket type, open one   */
/* to edit it. This is the landing screen whenever nothing is open.    */
/* ------------------------------------------------------------------ */

function LibraryView({ library, search, onSearchChange, onOpen, onDelete, onNew }) {
  const q = search.trim().toLowerCase();
  const filtered = q
    ? library.filter((e) => (e.doc.ticketType || "").toLowerCase().includes(q) || (e.doc.ticketKey || "").toLowerCase().includes(q))
    : library;

  return (
    <div className="max-w-3xl">
      <div className="flex items-center gap-3 mb-5">
        <div className="relative flex-1">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 pointer-events-none" />
          <input
            value={search}
            onChange={(e) => onSearchChange(e.target.value)}
            placeholder="Search by ticket type name or key…"
            className="w-full rounded-lg border border-slate-700 bg-[#0B101B] pl-9 pr-3 py-2.5 text-sm text-slate-100 placeholder:text-slate-600 outline-none transition focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20"
          />
        </div>
        <Btn variant="dashed" onClick={onNew}><Plus size={15} />New ticket type</Btn>
      </div>

      {!library.length && (
        <div className="rounded-xl border border-dashed border-slate-700 py-14 text-center text-sm text-slate-500">
          No ticket types yet — import a JSON export or Excel workbook, or start a new one.
        </div>
      )}

      {library.length > 0 && !filtered.length && (
        <div className="rounded-xl border border-dashed border-slate-700 py-14 text-center text-sm text-slate-500">
          Nothing matches "{search}".
        </div>
      )}

      <div className="space-y-2">
        {filtered.map((e) => (
          <div
            key={e.id}
            onClick={() => onOpen(e.id)}
            className="group flex items-center justify-between gap-3 rounded-xl border border-slate-800 bg-[#121826] hover:border-teal-500/40 px-4 py-3 cursor-pointer transition"
          >
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <span className="text-sm font-medium text-slate-100 truncate">{e.doc.ticketType || "(untitled ticket type)"}</span>
                {e.doc.deleted && (
                  <span className="text-[10px] uppercase tracking-wide rounded px-1.5 py-0.5 bg-rose-500/10 text-rose-300 shrink-0">Deleted</span>
                )}
              </div>
              <div className="text-xs font-mono text-slate-500 truncate">
                {e.doc.ticketKey || "no key"} · {e.doc.ticketSubType.length} sub-type{e.doc.ticketSubType.length === 1 ? "" : "s"}
              </div>
            </div>
            <IconBtn danger title="Remove from library" onClick={(ev) => { ev.stopPropagation(); onDelete(e); }}>
              <Trash2 size={14} />
            </IconBtn>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* App                                                                  */
/* ------------------------------------------------------------------ */

const SAMPLE = () => ({
  ticketKey: "DSHHPLv1",
  ticketType: "DS Household Plano Stock Request-v1",
  tenantId: "abc",
  deleted: false,
  creators: ["merchandiseanalyst", "corpadmin", "gmoperation", "gmretail", "Dist_Mgr", "Str_Mgr"],
  viewers: ["vendorRole", "corpadmin", "Str_Mgr", "Dist_Mgr"],
  assignee: ["Dist_Mgr", "Str_Mgr"],
  ticketSubType: [],
});

export default function App({ initialData } = {}) {
  const [library, setLibrary] = useState(() => {
    if (!initialData) return [];
    return parseTicketTypesJson(initialData).map((doc) => ({ id: uid(), doc }));
  });
  const [selectedId, setSelectedId] = useState(null);
  const [tab, setTab] = useState("subtypes");
  const [librarySearch, setLibrarySearch] = useState("");
  const [subtypeSearch, setSubtypeSearch] = useState("");
  const [importNotice, setImportNotice] = useState(null); // { warnings, missingSheets, summary } | null
  const fileInputRef = React.useRef(null);
  const excelInputRef = React.useRef(null);

  const selected = library.find((e) => e.id === selectedId) || null;
  const doc = selected ? selected.doc : null;

  const updateDoc = (patch) => setLibrary((lib) => lib.map((e) => (e.id === selectedId ? { ...e, doc: patch } : e)));

  const openEntry = (id) => { setSelectedId(id); setTab("subtypes"); setSubtypeSearch(""); };
  const backToLibrary = () => setSelectedId(null);

  const createNew = () => {
    const entry = { id: uid(), doc: emptyDoc() };
    setLibrary((lib) => [...lib, entry]);
    setSelectedId(entry.id);
    setTab("details");
  };

  const deleteEntry = (e) => {
    const label = e.doc.ticketType || e.doc.ticketKey || "this ticket type";
    if (window.confirm(`Remove "${label}" from the library? This can't be undone.`)) {
      setLibrary((lib) => lib.filter((x) => x.id !== e.id));
      if (selectedId === e.id) setSelectedId(null);
    }
  };

  const deleteSubType = (s) => {
    const label = s.ticketType || s.ticketKey || "this sub-type";
    if (window.confirm(`Delete sub-type "${label}"? This can't be undone.`)) {
      updateDoc({ ...doc, ticketSubType: doc.ticketSubType.filter((x) => x.id !== s.id) });
    }
  };

  const handleImportFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        const incomingDocs = parseTicketTypesJson(parsed);
        if (!incomingDocs.length) throw new Error("No ticket types found");
        const { library: merged, added, updated } = upsertLibraryEntries(library, incomingDocs);
        setLibrary(merged);
        setImportNotice({
          warnings: [],
          missingSheets: [],
          summary: `Imported JSON: ${incomingDocs.length} ticket type${incomingDocs.length === 1 ? "" : "s"} processed — ${added} new, ${updated} updated. Search below to open one.`,
        });
        setSelectedId(null);
        setLibrarySearch("");
      } catch (err) {
        window.alert("That file isn't valid JSON — double check it and try again.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const handleImportExcel = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = XLSX.read(reader.result, { type: "array" });
        const { doc: parsedDoc, warnings, missingSheets } = excelWorkbookToDoc(wb);
        const { library: merged, added, updated, touchedIds } = upsertLibraryEntries(library, [parsedDoc]);
        setLibrary(merged);
        setImportNotice({
          warnings,
          missingSheets,
          summary: `Imported Excel: ${added} new sub-type${added === 1 ? "" : "s"} added, ${updated} existing sub-type${updated === 1 ? "" : "s"} updated.`,
        });
        setSelectedId(touchedIds[0] || null);
        setTab("subtypes");
      } catch (err) {
        window.alert("Couldn't read that workbook — check it matches the expected sheet structure and try again.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const editorNav = doc ? [
    { id: "details", label: "Details", icon: <Ticket size={16} /> },
    { id: "roles", label: "Roles", icon: <Users size={16} /> },
    { id: "subtypes", label: `Sub-Types (${doc.ticketSubType.length})`, icon: <ListTree size={16} /> },
    { id: "json", label: "JSON", icon: <FileJson size={16} /> },
  ] : [];

  const loadSample = () => {
    const { library: merged, touchedIds } = upsertLibraryEntries(library, [SAMPLE()]);
    setLibrary(merged);
    setSelectedId(touchedIds[0] || null);
    setTab("details");
  };

  const clearLibrary = () => {
    if (window.confirm("Remove every ticket type from the library? This can't be undone.")) {
      setLibrary([]);
      setSelectedId(null);
      setImportNotice(null);
    }
  };

  const filteredSubtypes = doc ? (() => {
    const q = subtypeSearch.trim().toLowerCase();
    if (!q) return doc.ticketSubType;
    return doc.ticketSubType.filter((s) => (s.ticketType || "").toLowerCase().includes(q) || (s.ticketKey || "").toLowerCase().includes(q));
  })() : [];

  return (
    <div className="min-h-screen bg-[#080B12] flex text-slate-200" style={{ fontFamily: "Inter, ui-sans-serif, system-ui" }}>
      {/* sidebar */}
      <aside className="w-60 shrink-0 bg-[#0C1120] border-r border-slate-800/80 text-slate-300 flex flex-col">
        <div className="px-5 py-5 border-b border-slate-800/80">
          <div className="flex items-center gap-2 text-white font-semibold">
            <Workflow size={18} className="text-teal-400" />
            Schema Builder
          </div>
          <div className="text-[11px] text-slate-500 mt-0.5">{library.length} ticket type{library.length === 1 ? "" : "s"} in library</div>
        </div>
        <nav className="flex-1 px-2 py-4 space-y-1">
          <button
            onClick={backToLibrary}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm transition ${!selected ? "bg-teal-500/10 text-teal-300" : "hover:bg-white/5 text-slate-400"}`}
          >
            <Search size={16} />
            All Ticket Types
          </button>
          {selected && (
            <div className="mt-2 pt-2 border-t border-slate-800/60 space-y-1">
              {editorNav.map((n) => (
                <button
                  key={n.id}
                  onClick={() => setTab(n.id)}
                  className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm transition ${tab === n.id ? "bg-white/10 text-white" : "hover:bg-white/5 text-slate-400"}`}
                >
                  {n.icon}
                  {n.label}
                </button>
              ))}
            </div>
          )}
        </nav>
        <div className="px-3 py-4 border-t border-slate-800/80 space-y-1.5">
          <button onClick={loadSample} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-xs text-slate-400 hover:bg-white/5 hover:text-white transition">
            <RefreshCw size={13} /> Load sample data
          </button>
          <button onClick={clearLibrary} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-xs text-slate-400 hover:bg-white/5 hover:text-rose-300 transition">
            <Trash2 size={13} /> Clear library
          </button>
        </div>
      </aside>

      {/* main */}
      <main className="flex-1 min-w-0">
        <header className="bg-[#0C1120] border-b border-slate-800/80 px-8 py-4 flex items-center justify-between gap-4">
          <div className="min-w-0 flex items-center gap-3">
            {selected && (
              <button onClick={backToLibrary} title="Back to all ticket types" className="shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg text-slate-400 hover:bg-white/5 hover:text-slate-200 transition">
                <ArrowLeft size={16} />
              </button>
            )}
            <div className="min-w-0">
              <h1 className="text-base font-semibold text-slate-100 truncate">
                {selected ? (doc.ticketType || "Untitled ticket type") : "Ticket Types"}
              </h1>
              <p className="text-xs font-mono text-slate-500">
                {selected ? (doc.ticketKey || "no key set yet") : `${library.length} imported`}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <input ref={excelInputRef} type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={handleImportExcel} className="hidden" />
            <Btn variant="ghost" small onClick={() => excelInputRef.current && excelInputRef.current.click()}>
              <FileSpreadsheet size={14} />Import Excel
            </Btn>
            <input ref={fileInputRef} type="file" accept="application/json,.json" onChange={handleImportFile} className="hidden" />
            <Btn variant="ghost" small onClick={() => fileInputRef.current && fileInputRef.current.click()}>
              <FileJson size={14} />Import JSON
            </Btn>
            {selected ? (
              <Btn variant="accent" onClick={() => setTab("json")}><FileJson size={15} />Create JSON</Btn>
            ) : (
              <Btn variant="dashed" onClick={createNew}><Plus size={15} />New ticket type</Btn>
            )}
          </div>
        </header>

        {importNotice && (
          <div className="mx-8 mt-5 rounded-xl border border-amber-500/30 bg-amber-500/[0.06] px-4 py-3">
            <div className="flex items-start justify-between gap-3">
              <div className="flex items-start gap-2">
                <AlertTriangle size={16} className="text-amber-400 mt-0.5 shrink-0" />
                <div className="text-sm text-amber-200">
                  {importNotice.summary && <p className="font-medium mb-1">{importNotice.summary}</p>}
                  {(importNotice.warnings.length > 0 || importNotice.missingSheets.length > 0) && (
                    <ul className="space-y-0.5 text-xs list-disc pl-4 text-amber-200/90">
                      {importNotice.missingSheets.map((s, i) => (
                        <li key={`m${i}`}>Sheet "{s}" wasn't found — that section is empty.</li>
                      ))}
                      {importNotice.warnings.map((w, i) => (
                        <li key={`w${i}`}>{w}</li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>
              <button onClick={() => setImportNotice(null)} className="text-amber-400/70 hover:text-amber-300 shrink-0"><X size={16} /></button>
            </div>
          </div>
        )}

        <div className="px-8 py-6 max-w-4xl">
          {!selected && (
            <LibraryView
              library={library}
              search={librarySearch}
              onSearchChange={setLibrarySearch}
              onOpen={openEntry}
              onDelete={deleteEntry}
              onNew={createNew}
            />
          )}
          {selected && tab === "details" && <DetailsTab doc={doc} onChange={updateDoc} />}
          {selected && tab === "roles" && <RolesTab doc={doc} onChange={updateDoc} />}
          {selected && tab === "subtypes" && (
            <div>
              {doc.ticketSubType.length > 0 && (
                <div className="relative mb-4">
                  <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 pointer-events-none" />
                  <input
                    value={subtypeSearch}
                    onChange={(e) => setSubtypeSearch(e.target.value)}
                    placeholder="Filter sub-types by name or key…"
                    className="w-full rounded-lg border border-slate-700 bg-[#0B101B] pl-9 pr-3 py-2 text-sm text-slate-100 placeholder:text-slate-600 outline-none transition focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20"
                  />
                </div>
              )}
              {filteredSubtypes.map((s) => (
                <SubTypeEditor
                  key={s.id}
                  sub={s}
                  onChange={(patch) => updateDoc({ ...doc, ticketSubType: doc.ticketSubType.map((x) => (x.id === s.id ? patch : x)) })}
                  onDelete={() => deleteSubType(s)}
                />
              ))}
              {subtypeSearch.trim() && !filteredSubtypes.length && (
                <p className="text-sm text-slate-500 italic mb-4">No sub-types match "{subtypeSearch}".</p>
              )}
              <Btn variant="dashed" onClick={() => updateDoc({ ...doc, ticketSubType: [...doc.ticketSubType, newSubType()] })}>
                <Plus size={15} />Add sub-ticket type
              </Btn>
            </div>
          )}
          {selected && tab === "json" && <JsonTab doc={doc} />}
        </div>
      </main>
    </div>
  );
}
