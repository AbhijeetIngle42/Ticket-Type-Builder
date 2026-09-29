import React, { useState, useMemo, useCallback } from "react";
import {
  Plus, Trash2, ChevronDown, ChevronRight, Copy, Download,
  FileJson, Users, Ticket, ListTree, Workflow, Check, X,
  GripVertical, AlertTriangle, RefreshCw, FileSpreadsheet,
  Sparkles, LayoutTemplate, Star, Wand2, Diff
} from "lucide-react";
import * as XLSX from "xlsx-js-style";

/* ------------------------------------------------------------------ */
/* utils                                                               */
/* ------------------------------------------------------------------ */

let uidCounter = 0;
const uid = () => `id_${Date.now().toString(36)}_${(uidCounter++).toString(36)}`;

const FIELD_TYPES = [
  { value: "dropdown", label: "Dropdown" },
  { value: "radio", label: "Radio" },
  { value: "multiplechoice", label: "Multiple Choice" },
  { value: "casecadeDropdown", label: "Cascading Dropdown" },
  { value: "apiDropdown", label: "API Dropdown (advanced)" },
  { value: "matrix", label: "Matrix / Table" },
  { value: "date", label: "Date" },
  { value: "time", label: "Time" },
  { value: "number", label: "Number" },
  { value: "text", label: "Text" },
  { value: "fileUpload", label: "File Upload" },
];

// field types whose "options" list can reveal other fields conditionally
const OPTION_BASED_TYPES = ["dropdown", "radio", "multiplechoice"];

const NOTIF_TYPES = ["inApp", "email"];

const advancedDefaults = () => ({
  isMandatory: false, disabled: false, displayLabel: "", dependentOn: "", editRoles: [],
});

const newDropdownField = (type = "dropdown") => ({
  id: uid(), label: "", key: "", type, options: [], ...advancedDefaults(),
});
const newCascadeField = () => ({
  id: uid(), label: "", key: "", type: "casecadeDropdown",
  childLabel: "", childKey: "", parentOptionType: "multiplechoice", options: [], ...advancedDefaults(),
});
const newApiDropdownField = () => ({
  id: uid(), label: "", key: "", type: "apiDropdown",
  labelKey: "", valueKey: "", fillOnSelect: false, apiFiltersText: "{\n\n}", ...advancedDefaults(),
});
const newMatrixField = () => ({
  id: uid(), label: "", key: "", type: "matrix",
  columns: [], hideTotalColumn: false, ...advancedDefaults(),
});
const newSimpleField = (type) => ({ id: uid(), label: "", key: "", type, ...advancedDefaults() });

const newField = (type) => {
  if (OPTION_BASED_TYPES.includes(type)) return newDropdownField(type);
  if (type === "casecadeDropdown") return newCascadeField();
  if (type === "apiDropdown") return newApiDropdownField();
  if (type === "matrix") return newMatrixField();
  return newSimpleField(type);
};

const newParentOption = () => ({ id: uid(), name: "", key: "", children: [] });
const newChildOption = () => ({ id: uid(), value: "", key: "" });
const newSimpleOption = () => ({ id: uid(), value: "", key: "", dependentCustomFields: [] });
const newMatrixColumn = () => ({ id: uid(), label: "", key: "", type: "text", options: [], formula: "" });

const newAutoEscalation = () => ({
  id: uid(), assignee: "", days: "", overrideStatus: "",
  notification: true, notificationType: [], ccAddress: [],
});

const newStatus = () => ({
  id: uid(), status: "", label: "", public: false,
  roles: [], assigneeRoles: [], notification: false, notificationType: [],
  notificationRole: [], mandatoryCustomFields: [], overrideStatus: "",
  dependentStatus: [], customFieldsMetaData: [], autoEscalationConfig: [],
});

const newSubType = () => ({
  id: uid(), ticketKey: "", prefix: "", ticketType: "",
  deleted: false, isAutoEscalation: false, assigneeRoles: [],
  customFieldsMetaData: [], statusWorkFlow: [], _passthrough: {},
});

const emptyDoc = () => ({
  ticketKey: "", ticketType: "", tenantId: "", deleted: false,
  creators: [], viewers: [], assignee: [],
  ticketSubType: [], _passthrough: {},
});

/* Keys the editor models explicitly at the doc / sub-type level. Anything
   else found on import (_id, __v, createdAt, updatedAt, backend-specific
   metadata, etc.) isn't understood by this UI, but it must not be thrown
   away — it's captured into `_passthrough` on import and merged back into
   the exported JSON untouched, so round-tripping through this app never
   silently deletes fields it doesn't have a form for. */
const DOC_KNOWN_KEYS = new Set(["ticketKey", "ticketType", "tenantId", "creators", "viewers", "assignee", "deleted", "ticketSubType"]);
const SUBTYPE_KNOWN_KEYS = new Set(["ticketKey", "prefix", "ticketType", "customFieldsMetaData", "statusWorkFlow", "deleted", "isAutoEscalation", "assignee"]);

function extractPassthrough(src, knownKeys) {
  const passthrough = {};
  Object.keys(src || {}).forEach((k) => { if (!knownKeys.has(k)) passthrough[k] = src[k]; });
  return passthrough;
}

/* -- sub-type templates: a reusable "shape" (fields + status workflow)
   that bulk sub-type creation stamps out once per name. Only the
   ticketKey/ticketType are unique per generated sub-type; everything
   else is cloned from the template with brand-new ids. -- */
const newTemplate = (name = "New template") => ({
  id: uid(), name, description: "", isDefault: false,
  data: { prefix: "", isAutoEscalation: false, customFieldsMetaData: [], statusWorkFlow: [] },
});

/* deep-clone a template/field/status tree and assign a fresh id to every
   node that has one, so a generated sub-type never shares option/field/
   status ids with the template or with sibling sub-types made from it. */
function regenerateIds(node) {
  if (Array.isArray(node)) return node.map(regenerateIds);
  if (node && typeof node === "object") {
    const out = {};
    Object.keys(node).forEach((k) => { out[k] = regenerateIds(node[k]); });
    if (Object.prototype.hasOwnProperty.call(out, "id")) out.id = uid();
    return out;
  }
  return node;
}

/* deep-clean: drop empty strings / empty arrays / empty objects,
   keep booleans and numbers (incl. 0/false) as explicitly provided */
function clean(value) {
  if (Array.isArray(value)) {
    // Keep arrays even when empty — "no custom fields yet" / "no CC addresses"
    // is a real, meaningful value that was present in the source JSON, not
    // something to silently drop. Only genuinely undefined elements are
    // filtered out; the array itself always survives.
    return value.map(clean).filter((v) => v !== undefined);
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

function advancedToJson(f) {
  const out = {};
  if (f.isMandatory) out.isMandatory = true;
  if (f.disabled) out.disabled = true;
  if (f.displayLabel) out.displayLabel = f.displayLabel;
  if (f.dependentOn) out.dependentOn = f.dependentOn;
  if (f.editRoles && f.editRoles.length) out.roles = f.editRoles;
  return out;
}

function optionToJson(o) {
  const out = { value: o.value, key: o.key };
  if (o.dependentCustomFields && o.dependentCustomFields.length) out.dependentCustomFields = o.dependentCustomFields;
  return out;
}

function fieldToJson(f) {
  const adv = advancedToJson(f);
  if (OPTION_BASED_TYPES.includes(f.type)) {
    return { label: f.label, key: f.key, type: f.type, options: (f.options || []).map(optionToJson), ...adv };
  }
  if (f.type === "casecadeDropdown") {
    return {
      label: f.label, key: f.key, type: f.type,
      childLabel: f.childLabel, childKey: f.childKey,
      options: (f.options || []).map((p) => ({
        value: p.name, label: p.name, type: f.parentOptionType, key: p.key,
        options: (p.children || []).map((c) => ({ value: c.value, key: c.key })),
      })),
      ...adv,
    };
  }
  if (f.type === "apiDropdown") {
    let apiFilters;
    try { apiFilters = JSON.parse(f.apiFiltersText || "{}"); } catch (e) { apiFilters = undefined; }
    return {
      label: f.label, key: f.key, type: f.type,
      labelKey: f.labelKey, valueKey: f.valueKey, fillOnSelect: !!f.fillOnSelect,
      apiFilters, ...adv,
    };
  }
  if (f.type === "matrix") {
    return {
      label: f.label, key: f.key, type: f.type,
      hideTotalColumn: !!f.hideTotalColumn,
      row: (f.columns || []).map((c) => {
        const col = { label: c.label, key: c.key, type: c.type };
        if (OPTION_BASED_TYPES.includes(c.type)) col.options = (c.options || []).map(optionToJson);
        if (c.formula) col.formula = c.formula;
        return col;
      }),
      ...adv,
    };
  }
  return { label: f.label, key: f.key, type: f.type, ...adv };
}

function statusToJson(s) {
  return {
    status: s.status, label: s.label, public: s.public,
    roles: s.roles, assignee: (s.assigneeRoles || []).map((r) => ({ roleId: r })),
    notification: s.notification, notificationType: s.notificationType,
    notificationRole: s.notificationRole,
    mandatoryCustomFields: s.mandatoryCustomFields,
    overrideStatus: s.overrideStatus,
    dependentStatus: s.dependentStatus,
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
    ...(st._passthrough || {}),
    ticketKey: st.ticketKey, prefix: st.prefix, ticketType: st.ticketType,
    assignee: (st.assigneeRoles || []).map((r) => ({ roleId: r })),
    customFieldsMetaData: (st.customFieldsMetaData || []).map(fieldToJson),
    statusWorkFlow: (st.statusWorkFlow || []).map(statusToJson),
    deleted: st.deleted, isAutoEscalation: st.isAutoEscalation,
  };
}

function docToJson(doc) {
  const raw = {
    ...(doc._passthrough || {}),
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

function advancedFromJson(f) {
  return {
    isMandatory: !!f.isMandatory, disabled: !!f.disabled,
    displayLabel: f.displayLabel || "", dependentOn: f.dependentOn || "",
    editRoles: Array.isArray(f.roles) ? f.roles : [],
  };
}

function optionFromJson(o) {
  return {
    id: uid(), value: o.value || "", key: o.key != null ? String(o.key) : "",
    dependentCustomFields: Array.isArray(o.dependentCustomFields) ? o.dependentCustomFields : [],
  };
}

function fieldFromJson(f) {
  if (!f) return newField("text");
  const adv = advancedFromJson(f);
  if (OPTION_BASED_TYPES.includes(f.type)) {
    return {
      id: uid(), label: f.label || "", key: f.key || "", type: f.type,
      options: (f.options || []).map(optionFromJson), ...adv,
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
      ...adv,
    };
  }
  if (f.type === "apiDropdown") {
    return {
      id: uid(), label: f.label || "", key: f.key || "", type: "apiDropdown",
      labelKey: f.labelKey || "", valueKey: f.valueKey || "", fillOnSelect: !!f.fillOnSelect,
      apiFiltersText: f.apiFilters ? JSON.stringify(f.apiFilters, null, 2) : "{\n\n}", ...adv,
    };
  }
  if (f.type === "matrix") {
    return {
      id: uid(), label: f.label || "", key: f.key || "", type: "matrix",
      hideTotalColumn: !!f.hideTotalColumn,
      columns: (f.row || []).map((c) => ({
        id: uid(), label: c.label || "", key: c.key || "", type: c.type || "text",
        options: (c.options || []).map(optionFromJson), formula: c.formula || "",
      })),
      ...adv,
    };
  }
  return { id: uid(), label: f.label || "", key: f.key || "", type: f.type || "text", ...adv };
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
    dependentStatus: Array.isArray(s.dependentStatus) ? s.dependentStatus : [],
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
    assigneeRoles: rolesFromJson(st.assignee),
    customFieldsMetaData: (st.customFieldsMetaData || []).map(fieldFromJson),
    statusWorkFlow: (st.statusWorkFlow || []).map(statusFromJson),
    _passthrough: extractPassthrough(st, SUBTYPE_KNOWN_KEYS),
  };
}

function jsonToDoc(raw) {
  const src = Array.isArray(raw) ? raw[0] : raw;
  if (!src || typeof src !== "object") return emptyDoc();
  return {
    ticketKey: src.ticketKey || "", ticketType: src.ticketType || "", tenantId: src.tenantId || "",
    deleted: !!src.deleted,
    creators: rolesFromJson(src.creators),
    viewers: rolesFromJson(src.viewers),
    assignee: rolesFromJson(src.assignee),
    ticketSubType: (src.ticketSubType || []).map(subTypeFromJson),
    _passthrough: extractPassthrough(src, DOC_KNOWN_KEYS),
  };
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
// PascalCase key derived from a free-text name, e.g. "DS Stock gaps - Store 101" -> "DSStockGapsStore101"
const xToPascalKey = (label) => {
  const words = xnorm(label).match(/[A-Za-z0-9]+/g) || [];
  if (!words.length) return "Item";
  return words.map((w) => w[0].toUpperCase() + w.slice(1)).join("");
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

/* Returns an array — one entry per row — so a single "Ticket Details"
   sheet can describe several ticket types at once. Each row is keyed by
   its own Ticket Key, which every other sheet also references. */
function xParseTicketDetails(sheet) {
  const rows = xRowsOf(sheet);
  if (rows.length < 2) return [];
  const header = rows[0].map((h) => xnorm(h).toLowerCase());
  const col = (n) => xFindCol(header, n);
  return rows.slice(1).map((r) => ({
    ticketKey: xnorm(xCell(r, col("ticket key"))),
    ticketType: xnorm(xCell(r, col("ticket type"))),
    tenantId: xnorm(xCell(r, col("tenant"))),
    creators: xSplitRoles(xCell(r, col("creator"))),
    viewers: xSplitRoles(xCell(r, col("viewer"))),
    assignee: xSplitRoles(xCell(r, col("assignee"))),
    deleted: xIsYes(xCell(r, col("delete"))),
  })).filter((d) => d.ticketKey || d.ticketType);
}

/* Groups rows by the sheet's "Ticket Key" column (the ticket TYPE's key,
   not the subtype key) so a single "Sub Types" sheet can describe the
   sub-types of several ticket types at once. Returns per-ticket
   { subtypes, order } maps keyed by that Ticket Key. Rows with no
   Ticket Key fall under the "__default__" bucket, which lets a
   single-ticket-type sheet (no Ticket Key column data) still work. */
function xParseSubTypes(sheet) {
  const rows = xRowsOf(sheet);
  const byTicket = {};
  if (!rows.length) return byTicket;
  const header = rows[0].map((h) => xnorm(h).toLowerCase());
  const col = (n) => xFindCol(header, n);
  const hasTicketCol = col("ticket key") !== -1;
  for (const r of rows.slice(1)) {
    const name = xnorm(xCell(r, col("ticket subtype")));
    if (!name) continue;
    const ticketKey = hasTicketCol ? (xnorm(xCell(r, col("ticket key"))) || "__default__") : "__default__";
    const key = xnorm(xCell(r, col("subtype key"))) || name;
    byTicket[ticketKey] = byTicket[ticketKey] || { subtypes: {}, order: [] };
    byTicket[ticketKey].subtypes[name] = {
      ticketKey: key,
      prefix: xnorm(xCell(r, col("prefix"))),
      assigneeRoles: xSplitRoles(xCell(r, col("assignee"))),
      isAutoEscalation: xIsYes(xCell(r, col("auto escalation"))),
      deleted: xIsYes(xCell(r, col("delete"))),
      ticketType: name,
      customFieldsMetaData: [],
      statusWorkFlow: [],
    };
    byTicket[ticketKey].order.push(name);
  }
  return byTicket;
}

/* `subtypesByTicket` is the per-ticket map returned by xParseSubTypes.
   When the sheet has a leading "Ticket Key" header, every column shifts
   right by one and each row is routed to that ticket's subtype bucket;
   otherwise everything lands under "__default__" (single-ticket sheet). */
function xParseTicketStructure(sheet, subtypesByTicket) {
  const rows = xRowsOf(sheet);
  if (!rows.length) return;
  const headerRow0 = xnorm(xCell(rows[0], 0)).toLowerCase();
  const offset = headerRow0.includes("ticket key") ? 1 : 0;
  let currentTicketKey = "__default__";
  let currentSubtype = null;
  let currentField = null;
  const optionBased = ["dropdown", "radio", "multiplechoice"];
  const touched = new Set();
  for (const r of rows.slice(1)) {
    const ticketKeyCell = offset ? xnorm(xCell(r, 0)) : "";
    const subtypeCell = xnorm(xCell(r, offset + 0));
    const labelCell = xnorm(xCell(r, offset + 1));
    const typeCell = xnorm(xCell(r, offset + 2));
    const mandatoryCell = xnorm(xCell(r, offset + 3));
    const rolesCell = xnorm(xCell(r, offset + 4));

    if (ticketKeyCell) currentTicketKey = ticketKeyCell;
    if (subtypeCell) currentSubtype = subtypeCell;
    touched.add(currentTicketKey);

    const bucket = subtypesByTicket[currentTicketKey];

    if (labelCell) {
      const typeMatch = labelCell.match(/\((.*?)\)/);
      const explicitType = typeMatch ? typeMatch[1].trim().toLowerCase() : "";
      const cleanLabel = labelCell.replace(/\(.*?\)/g, "").trim();
      let ftype = "dropdown";
      if (explicitType.includes("radio")) ftype = "radio";
      else if (explicitType.includes("multi")) ftype = "multiplechoice";
      else if (explicitType.includes("drop")) ftype = "dropdown";
      else if (explicitType.includes("date")) ftype = "date";
      else if (explicitType.includes("time")) ftype = "time";
      else if (explicitType.includes("number")) ftype = "number";
      else if (explicitType.includes("file")) ftype = "fileUpload";
      else if (explicitType.includes("text")) ftype = "text";

      currentField = { label: cleanLabel, key: xToCamelKey(cleanLabel), type: ftype, editRoles: xSplitRoles(rolesCell), isMandatory: xIsYes(mandatoryCell), options: [] };
      if (typeCell && optionBased.includes(ftype)) currentField.options.push({ value: typeCell, key: "" });
      if (currentSubtype && bucket && bucket.subtypes[currentSubtype]) bucket.subtypes[currentSubtype].customFieldsMetaData.push(currentField);
      continue;
    }
    if (currentField && typeCell) currentField.options.push({ value: typeCell, key: "" });
  }
  touched.forEach((tk) => {
    const bucket = subtypesByTicket[tk];
    if (!bucket) return;
    Object.values(bucket.subtypes).forEach((st) => {
      st.customFieldsMetaData.forEach((f) => {
        if (optionBased.includes(f.type) && !f.options.length) f.type = "text";
        f.options.forEach((o, i) => { if (!o.key) o.key = String(i + 1); });
      });
    });
  });
}

/* Optional leading "Ticket Key" column groups rows per ticket type;
   without it, everything lands under "__default__". Returns a map of
   ticketKey -> { statusCode -> { label, roles } }. */
function xParseStatusAccessRoles(sheet) {
  const rows = xRowsOf(sheet);
  const byTicket = {};
  if (!rows.length) return byTicket;
  let headerIdx = rows.findIndex((r) => xnorm(xCell(r, 0)).toLowerCase().includes("status") || xnorm(xCell(r, 1)).toLowerCase().includes("status"));
  if (headerIdx === -1) return byTicket;
  const headerRow = rows[headerIdx].map((h) => xnorm(h).toLowerCase());
  const offset = headerRow[0] && headerRow[0].includes("ticket key") ? 1 : 0;
  for (const r of rows.slice(headerIdx + 1)) {
    const label = xnorm(xCell(r, offset + 0));
    if (!label) continue;
    const ticketKey = offset ? (xnorm(xCell(r, 0)) || "__default__") : "__default__";
    byTicket[ticketKey] = byTicket[ticketKey] || {};
    byTicket[ticketKey][xToStatusCode(label)] = { label, roles: xSplitRoles(xCell(r, offset + 1)) };
  }
  return byTicket;
}

/* Same optional leading "Ticket Key" column convention as the roles
   sheet. Returns ticketKey -> { statusCode -> mandate rule }. */
function xParseMandateRules(sheet) {
  const rows = xRowsOf(sheet);
  const byTicket = {};
  if (!rows.length) return byTicket;
  let headerIdx = rows.findIndex((r) => xnorm(xCell(r, 0)).toLowerCase().includes("status") || xnorm(xCell(r, 1)).toLowerCase().includes("status"));
  if (headerIdx === -1) return byTicket;
  const headerRow = rows[headerIdx].map((h) => xnorm(h).toLowerCase());
  const offset = headerRow[0] && headerRow[0].includes("ticket key") ? 1 : 0;
  for (const r of rows.slice(headerIdx + 1)) {
    const label = xnorm(xCell(r, offset + 0));
    if (!label) continue;
    const ticketKey = offset ? (xnorm(xCell(r, 0)) || "__default__") : "__default__";
    const imagesCell = xnorm(xCell(r, offset + 1));
    const commentsCell = xnorm(xCell(r, offset + 2));
    const assigneeYnCell = xnorm(xCell(r, offset + 3));
    const assigneeRoleCell = xnorm(xCell(r, offset + 4));
    let commentsLabel = commentsCell && !xIsYes(commentsCell) ? commentsCell : "Comments";
    commentsLabel = commentsLabel.replace(/\(mandatory\)/i, "").trim();
    byTicket[ticketKey] = byTicket[ticketKey] || {};
    byTicket[ticketKey][xToStatusCode(label)] = {
      mandatoryImages: xIsYes(imagesCell),
      mandatoryComments: xIsYes(commentsCell) || !!commentsCell,
      commentsLabel,
      mandatoryNewAssignee: xIsYes(assigneeYnCell),
      newAssigneeRole: assigneeRoleCell,
    };
  }
  return byTicket;
}

/* Recognizes an optional "TICKET: <ticketKey>" marker line above the
   existing "STATUS: <label>" blocks, so one sheet can hold escalation
   tables for several ticket types. Sheets with no TICKET markers put
   everything under "__default__" (single-ticket sheet, unchanged
   behavior). Returns { escalationByTicket, warnings } where
   escalationByTicket is ticketKey -> escalationByStatus. */
function xParseEscalationDetails(sheet) {
  const rows = xRowsOf(sheet);
  const escalationByTicket = {};
  const warnings = [];
  let currentTicket = "__default__";
  let currentStatus = null;
  let inTable = false;
  for (const r of rows) {
    const a0 = xnorm(xCell(r, 0));
    const tm = a0.match(/^ticket\s*:\s*(.+)$/i);
    if (tm) {
      currentTicket = tm[1].trim() || "__default__";
      currentStatus = null;
      inTable = false;
      continue;
    }
    const m = a0.match(/^status\s*:\s*(.+)$/i);
    if (m) {
      currentStatus = xToStatusCode(m[1]);
      escalationByTicket[currentTicket] = escalationByTicket[currentTicket] || {};
      escalationByTicket[currentTicket][currentStatus] = escalationByTicket[currentTicket][currentStatus] || {};
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
    const bucket = escalationByTicket[currentTicket][currentStatus][subtype] || [];
    if (bucket.length && chain.length) {
      warnings.push(`Sub-type '${subtype}' has more than one Issue Type row under STATUS:${currentStatus} (latest: '${issueType}') — escalation steps were merged into one chain.`);
    }
    escalationByTicket[currentTicket][currentStatus][subtype] = bucket.concat(chain);
  }
  return { escalationByTicket, warnings };
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
  if (["dropdown", "radio", "multiplechoice"].includes(f.type)) out.options = f.options.map((o) => ({ value: o.value, key: o.key }));
  if (f.isMandatory) out.isMandatory = true;
  if (f.editRoles && f.editRoles.length) out.roles = f.editRoles;
  return out;
}

/* Matrix fields (a field whose "value" is itself a mini-table with its own
   columns) can't fit the flat one-row-per-option grammar the "Ticket
   Structure" sheet uses, so they get a dedicated sheet whose layout
   mirrors what the matrix itself looks like — the same shape someone
   filling this in from scratch, with zero prior context, would
   naturally produce:
     Ticket Key | Ticket Subtype | Mandatory | Edit Roles | Field (Matrix) | Col1 (Type) | Col2 (Type) | ...
     ...                                                                  | option      | option      | ...
     ...                                                                  | option      | option      | ...
     <blank row separates the next field/subtype/ticket>
   Only the first row of each field block carries Ticket Key / Ticket
   Subtype / Mandatory / Edit Roles / the field's own name; every column
   header after it defines one matrix column, and every following row
   supplies one option value per dropdown column, aligned by position —
   exactly like the raw column layout of the matrix table itself.
   Returns ticketKey -> subtypeName -> [fieldDef, ...]. */
const MATRIX_COL_START = 5; // 0-indexed column F: A-D = Ticket Key/Subtype/Mandatory/Edit Roles, E = field name

function xParseMatrixFields(sheet) {
  const rows = xRowsOf(sheet);
  const byTicket = {};
  if (rows.length < 2) return byTicket;

  let currentTicketKey = "__default__";
  let currentSubtype = null;
  let currentFieldMeta = null; // { label, key, isMandatory, editRoles }
  let colDefs = null; // [{ label, key, type, options: [] }, ...] aligned to column index (0 = MATRIX_COL_START)

  const finalizeField = () => {
    if (currentFieldMeta && colDefs && colDefs.length && currentSubtype) {
      byTicket[currentTicketKey] = byTicket[currentTicketKey] || {};
      byTicket[currentTicketKey][currentSubtype] = byTicket[currentTicketKey][currentSubtype] || [];
      byTicket[currentTicketKey][currentSubtype].push({
        label: currentFieldMeta.label,
        key: currentFieldMeta.key,
        isMandatory: currentFieldMeta.isMandatory,
        editRoles: currentFieldMeta.editRoles,
        columns: colDefs.filter(Boolean).map((cd) => ({
          label: cd.label, key: cd.key, type: cd.type,
          options: cd.options.map((v, i) => ({ value: v, key: String(i + 1) })),
        })),
      });
    }
    currentFieldMeta = null;
    colDefs = null;
  };

  for (const r of rows.slice(1)) {
    const isBlankRow = !r.some((c) => c !== null && xnorm(c) !== "");
    if (isBlankRow) { finalizeField(); continue; }

    const fieldCell = xnorm(xCell(r, MATRIX_COL_START - 1)); // column E
    if (fieldCell) {
      // start of a new field block
      finalizeField();
      const tk = xnorm(xCell(r, 0));
      const st = xnorm(xCell(r, 1));
      if (tk) currentTicketKey = tk;
      if (st) currentSubtype = st;
      const mandatoryCell = xnorm(xCell(r, 2));
      const rolesCell = xnorm(xCell(r, 3));
      const cleanLabel = fieldCell.replace(/\(.*?\)/g, "").trim();
      currentFieldMeta = {
        label: cleanLabel,
        key: xToCamelKey(cleanLabel),
        isMandatory: xIsYes(mandatoryCell),
        editRoles: xSplitRoles(rolesCell),
      };
      colDefs = [];
      for (let ci = MATRIX_COL_START; ci < r.length; ci++) {
        const cell = xnorm(xCell(r, ci));
        if (!cell) { colDefs.push(null); continue; }
        const cm = cell.match(/\((.*?)\)/);
        const ctypeRaw = cm ? cm[1].trim().toLowerCase() : "";
        const clabel = cell.replace(/\(.*?\)/g, "").trim();
        let ctype = "text";
        if (ctypeRaw.includes("drop")) ctype = "dropdown";
        else if (ctypeRaw.includes("number")) ctype = "number";
        else if (ctypeRaw.includes("date")) ctype = "date";
        colDefs.push({ label: clabel, key: xToCamelKey(clabel), type: ctype, options: [] });
      }
      continue;
    }

    // option row: values line up with colDefs starting at MATRIX_COL_START
    if (colDefs) {
      for (let i = 0; i < colDefs.length; i++) {
        if (!colDefs[i]) continue;
        const val = xnorm(xCell(r, MATRIX_COL_START + i));
        if (val && colDefs[i].type === "dropdown") colDefs[i].options.push(val);
      }
    }
  }
  finalizeField();
  return byTicket;
}

/* Converts one parsed matrix-field block into the exact JSON shape
   fieldFromJson() expects for type "matrix" (same shape fieldToJson()
   produces), so it slots into customFieldsMetaData with zero special
   handling once merged in. */
function xMatrixFieldToOutput(mf) {
  const out = {
    label: mf.label, key: mf.key, type: "matrix",
    row: mf.columns.map((c) => {
      const col = { label: c.label, key: c.key, type: c.type };
      if (c.type === "dropdown") col.options = c.options;
      return col;
    }),
  };
  if (mf.isMandatory) out.isMandatory = true;
  if (mf.editRoles && mf.editRoles.length) out.roles = mf.editRoles;
  return out;
}

const CASCADE_COL = { ticketKey: 0, subtype: 1, mandatory: 2, roles: 3, field: 4, parentType: 5, childLabel: 6, childKey: 7, parentOption: 8, childOption: 9 };

/* Mirrors xParseMatrixFields' state machine: a field-header row (non-blank
   "Field (Type)" cell) starts a new block; a non-blank "Parent Option" cell
   starts a new parent within that block; a "Child Option"-only row appends
   another child to whichever parent came right before it; a fully blank
   row closes the field. Returns ticketKey -> subtypeName -> [fieldDef,...]. */
function xParseCascadeDropdowns(sheet) {
  const rows = xRowsOf(sheet);
  const byTicket = {};
  if (rows.length < 2) return byTicket;

  let currentTicketKey = "__default__";
  let currentSubtype = null;
  let currentFieldMeta = null;
  let parents = null;

  const finalizeField = () => {
    if (currentFieldMeta && parents && currentSubtype) {
      byTicket[currentTicketKey] = byTicket[currentTicketKey] || {};
      byTicket[currentTicketKey][currentSubtype] = byTicket[currentTicketKey][currentSubtype] || [];
      byTicket[currentTicketKey][currentSubtype].push({ ...currentFieldMeta, parents });
    }
    currentFieldMeta = null;
    parents = null;
  };

  for (const r of rows.slice(1)) {
    const isBlankRow = !r.some((c) => c !== null && xnorm(c) !== "");
    if (isBlankRow) { finalizeField(); continue; }

    const fieldCell = xnorm(xCell(r, CASCADE_COL.field));
    const parentCell = xnorm(xCell(r, CASCADE_COL.parentOption));
    const childCell = xnorm(xCell(r, CASCADE_COL.childOption));

    if (fieldCell) {
      finalizeField();
      const tk = xnorm(xCell(r, CASCADE_COL.ticketKey));
      const st = xnorm(xCell(r, CASCADE_COL.subtype));
      if (tk) currentTicketKey = tk;
      if (st) currentSubtype = st;
      const cleanLabel = fieldCell.replace(/\(.*?\)/g, "").trim();
      const childLabel = xnorm(xCell(r, CASCADE_COL.childLabel));
      currentFieldMeta = {
        label: cleanLabel, key: xToCamelKey(cleanLabel),
        isMandatory: xIsYes(xCell(r, CASCADE_COL.mandatory)),
        editRoles: xSplitRoles(xCell(r, CASCADE_COL.roles)),
        parentOptionType: xnorm(xCell(r, CASCADE_COL.parentType)) || "multiplechoice",
        childLabel, childKey: xnorm(xCell(r, CASCADE_COL.childKey)) || xToCamelKey(childLabel),
      };
      parents = [];
      if (parentCell) {
        const p = { value: parentCell, key: "", children: [] };
        if (childCell) p.children.push({ value: childCell, key: "" });
        parents.push(p);
      }
      continue;
    }

    if (!parents) continue; // stray row before any field header — ignore
    if (parentCell) {
      const p = { value: parentCell, key: "", children: [] };
      if (childCell) p.children.push({ value: childCell, key: "" });
      parents.push(p);
    } else if (childCell && parents.length) {
      parents[parents.length - 1].children.push({ value: childCell, key: "" });
    }
  }
  finalizeField();

  Object.values(byTicket).forEach((bySub) => Object.values(bySub).forEach((fields) => {
    fields.forEach((f) => f.parents.forEach((p, pi) => {
      if (!p.key) p.key = String(pi + 1);
      p.children.forEach((c, ci) => { if (!c.key) c.key = String(ci + 1); });
    }));
  }));

  return byTicket;
}

/* Converts one parsed cascade-field block into the exact JSON shape
   fieldFromJson() expects for type "casecadeDropdown". */
function xCascadeFieldToOutput(cf) {
  const out = {
    label: cf.label, key: cf.key, type: "casecadeDropdown",
    childLabel: cf.childLabel, childKey: cf.childKey,
    options: cf.parents.map((p) => ({
      value: p.value, label: p.value, type: cf.parentOptionType, key: p.key,
      options: p.children.map((c) => ({ value: c.value, key: c.key })),
    })),
  };
  if (cf.isMandatory) out.isMandatory = true;
  if (cf.editRoles && cf.editRoles.length) out.roles = cf.editRoles;
  return out;
}

/* API dropdown rows are one-per-field (no repeating option rows — the
   options are fetched live at runtime), so this is a flat table read like
   Status Access Roles: an optional leading Ticket Key/Ticket Subtype pair
   "sticks" down the sheet until a new one is given. */
function xParseApiDropdowns(sheet) {
  const rows = xRowsOf(sheet);
  const byTicket = {};
  if (rows.length < 2) return byTicket;
  const header = rows[0].map((h) => xnorm(h).toLowerCase());
  const col = (n) => xFindCol(header, n);
  let currentTicketKey = "__default__";
  let currentSubtype = null;
  for (const r of rows.slice(1)) {
    const tk = xnorm(xCell(r, col("ticket key")));
    const st = xnorm(xCell(r, col("ticket subtype")));
    if (tk) currentTicketKey = tk;
    if (st) currentSubtype = st;
    const fieldCell = xnorm(xCell(r, col("field")));
    if (!fieldCell || !currentSubtype) continue;
    const cleanLabel = fieldCell.replace(/\(.*?\)/g, "").trim();
    const field = {
      label: cleanLabel, key: xToCamelKey(cleanLabel),
      isMandatory: xIsYes(xCell(r, col("mandatory"))),
      editRoles: xSplitRoles(xCell(r, col("edit roles"))),
      labelKey: xnorm(xCell(r, col("label key"))),
      valueKey: xnorm(xCell(r, col("value key"))),
      fillOnSelect: xIsYes(xCell(r, col("fill on select"))),
      apiFiltersText: xnorm(xCell(r, col("api filters"))) || "{}",
    };
    byTicket[currentTicketKey] = byTicket[currentTicketKey] || {};
    byTicket[currentTicketKey][currentSubtype] = byTicket[currentTicketKey][currentSubtype] || [];
    byTicket[currentTicketKey][currentSubtype].push(field);
  }
  return byTicket;
}

function xApiFieldToOutput(af) {
  const out = { label: af.label, key: af.key, type: "apiDropdown", labelKey: af.labelKey, valueKey: af.valueKey, fillOnSelect: !!af.fillOnSelect };
  try { out.apiFilters = JSON.parse(af.apiFiltersText || "{}"); } catch (e) { out.apiFilters = {}; }
  if (af.isMandatory) out.isMandatory = true;
  if (af.editRoles && af.editRoles.length) out.roles = af.editRoles;
  return out;
}

/** Parses a workbook (from XLSX.read) built with the 6-sheet convention
 * into one or more raw JSON ticket types (docToJson() shape), then
 * converts each to editable state via jsonToDoc(). Every sheet below
 * "Ticket Details" resolves its rows against a ticket type via an
 * optional leading "Ticket Key" column — sheets without that column
 * are treated as describing a single ticket type ("__default__"),
 * so old single-ticket-type workbooks still import unchanged.
 * Returns { docs, warnings, missingSheets }. */
function excelWorkbookToDocs(wb) {
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
  // Matrix Fields is optional -- most ticket types have no matrix fields at
  // all, so its absence is never treated as "missing" like the 6 core sheets.
  const matrixSheet = xFindSheet(wb, "Matrix Fields");
  const matrixFieldsByTicket = matrixSheet ? xParseMatrixFields(matrixSheet) : {};
  // Cascading Dropdowns / API Dropdowns are optional too, same reasoning.
  const cascadeSheet = xFindSheet(wb, "Cascading Dropdowns");
  const cascadeFieldsByTicket = cascadeSheet ? xParseCascadeDropdowns(cascadeSheet) : {};
  const apiSheet = xFindSheet(wb, "API Dropdowns");
  const apiFieldsByTicket = apiSheet ? xParseApiDropdowns(apiSheet) : {};

  const detailsRows = sheets.details ? xParseTicketDetails(sheets.details) : [];
  const subtypesByTicket = sheets.subtypes ? xParseSubTypes(sheets.subtypes) : {};
  if (sheets.structure) xParseTicketStructure(sheets.structure, subtypesByTicket);
  const statusRolesByTicket = sheets.statusRoles ? xParseStatusAccessRoles(sheets.statusRoles) : {};
  const mandateByTicket = sheets.mandate ? xParseMandateRules(sheets.mandate) : {};
  const { escalationByTicket, warnings } = sheets.escalation ? xParseEscalationDetails(sheets.escalation) : { escalationByTicket: {}, warnings: [] };

  // A workbook might list ticket types only in "Ticket Details", or (in the
  // single-ticket legacy shape) have no Details rows at all but everything
  // else under "__default__". Build the working set of ticket keys from
  // whichever sheets actually have data.
  const allTicketKeys = new Set(detailsRows.map((d) => d.ticketKey || d.ticketType).filter(Boolean));
  if (!allTicketKeys.size) {
    Object.keys(subtypesByTicket).forEach((k) => allTicketKeys.add(k));
  }
  if (!allTicketKeys.size) allTicketKeys.add("__default__");

  const docs = [];
  allTicketKeys.forEach((ticketKeyLookup) => {
    const details = detailsRows.find((d) => d.ticketKey === ticketKeyLookup || d.ticketType === ticketKeyLookup) || {};
    // subtypes/roles/mandate/escalation are keyed by whatever the sheets
    // actually used (ticket key column value, or "__default__"); fall back
    // sensibly so a Details-only ticket key still picks up "__default__"
    // sheet data when there's exactly one ticket type in the workbook.
    const subBucketKey = subtypesByTicket[ticketKeyLookup] ? ticketKeyLookup : (subtypesByTicket.__default__ ? "__default__" : ticketKeyLookup);
    const subBucket = subtypesByTicket[subBucketKey] || { subtypes: {}, order: [] };
    const globalStatusRoles = statusRolesByTicket[ticketKeyLookup] || statusRolesByTicket.__default__ || {};
    const mandateRules = mandateByTicket[ticketKeyLookup] || mandateByTicket.__default__ || {};
    const escalationByStatus = escalationByTicket[ticketKeyLookup] || escalationByTicket.__default__ || {};
    const matrixBucketKey = matrixFieldsByTicket[ticketKeyLookup] ? ticketKeyLookup : (matrixFieldsByTicket.__default__ ? "__default__" : ticketKeyLookup);
    const matrixBySubtype = matrixFieldsByTicket[matrixBucketKey] || {};
    const cascadeBucketKey = cascadeFieldsByTicket[ticketKeyLookup] ? ticketKeyLookup : (cascadeFieldsByTicket.__default__ ? "__default__" : ticketKeyLookup);
    const cascadeBySubtype = cascadeFieldsByTicket[cascadeBucketKey] || {};
    const apiBucketKey = apiFieldsByTicket[ticketKeyLookup] ? ticketKeyLookup : (apiFieldsByTicket.__default__ ? "__default__" : ticketKeyLookup);
    const apiBySubtype = apiFieldsByTicket[apiBucketKey] || {};

    const ticketSubType = subBucket.order.map((name) => {
      const st = subBucket.subtypes[name];
      const matrixFields = (matrixBySubtype[name] || []).map(xMatrixFieldToOutput);
      const cascadeFields = (cascadeBySubtype[name] || []).map(xCascadeFieldToOutput);
      const apiFields = (apiBySubtype[name] || []).map(xApiFieldToOutput);
      return {
        ticketKey: st.ticketKey, prefix: st.prefix, ticketType: st.ticketType,
        assignee: (st.assigneeRoles || []).map((r) => ({ roleId: r })),
        customFieldsMetaData: [...st.customFieldsMetaData.map(xFieldToOutput), ...matrixFields, ...cascadeFields, ...apiFields],
        statusWorkFlow: xBuildStatusWorkflow(name, globalStatusRoles, mandateRules, escalationByStatus),
        deleted: st.deleted, isAutoEscalation: st.isAutoEscalation,
      };
    });

    const rawDoc = {
      ticketKey: details.ticketKey || (ticketKeyLookup !== "__default__" ? ticketKeyLookup : ""),
      ticketType: details.ticketType || "",
      tenantId: details.tenantId || "",
      deleted: !!details.deleted,
      creators: (details.creators || []).map((r) => ({ roleId: r })),
      viewers: (details.viewers || []).map((r) => ({ roleId: r })),
      assignee: (details.assignee || []).map((r) => ({ roleId: r })),
      ticketSubType,
    };
    docs.push(jsonToDoc(rawDoc));
  });

  return { docs, warnings, missingSheets };
}

/* Reverse of excelWorkbookToDocs: builds the same 9-sheet workbook,
   pre-filled with one or more ticket types' CURRENT data, so an editor
   with no context can open it, change cells, and hand it back. Every
   sheet after "Ticket Details" carries a leading "Ticket Key" column so
   rows from different ticket types can share one sheet without
   ambiguity; the Escalation sheet groups its "STATUS:" blocks under a
   "TICKET: <key>" marker per ticket type. Matrix, cascading-dropdown, and
   API-dropdown fields each get their own dedicated sheet (Matrix Fields /
   Cascading Dropdowns / API Dropdowns) since they don't fit the flat
   one-row-per-option grammar Ticket Structure uses. Any other field type
   the sheet convention doesn't yet know about is skipped and reported so
   nothing is silently lost. Status-level settings that the
   sheet format only tracks once per (ticket, status) pair are exported
   from the first sub-type that defines that code within that ticket
   type — a known limitation of the flat-sheet format itself. */
const XLSX_HEADER_STYLE = {
  font: { bold: true, color: { rgb: "FFFFFF" } },
  fill: { fgColor: { rgb: "1F3864" } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  border: {
    top: { style: "thin", color: { rgb: "B8C4DA" } },
    bottom: { style: "thin", color: { rgb: "B8C4DA" } },
    left: { style: "thin", color: { rgb: "B8C4DA" } },
    right: { style: "thin", color: { rgb: "B8C4DA" } },
  },
};

/* Applies the workbook's visual theme to a freshly-built sheet: bold white
   header row on a navy fill, sensible column widths, and an autofilter over
   the header row. Deliberately skips per-row banding -- some of these
   sheets run into the thousands of rows (Cascading Dropdowns can exceed
   4,000), and writing a style object onto every cell would noticeably
   bloat file size and slow generation for little readability gain on a
   data-entry sheet like this. Requires "xlsx-js-style" (a SheetJS fork
   with style-writing support) -- the plain "xlsx" package silently drops
   cell styles on write, which is why exports looked unstyled before. */
function styleSheet(ws, numCols) {
  for (let c = 0; c < numCols; c++) {
    const ref = XLSX.utils.encode_cell({ r: 0, c });
    if (ws[ref]) ws[ref].s = XLSX_HEADER_STYLE;
  }
  ws["!cols"] = Array.from({ length: numCols }, (_, c) => {
    const ref = XLSX.utils.encode_cell({ r: 0, c });
    const headerLen = ws[ref] && ws[ref].v ? String(ws[ref].v).length : 10;
    return { wch: Math.min(Math.max(headerLen + 4, 14), 34) };
  });
  const rowCount = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]).e.r + 1 : 1;
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rowCount - 1, c: numCols - 1 } }) };
}

function excelWorkbookFromDocs(docs) {
  const wb = XLSX.utils.book_new();
  const skippedFields = [];
  const list = docs.filter(Boolean);

  const detailsRows = [["Ticket Key", "Ticket Type", "Tenant", "Creator Roles", "Viewer Roles", "Assignee Roles", "Delete"]];
  list.forEach((doc) => {
    detailsRows.push([doc.ticketKey, doc.ticketType, doc.tenantId, doc.creators.join(", "), doc.viewers.join(", "), doc.assignee.join(", "), doc.deleted ? "Yes" : "No"]);
  });
  {
    const ws_ = XLSX.utils.aoa_to_sheet(detailsRows);
    styleSheet(ws_, (detailsRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Ticket Details");
  }

  const subRows = [["Ticket Key", "Ticket Subtype", "Subtype Key", "Prefix", "Assignee Roles", "Auto Escalation", "Delete"]];
  list.forEach((doc) => {
    doc.ticketSubType.forEach((s) => {
      subRows.push([doc.ticketKey, s.ticketType, s.ticketKey, s.prefix, (s.assigneeRoles || []).join(", "), s.isAutoEscalation ? "Yes" : "No", s.deleted ? "Yes" : "No"]);
    });
  });
  {
    const ws_ = XLSX.utils.aoa_to_sheet(subRows);
    styleSheet(ws_, (subRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Sub Types");
  }

  const EXCEL_TYPE_LABEL = { dropdown: "dropdown", radio: "radio", multiplechoice: "multichoice", date: "date", time: "time", number: "number", text: "text", fileUpload: "file" };
  const structRows = [["Ticket Key", "Ticket Subtype", "Field (Type)", "Option / Value", "Mandatory", "Edit Roles"]];
  const matrixFieldEntries = []; // { doc, sub, field } collected here, written to their own sheet below
  const cascadeFieldEntries = []; // { doc, sub, field } → "Cascading Dropdowns" sheet
  const apiDropdownEntries = []; // { doc, sub, field } → "API Dropdowns" sheet
  list.forEach((doc) => {
    doc.ticketSubType.forEach((s) => {
      s.customFieldsMetaData.forEach((f) => {
        if (f.type === "matrix") { matrixFieldEntries.push({ doc, sub: s, field: f }); return; }
        if (f.type === "casecadeDropdown") { cascadeFieldEntries.push({ doc, sub: s, field: f }); return; }
        if (f.type === "apiDropdown") { apiDropdownEntries.push({ doc, sub: s, field: f }); return; }
        if (!EXCEL_TYPE_LABEL[f.type]) { skippedFields.push(`${doc.ticketType || doc.ticketKey} / ${s.ticketType || s.ticketKey} → ${f.label || f.key} (${f.type})`); return; }
        const labelWithType = `${f.label} (${EXCEL_TYPE_LABEL[f.type]})`;
        const mandatoryCell = f.isMandatory ? "Yes" : "No";
        const opts = OPTION_BASED_TYPES.includes(f.type) ? (f.options || []) : [];
        if (opts.length === 0) {
          structRows.push([doc.ticketKey, s.ticketType, labelWithType, "", mandatoryCell, (f.editRoles || []).join(", ")]);
        } else {
          opts.forEach((o, i) => {
            structRows.push([i === 0 ? doc.ticketKey : "", i === 0 ? s.ticketType : "", i === 0 ? labelWithType : "", o.value, i === 0 ? mandatoryCell : "", i === 0 ? (f.editRoles || []).join(", ") : ""]);
          });
        }
      });
    });
  });
  {
    const ws_ = XLSX.utils.aoa_to_sheet(structRows);
    styleSheet(ws_, (structRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Ticket Structure");
  }

  // Matrix fields get their own sheet, laid out to visually mirror the
  // matrix table itself: one block per field, its own columns spelled out
  // as headers, and option values for any dropdown columns listed beneath.
  const matrixRows = [["Ticket Key", "Ticket Subtype", "Mandatory", "Edit Roles", "Field (Type) \u2192 columns (Type) \u2192"]];
  matrixFieldEntries.forEach(({ doc, sub, field }) => {
    const cols = field.columns || [];
    const colHeaders = cols.map((c) => `${c.label} (${c.type})`);
    matrixRows.push([doc.ticketKey, sub.ticketType, field.isMandatory ? "Yes" : "No", (field.editRoles || []).join(", "), `${field.label} (Matrix)`, ...colHeaders]);
    const maxOptions = Math.max(0, ...cols.map((c) => (c.type === "dropdown" ? (c.options || []).length : 0)));
    for (let i = 0; i < maxOptions; i++) {
      const row = ["", "", "", "", ""];
      cols.forEach((c) => {
        row.push(c.type === "dropdown" && c.options && c.options[i] ? c.options[i].value : "");
      });
      matrixRows.push(row);
    }
    matrixRows.push([]);
  });
  if (matrixRows.length <= 1) matrixRows.push(["No matrix fields configured."]);
  {
    const ws_ = XLSX.utils.aoa_to_sheet(matrixRows);
    styleSheet(ws_, (matrixRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Matrix Fields");
  }

  // Cascading dropdowns: a parent dropdown whose selected option reveals a
  // second ("child") dropdown. Laid out as one block per field, one row per
  // parent option, with each of that parent's children listed on their own
  // row directly beneath it (Parent Option left blank on those follow-on
  // rows so the grouping reads clearly) — same "field header row, then
  // option rows, blank row to close" convention as Ticket Structure /
  // Matrix Fields use elsewhere in this workbook.
  const cascadeRows = [["Ticket Key", "Ticket Subtype", "Mandatory", "Edit Roles", "Field (Type)", "Parent Option Type", "Child Field Label", "Child Field Key", "Parent Option", "Child Option"]];
  cascadeFieldEntries.forEach(({ doc, sub, field }) => {
    const parents = field.options || [];
    const metaCells = [doc.ticketKey, sub.ticketType, field.isMandatory ? "Yes" : "No", (field.editRoles || []).join(", "), `${field.label} (cascade)`, field.parentOptionType || "multiplechoice", field.childLabel, field.childKey];
    let firstRow = true;
    if (parents.length === 0) {
      cascadeRows.push([...metaCells, "", ""]);
      firstRow = false;
    }
    parents.forEach((p) => {
      const children = (p.children && p.children.length) ? p.children : [null];
      children.forEach((c, ci) => {
        const lead = firstRow ? metaCells : ["", "", "", "", "", "", "", ""];
        firstRow = false;
        cascadeRows.push([...lead, ci === 0 ? p.name : "", c ? c.value : ""]);
      });
    });
    cascadeRows.push([]); // blank row closes this field's block
  });
  if (cascadeRows.length <= 1) cascadeRows.push(["No cascading dropdown fields configured."]);
  {
    const ws_ = XLSX.utils.aoa_to_sheet(cascadeRows);
    styleSheet(ws_, (cascadeRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Cascading Dropdowns");
  }

  // API dropdowns: options are fetched live from an endpoint at runtime, so
  // there's nothing to enumerate — just the field's wiring, one row each.
  const apiRows = [["Ticket Key", "Ticket Subtype", "Mandatory", "Edit Roles", "Field (Type)", "Label Key", "Value Key", "Fill On Select", "API Filters (JSON)"]];
  apiDropdownEntries.forEach(({ doc, sub, field }) => {
    let filtersText = "{}";
    try { filtersText = JSON.stringify(JSON.parse(field.apiFiltersText || "{}")); } catch (e) { filtersText = field.apiFiltersText || "{}"; }
    apiRows.push([doc.ticketKey, sub.ticketType, field.isMandatory ? "Yes" : "No", (field.editRoles || []).join(", "), `${field.label} (api-dropdown)`, field.labelKey, field.valueKey, field.fillOnSelect ? "Yes" : "No", filtersText]);
  });
  if (apiRows.length <= 1) apiRows.push(["No API dropdown fields configured."]);
  {
    const ws_ = XLSX.utils.aoa_to_sheet(apiRows);
    styleSheet(ws_, (apiRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "API Dropdowns");
  }

  const roleRows = [["Ticket Key", "Status", "Roles Allowed"]];
  list.forEach((doc) => {
    const statusRoles = new Map();
    doc.ticketSubType.forEach((s) => s.statusWorkFlow.forEach((st) => {
      const code = xToStatusCode(st.status);
      const entry = statusRoles.get(code) || { label: st.label || st.status, roles: new Set() };
      (st.roles || []).forEach((r) => entry.roles.add(r));
      statusRoles.set(code, entry);
    }));
    statusRoles.forEach((v) => roleRows.push([doc.ticketKey, v.label, [...v.roles].join(", ")]));
  });
  {
    const ws_ = XLSX.utils.aoa_to_sheet(roleRows);
    styleSheet(ws_, (roleRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Status Access Roles");
  }

  const mandateRows = [["Ticket Key", "Status", "Mandatory Images", "Mandatory Comments", "Mandatory New Assignee", "New Assignee Role"]];
  list.forEach((doc) => {
    const mandateMap = new Map();
    doc.ticketSubType.forEach((s) => s.statusWorkFlow.forEach((st) => {
      const code = xToStatusCode(st.status);
      if (mandateMap.has(code)) return;
      const mand = st.mandatoryCustomFields || [];
      const commentField = (st.customFieldsMetaData || []).find((f) => mand.includes(f.key));
      mandateMap.set(code, {
        label: st.label || st.status,
        images: mand.includes("images") ? "Yes" : "No",
        comments: commentField ? commentField.label : (mand.includes("comments") ? "Yes" : "No"),
        newAssignee: mand.includes("newAssignee") ? "Yes" : "No",
        newAssigneeRole: (st.assigneeRoles || []).join(", "),
      });
    }));
    mandateMap.forEach((v) => mandateRows.push([doc.ticketKey, v.label, v.images, v.comments, v.newAssignee, v.newAssigneeRole]));
  });
  {
    const ws_ = XLSX.utils.aoa_to_sheet(mandateRows);
    styleSheet(ws_, (mandateRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Mandate rules for status change");
  }

  const escRows = [];
  list.forEach((doc) => {
    const byStatus = new Map();
    doc.ticketSubType.forEach((s) => s.statusWorkFlow.forEach((st) => {
      if (!st.autoEscalationConfig || !st.autoEscalationConfig.length) return;
      const code = xToStatusCode(st.status);
      const l = byStatus.get(code) || { label: st.label || st.status, rows: [] };
      l.rows.push({ subtype: s.ticketType || s.ticketKey, chain: st.autoEscalationConfig });
      byStatus.set(code, l);
    }));
    if (!byStatus.size) return;
    escRows.push([`TICKET: ${doc.ticketKey}`]);
    byStatus.forEach((v) => {
      escRows.push([`STATUS: ${v.label}`]);
      escRows.push(["Ticket Subtype", "Issue Type", "Days", "Role", "Days", "Role"]);
      v.rows.forEach((r) => {
        const flat = [];
        r.chain.forEach((step) => flat.push(step.days, step.assignee));
        escRows.push([r.subtype, r.subtype, ...flat]);
      });
      escRows.push([]);
    });
  });
  if (!escRows.length) escRows.push(["No auto-escalation rules configured."]);
  {
    const ws_ = XLSX.utils.aoa_to_sheet(escRows);
    styleSheet(ws_, (escRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Escalation details");
  }

  return { wb, skippedFields };
}

/* ------------------------------------------------------------------ */
/* Templates — reusable sub-type "shapes" used by bulk creation.       */
/* A template is stored independently of any ticket type. Import       */
/* accepts a raw sub-type JSON, a full ticket-type JSON, or a full     */
/* export ({ result: [...] }); whichever is passed, the first sub-type */
/* found becomes the template's fields + status workflow.              */
/* ------------------------------------------------------------------ */

function templateFromRawJson(raw, name) {
  let candidate = raw;
  if (candidate && Array.isArray(candidate.result)) candidate = candidate.result[0];
  if (Array.isArray(candidate)) candidate = candidate[0];
  if (candidate && Array.isArray(candidate.ticketSubType) && candidate.ticketSubType.length) {
    candidate = candidate.ticketSubType[0];
  }
  const st = subTypeFromJson(candidate || {});
  return {
    id: uid(),
    name: (name && name.trim()) || st.ticketType || "Imported template",
    description: "",
    isDefault: false,
    data: {
      prefix: st.prefix || "",
      isAutoEscalation: st.isAutoEscalation,
      assigneeRoles: st.assigneeRoles || [],
      customFieldsMetaData: st.customFieldsMetaData,
      statusWorkFlow: st.statusWorkFlow,
    },
  };
}

/* Builds a template's `data` shape directly from a sub-type that's already
   sitting in the library — lets bulk-create use "an existing sub-type" as
   its source without going through the Templates tab's JSON import step. */
function templateDataFromSubType(sub) {
  return {
    prefix: sub.prefix || "",
    isAutoEscalation: !!sub.isAutoEscalation,
    assigneeRoles: sub.assigneeRoles || [],
    customFieldsMetaData: sub.customFieldsMetaData || [],
    statusWorkFlow: sub.statusWorkFlow || [],
  };
}

/* Stamps out one new sub-type from a template + a chosen display name.
   Every id inside is regenerated so the result is fully independent of
   the template and of any other sub-type generated from it. */
function subTypeFromTemplate(template, name, ticketKey) {
  const cloned = regenerateIds(template.data || {});
  return {
    id: uid(),
    ticketKey,
    prefix: cloned.prefix || "",
    ticketType: name,
    deleted: false,
    isAutoEscalation: !!cloned.isAutoEscalation,
    assigneeRoles: cloned.assigneeRoles || [],
    customFieldsMetaData: cloned.customFieldsMetaData || [],
    statusWorkFlow: cloned.statusWorkFlow || [],
  };
}

/* ------------------------------------------------------------------ */
/* Smart update — apply a list of AI-proposed "change operations" to   */
/* the library deterministically. The AI never edits JSON directly;    */
/* it only proposes ops from this fixed vocabulary, each one is        */
/* resolved and applied (or rejected with a reason) by plain code so   */
/* nothing gets silently corrupted or hallucinated into the schema.    */
/* ------------------------------------------------------------------ */

const SMART_OP_TYPES = [
  "addCreatorRole", "removeCreatorRole",
  "addViewerRole", "removeViewerRole",
  "addDocAssigneeRole", "removeDocAssigneeRole",
  "addSubTypeField", "removeSubTypeField",
  "addStatus", "removeStatus",
  "addStatusRole", "removeStatusRole",
  "addStatusAssigneeRole", "removeStatusAssigneeRole",
  "addStatusField", "removeStatusField",
  "setStatusFieldMandatory",
  "note",
];

/* Resolves + applies a single op against `library` and returns a new
   library reference (only when applied). Never throws — unresolvable
   or already-satisfied ops come back as { applied: false, message }. */
function applyOpToLibrary(library, op) {
  if (!op || typeof op !== "object") return { library, applied: false, message: "Malformed operation" };
  if (op.op === "note") return { library, applied: false, message: op.description || "Flagged for manual review" };
  if (!SMART_OP_TYPES.includes(op.op)) return { library, applied: false, message: `Unrecognized operation "${op.op}"` };

  const docIdx = library.findIndex((d) => xnorm(d.ticketKey).toLowerCase() === xnorm(op.ticketKey).toLowerCase());
  if (docIdx < 0) return { library, applied: false, message: `Ticket type "${op.ticketKey || "?"}" not found` };
  const doc = library[docIdx];

  const withDoc = (nextDoc, message) => ({
    library: library.map((d, i) => (i === docIdx ? nextDoc : d)),
    applied: true,
    message,
  });

  switch (op.op) {
    case "addCreatorRole": {
      if (!op.role) return { library, applied: false, message: "Missing role" };
      if (doc.creators.includes(op.role)) return { library, applied: false, message: `"${op.role}" already a creator role on ${doc.ticketType || doc.ticketKey}` };
      return withDoc({ ...doc, creators: [...doc.creators, op.role] }, `Added "${op.role}" to creators on ${doc.ticketType || doc.ticketKey}`);
    }
    case "removeCreatorRole": {
      if (!doc.creators.includes(op.role)) return { library, applied: false, message: `"${op.role}" isn't a creator role on ${doc.ticketType || doc.ticketKey}` };
      return withDoc({ ...doc, creators: doc.creators.filter((r) => r !== op.role) }, `Removed "${op.role}" from creators on ${doc.ticketType || doc.ticketKey}`);
    }
    case "addViewerRole": {
      if (!op.role) return { library, applied: false, message: "Missing role" };
      if (doc.viewers.includes(op.role)) return { library, applied: false, message: `"${op.role}" already a viewer role on ${doc.ticketType || doc.ticketKey}` };
      return withDoc({ ...doc, viewers: [...doc.viewers, op.role] }, `Added "${op.role}" to viewers on ${doc.ticketType || doc.ticketKey}`);
    }
    case "removeViewerRole": {
      if (!doc.viewers.includes(op.role)) return { library, applied: false, message: `"${op.role}" isn't a viewer role on ${doc.ticketType || doc.ticketKey}` };
      return withDoc({ ...doc, viewers: doc.viewers.filter((r) => r !== op.role) }, `Removed "${op.role}" from viewers on ${doc.ticketType || doc.ticketKey}`);
    }
    case "addDocAssigneeRole": {
      if (!op.role) return { library, applied: false, message: "Missing role" };
      if (doc.assignee.includes(op.role)) return { library, applied: false, message: `"${op.role}" already an assignee role on ${doc.ticketType || doc.ticketKey}` };
      return withDoc({ ...doc, assignee: [...doc.assignee, op.role] }, `Added "${op.role}" to ticket-level assignee on ${doc.ticketType || doc.ticketKey}`);
    }
    case "removeDocAssigneeRole": {
      if (!doc.assignee.includes(op.role)) return { library, applied: false, message: `"${op.role}" isn't an assignee role on ${doc.ticketType || doc.ticketKey}` };
      return withDoc({ ...doc, assignee: doc.assignee.filter((r) => r !== op.role) }, `Removed "${op.role}" from ticket-level assignee on ${doc.ticketType || doc.ticketKey}`);
    }
    default:
      break;
  }

  // everything below is scoped to a sub-type
  if (!op.subTypeKey) return { library, applied: false, message: `Missing subTypeKey for "${op.op}"` };
  const subIdx = doc.ticketSubType.findIndex((s) => xnorm(s.ticketKey).toLowerCase() === xnorm(op.subTypeKey).toLowerCase());
  if (subIdx < 0) return { library, applied: false, message: `Sub-type "${op.subTypeKey}" not found in ${doc.ticketType || doc.ticketKey}` };
  const sub = doc.ticketSubType[subIdx];

  const withSub = (nextSub, message) => withDoc(
    { ...doc, ticketSubType: doc.ticketSubType.map((s, i) => (i === subIdx ? nextSub : s)) },
    message
  );

  if (op.op === "addSubTypeField") {
    if (!op.field || !op.field.key) return { library, applied: false, message: "Missing field definition" };
    if (sub.customFieldsMetaData.some((f) => f.key === op.field.key)) {
      return { library, applied: false, message: `Field "${op.field.key}" already exists on ${sub.ticketType || sub.ticketKey}` };
    }
    return withSub({ ...sub, customFieldsMetaData: [...sub.customFieldsMetaData, fieldFromJson(op.field)] }, `Added field "${op.field.label || op.field.key}" to ${sub.ticketType || sub.ticketKey}`);
  }
  if (op.op === "removeSubTypeField") {
    if (!sub.customFieldsMetaData.some((f) => f.key === op.fieldKey)) {
      return { library, applied: false, message: `Field "${op.fieldKey}" not found on ${sub.ticketType || sub.ticketKey}` };
    }
    return withSub({ ...sub, customFieldsMetaData: sub.customFieldsMetaData.filter((f) => f.key !== op.fieldKey) }, `Removed field "${op.fieldKey}" from ${sub.ticketType || sub.ticketKey}`);
  }
  if (op.op === "addStatus") {
    if (!op.status || !op.status.status) return { library, applied: false, message: "Missing status definition" };
    if (sub.statusWorkFlow.some((s) => xnorm(s.status).toLowerCase() === xnorm(op.status.status).toLowerCase())) {
      return { library, applied: false, message: `Status "${op.status.status}" already exists on ${sub.ticketType || sub.ticketKey}` };
    }
    return withSub({ ...sub, statusWorkFlow: [...sub.statusWorkFlow, statusFromJson(op.status)] }, `Added status "${op.status.label || op.status.status}" to ${sub.ticketType || sub.ticketKey}`);
  }
  if (op.op === "removeStatus") {
    if (!sub.statusWorkFlow.some((s) => xnorm(s.status).toLowerCase() === xnorm(op.statusCode).toLowerCase())) {
      return { library, applied: false, message: `Status "${op.statusCode}" not found on ${sub.ticketType || sub.ticketKey}` };
    }
    return withSub({ ...sub, statusWorkFlow: sub.statusWorkFlow.filter((s) => xnorm(s.status).toLowerCase() !== xnorm(op.statusCode).toLowerCase()) }, `Removed status "${op.statusCode}" from ${sub.ticketType || sub.ticketKey}`);
  }

  // everything below is additionally scoped to a status
  if (!op.statusCode) return { library, applied: false, message: `Missing statusCode for "${op.op}"` };
  const statusIdx = sub.statusWorkFlow.findIndex((s) => xnorm(s.status).toLowerCase() === xnorm(op.statusCode).toLowerCase());
  if (statusIdx < 0) return { library, applied: false, message: `Status "${op.statusCode}" not found on ${sub.ticketType || sub.ticketKey}` };
  const status = sub.statusWorkFlow[statusIdx];

  const withStatus = (nextStatus, message) => withSub(
    { ...sub, statusWorkFlow: sub.statusWorkFlow.map((s, i) => (i === statusIdx ? nextStatus : s)) },
    message
  );

  switch (op.op) {
    case "addStatusRole": {
      if (!op.role) return { library, applied: false, message: "Missing role" };
      if (status.roles.includes(op.role)) return { library, applied: false, message: `"${op.role}" already allowed to set "${status.label || op.statusCode}"` };
      return withStatus({ ...status, roles: [...status.roles, op.role] }, `Added "${op.role}" to roles allowed to set "${status.label || op.statusCode}"`);
    }
    case "removeStatusRole": {
      if (!status.roles.includes(op.role)) return { library, applied: false, message: `"${op.role}" not in roles for "${status.label || op.statusCode}"` };
      return withStatus({ ...status, roles: status.roles.filter((r) => r !== op.role) }, `Removed "${op.role}" from roles allowed to set "${status.label || op.statusCode}"`);
    }
    case "addStatusAssigneeRole": {
      if (!op.role) return { library, applied: false, message: "Missing role" };
      if (status.assigneeRoles.includes(op.role)) return { library, applied: false, message: `"${op.role}" already an assignee role on "${status.label || op.statusCode}"` };
      return withStatus({ ...status, assigneeRoles: [...status.assigneeRoles, op.role] }, `Added "${op.role}" to assignee roles on "${status.label || op.statusCode}"`);
    }
    case "removeStatusAssigneeRole": {
      if (!status.assigneeRoles.includes(op.role)) return { library, applied: false, message: `"${op.role}" isn't an assignee role on "${status.label || op.statusCode}"` };
      return withStatus({ ...status, assigneeRoles: status.assigneeRoles.filter((r) => r !== op.role) }, `Removed "${op.role}" from assignee roles on "${status.label || op.statusCode}"`);
    }
    case "addStatusField": {
      if (!op.field || !op.field.key) return { library, applied: false, message: "Missing field definition" };
      if (status.customFieldsMetaData.some((f) => f.key === op.field.key)) {
        return { library, applied: false, message: `Field "${op.field.key}" already exists on "${status.label || op.statusCode}"` };
      }
      const newFieldObj = fieldFromJson(op.field);
      const mandatory = !!op.field.isMandatory || !!op.field.mandatory;
      return withStatus({
        ...status,
        customFieldsMetaData: [...status.customFieldsMetaData, newFieldObj],
        mandatoryCustomFields: mandatory ? [...status.mandatoryCustomFields, newFieldObj.key] : status.mandatoryCustomFields,
      }, `Added ${mandatory ? "mandatory " : ""}field "${op.field.label || op.field.key}" to status "${status.label || op.statusCode}"`);
    }
    case "removeStatusField": {
      if (!status.customFieldsMetaData.some((f) => f.key === op.fieldKey)) {
        return { library, applied: false, message: `Field "${op.fieldKey}" not found on "${status.label || op.statusCode}"` };
      }
      return withStatus({
        ...status,
        customFieldsMetaData: status.customFieldsMetaData.filter((f) => f.key !== op.fieldKey),
        mandatoryCustomFields: status.mandatoryCustomFields.filter((k) => k !== op.fieldKey),
      }, `Removed field "${op.fieldKey}" from status "${status.label || op.statusCode}"`);
    }
    case "setStatusFieldMandatory": {
      if (!op.fieldKey) return { library, applied: false, message: "Missing fieldKey" };
      if (status.mandatoryCustomFields.includes(op.fieldKey)) return { library, applied: false, message: `"${op.fieldKey}" already mandatory on "${status.label || op.statusCode}"` };
      return withStatus({ ...status, mandatoryCustomFields: [...status.mandatoryCustomFields, op.fieldKey] }, `Marked "${op.fieldKey}" mandatory on status "${status.label || op.statusCode}"`);
    }
    default:
      return { library, applied: false, message: `Unrecognized operation "${op.op}"` };
  }
}

const SMART_UPDATE_SYSTEM_PROMPT = `You are a precise change-planner for a ticket-schema editor. You will be given CURRENT JSON (the real, current data for one or more ticket types, in the app's export format) and REQUIREMENTS (plain English or bullet points describing desired changes).

Output ONLY a JSON array of "change operations" — no prose, no markdown code fences, nothing before or after the array. Each element is an object with an "op" field plus whatever fields that op needs, and always a short human-readable "description" field.

Allowed "op" values and their fields:
- addCreatorRole / removeCreatorRole: { ticketKey, role, description }
- addViewerRole / removeViewerRole: { ticketKey, role, description }
- addDocAssigneeRole / removeDocAssigneeRole: { ticketKey, role, description }  (ticket-level "assignee" array)
- addSubTypeField: { ticketKey, subTypeKey, field, description }
- removeSubTypeField: { ticketKey, subTypeKey, fieldKey, description }
- addStatus: { ticketKey, subTypeKey, status, description }
- removeStatus: { ticketKey, subTypeKey, statusCode, description }
- addStatusRole / removeStatusRole: { ticketKey, subTypeKey, statusCode, role, description }  (who can set this status)
- addStatusAssigneeRole / removeStatusAssigneeRole: { ticketKey, subTypeKey, statusCode, role, description }  (assignee at this status)
- addStatusField: { ticketKey, subTypeKey, statusCode, field, description }
- removeStatusField: { ticketKey, subTypeKey, statusCode, fieldKey, description }
- setStatusFieldMandatory: { ticketKey, subTypeKey, statusCode, fieldKey, description }  (mark an EXISTING field mandatory, don't use this to add a new field)
- note: { ticketKey, description }  (use when a requirement is ambiguous or can't be confidently mapped to the ops above — explain what's unclear)

Rules:
- ticketKey, subTypeKey, statusCode, fieldKey must exactly match values already present in CURRENT JSON (case-sensitive) whenever you're referring to something that should already exist. Only invent a new statusCode/fieldKey when the requirement is clearly asking to add something new — keep new keys short camelCase, consistent with existing ones nearby.
- Any "status" object you output (for addStatus) must follow the EXACT JSON shape used by entries in that sub-type's "statusWorkFlow" in CURRENT JSON — including "assignee" as an array of {"roleId": "..."} objects, not plain strings.
- Any "field" object you output (for addSubTypeField / addStatusField) must follow the EXACT JSON shape used by entries in "customFieldsMetaData" in CURRENT JSON — including "isMandatory": true when the requirement says the field should be mandatory, and "options": [{"value":"...","key":"..."}] for dropdown/radio/multiplechoice types.
- One operation per atomic change. If a requirement implies several changes (e.g. "add role X and remove role Y"), emit separate ops.
- If CURRENT JSON contains multiple ticket types, infer which one each requirement applies to from context; if it's genuinely unclear, emit a "note" op instead of guessing.
- Never output anything except the JSON array.`;

/* Deterministic alternative to the AI planner: compares an old and a
   new doc (e.g. current state vs. a re-uploaded, hand-edited Excel
   export) and emits the exact same op vocabulary applyOpToLibrary()
   understands. No interpretation involved — every op is a plain fact
   about what differs, so this works with zero context from whoever
   edited the sheet. Anything the flat sheet format can change but this
   function can't safely auto-diff (a field's own type/options/label,
   a status's other settings, whole sub-types) comes back as a "note"
   for manual review instead of being guessed at. */
function diffDocsToOps(oldDoc, newDoc) {
  const ops = [];
  const ticketKey = newDoc.ticketKey || oldDoc.ticketKey;

  const diffRoleList = (oldList, newList, addOp, removeOp, extra = {}) => {
    (newList || []).filter((r) => !(oldList || []).includes(r)).forEach((r) =>
      ops.push({ op: addOp, ticketKey, role: r, ...extra, description: `Add "${r}"` }));
    (oldList || []).filter((r) => !(newList || []).includes(r)).forEach((r) =>
      ops.push({ op: removeOp, ticketKey, role: r, ...extra, description: `Remove "${r}"` }));
  };

  diffRoleList(oldDoc.creators, newDoc.creators, "addCreatorRole", "removeCreatorRole");
  diffRoleList(oldDoc.viewers, newDoc.viewers, "addViewerRole", "removeViewerRole");
  diffRoleList(oldDoc.assignee, newDoc.assignee, "addDocAssigneeRole", "removeDocAssigneeRole");

  const STATUS_META_KEYS = ["label", "public", "notification", "notificationType", "notificationRole", "overrideStatus", "dependentStatus"];
  const pickStatusMeta = (s) => { const o = {}; STATUS_META_KEYS.forEach((k) => { o[k] = s[k]; }); return o; };

  const oldSubs = new Map((oldDoc.ticketSubType || []).map((s) => [xnorm(s.ticketKey).toLowerCase(), s]));
  const newSubs = new Map((newDoc.ticketSubType || []).map((s) => [xnorm(s.ticketKey).toLowerCase(), s]));

  newSubs.forEach((newSub, key) => {
    const oldSub = oldSubs.get(key);
    if (!oldSub) {
      ops.push({ op: "note", ticketKey, description: `New sub-type "${newSub.ticketType || newSub.ticketKey}" found in the sheet — add it via Bulk Create or the Sub-Types tab (the diff never auto-creates sub-types).` });
      return;
    }
    const subTypeKey = newSub.ticketKey;
    const subLabel = newSub.ticketType || subTypeKey;

    const oldFields = new Map((oldSub.customFieldsMetaData || []).map((f) => [f.key, f]));
    const newFields = new Map((newSub.customFieldsMetaData || []).map((f) => [f.key, f]));
    newFields.forEach((f, fk) => {
      if (!oldFields.has(fk)) {
        ops.push({ op: "addSubTypeField", ticketKey, subTypeKey, field: fieldToJson(f), description: `Add field "${f.label || fk}" to ${subLabel}` });
      } else if (JSON.stringify(fieldToJson(oldFields.get(fk))) !== JSON.stringify(fieldToJson(f))) {
        ops.push({ op: "note", ticketKey, subTypeKey, description: `Field "${fk}" on ${subLabel} changed shape (label/type/options) — the diff only adds or removes whole fields, review this one in the Sub-Types tab.` });
      }
    });
    oldFields.forEach((f, fk) => {
      if (!newFields.has(fk)) {
        ops.push({ op: "removeSubTypeField", ticketKey, subTypeKey, fieldKey: fk, description: `Remove field "${f.label || fk}" from ${subLabel}` });
      }
    });

    const oldStatuses = new Map((oldSub.statusWorkFlow || []).map((s) => [xnorm(s.status).toLowerCase(), s]));
    const newStatuses = new Map((newSub.statusWorkFlow || []).map((s) => [xnorm(s.status).toLowerCase(), s]));

    newStatuses.forEach((s, sk) => {
      const statusCode = s.status;
      const statusLabel = s.label || statusCode;
      if (!oldStatuses.has(sk)) {
        ops.push({ op: "addStatus", ticketKey, subTypeKey, status: statusToJson(s), description: `Add status "${statusLabel}" to ${subLabel}` });
        return;
      }
      const oldStatus = oldStatuses.get(sk);
      diffRoleList(oldStatus.roles, s.roles, "addStatusRole", "removeStatusRole", { subTypeKey, statusCode });
      diffRoleList(oldStatus.assigneeRoles, s.assigneeRoles, "addStatusAssigneeRole", "removeStatusAssigneeRole", { subTypeKey, statusCode });

      const oldSFields = new Map((oldStatus.customFieldsMetaData || []).map((f) => [f.key, f]));
      const newSFields = new Map((s.customFieldsMetaData || []).map((f) => [f.key, f]));
      newSFields.forEach((f, fk) => {
        if (!oldSFields.has(fk)) {
          ops.push({ op: "addStatusField", ticketKey, subTypeKey, statusCode, field: fieldToJson(f), description: `Add field "${f.label || fk}" to status "${statusLabel}"` });
        } else if (JSON.stringify(fieldToJson(oldSFields.get(fk))) !== JSON.stringify(fieldToJson(f))) {
          ops.push({ op: "note", ticketKey, subTypeKey, statusCode, description: `Field "${fk}" on status "${statusLabel}" changed shape — review it manually in the Status Workflow editor.` });
        }
      });
      oldSFields.forEach((f, fk) => {
        if (!newSFields.has(fk)) {
          ops.push({ op: "removeStatusField", ticketKey, subTypeKey, statusCode, fieldKey: fk, description: `Remove field "${f.label || fk}" from status "${statusLabel}"` });
        }
      });

      (s.mandatoryCustomFields || []).forEach((fk) => {
        if (!(oldStatus.mandatoryCustomFields || []).includes(fk) && oldSFields.has(fk)) {
          ops.push({ op: "setStatusFieldMandatory", ticketKey, subTypeKey, statusCode, fieldKey: fk, description: `Mark "${fk}" mandatory on status "${statusLabel}"` });
        }
      });

      if (JSON.stringify(pickStatusMeta(oldStatus)) !== JSON.stringify(pickStatusMeta(s))) {
        ops.push({ op: "note", ticketKey, subTypeKey, statusCode, description: `Status "${statusLabel}" has other changes (label/visibility/notifications/override) — review it manually in the Status Workflow editor.` });
      }
    });

    oldStatuses.forEach((s, sk) => {
      if (!newStatuses.has(sk)) {
        ops.push({ op: "removeStatus", ticketKey, subTypeKey, statusCode: s.status, description: `Remove status "${s.label || s.status}" from ${subLabel}` });
      }
    });
  });

  oldSubs.forEach((oldSub, key) => {
    if (!newSubs.has(key)) {
      ops.push({ op: "note", ticketKey, subTypeKey: oldSub.ticketKey, description: `Sub-type "${oldSub.ticketType || oldSub.ticketKey}" is missing from the sheet — delete it manually in the Sub-Types tab if that's intentional (the diff never auto-deletes sub-types).` });
    }
  });

  return ops;
}

/* ------------------------------------------------------------------ */
/* tiny UI primitives                                                  */
/* ------------------------------------------------------------------ */

const Label = ({ children, mono }) => (
  <label className={`block text-[11px] font-semibold uppercase tracking-wider text-slate-400 mb-1.5 ${mono ? "font-mono normal-case tracking-normal" : ""}`}>
    {children}
  </label>
);

const Input = ({ value, onChange, placeholder, mono, ...rest }) => (
  <input
    value={value}
    onChange={(e) => onChange(e.target.value)}
    placeholder={placeholder}
    className={`w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-sm text-slate-100 placeholder:text-slate-600 outline-none transition focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20 ${mono ? "font-mono" : ""}`}
    {...rest}
  />
);

const Select = ({ value, onChange, options }) => (
  <select
    value={value}
    onChange={(e) => onChange(e.target.value)}
    className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-sm text-slate-100 outline-none transition focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20"
  >
    {options.map((o) => (
      <option key={o.value} value={o.value}>{o.label}</option>
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
      <span className={`inline-block h-3.5 w-3.5 transform rounded-full bg-white shadow transition ${checked ? "translate-x-4.5" : "translate-x-1"}`} style={{ transform: checked ? "translateX(18px)" : "translateX(4px)" }} />
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
    accent: "bg-teal-600 text-white hover:bg-teal-500",
    ghost: "bg-transparent text-slate-300 hover:bg-white/5 border border-white/10",
    dashed: "border border-dashed border-white/15 text-slate-500 hover:border-teal-400 hover:text-teal-300 hover:bg-teal-500/5",
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
      <button type="button" onClick={onRemove} className="hover:text-rose-500"><X size={12} /></button>
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
    <div className="rounded-lg border border-white/10 bg-[#0B0F1C] p-2">
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
        className="w-full text-sm outline-none placeholder:text-slate-300"
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
          className={`rounded-md px-2.5 py-1 text-xs font-mono border transition ${active ? "border-teal-500 bg-teal-500 text-white" : "border-white/10 text-slate-500 hover:border-white/25"}`}
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

const Accordion = ({ title, subtitle, tone = "slate", defaultOpen, forceOpen, domId, onDelete, children }) => {
  const [open, setOpen] = useState(!!defaultOpen);
  React.useEffect(() => { if (forceOpen) setOpen(true); }, [forceOpen]);
  const tones = { slate: "border-white/10", teal: "border-teal-500/30", amber: "border-amber-500/30" };
  return (
    <div id={domId} className={`rounded-xl border ${tones[tone]} bg-[#111528] mb-3 overflow-hidden ${forceOpen ? "ring-2 ring-teal-400/40" : ""}`}>
      <div className="flex items-center justify-between px-3 py-2.5 cursor-pointer select-none" onClick={() => setOpen(!open)}>
        <div className="flex items-center gap-2 min-w-0">
          {open ? <ChevronDown size={16} className="text-slate-400 shrink-0" /> : <ChevronRight size={16} className="text-slate-400 shrink-0" />}
          <div className="min-w-0">
            <div className="text-sm font-medium text-slate-100 truncate">{title}</div>
            {subtitle && <div className="text-xs text-slate-400 font-mono truncate">{subtitle}</div>}
          </div>
        </div>
        {onDelete && (
          <IconBtn danger title="Remove" onClick={(e) => { e.stopPropagation(); onDelete(); }}>
            <Trash2 size={14} />
          </IconBtn>
        )}
      </div>
      {open && <div className="px-3 pb-3 pt-1 border-t border-white/5">{children}</div>}
    </div>
  );
};

/* ------------------------------------------------------------------ */
/* Options editors                                                     */
/* ------------------------------------------------------------------ */

function SimpleOptionsEditor({ options, onChange, allFieldKeys }) {
  const [openId, setOpenId] = useState(null);
  const update = (id, patch) => onChange(options.map((o) => (o.id === id ? { ...o, ...patch } : o)));
  const remove = (id) => onChange(options.filter((o) => o.id !== id));
  const add = () => onChange([...options, newSimpleOption()]);
  return (
    <div>
      <div className="grid grid-cols-[1fr_1fr_28px_28px] gap-2 mb-1.5">
        <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">Value (label shown)</span>
        <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">Key</span>
        <span /><span />
      </div>
      <div className="space-y-2">
        {options.map((o) => {
          const hasCondition = o.dependentCustomFields && o.dependentCustomFields.length > 0;
          const open = openId === o.id;
          return (
            <div key={o.id} className="rounded-lg border border-white/5">
              <div className="grid grid-cols-[1fr_1fr_28px_28px] gap-2 items-center p-1">
                <Input value={o.value} onChange={(v) => update(o.id, { value: v })} placeholder="e.g. Marshmallow-Bl" />
                <Input mono value={o.key} onChange={(v) => update(o.id, { key: v })} placeholder="e.g. 24" />
                <IconBtn
                  title="Reveal other fields when this option is picked"
                  onClick={() => setOpenId(open ? null : o.id)}
                >
                  <ListTree size={14} className={hasCondition ? "text-teal-400" : ""} />
                </IconBtn>
                <IconBtn danger title="Remove option" onClick={() => remove(o.id)}><Trash2 size={14} /></IconBtn>
              </div>
              {open && (
                <div className="px-2 pb-2">
                  <Label>Reveal these field keys when "{o.value || "this option"}" is selected</Label>
                  <TagInput
                    values={o.dependentCustomFields || []}
                    onChange={(v) => update(o.id, { dependentCustomFields: v })}
                    placeholder="field key, press enter"
                    tone="teal"
                  />
                </div>
              )}
            </div>
          );
        })}
      </div>
      <button type="button" onClick={add} className="mt-2 text-xs font-medium text-teal-400 hover:text-teal-300 inline-flex items-center gap-1">
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
              className={`px-2.5 py-1 rounded-md text-xs font-mono border ${field.parentOptionType === t ? "border-teal-500 bg-teal-500 text-white" : "border-white/10 text-slate-500"}`}
            >
              {t}
            </button>
          ))}
        </div>
      </div>
      {options.map((p) => (
        <div key={p.id} className="rounded-lg border border-white/10 bg-white/[0.03] p-2.5 mb-2.5">
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
            <button type="button" onClick={() => addChild(p.id)} className="text-xs font-medium text-teal-400 hover:text-teal-300 inline-flex items-center gap-1 mt-1">
              <Plus size={12} /> Add child option
            </button>
          </div>
        </div>
      ))}
      <button type="button" onClick={addParent} className="text-xs font-medium text-teal-400 hover:text-teal-300 inline-flex items-center gap-1">
        <Plus size={13} /> Add parent option
      </button>
    </div>
  );
}

const MATRIX_COLUMN_TYPES = [
  { value: "text", label: "Text" },
  { value: "number", label: "Number" },
  { value: "date", label: "Date" },
  { value: "dropdown", label: "Dropdown" },
];

function MatrixColumnsEditor({ columns, onChange }) {
  const update = (id, patch) => onChange(columns.map((c) => (c.id === id ? { ...c, ...patch } : c)));
  const remove = (id) => onChange(columns.filter((c) => c.id !== id));
  const add = () => onChange([...columns, newMatrixColumn()]);
  return (
    <div>
      {columns.map((c) => (
        <div key={c.id} className="rounded-lg border border-white/10 bg-white/[0.03] p-2.5 mb-2.5">
          <div className="grid grid-cols-[1fr_1fr_140px_28px] gap-2 items-center mb-2">
            <Input value={c.label} onChange={(v) => update(c.id, { label: v })} placeholder="Column label, e.g. Rate" />
            <Input mono value={c.key} onChange={(v) => update(c.id, { key: v })} placeholder="columnKey" />
            <Select value={c.type} onChange={(v) => update(c.id, { type: v })} options={MATRIX_COLUMN_TYPES} />
            <IconBtn danger title="Remove column" onClick={() => remove(c.id)}><Trash2 size={14} /></IconBtn>
          </div>
          {c.type === "dropdown" && (
            <div className="pl-3 border-l-2 border-teal-500/20 mb-2">
              <SimpleOptionsEditor options={c.options || []} onChange={(opts) => update(c.id, { options: opts })} />
            </div>
          )}
          <Input mono value={c.formula} onChange={(v) => update(c.id, { formula: v })} placeholder="optional formula, e.g. rate*quantity" />
        </div>
      ))}
      <button type="button" onClick={add} className="text-xs font-medium text-teal-400 hover:text-teal-300 inline-flex items-center gap-1">
        <Plus size={13} /> Add column
      </button>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Custom field editor (used both at sub-type level and status level)  */
/* ------------------------------------------------------------------ */

function AdvancedFieldSettings({ field, onChange }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="mt-3 border-t border-white/5 pt-3">
      <button type="button" onClick={() => setOpen(!open)} className="text-xs font-medium text-slate-500 hover:text-slate-300 inline-flex items-center gap-1">
        {open ? <ChevronDown size={13} /> : <ChevronRight size={13} />} Advanced settings
      </button>
      {open && (
        <div className="mt-3 space-y-3">
          <div className="flex flex-wrap gap-6">
            <Toggle checked={field.isMandatory} onChange={(v) => onChange({ ...field, isMandatory: v })} label="Mandatory" />
            <Toggle checked={field.disabled} onChange={(v) => onChange({ ...field, disabled: v })} label="Disabled" />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <Label>Display label override</Label>
              <Input value={field.displayLabel} onChange={(v) => onChange({ ...field, displayLabel: v })} placeholder="optional — overrides Field label on screen" />
            </div>
            <div>
              <Label mono>Only show when field key equals…</Label>
              <Input mono value={field.dependentOn} onChange={(v) => onChange({ ...field, dependentOn: v })} placeholder="e.g. rootPettyCashRequired" />
            </div>
          </div>
          <div>
            <Label>Roles who can edit this field</Label>
            <TagInput values={field.editRoles || []} onChange={(v) => onChange({ ...field, editRoles: v })} placeholder="role id, press enter" tone="teal" />
          </div>
        </div>
      )}
    </div>
  );
}

function CustomFieldEditor({ field, onChange, onDelete, restrictToDropdown }) {
  const typeOptions = restrictToDropdown
    // ? FIELD_TYPES.filter((t) => t.value === "dropdown")
    ? FIELD_TYPES.filter((t) => t.value === "fileUpload" || t.value=="text" || t.value==="number" || t.value==="date" || t.value==="time" || t.value==="dropdown" || t.value==="multiplechoice" || t.value==="radio" || t.value==="casecadeDropdown" || t.value==="apiDropdown" || t.value==="matrix")
    : FIELD_TYPES;

  const changeType = (type) => {
    const base = { id: field.id, label: field.label, key: field.key, ...advancedDefaults() };
    onChange({ ...newField(type), ...base, type });
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

      {OPTION_BASED_TYPES.includes(field.type) && (
        <>
          <Label>Options</Label>
          <SimpleOptionsEditor options={field.options || []} onChange={(opts) => onChange({ ...field, options: opts })} />
        </>
      )}

      {field.type === "apiDropdown" && (
        <>
          <p className="text-xs text-slate-500 mb-3">Advanced type — options are fetched live from another ticket's data instead of a fixed list.</p>
          <div className="grid grid-cols-2 gap-3 mb-3">
            <div>
              <Label mono>Label key (path in API result)</Label>
              <Input mono value={field.labelKey} onChange={(v) => onChange({ ...field, labelKey: v })} placeholder="e.g. customFields.CorporateCode.value" />
            </div>
            <div>
              <Label mono>Value key (path in API result)</Label>
              <Input mono value={field.valueKey} onChange={(v) => onChange({ ...field, valueKey: v })} placeholder="e.g. customFields.CorporateCode.value" />
            </div>
          </div>
          <div className="mb-3">
            <Toggle checked={field.fillOnSelect} onChange={(v) => onChange({ ...field, fillOnSelect: v })} label="Auto-fill other fields on select" />
          </div>
          <Label mono>API filters (JSON)</Label>
          <textarea
            value={field.apiFiltersText}
            onChange={(e) => onChange({ ...field, apiFiltersText: e.target.value })}
            rows={5}
            className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-xs font-mono text-slate-100 outline-none focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20"
            placeholder={'{\n  "ticketType": "Customer Acquisition",\n  "status": "closed"\n}'}
          />
        </>
      )}

      {field.type === "matrix" && (
        <>
          <div className="mb-3">
            <Toggle checked={field.hideTotalColumn} onChange={(v) => onChange({ ...field, hideTotalColumn: v })} label="Hide total column" />
          </div>
          <Label>Columns</Label>
          <MatrixColumnsEditor columns={field.columns || []} onChange={(cols) => onChange({ ...field, columns: cols })} />
        </>
      )}

      {["date", "time", "text", "number", "fileUpload"].includes(field.type) && (
        <p className="text-xs text-slate-500 italic">No options needed for this field type.</p>
      )}

      <AdvancedFieldSettings field={field} onChange={onChange} />
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
        <div key={r.id} className="rounded-lg border border-amber-500/20 bg-amber-500/5 p-3 mb-2.5">
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
      <button type="button" onClick={add} className="text-xs font-medium text-amber-400 hover:text-amber-300 inline-flex items-center gap-1">
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
        <Label mono>Depends on other status codes</Label>
        <TagInput values={status.dependentStatus} onChange={(v) => onChange({ ...status, dependentStatus: v })} placeholder="status code, press enter" tone="teal" />
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
              <div className={`h-2.5 w-2.5 rounded-full ${s.status ? "bg-teal-500" : "bg-slate-700"}`} />
              <div className="mt-1.5 text-[11px] font-mono text-slate-500 whitespace-nowrap">{s.label || s.status || "…"}</div>
            </div>
            {i < statuses.length - 1 && <div className="h-px w-8 bg-slate-200 mb-4" />}
          </React.Fragment>
        ))}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Sub-type editor                                                     */
/* ------------------------------------------------------------------ */

function SubTypeEditor({ sub, onChange, onDelete, forceOpen }) {
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
      domId={`subtype-${sub.id}`}
      forceOpen={forceOpen}
      onDelete={onDelete}
    >
      <div className="flex gap-1 mb-4 border-b border-white/10">
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
          <div className="mb-3">
            <Label>Assignee roles — default routing for this sub-type</Label>
            <TagInput values={sub.assigneeRoles || []} onChange={(v) => onChange({ ...sub, assigneeRoles: v })} placeholder="e.g. HOHRPayroll — press enter" tone="teal" />
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
/* Bulk create sub-types from a template                               */
/* ------------------------------------------------------------------ */

function BulkCreatePanel({ templates, library, currentUid, onGenerate, onClose }) {
  const [text, setText] = useState("");
  const [sourceKind, setSourceKind] = useState(templates.length ? "template" : "subtype");
  const defaultTemplate = templates.find((t) => t.isDefault) || templates[0];
  const [templateId, setTemplateId] = useState(defaultTemplate ? defaultTemplate.id : "");
  const [subtypeRef, setSubtypeRef] = useState("");
  const [keyPrefix, setKeyPrefix] = useState("");
  const [reviewChoices, setReviewChoices] = useState({}); // name -> true (update) | false/undefined (no change)

  React.useEffect(() => {
    if (!templateId && defaultTemplate) setTemplateId(defaultTemplate.id);
  }, [defaultTemplate, templateId]);

  // Flatten every sub-type across every ticket type in the library into one
  // pickable list, so "use an existing sub-type" doesn't require a separate
  // template-import step. The currently open ticket type's own sub-types
  // are listed first since that's the most common source.
  const subtypeOptions = useMemo(() => {
    const opts = [];
    const ordered = [...library].sort((a, b) => (a._uid === currentUid ? -1 : b._uid === currentUid ? 1 : 0));
    ordered.forEach((d) => {
      d.ticketSubType.forEach((s) => {
        opts.push({
          value: `${d._uid}::${s.id}`,
          label: `${d.ticketType || d.ticketKey || "(untitled)"} → ${s.ticketType || s.ticketKey || "(untitled)"}`,
        });
      });
    });
    return opts;
  }, [library, currentUid]);

  React.useEffect(() => {
    if (!subtypeRef && subtypeOptions.length) setSubtypeRef(subtypeOptions[0].value);
  }, [subtypeOptions, subtypeRef]);

  const names = useMemo(() => {
    const parts = text.split(/[\n,]/).map((s) => s.trim()).filter(Boolean);
    const seen = new Set();
    return parts.filter((n) => {
      const k = n.toLowerCase();
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }, [text]);

  // Sub-types already sitting on the currently open ticket type — every
  // typed name is checked against these (by ticketType OR ticketKey,
  // case-insensitive) so we never silently create a duplicate.
  const currentSubtypes = useMemo(
    () => (library.find((d) => d._uid === currentUid) || {}).ticketSubType || [],
    [library, currentUid]
  );

  const categorized = useMemo(() => names.map((name) => {
    const norm = name.toLowerCase();
    const existing = currentSubtypes.find(
      (s) => xnorm(s.ticketType).toLowerCase() === norm || xnorm(s.ticketKey).toLowerCase() === norm
    );
    return { name, existing };
  }), [names, currentSubtypes]);

  const newOnes = categorized.filter((c) => !c.existing);
  const dupes = categorized.filter((c) => c.existing);
  const updateCount = dupes.filter((d) => reviewChoices[d.name]).length;

  const chosenTemplate = templates.find((t) => t.id === templateId);
  const chosenSubtypeOption = subtypeOptions.find((o) => o.value === subtypeRef);
  const hasAnySource = templates.length > 0 || subtypeOptions.length > 0;

  const handleGenerate = () => {
    if (!names.length) return;
    const sourceSpec = sourceKind === "template"
      ? (chosenTemplate ? { kind: "template", templateId } : null)
      : (subtypeRef ? { kind: "subtype", docUid: subtypeRef.split("::")[0], subId: subtypeRef.split("::")[1] } : null);
    if (!sourceSpec) return;

    const payload = {
      createNames: newOnes.map((c) => c.name),
      updateSubIds: dupes.filter((d) => reviewChoices[d.name]).map((d) => d.existing.id),
      skippedNames: dupes.filter((d) => !reviewChoices[d.name]).map((d) => d.name),
    };
    onGenerate(payload, sourceSpec, keyPrefix);
    setReviewChoices({});
  };

  return (
    <div className="rounded-xl border border-teal-500/30 bg-[#0D1424] p-4 mb-4">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2 text-sm font-semibold text-teal-300">
          <Sparkles size={15} /> Bulk create sub-types
        </div>
        <IconBtn onClick={onClose} title="Close"><X size={14} /></IconBtn>
      </div>

      {!hasAnySource ? (
        <p className="text-sm text-slate-400">
          Nothing to stamp out from yet. Either import a template on the <span className="text-teal-300 font-medium">Templates</span> tab,
          or add at least one sub-type first (here or on another ticket type) — either can be used as the source shape for bulk creation.
        </p>
      ) : (
        <>
          <div className="mb-3">
            <Label>Stamp out from</Label>
            <div className="flex gap-1.5">
              <button
                type="button"
                onClick={() => templates.length && setSourceKind("template")}
                disabled={!templates.length}
                className={`px-2.5 py-1.5 rounded-md text-xs font-medium border transition ${sourceKind === "template" ? "border-teal-500 bg-teal-500 text-white" : "border-white/10 text-slate-400"} ${!templates.length ? "opacity-40 cursor-not-allowed" : ""}`}
              >
                Saved template
              </button>
              <button
                type="button"
                onClick={() => subtypeOptions.length && setSourceKind("subtype")}
                disabled={!subtypeOptions.length}
                className={`px-2.5 py-1.5 rounded-md text-xs font-medium border transition ${sourceKind === "subtype" ? "border-teal-500 bg-teal-500 text-white" : "border-white/10 text-slate-400"} ${!subtypeOptions.length ? "opacity-40 cursor-not-allowed" : ""}`}
              >
                Existing sub-type
              </button>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-3 mb-3">
            <div>
              {sourceKind === "template" ? (
                <>
                  <Label>Template to stamp out</Label>
                  <Select
                    value={templateId}
                    onChange={setTemplateId}
                    options={templates.map((t) => ({ value: t.id, label: t.isDefault ? `${t.name} (default)` : t.name }))}
                  />
                </>
              ) : (
                <>
                  <Label>Sub-type to stamp out</Label>
                  <Select value={subtypeRef} onChange={setSubtypeRef} options={subtypeOptions} />
                </>
              )}
            </div>
            <div>
              <Label mono>Key prefix (optional)</Label>
              <Input mono value={keyPrefix} onChange={setKeyPrefix} placeholder="e.g. STK_" />
            </div>
          </div>
          <Label>Sub-type names — one per line, or comma separated</Label>
          <textarea
            value={text}
            onChange={(e) => setText(e.target.value)}
            rows={6}
            className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-sm text-slate-100 outline-none focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20 font-mono"
            placeholder={"DS Stock gaps - Store 101\nDS Stock gaps - Store 102\nDS Stock gaps - Store 103"}
          />
          <p className="text-xs text-slate-600 mt-1.5">
            Each new sub-type's key matches its name exactly (plus the prefix above, if set) — same convention the rest of the app uses.
          </p>

          {dupes.length > 0 && (
            <div className="mt-3 rounded-lg border border-amber-500/25 bg-amber-500/5 p-3">
              <div className="flex items-center gap-1.5 text-amber-400 text-xs font-semibold mb-2">
                <AlertTriangle size={13} />
                {dupes.length} of these already exist on this ticket type
              </div>
              <div className="flex items-center gap-2 mb-2">
                <button
                  type="button"
                  onClick={() => setReviewChoices(Object.fromEntries(dupes.map((d) => [d.name, true])))}
                  className="text-xs font-medium text-teal-400 hover:text-teal-300"
                >
                  Update all
                </button>
                <span className="text-slate-700">·</span>
                <button
                  type="button"
                  onClick={() => setReviewChoices({})}
                  className="text-xs font-medium text-slate-500 hover:text-slate-300"
                >
                  Leave all unchanged
                </button>
              </div>
              <div className="space-y-1.5">
                {dupes.map((d) => {
                  const willUpdate = !!reviewChoices[d.name];
                  return (
                    <div key={d.name} className="flex items-center justify-between gap-3 rounded-md bg-black/20 px-2.5 py-1.5">
                      <div className="min-w-0">
                        <div className="text-sm text-slate-200 truncate">{d.name}</div>
                        <div className="text-[11px] font-mono text-slate-500 truncate">already exists as {d.existing.ticketKey}</div>
                      </div>
                      <div className="flex gap-1 shrink-0">
                        <button
                          type="button"
                          onClick={() => setReviewChoices((c) => ({ ...c, [d.name]: false }))}
                          className={`px-2 py-1 rounded text-xs font-medium border transition ${!willUpdate ? "border-slate-500 bg-slate-600 text-white" : "border-white/10 text-slate-500"}`}
                        >
                          No change
                        </button>
                        <button
                          type="button"
                          onClick={() => setReviewChoices((c) => ({ ...c, [d.name]: true }))}
                          className={`px-2 py-1 rounded text-xs font-medium border transition ${willUpdate ? "border-teal-500 bg-teal-500 text-white" : "border-white/10 text-slate-500"}`}
                        >
                          Update
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          <div className="flex items-center justify-between mt-3 gap-4">
            <p className="text-xs text-slate-500">
              {newOnes.length} new sub-type{newOnes.length === 1 ? "" : "s"}
              {dupes.length > 0 && <>, {updateCount} update{updateCount === 1 ? "" : "s"}, {dupes.length - updateCount} left unchanged</>}
              {(sourceKind === "template" && chosenTemplate) || (sourceKind === "subtype" && chosenSubtypeOption) ? (
                <> — from <span className="text-teal-300">{sourceKind === "template" ? chosenTemplate.name : chosenSubtypeOption.label}</span></>
              ) : ""}.
            </p>
            <Btn variant="accent" onClick={handleGenerate}>
              <Sparkles size={14} />Generate
            </Btn>
          </div>
        </>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Templates tab                                                       */
/* ------------------------------------------------------------------ */

function TemplateImportBox({ onImport }) {
  const [pasteText, setPasteText] = useState("");
  const [name, setName] = useState("");
  const fileRef = React.useRef(null);

  const doImportText = () => {
    if (!pasteText.trim()) { window.alert("Paste some template JSON first."); return; }
    try {
      const parsed = JSON.parse(pasteText);
      onImport(parsed, name.trim());
      setPasteText("");
      setName("");
    } catch (e) {
      window.alert("That doesn't look like valid JSON — double check it and try again.");
    }
  };

  const doImportFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        onImport(parsed, name.trim() || file.name.replace(/\.json$/i, ""));
        setName("");
      } catch (err) {
        window.alert("That file isn't valid JSON — double check it and try again.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  return (
    <div className="rounded-xl border border-white/10 bg-[#111528] p-4 mb-5">
      <div className="flex items-center gap-2 text-sm font-semibold text-slate-100 mb-1">
        <LayoutTemplate size={15} className="text-teal-400" /> Import a template
      </div>
      <p className="text-xs text-slate-500 mb-3">
        Paste or upload a sub-type's JSON (its custom fields and status workflow). It becomes a reusable template —
        bulk creation fills in the ticket key and name per sub-type, everything else is cloned from here.
      </p>
      <div className="mb-3">
        <Label>Template name</Label>
        <Input value={name} onChange={setName} placeholder="e.g. Standard stock-gap subtype" />
      </div>
      <textarea
        value={pasteText}
        onChange={(e) => setPasteText(e.target.value)}
        rows={6}
        className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-xs font-mono text-slate-100 outline-none focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20 mb-3"
        placeholder='{ "ticketKey": "StockGap", "customFieldsMetaData": [...], "statusWorkFlow": [...] }'
      />
      <div className="flex items-center gap-2">
        <Btn variant="accent" small onClick={doImportText}><FileJson size={14} />Import pasted JSON</Btn>
        <input ref={fileRef} type="file" accept="application/json,.json" onChange={doImportFile} className="hidden" />
        <Btn variant="ghost" small onClick={() => fileRef.current && fileRef.current.click()}><Download size={14} />Upload JSON file</Btn>
      </div>
    </div>
  );
}

function TemplateFromLibraryBox({ library, onImport }) {
  const flat = useMemo(() => {
    const opts = [];
    library.forEach((d) => {
      d.ticketSubType.forEach((s) => {
        opts.push({
          value: `${d._uid}::${s.id}`,
          label: `${d.ticketType || d.ticketKey || "(untitled)"} → ${s.ticketType || s.ticketKey || "(untitled)"}`,
          sub: s,
        });
      });
    });
    return opts;
  }, [library]);

  const [ref, setRef] = useState(flat[0]?.value || "");
  const [name, setName] = useState("");

  React.useEffect(() => {
    if (!ref && flat.length) setRef(flat[0].value);
  }, [flat, ref]);

  if (!flat.length) return null;

  const chosen = flat.find((o) => o.value === ref);

  const doImport = () => {
    if (!chosen) return;
    const [docUid, subId] = chosen.value.split("::");
    onImport(docUid, subId, name.trim() || chosen.sub.ticketType || chosen.sub.ticketKey);
  };

  return (
    <div className="rounded-xl border border-white/10 bg-[#111528] p-4 mb-5">
      <div className="flex items-center gap-2 text-sm font-semibold text-slate-100 mb-1">
        <ListTree size={15} className="text-teal-400" /> Or turn an existing sub-type into a template
      </div>
      <p className="text-xs text-slate-500 mb-3">
        Pick any sub-type already in your library — its custom fields and status workflow become the template, no JSON copy/paste needed.
      </p>
      <div className="grid grid-cols-2 gap-3 mb-3">
        <div>
          <Label>Sub-type</Label>
          <Select value={ref} onChange={setRef} options={flat.map((o) => ({ value: o.value, label: o.label }))} />
        </div>
        <div>
          <Label>Template name</Label>
          <Input value={name} onChange={setName} placeholder={chosen ? (chosen.sub.ticketType || chosen.sub.ticketKey) : "template name"} />
        </div>
      </div>
      <Btn variant="accent" small onClick={doImport}><LayoutTemplate size={14} />Save as template</Btn>
    </div>
  );
}

function TemplateCard({ template, onChange, onDelete, onDuplicate, onSetDefault }) {
  const [tab, setTab] = useState("basic");
  const data = template.data;
  const updateData = (patch) => onChange({ ...template, data: { ...data, ...patch } });
  const tabs = [
    { id: "basic", label: "Basic" },
    { id: "fields", label: `Fields (${(data.customFieldsMetaData || []).length})` },
    { id: "workflow", label: `Workflow (${(data.statusWorkFlow || []).length})` },
  ];
  return (
    <Accordion
      title={template.name || "(untitled template)"}
      subtitle={`${(data.customFieldsMetaData || []).length} field(s) · ${(data.statusWorkFlow || []).length} status(es)`}
      tone={template.isDefault ? "amber" : "slate"}
      onDelete={onDelete}
    >
      <div className="flex items-center gap-2 mb-3">
        <button
          type="button"
          onClick={onSetDefault}
          className={`inline-flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium border transition ${template.isDefault ? "border-amber-400 bg-amber-500/10 text-amber-300" : "border-white/10 text-slate-400 hover:border-white/25"}`}
        >
          <Star size={13} className={template.isDefault ? "fill-amber-400 text-amber-400" : ""} />
          {template.isDefault ? "Default template" : "Make default"}
        </button>
        <Btn small variant="ghost" onClick={onDuplicate}><Copy size={13} />Duplicate</Btn>
      </div>

      <div className="flex gap-1 mb-4 border-b border-white/10">
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
          <div className="mb-3">
            <Label>Template name</Label>
            <Input value={template.name} onChange={(v) => onChange({ ...template, name: v })} placeholder="e.g. Standard stock-gap subtype" />
          </div>
          <div className="mb-3">
            <Label>Description</Label>
            <Input value={template.description} onChange={(v) => onChange({ ...template, description: v })} placeholder="optional note for your team" />
          </div>
          <div className="mb-3">
            <Label mono>Default prefix pattern</Label>
            <Input mono value={data.prefix} onChange={(v) => updateData({ prefix: v })} placeholder="e.g. ${locationId}" />
          </div>
          <Toggle checked={!!data.isAutoEscalation} onChange={(v) => updateData({ isAutoEscalation: v })} label="Auto-escalation enabled by default" />
        </div>
      )}

      {tab === "fields" && (
        <CustomFieldsList fields={data.customFieldsMetaData || []} onChange={(fs) => updateData({ customFieldsMetaData: fs })} />
      )}

      {tab === "workflow" && (
        <div>
          <StatusPipelinePreview statuses={data.statusWorkFlow || []} />
          {(data.statusWorkFlow || []).map((s) => (
            <StatusEditor
              key={s.id}
              status={s}
              onChange={(patch) => updateData({ statusWorkFlow: (data.statusWorkFlow || []).map((x) => (x.id === s.id ? patch : x)) })}
              onDelete={() => updateData({ statusWorkFlow: (data.statusWorkFlow || []).filter((x) => x.id !== s.id) })}
            />
          ))}
          <Btn small variant="dashed" onClick={() => updateData({ statusWorkFlow: [...(data.statusWorkFlow || []), newStatus()] })}>
            <Plus size={14} />Add status
          </Btn>
        </div>
      )}
    </Accordion>
  );
}

function TemplatesTab({ templates, library, onImport, onImportFromSubType, onCreateBlank, onChange, onDelete, onDuplicate, onSetDefault }) {
  return (
    <div>
      <TemplateImportBox onImport={onImport} />
      <TemplateFromLibraryBox library={library} onImport={onImportFromSubType} />
      <div className="flex items-center justify-between mb-3">
        <p className="text-xs text-slate-500">
          Templates are reused across every ticket type — pick one when bulk-creating sub-types.
        </p>
        <Btn small variant="dashed" onClick={onCreateBlank}><Plus size={14} />New blank template</Btn>
      </div>
      {templates.length === 0 ? (
        <div className="text-center py-10 text-slate-500 text-sm">
          No templates yet — import one above, or start from a blank template, to begin stamping out sub-types in bulk.
        </div>
      ) : (
        templates.map((t) => (
          <TemplateCard
            key={t.id}
            template={t}
            onChange={(patch) => onChange(t.id, patch)}
            onDelete={() => onDelete(t.id)}
            onDuplicate={() => onDuplicate(t.id)}
            onSetDefault={() => onSetDefault(t.id)}
          />
        ))
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Smart update — describe changes in plain English, review the        */
/* proposed diff, apply only what you approve.                         */
/* ------------------------------------------------------------------ */

function opTone(op, preview) {
  if (op.op === "note") return "amber";
  if (preview && !preview.applied) return "amber";
  if (op.op.toLowerCase().startsWith("remove")) return "rose";
  return "teal";
}

const OP_TONE_CLASSES = {
  teal: "border-teal-500/30 bg-teal-500/5",
  rose: "border-rose-500/30 bg-rose-500/5",
  amber: "border-amber-500/30 bg-amber-500/5",
};

function ChangeOpsReview({ ops, checked, previews, onToggle, onApply }) {
  const selectedCount = ops.filter((_, i) => checked[i]).length;
  return (
    <div className="rounded-xl border border-white/10 bg-[#111528] p-4">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2 text-sm font-semibold text-slate-100">
          <Diff size={15} className="text-teal-400" /> Proposed changes ({ops.length})
        </div>
        <Btn variant="accent" small onClick={onApply}>
          <Check size={14} />Apply {selectedCount} selected
        </Btn>
      </div>
      <div className="space-y-2">
        {ops.map((op, i) => {
          const preview = previews[i];
          const tone = opTone(op, preview);
          const disabled = op.op === "note" || (preview && !preview.applied);
          return (
            <label
              key={i}
              className={`flex items-start gap-3 rounded-lg border p-3 ${OP_TONE_CLASSES[tone]} ${disabled ? "opacity-60 cursor-not-allowed" : "cursor-pointer"}`}
            >
              <input
                type="checkbox"
                checked={!!checked[i] && !disabled}
                disabled={disabled}
                onChange={() => onToggle(i)}
                className="mt-0.5"
              />
              <div className="min-w-0 flex-1">
                <div className="flex items-center gap-2 flex-wrap mb-0.5">
                  <span className="text-xs font-mono text-slate-500">{op.op}</span>
                  {op.ticketKey && <span className="text-xs font-mono text-teal-400">{op.ticketKey}</span>}
                  {op.subTypeKey && <span className="text-xs font-mono text-slate-500">/ {op.subTypeKey}</span>}
                  {op.statusCode && <span className="text-xs font-mono text-slate-500">/ {op.statusCode}</span>}
                </div>
                <div className="text-sm text-slate-200">{op.description || "(no description)"}</div>
                {preview && !preview.applied && (
                  <div className="text-xs text-amber-400 mt-1">{preview.message}</div>
                )}
              </div>
            </label>
          );
        })}
      </div>
    </div>
  );
}

function ExcelDiffPanel({ library, onApply, onExport }) {
  const [fileName, setFileName] = useState("");
  const [targetLabel, setTargetLabel] = useState("");
  const [ops, setOps] = useState(null);
  const [checked, setChecked] = useState({});
  const [error, setError] = useState("");
  const fileRef = React.useRef(null);

  const previews = useMemo(() => {
    if (!ops) return {};
    const out = {};
    ops.forEach((op, i) => { out[i] = applyOpToLibrary(library, op); });
    return out;
  }, [ops, library]);

  const handleFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setError(""); setOps(null);
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = XLSX.read(reader.result, { type: "array" });
        const { docs: newDocs, missingSheets } = excelWorkbookToDocs(wb);
        const withKeys = newDocs.filter((d) => d.ticketKey);
        if (!withKeys.length) { setError("Couldn't find a Ticket Key in that sheet — check the Ticket Details tab of the workbook."); return; }

        const allOps = [];
        const matchedLabels = [];
        const unmatched = [];
        withKeys.forEach((newDoc) => {
          const oldDoc = library.find((d) => xnorm(d.ticketKey).toLowerCase() === xnorm(newDoc.ticketKey).toLowerCase());
          if (!oldDoc) { unmatched.push(newDoc.ticketType || newDoc.ticketKey); return; }
          matchedLabels.push(oldDoc.ticketType || oldDoc.ticketKey);
          allOps.push(...diffDocsToOps(oldDoc, newDoc));
        });

        if (!matchedLabels.length) {
          setError(`No existing ticket type here matches any Ticket Key in that sheet (${withKeys.map((d) => d.ticketKey).join(", ")}) — use "Import Excel" in the header instead to add them as new.`);
          return;
        }
        setTargetLabel(matchedLabels.join(", "));
        if (missingSheets.length) allOps.push({ op: "note", ticketKey: matchedLabels[0], description: `Sheet(s) not found in the upload: ${missingSheets.join(", ")} — that section wasn't compared.` });
        unmatched.forEach((label) => allOps.push({ op: "note", description: `"${label}" in the sheet doesn't match any ticket type already here — use "Import Excel" in the header to add it as new instead.` }));
        setOps(allOps);
        const initChecked = {};
        allOps.forEach((op, i) => { initChecked[i] = op.op !== "note"; });
        setChecked(initChecked);
      } catch (err) {
        setError("Couldn't read that workbook — make sure it's the sheet exported from this app (or matches its structure).");
      }
    };
    reader.readAsArrayBuffer(file);
    setFileName(file.name);
    e.target.value = "";
  };

  const apply = () => {
    if (!ops) return;
    const selected = ops.filter((_, i) => checked[i] && previews[i] && previews[i].applied);
    if (!selected.length) return;
    onApply(selected);
    setOps(null); setFileName("");
  };

  return (
    <div>
      <div className="rounded-xl border border-white/10 bg-[#111528] p-4 mb-5">
        <div className="flex items-center gap-2 text-sm font-semibold text-slate-100 mb-1">
          <FileSpreadsheet size={15} className="text-teal-400" /> Upload an edited Excel sheet
        </div>
        <p className="text-xs text-slate-500 mb-4">
          Works with no extra instructions or context from whoever makes the edits. Export the current sheet (one ticket
          type or all of them), hand it off, then upload whatever comes back here — every difference from what's
          already in the app is detected automatically, across every ticket type the sheet contains, and listed below
          for review before anything changes.
        </p>
        <div className="flex items-center gap-2">
          <Btn variant="ghost" onClick={onExport}><Download size={14} />Export all as Excel</Btn>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={handleFile} className="hidden" />
          <Btn variant="accent" onClick={() => fileRef.current && fileRef.current.click()}>
            <FileSpreadsheet size={14} />{fileName || "Upload edited sheet…"}
          </Btn>
        </div>
        {error && <p className="text-xs text-rose-400 mt-3">{error}</p>}
      </div>

      {ops && ops.length === 0 && (
        <p className="text-sm text-slate-500">No differences found against "{targetLabel}" — the sheet matches what's already here.</p>
      )}
      {ops && ops.length > 0 && (
        <ChangeOpsReview
          ops={ops}
          checked={checked}
          previews={previews}
          onToggle={(i) => setChecked((c) => ({ ...c, [i]: !c[i] }))}
          onApply={apply}
        />
      )}
    </div>
  );
}

function AiUpdatePanel({ library, onApply }) {
  const [scopeKey, setScopeKey] = useState("__all__");
  const [reqText, setReqText] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [ops, setOps] = useState(null);
  const [checked, setChecked] = useState({});

  const scopeOptions = [
    { value: "__all__", label: `All ticket types (${library.length})` },
    ...library.map((d) => ({ value: d._uid, label: d.ticketType || d.ticketKey || "(untitled)" })),
  ];
  const scopeDocs = scopeKey === "__all__" ? library : library.filter((d) => d._uid === scopeKey);

  const previews = useMemo(() => {
    if (!ops) return {};
    const out = {};
    ops.forEach((op, i) => { out[i] = applyOpToLibrary(library, op); });
    return out;
  }, [ops, library]);

  const analyze = async () => {
    if (!reqText.trim()) { setError("Describe what should change first."); return; }
    if (!scopeDocs.length) { setError("Nothing to analyze — import or create a ticket type first."); return; }
    setLoading(true); setError(""); setOps(null);
    try {
      const scopeJson = JSON.stringify(scopeDocs.map((d) => docToJson(d)).filter(Boolean), null, 2);
      const resp = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 4000,
          system: SMART_UPDATE_SYSTEM_PROMPT,
          messages: [{ role: "user", content: `CURRENT JSON:\n${scopeJson}\n\nREQUIREMENTS:\n${reqText}` }],
        }),
      });
      const data = await resp.json();
      const textBlocks = (data.content || []).filter((c) => c.type === "text").map((c) => c.text).join("\n");
      const cleaned = textBlocks.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(cleaned);
      if (!Array.isArray(parsed)) throw new Error("not an array");
      setOps(parsed);
      const initChecked = {};
      parsed.forEach((op, i) => { initChecked[i] = op.op !== "note"; });
      setChecked(initChecked);
    } catch (e) {
      setError("Couldn't get a clean response — try narrowing the scope, or rephrasing the requirements a bit more explicitly.");
    } finally {
      setLoading(false);
    }
  };

  const apply = () => {
    if (!ops) return;
    const selected = ops.filter((_, i) => checked[i] && previews[i] && previews[i].applied);
    if (!selected.length) return;
    onApply(selected);
    setOps(null);
    setReqText("");
  };

  return (
    <div>
      <div className="rounded-xl border border-white/10 bg-[#111528] p-4 mb-5">
        <div className="flex items-center gap-2 text-sm font-semibold text-slate-100 mb-1">
          <Wand2 size={15} className="text-teal-400" /> Describe changes in plain English
        </div>
        <p className="text-xs text-slate-500 mb-4">
          Faster when you already know exactly what to say, but it can only act on what you write — if the wording is
          vague or leaves something out, it'll either skip it or flag it as a "note" for you to handle manually rather
          than guess. For changes anyone can make with zero setup, use "Upload edited Excel sheet" instead.
        </p>
        <div className="mb-3">
          <Label>Scope</Label>
          <Select value={scopeKey} onChange={setScopeKey} options={scopeOptions} />
        </div>
        <Label>What should change?</Label>
        <textarea
          value={reqText}
          onChange={(e) => setReqText(e.target.value)}
          rows={7}
          className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-sm text-slate-100 outline-none focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20 mb-3"
          placeholder={"e.g.\n- On sub-type StockGap, remove assignee role Str_Mgr and add Regional_Mgr\n- Add a new status \"On Hold\" (code onHold) to StockGap, visible to viewers, settable by Dist_Mgr\n- On the On Hold status, add a mandatory text field \"Remark\" (key remark)"}
        />
        {error && <p className="text-xs text-rose-400 mb-3">{error}</p>}
        <Btn variant="accent" onClick={analyze}>
          <Wand2 size={14} />{loading ? "Analyzing…" : "Analyze changes"}
        </Btn>
      </div>

      {ops && ops.length === 0 && (
        <p className="text-sm text-slate-500">No changes were proposed — try being more specific about what should change.</p>
      )}

      {ops && ops.length > 0 && (
        <ChangeOpsReview
          ops={ops}
          checked={checked}
          previews={previews}
          onToggle={(i) => setChecked((c) => ({ ...c, [i]: !c[i] }))}
          onApply={apply}
        />
      )}
    </div>
  );
}

function SmartUpdateTab({ library, onApply, onExportExcel }) {
  const [mode, setMode] = useState("excel");
  return (
    <div className="max-w-3xl">
      <div className="flex gap-1 mb-4 border-b border-white/10">
        {[
          { id: "excel", label: "Upload edited Excel sheet" },
          { id: "ai", label: "Describe in plain English (AI)" },
        ].map((m) => (
          <button
            key={m.id}
            type="button"
            onClick={() => setMode(m.id)}
            className={`px-3 py-2 text-sm font-medium border-b-2 -mb-px transition ${mode === m.id ? "border-teal-400 text-teal-300" : "border-transparent text-slate-500 hover:text-slate-300"}`}
          >
            {m.label}
          </button>
        ))}
      </div>
      {mode === "excel"
        ? <ExcelDiffPanel library={library} onApply={onApply} onExport={onExportExcel} />
        : <AiUpdatePanel library={library} onApply={onApply} />}
    </div>
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
      <pre className="rounded-xl bg-slate-900 text-slate-100 text-xs leading-relaxed p-4 overflow-auto font-mono max-h-[70vh]">
        {json}
      </pre>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Library — searchable list of every ticket type + sub-type            */
/* ------------------------------------------------------------------ */

function matchesTerm(text, term) {
  return xnorm(text).toLowerCase().includes(term.toLowerCase());
}

function LibraryTab({ library, selectedUid, onSelectType, onSelectSubtype, onNew, onDuplicate, onDelete, onExportAll }) {
  const [term, setTerm] = useState("");
  const [expanded, setExpanded] = useState({});

  const rows = library.map((d) => {
    const selfMatch = !term || matchesTerm(d.ticketType, term) || matchesTerm(d.ticketKey, term);
    const subMatches = term ? d.ticketSubType.filter((s) => matchesTerm(s.ticketType, term) || matchesTerm(s.ticketKey, term)) : d.ticketSubType;
    const show = !term || selfMatch || subMatches.length > 0;
    return { d, selfMatch, subMatches, show };
  }).filter((r) => r.show);

  return (
    <div>
      <div className="flex items-center gap-2 mb-4">
        <div className="relative flex-1">
          <input
            value={term}
            onChange={(e) => setTerm(e.target.value)}
            placeholder="Search ticket types or sub-types…"
            className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] pl-3 pr-3 py-2.5 text-sm text-slate-100 placeholder:text-slate-600 outline-none focus:border-teal-400 focus:ring-2 focus:ring-teal-400/20"
          />
        </div>
        <Btn variant="ghost" onClick={onExportAll}><Download size={15} />Export all ({library.length})</Btn>
        <Btn variant="accent" onClick={onNew}><Plus size={15} />New ticket type</Btn>
      </div>

      {rows.length === 0 && (
        <div className="text-center py-16 text-slate-500 text-sm">
          {library.length === 0 ? "No ticket types yet — import a file or create one." : "No matches for that search."}
        </div>
      )}

      <div className="space-y-2">
        {rows.map(({ d, subMatches }) => {
          const isSelected = d._uid === selectedUid;
          const isOpen = term ? subMatches.length > 0 : !!expanded[d._uid];
          return (
            <div key={d._uid} className={`rounded-xl border bg-[#111528] overflow-hidden ${isSelected ? "border-teal-500/50" : "border-white/10"}`}>
              <div className="flex items-center gap-3 px-4 py-3">
                <button
                  onClick={() => setExpanded((e) => ({ ...e, [d._uid]: !e[d._uid] }))}
                  className="text-slate-500 hover:text-slate-300"
                >
                  {isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                </button>
                <button className="flex-1 min-w-0 text-left" onClick={() => onSelectType(d._uid)}>
                  <div className="text-sm font-medium text-slate-100 truncate">{d.ticketType || "(untitled ticket type)"}</div>
                  <div className="text-xs font-mono text-slate-500 truncate">{d.ticketKey || "no key"} · {d.ticketSubType.length} sub-type{d.ticketSubType.length === 1 ? "" : "s"}</div>
                </button>
                <IconBtn title="Duplicate" onClick={() => onDuplicate(d._uid)}><Copy size={14} /></IconBtn>
                <IconBtn danger title="Delete ticket type" onClick={() => onDelete(d._uid)}><Trash2 size={14} /></IconBtn>
              </div>
              {isOpen && (
                <div className="border-t border-white/5 divide-y divide-white/5">
                  {d.ticketSubType.length === 0 && (
                    <div className="px-4 py-2.5 text-xs text-slate-600 italic">No sub-types yet.</div>
                  )}
                  {(term ? subMatches : d.ticketSubType).map((s) => (
                    <button
                      key={s.id}
                      onClick={() => onSelectSubtype(d._uid, s.id)}
                      className="w-full flex items-center gap-2 px-4 py-2.5 pl-11 text-left hover:bg-white/5 transition"
                    >
                      <ListTree size={13} className="text-slate-600 shrink-0" />
                      <span className="text-sm text-slate-300 truncate">{s.ticketType || "(untitled sub-type)"}</span>
                      <span className="text-xs font-mono text-slate-600 truncate">{s.ticketKey}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          );
        })}
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

/* Merge an incoming parsed doc (from JSON or Excel import) into whatever
   is currently open, instead of wiping it. Sub-types are matched by
   ticketKey (case-insensitive): a match gets replaced in place, anything
   new gets appended. Top-level fields (name/tenant/roles) are only filled
   in when the current ones are still blank, so re-importing never clobbers
   details you've already typed. */
function upsertSubtypes(currentList, incomingList) {
  const result = [...currentList];
  const added = [];
  const updated = [];
  incomingList.forEach((inc) => {
    const incKey = xnorm(inc.ticketKey).toLowerCase();
    const idx = incKey ? result.findIndex((s) => xnorm(s.ticketKey).toLowerCase() === incKey) : -1;
    if (idx >= 0) {
      result[idx] = { ...inc, id: result[idx].id };
      updated.push(inc.ticketType || inc.ticketKey || "(unnamed)");
    } else {
      result.push(inc);
      added.push(inc.ticketType || inc.ticketKey || "(unnamed)");
    }
  });
  return { result, added, updated };
}

function isDocEmpty(doc) {
  return !doc.ticketKey && !doc.ticketType && doc.ticketSubType.length === 0;
}

function mergeIncomingDoc(current, incoming) {
  if (isDocEmpty(current)) {
    return { doc: incoming, added: incoming.ticketSubType.map((s) => s.ticketType || s.ticketKey), updated: [] };
  }
  const { result, added, updated } = upsertSubtypes(current.ticketSubType, incoming.ticketSubType);
  return {
    doc: {
      ticketKey: current.ticketKey || incoming.ticketKey,
      ticketType: current.ticketType || incoming.ticketType,
      tenantId: current.tenantId || incoming.tenantId,
      deleted: current.deleted,
      creators: current.creators.length ? current.creators : incoming.creators,
      viewers: current.viewers.length ? current.viewers : incoming.viewers,
      assignee: current.assignee.length ? current.assignee : incoming.assignee,
      ticketSubType: result,
    },
    added, updated,
  };
}

/* A raw import might be one ticket type, a bare array of them, or the
   { result: [...] } shape your real export uses. Normalize to an array
   of raw ticket-type objects before converting each to editable state. */
function extractRawDocs(parsed) {
  if (Array.isArray(parsed)) return parsed;
  if (parsed && Array.isArray(parsed.result)) return parsed.result;
  return [parsed];
}

/* Upsert a whole batch of incoming ticket types into the library, matched
   by ticketKey (case-insensitive). New ticket types are appended; existing
   ones get their sub-types merged via mergeIncomingDoc (never wiped). */
function importDocsIntoLibrary(library, incomingDocs) {
  let nextLibrary = [...library];
  const newTypes = [];
  const updatedTypes = [];
  incomingDocs.forEach((inc) => {
    const incKey = xnorm(inc.ticketKey).toLowerCase();
    const idx = incKey ? nextLibrary.findIndex((d) => xnorm(d.ticketKey).toLowerCase() === incKey) : -1;
    if (idx >= 0) {
      const { doc: merged, added, updated } = mergeIncomingDoc(nextLibrary[idx], inc);
      nextLibrary[idx] = { ...merged, _uid: nextLibrary[idx]._uid };
      const label = inc.ticketType || inc.ticketKey || "(unnamed)";
      if (added.length || updated.length) {
        updatedTypes.push(`${label} (+${added.length} sub-type${added.length === 1 ? "" : "s"} added, ${updated.length} updated)`);
      }
    } else {
      nextLibrary.push({ ...inc, _uid: uid() });
      newTypes.push(inc.ticketType || inc.ticketKey || "(unnamed)");
    }
  });
  return { library: nextLibrary, newTypes, updatedTypes };
}

export default function App({ initialData } = {}) {
  const initialLibrary = useMemo(() => {
    if (!initialData) return [];
    return extractRawDocs(initialData).map((raw) => ({ ...jsonToDoc(raw), _uid: uid() }));
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const [library, setLibrary] = useState(initialLibrary);
  const [selectedUid, setSelectedUid] = useState(initialLibrary[0]?._uid || null);
  const [tab, setTab] = useState(initialLibrary.length ? "details" : "library");
  const [importNotice, setImportNotice] = useState(null); // { title, warnings, missingSheets } | null
  const [focusSubtypeId, setFocusSubtypeId] = useState(null);
  const [templates, setTemplates] = useState([]);
  const [bulkOpen, setBulkOpen] = useState(false);
  const fileInputRef = React.useRef(null);
  const excelInputRef = React.useRef(null);

  const selectedIndex = library.findIndex((d) => d._uid === selectedUid);
  const selected = selectedIndex >= 0 ? library[selectedIndex] : null;

  const updateSelected = (patch) => {
    if (selectedIndex < 0) return;
    setLibrary(library.map((d, i) => (i === selectedIndex ? { ...patch, _uid: d._uid } : d)));
  };

  const selectType = (uidVal) => { setSelectedUid(uidVal); setFocusSubtypeId(null); setTab("details"); };
  const selectSubtype = (uidVal, subId) => { setSelectedUid(uidVal); setFocusSubtypeId(subId); setTab("subtypes"); };

  React.useEffect(() => {
    if (tab === "subtypes" && focusSubtypeId) {
      const t = setTimeout(() => {
        document.getElementById(`subtype-${focusSubtypeId}`)?.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 60);
      return () => clearTimeout(t);
    }
  }, [tab, focusSubtypeId, selectedUid]);

  const handleImportFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const rawDocs = extractRawDocs(JSON.parse(reader.result));
        const incomingDocs = rawDocs.map(jsonToDoc);
        const { library: nextLibrary, newTypes, updatedTypes } = importDocsIntoLibrary(library, incomingDocs);
        setLibrary(nextLibrary);
        const notes = [];
        if (newTypes.length) notes.push(`New ticket type(s): ${newTypes.join(", ")}`);
        if (updatedTypes.length) notes.push(...updatedTypes.map((t) => `Updated: ${t}`));
        setImportNotice(notes.length ? { title: `JSON imported (${incomingDocs.length} ticket type${incomingDocs.length === 1 ? "" : "s"})`, warnings: notes, missingSheets: [] } : null);
        if (!selectedUid && nextLibrary.length) setSelectedUid(nextLibrary[0]._uid);
        setTab("library");
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
        const { docs: parsedDocs, warnings, missingSheets } = excelWorkbookToDocs(wb);
        const { library: nextLibrary, newTypes, updatedTypes } = importDocsIntoLibrary(library, parsedDocs);
        setLibrary(nextLibrary);
        const notes = [...warnings];
        if (newTypes.length) notes.push(`New ticket type(s): ${newTypes.join(", ")}`);
        if (updatedTypes.length) notes.push(...updatedTypes.map((t) => `Updated: ${t}`));
        setImportNotice((notes.length || missingSheets.length) ? { title: `Excel imported (${parsedDocs.length} ticket type${parsedDocs.length === 1 ? "" : "s"})`, warnings: notes, missingSheets } : null);
        if (!selectedUid && nextLibrary.length) setSelectedUid(nextLibrary[0]._uid);
        setTab("library");
      } catch (err) {
        window.alert("Couldn't read that workbook — check it matches the expected sheet structure and try again.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const exportSelectedToExcel = () => {
    if (!selected) { window.alert("Pick a ticket type first (from Ticket Types) before exporting."); return; }
    const { wb, skippedFields } = excelWorkbookFromDocs([selected]);
    XLSX.writeFile(wb, `${selected.ticketKey || "ticket-type"}.xlsx`);
    if (skippedFields.length) {
      setImportNotice({
        title: "Excel exported",
        warnings: [`${skippedFields.length} field(s) use types the Excel template can't represent and were left out: ${skippedFields.join("; ")}. Edit those directly in the web app instead.`],
        missingSheets: [],
      });
    }
  };

  const exportAllToExcel = () => {
    if (!library.length) { window.alert("Nothing to export yet — import or create a ticket type first."); return; }
    const { wb, skippedFields } = excelWorkbookFromDocs(library);
    XLSX.writeFile(wb, "ticket-types.xlsx");
    if (skippedFields.length) {
      setImportNotice({
        title: "Excel exported",
        warnings: [`${skippedFields.length} field(s) use types the Excel template can't represent and were left out: ${skippedFields.join("; ")}. Edit those directly in the web app instead.`],
        missingSheets: [],
      });
    }
  };

  const exportAll = () => {
    const json = JSON.stringify({ result: library.map(docToJson) }, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "ticket-types.json";
    a.click();
    URL.revokeObjectURL(url);
  };

  const newTicketType = () => {
    const created = { ...emptyDoc(), _uid: uid() };
    setLibrary([...library, created]);
    setSelectedUid(created._uid);
    setTab("details");
  };

  const duplicateType = (uidVal) => {
    const src = library.find((d) => d._uid === uidVal);
    if (!src) return;
    const copy = jsonToDoc(docToJson(src) || {});
    copy._uid = uid();
    copy.ticketKey = src.ticketKey ? `${src.ticketKey}_copy` : "";
    copy.ticketType = src.ticketType ? `${src.ticketType} (copy)` : "";
    // it's a new entity — don't carry over backend-only passthrough fields
    // (e.g. _id) from the source, those belong to the original record only
    copy._passthrough = {};
    copy.ticketSubType = copy.ticketSubType.map((s) => ({ ...s, _passthrough: {} }));
    setLibrary([...library, copy]);
    setSelectedUid(copy._uid);
    setTab("details");
  };

  const deleteType = (uidVal) => {
    const target = library.find((d) => d._uid === uidVal);
    if (!target) return;
    if (!window.confirm(`Delete ticket type "${target.ticketType || target.ticketKey || "(unnamed)"}" and all its sub-types? This can't be undone.`)) return;
    setLibrary(library.filter((d) => d._uid !== uidVal));
    if (selectedUid === uidVal) { setSelectedUid(null); setTab("library"); }
  };

  /* -- templates -- */
  const importTemplate = (raw, name) => {
    const tpl = templateFromRawJson(raw, name);
    setTemplates((prev) => [...prev, { ...tpl, isDefault: prev.length === 0 }]);
  };
  const importTemplateFromSubType = (docUid, subId, name) => {
    const doc = library.find((d) => d._uid === docUid);
    const sub = doc && doc.ticketSubType.find((s) => s.id === subId);
    if (!sub) return;
    const tpl = {
      id: uid(),
      name: (name && name.trim()) || sub.ticketType || sub.ticketKey || "Imported template",
      description: "",
      isDefault: false,
      data: templateDataFromSubType(sub),
    };
    setTemplates((prev) => [...prev, { ...tpl, isDefault: prev.length === 0 }]);
  };
  const createBlankTemplate = () => {
    setTemplates((prev) => [...prev, { ...newTemplate(), isDefault: prev.length === 0 }]);
  };
  const updateTemplate = (id, patch) => setTemplates((prev) => prev.map((t) => (t.id === id ? patch : t)));
  const deleteTemplate = (id) => {
    const t = templates.find((x) => x.id === id);
    if (!t) return;
    if (!window.confirm(`Delete template "${t.name || "(untitled)"}"? This can't be undone.`)) return;
    setTemplates((prev) => {
      const next = prev.filter((x) => x.id !== id);
      if (t.isDefault && next.length) next[0] = { ...next[0], isDefault: true };
      return next;
    });
  };
  const duplicateTemplate = (id) => {
    const src = templates.find((t) => t.id === id);
    if (!src) return;
    const copy = { ...regenerateIds(src), id: uid(), name: `${src.name} (copy)`, isDefault: false };
    setTemplates((prev) => [...prev, copy]);
  };
  const setDefaultTemplate = (id) => setTemplates((prev) => prev.map((t) => ({ ...t, isDefault: t.id === id })));

  /* -- smart update: apply a batch of AI-proposed ops, log what happened -- */
  const applySmartUpdate = (selectedOps) => {
    let lib = library;
    const messages = [];
    selectedOps.forEach((op) => {
      const result = applyOpToLibrary(lib, op);
      if (result.applied) { lib = result.library; messages.push(result.message); }
      else messages.push(`Skipped: ${result.message}`);
    });
    setLibrary(lib);
    setImportNotice({
      title: `Smart update applied (${messages.length} change${messages.length === 1 ? "" : "s"})`,
      warnings: messages,
      missingSheets: [],
    });
  };

  /* -- bulk sub-type creation: source is either a saved template, or an
     existing sub-type picked live from the library (no template import
     step required). `payload` has already been split by BulkCreatePanel
     into genuinely-new names vs. names that matched an existing sub-type
     (where the user explicitly opted in per-item to update or skip) -- */
  const bulkCreateSubtypes = (payload, sourceSpec, keyPrefix) => {
    if (!selected || !sourceSpec) return;
    const { createNames = [], updateSubIds = [], skippedNames = [] } = payload || {};
    if (!createNames.length && !updateSubIds.length && !skippedNames.length) return;

    let templateData = null;
    let sourceLabel = "";

    if (sourceSpec.kind === "template") {
      const template = templates.find((t) => t.id === sourceSpec.templateId);
      if (!template) return;
      templateData = template.data;
      sourceLabel = template.name;
    } else if (sourceSpec.kind === "subtype") {
      const srcDoc = library.find((d) => d._uid === sourceSpec.docUid);
      const srcSub = srcDoc && srcDoc.ticketSubType.find((s) => s.id === sourceSpec.subId);
      if (!srcSub) return;
      templateData = templateDataFromSubType(srcSub);
      sourceLabel = `${srcSub.ticketType || srcSub.ticketKey}${srcDoc.ticketType || srcDoc.ticketKey ? ` (${srcDoc.ticketType || srcDoc.ticketKey})` : ""}`;
    } else {
      return;
    }

    // New sub-types: ticketKey matches the typed name exactly (plus the
    // optional prefix) — same convention ticketType uses elsewhere in the
    // app, so the two never drift apart the way they could before.
    const existingKeys = new Set(selected.ticketSubType.map((s) => xnorm(s.ticketKey).toLowerCase()));
    const created = [];
    createNames.forEach((name) => {
      const base = `${xnorm(keyPrefix)}${name}`;
      let key = base;
      let n = 2;
      while (existingKeys.has(key.toLowerCase())) { key = `${base} (${n})`; n += 1; }
      existingKeys.add(key.toLowerCase());
      created.push(subTypeFromTemplate({ data: templateData }, name, key));
    });

    // Existing sub-types the user opted to update: keep id/ticketKey/
    // ticketType/_passthrough, replace only the template-derived shape.
    const updateSubIdSet = new Set(updateSubIds);
    const updatedLabels = [];
    const nextExisting = selected.ticketSubType.map((s) => {
      if (!updateSubIdSet.has(s.id)) return s;
      const cloned = regenerateIds(templateData || {});
      updatedLabels.push(s.ticketType || s.ticketKey);
      return {
        ...s,
        prefix: cloned.prefix || s.prefix,
        isAutoEscalation: !!cloned.isAutoEscalation,
        customFieldsMetaData: cloned.customFieldsMetaData || [],
        statusWorkFlow: cloned.statusWorkFlow || [],
      };
    });

    updateSelected({ ...selected, ticketSubType: [...nextExisting, ...created] });

    const warnings = [`Stamped out from "${sourceLabel}".`];
    if (created.length) warnings.push(`Created ${created.length} new sub-type${created.length === 1 ? "" : "s"}.`);
    if (updatedLabels.length) warnings.push(`Updated ${updatedLabels.length} existing sub-type${updatedLabels.length === 1 ? "" : "s"} from the template: ${updatedLabels.join(", ")}.`);
    if (skippedNames.length) warnings.push(`Already existed, left unchanged: ${skippedNames.join(", ")}.`);

    setImportNotice({
      title: "Bulk create finished",
      warnings,
      missingSheets: [],
    });
    setBulkOpen(false);
  };

  const nav = [
    { id: "library", label: `Ticket Types (${library.length})`, icon: <ListTree size={16} /> },
    { id: "templates", label: `Templates (${templates.length})`, icon: <LayoutTemplate size={16} /> },
    { id: "smartupdate", label: "Smart Update", icon: <Wand2 size={16} /> },
    { id: "details", label: "Details", icon: <Ticket size={16} /> },
    { id: "roles", label: "Roles", icon: <Users size={16} /> },
    { id: "subtypes", label: `Sub-Types (${selected ? selected.ticketSubType.length : 0})`, icon: <ListTree size={16} /> },
    { id: "json", label: "JSON", icon: <FileJson size={16} /> },
  ];

  const loadSample = () => {
    const created = { ...SAMPLE(), _uid: uid() };
    setLibrary([...library, created]);
    setSelectedUid(created._uid);
    setTab("details");
  };

  return (
    <div className="min-h-screen bg-[#070A13] flex text-slate-200" style={{ fontFamily: "Inter, ui-sans-serif, system-ui" }}>
      {/* sidebar */}
      <aside className="w-60 shrink-0 bg-[#050710] text-slate-300 flex flex-col border-r border-white/5">
        <div className="px-5 py-5 border-b border-white/5">
          <div className="flex items-center gap-2 text-white font-semibold">
            <Workflow size={18} className="text-teal-400" />
            Ticket Type Builder
          </div>
          <div className="text-[11px] text-slate-500 mt-0.5">Ticket type dataset</div>
        </div>
        <nav className="flex-1 px-2 py-4 space-y-1">
          {nav.map((n) => {
            const disabled = n.id !== "library" && n.id !== "templates" && n.id !== "smartupdate" && !selected;
            return (
              <button
                key={n.id}
                onClick={() => !disabled && setTab(n.id)}
                disabled={disabled}
                className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm transition ${tab === n.id ? "bg-teal-500/10 text-teal-300" : disabled ? "text-slate-700 cursor-not-allowed" : "hover:bg-white/5 text-slate-400"}`}
              >
                {n.icon}
                {n.label}
              </button>
            );
          })}
        </nav>
        <div className="px-3 py-4 border-t border-white/5 space-y-1.5">
          <button onClick={loadSample} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-xs text-slate-400 hover:bg-white/5 hover:text-white transition">
            <RefreshCw size={13} /> Add sample ticket type
          </button>
        </div>
      </aside>

      {/* main */}
      <main className="flex-1 min-w-0">
        <header className="bg-[#0A0E1C] border-b border-white/10 px-8 py-4 flex items-center justify-between gap-4">
          <div className="min-w-0">
            <h1 className="text-base font-semibold text-slate-100 truncate">
              {tab === "library" ? "All ticket types" : tab === "templates" ? "Sub-type templates" : tab === "smartupdate" ? "Smart update" : (selected ? (selected.ticketType || "Untitled ticket type") : "No ticket type selected")}
            </h1>
            <p className="text-xs font-mono text-slate-500">
              {tab === "library" ? `${library.length} loaded` : tab === "templates" ? `${templates.length} template${templates.length === 1 ? "" : "s"} available for bulk creation` : tab === "smartupdate" ? "describe changes, review the diff, then apply" : (selected ? (selected.ticketKey || "no key set yet") : "pick one from Ticket Types")}
            </p>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <Btn variant="ghost" onClick={exportSelectedToExcel}>
              <Download size={15} />Export to Excel
            </Btn>
            <Btn variant="ghost" onClick={exportAllToExcel}>
              <FileSpreadsheet size={15} />Export All to Excel ({library.length})
            </Btn>
            <input ref={excelInputRef} type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={handleImportExcel} className="hidden" />
            <Btn variant="ghost" onClick={() => excelInputRef.current && excelInputRef.current.click()}>
              <FileSpreadsheet size={15} />Import Excel
            </Btn>
            <input ref={fileInputRef} type="file" accept="application/json,.json" onChange={handleImportFile} className="hidden" />
            <Btn variant="ghost" onClick={() => fileInputRef.current && fileInputRef.current.click()}>
              <FileJson size={15} />Import JSON
            </Btn>
            <Btn variant="accent" onClick={() => selected && setTab("json")}><FileJson size={15} />Create JSON</Btn>
          </div>
        </header>

        {importNotice && (importNotice.warnings.length > 0 || importNotice.missingSheets.length > 0) && (
          <div className="mx-8 mt-5 rounded-xl border border-amber-500/25 bg-amber-500/10 px-4 py-3">
            <div className="flex items-start justify-between gap-3">
              <div className="flex items-start gap-2">
                <AlertTriangle size={16} className="text-amber-500 mt-0.5 shrink-0" />
                <div className="text-sm text-amber-200/90">
                  <p className="font-medium mb-1">{importNotice.title || "Imported"} — a few things to note:</p>
                  <ul className="space-y-0.5 text-xs list-disc pl-4">
                    {importNotice.missingSheets.map((s, i) => (
                      <li key={`m${i}`}>Sheet "{s}" wasn't found — that section is empty.</li>
                    ))}
                    {importNotice.warnings.map((w, i) => (
                      <li key={`w${i}`}>{w}</li>
                    ))}
                  </ul>
                </div>
              </div>
              <button onClick={() => setImportNotice(null)} className="text-amber-400 hover:text-amber-200 shrink-0"><X size={16} /></button>
            </div>
          </div>
        )}

        <div className={(tab === "library" || tab === "templates" || tab === "smartupdate") ? "px-8 py-6" : "px-8 py-6 max-w-4xl"}>
          {tab === "library" && (
            <LibraryTab
              library={library}
              selectedUid={selectedUid}
              onSelectType={selectType}
              onSelectSubtype={selectSubtype}
              onNew={newTicketType}
              onDuplicate={duplicateType}
              onDelete={deleteType}
              onExportAll={exportAll}
            />
          )}
          {tab === "templates" && (
            <TemplatesTab
              templates={templates}
              library={library}
              onImport={importTemplate}
              onImportFromSubType={importTemplateFromSubType}
              onCreateBlank={createBlankTemplate}
              onChange={updateTemplate}
              onDelete={deleteTemplate}
              onDuplicate={duplicateTemplate}
              onSetDefault={setDefaultTemplate}
            />
          )}
          {tab === "smartupdate" && (
            <SmartUpdateTab library={library} onApply={applySmartUpdate} onExportExcel={exportAllToExcel} />
          )}
          {tab === "details" && selected && <DetailsTab doc={selected} onChange={updateSelected} />}
          {tab === "roles" && selected && <RolesTab doc={selected} onChange={updateSelected} />}
          {tab === "subtypes" && selected && (
            <div>
              <div className="flex items-center gap-2 mb-4">
                <Btn variant="dashed" onClick={() => updateSelected({ ...selected, ticketSubType: [...selected.ticketSubType, newSubType()] })}>
                  <Plus size={15} />Add sub-ticket type
                </Btn>
                <Btn variant="ghost" onClick={() => setBulkOpen((v) => !v)}>
                  <Sparkles size={15} />Bulk create from template
                </Btn>
              </div>

              {bulkOpen && (
                <BulkCreatePanel
                  templates={templates}
                  library={library}
                  currentUid={selected._uid}
                  onGenerate={bulkCreateSubtypes}
                  onClose={() => setBulkOpen(false)}
                />
              )}

              {selected.ticketSubType.map((s) => (
                <SubTypeEditor
                  key={s.id}
                  sub={s}
                  forceOpen={s.id === focusSubtypeId}
                  onChange={(patch) => updateSelected({ ...selected, ticketSubType: selected.ticketSubType.map((x) => (x.id === s.id ? patch : x)) })}
                  onDelete={() => {
                    if (window.confirm(`Remove sub-type "${s.ticketType || s.ticketKey || "(unnamed)"}"? This can't be undone.`)) {
                      updateSelected({ ...selected, ticketSubType: selected.ticketSubType.filter((x) => x.id !== s.id) });
                    }
                  }}
                />
              ))}
            </div>
          )}
          {tab === "json" && selected && <JsonTab doc={selected} />}
        </div>
      </main>
    </div>
  );
}