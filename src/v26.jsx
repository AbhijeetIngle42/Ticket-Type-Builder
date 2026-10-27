import React, { useState, useMemo, useCallback } from "react";
import {
  Plus, Trash2, ChevronDown, ChevronRight, Copy, Download,
  FileJson, Users, Ticket, ListTree, Workflow, Check, X,
  GripVertical, AlertTriangle, RefreshCw, FileSpreadsheet,
  Sparkles, LayoutTemplate, Star, Wand2, Diff, ClipboardCopy, ClipboardPaste, Eraser, ChevronUp, ShieldCheck
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

/* RFC 4122 v4 UUID. Every auto-escalation level carries one as its `_id`. */
function uuidv4() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return crypto.randomUUID();
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
  });
}

/* MongoDB ObjectId (24 hex chars) in Extended JSON form: { "$oid": "..." }.
   Top-level / sub-type `_id` must be exported as this object, not a plain
   string. (Auto-escalation level `_id`s are a separate thing: still uuids.) */
const OBJECT_ID_RE = /^[0-9a-fA-F]{24}$/;
function newObjectId() {
  const hex = (n) => Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16)).join("");
  // 4-byte timestamp + 5-byte random + 3-byte counter, like a real ObjectId
  const ts = Math.floor(Date.now() / 1000).toString(16).padStart(8, "0");
  return ts + hex(10) + hex(6);
}
const newOid = () => ({ $oid: newObjectId() });
/* "abc…" | {$oid:"abc…"} | undefined  ->  {$oid:"abc…"} (or left as-is if it
   isn't a valid ObjectId string, so we never corrupt unknown id formats). */
function toOid(v) {
  if (v && typeof v === "object" && typeof v.$oid === "string") return v;
  if (typeof v === "string" && OBJECT_ID_RE.test(v.trim())) return { $oid: v.trim() };
  return v;
}

/* Status codes are always lowercase with no spaces/symbols
   ("Resolved and Closed" -> "resolvedandclosed"). */
const toStatusCode = (v) => String(v == null ? "" : v).toLowerCase().replace(/[^a-z0-9]/g, "");
/* A status-level fileUpload field always uses this fixed key. */
const STATUS_FILE_UPLOAD_KEY = "uploadImage";

/* Applies the status rules: clean status/override/dependent codes, fixed
   key for fileUpload fields, and mandatoryCustomFields kept in sync with
   any key that had to change. */
function normalizeStatus(s) {
  const keyMap = {};
  const fields = (s.customFieldsMetaData || []).map((f) => {
    if (f && f.type === "fileUpload" && f.key !== STATUS_FILE_UPLOAD_KEY) {
      if (f.key) keyMap[f.key] = STATUS_FILE_UPLOAD_KEY;
      return { ...f, key: STATUS_FILE_UPLOAD_KEY };
    }
    return f;
  });
  return {
    ...s,
    status: toStatusCode(s.status),
    overrideStatus: toStatusCode(s.overrideStatus),
    dependentStatus: (s.dependentStatus || []).map(toStatusCode).filter(Boolean),
    mandatoryCustomFields: [...new Set((s.mandatoryCustomFields || []).map((k) => keyMap[k] || k))],
    customFieldsMetaData: fields,
    autoEscalationConfig: (s.autoEscalationConfig || []).map((e) => ({ ...e, overrideStatus: toStatusCode(e.overrideStatus) })),
  };
}

const newAutoEscalation = () => ({
  id: uid(), _id: uuidv4(), assignee: "", days: "", overrideStatus: "",
  notification: true, notificationType: ["inApp", "email"], ccAddress: [],
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
  ticketSubType: [], _passthrough: { _id: newOid() },
});

/* Keys the editor models explicitly at the doc / sub-type level. Anything
   else found on import (_id, __v, createdAt, updatedAt, backend-specific
   metadata, etc.) isn't understood by this UI, but it must not be thrown
   away — it's captured into `_passthrough` on import and merged back into
   the exported JSON untouched, so round-tripping through this app never
   silently deletes fields it doesn't have a form for. */
const DOC_KNOWN_KEYS = new Set(["ticketKey", "ticketType", "tenantId", "creators", "viewers", "assignee", "deleted", "ticketSubType"]);
// "roles" is listed on purpose even though the editor has no field for it:
// viewer/creator roles belong to the ticket type only, and a sub-type only
// carries `assignee`. Older live JSON has a stray sub-type-level `roles`
// array, so it is dropped on import instead of being passed through.
const SUBTYPE_KNOWN_KEYS = new Set(["ticketKey", "prefix", "ticketType", "customFieldsMetaData", "statusWorkFlow", "deleted", "isAutoEscalation", "assignee", "roles"]);

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
    // escalation levels carry a string uuid `_id`; a clone must get its own
    if (typeof out._id === "string") out._id = uuidv4();
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

// `explicitMandatory` is set for sub-type level custom fields, which always
// carry an explicit isMandatory (true or false) like the live schema does.
// Status-level fields (remarks etc.) keep the old "only when true" shape.
function advancedToJson(f, explicitMandatory) {
  const out = {};
  if (explicitMandatory) out.isMandatory = !!f.isMandatory;
  else if (f.isMandatory) out.isMandatory = true;
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

function fieldToJson(f, subTypeLevel = false) {
  const adv = advancedToJson(f, subTypeLevel);
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

function statusToJson(s0) {
  const s = normalizeStatus(s0);
  return {
    status: s.status, label: s.label, public: s.public,
    roles: s.roles, assignee: (s.assigneeRoles || []).map((r) => ({ roleId: r })),
    notification: s.notification, notificationType: s.notificationType,
    notificationRole: s.notificationRole,
    mandatoryCustomFields: s.mandatoryCustomFields,
    overrideStatus: s.overrideStatus,
    dependentStatus: s.dependentStatus,
    customFieldsMetaData: (s.customFieldsMetaData || []).map((f) => fieldToJson(f)),
    autoEscalationConfig: (s.autoEscalationConfig || []).map((e) => ({
      assignee: e.assignee, days: e.days === "" ? undefined : Number(e.days),
      notification: e.notification, overrideStatus: e.overrideStatus,
      notificationType: e.notificationType, ccAddress: e.ccAddress,
      _id: e._id || uuidv4(),
    })),
  };
}

/* A sub-type "has auto-escalation" as soon as any of its statuses has an
   escalation level; isAutoEscalation is then always true on export. */
const subTypeHasEscalation = (st) => (st.statusWorkFlow || []).some((w) => (w.autoEscalationConfig || []).length > 0);

function subTypeToJson(st) {
  const pt = { ...(st._passthrough || {}) };
  if ("_id" in pt) pt._id = toOid(pt._id);
  return {
    ...pt,
    ticketKey: st.ticketKey, prefix: st.prefix, ticketType: st.ticketType,
    assignee: (st.assigneeRoles || []).map((r) => ({ roleId: r })),
    customFieldsMetaData: (st.customFieldsMetaData || []).map((f) => fieldToJson(f, true)),
    statusWorkFlow: (st.statusWorkFlow || []).map(statusToJson),
    deleted: st.deleted, isAutoEscalation: !!st.isAutoEscalation || subTypeHasEscalation(st),
  };
}

function docToJson(doc) {
  const pt = { ...(doc._passthrough || {}) };
  if ("_id" in pt) pt._id = toOid(pt._id); // string id -> { "$oid": "..." }
  const raw = {
    ...pt,
    ticketKey: doc.ticketKey, ticketType: doc.ticketType, tenantId: doc.tenantId,
    creators: doc.creators.map((r) => ({ roleId: r })),
    viewers: doc.viewers.map((r) => ({ roleId: r })),
    assignee: doc.assignee.map((r) => ({ roleId: r })),
    deleted: doc.deleted,
    ticketSubType: doc.ticketSubType.map(subTypeToJson),
  };
  const cleaned = clean(raw) || {};
  // tenantId is a required field: clean() drops empty strings, so put it
  // back (in its usual position) as "" instead of omitting the key.
  const out = {};
  Object.keys(raw).forEach((k) => {
    if (k === "tenantId") out.tenantId = doc.tenantId || "";
    else if (k in cleaned) out[k] = cleaned[k];
  });
  return out;
}

/* Creators and viewers are required on every ticket type: returns the names
   of the required role lists that are empty (used to block exports). */
function missingRequiredRoles(doc) {
  const missing = [];
  if (!doc.creators || !doc.creators.length) missing.push("Creators");
  if (!doc.viewers || !doc.viewers.length) missing.push("Viewers");
  return missing;
}

/* The location placeholder in a sub-type prefix must be written exactly
   "${locationId}". Any other spelling of it (${LocationId}, ${locationid},
   ${location_id}, {locationId}, ${ locationId }) is rewritten to the exact form. */
const PREFIX_PLACEHOLDER = "${locationId}";
function normalizePrefix(prefix) {
  const from = String(prefix == null ? "" : prefix);
  const value = from.replace(/\$?\{\s*location[\s_-]?id\s*\}/gi, () => PREFIX_PLACEHOLDER);
  return { value, changed: value !== from };
}

/* Fix every sub-type prefix of a doc. Returns the same doc object when nothing
   needed fixing, plus a list of what changed (for the review toast). */
function normalizeDocPrefixes(doc) {
  const fixes = [];
  const subs = (doc.ticketSubType || []).map((s) => {
    const { value, changed } = normalizePrefix(s.prefix);
    if (!changed) return s;
    fixes.push({ ticket: doc.ticketType || doc.ticketKey || "(untitled ticket type)", sub: s.ticketType || s.ticketKey || "(untitled sub-type)", from: s.prefix, to: value });
    return { ...s, prefix: value };
  });
  return fixes.length ? { doc: { ...doc, ticketSubType: subs }, fixes } : { doc, fixes };
}

/* Assignee roles that auto-escalation relies on: the assignee of each escalation
   level. A status's own "new assignee" role is NOT included — it needs no
   matching entry in the sub-type / ticket-type lists. Escalation levels with
   no assignee are returned separately (a level must always have one). */
function usedAssigneeRoles(sub) {
  const roles = new Set();
  const blankLevels = [];
  (sub.statusWorkFlow || []).forEach((w) => {
    (w.autoEscalationConfig || []).forEach((e, i) => {
      const a = xnorm(e.assignee);
      if (a) roles.add(a); else blankLevels.push(`status "${w.label || w.status || "(untitled)"}" escalation level ${i + 1}`);
    });
  });
  return { roles: [...roles], blankLevels };
}

/* Rules (checked before every export):
   - every sub-type needs assignee roles, on the sub-type or the ticket type;
   - every escalation level needs an assignee role;
   - every escalation-level assignee role (only where auto-escalation exists) must
     be listed in BOTH the sub-type's and the ticket type's assignee roles.
   Returns [{ text, fixable }]; `fixable` ones are solved by syncAssigneeRoles(). */
function assigneeProblems(doc) {
  const out = [];
  const root = doc.assignee || [];
  (doc.ticketSubType || []).forEach((s) => {
    const name = s.ticketType || s.ticketKey || "(untitled sub-type)";
    const own = s.assigneeRoles || [];
    if (!own.length && !root.length) out.push({ text: `${name}: no assignee roles — set them on the sub-type or the ticket type` });
    const { roles, blankLevels } = usedAssigneeRoles(s);
    blankLevels.forEach((b) => out.push({ text: `${name}: ${b} has no assignee role` }));
    const missSub = roles.filter((r) => !own.includes(r));
    const missRoot = roles.filter((r) => !root.includes(r));
    if (missSub.length || missRoot.length) {
      const where = [missSub.length ? `sub-type (${missSub.join(", ")})` : "", missRoot.length ? `ticket type (${missRoot.join(", ")})` : ""].filter(Boolean).join(" and ");
      out.push({ fixable: true, text: `${name}: escalation assignee roles are missing from the ${where}` });
    }
  });
  return out;
}

/* Adds every escalation assignee role to both the sub-type's and the ticket type's assignee roles. */
function syncAssigneeRoles(doc) {
  const rootAdd = [];
  const subs = (doc.ticketSubType || []).map((s) => {
    const { roles } = usedAssigneeRoles(s);
    roles.forEach((r) => { if (!rootAdd.includes(r)) rootAdd.push(r); });
    const add = roles.filter((r) => !(s.assigneeRoles || []).includes(r));
    return add.length ? { ...s, assigneeRoles: [...(s.assigneeRoles || []), ...add] } : s;
  });
  return { ...doc, assignee: [...(doc.assignee || []), ...rootAdd.filter((r) => !(doc.assignee || []).includes(r))], ticketSubType: subs };
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
  return normalizeStatus({
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
      id: uid(), _id: typeof e._id === "string" && e._id ? e._id : uuidv4(),
      assignee: e.assignee || "", days: e.days != null ? String(e.days) : "",
      overrideStatus: e.overrideStatus || "", notification: !!e.notification,
      notificationType: Array.isArray(e.notificationType) ? e.notificationType : [],
      ccAddress: Array.isArray(e.ccAddress) ? e.ccAddress : [],
    })),
  });
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

/* ---- Less typing: blank cells are filled in from what is already known ----
   Prefix        <- Ticket Details "Prefix", else ${locationId}-<TICKET KEY>
   Subtype key   <- the sub-type's name
   Assignee      <- the ticket type's Assignee Roles
   Sub-types     <- the "Sub Types" list in Ticket Details (one comma separated cell)
   Status flow   <- one block with a blank Ticket Key / Ticket Subtype serves every sub-type
   Role cells accept <assignee> <creators> <viewers> <ticket assignee> instead of typing role ids. */
const xDerivedPrefix = (ticketKey) => "${locationId}-" + xnorm(ticketKey).toUpperCase().replace(/[^A-Z0-9&]+/g, "-").replace(/^-+|-+$/g, "");
const xSplitList = (v) => xnorm(v).split(/[,;\n]+/).map((x) => x.trim()).filter(Boolean);
const X_ALL = "*";
const xIsAllLabel = (v) => /^\(?\s*all(\s+sub-?\s?types?)?\s*\)?$/i.test(xnorm(v));
/* "<assignee>" -> that sub-type's assignee roles, "<creators>", "<viewers>", "<ticket assignee>" -> the ticket type's lists */
function xResolveRoleTokens(roles, ctx) {
  const out = [];
  (roles || []).forEach((r) => {
    const raw = xnorm(r);
    if (!/^[<{[]/.test(raw)) { out.push(raw); return; }
    const k = raw.replace(/[<>{}[\]]/g, "").replace(/\s+/g, " ").trim().toLowerCase();
    if (/^(ticket|root) (type )?assignee/.test(k)) out.push(...ctx.ticketAssignee);
    else if (/^(sub ?type )?assignee/.test(k)) out.push(...ctx.assignee);
    else if (/^creator/.test(k)) out.push(...ctx.creators);
    else if (/^viewer/.test(k)) out.push(...ctx.viewers);
    else out.push(raw);
  });
  return [...new Set(out.filter(Boolean))];
}

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
    prefix: xnorm(xCell(r, col("prefix"))),
    subTypeNames: xSplitList(xCell(r, col("sub type"))),
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
  let lastKey = "__default__"; // Ticket Key sticks down: only the first row of a ticket type needs it
  for (const r of rows.slice(1)) {
    const name = xnorm(xCell(r, col("ticket subtype")));
    const keyCell = hasTicketCol ? xnorm(xCell(r, col("ticket key"))) : "";
    if (keyCell) lastKey = keyCell;
    if (!name) continue;
    const ticketKey = hasTicketCol ? lastKey : "__default__";
    const key = xnorm(xCell(r, col("subtype key"))) || name;
    byTicket[ticketKey] = byTicket[ticketKey] || { subtypes: {}, order: [] };
    byTicket[ticketKey].subtypes[name] = {
      ticketKey: key,
      prefix: xnorm(xCell(r, col("prefix"))),
      assigneeRoles: xSplitRoles(xCell(r, col("assignee"))),
      prefixGiven: !!xnorm(xCell(r, col("prefix"))),
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
  const commonFields = {}; // ticketKey -> fields written with no sub-type: added to every sub-type
  for (const r of rows.slice(1)) {
    const ticketKeyCell = offset ? xnorm(xCell(r, 0)) : "";
    const subtypeCell = xnorm(xCell(r, offset + 0));
    const labelCell = xnorm(xCell(r, offset + 1));
    const typeCell = xnorm(xCell(r, offset + 2));
    const mandatoryCell = xnorm(xCell(r, offset + 3));
    const rolesCell = xnorm(xCell(r, offset + 4));

    if (ticketKeyCell) { if (ticketKeyCell !== currentTicketKey) currentSubtype = null; currentTicketKey = ticketKeyCell; }
    if (subtypeCell) currentSubtype = xIsAllLabel(subtypeCell) ? X_ALL : subtypeCell;
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
      if (!currentSubtype || currentSubtype === X_ALL) (commonFields[currentTicketKey] = commonFields[currentTicketKey] || []).push(currentField);
      else if (bucket && bucket.subtypes[currentSubtype]) bucket.subtypes[currentSubtype].customFieldsMetaData.push(currentField);
      continue;
    }
    if (currentField && typeCell) currentField.options.push({ value: typeCell, key: "" });
  }
  touched.forEach((tk) => {
    const bucket = subtypesByTicket[tk];
    if (!bucket) return;
    Object.values(bucket.subtypes).forEach((st) => {
      if ((commonFields[tk] || []).length) st.customFieldsMetaData = [...commonFields[tk].map((f) => JSON.parse(JSON.stringify(f))), ...st.customFieldsMetaData];
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
const X_ALL_SUBTYPES = "(all sub-types)";
/* Bucket key for a whole ticket type, or for one sub-type of it. */
const xScopeKey = (ticketKey, subtype) => (subtype && subtype !== X_ALL_SUBTYPES ? `${ticketKey}||${subtype}` : ticketKey);

/* Shared reader for the two status sheets (Status Access Roles / Mandate
   rules). Optional leading columns, detected from the header row:
     [Ticket Key] [Ticket Subtype] Status ...
   Ticket Key and Ticket Subtype "stick down" the sheet until a new value is
   given, so a block only needs them on its first row and blocks can be
   separated by blank rows. A blank or "(all sub-types)" sub-type means the
   rows apply to every sub-type of that ticket type; a named sub-type gets
   its own block, which replaces the shared one for that sub-type. */
function xStatusSheetRows(sheet) {
  const rows = xRowsOf(sheet);
  const headerIdx = rows.findIndex((r) => r.slice(0, 4).some((c) => xnorm(c).toLowerCase().includes("status")));
  if (headerIdx === -1) return { headerRow: [], offset: 0, items: [] };
  const headerRow = rows[headerIdx].map((h) => xnorm(h).toLowerCase());
  const hasKey = !!headerRow[0] && headerRow[0].includes("ticket key");
  const hasSub = hasKey && !!headerRow[1] && headerRow[1].includes("subtype");
  const offset = (hasKey ? 1 : 0) + (hasSub ? 1 : 0);
  const items = [];
  let curKey = "__default__";
  let curSub = "";
  for (const r of rows.slice(headerIdx + 1)) {
    if (hasKey) {
      const k = xnorm(xCell(r, 0));
      if (k) { curKey = k; curSub = hasSub ? xnorm(xCell(r, 1)) : ""; }
      else if (hasSub && xnorm(xCell(r, 1))) curSub = xnorm(xCell(r, 1));
    }
    const label = xnorm(xCell(r, offset));
    if (!label) continue;
    // optional "Status Code" column (written by the exporter): the real status code, so two
    // statuses that share a label ("Approve", "Reject" ...) stay two different statuses
    const codeCol = headerRow.findIndex((h) => h.includes("status code"));
    const code = codeCol >= 0 ? toStatusCode(xCell(r, codeCol)) : "";
    items.push({ scope: xScopeKey(curKey, curSub), r, label, code: code || xToStatusCode(label) });
  }
  return { headerRow, offset, items };
}

/* Returns scopeKey (ticketKey or ticketKey||subtype) -> { statusCode -> { label, roles, public } }.
   Status codes keep the order the rows appear in the sheet. */
function xParseStatusAccessRoles(sheet) {
  const { headerRow, offset, items } = xStatusSheetRows(sheet);
  const publicCol = headerRow.findIndex((h) => h.includes("public"));
  const byScope = {};
  for (const { scope, r, label, code } of items) {
    byScope[scope] = byScope[scope] || {};
    // optional "Public" column: Yes/No overrides the default (only "open" is public)
    const publicCell = publicCol >= 0 ? xnorm(xCell(r, publicCol)) : "";
    byScope[scope][code] = {
      label, roles: xSplitRoles(xCell(r, offset + 1)),
      public: publicCell ? xIsYes(publicCell) : undefined,
    };
  }
  return byScope;
}

/* Same optional leading "Ticket Key" column convention as the roles
   sheet. Returns ticketKey -> { statusCode -> mandate rule }. */
/* Mandate rules columns (read by header name; old files without the new
   columns fall back to the old positions):
     Mandatory Images | Mandatory Comments | Comment Field Label | Comment Field Type | Mandatory New Assignee | New Assignee Role
   Be forgiving about how people fill them in: "Yes (anything)" still counts
   as Yes, "No / N/A / -" in a role cell means no role, and the old style of
   putting the label itself in "Mandatory Comments" ("Rejection Remarks",
   "STO No. (Number)" or "Yes (Rejection Remarks)") keeps working. */
const xYesish = (v) => /^(yes|y|true|1)(\s|\(|$)/i.test(xnorm(v));
const xNoish = (v) => ["", "no", "n", "false", "0", "na", "n/a", "-", "none"].includes(xnorm(v).toLowerCase());

function xParseMandateRules(sheet) {
  const { headerRow, offset, items } = xStatusSheetRows(sheet);
  const find = (test, fallback) => { const i = headerRow.findIndex(test); return i >= 0 ? i : fallback; };
  const imagesCol = find((h) => h.includes("image"), offset + 1);
  const commentsCol = find((h) => h.includes("comment") && !h.includes("label") && !h.includes("type"), offset + 2);
  const labelCol = find((h) => h.includes("comment") && h.includes("label"), -1);
  const typeCol = find((h) => h.includes("comment") && h.includes("type"), -1);
  const assigneeYnCol = find((h) => h.includes("assignee") && !h.includes("role"), offset + 3);
  const roleCol = find((h) => h.includes("assignee") && h.includes("role"), offset + 4);
  const byTicket = {};
  for (const { scope: ticketKey, r, label, code } of items) {
    const imagesCell = xnorm(xCell(r, imagesCol));
    const commentsCell = xnorm(xCell(r, commentsCol));
    const labelCell = labelCol >= 0 ? xnorm(xCell(r, labelCol)) : "";
    const typeCell = typeCol >= 0 ? xnorm(xCell(r, typeCol)).toLowerCase() : "";
    const assigneeYnCell = xnorm(xCell(r, assigneeYnCol));
    const roleCell = xnorm(xCell(r, roleCol));

    // "No" (written by the exporter) must read back as "not mandatory", never as a field labelled "No"
    const commentsOff = xNoish(commentsCell) && !labelCell;
    let rawLabel = labelCell;
    if (!rawLabel && !commentsOff) {
      const bracket = commentsCell.match(/^(?:yes|y|true)\s*\((.+)\)\s*$/i); // "Yes (Rejection Remarks)"
      if (bracket) rawLabel = bracket[1];
      else if (!xYesish(commentsCell)) rawLabel = commentsCell;               // "Rejection Remarks"
    }
    rawLabel = rawLabel.replace(/\(mandatory\)/i, "").trim();
    // a "(Number)" / "(Date)" / "(Text)" suffix on the label sets the remark field's type
    const typeMatch = rawLabel.match(/\((text|number|date)\)\s*$/i);
    if (typeMatch) rawLabel = rawLabel.replace(/\((text|number|date)\)\s*$/i, "").trim();
    const commentsType = typeCell === "fileupload" ? "fileUpload"
      : ["text", "number", "date"].includes(typeCell) ? typeCell : (typeMatch ? typeMatch[1].toLowerCase() : "text");

    byTicket[ticketKey] = byTicket[ticketKey] || {};
    byTicket[ticketKey][code] = {
      mandatoryImages: xYesish(imagesCell),
      mandatoryComments: !commentsOff,
      commentsLabel: rawLabel || "Comments",
      commentsType,
      mandatoryNewAssignee: xYesish(assigneeYnCell),
      newAssigneeRole: xNoish(roleCell) ? "" : roleCell,
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
function xParseEscalationDetailsLegacy(sheet) {
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

/* Escalation details, one row per escalation level (current layout):
     Ticket Key | Ticket Subtype | Status | Level | Days | Assignee Role |
     Notification (Yes/No) | Notification Type | CC Email Addresses | Override Status
   Ticket Key / Ticket Subtype / Status "stick down" until a new value is given,
   so they only need to appear on the first row of a group. Columns are found by
   header name. Level (1 = escalates first) sets the order of the levels, so to
   shuffle the assignees just change the Level numbers. Notification
   defaults to Yes, in-app + email when those cells are empty. */
function xNotifTypes(cell) {
  const wanted = xnorm(cell).split(/[,;/\n]+/).map((t) => t.toLowerCase().replace(/[^a-z]/g, "")).filter(Boolean);
  const out = [];
  wanted.forEach((w) => {
    const hit = NOTIF_TYPES.find((n) => n.toLowerCase() === w) || (w === "app" || w === "inapp" ? "inApp" : null);
    if (hit && !out.includes(hit)) out.push(hit);
  });
  return out;
}
function xParseEscalationRows(rows, headerIdx) {
  const header = rows[headerIdx].map((h) => xnorm(h).toLowerCase());
  const find = (test) => header.findIndex(test);
  const cSub = find((h) => h.includes("subtype"));
  const cStatus = find((h) => h === "status" || h.startsWith("status"));
  const cLevel = find((h) => h.startsWith("level"));
  const cDays = find((h) => h.includes("days"));
  const cRole = find((h) => h.includes("assignee") || h.includes("role"));
  const cNotif = find((h) => h.startsWith("notification") && !h.includes("type"));
  const cType = find((h) => h.includes("notification") && h.includes("type"));
  const cCc = find((h) => h.includes("cc"));
  const cOver = find((h) => h.includes("override"));
  const cCode = find((h) => h.includes("status code"));
  const escalationByTicket = {};
  const warnings = [];
  let key = "__default__";
  let sub = "";
  let status = "";
  for (const r of rows.slice(headerIdx + 1)) {
    const k = xnorm(xCell(r, 0));
    if (k) { key = k; sub = ""; status = ""; }
    const sv = cSub >= 0 ? xnorm(xCell(r, cSub)) : "";
    if (sv) { sub = xIsAllLabel(sv) ? X_ALL : sv; status = ""; }
    const stv = cStatus >= 0 ? xnorm(xCell(r, cStatus)) : "";
    const codeV = cCode >= 0 ? toStatusCode(xCell(r, cCode)) : "";
    if (stv || codeV) status = codeV || xToStatusCode(stv);
    const days = cDays >= 0 ? xnorm(xCell(r, cDays)) : "";
    const role = cRole >= 0 ? xnorm(xCell(r, cRole)) : "";
    if (!days && !role) continue;
    if (!status) { warnings.push(`Escalation row (${role || days}) has no Status above it — skipped.`); continue; }
    const subUse = sub || X_ALL; // a blank Ticket Subtype means every sub-type of that ticket type
    if (!role) { warnings.push(`Escalation for '${subUse === X_ALL ? "all sub-types" : subUse}' under ${status} has Days but no Assignee Role — skipped that level.`); continue; }
    const notifCell = cNotif >= 0 ? xnorm(xCell(r, cNotif)) : "";
    const types = cType >= 0 ? xNotifTypes(xCell(r, cType)) : [];
    const ccCell = cCc >= 0 ? xnorm(xCell(r, cCc)) : "";
    const overCell = cOver >= 0 ? xnorm(xCell(r, cOver)) : "";
    const lvl = cLevel >= 0 ? parseFloat(xnorm(xCell(r, cLevel))) : NaN;
    const entry = {
      _lv: Number.isFinite(lvl) ? lvl : Infinity, // blank / invalid Level -> after the numbered ones, in sheet order
      assignee: role,
      days,
      notification: notifCell ? !xNoish(notifCell) : true,
      notificationType: types.length ? types : (cType >= 0 && xnorm(xCell(r, cType)) ? [] : [...NOTIF_TYPES]),
      ccAddress: ccCell.split(/[,;\s]+/).map((x) => x.trim()).filter(Boolean),
      overrideStatus: overCell ? xToStatusCode(overCell) : "",
    };
    const bucket = (escalationByTicket[key] = escalationByTicket[key] || {});
    const bySub = (bucket[status] = bucket[status] || {});
    (bySub[subUse] = bySub[subUse] || []).push(entry);
  }
  // Level decides the order inside a (ticket, sub-type, status) group; rows may sit anywhere in the sheet.
  Object.values(escalationByTicket).forEach((byStatus) => Object.values(byStatus).forEach((bySub) => Object.keys(bySub).forEach((sub) => {
    bySub[sub] = bySub[sub].map((e, i) => ({ e, i })).sort((a, b) => (a.e._lv - b.e._lv) || (a.i - b.i)).map(({ e }) => { const { _lv, ...rest } = e; return rest; });
  })));
  return { escalationByTicket, warnings };
}

function xParseEscalationDetails(sheet) {
  const rows = xRowsOf(sheet);
  const headerIdx = rows.findIndex((r) => xnorm(xCell(r, 0)).toLowerCase() === "ticket key");
  // the older block layout (TICKET:/STATUS: markers, Days/Role pairs) is still read
  return headerIdx >= 0 ? xParseEscalationRows(rows, headerIdx) : xParseEscalationDetailsLegacy(sheet);
}

function xBuildStatusWorkflow(subtypeName, globalStatusRoles, mandateRules, escalationByStatus) {
  const preferredOrder = ["open", "inprogress", "resolved", "closed"];
  const allCodes = new Set([...Object.keys(globalStatusRoles), ...Object.keys(mandateRules), ...Object.keys(escalationByStatus)]);
  // Statuses keep the order of the Status Access Roles sheet (the workflow
  // order); any code that only appears in the mandate/escalation sheets is
  // appended afterwards, standard statuses first, the rest alphabetical.
  const orderedCodes = Object.keys(globalStatusRoles);
  const leftover = [...allCodes].filter((c) => !orderedCodes.includes(c));
  preferredOrder.filter((c) => leftover.includes(c)).forEach((c) => orderedCodes.push(c));
  leftover.filter((c) => !orderedCodes.includes(c)).sort().forEach((c) => orderedCodes.push(c));

  return orderedCodes.map((code) => {
    const roleInfo = globalStatusRoles[code] || {};
    const mandate = mandateRules[code] || {};
    const escRows = (escalationByStatus[code] || {})[subtypeName] || (escalationByStatus[code] || {})[X_ALL] || [];

    const mandatoryFields = [];
    const customFields = [];
    if (mandate.mandatoryImages) mandatoryFields.push("images");
    if (mandate.mandatoryComments) {
      const key = mandate.commentsType === "fileUpload" ? STATUS_FILE_UPLOAD_KEY : xToCamelKey(mandate.commentsLabel || "Comments");
      mandatoryFields.push(key);
      customFields.push({ label: mandate.commentsLabel || "Comments", key, type: mandate.commentsType || "text" });
    }
    if (mandate.mandatoryNewAssignee) mandatoryFields.push("newAssignee");

    const assigneeRoles = mandate.newAssigneeRole ? xSplitRoles(mandate.newAssigneeRole) : [];
    const autoEscalation = escRows.map((e) => ({
      assignee: e.assignee,
      days: /^\d+$/.test(xnorm(e.days)) ? Number(e.days) : e.days,
      notification: e.notification !== undefined ? e.notification : true,
      notificationType: e.notificationType || ["inApp", "email"],
      ...(e.ccAddress ? { ccAddress: e.ccAddress } : {}),
      ...(e.overrideStatus ? { overrideStatus: e.overrideStatus } : {}),
    }));

    return {
      status: code,
      label: roleInfo.label || code.charAt(0).toUpperCase() + code.slice(1),
      // Only "open" is public by default; every other status is private unless
      // the sheet's optional "Public" column says Yes.
      public: roleInfo.public !== undefined ? roleInfo.public : code === "open",
      roles: roleInfo.roles || [],
      assignee: assigneeRoles.map((r) => ({ roleId: r })),
      // Specs never say how to notify, so every status gets in-app + email by default.
      notification: true,
      notificationType: ["inApp", "email"],
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
  // Sub-types typed in Ticket Details ("Sub Types" cell) are created here, before the field sheet is read,
  // so common fields reach them too. Everything they don't say is filled in below.
  const autoNote = { created: 0, prefix: 0, assignee: 0 };
  detailsRows.forEach((d) => {
    const effKey = d.ticketKey || d.ticketType;
    if (!d.subTypeNames.length) return;
    const bk = !subtypesByTicket[effKey] && subtypesByTicket.__default__ && detailsRows.length === 1 ? "__default__" : effKey;
    const bucket = (subtypesByTicket[bk] = subtypesByTicket[bk] || { subtypes: {}, order: [] });
    d.subTypeNames.forEach((name) => {
      if (bucket.subtypes[name]) return;
      bucket.subtypes[name] = { ticketKey: name, prefix: "", prefixGiven: false, assigneeRoles: [], isAutoEscalation: false, deleted: false, ticketType: name, customFieldsMetaData: [], statusWorkFlow: [] };
      bucket.order.push(name);
      autoNote.created++;
    });
  });
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
    // a sub-type's own block (Ticket Subtype column) wins over the ticket type's shared block
    const ticketKeyEff = details.ticketKey || details.ticketType || (ticketKeyLookup !== "__default__" ? ticketKeyLookup : "");
    const ticketCtx = { creators: details.creators || [], viewers: details.viewers || [], ticketAssignee: details.assignee || [] };
    const effAssignee = (name) => { const st0 = subBucket.subtypes[name]; return st0 && st0.assigneeRoles && st0.assigneeRoles.length ? st0.assigneeRoles : (details.assignee || []); };
    // a sub-type's own block (Ticket Subtype column) wins over the ticket type's, which wins over the all-ticket-types block
    // scope rank of the block that applies: 3 = this sub-type's own, 2 = the ticket type's, 1 = the all-ticket-types block
    const pickScope = (byScope, name) => {
      const own = byScope[xScopeKey(ticketKeyLookup, name)];
      if (own && xScopeKey(ticketKeyLookup, name) !== ticketKeyLookup) return { map: own, rank: 3 };
      if (byScope[ticketKeyLookup]) return { map: byScope[ticketKeyLookup], rank: 2 };
      if (byScope.__default__) return { map: byScope.__default__, rank: 1 };
      return { map: {}, rank: 0 };
    };
    const rawRolesFor = (name) => pickScope(statusRolesByTicket, name).map;
    const rawMandateFor = (name) => {
      const roles = pickScope(statusRolesByTicket, name);
      const mand = pickScope(mandateByTicket, name);
      // a broader mandate block must not invent statuses the (narrower) status flow doesn't have
      if (mand.rank < roles.rank && Object.keys(roles.map).length) {
        const kept = {};
        Object.keys(mand.map).forEach((c) => { if (roles.map[c]) kept[c] = mand.map[c]; });
        return kept;
      }
      return mand.map;
    };
    const rolesFor = (name) => {
      const ctx = { ...ticketCtx, assignee: effAssignee(name) };
      const src = rawRolesFor(name);
      const out = {};
      Object.keys(src).forEach((c) => { out[c] = { ...src[c], roles: xResolveRoleTokens(src[c].roles, ctx) }; });
      return out;
    };
    const mandateFor = (name) => {
      const ctx = { ...ticketCtx, assignee: effAssignee(name) };
      const src = rawMandateFor(name);
      const out = {};
      Object.keys(src).forEach((c) => { out[c] = { ...src[c], newAssigneeRole: src[c].newAssigneeRole ? xResolveRoleTokens(xSplitRoles(src[c].newAssigneeRole), ctx).join(", ") : "" }; });
      return out;
    };
    const escalationByStatus = escalationByTicket[ticketKeyLookup] || escalationByTicket.__default__ || {};
    const matrixBucketKey = matrixFieldsByTicket[ticketKeyLookup] ? ticketKeyLookup : (matrixFieldsByTicket.__default__ ? "__default__" : ticketKeyLookup);
    const matrixBySubtype = matrixFieldsByTicket[matrixBucketKey] || {};
    const cascadeBucketKey = cascadeFieldsByTicket[ticketKeyLookup] ? ticketKeyLookup : (cascadeFieldsByTicket.__default__ ? "__default__" : ticketKeyLookup);
    const cascadeBySubtype = cascadeFieldsByTicket[cascadeBucketKey] || {};
    const apiBucketKey = apiFieldsByTicket[ticketKeyLookup] ? ticketKeyLookup : (apiFieldsByTicket.__default__ ? "__default__" : ticketKeyLookup);
    const apiBySubtype = apiFieldsByTicket[apiBucketKey] || {};

    // what the workbook actually said about each sub-type, so a re-upload only
    // touches the parts it covers (see mergeExcelDoc)
    const seenList = [];
    const escTicketHasRows = Object.keys(escalationByStatus).length > 0;
    const ticketSubType = subBucket.order.map((name) => {
      const st = subBucket.subtypes[name];
      seenList.push({
        roles: sheets.statusRoles ? new Set(Object.keys(rolesFor(name))) : null,
        mandate: sheets.mandate ? new Set(Object.keys(mandateFor(name))) : null,
        esc: sheets.escalation && escTicketHasRows ? new Set(Object.keys(escalationByStatus).filter((c) => escalationByStatus[c] && (escalationByStatus[c][name] || escalationByStatus[c][X_ALL]))) : null,
        // prefix / assignee the sheet left blank were filled in with defaults: a re-upload must not overwrite a real value with them
        prefixDefaulted: !st.prefixGiven,
        assigneeDefaulted: !(st.assigneeRoles && st.assigneeRoles.length),
      });
      const subPrefix = st.prefix || details.prefix || xDerivedPrefix(ticketKeyEff);
      const subAssignee = effAssignee(name);
      if (!st.prefix && !details.prefix) autoNote.prefix++;
      if (!(st.assigneeRoles && st.assigneeRoles.length) && subAssignee.length) autoNote.assignee++;
      const matrixFields = (matrixBySubtype[name] || []).map(xMatrixFieldToOutput);
      const cascadeFields = (cascadeBySubtype[name] || []).map(xCascadeFieldToOutput);
      const apiFields = (apiBySubtype[name] || []).map(xApiFieldToOutput);
      const statusWorkFlow = xBuildStatusWorkflow(name, rolesFor(name), mandateFor(name), escalationByStatus);
      return {
        ticketKey: st.ticketKey, prefix: subPrefix, ticketType: st.ticketType,
        assignee: subAssignee.map((r) => ({ roleId: r })),
        customFieldsMetaData: [...st.customFieldsMetaData.map(xFieldToOutput), ...matrixFields, ...cascadeFields, ...apiFields],
        statusWorkFlow,
        // escalation rules for this sub-type switch auto-escalation on
        deleted: st.deleted, isAutoEscalation: !!st.isAutoEscalation || statusWorkFlow.some((w) => (w.autoEscalationConfig || []).length > 0),
      };
    });

    const rawDoc = {
      ticketKey: details.ticketKey || details.ticketType || (ticketKeyLookup !== "__default__" ? ticketKeyLookup : ""),
      ticketType: details.ticketType || details.ticketKey || "",
      tenantId: details.tenantId || "",
      deleted: !!details.deleted,
      creators: (details.creators || []).map((r) => ({ roleId: r })),
      viewers: (details.viewers || []).map((r) => ({ roleId: r })),
      assignee: (details.assignee || []).map((r) => ({ roleId: r })),
      ticketSubType,
    };
    const stateDoc = jsonToDoc(rawDoc);
    stateDoc.ticketSubType.forEach((st, i) => { st._excelSeen = seenList[i]; });
    docs.push(stateDoc);
  });

  const filled = [
    autoNote.created ? `${autoNote.created} sub-type(s) created from the "Sub Types" list in Ticket Details` : "",
    autoNote.prefix ? `prefix filled in for ${autoNote.prefix} sub-type(s)` : "",
    autoNote.assignee ? `assignee roles filled in from the ticket type for ${autoNote.assignee} sub-type(s)` : "",
  ].filter(Boolean);
  if (filled.length) warnings.push(`Auto-filled: ${filled.join(", ")}.`);
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

/* The "starter" workbook: a short sheet with FORMULAS wired between its tabs, so only the real
   information is typed once and the repetitive rows fill themselves in Excel:
     Ticket Details  (type)  ticket key, roles, optional prefix
     Sub Types       (type)  ticket key (first row of a group) + sub-type name  ->  key, prefix, assignee, ticket key fill in
     Default Status Flow (type once)  the status flow every sub-type gets
     Status Access Roles / Mandate rules  = formulas: every sub-type x every status of the default flow
   Any generated cell can still be overtyped to make one sub-type different. The result is an ordinary
   workbook the importer reads (it reads the calculated values), and a minimal sheet also works without formulas. */
const STARTER_SUBS = 150;      // sub-type rows that are wired up (copy the last row down for more)
const STARTER_STATUSES = 8;    // statuses the default flow can hold
function excelStarterWorkbook() {
  const wb = XLSX.utils.book_new();
  const TD = "'Ticket Details'", ST = "'Sub Types'", DF = "'Default Status Flow'";
  const subEnd = STARTER_SUBS + 1, flowEnd = STARTER_STATUSES + 1;
  const AUTO_FILL = { fill: { fgColor: { rgb: "EAF4EA" } } };
  const AUTO_HEAD = { ...XLSX_HEADER_STYLE, fill: { fgColor: { rgb: "2E7D5B" } } };
  const HELPER = { font: { color: { rgb: "9AA5B8" }, sz: 9 } };
  const F = (f) => ({ t: "s", f, v: "" });

  // ---------------- Start here ----------------
  const guide = [
    ["START HERE — type each thing once, the green columns fill themselves in"],
    [""],
    ["1. Ticket Details", "One row per ticket type: Ticket Key, roles (comma separated) and, if you want, a Prefix. Leave Prefix blank to get  ${locationId}-<TICKET KEY>."],
    [""],
    ["2. Sub Types", "Type the Ticket Key on the first row of each ticket type (it carries down) and one sub-type name per row. Subtype Key, Prefix, Assignee Roles and the carried-down Ticket Key are green: they come from Ticket Details."],
    ["", "To make one sub-type different, just type over its green cell (its own key, prefix or assignee)."],
    [""],
    ["3. Default Status Flow", "Type the status flow ONCE (up to " + STARTER_STATUSES + " statuses). In a role cell you can write <assignee> (the sub-type's assignee), <creators>, <viewers> or <ticket assignee>."],
    [""],
    ["4. Status Access Roles / Mandate rules", "Both sheets are formulas: every sub-type gets every status of the Default Status Flow. Nothing to type here. To change ONE sub-type, type over its cells."],
    [""],
    ["5. Ticket Structure / Escalation details", "Normal sheets. Leave Ticket Subtype blank (or write (all sub-types)) to apply a field or escalation levels to every sub-type of that ticket type."],
    [""],
    ["Then", "Save the file and use Import Excel in the app. Open any ticket type to review everything that was filled in."],
    ["Limits", "Wired for " + STARTER_SUBS + " sub-types and " + STARTER_STATUSES + " statuses. More sub-types: copy the last row of Sub Types down. Keep the Default Status Flow filled from the top with no blank rows between statuses."],
  ];
  const wsGuide = XLSX.utils.aoa_to_sheet(guide);
  wsGuide["!cols"] = [{ wch: 38 }, { wch: 170 }];
  wsGuide.A1.s = { font: { bold: true, sz: 14, color: { rgb: "1F3864" } } };
  guide.forEach((row, i) => { const ref = XLSX.utils.encode_cell({ r: i, c: 0 }); if (i > 0 && row[0] && wsGuide[ref]) wsGuide[ref].s = { font: { bold: true } }; });
  XLSX.utils.book_append_sheet(wb, wsGuide, "Start here");

  // builds a sheet cell by cell so formula cells can carry a style
  const sheet = (headers, widths, autoCols, rowCount, cellFor, helperCols = []) => {
    const ws = {};
    headers.forEach((h, c) => { ws[XLSX.utils.encode_cell({ r: 0, c })] = { t: "s", v: h }; });
    for (let r = 1; r <= rowCount; r++) {
      headers.forEach((_, c) => {
        const f = cellFor(r + 1, c); // 1-based Excel row number
        if (f == null) return;
        ws[XLSX.utils.encode_cell({ r, c })] = { ...F(f), s: helperCols.includes(c) ? HELPER : AUTO_FILL };
      });
    }
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rowCount, c: headers.length - 1 } });
    styleSheet(ws, headers.length);
    autoCols.forEach((c) => { ws[XLSX.utils.encode_cell({ r: 0, c })].s = AUTO_HEAD; });
    ws["!cols"] = widths.map((wch) => ({ wch }));
    return ws;
  };
  const plain = (headers, widths) => { const ws = XLSX.utils.aoa_to_sheet([headers]); styleSheet(ws, headers.length); ws["!cols"] = widths.map((wch) => ({ wch })); return ws; };

  // ---------------- Ticket Details (typed) ----------------
  XLSX.utils.book_append_sheet(wb, plain(["Ticket Key", "Ticket Type", "Tenant", "Creator Roles", "Viewer Roles", "Assignee Roles", "Prefix (optional)", "Delete"], [20, 22, 14, 26, 32, 26, 26, 10]), "Ticket Details");

  // ---------------- Sub Types: type A (first row of a group) + B, the rest is formulas ----------------
  const tdLookup = (col, key) => `INDEX(${TD}!$${col}:$${col},MATCH(${key},${TD}!$A:$A,0))`;
  const subFormula = (r, c) => {
    if (c === 0 || c === 1) return null;                                   // A, B typed
    if (c === 2) return `IF($B${r}="","",$B${r})`;                       // Subtype Key = name
    if (c === 3) return `IF($B${r}="","",IF(IFERROR(${tdLookup("G", `$H${r}`)}&"","")<>"",${tdLookup("G", `$H${r}`)},"${"$"}{locationId}-"&UPPER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($H${r}," ","-"),"(",""),")",""))))`; // Prefix
    if (c === 4) return `IF($B${r}="","",IFERROR(${tdLookup("F", `$H${r}`)}&"",""))`;   // Assignee roles from the ticket type
    if (c === 5 || c === 6) return `IF($B${r}="","","No")`;               // Auto Escalation / Delete
    if (c === 7) return `IF($B${r}="","",$I${r})`;                       // Ticket Key (auto)
    if (c === 8) return r === 2 ? `IF($A2<>"",$A2,"")` : `IF($A${r}<>"",$A${r},$I${r - 1})`; // carried-down key (helper)
    return null;
  };
  XLSX.utils.book_append_sheet(wb, sheet(["Ticket Key", "Ticket Subtype", "Subtype Key", "Prefix", "Assignee Roles", "Auto Escalation", "Delete", "Ticket Key (auto)", "(helper)"], [20, 32, 28, 28, 28, 16, 10, 20, 12], [2, 3, 4, 5, 6, 7], STARTER_SUBS, subFormula, [8]), "Sub Types");

  // ---------------- Default Status Flow (typed once) ----------------
  const flowHeaders = ["Status", "Roles Allowed", "Public", "Mandatory Images (Yes/No)", "Mandatory Comments (Yes/No)", "Comment Field Label", "Comment Field Type", "Mandatory New Assignee (Yes/No)", "New Assignee Role"];
  const flowRows = [
    ["Open", "", "Yes", "No", "No", "", "", "No", ""],
    ["In Progress", "<assignee>", "No", "No", "No", "", "", "No", ""],
    ["Closed", "<assignee>", "No", "No", "Yes", "Closure Remarks", "text", "No", ""],
    ["Reopen", "<creators>", "No", "No", "Yes", "Reopen Remarks", "text", "No", ""],
  ];
  const wsFlow = XLSX.utils.aoa_to_sheet([flowHeaders, ...flowRows]);
  styleSheet(wsFlow, flowHeaders.length);
  wsFlow["!cols"] = [26, 30, 10, 18, 18, 26, 16, 20, 26].map((wch) => ({ wch }));
  wsFlow["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: STARTER_STATUSES, c: flowHeaders.length - 1 } });
  XLSX.utils.book_append_sheet(wb, wsFlow, "Default Status Flow");

  // ---------------- generated: every sub-type x every status of the default flow ----------------
  const n = `COUNTA(${DF}!$A$2:$A$${flowEnd})`;
  const rowsTotal = STARTER_SUBS * STARTER_STATUSES;
  const flowAt = (col, r) => `INDEX(${DF}!$${col}$2:$${col}$${flowEnd},$H${r})&""`;
  const resolveTokens = (text, r) => {
    let t = text;
    t = `SUBSTITUTE(${t},"<ticket assignee>",IFERROR(${tdLookup("F", `$A${r}`)}&"",""))`;
    t = `SUBSTITUTE(${t},"<assignee>",INDEX(${ST}!$E$2:$E$${subEnd},$G${r})&"")`;
    t = `SUBSTITUTE(${t},"<creators>",IFERROR(${tdLookup("D", `$A${r}`)}&"",""))`;
    t = `SUBSTITUTE(${t},"<viewers>",IFERROR(${tdLookup("E", `$A${r}`)}&"",""))`;
    return t;
  };
  // Status Access Roles: A key | B sub-type | C status | D roles | E public | F (unused spacer) | G sub# | H status#
  const rolesFormula = (r, c) => {
    if (c === 6) return `IF(${n}=0,"",IF(INT((ROW()-2)/${n})+1>${STARTER_SUBS},"",IF(INDEX(${ST}!$B$2:$B$${subEnd},INT((ROW()-2)/${n})+1)="","",INT((ROW()-2)/${n})+1)))`;
    if (c === 7) return `IF($G${r}="","",MOD(ROW()-2,${n})+1)`;
    const live = `$G${r}<>""`;
    if (c === 0) return `IF(${live},IF(INDEX(${ST}!$B$2:$B$${subEnd},$G${r})="","",INDEX(${ST}!$H$2:$H$${subEnd},$G${r})),"")`;
    if (c === 1) return `IF($A${r}="","",INDEX(${ST}!$B$2:$B$${subEnd},$G${r}))`;
    if (c === 2) return `IF($A${r}="","",INDEX(${DF}!$A$2:$A$${flowEnd},$H${r}))`;
    if (c === 3) return `IF($A${r}="","",${resolveTokens(flowAt("B", r), r)})`;
    if (c === 4) return `IF($A${r}="","",${flowAt("C", r)})`;
    return null;
  };
  XLSX.utils.book_append_sheet(wb, sheet(["Ticket Key", "Ticket Subtype", "Status", "Roles Allowed", "Public", "", "Sub #", "Status #"], [20, 32, 22, 34, 10, 3, 8, 9], [0, 1, 2, 3, 4], rowsTotal, rolesFormula, [5, 6, 7]), "Status Access Roles");

  // Mandate rules: A key | B sub-type | C status | D images | E comments | F label | G type | H new assignee | I role | J (spacer) | K sub# | L status#
  const mandHelpSub = "K", mandHelpSt = "L";
  const mandFormula = (r, c) => {
    if (c === 10) return `IF(${n}=0,"",IF(INT((ROW()-2)/${n})+1>${STARTER_SUBS},"",IF(INDEX(${ST}!$B$2:$B$${subEnd},INT((ROW()-2)/${n})+1)="","",INT((ROW()-2)/${n})+1)))`;
    if (c === 11) return `IF($K${r}="","",MOD(ROW()-2,${n})+1)`;
    const sub = `$${mandHelpSub}${r}`, st = `$${mandHelpSt}${r}`;
    const at = (col) => `INDEX(${DF}!$${col}$2:$${col}$${flowEnd},${st})&""`;
    if (c === 0) return `IF(${sub}<>"",IF(INDEX(${ST}!$B$2:$B$${subEnd},${sub})="","",INDEX(${ST}!$H$2:$H$${subEnd},${sub})),"")`;
    if (c === 1) return `IF($A${r}="","",INDEX(${ST}!$B$2:$B$${subEnd},${sub}))`;
    if (c === 2) return `IF($A${r}="","",INDEX(${DF}!$A$2:$A$${flowEnd},${st}))`;
    if (c >= 3 && c <= 7) return `IF($A${r}="","",${at(String.fromCharCode(65 + c))})`;      // D..H straight from the default flow
    if (c === 8) {
      let t = at("I");
      t = `SUBSTITUTE(${t},"<ticket assignee>",IFERROR(${tdLookup("F", `$A${r}`)}&"",""))`;
      t = `SUBSTITUTE(${t},"<assignee>",INDEX(${ST}!$E$2:$E$${subEnd},${sub})&"")`;
      t = `SUBSTITUTE(${t},"<creators>",IFERROR(${tdLookup("D", `$A${r}`)}&"",""))`;
      t = `SUBSTITUTE(${t},"<viewers>",IFERROR(${tdLookup("E", `$A${r}`)}&"",""))`;
      return `IF($A${r}="","",${t})`;
    }
    return null;
  };
  XLSX.utils.book_append_sheet(wb, sheet(["Ticket Key", "Ticket Subtype", "Status", "Mandatory Images (Yes/No)", "Mandatory Comments (Yes/No)", "Comment Field Label", "Comment Field Type", "Mandatory New Assignee (Yes/No)", "New Assignee Role", "", "Sub #", "Status #"], [20, 32, 22, 16, 16, 24, 16, 18, 26, 3, 8, 9], [0, 1, 2, 3, 4, 5, 6, 7, 8], rowsTotal, mandFormula, [9, 10, 11]), "Mandate rules for status change");

  // ---------------- normal typed sheets ----------------
  XLSX.utils.book_append_sheet(wb, plain(["Ticket Key", "Ticket Subtype", "Field (Type)", "Option / Value", "Mandatory", "Edit Roles"], [20, 30, 30, 40, 12, 30]), "Ticket Structure");
  XLSX.utils.book_append_sheet(wb, plain(["Ticket Key", "Ticket Subtype", "Status", "Level (1 = first)", "Days", "Assignee Role", "Notification (Yes/No)", "Notification Type", "CC Email Addresses", "Override Status"], [20, 30, 18, 14, 8, 22, 14, 18, 40, 18]), "Escalation details");
  return wb;
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
      subRows.push([doc.ticketKey, s.ticketType, s.ticketKey, s.prefix, (s.assigneeRoles || []).join(", "), (s.isAutoEscalation || subTypeHasEscalation(s)) ? "Yes" : "No", s.deleted ? "Yes" : "No"]);
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

  // Status Access Roles / Mandate rules are written as readable blocks: one
  // block per sub-type, grouped by ticket type, always naming the sub-type.
  // Ticket Key / Ticket Subtype appear on the first row of a block only, with
  // a blank row between blocks. (The importer still accepts a blank or
  // "(all sub-types)" sub-type for hand-written sheets.)
  const statusBlocks = (doc, rowsOf) => doc.ticketSubType.map((s) => ({ name: s.ticketType || s.ticketKey, rows: rowsOf(s) }));
  const pushStatusBlocks = (sheetRows, doc, blocks) => {
    blocks.forEach((b) => {
      b.rows.forEach((row, i) => sheetRows.push([i === 0 ? doc.ticketKey : "", i === 0 ? b.name : "", ...row]));
      sheetRows.push([]);
    });
  };

  const roleRows = [["Ticket Key", "Ticket Subtype", "Status", "Roles Allowed", "Public", "Status Code (don't edit)"]];
  list.forEach((doc) => {
    pushStatusBlocks(roleRows, doc, statusBlocks(doc, (s) => s.statusWorkFlow.map((st) => [st.label || st.status, (st.roles || []).join(", "), st.public ? "Yes" : "No", st.status])));
  });
  {
    const ws_ = XLSX.utils.aoa_to_sheet(roleRows);
    styleSheet(ws_, (roleRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Status Access Roles");
  }

  const mandateRows = [["Ticket Key", "Ticket Subtype", "Status", "Mandatory Images (Yes/No)", "Mandatory Comments (Yes/No)", "Comment Field Label", "Comment Field Type", "Mandatory New Assignee (Yes/No)", "New Assignee Role", "Status Code (don't edit)"]];
  list.forEach((doc) => {
    pushStatusBlocks(mandateRows, doc, statusBlocks(doc, (s) => s.statusWorkFlow.map((st) => {
      const mand = st.mandatoryCustomFields || [];
      const commentField = (st.customFieldsMetaData || []).find((f) => mand.includes(f.key));
      return [
        st.label || st.status,
        mand.includes("images") ? "Yes" : "No",
        commentField || mand.includes("comments") ? "Yes" : "No",
        commentField ? commentField.label : "",
        commentField ? (commentField.type || "text") : "",
        mand.includes("newAssignee") ? "Yes" : "No",
        (st.assigneeRoles || []).join(", "),
        st.status,
      ];
    })));
  });
  {
    const ws_ = XLSX.utils.aoa_to_sheet(mandateRows);
    styleSheet(ws_, (mandateRows[0] || []).length);
    XLSX.utils.book_append_sheet(wb, ws_, "Mandate rules for status change");
  }

  // Escalation details: one row per escalation level, every field kept (days,
  // assignee, notify on/off, channels, CC emails, override status). Every row
  // carries its own Ticket Key / Sub-type / Status (repeats are greyed) so rows can
  // be sorted, filtered, cut and pasted freely; the Level number sets the order.
  // Ticket types are separated by a blank row and shaded alternately.
  const ESC_HEADER = ["Ticket Key", "Ticket Subtype", "Status", "Level (1 = first)", "Days", "Assignee Role", "Notification (Yes/No)", "Notification Type", "CC Email Addresses", "Override Status", "Status Code (don't edit)"];
  const escRows = [ESC_HEADER];
  const escBand = [-1]; // per sheet row: which ticket type it belongs to (-1 = header / spacer)
  const escRepeat = [[]]; // per sheet row: [key, sub-type, status] cells that repeat the row above
  list.forEach((doc, di) => {
    let firstOfTicket = true;
    doc.ticketSubType.forEach((s) => {
      let firstOfSub = true;
      s.statusWorkFlow.forEach((st) => {
        (st.autoEscalationConfig || []).forEach((e, i) => {
          const days = /^\d+$/.test(String(e.days).trim()) ? Number(e.days) : e.days;
          escRows.push([
            doc.ticketKey,
            s.ticketType || s.ticketKey,
            st.label || st.status,
            i + 1,
            days === undefined ? "" : days,
            e.assignee || "",
            e.notification === false ? "No" : "Yes",
            (e.notificationType || []).join(", "),
            (e.ccAddress || []).join(", "),
            e.overrideStatus || "",
            st.status,
          ]);
          escBand.push(di);
          escRepeat.push([!firstOfTicket, !firstOfSub, i > 0]);
          firstOfTicket = false;
          firstOfSub = false;
        });
      });
    });
    if (!firstOfTicket) { escRows.push([]); escBand.push(-1); escRepeat.push([]); }
  });
  if (escRows.length === 1) escRows.push(["No auto-escalation rules configured."]);
  else if (!escRows[escRows.length - 1].length) { escRows.pop(); escBand.pop(); escRepeat.pop(); }
  {
    const ws_ = XLSX.utils.aoa_to_sheet(escRows);
    styleSheet(ws_, ESC_HEADER.length);
    ws_["!cols"] = [26, 26, 16, 14, 8, 22, 14, 18, 40, 18, 28].map((wch) => ({ wch }));
    escBand.forEach((band, r) => {
      if (band < 0 || r === 0) return;
      for (let c = 0; c < ESC_HEADER.length; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (!ws_[ref]) ws_[ref] = { t: "s", v: "" };
        ws_[ref].s = {
          ...(band % 2 ? { fill: { fgColor: { rgb: "EEF2FA" } } } : {}),
          ...(c <= 2 ? { font: escRepeat[r][c] ? { color: { rgb: "9AA5B8" } } : { bold: true } } : {}),
          alignment: { vertical: "center", wrapText: c === 8, horizontal: c === 3 || c === 4 || c === 6 ? "center" : "left" },
        };
      }
    });
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

/* ---- Bulk escalation: add the same escalation levels to one status of every
   sub-type of the chosen ticket types, in one go. ---- */

// "ACCOUNTS  ISSUE" and "Accounts Issues" are the same ticket type: compare case-, space- and plural-insensitively.
const xnKey = (v) => xnorm(v).toLowerCase().replace(/[^a-z0-9]/g, "").replace(/s$/, "");
function matchTicketType(library, text) {
  if (/^(all|\*|all ticket types)$/i.test(xnorm(text))) return "*";
  const k = xnKey(text);
  if (!k) return "";
  const hit = library.find((d) => xnKey(d.ticketType) === k || xnKey(d.ticketKey) === k);
  return hit ? hit._uid : "";
}

/* Pasted spreadsheet table -> { levels:[{days}], rows:[{id,target,label,roles[]}] }.
   Row = ticket type + one assignee cell per level. A header such as
   "Assignee 2 after 72 hours" / "after 3 days" sets that level's days. */
function parseBulkEscalationPaste(text, library) {
  const lines = String(text || "").split(/\r?\n/).map((l) => l.split("\t").map((c) => c.trim())).filter((c) => c.some(Boolean));
  if (!lines.length) return { levels: [], rows: [] };
  let header = null;
  if (!matchTicketType(library, lines[0][0]) && /ticket|type|assignee|level|role/i.test(lines[0].join(" "))) header = lines.shift();
  const width = Math.max(1, ...lines.map((l) => l.length - 1));
  const levels = [];
  for (let i = 0; i < width; i++) {
    const h = (header && header[i + 1]) || "";
    const hr = h.match(/(\d+(?:\.\d+)?)\s*(?:hours?|hrs?|h)\b/i);
    const dy = h.match(/(\d+(?:\.\d+)?)\s*days?/i);
    const days = hr ? Math.round((parseFloat(hr[1]) / 24) * 100) / 100 : dy ? parseFloat(dy[1]) : null;
    levels.push({ days: String(days != null ? days : i === 0 ? 0 : (parseFloat(levels[i - 1].days) || 0) + 1) });
  }
  const rows = lines.map((l) => ({ id: uid(), target: matchTicketType(library, l[0]), label: l[0] || "", roles: levels.map((_, i) => l[i + 1] || "") }));
  return { levels, rows };
}

/* Dry-run / apply in one: returns the new library plus a per-row report.
   opts: { rows, levels, status, mode: "add"|"replace", sortByDays, extraCc[], ccOnReplace: "keep"|"remove" }
   - add: an assignee already on the status only gets its days updated; others are appended
          (existing levels, with their CC emails, are left alone)
   - replace: the status's levels are replaced by the ones given. What happens to the
              CC emails that were on the replaced levels is decided by ccOnReplace:
                remove -> new levels get only extraCc
                keep   -> a role that already had a level keeps that level's own CC list;
                          any other new role gets every CC email that was on the status;
                          extraCc is added on top (no duplicates)
   - sortByDays: lowest days first (equal days keep their order)
   Roles used by the levels are added to the sub-type and ticket type assignee roles. */
function bulkEscalationPlan(library, opts) {
  const code = xToStatusCode(opts.status);
  const daysOf = (e) => { const n = parseFloat(e.days); return Number.isFinite(n) ? n : Infinity; };
  const dedupe = (list) => { const seen = new Set(); return list.filter((m) => { const k = String(m).trim().toLowerCase(); if (!k || seen.has(k)) return false; seen.add(k); return true; }); };
  const keepCc = opts.mode === "replace" && opts.ccOnReplace === "keep";
  const oldCc = new Map(); // lower-case email -> email, across every replaced level
  let lib = library;
  const rows = opts.rows.map((row) => {
    const rep = { id: row.id, subs: 0, added: 0, updated: 0, removed: 0, ccLevels: 0, missing: [], skipped: "" };
    const desired = [];
    opts.levels.forEach((lv, i) => xSplitRoles(row.roles[i]).forEach((r) => desired.push({ assignee: r, days: String(lv.days).trim() })));
    if (!row.target) { rep.skipped = "pick a ticket type"; return rep; }
    if (!desired.length) { rep.skipped = "no assignees"; return rep; }
    lib = lib.map((d) => {
      if (row.target !== "*" && d._uid !== row.target) return d;
      const subs = d.ticketSubType.map((sub) => {
        const idx = sub.statusWorkFlow.findIndex((w) => xToStatusCode(w.status) === code || xToStatusCode(w.label) === code);
        if (idx < 0) { rep.missing.push(`${d.ticketType || d.ticketKey} › ${sub.ticketType || sub.ticketKey}`); return sub; }
        rep.subs++;
        const w = sub.statusWorkFlow[idx];
        const existing = w.autoEscalationConfig || [];
        let chain = opts.mode === "replace" ? [] : existing.map((e) => ({ ...e }));
        if (opts.mode === "replace") {
          rep.removed += existing.length;
          existing.forEach((e) => { if ((e.ccAddress || []).length) { rep.ccLevels++; e.ccAddress.forEach((m) => oldCc.set(String(m).trim().toLowerCase(), String(m).trim())); } });
        }
        const allOld = dedupe(existing.flatMap((e) => e.ccAddress || []));
        const extra = [...new Set(opts.extraCc || [])];
        desired.forEach((dl) => {
          const at = chain.findIndex((e) => xnorm(e.assignee).toLowerCase() === dl.assignee.toLowerCase());
          if (at >= 0) {
            if (String(chain[at].days) !== dl.days) { chain[at] = { ...chain[at], days: dl.days }; rep.updated++; }
          } else {
            let cc = extra;
            if (keepCc) {
              const same = existing.find((e) => xnorm(e.assignee).toLowerCase() === dl.assignee.toLowerCase());
              cc = dedupe([...(same ? (same.ccAddress || []) : allOld), ...extra]);
            }
            chain.push({ id: uid(), _id: uuidv4(), assignee: dl.assignee, days: dl.days, overrideStatus: "", notification: true, notificationType: [...NOTIF_TYPES], ccAddress: [...cc] });
            rep.added++;
          }
        });
        if (opts.sortByDays) chain = chain.map((e, i) => ({ e, i })).sort((a, b) => (daysOf(a.e) - daysOf(b.e)) || (a.i - b.i)).map((x) => x.e);
        return { ...sub, isAutoEscalation: true, statusWorkFlow: sub.statusWorkFlow.map((x, i) => (i === idx ? { ...x, autoEscalationConfig: chain } : x)) };
      });
      return syncAssigneeRoles({ ...d, ticketSubType: subs });
    });
    return rep;
  });
  const totals = rows.reduce((t, r) => ({ subs: t.subs + r.subs, added: t.added + r.added, updated: t.updated + r.updated, removed: t.removed + r.removed, ccLevels: t.ccLevels + r.ccLevels }), { subs: 0, added: 0, updated: 0, removed: 0, ccLevels: 0 });
  // needsCcChoice: replacing levels that carry CC emails and the user hasn't said keep/remove yet
  const needsCcChoice = opts.mode === "replace" && totals.ccLevels > 0 && !opts.ccOnReplace;
  return { library: lib, rows, totals, ccEmails: [...oldCc.values()], needsCcChoice };
}

/* ------------------------------------------------------------------ */
/* Bulk mandatory fields on status change                              */
/* For the chosen ticket types + statuses, make sure each given field  */
/* exists on the status (added if missing) and is listed in that       */
/* status's mandatoryCustomFields. Mode "remove" only takes the key    */
/* back out of mandatoryCustomFields (the field itself stays).         */
/* A fileUpload field is matched by type, and a new one always gets    */
/* the fixed key "uploadImage".                                        */
/* ------------------------------------------------------------------ */
const MANDATORY_FIELD_TYPES = [
  { value: "fileUpload", label: "Image / file upload" },
  { value: "text", label: "Text / comment" },
  { value: "number", label: "Number" },
  { value: "date", label: "Date" },
  { value: "time", label: "Time" },
];

function labelToKey(label) {
  const words = String(label || "").replace(/[^a-zA-Z0-9]+/g, " ").trim().split(" ").filter(Boolean);
  return words.map((w, i) => (i === 0 ? w.toLowerCase() : w[0].toUpperCase() + w.slice(1).toLowerCase())).join("");
}
const mandatoryFieldKey = (f) => (f.type === "fileUpload" ? STATUS_FILE_UPLOAD_KEY : (xnorm(f.key) || labelToKey(f.label)));

/* A "conflict" = the status already has a field with the SAME LABEL but a
   different key (or type) than the one being added. The caller chooses what
   to do per conflict (opts.resolutions[id]) or for all (opts.defaultRes):
     keep     -> reuse the existing field as-is (its key) and make that mandatory
     override -> rewrite the existing field to the new key/type/label
     skip     -> change nothing for that field on that status
   A conflict with no choice yet is treated as skip and counted as unresolved. */
function bulkMandatoryPlan(library, opts) {
  const codes = (opts.statuses || []).map(xToStatusCode).filter(Boolean);
  const fields = (opts.fields || []).filter((f) => mandatoryFieldKey(f) && (opts.mode === "remove" || f.type === "fileUpload" || xnorm(f.label)));
  const totals = { types: 0, subs: 0, statuses: 0, added: 0, marked: 0, removed: 0, overridden: 0, conflicts: 0, unresolved: 0 };
  const reps = [];
  const conflicts = [];
  if (!codes.length || !fields.length) return { library, reps, conflicts, totals, ready: false };
  const sameLabel = (a, b) => xnorm(a).toLowerCase() === xnorm(b).toLowerCase() && xnorm(a) !== "";
  const lib = library.map((d) => {
    if (!opts.allTypes && !opts.picked.includes(d._uid)) return d;
    const rep = { uid: d._uid, name: d.ticketType || d.ticketKey || "(untitled)", subs: 0, statuses: 0, added: 0, marked: 0, removed: 0, overridden: 0, conflicts: 0, missing: [] };
    const subs = d.ticketSubType.map((sub) => {
      let hit = false;
      const wf = sub.statusWorkFlow.map((w) => {
        const wCode = xToStatusCode(w.status);
        if (!codes.includes(wCode) && !codes.includes(xToStatusCode(w.label))) return w;
        hit = true;
        let cf = [...(w.customFieldsMetaData || [])];
        let mand = [...(w.mandatoryCustomFields || [])];
        let changed = false;
        fields.forEach((f) => {
          const key = mandatoryFieldKey(f);
          const label = xnorm(f.label) || key;
          const existing = cf.find((x) => (f.type === "fileUpload" ? x.type === "fileUpload" : xnorm(x.key).toLowerCase() === key.toLowerCase()));
          const labelHit = existing ? null : cf.find((x) => sameLabel(x.label, label));

          if (opts.mode === "remove") {
            const useKey = (existing || labelHit || {}).key || key;
            if (mand.includes(useKey)) { mand = mand.filter((k) => k !== useKey); rep.removed++; changed = true; }
            return;
          }

          if (labelHit) {
            const cid = `${sub.id}|${wCode}|${f.id}`;
            const res = (opts.resolutions && opts.resolutions[cid]) || opts.defaultRes || "";
            conflicts.push({
              id: cid, typeName: rep.name, subName: sub.ticketType || sub.ticketKey || "(unnamed)",
              statusLabel: w.label || w.status, fieldLabel: label,
              existingKey: labelHit.key, existingType: labelHit.type, newKey: key, newType: f.type, res,
            });
            rep.conflicts++;
            if (!res) return; // unresolved: leave untouched
            if (res === "skip") return;
            if (res === "keep") {
              if (!mand.includes(labelHit.key)) { mand.push(labelHit.key); rep.marked++; changed = true; }
              return;
            }
            // override: rewrite the existing field in place, swap the old key for the new one
            cf = cf.map((x) => (x === labelHit ? { ...newSimpleField(f.type), id: labelHit.id, label, key } : x));
            mand = mand.filter((k) => k !== labelHit.key);
            if (!mand.includes(key)) mand.push(key);
            rep.overridden++; rep.marked++; changed = true;
            return;
          }

          const useKey = existing ? existing.key : key;
          if (!existing) { cf.push({ ...newSimpleField(f.type), label, key: useKey }); rep.added++; changed = true; }
          if (!mand.includes(useKey)) { mand.push(useKey); rep.marked++; changed = true; }
        });
        if (!changed) return w;
        rep.statuses++;
        return { ...w, customFieldsMetaData: cf, mandatoryCustomFields: mand };
      });
      if (hit) rep.subs++; else rep.missing.push(sub.ticketType || sub.ticketKey || "(unnamed)");
      return { ...sub, statusWorkFlow: wf };
    });
    reps.push(rep);
    return { ...d, ticketSubType: subs };
  });
  reps.forEach((r) => { totals.types++; ["subs", "statuses", "added", "marked", "removed", "overridden", "conflicts"].forEach((k) => { totals[k] += r[k]; }); });
  totals.unresolved = conflicts.filter((c) => !c.res).length;
  return { library: lib, reps, conflicts, totals, ready: true };
}

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
    isAutoEscalation: !!sub.isAutoEscalation || subTypeHasEscalation(sub),
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
    if (sub.statusWorkFlow.some((s) => toStatusCode(s.status) === toStatusCode(op.status.status))) {
      return { library, applied: false, message: `Status "${op.status.status}" already exists on ${sub.ticketType || sub.ticketKey}` };
    }
    return withSub({ ...sub, statusWorkFlow: [...sub.statusWorkFlow, statusFromJson(op.status)] }, `Added status "${op.status.label || op.status.status}" to ${sub.ticketType || sub.ticketKey}`);
  }
  if (op.op === "removeStatus") {
    if (!sub.statusWorkFlow.some((s) => toStatusCode(s.status) === toStatusCode(op.statusCode))) {
      return { library, applied: false, message: `Status "${op.statusCode}" not found on ${sub.ticketType || sub.ticketKey}` };
    }
    return withSub({ ...sub, statusWorkFlow: sub.statusWorkFlow.filter((s) => toStatusCode(s.status) !== toStatusCode(op.statusCode)) }, `Removed status "${op.statusCode}" from ${sub.ticketType || sub.ticketKey}`);
  }

  // everything below is additionally scoped to a status
  if (!op.statusCode) return { library, applied: false, message: `Missing statusCode for "${op.op}"` };
  const statusIdx = sub.statusWorkFlow.findIndex((s) => toStatusCode(s.status) === toStatusCode(op.statusCode));
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
      if (newFieldObj.type === "fileUpload") newFieldObj.key = STATUS_FILE_UPLOAD_KEY;
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

    const oldStatuses = new Map((oldSub.statusWorkFlow || []).map((s) => [toStatusCode(s.status), s]));
    const newStatuses = new Map((newSub.statusWorkFlow || []).map((s) => [toStatusCode(s.status), s]));

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

const Label = ({ children, mono, hint }) => (
  <label className="block mb-1.5">
    <span className={`block text-[11px] font-semibold uppercase tracking-wider text-white ${mono ? "font-mono normal-case tracking-normal" : ""}`}>
      {children}
    </span>
    {hint && <span className="block text-[11px] font-normal normal-case tracking-normal text-white mt-0.5">{hint}</span>}
  </label>
);

const Input = ({ value, onChange, placeholder, mono, ...rest }) => (
  <input
    value={value}
    onChange={(e) => onChange(e.target.value)}
    placeholder={placeholder}
    className={`w-full rounded-xl border border-white/10 bg-[#0B0F1C] px-3 py-2 text-sm text-white placeholder:text-slate-600 outline-none transition hover:border-white/20 focus:border-indigo-400 focus:ring-2 focus:ring-indigo-400/20 ${mono ? "font-mono" : ""}`}
    {...rest}
  />
);

const Select = ({ value, onChange, options }) => (
  <select
    value={value}
    onChange={(e) => onChange(e.target.value)}
    className="w-full rounded-xl border border-white/10 bg-[#0B0F1C] px-3 py-2 text-sm text-white outline-none transition hover:border-white/20 focus:border-indigo-400 focus:ring-2 focus:ring-indigo-400/20"
  >
    {options.map((o) => (
      <option key={o.value} value={o.value}>{o.label}</option>
    ))}
  </select>
);

const Toggle = ({ checked, onChange, label }) => (
  <button
    type="button"
    role="switch"
    aria-checked={checked}
    onClick={() => onChange(!checked)}
    className="group flex items-center gap-3 text-sm text-white outline-none"
  >
    <span
      className={`relative inline-flex h-6 w-10 shrink-0 items-center rounded-full border transition-colors ${
        checked ? "border-indigo-400 bg-indigo-400" : "border-white/15 bg-slate-800"
      } group-focus-visible:ring-2 group-focus-visible:ring-indigo-400/40 group-focus-visible:ring-offset-2 group-focus-visible:ring-offset-[#0B0F1C]`}
    >
      <span
        className="inline-block h-4 w-4 transform rounded-full bg-white shadow-sm transition-transform duration-150"
        style={{ transform: checked ? "translateX(18px)" : "translateX(2px)" }}
      />
    </span>
    <span className="select-none">{label}</span>
  </button>
);

const IconBtn = ({ onClick, title, danger, children }) => (
  <button
    type="button"
    onClick={onClick}
    title={title}
    className={`inline-flex items-center justify-center h-7 w-7 rounded-md transition ${danger ? "text-rose-400 hover:bg-rose-500/10 hover:text-rose-300" : "text-white hover:bg-white/5 hover:text-white"}`}
  >
    {children}
  </button>
);

const Btn = ({ onClick, children, variant = "primary", small }) => {
  const base = "inline-flex items-center gap-1.5 rounded-xl font-medium transition active:scale-[0.97]";
  const size = small ? "px-2.5 py-1.5 text-xs" : "px-3.5 py-2 text-sm";
  const styles = {
    primary: "bg-slate-100 text-slate-900 hover:bg-white",
    accent: "bg-gradient-to-b from-indigo-400 to-indigo-500 text-white shadow-md shadow-indigo-500/20 hover:from-indigo-300 hover:to-indigo-400 hover:shadow-indigo-400/30",
    ghost: "bg-transparent text-white hover:bg-white/5 hover:border-white/20 border border-white/10",
    dashed: "border border-dashed border-white/15 text-white hover:border-indigo-400 hover:text-indigo-300 hover:bg-indigo-500/5",
    danger: "bg-rose-500/10 text-rose-200 border border-rose-500/40 hover:bg-rose-500/20 hover:border-rose-400",
  };
  return (
    <button type="button" onClick={onClick} className={`${base} ${size} ${styles[variant]}`}>
      {children}
    </button>
  );
};

const Checkbox = ({ checked, onChange, title, disabled }) => (
  <input
    type="checkbox"
    checked={!!checked}
    disabled={disabled}
    title={title}
    onChange={(e) => onChange(e.target.checked)}
    onClick={(e) => e.stopPropagation()}
    className="h-4 w-4 shrink-0 cursor-pointer rounded accent-indigo-400 disabled:cursor-not-allowed disabled:opacity-50"
  />
);

/* Bottom-right toast stack. Each toast: { id, type: "error"|"success"|"info", title, lines[] } */
const ToastStack = ({ toasts, onDismiss }) => {
  if (!toasts.length) return null;
  const tones = {
    error: { box: "border-rose-500/40 bg-[#2a1220]", icon: <AlertTriangle size={18} className="text-rose-400 mt-0.5 shrink-0" />, title: "text-rose-100" },
    warning: { box: "border-amber-500/40 bg-[#2b2110]", icon: <AlertTriangle size={18} className="text-amber-400 mt-0.5 shrink-0" />, title: "text-amber-100" },
    success: { box: "border-emerald-500/40 bg-[#10241c]", icon: <Check size={18} className="text-emerald-400 mt-0.5 shrink-0" />, title: "text-emerald-100" },
    info: { box: "border-indigo-400/40 bg-[#14183a]", icon: <Sparkles size={18} className="text-indigo-300 mt-0.5 shrink-0" />, title: "text-indigo-100" },
  };
  return (
    <div className="fixed bottom-5 right-5 z-50 flex w-[380px] max-w-[calc(100vw-2.5rem)] flex-col gap-2" role="status" aria-live="polite">
      {toasts.map((t) => {
        const tone = tones[t.type] || tones.info;
        return (
          <div key={t.id} className={`flex items-start gap-3 rounded-xl border px-4 py-3 shadow-2xl shadow-black/40 ${tone.box}`}>
            {tone.icon}
            <div className="min-w-0 flex-1">
              <div className={`text-sm font-semibold ${tone.title}`}>{t.title}</div>
              {t.lines && t.lines.length > 0 && (
                <ul className="mt-1 space-y-0.5 text-xs text-white">
                  {t.lines.map((l, i) => <li key={i} className="break-words">{l}</li>)}
                </ul>
              )}
              {t.action && (
                <button
                  type="button"
                  onClick={() => { t.action.onClick(); onDismiss(t.id); }}
                  className="mt-2 rounded-lg border border-white/25 bg-white/10 px-2.5 py-1 text-xs font-medium text-white hover:bg-white/20"
                >
                  {t.action.label}
                </button>
              )}
            </div>
            <button type="button" onClick={() => onDismiss(t.id)} className="shrink-0 text-white/70 hover:text-white" title="Dismiss"><X size={15} /></button>
          </div>
        );
      })}
    </div>
  );
};

const Chip = ({ text, onRemove, tone = "slate" }) => {
  const tones = {
    slate: "bg-white/5 text-white",
    indigo: "bg-indigo-500/10 text-indigo-300",
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
        className="w-full text-sm outline-none bg-transparent text-white placeholder:text-slate-500"
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
          className={`rounded-md px-2.5 py-1 text-xs font-mono border transition ${active ? "border-indigo-400 bg-indigo-400 text-white" : "border-white/10 text-white hover:border-white/25"}`}
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
      <div className="flex items-center gap-2 text-white font-semibold text-sm">
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

const Accordion = ({ title, subtitle, tone = "slate", defaultOpen, forceOpen, domId, onDelete, onCopy, select, children }) => {
  const [open, setOpen] = useState(!!defaultOpen);
  React.useEffect(() => { if (forceOpen) setOpen(true); }, [forceOpen]);
  const tones = { slate: "border-white/10", indigo: "border-indigo-500/30", amber: "border-amber-500/30" };
  return (
    <div id={domId} className={`rounded-xl border ${tones[tone]} bg-[#111528] mb-3 overflow-hidden ${forceOpen ? "ring-2 ring-indigo-400/40" : ""}`}>
      <div className="flex items-center justify-between px-3 py-2.5 cursor-pointer select-none" onClick={() => setOpen(!open)}>
        <div className="flex items-center gap-2 min-w-0">
          {select && <Checkbox checked={select.checked} onChange={select.onChange} title="Select" />}
          {open ? <ChevronDown size={16} className="text-white shrink-0" /> : <ChevronRight size={16} className="text-white shrink-0" />}
          <div className="min-w-0">
            <div className="text-sm font-medium text-white truncate">{title}</div>
            {subtitle && <div className="text-xs text-white font-mono truncate">{subtitle}</div>}
          </div>
        </div>
        <div className="flex items-center gap-0.5 shrink-0">
        {onCopy && (
          <IconBtn title="Copy (paste it into another ticket type)" onClick={(e) => { e.stopPropagation(); onCopy(); }}>
            <ClipboardCopy size={14} />
          </IconBtn>
        )}
        {onDelete && (
          <IconBtn danger title="Remove" onClick={(e) => { e.stopPropagation(); onDelete(); }}>
            <Trash2 size={14} />
          </IconBtn>
        )}
        </div>
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
        <span className="text-[11px] font-semibold uppercase tracking-wider text-white">Value (label shown)</span>
        <span className="text-[11px] font-semibold uppercase tracking-wider text-white">Key</span>
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
                  <ListTree size={14} className={hasCondition ? "text-indigo-400" : ""} />
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
                    tone="indigo"
                  />
                </div>
              )}
            </div>
          );
        })}
      </div>
      <button type="button" onClick={add} className="mt-2 text-xs font-medium text-indigo-400 hover:text-indigo-300 inline-flex items-center gap-1">
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
        <span className="text-xs text-white">Parent option style:</span>
        <div className="flex gap-1.5">
          {["multiplechoice", "dropdown"].map((t) => (
            <button
              key={t}
              type="button"
              onClick={() => onChange({ ...field, parentOptionType: t })}
              className={`px-2.5 py-1 rounded-md text-xs font-mono border ${field.parentOptionType === t ? "border-indigo-400 bg-indigo-400 text-white" : "border-white/10 text-white"}`}
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
          <div className="pl-3 border-l-2 border-indigo-500/20 space-y-1.5">
            {p.children.map((c) => (
              <div key={c.id} className="grid grid-cols-[1fr_120px_28px] gap-2 items-center">
                <Input value={c.value} onChange={(v) => updateChild(p.id, c.id, { value: v })} placeholder="Child value, e.g. Buttercup" />
                <Input mono value={c.key} onChange={(v) => updateChild(p.id, c.id, { key: v })} placeholder="key" />
                <IconBtn danger title="Remove child" onClick={() => removeChild(p.id, c.id)}><Trash2 size={14} /></IconBtn>
              </div>
            ))}
            <button type="button" onClick={() => addChild(p.id)} className="text-xs font-medium text-indigo-400 hover:text-indigo-300 inline-flex items-center gap-1 mt-1">
              <Plus size={12} /> Add child option
            </button>
          </div>
        </div>
      ))}
      <button type="button" onClick={addParent} className="text-xs font-medium text-indigo-400 hover:text-indigo-300 inline-flex items-center gap-1">
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
            <div className="pl-3 border-l-2 border-indigo-500/20 mb-2">
              <SimpleOptionsEditor options={c.options || []} onChange={(opts) => update(c.id, { options: opts })} />
            </div>
          )}
          <Input mono value={c.formula} onChange={(v) => update(c.id, { formula: v })} placeholder="optional formula, e.g. rate*quantity" />
        </div>
      ))}
      <button type="button" onClick={add} className="text-xs font-medium text-indigo-400 hover:text-indigo-300 inline-flex items-center gap-1">
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
      <button type="button" onClick={() => setOpen(!open)} className="text-xs font-medium text-white hover:text-white inline-flex items-center gap-1">
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
            <TagInput values={field.editRoles || []} onChange={(v) => onChange({ ...field, editRoles: v })} placeholder="role id, press enter" tone="indigo" />
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
    const base = { id: field.id, label: field.label, key: (restrictToDropdown && type === "fileUpload") ? STATUS_FILE_UPLOAD_KEY : field.key, ...advancedDefaults() };
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
          <Label hint="What people see on the form">Field label</Label>
          <Input value={field.label} onChange={(v) => onChange({ ...field, label: v })} placeholder="e.g. Division Code" />
        </div>
        <div>
          <Label mono hint="Internal id this value is stored under — no spaces">Field key</Label>
          {restrictToDropdown && field.type === "fileUpload"
            ? <Input mono value={STATUS_FILE_UPLOAD_KEY} onChange={() => {}} readOnly title="File upload fields always use this key" />
            : <Input mono value={field.key} onChange={(v) => onChange({ ...field, key: v })} placeholder="e.g. divisionCode" />}
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
          <p className="text-xs text-white mb-3">Advanced type — options are fetched live from another ticket's data instead of a fixed list.</p>
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
            className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-xs font-mono text-white outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-400/20"
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
        <p className="text-xs text-white italic">No options needed for this field type.</p>
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
  const [dragIdx, setDragIdx] = useState(null);
  const [overIdx, setOverIdx] = useState(null);
  const update = (id, patch) => onChange(rules.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  const remove = (id) => onChange(rules.filter((r) => r.id !== id));
  const add = () => onChange([...rules, newAutoEscalation()]);
  // Level order = array order (level 1 escalates first). Move one level up/down, or drag it to a new place.
  const move = (from, to) => {
    if (from == null || to < 0 || to >= rules.length || from === to) return;
    const next = [...rules];
    const [m] = next.splice(from, 1);
    next.splice(to, 0, m);
    onChange(next);
  };
  return (
    <div>
      {rules.length > 1 && <p className="text-xs text-white/70 mb-2">Levels escalate in order, Level 1 first. Use the arrows or drag the grip to reorder.</p>}
      {rules.map((r, i) => (
        <div
          key={r.id}
          onDragOver={(e) => { if (dragIdx != null) { e.preventDefault(); setOverIdx(i); } }}
          onDrop={(e) => { e.preventDefault(); move(dragIdx, i); setDragIdx(null); setOverIdx(null); }}
          className={`rounded-lg border bg-amber-500/5 p-3 mb-2.5 transition ${dragIdx === i ? "opacity-40" : ""} ${overIdx === i && dragIdx != null && dragIdx !== i ? "border-indigo-400 ring-2 ring-indigo-400/40" : "border-amber-500/20"}`}
        >
          <div className="flex justify-between items-center mb-2">
            <div className="flex items-center gap-1.5 text-amber-400 text-xs font-semibold">
              <span
                draggable
                onDragStart={(e) => { e.dataTransfer.effectAllowed = "move"; setDragIdx(i); }}
                onDragEnd={() => { setDragIdx(null); setOverIdx(null); }}
                className="cursor-grab active:cursor-grabbing text-white/70 hover:text-white"
                title="Drag to reorder"
              >
                <GripVertical size={14} />
              </span>
              <AlertTriangle size={13} /> Level {i + 1}
              <span className="font-normal text-white/60">{i === 0 ? "— escalates first" : "— then"}</span>
            </div>
            <div className="flex items-center">
              {i > 0 && <IconBtn title="Move up (escalate earlier)" onClick={() => move(i, i - 1)}><ChevronUp size={14} /></IconBtn>}
              {i < rules.length - 1 && <IconBtn title="Move down (escalate later)" onClick={() => move(i, i + 1)}><ChevronDown size={14} /></IconBtn>}
              <IconBtn danger onClick={() => remove(r.id)}><Trash2 size={14} /></IconBtn>
            </div>
          </div>
          <div className="grid grid-cols-3 gap-2 mb-2">
            <div>
              <Label mono hint="Required">Assignee role *</Label>
              <div className={xnorm(r.assignee) ? "" : "rounded-xl ring-1 ring-rose-500/60"}>
                <Input mono value={r.assignee} onChange={(v) => update(r.id, { assignee: v })} placeholder="Dist_Mgr" />
              </div>
            </div>
            <div>
              <Label>Days</Label>
              <Input value={r.days} onChange={(v) => update(r.id, { days: v.replace(/[^0-9]/g, "") })} placeholder="2" />
            </div>
            <div>
              <Label mono>Override status</Label>
              <Input mono value={r.overrideStatus} onChange={(v) => update(r.id, { overrideStatus: toStatusCode(v) })} placeholder="rejected" />
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
        <Plus size={13} /> Add escalation level
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
      tone="indigo"
      onDelete={onDelete}
    >
      <div className="grid grid-cols-2 gap-3 mb-3">
        <div>
          <Label mono hint="Internal id used by the workflow engine — lowercase, no spaces (e.g. resolvedandclosed)">Status code</Label>
          <Input mono value={status.status} onChange={(v) => onChange({ ...status, status: toStatusCode(v) })} placeholder="e.g. open" />
        </div>
        <div>
          <Label hint="What people see on the ticket">Display label</Label>
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
          <TagInput values={status.roles} onChange={(v) => onChange({ ...status, roles: v })} placeholder="e.g. corpadmin" tone="indigo" />
        </div>
        <div>
          <Label>Notify roles</Label>
          <TagInput values={status.notificationRole} onChange={(v) => onChange({ ...status, notificationRole: v })} placeholder="e.g. Dist_Mgr" tone="indigo" />
        </div>
      </div>

      <div className="grid grid-cols-2 gap-3 mb-3">
        <div>
          <Label>Assignee roles at this status</Label>
          <TagInput values={status.assigneeRoles} onChange={(v) => onChange({ ...status, assigneeRoles: v })} placeholder="e.g. corpadmin" tone="indigo" />
        </div>
        <div>
          <Label mono hint="Jump straight to this status code instead, if set">Override status on completion</Label>
          <Input mono value={status.overrideStatus} onChange={(v) => onChange({ ...status, overrideStatus: toStatusCode(v) })} placeholder="e.g. closed" />
        </div>
      </div>

      <div className="mb-3">
        <Label hint="Ticket can't move here until these field keys are filled in">Mandatory fields to move into this status</Label>
        <TagInput values={status.mandatoryCustomFields} onChange={(v) => onChange({ ...status, mandatoryCustomFields: v })} placeholder="field key, press enter" tone="amber" />
      </div>

      <div className="mb-3">
        <Label mono hint="This status can only be reached after all of these are done">Depends on other status codes</Label>
        <TagInput values={status.dependentStatus} onChange={(v) => onChange({ ...status, dependentStatus: v.map(toStatusCode).filter(Boolean) })} placeholder="status code, press enter" tone="indigo" />
      </div>

      <div className="mb-3">
        <Label>Custom fields captured at this status</Label>
        <CustomFieldsList
          fields={status.customFieldsMetaData}
          restrictToDropdown
          onChange={(fs) => {
            // keep mandatory list pointing at the right key when a field key changes
            let mand = status.mandatoryCustomFields;
            fs.forEach((nf) => {
              const of = status.customFieldsMetaData.find((x) => x.id === nf.id);
              if (of && of.key && of.key !== nf.key) mand = mand.map((k) => (k === of.key ? nf.key : k));
            });
            onChange({ ...status, customFieldsMetaData: fs, mandatoryCustomFields: [...new Set(mand)] });
          }}
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
              <div className={`h-2.5 w-2.5 rounded-full ${s.status ? "bg-indigo-400" : "bg-slate-700"}`} />
              <div className="mt-1.5 text-[11px] font-mono text-white whitespace-nowrap">{s.label || s.status || "…"}</div>
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

function SubTypeEditor({ sub, onChange, onDelete, onCopy, forceOpen, select }) {
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
      onCopy={onCopy}
      select={select}
    >
      <div className="flex gap-1 mb-4 border-b border-white/10">
        {tabs.map((t) => (
          <button
            key={t.id}
            type="button"
            onClick={() => setTab(t.id)}
            className={`px-3 py-2 text-xs font-medium border-b-2 -mb-px transition ${tab === t.id ? "border-indigo-400 text-indigo-300" : "border-transparent text-white hover:text-white"}`}
          >
            {t.label}
          </button>
        ))}
      </div>

      {tab === "basic" && (
        <div>
          <div className="grid grid-cols-2 gap-3 mb-3">
            <div>
              <Label mono hint="Internal code for this sub-type — used in ticket IDs, no spaces">Sub-type key</Label>
              <Input mono value={sub.ticketKey} onChange={(v) => onChange({ ...sub, ticketKey: v })} placeholder="e.g. StockGap" />
            </div>
            <div>
              <Label mono hint="Ticket ID template — write the location placeholder exactly as ${locationId} (fixed automatically)">Prefix</Label>
              <Input mono value={sub.prefix} onChange={(v) => onChange({ ...sub, prefix: v })} placeholder="e.g. ${locationId}" />
            </div>
          </div>
          <div className="mb-3">
            <Label hint="The display name people see when choosing this sub-type">Sub-type name</Label>
            <Input value={sub.ticketType} onChange={(v) => onChange({ ...sub, ticketType: v })} placeholder="e.g. DS Stock gaps in plotted plannos" />
          </div>
          <div className="mb-3">
            <Label>Assignee roles — default routing for this sub-type</Label>
            <TagInput values={sub.assigneeRoles || []} onChange={(v) => onChange({ ...sub, assigneeRoles: v })} placeholder="e.g. HOHRPayroll — press enter" tone="indigo" />
          </div>
          <div className="flex gap-6">
            <Toggle checked={sub.deleted} onChange={(v) => onChange({ ...sub, deleted: v })} label="Deleted" />
            <Toggle
              checked={sub.isAutoEscalation || subTypeHasEscalation(sub)}
              onChange={(v) => { if (!v && subTypeHasEscalation(sub)) return; onChange({ ...sub, isAutoEscalation: v }); }}
              label="Auto-escalation enabled"
            />
          </div>
          {subTypeHasEscalation(sub) && (
            <p className="mt-2 text-xs text-white/70">Always on while this sub-type has escalation rules (Status Workflow tab). Remove the rules to switch it off.</p>
          )}
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
              onChange={(patch) => {
                const next = { ...sub, statusWorkFlow: sub.statusWorkFlow.map((x) => (x.id === s.id ? patch : x)) };
                // adding an escalation rule switches auto-escalation on for the sub-type
                onChange(subTypeHasEscalation(next) ? { ...next, isAutoEscalation: true } : next);
              }}
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
    <div className="rounded-xl border border-indigo-500/30 bg-[#0D1424] p-4 mb-4">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2 text-sm font-semibold text-indigo-300">
          <Sparkles size={15} /> Bulk create sub-types
        </div>
        <IconBtn onClick={onClose} title="Close"><X size={14} /></IconBtn>
      </div>

      {!hasAnySource ? (
        <p className="text-sm text-white">
          Nothing to stamp out from yet. Either import a template on the <span className="text-indigo-300 font-medium">Templates</span> tab,
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
                className={`px-2.5 py-1.5 rounded-md text-xs font-medium border transition ${sourceKind === "template" ? "border-indigo-400 bg-indigo-400 text-white" : "border-white/10 text-white"} ${!templates.length ? "opacity-40 cursor-not-allowed" : ""}`}
              >
                Saved template
              </button>
              <button
                type="button"
                onClick={() => subtypeOptions.length && setSourceKind("subtype")}
                disabled={!subtypeOptions.length}
                className={`px-2.5 py-1.5 rounded-md text-xs font-medium border transition ${sourceKind === "subtype" ? "border-indigo-400 bg-indigo-400 text-white" : "border-white/10 text-white"} ${!subtypeOptions.length ? "opacity-40 cursor-not-allowed" : ""}`}
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
            className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-sm text-white outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-400/20 font-mono"
            placeholder={"DS Stock gaps - Store 101\nDS Stock gaps - Store 102\nDS Stock gaps - Store 103"}
          />
          <p className="text-xs text-white mt-1.5">
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
                  className="text-xs font-medium text-indigo-400 hover:text-indigo-300"
                >
                  Update all
                </button>
                <span className="text-slate-700">·</span>
                <button
                  type="button"
                  onClick={() => setReviewChoices({})}
                  className="text-xs font-medium text-white hover:text-white"
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
                        <div className="text-sm text-white truncate">{d.name}</div>
                        <div className="text-[11px] font-mono text-white truncate">already exists as {d.existing.ticketKey}</div>
                      </div>
                      <div className="flex gap-1 shrink-0">
                        <button
                          type="button"
                          onClick={() => setReviewChoices((c) => ({ ...c, [d.name]: false }))}
                          className={`px-2 py-1 rounded text-xs font-medium border transition ${!willUpdate ? "border-slate-500 bg-slate-600 text-white" : "border-white/10 text-white"}`}
                        >
                          No change
                        </button>
                        <button
                          type="button"
                          onClick={() => setReviewChoices((c) => ({ ...c, [d.name]: true }))}
                          className={`px-2 py-1 rounded text-xs font-medium border transition ${willUpdate ? "border-indigo-400 bg-indigo-400 text-white" : "border-white/10 text-white"}`}
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
            <p className="text-xs text-white">
              {newOnes.length} new sub-type{newOnes.length === 1 ? "" : "s"}
              {dupes.length > 0 && <>, {updateCount} update{updateCount === 1 ? "" : "s"}, {dupes.length - updateCount} left unchanged</>}
              {(sourceKind === "template" && chosenTemplate) || (sourceKind === "subtype" && chosenSubtypeOption) ? (
                <> — from <span className="text-indigo-300">{sourceKind === "template" ? chosenTemplate.name : chosenSubtypeOption.label}</span></>
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
      <div className="flex items-center gap-2 text-sm font-semibold text-white mb-1">
        <LayoutTemplate size={15} className="text-indigo-400" /> Import a template
      </div>
      <p className="text-xs text-white mb-3">
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
        className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-xs font-mono text-white outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-400/20 mb-3"
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
      <div className="flex items-center gap-2 text-sm font-semibold text-white mb-1">
        <ListTree size={15} className="text-indigo-400" /> Or turn an existing sub-type into a template
      </div>
      <p className="text-xs text-white mb-3">
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
          className={`inline-flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium border transition ${template.isDefault ? "border-amber-400 bg-amber-500/10 text-amber-300" : "border-white/10 text-white hover:border-white/25"}`}
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
            className={`px-3 py-2 text-xs font-medium border-b-2 -mb-px transition ${tab === t.id ? "border-indigo-400 text-indigo-300" : "border-transparent text-white hover:text-white"}`}
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
        <p className="text-xs text-white">
          Templates are reused across every ticket type — pick one when bulk-creating sub-types.
        </p>
        <Btn small variant="dashed" onClick={onCreateBlank}><Plus size={14} />New blank template</Btn>
      </div>
      {templates.length === 0 ? (
        <div className="text-center py-10 text-white text-sm">
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
  return "indigo";
}

const OP_TONE_CLASSES = {
  indigo: "border-indigo-500/30 bg-indigo-500/5",
  rose: "border-rose-500/30 bg-rose-500/5",
  amber: "border-amber-500/30 bg-amber-500/5",
};

function ChangeOpsReview({ ops, checked, previews, onToggle, onApply }) {
  const selectedCount = ops.filter((_, i) => checked[i]).length;
  return (
    <div className="rounded-xl border border-white/10 bg-[#111528] p-4">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2 text-sm font-semibold text-white">
          <Diff size={15} className="text-indigo-400" /> Proposed changes ({ops.length})
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
                  <span className="text-xs font-mono text-white">{op.op}</span>
                  {op.ticketKey && <span className="text-xs font-mono text-indigo-400">{op.ticketKey}</span>}
                  {op.subTypeKey && <span className="text-xs font-mono text-white">/ {op.subTypeKey}</span>}
                  {op.statusCode && <span className="text-xs font-mono text-white">/ {op.statusCode}</span>}
                </div>
                <div className="text-sm text-white">{op.description || "(no description)"}</div>
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
        <div className="flex items-center gap-2 text-sm font-semibold text-white mb-1">
          <FileSpreadsheet size={15} className="text-indigo-400" /> Upload an edited Excel sheet
        </div>
        <p className="text-xs text-white mb-4">
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
        <p className="text-sm text-white">No differences found against "{targetLabel}" — the sheet matches what's already here.</p>
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
        <div className="flex items-center gap-2 text-sm font-semibold text-white mb-1">
          <Wand2 size={15} className="text-indigo-400" /> Describe changes in plain English
        </div>
        <p className="text-xs text-white mb-4">
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
          className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] px-3 py-2 text-sm text-white outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-400/20 mb-3"
          placeholder={"e.g.\n- On sub-type StockGap, remove assignee role Str_Mgr and add Regional_Mgr\n- Add a new status \"On Hold\" (code onHold) to StockGap, visible to viewers, settable by Dist_Mgr\n- On the On Hold status, add a mandatory text field \"Remark\" (key remark)"}
        />
        {error && <p className="text-xs text-rose-400 mb-3">{error}</p>}
        <Btn variant="accent" onClick={analyze}>
          <Wand2 size={14} />{loading ? "Analyzing…" : "Analyze changes"}
        </Btn>
      </div>

      {ops && ops.length === 0 && (
        <p className="text-sm text-white">No changes were proposed — try being more specific about what should change.</p>
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
            className={`px-3 py-2 text-sm font-medium border-b-2 -mb-px transition ${mode === m.id ? "border-indigo-400 text-indigo-300" : "border-transparent text-white hover:text-white"}`}
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
          <Label mono hint="Unique internal code for this whole ticket type, no spaces">Ticket key</Label>
          <Input mono value={doc.ticketKey} onChange={(v) => onChange({ ...doc, ticketKey: v })} placeholder="e.g. DSHHPLv1" />
        </div>
        <div>
          <Label mono hint="Which tenant/organization this ticket type belongs to">Tenant ID</Label>
          <Input mono value={doc.tenantId} onChange={(v) => onChange({ ...doc, tenantId: v })} placeholder="e.g. abc" />
        </div>
      </div>
      <div className="mb-4">
        <Label hint="The display name people see when creating this ticket">Ticket type name</Label>
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
        <Label hint="Required — export is blocked until at least one role is added">Creators * — roles allowed to raise this ticket</Label>
        <div className={doc.creators.length ? "" : "rounded-xl ring-1 ring-rose-500/60"}>
          <TagInput values={doc.creators} onChange={(v) => onChange({ ...doc, creators: v })} placeholder="e.g. merchandiseanalyst — press enter" tone="indigo" />
        </div>
        {!doc.creators.length && <p className="mt-1 text-xs text-rose-300">Add at least one creator role.</p>}
      </div>
      <div>
        <Label hint="Required — export is blocked until at least one role is added">Viewers * — roles allowed to view this ticket</Label>
        <div className={doc.viewers.length ? "" : "rounded-xl ring-1 ring-rose-500/60"}>
          <TagInput values={doc.viewers} onChange={(v) => onChange({ ...doc, viewers: v })} placeholder="e.g. vendorRole — press enter" tone="indigo" />
        </div>
        {!doc.viewers.length && <p className="mt-1 text-xs text-rose-300">Add at least one viewer role.</p>}
      </div>
      <div>
        <Label>Assignees — roles this ticket can be routed to</Label>
        <TagInput values={doc.assignee} onChange={(v) => onChange({ ...doc, assignee: v })} placeholder="e.g. Dist_Mgr — press enter" tone="indigo" />
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* JSON tab                                                             */
/* ------------------------------------------------------------------ */

function JsonTab({ doc, onBeforeExport }) {
  const [copied, setCopied] = useState(false);
  const json = useMemo(() => JSON.stringify(docToJson(doc) || {}, null, 2), [doc]);
  const missing = missingRequiredRoles(doc);
  const problems = [
    ...(missing.length ? [`${missing.join(" and ")} ${missing.length === 1 ? "is" : "are"} required but empty (Roles tab)`] : []),
    ...assigneeProblems(doc).map((p) => p.text),
  ];

  const copy = async () => {
    if (onBeforeExport && !onBeforeExport()) return;
    try {
      await navigator.clipboard.writeText(json);
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    } catch (e) { /* clipboard unavailable */ }
  };

  const download = () => {
    if (onBeforeExport && !onBeforeExport()) return;
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
      {problems.length > 0 && (
        <div className="mb-3 flex items-start gap-2 rounded-xl border border-rose-500/40 bg-rose-500/10 px-4 py-3 text-sm text-rose-100">
          <AlertTriangle size={16} className="mt-0.5 shrink-0 text-rose-400" />
          <div>
            <div className="font-medium">Copy and Download are blocked until this is fixed:</div>
            <ul className="mt-1 list-disc space-y-0.5 pl-4 text-xs">{problems.map((p, i) => <li key={i}>{p}</li>)}</ul>
          </div>
        </div>
      )}
      <div className="flex items-center justify-between mb-3">
        <p className="text-sm text-white">Fields left blank are automatically left out of the generated JSON.</p>
        <div className="flex gap-2">
          <Btn variant="ghost" small onClick={copy}>
            {copied ? <Check size={14} /> : <Copy size={14} />} {copied ? "Copied" : "Copy"}
          </Btn>
          <Btn variant="accent" small onClick={download}><Download size={14} />Download JSON</Btn>
        </div>
      </div>
      <pre className="rounded-xl bg-slate-900 text-white text-xs leading-relaxed p-4 overflow-auto font-mono max-h-[70vh]">
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

function LibraryTab({ library, selectedUid, onSelectType, onSelectSubtype, onNew, onDuplicate, onDelete, onBulkDelete, onExportAll, clip, onCopySubs, onPasteSubs }) {
  const [term, setTerm] = useState("");
  const [expanded, setExpanded] = useState({});
  const [pickedTypes, setPickedTypes] = useState(() => new Set()); // ticket type _uid
  const [pickedSubs, setPickedSubs] = useState(() => new Set()); // `${_uid}::${sub.id}`

  const rows = library.map((d) => {
    const selfMatch = !term || matchesTerm(d.ticketType, term) || matchesTerm(d.ticketKey, term);
    const subMatches = term ? d.ticketSubType.filter((s) => matchesTerm(s.ticketType, term) || matchesTerm(s.ticketKey, term)) : d.ticketSubType;
    const show = !term || selfMatch || subMatches.length > 0;
    return { d, selfMatch, subMatches, show };
  }).filter((r) => r.show);

  // Selection is derived against the live library so stale picks (after a delete/import) never count.
  const liveTypeUids = new Set(library.map((d) => d._uid));
  const typeUids = [...pickedTypes].filter((u) => liveTypeUids.has(u));
  const subPicks = [];
  library.forEach((d) => {
    if (pickedTypes.has(d._uid)) return; // the whole type goes, its sub-types are implied
    d.ticketSubType.forEach((s) => { if (pickedSubs.has(`${d._uid}::${s.id}`)) subPicks.push({ uid: d._uid, subId: s.id }); });
  });
  const allVisibleTypesPicked = rows.length > 0 && rows.every(({ d }) => pickedTypes.has(d._uid));
  const toggle = (setter, key, on) => setter((prev) => { const next = new Set(prev); if (on) next.add(key); else next.delete(key); return next; });
  const clearPicks = () => { setPickedTypes(new Set()); setPickedSubs(new Set()); };
  const toggleAllVisible = (on) => setPickedTypes((prev) => { const next = new Set(prev); rows.forEach(({ d }) => (on ? next.add(d._uid) : next.delete(d._uid))); return next; });
  const deleteSelected = () => { if (onBulkDelete({ typeUids, subs: subPicks })) clearPicks(); };
  // Copy everything ticked: whole ticket types contribute all their sub-types.
  const copySelected = () => {
    const picked = [
      ...library.filter((d) => pickedTypes.has(d._uid)).flatMap((d) => d.ticketSubType.map((s) => ({ s, from: d.ticketType || d.ticketKey }))),
      ...subPicks.map(({ uid: u, subId }) => { const d = library.find((x) => x._uid === u); return { s: d.ticketSubType.find((x) => x.id === subId), from: d.ticketType || d.ticketKey }; }),
    ].filter((p) => p.s);
    onCopySubs(picked.map((p) => p.s), [...new Set(picked.map((p) => p.from))].join(", "));
  };

  return (
    <div>
      <div className="flex items-center gap-2 mb-4">
        <div className="relative flex-1">
          <input
            value={term}
            onChange={(e) => setTerm(e.target.value)}
            placeholder="Search ticket types or sub-types…"
            className="w-full rounded-lg border border-white/10 bg-[#0B0F1C] pl-3 pr-3 py-2.5 text-sm text-white placeholder:text-slate-600 outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-400/20"
          />
        </div>
        <Btn variant="ghost" onClick={onExportAll}><Download size={15} />Export all ({library.length})</Btn>
        <Btn variant="accent" onClick={onNew}><Plus size={15} />New ticket type</Btn>
      </div>

      {rows.length === 0 && (
        <div className="text-center py-16 text-white text-sm">
          {library.length === 0 ? "No ticket types yet — import a file or create one." : "No matches for that search."}
        </div>
      )}

      {rows.length > 0 && (
        <div className="mb-3 flex flex-wrap items-center gap-3 rounded-xl border border-white/10 bg-[#0D1220] px-4 py-2.5">
          <label className="flex cursor-pointer items-center gap-2 text-sm text-white">
            <Checkbox checked={allVisibleTypesPicked} onChange={toggleAllVisible} title="Select all ticket types" />
            Select all ticket types
          </label>
          {(typeUids.length > 0 || subPicks.length > 0) ? (
            <>
              <span className="text-sm text-indigo-200">
                {[typeUids.length ? `${typeUids.length} ticket type${typeUids.length === 1 ? "" : "s"}` : "", subPicks.length ? `${subPicks.length} sub-type${subPicks.length === 1 ? "" : "s"}` : ""].filter(Boolean).join(" + ")} selected
              </span>
              <div className="ml-auto flex items-center gap-2">
                <Btn variant="ghost" small onClick={clearPicks}>Clear</Btn>
                <Btn variant="ghost" small onClick={copySelected}><ClipboardCopy size={13} />Copy sub-types</Btn>
                <Btn variant="danger" small onClick={deleteSelected}><Trash2 size={13} />Delete selected</Btn>
              </div>
            </>
          ) : (
            <span className="text-xs text-white/70">Tick ticket types, or expand one and tick its sub-types, to copy or delete several at once.</span>
          )}
        </div>
      )}

      <div className="space-y-2">
        {rows.map(({ d, subMatches }) => {
          const isSelected = d._uid === selectedUid;
          const isOpen = term ? subMatches.length > 0 : !!expanded[d._uid];
          const typePicked = pickedTypes.has(d._uid);
          return (
            <div key={d._uid} className={`rounded-xl border bg-[#111528] overflow-hidden ${typePicked ? "border-rose-400/50" : isSelected ? "border-indigo-500/50" : "border-white/10"}`}>
              <div className="flex items-center gap-3 px-4 py-3">
                <Checkbox checked={typePicked} onChange={(on) => toggle(setPickedTypes, d._uid, on)} title="Select this ticket type" />
                <button
                  onClick={() => setExpanded((e) => ({ ...e, [d._uid]: !e[d._uid] }))}
                  className="text-white hover:text-white"
                >
                  {isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                </button>
                <button className="flex-1 min-w-0 text-left" onClick={() => onSelectType(d._uid)}>
                  <div className="text-sm font-medium text-white truncate">{d.ticketType || "(untitled ticket type)"}</div>
                  <div className="text-xs font-mono text-white truncate">{d.ticketKey || "no key"} · {d.ticketSubType.length} sub-type{d.ticketSubType.length === 1 ? "" : "s"}</div>
                </button>
                <IconBtn title="Copy all its sub-types (then paste into another ticket type)" onClick={() => onCopySubs(d.ticketSubType, d.ticketType || d.ticketKey)}><ClipboardCopy size={14} /></IconBtn>
                {clip && <IconBtn title={`Paste ${clip.subs.length} copied sub-type${clip.subs.length === 1 ? "" : "s"} here`} onClick={() => onPasteSubs(d._uid)}><ClipboardPaste size={14} /></IconBtn>}
                <IconBtn title="Duplicate ticket type" onClick={() => onDuplicate(d._uid)}><Copy size={14} /></IconBtn>
                <IconBtn danger title="Delete ticket type" onClick={() => onDelete(d._uid)}><Trash2 size={14} /></IconBtn>
              </div>
              {isOpen && (
                <div className="border-t border-white/5 divide-y divide-white/5">
                  {d.ticketSubType.length === 0 && (
                    <div className="px-4 py-2.5 text-xs text-white italic">No sub-types yet.</div>
                  )}
                  {(term ? subMatches : d.ticketSubType).map((s) => (
                    <div key={s.id} className="flex items-center gap-3 px-4 py-2.5 pl-11 hover:bg-white/5 transition">
                      <Checkbox
                        checked={typePicked || pickedSubs.has(`${d._uid}::${s.id}`)}
                        disabled={typePicked}
                        onChange={(on) => toggle(setPickedSubs, `${d._uid}::${s.id}`, on)}
                        title={typePicked ? "Included because the whole ticket type is selected" : "Select this sub-type"}
                      />
                      <button onClick={() => onSelectSubtype(d._uid, s.id)} className="flex min-w-0 flex-1 items-center gap-2 text-left">
                        <ListTree size={13} className="text-white shrink-0" />
                        <span className="text-sm text-white truncate">{s.ticketType || "(untitled sub-type)"}</span>
                        <span className="text-xs font-mono text-white truncate">{s.ticketKey}</span>
                      </button>
                      <IconBtn title="Copy this sub-type" onClick={() => onCopySubs([s], d.ticketType || d.ticketKey)}><ClipboardCopy size={14} /></IconBtn>
                    </div>
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
/* Bulk escalation                                                     */
/* ------------------------------------------------------------------ */

function StepTitle({ n, title, hint }) {
  return (
    <div className="mb-2 flex items-start gap-2.5">
      <span className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-indigo-500 text-[11px] font-bold text-white">{n}</span>
      <div>
        <p className="text-sm font-semibold text-white">{title}</p>
        {hint && <p className="text-xs text-white/70">{hint}</p>}
      </div>
    </div>
  );
}

/* Big selectable option (radio behaviour) with a one-line explanation */
function ChoiceCard({ selected, onClick, title, desc, tone }) {
  const ring = selected ? (tone === "danger" ? "border-rose-400 bg-rose-500/10" : "border-indigo-400 bg-indigo-500/10") : "border-white/10 hover:border-white/30";
  return (
    <button type="button" onClick={onClick} className={`flex w-full items-start gap-3 rounded-xl border p-3 text-left ${ring}`}>
      <span className={`mt-0.5 flex h-4 w-4 shrink-0 items-center justify-center rounded-full border ${selected ? "border-white bg-white" : "border-white/40"}`}>
        {selected && <span className={`h-2 w-2 rounded-full ${tone === "danger" ? "bg-rose-500" : "bg-indigo-500"}`} />}
      </span>
      <span>
        <span className="block text-sm font-medium text-white">{title}</span>
        <span className="block text-xs text-white/70">{desc}</span>
      </span>
    </button>
  );
}

function BulkEscalationTab({ library, onApply }) {
  const [status, setStatus] = useState("");
  const [mode, setMode] = useState("add");
  const [ccOnReplace, setCcOnReplace] = useState(""); // "" (not chosen yet) | keep | remove
  const [sortByDays, setSortByDays] = useState(true);
  const [extraCcText, setExtraCcText] = useState("");
  const [levels, setLevels] = useState([{ days: "0" }, { days: "3" }]);
  const [rows, setRows] = useState(() => [{ id: uid(), target: "", label: "", roles: ["", ""] }]);
  const [pasted, setPasted] = useState(false);
  const [showPaste, setShowPaste] = useState(false);

  const cellInput = "w-full rounded-md border border-white/10 bg-[#0B0F1C] px-2 py-1.5 text-sm text-white placeholder:text-slate-600 outline-none focus:border-indigo-400";
  const extraCc = extraCcText.split(/[,;\s]+/).map((x) => x.trim()).filter(Boolean);
  const plan = useMemo(
    () => bulkEscalationPlan(library, { rows, levels, status, mode, sortByDays, extraCc, ccOnReplace }),
    [library, rows, levels, status, mode, sortByDays, extraCcText, ccOnReplace], // eslint-disable-line react-hooks/exhaustive-deps
  );
  const changes = plan.totals.added + plan.totals.updated + plan.totals.removed;

  // statuses that exist in the library (click to fill the box) and role names already used (autocomplete)
  const foundStatuses = useMemo(() => {
    const map = new Map();
    library.forEach((d) => d.ticketSubType.forEach((sub) => sub.statusWorkFlow.forEach((w) => {
      const c = xToStatusCode(w.status || w.label);
      if (!c) return;
      const cur = map.get(c) || { code: c, label: w.label || w.status, count: 0 };
      cur.count++;
      map.set(c, cur);
    })));
    return [...map.values()].sort((a, b) => b.count - a.count).slice(0, 14);
  }, [library]);
  const knownRoles = useMemo(() => [...new Set(library.flatMap((d) => d.assignee || []))].sort(), [library]);

  const setRow = (id, patch) => setRows((rs) => rs.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  const setRole = (id, i, v) => setRows((rs) => rs.map((r) => (r.id === id ? { ...r, roles: r.roles.map((x, j) => (j === i ? v : x)) } : r)));
  const addRow = () => setRows((rs) => [...rs, { id: uid(), target: "", label: "", roles: levels.map(() => "") }]);
  const addLevel = () => {
    setLevels((ls) => [...ls, { days: String((parseFloat(ls[ls.length - 1]?.days) || 0) + 1) }]);
    setRows((rs) => rs.map((r) => ({ ...r, roles: [...r.roles, ""] })));
  };
  const removeLevel = (i) => {
    if (levels.length < 2) return;
    setLevels((ls) => ls.filter((_, j) => j !== i));
    setRows((rs) => rs.map((r) => ({ ...r, roles: r.roles.filter((_, j) => j !== i) })));
  };
  const onPaste = (e) => {
    const text = e.clipboardData.getData("text");
    if (!text.trim()) return;
    e.preventDefault();
    const parsed = parseBulkEscalationPaste(text, library);
    if (!parsed.rows.length) return;
    setLevels(parsed.levels);
    setRows(parsed.rows);
    setPasted(true);
  };
  const apply = () => {
    if (!status.trim()) { window.alert("Step 1: type the status you want to update (for example: open)."); return; }
    if (plan.needsCcChoice) { window.alert("Step 2: the levels being replaced have CC emails.\n\nChoose whether to remove those CC emails or keep them on the new levels."); return; }
    if (!changes) return;
    const t = plan.totals;
    const where = t.subs;
    const parts = [t.added ? `add ${t.added} level${t.added === 1 ? "" : "s"}` : "", t.updated ? `update days on ${t.updated}` : "", t.removed ? `remove ${t.removed} existing level${t.removed === 1 ? "" : "s"}` : ""].filter(Boolean);
    const ccLine = mode === "replace" && t.ccLevels > 0 ? `\n\nCC emails on the replaced levels will be ${ccOnReplace === "keep" ? "KEPT (copied onto the new levels)" : "REMOVED"}.` : "";
    if (!window.confirm(`On the "${status}" status of ${where} sub-type${where === 1 ? "" : "s"}: ${parts.join(", ")}.${ccLine}\n\nContinue?`)) return;
    onApply(plan.library, { ...plan.totals, ccKept: mode === "replace" && ccOnReplace === "keep" });
  };

  return (
    <div className="space-y-6">
      <div className="rounded-xl border border-white/10 bg-[#0D1220] p-4 text-sm text-white">
        <p className="mb-1 font-medium">Set up auto-escalation for many sub-types at once</p>
        <p className="text-xs text-white/70">
          Answer the four steps below. Nothing is changed until you press <span className="font-semibold">Apply</span> at the bottom, and the table shows exactly what will happen first.
        </p>
      </div>

      {/* 1 — status */}
      <div>
        <StepTitle n={1} title="Which status should get the escalation?" hint="Every sub-type of the ticket types you pick in step 3 is updated on this one status." />
        <div className="md:w-1/2"><Input mono value={status} onChange={setStatus} placeholder="e.g. open" /></div>
        {foundStatuses.length > 0 && (
          <div className="mt-2 flex flex-wrap items-center gap-1.5">
            <span className="text-xs text-white/60">Statuses in your data — click one:</span>
            {foundStatuses.map((st) => (
              <button
                key={st.code} type="button" onClick={() => setStatus(st.code)}
                className={`rounded-full border px-2.5 py-0.5 text-xs ${xToStatusCode(status) === st.code ? "border-indigo-400 bg-indigo-500/20 text-indigo-100" : "border-white/10 text-white/80 hover:border-white/30"}`}
              >{st.label || st.code} <span className="text-white/50">· {st.count}</span></button>
            ))}
          </div>
        )}
      </div>

      {/* 2 — existing levels */}
      <div>
        <StepTitle n={2} title="What if that status already has escalation levels?" />
        <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
          <ChoiceCard
            selected={mode === "add"} onClick={() => setMode("add")}
            title="Keep them and add mine"
            desc="Existing levels (and their CC emails) stay as they are. If one of your roles is already there, only its days are updated."
          />
          <ChoiceCard
            selected={mode === "replace"} onClick={() => setMode("replace")} tone="danger"
            title="Replace them with mine"
            desc="Existing levels on that status are deleted. Only the levels you enter below remain."
          />
        </div>

        {mode === "replace" && (
          <div className={`mt-3 rounded-xl border p-4 ${plan.needsCcChoice ? "border-amber-500/40 bg-amber-500/10" : "border-white/10 bg-[#111528]"}`}>
            <p className="text-sm font-medium text-white">What should happen to the CC emails on the levels being replaced?</p>
            {plan.totals.ccLevels === 0 ? (
              <p className="mt-1 text-xs text-white/70">None of the levels being replaced have CC emails, so there is nothing to decide here.</p>
            ) : (
              <>
                <p className="mt-1 text-xs text-white/70">
                  {plan.totals.ccLevels} level{plan.totals.ccLevels === 1 ? "" : "s"} being replaced {plan.totals.ccLevels === 1 ? "has" : "have"} CC emails
                  {plan.ccEmails.length ? <>: <span className="font-mono text-white/90">{plan.ccEmails.slice(0, 4).join(", ")}{plan.ccEmails.length > 4 ? ` +${plan.ccEmails.length - 4} more` : ""}</span></> : ""}.
                </p>
                <div className="mt-3 grid grid-cols-1 gap-3 md:grid-cols-2">
                  <ChoiceCard
                    selected={ccOnReplace === "remove"} onClick={() => setCcOnReplace("remove")}
                    title="Remove them"
                    desc="The new levels start with no CC emails (apart from the ones you type in step 4)."
                  />
                  <ChoiceCard
                    selected={ccOnReplace === "keep"} onClick={() => setCcOnReplace("keep")}
                    title="Keep them"
                    desc="Copy the old CC emails onto the new levels. A role that already had a level keeps that level's own CC list; new roles get all of the old CC emails."
                  />
                </div>
                {plan.needsCcChoice && <p className="mt-2 text-xs font-medium text-amber-200">Please choose one — Apply is disabled until you do.</p>}
              </>
            )}
          </div>
        )}
      </div>

      {/* 3 — who and when */}
      <div>
        <div className="flex items-start justify-between gap-3">
          <StepTitle n={3} title="Who should it escalate to, and after how many days?" hint="One row per ticket type. For each level, type the assignee role and set the number of days in the column title." />
          <button type="button" onClick={() => setShowPaste((v) => !v)} className="shrink-0 text-xs font-medium text-indigo-300 hover:text-indigo-200">
            {showPaste ? "Hide Excel paste" : "Have it in Excel? Paste it"}
          </button>
        </div>
        {showPaste && (
          <div className="mb-3 rounded-xl border border-white/10 bg-[#0D1220] p-3">
            <p className="text-xs text-white/70">Copy the table in Excel — ticket type in the first column, one role column per level — click the box and press Ctrl+V. A header like “Assignee 2 after 72 hours” sets that level's days.</p>
            <textarea
              onPaste={onPaste} rows={2} placeholder="Click here, then Ctrl+V"
              className="mt-2 w-full rounded-lg border border-dashed border-white/20 bg-[#0B0F1C] px-3 py-2 text-xs text-white placeholder:text-slate-500 outline-none focus:border-indigo-400"
            />
            {pasted && <p className="mt-1 text-xs text-emerald-300">Table loaded — check the rows below.</p>}
          </div>
        )}
        <datalist id="bulk-esc-roles">{knownRoles.map((r) => <option key={r} value={r} />)}</datalist>
        <div className="overflow-x-auto rounded-xl border border-white/10 bg-[#111528]">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-[11px] uppercase tracking-wider text-white">
                <th className="px-3 py-2.5 min-w-[210px]">Ticket type</th>
                {levels.map((lv, i) => (
                  <th key={i} className="px-3 py-2.5 min-w-[200px]">
                    <div className="flex items-center gap-2">
                      <span>Level {i + 1}</span>
                      <span className="font-normal normal-case text-white/70">after</span>
                      <input
                        value={lv.days}
                        onChange={(e) => setLevels((ls) => ls.map((x, j) => (j === i ? { days: e.target.value.replace(/[^0-9.]/g, "") } : x)))}
                        className="w-14 rounded-md border border-white/10 bg-[#0B0F1C] px-2 py-1 text-center text-sm text-white outline-none focus:border-indigo-400"
                        title="Days after which this level is notified (0 = straight away)"
                      />
                      <span className="font-normal normal-case text-white/70">days</span>
                      {levels.length > 1 && <IconBtn danger title="Remove this level" onClick={() => removeLevel(i)}><X size={13} /></IconBtn>}
                    </div>
                  </th>
                ))}
                <th className="px-3 py-2.5"><button type="button" onClick={addLevel} className="inline-flex items-center gap-1 text-xs font-medium normal-case text-indigo-300 hover:text-indigo-200"><Plus size={13} />Add level</button></th>
                <th className="px-3 py-2.5 min-w-[200px]">What will happen</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {rows.map((r, ri) => {
                const rep = plan.rows[ri] || {};
                return (
                  <tr key={r.id} className="border-t border-white/5 align-top">
                    <td className="px-3 py-2">
                      <select value={r.target} onChange={(e) => setRow(r.id, { target: e.target.value })} className={`${cellInput} ${r.target ? "" : "ring-1 ring-rose-500/60"}`}>
                        <option value="">{r.label ? `“${r.label}” — pick a ticket type` : "Pick a ticket type…"}</option>
                        <option value="*">All ticket types</option>
                        {library.map((d) => <option key={d._uid} value={d._uid}>{d.ticketType || d.ticketKey}</option>)}
                      </select>
                    </td>
                    {levels.map((_, i) => (
                      <td key={i} className="px-3 py-2">
                        <input list="bulk-esc-roles" value={r.roles[i] || ""} onChange={(e) => setRole(r.id, i, e.target.value)} placeholder="Role, e.g. Manager" className={`${cellInput} font-mono`} />
                      </td>
                    ))}
                    <td />
                    <td className="px-3 py-2 text-xs leading-5">
                      {rep.skipped ? (
                        <span className="text-amber-300">Not applied yet — {rep.skipped}</span>
                      ) : (
                        <>
                          <span className="text-white">{rep.subs} sub-type{rep.subs === 1 ? "" : "s"}</span>
                          <span className="text-emerald-300"> · +{rep.added} added</span>
                          {rep.updated > 0 && <span className="text-sky-300"> · {rep.updated} days updated</span>}
                          {rep.removed > 0 && <span className="text-rose-300"> · −{rep.removed} replaced</span>}
                          {rep.missing.length > 0 && <div className="text-amber-300" title={rep.missing.join("\n")}>{rep.missing.length} sub-type{rep.missing.length === 1 ? "" : "s"} skipped: no “{status || "…"}” status</div>}
                        </>
                      )}
                    </td>
                    <td className="px-2 py-2">{rows.length > 1 && <IconBtn danger title="Remove this row" onClick={() => setRows((rs) => rs.filter((x) => x.id !== r.id))}><Trash2 size={14} /></IconBtn>}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="mt-2"><Btn variant="dashed" small onClick={addRow}><Plus size={13} />Add another ticket type</Btn></div>
        <p className="mt-2 text-xs text-white/70">Tip: choose “All ticket types” in a row to use the same roles everywhere. Put several roles in one box by separating them with commas.</p>
      </div>

      {/* 4 — options */}
      <div>
        <StepTitle n={4} title="Optional extras" />
        <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
          <div>
            <Label hint="Added to the NEW levels only. Separate several with commas. Leave empty for none.">CC emails for the new levels</Label>
            <Input value={extraCcText} onChange={setExtraCcText} placeholder="ops@company.com, lead@company.com" />
          </div>
          <label className="flex cursor-pointer items-start gap-2 pt-1 text-sm text-white">
            <Checkbox checked={sortByDays} onChange={setSortByDays} title="Sort levels by days" />
            <span>Sort the levels by days<span className="block text-xs text-white/70">Lowest days first, so the order always matches the timing.</span></span>
          </label>
        </div>
      </div>

      <div className="flex flex-wrap items-center gap-3 border-t border-white/10 pt-4">
        <span className="text-sm text-white">
          {plan.needsCcChoice
            ? "Choose what to do with the CC emails (step 2) to continue"
            : changes ? `${plan.totals.added} to add · ${plan.totals.updated} to update · ${plan.totals.removed} to remove, across ${plan.totals.subs} sub-types` : "Nothing to change yet — fill in the steps above"}
        </span>
        <div className="ml-auto"><Btn variant="accent" onClick={apply}><Check size={15} />Apply to all sub-types</Btn></div>
      </div>
      <p className="text-xs text-white/70">The roles you enter are also added to each sub-type's and ticket type's assignee roles, which the export check requires. Sub-types stay switched on for auto-escalation.</p>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Bulk mandatory fields tab                                           */
/* ------------------------------------------------------------------ */

const presetResolutionImages = () => ({ id: uid(), type: "fileUpload", label: "Resolution Images", key: "" });
const presetActionTaken = () => ({ id: uid(), type: "text", label: "Action Taken", key: "" });

function BulkMandatoryTab({ library, onApply }) {
  const [allTypes, setAllTypes] = useState(true);
  const [picked, setPicked] = useState([]);
  const [statusText, setStatusText] = useState("");
  const [mode, setMode] = useState("add");
  const [fields, setFields] = useState([presetResolutionImages(), presetActionTaken()]);
  const [defaultRes, setDefaultRes] = useState(""); // "" | keep | override | skip — for same-label/different-key conflicts
  const [resolutions, setResolutions] = useState({}); // per-conflict overrides of the default

  const cellInput = "w-full rounded-md border border-white/10 bg-[#0B0F1C] px-2 py-1.5 text-sm text-white placeholder:text-slate-600 outline-none focus:border-indigo-400";
  const statuses = statusText.split(/[,;\n]+/).map((x) => x.trim()).filter(Boolean);
  const statusCodes = statuses.map(xToStatusCode);

  // statuses that actually exist in the chosen ticket types (click to add)
  const found = useMemo(() => {
    const map = new Map();
    library.forEach((d) => {
      if (!allTypes && !picked.includes(d._uid)) return;
      d.ticketSubType.forEach((sub) => sub.statusWorkFlow.forEach((w) => {
        const code = xToStatusCode(w.status || w.label);
        if (!code) return;
        const cur = map.get(code) || { code, label: w.label || w.status, count: 0 };
        cur.count++;
        map.set(code, cur);
      }));
    });
    return [...map.values()].sort((a, b) => b.count - a.count);
  }, [library, allTypes, picked]);

  const plan = useMemo(
    () => bulkMandatoryPlan(library, { allTypes, picked, statuses, fields, mode, defaultRes, resolutions }),
    [library, allTypes, picked, statusText, fields, mode, defaultRes, resolutions], // eslint-disable-line react-hooks/exhaustive-deps
  );
  const changes = plan.totals.added + plan.totals.marked + plan.totals.removed;
  const RES_LABEL = { keep: "Keep existing", override: "Override", skip: "No change" };

  const toggleStatus = (code) => {
    if (statusCodes.includes(code)) setStatusText(statuses.filter((x) => xToStatusCode(x) !== code).join(", "));
    else setStatusText([...statuses, code].join(", "));
  };
  const setField = (id, patch) => setFields((fs) => fs.map((f) => (f.id === id ? { ...f, ...patch } : f)));
  const togglePicked = (uidVal, on) => setPicked((p) => (on ? [...p, uidVal] : p.filter((x) => x !== uidVal)));

  const apply = () => {
    const t = plan.totals;
    if (t.unresolved > 0) {
      window.alert(`${t.unresolved} field${t.unresolved === 1 ? "" : "s"} already exist${t.unresolved === 1 ? "s" : ""} with the same label but a different key.\n\nChoose Keep existing, Override or No change (in the yellow box) before applying.`);
      return;
    }
    if (!changes) return;
    const parts = mode === "remove"
      ? [`stop requiring ${t.removed} field${t.removed === 1 ? "" : "s"}`]
      : [t.added ? `add ${t.added} missing field${t.added === 1 ? "" : "s"}` : "", t.marked ? `make ${t.marked} field${t.marked === 1 ? "" : "s"} mandatory` : "", t.overridden ? `override ${t.overridden} existing same-label field${t.overridden === 1 ? "" : "s"} (new key)` : ""].filter(Boolean);
    if (!window.confirm(`On the ${statuses.join(", ")} status${statuses.length === 1 ? "" : "es"} of ${t.subs} sub-type${t.subs === 1 ? "" : "s"}: ${parts.join(" and ")}.\n\nContinue?`)) return;
    onApply(plan.library, plan.totals, mode);
  };

  return (
    <div className="space-y-5">
      <div className="rounded-xl border border-white/10 bg-[#0D1220] p-4 text-sm text-white">
        <p className="font-medium mb-1">Make fields mandatory when a ticket moves to a status — for many sub-types at once</p>
        <p className="text-xs text-white/70">
          1) pick the ticket types, 2) pick the statuses (e.g. resolved, closed), 3) say which fields must be filled. A field that isn't on the status yet is
          added to it, then listed in its mandatory fields. Image uploads always use the key <span className="font-mono">{STATUS_FILE_UPLOAD_KEY}</span>.
        </p>
      </div>

      {/* 1. ticket types */}
      <div>
        <Label hint="Which ticket types should be changed">1 · Ticket types</Label>
        <div className="rounded-xl border border-white/10 bg-[#111528] p-3">
          <label className="flex cursor-pointer items-center gap-2 text-sm text-white">
            <Checkbox checked={allTypes} onChange={setAllTypes} title="Apply to every ticket type" />All ticket types ({library.length})
          </label>
          {!allTypes && (
            <div className="mt-2 grid grid-cols-1 gap-x-4 gap-y-1.5 md:grid-cols-2 max-h-56 overflow-y-auto">
              {library.length === 0 && <p className="text-xs text-white/70">No ticket types loaded yet.</p>}
              {library.map((d) => (
                <label key={d._uid} className="flex cursor-pointer items-center gap-2 text-sm text-white">
                  <Checkbox checked={picked.includes(d._uid)} onChange={(on) => togglePicked(d._uid, on)} title="Include this ticket type" />
                  <span className="truncate">{d.ticketType || d.ticketKey || "(untitled)"}</span>
                  <span className="text-xs text-white/60">{d.ticketSubType.length} sub</span>
                </label>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* 2. statuses */}
      <div>
        <Label hint="Status code or label, comma separated">2 · Statuses</Label>
        <Input mono value={statusText} onChange={setStatusText} placeholder="resolved, closed" />
        {found.length > 0 && (
          <div className="mt-2 flex flex-wrap items-center gap-1.5">
            <span className="text-xs text-white/60">In these ticket types:</span>
            {found.map((st) => {
              const on = statusCodes.includes(st.code);
              return (
                <button
                  key={st.code} type="button" onClick={() => toggleStatus(st.code)}
                  className={`rounded-full border px-2.5 py-0.5 text-xs ${on ? "border-indigo-400 bg-indigo-500/20 text-indigo-100" : "border-white/10 text-white/80 hover:border-white/30"}`}
                  title={`${st.count} sub-type${st.count === 1 ? "" : "s"} have this status`}
                >{st.label || st.code} <span className="text-white/50">· {st.count}</span></button>
              );
            })}
          </div>
        )}
      </div>

      {/* 3. fields */}
      <div>
        <div className="flex items-end justify-between gap-3">
          <Label hint="What must be filled when moving to those statuses">3 · Mandatory fields</Label>
          <div className="mb-1 flex gap-2">
            <Btn variant="ghost" small onClick={() => setFields((fs) => [...fs, presetResolutionImages()])}><Plus size={13} />Resolution Images</Btn>
            <Btn variant="ghost" small onClick={() => setFields((fs) => [...fs, presetActionTaken()])}><Plus size={13} />Action Taken</Btn>
            <Btn variant="dashed" small onClick={() => setFields((fs) => [...fs, { id: uid(), type: "text", label: "", key: "" }])}><Plus size={13} />Other field</Btn>
          </div>
        </div>
        <div className="overflow-x-auto rounded-xl border border-white/10 bg-[#111528]">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-[11px] uppercase tracking-wider text-white">
                <th className="px-3 py-2.5 min-w-[170px]">Type</th>
                <th className="px-3 py-2.5 min-w-[200px]">Label</th>
                <th className="px-3 py-2.5 min-w-[170px]">Key</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {fields.length === 0 && <tr><td colSpan={4} className="px-3 py-6 text-center text-xs text-white/70">Add at least one field.</td></tr>}
              {fields.map((f) => (
                <tr key={f.id} className="border-t border-white/5">
                  <td className="px-3 py-2">
                    <select value={f.type} onChange={(e) => setField(f.id, { type: e.target.value })} className={cellInput}>
                      {MANDATORY_FIELD_TYPES.map((t) => <option key={t.value} value={t.value}>{t.label}</option>)}
                    </select>
                  </td>
                  <td className="px-3 py-2"><input value={f.label} onChange={(e) => setField(f.id, { label: e.target.value })} placeholder="e.g. Action Taken" className={cellInput} /></td>
                  <td className="px-3 py-2">
                    {f.type === "fileUpload"
                      ? <input value={STATUS_FILE_UPLOAD_KEY} disabled className={`${cellInput} font-mono opacity-60`} title="Fixed key for image uploads" />
                      : <input value={f.key} onChange={(e) => setField(f.id, { key: e.target.value })} placeholder={labelToKey(f.label) || "auto from label"} className={`${cellInput} font-mono`} />}
                  </td>
                  <td className="px-2 py-2"><IconBtn danger title="Remove field" onClick={() => setFields((fs) => fs.filter((x) => x.id !== f.id))}><Trash2 size={14} /></IconBtn></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {mode === "add" && plan.conflicts.length > 0 && (
        <div className="rounded-xl border border-amber-500/30 bg-amber-500/10 p-4">
          <div className="flex items-start gap-2">
            <AlertTriangle size={16} className="mt-0.5 shrink-0 text-amber-400" />
            <div className="min-w-0 flex-1">
              <p className="text-sm font-medium text-amber-100">
                {plan.conflicts.length} field{plan.conflicts.length === 1 ? "" : "s"} already exist with the same label but a different key
                {plan.totals.unresolved > 0 && <span className="text-rose-300"> — {plan.totals.unresolved} still need a decision</span>}
              </p>
              <p className="mt-0.5 text-xs text-amber-200/80">
                Keep existing = leave that field as it is and just make its current key mandatory · Override = change its key/type to the new one · No change = skip it.
              </p>
              <div className="mt-3 flex flex-wrap items-center gap-2">
                <span className="text-xs text-amber-100">For all of them:</span>
                <select value={defaultRes} onChange={(e) => setDefaultRes(e.target.value)} className={`${cellInput} !w-auto`}>
                  <option value="">— choose —</option>
                  <option value="keep">Keep existing</option>
                  <option value="override">Override with the new key</option>
                  <option value="skip">No change</option>
                </select>
                {Object.keys(resolutions).length > 0 && <button type="button" onClick={() => setResolutions({})} className="text-xs text-amber-200 underline hover:text-amber-100">reset individual choices</button>}
              </div>
              <div className="mt-3 max-h-64 overflow-auto rounded-lg border border-white/10 bg-[#0B0F1C]">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="text-left uppercase tracking-wider text-white/70">
                      <th className="px-2 py-1.5">Where</th>
                      <th className="px-2 py-1.5">Label</th>
                      <th className="px-2 py-1.5">Existing key</th>
                      <th className="px-2 py-1.5">New key</th>
                      <th className="px-2 py-1.5">Decision</th>
                    </tr>
                  </thead>
                  <tbody>
                    {plan.conflicts.map((c) => (
                      <tr key={c.id} className="border-t border-white/5">
                        <td className="px-2 py-1.5 text-white">{c.typeName} › {c.subName} › {c.statusLabel}</td>
                        <td className="px-2 py-1.5 text-white">{c.fieldLabel}</td>
                        <td className="px-2 py-1.5 font-mono text-white">{c.existingKey}<span className="text-white/50"> ({c.existingType})</span></td>
                        <td className="px-2 py-1.5 font-mono text-white">{c.newKey}<span className="text-white/50"> ({c.newType})</span></td>
                        <td className="px-2 py-1.5">
                          <select
                            value={resolutions[c.id] || ""}
                            onChange={(e) => setResolutions((r) => { const n = { ...r }; if (e.target.value) n[c.id] = e.target.value; else delete n[c.id]; return n; })}
                            className={`${cellInput} !py-1 ${c.res ? "" : "ring-1 ring-rose-500/60"}`}
                          >
                            <option value="">{defaultRes ? `Default (${RES_LABEL[defaultRes]})` : "Choose…"}</option>
                            <option value="keep">Keep existing</option>
                            <option value="override">Override</option>
                            <option value="skip">No change</option>
                          </select>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
        <div>
          <Label>Action</Label>
          <select value={mode} onChange={(e) => setMode(e.target.value)} className={cellInput}>
            <option value="add">Make these fields mandatory (add them if missing)</option>
            <option value="remove">Stop requiring them (the fields stay on the status)</option>
          </select>
        </div>
      </div>

      {/* preview */}
      <div className="overflow-x-auto rounded-xl border border-white/10 bg-[#111528]">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-[11px] uppercase tracking-wider text-white">
              <th className="px-3 py-2.5">Ticket type</th>
              <th className="px-3 py-2.5">Sub-types hit</th>
              <th className="px-3 py-2.5">Result</th>
            </tr>
          </thead>
          <tbody>
            {!plan.ready && <tr><td colSpan={3} className="px-3 py-6 text-center text-xs text-white/70">Pick the statuses and fields above to see what will change.</td></tr>}
            {plan.ready && plan.reps.length === 0 && <tr><td colSpan={3} className="px-3 py-6 text-center text-xs text-white/70">No ticket types selected.</td></tr>}
            {plan.reps.map((r) => (
              <tr key={r.uid} className="border-t border-white/5 align-top">
                <td className="px-3 py-2 text-white">{r.name}</td>
                <td className="px-3 py-2 text-white">{r.subs}</td>
                <td className="px-3 py-2 text-xs leading-5">
                  {mode === "remove"
                    ? <span className="text-rose-300">−{r.removed} no longer mandatory</span>
                    : <><span className="text-emerald-300">+{r.added} field{r.added === 1 ? "" : "s"} added</span><span className="text-sky-300"> · {r.marked} marked mandatory</span>{r.overridden > 0 && <span className="text-amber-300"> · {r.overridden} overridden</span>}{r.conflicts > 0 && <span className="text-amber-300"> · {r.conflicts} same-label conflict{r.conflicts === 1 ? "" : "s"}</span>}</>}
                  {r.missing.length > 0 && <div className="text-amber-300" title={r.missing.join("\n")}>{r.missing.length} sub-type{r.missing.length === 1 ? "" : "s"} have none of these statuses</div>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="flex flex-wrap items-center gap-3">
        <span className="text-sm text-white">
          {changes
            ? (mode === "remove" ? `${plan.totals.removed} to un-require` : `${plan.totals.added} fields to add · ${plan.totals.marked} to mark mandatory${plan.totals.overridden ? ` · ${plan.totals.overridden} to override` : ""}`) + `, across ${plan.totals.subs} sub-types`
            : "Nothing to change yet"}
        </span>
        <div className="ml-auto"><Btn variant="accent" onClick={apply}><Check size={15} />Apply to selected sub-types</Btn></div>
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
      // keep the document's own extras (its Mongo _id above all), or a re-import would silently drop them
      _passthrough: { ...(incoming._passthrough || {}), ...(current._passthrough || {}) },
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

/* ------------------------------------------------------------------ */
/* Excel re-upload = MERGE, never replace                              */
/* The workbook only carries part of a ticket type (no notification     */
/* settings, dependent statuses, status-level remarks / upload / radio  */
/* fields, field keys, option keys, dependentOn, ids ...). So when a    */
/* sub-type already exists, only what the sheets actually say is        */
/* applied and everything else is kept exactly as it was.              */
/* ------------------------------------------------------------------ */
const EXCEL_HOLDS_TYPES = ["dropdown", "radio", "multiplechoice", "date", "time", "number", "text", "fileUpload", "matrix", "casecadeDropdown", "apiDropdown"];
const xSame = (a, b) => xnorm(a).toLowerCase() === xnorm(b).toLowerCase();
const xSameSet = (a, b) => {
  const f = (l) => (l || []).map((v) => xnorm(v).toLowerCase()).sort().join("|");
  return f(a) === f(b);
};

/* options: the sheet shows values only, so an option that still exists keeps its key and dependentCustomFields */
function mergeExcelOptions(prevOpts, incOpts) {
  return (incOpts || []).map((o) => {
    const prev = (prevOpts || []).find((p) => xSame(p.value, o.value));
    return prev ? { ...prev, value: o.value } : o;
  });
}

/* one sub-type-level field: the sheet owns label / type / mandatory / edit roles / options;
   key, disabled, displayLabel, dependentOn, hideTotalColumn, formulas ... stay as they were */
function mergeExcelField(cur, inc) {
  if (cur.type !== inc.type && !(OPTION_BASED_TYPES.includes(cur.type) && OPTION_BASED_TYPES.includes(inc.type))) {
    return { ...inc, id: cur.id, key: cur.key || inc.key, disabled: cur.disabled, displayLabel: cur.displayLabel, dependentOn: cur.dependentOn };
  }
  const out = { ...cur, label: inc.label, type: inc.type, isMandatory: inc.isMandatory, editRoles: inc.editRoles };
  if (OPTION_BASED_TYPES.includes(inc.type)) out.options = mergeExcelOptions(cur.options, inc.options);
  else if (inc.type === "casecadeDropdown") Object.assign(out, { childLabel: inc.childLabel, childKey: inc.childKey, parentOptionType: inc.parentOptionType, options: inc.options });
  else if (inc.type === "apiDropdown") Object.assign(out, { labelKey: inc.labelKey, valueKey: inc.valueKey, fillOnSelect: inc.fillOnSelect, apiFiltersText: inc.apiFiltersText });
  else if (inc.type === "matrix") {
    out.columns = (inc.columns || []).map((c) => {
      const prev = (cur.columns || []).find((p) => p.key === c.key) || (cur.columns || []).find((p) => xSame(p.label, c.label));
      return prev ? { ...prev, label: c.label, type: c.type, options: OPTION_BASED_TYPES.includes(c.type) ? mergeExcelOptions(prev.options, c.options) : prev.options } : c;
    });
  }
  return out;
}

function mergeExcelFieldList(curFields, incFields) {
  const used = new Set();
  const out = (incFields || []).map((inc) => {
    const free = (curFields || []).filter((c) => !used.has(c));
    const prev = free.find((c) => xSame(c.label, inc.label) && c.type === inc.type) || free.find((c) => xSame(c.label, inc.label)) || free.find((c) => c.key && c.key === inc.key);
    if (!prev) return inc;
    used.add(prev);
    return mergeExcelField(prev, inc);
  });
  let kept = 0;
  // field types the workbook can't hold were never written to it, so they must survive a re-upload
  (curFields || []).forEach((c, i) => {
    if (used.has(c) || EXCEL_HOLDS_TYPES.includes(c.type)) return;
    out.splice(Math.min(i, out.length), 0, c);
    kept++;
  });
  return { fields: out, kept };
}

/* what the Mandate rules sheet shows for a status */
function mandateRuleOf(st) {
  const mand = st.mandatoryCustomFields || [];
  const field = (st.customFieldsMetaData || []).find((f) => mand.includes(f.key));
  return { images: mand.includes("images"), newAssignee: mand.includes("newAssignee"), commentOn: !!field || mand.includes("comments"), field, roles: st.assigneeRoles || [] };
}

/* apply only the DIFFERENCE between the sheet and what the app would have written */
function applyExcelMandate(cur, inc) {
  const a = mandateRuleOf(cur), b = mandateRuleOf(inc);
  let mand = [...(cur.mandatoryCustomFields || [])];
  let fields = [...(cur.customFieldsMetaData || [])];
  const setFlag = (key, on) => { mand = mand.filter((k) => k !== key); if (on) mand.push(key); };
  if (a.images !== b.images) setFlag("images", b.images);
  if (a.newAssignee !== b.newAssignee) setFlag("newAssignee", b.newAssignee);

  const aLabel = a.field ? a.field.label : "Comments";
  const aType = a.field ? (a.field.type || "text") : "text";
  const bField = b.field;
  const bLabel = bField ? bField.label : "Comments";
  const bType = bField ? (bField.type || "text") : "text";
  const simpleType = ["text", "number", "date", "fileUpload"].includes(aType); // the sheet can only express these
  const typeChanged = simpleType && aType !== bType;
  if (a.commentOn !== b.commentOn || (b.commentOn && (!xSame(aLabel, bLabel) || typeChanged))) {
    if (!b.commentOn) {
      if (a.field) mand = mand.filter((k) => k !== a.field.key);
      mand = mand.filter((k) => k !== "comments");     // the field itself stays, it just stops being required
    } else if (a.field) {
      fields = fields.map((f) => (f === a.field ? { ...f, label: bLabel, ...(typeChanged ? { type: bType } : {}) } : f)); // rename in place, keep its key
    } else if (bField) {
      if (!fields.some((f) => f.key === bField.key)) fields.push(bField);
      if (!mand.includes(bField.key)) mand.push(bField.key);
    }
  }
  const out = { ...cur, mandatoryCustomFields: mand, customFieldsMetaData: fields };
  if (!xSameSet(cur.assigneeRoles, inc.assigneeRoles)) out.assigneeRoles = inc.assigneeRoles || [];
  return out;
}

function mergeExcelStatus(cur, inc, seen) {
  const code = toStatusCode(cur.status);
  let out = { ...cur };                                  // notification settings, dependentStatus, remarks/upload/radio fields ... all kept
  if (seen.roles && seen.roles.has(code)) out = { ...out, label: inc.label, roles: inc.roles, public: inc.public };
  if (seen.mandate && seen.mandate.has(code)) out = applyExcelMandate(out, inc);
  if (seen.esc) {
    if (!seen.esc.has(code)) out = { ...out, autoEscalationConfig: [] };            // the sheet lists every level, so a status with none left has none
    else {
      const used = new Set();
      out.autoEscalationConfig = (inc.autoEscalationConfig || []).map((lv) => {
        const prev = (cur.autoEscalationConfig || []).find((p) => !used.has(p) && xSame(p.assignee, lv.assignee));
        if (!prev) return lv;
        used.add(prev);                                                              // keeps the level's _id
        return { ...prev, assignee: lv.assignee, days: lv.days, notification: lv.notification, notificationType: lv.notificationType, ccAddress: lv.ccAddress, overrideStatus: lv.overrideStatus };
      });
    }
  }
  return out;
}

function mergeExcelSubType(cur, inc, seen) {
  const notes = [];
  const curStatuses = cur.statusWorkFlow || [];
  const byCode = new Map(curStatuses.map((s) => [toStatusCode(s.status), s]));
  const used = new Set();
  const ambiguous = [];
  const statuses = (inc.statusWorkFlow || []).map((is) => {
    const code = toStatusCode(is.status);
    const c = byCode.get(code);
    if (!c) {
      // a workbook exported before the "Status Code" column existed only has labels: a row whose label
      // belongs to an existing status (e.g. several "Approve" / "Reject") can't be told apart, so it is skipped
      // rather than added as a made-up extra status
      if (curStatuses.some((x) => xSame(x.label, is.label))) { ambiguous.push(is.label); return null; }
      return is;
    }
    if (used.has(code)) return null;
    used.add(code);
    return mergeExcelStatus(c, is, seen);
  }).filter(Boolean);
  if (ambiguous.length) notes.push(`${ambiguous.length} row(s) share a label with existing statuses (${[...new Set(ambiguous)].join(", ")}) and have no Status Code, so they were ignored. Export a fresh workbook to get the Status Code column.`);
  const left = curStatuses.filter((s) => !used.has(toStatusCode(s.status)));
  curStatuses.forEach((s, i) => { if (left.includes(s)) statuses.splice(Math.min(i, statuses.length), 0, s); });
  if (left.length) notes.push(`${left.length} status(es) in the app aren't in the sheet — kept as they were (${left.map((s) => s.label || s.status).join(", ")})`);
  const { fields, kept } = mergeExcelFieldList(cur.customFieldsMetaData, inc.customFieldsMetaData);
  if (kept) notes.push(`${kept} field(s) of a type Excel can't hold were kept`);
  return {
    sub: {
      ...cur,
      ticketType: inc.ticketType || cur.ticketType,
      prefix: seen.prefixDefaulted ? (cur.prefix || inc.prefix) : (inc.prefix || cur.prefix),
      assigneeRoles: !seen.assigneeDefaulted && (inc.assigneeRoles || []).length ? inc.assigneeRoles : (cur.assigneeRoles && cur.assigneeRoles.length ? cur.assigneeRoles : (inc.assigneeRoles || [])),
      deleted: inc.deleted,
      isAutoEscalation: inc.isAutoEscalation,
      customFieldsMetaData: fields,
      statusWorkFlow: statuses,
    },
    notes,
  };
}

function mergeExcelDoc(current, incoming) {
  const strip = (s) => { const { _excelSeen, ...rest } = s; return rest; };
  if (isDocEmpty(current)) {
    return { doc: { ...incoming, ticketSubType: incoming.ticketSubType.map(strip) }, added: incoming.ticketSubType.map((s) => s.ticketType || s.ticketKey), updated: [], notes: [] };
  }
  const list = [...current.ticketSubType];
  const added = [], updated = [], notes = [];
  incoming.ticketSubType.forEach((inc) => {
    const key = xnorm(inc.ticketKey).toLowerCase();
    const idx = key ? list.findIndex((s) => xnorm(s.ticketKey).toLowerCase() === key) : -1;
    const label = inc.ticketType || inc.ticketKey || "(unnamed)";
    if (idx >= 0) {
      const r = mergeExcelSubType(list[idx], strip(inc), inc._excelSeen || {});
      list[idx] = r.sub;
      updated.push(label);
      r.notes.forEach((n) => notes.push(`${current.ticketType || current.ticketKey} › ${label}: ${n}`));
    } else {
      list.push(strip(inc));
      added.push(label);
    }
  });
  return {
    doc: {
      ...current,
      ticketKey: current.ticketKey || incoming.ticketKey,
      ticketType: current.ticketType || incoming.ticketType,
      tenantId: current.tenantId || incoming.tenantId,
      creators: current.creators.length ? current.creators : incoming.creators,
      viewers: current.viewers.length ? current.viewers : incoming.viewers,
      assignee: current.assignee.length ? current.assignee : incoming.assignee,
      ticketSubType: list,
    },
    added, updated, notes,
  };
}

/* Upsert a whole batch of incoming ticket types into the library, matched
   by ticketKey (case-insensitive). New ticket types are appended; existing
   ones get their sub-types merged via mergeIncomingDoc (never wiped). */
function importDocsIntoLibrary(library, incomingDocs, opts = {}) {
  let nextLibrary = [...library];
  const newTypes = [];
  const updatedTypes = [];
  const mergeNotes = [];
  incomingDocs.forEach((inc) => {
    const incKey = xnorm(inc.ticketKey).toLowerCase();
    const idx = incKey ? nextLibrary.findIndex((d) => xnorm(d.ticketKey).toLowerCase() === incKey) : -1;
    if (idx >= 0) {
      const r = opts.excel ? mergeExcelDoc(nextLibrary[idx], inc) : mergeIncomingDoc(nextLibrary[idx], inc);
      const { doc: merged, added, updated } = r;
      if (r.notes) mergeNotes.push(...r.notes);
      nextLibrary[idx] = { ...merged, _uid: nextLibrary[idx]._uid };
      const label = inc.ticketType || inc.ticketKey || "(unnamed)";
      if (added.length || updated.length) {
        updatedTypes.push(`${label} (+${added.length} sub-type${added.length === 1 ? "" : "s"} added, ${updated.length} updated)`);
      }
    } else {
      const fresh = opts.excel ? { ...inc, ticketSubType: inc.ticketSubType.map(({ _excelSeen, ...rest }) => rest), _passthrough: { _id: newOid(), ...(inc._passthrough || {}) } } : inc;
      nextLibrary.push({ ...fresh, _uid: uid() });
      newTypes.push(inc.ticketType || inc.ticketKey || "(unnamed)");
    }
  });
  return { library: nextLibrary, newTypes, updatedTypes, mergeNotes };
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
  const [toasts, setToasts] = useState([]);
  const [clip, setClip] = useState(null); // { subs, from } — sub-types copied, ready to paste into any ticket type
  const [subPick, setSubPick] = useState({ uid: null, ids: [] }); // sub-types ticked on the Sub-Types tab, per ticket type
  const fileInputRef = React.useRef(null);
  const excelInputRef = React.useRef(null);

  const selectedIndex = library.findIndex((d) => d._uid === selectedUid);
  const selected = selectedIndex >= 0 ? library[selectedIndex] : null;

  const dismissToast = (id) => setToasts((prev) => prev.filter((t) => t.id !== id));
  const showToast = ({ type = "info", title, lines = [], duration, action }) => {
    const id = uid();
    setToasts((prev) => [...prev.slice(-3), { id, type, title, lines, action }]);
    setTimeout(() => dismissToast(id), duration || (type === "error" ? (action ? 20000 : 9000) : 4500));
  };

  // Add the escalation assignee roles to both the sub-type's and the ticket type's assignee roles.
  const fixAssigneeRoles = (uids) => {
    setLibrary((lib) => lib.map((d) => (uids.includes(d._uid) ? syncAssigneeRoles(d) : d)));
    showToast({ type: "success", title: "Assignee roles updated", lines: ["The escalation assignee roles were added to the sub-type and ticket type. Export again."] });
  };

  // Required before any export: creators + viewers on every ticket type, and the assignee-role rules.
  const checkDocsForExport = (docs, { goToRoles = false } = {}) => {
    const lines = [];
    const rolesFirst = [];
    const fixableUids = [];
    docs.forEach((d) => {
      const label = d.ticketType || d.ticketKey || "(untitled ticket type)";
      const missing = missingRequiredRoles(d);
      if (missing.length) { lines.push(`${label}: ${missing.join(" and ")} ${missing.length === 1 ? "is" : "are"} not set`); rolesFirst.push(d); }
      assigneeProblems(d).forEach((p) => {
        lines.push(`${label} › ${p.text}`);
        if (p.fixable && !fixableUids.includes(d._uid)) fixableUids.push(d._uid);
      });
    });
    if (!lines.length) return true;
    showToast({
      type: "error",
      title: "Export blocked — fix these first",
      lines,
      action: fixableUids.length ? { label: "Add the missing assignee roles", onClick: () => fixAssigneeRoles(fixableUids) } : undefined,
    });
    if (goToRoles && docs.length === 1 && rolesFirst.length === 1) { setSelectedUid(rolesFirst[0]._uid); setTab("roles"); }
    return false;
  };

  // Ticks only count for the ticket type they were made on (and only for sub-types that still exist).
  const pickedSubIds = new Set(
    subPick.uid === selectedUid && selected ? subPick.ids.filter((id) => selected.ticketSubType.some((s) => s.id === id)) : [],
  );
  const setPickedSubIds = (ids) => setSubPick({ uid: selectedUid, ids: [...ids] });

  // Soft heads-up (amber, auto-dismisses) after prefixes were rewritten to ${locationId}.
  const notifyPrefixFixes = (fixes) => {
    if (!fixes.length) return;
    showToast({
      type: "warning",
      title: `Prefix corrected in ${fixes.length} sub-type${fixes.length === 1 ? "" : "s"} — please review`,
      lines: [
        `The location placeholder must be written exactly as ${PREFIX_PLACEHOLDER}.`,
        ...fixes.slice(0, 5).map((f) => `${f.ticket} › ${f.sub}: "${f.from}" → "${f.to}"`),
        ...(fixes.length > 5 ? [`…and ${fixes.length - 5} more`] : []),
      ],
      duration: 9000,
    });
  };

  const updateSelected = (patch) => {
    if (selectedIndex < 0) return;
    const { doc: fixedDoc, fixes } = normalizeDocPrefixes(patch);
    setLibrary(library.map((d, i) => (i === selectedIndex ? { ...fixedDoc, _uid: d._uid } : d)));
    notifyPrefixFixes(fixes);
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
        const prefixFixes = [];
        const incomingDocs = rawDocs.map(jsonToDoc).map((d) => { const r = normalizeDocPrefixes(d); prefixFixes.push(...r.fixes); return r.doc; });
        const { library: nextLibrary, newTypes, updatedTypes } = importDocsIntoLibrary(library, incomingDocs);
        setLibrary(nextLibrary);
        notifyPrefixFixes(prefixFixes);
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
        const { docs: rawParsedDocs, warnings, missingSheets } = excelWorkbookToDocs(wb);
        const prefixFixes = [];
        const parsedDocs = rawParsedDocs.map((d) => { const r = normalizeDocPrefixes(d); prefixFixes.push(...r.fixes); return r.doc; });
        const { library: nextLibrary, newTypes, updatedTypes, mergeNotes } = importDocsIntoLibrary(library, parsedDocs, { excel: true });
        setLibrary(nextLibrary);
        notifyPrefixFixes(prefixFixes);
        const notes = [...warnings];
        if (updatedTypes.length) notes.push("Existing ticket types were merged, not replaced: anything the workbook doesn't carry (notification settings, dependent statuses, status remarks / upload / radio fields, field and option keys, ids) was kept as it was.");
        notes.push(...(mergeNotes || []));
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

  const downloadStarterSheet = () => {
    XLSX.writeFile(excelStarterWorkbook(), "ticket-types-starter.xlsx", { compression: true });
    showToast({ type: "info", title: "Starter sheet downloaded", lines: ["Type the ticket types, then one sub-type per row. The green columns (keys, prefixes, assignees, status flow for every sub-type) fill themselves in Excel — then import the saved file."] });
  };

  const exportSelectedToExcel = () => {
    if (!selected) { window.alert("Pick a ticket type first (from Ticket Types) before exporting."); return; }
    if (!checkDocsForExport([selected], { goToRoles: true })) return;
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
    if (!checkDocsForExport(library)) return;
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
    if (!checkDocsForExport(library)) return;
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
    copy._passthrough = { _id: newOid() };
    // ...and every escalation level of the copy gets its own uuid
    copy.ticketSubType = copy.ticketSubType.map((s) => ({
      ...s,
      _passthrough: {},
      statusWorkFlow: s.statusWorkFlow.map((w) => ({ ...w, autoEscalationConfig: w.autoEscalationConfig.map((e) => ({ ...e, _id: uuidv4() })) })),
    }));
    setLibrary([...library, copy]);
    setSelectedUid(copy._uid);
    setTab("details");
  };

  /* Copy sub-types to the in-app clipboard; paste them into any ticket type
     (or the same one). Unlike Duplicate, this works across ticket types. */
  const copySubs = (subs, from) => {
    if (!subs.length) { showToast({ type: "info", title: "Nothing to copy", lines: ["That ticket type has no sub-types yet."] }); return; }
    setClip({ subs: subs.map((s) => JSON.parse(JSON.stringify(s))), from: from || "" });
    showToast({ type: "success", title: `Copied ${subs.length} sub-type${subs.length === 1 ? "" : "s"}`, lines: ["Open another ticket type and press Paste (Sub-Types tab, or the paste icon in Ticket Types)."] });
  };

  const pasteSubs = (docUid) => {
    const target = library.find((d) => d._uid === docUid);
    if (!clip || !target) return;
    const keys = new Set(target.ticketSubType.map((s) => xnorm(s.ticketKey).toLowerCase()));
    const renamed = [];
    const pasted = clip.subs.map((src) => {
      const fresh = regenerateIds(src); // new ids for the sub-type, fields, statuses, options and escalation levels
      fresh._passthrough = {}; // a pasted sub-type is a new record: no backend-only fields from the source
      let key = fresh.ticketKey || "";
      if (xnorm(key) && keys.has(xnorm(key).toLowerCase())) {
        const base = `${key}_copy`;
        let n = 2;
        key = base;
        while (keys.has(key.toLowerCase())) { key = `${base}${n}`; n += 1; }
        fresh.ticketKey = key;
        if (fresh.ticketType) fresh.ticketType = `${fresh.ticketType} (copy)`;
        renamed.push(fresh.ticketType || key);
      }
      keys.add(xnorm(key).toLowerCase());
      return fresh;
    });
    const { doc: fixedDoc, fixes } = normalizeDocPrefixes({ ...target, ticketSubType: [...target.ticketSubType, ...pasted] });
    setLibrary(library.map((d) => (d._uid === docUid ? { ...fixedDoc, _uid: d._uid } : d)));
    notifyPrefixFixes(fixes);
    setSelectedUid(docUid);
    setFocusSubtypeId(null);
    setTab("subtypes");
    showToast({
      type: "success",
      title: `Pasted ${pasted.length} sub-type${pasted.length === 1 ? "" : "s"} into "${target.ticketType || target.ticketKey || "(untitled)"}"`,
      lines: renamed.length ? [`Key already existed there, so renamed: ${renamed.join(", ")}`] : [],
    });
  };

  // Wipe everything (what Ctrl+R used to be for): ticket types, templates, clipboard and selections.
  const clearAll = () => {
    if (!library.length && !templates.length) {
      showToast({ type: "info", title: "Nothing to clear", lines: ["The workspace is already empty."] });
      return;
    }
    const tCount = library.length;
    const subCount = library.reduce((n, d) => n + d.ticketSubType.length, 0);
    const parts = [`${tCount} ticket type${tCount === 1 ? "" : "s"} (${subCount} sub-type${subCount === 1 ? "" : "s"})`];
    if (templates.length) parts.push(`${templates.length} template${templates.length === 1 ? "" : "s"}`);
    if (!window.confirm(`Clear ALL data — ${parts.join(" and ")}?\n\nThis wipes the whole workspace and can't be undone. Export first if you need a copy.`)) return;
    setLibrary([]);
    setTemplates([]);
    setClip(null);
    setSelectedUid(null);
    setFocusSubtypeId(null);
    setSubPick({ uid: null, ids: [] });
    setImportNotice(null);
    setBulkOpen(false);
    setTab("library");
    showToast({ type: "success", title: "Workspace cleared", lines: ["All ticket types and templates were removed."] });
  };

  const deleteType = (uidVal) => {
    const target = library.find((d) => d._uid === uidVal);
    if (!target) return;
    if (!window.confirm(`Delete ticket type "${target.ticketType || target.ticketKey || "(unnamed)"}" and all its sub-types? This can't be undone.`)) return;
    setLibrary(library.filter((d) => d._uid !== uidVal));
    if (selectedUid === uidVal) { setSelectedUid(null); setTab("library"); }
  };

  /* Delete several ticket types and/or sub-types in one go (one confirmation).
     Returns true when something was deleted so the caller can clear its ticks. */
  const bulkDelete = ({ typeUids = [], subs = [] }) => {
    const typeSet = new Set(typeUids);
    const subMap = {};
    subs.forEach(({ uid: u, subId }) => { if (!typeSet.has(u)) (subMap[u] = subMap[u] || new Set()).add(subId); });
    const subCount = Object.values(subMap).reduce((n, s) => n + s.size, 0);
    const typeSubCount = library.filter((d) => typeSet.has(d._uid)).reduce((n, d) => n + d.ticketSubType.length, 0);
    if (!typeSet.size && !subCount) return false;
    const parts = [];
    if (typeSet.size) parts.push(`${typeSet.size} ticket type${typeSet.size === 1 ? "" : "s"} (with ${typeSubCount} sub-type${typeSubCount === 1 ? "" : "s"})`);
    if (subCount) parts.push(`${subCount} sub-type${subCount === 1 ? "" : "s"}`);
    if (!window.confirm(`Delete ${parts.join(" and ")}? This can't be undone.`)) return false;
    setLibrary(library
      .filter((d) => !typeSet.has(d._uid))
      .map((d) => (subMap[d._uid] ? { ...d, ticketSubType: d.ticketSubType.filter((s) => !subMap[d._uid].has(s.id)) } : d)));
    if (typeSet.has(selectedUid)) { setSelectedUid(null); setTab("library"); }
    showToast({ type: "success", title: `Deleted ${parts.join(" and ")}` });
    return true;
  };

  const deletePickedSubtypes = () => {
    if (!selected || !pickedSubIds.size) return;
    if (bulkDelete({ subs: [...pickedSubIds].map((subId) => ({ uid: selected._uid, subId })) })) setPickedSubIds([]);
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

  const applyBulkEscalation = (nextLibrary, totals) => {
    setLibrary(nextLibrary);
    showToast({
      type: "success",
      title: "Escalation levels applied",
      lines: [`${totals.added} added, ${totals.updated} days updated${totals.removed ? `, ${totals.removed} replaced` : ""} across ${totals.subs} sub-types.${totals.removed && totals.ccLevels ? (totals.ccKept ? " Old CC emails were kept on the new levels." : " Old CC emails were removed.") : ""} Assignee roles were added where needed.`],
    });
  };

  const applyBulkMandatory = (nextLibrary, totals, mode) => {
    setLibrary(nextLibrary);
    showToast({
      type: "success",
      title: mode === "remove" ? "Mandatory fields removed" : "Mandatory fields applied",
      lines: [mode === "remove"
        ? `${totals.removed} field${totals.removed === 1 ? "" : "s"} no longer mandatory across ${totals.subs} sub-types.`
        : `${totals.added} field${totals.added === 1 ? "" : "s"} added, ${totals.marked} marked mandatory${totals.overridden ? `, ${totals.overridden} overridden` : ""} across ${totals.subs} sub-types.`],
    });
  };

  const nav = [
    { id: "library", label: `Ticket Types (${library.length})`, icon: <ListTree size={16} /> },
    { id: "templates", label: `Templates (${templates.length})`, icon: <LayoutTemplate size={16} /> },
    { id: "smartupdate", label: "Smart Update", icon: <Wand2 size={16} /> },
    { id: "bulkesc", label: "Bulk Escalation", icon: <AlertTriangle size={16} /> },
    { id: "bulkmand", label: "Bulk Mandatory", icon: <ShieldCheck size={16} /> },
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
    <div
      className="min-h-screen bg-[#070A13] bg-[radial-gradient(ellipse_90%_50%_at_50%_-10%,rgba(129,140,248,0.14),rgba(7,10,19,0)_60%)] flex text-white"
      style={{ fontFamily: "Inter, ui-sans-serif, system-ui" }}
    >
      {/* sidebar */}
      <aside className="w-60 shrink-0 bg-[#050710] text-white flex flex-col border-r border-white/5">
        <div className="px-5 py-5 border-b border-white/5">
          <div className="flex items-center gap-2 text-white font-semibold">
            <Workflow size={18} className="text-indigo-400" />
            Ticket Type Builder
          </div>
          <div className="text-[11px] text-white mt-0.5">Ticket type dataset</div>
        </div>
        <nav className="flex-1 px-2 py-4 space-y-1">
          {nav.map((n) => {
            const disabled = n.id !== "library" && n.id !== "templates" && n.id !== "smartupdate" && n.id !== "bulkesc" && n.id !== "bulkmand" && !selected;
            return (
              <button
                key={n.id}
                onClick={() => !disabled && setTab(n.id)}
                disabled={disabled}
                className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-xl text-sm font-medium transition ${tab === n.id ? "bg-indigo-400/15 text-indigo-200 shadow-[inset_0_0_0_1px_rgba(129,140,248,0.25)]" : disabled ? "text-slate-600 cursor-not-allowed" : "hover:bg-white/5 text-white"}`}
              >
                {n.icon}
                {n.label}
              </button>
            );
          })}
        </nav>
        <div className="px-3 py-4 border-t border-white/5 space-y-1.5">
          <button onClick={downloadStarterSheet} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-xs text-white hover:bg-white/5 hover:text-white transition" title="An Excel sheet with formulas: type ticket types + sub-type names once, keys / prefixes / assignees / status flow fill in by themselves">
            <FileSpreadsheet size={13} /> Download starter sheet
          </button>
          <button onClick={loadSample} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-xs text-white hover:bg-white/5 hover:text-white transition">
            <RefreshCw size={13} /> Add sample ticket type
          </button>
        </div>
      </aside>

      {/* main */}
      <main className="flex-1 min-w-0">
        <header className="bg-[#0A0E1C] border-b border-white/10 px-8 py-4 flex items-center justify-between gap-4">
          <div className="min-w-0">
            <h1 className="text-base font-semibold text-white truncate">
              {tab === "library" ? "All ticket types" : tab === "templates" ? "Sub-type templates" : tab === "smartupdate" ? "Smart update" : tab === "bulkesc" ? "Bulk escalation" : tab === "bulkmand" ? "Bulk mandatory fields" : (selected ? (selected.ticketType || "Untitled ticket type") : "No ticket type selected")}
            </h1>
            <p className="text-xs font-mono text-white">
              {tab === "library" ? `${library.length} loaded` : tab === "templates" ? `${templates.length} template${templates.length === 1 ? "" : "s"} available for bulk creation` : tab === "smartupdate" ? "describe changes, review the diff, then apply" : tab === "bulkesc" ? "bulk-add escalation levels" : tab === "bulkmand" ? "require fields when a status is reached" : (selected ? (selected.ticketKey || "no key set yet") : "pick one from Ticket Types")}
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
            <Btn variant="danger" onClick={clearAll}><Eraser size={15} />Clear all</Btn>
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

        <div className={(tab === "library" || tab === "templates" || tab === "smartupdate" || tab === "bulkesc" || tab === "bulkmand") ? "px-8 py-6" : "px-8 py-6 max-w-4xl"}>
          {tab === "library" && (
            <LibraryTab
              library={library}
              selectedUid={selectedUid}
              onSelectType={selectType}
              onSelectSubtype={selectSubtype}
              onNew={newTicketType}
              onDuplicate={duplicateType}
              onDelete={deleteType}
              onBulkDelete={bulkDelete}
              onExportAll={exportAll}
              clip={clip}
              onCopySubs={copySubs}
              onPasteSubs={pasteSubs}
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
          {tab === "bulkesc" && <BulkEscalationTab library={library} onApply={applyBulkEscalation} />}
          {tab === "bulkmand" && <BulkMandatoryTab library={library} onApply={applyBulkMandatory} />}
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
                <Btn variant="ghost" onClick={() => copySubs(selected.ticketSubType, selected.ticketType || selected.ticketKey)}>
                  <ClipboardCopy size={15} />Copy all
                </Btn>
                {clip && (
                  <>
                    <Btn variant="accent" onClick={() => pasteSubs(selected._uid)}>
                      <ClipboardPaste size={15} />Paste {clip.subs.length} copied sub-type{clip.subs.length === 1 ? "" : "s"}{clip.from ? ` (from ${clip.from})` : ""}
                    </Btn>
                    <IconBtn title="Clear the clipboard" onClick={() => setClip(null)}><X size={14} /></IconBtn>
                  </>
                )}
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

              {selected.ticketSubType.length > 0 && (
                <div className="mb-3 flex flex-wrap items-center gap-3 rounded-xl border border-white/10 bg-[#0D1220] px-4 py-2.5">
                  <label className="flex cursor-pointer items-center gap-2 text-sm text-white">
                    <Checkbox
                      checked={pickedSubIds.size > 0 && pickedSubIds.size === selected.ticketSubType.length}
                      onChange={(on) => setPickedSubIds(on ? selected.ticketSubType.map((s) => s.id) : [])}
                      title="Select all sub-types"
                    />
                    Select all ({selected.ticketSubType.length})
                  </label>
                  {pickedSubIds.size > 0 ? (
                    <>
                      <span className="text-sm text-indigo-200">{pickedSubIds.size} selected</span>
                      <div className="ml-auto flex items-center gap-2">
                        <Btn variant="ghost" small onClick={() => setPickedSubIds([])}>Clear</Btn>
                        <Btn variant="ghost" small onClick={() => copySubs(selected.ticketSubType.filter((x) => pickedSubIds.has(x.id)), selected.ticketType || selected.ticketKey)}><ClipboardCopy size={13} />Copy selected</Btn>
                        <Btn variant="danger" small onClick={deletePickedSubtypes}><Trash2 size={13} />Delete selected</Btn>
                      </div>
                    </>
                  ) : (
                    <span className="text-xs text-white/70">Tick sub-types to copy or delete several at once.</span>
                  )}
                </div>
              )}

              {selected.ticketSubType.map((s) => (
                <SubTypeEditor
                  key={s.id}
                  sub={s}
                  select={{
                    checked: pickedSubIds.has(s.id),
                    onChange: (on) => setPickedSubIds(on ? [...pickedSubIds, s.id] : [...pickedSubIds].filter((x) => x !== s.id)),
                  }}
                  forceOpen={s.id === focusSubtypeId}
                  onCopy={() => copySubs([s], selected.ticketType || selected.ticketKey)}
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
          {tab === "json" && selected && <JsonTab doc={selected} onBeforeExport={() => checkDocsForExport([selected], { goToRoles: true })} />}
        </div>
      </main>
      <ToastStack toasts={toasts} onDismiss={dismissToast} />
    </div>
  );
}