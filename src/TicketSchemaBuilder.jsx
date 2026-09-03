import React, { useState, useMemo, useCallback } from "react";
import {
  Plus, Trash2, ChevronDown, ChevronRight, Copy, Download,
  FileJson, Users, Ticket, ListTree, Workflow, Check, X,
  GripVertical, AlertTriangle, RefreshCw
} from "lucide-react";

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
    className={`w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-800 placeholder:text-slate-300 outline-none transition focus:border-teal-500 focus:ring-2 focus:ring-teal-100 ${mono ? "font-mono" : ""}`}
    {...rest}
  />
);

const Select = ({ value, onChange, options }) => (
  <select
    value={value}
    onChange={(e) => onChange(e.target.value)}
    className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-800 outline-none transition focus:border-teal-500 focus:ring-2 focus:ring-teal-100"
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
    className="flex items-center gap-2 text-sm text-slate-700"
  >
    <span className={`relative inline-flex h-5 w-9 items-center rounded-full transition ${checked ? "bg-teal-500" : "bg-slate-200"}`}>
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
    className={`inline-flex items-center justify-center h-7 w-7 rounded-md transition ${danger ? "text-rose-400 hover:bg-rose-50 hover:text-rose-600" : "text-slate-400 hover:bg-slate-100 hover:text-slate-700"}`}
  >
    {children}
  </button>
);

const Btn = ({ onClick, children, variant = "primary", small }) => {
  const base = "inline-flex items-center gap-1.5 rounded-lg font-medium transition";
  const size = small ? "px-2.5 py-1.5 text-xs" : "px-3.5 py-2 text-sm";
  const styles = {
    primary: "bg-slate-900 text-white hover:bg-slate-700",
    accent: "bg-teal-600 text-white hover:bg-teal-500",
    ghost: "bg-transparent text-slate-600 hover:bg-slate-100 border border-slate-200",
    dashed: "border border-dashed border-slate-300 text-slate-500 hover:border-teal-400 hover:text-teal-600 hover:bg-teal-50/50",
  };
  return (
    <button type="button" onClick={onClick} className={`${base} ${size} ${styles[variant]}`}>
      {children}
    </button>
  );
};

const Chip = ({ text, onRemove, tone = "slate" }) => {
  const tones = {
    slate: "bg-slate-100 text-slate-700",
    teal: "bg-teal-50 text-teal-700",
    amber: "bg-amber-50 text-amber-700",
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
    <div className="rounded-lg border border-slate-200 bg-white p-2">
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
          className={`rounded-md px-2.5 py-1 text-xs font-mono border transition ${active ? "border-teal-500 bg-teal-500 text-white" : "border-slate-200 text-slate-500 hover:border-slate-300"}`}
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
      <div className="flex items-center gap-2 text-slate-800 font-semibold text-sm">
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
  const tones = { slate: "border-slate-200", teal: "border-teal-200", amber: "border-amber-200" };
  return (
    <div className={`rounded-xl border ${tones[tone]} bg-white mb-3 overflow-hidden`}>
      <div className="flex items-center justify-between px-3 py-2.5 cursor-pointer select-none" onClick={() => setOpen(!open)}>
        <div className="flex items-center gap-2 min-w-0">
          {open ? <ChevronDown size={16} className="text-slate-400 shrink-0" /> : <ChevronRight size={16} className="text-slate-400 shrink-0" />}
          <div className="min-w-0">
            <div className="text-sm font-medium text-slate-800 truncate">{title}</div>
            {subtitle && <div className="text-xs text-slate-400 font-mono truncate">{subtitle}</div>}
          </div>
        </div>
        {onDelete && (
          <IconBtn danger title="Remove" onClick={(e) => { e.stopPropagation(); onDelete(); }}>
            <Trash2 size={14} />
          </IconBtn>
        )}
      </div>
      {open && <div className="px-3 pb-3 pt-1 border-t border-slate-100">{children}</div>}
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
        <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">Value (label shown)</span>
        <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">Key</span>
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
      <button type="button" onClick={add} className="mt-2 text-xs font-medium text-teal-600 hover:text-teal-700 inline-flex items-center gap-1">
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
              className={`px-2.5 py-1 rounded-md text-xs font-mono border ${field.parentOptionType === t ? "border-teal-500 bg-teal-500 text-white" : "border-slate-200 text-slate-500"}`}
            >
              {t}
            </button>
          ))}
        </div>
      </div>
      {options.map((p) => (
        <div key={p.id} className="rounded-lg border border-slate-100 bg-slate-50/60 p-2.5 mb-2.5">
          <div className="grid grid-cols-[1fr_120px_28px] gap-2 items-center mb-2">
            <Input value={p.name} onChange={(v) => updateParent(p.id, { name: v })} placeholder="Parent value, e.g. HC-House Hold-Bath" />
            <Input mono value={p.key} onChange={(v) => updateParent(p.id, { key: v })} placeholder="key" />
            <IconBtn danger title="Remove parent" onClick={() => removeParent(p.id)}><Trash2 size={14} /></IconBtn>
          </div>
          <div className="pl-3 border-l-2 border-teal-100 space-y-1.5">
            {p.children.map((c) => (
              <div key={c.id} className="grid grid-cols-[1fr_120px_28px] gap-2 items-center">
                <Input value={c.value} onChange={(v) => updateChild(p.id, c.id, { value: v })} placeholder="Child value, e.g. Buttercup" />
                <Input mono value={c.key} onChange={(v) => updateChild(p.id, c.id, { key: v })} placeholder="key" />
                <IconBtn danger title="Remove child" onClick={() => removeChild(p.id, c.id)}><Trash2 size={14} /></IconBtn>
              </div>
            ))}
            <button type="button" onClick={() => addChild(p.id)} className="text-xs font-medium text-teal-600 hover:text-teal-700 inline-flex items-center gap-1 mt-1">
              <Plus size={12} /> Add child option
            </button>
          </div>
        </div>
      ))}
      <button type="button" onClick={addParent} className="text-xs font-medium text-teal-600 hover:text-teal-700 inline-flex items-center gap-1">
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
        <p className="text-xs text-slate-400 italic">No options needed for this field type.</p>
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
        <div key={r.id} className="rounded-lg border border-amber-100 bg-amber-50/40 p-3 mb-2.5">
          <div className="flex justify-between items-center mb-2">
            <div className="flex items-center gap-1.5 text-amber-600 text-xs font-semibold"><AlertTriangle size={13} /> Escalation rule</div>
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
      <button type="button" onClick={add} className="text-xs font-medium text-amber-600 hover:text-amber-700 inline-flex items-center gap-1">
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
              <div className={`h-2.5 w-2.5 rounded-full ${s.status ? "bg-teal-500" : "bg-slate-300"}`} />
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
      <div className="flex gap-1 mb-4 border-b border-slate-100">
        {tabs.map((t) => (
          <button
            key={t.id}
            type="button"
            onClick={() => setTab(t.id)}
            className={`px-3 py-2 text-xs font-medium border-b-2 -mb-px transition ${tab === t.id ? "border-teal-500 text-teal-700" : "border-transparent text-slate-400 hover:text-slate-600"}`}
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
      <pre className="rounded-xl bg-slate-900 text-slate-100 text-xs leading-relaxed p-4 overflow-auto font-mono max-h-[70vh]">
        {json}
      </pre>
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

export default function App() {
  const [doc, setDoc] = useState(emptyDoc);
  const [tab, setTab] = useState("details");

  const nav = [
    { id: "details", label: "Details", icon: <Ticket size={16} /> },
    { id: "roles", label: "Roles", icon: <Users size={16} /> },
    { id: "subtypes", label: `Sub-Types (${doc.ticketSubType.length})`, icon: <ListTree size={16} /> },
    { id: "json", label: "JSON", icon: <FileJson size={16} /> },
  ];

  const loadSample = () => setDoc(SAMPLE());
  const resetAll = () => { if (window.confirm("Clear everything and start a new ticket type?")) setDoc(emptyDoc()); };

  return (
    <div className="min-h-screen bg-[#F4F6F9] flex text-slate-800" style={{ fontFamily: "Inter, ui-sans-serif, system-ui" }}>
      {/* sidebar */}
      <aside className="w-60 shrink-0 bg-[#151B2E] text-slate-300 flex flex-col">
        <div className="px-5 py-5 border-b border-white/5">
          <div className="flex items-center gap-2 text-white font-semibold">
            <Workflow size={18} className="text-teal-400" />
            Schema Builder
          </div>
          <div className="text-[11px] text-slate-500 mt-0.5">Ticket type dataset</div>
        </div>
        <nav className="flex-1 px-2 py-4 space-y-1">
          {nav.map((n) => (
            <button
              key={n.id}
              onClick={() => setTab(n.id)}
              className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm transition ${tab === n.id ? "bg-white/10 text-white" : "hover:bg-white/5 text-slate-400"}`}
            >
              {n.icon}
              {n.label}
            </button>
          ))}
        </nav>
        <div className="px-3 py-4 border-t border-white/5 space-y-1.5">
          <button onClick={loadSample} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-xs text-slate-400 hover:bg-white/5 hover:text-white transition">
            <RefreshCw size={13} /> Load sample data
          </button>
          <button onClick={resetAll} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-xs text-slate-400 hover:bg-white/5 hover:text-rose-300 transition">
            <Trash2 size={13} /> Start new
          </button>
        </div>
      </aside>

      {/* main */}
      <main className="flex-1 min-w-0">
        <header className="bg-white border-b border-slate-200 px-8 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-base font-semibold text-slate-800">
              {doc.ticketType || "Untitled ticket type"}
            </h1>
            <p className="text-xs font-mono text-slate-400">{doc.ticketKey || "no key set yet"}</p>
          </div>
          <Btn variant="accent" onClick={() => setTab("json")}><FileJson size={15} />Create JSON</Btn>
        </header>

        <div className="px-8 py-6 max-w-4xl">
          {tab === "details" && <DetailsTab doc={doc} onChange={setDoc} />}
          {tab === "roles" && <RolesTab doc={doc} onChange={setDoc} />}
          {tab === "subtypes" && (
            <div>
              {doc.ticketSubType.map((s) => (
                <SubTypeEditor
                  key={s.id}
                  sub={s}
                  onChange={(patch) => setDoc({ ...doc, ticketSubType: doc.ticketSubType.map((x) => (x.id === s.id ? patch : x)) })}
                  onDelete={() => setDoc({ ...doc, ticketSubType: doc.ticketSubType.filter((x) => x.id !== s.id) })}
                />
              ))}
              <Btn variant="dashed" onClick={() => setDoc({ ...doc, ticketSubType: [...doc.ticketSubType, newSubType()] })}>
                <Plus size={15} />Add sub-ticket type
              </Btn>
            </div>
          )}
          {tab === "json" && <JsonTab doc={doc} />}
        </div>
      </main>
    </div>
  );
}
