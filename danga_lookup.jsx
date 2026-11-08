import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { Search, Upload, Trash2, Pencil, Plus, X, ChevronDown, ChevronUp, Check, Save, CalendarDays, Printer, Download } from "lucide-react";

const PRODUCTS_KEY = "products";
const REASONS_KEY = "custom_reasons";
const FEE_KEY = "fee_rate";
const DELIVERIES_KEY = "deliveries";
// window.storage is provided by index.html and syncs these keys to Supabase with a browser fallback.
const STATUS_LIST = ["완료", "연기", "취소"];
const STATUS_STYLE = {
  완료: "bg-emerald-100 text-emerald-700",
  연기: "bg-amber-100 text-amber-800",
  취소: "bg-red-100 text-red-700",
};
const PRINT_CSS = `
.print-only { display: none; }
.print-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.print-table th, .print-table td { border: 1px solid #333; padding: 4px 6px; vertical-align: top; word-break: break-all; }
.print-table th { background: #eee; }
.print-table .num { text-align: right; white-space: nowrap; }
.print-table tfoot td { font-weight: 700; background: #f5f5f5; }
table.print-table.delivery-table { table-layout: fixed; margin: 0 auto; font-size: 15px; }
table.print-table.delivery-table th, table.print-table.delivery-table td { text-align: center; vertical-align: middle; padding: 8px 10px; word-break: keep-all; overflow-wrap: anywhere; }
table.print-table.delivery-table th:nth-child(1), table.print-table.delivery-table td:nth-child(1), table.print-table.delivery-table th:nth-child(2), table.print-table.delivery-table td:nth-child(2) { padding-left: 2px; padding-right: 2px; }
.delivery-title { font-size: 18px; font-weight: 700; margin-bottom: 8px; text-align: center; }
table.print-table.delivery-table td.left, table.print-table.delivery-table td.time-cell { text-align: left; }
.tp-on { display: inline-block; border: 1.5px solid #111; border-radius: 3px; padding: 0 6px; font-weight: 700; }
.tp-off { display: inline-block; padding: 0 7.5px; color: #9ca3af; }
.tp-note { margin-left: 4px; }
table.print-table.settle-table th, table.print-table.settle-table td { text-align: center; vertical-align: middle; }
table.print-table.settle-table td.left { text-align: left; }
table.print-table.settle-table td.num { text-align: right; }
html { color-scheme: light; }
body { margin: 0; background: #f8fafc; font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
.min-h-screen { background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%); }
.min-h-screen .no-print button { border-radius: 9px; transition: background-color .15s ease, border-color .15s ease, color .15s ease, box-shadow .15s ease, transform .15s ease; }
.min-h-screen .no-print button:hover:not(:disabled) { transform: translateY(-1px); }
.min-h-screen .no-print input:not([type="checkbox"]):not([type="radio"]), .min-h-screen .no-print select { border-radius: 8px; }
.min-h-screen .no-print input:focus, .min-h-screen .no-print select:focus { box-shadow: 0 0 0 3px rgba(99, 102, 241, .15); }
.min-h-screen .no-print section.bg-white { border-radius: 16px; box-shadow: 0 8px 24px rgba(15, 23, 42, .06); }
.min-h-screen .no-print .bg-white.border, .min-h-screen .no-print .bg-white.border-2 { border-radius: 12px; }
@media print {
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  @page { size: A4 landscape; margin: 8mm; }
  .no-print { display: none !important; }
  .print-only { display: block !important; }
  .min-h-screen { min-height: 0 !important; }
  .bg-slate-50 { background: #fff !important; }
  .print-table { font-size: 10px; }
  .print-table th, .print-table td { padding: 3px 4px; }
  table.print-table.delivery-table { font-size: 9px; }
  table.print-table.delivery-table th, table.print-table.delivery-table td { padding: 5px 6px; }
  .delivery-title { font-size: 9px; margin-bottom: 4px; }
  .tp-on { border-width: 1px; border-radius: 2px; padding: 0 3px; }
  .tp-off { padding: 0 4px; }
}
`;
const DEFAULT_REASONS = ["기본", "재출고", "프로모션", "내림서비스"];

const REASON_STYLE = {
  기본: "bg-slate-100 text-slate-700",
  재출고: "bg-blue-100 text-blue-700",
  프로모션: "bg-amber-100 text-amber-800",
  내림서비스: "bg-purple-100 text-purple-700",
};
function reasonStyle(reason) {
  return REASON_STYLE[reason] || "bg-slate-100 text-slate-700";
}

function tokenize(query) {
  return query
    .split(/[\s,]+/)
    .map((t) => t.replace(/^\+|\+$/g, "").trim())
    .filter((t) => t.length > 0);
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function normalizeName(s) {
  return String(s || "").trim().replace(/\s+/g, " ");
}

function Highlighted({ text, tokens }) {
  if (!tokens.length) return <span>{text}</span>;
  const escaped = [...tokens].sort((a, b) => b.length - a.length).map(escapeRegex);
  if (!escaped.length) return <span>{text}</span>;
  const regex = new RegExp("(" + escaped.join("|") + ")", "gi");
  const parts = text.split(regex);
  const lowerTokens = escaped.map((t) => t.toLowerCase());
  return (
    <span>
      {parts.map((part, i) =>
        lowerTokens.includes(part.toLowerCase()) ? (
          <mark key={i} className="bg-amber-200 text-slate-900 px-0.5 rounded-sm">
            {part}
          </mark>
        ) : (
          <span key={i}>{part}</span>
        )
      )}
    </span>
  );
}

function formatPrice(n) {
  const num = Number(n) || 0;
  return "₩" + num.toLocaleString("ko-KR");
}
function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function dateInputStr(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
const DEFAULT_FEE_RATE = 15;
function netOf(price, rate) {
  const r = Number(rate);
  const safe = Number.isFinite(r) ? Math.min(Math.max(r, 0), 100) : DEFAULT_FEE_RATE;
  return Math.round(((Number(price) || 0) * (100 - safe)) / 100);
}
function toNum(v) {
  const n = Number(String(v ?? "").replace(/[^0-9]/g, ""));
  return Number.isFinite(n) ? n : 0;
}
function fmtNum(n) {
  return n ? Number(n).toLocaleString("ko-KR") : "";
}
function deliveryTotal(d) {
  return (d.install || 0) + (d.lift || 0) + (d.unload || 0) + (d.etc || 0);
}
function shiftDate(dateStr, days) {
  const [y, m, d] = dateStr.split("-").map(Number);
  const dt = new Date(y, m - 1, d + days);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
}
function buildGroups(list, asOf) {
  const nameMap = new Map();
  for (const r of list) {
    const nkey = normalizeName(r.name);
    if (!nameMap.has(nkey)) nameMap.set(nkey, new Map());
    const rm = nameMap.get(nkey);
    const key = (r.reason || "기본").trim() || "기본";
    if (!rm.has(key)) rm.set(key, []);
    rm.get(key).push(r);
  }
  return Array.from(nameMap.entries())
    .map(([name, rm]) => ({
      name,
      reasonGroups: Array.from(rm.entries()).map(([reason, entries]) => {
        const sorted = [...entries].sort((a, b) => (b.effectiveDate || "").localeCompare(a.effectiveDate || ""));
        const idx = sorted.findIndex((e) => (e.effectiveDate || "") <= asOf);
        return { reason, current: idx > -1 ? sorted[idx] : sorted[sorted.length - 1] };
      }),
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}
const SHEET_HEAD = ["구분", "상태", "배송일", "송장번호", "박스", "상품명", "수량", "주소", "메모", "설치비", "양중비", "내림비", "기타비용", "합계"];
const FLAG_LIST = [
  { key: "direct", label: "직접설치", color: "#dc2626", badge: "bg-red-100 text-red-700" },
  { key: "reship", label: "재출고", color: "#2563eb", badge: "bg-blue-100 text-blue-700" },
  { key: "postpone", label: "연기", color: "#db2777", badge: "bg-pink-100 text-pink-700" },
  { key: "exchange", label: "교환", color: "#ea580c", badge: "bg-orange-100 text-orange-700" },
  { key: "partial", label: "부분출고", color: "#16a34a", badge: "bg-green-100 text-green-700" },
];
function flagsText(d) {
  return FLAG_LIST.filter((f) => d.flags && d.flags[f.key])
    .map((f) => f.label)
    .join(", ");
}
function shortDate(str) {
  const m = String(str || "").match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  return m ? `${Number(m[2])}/${Number(m[3])}` : String(str || "");
}
function FlagLines({ d }) {
  return FLAG_LIST.filter((f) => d.flags && d.flags[f.key]).map((f) => (
    <div key={f.key} style={{ color: f.color, fontWeight: 700 }}>
      {f.label}
    </div>
  ));
}
const MODE_LABEL = { net: "수수료 제외", applied: "적용단가", manual: "직접 수정" };
function uid() {
  return Date.now() + "-" + Math.random().toString(36).slice(2);
}
function hasNum(v) {
  return typeof v === "number" && Number.isFinite(v);
}
function optNum(v) {
  const n = Number(String(v ?? "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) && n > 0 ? n : null;
}
function itemsOf(d) {
  if (Array.isArray(d.items) && d.items.length) {
    return d.items.map((it) => ({ ...it, qty: Math.max(1, Math.floor(Number(it.qty) || 1)) }));
  }
  if (d.productName || d.install) {
    return [{ id: "legacy", name: d.productName || "", reason: d.reason || "기본", price: d.install || 0, qty: 1 }];
  }
  return [];
}
function sumItems(items) {
  return items.reduce((sum, it) => sum + (Number(it.price) || 0) * Math.max(1, Number(it.qty) || 1), 0);
}
function itemLabel(it) {
  return `${it.name || ""}${it.reason && it.reason !== "기본" ? ` (${it.reason})` : ""}`;
}
function productLabel(d) {
  const its = itemsOf(d);
  return its.length ? its.map((it) => `${itemLabel(it)}${Number(it.qty) > 1 ? ` × ${it.qty}` : ""}`).join(" + ") : d.productName || "";
}
function applyItems(d, items) {
  const reasons = [...new Set(items.map((it) => it.reason || "기본"))];
  return {
    ...d,
    items,
    install: sumItems(items),
    productName: items.map((it) => it.name).join(" + "),
    reason: reasons.join("+") || "기본",
  };
}
function isSettleSel(d) {
  return typeof d.settle === "boolean" ? d.settle : d.status === "완료";
}
function fmt0(v) {
  return (v || 0).toLocaleString("ko-KR");
}
function setNumFmt(ws, cols, fromRow, toRow) {
  for (let r = fromRow; r <= toRow; r++) {
    cols.forEach((c) => {
      if (ws[c + r]) ws[c + r].z = "#,##0";
    });
  }
}
function buildSettleSheet(rows, title) {
  const aoa = [[title], SHEET_HEAD];
  const merges = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 13 } }];
  const groups = [];
  const sums = [0, 0, 0, 0, 0];
  rows.forEach((d) => {
    const its = itemsOf(d);
    const lines = its.length ? its : [null];
    const first = aoa.length + 1;
    lines.forEach((it, k) => {
      const quantity = it ? Math.max(1, Number(it.qty) || 1) : 1;
      const inst = it ? (Number(it.price) || 0) * quantity : d.install || 0;
      aoa.push([
        k === 0 ? flagsText(d) : null,
        k === 0 ? d.status || "" : null,
        k === 0 ? shortDate(d.date || "") : null,
        k === 0 ? d.invoice || "" : null,
        k === 0 ? d.box || "" : null,
        it ? itemLabel(it) : "",
        quantity,
        k === 0 ? d.address || "" : null,
        k === 0 ? d.memo || "" : null,
        inst,
        k === 0 ? d.lift || 0 : null,
        k === 0 ? d.unload || 0 : null,
        k === 0 ? d.etc || 0 : null,
        k === 0 ? deliveryTotal(d) : null,
      ]);
      sums[0] += inst;
    });
    sums[1] += d.lift || 0;
    sums[2] += d.unload || 0;
    sums[3] += d.etc || 0;
    sums[4] += deliveryTotal(d);
    groups.push({ d, first, last: first + lines.length - 1 });
    if (lines.length > 1) merges.push({ s: { r: first - 1, c: 7 }, e: { r: first + lines.length - 2, c: 7 } });
  });
  const firstRow = 3;
  const lastRow = aoa.length;
  const foot = aoa.length + 1;
  aoa.push([`선택 ${rows.length}건 합계`, null, null, null, null, null, null, null, null, ...sums]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  groups.forEach((g) => {
    ws["N" + g.first] = {
      t: "n",
      f:
        g.last === g.first
          ? `SUM(J${g.first}:M${g.first})`
          : `SUM(J${g.first}:J${g.last})+SUM(K${g.first}:M${g.first})`,
      v: deliveryTotal(g.d),
    };
  });
  ["J", "K", "L", "M", "N"].forEach((c, k) => {
    ws[c + foot] = lastRow >= firstRow
      ? { t: "n", f: `SUM(${c}${firstRow}:${c}${lastRow})`, v: sums[k] }
      : { t: "n", v: 0 };
  });
  setNumFmt(ws, ["J", "K", "L", "M", "N"], firstRow, foot);
  merges.push({ s: { r: foot - 1, c: 0 }, e: { r: foot - 1, c: 8 } });
  ws["!merges"] = merges;
  ws["!cols"] = [6, 6, 11, 14, 8, 42, 8, 36, 20, 10, 10, 10, 10, 12].map((w) => ({ wch: w }));
  return ws;
}
function Field({ label, className = "", children }) {
  return (
    <label className={`block ${className}`}>
      <span className="block text-xs text-slate-500 mb-0.5">{label}</span>
      {children}
    </label>
  );
}
function ExcelFileIcon({ size = 16, ...props }) {
  return (
    <Icon size={size} {...props}>
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <path d="M14 2v6h6" />
      <path d="m8 12 4 6m0-6-4 6" />
      <path d="M14 13h3m-3 3h3" />
    </Icon>
  );
}
function excelCellToDateStr(raw) {
  if (raw === null || raw === undefined || raw === "") return "";
  if (typeof raw === "number") {
    try {
      const d = XLSX.SSF.parse_date_code(raw);
      if (d) return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
    } catch (e) {}
  }
  const s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}$/.test(s)) return s + "-01";
  const kor = s.match(/^(\d{2}|\d{4})\s*년\s*(\d{1,2})\s*월$/);
  if (kor) {
    const y = kor[1].length === 2 ? 2000 + Number(kor[1]) : Number(kor[1]);
    return `${y}-${String(kor[2]).padStart(2, "0")}-01`;
  }
  if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(s)) {
    const [y, m, d] = s.split("/");
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  return "";
}

export default function DangaLookup() {
  const [products, setProducts] = useState([]);
  const [customReasons, setCustomReasons] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [saveError, setSaveError] = useState(false);
  const [syncStatus, setSyncStatus] = useState(() => window.storage?.syncStatus || "connecting");
  const [query, setQuery] = useState("");
  const [manageOpen, setManageOpen] = useState(false);

  const [newName, setNewName] = useState("");
  const [newPrice, setNewPrice] = useState("");
  const [newApplied, setNewApplied] = useState("");
  const [newReason, setNewReason] = useState("기본");
  const [newEffectiveDate, setNewEffectiveDate] = useState(() => todayStr());
  const [addingCustomReason, setAddingCustomReason] = useState(false);
  const [customReasonInput, setCustomReasonInput] = useState("");

  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState({ name: "", price: "", applied: "", reason: "기본", effectiveDate: "" });
  const [expandedHistoryKeys, setExpandedHistoryKeys] = useState(() => new Set());
  function toggleHistory(key) {
    setExpandedHistoryKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }

  const [uploadMsg, setUploadMsg] = useState("");
  const fileInputRef = useRef(null);

  const [feeRate, setFeeRate] = useState(DEFAULT_FEE_RATE);

  const [recordsOpen, setRecordsOpen] = useState(true);

  const [confirmResetAll, setConfirmResetAll] = useState(false);
  const [confirmDeleteProductId, setConfirmDeleteProductId] = useState(null);

  const [view, setView] = useState("lookup");
  const [deliveries, setDeliveries] = useState([]);
  const [deliveryDate, setDeliveryDate] = useState(() => todayStr());
  const [pickQuery, setPickQuery] = useState("");
  const [pickTarget, setPickTarget] = useState(null);
  const [pickLimit, setPickLimit] = useState(40);
  const [confirmDeleteDeliveryId, setConfirmDeleteDeliveryId] = useState(null);
  const [showDeliveryPreview, setShowDeliveryPreview] = useState(false);
  const [settleListOpen, setSettleListOpen] = useState(true);
  const [settleOpenRows, setSettleOpenRows] = useState(() => new Set());
  const [expandedSettleDate, setExpandedSettleDate] = useState(null);
  const [recordsMode, setRecordsMode] = useState("month");
  const [recordsDate, setRecordsDate] = useState(() => deliveryDate);
  useEffect(() => { setRecordsDate(deliveryDate); }, [deliveryDate]);
  const [exportMsg, setExportMsg] = useState("");
  const [showSettlePreview, setShowSettlePreview] = useState(false);
  const [pickerOpen, setPickerOpen] = useState(true);
  const [listOpen, setListOpen] = useState(true);
  const [openRows, setOpenRows] = useState(() => new Set());

  function changeFeeRate(v) {
    const n = Number(v);
    const next = Number.isFinite(n) ? Math.min(Math.max(n, 0), 100) : 0;
    setFeeRate(next);
    try {
      window.storage.set(FEE_KEY, String(next), false).catch(() => {});
    } catch (e) {}
  }

  const allReasons = useMemo(() => {
    const set = new Set([...DEFAULT_REASONS, ...customReasons]);
    return Array.from(set);
  }, [customReasons]);

  const sortedProducts = useMemo(
    () => [...products].sort((a, b) => a.name.localeCompare(b.name) || (a.reason || "").localeCompare(b.reason || "") || (b.effectiveDate || "").localeCompare(a.effectiveDate || "")),
    [products]
  );
  const [manageQuery, setManageQuery] = useState("");
  const manageTokens = useMemo(() => tokenize(manageQuery), [manageQuery]);
  const filteredManageProducts = useMemo(() => {
    if (!manageTokens.length) return sortedProducts;
    const lowerTokens = manageTokens.map((t) => t.toLowerCase());
    return sortedProducts.filter((p) => {
      const lname = p.name.toLowerCase();
      return lowerTokens.every((t) => lname.includes(t));
    });
  }, [sortedProducts, manageTokens]);

  useEffect(() => {
    const onStorageStatus = (event) => setSyncStatus(event.detail?.status || window.storage?.syncStatus || "local");
    window.addEventListener("danga-storage-status", onStorageStatus);
    if (window.storage?.syncStatus) setSyncStatus(window.storage.syncStatus);
    return () => window.removeEventListener("danga-storage-status", onStorageStatus);
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const resFee = await window.storage.get(FEE_KEY, false);
        if (resFee && resFee.value) {
          const n = Number(resFee.value);
          if (Number.isFinite(n)) setFeeRate(Math.min(Math.max(n, 0), 100));
        }
      } catch (e) {
        // 최초 사용, 무시
      }
      try {
        const resDel = await window.storage.get(DELIVERIES_KEY, false);
        if (resDel && resDel.value) setDeliveries(JSON.parse(resDel.value));
      } catch (e) {
        // 최초 사용, 무시
      }
      try {
        const res = await window.storage.get(PRODUCTS_KEY, false);
        if (res && res.value) setProducts(JSON.parse(res.value));
      } catch (e) {
        // 최초 사용, 무시
      }
      try {
        const res2 = await window.storage.get(REASONS_KEY, false);
        if (res2 && res2.value) setCustomReasons(JSON.parse(res2.value));
      } catch (e) {
        // 최초 사용, 무시
      }
      setLoaded(true);
    })();
  }, []);

  async function persistProducts(next) {
    setProducts(next);
    try {
      const result = await window.storage.set(PRODUCTS_KEY, JSON.stringify(next), false);
      setSaveError(!result);
    } catch (e) {
      console.error(e);
      setSaveError(true);
    }
  }

  async function persistReasons(next) {
    setCustomReasons(next);
    try {
      await window.storage.set(REASONS_KEY, JSON.stringify(next), false);
    } catch (e) {
      console.error(e);
    }
  }

  function addCustomReasonToList(reason) {
    const r = reason.trim();
    if (!r || DEFAULT_REASONS.includes(r) || customReasons.includes(r)) return;
    persistReasons([...customReasons, r]);
  }

  const tokens = useMemo(() => tokenize(query), [query]);

  const results = useMemo(() => {
    if (!tokens.length) return [];
    const lowerTokens = tokens.map((t) => t.toLowerCase());
    const scored = products
      .map((p) => {
        const lname = p.name.toLowerCase();
        const matchCount = lowerTokens.filter((t) => lname.includes(t)).length;
        return { ...p, matchCount };
      })
      .filter((p) => p.matchCount > 0);
    const fullMatch = scored.filter((p) => p.matchCount === tokens.length);
    const list = fullMatch.length ? fullMatch : scored;
    return list.sort(
      (a, b) =>
        b.matchCount - a.matchCount ||
        a.name.localeCompare(b.name) ||
        (a.reason || "").localeCompare(b.reason || "")
    );
  }, [products, tokens]);

  const groupedResults = useMemo(() => {
    const today = todayStr();
    const nameMap = new Map();
    for (const r of results) {
      const nkey = normalizeName(r.name);
      if (!nameMap.has(nkey)) nameMap.set(nkey, new Map());
      const reasonMap = nameMap.get(nkey);
      const key = (r.reason || "기본").trim() || "기본";
      if (!reasonMap.has(key)) reasonMap.set(key, []);
      reasonMap.get(key).push(r);
    }
    return Array.from(nameMap.entries()).map(([name, reasonMap]) => ({
      name,
      reasonGroups: Array.from(reasonMap.entries()).map(([reason, entries]) => {
        const sorted = [...entries].sort((a, b) => (b.effectiveDate || "").localeCompare(a.effectiveDate || ""));
        const currentIdx = sorted.findIndex((e) => (e.effectiveDate || "") <= today);
        const current = currentIdx > -1 ? sorted[currentIdx] : sorted[sorted.length - 1];
        const history = sorted.filter((e) => e.id !== (current && current.id));
        return { reason, current, history };
      }),
    }));
  }, [results]);

  function handleAdd() {
    const name = newName.trim();
    const price = optNum(newPrice);
    const applied = optNum(newApplied);
    const reason = (addingCustomReason ? customReasonInput.trim() : newReason) || "기본";
    const effectiveDate = newEffectiveDate || todayStr();
    if (!name || (!price && !applied)) return;
    if (addingCustomReason && reason) addCustomReasonToList(reason);
    const next = [
      ...products,
      { id: Date.now() + "-" + Math.random().toString(36).slice(2), name, price, applied, reason, effectiveDate },
    ];
    persistProducts(next);
    setNewName("");
    setNewPrice("");
    setNewApplied("");
    setAddingCustomReason(false);
    setCustomReasonInput("");
  }

  function startEdit(p) {
    setEditingId(p.id);
    setEditDraft({
      name: p.name,
      price: hasNum(p.price) ? String(p.price) : "",
      applied: hasNum(p.applied) ? String(p.applied) : "",
      reason: p.reason || "기본",
      effectiveDate: p.effectiveDate || todayStr(),
    });
  }

  function cancelEdit() {
    setEditingId(null);
    setEditDraft({ name: "", price: "", applied: "", reason: "기본", effectiveDate: "" });
  }

  function saveEdit(id) {
    const name = editDraft.name.trim();
    const price = optNum(editDraft.price);
    const applied = optNum(editDraft.applied);
    const reason = editDraft.reason.trim() || "기본";
    const effectiveDate = editDraft.effectiveDate || todayStr();
    if (!name || (!price && !applied)) return;
    addCustomReasonToList(reason);
    const next = products.map((p) => (p.id === id ? { ...p, name, price, applied, reason, effectiveDate } : p));
    persistProducts(next);
    cancelEdit();
  }

  function handleDelete(id) {
    persistProducts(products.filter((p) => p.id !== id));
    setConfirmDeleteProductId(null);
  }

  function handleResetAll() {
    persistProducts([]);
    setConfirmResetAll(false);
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        let headerIdx = -1,
          nameCol = 0,
          priceCol = 1,
          appliedCol = -1,
          reasonCol = -1,
          dateCol = -1;
        for (let i = 0; i < Math.min(5, rows.length); i++) {
          const row = (rows[i] || []).map((c) => String(c));
          const nIdx = row.findIndex((c) => /상품|품명|제품/.test(c));
          let pIdx = row.findIndex((c) => /원단가/.test(c));
          if (pIdx === -1) pIdx = row.findIndex((c) => /단가|가격/.test(c) && !/적용단가/.test(c));
          const aIdx = row.findIndex((c) => /적용단가/.test(c));
          const rIdx = row.findIndex((c) => /구분|사유|타입|유형/.test(c));
          const dIdx = row.findIndex((c) => /적용일|적용월|기준일|날짜/.test(c));
          if (nIdx > -1 || pIdx > -1 || aIdx > -1 || rIdx > -1 || dIdx > -1) {
            headerIdx = i;
            if (nIdx > -1) nameCol = nIdx;
            if (pIdx > -1) priceCol = pIdx;
            if (aIdx > -1) appliedCol = aIdx;
            if (rIdx > -1) reasonCol = rIdx;
            if (dIdx > -1) dateCol = dIdx;
            break;
          }
        }
        const startRow = headerIdx > -1 ? headerIdx + 1 : 0;

        const imported = [];
        const newCustomReasons = [];
        for (let i = startRow; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          const name = String(row[nameCol] ?? "").trim();
          const price = optNum(row[priceCol]);
          const applied = appliedCol > -1 ? optNum(row[appliedCol]) : null;
          const reasonRaw = reasonCol > -1 ? String(row[reasonCol] ?? "").trim() : "";
          const reason = reasonRaw || "기본";
          const effectiveDate = dateCol > -1 ? excelCellToDateStr(row[dateCol]) || todayStr() : todayStr();
          if (!name || (!price && !applied)) continue;
          if (!DEFAULT_REASONS.includes(reason) && !newCustomReasons.includes(reason)) {
            newCustomReasons.push(reason);
          }
          imported.push({
            id: Date.now() + "-" + i + "-" + Math.random().toString(36).slice(2),
            name,
            price,
            applied,
            reason,
            effectiveDate,
          });
        }

        if (imported.length === 0) {
          setUploadMsg(
            "불러올 수 있는 데이터가 없어요. '상품명/품명', '원단가(또는 단가/가격)', '적용단가' 열이 있는지 확인해주세요. '구분/사유', '적용일' 열이 있으면 자동 인식해요."
          );
          return;
        }
        if (newCustomReasons.length) {
          const merged = Array.from(new Set([...customReasons, ...newCustomReasons]));
          persistReasons(merged);
        }
        persistProducts([...products, ...imported]);
        setUploadMsg(`${imported.length}건을 추가했어요.`);
      } catch (err) {
        console.error(err);
        setUploadMsg("파일을 읽는 중 문제가 생겼어요. CSV 또는 엑셀(xlsx) 파일인지 확인해주세요.");
      }
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  async function persistDeliveries(next) {
    setDeliveries(next);
    try {
      const result = await window.storage.set(DELIVERIES_KEY, JSON.stringify(next), false);
      setSaveError(!result);
    } catch (e) {
      console.error(e);
      setSaveError(true);
    }
  }
  function updateDelivery(id, patch) {
    persistDeliveries(deliveries.map((d) => (d.id === id ? { ...d, ...patch } : d)));
  }
  function updateItem(stopId, itemId, patch) {
    persistDeliveries(
      deliveries.map((d) =>
        d.id === stopId ? applyItems(d, itemsOf(d).map((it) => (it.id === itemId ? { ...it, ...patch } : it))) : d
      )
    );
  }
  function removeItem(stopId, itemId) {
    persistDeliveries(
      deliveries.map((d) => {
        if (d.id !== stopId) return d;
        const items = itemsOf(d).filter((it) => it.id !== itemId);
        return items.length ? applyItems(d, items) : d;
      })
    );
  }
  function toggleFlag(stopId, key) {
    persistDeliveries(
      deliveries.map((d) =>
        d.id === stopId ? { ...d, flags: { ...(d.flags || {}), [key]: !(d.flags && d.flags[key]) } } : d
      )
    );
  }
  function scrollToRow(id) {
    setTimeout(() => {
      try {
        const el = document.getElementById("delivery-" + id);
        if (el && el.scrollIntoView) el.scrollIntoView({ behavior: "smooth", block: "center" });
      } catch (e) {}
    }, 80);
  }
  function toggleRow(id) {
    setOpenRows((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }
  function makeItem(entry, mode) {
    const useApplied = mode === "applied" && hasNum(entry.applied);
    return {
      id: uid(),
      name: normalizeName(entry.name),
      reason: (entry.reason || "기본").trim() || "기본",
      qty: 1,
      price: useApplied ? entry.applied : netOf(entry.price, feeRate),
      priceMode: useApplied ? "applied" : "net",
      basePrice: hasNum(entry.price) ? entry.price : null,
      appliedPrice: hasNum(entry.applied) ? entry.applied : null,
      feeRate,
      productId: entry.id,
    };
  }
  function pickProduct(entry, mode) {
    const item = makeItem(entry, mode);
    let targetId;
    if (pickTarget && pickTarget.itemId) {
      targetId = pickTarget.stopId;
      persistDeliveries(
        deliveries.map((d) =>
          d.id === targetId
            ? applyItems(d, itemsOf(d).map((it) => (it.id === pickTarget.itemId ? { ...item, id: it.id } : it)))
            : d
        )
      );
    } else if (pickTarget) {
      targetId = pickTarget.stopId;
      persistDeliveries(deliveries.map((d) => (d.id === targetId ? applyItems(d, [...itemsOf(d), item]) : d)));
    } else {
      targetId = uid();
      persistDeliveries([
        ...deliveries,
        applyItems(
          {
            id: targetId,
            date: deliveryDate,
            status: "완료",
            invoice: "",
            box: "",
            address: "",
            memo: "",
            lift: 0,
            unload: 0,
            etc: 0,
            flags: {},
          },
          [item]
        ),
      ]);
    }
    setPickTarget(null);
    setOpenRows((prev) => new Set(prev).add(targetId));
    setListOpen(true);
    setPickerOpen(false);
    scrollToRow(targetId);
  }
  function openPickerFor(target) {
    setPickTarget(target);
    setPickerOpen(true);
    setTimeout(() => {
      try {
        const el = document.getElementById("delivery-picker");
        if (el && el.scrollIntoView) el.scrollIntoView({ behavior: "smooth", block: "start" });
      } catch (e) {}
    }, 60);
  }
  function handleDeleteDelivery(id) {
    persistDeliveries(deliveries.filter((d) => d.id !== id));
    setConfirmDeleteDeliveryId(null);
    setPickTarget((t) => (t && t.stopId === id ? null : t));
  }

  const dayRows = useMemo(() => deliveries.filter((d) => d.date === deliveryDate), [deliveries, deliveryDate]);
  const pickTokens = useMemo(() => tokenize(pickQuery), [pickQuery]);
  const nameSuggestions = useMemo(() => {
    const seen = new Set();
    const out = [];
    for (let i = deliveries.length - 1; i >= 0 && out.length < 300; i--) {
      itemsOf(deliveries[i]).forEach((it) => {
        const n = normalizeName(it.name);
        if (n && !seen.has(n)) {
          seen.add(n);
          out.push(n);
        }
      });
    }
    return out;
  }, [deliveries]);
  const pickSearch = useMemo(() => {
    if (!pickTokens.length) return { groups: [], partial: false };
    const lower = pickTokens.map((t) => t.toLowerCase());
    const scored = products
      .map((p) => {
        const ln = p.name.toLowerCase();
        return { p, matchCount: lower.filter((t) => ln.includes(t)).length };
      })
      .filter((x) => x.matchCount > 0);
    const full = scored.filter((x) => x.matchCount === lower.length);
    const list = full.length ? full : scored;
    const best = new Map();
    list.forEach((x) => {
      const k = normalizeName(x.p.name);
      best.set(k, Math.max(best.get(k) || 0, x.matchCount));
    });
    const groups = buildGroups(
      list.map((x) => x.p),
      deliveryDate
    ).sort((a, b) => (best.get(b.name) || 0) - (best.get(a.name) || 0) || a.name.localeCompare(b.name));
    return { groups, partial: full.length === 0 };
  }, [products, pickTokens, deliveryDate]);
  const pickAllGroups = pickSearch.groups;
  const pickPartial = pickSearch.partial;
  const pickGroups = pickAllGroups.slice(0, pickLimit);
  const daySummary = useMemo(() => {
    const by = { 완료: { count: 0, total: 0 }, 연기: { count: 0, total: 0 }, 취소: { count: 0, total: 0 } };
    const done = { install: 0, lift: 0, unload: 0, etc: 0, total: 0 };
    dayRows.forEach((d) => {
      const t = deliveryTotal(d);
      const slot = by[d.status] || by.완료;
      slot.count += 1;
      slot.total += t;
      if (slot === by.완료) {
        done.install += d.install || 0;
        done.lift += d.lift || 0;
        done.unload += d.unload || 0;
        done.etc += d.etc || 0;
        done.total += t;
      }
    });
    return { by, done };
  }, [dayRows]);

  const inputCls = "w-full border border-slate-300 px-2 py-1.5 text-base focus:outline-none focus:border-slate-900";
  const moneyCls = inputCls + " font-mono text-right";

  const sheetRows = dayRows
    .map((d, i) => ({ d, i }))
    .sort((x, y) => {
      const xa = normalizeName(x.d.address);
      const ya = normalizeName(y.d.address);
      if (!xa && !ya) return x.i - y.i;
      if (!xa) return 1;
      if (!ya) return -1;
      return xa.localeCompare(ya) || x.i - y.i;
    })
    .map((x) => x.d);
  const addrSpans = sheetRows.map((d, i) => {
    const a = normalizeName(d.address);
    if (!a) return 1;
    if (i > 0 && normalizeName(sheetRows[i - 1].address) === a) return 0;
    let n = 1;
    while (i + n < sheetRows.length && normalizeName(sheetRows[i + n].address) === a) n++;
    return n;
  });
  const deliverySheet = (
    <div>
      <h2 className="delivery-title">배송 리스트 · {deliveryDate}</h2>
      <table className="print-table delivery-table">
        <colgroup>
          <col style={{ width: "5%" }} />
          <col style={{ width: "5%" }} />
          <col style={{ width: "37%" }} />
          <col style={{ width: "38%" }} />
          <col style={{ width: "15%" }} />
        </colgroup>
        <thead>
          <tr>
            <th>구분</th>
            <th>박스</th>
            <th>상품명</th>
            <th>주소</th>
            <th>시간</th>
          </tr>
        </thead>
        <tbody>
          {sheetRows.map((d, idx) => (
            <tr key={d.id}>
              <td>
                <FlagLines d={d} />
              </td>
              <td>{d.box}</td>
              <td className="left">
                {itemsOf(d).map((it, i) => (
                  <div key={it.id || i}>{itemLabel(it)}</div>
                ))}
              </td>
              {addrSpans[idx] > 0 && (
                <td className="left" rowSpan={addrSpans[idx]}>
                  {d.address}
                </td>
              )}
              <td className="time-cell">
                <div>
                  <span className={d.timePeriod === "AM" ? "tp-on" : "tp-off"}>AM</span>
                  {d.timePeriod === "AM" && d.timeNote ? <span className="tp-note">{d.timeNote}</span> : null}
                </div>
                <div>
                  <span className={d.timePeriod === "PM" ? "tp-on" : "tp-off"}>PM</span>
                  {d.timePeriod === "PM" && d.timeNote ? <span className="tp-note">{d.timeNote}</span> : null}
                </div>
                {!d.timePeriod && d.timeNote ? <div className="tp-note">{d.timeNote}</div> : null}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );

  const pickTitle = pickTarget
    ? pickTarget.itemId
      ? "바꿀 상품을 선택하세요"
      : "이 배송지에 추가할 상품을 선택하세요"
    : "상품 선택해서 배송 건 추가";

  const deliveryView = (
    <div>
      <datalist id="delivery-name-suggestions">
        {nameSuggestions.map((n) => (
          <option key={n} value={n} />
        ))}
      </datalist>
      <div className="flex flex-wrap items-center gap-2">
        <button
          onClick={() => setDeliveryDate(shiftDate(deliveryDate, -1))}
          className="px-3 py-2 border border-slate-300 bg-white text-sm hover:bg-slate-100"
        >
          ◀ 전날
        </button>
        <input
          type="date"
          value={deliveryDate}
          onChange={(e) => e.target.value && setDeliveryDate(e.target.value)}
          className="border-2 border-slate-900 bg-white px-2 py-1.5 text-base font-mono focus:outline-none focus:border-amber-500"
        />
        <button
          onClick={() => setDeliveryDate(shiftDate(deliveryDate, 1))}
          className="px-3 py-2 border border-slate-300 bg-white text-sm hover:bg-slate-100"
        >
          다음날 ▶
        </button>
        <button
          onClick={() => setDeliveryDate(todayStr())}
          className="px-3 py-2 border border-slate-300 bg-white text-sm hover:bg-slate-100"
        >
          오늘
        </button>
        <div className="flex-1" />
        <button
          onClick={() => setShowDeliveryPreview((v) => !v)}
          disabled={dayRows.length === 0}
          className="px-4 py-2 border-2 border-slate-900 bg-white text-sm font-semibold hover:bg-slate-100 disabled:opacity-40"
        >
          {showDeliveryPreview ? "미리보기 닫기" : "미리보기"}
        </button>
        <button
          onClick={() => window.print()}
          className="flex items-center gap-1 bg-slate-900 text-white font-semibold text-sm px-4 py-2 hover:bg-slate-800"
        >
          <Printer size={16} />
          프린트
        </button>
      </div>
      <p className="text-xs text-slate-500 mt-2">
        {deliveryDate} · 배송 {dayRows.length}건 (완료 {daySummary.by.완료.count} · 연기 {daySummary.by.연기.count} · 취소{" "}
        {daySummary.by.취소.count})
      </p>

      {showDeliveryPreview && dayRows.length > 0 && (
        <div className="mt-4 bg-white border-2 border-slate-900 p-4">
          <div className="flex items-center justify-between mb-3">
            <span className="text-sm font-semibold">배송 리스트 미리보기 (인쇄 모양 그대로)</span>
            <button onClick={() => setShowDeliveryPreview(false)} className="text-xs text-slate-500 underline">
              닫기
            </button>
          </div>
          <div className="overflow-x-auto">
            <div style={{ minWidth: 720 }}>{deliverySheet}</div>
          </div>
        </div>
      )}

      {/* 배송 내역 */}
      <div className="mt-4 bg-white border border-slate-200">
        <div className="flex items-center gap-3 px-4 py-3">
          <button
            onClick={() => setListOpen((v) => !v)}
            aria-expanded={listOpen}
            className="flex flex-1 items-center gap-2 text-left text-sm font-semibold"
          >
            배송 내역 ({dayRows.length}건)
            {listOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
          </button>
          {listOpen && dayRows.length > 0 && (
            <span className="flex items-center gap-3 text-xs">
              <button
                onClick={() => setOpenRows((prev) => new Set([...prev, ...dayRows.map((d) => d.id)]))}
                className="underline text-slate-600"
              >
                모두 펼치기
              </button>
              <button
                onClick={() => setOpenRows((prev) => new Set([...prev].filter((id) => !dayRows.some((d) => d.id === id))))}
                className="underline text-slate-600"
              >
                모두 접기
              </button>
            </span>
          )}
        </div>
        {listOpen && (
          <div className="border-t border-slate-100 bg-slate-50 p-3 space-y-2">
            {dayRows.length === 0 ? (
              <p className="text-sm text-slate-400 py-4 text-center">
                {deliveryDate}에 등록된 배송 건이 없어요. 아래 "상품 선택해서 배송 건 추가"에서 상품을 검색해 추가해주세요.
              </p>
            ) : (
              dayRows.map((d, idx) => {
                const open = openRows.has(d.id);
                const its = itemsOf(d);
                const flags = d.flags || {};
                return (
                  <div
                    key={d.id}
                    id={"delivery-" + d.id}
                    className={`bg-white border ${open ? "border-slate-900" : "border-slate-200"}`}
                  >
                    <div className="flex items-center gap-2 px-3 py-2">
                      <button
                        onClick={() => toggleRow(d.id)}
                        aria-expanded={open}
                        className="flex flex-1 min-w-0 items-center gap-2 text-left"
                      >
                        {open ? (
                          <ChevronUp size={16} className="shrink-0 text-slate-400" />
                        ) : (
                          <ChevronDown size={16} className="shrink-0 text-slate-400" />
                        )}
                        <span className="text-xs font-bold text-slate-400 w-5 shrink-0">{idx + 1}</span>
                        <span className="min-w-0 flex-1">
                          <span className="flex items-center gap-1 min-w-0">
                            {FLAG_LIST.filter((f) => flags[f.key]).map((f) => (
                              <span
                                key={f.key}
                                className={`text-xs font-semibold px-1.5 py-0.5 rounded-sm shrink-0 ${f.badge}`}
                              >
                                {f.label}
                              </span>
                            ))}
                            {d.timePeriod && (
                              <span className="text-xs font-semibold px-1.5 py-0.5 rounded-sm shrink-0 bg-slate-200 text-slate-700">
                                {d.timePeriod}
                                {d.timeNote ? ` ${d.timeNote}` : ""}
                              </span>
                            )}
                            <span className="block text-sm truncate">
                              {its.length ? itemLabel(its[0]) : "상품 없음"}
                              {its.length > 1 ? ` 외 ${its.length - 1}건` : ""}
                            </span>
                          </span>
                          <span className="block text-xs text-slate-500 truncate">
                            {[d.address, d.invoice].filter(Boolean).join(" · ") || "주소·송장 미입력"}
                          </span>
                        </span>
                      </button>
                      <span
                        className={`shrink-0 rounded-full px-2.5 py-1 text-xs font-semibold ${STATUS_STYLE[d.status] || "bg-slate-100 text-slate-700"}`}
                      >
                        {d.status}
                      </span>
                    </div>

                    {open && (
                      <div className="border-t border-slate-200 p-3 sm:p-4">
                        <div className="flex flex-wrap items-center gap-3 mb-3">
                          <input
                            type="date"
                            value={d.date}
                            onChange={(e) => e.target.value && updateDelivery(d.id, { date: e.target.value })}
                            title="배송일 (바꾸면 해당 날짜 목록으로 이동해요)"
                            className="border border-slate-300 px-2 py-1 text-base font-mono focus:outline-none focus:border-slate-900"
                          />
                          {FLAG_LIST.map((f) => (
                            <label key={f.key} className="flex items-center gap-1 text-sm cursor-pointer">
                              <input
                                type="checkbox"
                                checked={!!flags[f.key]}
                                onChange={() => toggleFlag(d.id, f.key)}
                                className="w-4 h-4"
                              />
                              {f.label}
                            </label>
                          ))}
                          <div className="flex-1" />
                          {confirmDeleteDeliveryId !== d.id ? (
                            <button
                              onClick={() => setConfirmDeleteDeliveryId(d.id)}
                              className="p-2 text-slate-400 hover:text-red-600"
                              title="삭제"
                            >
                              <Trash2 size={14} />
                            </button>
                          ) : (
                            <span className="flex items-center gap-1 text-xs">
                              <button
                                onClick={() => handleDeleteDelivery(d.id)}
                                className="text-red-600 font-semibold hover:underline px-1"
                              >
                                삭제
                              </button>
                              <button
                                onClick={() => setConfirmDeleteDeliveryId(null)}
                                className="text-slate-400 hover:underline px-1"
                              >
                                취소
                              </button>
                            </span>
                          )}
                        </div>

                        <div className="flex flex-wrap items-center gap-2 mb-3">
                          <span className="text-xs text-slate-500">시간</span>
                          {["AM", "PM"].map((tp) => (
                            <button
                              key={tp}
                              onClick={() => updateDelivery(d.id, { timePeriod: d.timePeriod === tp ? "" : tp })}
                              aria-pressed={d.timePeriod === tp}
                              className={`px-3 py-1 text-sm font-semibold border ${
                                d.timePeriod === tp
                                  ? "border-slate-900 bg-slate-900 text-white"
                                  : "border-slate-300 bg-white text-slate-600 hover:bg-slate-100"
                              }`}
                            >
                              {tp}
                            </button>
                          ))}
                          <input
                            type="text"
                            value={d.timeNote || ""}
                            onChange={(e) => updateDelivery(d.id, { timeNote: e.target.value })}
                            placeholder="시간 상세 (예: 10시 이후)"
                            aria-label="시간 상세"
                            className="flex-1 basis-40 min-w-0 border border-slate-300 px-2 py-1 text-base focus:outline-none focus:border-slate-900"
                          />
                        </div>
                        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                          <Field label="송장번호">
                            <input
                              type="text"
                              value={d.invoice || ""}
                              onChange={(e) => updateDelivery(d.id, { invoice: e.target.value })}
                              className={inputCls + " font-mono"}
                            />
                          </Field>
                          <Field label="박스">
                            <input
                              type="text"
                              value={d.box || ""}
                              onChange={(e) => updateDelivery(d.id, { box: e.target.value })}
                              className={inputCls}
                            />
                          </Field>
                          <Field label="주소" className="col-span-2">
                            <input
                              type="text"
                              value={d.address || ""}
                              onChange={(e) => updateDelivery(d.id, { address: e.target.value })}
                              className={inputCls}
                            />
                          </Field>
                          <Field label="메모" className="col-span-2 sm:col-span-4">
                            <input
                              type="text"
                              value={d.memo || ""}
                              onChange={(e) => updateDelivery(d.id, { memo: e.target.value })}
                              className={inputCls}
                            />
                          </Field>
                        </div>

                        <div className="mt-3 border border-slate-200">
                          <div className="flex items-center justify-between bg-slate-50 px-3 py-2">
                            <span className="text-xs font-semibold text-slate-600">이 배송지의 상품 ({its.length})</span>
                            <button
                              onClick={() => openPickerFor({ stopId: d.id })}
                              className="flex items-center gap-1 text-xs font-semibold text-slate-700 border border-slate-300 bg-white px-2 py-1 hover:bg-slate-100"
                            >
                              <Plus size={12} />
                              상품 추가
                            </button>
                          </div>
                          <ul className="divide-y divide-slate-100">
                            {its.map((it) => (
                              <li key={it.id} className="flex flex-wrap items-center gap-2 px-3 py-2">
                                <span
                                  className={`text-xs font-semibold px-2 py-0.5 rounded-sm ${reasonStyle(
                                    it.reason || "기본"
                                  )}`}
                                >
                                  {it.reason || "기본"}
                                </span>
                                <input
                                  type="text"
                                  list="delivery-name-suggestions"
                                  value={it.name || ""}
                                  onChange={(e) => updateItem(d.id, it.id, { name: e.target.value })}
                                  placeholder="상품명 (이전에 입력한 상품명이 추천돼요)"
                                  title="상품명 (직접 수정할 수 있어요)"
                                  className="flex-1 basis-48 min-w-0 border border-slate-300 px-2 py-1.5 text-base focus:outline-none focus:border-slate-900"
                                />
                                <label className="flex items-center gap-1 text-xs text-slate-500 whitespace-nowrap">
                                  수량
                                  <input
                                    type="number"
                                    min="1"
                                    step="1"
                                    inputMode="numeric"
                                    aria-label="상품 수량"
                                    value={it.qty || 1}
                                    onChange={(e) => updateItem(d.id, it.id, { qty: Math.max(1, parseInt(e.target.value, 10) || 1) })}
                                    className="w-16 border border-slate-300 px-1.5 py-1 text-base text-center focus:outline-none focus:border-slate-900"
                                  />
                                </label>
                                <span className="text-xs text-slate-500 whitespace-nowrap">
                                  {formatPrice(it.price)} / 개 · 소계 {formatPrice((Number(it.price) || 0) * (it.qty || 1))}
                                  {MODE_LABEL[it.priceMode] ? (
                                    <span className="text-slate-400"> ({MODE_LABEL[it.priceMode]})</span>
                                  ) : null}
                                </span>
                                <button
                                  onClick={() => openPickerFor({ stopId: d.id, itemId: it.id })}
                                  className="text-xs text-slate-500 underline hover:text-slate-700"
                                >
                                  상품 변경
                                </button>
                                {its.length > 1 && (
                                  <button
                                    onClick={() => removeItem(d.id, it.id)}
                                    className="p-1 text-slate-400 hover:text-red-600"
                                    title="이 상품 빼기"
                                  >
                                    <X size={14} />
                                  </button>
                                )}
                              </li>
                            ))}
                          </ul>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })
            )}
          </div>
        )}
      </div>

      {/* 상품 선택해서 배송 건 추가 (배송 내역 아래) */}
      <div id="delivery-picker" className="mt-4 bg-white border border-slate-200">
        <button
          onClick={() => setPickerOpen((v) => !v)}
          aria-expanded={pickerOpen}
          className="w-full flex items-center justify-between gap-2 px-4 py-3 text-left"
        >
          <span className="flex items-center gap-2 text-sm font-semibold">
            <Plus size={16} />
            {pickTitle}
          </span>
          {pickerOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
        </button>
        {pickerOpen && (
          <div className="px-4 pb-4 pt-3 border-t border-slate-100">
            <p className="text-xs text-slate-500 mb-2">
              상품을 검색하고 수수료 제외 단가 또는 적용단가를 선택하면 {deliveryDate} 배송에 추가돼요. 수수료율은 단가 조회에서 변경할 수 있어요.
            </p>
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
              <input
                type="text"
                value={pickQuery}
                onChange={(e) => {
                  setPickQuery(e.target.value);
                  setPickLimit(40);
                }}
                placeholder="예: 이즈 카이"
                className="w-full border-2 border-slate-900 pl-9 pr-3 py-2 text-base focus:outline-none focus:border-amber-500"
              />
            </div>
            {pickTarget && (
              <div className="mt-2 flex items-center gap-2 text-xs">
                <span className="text-amber-700">
                  {pickTarget.itemId ? "선택한 상품을 바꾸는 중이에요." : "선택한 배송지에 상품을 추가하는 중이에요."}
                </span>
                <button onClick={() => setPickTarget(null)} className="underline text-slate-500">
                  취소
                </button>
              </div>
            )}
            {pickQuery.trim() &&
              (pickGroups.length === 0 ? (
                <p className="text-sm text-slate-500 mt-3">일치하는 상품이 없어요.</p>
              ) : (
                <>
                  {pickPartial && (
                    <p className="mt-3 text-xs text-amber-700">
                      검색어를 모두 포함한 상품이 없어서, 일부만 일치하는 상품을 보여줘요.
                    </p>
                  )}
                  <div className="mt-3 border border-slate-200 divide-y divide-slate-200 max-h-80 overflow-y-auto">
                    {pickGroups.map((g) => (
                      <div key={g.name} className="px-3 py-2">
                        <div className="text-sm leading-snug mb-1">
                          <Highlighted text={g.name} tokens={pickTokens} />
                        </div>
                        <div className="flex flex-wrap gap-2">
                          {g.reasonGroups.map((rg) => {
                            const cur = rg.current;
                            const hasNet = hasNum(cur.price) && cur.price > 0;
                            const hasApp = hasNum(cur.applied);
                            const netVal = hasNet ? netOf(cur.price, feeRate) : 0;
                            const diff = hasNet && hasApp ? cur.applied - netVal : 0;
                            return (
                              <div
                                key={rg.reason}
                                className="flex flex-wrap items-center gap-2 border border-slate-200 bg-slate-50 pl-2 pr-2 py-1"
                              >
                                <span
                                  className={`text-xs font-semibold px-2 py-0.5 rounded-sm ${reasonStyle(rg.reason)}`}
                                >
                                  {rg.reason}
                                </span>
                                {hasNet && (
                                  <button
                                    onClick={() => pickProduct(cur, "net")}
                                    title="수수료 제외 단가로 선택"
                                    className="flex flex-col items-end leading-tight border border-blue-200 bg-white hover:bg-blue-50 px-2 py-0.5"
                                  >
                                    <span className="text-xs text-slate-400">수수료 {feeRate}% 제외</span>
                                    <span className="font-mono text-blue-700 font-semibold text-sm">
                                      {formatPrice(netVal)}
                                    </span>
                                  </button>
                                )}
                                {hasApp && (
                                  <button
                                    onClick={() => pickProduct(cur, "applied")}
                                    title="적용단가로 선택"
                                    className="flex flex-col items-end leading-tight border border-purple-200 bg-white hover:bg-purple-50 px-2 py-0.5"
                                  >
                                    <span className="text-xs text-slate-400">적용단가</span>
                                    <span className="font-mono text-purple-700 font-semibold text-sm">
                                      {formatPrice(cur.applied)}
                                    </span>
                                  </button>
                                )}
                                {hasNet && hasApp && diff !== 0 && (
                                  <span className="text-xs text-slate-400">
                                    적용단가가 {formatPrice(Math.abs(diff))} {diff > 0 ? "높아요" : "낮아요"}
                                  </span>
                                )}
                                {!hasNet && !hasApp && <span className="text-xs text-slate-400">단가 없음</span>}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    ))}
                  </div>
                  <div className="mt-2 flex flex-wrap items-center gap-3 text-xs text-slate-500">
                    <span>
                      단가 조회의 전체 상품에서 검색해요 · 결과 {pickAllGroups.length}개 중 {pickGroups.length}개 표시
                    </span>
                    {pickAllGroups.length > pickGroups.length && (
                      <button
                        onClick={() => setPickLimit((n) => n + 40)}
                        className="underline font-semibold text-slate-700"
                      >
                        더 보기 (+40)
                      </button>
                    )}
                  </div>
                </>
              ))}
          </div>
        )}
      </div>
    </div>
  );

  const settleSel = useMemo(() => dayRows.filter(isSettleSel), [dayRows]);
  const settleSummary = useMemo(() => {
    const t = { install: 0, lift: 0, unload: 0, etc: 0, total: 0 };
    settleSel.forEach((d) => {
      t.install += d.install || 0;
      t.lift += d.lift || 0;
      t.unload += d.unload || 0;
      t.etc += d.etc || 0;
      t.total += deliveryTotal(d);
    });
    return t;
  }, [settleSel]);
  const deliveryRecords = useMemo(() => {
    const byDate = new Map();
    deliveries.filter(isSettleSel).forEach((d) => {
      if (!byDate.has(d.date)) byDate.set(d.date, []);
      byDate.get(d.date).push(d);
    });
    return Array.from(byDate.entries())
      .map(([date, items]) => ({
        date,
        items,
        count: items.length,
        install: items.reduce((sum, d) => sum + (d.install || 0), 0),
        lift: items.reduce((sum, d) => sum + (d.lift || 0), 0),
        unload: items.reduce((sum, d) => sum + (d.unload || 0), 0),
        etc: items.reduce((sum, d) => sum + (d.etc || 0), 0),
        total: items.reduce((sum, d) => sum + deliveryTotal(d), 0),
      }))
      .sort((x, y) => (x.date < y.date ? 1 : -1));
  }, [deliveries]);

  const recordsByDate = useMemo(() => new Map(deliveryRecords.map((r) => [r.date, r])), [deliveryRecords]);
  const recordsYear = Number(recordsDate.slice(0, 4));
  const recordsMonth = Number(recordsDate.slice(5, 7));
  const recordsMonthPrefix = recordsDate.slice(0, 7);
  const recordsYearPrefix = recordsDate.slice(0, 4);
  const dailyRecord = recordsByDate.get(recordsDate);
  const monthlyRecords = useMemo(() => deliveryRecords.filter((r) => r.date.startsWith(recordsMonthPrefix)), [deliveryRecords, recordsMonthPrefix]);
  const yearlyRecords = useMemo(() => deliveryRecords.filter((r) => r.date.startsWith(recordsYearPrefix)), [deliveryRecords, recordsYearPrefix]);
  const annualMonths = useMemo(() => Array.from({ length: 12 }, (_, i) => {
    const prefix = `${recordsYearPrefix}-${String(i + 1).padStart(2, "0")}`;
    const rows = deliveryRecords.filter((r) => r.date.startsWith(prefix));
    return { month: i + 1, prefix, days: rows.length, count: rows.reduce((sum, r) => sum + r.count, 0), total: rows.reduce((sum, r) => sum + r.total, 0) };
  }), [deliveryRecords, recordsYearPrefix]);
  const calendarCells = useMemo(() => {
    const first = new Date(recordsYear, recordsMonth - 1, 1);
    const offset = first.getDay();
    const count = new Date(recordsYear, recordsMonth, 0).getDate();
    return [...Array(offset).fill(null), ...Array.from({ length: count }, (_, i) => dateInputStr(new Date(recordsYear, recordsMonth - 1, i + 1)))];
  }, [recordsYear, recordsMonth]);
  const recordsDayRows = useMemo(() => deliveries.filter((d) => d.date === recordsDate), [deliveries, recordsDate]);
  const recordsMonthRows = useMemo(() => deliveries.filter((d) => String(d.date || "").startsWith(recordsMonthPrefix)), [deliveries, recordsMonthPrefix]);
  const recordsYearRows = useMemo(() => deliveries.filter((d) => String(d.date || "").startsWith(recordsYearPrefix)), [deliveries, recordsYearPrefix]);

  function toggleSettle(d) {
    updateDelivery(d.id, { settle: !isSettleSel(d) });
  }
  function setAllSettle(on) {
    const ids = new Set(dayRows.map((d) => d.id));
    persistDeliveries(deliveries.map((d) => (ids.has(d.id) ? { ...d, settle: on } : d)));
  }
  function toggleSettleRow(id) {
    setSettleOpenRows((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }
  function downloadWorkbook(wb, filename) {
    try {
      XLSX.writeFile(wb, filename);
      setExportMsg("");
    } catch (e) {
      console.error(e);
      setExportMsg("엑셀 파일을 만들지 못했어요. 배포된 사이트에서 다시 시도해주세요.");
    }
  }
  function exportDayExcel() {
    const rows = [...dayRows].sort((a, b) => (a.date || "").localeCompare(b.date || ""));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, buildSettleSheet(rows, `배송 내역 · ${deliveryDate} (전체 ${rows.length}건)`), "배송내역");
    downloadWorkbook(wb, `배송내역_${deliveryDate}.xlsx`);
  }
  function exportAllExcel() {
    const rows = [...deliveries].sort((a, b) => (a.date || "").localeCompare(b.date || ""));
    const title = rows.length
      ? `배송 내역 · 전체 (${rows[0].date} ~ ${rows[rows.length - 1].date}, ${rows.length}건)`
      : "배송 내역 · 전체";
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, buildSettleSheet(rows, title), "배송내역");
    downloadWorkbook(wb, `배송내역_전체_${todayStr().replace(/-/g, "")}.xlsx`);
  }

  function shiftRecordsCursor(delta) {
    const d = new Date(`${recordsDate}T12:00:00`);
    if (recordsMode === "day") d.setDate(d.getDate() + delta);
    else if (recordsMode === "month") { d.setDate(1); d.setMonth(d.getMonth() + delta); }
    else d.setFullYear(d.getFullYear() + delta);
    const next = dateInputStr(d);
    setRecordsDate(next);
    setDeliveryDate(next);
  }

  function exportRecordsPeriod(period) {
    const key = period === "day" ? recordsDate : period === "month" ? recordsMonthPrefix : recordsYearPrefix;
    const rows = deliveries
      .filter((d) => period === "day" ? d.date === key : String(d.date || "").startsWith(key))
      .sort((a, b) => (a.date || "").localeCompare(b.date || ""));
    if (!rows.length) return;
    const title = `배송 내역 · ${period === "day" ? key : period === "month" ? `${key}월` : `${key}년`} (전체 ${rows.length}건)`;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, buildSettleSheet(rows, title), "배송내역");
    downloadWorkbook(wb, `배송내역_${key}.xlsx`);
  }

  const headTexts = {
    lookup: [
      "가구 배송 단가 조회",
      "상품명 또는 키워드를 입력하면 단가를 보여드려요",
      "같은 상품도 재출고·프로모션·내림서비스 등 사유별로 단가가 다르면 전부 따로 보여줘요.",
    ],
    delivery: [
      "일별 배송 리스트",
      "날짜별로 배송 건을 정리하고 프린트해요",
      "배송지마다 상품과 수량을 관리하고, 배송 상태는 정산 페이지에서 선택할 수 있어요.",
    ],
    settle: [
      "정산 페이지",
      "배송 건을 체크해 정산서를 미리 보고 프린트·엑셀로 받아요",
      "상품별 단가와 수량을 수정하고, 정산 기록을 달력에서 일·월·년 단위로 확인하고 엑셀로 내려받을 수 있어요.",
    ],
  };

  const printView = <div className="print-only">{deliverySheet}</div>;

  const settleSheet = (
    <div>
      <h2 style={{ fontSize: 16, fontWeight: 700, marginBottom: 6 }}>
        정산서 · {deliveryDate} (선택 {settleSel.length}건)
      </h2>
      <table className="print-table settle-table">
        <thead>
          <tr>
            <th>구분</th>
            <th>상태</th>
            <th>배송일</th>
            <th>송장번호</th>
            <th>박스</th>
            <th>상품명</th>
            <th>수량</th>
            <th>주소</th>
            <th>메모</th>
            <th>설치비</th>
            <th>양중비</th>
            <th>내림비</th>
            <th>기타비용</th>
            <th>합계</th>
          </tr>
        </thead>
        <tbody>
          {settleSel.flatMap((d) => {
            const its = itemsOf(d);
            const lines = its.length ? its : [null];
            return lines.map((it, k) => (
              <tr key={d.id + "-" + k}>
                <td>{k === 0 ? <FlagLines d={d} /> : null}</td>
                <td>{k === 0 ? d.status : null}</td>
                <td>{k === 0 ? shortDate(d.date) : null}</td>
                <td>{k === 0 ? d.invoice : null}</td>
                <td>{k === 0 ? d.box : null}</td>
                <td className="left">{it ? itemLabel(it) : ""}</td>
                <td className="num">{it ? fmtNum(it.qty || 1) : "1"}</td>
                {k === 0 && (
                  <td className="left" rowSpan={lines.length}>
                    {d.address}
                  </td>
                )}
                <td>{k === 0 ? d.memo : null}</td>
                <td className="num">{fmtNum(it ? it.price * (it.qty || 1) : k === 0 ? d.install : 0)}</td>
                <td className="num">{k === 0 ? fmtNum(d.lift) : null}</td>
                <td className="num">{k === 0 ? fmtNum(d.unload) : null}</td>
                <td className="num">{k === 0 ? fmtNum(d.etc) : null}</td>
                <td className="num">{k === 0 ? fmtNum(deliveryTotal(d)) : null}</td>
              </tr>
            ));
          })}
        </tbody>
        <tfoot>
          <tr>
            <td colSpan={9}>선택 {settleSel.length}건 합계</td>
            <td className="num">{fmtNum(settleSummary.install)}</td>
            <td className="num">{fmtNum(settleSummary.lift)}</td>
            <td className="num">{fmtNum(settleSummary.unload)}</td>
            <td className="num">{fmtNum(settleSummary.etc)}</td>
            <td className="num">{fmtNum(settleSummary.total)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
  );
  const settlePrintView = <div className="print-only">{settleSheet}</div>;

  const settleView = (
    <div>
      <div className="flex flex-wrap items-center gap-2">
        <button
          onClick={() => setDeliveryDate(shiftDate(deliveryDate, -1))}
          className="px-3 py-2 border border-slate-300 bg-white text-sm hover:bg-slate-100"
        >
          ◀ 전날
        </button>
        <input
          type="date"
          value={deliveryDate}
          onChange={(e) => e.target.value && setDeliveryDate(e.target.value)}
          className="border-2 border-slate-900 bg-white px-2 py-1.5 text-base font-mono focus:outline-none focus:border-amber-500"
        />
        <button
          onClick={() => setDeliveryDate(shiftDate(deliveryDate, 1))}
          className="px-3 py-2 border border-slate-300 bg-white text-sm hover:bg-slate-100"
        >
          다음날 ▶
        </button>
        <button
          onClick={() => setDeliveryDate(todayStr())}
          className="px-3 py-2 border border-slate-300 bg-white text-sm hover:bg-slate-100"
        >
          오늘
        </button>
        <div className="flex-1" />
        <button
          onClick={() => setShowSettlePreview((v) => !v)}
          disabled={settleSel.length === 0}
          className="px-4 py-2 border-2 border-slate-900 bg-white text-sm font-semibold hover:bg-slate-100 disabled:opacity-40"
        >
          {showSettlePreview ? "미리보기 닫기" : "미리보기"}
        </button>
        <button
          onClick={exportDayExcel}
          disabled={dayRows.length === 0}
          title="선택한 날짜의 완료·연기·취소 배송 건을 모두 엑셀 다운로드"
          className="flex items-center gap-1 px-4 py-2 border-2 border-slate-900 bg-white text-sm font-semibold hover:bg-slate-100 disabled:opacity-40"
        >
          <Download size={16} />
          엑셀 (당일 전체)
        </button>
        <button
          onClick={() => window.print()}
          disabled={settleSel.length === 0}
          className="flex items-center gap-1 bg-slate-900 text-white font-semibold text-sm px-4 py-2 hover:bg-slate-800 disabled:opacity-40"
        >
          <Printer size={16} />
          프린트
        </button>
      </div>
      <p className="text-xs text-slate-500 mt-2">
        {deliveryDate} · 배송 {dayRows.length}건 중 <b>{settleSel.length}건 선택</b> · 선택 합계{" "}
        <b className="text-emerald-700">{formatPrice(settleSummary.total)}</b>
      </p>
      {exportMsg && <p className="text-xs text-red-600 mt-1">{exportMsg}</p>}

      {showSettlePreview && settleSel.length > 0 && (
        <div className="mt-4 bg-white border-2 border-slate-900 p-4">
          <div className="flex items-center justify-between mb-3">
            <span className="text-sm font-semibold">정산서 미리보기 (인쇄 모양 그대로)</span>
            <button onClick={() => setShowSettlePreview(false)} className="text-xs text-slate-500 underline">
              닫기
            </button>
          </div>
          <div className="overflow-x-auto">
            <div style={{ minWidth: 860 }}>{settleSheet}</div>
          </div>
        </div>
      )}

      <div className="mt-4 bg-white border border-slate-200">
        <div className="flex items-center gap-3 px-4 py-3">
          <button
            onClick={() => setSettleListOpen((v) => !v)}
            aria-expanded={settleListOpen}
            className="flex flex-1 items-center gap-2 text-left text-sm font-semibold"
          >
            정산 대상 배송 내역 ({dayRows.length}건)
            {settleListOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
          </button>
          {settleListOpen && dayRows.length > 0 && (
            <span className="flex items-center gap-3 text-xs">
              <button
                onClick={() => setSettleOpenRows((prev) => new Set([...prev, ...dayRows.map((d) => d.id)]))}
                className="underline text-slate-600"
              >
                모두 펼치기
              </button>
              <button
                onClick={() =>
                  setSettleOpenRows((prev) => new Set([...prev].filter((id) => !dayRows.some((d) => d.id === id))))
                }
                className="underline text-slate-600"
              >
                모두 접기
              </button>
            </span>
          )}
        </div>
        {settleListOpen && (
          <div className="border-t border-slate-100 bg-slate-50 p-3 space-y-2">
            {dayRows.length > 0 && (
              <div className="flex flex-wrap items-center gap-3 text-xs pb-1">
                <button onClick={() => setAllSettle(true)} className="underline text-slate-600">
                  전체 선택
                </button>
                <button onClick={() => setAllSettle(false)} className="underline text-slate-600">
                  전체 해제
                </button>
                <span className="text-slate-400">처음에는 완료 건만 체크돼 있어요. 체크한 건이 정산 기록에 쌓여요.</span>
              </div>
            )}
            {dayRows.length === 0 ? (
              <p className="text-sm text-slate-400 py-4 text-center">
                {deliveryDate}에 배송 건이 없어요. 「배송 리스트」 탭에서 먼저 추가해주세요.
              </p>
            ) : (
              dayRows.map((d, idx) => {
                const sel = isSettleSel(d);
                const open = settleOpenRows.has(d.id);
                return (
                  <div
                    key={d.id}
                    className={`bg-white border ${sel ? "border-slate-900" : "border-slate-200 opacity-70"}`}
                  >
                    <div className="flex items-center gap-2 px-3 py-2">
                      <input
                        type="checkbox"
                        checked={sel}
                        onChange={() => toggleSettle(d)}
                        className="w-5 h-5 shrink-0"
                        title="정산서·정산 기록에 포함"
                      />
                      <button
                        onClick={() => toggleSettleRow(d.id)}
                        aria-expanded={open}
                        className="flex flex-1 min-w-0 items-center gap-2 text-left"
                      >
                        {open ? (
                          <ChevronUp size={16} className="shrink-0 text-slate-400" />
                        ) : (
                          <ChevronDown size={16} className="shrink-0 text-slate-400" />
                        )}
                        <span className="text-xs font-bold text-slate-400 w-5 shrink-0">{idx + 1}</span>
                        <span className="min-w-0 flex-1">
                          <span className="flex items-center gap-1 min-w-0">
                            <span
                              className={`text-xs font-semibold px-1.5 py-0.5 rounded-sm shrink-0 ${
                                STATUS_STYLE[d.status] || ""
                              }`}
                            >
                              {d.status}
                            </span>
                            {d.reason && d.reason !== "기본" && (
                              <span
                                className={`text-xs font-semibold px-1.5 py-0.5 rounded-sm shrink-0 ${reasonStyle(
                                  d.reason
                                )}`}
                              >
                                {d.reason}
                              </span>
                            )}
                            {FLAG_LIST.filter((f) => d.flags && d.flags[f.key]).map((f) => (
                              <span
                                key={f.key}
                                className={`text-xs font-semibold px-1.5 py-0.5 rounded-sm shrink-0 ${f.badge}`}
                              >
                                {f.label}
                              </span>
                            ))}
                            <span className="block text-sm truncate">
                              {itemsOf(d).length ? itemLabel(itemsOf(d)[0]) : d.productName}
                              {itemsOf(d).length > 1 ? ` 외 ${itemsOf(d).length - 1}건` : ""}
                            </span>
                          </span>
                          <span className="block text-xs text-slate-500 truncate">{d.address || "주소 미입력"}</span>
                        </span>
                      </button>
                      <select
                        value={d.status}
                        onChange={(e) => updateDelivery(d.id, { status: e.target.value })}
                        aria-label="배송 상태"
                        className={`shrink-0 border px-2 py-1.5 text-xs font-semibold ${STATUS_STYLE[d.status] || "bg-slate-100 text-slate-700"}`}
                      >
                        {STATUS_LIST.map((st) => (
                          <option key={st} value={st}>{st}</option>
                        ))}
                      </select>
                      <span className="font-mono text-emerald-700 font-semibold text-sm whitespace-nowrap">
                        {formatPrice(deliveryTotal(d))}
                      </span>
                    </div>
                    {open && (
                      <div className="border-t border-slate-200 p-3 sm:p-4">
                        <div className="mb-3 border border-slate-200">
                          <p className="bg-slate-50 px-3 py-1.5 text-xs font-semibold text-slate-600">
                            상품별 단가 (수정할 수 있어요)
                          </p>
                          <ul className="divide-y divide-slate-100">
                            {itemsOf(d).map((it) => (
                              <li key={it.id} className="flex flex-wrap items-center gap-2 px-3 py-2">
                                <span
                                  className={`text-xs font-semibold px-2 py-0.5 rounded-sm ${reasonStyle(
                                    it.reason || "기본"
                                  )}`}
                                >
                                  {it.reason || "기본"}
                                </span>
                                <span className="flex-1 basis-48 min-w-0 text-sm break-words">{it.name}</span>
                                {MODE_LABEL[it.priceMode] && (
                                  <span className="text-xs text-slate-400">{MODE_LABEL[it.priceMode]}</span>
                                )}
                                <label className="flex items-center gap-1 text-xs text-slate-500 whitespace-nowrap">
                                  수량
                                  <input
                                    type="number"
                                    min="1"
                                    step="1"
                                    inputMode="numeric"
                                    aria-label="상품 수량"
                                    value={it.qty || 1}
                                    onChange={(e) => updateItem(d.id, it.id, { qty: Math.max(1, parseInt(e.target.value, 10) || 1) })}
                                    className="w-16 border border-slate-300 px-1.5 py-1 text-base text-center focus:outline-none focus:border-slate-900"
                                  />
                                </label>
                                <input
                                  type="text"
                                  inputMode="numeric"
                                  aria-label="단가"
                                  value={fmtNum(it.price)}
                                  onChange={(e) =>
                                    updateItem(d.id, it.id, { price: toNum(e.target.value), priceMode: "manual" })
                                  }
                                  placeholder="0"
                                  className="w-32 sm:w-36 border border-slate-300 px-2 py-1.5 text-base font-mono text-right focus:outline-none focus:border-slate-900"
                                />
                              </li>
                            ))}
                          </ul>
                        </div>
                        <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                          <Field label="설치비 (단가 합계)">
                            <div className="w-full border border-slate-200 bg-slate-50 px-2 py-1.5 text-base font-mono text-right text-slate-700">
                              {formatPrice(d.install)}
                            </div>
                          </Field>
                          <Field label="양중비">
                            <input
                              type="text"
                              inputMode="numeric"
                              value={fmtNum(d.lift)}
                              onChange={(e) => updateDelivery(d.id, { lift: toNum(e.target.value) })}
                              placeholder="0"
                              className={moneyCls}
                            />
                          </Field>
                          <Field label="내림비">
                            <input
                              type="text"
                              inputMode="numeric"
                              value={fmtNum(d.unload)}
                              onChange={(e) => updateDelivery(d.id, { unload: toNum(e.target.value) })}
                              placeholder="0"
                              className={moneyCls}
                            />
                          </Field>
                          <Field label="기타비용">
                            <input
                              type="text"
                              inputMode="numeric"
                              value={fmtNum(d.etc)}
                              onChange={(e) => updateDelivery(d.id, { etc: toNum(e.target.value) })}
                              placeholder="0"
                              className={moneyCls}
                            />
                          </Field>
                          <Field label="합계" className="col-span-2 sm:col-span-1">
                            <div className="w-full border border-slate-200 bg-slate-50 px-2 py-1.5 text-base font-mono text-right font-bold text-emerald-700">
                              {formatPrice(deliveryTotal(d))}
                            </div>
                          </Field>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })
            )}
          </div>
        )}
      </div>

      {/* 정산 기록 · 달력 */}
      <section className="mt-6 bg-white border border-slate-200 overflow-hidden">
        <div className="flex flex-wrap items-center gap-3 px-4 py-3">
          <button
            onClick={() => setRecordsOpen((v) => !v)}
            aria-expanded={recordsOpen}
            className="flex flex-1 min-w-0 flex-wrap items-center gap-2 text-left text-sm font-semibold text-slate-700 hover:text-slate-900"
          >
            <CalendarDays size={17} />
            <span>정산 기록</span>
            <span className="text-xs font-normal text-slate-400">선택된 배송 {deliveryRecords.reduce((sum, r) => sum + r.count, 0)}건 · {deliveryRecords.length}일</span>
            {recordsOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
          </button>
          <span className={"inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold " + (syncStatus === "supabase" ? "bg-emerald-50 text-emerald-700" : syncStatus === "saving" ? "bg-amber-50 text-amber-700" : "bg-slate-100 text-slate-600")} title="배송·상품 정보는 변경할 때마다 자동 저장되고 Supabase와 동기화됩니다.">
            <span className={"h-1.5 w-1.5 rounded-full " + (syncStatus === "supabase" ? "bg-emerald-500" : syncStatus === "saving" ? "bg-amber-500" : "bg-slate-400")} />
            {syncStatus === "supabase" ? "Supabase 저장" : syncStatus === "saving" ? "저장 중" : syncStatus === "connecting" ? "연결 중" : "브라우저 저장"}
          </span>
          <button
            onClick={exportAllExcel}
            title="완료·연기·취소를 포함한 전체 배송 내역 엑셀 다운로드"
            className="flex items-center gap-1 shrink-0 px-3 py-1.5 border border-slate-300 bg-white text-xs sm:text-sm font-semibold hover:bg-slate-100"
          >
            <Download size={14} /> 전체 엑셀
          </button>
        </div>
        {recordsOpen && (
          <div className="border-t border-slate-100 p-4 sm:p-5">
            <div className="flex flex-wrap items-center gap-2">
              <div className="inline-flex rounded-xl bg-slate-100 p-1">
                {[['day', '일별'], ['month', '월별'], ['year', '연별']].map(([mode, label]) => (
                  <button
                    key={mode}
                    onClick={() => setRecordsMode(mode)}
                    aria-pressed={recordsMode === mode}
                    className={"px-3 py-1.5 text-sm font-semibold " + (recordsMode === mode ? "bg-white text-slate-900 shadow-sm" : "text-slate-500 hover:text-slate-800")}
                  >
                    {label}
                  </button>
                ))}
              </div>
              <div className="flex-1" />
              <button onClick={() => shiftRecordsCursor(-1)} className="px-3 py-2 border border-slate-300 bg-white text-sm hover:bg-slate-100" aria-label="이전 기간">◀</button>
              <strong className="min-w-28 text-center text-sm sm:text-base text-slate-800">
                {recordsMode === "day" ? recordsDate : recordsMode === "month" ? recordsYear + "년 " + String(recordsMonth).padStart(2, "0") + "월" : recordsYear + "년"}
              </strong>
              <button onClick={() => shiftRecordsCursor(1)} className="px-3 py-2 border border-slate-300 bg-white text-sm hover:bg-slate-100" aria-label="다음 기간">▶</button>
              <button onClick={() => { const d = todayStr(); setRecordsDate(d); setDeliveryDate(d); }} className="px-3 py-2 border border-slate-300 bg-white text-sm font-semibold hover:bg-slate-100">오늘</button>
            </div>



            {recordsMode === "month" && (
              <div className="mt-4">
                <div className="mb-3 flex flex-wrap items-center justify-between gap-2 rounded-xl bg-indigo-50 px-4 py-3">
                  <span className="text-sm font-semibold text-slate-700">{recordsYear}년 {recordsMonth}월 정산</span>
                  <div className="flex items-center gap-3">
                    <span className="text-sm text-slate-600">{monthlyRecords.reduce((sum, r) => sum + r.count, 0)}건 · <b className="font-mono text-indigo-700">{formatPrice(monthlyRecords.reduce((sum, r) => sum + r.total, 0))}</b></span>
                    <button onClick={() => exportRecordsPeriod("month")} disabled={!recordsMonthRows.length} aria-label="월별 엑셀 다운로드" title="월별 전체 배송 엑셀 다운로드 (완료·연기·취소 포함)" className="inline-flex h-9 w-9 items-center justify-center rounded-lg border border-emerald-200 bg-white text-emerald-700 shadow-sm hover:bg-emerald-50 disabled:opacity-40">
                      <ExcelFileIcon size={17} />
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-7 gap-1 text-center">
                  {["일", "월", "화", "수", "목", "금", "토"].map((day) => <div key={day} className="py-2 text-xs font-semibold text-slate-500">{day}</div>)}
                  {calendarCells.map((date, index) => {
                    if (!date) return <div key={"blank-" + index} className="min-h-20 sm:min-h-24 rounded-xl bg-slate-50/70" />;
                    const record = recordsByDate.get(date);
                    const selected = date === recordsDate;
                    return (
                      <button
                        key={date}
                        onClick={() => { setRecordsDate(date); setDeliveryDate(date); }}
                        aria-pressed={selected}
                        className={"min-h-20 sm:min-h-24 rounded-xl border p-1.5 sm:p-2 text-left transition " + (selected ? "border-indigo-500 bg-indigo-50 ring-2 ring-indigo-100" : record ? "border-indigo-100 bg-white hover:bg-indigo-50" : "border-slate-100 bg-white hover:bg-slate-50")}
                      >
                        <span className={"text-xs sm:text-sm font-semibold " + (selected ? "text-indigo-700" : "text-slate-700")}>{Number(date.slice(8, 10))}</span>
                        {record ? (
                          <span className="mt-1 block">
                            <span className="block text-[10px] sm:text-xs text-slate-500">{record.count}건</span>
                            <span className="block truncate font-mono text-[10px] sm:text-xs font-semibold text-emerald-700">{formatPrice(record.total)}</span>
                          </span>
                        ) : <span className="mt-1 hidden text-[10px] text-slate-300 sm:block">내역 없음</span>}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {recordsMode === "year" && (
              <div className="mt-4">
                <div className="mb-3 flex items-center justify-between rounded-xl bg-indigo-50 px-4 py-3">
                  <span className="text-sm font-semibold text-slate-700">{recordsYear}년 전체</span>
                  <div className="flex items-center gap-3">
                    <span className="text-sm text-slate-600">{yearlyRecords.reduce((sum, r) => sum + r.count, 0)}건 · <b className="font-mono text-indigo-700">{formatPrice(yearlyRecords.reduce((sum, r) => sum + r.total, 0))}</b></span>
                    <button onClick={() => exportRecordsPeriod("year")} disabled={!recordsYearRows.length} aria-label="연별 엑셀 다운로드" title="연별 전체 배송 엑셀 다운로드 (완료·연기·취소 포함)" className="inline-flex h-9 w-9 items-center justify-center rounded-lg border border-emerald-200 bg-white text-emerald-700 shadow-sm hover:bg-emerald-50 disabled:opacity-40">
                      <ExcelFileIcon size={17} />
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-2 sm:grid-cols-3 lg:grid-cols-4">
                  {annualMonths.map((month) => (
                    <button
                      key={month.prefix}
                      onClick={() => { const d = month.prefix + "-01"; setRecordsDate(d); setDeliveryDate(d); setRecordsMode("month"); }}
                      className={"rounded-xl border p-3 text-left transition " + (month.days ? "border-indigo-100 bg-white hover:border-indigo-300 hover:bg-indigo-50" : "border-slate-100 bg-slate-50 text-slate-400")}
                    >
                      <span className="block font-semibold">{month.month}월</span>
                      <span className="mt-1 block text-xs text-slate-500">{month.count}건 · {month.days}일</span>
                      <span className="mt-1 block truncate font-mono text-sm font-semibold text-emerald-700">{formatPrice(month.total)}</span>
                    </button>
                  ))}
                </div>
              </div>
            )}

            {recordsMode === "day" && (
              <div className="mt-4 rounded-xl border border-slate-200 bg-slate-50 p-4">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <strong className="text-sm text-slate-800">{shortDate(recordsDate)} 배송 내역</strong>
                  <span className="text-sm text-slate-500">{recordsDayRows.length ? recordsDayRows.length + "건 · " + formatPrice(recordsDayRows.reduce((sum, d) => sum + deliveryTotal(d), 0)) : "배송 내역 없음"}</span>
                </div>
              </div>
            )}

            <div className="mt-4 rounded-xl border border-slate-200 bg-white">
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 px-4 py-3">
                <strong className="text-sm text-slate-800">{shortDate(recordsDate)} 상세 내역</strong>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-slate-500">{recordsDayRows.length}건</span>
                  <button onClick={() => exportRecordsPeriod("day")} disabled={!recordsDayRows.length} aria-label="일별 엑셀 다운로드" title="일별 전체 배송 엑셀 다운로드 (완료·연기·취소 포함)" className="inline-flex h-9 w-9 items-center justify-center rounded-lg border border-emerald-200 bg-white text-emerald-700 shadow-sm hover:bg-emerald-50 disabled:opacity-40">
                    <ExcelFileIcon size={17} />
                  </button>
                </div>
              </div>
              {!recordsDayRows.length ? (
                <p className="px-4 py-6 text-center text-sm text-slate-400">선택한 날짜에 배송 내역이 없어요.</p>
              ) : (
                <div className="divide-y divide-slate-100">
                  {recordsDayRows.map((d) => (
                    <div key={d.id} className="flex flex-wrap items-center gap-2 px-4 py-3">
                      <span className={"rounded-full px-2 py-1 text-xs font-semibold " + (STATUS_STYLE[d.status] || "bg-slate-100 text-slate-700")}>{d.status}</span>
                      <span className="min-w-0 flex-1 basis-48">
                        <span className="block truncate text-sm font-medium text-slate-800">{productLabel(d) || "상품 미입력"}</span>
                        <span className="block truncate text-xs text-slate-500">{d.address || "주소 미입력"}</span>
                      </span>
                      <span className="text-xs text-slate-500">설치 {fmt0(d.install)} · 양중 {fmt0(d.lift)} · 내림 {fmt0(d.unload)} · 기타 {fmt0(d.etc)}</span>
                      <span className="w-24 text-right font-mono text-sm font-semibold text-emerald-700">{formatPrice(deliveryTotal(d))}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}
      </section>
    </div>
  );

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <style>{PRINT_CSS}</style>
      {/* 헤더 */}
      <div className="no-print sticky top-0 z-20 bg-gradient-to-r from-slate-950 via-slate-900 to-slate-800 text-white px-5 py-5 sm:px-8 sm:py-7 shadow-lg">
        <div className={`${view === "lookup" ? "max-w-3xl" : "max-w-5xl"} mx-auto`}>
          <p className={"text-xs font-semibold tracking-widest mb-1 " + (view === "lookup" ? "text-indigo-300" : view === "delivery" ? "text-emerald-300" : "text-amber-300")}>{headTexts[view][0]}</p>
          <h1 className="text-xl sm:text-2xl font-bold">{headTexts[view][1]}</h1>
          <p className="text-slate-400 text-sm mt-1">{headTexts[view][2]}</p>
          <div className="mt-5 inline-flex flex-wrap items-center gap-1.5 rounded-2xl border border-white/10 bg-slate-950/40 p-1.5" role="tablist" aria-label="페이지 메뉴">
            <button
              type="button"
              role="tab"
              aria-selected={view === "lookup"}
              onClick={() => setView("lookup")}
              className={"inline-flex min-w-32 items-center justify-center gap-2 rounded-xl border px-4 py-2.5 text-sm font-semibold transition " + (view === "lookup" ? "border-indigo-400 bg-indigo-500 text-white shadow-lg shadow-indigo-950/20" : "border-indigo-300/25 bg-indigo-500/10 text-indigo-100 hover:bg-indigo-500/20")}
            >
              <span className="grid h-6 w-6 place-items-center rounded-full bg-black/10 text-[11px]">01</span>
              <span>단가 조회</span>
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={view === "delivery"}
              onClick={() => setView("delivery")}
              className={"inline-flex min-w-32 items-center justify-center gap-2 rounded-xl border px-4 py-2.5 text-sm font-semibold transition " + (view === "delivery" ? "border-emerald-400 bg-emerald-500 text-white shadow-lg shadow-emerald-950/20" : "border-emerald-300/25 bg-emerald-500/10 text-emerald-100 hover:bg-emerald-500/20")}
            >
              <span className="grid h-6 w-6 place-items-center rounded-full bg-black/10 text-[11px]">02</span>
              <span>배송 리스트</span>
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={view === "settle"}
              onClick={() => setView("settle")}
              className={"inline-flex min-w-32 items-center justify-center gap-2 rounded-xl border px-4 py-2.5 text-sm font-semibold transition " + (view === "settle" ? "border-amber-300 bg-amber-400 text-slate-950 shadow-lg shadow-amber-950/20" : "border-amber-200/25 bg-amber-400/10 text-amber-100 hover:bg-amber-400/20")}
            >
              <span className="grid h-6 w-6 place-items-center rounded-full bg-black/10 text-[11px]">03</span>
              <span>정산 페이지</span>
            </button>
          </div>
        </div>
      </div>

      <div className={`no-print ${view === "lookup" ? "max-w-3xl" : "max-w-5xl"} mx-auto px-5 sm:px-8 py-6`}>
        {view === "delivery" ? (
          deliveryView
        ) : view === "settle" ? (
          settleView
        ) : (
          <>
        {/* 검색창 */}
        <div className="relative">
          <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={20} />
          <input
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="예: 이즈 카이  또는  동서가구 이즈 카이 1단 3서랍 LED 수납침대..."
            className="w-full bg-white border-2 border-slate-900 rounded-none pl-12 pr-4 py-3 text-base focus:outline-none focus:border-amber-500"
          />
        </div>
        <div className="mt-2 flex flex-wrap items-center gap-2 text-xs text-slate-500">
          <span>수수료율</span>
          <input
            type="number"
            inputMode="decimal"
            min="0"
            max="100"
            step="0.1"
            value={feeRate}
            onChange={(e) => changeFeeRate(e.target.value)}
            className="w-16 border border-slate-300 px-2 py-1 text-base font-mono text-slate-900 focus:outline-none focus:border-slate-900"
          />
          <span>% · 수수료 제외 단가 = 원단가 × {Number((100 - feeRate).toFixed(2))}% (원 단위 반올림)</span>
        </div>

        {/* 결과 영역 */}
        <div className="mt-4">
          {!loaded ? (
            <p className="text-slate-400 text-sm py-6">불러오는 중...</p>
          ) : !query.trim() ? (
            <p className="text-slate-400 text-sm py-6">
              검색어를 입력해주세요. 등록된 상품 {products.length}건 중에서 찾아드려요.
            </p>
          ) : results.length === 0 ? (
            <p className="text-slate-500 text-sm py-6">
              일치하는 상품이 없어요. 상품명을 다시 확인하거나 아래 상품 관리에서 등록해주세요.
            </p>
          ) : (
            <div>
              <p className="text-xs text-slate-500 mb-2">
                상품 {groupedResults.length}건 · 단가 {results.length}건 검색됨
              </p>
              <div className="border border-slate-200 divide-y divide-slate-200 bg-white">
                {groupedResults.map((g) => (
                  <div key={g.name} className="px-4 py-3">
                    <div className="text-sm leading-snug mb-2">
                      <Highlighted text={g.name} tokens={tokens} />
                    </div>
                    <div className="flex flex-col gap-2">
                      {g.reasonGroups.map((rg) => {
                        const cur = rg.current;
                        const histKey = g.name + "|" + rg.reason;
                        const isHistOpen = expandedHistoryKeys.has(histKey);
                        return (
                          <div key={rg.reason} className="flex flex-col gap-1">
                            <div className="flex items-center flex-wrap gap-2">
                              <div className="flex items-center gap-2 border border-slate-200 bg-slate-50 pl-2 pr-3 py-1">
                                <span
                                  className={`text-xs font-semibold px-2 py-0.5 whitespace-nowrap rounded-sm ${reasonStyle(
                                    rg.reason
                                  )}`}
                                >
                                  {rg.reason || "기본"}
                                </span>
                                <span className="flex flex-col leading-tight">
                                  {cur && hasNum(cur.price) && cur.price > 0 && (
                                    <>
                                      <span className="flex items-baseline justify-between gap-2 whitespace-nowrap">
                                        <span className="text-xs text-slate-400">원단가</span>
                                        <span className="font-mono text-emerald-700 font-semibold text-sm">
                                          {formatPrice(cur.price)}
                                        </span>
                                      </span>
                                      <span className="flex items-baseline justify-between gap-2 whitespace-nowrap">
                                        <span className="text-xs text-slate-400">수수료 제외</span>
                                        <span className="font-mono text-blue-700 font-semibold text-sm">
                                          {formatPrice(netOf(cur.price, feeRate))}
                                        </span>
                                      </span>
                                    </>
                                  )}
                                  {cur && hasNum(cur.applied) && (
                                    <span className="flex items-baseline justify-between gap-2 whitespace-nowrap">
                                      <span className="text-xs text-slate-400">적용단가</span>
                                      <span className="font-mono text-purple-700 font-semibold text-sm">
                                        {formatPrice(cur.applied)}
                                      </span>
                                    </span>
                                  )}
                                  {cur && !(hasNum(cur.price) && cur.price > 0) && !hasNum(cur.applied) && (
                                    <span className="text-xs text-slate-400">단가 없음</span>
                                  )}
                                </span>
                                {cur && cur.effectiveDate && (
                                  <span className="text-xs text-slate-400 whitespace-nowrap">{cur.effectiveDate}~</span>
                                )}
                              </div>
                              {rg.history.length > 0 && (
                                <button
                                  onClick={() => toggleHistory(histKey)}
                                  className="text-xs text-slate-400 underline hover:text-slate-600"
                                >
                                  이전 단가 {rg.history.length}건 {isHistOpen ? "숨기기" : "보기"}
                                </button>
                              )}
                            </div>
                            {isHistOpen && (
                              <div className="ml-1 pl-3 border-l-2 border-slate-200 flex flex-col gap-1">
                                {rg.history.map((h) => (
                                  <div key={h.id} className="flex items-center gap-2 text-xs text-slate-500 w-fit">
                                    <span className="whitespace-nowrap">{h.effectiveDate || "날짜 없음"}</span>
                                    {hasNum(h.price) && h.price > 0 && (
                                      <>
                                        <span className="font-mono">원 {formatPrice(h.price)}</span>
                                        <span className="font-mono text-blue-700">
                                          제외 {formatPrice(netOf(h.price, feeRate))}
                                        </span>
                                      </>
                                    )}
                                    {hasNum(h.applied) && (
                                      <span className="font-mono text-purple-700">적용 {formatPrice(h.applied)}</span>
                                    )}
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* 상품 관리 토글 */}
        <div className="mt-8 border-t border-slate-200 pt-4">
          <button
            onClick={() => setManageOpen((v) => !v)}
            className="flex items-center gap-1 text-sm font-semibold text-slate-700 hover:text-slate-900"
          >
            상품 관리 ({products.length}건)
            {manageOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
          </button>

          {manageOpen && (
            <div className="mt-4 space-y-6">
              {/* 엑셀 업로드 */}
              <div className="bg-white border border-slate-200 p-4">
                <p className="text-sm font-semibold mb-2">엑셀/CSV로 한번에 등록</p>
                <p className="text-xs text-slate-500 mb-3">
                  '상품명/품명', '원단가(또는 단가)', '적용단가' 열은 자동 인식해요. '구분/사유', '적용일' 열이 있으면
                  재출고·프로모션·월별 단가변경도 같이 불러와요. 없으면 전부 '기본'·오늘 날짜로 등록돼요.
                </p>
                <label className="inline-flex items-center gap-2 cursor-pointer bg-slate-900 text-white text-sm px-4 py-2 hover:bg-slate-800">
                  <Upload size={16} />
                  파일 선택
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".csv,.xlsx,.xls"
                    onChange={handleFile}
                    className="hidden"
                  />
                </label>
                {uploadMsg && <p className="text-xs text-slate-600 mt-2">{uploadMsg}</p>}
              </div>

              {/* 수동 추가 */}
              <div className="bg-white border border-slate-200 p-4">
                <p className="text-sm font-semibold mb-1">상품 직접 추가</p>
                <p className="text-xs text-slate-500 mb-3">
                  같은 상품·같은 구분이라도 적용일을 다르게 넣으면 단가 변경 이력으로 따로 쌓여요. 검색에는 오늘
                  기준 최신 단가만 크게 보여요.
                </p>
                <div className="flex flex-col sm:flex-row gap-2">
                  <input
                    type="text"
                    value={newName}
                    onChange={(e) => setNewName(e.target.value)}
                    placeholder="상품명 (전체 옵션명 그대로)"
                    className="flex-1 border border-slate-300 px-3 py-2 text-base focus:outline-none focus:border-slate-900"
                  />
                  <input
                    type="text"
                    inputMode="numeric"
                    value={newPrice}
                    onChange={(e) => setNewPrice(e.target.value)}
                    placeholder="원단가 (모르면 비워둠)"
                    className="sm:w-36 border border-slate-300 px-3 py-2 text-base font-mono focus:outline-none focus:border-slate-900"
                  />
                  <input
                    type="text"
                    inputMode="numeric"
                    value={newApplied}
                    onChange={(e) => setNewApplied(e.target.value)}
                    placeholder="적용단가"
                    className="sm:w-28 border border-slate-300 px-3 py-2 text-base font-mono focus:outline-none focus:border-slate-900"
                  />
                  <input
                    type="date"
                    value={newEffectiveDate}
                    onChange={(e) => setNewEffectiveDate(e.target.value)}
                    title="적용일 (이 날짜부터 이 단가가 적용돼요)"
                    className="sm:w-36 border border-slate-300 px-2 py-2 text-base focus:outline-none focus:border-slate-900"
                  />
                  {!addingCustomReason ? (
                    <select
                      value={newReason}
                      onChange={(e) => {
                        if (e.target.value === "__custom__") {
                          setAddingCustomReason(true);
                          setCustomReasonInput("");
                        } else {
                          setNewReason(e.target.value);
                        }
                      }}
                      className="sm:w-32 border border-slate-300 px-2 py-2 text-base focus:outline-none focus:border-slate-900 bg-white"
                    >
                      {allReasons.map((r) => (
                        <option key={r} value={r}>
                          {r}
                        </option>
                      ))}
                      <option value="__custom__">+ 새 구분 추가</option>
                    </select>
                  ) : (
                    <div className="flex gap-1 sm:w-40">
                      <input
                        type="text"
                        value={customReasonInput}
                        onChange={(e) => setCustomReasonInput(e.target.value)}
                        placeholder="새 구분명"
                        className="flex-1 border border-slate-300 px-2 py-2 text-base focus:outline-none focus:border-slate-900"
                      />
                      <button
                        onClick={() => {
                          setAddingCustomReason(false);
                          setNewReason("기본");
                        }}
                        className="px-2 text-slate-400 hover:text-slate-700"
                        title="취소"
                      >
                        <X size={16} />
                      </button>
                    </div>
                  )}
                  <button
                    onClick={handleAdd}
                    className="flex items-center justify-center gap-1 bg-amber-500 text-slate-900 font-semibold text-sm px-4 py-2 hover:bg-amber-400"
                  >
                    <Plus size={16} />
                    추가
                  </button>
                </div>
              </div>

              {/* 등록 목록 */}
              <div className="bg-white border border-slate-200">
                <div className="flex items-center justify-between px-4 py-3 border-b border-slate-200">
                  <p className="text-sm font-semibold">등록된 상품 목록</p>
                  {products.length > 0 &&
                    (!confirmResetAll ? (
                      <button onClick={() => setConfirmResetAll(true)} className="text-xs text-red-600 hover:underline">
                        전체 삭제
                      </button>
                    ) : (
                      <span className="flex items-center gap-2 text-xs">
                        <span className="text-slate-500">정말요?</span>
                        <button onClick={handleResetAll} className="text-red-600 font-semibold hover:underline">
                          삭제
                        </button>
                        <button onClick={() => setConfirmResetAll(false)} className="text-slate-400 hover:underline">
                          취소
                        </button>
                      </span>
                    ))}
                </div>
                {products.length > 0 && (
                  <div className="px-4 py-2 border-b border-slate-200 relative">
                    <Search size={14} className="absolute left-6 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      value={manageQuery}
                      onChange={(e) => setManageQuery(e.target.value)}
                      placeholder="목록에서 상품 찾기 (수정할 상품 검색)"
                      className="w-full border border-slate-300 pl-7 pr-2 py-1.5 text-base focus:outline-none focus:border-slate-900"
                    />
                    {manageQuery.trim() && (
                      <p className="text-xs text-slate-400 mt-1">{filteredManageProducts.length}건 일치</p>
                    )}
                  </div>
                )}
                {saveError && (
                  <p className="text-xs text-red-600 px-4 py-2">
                    저장 중 문제가 생겼어요. 새로고침 후 다시 시도해주세요.
                  </p>
                )}
                {products.length === 0 ? (
                  <p className="text-sm text-slate-400 px-4 py-6">등록된 상품이 없어요.</p>
                ) : filteredManageProducts.length === 0 ? (
                  <p className="text-sm text-slate-400 px-4 py-6">검색어와 일치하는 상품이 없어요.</p>
                ) : (
                  <div className="divide-y divide-slate-100 max-h-96 overflow-y-auto">
                    {filteredManageProducts.map((p) => (
                      <div key={p.id} className="flex flex-wrap items-center gap-2 px-4 py-2">
                        {editingId === p.id ? (
                          <>
                            <input
                              type="text"
                              value={editDraft.name}
                              onChange={(e) => setEditDraft((d) => ({ ...d, name: e.target.value }))}
                              className="w-full sm:flex-1 border border-slate-300 px-2 py-1.5 text-base focus:outline-none focus:border-slate-900"
                            />
                            <input
                              type="text"
                              inputMode="numeric"
                              value={editDraft.price}
                              onChange={(e) => setEditDraft((d) => ({ ...d, price: e.target.value }))}
                              placeholder="원단가"
                              className="w-24 border border-slate-300 px-2 py-1.5 text-base font-mono focus:outline-none focus:border-slate-900"
                            />
                            <input
                              type="text"
                              inputMode="numeric"
                              value={editDraft.applied}
                              onChange={(e) => setEditDraft((d) => ({ ...d, applied: e.target.value }))}
                              placeholder="적용단가"
                              className="w-24 border border-slate-300 px-2 py-1.5 text-base font-mono focus:outline-none focus:border-slate-900"
                            />
                            <input
                              type="date"
                              value={editDraft.effectiveDate}
                              onChange={(e) => setEditDraft((d) => ({ ...d, effectiveDate: e.target.value }))}
                              className="w-36 border border-slate-300 px-2 py-1.5 text-base focus:outline-none focus:border-slate-900"
                            />
                            <input
                              list="reason-options"
                              value={editDraft.reason}
                              onChange={(e) => setEditDraft((d) => ({ ...d, reason: e.target.value }))}
                              className="w-24 border border-slate-300 px-2 py-1.5 text-base focus:outline-none focus:border-slate-900"
                            />
                            <button
                              onClick={() => saveEdit(p.id)}
                              className="p-2 text-emerald-700 hover:bg-emerald-50"
                              title="저장"
                            >
                              <Check size={16} />
                            </button>
                            <button onClick={cancelEdit} className="p-2 text-slate-400 hover:bg-slate-100" title="취소">
                              <X size={16} />
                            </button>
                          </>
                        ) : (
                          <>
                            <div className="flex-1 basis-40 text-sm truncate">{p.name}</div>
                            <span
                              className={`text-xs font-semibold px-2 py-0.5 whitespace-nowrap rounded-sm ${reasonStyle(
                                p.reason
                              )}`}
                            >
                              {p.reason || "기본"}
                            </span>
                            <span className="text-xs text-slate-400 font-mono whitespace-nowrap">
                              {p.effectiveDate || "-"}
                            </span>
                            <div className="font-mono text-emerald-700 text-sm whitespace-nowrap">
                              {hasNum(p.price) && p.price > 0 ? `원 ${formatPrice(p.price)}` : "원단가 -"}
                            </div>
                            {hasNum(p.applied) && (
                              <div className="font-mono text-purple-700 text-sm whitespace-nowrap">
                                적용 {formatPrice(p.applied)}
                              </div>
                            )}
                            <button
                              onClick={() => startEdit(p)}
                              className="p-2 text-slate-500 hover:bg-slate-100"
                              title="수정"
                            >
                              <Pencil size={14} />
                            </button>
                            {confirmDeleteProductId !== p.id ? (
                              <button
                                onClick={() => setConfirmDeleteProductId(p.id)}
                                className="p-2 text-slate-500 hover:bg-red-50 hover:text-red-600"
                                title="삭제"
                              >
                                <Trash2 size={14} />
                              </button>
                            ) : (
                              <span className="flex items-center gap-1 text-xs shrink-0">
                                <button
                                  onClick={() => handleDelete(p.id)}
                                  className="text-red-600 font-semibold hover:underline px-1"
                                >
                                  삭제
                                </button>
                                <button
                                  onClick={() => setConfirmDeleteProductId(null)}
                                  className="text-slate-400 hover:underline px-1"
                                >
                                  취소
                                </button>
                              </span>
                            )}
                          </>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
              <datalist id="reason-options">
                {allReasons.map((r) => (
                  <option key={r} value={r} />
                ))}
              </datalist>
            </div>
          )}
        </div>

          </>
        )}
      </div>

      {view === "delivery" && printView}
      {view === "settle" && settlePrintView}

    </div>
  );
}

