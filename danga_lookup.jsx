import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { Search, Upload, Trash2, Pencil, Plus, X, ChevronDown, ChevronUp, Check, Save, CalendarDays } from "lucide-react";

const PRODUCTS_KEY = "products";
const REASONS_KEY = "custom_reasons";
const SETTLEMENTS_KEY = "settlement_records";
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
  return new Date().toISOString().slice(0, 10);
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
  const [query, setQuery] = useState("");
  const [manageOpen, setManageOpen] = useState(false);

  const [newName, setNewName] = useState("");
  const [newPrice, setNewPrice] = useState("");
  const [newReason, setNewReason] = useState("기본");
  const [newEffectiveDate, setNewEffectiveDate] = useState(() => todayStr());
  const [addingCustomReason, setAddingCustomReason] = useState(false);
  const [customReasonInput, setCustomReasonInput] = useState("");

  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState({ name: "", price: "", reason: "기본", effectiveDate: "" });
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

  const [selection, setSelection] = useState([]);
  const [cartOpen, setCartOpen] = useState(false);
  const [showDateInput, setShowDateInput] = useState(false);
  const [saveDate, setSaveDate] = useState(() => new Date().toISOString().slice(0, 10));

  const [settlements, setSettlements] = useState([]);
  const [recordsOpen, setRecordsOpen] = useState(false);
  const [expandedRecordId, setExpandedRecordId] = useState(null);

  const [confirmClear, setConfirmClear] = useState(false);
  const [confirmResetAll, setConfirmResetAll] = useState(false);
  const [confirmDeleteProductId, setConfirmDeleteProductId] = useState(null);
  const [confirmDeleteRecordId, setConfirmDeleteRecordId] = useState(null);

  function addToSelection(v) {
    setCartOpen(true);
    setSelection((prev) => {
      const idx = prev.findIndex((item) => item.id === v.id);
      if (idx > -1) {
        const next = [...prev];
        next[idx] = { ...next[idx], qty: next[idx].qty + 1 };
        return next;
      }
      return [...prev, { id: v.id, name: v.name, reason: v.reason, price: v.price, qty: 1 }];
    });
  }

  function incrementQty(id) {
    setSelection((prev) => prev.map((item) => (item.id === id ? { ...item, qty: item.qty + 1 } : item)));
  }

  function decrementQty(id) {
    setSelection((prev) =>
      prev
        .map((item) => (item.id === id ? { ...item, qty: item.qty - 1 } : item))
        .filter((item) => item.qty > 0)
    );
  }

  function removeFromSelection(id) {
    setSelection((prev) => prev.filter((item) => item.id !== id));
  }

  function clearSelection() {
    setSelection([]);
    setConfirmClear(false);
  }

  async function persistSettlements(next) {
    setSettlements(next);
    try {
      await window.storage.set(SETTLEMENTS_KEY, JSON.stringify(next), false);
    } catch (e) {
      console.error(e);
    }
  }

  function handleSaveSettlement() {
    if (!selection.length || !saveDate) return;
    const existing = settlements.find((r) => r.date === saveDate);

    if (existing) {
      const mergedItems = [...existing.items];
      selection.forEach((item) => {
        const idx = mergedItems.findIndex((m) => m.id === item.id);
        if (idx > -1) {
          mergedItems[idx] = { ...mergedItems[idx], qty: mergedItems[idx].qty + item.qty };
        } else {
          mergedItems.push(item);
        }
      });
      const mergedTotalCount = mergedItems.reduce((sum, i) => sum + i.qty, 0);
      const mergedTotalPrice = mergedItems.reduce((sum, i) => sum + i.price * i.qty, 0);
      const updated = {
        ...existing,
        items: mergedItems,
        totalCount: mergedTotalCount,
        totalPrice: mergedTotalPrice,
        savedAt: new Date().toISOString(),
      };
      persistSettlements(settlements.map((r) => (r.id === existing.id ? updated : r)));
    } else {
      const record = {
        id: Date.now() + "-" + Math.random().toString(36).slice(2),
        date: saveDate,
        items: selection,
        totalCount,
        totalPrice,
        savedAt: new Date().toISOString(),
      };
      persistSettlements([record, ...settlements]);
    }

    setSelection([]);
    setShowDateInput(false);
    setCartOpen(false);
    setRecordsOpen(true);
  }

  function handleDeleteRecord(id) {
    persistSettlements(settlements.filter((r) => r.id !== id));
    setConfirmDeleteRecordId(null);
  }

  const selectionMap = useMemo(() => {
    const m = new Map();
    selection.forEach((item) => m.set(item.id, item.qty));
    return m;
  }, [selection]);

  const totalCount = useMemo(() => selection.reduce((sum, i) => sum + i.qty, 0), [selection]);
  const totalPrice = useMemo(() => selection.reduce((sum, i) => sum + i.price * i.qty, 0), [selection]);

  const allReasons = useMemo(() => {
    const set = new Set([...DEFAULT_REASONS, ...customReasons]);
    return Array.from(set);
  }, [customReasons]);

  const sortedProducts = useMemo(
    () => [...products].sort((a, b) => a.name.localeCompare(b.name) || (a.reason || "").localeCompare(b.reason || "") || (b.effectiveDate || "").localeCompare(a.effectiveDate || "")),
    [products]
  );

  useEffect(() => {
    (async () => {
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
      try {
        const res3 = await window.storage.get(SETTLEMENTS_KEY, false);
        if (res3 && res3.value) {
          const loaded = JSON.parse(res3.value);
          const byDate = new Map();
          loaded.forEach((r) => {
            if (!byDate.has(r.date)) {
              byDate.set(r.date, { ...r, items: [...r.items] });
            } else {
              const acc = byDate.get(r.date);
              r.items.forEach((item) => {
                const idx = acc.items.findIndex((m) => m.id === item.id);
                if (idx > -1) acc.items[idx] = { ...acc.items[idx], qty: acc.items[idx].qty + item.qty };
                else acc.items.push(item);
              });
            }
          });
          const merged = Array.from(byDate.values()).map((r) => ({
            ...r,
            totalCount: r.items.reduce((s, i) => s + i.qty, 0),
            totalPrice: r.items.reduce((s, i) => s + i.price * i.qty, 0),
          }));
          setSettlements(merged);
          if (merged.length !== loaded.length) {
            window.storage.set(SETTLEMENTS_KEY, JSON.stringify(merged), false).catch(() => {});
          }
        }
      } catch (e) {
        // 최초 사용, 무시
      } finally {
        setLoaded(true);
      }
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
      if (!nameMap.has(r.name)) nameMap.set(r.name, new Map());
      const reasonMap = nameMap.get(r.name);
      const key = r.reason || "기본";
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
    const price = Number(String(newPrice).replace(/[^0-9.-]/g, ""));
    const reason = (addingCustomReason ? customReasonInput.trim() : newReason) || "기본";
    const effectiveDate = newEffectiveDate || todayStr();
    if (!name || !price) return;
    if (addingCustomReason && reason) addCustomReasonToList(reason);
    const next = [
      ...products,
      { id: Date.now() + "-" + Math.random().toString(36).slice(2), name, price, reason, effectiveDate },
    ];
    persistProducts(next);
    setNewName("");
    setNewPrice("");
    setAddingCustomReason(false);
    setCustomReasonInput("");
  }

  function startEdit(p) {
    setEditingId(p.id);
    setEditDraft({ name: p.name, price: String(p.price), reason: p.reason || "기본", effectiveDate: p.effectiveDate || todayStr() });
  }

  function cancelEdit() {
    setEditingId(null);
    setEditDraft({ name: "", price: "", reason: "기본", effectiveDate: "" });
  }

  function saveEdit(id) {
    const name = editDraft.name.trim();
    const price = Number(String(editDraft.price).replace(/[^0-9.-]/g, ""));
    const reason = editDraft.reason.trim() || "기본";
    const effectiveDate = editDraft.effectiveDate || todayStr();
    if (!name || !price) return;
    addCustomReasonToList(reason);
    const next = products.map((p) => (p.id === id ? { ...p, name, price, reason, effectiveDate } : p));
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
          reasonCol = -1,
          dateCol = -1;
        for (let i = 0; i < Math.min(5, rows.length); i++) {
          const row = (rows[i] || []).map((c) => String(c));
          const nIdx = row.findIndex((c) => /상품|품명|제품/.test(c));
          const pIdx = row.findIndex((c) => /단가|가격/.test(c));
          const rIdx = row.findIndex((c) => /구분|사유|타입|유형/.test(c));
          const dIdx = row.findIndex((c) => /적용일|적용월|기준일|날짜/.test(c));
          if (nIdx > -1 || pIdx > -1 || rIdx > -1 || dIdx > -1) {
            headerIdx = i;
            if (nIdx > -1) nameCol = nIdx;
            if (pIdx > -1) priceCol = pIdx;
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
          const price = Number(String(row[priceCol] ?? "").replace(/[^0-9.-]/g, ""));
          const reasonRaw = reasonCol > -1 ? String(row[reasonCol] ?? "").trim() : "";
          const reason = reasonRaw || "기본";
          const effectiveDate = dateCol > -1 ? excelCellToDateStr(row[dateCol]) || todayStr() : todayStr();
          if (!name || !price) continue;
          if (!DEFAULT_REASONS.includes(reason) && !newCustomReasons.includes(reason)) {
            newCustomReasons.push(reason);
          }
          imported.push({
            id: Date.now() + "-" + i + "-" + Math.random().toString(36).slice(2),
            name,
            price,
            reason,
            effectiveDate,
          });
        }

        if (imported.length === 0) {
          setUploadMsg(
            "불러올 수 있는 데이터가 없어요. '상품명/품명', '단가/가격' 열이 있는지 확인해주세요. '구분/사유', '적용일' 열이 있으면 자동 인식해요."
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

  return (
    <div className={`min-h-screen bg-slate-50 text-slate-900 ${selection.length > 0 ? "pb-20" : ""}`}>
      {/* 헤더 */}
      <div className="bg-slate-900 text-white px-5 py-6 sm:px-8 sm:py-8">
        <div className="max-w-3xl mx-auto">
          <p className="text-amber-400 text-xs font-semibold tracking-widest mb-1">가구 배송 단가 조회</p>
          <h1 className="text-xl sm:text-2xl font-bold">상품명 또는 키워드를 입력하면 단가를 보여드려요</h1>
          <p className="text-slate-400 text-sm mt-1">
            같은 상품도 재출고·프로모션·내림서비스 등 사유별로 단가가 다르면 전부 따로 보여줘요.
          </p>
        </div>
      </div>

      <div className="max-w-3xl mx-auto px-5 sm:px-8 py-6">
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
                        const qty = cur ? selectionMap.get(cur.id) || 0 : 0;
                        const histKey = g.name + "|" + rg.reason;
                        const isHistOpen = expandedHistoryKeys.has(histKey);
                        return (
                          <div key={rg.reason} className="flex flex-col gap-1">
                            <div className="flex items-center flex-wrap gap-2">
                              <button
                                onClick={() => cur && addToSelection(cur)}
                                className={`flex items-center gap-2 border pl-2 pr-3 py-1 transition-colors ${
                                  qty > 0
                                    ? "border-amber-500 bg-amber-50"
                                    : "border-slate-200 bg-slate-50 hover:bg-slate-100"
                                }`}
                                title="클릭하면 정산 목록에 추가돼요"
                              >
                                <span
                                  className={`text-xs font-semibold px-2 py-0.5 whitespace-nowrap rounded-sm ${reasonStyle(
                                    rg.reason
                                  )}`}
                                >
                                  {rg.reason || "기본"}
                                </span>
                                <span className="font-mono text-emerald-700 font-semibold text-sm whitespace-nowrap">
                                  {cur ? formatPrice(cur.price) : "-"}
                                </span>
                                {cur && cur.effectiveDate && (
                                  <span className="text-xs text-slate-400 whitespace-nowrap">{cur.effectiveDate}~</span>
                                )}
                                {qty > 0 && (
                                  <span className="bg-amber-500 text-slate-900 text-xs font-bold rounded-full w-5 h-5 flex items-center justify-center">
                                    {qty}
                                  </span>
                                )}
                              </button>
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
                                  <button
                                    key={h.id}
                                    onClick={() => addToSelection(h)}
                                    className="flex items-center gap-2 text-xs text-slate-500 hover:text-slate-700 w-fit"
                                    title="클릭하면 정산 목록에 추가돼요"
                                  >
                                    <span className="whitespace-nowrap">{h.effectiveDate || "날짜 없음"}</span>
                                    <span className="font-mono">{formatPrice(h.price)}</span>
                                  </button>
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
                  '상품명/품명', '단가/가격' 열은 자동 인식해요. '구분/사유' 열이 있으면 재출고·프로모션 등도 같이
                  불러와요. 없으면 전부 '기본'으로 등록돼요.
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
                    placeholder="단가 (숫자)"
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
                {saveError && (
                  <p className="text-xs text-red-600 px-4 py-2">
                    저장 중 문제가 생겼어요. 새로고침 후 다시 시도해주세요.
                  </p>
                )}
                {products.length === 0 ? (
                  <p className="text-sm text-slate-400 px-4 py-6">등록된 상품이 없어요.</p>
                ) : (
                  <div className="divide-y divide-slate-100 max-h-96 overflow-y-auto">
                    {sortedProducts.map((p) => (
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
                              {formatPrice(p.price)}
                            </div>
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

        {/* 정산 기록 */}
        <div className="mt-4 border-t border-slate-200 pt-4">
          <button
            onClick={() => setRecordsOpen((v) => !v)}
            className="flex items-center gap-1 text-sm font-semibold text-slate-700 hover:text-slate-900"
          >
            <CalendarDays size={16} />
            정산 기록 ({settlements.length}건)
            {recordsOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
          </button>

          {recordsOpen && (
            <div className="mt-4 bg-white border border-slate-200">
              {settlements.length === 0 ? (
                <p className="text-sm text-slate-400 px-4 py-6">
                  저장된 정산 기록이 없어요. 정산 목록에서 "저장"을 누르면 날짜별로 여기 쌓여요.
                </p>
              ) : (
                <div className="divide-y divide-slate-100">
                  {[...settlements]
                    .sort((a, b) => (a.date < b.date ? 1 : -1))
                    .map((r) => (
                      <div key={r.id}>
                        <div
                          className="flex items-center gap-3 px-4 py-3 cursor-pointer hover:bg-slate-50"
                          onClick={() => setExpandedRecordId((id) => (id === r.id ? null : r.id))}
                        >
                          <span className="text-sm font-semibold w-28 shrink-0">{r.date}</span>
                          <span className="text-xs text-slate-500 flex-1">{r.totalCount}건</span>
                          <span className="font-mono text-emerald-700 font-semibold text-sm">
                            {formatPrice(r.totalPrice)}
                          </span>
                          {confirmDeleteRecordId !== r.id ? (
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setConfirmDeleteRecordId(r.id);
                              }}
                              className="p-1 text-slate-400 hover:text-red-600"
                              title="삭제"
                            >
                              <Trash2 size={14} />
                            </button>
                          ) : (
                            <span className="flex items-center gap-1 text-xs shrink-0" onClick={(e) => e.stopPropagation()}>
                              <button
                                onClick={() => handleDeleteRecord(r.id)}
                                className="text-red-600 font-semibold hover:underline px-1"
                              >
                                삭제
                              </button>
                              <button
                                onClick={() => setConfirmDeleteRecordId(null)}
                                className="text-slate-400 hover:underline px-1"
                              >
                                취소
                              </button>
                            </span>
                          )}
                          {expandedRecordId === r.id ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                        </div>
                        {expandedRecordId === r.id && (
                          <div className="px-4 pb-3 bg-slate-50">
                            {r.items.map((item) => (
                              <div key={item.id} className="flex items-center gap-2 py-1.5 text-sm">
                                <span className="flex-1 truncate">{item.name}</span>
                                <span
                                  className={`text-xs font-semibold px-2 py-0.5 rounded-sm ${reasonStyle(
                                    item.reason
                                  )}`}
                                >
                                  {item.reason || "기본"}
                                </span>
                                <span className="text-xs text-slate-500 w-10 text-right">×{item.qty}</span>
                                <span className="font-mono text-emerald-700 w-24 text-right">
                                  {formatPrice(item.price * item.qty)}
                                </span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    ))}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* 하단 고정 정산 목록 바 */}
      {selection.length > 0 && (
        <div className="fixed bottom-0 left-0 right-0 bg-slate-900 text-white shadow-lg z-20">
          {cartOpen && (
            <div className="max-w-3xl mx-auto px-5 sm:px-8 pt-3 max-h-64 overflow-y-auto">
              {selection.map((item) => (
                <div
                  key={item.id}
                  className="flex items-center justify-between gap-2 py-2 border-b border-slate-700 text-sm"
                >
                  <div className="flex-1 min-w-0">
                    <p className="truncate">{item.name}</p>
                    <span className={`inline-block mt-0.5 text-xs font-semibold px-2 py-0.5 rounded-sm ${reasonStyle(item.reason)}`}>
                      {item.reason || "기본"}
                    </span>
                  </div>
                  <div className="flex items-center gap-1 shrink-0">
                    <button
                      onClick={() => decrementQty(item.id)}
                      className="w-8 h-8 bg-slate-700 hover:bg-slate-600 rounded-sm text-base"
                    >
                      −
                    </button>
                    <span className="w-6 text-center">{item.qty}</span>
                    <button
                      onClick={() => incrementQty(item.id)}
                      className="w-8 h-8 bg-slate-700 hover:bg-slate-600 rounded-sm text-base"
                    >
                      +
                    </button>
                  </div>
                  <div className="font-mono text-amber-400 w-24 text-right shrink-0">
                    {formatPrice(item.price * item.qty)}
                  </div>
                  <button
                    onClick={() => removeFromSelection(item.id)}
                    className="text-slate-400 hover:text-red-400 shrink-0"
                  >
                    <X size={14} />
                  </button>
                </div>
              ))}
            </div>
          )}
          <div className="max-w-3xl mx-auto px-5 sm:px-8 py-3 flex flex-col sm:flex-row sm:items-center gap-2 sm:gap-4">
            <div className="flex items-center justify-between sm:flex-1 sm:gap-4">
              <button onClick={() => setCartOpen((v) => !v)} className="text-sm underline flex items-center gap-1 shrink-0">
                정산 목록 {totalCount}건
                {cartOpen ? <ChevronDown size={14} /> : <ChevronUp size={14} />}
              </button>
              <div className="font-mono font-bold text-amber-400 text-base sm:text-lg shrink-0">
                {formatPrice(totalPrice)}
              </div>
            </div>
            {!showDateInput ? (
              <div className="flex items-center justify-end gap-3 shrink-0">
                <button
                  onClick={() => setShowDateInput(true)}
                  className="flex items-center gap-1 bg-amber-500 text-slate-900 font-semibold text-sm px-3 py-1.5 hover:bg-amber-400"
                >
                  <Save size={14} />
                  저장
                </button>
                {!confirmClear ? (
                  <button onClick={() => setConfirmClear(true)} className="text-xs text-slate-400 hover:text-white">
                    초기화
                  </button>
                ) : (
                  <span className="flex items-center gap-1 text-xs">
                    <button onClick={clearSelection} className="text-red-400 font-semibold hover:underline">
                      비우기
                    </button>
                    <button onClick={() => setConfirmClear(false)} className="text-slate-400 hover:underline">
                      취소
                    </button>
                  </span>
                )}
              </div>
            ) : (
              <div className="flex items-center justify-end gap-2 shrink-0 flex-wrap">
                <input
                  type="date"
                  value={saveDate}
                  onChange={(e) => setSaveDate(e.target.value)}
                  className="bg-slate-800 border border-slate-600 text-white text-base px-2 py-1.5 focus:outline-none focus:border-amber-500"
                />
                <button
                  onClick={handleSaveSettlement}
                  className="flex items-center gap-1 bg-amber-500 text-slate-900 font-semibold text-sm px-3 py-1.5 hover:bg-amber-400"
                >
                  <Check size={14} />
                  확인
                </button>
                <button
                  onClick={() => setShowDateInput(false)}
                  className="text-slate-400 hover:text-white px-1"
                >
                  <X size={16} />
                </button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
