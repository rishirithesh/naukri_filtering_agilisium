import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

// ─── Constants ────────────────────────────────────────────────────────────────
const FILTER_COLS = {
  Department: "Department",
  Role: "Role",
  Industry: "Industry",
  "UG Degree": "Under Graduation degree",
  "UG University": "UG University/institute Name",
  "PG Degree": "Post graduation degree",
};

const TABLE_COLS = [
  { key: "Name", label: "Name", width: 160 },
  { key: "Role", label: "Role", width: 200 },
  { key: "Department", label: "Department", width: 180 },
  { key: "Industry", label: "Industry", width: 180 },
  { key: "Current Location", label: "Location", width: 150 },
  { key: "Total Experience", label: "Experience", width: 150 },
  { key: "Key Skills", label: "Key Skills", width: 280 },
  { key: "Annual Salary", label: "Salary", width: 140 },
  { key: "Notice period/ Availability to join", label: "Notice Period", width: 170 },
  { key: "Under Graduation degree", label: "UG Degree", width: 180 },
  { key: "Post graduation degree", label: "PG Degree", width: 180 },
  { key: "Email ID", label: "Email", width: 220 },
  { key: "Phone Number", label: "Phone", width: 150 },
];

const INITIAL_COLUMN_FILTERS = TABLE_COLS.reduce((acc, col) => {
  acc[col.key] = "";
  return acc;
}, {});

const ROW_HEIGHT = 52;
const OVERSCAN = 8;

// ─── Utility ──────────────────────────────────────────────────────────────────
function useDebounce(value, delay) {
  const [debouncedValue, setDebouncedValue] = useState(value);
  useEffect(() => {
    const handler = setTimeout(() => setDebouncedValue(value), delay);
    return () => clearTimeout(handler);
  }, [value, delay]);
  return debouncedValue;
}

function exportToCSV(rows, filename) {
  const csv = Papa.unparse(rows);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function exportToXLSX(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Candidates");
  XLSX.writeFile(wb, filename);
}

// ─── Multi-Select Dropdown ────────────────────────────────────────────────────
function MultiSelectDropdown({ label, options, selected, onChange }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const ref = useRef(null);
  const listRef = useRef(null);
  const debouncedSearch = useDebounce(search, 120);

  useEffect(() => {
    function handleClick(e) {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, []);

  const filtered = useMemo(() => {
    if (!debouncedSearch) return options;
    const q = debouncedSearch.toLowerCase();
    return options.filter((o) => o.toLowerCase().includes(q));
  }, [options, debouncedSearch]);

  // Virtualize if large list
  const ITEM_H = 34;
  const maxVisible = 10;
  const [scrollTop, setScrollTop] = useState(0);
  const totalH = filtered.length * ITEM_H;
  const containerH = Math.min(filtered.length * ITEM_H, maxVisible * ITEM_H);
  const startIdx = Math.max(0, Math.floor(scrollTop / ITEM_H) - 2);
  const endIdx = Math.min(filtered.length - 1, Math.floor((scrollTop + containerH) / ITEM_H) + 2);
  const visibleItems = filtered.slice(startIdx, endIdx + 1);

  function toggle(val) {
    if (selected.includes(val)) onChange(selected.filter((s) => s !== val));
    else onChange([...selected, val]);
  }

  return (
    <div ref={ref} style={{ position: "relative", marginBottom: 8 }}>
      <button
        onClick={() => { setOpen(!open); setSearch(""); setScrollTop(0); }}
        style={{
          width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
          background: selected.length ? "#1a2a3a" : "#0f1923",
          border: `1px solid ${selected.length ? "#3b9eff" : "#1e3a5f"}`,
          borderRadius: 8, padding: "8px 12px", cursor: "pointer", color: "#e2eaf4",
          fontSize: 13, fontFamily: "'DM Sans', sans-serif", transition: "all 0.15s",
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 160 }}>
          {selected.length === 0 ? label : selected.length === 1 ? selected[0] : `${label} (${selected.length})`}
        </span>
        <span style={{ marginLeft: 6, color: "#5a8fc4", fontSize: 10 }}>{open ? "▲" : "▼"}</span>
      </button>
      {open && (
        <div style={{
          position: "absolute", left: 0, right: 0, zIndex: 1000,
          background: "#0d1e2e", border: "1px solid #1e3a5f",
          borderRadius: 8, boxShadow: "0 8px 32px rgba(0,0,0,0.5)", overflow: "hidden",
        }}>
          <div style={{ padding: "8px 8px 4px" }}>
            <input
              autoFocus
              value={search}
              onChange={(e) => { setSearch(e.target.value); setScrollTop(0); if (listRef.current) listRef.current.scrollTop = 0; }}
              placeholder="Search..."
              style={{
                width: "100%", background: "#1a2a3a", border: "1px solid #1e3a5f",
                borderRadius: 6, padding: "6px 10px", color: "#e2eaf4", fontSize: 13,
                fontFamily: "'DM Sans', sans-serif", outline: "none", boxSizing: "border-box",
              }}
            />
          </div>
          {selected.length > 0 && (
            <div style={{ padding: "2px 8px 4px", display: "flex", gap: 4, flexWrap: "wrap" }}>
              <button onClick={() => onChange([])} style={{
                background: "#1e3a5f", border: "none", borderRadius: 4, padding: "2px 8px",
                color: "#5a8fc4", fontSize: 11, cursor: "pointer",
              }}>Clear all</button>
            </div>
          )}
          <div
            ref={listRef}
            onScroll={(e) => setScrollTop(e.target.scrollTop)}
            style={{ height: containerH, overflowY: "auto", position: "relative" }}
          >
            <div style={{ height: totalH, position: "relative" }}>
              {visibleItems.map((opt, i) => {
                const idx = startIdx + i;
                const checked = selected.includes(opt);
                return (
                  <div
                    key={opt}
                    onClick={() => toggle(opt)}
                    style={{
                      position: "absolute", top: idx * ITEM_H, left: 0, right: 0,
                      height: ITEM_H, display: "flex", alignItems: "center", gap: 8,
                      padding: "0 12px", cursor: "pointer", fontSize: 13,
                      color: checked ? "#3b9eff" : "#b0c8e4",
                      background: checked ? "rgba(59,158,255,0.08)" : "transparent",
                      fontFamily: "'DM Sans', sans-serif",
                      transition: "background 0.1s",
                    }}
                    onMouseEnter={(e) => { if (!checked) e.currentTarget.style.background = "rgba(255,255,255,0.04)"; }}
                    onMouseLeave={(e) => { if (!checked) e.currentTarget.style.background = "transparent"; }}
                  >
                    <div style={{
                      width: 16, height: 16, border: `2px solid ${checked ? "#3b9eff" : "#2a4a6f"}`,
                      borderRadius: 4, background: checked ? "#3b9eff" : "transparent",
                      display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                      transition: "all 0.15s",
                    }}>
                      {checked && <span style={{ color: "#fff", fontSize: 10, lineHeight: 1 }}>✓</span>}
                    </div>
                    <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{opt}</span>
                  </div>
                );
              })}
            </div>
          </div>
          {filtered.length === 0 && (
            <div style={{ padding: "16px 12px", color: "#5a8fc4", fontSize: 13, textAlign: "center" }}>No results</div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── Virtualized Table ────────────────────────────────────────────────────────
function VirtualTable({ rows, visibleCols, columnFilters, onColumnFilterChange }) {
  const headerRef = useRef(null);
  const bodyRef = useRef(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [containerH, setContainerH] = useState(600);

  useEffect(() => {
    const ro = new ResizeObserver((entries) => {
      if (entries[0]) setContainerH(entries[0].contentRect.height);
    });
    if (bodyRef.current) ro.observe(bodyRef.current);
    return () => ro.disconnect();
  }, []);

  const totalHeight = rows.length * ROW_HEIGHT;
  const startIdx = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN);
  const endIdx = Math.min(rows.length - 1, Math.ceil((scrollTop + containerH) / ROW_HEIGHT) + OVERSCAN);
  const visibleRows = rows.slice(startIdx, endIdx + 1);

  const totalWidth = visibleCols.reduce((s, c) => s + c.width, 0);

  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden", minHeight: 0 }}>
      {/* Header */}
      <div ref={headerRef} style={{ overflowX: "auto", flexShrink: 0 }} onScroll={(e) => {
        if (bodyRef.current) bodyRef.current.scrollLeft = e.target.scrollLeft;
      }}>
        <div style={{
          display: "flex", minWidth: totalWidth,
          background: "#0a1628", borderBottom: "2px solid #1e3a5f",
          position: "sticky", top: 0, zIndex: 10,
        }}>
          {visibleCols.map((col) => (
            <div key={col.key} style={{
              width: col.width, minWidth: col.width, padding: "12px 14px",
              fontSize: 11, fontWeight: 700, letterSpacing: "0.08em",
              color: "#5a8fc4", textTransform: "uppercase", flexShrink: 0,
              fontFamily: "'DM Sans', sans-serif",
            }}>
              {col.label}
            </div>
          ))}
        </div>
        <div style={{
          display: "flex", minWidth: totalWidth,
          background: "#07111f", borderBottom: "1px solid #1b364f",
        }}>
          {visibleCols.map((col) => (
            <div key={col.key} style={{
              width: col.width, minWidth: col.width, padding: "8px 14px",
              flexShrink: 0,
            }}>
              <input
                value={columnFilters[col.key] ?? ""}
                onChange={(e) => onColumnFilterChange(col.key, e.target.value)}
                placeholder="Filter..."
                style={{
                  width: "100%", background: "#081523", border: "1px solid #1f3a59",
                  borderRadius: 6, padding: "7px 10px", color: "#e2eaf4",
                  fontSize: 12, fontFamily: "'DM Sans', sans-serif", outline: "none",
                }}
              />
            </div>
          ))}
        </div>
      </div>
      {/* Body */}
      <div
        ref={bodyRef}
        onScroll={(e) => {
          setScrollTop(e.target.scrollTop);
          if (headerRef.current) headerRef.current.scrollLeft = e.target.scrollLeft;
        }}
        style={{ flex: 1, overflowY: "auto", overflowX: "auto", position: "relative" }}
      >
        <div style={{ height: totalHeight, minWidth: totalWidth, position: "relative" }}>
          {visibleRows.map((row, i) => {
            const idx = startIdx + i;
            const isEven = idx % 2 === 0;
            return (
              <div
                key={idx}
                style={{
                  position: "absolute", top: idx * ROW_HEIGHT,
                  left: 0, right: 0, height: ROW_HEIGHT,
                  display: "flex", alignItems: "center",
                  minWidth: totalWidth,
                  background: isEven ? "#0d1e2e" : "#0a1824",
                  borderBottom: "1px solid rgba(30,58,95,0.3)",
                  transition: "background 0.1s",
                }}
                onMouseEnter={(e) => { e.currentTarget.style.background = "#102236"; }}
                onMouseLeave={(e) => { e.currentTarget.style.background = isEven ? "#0d1e2e" : "#0a1824"; }}
              >
                {visibleCols.map((col) => {
                  const val = row[col.key] ?? "";
                  const str = String(val);
                  return (
                    <div key={col.key} style={{
                      width: col.width, minWidth: col.width, padding: "0 14px",
                      fontSize: 13, color: "#b0c8e4", flexShrink: 0,
                      fontFamily: "'DM Sans', sans-serif",
                      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                    }} title={str}>
                      {col.key === "Key Skills" ? (
                        <span style={{ color: "#7ab3e0" }}>{str.slice(0, 60)}{str.length > 60 ? "…" : ""}</span>
                      ) : col.key === "Name" ? (
                        <span style={{ color: "#e2eaf4", fontWeight: 600 }}>{str}</span>
                      ) : str}
                    </div>
                  );
                })}
              </div>
            );
          })}
          {rows.length === 0 && (
            <div style={{
              position: "absolute", top: 80, left: 0, right: 0,
              textAlign: "center", color: "#5a8fc4", fontSize: 15,
              fontFamily: "'DM Sans', sans-serif",
            }}>
              No candidates match the current filters.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Main Dashboard ───────────────────────────────────────────────────────────
export default function RecruiterDashboard() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState(null);
  const [filters, setFilters] = useState({
    Department: [], Role: [], Industry: [],
    "Under Graduation degree": [], "UG University/institute Name": [], "Post graduation degree": [],
  });
  const [columnFilters, setColumnFilters] = useState(INITIAL_COLUMN_FILTERS);
  const [skillInput, setSkillInput] = useState("");
  const [globalSearch, setGlobalSearch] = useState("");
  const [exportMenu, setExportMenu] = useState(false);
  const fileInputRef = useRef(null);
  const exportRef = useRef(null);

  const debouncedSkill = useDebounce(skillInput, 200);
  const debouncedGlobal = useDebounce(globalSearch, 180);
  const debouncedColumnFilters = useDebounce(columnFilters, 180);

  // Close export menu on outside click
  useEffect(() => {
    function handle(e) { if (exportRef.current && !exportRef.current.contains(e.target)) setExportMenu(false); }
    document.addEventListener("mousedown", handle);
    return () => document.removeEventListener("mousedown", handle);
  }, []);

  // Parse file
  const handleFile = useCallback((file) => {
    if (!file) return;
    setLoading(true);
    setLoadError(null);
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: (result) => { setData(result.data); setLoading(false); },
        error: (err) => { setLoadError(err.message); setLoading(false); },
      });
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "binary" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
          setData(json);
          setLoading(false);
        } catch (err) { setLoadError(err.message); setLoading(false); }
      };
      reader.readAsBinaryString(file);
    }
  }, []);

  // Distinct values per filter col (memoized)
  const distinctValues = useMemo(() => {
    const result = {};
    for (const [label, col] of Object.entries(FILTER_COLS)) {
      const seen = new Set();
      for (const row of data) {
        const v = row[col];
        if (v && String(v).trim()) seen.add(String(v).trim());
      }
      result[col] = Array.from(seen).sort();
    }
    return result;
  }, [data]);

  // Skill terms
  const skillTerms = useMemo(() => {
    if (!debouncedSkill.trim()) return [];
    return debouncedSkill.split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
  }, [debouncedSkill]);

  // Filter data
  const filteredData = useMemo(() => {
    if (!data.length) return [];
    return data.filter((row) => {
      // Dropdown filters: AND across fields, OR within
      for (const [label, col] of Object.entries(FILTER_COLS)) {
        const sel = filters[col];
        if (sel && sel.length > 0) {
          const val = String(row[col] ?? "").trim();
          if (!sel.includes(val)) return false;
        }
      }
      // Per-column filters
      for (const col of TABLE_COLS) {
        const term = String(debouncedColumnFilters[col.key] ?? "").trim().toLowerCase();
        if (term) {
          const actual = String(row[col.key] ?? "").toLowerCase();
          if (!actual.includes(term)) return false;
        }
      }
      // Skills filter
      if (skillTerms.length > 0) {
        const skills = String(row["Key Skills"] ?? "").toLowerCase();
        if (!skillTerms.every((t) => skills.includes(t))) return false;
      }
      // Global search
      if (debouncedGlobal.trim()) {
        const q = debouncedGlobal.toLowerCase();
        const searchable = TABLE_COLS.map((c) => String(row[c.key] ?? "")).join(" ").toLowerCase();
        if (!searchable.includes(q)) return false;
      }
      return true;
    });
  }, [data, filters, debouncedColumnFilters, skillTerms, debouncedGlobal]);

  // Active filter chips
  const activeChips = useMemo(() => {
    const chips = [];
    for (const [label, col] of Object.entries(FILTER_COLS)) {
      const sel = filters[col] ?? [];
      for (const v of sel) chips.push({ type: "dropdown", label: `${label}: ${v}`, col, val: v });
    }
    for (const col of TABLE_COLS) {
      const term = String(columnFilters[col.key] ?? "").trim();
      if (term) chips.push({ type: "column", label: `${col.label}: ${term}`, col: col.key, term });
    }
    for (const t of skillTerms) chips.push({ type: "skill", label: `Skill: ${t}`, term: t });
    return chips;
  }, [filters, columnFilters, skillTerms]);

  function removeChip(chip) {
    if (chip.type === "dropdown") {
      setFilters((f) => ({ ...f, [chip.col]: f[chip.col].filter((v) => v !== chip.val) }));
    } else if (chip.type === "column") {
      setColumnFilters((prev) => ({ ...prev, [chip.col]: "" }));
    } else {
      setSkillInput((prev) => {
        const terms = prev.split(",").map((s) => s.trim()).filter((s) => s.toLowerCase() !== chip.term);
        return terms.join(", ");
      });
    }
  }

  function clearAll() {
    setFilters({ Department: [], Role: [], Industry: [], "Under Graduation degree": [], "UG University/institute Name": [], "Post graduation degree": [] });
    setColumnFilters(INITIAL_COLUMN_FILTERS);
    setSkillInput("");
    setGlobalSearch("");
  }

  const uniqueDepts = useMemo(() => new Set(filteredData.map((r) => r["Department"])).size, [filteredData]);
  const uniqueRoles = useMemo(() => new Set(filteredData.map((r) => r["Role"])).size, [filteredData]);

  const S = styles;

  if (!data.length && !loading) {
    return (
      <div style={S.uploadScreen}>
        <div style={S.uploadCard}>
          <div style={S.logoMark}>Agilisium Consulting</div>
          <h1 style={S.uploadTitle}>Naukri Sheet Filtering</h1>
          <p style={S.uploadSub}>Upload your candidate export to get started</p>
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,.xlsx,.xls"
            style={{ display: "none" }}
            onChange={(e) => handleFile(e.target.files[0])}
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            style={S.uploadBtn}
            onMouseEnter={(e) => { e.currentTarget.style.background = "#2a7ae4"; }}
            onMouseLeave={(e) => { e.currentTarget.style.background = "#1a6be0"; }}
          >
            Upload CSV / Excel
          </button>
          {loadError && <p style={{ color: "#f87171", marginTop: 12, fontSize: 13 }}>{loadError}</p>}
          <p style={{ color: "#3a5a7a", fontSize: 12, marginTop: 20 }}>Supports .csv, .xlsx, .xls · Runs entirely in browser</p>
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div style={S.uploadScreen}>
        <div style={{ textAlign: "center", color: "#5a8fc4" }}>
          <div style={{ fontSize: 36, marginBottom: 16, animation: "spin 1s linear infinite" }}>⟳</div>
          <p style={{ fontFamily: "'DM Sans', sans-serif", fontSize: 16 }}>Parsing candidates…</p>
        </div>
      </div>
    );
  }

  return (
    <div style={S.root}>
      {/* ── Header ── */}
      <header style={S.header}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <span style={S.logoMark}>Agilisium</span>
          <div>
            <div style={S.appName}>Naukri Sheet Filtering</div>
            <div style={{ fontSize: 11, color: "#3a5a7a", letterSpacing: "0.05em" }}>CANDIDATE FILTERING PLATFORM</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
          {/* Stats */}
          <div style={S.statPill}>
            <span style={{ color: "#5a8fc4", fontSize: 11, letterSpacing: "0.06em" }}>TOTAL</span>
            <span style={{ color: "#e2eaf4", fontWeight: 700, fontSize: 18 }}>{data.length.toLocaleString()}</span>
          </div>
          <div style={{ ...S.statPill, borderColor: "#3b9eff" }}>
            <span style={{ color: "#3b9eff", fontSize: 11, letterSpacing: "0.06em" }}>FILTERED</span>
            <span style={{ color: "#3b9eff", fontWeight: 700, fontSize: 18 }}>{filteredData.length.toLocaleString()}</span>
          </div>
          <div style={S.statPill}>
            <span style={{ color: "#5a8fc4", fontSize: 11, letterSpacing: "0.06em" }}>DEPTS</span>
            <span style={{ color: "#e2eaf4", fontWeight: 700, fontSize: 18 }}>{uniqueDepts}</span>
          </div>
          <div style={S.statPill}>
            <span style={{ color: "#5a8fc4", fontSize: 11, letterSpacing: "0.06em" }}>ROLES</span>
            <span style={{ color: "#e2eaf4", fontWeight: 700, fontSize: 18 }}>{uniqueRoles}</span>
          </div>
          {/* Actions */}
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,.xlsx,.xls"
            style={{ display: "none" }}
            onChange={(e) => handleFile(e.target.files[0])}
          />
          <button onClick={() => fileInputRef.current?.click()} style={S.headerBtn}>
            ↑ Upload
          </button>
          <div ref={exportRef} style={{ position: "relative" }}>
            <button
              onClick={() => setExportMenu(!exportMenu)}
              style={{ ...S.headerBtn, background: "#1a6be0" }}
              onMouseEnter={(e) => { e.currentTarget.style.background = "#2a7ae4"; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = "#1a6be0"; }}
            >
              ↓ Export ▾
            </button>
            {exportMenu && (
              <div style={S.exportMenu}>
                <button onClick={() => { exportToCSV(filteredData, "filtered_candidates.csv"); setExportMenu(false); }} style={S.exportItem}>
                  Export as CSV
                </button>
                <button onClick={() => { exportToXLSX(filteredData, "filtered_candidates.xlsx"); setExportMenu(false); }} style={S.exportItem}>
                  Export as Excel
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* ── Body ── */}
      <div style={S.body}>
        {/* ── Sidebar ── */}
        <aside style={S.sidebar}>
          <div style={S.sidebarTitle}>FILTERS</div>
          {Object.entries(FILTER_COLS).map(([label, col]) => (
            <div key={col} style={{ marginBottom: 4 }}>
              <div style={S.filterLabel}>{label}</div>
              <MultiSelectDropdown
                label={label}
                options={distinctValues[col] ?? []}
                selected={filters[col] ?? []}
                onChange={(sel) => setFilters((f) => ({ ...f, [col]: sel }))}
              />
            </div>
          ))}
          {/* Skills */}
          <div style={{ marginTop: 8 }}>
            <div style={S.filterLabel}>Key Skills</div>
            <textarea
              value={skillInput}
              onChange={(e) => setSkillInput(e.target.value)}
              placeholder="Python, SQL, Power BI"
              rows={3}
              style={{
                width: "100%", background: "#0f1923", border: "1px solid #1e3a5f",
                borderRadius: 8, padding: "8px 12px", color: "#e2eaf4", fontSize: 13,
                fontFamily: "'DM Sans', sans-serif", outline: "none", resize: "none",
                boxSizing: "border-box", lineHeight: 1.5,
              }}
            />
            <div style={{ color: "#3a5a7a", fontSize: 11, marginTop: 4 }}>
              Comma-separated · partial match · case-insensitive
            </div>
          </div>
          {activeChips.length > 0 && (
            <button
              onClick={clearAll}
              style={{ marginTop: 16, width: "100%", background: "transparent", border: "1px solid #3a5a7a", borderRadius: 6, padding: "7px 0", color: "#5a8fc4", fontSize: 13, cursor: "pointer", fontFamily: "'DM Sans', sans-serif" }}
              onMouseEnter={(e) => { e.currentTarget.style.borderColor = "#3b9eff"; e.currentTarget.style.color = "#3b9eff"; }}
              onMouseLeave={(e) => { e.currentTarget.style.borderColor = "#3a5a7a"; e.currentTarget.style.color = "#5a8fc4"; }}
            >
              Clear All Filters
            </button>
          )}
        </aside>

        {/* ── Main ── */}
        <main style={S.main}>
          {/* Chips + search row */}
          <div style={S.controlRow}>
            <div style={{ flex: 1, display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
              {activeChips.length === 0 && (
                <span style={{ color: "#3a5a7a", fontSize: 13, fontFamily: "'DM Sans', sans-serif" }}>No active filters</span>
              )}
              {activeChips.map((chip, i) => (
                <div key={i} style={S.chip}>
                  <span>{chip.label}</span>
                  <button onClick={() => removeChip(chip)} style={S.chipX}>×</button>
                </div>
              ))}
            </div>
            <input
              value={globalSearch}
              onChange={(e) => setGlobalSearch(e.target.value)}
              placeholder="Search all columns…"
              style={S.searchInput}
            />
          </div>

          {/* Table */}
          <VirtualTable
            rows={filteredData}
            visibleCols={TABLE_COLS}
            columnFilters={columnFilters}
            onColumnFilterChange={(key, value) => setColumnFilters((prev) => ({ ...prev, [key]: value }))}
          />
        </main>
      </div>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        body { margin: 0; background: #070e18; }
        @keyframes spin { to { transform: rotate(360deg); } }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #0a1628; }
        ::-webkit-scrollbar-thumb { background: #1e3a5f; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: #2a5a8f; }
      `}</style>
    </div>
  );
}

// ─── Styles ───────────────────────────────────────────────────────────────────
const styles = {
  root: {
    height: "100vh", display: "flex", flexDirection: "column",
    background: "#070e18", fontFamily: "'DM Sans', sans-serif",
    overflow: "hidden",
  },
  uploadScreen: {
    height: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
    background: "radial-gradient(ellipse at 50% 0%, #0a1e38 0%, #070e18 70%)",
  },
  uploadCard: {
    textAlign: "center", padding: "48px 56px",
    background: "rgba(13,30,46,0.8)", border: "1px solid #1e3a5f",
    borderRadius: 16, backdropFilter: "blur(12px)",
    boxShadow: "0 24px 64px rgba(0,0,0,0.5)",
  },
  uploadTitle: {
    color: "#e2eaf4", fontSize: 32, fontWeight: 700, margin: "12px 0 8px",
    letterSpacing: "-0.01em",
  },
  uploadSub: { color: "#5a8fc4", fontSize: 15, margin: "0 0 32px" },
  uploadBtn: {
    background: "#1a6be0", border: "none", borderRadius: 10, padding: "14px 36px",
    color: "#fff", fontSize: 15, fontWeight: 600, cursor: "pointer",
    fontFamily: "'DM Sans', sans-serif", transition: "background 0.15s",
  },
  logoMark: { fontSize: 24, color: "#3b9eff" },
  header: {
    display: "flex", alignItems: "center", justifyContent: "space-between",
    padding: "0 24px", height: 64, background: "#0a1628",
    borderBottom: "1px solid #1e3a5f", flexShrink: 0, gap: 16,
  },
  appName: {
    color: "#e2eaf4", fontSize: 17, fontWeight: 700, letterSpacing: "-0.01em",
  },
  statPill: {
    display: "flex", flexDirection: "column", alignItems: "center",
    padding: "4px 14px", border: "1px solid #1e3a5f", borderRadius: 8,
    background: "#0d1e2e", minWidth: 64,
  },
  headerBtn: {
    background: "#1a2a3a", border: "1px solid #1e3a5f", borderRadius: 8,
    padding: "8px 16px", color: "#b0c8e4", fontSize: 13, fontWeight: 600,
    cursor: "pointer", fontFamily: "'DM Sans', sans-serif", transition: "all 0.15s",
  },
  exportMenu: {
    position: "absolute", right: 0, top: "calc(100% + 6px)", zIndex: 100,
    background: "#0d1e2e", border: "1px solid #1e3a5f", borderRadius: 8,
    boxShadow: "0 8px 24px rgba(0,0,0,0.4)", overflow: "hidden", minWidth: 160,
  },
  exportItem: {
    display: "block", width: "100%", padding: "11px 16px",
    background: "transparent", border: "none", color: "#b0c8e4", fontSize: 13,
    cursor: "pointer", textAlign: "left", fontFamily: "'DM Sans', sans-serif",
    transition: "background 0.1s",
  },
  body: {
    display: "flex", flex: 1, overflow: "hidden",
  },
  sidebar: {
    width: 240, flexShrink: 0, background: "#0a1628",
    borderRight: "1px solid #1e3a5f", overflowY: "auto", padding: "16px 12px",
  },
  sidebarTitle: {
    color: "#3a5a7a", fontSize: 10, fontWeight: 700, letterSpacing: "0.12em",
    marginBottom: 16, padding: "0 4px",
  },
  filterLabel: {
    color: "#5a8fc4", fontSize: 11, fontWeight: 600, letterSpacing: "0.06em",
    marginBottom: 4, padding: "0 2px",
  },
  main: {
    flex: 1, display: "flex", flexDirection: "column", overflow: "hidden",
    padding: "12px 16px", gap: 10,
  },
  controlRow: {
    display: "flex", alignItems: "flex-start", gap: 12, flexShrink: 0,
    background: "#0d1e2e", border: "1px solid #1e3a5f",
    borderRadius: 10, padding: "10px 14px", flexWrap: "wrap",
  },
  chip: {
    display: "flex", alignItems: "center", gap: 4,
    background: "rgba(59,158,255,0.12)", border: "1px solid rgba(59,158,255,0.3)",
    borderRadius: 20, padding: "3px 10px 3px 10px",
    color: "#7ab3e0", fontSize: 12, fontFamily: "'DM Sans', sans-serif",
  },
  chipX: {
    background: "none", border: "none", color: "#3b9eff", fontSize: 16,
    cursor: "pointer", padding: "0 0 0 4px", lineHeight: 1,
  },
  searchInput: {
    background: "#0f1923", border: "1px solid #1e3a5f",
    borderRadius: 8, padding: "7px 14px", color: "#e2eaf4", fontSize: 13,
    fontFamily: "'DM Sans', sans-serif", outline: "none", minWidth: 220,
  },
};
