import type { CSSProperties } from "react";
import { FiRotateCcw, FiSearch } from "react-icons/fi";

const fieldStyle: CSSProperties = {
  boxSizing: "border-box", height: 38, minWidth: 150, padding: "0 11px",
  border: "1px solid #cbd5e1", borderRadius: 9, background: "#fff",
  color: "#1e293b", fontSize: 13.5,
};

export function FilterSearch({ value, onChange, placeholder, ariaLabel, minWidth = 240 }: {
  value: string; onChange: (value: string) => void; placeholder: string; ariaLabel: string; minWidth?: number;
}) {
  return <label style={{ display: "flex", alignItems: "center", gap: 8, height: 38, minWidth, flex: "1 1 240px", boxSizing: "border-box", padding: "0 10px", border: "1px solid #cbd5e1", borderRadius: 9, background: "#fff", color: "#64748b" }}>
    <FiSearch size={15} aria-hidden="true" />
    <input aria-label={ariaLabel} value={value} onChange={(event) => onChange(event.target.value)} placeholder={placeholder} style={{ width: "100%", minWidth: 0, border: 0, outline: 0, background: "transparent", color: "#0f172a", fontSize: 13.5 }} />
    {value && <button type="button" title="Очистити пошук" aria-label="Очистити пошук" onClick={() => onChange("")} style={{ border: 0, background: "transparent", color: "#64748b", fontSize: 19, lineHeight: 1, cursor: "pointer" }}>×</button>}
  </label>;
}

export function FilterSelect({ value, onChange, ariaLabel, options, minWidth = 150 }: {
  value: string; onChange: (value: string) => void; ariaLabel: string; options: Array<{ value: string; label: string }>; minWidth?: number;
}) {
  return <select aria-label={ariaLabel} value={value} onChange={(event) => onChange(event.target.value)} style={{ ...fieldStyle, minWidth, flex: "0 1 auto" }}>
    {options.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
  </select>;
}

export function FilterReset({ onClick }: { onClick: () => void }) {
  return <button type="button" onClick={onClick} style={{ display: "inline-flex", alignItems: "center", gap: 6, height: 36, padding: "0 11px", border: "1px solid #cbd5e1", borderRadius: 9, background: "#fff", color: "#475569", fontSize: 13, whiteSpace: "nowrap", cursor: "pointer" }}><FiRotateCcw size={14} />Скинути фільтри</button>;
}

export const filterBarStyle: CSSProperties = { display: "flex", alignItems: "center", flexWrap: "wrap", gap: 9, padding: "12px 16px", borderBottom: "1px solid #e2e8f0" };
