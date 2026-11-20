import { useEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties } from "react";
import { createPortal } from "react-dom";
import { FiCheck, FiChevronDown, FiSearch, FiX } from "react-icons/fi";

export type SearchableSelectOption = {
  value: string;
  label: string;
  sublabel?: string;
};

type SearchableSelectProps = {
  options: SearchableSelectOption[];
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  ariaLabel?: string;
  disabled?: boolean;
  style?: CSSProperties;
  triggerHeight?: number;
  emptyMessage?: string;
};

export default function SearchableSelect({
  options,
  value,
  onChange,
  placeholder = "Оберіть значення",
  ariaLabel = placeholder,
  disabled = false,
  style,
  triggerHeight = 38,
  emptyMessage = "Нічого не знайдено",
}: SearchableSelectProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  const popupRef = useRef<HTMLDivElement>(null);
  const searchRef = useRef<HTMLInputElement>(null);
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [popupPosition, setPopupPosition] = useState({ top: 0, left: 0, width: 0, above: false });
  const selected = options.find((option) => option.value === value);
  const filtered = useMemo(() => {
    const term = query.trim().toLocaleLowerCase("uk-UA");
    if (!term) return options;
    return options.filter((option) => `${option.label} ${option.sublabel ?? ""}`.toLocaleLowerCase("uk-UA").includes(term));
  }, [options, query]);

  useEffect(() => {
    if (!open) return;
    const updatePosition = () => {
      const rect = rootRef.current?.getBoundingClientRect();
      if (!rect) return;
      const above = window.innerHeight - rect.bottom < 300 && rect.top > 300;
      setPopupPosition({ top: above ? rect.top - 5 : rect.bottom + 5, left: rect.left, width: rect.width, above });
    };
    const closeOutside = (event: PointerEvent) => {
      const target = event.target as Node;
      if (!rootRef.current?.contains(target) && !popupRef.current?.contains(target)) setOpen(false);
    };
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setOpen(false);
        setQuery("");
      }
    };
    document.addEventListener("pointerdown", closeOutside);
    document.addEventListener("keydown", closeOnEscape);
    window.addEventListener("resize", updatePosition);
    window.addEventListener("scroll", updatePosition, true);
    updatePosition();
    requestAnimationFrame(() => searchRef.current?.focus());
    return () => {
      document.removeEventListener("pointerdown", closeOutside);
      document.removeEventListener("keydown", closeOnEscape);
      window.removeEventListener("resize", updatePosition);
      window.removeEventListener("scroll", updatePosition, true);
    };
  }, [open]);

  const choose = (option: SearchableSelectOption) => {
    onChange(option.value);
    setOpen(false);
    setQuery("");
  };

  return <div ref={rootRef} style={{ position: "relative", width: "100%", minWidth: 0, ...style }}>
    <button
      type="button"
      aria-label={ariaLabel}
      aria-haspopup="listbox"
      aria-expanded={open}
      disabled={disabled}
      onClick={() => setOpen((current) => !current)}
      style={{ boxSizing: "border-box", display: "flex", alignItems: "center", gap: 8, width: "100%", height: triggerHeight, padding: "0 10px", border: "1px solid #cbd5e1", borderRadius: 10, background: disabled ? "#f8fafc" : "#fff", color: "#0f172a", font: "inherit", fontSize: 13.5, textAlign: "left", cursor: disabled ? "not-allowed" : "pointer" }}
    >
      <FiSearch aria-hidden="true" size={15} color="#64748b" style={{ flexShrink: 0 }} />
      <span style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: selected ? "#0f172a" : "#64748b" }}>{selected?.label ?? placeholder}</span>
      <FiChevronDown aria-hidden="true" size={15} color="#64748b" style={{ flexShrink: 0, transform: open ? "rotate(180deg)" : undefined }} />
    </button>
    {value && <button type="button" aria-label="Очистити вибір" onClick={() => { onChange(""); setQuery(""); }} style={{ position: "absolute", top: "50%", right: 30, transform: "translateY(-50%)", display: "inline-flex", alignItems: "center", justifyContent: "center", width: 20, height: 20, border: 0, borderRadius: 4, background: "#fff", color: "#64748b", cursor: "pointer" }}><FiX size={15} /></button>}
    {open && createPortal(<div ref={popupRef} style={{ position: "fixed", top: popupPosition.top, left: popupPosition.left, width: popupPosition.width, transform: popupPosition.above ? "translateY(-100%)" : undefined, zIndex: 99999, overflow: "hidden", border: "1px solid #cbd5e1", borderRadius: 10, background: "#fff", boxShadow: "0 10px 25px rgba(15, 23, 42, 0.14)" }}>
      <div style={{ position: "sticky", top: 0, zIndex: 1, display: "flex", alignItems: "center", gap: 8, padding: "7px 10px", borderBottom: "1px solid #e2e8f0", background: "#fff" }}>
        <FiSearch aria-hidden="true" size={15} color="#64748b" />
        <input ref={searchRef} value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Пошук…" aria-label="Пошук у списку" style={{ width: "100%", minWidth: 0, height: 28, padding: 0, border: 0, outline: 0, background: "transparent", color: "#0f172a", fontSize: 13.5 }} />
        {query && <button type="button" aria-label="Очистити пошук" onClick={() => setQuery("")} style={{ display: "inline-flex", border: 0, background: "transparent", color: "#64748b", cursor: "pointer" }}><FiX size={15} /></button>}
      </div>
      <div role="listbox" aria-label={ariaLabel} style={{ maxHeight: 260, overflowY: "auto" }}>
        {filtered.length ? filtered.map((option) => {
          const isSelected = option.value === value;
          return <button key={option.value} type="button" role="option" aria-selected={isSelected} onClick={() => choose(option)} onMouseEnter={(event) => { event.currentTarget.style.background = "#eff6ff"; }} onMouseLeave={(event) => { event.currentTarget.style.background = isSelected ? "#eff6ff" : "#fff"; }} style={{ boxSizing: "border-box", display: "flex", alignItems: "center", gap: 8, width: "100%", minHeight: 36, padding: "8px 12px", border: 0, borderBottom: "1px solid #f1f5f9", background: isSelected ? "#eff6ff" : "#fff", color: isSelected ? "#1d4ed8" : "#1e293b", fontSize: 13.5, textAlign: "left", cursor: "pointer" }}>
            <span style={{ flex: 1, minWidth: 0 }}><span style={{ display: "block", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{option.label}</span>{option.sublabel && <span style={{ display: "block", marginTop: 2, color: "#64748b", fontSize: 12, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{option.sublabel}</span>}</span>
            {isSelected && <FiCheck aria-hidden="true" size={16} style={{ flexShrink: 0 }} />}
          </button>;
        }) : <div style={{ padding: "14px 12px", color: "#64748b", fontSize: 13, textAlign: "center" }}>{emptyMessage}</div>}
      </div>
    </div>, document.body)}
  </div>;
}
