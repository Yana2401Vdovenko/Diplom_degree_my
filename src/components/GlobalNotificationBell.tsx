import { useEffect, useRef, useState } from "react";
import { FiBell, FiVolume2, FiVolumeX } from "react-icons/fi";
import { isNotificationSoundEnabled, notificationEventName, readNotifications, saveNotifications, setNotificationSoundEnabled } from "../utils/notifications";
import type { AppNotification } from "../utils/notifications";

const typeLabel: Record<AppNotification["type"], string> = { success: "Успіх", info: "Інформація", warning: "Попередження", error: "Помилка" };
const typeColor: Record<AppNotification["type"], string> = { success: "#16a34a", info: "#2563eb", warning: "#d97706", error: "#dc2626" };

export default function GlobalNotificationBell() {
  const [items, setItems] = useState<AppNotification[]>(readNotifications);
  const [open, setOpen] = useState(false);
  const [soundEnabled, setSoundEnabled] = useState(isNotificationSoundEnabled);
  const rootRef = useRef<HTMLDivElement>(null);
  const unreadCount = items.filter((item) => !item.read).length;

  useEffect(() => {
    const onNotification = (event: Event) => {
      const item = (event as CustomEvent<AppNotification>).detail;
      if (!item?.message) return;
      setItems((current) => {
        const next = [item, ...current.filter((existing) => existing.id !== item.id)].slice(0, 100);
        saveNotifications(next);
        return next;
      });
    };
    const onOutsideClick = (event: MouseEvent) => {
      if (rootRef.current && !rootRef.current.contains(event.target as Node)) setOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => { if (event.key === "Escape") setOpen(false); };
    window.addEventListener(notificationEventName, onNotification);
    document.addEventListener("mousedown", onOutsideClick);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      window.removeEventListener(notificationEventName, onNotification);
      document.removeEventListener("mousedown", onOutsideClick);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, []);

  const markAllRead = () => setItems((current) => {
    const next = current.map((item) => ({ ...item, read: true }));
    saveNotifications(next);
    return next;
  });
  const clear = () => {
    setItems([]);
    saveNotifications([]);
  };
  const toggleSound = () => setSoundEnabled((current) => {
    const next = !current;
    setNotificationSoundEnabled(next);
    return next;
  });

  return <div ref={rootRef} className="global-notification-bell" style={{ position: "fixed", top: 16, right: 20, zIndex: 20000, fontFamily: "Inter, Arial, sans-serif" }}>
    <button type="button" aria-label={unreadCount ? `Сповіщення, непрочитаних: ${unreadCount}` : "Сповіщення"} aria-expanded={open} onClick={() => setOpen((current) => !current)} style={{ position: "relative", display: "grid", placeItems: "center", width: 42, height: 42, border: "1px solid #dbe3ee", borderRadius: 12, background: "#fff", color: "#334155", boxShadow: "0 4px 14px rgba(15,23,42,.12)", cursor: "pointer" }}>
      <FiBell size={19} />{unreadCount > 0 && <span aria-hidden="true" style={{ position: "absolute", top: -5, right: -5, minWidth: 18, height: 18, padding: "0 4px", display: "grid", placeItems: "center", border: "2px solid white", borderRadius: 999, background: "#2563eb", color: "#fff", fontSize: 10, fontWeight: 700 }}>{unreadCount > 99 ? "99+" : unreadCount}</span>}
    </button>
    {open && <section role="dialog" aria-label="Центр сповіщень" style={{ position: "absolute", top: 50, right: 0, width: "min(420px, calc(100vw - 28px))", maxHeight: "min(72vh, 560px)", display: "flex", flexDirection: "column", overflow: "hidden", border: "1px solid #e2e8f0", borderRadius: 14, background: "#fff", color: "#0f172a", boxShadow: "0 20px 50px rgba(15,23,42,.2)" }}>
      <header style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "14px 16px", borderBottom: "1px solid #e2e8f0" }}><strong style={{ fontSize: 14 }}>Сповіщення</strong><div style={{ display: "flex", alignItems: "center", gap: 10 }}><button type="button" onClick={markAllRead} style={{ border: 0, background: "transparent", color: "#2563eb", fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" }}>Позначити всі як прочитані</button><button type="button" aria-label={soundEnabled ? "Вимкнути звук сповіщень" : "Увімкнути звук сповіщень"} title={soundEnabled ? "Вимкнути звук" : "Увімкнути звук"} onClick={toggleSound} style={{ display: "grid", placeItems: "center", width: 28, height: 28, border: "1px solid #e2e8f0", borderRadius: 7, background: "#fff", color: soundEnabled ? "#2563eb" : "#94a3b8", cursor: "pointer" }}>{soundEnabled ? <FiVolume2 size={15} /> : <FiVolumeX size={15} />}</button><button type="button" onClick={clear} style={{ border: 0, background: "transparent", color: "#64748b", fontSize: 12, cursor: "pointer" }}>Очистити</button></div></header>
      <div style={{ overflowY: "auto", padding: 8 }}>{items.length === 0 ? <p style={{ margin: 0, padding: 24, textAlign: "center", color: "#64748b", fontSize: 13 }}>Немає сповіщень</p> : items.map((item) => <article key={item.id} style={{ display: "flex", gap: 10, padding: "11px 10px", borderBottom: "1px solid #f1f5f9", borderRadius: 8, background: item.read ? "#fff" : "#eff6ff" }}><span aria-hidden="true" style={{ width: 8, height: 8, flex: "0 0 auto", marginTop: 5, borderRadius: 999, background: typeColor[item.type] }} /><div style={{ minWidth: 0, flex: 1 }}><p style={{ margin: 0, color: "#334155", fontSize: 13, lineHeight: 1.45, whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>{item.message}</p><div style={{ display: "flex", justifyContent: "space-between", gap: 8, marginTop: 6, color: "#94a3b8", fontSize: 11 }}><span>{item.source} · {typeLabel[item.type]}</span><time>{new Date(item.createdAt).toLocaleString("uk-UA", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })}</time></div></div></article>)}</div>
    </section>}
  </div>;
}
