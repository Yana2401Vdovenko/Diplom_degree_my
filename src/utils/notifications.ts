export type NotificationType = "success" | "info" | "warning" | "error";
export type AppNotification = {
  id: string;
  message: string;
  type: NotificationType;
  source: string;
  createdAt: string;
  read: boolean;
};

const STORAGE_KEY = "app-notifications";
const LEGACY_STORAGE_KEY = "workload-notifications";
const EVENT_NAME = "app-notification";
const SOUND_KEY = "app-notification-sound-enabled";

export const isNotificationSoundEnabled = () => {
  try { return localStorage.getItem(SOUND_KEY) !== "false"; } catch { return true; }
};

export const setNotificationSoundEnabled = (enabled: boolean) => {
  try { localStorage.setItem(SOUND_KEY, String(enabled)); } catch { /* Keep the setting for the current session when storage is unavailable. */ }
};

export const playNotificationSound = () => {
  if (!isNotificationSoundEnabled() || typeof window === "undefined") return;
  try {
    const audioWindow = window as Window & { webkitAudioContext?: typeof AudioContext };
    const AudioContextConstructor = window.AudioContext ?? audioWindow.webkitAudioContext;
    if (!AudioContextConstructor) return;
    const context = new AudioContextConstructor();
    const playTones = () => {
      const start = context.currentTime;
      [
        { frequency: 587, at: start, duration: 0.12 },
        { frequency: 880, at: start + 0.075, duration: 0.14 },
      ].forEach(({ frequency, at, duration }) => {
        const oscillator = context.createOscillator();
        const gain = context.createGain();
        oscillator.type = "sine";
        oscillator.frequency.setValueAtTime(frequency, at);
        gain.gain.setValueAtTime(0.0001, at);
        gain.gain.exponentialRampToValueAtTime(0.075, at + 0.018);
        gain.gain.exponentialRampToValueAtTime(0.0001, at + duration);
        oscillator.connect(gain);
        gain.connect(context.destination);
        oscillator.start(at);
        oscillator.stop(at + duration);
      });
      window.setTimeout(() => { void context.close().catch(() => undefined); }, 450);
    };
    if (context.state === "suspended") void context.resume().then(playTones).catch(() => { void context.close().catch(() => undefined); });
    else playTones();
  } catch { /* Audio is optional and must never interrupt notifications. */ }
};

const normalize = (value: unknown): AppNotification | null => {
  if (!value || typeof value !== "object") return null;
  const row = value as Partial<AppNotification>;
  if (!row.message) return null;
  const type: NotificationType = ["success", "info", "warning", "error"].includes(String(row.type)) ? row.type as NotificationType : "info";
  const createdAt = row.createdAt && !Number.isNaN(new Date(row.createdAt).getTime()) ? row.createdAt : new Date().toISOString();
  return { id: String(row.id ?? `${Date.now()}-${Math.random()}`), message: String(row.message), type, source: String(row.source ?? "Система"), createdAt, read: Boolean(row.read) };
};

export const readNotifications = (): AppNotification[] => {
  try {
    const current = JSON.parse(sessionStorage.getItem(STORAGE_KEY) ?? "[]") as unknown[];
    const legacy = JSON.parse(sessionStorage.getItem(LEGACY_STORAGE_KEY) ?? "[]") as unknown[];
    const merged = [...current, ...legacy].map(normalize).filter((item): item is AppNotification => Boolean(item));
    const unique = [...new Map(merged.map((item) => [item.id, item])).values()].slice(0, 100);
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(unique));
    sessionStorage.removeItem(LEGACY_STORAGE_KEY);
    return unique;
  } catch { return []; }
};

export const saveNotifications = (items: AppNotification[]) => {
  try { sessionStorage.setItem(STORAGE_KEY, JSON.stringify(items.slice(0, 100))); } catch { /* Notifications remain available in memory. */ }
};

export const notify = (message: string, type: NotificationType = "info", source = "Система") => {
  const item: AppNotification = { id: `${Date.now()}-${Math.random()}`, message, type, source, createdAt: new Date().toISOString(), read: false };
  const items = [item, ...readNotifications()].slice(0, 100);
  saveNotifications(items);
  playNotificationSound();
  window.dispatchEvent(new CustomEvent<AppNotification>(EVENT_NAME, { detail: item }));
};

export const notificationEventName = EVENT_NAME;
