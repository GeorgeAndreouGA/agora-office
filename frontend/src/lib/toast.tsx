import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';

type Tone = 'info' | 'success' | 'error';
interface Toast {
  id: number;
  tone: Tone;
  text: string;
}

interface ToastApi {
  show(text: string, tone?: Tone): void;
}

const ToastContext = createContext<ToastApi>({ show: () => undefined });

export function ToastProvider({ children }: { children: ReactNode }) {
  const [toasts, setToasts] = useState<Toast[]>([]);
  const next = useRef(1);
  const timers = useRef(new Set<number>());

  // Don't let a pending dismissal fire setState after the provider is gone.
  useEffect(() => {
    const pending = timers.current;
    return () => {
      pending.forEach((h) => window.clearTimeout(h));
      pending.clear();
    };
  }, []);

  const show = useCallback((text: string, tone: Tone = 'info') => {
    const id = next.current++;
    setToasts((t) => [...t.slice(-3), { id, tone, text }]);
    const handle = window.setTimeout(() => {
      timers.current.delete(handle);
      setToasts((t) => t.filter((x) => x.id !== id));
    }, tone === 'error' ? 9000 : 4500);
    timers.current.add(handle);
  }, []);

  const api = useMemo(() => ({ show }), [show]);

  return (
    <ToastContext.Provider value={api}>
      {children}
      {/* Two always-mounted live regions (a region must exist before its content changes to be
          announced reliably): errors interrupt, everything else waits its turn. */}
      <div className="toasts">
        <div className="toasts__live" role="status" aria-live="polite">
          {toasts
            .filter((t) => t.tone !== 'error')
            .map((t) => (
              <div key={t.id} className={`toast toast--${t.tone}`}>
                {t.text}
              </div>
            ))}
        </div>
        <div className="toasts__live" role="alert" aria-live="assertive">
          {toasts
            .filter((t) => t.tone === 'error')
            .map((t) => (
              <div key={t.id} className={`toast toast--${t.tone}`}>
                {t.text}
              </div>
            ))}
        </div>
      </div>
    </ToastContext.Provider>
  );
}

export function useToast(): ToastApi {
  return useContext(ToastContext);
}
