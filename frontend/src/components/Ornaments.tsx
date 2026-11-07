import { useId } from 'react';

/** The Greek key band. Decorative only. */
export function Meander({ className = '' }: { className?: string }) {
  const id = `meander-${useId().replace(/:/g, '')}`;
  return (
    <svg className={`meander ${className}`} width="100%" height="22" aria-hidden="true" focusable="false">
      <defs>
        <pattern id={id} width="22" height="22" patternUnits="userSpaceOnUse">
          <path
            d="M0 20 H22 M0 2 H22 M17 20 V6 H5 V16 H12 V11"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          />
        </pattern>
      </defs>
      <rect width="100%" height="22" fill={`url(#${id})`} />
    </svg>
  );
}

/** Temple pediment drawn across the top of a card. */
export function Pediment({ tone }: { tone: 'terracotta' | 'olive' | 'aegean' }) {
  return (
    <svg className={`pediment pediment--${tone}`} viewBox="0 0 400 56" preserveAspectRatio="none" aria-hidden="true">
      <path d="M0 56 L200 6 L400 56" className="pediment__fill" strokeWidth="2" />
      <path d="M40 56 L200 18 L360 56" fill="none" className="pediment__line" strokeWidth="1" />
    </svg>
  );
}

const ICON_PROPS = {
  fill: 'none',
  stroke: 'currentColor',
  strokeWidth: 1.6,
  strokeLinecap: 'round' as const,
  strokeLinejoin: 'round' as const,
  'aria-hidden': true,
  focusable: false,
};

export function ScrollIcon({ size = 34 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" {...ICON_PROPS}>
      <path d="M6 4h11a2 2 0 0 1 2 2v12" />
      <path d="M6 4a2 2 0 0 0-2 2v1h4V6a2 2 0 0 0-2-2z" />
      <path d="M8 7v11a2 2 0 0 0 2 2h9a2 2 0 0 0 2-2v-1H12v1a2 2 0 0 1-2 2" />
      <path d="M11 9h5M11 12h5M11 15h3" />
    </svg>
  );
}

export function AbacusIcon({ size = 34 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" {...ICON_PROPS}>
      <rect x="3" y="3" width="18" height="18" rx="1.5" />
      <path d="M3 8h18M3 13h18M3 18h18" />
      <circle cx="8" cy="8" r="1.4" fill="currentColor" />
      <circle cx="14" cy="13" r="1.4" fill="currentColor" />
      <circle cx="10" cy="18" r="1.4" fill="currentColor" />
      <circle cx="17" cy="8" r="1.4" fill="currentColor" />
    </svg>
  );
}

/** A painted panel on an easel: Pinax. */
export function PinaxIcon({ size = 34 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" {...ICON_PROPS}>
      <rect x="3" y="3" width="18" height="12" rx="1" />
      <path d="M6 12l3.5-4 3 3 2-2 3.5 3" />
      <circle cx="15.5" cy="6.5" r="1.2" />
      <path d="M12 15v3M8 21l4-3 4 3" />
    </svg>
  );
}

export function DropIcon() {
  return (
    <svg width="36" height="36" viewBox="0 0 24 24" {...ICON_PROPS}>
      <path d="M12 3v12M7 10l5 5 5-5" />
      <path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2" />
    </svg>
  );
}

export function MemoryIcon() {
  return (
    <svg width="28" height="28" viewBox="0 0 24 24" {...ICON_PROPS}>
      <rect x="4" y="6" width="16" height="12" rx="1" />
      <path d="M8 6V4M12 6V4M16 6V4M8 18v2M12 18v2M16 18v2" />
    </svg>
  );
}

export function ClockIcon() {
  return (
    <svg width="28" height="28" viewBox="0 0 24 24" {...ICON_PROPS}>
      <circle cx="12" cy="12" r="8" />
      <path d="M12 7v5l3 2" />
    </svg>
  );
}

export function ShieldIcon() {
  return (
    <svg width="28" height="28" viewBox="0 0 24 24" {...ICON_PROPS}>
      <path d="M12 3l7 3v5c0 4.5-3 8-7 10-4-2-7-5.5-7-10V6z" />
      <path d="M9 12l2 2 4-4" />
    </svg>
  );
}

export function BackIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" {...ICON_PROPS} strokeWidth={1.8}>
      <path d="M15 18l-6-6 6-6" />
    </svg>
  );
}
