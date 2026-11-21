/**
 * Versión simplificada del logo de Fiordo (círculo marrón con olas en crema).
 * Reemplazar por el logo original cuando esté: guardarlo como public/logo.png y usar <img src="/logo.png">.
 */
export default function Logo({ className = "logo" }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 64 64" role="img" aria-label="Fiordo">
      <circle cx="32" cy="32" r="30.5" fill="#762303" stroke="#FFEEDA" strokeWidth="2" />
      <g fill="none" stroke="#FFEEDA" strokeWidth="2.6" strokeLinecap="round">
        <path d="M21 30c1-7 6-11 11-11 4 0 6 2 6 5" />
        <circle cx="34.5" cy="22.5" r="1" fill="#FFEEDA" stroke="none" />
        <path d="M14 36c4-3 8-3 12 0s8 3 12 0 8-3 12 0" />
        <path d="M18 43c3.5-2.5 7-2.5 10.5 0s7 2.5 10.5 0 7-2.5 7.5-1" />
      </g>
    </svg>
  );
}
