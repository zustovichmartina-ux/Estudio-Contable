/**
 * Colores de Fiordo Café usados fuera del CSS (PDF y barra del celular).
 * Mantener alineados con los tokens de app/globals.css.
 */
export const MARCA = {
  primario: "#762303", // marrón del logo
  primarioClaro: "#FFEDD9", // crema
  acento: "#A0553A", // terracota de promos
  tinta: "#3B1A0C",
  gris: "#7A5847",
  linea: "#ECDAC6",
  ok: "#4C7A34",
  alerta: "#A86F0C",
  mal: "#C0392B",
};

export const hexARgb = (hex: string): [number, number, number] => {
  const n = parseInt(hex.slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
};
