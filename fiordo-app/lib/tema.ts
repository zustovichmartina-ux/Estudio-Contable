/**
 * Colores de marca usados fuera del CSS (PDF y color de la barra del celular).
 * PROVISORIOS hasta tener el estilo real de Fiordo Café: cambiar acá y en app/globals.css (:root).
 */
export const MARCA = {
  primario: "#1E4A5C", // azul fiordo
  primarioClaro: "#E4EEF1",
  acento: "#B5794A", // café
  tinta: "#15232B",
  gris: "#5E6F78",
  linea: "#D6E0E4",
  ok: "#2E8049",
  alerta: "#B7800F",
  mal: "#BC3B2C",
};

export const hexARgb = (hex: string): [number, number, number] => {
  const n = parseInt(hex.slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
};
