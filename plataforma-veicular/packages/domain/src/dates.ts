/**
 * Datas civis (sem horário) no formato ISO `YYYY-MM-DD`.
 * Toda a aritmética é feita em UTC para não sofrer com fuso horário nem com
 * horário de verão. Vencimentos e lembretes são datas civis, não instantes.
 */
export type IsoDate = string;

const ISO_RE = /^(\d{4})-(\d{2})-(\d{2})$/;
const DAY_MS = 86_400_000;

export function parseIsoDate(d: IsoDate): Date {
  const m = ISO_RE.exec(d);
  if (!m) throw new RangeError(`Data inválida (esperado AAAA-MM-DD): ${d}`);
  const [y, mo, da] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const dt = new Date(Date.UTC(y, mo - 1, da));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== da) {
    throw new RangeError(`Data inexistente: ${d}`);
  }
  return dt;
}

export const toIsoDate = (dt: Date): IsoDate => dt.toISOString().slice(0, 10);

export function addDays(d: IsoDate, days: number): IsoDate {
  return toIsoDate(new Date(parseIsoDate(d).getTime() + days * DAY_MS));
}

/** Soma meses preservando o fim do mês: 31/01 + 1 mês = 28/02 (ou 29/02). */
export function addMonths(d: IsoDate, months: number): IsoDate {
  const dt = parseIsoDate(d);
  const y = dt.getUTCFullYear();
  const m = dt.getUTCMonth() + months;
  const lastDay = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return toIsoDate(new Date(Date.UTC(y, m, Math.min(dt.getUTCDate(), lastDay))));
}

/** Dias corridos de `a` até `b` (positivo se b é posterior). */
export function daysBetween(a: IsoDate, b: IsoDate): number {
  return Math.round((parseIsoDate(b).getTime() - parseIsoDate(a).getTime()) / DAY_MS);
}

export const minDate = (a: IsoDate, b: IsoDate): IsoDate => (a <= b ? a : b);
