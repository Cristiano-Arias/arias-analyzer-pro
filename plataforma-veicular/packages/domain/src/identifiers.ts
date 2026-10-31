/**
 * Identificadores veiculares: placa (padrão antigo e Mercosul) e VIN.
 *
 * Validamos apenas o FORMATO. Validar que a placa ou o chassi existem exige
 * consulta a uma base oficial ou autorizada, que não temos (risco R-01).
 */

export type PlateFormat = 'legacy' | 'mercosul';

export interface PlateResult {
  valid: boolean;
  /** Placa normalizada: maiúsculas, sem hífen nem espaços. */
  normalized: string;
  format?: PlateFormat;
}

const LEGACY_RE = /^[A-Z]{3}[0-9]{4}$/;
const MERCOSUL_RE = /^[A-Z]{3}[0-9][A-Z][0-9]{2}$/;

export function parsePlate(input: string): PlateResult {
  const normalized = input.toUpperCase().replace(/[\s-]/g, '');
  if (LEGACY_RE.test(normalized)) return { valid: true, normalized, format: 'legacy' };
  if (MERCOSUL_RE.test(normalized)) return { valid: true, normalized, format: 'mercosul' };
  return { valid: false, normalized };
}

/**
 * Converte uma placa antiga no equivalente Mercosul. Regra oficial de
 * conversão: o 5º caractere (dígito 0–9) vira a letra A–J.
 * Ex.: ABC1234 -> ABC1C34.
 */
export function legacyToMercosul(plate: string): string {
  const p = parsePlate(plate);
  if (!p.valid || p.format !== 'legacy') throw new RangeError(`Placa não está no padrão antigo: ${plate}`);
  const n = p.normalized;
  return n.slice(0, 4) + String.fromCharCode(65 + Number(n[4])) + n.slice(5);
}

/**
 * Duas placas representam o mesmo veículo? Considera a conversão para
 * Mercosul, para evitar cadastro duplicado.
 */
export function samePlate(a: string, b: string): boolean {
  const canon = (x: string): string | null => {
    const p = parsePlate(x);
    if (!p.valid) return null;
    return p.format === 'legacy' ? legacyToMercosul(p.normalized) : p.normalized;
  };
  const ca = canon(a);
  return ca !== null && ca === canon(b);
}

export interface VinResult {
  valid: boolean;
  normalized: string;
  /** Código do fabricante (WMI): 3 primeiros caracteres. */
  wmi?: string;
  reason?: string;
}

/**
 * VIN (ISO 3779): 17 caracteres alfanuméricos, sem I, O ou Q.
 * O dígito verificador da posição 9 é obrigatório só na América do Norte,
 * por isso NÃO é usado para rejeitar VINs brasileiros.
 */
export function parseVin(input: string): VinResult {
  const normalized = input.toUpperCase().replace(/[\s-]/g, '');
  if (normalized.length !== 17) return { valid: false, normalized, reason: 'O VIN deve ter 17 caracteres.' };
  if (/[IOQ]/.test(normalized)) return { valid: false, normalized, reason: 'O VIN não usa as letras I, O ou Q.' };
  if (!/^[A-Z0-9]{17}$/.test(normalized)) return { valid: false, normalized, reason: 'O VIN aceita apenas letras e números.' };
  return { valid: true, normalized, wmi: normalized.slice(0, 3) };
}
