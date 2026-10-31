/**
 * Cálculo de consumo.
 *
 * COMBUSTÍVEL — método "tanque cheio a tanque cheio":
 *   Um trecho começa num abastecimento com tanque cheio e termina no próximo
 *   abastecimento com tanque cheio. O combustível consumido no trecho é a soma
 *   de tudo o que foi colocado DEPOIS do início, incluindo parciais e o
 *   abastecimento final. Trechos com abastecimento esquecido (missedPrevious)
 *   não podem ser calculados com segurança e são descartados (com motivo).
 *   A média geral é ponderada (distância total / combustível total), e não a
 *   média simples dos trechos.
 *
 * ENERGIA ELÉTRICA — estimativa:
 *   A energia recarregada entre a primeira e a última recarga, dividida pela
 *   distância. Se houver SOC final da primeira e da última recarga e a
 *   capacidade útil, corrige a variação de carga da bateria. O resultado é
 *   SEMPRE marcado como estimativa: inclui perdas de recarga quando medido no
 *   carregador e depende da precisão do SOC informado.
 */
import { ENERGY_SOURCES, type EnergyCode, type TankGroup } from './taxonomy.js';
import type { Cents } from './money.js';

export interface FuelLogInput {
  id: string;
  odometerKm: number;
  /** Instante ISO; usado para desempate quando o odômetro é igual. */
  filledAt: string;
  energy: EnergyCode;
  quantity: number;
  totalCents: Cents;
  isFull: boolean;
  missedPrevious: boolean;
}

export type SegmentInvalidReason = 'missed_fill' | 'non_increasing_odometer';

export interface ConsumptionSegment {
  startLogId: string;
  endLogId: string;
  tankGroup: TankGroup;
  distanceKm: number;
  quantity: number;
  costCents: Cents;
  /** km por unidade (km/L ou km/m³); null quando o trecho é inválido. */
  kmPerUnit: number | null;
  /** Energias presentes no trecho (ex.: gasolina e etanol num flex). */
  energies: EnergyCode[];
  invalidReason?: SegmentInvalidReason;
  /** Marcado pela detecção de anomalia. */
  outlier?: boolean;
}

export interface ConsumptionSummary {
  tankGroup: TankGroup;
  segments: ConsumptionSegment[];
  validSegments: number;
  /** Média ponderada; null se não houver nenhum trecho válido. */
  averageKmPerUnit: number | null;
  /** Inverso em L/100 km (ou m³/100 km). */
  averageUnitsPer100Km: number | null;
  /** Custo de combustível por km nos trechos válidos. */
  costPerKmCents: number | null;
  totalDistanceKm: number;
}

function sortLogs<T extends { odometerKm: number; filledAt: string }>(logs: readonly T[]): T[] {
  return [...logs].sort((a, b) => a.odometerKm - b.odometerKm || a.filledAt.localeCompare(b.filledAt));
}

/** Calcula o consumo por grupo de reservatório (líquido, gás). */
export function computeFuelConsumption(logs: readonly FuelLogInput[]): ConsumptionSummary[] {
  const byGroup = new Map<TankGroup, FuelLogInput[]>();
  for (const l of logs) {
    const g = ENERGY_SOURCES[l.energy].tankGroup;
    if (g === 'battery') continue;
    if (!(l.quantity > 0)) throw new RangeError(`Quantidade inválida no registro ${l.id}`);
    byGroup.set(g, [...(byGroup.get(g) ?? []), l]);
  }
  return [...byGroup.entries()].map(([group, groupLogs]) => summarize(group, sortLogs(groupLogs)));
}

function summarize(group: TankGroup, logs: FuelLogInput[]): ConsumptionSummary {
  const segments: ConsumptionSegment[] = [];
  let startIdx = logs.findIndex((l) => l.isFull);

  while (startIdx >= 0 && startIdx < logs.length - 1) {
    const start = logs[startIdx]!;
    let qty = 0;
    let cost = 0;
    let missed = false;
    const energies = new Set<EnergyCode>();
    let endIdx = -1;
    for (let i = startIdx + 1; i < logs.length; i++) {
      const l = logs[i]!;
      qty += l.quantity;
      cost += l.totalCents;
      energies.add(l.energy);
      if (l.missedPrevious) missed = true;
      if (l.isFull) { endIdx = i; break; }
    }
    if (endIdx < 0) break; // trecho em aberto: aguardando próximo tanque cheio

    const end = logs[endIdx]!;
    const distance = end.odometerKm - start.odometerKm;
    const seg: ConsumptionSegment = {
      startLogId: start.id, endLogId: end.id, tankGroup: group,
      distanceKm: distance, quantity: round(qty, 3), costCents: cost,
      kmPerUnit: null, energies: [...energies],
    };
    if (missed) seg.invalidReason = 'missed_fill';
    else if (distance <= 0) seg.invalidReason = 'non_increasing_odometer';
    else seg.kmPerUnit = round(distance / qty, 2);
    segments.push(seg);
    startIdx = endIdx;
  }

  flagOutliers(segments);
  const valid = segments.filter((s) => s.kmPerUnit !== null);
  const dist = valid.reduce((a, s) => a + s.distanceKm, 0);
  const qty = valid.reduce((a, s) => a + s.quantity, 0);
  const cost = valid.reduce((a, s) => a + s.costCents, 0);
  return {
    tankGroup: group,
    segments,
    validSegments: valid.length,
    averageKmPerUnit: qty > 0 ? round(dist / qty, 2) : null,
    averageUnitsPer100Km: dist > 0 ? round((qty / dist) * 100, 2) : null,
    costPerKmCents: dist > 0 ? round(cost / dist, 2) : null,
    totalDistanceKm: dist,
  };
}

/**
 * Detecção robusta de consumo fora do padrão (z-score modificado de
 * Iglewicz & Hoaglin: |0,6745·(x − mediana)/MAD| > 3,5). Exige ao menos 5
 * trechos válidos. Se MAD = 0 (todos iguais), marca desvios acima de 20%.
 */
export function flagOutliers(segments: ConsumptionSegment[], minSamples = 5): void {
  const values = segments.map((s) => s.kmPerUnit).filter((v): v is number => v !== null);
  if (values.length < minSamples) return;
  const med = median(values);
  const mad = median(values.map((v) => Math.abs(v - med)));
  for (const s of segments) {
    if (s.kmPerUnit === null) continue;
    s.outlier = mad === 0
      ? Math.abs(s.kmPerUnit - med) / med > 0.2
      : Math.abs((0.6745 * (s.kmPerUnit - med)) / mad) > 3.5;
  }
}

export function median(values: readonly number[]): number {
  if (values.length === 0) throw new RangeError('Mediana de lista vazia.');
  const s = [...values].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid]! : (s[mid - 1]! + s[mid]!) / 2;
}

const round = (v: number, d: number): number => Math.round(v * 10 ** d) / 10 ** d;

// ---------------------------------------------------------------------------
// Energia elétrica
// ---------------------------------------------------------------------------
export interface ChargingInput {
  id: string;
  startedAt: string;
  odometerKm: number | null;
  energyKwh: number;
  costCents: Cents;
  socEndPct?: number | null;
}

export interface ElectricEfficiency {
  /** Sempre 'estimated': ver comentário do módulo. */
  confidence: 'estimated';
  distanceKm: number;
  energyKwh: number;
  kwhPer100Km: number;
  kmPerKwh: number;
  costPerKmCents: number;
  socCorrected: boolean;
  sessionsUsed: number;
}

export function estimateElectricEfficiency(
  sessions: readonly ChargingInput[],
  usableCapacityKwh?: number | null,
): ElectricEfficiency | null {
  const withOdo = [...sessions]
    .filter((s): s is ChargingInput & { odometerKm: number } => s.odometerKm !== null)
    .sort((a, b) => a.odometerKm - b.odometerKm || a.startedAt.localeCompare(b.startedAt));
  if (withOdo.length < 2) return null;

  const first = withOdo[0]!;
  const last = withOdo[withOdo.length - 1]!;
  const distance = last.odometerKm - first.odometerKm;
  if (distance <= 0) return null;

  // A energia da 1ª recarga repõe o que foi rodado ANTES do período; por isso
  // entram as recargas da 2ª até a última.
  const after = withOdo.slice(1);
  let energy = after.reduce((a, s) => a + s.energyKwh, 0);
  const cost = after.reduce((a, s) => a + s.costCents, 0);

  let socCorrected = false;
  if (usableCapacityKwh && usableCapacityKwh > 0 && first.socEndPct != null && last.socEndPct != null) {
    energy -= ((last.socEndPct - first.socEndPct) / 100) * usableCapacityKwh;
    socCorrected = true;
  }
  if (energy <= 0) return null;

  return {
    confidence: 'estimated',
    distanceKm: distance,
    energyKwh: round(energy, 2),
    kwhPer100Km: round((energy / distance) * 100, 1),
    kmPerKwh: round(distance / energy, 2),
    costPerKmCents: round(cost / distance, 2),
    socCorrected,
    sessionsUsed: withOdo.length,
  };
}
