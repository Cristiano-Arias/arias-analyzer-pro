/**
 * Custo real do veículo: total, por categoria e por quilômetro num período.
 */
import type { Cents } from './money.js';
import type { IsoDate } from './dates.js';

export interface ExpenseInput {
  occurredOn: IsoDate;
  amountCents: Cents;
  category: string;
}

export interface CostReport {
  from: IsoDate;
  to: IsoDate;
  totalCents: Cents;
  byCategory: Record<string, Cents>;
  distanceKm: number | null;
  /** Centavos por km (fracionário); null se a distância for desconhecida. */
  costPerKmCents: number | null;
  /** 'low' quando a distância é pequena demais para um R$/km representativo. */
  reliability: 'ok' | 'low' | 'unavailable';
}

/**
 * @param odometerKms leituras de odômetro do período (de qualquer origem:
 *   abastecimento, manutenção, leitura manual).
 * @param minReliableKm distância mínima para considerar o R$/km confiável.
 */
export function costReport(
  expenses: readonly ExpenseInput[],
  odometerKms: readonly { on: IsoDate; km: number }[],
  from: IsoDate,
  to: IsoDate,
  minReliableKm = 300,
): CostReport {
  if (from > to) throw new RangeError('Período inválido: início após o fim.');
  const inRange = (d: IsoDate): boolean => d >= from && d <= to;

  const byCategory: Record<string, Cents> = {};
  let total = 0;
  for (const e of expenses) {
    if (!inRange(e.occurredOn)) continue;
    byCategory[e.category] = (byCategory[e.category] ?? 0) + e.amountCents;
    total += e.amountCents;
  }

  const kms = odometerKms.filter((o) => inRange(o.on)).map((o) => o.km);
  const distance = kms.length >= 2 ? Math.max(...kms) - Math.min(...kms) : null;
  const usable = distance !== null && distance > 0;
  return {
    from, to, totalCents: total, byCategory,
    distanceKm: distance,
    costPerKmCents: usable ? Math.round((total / distance!) * 100) / 100 : null,
    reliability: !usable ? 'unavailable' : distance! < minReliableKm ? 'low' : 'ok',
  };
}
