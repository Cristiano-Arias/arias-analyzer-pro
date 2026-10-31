/**
 * Vencimento de manutenção por data e/ou quilometragem: vale o que ocorrer
 * primeiro. A previsão de data para itens por km usa a média diária de
 * rodagem do próprio veículo.
 */
import { addDays, addMonths, daysBetween, minDate, type IsoDate } from './dates.js';

export type IntervalSource = 'manufacturer_manual' | 'user' | 'generic_template';

export interface PlanItemInput {
  intervalKm?: number | null;
  intervalMonths?: number | null;
  intervalSource: IntervalSource;
  lastDoneOn?: IsoDate | null;
  lastDoneKm?: number | null;
  remindBeforeKm?: number;
  remindBeforeDays?: number;
}

export type DueStatus = 'ok' | 'due_soon' | 'overdue' | 'needs_baseline';

export interface DueResult {
  status: DueStatus;
  dueKm: number | null;
  dueDate: IsoDate | null;
  kmRemaining: number | null;
  daysRemaining: number | null;
  /** Data prevista em que o km será atingido (estimativa por média diária). */
  predictedDateForKm: IsoDate | null;
  /** Primeira data esperada considerando data e km. */
  effectiveDueDate: IsoDate | null;
  triggeredBy: Array<'km' | 'date'>;
  /** Avisar que o intervalo é genérico e deve ser conferido no manual. */
  genericIntervalWarning: boolean;
}

export function computeDue(
  item: PlanItemInput,
  ctx: { today: IsoDate; currentKm: number; avgKmPerDay?: number | null },
): DueResult {
  if (!item.intervalKm && !item.intervalMonths) {
    throw new RangeError('O item precisa de intervalo em km e/ou em meses.');
  }
  const remindKm = item.remindBeforeKm ?? 1000;
  const remindDays = item.remindBeforeDays ?? 15;
  const base: DueResult = {
    status: 'ok', dueKm: null, dueDate: null, kmRemaining: null, daysRemaining: null,
    predictedDateForKm: null, effectiveDueDate: null, triggeredBy: [],
    genericIntervalWarning: item.intervalSource === 'generic_template',
  };

  const kmKnown = !!item.intervalKm && item.lastDoneKm != null;
  const dateKnown = !!item.intervalMonths && !!item.lastDoneOn;
  // Sem ponto de partida não inventamos: pedimos ao usuário a última execução.
  if (!kmKnown && !dateKnown) return { ...base, status: 'needs_baseline' };

  const r = { ...base };
  let overdue = false;
  let soon = false;

  if (kmKnown) {
    r.dueKm = item.lastDoneKm! + item.intervalKm!;
    r.kmRemaining = r.dueKm - ctx.currentKm;
    if (r.kmRemaining <= 0) { overdue = true; r.triggeredBy.push('km'); }
    else if (r.kmRemaining <= remindKm) { soon = true; r.triggeredBy.push('km'); }
    if (ctx.avgKmPerDay && ctx.avgKmPerDay > 0) {
      r.predictedDateForKm = addDays(ctx.today, Math.max(0, Math.floor(r.kmRemaining / ctx.avgKmPerDay)));
    }
  }
  if (dateKnown) {
    r.dueDate = addMonths(item.lastDoneOn!, item.intervalMonths!);
    r.daysRemaining = daysBetween(ctx.today, r.dueDate);
    if (r.daysRemaining <= 0) { overdue = true; r.triggeredBy.push('date'); }
    else if (r.daysRemaining <= remindDays) { soon = true; r.triggeredBy.push('date'); }
  }
  // Previsão por km que cai dentro da janela de aviso também dispara aviso.
  if (!overdue && r.predictedDateForKm && daysBetween(ctx.today, r.predictedDateForKm) <= remindDays && !r.triggeredBy.includes('km')) {
    soon = true; r.triggeredBy.push('km');
  }

  const candidates = [r.dueDate, r.predictedDateForKm].filter((d): d is IsoDate => !!d);
  r.effectiveDueDate = candidates.length ? candidates.reduce(minDate) : null;
  r.status = overdue ? 'overdue' : soon ? 'due_soon' : 'ok';
  return r;
}

export interface OdometerReading { readOn: IsoDate; km: number }

/**
 * Média de km/dia a partir das leituras de odômetro dentro de uma janela
 * (padrão: 180 dias). Exige ao menos 2 leituras com 7+ dias de intervalo;
 * caso contrário, devolve null (sem estimativa, em vez de um chute).
 */
export function estimateKmPerDay(readings: readonly OdometerReading[], today: IsoDate, windowDays = 180): number | null {
  const from = addDays(today, -windowDays);
  const rs = readings.filter((r) => r.readOn >= from && r.readOn <= today)
    .sort((a, b) => a.readOn.localeCompare(b.readOn) || a.km - b.km);
  if (rs.length < 2) return null;
  const first = rs[0]!;
  const last = rs[rs.length - 1]!;
  const days = daysBetween(first.readOn, last.readOn);
  const km = last.km - first.km;
  if (days < 7 || km < 0) return null;
  return Math.round((km / days) * 10) / 10;
}
