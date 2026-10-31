import { describe, expect, it } from 'vitest';
import {
  addDays, addMonths, computeDue, costReport, daysBetween, documentStatus,
  estimateKmPerDay, parseIsoDate, planDocumentReminders,
} from '../src/index.js';

const today = '2026-10-31';

describe('datas civis', () => {
  it('soma meses preservando fim de mês e ano bissexto', () => {
    expect(addMonths('2026-01-31', 1)).toBe('2026-02-28');
    expect(addMonths('2024-01-31', 1)).toBe('2024-02-29');
    expect(addMonths('2026-11-15', 3)).toBe('2027-02-15');
    expect(addDays('2026-12-31', 1)).toBe('2027-01-01');
    expect(daysBetween('2026-10-31', '2026-11-07')).toBe(7);
  });
  it('rejeita datas malformadas ou inexistentes', () => {
    expect(() => parseIsoDate('31/10/2026')).toThrow(RangeError);
    expect(() => parseIsoDate('2026-02-30')).toThrow(RangeError);
  });
});

describe('vencimento de manutenção (o que vier primeiro)', () => {
  const oil = { intervalKm: 10_000, intervalMonths: 12, intervalSource: 'manufacturer_manual' as const, lastDoneOn: '2026-01-10', lastDoneKm: 20_000 };

  it('próximo do km: avisa e prevê a data pela média diária', () => {
    const r = computeDue(oil, { today, currentKm: 29_200, avgKmPerDay: 30 });
    expect(r.status).toBe('due_soon');
    expect(r.dueKm).toBe(30_000);
    expect(r.kmRemaining).toBe(800);
    expect(r.dueDate).toBe('2027-01-10');
    expect(r.daysRemaining).toBe(71);
    expect(r.predictedDateForKm).toBe('2026-11-26'); // 800/30 = 26 dias
    expect(r.effectiveDueDate).toBe('2026-11-26');
    expect(r.triggeredBy).toEqual(['km']);
    expect(r.genericIntervalWarning).toBe(false);
  });

  it('vencido por data mesmo rodando pouco', () => {
    const r = computeDue({ ...oil, lastDoneOn: '2025-10-01' }, { today, currentKm: 21_000 });
    expect(r.status).toBe('overdue');
    expect(r.triggeredBy).toEqual(['date']);
    expect(r.effectiveDueDate).toBe('2026-10-01');
  });

  it('vencido por km', () => {
    const r = computeDue(oil, { today, currentKm: 30_100 });
    expect(r.status).toBe('overdue');
    expect(r.kmRemaining).toBe(-100);
  });

  it('aviso antecipado pela previsão de km, mesmo longe do limite em km', () => {
    const r = computeDue({ ...oil, remindBeforeKm: 500 }, { today, currentKm: 27_000, avgKmPerDay: 300 });
    expect(r.status).toBe('due_soon');
    expect(r.predictedDateForKm).toBe('2026-11-10');
  });

  it('em dia', () => {
    const r = computeDue(oil, { today, currentKm: 22_000 });
    expect(r.status).toBe('ok');
    expect(r.predictedDateForKm).toBeNull();
    expect(r.effectiveDueDate).toBe('2027-01-10');
  });

  it('sem histórico: pede ponto de partida em vez de inventar', () => {
    const r = computeDue({ intervalKm: 10_000, intervalSource: 'user' }, { today, currentKm: 5_000 });
    expect(r.status).toBe('needs_baseline');
  });

  it('intervalo genérico gera aviso para conferir o manual', () => {
    const r = computeDue({ intervalMonths: 6, intervalSource: 'generic_template', lastDoneOn: '2026-10-20' }, { today, currentKm: 0 });
    expect(r.genericIntervalWarning).toBe(true);
    expect(r.status).toBe('ok');
  });

  it('exige ao menos um intervalo', () => {
    expect(() => computeDue({ intervalSource: 'user' }, { today, currentKm: 0 })).toThrow(RangeError);
  });
});

describe('média de km por dia', () => {
  it('calcula pela janela de leituras', () => {
    expect(estimateKmPerDay([{ readOn: '2026-09-01', km: 10_000 }, { readOn: '2026-10-31', km: 12_400 }], today)).toBe(40);
  });
  it('null com poucos dados ou intervalo curto', () => {
    expect(estimateKmPerDay([{ readOn: '2026-10-30', km: 1 }], today)).toBeNull();
    expect(estimateKmPerDay([{ readOn: '2026-10-28', km: 1 }, { readOn: '2026-10-30', km: 50 }], today)).toBeNull();
    expect(estimateKmPerDay([{ readOn: '2020-01-01', km: 1 }, { readOn: '2020-06-01', km: 5000 }], today)).toBeNull();
  });
});

describe('documentos', () => {
  it('situação de validade', () => {
    expect(documentStatus({ id: 'd', expiresOn: '2026-11-07' }, today)).toEqual({ status: 'expiring_soon', daysToExpiry: 7 });
    expect(documentStatus({ id: 'd', expiresOn: '2026-10-31' }, today).status).toBe('expiring_soon');
    expect(documentStatus({ id: 'd', expiresOn: '2026-10-30' }, today).status).toBe('expired');
    expect(documentStatus({ id: 'd', expiresOn: '2027-06-01' }, today).status).toBe('valid');
    expect(documentStatus({ id: 'd' }, today).status).toBe('no_expiry');
    expect(documentStatus({ id: 'd', expiresOn: '2026-12-15', reminderDays: [60] }, today).status).toBe('expiring_soon');
  });

  it('agenda apenas lembretes futuros, com chave idempotente', () => {
    const r = planDocumentReminders({ id: 'ipva', expiresOn: '2026-11-07' }, today);
    expect(r).toEqual([
      { dedupeKey: 'doc:ipva:2026-11-07:7d', sendOn: '2026-10-31', daysBefore: 7 },
      { dedupeKey: 'doc:ipva:2026-11-07:1d', sendOn: '2026-11-06', daysBefore: 1 },
    ]);
    expect(planDocumentReminders({ id: 'x', expiresOn: '2027-01-01', reminderDays: [10, 10, -1, 1.5] }, today)).toHaveLength(1);
    expect(planDocumentReminders({ id: 'x' }, today)).toEqual([]);
  });
});

describe('relatório de custos', () => {
  const expenses = [
    { occurredOn: '2026-01-05', amountCents: 20_000, category: 'fuel' },
    { occurredOn: '2026-01-20', amountCents: 50_000, category: 'maintenance' },
    { occurredOn: '2026-02-01', amountCents: 99_999, category: 'fuel' },
  ];
  it('total, por categoria e R$/km do período', () => {
    const r = costReport(expenses, [{ on: '2026-01-01', km: 10_000 }, { on: '2026-01-31', km: 11_000 }], '2026-01-01', '2026-01-31');
    expect(r.totalCents).toBe(70_000);
    expect(r.byCategory).toEqual({ fuel: 20_000, maintenance: 50_000 });
    expect(r.distanceKm).toBe(1000);
    expect(r.costPerKmCents).toBe(70);
    expect(r.reliability).toBe('ok');
  });
  it('sem distância conhecida: R$/km indisponível; distância curta: baixa confiabilidade', () => {
    expect(costReport(expenses, [], '2026-01-01', '2026-01-31')).toMatchObject({ costPerKmCents: null, reliability: 'unavailable' });
    expect(costReport(expenses, [{ on: '2026-01-02', km: 0 }, { on: '2026-01-03', km: 100 }], '2026-01-01', '2026-01-31').reliability).toBe('low');
    expect(() => costReport([], [], '2026-02-01', '2026-01-01')).toThrow(RangeError);
  });
});
