import { describe, expect, it } from 'vitest';
import {
  computeFuelConsumption, estimateElectricEfficiency, flagOutliers, median,
  type ConsumptionSegment, type FuelLogInput,
} from '../src/index.js';

let seq = 0;
const log = (odometerKm: number, quantity: number, opts: Partial<FuelLogInput> = {}): FuelLogInput => ({
  id: `f${++seq}`, odometerKm, quantity, filledAt: `2026-01-${String(10 + seq).padStart(2, '0')}T10:00:00Z`,
  energy: 'gasoline', totalCents: Math.round(quantity * 600), isFull: true, missedPrevious: false, ...opts,
});

describe('consumo tanque cheio a tanque cheio', () => {
  it('soma parciais, descarta trecho com abastecimento esquecido e pondera a média', () => {
    const logs = [
      log(10_000, 40),                         // início (cheio)
      log(10_300, 10, { isFull: false }),      // parcial
      log(10_500, 30),                         // fim: 500 km / 40 L = 12,5
      log(10_900, 35, { missedPrevious: true }),// trecho inválido
      log(11_300, 32),                         // 400 km / 32 L = 12,5
    ];
    const [s] = computeFuelConsumption([...logs].reverse()); // ordem de entrada não importa
    expect(s!.tankGroup).toBe('liquid_otto');
    expect(s!.segments).toHaveLength(3);
    expect(s!.segments[0]).toMatchObject({ distanceKm: 500, quantity: 40, kmPerUnit: 12.5 });
    expect(s!.segments[1]).toMatchObject({ kmPerUnit: null, invalidReason: 'missed_fill' });
    expect(s!.segments[2]).toMatchObject({ distanceKm: 400, quantity: 32, kmPerUnit: 12.5 });
    expect(s!.validSegments).toBe(2);
    expect(s!.averageKmPerUnit).toBe(12.5);
    expect(s!.averageUnitsPer100Km).toBe(8);
    expect(s!.totalDistanceKm).toBe(900);
    // custo: (40 + 32) L * R$ 6,00 = R$ 432,00 em 900 km = 48 centavos/km
    expect(s!.costPerKmCents).toBe(48);
  });

  it('flex: gasolina e etanol no mesmo trecho; GNV calculado separadamente', () => {
    const res = computeFuelConsumption([
      log(0, 40),
      log(400, 40, { energy: 'ethanol' }),
      log(100, 10, { energy: 'cng' }),
      log(300, 15, { energy: 'cng' }),
    ]);
    const liquid = res.find((r) => r.tankGroup === 'liquid_otto')!;
    const gas = res.find((r) => r.tankGroup === 'gas')!;
    expect(liquid.segments[0]!.energies).toEqual(['ethanol']);
    expect(liquid.averageKmPerUnit).toBe(10);
    expect(gas.averageKmPerUnit).toBe(13.33);
  });

  it('ignora parciais antes do primeiro tanque cheio e trecho em aberto', () => {
    const [s] = computeFuelConsumption([log(0, 20, { isFull: false }), log(100, 40), log(300, 15, { isFull: false })]);
    expect(s!.segments).toHaveLength(0);
    expect(s!.averageKmPerUnit).toBeNull();
    expect(s!.costPerKmCents).toBeNull();
  });

  it('marca odômetro que não avança como inválido', () => {
    const [s] = computeFuelConsumption([log(500, 40), log(500, 30)]);
    expect(s!.segments[0]!.invalidReason).toBe('non_increasing_odometer');
  });

  it('rejeita quantidade não positiva e ignora energia elétrica', () => {
    expect(() => computeFuelConsumption([log(0, 0)])).toThrow(RangeError);
    expect(computeFuelConsumption([log(0, 10, { energy: 'electricity' })])).toEqual([]);
  });
});

describe('detecção de consumo fora do padrão', () => {
  const seg = (v: number | null): ConsumptionSegment => ({
    startLogId: 'a', endLogId: 'b', tankGroup: 'liquid_otto', distanceKm: 1, quantity: 1, costCents: 0, kmPerUnit: v, energies: [],
  });
  it('marca o trecho anômalo pelo z-score modificado', () => {
    const s = [12, 12.2, 11.9, 12.1, 12, 6].map(seg);
    flagOutliers(s);
    expect(s.map((x) => x.outlier)).toEqual([false, false, false, false, false, true]);
  });
  it('MAD = 0: usa desvio relativo de 20%', () => {
    const s = [10, 10, 10, 10, 10, 13, null].map(seg);
    flagOutliers(s);
    expect(s[5]!.outlier).toBe(true);
    expect(s[0]!.outlier).toBe(false);
    expect(s[6]!.outlier).toBeUndefined();
  });
  it('não marca nada com menos de 5 amostras', () => {
    const s = [10, 30, 10].map(seg);
    flagOutliers(s);
    expect(s.every((x) => x.outlier === undefined)).toBe(true);
  });
  it('mediana', () => {
    expect(median([3, 1, 2])).toBe(2);
    expect(median([4, 1, 2, 3])).toBe(2.5);
    expect(() => median([])).toThrow();
  });
});

describe('eficiência elétrica (estimativa)', () => {
  const sessions = [
    { id: 'c1', startedAt: '2026-03-01T20:00:00Z', odometerKm: 1000, energyKwh: 30, costCents: 3000, socEndPct: 80 },
    { id: 'c2', startedAt: '2026-03-05T20:00:00Z', odometerKm: 1300, energyKwh: 45, costCents: 4500, socEndPct: 80 },
    { id: 'c3', startedAt: '2026-03-09T20:00:00Z', odometerKm: 1600, energyKwh: 40, costCents: 4000, socEndPct: 90 },
    { id: 'c4', startedAt: '2026-03-10T20:00:00Z', odometerKm: null, energyKwh: 5, costCents: 500 },
  ];
  it('ignora a 1ª recarga e sessões sem odômetro', () => {
    const e = estimateElectricEfficiency(sessions)!;
    expect(e.confidence).toBe('estimated');
    expect(e.distanceKm).toBe(600);
    expect(e.energyKwh).toBe(85);
    expect(e.kwhPer100Km).toBe(14.2);
    expect(e.costPerKmCents).toBe(14.17);
    expect(e.socCorrected).toBe(false);
    expect(e.sessionsUsed).toBe(3);
  });
  it('corrige pela variação de SOC quando a capacidade útil é conhecida', () => {
    const e = estimateElectricEfficiency(sessions, 60)!;
    expect(e.socCorrected).toBe(true);
    expect(e.energyKwh).toBe(79);             // 85 − (90−80)% × 60
    expect(e.kwhPer100Km).toBe(13.2);
    expect(e.kmPerKwh).toBe(7.59);
  });
  it('retorna null quando não há dados suficientes (não inventa)', () => {
    expect(estimateElectricEfficiency(sessions.slice(0, 1))).toBeNull();
    expect(estimateElectricEfficiency([sessions[0]!, { ...sessions[1]!, odometerKm: 1000 }])).toBeNull();
    expect(estimateElectricEfficiency([
      { ...sessions[0]!, socEndPct: 0 }, { ...sessions[1]!, energyKwh: 1, socEndPct: 100 },
    ], 60)).toBeNull();
  });
});
