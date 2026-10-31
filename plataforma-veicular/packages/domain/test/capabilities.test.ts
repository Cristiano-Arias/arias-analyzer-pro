import { describe, expect, it } from 'vitest';
import { getCapabilities, validateProfile, type VehicleProfile } from '../src/index.js';

const flexCar: VehicleProfile = { category: 'car', propulsion: 'ice', energySources: ['gasoline', 'ethanol'] };
const bevCar: VehicleProfile = { category: 'car', propulsion: 'bev', energySources: ['electricity'] };

describe('matriz de capacidades', () => {
  it('carro flex: abastecimento sim, recarga e bateria não', () => {
    const c = getCapabilities(flexCar);
    expect(c.fuelLogs).toBe(true);
    expect(c.chargingLogs).toBe(false);
    expect(c.fields.batteryCapacityKwh).toBe('hidden');
    expect(c.fields.fuelTankLiters).toBe('optional');
    expect(c.maintenanceCategories).toContain('spark_plugs');
    expect(c.maintenanceCategories).toContain('engine_oil');
    expect(c.maintenanceCategories).not.toContain('hv_battery');
    expect(c.dashboard).toContain('fuel_efficiency');
    expect(c.dashboard).not.toContain('energy_efficiency');
    expect(c.highVoltageSafetyNotice).toBe(false);
  });

  it('elétrico (BEV): sem itens de motor a combustão, com bateria e aviso de alta tensão', () => {
    const c = getCapabilities(bevCar);
    expect(c.fuelLogs).toBe(false);
    expect(c.chargingLogs).toBe(true);
    expect(c.fields.batteryCapacityKwh).toBe('required');
    expect(c.fields.fuelTankLiters).toBe('hidden');
    expect(c.fields.engineDescription).toBe('hidden');
    for (const ice of ['engine_oil', 'spark_plugs', 'fuel_filter', 'exhaust', 'timing_belt'] as const) {
      expect(c.maintenanceCategories).not.toContain(ice);
    }
    expect(c.maintenanceCategories).toEqual(expect.arrayContaining(['hv_battery', 'battery_thermal_system', 'ev_reduction_gear', 'battery_12v']));
    expect(c.dashboard).toEqual(expect.arrayContaining(['energy_efficiency', 'charging_history', 'battery_info']));
    expect(c.dashboard).not.toContain('fuel_efficiency');
    expect(c.highVoltageSafetyNotice).toBe(true);
    expect(c.modules.batteryHealth).toBe(true);
  });

  it('PHEV: combustível e recarga; bateria opcional', () => {
    const c = getCapabilities({ category: 'car', propulsion: 'phev', energySources: ['gasoline', 'electricity'] });
    expect(c.fuelLogs && c.chargingLogs).toBe(true);
    expect(c.fields.batteryCapacityKwh).toBe('optional');
    expect(c.maintenanceCategories).toEqual(expect.arrayContaining(['engine_oil', 'hv_battery']));
    expect(c.maintenanceCategories).not.toContain('ev_reduction_gear');
  });

  it('MHEV (48 V): não é alta tensão e não recarrega na tomada', () => {
    const c = getCapabilities({ category: 'car', propulsion: 'mhev', energySources: ['gasoline'] });
    expect(c.highVoltageSafetyNotice).toBe(false);
    expect(c.chargingLogs).toBe(false);
    expect(c.maintenanceCategories).not.toContain('hv_battery');
  });

  it('HEV: alta tensão, sem recarga externa', () => {
    const c = getCapabilities({ category: 'car', propulsion: 'hev', energySources: ['gasoline'] });
    expect(c.highVoltageSafetyNotice).toBe(true);
    expect(c.chargingLogs).toBe(false);
    expect(c.modules.batteryHealth).toBe(true);
    expect(c.dashboard).not.toContain('charging_history');
  });

  it('caminhão pesado diesel: Arla 32 presumido, eixos obrigatórios, sem velas', () => {
    const c = getCapabilities({ category: 'truck_heavy', propulsion: 'ice', energySources: ['diesel_s10'] });
    expect(c.defLogs).toBe(true);
    expect(c.fields.axleConfiguration).toBe('required');
    expect(c.maintenanceCategories).toContain('def_system');
    expect(c.maintenanceCategories).toContain('wheel_bearings');
    expect(c.maintenanceCategories).not.toContain('spark_plugs');
    expect(c.dashboard.slice(0, 3)).toEqual(['cost_per_km', 'trips', 'tires']);
    expect(c.modules).toMatchObject({ trips: true, tiresByAxle: true });
  });

  it('caminhão diesel sem SCR informado explicitamente: sem Arla', () => {
    const c = getCapabilities({ category: 'truck_medium', propulsion: 'ice', energySources: ['diesel_s500'], hasDefSystem: false });
    expect(c.defLogs).toBe(false);
    expect(c.maintenanceCategories).not.toContain('def_system');
  });

  it('picape diesel leve: Arla depende do modelo => pergunta ao usuário', () => {
    const c = getCapabilities({ category: 'pickup', propulsion: 'ice', energySources: ['diesel_s10'] });
    expect(c.fields.hasDefSystem).toBe('ask');
    expect(c.defLogs).toBe(false);
    const withDef = getCapabilities({ category: 'pickup', propulsion: 'ice', energySources: ['diesel_s10'], hasDefSystem: true });
    expect(withDef.defLogs).toBe(true);
  });

  it('carro com kit GNV: velas e inspeção do kit', () => {
    const c = getCapabilities({ category: 'car', propulsion: 'ice', energySources: ['gasoline', 'ethanol', 'cng'] });
    expect(c.maintenanceCategories).toEqual(expect.arrayContaining(['spark_plugs', 'cng_kit_inspection']));
  });

  it('semirreboque: sem propulsão, odômetro opcional, sem itens de motor', () => {
    const c = getCapabilities({ category: 'semi_trailer', propulsion: null, energySources: [] });
    expect(c.odometer).toBe('optional');
    expect(c.fuelLogs || c.chargingLogs || c.defLogs).toBe(false);
    expect(c.maintenanceCategories).not.toContain('engine_oil');
    expect(c.maintenanceCategories).not.toContain('battery_12v');
    expect(c.maintenanceCategories).toEqual(expect.arrayContaining(['brakes', 'tires', 'wheel_bearings']));
    expect(c.fields.transmission).toBe('hidden');
  });

  it('motocicleta: corrente; sem filtro de cabine nem alinhamento de carro', () => {
    const c = getCapabilities({ category: 'motorcycle', propulsion: 'ice', energySources: ['gasoline'] });
    expect(c.maintenanceCategories).toContain('drive_chain');
    expect(c.maintenanceCategories).not.toContain('cabin_filter');
    expect(c.maintenanceCategories).not.toContain('alignment_balancing');
  });

  it('moto elétrica: recarga e bateria', () => {
    const c = getCapabilities({ category: 'motorcycle', propulsion: 'bev', energySources: ['electricity'] });
    expect(c.chargingLogs).toBe(true);
    expect(c.maintenanceCategories).toContain('drive_chain');
    expect(c.maintenanceCategories).not.toContain('engine_oil');
  });
});

describe('validação de perfil', () => {
  const cases: Array<[string, VehicleProfile]> = [
    ['BEV com gasolina', { category: 'car', propulsion: 'bev', energySources: ['electricity', 'gasoline'] }],
    ['PHEV sem eletricidade', { category: 'car', propulsion: 'phev', energySources: ['gasoline'] }],
    ['HEV com eletricidade', { category: 'car', propulsion: 'hev', energySources: ['gasoline', 'electricity'] }],
    ['diesel + gasolina', { category: 'car', propulsion: 'ice', energySources: ['diesel_s10', 'gasoline'] }],
    ['sem energia', { category: 'car', propulsion: 'ice', energySources: [] }],
    ['motorizado sem propulsão', { category: 'car', propulsion: null, energySources: [] }],
    ['reboque com propulsão', { category: 'trailer', propulsion: 'ice', energySources: ['diesel_s10'] }],
    ['reboque com energia', { category: 'trailer', propulsion: null, energySources: ['diesel_s10'] }],
    ['Arla em motor a gasolina', { category: 'car', propulsion: 'ice', energySources: ['gasoline'], hasDefSystem: true }],
    ['BEV sem combustível mas com eletricidade ausente', { category: 'car', propulsion: 'bev', energySources: ['gasoline'] }],
    ['categoria desconhecida', { category: 'xyz' as never, propulsion: 'ice', energySources: ['gasoline'] }],
  ];
  it.each(cases)('rejeita: %s', (_, p) => {
    expect(validateProfile(p).length).toBeGreaterThan(0);
    expect(() => getCapabilities(p)).toThrow(RangeError);
  });

  it('aceita perfis coerentes', () => {
    expect(validateProfile(flexCar)).toEqual([]);
    expect(validateProfile(bevCar)).toEqual([]);
    expect(validateProfile({ category: 'bus', propulsion: 'bev', energySources: ['electricity'] })).toEqual([]);
  });
});
