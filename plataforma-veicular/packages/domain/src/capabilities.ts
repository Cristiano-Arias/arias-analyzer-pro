/**
 * Matriz de capacidades: a partir do perfil do veículo, decide quais módulos,
 * campos, categorias de manutenção e widgets do dashboard se aplicam.
 *
 * Requisito do produto: "Não apresentar funcionalidades incompatíveis com a
 * tecnologia cadastrada". A interface consulta SOMENTE esta função; não há
 * `if (bev)` espalhado pelas telas.
 */
import {
  ENERGY_SOURCES, PROPULSION_TYPES, VEHICLE_CATEGORIES, isDiesel,
  type EnergyCode, type PropulsionCode, type VehicleCategoryCode,
} from './taxonomy.js';

export interface VehicleProfile {
  category: VehicleCategoryCode;
  /** null para veículos sem motor (reboque, semirreboque). */
  propulsion: PropulsionCode | null;
  /** Energias aceitas (flex = gasolina + etanol; kit GNV adiciona 'cng'). */
  energySources: readonly EnergyCode[];
  /**
   * Possui sistema SCR com Arla 32. Em pesados a diesel, assume-se `true` se
   * não for informado. Em leves a diesel depende do modelo, por isso o
   * usuário é perguntado (valor `undefined` => campo 'ask').
   */
  hasDefSystem?: boolean | undefined;
}

export type FieldVisibility = 'required' | 'optional' | 'hidden';

export type MaintenanceCategory =
  | 'engine_oil' | 'oil_filter' | 'air_filter' | 'cabin_filter' | 'fuel_filter'
  | 'spark_plugs' | 'timing_belt' | 'accessory_belt' | 'coolant' | 'brakes'
  | 'brake_fluid' | 'tires' | 'alignment_balancing' | 'suspension' | 'steering'
  | 'transmission' | 'exhaust' | 'battery_12v' | 'electrical' | 'body'
  | 'hv_battery' | 'battery_thermal_system' | 'ev_reduction_gear'
  | 'drive_chain' | 'def_system' | 'wheel_bearings' | 'cng_kit_inspection'
  | 'general_inspection';

export type DashboardWidget =
  | 'fuel_efficiency' | 'energy_efficiency' | 'charging_history' | 'battery_info'
  | 'maintenance_due' | 'documents_due' | 'costs' | 'cost_per_km'
  | 'trips' | 'tires' | 'def_consumption';

export interface VehicleCapabilities {
  fuelLogs: boolean;
  chargingLogs: boolean;
  defLogs: boolean;
  /** Para veículos sem motor, o odômetro é opcional (hodômetro de cubo). */
  odometer: FieldVisibility;
  fields: {
    fuelTankLiters: FieldVisibility;
    batteryCapacityKwh: FieldVisibility;
    transmission: FieldVisibility;
    engineDescription: FieldVisibility;
    axleConfiguration: FieldVisibility;
    hasDefSystem: FieldVisibility | 'ask';
  };
  /** Exibir alerta de segurança: não manipular sistema de alta tensão. */
  highVoltageSafetyNotice: boolean;
  /** Módulos avançados da Fase 5 relevantes para o perfil. */
  modules: { trips: boolean; tiresByAxle: boolean; batteryHealth: boolean };
  maintenanceCategories: MaintenanceCategory[];
  dashboard: DashboardWidget[];
}

export interface ProfileIssue { field: keyof VehicleProfile; message: string }

/** Valida a coerência do perfil. Retorna lista vazia se o perfil for válido. */
export function validateProfile(p: VehicleProfile): ProfileIssue[] {
  const issues: ProfileIssue[] = [];
  const cat = VEHICLE_CATEGORIES[p.category];
  if (!cat) return [{ field: 'category', message: 'Categoria desconhecida.' }];

  if (!cat.isPowered) {
    if (p.propulsion !== null) issues.push({ field: 'propulsion', message: `${cat.namePt} não possui propulsão própria.` });
    if (p.energySources.length > 0) issues.push({ field: 'energySources', message: `${cat.namePt} não consome energia de tração.` });
    return issues;
  }
  if (p.propulsion === null) {
    return [{ field: 'propulsion', message: 'Informe o tipo de propulsão.' }];
  }
  const prop = PROPULSION_TYPES[p.propulsion];
  const groups = new Set(p.energySources.map((e) => ENERGY_SOURCES[e].tankGroup));
  const hasElectric = groups.has('battery');
  const hasFuel = [...groups].some((g) => g !== 'battery');

  if (p.energySources.length === 0) issues.push({ field: 'energySources', message: 'Informe ao menos uma fonte de energia.' });
  if (prop.isPlugIn && !hasElectric) issues.push({ field: 'energySources', message: `${prop.namePt} deve incluir energia elétrica.` });
  if (!prop.isPlugIn && hasElectric) issues.push({ field: 'energySources', message: `${prop.namePt} não é recarregado na tomada; remova energia elétrica.` });
  if (prop.hasCombustion && !hasFuel) issues.push({ field: 'energySources', message: `${prop.namePt} exige ao menos um combustível.` });
  if (!prop.hasCombustion && hasFuel) issues.push({ field: 'energySources', message: `${prop.namePt} não usa combustível.` });
  if (groups.has('liquid_otto') && groups.has('liquid_diesel')) {
    issues.push({ field: 'energySources', message: 'Um mesmo motor não aceita diesel e gasolina/etanol.' });
  }
  if (p.hasDefSystem === true && !p.energySources.some(isDiesel)) {
    issues.push({ field: 'hasDefSystem', message: 'Arla 32 se aplica apenas a motores diesel.' });
  }
  return issues;
}

const ALWAYS: MaintenanceCategory[] = [
  'brakes', 'tires', 'suspension', 'electrical', 'body', 'general_inspection',
];

export function getCapabilities(p: VehicleProfile): VehicleCapabilities {
  const issues = validateProfile(p);
  if (issues.length > 0) {
    throw new RangeError(`Perfil de veículo inválido: ${issues.map((i) => i.message).join(' ')}`);
  }
  const cat = VEHICLE_CATEGORIES[p.category];
  const prop = p.propulsion ? PROPULSION_TYPES[p.propulsion] : null;
  const combustion = prop?.hasCombustion ?? false;
  const traction = prop?.hasTractionBattery ?? false;
  const plugIn = prop?.isPlugIn ?? false;
  const diesel = p.energySources.some(isDiesel);
  const gas = p.energySources.some((e) => ENERGY_SOURCES[e].tankGroup === 'gas');
  const otto = p.energySources.some((e) => ENERGY_SOURCES[e].tankGroup === 'liquid_otto');

  let def: boolean;
  let defField: FieldVisibility | 'ask';
  if (!diesel) { def = false; defField = 'hidden'; }
  else if (p.hasDefSystem !== undefined) { def = p.hasDefSystem; defField = 'optional'; }
  else if (cat.isHeavy) { def = true; defField = 'optional'; }
  else { def = false; defField = 'ask'; }

  const m = new Set<MaintenanceCategory>(ALWAYS);
  if (cat.isPowered) { m.add('battery_12v'); m.add('brake_fluid'); }
  if (!cat.isPowered || cat.isHeavy) m.add('wheel_bearings');
  if (cat.isPowered && !cat.isTwoWheeler) { m.add('steering'); m.add('cabin_filter'); m.add('alignment_balancing'); }
  if (cat.isTwoWheeler) m.add('drive_chain');
  if (combustion) {
    ['engine_oil', 'oil_filter', 'air_filter', 'fuel_filter', 'coolant', 'exhaust', 'transmission', 'accessory_belt', 'timing_belt']
      .forEach((c) => m.add(c as MaintenanceCategory));
    if (otto || gas) m.add('spark_plugs');
    if (gas) m.add('cng_kit_inspection');
  }
  if (traction && prop?.hasHighVoltage) { m.add('hv_battery'); m.add('battery_thermal_system'); m.add('coolant'); }
  if (prop?.code === 'bev') m.add('ev_reduction_gear');
  if (def) m.add('def_system');
  if (cat.isTwoWheeler) m.delete('cabin_filter');

  const dashboard: DashboardWidget[] = [];
  if (cat.isHeavy) dashboard.push('cost_per_km', 'trips', 'tires');
  if (combustion) dashboard.push('fuel_efficiency');
  if (plugIn) dashboard.push('energy_efficiency', 'charging_history', 'battery_info');
  if (def) dashboard.push('def_consumption');
  dashboard.push('maintenance_due', 'documents_due', 'costs');
  if (!cat.isHeavy) dashboard.push('cost_per_km');

  return {
    fuelLogs: combustion,
    chargingLogs: plugIn,
    defLogs: def,
    odometer: cat.isPowered ? 'required' : 'optional',
    fields: {
      fuelTankLiters: combustion ? 'optional' : 'hidden',
      batteryCapacityKwh: prop?.code === 'bev' ? 'required' : plugIn ? 'optional' : 'hidden',
      transmission: cat.isPowered ? 'optional' : 'hidden',
      engineDescription: combustion ? 'optional' : 'hidden',
      axleConfiguration: cat.isHeavy ? 'required' : 'hidden',
      hasDefSystem: defField,
    },
    highVoltageSafetyNotice: prop?.hasHighVoltage ?? false,
    modules: {
      trips: cat.isHeavy,
      tiresByAxle: cat.isHeavy,
      batteryHealth: plugIn || (prop?.hasHighVoltage ?? false),
    },
    maintenanceCategories: [...m].sort(),
    dashboard,
  };
}
