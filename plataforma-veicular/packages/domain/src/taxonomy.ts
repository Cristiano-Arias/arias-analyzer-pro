/**
 * Taxonomia de veículos, propulsões e fontes de energia.
 *
 * Espelha as tabelas de referência do banco (ref_vehicle_categories,
 * ref_propulsion_types, ref_energy_sources). As propriedades booleanas
 * descrevem características físicas; as regras de negócio (ver capabilities.ts)
 * derivam delas. Assim, uma categoria nova costuma exigir só uma entrada aqui.
 */

export type VehicleCategoryCode =
  | 'car' | 'motorcycle' | 'scooter' | 'pickup' | 'van' | 'light_commercial'
  | 'truck_light' | 'truck_medium' | 'truck_heavy' | 'truck_extra_heavy'
  | 'tractor_unit' | 'bus' | 'minibus' | 'semi_trailer' | 'trailer';

export interface VehicleCategory {
  readonly code: VehicleCategoryCode;
  readonly namePt: string;
  /** Possui motor próprio (reboques e semirreboques não possuem). */
  readonly isPowered: boolean;
  /** Veículo pesado: exige gestão por eixo, viagens e indicadores operacionais. */
  readonly isHeavy: boolean;
  /** Duas rodas: corrente/transmissão final, sem alinhamento convencional. */
  readonly isTwoWheeler: boolean;
}

export const VEHICLE_CATEGORIES: Readonly<Record<VehicleCategoryCode, VehicleCategory>> = {
  car:               { code: 'car', namePt: 'Automóvel de passeio', isPowered: true, isHeavy: false, isTwoWheeler: false },
  motorcycle:        { code: 'motorcycle', namePt: 'Motocicleta', isPowered: true, isHeavy: false, isTwoWheeler: true },
  scooter:           { code: 'scooter', namePt: 'Scooter / ciclomotor', isPowered: true, isHeavy: false, isTwoWheeler: true },
  pickup:            { code: 'pickup', namePt: 'Picape', isPowered: true, isHeavy: false, isTwoWheeler: false },
  van:               { code: 'van', namePt: 'Van / utilitário', isPowered: true, isHeavy: false, isTwoWheeler: false },
  light_commercial:  { code: 'light_commercial', namePt: 'Veículo comercial leve', isPowered: true, isHeavy: false, isTwoWheeler: false },
  truck_light:       { code: 'truck_light', namePt: 'Caminhão leve', isPowered: true, isHeavy: true, isTwoWheeler: false },
  truck_medium:      { code: 'truck_medium', namePt: 'Caminhão médio', isPowered: true, isHeavy: true, isTwoWheeler: false },
  truck_heavy:       { code: 'truck_heavy', namePt: 'Caminhão pesado', isPowered: true, isHeavy: true, isTwoWheeler: false },
  truck_extra_heavy: { code: 'truck_extra_heavy', namePt: 'Caminhão extrapesado', isPowered: true, isHeavy: true, isTwoWheeler: false },
  tractor_unit:      { code: 'tractor_unit', namePt: 'Cavalo mecânico', isPowered: true, isHeavy: true, isTwoWheeler: false },
  bus:               { code: 'bus', namePt: 'Ônibus', isPowered: true, isHeavy: true, isTwoWheeler: false },
  minibus:           { code: 'minibus', namePt: 'Micro-ônibus', isPowered: true, isHeavy: true, isTwoWheeler: false },
  semi_trailer:      { code: 'semi_trailer', namePt: 'Semirreboque', isPowered: false, isHeavy: true, isTwoWheeler: false },
  trailer:           { code: 'trailer', namePt: 'Reboque', isPowered: false, isHeavy: true, isTwoWheeler: false },
};

export type PropulsionCode = 'ice' | 'mhev' | 'hev' | 'phev' | 'erev' | 'bev';

export interface PropulsionType {
  readonly code: PropulsionCode;
  readonly namePt: string;
  readonly hasCombustion: boolean;
  readonly hasTractionBattery: boolean;
  readonly isPlugIn: boolean;
  /**
   * Sistema de tração acima de 60 V CC (faixa considerada de alta tensão).
   * MHEV costuma operar em 48 V e, por isso, fica fora dessa faixa.
   */
  readonly hasHighVoltage: boolean;
}

export const PROPULSION_TYPES: Readonly<Record<PropulsionCode, PropulsionType>> = {
  ice:  { code: 'ice',  namePt: 'Combustão interna',           hasCombustion: true,  hasTractionBattery: false, isPlugIn: false, hasHighVoltage: false },
  mhev: { code: 'mhev', namePt: 'Híbrido leve (MHEV)',         hasCombustion: true,  hasTractionBattery: true,  isPlugIn: false, hasHighVoltage: false },
  hev:  { code: 'hev',  namePt: 'Híbrido (HEV)',               hasCombustion: true,  hasTractionBattery: true,  isPlugIn: false, hasHighVoltage: true },
  phev: { code: 'phev', namePt: 'Híbrido plug-in (PHEV)',      hasCombustion: true,  hasTractionBattery: true,  isPlugIn: true,  hasHighVoltage: true },
  erev: { code: 'erev', namePt: 'Elétrico c/ extensor (EREV)', hasCombustion: true,  hasTractionBattery: true,  isPlugIn: true,  hasHighVoltage: true },
  bev:  { code: 'bev',  namePt: 'Elétrico a bateria (BEV)',    hasCombustion: false, hasTractionBattery: true,  isPlugIn: true,  hasHighVoltage: true },
};

export type EnergyCode =
  | 'gasoline' | 'gasoline_premium' | 'ethanol' | 'diesel_s10' | 'diesel_s500'
  | 'cng' | 'biomethane' | 'electricity';

export type EnergyUnit = 'L' | 'm3' | 'kWh';

/**
 * Grupo de reservatório: abastecimentos só podem ser comparados dentro do
 * mesmo grupo. Ex.: gasolina e etanol se misturam no mesmo tanque (flex),
 * mas o GNV fica num cilindro separado.
 */
export type TankGroup = 'liquid_otto' | 'liquid_diesel' | 'gas' | 'battery';

export interface EnergySource {
  readonly code: EnergyCode;
  readonly namePt: string;
  readonly unit: EnergyUnit;
  readonly tankGroup: TankGroup;
}

export const ENERGY_SOURCES: Readonly<Record<EnergyCode, EnergySource>> = {
  gasoline:         { code: 'gasoline', namePt: 'Gasolina', unit: 'L', tankGroup: 'liquid_otto' },
  gasoline_premium: { code: 'gasoline_premium', namePt: 'Gasolina aditivada/premium', unit: 'L', tankGroup: 'liquid_otto' },
  ethanol:          { code: 'ethanol', namePt: 'Etanol', unit: 'L', tankGroup: 'liquid_otto' },
  diesel_s10:       { code: 'diesel_s10', namePt: 'Diesel S10', unit: 'L', tankGroup: 'liquid_diesel' },
  diesel_s500:      { code: 'diesel_s500', namePt: 'Diesel S500', unit: 'L', tankGroup: 'liquid_diesel' },
  cng:              { code: 'cng', namePt: 'GNV', unit: 'm3', tankGroup: 'gas' },
  biomethane:       { code: 'biomethane', namePt: 'Biometano', unit: 'm3', tankGroup: 'gas' },
  electricity:      { code: 'electricity', namePt: 'Energia elétrica', unit: 'kWh', tankGroup: 'battery' },
};

export const isDiesel = (e: EnergyCode): boolean => ENERGY_SOURCES[e].tankGroup === 'liquid_diesel';
