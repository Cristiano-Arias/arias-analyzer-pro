/**
 * Procedência dos dados. Todo valor medido ou importado carrega a origem e o
 * grau de confiança, para que a interface e a IA distingam fatos confirmados
 * de estimativas (requisito: não simular dados de bateria/telemetria).
 */
export type DataSource =
  | 'user'              // digitado pelo usuário
  | 'ocr'               // extraído de documento/imagem (precisa de confirmação)
  | 'service_report'    // laudo/relatório de oficina
  | 'obd'               // leitor OBD-II
  | 'manufacturer_api'  // API oficial da montadora
  | 'telematics'        // plataforma de telemetria
  | 'charger_api'       // operador de recarga
  | 'derived';          // calculado pelo sistema a partir de outros dados

export type Confidence = 'confirmed' | 'estimated' | 'hypothesis';

export interface Sourced<T> {
  value: T;
  source: DataSource;
  confidence: Confidence;
  /** Instante ISO da obtenção. */
  observedAt: string;
}

/** Dados vindos de OCR nunca são "confirmados" até o usuário revisar. */
export function sourced<T>(value: T, source: DataSource, observedAt: string, confidence?: Confidence): Sourced<T> {
  const c: Confidence = confidence ?? (source === 'derived' ? 'estimated' : source === 'ocr' ? 'estimated' : 'confirmed');
  if (source === 'ocr' && c === 'confirmed') {
    throw new RangeError('Dado de OCR só pode ser confirmado após revisão do usuário (use source "user").');
  }
  return { value, source, confidence: c, observedAt };
}
