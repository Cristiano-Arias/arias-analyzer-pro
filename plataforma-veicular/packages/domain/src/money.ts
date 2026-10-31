/**
 * Dinheiro sempre em CENTAVOS inteiros (nunca float) para evitar erros de
 * arredondamento em somas. Preço unitário de combustível usa milésimos de
 * real (R$ 5,879/L = 5879), como nas bombas.
 */
export type Cents = number;

export function assertCents(v: number, label = 'valor'): asserts v is Cents {
  if (!Number.isSafeInteger(v) || v < 0) throw new RangeError(`${label} deve ser inteiro não negativo em centavos: ${v}`);
}

export const sumCents = (values: readonly Cents[]): Cents => values.reduce((a, b) => a + b, 0);

/** Arredondamento comercial (meio para cima) de um valor em centavos fracionário. */
export const roundCents = (v: number): Cents => Math.round(v + Number.EPSILON * Math.sign(v));

/** Total em centavos a partir de quantidade e preço em milésimos de real. */
export function totalFromUnitPrice(quantity: number, unitPriceMilli: number): Cents {
  if (!(quantity > 0) || !Number.isInteger(unitPriceMilli) || unitPriceMilli <= 0) {
    throw new RangeError('Quantidade e preço unitário devem ser positivos.');
  }
  return roundCents((quantity * unitPriceMilli) / 10);
}

const BRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
export const formatBRL = (cents: Cents): string => BRL.format(cents / 100);
