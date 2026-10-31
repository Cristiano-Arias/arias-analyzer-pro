import { describe, expect, it } from 'vitest';
import {
  assertCents, formatBRL, legacyToMercosul, parsePlate, parseVin, samePlate,
  sourced, sumCents, totalFromUnitPrice,
} from '../src/index.js';

describe('placa', () => {
  it('reconhece padrão antigo e Mercosul, normalizando', () => {
    expect(parsePlate('abc-1234')).toEqual({ valid: true, normalized: 'ABC1234', format: 'legacy' });
    expect(parsePlate(' ABC 1D23 ')).toEqual({ valid: true, normalized: 'ABC1D23', format: 'mercosul' });
    expect(parsePlate('AB12345').valid).toBe(false);
    expect(parsePlate('ABCD123').valid).toBe(false);
  });
  it('converte antiga para Mercosul (5º dígito -> letra A–J)', () => {
    expect(legacyToMercosul('ABC1234')).toBe('ABC1C34');
    expect(legacyToMercosul('XYZ9087')).toBe('XYZ9A87');
    expect(() => legacyToMercosul('ABC1D23')).toThrow(RangeError);
  });
  it('identifica a mesma placa nos dois padrões', () => {
    expect(samePlate('ABC-1234', 'abc1c34')).toBe(true);
    expect(samePlate('ABC1234', 'ABC1D34')).toBe(false);
    expect(samePlate('inválida', 'inválida')).toBe(false);
  });
});

describe('VIN', () => {
  it('aceita VIN com formato válido e extrai o WMI', () => {
    expect(parseVin('9bwzzz377vt004251')).toEqual({ valid: true, normalized: '9BWZZZ377VT004251', wmi: '9BW' });
  });
  it('rejeita com motivo', () => {
    expect(parseVin('123').reason).toMatch(/17/);
    expect(parseVin('9BWZZZ377VT00425O').reason).toMatch(/I, O ou Q/);
    expect(parseVin('9BWZZZ377VT00425*').reason).toMatch(/letras e números/);
  });
});

describe('dinheiro', () => {
  it('total a partir de litros e preço em milésimos', () => {
    expect(totalFromUnitPrice(40, 5879)).toBe(23_516);
    // 37,452 L × R$ 6,199 = R$ 232,1649… -> R$ 232,16
    expect(totalFromUnitPrice(37.452, 6199)).toBe(23_216);
    // meio centavo arredonda para cima: 1 L × R$ 1,005 = R$ 1,01
    expect(totalFromUnitPrice(1, 1005)).toBe(101);
  });
  it('rejeita entradas inválidas', () => {
    expect(() => totalFromUnitPrice(0, 5879)).toThrow(RangeError);
    expect(() => totalFromUnitPrice(1, 5.5)).toThrow(RangeError);
    expect(() => assertCents(1.5)).toThrow(RangeError);
    expect(() => assertCents(-1)).toThrow(RangeError);
    expect(() => assertCents(100)).not.toThrow();
  });
  it('soma e formata em BRL', () => {
    expect(sumCents([100, 250])).toBe(350);
    expect(formatBRL(23_516).replace(/\s/g, ' ')).toBe('R$ 235,16');
  });
});

describe('procedência', () => {
  it('define confiança padrão pela origem', () => {
    expect(sourced(80, 'user', '2026-10-31T00:00:00Z').confidence).toBe('confirmed');
    expect(sourced(80, 'ocr', '2026-10-31T00:00:00Z').confidence).toBe('estimated');
    expect(sourced(80, 'derived', '2026-10-31T00:00:00Z').confidence).toBe('estimated');
    expect(sourced(80, 'obd', '2026-10-31T00:00:00Z', 'hypothesis').confidence).toBe('hypothesis');
  });
  it('OCR não pode ser marcado como confirmado sem revisão', () => {
    expect(() => sourced(1, 'ocr', '2026-10-31T00:00:00Z', 'confirmed')).toThrow(RangeError);
  });
});
