/**
 * Situação de documentos e agenda de lembretes.
 */
import { addDays, daysBetween, type IsoDate } from './dates.js';

export type DocumentStatus = 'no_expiry' | 'valid' | 'expiring_soon' | 'expired';

export interface DocumentInput {
  id: string;
  expiresOn?: IsoDate | null;
  /** Dias de antecedência para avisar (padrão 30, 7, 1). */
  reminderDays?: readonly number[];
  paymentStatus?: 'not_applicable' | 'pending' | 'paid' | 'overdue';
}

export const DEFAULT_REMINDER_DAYS = [30, 7, 1] as const;

export function documentStatus(doc: DocumentInput, today: IsoDate): { status: DocumentStatus; daysToExpiry: number | null } {
  if (!doc.expiresOn) return { status: 'no_expiry', daysToExpiry: null };
  const days = daysBetween(today, doc.expiresOn);
  const window = Math.max(...(doc.reminderDays?.length ? doc.reminderDays : DEFAULT_REMINDER_DAYS));
  // Vence no fim do dia: no próprio dia do vencimento ainda é válido.
  if (days < 0) return { status: 'expired', daysToExpiry: days };
  if (days <= window) return { status: 'expiring_soon', daysToExpiry: days };
  return { status: 'valid', daysToExpiry: days };
}

export interface ScheduledReminder {
  /** Chave idempotente: o agendador nunca envia o mesmo aviso duas vezes. */
  dedupeKey: string;
  sendOn: IsoDate;
  daysBefore: number;
}

/**
 * Lembretes futuros (a partir de hoje, inclusive) de um documento. Documentos
 * já pagos não geram aviso de pagamento, mas continuam gerando aviso de
 * validade (ex.: CNH não tem pagamento, mas vence).
 */
export function planDocumentReminders(doc: DocumentInput, today: IsoDate): ScheduledReminder[] {
  if (!doc.expiresOn) return [];
  const days = [...new Set(doc.reminderDays?.length ? doc.reminderDays : DEFAULT_REMINDER_DAYS)]
    .filter((d) => Number.isInteger(d) && d >= 0)
    .sort((a, b) => b - a);
  return days
    .map((d) => ({ dedupeKey: `doc:${doc.id}:${doc.expiresOn}:${d}d`, sendOn: addDays(doc.expiresOn!, -d), daysBefore: d }))
    .filter((r) => r.sendOn >= today);
}
