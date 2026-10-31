# Estratégia de sincronização (offline-first)

## Objetivo
- O app móvel funciona sem conexão: lê tudo e cria e edita registros.
- Ao reconectar, sincroniza sozinho.
- O usuário começa algo num dispositivo e continua em outro (meta: ≤ 30 s com ambos online).

## Modelo
- **Fonte da verdade:** PostgreSQL. Cada dispositivo guarda uma réplica parcial (apenas as contas e os veículos que o usuário pode acessar).
- **IDs gerados no cliente** (UUID v4 ou v7). Criar offline não exige ir ao servidor, e não há remapeamento de IDs.
- Cada tabela sincronizável tem `version`, `updated_at`, `deleted_at` e `server_seq`. O `server_seq` é atribuído por *trigger* a partir de uma **sequência global**, e não de relógio. Por isso o cursor não sofre com relógios dessincronizados.

## Protocolo

### Pull
```http
POST /v1/sync/pull
{ "cursor": 18233, "tables": ["vehicles","fuel_logs", ...], "limit": 500 }
→ 200
{ "changes": { "fuel_logs": { "upserted": [...], "deleted": ["<id>"] }, ... },
  "cursor": 18710, "hasMore": false }
```
O servidor devolve as linhas com `server_seq > cursor` visíveis ao usuário. O filtro é a RLS (`account_id`/compartilhamento), e não um parâmetro do cliente. Linhas com `deleted_at` vão como `deleted`. A paginação segue até `hasMore=false`.

> Detalhe importante: uma transação longa pode confirmar um `server_seq` menor **depois** que um pull já leu um maior. Para não perder linhas, o pull só devolve até o `server_seq` menor que o de qualquer transação ainda aberta. O servidor calcula esse limite com `pg_snapshot_xmin(pg_current_snapshot())` e uma tabela de controle por transação (*watermark*). Isso será implementado e testado no módulo `sync` (E-05).

### Push
```http
POST /v1/sync/push
{ "mutations": [
  { "table": "fuel_logs", "op": "upsert", "id": "...", "baseVersion": null, "data": {...} },
  { "table": "vehicles",  "op": "upsert", "id": "...", "baseVersion": 7, "data": { "nickname": "..." } },
  { "table": "documents", "op": "delete", "id": "...", "baseVersion": 3 } ] }
→ 200 { "results": [ { "id": "...", "status": "applied", "version": 8 },
                     { "id": "...", "status": "conflict", "server": {...} } ] }
```
- Cada mutação passa pelo **mesmo serviço de aplicação** da API online: validação com `@garagem/domain`, autorização e auditoria.
- É **idempotente**: reenviar a mesma mutação (mesmo `id` + `baseVersion`) não duplica nada.

## Resolução de conflitos
| Situação | Regra |
|----------|-------|
| Inserção de registro de log (abastecimento, despesa, recarga) | Sem conflito possível: o ID é único e gerado no cliente |
| Edição com `baseVersion` = versão atual | Aplica |
| Edição com `baseVersion` desatualizada | **Mescla por campo:** se os campos alterados pelo cliente não foram alterados no servidor desde `baseVersion` (consulta ao `audit_log`), aplica. Senão, devolve `conflict` com a versão do servidor, e o app mostra uma tela simples de escolha |
| Exclusão vs. edição | A exclusão vence, mas fica recuperável (exclusão lógica) por 30 dias |
| Odômetro menor que o último conhecido | Aceito como registro histórico, se a data for anterior. Se não, gera aviso de validação, não erro silencioso |

## Testes previstos
- Unitários do motor de mescla.
- Integração com dois "dispositivos" simulados contra Postgres real: edição concorrente, reenvio idempotente, exclusão concorrente, transação longa vs. pull (*watermark*).
- E2E: modo avião no emulador → criar 3 registros → reconectar → conferir no web.
