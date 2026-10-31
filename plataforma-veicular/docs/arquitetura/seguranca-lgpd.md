# Segurança e privacidade (LGPD)

Referência: OWASP ASVS 4.0 nível 2 e OWASP Mobile Top 10.

## Autenticação e sessões
- Senhas com **Argon2id** (parâmetros recomendados pelo OWASP). Mínimo de 10 caracteres e verificação contra lista de senhas vazadas (k-anonimato).
- *Access token* JWT de 15 min (assinatura EdDSA, `kid` para rotação de chaves). *Refresh token* opaco de 256 bits, **rotativo**, guardado só como hash SHA-256. Reuso de token já rotacionado revoga a família inteira, o que indica roubo.
- No app: tokens em armazenamento seguro (Keychain/Keystore via `expo-secure-store`). Na web: *refresh token* em cookie `HttpOnly; Secure; SameSite=Strict`, e o *access token* só em memória.
- Recuperação de senha com token de uso único, 30 min de validade e resposta idêntica para e-mail existente ou não (evita enumeração de contas).
- Rate limiting: login com 5 tentativas a cada 15 min por conta e por IP, com atraso progressivo. API geral com limite por usuário e por IP.

## Autorização e isolamento
- **Camada 1:** *guards* da API verificam o papel do usuário na conta (`owner`, `admin`, `manager`, `driver`, `viewer`) e o compartilhamento do veículo.
- **Camada 2:** **RLS no PostgreSQL** (`FORCE ROW LEVEL SECURITY`). A API conecta com um papel sem privilégio de dono e define `app.user_id`, `app.account_ids` e `app.shared_vehicle_ids` por transação (`SET LOCAL`). Sem esse contexto, nenhuma linha é visível (*fail-closed*). **Testado** em `db/tests/0001_rls_sync_test.sql`.
- A tabela de auditoria só pode ser gravada por *trigger* `SECURITY DEFINER`. O papel da aplicação não consegue alterá-la nem apagá-la (testado).

## Dados e arquivos
- TLS 1.2+ em tudo. HSTS na web.
- Criptografia em repouso: disco do banco e do storage (padrão do provedor). Campos especialmente sensíveis, como o número da CNH, ganham criptografia de aplicação (envelope com KMS) na F4.
- Arquivos em **bucket privado**. Upload e download só por URL assinada de curta duração, com validação de tipo MIME e tamanho no servidor e varredura antivírus antes de marcar como `ready`.
- Logs nunca registram tokens, senhas, documentos nem conteúdo de arquivos (filtros no logger).

## LGPD
| Exigência | Implementação |
|-----------|---------------|
| Base legal por finalidade | Execução de contrato (funcionalidades principais), consentimento (marketing e processamento por IA), legítimo interesse (segurança e antifraude), registrados em `consents` com versão |
| Direitos do titular (art. 18) | Exportação (`POST /me/export`), correção (edição), exclusão (`DELETE /me`), revogação de consentimento |
| Minimização | Placa e VIN são opcionais; não coletamos CPF no MVP; localização só com ação explícita do usuário |
| IA (Fase 5) | Dados não são usados para treinar modelos. Provedores contratados sem retenção para treinamento; envio mínimo de contexto |
| Retenção | Conta excluída: anonimização imediata do perfil e exclusão física dos dados em até 30 dias, salvo obrigação legal |
| Incidentes | Plano de resposta com comunicação à ANPD e aos titulares nos prazos da Resolução CD/ANPD nº 15/2024 |
| Encarregado (DPO) | Definir antes do lançamento (tarefa jurídica) |

## Backups e recuperação de desastres
- PITR de 7 dias (MVP) e 30 dias (empresarial). Backup lógico diário em outra região, com criptografia.
- **RPO ≤ 5 min, RTO ≤ 4 h** (MVP). Teste de restauração mensal documentado.

## Processo
- Dependências verificadas no CI (`npm audit` e Dependabot), SAST e varredura de segredos.
- Pentest externo antes do lançamento público (F6a).
