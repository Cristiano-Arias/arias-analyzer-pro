# ADRs 0001–0008 — Decisões técnicas iniciais

Formato: contexto → alternativas → decisão → consequências. Status: **Aceita (revisável ao fim da F2)**.

Critérios usados em todas as comparações: custo, escalabilidade, segurança, manutenção, facilidade de desenvolvimento, integrações, cobertura multiplataforma e disponibilidade de profissionais no Brasil.

---

## ADR-0001 — TypeScript em todo o stack, em monorepo

**Contexto.** As mesmas regras de negócio (consumo, vencimento de manutenção, matriz de capacidades por propulsão) precisam rodar **no celular offline** e **no servidor**. Duplicar essas regras em duas linguagens gera divergência.

**Alternativas.** (a) Dart/Flutter no cliente e Python ou Node no servidor; (b) TypeScript ponta a ponta.

**Decisão.** TypeScript ponta a ponta, num monorepo (`npm workspaces`) com o pacote compartilhado `@garagem/domain`, que é puro, sem I/O e 100% testável.

**Consequências.** + Uma única implementação das regras, testada uma vez. + Maior base de profissionais no Brasil. − O cliente fica restrito ao ecossistema JS para mobile (ver ADR-0002).

## ADR-0002 — Apps: React Native + Expo para iOS, Android e web

| Critério | Flutter | React Native + Expo |
|----------|---------|---------------------|
| Desempenho de interface | Excelente | Muito bom (Nova Arquitetura) |
| Web | Funciona, mas com bundle pesado, acessibilidade e SEO mais fracos | react-native-web gera DOM real, acessível |
| Compartilhar código com o backend | Não (Dart) | **Sim (TS)** |
| Build iOS sem Mac local | Exige Mac ou CI com macOS | **EAS Build na nuvem** |
| Atualizações sem passar pela loja | Limitado | **EAS Update** (para JS, dentro das regras das lojas) |
| Profissionais no Brasil | Bom | **Muito bom** |

**Decisão.** Expo (SDK estável mais recente, com Expo Router) num **único app** para Android, iOS, iPadOS e web (layout adaptativo por *breakpoints*).
Se o painel de gestão de frotas (Fase 5) precisar de tabelas densas típicas de desktop, poderá virar um app web separado (Next.js) reaproveitando `@garagem/domain` e o cliente da API.

**Consequências.** − Dependemos do ciclo de versões do Expo. − Bibliotecas nativas fora do Expo exigem *config plugins*.

## ADR-0003 — Backend: monólito modular em NestJS + PostgreSQL

**Alternativas.**
- **Supabase (BaaS):** muito rápido para começar (autenticação, storage, RLS, realtime) e barato. Por outro lado, a lógica de negócio fica espalhada entre políticas SQL e *edge functions*, e a dependência do fornecedor é maior. Ainda é uma alternativa válida se o time for muito pequeno.
- **Firebase:** banco NoSQL, ruim para relatórios financeiros e frotas com consultas relacionais. Descartado como banco principal (pode ser usado só para FCM).
- **FastAPI (Python):** excelente, mas quebra o compartilhamento de código da ADR-0001.
- **Microsserviços:** complexidade operacional desnecessária no MVP.

**Decisão.** **Monólito modular NestJS** (módulos alinhados aos do produto: `auth`, `vehicles`, `maintenance`, `documents`, `finance`, `fuel`, `sync`, `notifications`), com **PostgreSQL 16** como fonte única da verdade e SQL versionado em `db/migrations`.
Cada módulo expõe apenas um serviço público; nenhum módulo acessa tabelas de outro diretamente. Assim, extrair um serviço depois é viável.

## ADR-0004 — API REST + OpenAPI; sincronização por pull/push

**Decisão.** REST (JSON) documentado com OpenAPI 3.1. O cliente TS é gerado a partir do contrato. GraphQL foi descartado no MVP porque as telas são previsíveis e o cache HTTP e o rate limiting ficam mais simples com REST.
Sincronização offline por dois endpoints (`POST /sync/pull` e `POST /sync/push`), descritos em `docs/arquitetura/sincronizacao.md`.

## ADR-0005 — Banco local no app: SQLite (mobile) / IndexedDB (web)

**Decisão.** Usar uma camada local com o protocolo de sincronização acima. Candidato principal: **WatermelonDB** (SQLite no mobile, LokiJS/IndexedDB na web), que é maduro e usa exatamente o modelo pull/push com *changes* por tabela.
Alternativas avaliadas: PowerSync e ElectricSQL (sincronização gerenciada direto do Postgres). Funcionam bem, mas adicionam um fornecedor e um custo; ficam como opção se a sincronização própria se mostrar cara de manter.
A escolha final do cliente local será confirmada com uma prova de conceito na F2. **O protocolo do servidor não depende dela.**

## ADR-0006 — Infraestrutura: contêineres em PaaS no MVP, AWS São Paulo na escala

**Contexto.** A LGPD não exige que os dados fiquem no Brasil, mas a transferência internacional precisa de base legal (cláusulas-padrão, Resolução ANPD 19/2024). Clientes empresariais costumam exigir dados no país.

**Decisão.**
- **MVP:** API em contêiner Docker numa PaaS (ex.: Render, Railway, Fly.io), Postgres gerenciado com PITR e storage S3-compatível (S3 ou Cloudflare R2).
- **Escala ou plano empresarial:** migrar para AWS `sa-east-1` (ECS Fargate + RDS + S3). Como tudo está em contêiner e em SQL padrão, a migração não exige reescrita.
- Filas: **pg-boss** (filas dentro do próprio Postgres), sem Redis no MVP.
- Push: Expo Push Service (FCM/APNs). E-mail transacional: provedor com DKIM/SPF (ex.: Amazon SES ou Resend).

## ADR-0007 — Autenticação própria com bibliotecas consolidadas

**Alternativas.** Serviços gerenciados (Auth0, Clerk, Firebase Auth, Supabase Auth) vs. módulo próprio.

**Decisão.** Módulo `auth` próprio, porque o custo por usuário ativo dos serviços gerenciados pesa num B2C com plano gratuito e o login de frota terá papéis específicos:
- senhas com **Argon2id**;
- *access token* JWT de 15 min + *refresh token* opaco, rotativo, com detecção de reuso (reuso revoga a família inteira de tokens), guardado como hash;
- sessões listáveis e revogáveis por dispositivo;
- login social (Google/Apple) via OIDC, com verificação de assinatura dos *ID tokens*.
A interface `IdentityProvider` permite trocar por um serviço gerenciado se a auditoria de segurança recomendar.

## ADR-0008 — Isolamento de dados: `owner` + `organization` + RLS

**Decisão.** Todo dado de negócio pertence a uma **conta-dona** (`account_id`). Uma conta é *pessoal* (um usuário) ou *organizacional* (uma empresa). Assim, a pessoa física e a frota usam o **mesmo modelo de dados**: a pessoa física é simplesmente uma conta com um único membro.
A autorização acontece em duas camadas: (1) *guards* na API; (2) **Row-Level Security no PostgreSQL**, usando `current_setting('app.account_ids')`, como defesa em profundidade contra bugs de autorização.
Implementado e testado em `db/migrations/0001_init.sql`.
