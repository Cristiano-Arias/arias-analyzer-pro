# Garagem360 — Plataforma de gestão automotiva e veicular

> **Nome provisório** (ver tarefa D-01). Projeto isolado nesta pasta: não interfere na aplicação Flask existente na raiz do repositório.

**Status:** Fase 1 (descoberta e planejamento) concluída. As fundações da Fase 2 (domínio e banco) estão implementadas e testadas.

## Documentação

| Documento | Conteúdo |
|-----------|----------|
| [Visão do produto](docs/produto/01-visao-do-produto.md) | Problema, público, proposta de valor, diferenciais, concorrentes, métricas |
| [Requisitos](docs/produto/02-requisitos.md) | Funcionais (com prioridade MoSCoW) e não funcionais |
| [MVP e roadmap](docs/produto/03-mvp-e-roadmap.md) | O que entra e o que fica fora do MVP, fases, entregáveis |
| [Riscos](docs/produto/04-riscos.md) | Riscos técnicos e comerciais, com mitigação |
| [ADRs 0001–0008](docs/adr/0001-0008-decisoes-iniciais.md) | Escolhas de stack, comparadas e justificadas |
| [Arquitetura](docs/arquitetura/visao-geral.md) | Diagramas, camadas, estrutura, API, ambientes |
| [Modelo de dados](docs/arquitetura/modelo-de-dados.md) | Diagrama ER e decisões de modelagem |
| [Sincronização](docs/arquitetura/sincronizacao.md) | Protocolo offline-first e resolução de conflitos |
| [Segurança e LGPD](docs/arquitetura/seguranca-lgpd.md) | Autenticação, isolamento, LGPD, backups |

## Stack (resumo)
TypeScript ponta a ponta · **Expo/React Native** (Android, iOS, web) · **NestJS** (monólito modular) · **PostgreSQL 16** com RLS · storage S3-compatível · pg-boss (filas) · Expo Push.

## O que já funciona

### `packages/domain` — regras de negócio compartilhadas
Puro (sem I/O). Roda igual no app offline e na API.
- **Matriz de capacidades** por categoria e propulsão (combustão, flex, GNV, MHEV, HEV, PHEV, EREV, BEV, diesel com ou sem Arla, reboques). Define campos, módulos, itens de manutenção e widgets do dashboard.
- **Consumo** tanque cheio a tanque cheio (parciais, abastecimentos esquecidos, flex, GNV separado), média ponderada e detecção robusta de anomalias.
- **Eficiência elétrica** estimada (kWh/100 km), com correção por SOC, sempre marcada como estimativa.
- **Vencimento de manutenção** por data e/ou km (o que vier primeiro), com previsão pela média diária e aviso de intervalo genérico.
- **Documentos:** situação de validade e lembretes idempotentes.
- **Custos:** total, por categoria e R$/km com indicador de confiabilidade.
- Placa (antiga e Mercosul, com conversão), VIN, dinheiro em centavos e procedência de dados.

### `db/` — esquema do MVP
Multiempresa, RLS *fail-closed*, auditoria imutável, colunas de sincronização e restrições de domínio.

## Como rodar

```bash
cd plataforma-veicular
npm install
npm test            # testes do domínio + cobertura (limite mínimo: 90%)
npm run typecheck
# Banco (requer PostgreSQL 16 e um superusuário local):
./db/run-tests.sh   # cria um banco efêmero, aplica migrações e roda os testes SQL
```

### Última verificação (31/10/2026, neste ambiente)
- `npm test`: **62 testes aprovados**. Cobertura: 100% de linhas e funções, 98,7% de ramos.
- `npm run typecheck`: sem erros (TypeScript 5.9, modo `strict`).
- `./db/run-tests.sh` (PostgreSQL 16.15): **todos os testes SQL aprovados**: isolamento de leitura e escrita entre contas, *fail-closed*, compartilhamento só de leitura, versão e cursor de sincronização, auditoria com diff e protegida contra a aplicação, 7 restrições de domínio e extensibilidade de categoria.

## Limitações conhecidas (transparência)
- **Ainda não existem API, app nem telas.** São os próximos entregáveis (E-05 a E-07).
- Não há consulta por placa/VIN, consulta de débitos, OCR, IA nem telemetria. Nenhuma dessas funcionalidades é simulada (ver riscos R-01 a R-05).
- Builds iOS e Android não podem ser gerados nem testados neste ambiente de nuvem (não há SDKs mobile nem dispositivos). O plano é usar EAS Build e testes em dispositivos reais.
- Os limites dos planos na migração são provisórios. **Não há preços nem cobrança** (decisão D-02).
