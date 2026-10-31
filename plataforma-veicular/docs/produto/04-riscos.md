# 04 — Riscos técnicos e comerciais

Escala: Probabilidade (P) e Impacto (I) de 1 a 3. Exposição = P × I.

## Riscos técnicos e de dados

| ID | Risco | P | I | Exp. | Mitigação |
|----|-------|---|---|------|-----------|
| R-01 | **Não existe API pública oficial de consulta de veículo por placa/VIN.** Os dados do RENAVAM/SENATRAN têm acesso restrito; fornecedores comerciais existem, mas com custo por consulta e exigências de LGPD | 3 | 2 | 6 | Cadastro manual excelente no MVP; interface `VehicleLookupProvider` pronta; contratar só após avaliação jurídica; **nunca preencher dados sem fonte** |
| R-02 | Catálogo de marcas, modelos e versões: não há fonte aberta oficial e completa. Bases como a tabela FIPE têm termos de uso a avaliar | 2 | 2 | 4 | Texto livre + autocompletar a partir do histórico da própria base; avaliar licenciamento antes de importar dados de terceiros |
| R-03 | **Planos de manutenção do fabricante não estão disponíveis em formato aberto** | 3 | 2 | 6 | O usuário informa o intervalo do manual; modelos genéricos marcados como “genérico — confira o manual”; campo `interval_source` obrigatório; licenciar base técnica no futuro |
| R-04 | Consultas oficiais (débitos, multas, licenciamento) sem API aberta | 3 | 2 | 6 | Fora do MVP; avaliar parceiros autorizados; a interface informa a limitação |
| R-05 | Dados de bateria (SOH/SOC) e telemetria dependem de APIs de montadoras (acesso restrito, por marca) ou de OBD-II (sem padrão para EV) | 3 | 2 | 6 | MVP usa dados informados pelo usuário, com procedência; arquitetura de *providers* com campo `source` e `confidence` |
| R-06 | Sincronização offline com conflitos (dois aparelhos editando o mesmo registro) | 2 | 3 | 6 | Protocolo pull/push com versão por linha; registros de log (abastecimento, despesa) são quase sempre só de inserção; para edição concorrente, vence a última escrita por campo, com auditoria; testes dedicados |
| R-07 | Complexidade multiplataforma (iOS, Android e web com uma equipe pequena) | 2 | 3 | 6 | Uma única base de interface (Expo + react-native-web) e TypeScript em todo o stack; builds na nuvem (EAS) |
| R-08 | Custos de IA crescendo sem controle | 2 | 2 | 4 | Fase 5; cotas por plano, cache, registro de tokens por usuário, camada de provedor intercambiável |
| R-09 | Vazamento de documentos (CNH, CRLV) e dados financeiros | 1 | 3 | 3 | Bucket privado com URLs assinadas de curta duração, RLS, criptografia em repouso, testes de autorização, pentest antes do lançamento |
| R-10 | Rejeição na App Store (login social sem Apple, exclusão de conta, permissões) | 2 | 2 | 4 | Sign in with Apple se houver login social; exclusão de conta dentro do app (exigida pelas lojas); justificativas de permissão claras |

## Riscos comerciais

| ID | Risco | P | I | Exp. | Mitigação |
|----|-------|---|---|------|-----------|
| C-01 | Mercado com apps gratuitos consolidados de registro de consumo | 3 | 2 | 6 | Diferenciar por documentos, multipropulsão, passaporte e custo real; importação de dados (CSV) de outros apps |
| C-02 | Baixo engajamento: o usuário esquece de registrar | 3 | 3 | 9 | Registro em 3 toques, lembretes úteis (não spam), widgets, OCR de cupom de posto (Fase 5) |
| C-03 | Disposição a pagar baixa no B2C | 2 | 3 | 6 | Receita principal em P3–P5 (profissional/empresarial); premium com valor claro (IA, relatórios, vários veículos) |
| C-04 | Escopo excessivo atrasando o lançamento | 3 | 3 | 9 | MVP rígido (doc 03), fases com critérios de saída, nenhuma funcionalidade nova sem métrica |
| C-05 | Riscos jurídicos (LGPD, Código de Defesa do Consumidor, termos) | 2 | 3 | 6 | Revisão jurídica obrigatória antes do lançamento (F7); encarregado (DPO) definido |
