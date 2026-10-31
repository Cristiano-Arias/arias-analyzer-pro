# 02 — Requisitos

Os IDs seguem o padrão `RF-<módulo>-<n>` (funcional) e `RNF-<n>` (não funcional).
Prioridade (MoSCoW) aplicada ao **MVP**: **M** = obrigatório, **S** = deveria, **C** = poderia, **W** = fica para depois do MVP.

## 1. Requisitos funcionais

### Módulo 0 — Conta e acesso
| ID | Requisito | MVP |
|----|-----------|-----|
| RF-00-1 | Criar conta com e-mail e senha, com verificação de e-mail | M |
| RF-00-2 | Login, logout, logout de todos os dispositivos | M |
| RF-00-3 | Recuperação de senha por e-mail (token de uso único e curta duração) | M |
| RF-00-4 | Login com Google e Apple (Apple é obrigatório no iOS quando se oferece login social) | S |
| RF-00-5 | Consentimentos LGPD registrados com versão do termo e data/hora | M |
| RF-00-6 | Exportar todos os meus dados (JSON + arquivos) | M |
| RF-00-7 | Excluir conta (exclusão lógica imediata, física após o prazo legal ou de retenção) | M |
| RF-00-8 | Compartilhar veículo com outro usuário (papéis: dono, editor, leitor) | S |

### Módulo 1 — Garagem digital
| ID | Requisito | MVP |
|----|-----------|-----|
| RF-01-1 | Cadastrar vários veículos (categoria, tipo de carroceria, propulsão, marca, modelo, versão, ano fab./modelo) | M |
| RF-01-2 | Cadastro manual com campos adaptados à propulsão (tanque só se houver combustível; bateria só se for eletrificado) | M |
| RF-01-3 | Placa (formatos antigo e Mercosul, com validação) e VIN (17 caracteres, com validação de formato) | M |
| RF-01-4 | Odômetro atual, apelido, foto, data e valor de aquisição | M |
| RF-01-5 | Consulta automática por placa ou VIN | W — depende de fornecedor autorizado (ver riscos R-01) |
| RF-01-6 | Catálogo de marcas e modelos para autocompletar | S — origem dos dados a definir (R-02); com texto livre como alternativa |
| RF-01-7 | Arquivar ou vender veículo, mantendo o histórico | M |

### Módulo 2 — Manutenção
| ID | Requisito | MVP |
|----|-----------|-----|
| RF-02-1 | Registrar manutenção: data, km, categoria, descrição, oficina, responsável, itens (peça/serviço, quantidade, valor unitário), mão de obra, total, observações | M |
| RF-02-2 | Anexar nota fiscal e fotos | M |
| RF-02-3 | Garantia da manutenção (prazo em dias e/ou km) | S |
| RF-02-4 | Planos de manutenção por veículo: itens com intervalo em km e/ou meses, com **origem do intervalo** (manual do fabricante, usuário, modelo genérico) | M |
| RF-02-5 | Cálculo de “próxima manutenção”: o que vencer primeiro (data ou km), com previsão de data a partir da média de km/dia | M |
| RF-02-6 | Modelos genéricos de plano **identificados como genéricos**, com aviso para conferir o manual | S |
| RF-02-7 | Planos oficiais por fabricante, modelo e versão | W — depende de licenciamento de dados (R-03) |

### Módulo 3 — Documentos
| ID | Requisito | MVP |
|----|-----------|-----|
| RF-03-1 | Cadastrar documento (tipo, número, emissão, vencimento, valor, status de pagamento) vinculado a veículo ou pessoa (CNH) | M |
| RF-03-2 | Upload de PDF ou imagem com armazenamento privado | M |
| RF-03-3 | Alertas antecipados configuráveis (padrão: 30, 7 e 1 dia antes) | M |
| RF-03-4 | Histórico de documentos por veículo | M |
| RF-03-5 | OCR para preencher campos | W (Fase 5) |
| RF-03-6 | Consulta oficial de débitos, licenciamento e multas | W — sem integração validada; a interface fica pronta com aviso claro |

### Módulo 4 — Financeiro
| ID | Requisito | MVP |
|----|-----------|-----|
| RF-04-1 | Registrar despesa (categoria, valor, data, veículo, km opcional, forma de pagamento, anexo) | M |
| RF-04-2 | Manutenções e abastecimentos geram despesa automaticamente (sem lançamento duplicado) | M |
| RF-04-3 | Registrar receita (para P3/P4: corridas, fretes) | S |
| RF-04-4 | Relatório mensal e anual por veículo e por categoria | M |
| RF-04-5 | Custo por km (total e por categoria) | M |
| RF-04-6 | Comparação entre veículos | S |
| RF-04-7 | Orçamento mensal com alerta ao ultrapassar | C |
| RF-04-8 | Exportar CSV e PDF | S (CSV no MVP) |
| RF-04-9 | Despesas recorrentes ou parceladas (financiamento, seguro) | C |

### Módulo 5 — Abastecimento e consumo
| ID | Requisito | MVP |
|----|-----------|-----|
| RF-05-1 | Registrar abastecimento: data, km, combustível, litros (ou m³ para GNV), preço unitário, total, posto, tanque cheio (sim/não) | M |
| RF-05-2 | Consumo pelo método “tanque cheio a tanque cheio”, lidando com abastecimentos parciais e registros faltantes | M |
| RF-05-3 | Consumo médio e histórico; custo por km de combustível | M |
| RF-05-4 | Detecção de consumo fora do padrão (desvio robusto pela mediana) | S |
| RF-05-5 | Comparação de postos (preço médio pago) | C |
| RF-05-6 | Registro de recarga (kWh, local residencial/público, conector, potência, custo, SOC inicial/final opcional) | M para BEV/PHEV/EREV |

### Notificações, dashboard e histórico (transversais)
| ID | Requisito | MVP |
|----|-----------|-----|
| RF-NT-1 | Notificações push e e-mail para vencimentos de documentos e manutenção | M |
| RF-NT-2 | Central de notificações no app (lidas/não lidas) | M |
| RF-DB-1 | Dashboard por veículo adaptado à propulsão (combustão: consumo; elétrico: kWh/100 km e recargas) | M |
| RF-HI-1 | Linha do tempo unificada (manutenção, abastecimento, documentos, despesas) | M |
| RF-BU-1 | Busca global (veículos, registros, documentos) | S |
| RF-CA-1 | Calendário de vencimentos | S |

### Fase 5 em diante (fora do MVP)
EV/híbridos avançado (SOH, telemetria), caminhões, pneus, viagens, IA, OCR, serviços/marketplace, passaporte com link público, emergências e frotas multiempresa completas.
Os requisitos detalhados desses módulos estão no prompt-mestre e serão refinados no início de cada fase.

## 2. Requisitos não funcionais

| ID | Requisito | Meta |
|----|-----------|------|
| RNF-1 | Plataformas | Android 8+ (API 26+), iOS 16+, iPadOS, tablets Android; web nos navegadores modernos (2 últimas versões de Chrome, Edge, Firefox, Safari) |
| RNF-2 | Offline | Leitura e criação/edição de registros sem conexão no app móvel; sincronização automática ao reconectar |
| RNF-3 | Sincronização | Mudanças visíveis em outro dispositivo em ≤ 30 s com ambos online |
| RNF-4 | Desempenho | API p95 < 300 ms nas leituras comuns; app abre em < 2 s em aparelho intermediário (dados locais) |
| RNF-5 | Disponibilidade | 99,5% mensal no MVP; 99,9% no plano empresarial |
| RNF-6 | Segurança | OWASP ASVS nível 2 como referência; TLS 1.2+; senhas com Argon2id; dados isolados por usuário/organização com RLS no banco |
| RNF-7 | Privacidade | LGPD: base legal por finalidade, consentimento versionado, exportação, exclusão, registro de operações |
| RNF-8 | Acessibilidade | WCAG 2.1 AA; suporte a TalkBack/VoiceOver; fonte dinâmica |
| RNF-9 | Internacionalização | Interface em pt-BR; código preparado para i18n; valores em centavos (inteiros), moeda BRL; datas em UTC no servidor |
| RNF-10 | Observabilidade | Logs estruturados, métricas, rastreamento de erros (Sentry ou equivalente), *correlation id* por requisição |
| RNF-11 | Backups | Backup diário com PITR (recuperação para um ponto no tempo) de 7 dias; teste de restauração mensal |
| RNF-12 | Testes | Domínio ≥ 90% de cobertura; API com testes de integração em banco real; fluxos críticos com E2E |
| RNF-13 | Portabilidade | Serviços em contêiner (Docker), sem dependência que impeça migrar de provedor |
