# 01 — Visão do Produto

> **Nome provisório:** Garagem360. Antes de qualquer uso público, é preciso fazer uma busca de marca no INPI e verificar a disponibilidade de domínio e de nomes nas lojas de aplicativos.
> **Status:** Fase 1 — Descoberta. Documento vivo.

## 1. Problema

Quem tem um veículo no Brasil lida com muitas obrigações espalhadas em lugares diferentes:

- **Documentos e prazos** (IPVA, licenciamento, seguro, CNH, multas) ficam em sites de vários órgãos e no app oficial Carteira Digital de Trânsito, sem um calendário pessoal unificado.
- **O histórico de manutenção** fica em notas fiscais de papel, no WhatsApp da oficina ou na memória. Isso aumenta o risco de esquecer revisões, perder a garantia e vender o veículo por menos.
- **O custo real** de manter o veículo (R$/km) quase ninguém conhece. As despesas se misturam ao orçamento pessoal.
- **Veículos eletrificados** trazem dúvidas novas: saúde da bateria, custo da recarga e manutenção específica. Hoje existem poucas ferramentas neutras para isso, e menos ainda em português.
- **Caminhoneiros autônomos e pequenas frotas** usam planilhas. As soluções corporativas de telemetria são caras e foram pensadas para frotas grandes.

## 2. Público-alvo (segmentos em ordem de prioridade)

| # | Segmento | Dor principal | Plano provável |
|---|----------|---------------|----------------|
| P1 | **Proprietário pessoa física** (1–3 veículos, carro ou moto) | Esquecer prazos, não saber quanto gasta, perder histórico | Gratuito → Premium |
| P2 | **Proprietário de elétrico/híbrido** | Custo de recarga, autonomia, dúvidas sobre a bateria | Premium |
| P3 | **Motorista profissional** (aplicativo, táxi, motofrete) | Custo por km para calcular o lucro | Profissional |
| P4 | **Caminhoneiro autônomo / transportador pequeno** (1–10 veículos) | Custo por viagem, pneus, Arla 32, documentos de transporte | Profissional |
| P5 | **Empresas com frota** (10+ veículos) | Indicadores consolidados, motoristas, auditoria | Empresarial |

**Decisão:** o MVP se concentra em **P1 e P2, com a base pronta para P3**. Os segmentos P4 e P5 exigem funcionalidades pesadas (viagens, pneus por eixo, multiempresa com perfis) e vêm na Fase 5. **O modelo de dados, porém, já nasce multiempresa** (ver arquitetura), para não precisar de reescrita depois.

## 3. Proposta de valor

> “Um único aplicativo para cuidar de tudo relacionado ao seu veículo, simplificando sua rotina, reduzindo custos e ajudando você a tomar melhores decisões.”

Desdobrada em três promessas verificáveis:

1. **Nunca mais esquecer um prazo:** documentos e manutenções com lembretes por data **e** por quilometragem.
2. **Saber quanto o veículo custa de verdade:** R$/km por veículo e por categoria, calculado a partir de dados reais.
3. **Ter um histórico que valoriza o veículo:** passaporte digital exportável na hora da venda.

## 4. Diferenciais pretendidos

| Diferencial | Por que é defensável |
|-------------|----------------------|
| **Adaptação automática ao tipo de propulsão** (combustão, flex, GNV, MHEV, HEV, PHEV, EREV, BEV) | A maioria dos apps de consumo trata elétrico como exceção. Aqui existe uma *matriz de capacidades* no núcleo do domínio (já implementada no pacote `@garagem/domain`). |
| **Procedência explícita dos dados** (informado pelo usuário, OCR, OBD, API do fabricante, estimativa) | Gera confiança e atende ao requisito de “não simular dados”. Toda informação mostra de onde veio. |
| **Do carro ao caminhão na mesma plataforma** | Concorrentes costumam ser B2C (só consumo) ou B2B (telemetria cara). Atendemos a faixa intermediária. |
| **Assistente de IA com contexto do veículo** e distinção entre fato, estimativa e hipótese | Diferencial da Fase 5, com custo controlado por plano. |
| **Passaporte digital** compartilhável por link com escopo e prazo de validade | Ajuda na revenda e cria efeito de rede (o comprador conhece o app). |

## 5. Análise preliminar de concorrentes

> ⚠️ **Limitação:** esta análise se baseia em conhecimento público geral e **não foi validada com pesquisa de campo nem com testes atualizados dos apps**. Funcionalidades e preços mudam com frequência. A validação está prevista como tarefa da Fase 1 (ver `03-mvp-e-roadmap.md`, item D-04).

| Categoria | Exemplos | Pontos fortes | Lacunas que exploramos |
|-----------|----------|---------------|------------------------|
| Apps de consumo e despesas (B2C) | Drivvo, Fuelio, Simply Auto, Fuelly | Base grande de usuários, registro rápido de abastecimento | Pouco foco em documentos brasileiros, elétricos tratados de forma superficial, sem passaporte de venda |
| App governamental | Carteira Digital de Trânsito (SENATRAN) | Fonte oficial de CRLV-e, CNH-e e multas | Não cuida de manutenção, custos ou histórico. **É complemento, não concorrente.** Não há API pública aberta para terceiros. |
| Pagamento de débitos | Empresas de pagamento de IPVA/multas | Pagam débitos veiculares | Transacional, sem gestão. **Possíveis parceiros de integração.** |
| Telemetria e gestão de frota (B2B) | Cobli, Infleet, Sascar, Fleetio (EUA) | Rastreamento, telemetria, relatórios | Custo e hardware dedicado; pouco acessíveis para autônomos e frotas pequenas |
| Apps de montadoras | Apps conectados das marcas | Dados reais do veículo (SOC, localização) | Restritos a uma marca, sem visão financeira, sem multimarca |

## 6. Princípios do produto

1. **Útil sem integração:** tudo funciona com entrada manual bem projetada. Integrações aceleram, mas não são pré-requisito.
2. **Nunca inventar dado:** se não temos a informação, mostramos “não informado”, e não um valor plausível.
3. **Simples por padrão, profundo quando necessário:** campos avançados aparecem conforme o tipo de veículo e o perfil.
4. **Privacidade como recurso:** dados do usuário não treinam modelos; exportação e exclusão de conta disponíveis desde o MVP.
5. **Offline primeiro no celular:** registrar um abastecimento no posto sem sinal precisa funcionar.

## 7. Métricas de sucesso (North Star e de apoio)

- **North Star:** número de *veículos ativos por semana* (veículo com ao menos um registro ou interação com lembrete em 7 dias).
- Ativação: % de novas contas que cadastram veículo **e** um primeiro registro em 24 h (meta inicial: 40%).
- Retenção D30 ≥ 25% (referência típica de apps utilitários; a confirmar no beta).
- % de lembretes atendidos (manutenção ou documento registrado até 7 dias depois do alerta).
- Conversão gratuito → premium (meta a definir com a estratégia de preços — **não definida**).
