# Report Meta Ads + Zoho CRM | PipeLovers

Dashboard de performance de criativos que cruza exportações do **Meta Ads** com o **Zoho CRM** para mostrar quais anúncios geram mais **leads**, **reuniões** e, principalmente, **assinaturas** — e a que custo. Prioridade de leitura: Assinaturas > Reuniões > Leads.

100% client-side e 100% estático — apenas HTML + CSS + JavaScript puro, com SheetJS e Chart.js via CDN. Sem backend, sem build step, sem variáveis de ambiente. Feito para rodar direto no **GitHub Pages**.

Visual em tema claro com banner gradiente roxo/navy no topo e navegação por abas (📁 Upload, 📊 Visão Geral, 🏆 Rankings & Gráficos, 📋 Tabela Detalhada), inspirado no painel de churn da PipeLovers.

## Estrutura do projeto

```
.
├── index.html      # estrutura da página (com navegação por abas)
├── style.css       # tema claro com banner gradiente
├── script.js       # leitura de planilhas, cruzamento, métricas, gráficos, abas, report do Slack
└── README.md
```

> **Arquivos que podem ser removidos:** `send-slack-report.js`, `vercel.json`, `.env.example` e `package.json` foram criados em uma versão anterior que dependia de um backend na Vercel. Como o projeto agora é 100% GitHub Pages, esses arquivos ficaram sem função — pode excluí-los do repositório com segurança. Veja "Enviando o report para o Slack" abaixo para entender como o envio funciona nesta versão.

## Como funciona

1. Faça upload da planilha exportada do **Meta Ads** e da planilha exportada do **Zoho CRM**.
2. Clique em **Atualizar Dashboard**. O script:
   - Detecta automaticamente a linha de cabeçalho real em cada planilha (útil quando o Zoho exporta linhas de lixo antes do cabeçalho).
   - Normaliza nomes de colunas, acentos, espaços e caracteres invisíveis.
   - Agrupa as linhas do Meta por criativo (campanha + conjunto + anúncio).
   - Cruza cada negócio do Zoho com o criativo correspondente: 1ª prioridade pelo nome do anúncio (`Nome do anúncio` ↔ `Meta Ads - Anuncio`), 2ª prioridade por `Meta Ads - ADs ID`, quando disponível.
   - Calcula métricas, rankings, badges e insights automáticos.
3. Use os filtros (data, campanha, conjunto de anúncios, criativo) na aba **📁 Upload** para refinar a visão — tudo recalcula em tempo real.
4. Navegue pelas abas **📊 Visão Geral**, **🏆 Rankings & Gráficos** e **📋 Tabela Detalhada** para ver cards, badges, insights, gráficos e a tabela completa. Ao atualizar o dashboard com sucesso, a aba "Visão Geral" abre automaticamente.
5. Exporte os dados filtrados em CSV ou XLSX a qualquer momento.
6. Clique em **Enviar Report** para gerar o texto do resumo (ver seção do Slack abaixo) — geração manual, nunca automática.

### Colunas esperadas

**Meta Ads:** Nome da campanha, Nome do conjunto de anúncios, Nome do anúncio, Impressões, Frequência, Valor usado (BRL), Cliques (todos), CPM, CTR, CPC, Leads, Data início, Data fim.

**Zoho CRM:** Nome Negócios, Nome Contato, Origem, Meta Ads - ADs ID, Meta Ads - Anuncio, Meta Ads - Campanha, Meta Ads - Lead ID, Meta Ads Campanha ID, ICP, Stage, Hora de Criação.

### Regras de negócio

- **Reunião Gerada**: `Stage` igual a "Reunião Agendada" ou "Reunião Realizada".
- **Assinatura Realizada**: `Stage` contém qualquer um dos termos configurados em `CONFIG.signupStageKeywords` (em `script.js`): `assinatura`, `assinado`, `fechado ganho`, `ganho`, `venda realizada`. Ajuste essa lista conforme o vocabulário do seu funil.
- **Fórmulas**: CPL Meta = Valor Gasto ÷ Leads Meta · Custo por Reunião = Valor Gasto ÷ Reuniões · Custo por Assinatura = Valor Gasto ÷ Assinaturas · Taxa Lead→Reunião = Reuniões ÷ Leads Zoho × 100 · Taxa Lead→Assinatura = Assinaturas ÷ Leads Zoho × 100.

## Rodando localmente

Não há build step. Basta servir os arquivos estáticos:

```bash
npx serve .
# ou
python3 -m http.server 8080
```

Abra o endereço indicado no terminal no navegador.

## Deploy (GitHub Pages)

1. Faça commit e push do projeto para o repositório.
2. Em **Settings → Pages**, selecione a branch (`main`) e a pasta raiz (`/`).
3. Acesse a URL gerada pelo GitHub Pages.

Não há mais nenhuma etapa de backend, variável de ambiente ou serviço externo a configurar — o site funciona inteiro a partir desses arquivos estáticos.

## Enviando o report para o Slack

Como o GitHub Pages só serve arquivos estáticos, o site não tem como guardar nenhum segredo (token, webhook) com segurança — qualquer coisa escrita em `script.js` fica visível a qualquer pessoa que abrir "Ver código-fonte" da página. Por isso o envio ao Slack não acontece automaticamente a partir do site. O fluxo é:

1. No dashboard, clique em **Enviar Report**.
2. O script monta o texto do resumo no formato oficial (📊 Report Meta Ads + Zoho, Resumo do Dia, Acumulado, Top Criativos, Insight do período) e copia para a área de transferência (ou baixa `report-slack.txt`, caso o navegador bloqueie o clipboard).
3. Cole esse texto no Slack manualmente, **ou** volte para a conversa com o Claude no Cowork e peça para ele enviar — o Claude já está conectado ao workspace do Slack (PipeLovers) e pode postar a mensagem direto no canal combinado usando esse conector, sem precisar de webhook nem de backend.

O botão mostra os estados "Gerando report...", "Report enviado com sucesso." (texto copiado/baixado) e "Erro ao enviar report." (se algo falhar ao montar o texto). O envio nunca é automático — só acontece quando alguém clica no botão.

## Atualizar Dashboard vs. Enviar Report

- **Atualizar Dashboard**: lê os dois arquivos selecionados, refaz todo o cruzamento e re-renderiza cards, gráficos, tabela, rankings e insights. Pode ser clicado quantas vezes for necessário, inclusive com arquivos diferentes.
- **Enviar Report**: usa os dados já processados na tela (resumo do dia + acumulado + top criativos + insight automático), monta o texto no formato do template e copia/baixa para você enviar ao Slack (manualmente ou via Claude), somente quando clicado.

## Personalização rápida

Os principais pontos de ajuste estão no topo de `script.js`, no objeto `CONFIG`:

```js
const CONFIG = {
  meetingStages: ["reuniao agendada", "reuniao realizada"],
  signupStageKeywords: ["assinatura", "assinado", "fechado ganho", "ganho", "venda realizada"],
  topN: 8,        // itens exibidos nos gráficos "Top N"
  rankingSize: 5, // itens exibidos em cada ranking
};
```
