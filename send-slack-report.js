/**
 * IMPORTANTE: este arquivo deve ficar em "api/send-slack-report.js" dentro
 * do repositório (crie a pasta "api" na raiz do projeto e mova este arquivo
 * para dentro dela). Essa é a convenção de "Serverless Functions" da Vercel:
 * qualquer arquivo dentro de /api vira automaticamente uma rota
 * "/api/<nome-do-arquivo>".
 *
 * Endpoint: POST /api/send-slack-report
 *
 * Recebe o payload já calculado no front-end (resumo do dia, acumulado,
 * top criativos e insight do período) e envia para o Slack via Incoming
 * Webhook. A URL do webhook NUNCA é exposta ao cliente — fica apenas na
 * variável de ambiente SLACK_WEBHOOK_URL, configurada no servidor/host
 * (Vercel → Project Settings → Environment Variables).
 *
 * Compatível com a runtime de Vercel Functions (Node.js).
 */

function fmtCurrency(n) {
  if (n === null || n === undefined || !isFinite(n)) return "R$ 0,00";
  return Number(n).toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 2 });
}

function fmtInt(n) {
  if (n === null || n === undefined || !isFinite(n)) return "0";
  return Math.round(Number(n)).toLocaleString("pt-BR");
}

function buildSlackText(payload) {
  const { data, resumoDia, acumulado, topCriativos, insightPeriodo } = payload;

  return (
    `📊 *Report Meta Ads + Zoho* | PipeLovers\n` +
    `Data: ${data}\n\n` +
    `*Resumo do Dia*\n` +
    `• Valor gasto: ${fmtCurrency(resumoDia.valorGasto)}  ` +
    `• Leads Meta: ${fmtInt(resumoDia.leadsMeta)}  ` +
    `• Leads Zoho: ${fmtInt(resumoDia.leadsZoho)}  ` +
    `• Reuniões geradas: ${fmtInt(resumoDia.reunioes)}  ` +
    `• Assinaturas realizadas: ${fmtInt(resumoDia.assinaturas)}  ` +
    `• CPL Meta: ${fmtCurrency(resumoDia.cplMeta)}  ` +
    `• Custo por reunião: ${fmtCurrency(resumoDia.custoReuniao)}  ` +
    `• Custo por assinatura: ${fmtCurrency(resumoDia.custoAssinatura)}\n\n` +
    `*Acumulado*\n` +
    `• Investimento total: ${fmtCurrency(acumulado.valorGasto)}  ` +
    `• Leads Meta: ${fmtInt(acumulado.leadsMeta)}  ` +
    `• Leads Zoho: ${fmtInt(acumulado.leadsZoho)}  ` +
    `• Reuniões geradas: ${fmtInt(acumulado.reunioes)}  ` +
    `• Assinaturas realizadas: ${fmtInt(acumulado.assinaturas)}\n\n` +
    `*Top Criativos*\n` +
    `🥇 Mais assinaturas: ${topCriativos.assinaturas}\n` +
    `🥈 Mais reuniões: ${topCriativos.reunioes}\n` +
    `🥉 Mais leads: ${topCriativos.leads}\n\n` +
    `*Insight do período*\n` +
    `${insightPeriodo}`
  );
}

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.statusCode = 405;
    res.setHeader("Content-Type", "application/json");
    res.end(JSON.stringify({ error: "Método não permitido. Use POST." }));
    return;
  }

  const webhookUrl = process.env.SLACK_WEBHOOK_URL;
  if (!webhookUrl) {
    res.statusCode = 500;
    res.setHeader("Content-Type", "application/json");
    res.end(JSON.stringify({ error: "SLACK_WEBHOOK_URL não configurada no servidor." }));
    return;
  }

  try {
    let body = req.body;
    if (!body || typeof body === "string") {
      body = await new Promise((resolve, reject) => {
        let raw = "";
        req.on("data", (chunk) => (raw += chunk));
        req.on("end", () => {
          try {
            resolve(raw ? JSON.parse(raw) : {});
          } catch (e) {
            reject(e);
          }
        });
        req.on("error", reject);
      });
    }

    const text = buildSlackText(body);

    const slackRes = await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text }),
    });

    if (!slackRes.ok) {
      const errText = await slackRes.text().catch(() => "");
      throw new Error(`Slack respondeu com status ${slackRes.status}: ${errText}`);
    }

    res.statusCode = 200;
    res.setHeader("Content-Type", "application/json");
    res.end(JSON.stringify({ ok: true }));
  } catch (err) {
    console.error("Erro ao enviar report para o Slack:", err);
    res.statusCode = 500;
    res.setHeader("Content-Type", "application/json");
    res.end(JSON.stringify({ error: err.message || "Erro desconhecido ao enviar report." }));
  }
};
