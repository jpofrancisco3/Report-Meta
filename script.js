/* ==========================================================================
   Report Meta Ads + Zoho CRM | PipeLovers
   Pipeline 100% client-side: leitura de planilhas (SheetJS), cruzamento
   Meta Ads x Zoho CRM, métricas, rankings, insights, gráficos (Chart.js)
   e envio de report para o Slack via endpoint serverless /api/send-slack-report.
   ========================================================================== */

// ---------------------------------------------------------------------------
// CONFIGURAÇÃO (fácil de ajustar)
// ---------------------------------------------------------------------------
const CONFIG = {
  // Stages do Zoho que contam como "Reunião Gerada"
  meetingStages: ["reuniao agendada", "reuniao realizada"],
  // Stage do Zoho "CONTÉM" um destes termos -> conta como Assinatura
  signupStageKeywords: ["assinatura", "assinado", "fechado ganho", "ganho", "venda realizada"],
  // Quantos itens mostrar nos gráficos de "top N"
  topN: 8,
  // Quantos itens mostrar em cada ranking
  rankingSize: 5,
};

// ---------------------------------------------------------------------------
// ESTADO GLOBAL
// ---------------------------------------------------------------------------
const state = {
  metaRows: [],      // linhas brutas normalizadas da planilha Meta Ads
  zohoRows: [],       // linhas brutas normalizadas da planilha Zoho CRM
  creatives: [],       // dados agregados por criativo (após cruzamento)
  filtered: [],         // `creatives` após aplicar filtros globais
  sort: { key: "valorGasto", dir: "desc" },
  charts: {},
  metaFile: null,
  zohoFile: null,
};

// ---------------------------------------------------------------------------
// HELPERS DE NORMALIZACAO / TEXTO
// Os conjuntos de caracteres "invisiveis" e o range de marcas diacriticas
// combinantes (que sobra apos normalize("NFD")) sao montados a partir de
// CODE POINTS NUMERICOS (String.fromCharCode), e nao de caracteres colados
// diretamente no codigo-fonte. Isso evita qualquer ambiguidade de encoding.
// ---------------------------------------------------------------------------
function charFromCode(code) {
  return String.fromCharCode(code);
}
function buildCharRangeRegex(startCode, endCode) {
  return new RegExp("[" + charFromCode(startCode) + "-" + charFromCode(endCode) + "]", "g");
}
function buildCharSetRegex(codePoints, extraPattern) {
  const chars = codePoints.map(charFromCode).join("");
  return new RegExp("[" + chars + (extraPattern || "") + "]", "g");
}

// Combining Diacritical Marks: U+0300 a U+036F.
const COMBINING_DIACRITICS_RE = buildCharRangeRegex(0x0300, 0x036f);
// NBSP(0x00A0), espaco fino(0x2009), ZWSP(0x200B), ZWNJ(0x200C), ZWJ(0x200D), BOM/ZWNBSP(0xFEFF) + \s.
const INVISIBLE_CODE_POINTS = [0x00a0, 0x2009, 0x200b, 0x200c, 0x200d, 0xfeff];
const INVISIBLE_CHARS_RE = buildCharSetRegex(INVISIBLE_CODE_POINTS, "\\s");

function stripAccents(str) {
  return String(str).normalize("NFD").replace(COMBINING_DIACRITICS_RE, "");
}

function normalizeKey(str) {
  if (str === null || str === undefined) return "";
  return stripAccents(String(str))
    .replace(INVISIBLE_CHARS_RE, " ")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function cleanDisplay(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(INVISIBLE_CHARS_RE, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function toNumber(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isFinite(val) ? val : 0;
  let s = String(val).trim();
  // remove símbolos de moeda e espaços
  s = s.replace(/R\$\s?/gi, "").replace(/%/g, "").trim();
  // formato BR: 1.234,56  -> 1234.56  | formato US: 1,234.56 -> 1234.56
  const hasComma = s.includes(",");
  const hasDot = s.includes(".");
  if (hasComma && hasDot) {
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      s = s.replace(/,/g, "");
    }
  } else if (hasComma && !hasDot) {
    s = s.replace(",", ".");
  }
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function parseDateValue(val) {
  if (val === null || val === undefined || val === "") return null;
  if (val instanceof Date) return val;
  if (typeof val === "number") {
    // número de série do Excel
    const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(val) : null;
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, parsed.S || 0);
  }
  const s = String(val).trim();
  // dd/mm/yyyy ou dd-mm-yyyy
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = "20" + y;
    return new Date(Number(y), Number(mo) - 1, Number(d));
  }
  // yyyy-mm-dd
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    const [, y, mo, d] = m;
    return new Date(Number(y), Number(mo) - 1, Number(d));
  }
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

function fmtCurrency(n) {
  if (!isFinite(n)) return "—";
  return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 2 });
}
function fmtInt(n) {
  if (!isFinite(n)) return "—";
  return Math.round(n).toLocaleString("pt-BR");
}
function fmtPct(n) {
  if (!isFinite(n)) return "—";
  return n.toLocaleString("pt-BR", { maximumFractionDigits: 2 }) + "%";
}
function fmtRatio(n, suffix) {
  if (!isFinite(n) || n === 0) return "—";
  return n.toLocaleString("pt-BR", { maximumFractionDigits: 2 }) + (suffix || "");
}
function safeDiv(a, b) {
  if (!b) return NaN;
  return a / b;
}

// ---------------------------------------------------------------------------
// LEITURA DE PLANILHAS (SheetJS) + DETECÇÃO AUTOMÁTICA DE CABEÇALHO
// ---------------------------------------------------------------------------
async function readWorkbook(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  // matriz bruta (linha por linha), preservando linhas vazias/lixo antes do header
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  return matrix;
}

// Tenta localizar, dentre as primeiras N linhas, qual é a linha de cabeçalho real,
// comparando cada linha com um conjunto de palavras-chave esperadas.
function detectHeaderRow(matrix, expectedKeywords, maxScanRows = 15) {
  let bestRow = 0;
  let bestScore = -1;
  const limit = Math.min(matrix.length, maxScanRows);
  for (let i = 0; i < limit; i++) {
    const row = matrix[i] || [];
    const normalizedCells = row.map((c) => normalizeKey(c));
    let score = 0;
    expectedKeywords.forEach((kw) => {
      if (normalizedCells.some((c) => c.includes(kw))) score++;
    });
    // linhas com mais células não-vazias também ajudam a desempatar
    const nonEmpty = row.filter((c) => cleanDisplay(c) !== "").length;
    score += nonEmpty * 0.01;
    if (score > bestScore) {
      bestScore = score;
      bestRow = i;
    }
  }
  return bestRow;
}

function matrixToObjects(matrix, headerRowIdx) {
  const headerRow = matrix[headerRowIdx] || [];
  const headers = headerRow.map((h) => cleanDisplay(h));
  const rows = [];
  for (let i = headerRowIdx + 1; i < matrix.length; i++) {
    const raw = matrix[i];
    if (!raw || raw.every((c) => cleanDisplay(c) === "")) continue; // ignora linhas vazias
    const obj = {};
    headers.forEach((h, idx) => {
      if (!h) return;
      obj[h] = raw[idx];
    });
    rows.push(obj);
  }
  return { headers, rows };
}

// Mapeia, de forma flexível, o nome de cabeçalho real para uma chave lógica,
// procurando por aliases normalizados (sem acento, case-insensitive).
function buildHeaderMap(headers, aliasMap) {
  const map = {};
  const normalizedHeaders = headers.map((h) => ({ original: h, norm: normalizeKey(h) }));
  Object.entries(aliasMap).forEach(([logicalKey, aliases]) => {
    const found = normalizedHeaders.find((h) =>
      aliases.some((alias) => h.norm === normalizeKey(alias) || h.norm.includes(normalizeKey(alias)))
    );
    if (found) map[logicalKey] = found.original;
  });
  return map;
}

// ---------------------------------------------------------------------------
// PARSE META ADS
// ---------------------------------------------------------------------------
const META_ALIASES = {
  campanha: ["nome da campanha", "campanha"],
  conjunto: ["nome do conjunto de anuncios", "conjunto de anuncios", "ad set name"],
  anuncio: ["nome do anuncio", "anuncio", "ad name"],
  impressoes: ["impressoes", "impressions"],
  frequencia: ["frequencia", "frequency"],
  valorUsado: ["valor usado (brl)", "valor usado", "amount spent"],
  cliques: ["cliques (todos)", "cliques", "clicks (all)", "link clicks"],
  cpm: ["cpm"],
  ctr: ["ctr"],
  cpc: ["cpc"],
  leads: ["leads", "resultados", "results"],
  dataInicio: ["data inicio", "data de inicio", "reporting starts", "start date"],
  dataFim: ["data fim", "data de termino", "reporting ends", "end date"],
  adId: ["id do anuncio", "ad id", "identificacao do anuncio"],
};

function parseMetaAds(matrix) {
  const headerIdx = detectHeaderRow(matrix, [
    "nome da campanha", "nome do anuncio", "impressoes", "cliques", "leads",
  ]);
  const { headers, rows } = matrixToObjects(matrix, headerIdx);
  const map = buildHeaderMap(headers, META_ALIASES);

  return rows
    .map((r) => {
      const anuncio = cleanDisplay(r[map.anuncio]);
      if (!anuncio) return null;
      return {
        campanha: cleanDisplay(r[map.campanha]) || "(sem campanha)",
        conjunto: cleanDisplay(r[map.conjunto]) || "(sem conjunto)",
        anuncio,
        anuncioKey: normalizeKey(anuncio),
        impressoes: toNumber(r[map.impressoes]),
        frequencia: toNumber(r[map.frequencia]),
        valorGasto: toNumber(r[map.valorUsado]),
        cliques: toNumber(r[map.cliques]),
        leads: toNumber(r[map.leads]),
        dataInicio: map.dataInicio ? parseDateValue(r[map.dataInicio]) : null,
        dataFim: map.dataFim ? parseDateValue(r[map.dataFim]) : null,
        adId: map.adId ? normalizeKey(r[map.adId]) : "",
      };
    })
    .filter(Boolean);
}

// ---------------------------------------------------------------------------
// PARSE ZOHO CRM
// ---------------------------------------------------------------------------
const ZOHO_ALIASES = {
  nomeNegocio: ["nome negocios", "nome do negocio", "deal name"],
  nomeContato: ["nome contato", "nome do contato", "contact name"],
  origem: ["origem", "lead source"],
  metaAdsId: ["meta ads - ads id", "meta ads ads id", "ad id"],
  metaAdsAnuncio: ["meta ads - anuncio", "meta ads anuncio"],
  metaAdsCampanha: ["meta ads - campanha", "meta ads campanha"],
  metaAdsLeadId: ["meta ads - lead id", "meta ads lead id"],
  metaAdsCampanhaId: ["meta ads campanha id", "meta ads - campanha id"],
  icp: ["icp"],
  stage: ["stage", "estagio", "etapa"],
  horaCriacao: ["hora de criacao", "data de criacao", "created time"],
};

function parseZohoCRM(matrix) {
  const headerIdx = detectHeaderRow(matrix, [
    "nome negocios", "stage", "origem", "hora de criacao", "meta ads",
  ]);
  const { headers, rows } = matrixToObjects(matrix, headerIdx);
  const map = buildHeaderMap(headers, ZOHO_ALIASES);

  return rows
    .map((r) => {
      const stage = cleanDisplay(r[map.stage]);
      const metaAdsAnuncio = map.metaAdsAnuncio ? cleanDisplay(r[map.metaAdsAnuncio]) : "";
      if (!stage && !metaAdsAnuncio) return null;
      return {
        nomeNegocio: cleanDisplay(r[map.nomeNegocio]),
        nomeContato: cleanDisplay(r[map.nomeContato]),
        origem: cleanDisplay(r[map.origem]),
        metaAdsId: map.metaAdsId ? normalizeKey(r[map.metaAdsId]) : "",
        metaAdsAnuncio,
        metaAdsAnuncioKey: normalizeKey(metaAdsAnuncio),
        metaAdsCampanha: map.metaAdsCampanha ? cleanDisplay(r[map.metaAdsCampanha]) : "",
        metaAdsLeadId: map.metaAdsLeadId ? cleanDisplay(r[map.metaAdsLeadId]) : "",
        metaAdsCampanhaId: map.metaAdsCampanhaId ? cleanDisplay(r[map.metaAdsCampanhaId]) : "",
        icp: map.icp ? cleanDisplay(r[map.icp]) : "",
        stage,
        stageKey: normalizeKey(stage),
        horaCriacao: map.horaCriacao ? parseDateValue(r[map.horaCriacao]) : null,
      };
    })
    .filter(Boolean);
}

// ---------------------------------------------------------------------------
// REGRAS DE STAGE
// ---------------------------------------------------------------------------
function isMeetingStage(stageKey) {
  return CONFIG.meetingStages.some((s) => stageKey === normalizeKey(s));
}
function isSignupStage(stageKey) {
  return CONFIG.signupStageKeywords.some((kw) => stageKey.includes(normalizeKey(kw)));
}

// ---------------------------------------------------------------------------
// CRUZAMENTO META ADS x ZOHO CRM
// ---------------------------------------------------------------------------
function buildCreatives(metaRows, zohoRows) {
  // Agrupa linhas do Meta por criativo (campanha + conjunto + anuncio), somando métricas
  const groups = new Map();
  metaRows.forEach((row) => {
    const key = `${normalizeKey(row.campanha)}|${normalizeKey(row.conjunto)}|${row.anuncioKey}`;
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        anuncio: row.anuncio,
        anuncioKey: row.anuncioKey,
        campanha: row.campanha,
        conjunto: row.conjunto,
        adId: row.adId,
        impressoes: 0,
        cliques: 0,
        valorGasto: 0,
        leadsMeta: 0,
        dataInicio: null,
        dataFim: null,
        zohoDeals: [],
      });
    }
    const g = groups.get(key);
    g.impressoes += row.impressoes;
    g.cliques += row.cliques;
    g.valorGasto += row.valorGasto;
    g.leadsMeta += row.leads;
    if (row.adId) g.adId = row.adId;
    if (row.dataInicio && (!g.dataInicio || row.dataInicio < g.dataInicio)) g.dataInicio = row.dataInicio;
    if (row.dataFim && (!g.dataFim || row.dataFim > g.dataFim)) g.dataFim = row.dataFim;
  });

  const creatives = Array.from(groups.values());

  // Índices para cruzamento: por nome de anúncio normalizado e por ID (quando existir)
  const byAnuncioKey = new Map();
  const byAdId = new Map();
  creatives.forEach((c) => {
    if (!byAnuncioKey.has(c.anuncioKey)) byAnuncioKey.set(c.anuncioKey, []);
    byAnuncioKey.get(c.anuncioKey).push(c);
    if (c.adId) {
      if (!byAdId.has(c.adId)) byAdId.set(c.adId, []);
      byAdId.get(c.adId).push(c);
    }
  });

  const unmatchedZoho = [];

  zohoRows.forEach((deal) => {
    let matches = [];
    // Prioridade 1: Nome do anúncio (Meta) == Meta Ads - Anuncio (Zoho)
    if (deal.metaAdsAnuncioKey && byAnuncioKey.has(deal.metaAdsAnuncioKey)) {
      matches = byAnuncioKey.get(deal.metaAdsAnuncioKey);
    }
    // Prioridade 2: Meta Ads - ADs ID, quando existir dos dois lados
    if (matches.length === 0 && deal.metaAdsId && byAdId.has(deal.metaAdsId)) {
      matches = byAdId.get(deal.metaAdsId);
    }
    if (matches.length > 0) {
      // se houver mais de um criativo com o mesmo nome (ex: campanhas diferentes),
      // distribui o negócio para todos eles é incorreto — usamos o primeiro,
      // priorizando o que também combina pela campanha informada no Zoho, se houver.
      let target = matches[0];
      if (matches.length > 1 && deal.metaAdsCampanha) {
        const byCampanha = matches.find((m) => normalizeKey(m.campanha) === normalizeKey(deal.metaAdsCampanha));
        if (byCampanha) target = byCampanha;
      }
      target.zohoDeals.push(deal);
    } else {
      unmatchedZoho.push(deal);
    }
  });

  // Calcula métricas derivadas por criativo
  creatives.forEach((c) => {
    const deals = c.zohoDeals;
    c.leadsZoho = deals.length;
    c.reunioes = deals.filter((d) => isMeetingStage(d.stageKey)).length;
    c.assinaturas = deals.filter((d) => isSignupStage(d.stageKey)).length;
    c.ctr = safeDiv(c.cliques, c.impressoes) * 100;
    c.cpc = safeDiv(c.valorGasto, c.cliques);
    c.cpm = safeDiv(c.valorGasto, c.impressoes) * 1000;
    c.cplMeta = safeDiv(c.valorGasto, c.leadsMeta);
    c.custoReuniao = safeDiv(c.valorGasto, c.reunioes);
    c.custoAssinatura = safeDiv(c.valorGasto, c.assinaturas);
    c.taxaReuniao = safeDiv(c.reunioes, c.leadsZoho) * 100;
    c.taxaAssinatura = safeDiv(c.assinaturas, c.leadsZoho) * 100;
  });

  return { creatives, unmatchedZoho };
}

// ---------------------------------------------------------------------------
// FILTROS GLOBAIS
// ---------------------------------------------------------------------------
function populateFilterOptions(creatives) {
  const campanhaSel = document.getElementById("filterCampanha");
  const conjuntoSel = document.getElementById("filterConjunto");
  const campanhas = [...new Set(creatives.map((c) => c.campanha))].sort();
  const conjuntos = [...new Set(creatives.map((c) => c.conjunto))].sort();

  campanhaSel.innerHTML = '<option value="">Todas</option>' + campanhas.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
  conjuntoSel.innerHTML = '<option value="">Todos</option>' + conjuntos.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
}

function applyFilters() {
  const start = document.getElementById("filterStart").value ? new Date(document.getElementById("filterStart").value) : null;
  const end = document.getElementById("filterEnd").value ? new Date(document.getElementById("filterEnd").value) : null;
  const campanha = document.getElementById("filterCampanha").value;
  const conjunto = document.getElementById("filterConjunto").value;
  const criativoSearch = normalizeKey(document.getElementById("filterCriativo").value);

  state.filtered = state.creatives.filter((c) => {
    if (campanha && c.campanha !== campanha) return false;
    if (conjunto && c.conjunto !== conjunto) return false;
    if (criativoSearch && !c.anuncioKey.includes(criativoSearch)) return false;
    if (start || end) {
      const overlapsRange = (() => {
        if (!c.dataInicio && !c.dataFim) return true; // sem data, não filtra fora
        const s = c.dataInicio || c.dataFim;
        const e = c.dataFim || c.dataInicio;
        if (start && e < start) return false;
        if (end && s > end) return false;
        return true;
      })();
      if (!overlapsRange) return false;
    }
    return true;
  });

  renderAll();
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ---------------------------------------------------------------------------
// RENDERIZAÇÃO: CARDS
// ---------------------------------------------------------------------------
function renderCards(creatives) {
  const sum = (key) => creatives.reduce((acc, c) => acc + (c[key] || 0), 0);
  const valorGasto = sum("valorGasto");
  const impressoes = sum("impressoes");
  const cliques = sum("cliques");
  const leadsMeta = sum("leadsMeta");
  const leadsZoho = sum("leadsZoho");
  const reunioes = sum("reunioes");
  const assinaturas = sum("assinaturas");

  document.getElementById("cardInvestimento").textContent = fmtCurrency(valorGasto);
  document.getElementById("cardImpressoes").textContent = fmtInt(impressoes);
  document.getElementById("cardCliques").textContent = fmtInt(cliques);
  document.getElementById("cardLeadsMeta").textContent = fmtInt(leadsMeta);
  document.getElementById("cardLeadsZoho").textContent = fmtInt(leadsZoho);
  document.getElementById("cardReunioes").textContent = fmtInt(reunioes);
  document.getElementById("cardAssinaturas").textContent = fmtInt(assinaturas);
  document.getElementById("cardCplMeta").textContent = fmtCurrency(safeDiv(valorGasto, leadsMeta));
  document.getElementById("cardCustoReuniao").textContent = fmtCurrency(safeDiv(valorGasto, reunioes));
  document.getElementById("cardCustoAssinatura").textContent = fmtCurrency(safeDiv(valorGasto, assinaturas));

  return { valorGasto, impressoes, cliques, leadsMeta, leadsZoho, reunioes, assinaturas };
}

// ---------------------------------------------------------------------------
// RENDERIZAÇÃO: BADGES
// ---------------------------------------------------------------------------
function topBy(creatives, key, tieBreakers = []) {
  if (creatives.length === 0) return null;
  return [...creatives].sort((a, b) => {
    if (b[key] !== a[key]) return b[key] - a[key];
    for (const tb of tieBreakers) {
      const diff = tb.dir === "asc" ? a[tb.key] - b[tb.key] : b[tb.key] - a[tb.key];
      if (diff) return diff;
    }
    return 0;
  })[0];
}

function renderBadges(creatives) {
  const bestAssinaturas = topBy(creatives, "assinaturas", [{ key: "custoAssinatura", dir: "asc" }, { key: "taxaAssinatura", dir: "desc" }]);
  const bestReunioes = topBy(creatives, "reunioes", [{ key: "custoReuniao", dir: "asc" }, { key: "taxaReuniao", dir: "desc" }]);
  const bestLeads = topBy(creatives, "leadsZoho", [{ key: "cplMeta", dir: "asc" }, { key: "ctr", dir: "desc" }]);

  document.getElementById("badgeAssinaturas").textContent = bestAssinaturas && bestAssinaturas.assinaturas > 0
    ? `${bestAssinaturas.anuncio} (${bestAssinaturas.assinaturas})` : "Sem dados ainda";
  document.getElementById("badgeReunioes").textContent = bestReunioes && bestReunioes.reunioes > 0
    ? `${bestReunioes.anuncio} (${bestReunioes.reunioes})` : "Sem dados ainda";
  document.getElementById("badgeLeads").textContent = bestLeads && bestLeads.leadsZoho > 0
    ? `${bestLeads.anuncio} (${bestLeads.leadsZoho})` : "Sem dados ainda";

  return { bestAssinaturas, bestReunioes, bestLeads };
}

// ---------------------------------------------------------------------------
// RENDERIZAÇÃO: RANKINGS
// ---------------------------------------------------------------------------
function renderRankingList(elId, items, primaryKey, primaryFmt) {
  const el = document.getElementById(elId);
  if (items.length === 0) {
    el.innerHTML = `<li><span class="rank-name">Sem dados</span></li>`;
    return;
  }
  el.innerHTML = items
    .map((c, i) => `<li><span class="rank-name">${i + 1}. ${escapeHtml(c.anuncio)}</span><span class="rank-meta">${primaryFmt(c[primaryKey])} · ${escapeHtml(c.campanha)}</span></li>`)
    .join("");
}

function renderRankings(creatives) {
  const ranking1 = [...creatives]
    .sort((a, b) => b.assinaturas - a.assinaturas || a.custoAssinatura - b.custoAssinatura || b.taxaAssinatura - a.taxaAssinatura)
    .slice(0, CONFIG.rankingSize);
  const ranking2 = [...creatives]
    .sort((a, b) => b.reunioes - a.reunioes || a.custoReuniao - b.custoReuniao || b.taxaReuniao - a.taxaReuniao)
    .slice(0, CONFIG.rankingSize);
  const ranking3 = [...creatives]
    .sort((a, b) => b.leadsZoho - a.leadsZoho || a.cplMeta - b.cplMeta || b.ctr - a.ctr)
    .slice(0, CONFIG.rankingSize);

  renderRankingList("ranking1List", ranking1, "assinaturas", (v) => `${fmtInt(v)} assinaturas`);
  renderRankingList("ranking2List", ranking2, "reunioes", (v) => `${fmtInt(v)} reuniões`);
  renderRankingList("ranking3List", ranking3, "leadsZoho", (v) => `${fmtInt(v)} leads`);
}

// ---------------------------------------------------------------------------
// RENDERIZAÇÃO: INSIGHTS AUTOMÁTICOS
// ---------------------------------------------------------------------------
function renderInsights(creatives) {
  const list = document.getElementById("insightsList");
  if (creatives.length === 0) {
    list.innerHTML = "<li>Sem dados suficientes para gerar insights.</li>";
    return;
  }
  const insights = [];

  const maisLeads = topBy(creatives, "leadsZoho");
  const maisReunioes = topBy(creatives, "reunioes");
  const maisAssinaturas = topBy(creatives, "assinaturas");
  const validCustoReuniao = creatives.filter((c) => c.reunioes > 0);
  const validCustoAssinatura = creatives.filter((c) => c.assinaturas > 0);
  const menorCustoReuniao = validCustoReuniao.length ? [...validCustoReuniao].sort((a, b) => a.custoReuniao - b.custoReuniao)[0] : null;
  const menorCustoAssinatura = validCustoAssinatura.length ? [...validCustoAssinatura].sort((a, b) => a.custoAssinatura - b.custoAssinatura)[0] : null;
  const validTaxaReuniao = creatives.filter((c) => c.leadsZoho > 0);
  const melhorTaxaReuniao = validTaxaReuniao.length ? [...validTaxaReuniao].sort((a, b) => b.taxaReuniao - a.taxaReuniao)[0] : null;
  const melhorTaxaAssinatura = validTaxaReuniao.length ? [...validTaxaReuniao].sort((a, b) => b.taxaAssinatura - a.taxaAssinatura)[0] : null;

  if (maisLeads && maisLeads.leadsZoho > 0) insights.push(`O criativo que mais gerou leads foi <strong>${escapeHtml(maisLeads.anuncio)}</strong>, com ${fmtInt(maisLeads.leadsZoho)} leads.`);
  if (maisReunioes && maisReunioes.reunioes > 0) insights.push(`O criativo que mais gerou reuniões foi <strong>${escapeHtml(maisReunioes.anuncio)}</strong>, com ${fmtInt(maisReunioes.reunioes)} reuniões.`);
  if (maisAssinaturas && maisAssinaturas.assinaturas > 0) insights.push(`O criativo que mais gerou assinaturas foi <strong>${escapeHtml(maisAssinaturas.anuncio)}</strong>, com ${fmtInt(maisAssinaturas.assinaturas)} assinaturas.`);
  if (menorCustoReuniao) insights.push(`O menor custo por reunião foi de <strong>${fmtCurrency(menorCustoReuniao.custoReuniao)}</strong>, no criativo ${escapeHtml(menorCustoReuniao.anuncio)}.`);
  if (menorCustoAssinatura) insights.push(`O menor custo por assinatura foi de <strong>${fmtCurrency(menorCustoAssinatura.custoAssinatura)}</strong>, no criativo ${escapeHtml(menorCustoAssinatura.anuncio)}.`);
  if (melhorTaxaReuniao && melhorTaxaReuniao.taxaReuniao > 0) insights.push(`A melhor taxa de Lead → Reunião foi de <strong>${fmtPct(melhorTaxaReuniao.taxaReuniao)}</strong>, no criativo ${escapeHtml(melhorTaxaReuniao.anuncio)}.`);
  if (melhorTaxaAssinatura && melhorTaxaAssinatura.taxaAssinatura > 0) insights.push(`A melhor taxa de Lead → Assinatura foi de <strong>${fmtPct(melhorTaxaAssinatura.taxaAssinatura)}</strong>, no criativo ${escapeHtml(melhorTaxaAssinatura.anuncio)}.`);

  // Melhor / pior campanha (agregando criativos por campanha)
  const porCampanha = new Map();
  creatives.forEach((c) => {
    if (!porCampanha.has(c.campanha)) {
      porCampanha.set(c.campanha, { campanha: c.campanha, assinaturas: 0, reunioes: 0, leadsZoho: 0, valorGasto: 0 });
    }
    const g = porCampanha.get(c.campanha);
    g.assinaturas += c.assinaturas;
    g.reunioes += c.reunioes;
    g.leadsZoho += c.leadsZoho;
    g.valorGasto += c.valorGasto;
  });
  const campanhas = Array.from(porCampanha.values());
  if (campanhas.length > 1) {
    const ordenadas = [...campanhas].sort((a, b) => b.assinaturas - a.assinaturas || b.reunioes - a.reunioes || b.leadsZoho - a.leadsZoho);
    const melhor = ordenadas[0];
    const pior = ordenadas[ordenadas.length - 1];
    insights.push(`A melhor campanha foi <strong>${escapeHtml(melhor.campanha)}</strong> (${fmtInt(melhor.assinaturas)} assinaturas, ${fmtInt(melhor.reunioes)} reuniões, ${fmtInt(melhor.leadsZoho)} leads).`);
    insights.push(`A pior campanha foi <strong>${escapeHtml(pior.campanha)}</strong> (${fmtInt(pior.assinaturas)} assinaturas, ${fmtInt(pior.reunioes)} reuniões, ${fmtInt(pior.leadsZoho)} leads).`);
  }

  list.innerHTML = insights.length ? insights.map((i) => `<li>${i}</li>`).join("") : "<li>Sem dados suficientes para gerar insights.</li>";
  return insights;
}

// ---------------------------------------------------------------------------
// RENDERIZAÇÃO: TABELA
// ---------------------------------------------------------------------------
function renderTable() {
  const search = normalizeKey(document.getElementById("tableSearch").value);
  let rows = state.filtered.filter((c) => !search || c.anuncioKey.includes(search) || normalizeKey(c.campanha).includes(search));

  const { key, dir } = state.sort;
  rows = [...rows].sort((a, b) => {
    let va = a[key], vb = b[key];
    if (typeof va === "string") { va = va.toLowerCase(); vb = String(vb).toLowerCase(); }
    if (va < vb) return dir === "asc" ? -1 : 1;
    if (va > vb) return dir === "asc" ? 1 : -1;
    return 0;
  });

  const tbody = document.getElementById("mainTableBody");
  if (rows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="15" style="text-align:center;color:var(--text-muted);padding:24px;">Nenhum criativo encontrado com os filtros atuais.</td></tr>`;
    return;
  }

  tbody.innerHTML = rows
    .map(
      (c) => `<tr>
        <td>${escapeHtml(c.anuncio)}</td>
        <td>${escapeHtml(c.campanha)}</td>
        <td>${fmtCurrency(c.valorGasto)}</td>
        <td>${fmtInt(c.impressoes)}</td>
        <td>${fmtInt(c.cliques)}</td>
        <td>${fmtPct(c.ctr)}</td>
        <td>${fmtCurrency(c.cpc)}</td>
        <td>${fmtInt(c.leadsMeta)}</td>
        <td>${fmtInt(c.leadsZoho)}</td>
        <td>${fmtInt(c.reunioes)}</td>
        <td>${fmtInt(c.assinaturas)}</td>
        <td>${fmtCurrency(c.custoReuniao)}</td>
        <td>${fmtCurrency(c.custoAssinatura)}</td>
        <td>${fmtPct(c.taxaReuniao)}</td>
        <td>${fmtPct(c.taxaAssinatura)}</td>
      </tr>`
    )
    .join("");

  document.querySelectorAll("#mainTable thead th").forEach((th) => {
    th.classList.remove("sorted-asc", "sorted-desc");
    if (th.dataset.key === key) th.classList.add(dir === "asc" ? "sorted-asc" : "sorted-desc");
  });
}

// ---------------------------------------------------------------------------
// RENDERIZAÇÃO: GRÁFICOS (Chart.js)
// ---------------------------------------------------------------------------
const CHART_PALETTE = ["#6c5ce7", "#8b7bff", "#2ecc8f", "#f5b942", "#ff6b6b", "#4fb6e8", "#d98a4b", "#9aa4b8"];

function destroyChart(name) {
  if (state.charts[name]) {
    state.charts[name].destroy();
    delete state.charts[name];
  }
}

function makeBarChart(canvasId, name, labels, data, label, horizontal = true) {
  destroyChart(name);
  const ctx = document.getElementById(canvasId).getContext("2d");
  state.charts[name] = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{ label, data, backgroundColor: CHART_PALETTE, borderRadius: 6 }],
    },
    options: {
      indexAxis: horizontal ? "y" : "x",
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: "#9aa4b8" }, grid: { color: "#232a3b" } },
        y: { ticks: { color: "#9aa4b8" }, grid: { color: "#232a3b" } },
      },
    },
  });
}

function renderCharts(creatives) {
  const byLeads = [...creatives].filter((c) => c.leadsZoho > 0).sort((a, b) => b.leadsZoho - a.leadsZoho).slice(0, CONFIG.topN);
  const byReunioes = [...creatives].filter((c) => c.reunioes > 0).sort((a, b) => b.reunioes - a.reunioes).slice(0, CONFIG.topN);
  const byAssinaturas = [...creatives].filter((c) => c.assinaturas > 0).sort((a, b) => b.assinaturas - a.assinaturas).slice(0, CONFIG.topN);
  const byInvestimento = [...creatives].sort((a, b) => b.valorGasto - a.valorGasto).slice(0, CONFIG.topN);
  const byCpl = [...creatives].filter((c) => c.leadsMeta > 0).sort((a, b) => a.cplMeta - b.cplMeta).slice(0, CONFIG.topN);
  const byCustoReuniao = [...creatives].filter((c) => c.reunioes > 0).sort((a, b) => a.custoReuniao - b.custoReuniao).slice(0, CONFIG.topN);
  const byCustoAssinatura = [...creatives].filter((c) => c.assinaturas > 0).sort((a, b) => a.custoAssinatura - b.custoAssinatura).slice(0, CONFIG.topN);

  makeBarChart("chartLeads", "leads", byLeads.map((c) => c.anuncio), byLeads.map((c) => c.leadsZoho), "Leads Zoho");
  makeBarChart("chartReunioes", "reunioes", byReunioes.map((c) => c.anuncio), byReunioes.map((c) => c.reunioes), "Reuniões");
  makeBarChart("chartAssinaturas", "assinaturas", byAssinaturas.map((c) => c.anuncio), byAssinaturas.map((c) => c.assinaturas), "Assinaturas");
  makeBarChart("chartInvestimento", "investimento", byInvestimento.map((c) => c.anuncio), byInvestimento.map((c) => Number(c.valorGasto.toFixed(2))), "Investimento (R$)");
  makeBarChart("chartCpl", "cpl", byCpl.map((c) => c.anuncio), byCpl.map((c) => Number(c.cplMeta.toFixed(2))), "CPL (R$)");
  makeBarChart("chartCustoReuniao", "custoReuniao", byCustoReuniao.map((c) => c.anuncio), byCustoReuniao.map((c) => Number(c.custoReuniao.toFixed(2))), "Custo/Reunião (R$)");
  makeBarChart("chartCustoAssinatura", "custoAssinatura", byCustoAssinatura.map((c) => c.anuncio), byCustoAssinatura.map((c) => Number(c.custoAssinatura.toFixed(2))), "Custo/Assinatura (R$)");
}

// ---------------------------------------------------------------------------
// RENDER GERAL
// ---------------------------------------------------------------------------
function renderAll() {
  const creatives = state.filtered;
  renderCards(creatives);
  renderBadges(creatives);
  renderRankings(creatives);
  renderInsights(creatives);
  renderTable();
  renderCharts(creatives);

  const hasData = state.creatives.length > 0;
  ["filtersPanel", "cardsGrid", "badgesGrid", "insightsPanel", "rankingsGrid", "chartsGrid", "tablePanel"].forEach((id) => {
    document.getElementById(id).hidden = !hasData;
  });
  document.getElementById("exportCsvBtn").disabled = !hasData;
  document.getElementById("exportXlsxBtn").disabled = !hasData;
  document.getElementById("sendReportBtn").disabled = !hasData;
}

// ---------------------------------------------------------------------------
// TOAST / FEEDBACK VISUAL
// ---------------------------------------------------------------------------
let toastTimer = null;
function showToast(message, type = "info", duration = 3500) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.className = `toast show toast--${type}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("show"), duration);
}

// ---------------------------------------------------------------------------
// UPLOAD HANDLERS
// ---------------------------------------------------------------------------
function checkReadyToUpdate() {
  document.getElementById("updateDashboardBtn").disabled = !(state.metaFile && state.zohoFile);
}

document.getElementById("metaFile").addEventListener("change", (e) => {
  state.metaFile = e.target.files[0] || null;
  const status = document.getElementById("metaFileStatus");
  status.textContent = state.metaFile ? `Selecionado: ${state.metaFile.name}` : "Nenhum arquivo selecionado";
  status.className = state.metaFile ? "file-status ok" : "file-status";
  checkReadyToUpdate();
});

document.getElementById("zohoFile").addEventListener("change", (e) => {
  state.zohoFile = e.target.files[0] || null;
  const status = document.getElementById("zohoFileStatus");
  status.textContent = state.zohoFile ? `Selecionado: ${state.zohoFile.name}` : "Nenhum arquivo selecionado";
  status.className = state.zohoFile ? "file-status ok" : "file-status";
  checkReadyToUpdate();
});

document.getElementById("updateDashboardBtn").addEventListener("click", async () => {
  const processStatus = document.getElementById("processStatus");
  try {
    processStatus.textContent = "Lendo e processando planilhas...";
    processStatus.className = "process-status";

    const [metaMatrix, zohoMatrix] = await Promise.all([
      readWorkbook(state.metaFile),
      readWorkbook(state.zohoFile),
    ]);

    state.metaRows = parseMetaAds(metaMatrix);
    state.zohoRows = parseZohoCRM(zohoMatrix);

    if (state.metaRows.length === 0) throw new Error("Não foi possível identificar linhas válidas na planilha do Meta Ads.");
    if (state.zohoRows.length === 0) throw new Error("Não foi possível identificar linhas válidas na planilha do Zoho CRM.");

    const { creatives, unmatchedZoho } = buildCreatives(state.metaRows, state.zohoRows);
    state.creatives = creatives;
    state.filtered = creatives;

    populateFilterOptions(creatives);
    renderAll();

    const matchedDeals = state.zohoRows.length - unmatchedZoho.length;
    processStatus.textContent = `Dashboard atualizado: ${creatives.length} criativos, ${matchedDeals}/${state.zohoRows.length} negócios do Zoho cruzados.`;
    processStatus.className = "process-status success";
    showToast("Dashboard atualizado com sucesso.", "success");
  } catch (err) {
    console.error(err);
    processStatus.textContent = `Erro ao processar planilhas: ${err.message}`;
    processStatus.className = "process-status error";
    showToast("Erro ao processar planilhas.", "error");
  }
});

document.getElementById("clearFiltersBtn").addEventListener("click", () => {
  document.getElementById("filterStart").value = "";
  document.getElementById("filterEnd").value = "";
  document.getElementById("filterCampanha").value = "";
  document.getElementById("filterConjunto").value = "";
  document.getElementById("filterCriativo").value = "";
  applyFilters();
});

["filterStart", "filterEnd", "filterCampanha", "filterConjunto"].forEach((id) =>
  document.getElementById(id).addEventListener("change", applyFilters)
);
document.getElementById("filterCriativo").addEventListener("input", debounce(applyFilters, 250));
document.getElementById("tableSearch").addEventListener("input", debounce(renderTable, 200));

document.querySelectorAll("#mainTable thead th").forEach((th) => {
  th.addEventListener("click", () => {
    const key = th.dataset.key;
    if (state.sort.key === key) {
      state.sort.dir = state.sort.dir === "asc" ? "desc" : "asc";
    } else {
      state.sort.key = key;
      state.sort.dir = "desc";
    }
    renderTable();
  });
});

function debounce(fn, wait) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), wait);
  };
}

// ---------------------------------------------------------------------------
// EXPORTAÇÃO CSV / XLSX
// ---------------------------------------------------------------------------
function getExportRows() {
  return state.filtered.map((c) => ({
    "Nome do anúncio": c.anuncio,
    Campanha: c.campanha,
    "Conjunto de anúncios": c.conjunto,
    "Valor gasto": Number(c.valorGasto.toFixed(2)),
    Impressões: c.impressoes,
    Cliques: c.cliques,
    "CTR (%)": Number(c.ctr.toFixed(2)),
    "CPC (R$)": Number((c.cpc || 0).toFixed(2)),
    "Leads Meta": c.leadsMeta,
    "Leads Zoho": c.leadsZoho,
    "Reuniões Geradas": c.reunioes,
    "Assinaturas Realizadas": c.assinaturas,
    "Custo por Reunião (R$)": isFinite(c.custoReuniao) ? Number(c.custoReuniao.toFixed(2)) : "",
    "Custo por Assinatura (R$)": isFinite(c.custoAssinatura) ? Number(c.custoAssinatura.toFixed(2)) : "",
    "Taxa Lead→Reunião (%)": isFinite(c.taxaReuniao) ? Number(c.taxaReuniao.toFixed(2)) : "",
    "Taxa Lead→Assinatura (%)": isFinite(c.taxaAssinatura) ? Number(c.taxaAssinatura.toFixed(2)) : "",
  }));
}

document.getElementById("exportCsvBtn").addEventListener("click", () => {
  const rows = getExportRows();
  const ws = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(ws);
  downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8;" }), "report-meta-zoho.csv");
});

document.getElementById("exportXlsxBtn").addEventListener("click", () => {
  const rows = getExportRows();
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Report");
  XLSX.writeFile(wb, "report-meta-zoho.xlsx");
});

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------------
// ENVIAR REPORT PARA O SLACK
// ---------------------------------------------------------------------------
function isSameDay(d1, d2) {
  if (!d1 || !d2) return false;
  return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
}

function buildSlackPayload() {
  const today = new Date();
  const all = state.creatives;

  // "Hoje": negócios do Zoho criados hoje (para leads/reuniões/assinaturas),
  // e gasto do Meta cuja janela de datas inclui hoje.
  const dealsToday = state.zohoRows.filter((d) => isSameDay(d.horaCriacao, today));
  const leadsZohoHoje = dealsToday.length;
  const reunioesHoje = dealsToday.filter((d) => isMeetingStage(d.stageKey)).length;
  const assinaturasHoje = dealsToday.filter((d) => isSignupStage(d.stageKey)).length;
  const metaRowsHoje = state.metaRows.filter(
    (r) => (r.dataInicio && isSameDay(r.dataInicio, today)) || (r.dataFim && isSameDay(r.dataFim, today)) || (!r.dataInicio && !r.dataFim)
  );
  const valorGastoHoje = metaRowsHoje.reduce((acc, r) => acc + r.valorGasto, 0);
  const leadsMetaHoje = metaRowsHoje.reduce((acc, r) => acc + r.leads, 0);

  // Acumulado: totais de todo o dataset carregado.
  const acumulado = {
    valorGasto: all.reduce((a, c) => a + c.valorGasto, 0),
    leadsMeta: all.reduce((a, c) => a + c.leadsMeta, 0),
    leadsZoho: all.reduce((a, c) => a + c.leadsZoho, 0),
    reunioes: all.reduce((a, c) => a + c.reunioes, 0),
    assinaturas: all.reduce((a, c) => a + c.assinaturas, 0),
  };

  const bestAssinaturas = topBy(all, "assinaturas", [{ key: "custoAssinatura", dir: "asc" }]);
  const bestReunioes = topBy(all, "reunioes", [{ key: "custoReuniao", dir: "asc" }]);
  const bestLeads = topBy(all, "leadsZoho", [{ key: "cplMeta", dir: "asc" }]);

  const insightPeriodo = (bestAssinaturas && bestAssinaturas.assinaturas > 0)
    ? `O criativo "${bestAssinaturas.anuncio}" lidera em assinaturas (${bestAssinaturas.assinaturas}), com custo por assinatura de ${fmtCurrency(bestAssinaturas.custoAssinatura)}.`
    : "Ainda não há assinaturas registradas no período carregado.";

  return {
    data: today.toLocaleDateString("pt-BR"),
    resumoDia: {
      valorGasto: valorGastoHoje,
      leadsMeta: leadsMetaHoje,
      leadsZoho: leadsZohoHoje,
      reunioes: reunioesHoje,
      assinaturas: assinaturasHoje,
      cplMeta: safeDiv(valorGastoHoje, leadsMetaHoje),
      custoReuniao: safeDiv(valorGastoHoje, reunioesHoje),
      custoAssinatura: safeDiv(valorGastoHoje, assinaturasHoje),
    },
    acumulado,
    topCriativos: {
      assinaturas: bestAssinaturas ? `${bestAssinaturas.anuncio} (${bestAssinaturas.assinaturas})` : "—",
      reunioes: bestReunioes ? `${bestReunioes.anuncio} (${bestReunioes.reunioes})` : "—",
      leads: bestLeads ? `${bestLeads.anuncio} (${bestLeads.leadsZoho})` : "—",
    },
    insightPeriodo,
  };
}

// Este projeto é 100% estático (GitHub Pages), sem backend e sem webhook
// armazenado no front-end. O envio real para o Slack acontece através do
// conector de Slack já conectado na sessão do Cowork: ao clicar em
// "Enviar Report", o texto da mensagem (no formato oficial do template) é
// copiado para a área de transferência. Basta colar no Slack manualmente,
// ou pedir para o Claude (aqui no Cowork) enviar o texto copiado para você.
function buildSlackMessageText(payload) {
  const r = payload.resumoDia;
  const a = payload.acumulado;
  const t = payload.topCriativos;
  return [
    "📊 Report Meta Ads + Zoho | PipeLovers",
    `Data: ${payload.data}`,
    "",
    "Resumo do Dia",
    `• Valor gasto: ${fmtCurrency(r.valorGasto)} • Leads Meta: ${fmtInt(r.leadsMeta)} • Leads Zoho: ${fmtInt(r.leadsZoho)} • Reuniões geradas: ${fmtInt(r.reunioes)} • Assinaturas realizadas: ${fmtInt(r.assinaturas)} • CPL Meta: ${fmtCurrency(r.cplMeta)} • Custo por reunião: ${fmtCurrency(r.custoReuniao)} • Custo por assinatura: ${fmtCurrency(r.custoAssinatura)}`,
    "",
    "Acumulado",
    `• Investimento total: ${fmtCurrency(a.valorGasto)} • Leads Meta: ${fmtInt(a.leadsMeta)} • Leads Zoho: ${fmtInt(a.leadsZoho)} • Reuniões geradas: ${fmtInt(a.reunioes)} • Assinaturas realizadas: ${fmtInt(a.assinaturas)}`,
    "",
    "Top Criativos",
    `🥇 Mais assinaturas: ${t.assinaturas}`,
    `🥈 Mais reuniões: ${t.reunioes}`,
    `🥉 Mais leads: ${t.leads}`,
    "",
    "Insight do período",
    payload.insightPeriodo,
  ].join("\n");
}

document.getElementById("sendReportBtn").addEventListener("click", async () => {
  const btn = document.getElementById("sendReportBtn");
  btn.disabled = true;
  showToast("Gerando report...", "info", 60000);
  try {
    const payload = buildSlackPayload();
    const text = buildSlackMessageText(payload);
    let copied = false;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      try {
        await navigator.clipboard.writeText(text);
        copied = true;
      } catch (clipErr) {
        copied = false;
      }
    }
    if (!copied) {
      downloadBlob(new Blob([text], { type: "text/plain;charset=utf-8;" }), "report-slack.txt");
    }
    showToast(
      copied
        ? "Report enviado com sucesso: texto copiado. Cole no Slack ou peça para o Claude (Cowork) enviar."
        : "Report enviado com sucesso: arquivo report-slack.txt baixado. Envie esse texto no Slack ou peça para o Claude (Cowork) enviar.",
      "success",
      6000
    );
  } catch (err) {
    console.error(err);
    showToast(`Erro ao gerar report: ${err.message}`, "error", 5000);
  } finally {
    btn.disabled = !(state.creatives.length > 0);
  }
});
