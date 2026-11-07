/* =========================================================
   PLANILHA DE LEADS — Larissa Johann (lariink.com)
   1) Cole este arquivo no Apps Script da planilha (Extensões > Apps Script).
   2) Rode a função "configurar" UMA vez (cria as abas).
   3) Implantar > Nova implantação > App da Web
      Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
   4) Copie a URL que termina em /exec e cole em sheetUrl no index.html.
   ========================================================= */

const ABAS = {
  leads:   "Leads",
  fichas:  "Fichas",
  visitas: "Visitas",
  eventos: "Eventos",
  painel:  "Painel",
  erros:   "Erros"
};

const UTM = ["utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term", "fbclid", "gclid"];

/* colunas fixas da aba Leads; campos novos do formulário viram colunas no fim, sozinhos */
const COL_LEADS = [
  "data_hora", "lead_id", "status", "nome", "whatsapp", "idade",
  "estilo", "quando", "ideia", "local", "tamanho", "pontos", "perfil", "funil",
  ...UTM,
  "primeira_origem", "primeira_campanha", "pagina", "referrer", "dispositivo", "visitor_id", "visitas_antes"
];
const COL_VISITAS = ["data_hora", "visitor_id", "sessao", "pagina", "referrer", "dispositivo", ...UTM, "primeira_origem", "primeira_campanha"];
const COL_EVENTOS = ["data_hora", "evento", "detalhe", "visitor_id", "sessao", "pagina", "utm_source", "utm_campaign"];
const STATUS = ["Novo", "Em contato", "Orçamento enviado", "Agendado", "Fechado", "Perdido"];

/* ---------- recebe os dados da página ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const tipo = String(d.tipo || "lead");
    d.data_hora = new Date();
    if (tipo === "visita") gravar_(ABAS.visitas, COL_VISITAS, d);
    else if (tipo === "lead") { gravarLead_(d); }
    else gravar_(ABAS.eventos, COL_EVENTOS, Object.assign({}, d, { evento: tipo }));
    return json_({ ok: true });
  } catch (err) {
    try { aba_(ABAS.erros, ["data_hora", "erro", "recebido"]).appendRow([new Date(), String(err), e && e.postData ? e.postData.contents : ""]); } catch (x) {}
    return json_({ ok: false, erro: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* abrir a URL /exec no navegador mostra isto: serve de teste */
function doGet() { return json_({ ok: true, msg: "Planilha da Larissa recebendo dados." }); }

/* ---------- lead: linha na aba Leads + bloco na aba Fichas ---------- */
function gravarLead_(d) {
  if (!d.lead_id) d.lead_id = "L" + Date.now();
  const sh = aba_(ABAS.leads, COL_LEADS);
  // evita duplicar se a mesma aplicação chegar duas vezes
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().flat() : [];
  if (ids.indexOf(d.lead_id) !== -1) return;
  d.status = "Novo";
  gravar_(ABAS.leads, COL_LEADS, d);
  const linha = sh.getLastRow();
  const num = String(d.whatsapp || "").replace(/\D/g, "");
  if (num) sh.getRange(linha, COL_LEADS.indexOf("whatsapp") + 1)
    .setFormula('=HYPERLINK("https://wa.me/55' + num.replace(/^55/, "") + '";"' + d.whatsapp + '")');
  const cor = { Quente: "#F4C7C3", Morno: "#FCE8B2", Frio: "#C9DAF8" }[d.perfil];
  if (cor) sh.getRange(linha, COL_LEADS.indexOf("perfil") + 1).setBackground(cor);
  ficha_(d);
}

function ficha_(d) {
  const sh = aba_(ABAS.fichas, null);
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 420);
  const blocos = [
    ["LEAD " + d.lead_id + "  ·  " + Utilities.formatDate(d.data_hora, "America/Sao_Paulo", "dd/MM/yyyy HH:mm"), ""],
    ["CONTATO", ""], ["Nome", d.nome], ["WhatsApp", d.whatsapp], ["Idade", d.idade],
    ["PROJETO", ""], ["Estilo", d.estilo], ["Quando", d.quando], ["Ideia", d.ideia], ["Local do corpo", d.local], ["Tamanho", d.tamanho],
    ["QUALIFICAÇÃO", ""], ["Perfil", d.perfil], ["Pontos", (d.pontos || 0) + "%"],
    ["ORIGEM", ""], ["Fonte (utm_source)", d.utm_source || "direto"], ["Meio (utm_medium)", d.utm_medium], ["Campanha", d.utm_campaign],
    ["Anúncio (utm_content)", d.utm_content], ["Termo (utm_term)", d.utm_term], ["Primeira origem", d.primeira_origem],
    ["Página", d.pagina], ["Veio de", d.referrer], ["Dispositivo", d.dispositivo], ["Visitas antes do lead", d.visitas_antes]
  ].map(r => [r[0], r[1] == null ? "" : r[1]]);
  // ficha mais nova em cima
  sh.insertRowsBefore(1, blocos.length + 1);
  const r = sh.getRange(1, 1, blocos.length, 2).setValues(blocos)
    .setVerticalAlignment("top").setWrap(true).setFontFamily("Arial");
  sh.getRange(1, 1, blocos.length, 2).setBorder(true, true, true, true, null, null, "#999999", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(1, 1, 1, 2).merge().setBackground("#3B2F2A").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  blocos.forEach((b, k) => {
    if (k && b[1] === "" && /^[A-ZÇÃ]+$/.test(b[0])) sh.getRange(k + 1, 1, 1, 2).merge().setBackground("#EFE6DD").setFontWeight("bold");
    else if (k) sh.getRange(k + 1, 1).setFontColor("#6B5E57");
  });
  const cor = { Quente: "#F4C7C3", Morno: "#FCE8B2", Frio: "#C9DAF8" }[d.perfil];
  if (cor) sh.getRange(13, 2).setBackground(cor).setFontWeight("bold");
  sh.getRange(blocos.length + 1, 1, 1, 2).clearFormat();
}

/* ---------- utilidades ---------- */
function gravar_(nome, colunas, d) {
  const sh = aba_(nome, colunas);
  let cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const extras = Object.keys(d).filter(k => k !== "tipo" && cab.indexOf(k) === -1 && typeof d[k] !== "object");
  if (extras.length) {
    sh.getRange(1, cab.length + 1, 1, extras.length).setValues([extras]).setFontWeight("bold").setBackground("#3B2F2A").setFontColor("#FFFFFF");
    cab = cab.concat(extras);
  }
  sh.appendRow(cab.map(k => {
    const v = d[k];
    if (v == null) return "";
    if (v instanceof Date) return v;
    const s = String(v).slice(0, 5000);
    return /^[=+\-@]/.test(s) ? "'" + s : s; // impede fórmula injetada pelo formulário
  }));
}

function aba_(nome, colunas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    if (colunas) {
      sh.getRange(1, 1, 1, colunas.length).setValues([colunas])
        .setFontWeight("bold").setBackground("#3B2F2A").setFontColor("#FFFFFF");
      sh.setFrozenRows(1);
      sh.getRange("A:A").setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
  }
  return sh;
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- rode UMA vez: cria e organiza as abas ---------- */
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone("America/Sao_Paulo");
  const leads = aba_(ABAS.leads, COL_LEADS);
  aba_(ABAS.fichas, null);
  aba_(ABAS.visitas, COL_VISITAS);
  aba_(ABAS.eventos, COL_EVENTOS);
  aba_(ABAS.erros, ["data_hora", "erro", "recebido"]);

  // status com lista suspensa
  const regra = SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build();
  leads.getRange(2, COL_LEADS.indexOf("status") + 1, 2000, 1).setDataValidation(regra);
  leads.setColumnWidth(COL_LEADS.indexOf("ideia") + 1, 300);

  // painel
  const p = aba_(ABAS.painel, null);
  p.clear();
  const L = "Leads!", V = "Visitas!";
  const col = c => String.fromCharCode(65 + COL_LEADS.indexOf(c));
  const vcol = c => String.fromCharCode(65 + COL_VISITAS.indexOf(c));
  p.getRange("A1").setValue("PAINEL — Larissa Johann").setFontSize(16).setFontWeight("bold");
  const num = [
    ["Visitas na página", "=MAX(0;COUNTA(" + V + "A:A)-1)"],
    ["Visitantes únicos", "=IFERROR(COUNTUNIQUE(" + V + vcol("visitor_id") + "2:" + vcol("visitor_id") + ");0)"],
    ["Popups abertos", "=COUNTIF(Eventos!B:B;\"popup_aberto\")"],
    ["Leads (aplicações)", "=MAX(0;COUNTA(" + L + "B:B)-1)"],
    ["Conversão visitante → lead", "=IFERROR(B6/B4;0)"],
    ["Leads Quentes", "=COUNTIF(" + L + col("perfil") + ":" + col("perfil") + ";\"Quente\")"],
    ["Leads Mornos", "=COUNTIF(" + L + col("perfil") + ":" + col("perfil") + ";\"Morno\")"],
    ["Leads Frios", "=COUNTIF(" + L + col("perfil") + ":" + col("perfil") + ";\"Frio\")"],
    ["Cliques no WhatsApp", "=COUNTIF(Eventos!B:B;\"whatsapp*\")"]
  ];
  p.getRange(3, 1, num.length, 2).setValues(num);
  p.getRange("B7").setNumberFormat("0.0%");
  p.getRange("A3:A11").setFontWeight("bold");

  p.getRange("D3").setValue("LEADS POR ORIGEM (utm_source)").setFontWeight("bold");
  p.getRange("D4").setFormula('=IFERROR(QUERY(' + L + 'A:AZ;"select ' + col("utm_source") + ', count(B) where B is not null group by ' + col("utm_source") + ' order by count(B) desc label ' + col("utm_source") + ' \'Origem\', count(B) \'Leads\'";1);"sem dados")');
  p.getRange("G3").setValue("LEADS POR CAMPANHA").setFontWeight("bold");
  p.getRange("G4").setFormula('=IFERROR(QUERY(' + L + 'A:AZ;"select ' + col("utm_campaign") + ', count(B) where B is not null group by ' + col("utm_campaign") + ' order by count(B) desc label ' + col("utm_campaign") + ' \'Campanha\', count(B) \'Leads\'";1);"sem dados")');
  p.getRange("J3").setValue("VISITAS POR ORIGEM").setFontWeight("bold");
  p.getRange("J4").setFormula('=IFERROR(QUERY(' + V + 'A:Z;"select ' + vcol("utm_source") + ', count(B) where B is not null group by ' + vcol("utm_source") + ' order by count(B) desc label ' + vcol("utm_source") + ' \'Origem\', count(B) \'Visitas\'";1);"sem dados")');
  p.getRange("M3").setValue("LEADS POR ANÚNCIO (utm_content)").setFontWeight("bold");
  p.getRange("M4").setFormula('=IFERROR(QUERY(' + L + 'A:AZ;"select ' + col("utm_content") + ', count(B) where B is not null group by ' + col("utm_content") + ' order by count(B) desc label ' + col("utm_content") + ' \'Anúncio\', count(B) \'Leads\'";1);"sem dados")');

  // ordem das abas
  [ABAS.painel, ABAS.leads, ABAS.fichas, ABAS.visitas, ABAS.eventos, ABAS.erros].forEach((n, k) => {
    ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(k + 1);
  });
  const vazia = ss.getSheetByName("Página1") || ss.getSheetByName("Sheet1");
  if (vazia && ss.getSheets().length > 1) ss.deleteSheet(vazia);
  ss.setActiveSheet(ss.getSheetByName(ABAS.painel));
}

/* rode para testar sem a página: deve aparecer 1 lead, 1 ficha e 1 visita */
function testar() {
  const base = { visitor_id: "teste", sessao: "s1", pagina: "/", utm_source: "instagram", utm_medium: "paid", utm_campaign: "teste", dispositivo: "celular" };
  doPost({ postData: { contents: JSON.stringify(Object.assign({ tipo: "visita" }, base)) } });
  doPost({ postData: { contents: JSON.stringify(Object.assign({ tipo: "lead", lead_id: "TESTE-" + Date.now(), nome: "Teste", whatsapp: "(45) 99999-9999", idade: "28", estilo: "Traços finos", quando: "O quanto antes, ainda este mês", ideia: "Flor", local: "Antebraço", tamanho: "8 cm", pontos: 100, perfil: "Quente" }, base)) } });
}
