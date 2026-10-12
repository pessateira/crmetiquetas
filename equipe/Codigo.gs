// CRM Etiquetas: ponto de encontro da equipe (Google Planilhas + Apps Script).
// Cada pessoa da equipe manda para cá uma cópia do funil do WhatsApp dela,
// e todos leem daqui para ver o funil junto, com filtro por pessoa.
// Nada aqui mexe no WhatsApp.
//
// COMO INSTALAR: veja o arquivo COMO-INSTALAR.md (parte "Ver o CRM da equipe").

// 1) Troque pelo código secreto da sua equipe (o mesmo que vai no CRM de todos).
const CODIGO = 'troque-este-codigo';

const ABA_FUNIL = 'funil';
const ABA_NOTAS = 'anotacoes';
const COLS_FUNIL = ['pessoa', 'id da conversa', 'nome', 'telefone', 'etapa', 'última mensagem', 'atualizado em', 'dados'];
const COLS_NOTAS = ['pessoa', 'id da conversa', 'nome', 'valor estimado/mês', 'produto', 'nível', 'origem', 'próximo passo', 'data', 'anotações', 'alterado por', 'alterado em', 'ficha'];

function doPost(e) {
  let req = {};
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return resposta({ ok: false, erro: 'pedido inválido' });
  }
  if (CODIGO === 'troque-este-codigo') return resposta({ ok: false, erro: 'Troque o CODIGO no Apps Script antes de usar.' });
  if (req.codigo !== CODIGO) return resposta({ ok: false, erro: 'código da equipe errado' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (req.acao === 'teste') return resposta({ ok: true });
    if (req.acao === 'publicar') return resposta(publicar(req.pessoa, req.cartoes || []));
    if (req.acao === 'anotar') return resposta(anotar(req));
    if (req.acao === 'ler') return resposta(ler());
    return resposta({ ok: false, erro: 'ação desconhecida' });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return resposta({ ok: true, mensagem: 'CRM Etiquetas: ponto de encontro da equipe funcionando.' });
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function aba(nome, colunas) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let a = planilha.getSheetByName(nome);
  if (!a) {
    a = planilha.insertSheet(nome);
    a.getRange(1, 1, 1, colunas.length).setValues([colunas]);
  }
  return a;
}

function linhas(a) {
  const valores = a.getDataRange().getValues();
  return valores.slice(1).filter((l) => l[0] !== '');
}

function reescrever(a, colunas, corpo) {
  a.clearContents();
  a.getRange(1, 1, 1, colunas.length).setValues([colunas]);
  if (corpo.length) a.getRange(2, 1, corpo.length, colunas.length).setValues(corpo);
}

// Troca a foto do funil de uma pessoa pela nova (as anotações ficam em outra aba).
function publicar(pessoa, cartoes) {
  pessoa = String(pessoa || '').trim();
  if (!pessoa) return { ok: false, erro: 'falta o nome da pessoa' };
  const a = aba(ABA_FUNIL, COLS_FUNIL);
  const agora = new Date().toISOString();
  const outros = linhas(a).filter((l) => l[0] !== pessoa);
  const meus = cartoes.slice(0, 2000).map((c) => [
    pessoa,
    String(c.id),
    String(c.name || ''),
    String(c.phone || ''),
    String(c.stageName || 'Sem etapa'),
    c.lastAt ? new Date(c.lastAt * 1000).toISOString() : '',
    agora,
    JSON.stringify(c),
  ]);
  reescrever(a, COLS_FUNIL, outros.concat(meus));
  return { ok: true, total: meus.length };
}

// Ficha do cliente (selos, próximo passo, valor, anotações): é da equipe.
function anotar(req) {
  const a = aba(ABA_NOTAS, COLS_NOTAS);
  const corpo = linhas(a);
  const dono = String(req.dono || '');
  const id = String(req.id || '');
  if (!dono || !id) return { ok: false, erro: 'faltam dados' };
  const f = req.ficha || {};
  const linha = [
    dono,
    id,
    String(req.nome || ''),
    f.value == null || f.value === '' ? '' : Number(f.value),
    String(f.produto || ''),
    String(f.nivel || ''),
    String(f.origem || ''),
    String(f.nextStep || ''),
    String(f.nextDate || ''),
    String(f.notes || ''),
    String(req.pessoa || ''),
    new Date().toISOString(),
    JSON.stringify(f),
  ];
  const i = corpo.findIndex((l) => l[0] === dono && String(l[1]) === id);
  if (i >= 0) corpo[i] = linha;
  else corpo.push(linha);
  reescrever(a, COLS_NOTAS, corpo);
  return { ok: true };
}

function ler() {
  const cartoes = linhas(aba(ABA_FUNIL, COLS_FUNIL)).map((l) => {
    let dados = {};
    try {
      dados = JSON.parse(l[7]);
    } catch (err) {}
    return Object.assign(dados, { owner: l[0], id: String(l[1]), updatedAt: l[6] });
  });
  const notas = linhas(aba(ABA_NOTAS, COLS_NOTAS)).map((l) => {
    let ficha = {};
    try {
      ficha = JSON.parse(l[12]);
    } catch (err) {}
    return Object.assign(ficha, { owner: l[0], id: String(l[1]), by: l[10], at: l[11] });
  });
  return { ok: true, cartoes, notas };
}
