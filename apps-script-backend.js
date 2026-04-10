// =====================================================
// CONTROLE RANCHO PE — Google Apps Script Backend
// Cole este código no Apps Script do seu Google Sheets
// =====================================================

const SHEET_PRODUTOS   = 'Table1';
const SHEET_USUARIOS   = 'Table2';
const SHEET_VENDAS     = 'Table3';
const SHEET_INVENTARIO = 'Table4';
const SHEET_CONTROLE   = 'Table5';

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function formatDate(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  const dt = new Date(d);
  return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`;
}

function isSameDate(d1, d2) {
  return formatDate(d1) === formatDate(d2);
}

// ─── GET ─────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === 'login') {
      result = handleLogin(e.parameter.usuario, e.parameter.senha);
    } else if (action === 'getAllData') {
      result = handleGetAllData();
    } else {
      result = { success: false, error: 'Ação desconhecida' };
    }
  } catch(err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── POST ─────────────────────────────────────────────
function doPost(e) {
  let body, result;
  try {
    body = JSON.parse(e.postData.contents);
    const action = body.action;
    const data = body.data;

    if (action === 'addVenda') {
      result = handleAddVenda(data);
    } else if (action === 'addInventario') {
      result = handleAddInventario(data);
    } else if (action === 'calcularControle') {
      result = handleCalcularControle(data);
    } else {
      result = { success: false, error: 'Ação desconhecida' };
    }
  } catch(err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── LOGIN ────────────────────────────────────────────
function handleLogin(usuario, senha) {
  const usuarios = sheetToObjects(SHEET_USUARIOS);
  const found = usuarios.find(u =>
    String(u.Usuario).toLowerCase() === String(usuario).toLowerCase() &&
    String(u.Senha) === String(senha)
  );
  return { success: !!found };
}

// ─── GET ALL DATA ─────────────────────────────────────
function handleGetAllData() {
  const produtos   = sheetToObjects(SHEET_PRODUTOS).map(p => ({
    ...p, Preco: parseFloat(p.Preco)||0, Estoque_Inicial: parseFloat(p.Estoque_Inicial)||0
  }));
  const vendas     = sheetToObjects(SHEET_VENDAS).map(v => ({
    ...v, Data: formatDate(v.Data), Quantidade: parseFloat(v.Quantidade)||0
  }));
  const inventario = sheetToObjects(SHEET_INVENTARIO).map(i => ({
    ...i, Data: formatDate(i.Data), Estoque_Real: parseFloat(i.Estoque_Real)||0
  }));
  const controle   = sheetToObjects(SHEET_CONTROLE).map(c => ({
    ...c, Data: formatDate(c.Data)
  }));

  return { success: true, data: { produtos, vendas, inventario, controle } };
}

// ─── ADD VENDA ────────────────────────────────────────
function handleAddVenda(data) {
  const sheet = getSheet(SHEET_VENDAS);
  sheet.appendRow([data.Data, data.Usuario, data.Produto, data.Quantidade]);
  return { success: true };
}

// ─── ADD INVENTÁRIO ───────────────────────────────────
function handleAddInventario(data) {
  const sheet = getSheet(SHEET_INVENTARIO);

  // Idempotente: atualiza se já existir registro do produto na mesma data
  const all = sheet.getDataRange().getValues();
  const headers = all[0];
  const dataIdx    = headers.indexOf('Data');
  const prodIdx    = headers.indexOf('Produto');
  const estoqIdx   = headers.indexOf('Estoque_Real');

  for (let i = 1; i < all.length; i++) {
    if (isSameDate(all[i][dataIdx], data.Data) && all[i][prodIdx] === data.Produto) {
      sheet.getRange(i+1, estoqIdx+1).setValue(data.Estoque_Real);
      return { success: true, updated: true };
    }
  }

  sheet.appendRow([data.Data, data.Produto, data.Estoque_Real]);
  return { success: true };
}

// ─── CALCULAR CONTROLE ────────────────────────────────
function handleCalcularControle(data) {
  const targetDate = data.Data; // yyyy-mm-dd

  const produtos   = sheetToObjects(SHEET_PRODUTOS);
  const vendas     = sheetToObjects(SHEET_VENDAS);
  const inventario = sheetToObjects(SHEET_INVENTARIO);
  const ctrlSheet  = getSheet(SHEET_CONTROLE);

  const resultados = [];

  for (const prod of produtos) {
    const nomeProd = prod.Produto;
    const estoqueInicial = parseFloat(prod.Estoque_Inicial) || 0;

    // 1. Estoque Ontem: último inventário com data < hoje
    const invAnteriores = inventario
      .filter(i => i.Produto === nomeProd && formatDate(i.Data) < targetDate)
      .sort((a,b) => formatDate(b.Data).localeCompare(formatDate(a.Data)));

    const estoqueOntem = invAnteriores.length > 0
      ? parseFloat(invAnteriores[0].Estoque_Real) || 0
      : estoqueInicial;

    // 2. Vendido hoje
    const vendidoHoje = vendas
      .filter(v => v.Produto === nomeProd && isSameDate(formatDate(v.Data), targetDate))
      .reduce((s, v) => s + (parseFloat(v.Quantidade)||0), 0);

    // 3. Estoque Real hoje
    const invHoje = inventario.find(i => i.Produto === nomeProd && isSameDate(formatDate(i.Data), targetDate));
    const estoqueReal = invHoje ? parseFloat(invHoje.Estoque_Real) || 0 : 0;

    // 4. Cálculo
    const reposicao = 0;
    const esperado = estoqueOntem - vendidoHoje + reposicao;
    const diferenca = estoqueReal - esperado;

    resultados.push({
      Data: targetDate,
      Produto: nomeProd,
      Estoque_Ontem: estoqueOntem,
      Vendido: vendidoHoje,
      Reposicao: reposicao,
      Esperado: esperado,
      Real: estoqueReal,
      Diferenca: diferenca
    });
  }

  // Idempotente: remove registros do dia e reinsere
  const allCtrl = ctrlSheet.getDataRange().getValues();
  const headers  = allCtrl[0];
  const dataIdx  = headers.indexOf('Data');

  // Encontra linhas existentes do dia (de baixo pra cima para deletar sem deslocar índices)
  const toDelete = [];
  for (let i = allCtrl.length - 1; i >= 1; i--) {
    if (isSameDate(allCtrl[i][dataIdx], targetDate)) {
      toDelete.push(i + 1); // 1-based
    }
  }
  toDelete.forEach(rowNum => ctrlSheet.deleteRow(rowNum));

  // Insere novos
  resultados.forEach(r => {
    ctrlSheet.appendRow([r.Data, r.Produto, r.Estoque_Ontem, r.Vendido, r.Reposicao, r.Esperado, r.Real, r.Diferenca]);
  });

  return { success: true, count: resultados.length };
}
