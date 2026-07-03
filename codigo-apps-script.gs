/**
 * RIFA DA RAQUEL — Backend (Google Apps Script)
 * ------------------------------------------------
 * Este script deve ser criado DENTRO da planilha da rifa:
 * Extensões > Apps Script > cole este código > Implantar > App da Web
 *   - Executar como: Eu (você)
 *   - Quem pode acessar: Qualquer pessoa
 * Copie a URL gerada e cole na constante API_URL do index.html
 */

const ABA_NUMEROS = 'números';        // aba com todos os números da rifa
const ABA_COMPRADORES = 'COMPRADORES'; // aba onde as reservas são gravadas

/** Busca uma aba ignorando maiúsculas/acentos (ex.: "Números", "NUMEROS") */
function obterAba(nome) {
  const normalizar = s => s.toString().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  const alvo = normalizar(nome);
  const abas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (const aba of abas) {
    if (normalizar(aba.getName()) === alvo) return aba;
  }
  throw new Error('Aba não encontrada: ' + nome);
}

/** Lê todos os números cadastrados na aba "números" (qualquer célula no formato 001A / 015B) */
function listarTodosNumeros() {
  const valores = obterAba(ABA_NUMEROS).getDataRange().getValues();
  const numeros = [];
  const regex = /^\d{1,4}[AB]$/i;
  valores.forEach(linha => linha.forEach(celula => {
    const v = celula.toString().trim().toUpperCase();
    if (regex.test(v)) numeros.push(v);
  }));
  return numeros;
}

/** Lê os números já vendidos/reservados na aba COMPRADORES (coluna NÚMERO) */
function listarVendidos() {
  const aba = obterAba(ABA_COMPRADORES);
  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return [];
  const cabecalho = dados[0].map(c => c.toString().toUpperCase().trim());
  let col = cabecalho.findIndex(c => c.includes('NÚMERO') || c.includes('NUMERO'));
  if (col === -1) col = 4; // coluna E como padrão
  const vendidos = [];
  for (let i = 1; i < dados.length; i++) {
    const celula = dados[i][col];
    if (!celula) continue;
    celula.toString().split(/[;,\s]+/).forEach(n => {
      const v = n.trim().toUpperCase();
      if (v) vendidos.push(v);
    });
  }
  return vendidos;
}

/** GET → devolve números e vendidos para a página */
function doGet() {
  const resposta = {
    numeros: listarTodosNumeros(),
    vendidos: listarVendidos()
  };
  return ContentService
    .createTextOutput(JSON.stringify(resposta))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST → grava uma reserva na aba COMPRADORES */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita duas pessoas gravarem o mesmo número ao mesmo tempo

  try {
    const dados = JSON.parse(e.postData.contents);
    const nome = (dados.nome || '').toString().trim().toUpperCase();
    const telefone = (dados.telefone || '').toString().trim();
    const numeros = (dados.numeros || []).map(n => n.toString().trim().toUpperCase());

    if (!nome || !telefone || numeros.length === 0) {
      return responder({ ok: false, erro: 'Dados incompletos.' });
    }

    // Verifica se algum número já foi vendido
    const vendidos = new Set(listarVendidos());
    const conflitos = numeros.filter(n => vendidos.has(n));
    if (conflitos.length) {
      return responder({ ok: false, conflitos: conflitos });
    }

    // Calcula o valor total no servidor (não confia no valor vindo da página)
    const qtd = numeros.length;
    const precoUnit = qtd >= 3 ? 5.00 : 6.00;
    const total = qtd * precoUnit;

    // Monta a coluna PRÊMIO
    const premios = [];
    if (numeros.some(n => n.endsWith('A'))) premios.push('BODY SPLASH');
    if (numeros.some(n => n.endsWith('B'))) premios.push('FURADEIRA');

    // Grava: DATA | NOME COMPLETO | TELEFONE | PRÊMIO | NÚMERO | VALOR TOTAL
    const dataHoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
    obterAba(ABA_COMPRADORES).appendRow([
      dataHoje,
      nome,
      telefone,
      premios.join(';'),
      numeros.join(';'),
      total
    ]);

    return responder({ ok: true, total: total });

  } catch (erro) {
    return responder({ ok: false, erro: erro.message });
  } finally {
    lock.releaseLock();
  }
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
