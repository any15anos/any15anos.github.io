/**
 * Confirmação de presença — 15 anos da Any Karoliny
 *
 * Grava na planilha "Confirmações 15 anos Any Karoliny" (PLANILHA_ID abaixo).
 * Para reaproveitar: troque PLANILHA_ID pelo ID de outra planilha.
 * Implantação: Implantar > Nova implantação > tipo "App da Web".
 *      - Executar como: Eu
 *      - Quem pode acessar: Qualquer pessoa
 * A URL que termina em /exec vai em scriptUrl no index.html.
 *
 * Cada confirmação vira uma linha na aba "Confirmações".
 */

const PLANILHA_ID = '1uMZ8lLoZtVrH6eTJVxUPQykj4sknPSqa3KQULW551ms';
const NOME_ABA = 'Confirmações';
const FUSO = 'America/Sao_Paulo'; // horário de Brasília
const CABECALHO = ['Data/hora', 'Nome', 'WhatsApp', 'Presença', 'Adultos', 'Crianças', 'Total', 'Recado'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const aba = pegarAba_();

    const adultos = paraNumero_(p.adultos);
    const criancas = paraNumero_(p.criancas);
    const vai = p.presenca === 'sim';

    aba.appendRow([
      Utilities.formatDate(new Date(), FUSO, 'dd/MM/yyyy HH:mm'),
      limpar_(p.nome),
      "'" + String(p.telefone || '').replace(/[^\d() -]/g, '').slice(0, 20), // apóstrofo mantém como texto
      vai ? 'Vai ✅' : 'Não vai ❌',
      vai ? adultos : 0,
      vai ? criancas : 0,
      vai ? adultos + criancas : 0,
      limpar_(p.recado)
    ]);

    return resposta_({ ok: true });
  } catch (err) {
    return resposta_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Abrir a URL /exec no navegador mostra este resumo — útil para testar.
function doGet() {
  const aba = pegarAba_();
  const linhas = Math.max(aba.getLastRow() - 1, 0);
  let total = 0;
  if (linhas > 0) {
    aba.getRange(2, 7, linhas, 1).getValues().forEach(function (l) { total += Number(l[0]) || 0; });
  }
  return resposta_({ ok: true, respostas: linhas, pessoasConfirmadas: total });
}

function pegarAba_() {
  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  let aba = planilha.getSheetByName(NOME_ABA);
  if (!aba) {
    aba = planilha.insertSheet(NOME_ABA);
    aba.appendRow(CABECALHO);
    aba.getRange(1, 1, 1, CABECALHO.length)
      .setFontWeight('bold').setBackground('#9b7fd1').setFontColor('#ffffff');
    aba.setFrozenRows(1);
    aba.setColumnWidth(2, 220);
    aba.setColumnWidth(8, 320);
  }
  return aba;
}

function paraNumero_(v) {
  const n = parseInt(v, 10);
  return isNaN(n) || n < 0 ? 0 : Math.min(n, 20);
}

function limpar_(v) {
  // evita que um texto começando com = + - @ vire fórmula na planilha
  return String(v || '').trim().slice(0, 500).replace(/^[=+\-@]/, "'$&");
}

function resposta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
