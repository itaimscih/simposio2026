/**
 * Google Apps Script — Backend de Check-in
 * II Simposio de Prevencao de IRAS e Stewardship
 *
 * COMO CONFIGURAR:
 * 1. Acesse script.google.com > Novo projeto
 * 2. Cole este codigo e salve
 * 3. Substitua SPREADSHEET_ID pelo ID da planilha de respostas
 * 4. Substitua CHECKIN_PIN pela senha desejada
 * 5. Implantar > Nova implantacao > Aplicativo da web
 *    - Executar como: Eu
 *    - Quem tem acesso: Qualquer pessoa
 * 6. Copie a URL gerada e cole no checkin.html (SCRIPT_URL)
 *
 * PLANILHA (respostas do Google Forms):
 * A: Timestamp | B: Nome | C: E-mail | D: Profissao
 * E: Instituicao | F: CPF | G: Conselho | H: Status
 *
 * COLUNA I: usada para check-in automaticamente pelo script.
 * Nao e necessario adicionar cabecalho manualmente — o Forms
 * nao bloqueia escrita nas linhas de dados, apenas no cabecalho.
 */

const SPREADSHEET_ID = 'SUBSTITUA_PELO_ID_DA_SUA_PLANILHA';
const SHEET_NAME = 'Respostas ao formulario 1'; // nome padrao do Google Forms
const CHECKIN_PIN = '2026'; // altere para a senha desejada

function doGet(e) {
  const action = e.parameter.action;
  const pin = e.parameter.pin;
  const query = e.parameter.q || '';

  if (action === 'auth') {
    return json({ valid: pin === CHECKIN_PIN });
  }

  // Verificar PIN em todas as operacoes
  if (pin !== CHECKIN_PIN) {
    return json({ error: 'Acesso negado', results: [] });
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return json({ results: [] });

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  if (action === 'search' && query.length >= 2) {
    const q = query.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
    const results = [];

    for (let i = 0; i < data.length; i++) {
      // Coluna B = nome (indice 1), Coluna C = email (indice 2)
      const nome = (data[i][1] || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      const email = (data[i][2] || '').toLowerCase();
      if (nome.includes(q) || email.includes(q)) {
        results.push({
          row: i + 2,
          nome: data[i][1] || '',
          profissao: data[i][3] || '',
          instituicao: data[i][4] || '',
          checkedIn: !!data[i][8]  // Coluna I = check-in
        });
        if (results.length >= 15) break; // maximo 15 resultados
      }
    }
    return json({ results: results });

  } else if (action === 'all') {
    const results = data.map(function(row, i) {
      return {
        row: i + 2,
        nome: row[1] || '',
        profissao: row[3] || '',
        instituicao: row[4] || '',
        checkedIn: !!row[8]
      };
    });
    return json({ results: results });
  }

  return json({ results: [] });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const pin = data.pin;

  if (pin !== CHECKIN_PIN) {
    return json({ success: false, message: 'PIN invalido' });
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);

  if (data.action === 'checkin') {
    // Marcar check-in (coluna I)
    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'America/Sao_Paulo', 'dd/MM HH:mm');
    sheet.getRange(data.row, 9).setValue(timestamp);
    return json({ success: true, timestamp: timestamp });
  }

  if (data.action === 'undo') {
    sheet.getRange(data.row, 9).setValue('');
    return json({ success: true, message: 'Check-in removido' });
  }

  return json({ success: false, message: 'Acao desconhecida' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
