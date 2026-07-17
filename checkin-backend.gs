/**
 * Google Apps Script — Backend de Check-in (v2 com suporte JSONP)
 * II Simposio de Prevencao de IRAS e Stewardship
 */
const SPREADSHEET_ID = '1dqXEYveKt5zd58gokpkFVrEmZOdU5nwFLLeAlIKVDl8';
const SHEET_NAME = 'Respostas ao formulario 1';
const CHECKIN_PIN = '202420252026';

function doGet(e) {
  const action = e.parameter.action;
  const pin = e.parameter.pin;
  const query = e.parameter.q || '';
  const callback = e.parameter.callback || '';

  if (pin !== CHECKIN_PIN) {
    return json({ error: 'PIN invalido', results: [] }, callback);
  }

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) return json({ error: 'Planilha nao encontrada', results: [] }, callback);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return json({ results: [] }, callback);

    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    if (action === 'search' && query.length >= 2) {
      const q = query.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      const results = [];
      for (let i = 0; i < data.length; i++) {
        const nome = (data[i][1] || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
        const email = (data[i][2] || '').toLowerCase();
        if (nome.includes(q) || email.includes(q)) {
          results.push({
            row: i + 2, nome: data[i][1] || '', profissao: data[i][3] || '',
            instituicao: data[i][4] || '', checkedIn: !!data[i][8]
          });
          if (results.length >= 15) break;
        }
      }
      return json({ results: results }, callback);
    }

    if (action === 'all') {
      const results = data.map(function(row, i) {
        return {
          row: i + 2, nome: row[1] || '', profissao: row[3] || '',
          instituicao: row[4] || '', checkedIn: !!row[8]
        };
      });
      return json({ results: results }, callback);
    }

    return json({ results: [] }, callback);

  } catch (err) {
    return json({ error: err.toString(), results: [] }, callback);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.pin !== CHECKIN_PIN) {
      return json({ success: false, message: 'PIN invalido' });
    }
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);

    if (data.action === 'checkin') {
      const now = new Date();
      const ts = Utilities.formatDate(now, 'America/Sao_Paulo', 'dd/MM HH:mm');
      sheet.getRange(data.row, 9).setValue(ts);
      return json({ success: true, timestamp: ts });
    }
    if (data.action === 'undo') {
      sheet.getRange(data.row, 9).setValue('');
      return json({ success: true, message: 'Check-in removido' });
    }
    return json({ success: false, message: 'Acao desconhecida' });
  } catch (err) {
    return json({ success: false, message: err.toString() });
  }
}

function json(obj, callback) {
  var out = JSON.stringify(obj);
  if (callback) {
    out = callback + '(' + out + ')';
  }
  return ContentService.createTextOutput(out)
    .setMimeType(callback ? 'application/javascript' : ContentService.MimeType.JSON);
}
