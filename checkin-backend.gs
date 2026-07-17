/**
 * Google Apps Script — Backend de Check-in (v2 com suporte JSONP)
 * II Simposio de Prevencao de IRAS e Stewardship
 */
const SPREADSHEET_ID = '1phdlEjm__vtHIlLDl_V1AAe21jNjsxmsB61nVDj8Ufk';
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

  if (action === 'auth') {
    return json({ valid: true }, callback);
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    if (!sheet) return json({ error: 'Planilha nao encontrada', results: [] }, callback);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return json({ results: [] }, callback);

    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

    if (action === 'search' && query.length >= 2) {
      const q = query.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      const results = [];
      for (let i = 0; i < data.length; i++) {
        const nome = (data[i][1] || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
        const email = (data[i][2] || '').toLowerCase();
        if (nome.includes(q) || email.includes(q)) {
          results.push({
            row: i + 2, nome: data[i][1] || '', profissao: data[i][3] || '',
            instituicao: data[i][4] || '',
            d1: !!data[i][8], d2: !!data[i][9]
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
          instituicao: row[4] || '',
          d1: !!row[8], d2: !!row[9]
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    if (data.action === 'checkin-d1') {
      const now = new Date();
      const ts = Utilities.formatDate(now, 'America/Sao_Paulo', 'dd/MM HH:mm');
      sheet.getRange(data.row, 9).setValue(ts);
      return json({ success: true, timestamp: ts, day: 1 });
    }
    if (data.action === 'checkin-d2') {
      const now = new Date();
      const ts = Utilities.formatDate(now, 'America/Sao_Paulo', 'dd/MM HH:mm');
      sheet.getRange(data.row, 10).setValue(ts);
      return json({ success: true, timestamp: ts, day: 2 });
    }
    if (data.action === 'undo-d1') {
      sheet.getRange(data.row, 9).setValue('');
      return json({ success: true, message: 'Check-in Dia 1 removido' });
    }
    if (data.action === 'undo-d2') {
      sheet.getRange(data.row, 10).setValue('');
      return json({ success: true, message: 'Check-in Dia 2 removido' });
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
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
