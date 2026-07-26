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

  // Verificacao de certificado (publico, sem PIN)
  if (action === 'verify') {
    const code = (e.parameter.c || '').trim().toUpperCase();
    if (!code) return json({ found: false, error: 'Codigo nao informado' }, callback);
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var vSheet = ss.getSheetByName('Verificacao');
      if (!vSheet) return json({ found: false, error: 'Nenhum certificado registrado' }, callback);
      const vLastRow = vSheet.getLastRow();
      if (vLastRow < 2) return json({ found: false, error: 'Nenhum certificado registrado' }, callback);
      const vData = vSheet.getRange(2, 1, vLastRow - 1, 3).getValues();
      for (let i = 0; i < vData.length; i++) {
        if ((vData[i][0] || '').toString().trim().toUpperCase() === code) {
          return json({
            found: true,
            nome: vData[i][1] || '',
            dataEmissao: vData[i][2] || '',
            evento: 'II Simposio de Prevencao de IRAS e Stewardship de Antimicrobianos — Regional Sul, Rede D\'Or'
          }, callback);
        }
      }
      return json({ found: false, error: 'Codigo nao encontrado' }, callback);
    } catch (err) {
      return json({ found: false, error: err.toString() }, callback);
    }
  }

  if (pin !== CHECKIN_PIN) {
    return json({ error: 'PIN invalido', results: [] }, callback);
  }

  if (action === 'auth') {
    return json({ valid: true }, callback);
  }

  // Acoes de check-in via GET (JSONP, bypass CORS)
  var ckRow = parseInt(e.parameter.row) || 0;
  if (ckRow > 0 && (action === 'checkin-d1' || action === 'checkin-d2' || action === 'undo-d1' || action === 'undo-d2')) {
    try {
      const ss2 = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sh2 = ss2.getSheetByName(SHEET_NAME) || ss2.getSheets()[0];
      if (action === 'checkin-d1') {
        const ts = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM HH:mm');
        sh2.getRange(ckRow, 9).setValue(ts);
        return json({ success: true, timestamp: ts, day: 1 }, callback);
      }
      if (action === 'checkin-d2') {
        const ts = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM HH:mm');
        sh2.getRange(ckRow, 10).setValue(ts);
        return json({ success: true, timestamp: ts, day: 2 }, callback);
      }
      if (action === 'undo-d1') {
        sh2.getRange(ckRow, 9).setValue('');
        return json({ success: true, message: 'Check-in Dia 1 removido' }, callback);
      }
      if (action === 'undo-d2') {
        sh2.getRange(ckRow, 10).setValue('');
        return json({ success: true, message: 'Check-in Dia 2 removido' }, callback);
      }
    } catch (err) {
      return json({ success: false, message: err.toString() }, callback);
    }
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
