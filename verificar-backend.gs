/**
 * Google Apps Script — Backend de Verificacao de Certificados
 * Adicione ao mesmo projeto do checkin ou crie um novo.
 *
 * Deploy: Implantar > Aplicativo da web > Qualquer pessoa
 * Depois atualize VERIF_URL no verificar.html com a URL gerada.
 */

const SPREADSHEET_ID = '1phdlEjm__vtHIlLDl_V1AAe21jNjsxmsB61nVDj8Ufk';
const SHEET_NAME = 'Verificacao';

function doGet(e) {
  const code = (e.parameter.c || '').trim().toUpperCase();
  const callback = e.parameter.callback || '';

  if (!code) {
    return json({ found: false, error: 'Codigo nao informado' }, callback);
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Codigo', 'Nome', 'Data de Emissao']);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return json({ found: false, error: 'Nenhum certificado registrado' }, callback);

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      const rowCode = (data[i][0] || '').toString().trim().toUpperCase();
      if (rowCode === code) {
        return json({
          found: true,
          nome: data[i][1] || '',
          dataEmissao: data[i][2] || '',
          evento: 'II Simposio de Prevencao de IRAS e Stewardship de Antimicrobianos — Regional Sul, Rede D\'Or'
        }, callback);
      }
    }
    return json({ found: false, error: 'Codigo nao encontrado' }, callback);

  } catch (err) {
    return json({ found: false, error: err.toString() }, callback);
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
