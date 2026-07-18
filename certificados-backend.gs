/**
 * Google Apps Script — Gerador de Certificados
 * II Simpósio de Prevenção de IRAS e Stewardship
 *
 * USO:
 * 1. Cole este código no Apps Script
 * 2. Execute testarCertificado() para validar
 * 3. Execute gerarTodosCertificados() apos o evento
 * (Nao requer servicos adicionais — usa REST API do Drive)
 *
 * CRITÉRIO: participantes com check-in em pelo menos 1 dia (col I ou J)
 */

const SPREADSHEET_ID = '1phdlEjm__vtHIlLDl_V1AAe21jNjsxmsB61nVDj8Ufk';
const SHEET_NAME = 'Respostas ao formulario 1';
const FOLDER_NAME = 'Certificados_II_Simposio';
const CARGA_HORARIA = '16 horas';

function gerarTodosCertificados() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('Nenhum inscrito encontrado.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const folder = getOrCreateFolder(FOLDER_NAME);

  let gerados = 0, pulados = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const nome = (row[1] || '').trim();
    const cpf = formatCPF(row[5] || '');
    const conselho = (row[6] || '').trim();
    const d1 = row[8], d2 = row[9];

    if (!d1 && !d2) { pulados++; continue; }

    const dias = [];
    if (d1) dias.push('14/08');
    if (d2) dias.push('15/08');

    try {
      const html = buildHTML(nome, cpf, conselho, dias);
      const pdfBlob = htmlToPdf(html);
      const fn = 'Certificado_' + nome.replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s+/g, '_') + '.pdf';
      folder.createFile(pdfBlob).setName(fn);
      gerados++;
    } catch (err) {
      Logger.log('Erro: ' + nome + ' - ' + err);
    }
  }

  Logger.log('Pronto! ' + gerados + ' certificados gerados. ' + pulados + ' sem check-in.');
}

// ── Template HTML ────────────────────────────────

function buildHTML(nome, cpf, conselho, dias) {
  const nomeUpper = nome.toUpperCase();
  const docInfo = cpf ? 'CPF ' + cpf : '';
  const regInfo = conselho || '';
  const diasTexto = formatarDias(dias);
  const dataEmissao = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
  const docLine = [docInfo, regInfo].filter(Boolean).join(' — ');

  return '<!DOCTYPE html>\n<html lang="pt-BR">\n<head>\n<meta charset="UTF-8">\n<style>\n' +
    '@page{size:A4 landscape;margin:0;}\n' +
    '*{margin:0;padding:0;box-sizing:border-box;}\n' +
    'html,body{width:100%;height:100%;font-family:Verdana,sans-serif;}\n' +
    // Use a single table that fills the entire page
    '.cert-table{width:100%;height:100%;background-color:#002855;}\n' +
    '.cert-table td{vertical-align:middle;text-align:center;padding:30px 50px;}\n' +
    // Inner content card
    '.card{display:inline-block;background-color:#001a35;border:1px solid #C6A27C;padding:40px 60px;text-align:center;color:#fff;max-width:900px;}\n' +
    // Accent bars inside card
    '.bar-top{width:100%;height:3px;background-color:#003DA5;margin-bottom:30px;}\n' +
    '.bar-bottom{width:100%;height:3px;background-color:#C6A27C;margin-top:30px;}\n' +
    // Typography
    '.gold-line{width:60px;height:2px;background-color:#C6A27C;margin:0 auto 14px;}\n' +
    '.title{font-size:24px;font-weight:700;letter-spacing:3px;color:#F7CEA7;text-transform:uppercase;margin-bottom:16px;}\n' +
    '.subtitle{font-size:14px;color:#71C5E8;margin-bottom:26px;line-height:1.5;}\n' +
    '.name{font-size:28px;font-weight:700;color:#FFFFFF;margin-bottom:8px;}\n' +
    '.body-text{font-size:12.5px;color:#C8D6E5;line-height:1.9;margin:0 auto 20px;}\n' +
    '.doc-line{font-size:11px;color:#F7CEA7;margin-bottom:30px;}\n' +
    '.hours{font-size:11px;color:#C6A27C;font-weight:600;margin-bottom:30px;}\n' +
    '.sig-line{width:220px;height:1px;background-color:#5A7090;margin:0 auto 3px;}\n' +
    '.sig-name{font-size:12px;font-weight:600;color:#FFFFFF;}\n' +
    '.sig-role{font-size:9px;color:#71C5E8;}\n' +
    '.footer-text{font-size:8px;color:#5A7090;margin-top:22px;}\n' +
    '</style>\n</head>\n<body>\n' +
    '<table class="cert-table"><tr><td>\n' +
    '<div class="card">\n' +
    '<div class="bar-top"></div>\n' +
    '<div class="gold-line"></div>\n' +
    '<div class="title">Certificado de Participação</div>\n' +
    '<div class="subtitle">II Simpósio de Prevenção de IRAS e Stewardship de Antimicrobianos<br>Regional Sul — Rede D\'Or</div>\n' +
    '<div class="name">' + nomeUpper + '</div>\n' +
    (docLine ? '<div class="doc-line">' + docLine + '</div>\n' : '') +
    '<div class="body-text">' +
    'Certificamos que participou do II Simpósio de Prevenção de IRAS e Stewardship de Antimicrobianos — Regional Sul, ' +
    'realizado ' + diasTexto + ' no Hospital São Luiz Itaim, Auditório Térreo, São Paulo, SP, ' +
    'promovido pela Rede D\'Or — Regional Sul (Hospitais São Luiz Itaim, Vila Nova Star e Maternidade Star).' +
    '</div>\n' +
    '<div class="hours">Carga horária: ' + CARGA_HORARIA + '</div>\n' +
    '<div class="sig-line"></div>\n' +
    '<div class="sig-name">Natanael S Adiwardana</div>\n' +
    '<div class="sig-role">Comissão Organizadora</div>\n' +
    '<div class="footer-text">São Paulo, ' + dataEmissao + ' — Rede D\'Or São Luiz</div>\n' +
    '<div class="bar-bottom"></div>\n' +
    '</div>\n' +
    '</td></tr></table>\n</body>\n</html>';
}

// ── TESTE ───────────────────────────────────────

function testarCertificado() {
  const html = buildHTML(
    'Maria Silva Santos',
    '123.456.789-00',
    'CRM 123456',
    ['14/08', '15/08']
  );
  const pdfBlob = htmlToPdf(html);
  const folder = getOrCreateFolder(FOLDER_NAME);
  folder.createFile(pdfBlob).setName('TESTE_Certificado_Maria_Silva_Santos.pdf');
  Logger.log('Certificado de teste gerado na pasta ' + FOLDER_NAME);
}

// ── Conversão HTML → PDF ────────────────────────
// Usa REST API diretamente — nao requer servicos adicionais

function htmlToPdf(html) {
  const token = ScriptApp.getOAuthToken();
  const boundary = 'cert' + Math.random().toString(36).slice(2);

  // Upload HTML e converte para Google Doc
  const metadata = { name: 'temp', mimeType: 'application/vnd.google-apps.document' };
  const multipart = '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: text/html; charset=UTF-8\r\n\r\n' +
    html + '\r\n' +
    '--' + boundary + '--';

  const uploadResp = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&convert=true',
    { method: 'POST', headers: { Authorization: 'Bearer ' + token,
      'Content-Type': 'multipart/related; boundary=' + boundary },
      payload: multipart, muteHttpExceptions: true });

  if (uploadResp.getResponseCode() !== 200) {
    throw new Error('Upload falhou: ' + uploadResp.getContentText());
  }

  const fileId = JSON.parse(uploadResp.getContentText()).id;

  // Exporta como PDF
  const pdfBlob = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + fileId + '/export?mimeType=application/pdf',
    { headers: { Authorization: 'Bearer ' + token } }).getBlob();

  // Remove arquivo temporario
  UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + fileId,
    { method: 'DELETE', headers: { Authorization: 'Bearer ' + token } });

  return pdfBlob;
}

// ── Helpers ──────────────────────────────────────

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function formatCPF(cpf) {
  var nums = String(cpf).replace(/\D/g, '');
  if (nums.length !== 11) return cpf;
  return nums.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
}

function formatarDias(dias) {
  if (dias.length === 2) return 'nos dias 14 e 15 de agosto de 2026';
  if (dias[0] === '14/08') return 'no dia 14 de agosto de 2026';
  return 'no dia 15 de agosto de 2026';
}
