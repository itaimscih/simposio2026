/**
 * Google Apps Script — Gerador de Certificados
 * II Simpósio de Prevenção de IRAS e Stewardship
 *
 * COMO USAR:
 * 1. Cole este código no Apps Script (pode ser mesmo projeto do check-in)
 * 2. Execute a função gerarTodosCertificados()
 * 3. Autorize os escopos (Drive, Spreadsheet)
 * 4. Os PDFs aparecerão na pasta "Certificados_II_Simposio" no seu Drive
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
      const blob = HtmlService.createHtmlOutput(html)
          .setWidth(1122).setHeight(793).getBlob().setName('certificado.pdf');
      const fn = 'Certificado_' + nome.replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s+/g, '_') + '.pdf';
      folder.createFile(blob).setName(fn);
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
    '*{margin:0;padding:0;box-sizing:border-box;}\n' +
    'body{font-family:Verdana,sans-serif;width:1122px;height:793px;overflow:hidden;}\n' +
    '.page{width:100%;height:100%;background:linear-gradient(155deg,#001230 0%,#002855 40%,#001a3d 100%);position:relative;overflow:hidden;color:#fff;}\n' +
    '.page::before{content:"";position:absolute;top:-120px;right:-120px;width:500px;height:500px;border-radius:50%;background:rgba(113,197,232,0.08);}\n' +
    '.page::after{content:"";position:absolute;bottom:-80px;left:-80px;width:350px;height:350px;border-radius:50%;background:rgba(198,162,124,0.08);}\n' +
    '.corner{position:absolute;width:180px;height:180px;border-radius:50%;}\n' +
    '.corner-tl{top:-60px;left:-60px;background:rgba(0,61,165,0.5);}\n' +
    '.corner-br{bottom:-40px;right:-40px;background:rgba(113,197,232,0.2);}\n' +
    '.content{position:relative;z-index:2;text-align:center;padding:80px 100px;height:100%;display:flex;flex-direction:column;justify-content:center;}\n' +
    '.gold-line{width:80px;height:2px;background:#C6A27C;margin:0 auto 20px;border-radius:1px;}\n' +
    '.title{font-size:28px;font-weight:700;letter-spacing:2px;color:#F7CEA7;text-transform:uppercase;margin-bottom:30px;}\n' +
    '.subtitle{font-size:16px;color:#71C5E8;margin-bottom:40px;}\n' +
    '.name{font-size:32px;font-weight:700;color:#fff;margin-bottom:16px;letter-spacing:1px;}\n' +
    '.body-text{font-size:14px;color:rgba(255,255,255,0.75);line-height:1.8;max-width:700px;margin:0 auto 30px;}\n' +
    '.doc-line{font-size:12px;color:#F7CEA7;margin-bottom:50px;}\n' +
    '.signature-area{margin-top:10px;}\n' +
    '.sig-line{width:300px;height:1px;background:rgba(255,255,255,0.3);margin:0 auto 4px;}\n' +
    '.sig-name{font-size:13px;font-weight:600;color:#fff;}\n' +
    '.sig-role{font-size:10px;color:#71C5E8;}\n' +
    '.footer{position:absolute;bottom:30px;left:0;right:0;text-align:center;font-size:9px;color:rgba(255,255,255,0.35);}\n' +
    '.hours{margin-top:16px;font-size:12px;color:#C6A27C;font-weight:600;}\n' +
    '</style>\n</head>\n<body>\n<div class="page">\n' +
    '<div class="corner corner-tl"></div>\n<div class="corner corner-br"></div>\n' +
    '<div class="content">\n' +
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
    '<div class="signature-area">\n' +
    '<div class="sig-line"></div>\n' +
    '<div class="sig-name">Natanael S Adiwardana</div>\n' +
    '<div class="sig-role">Comissão Organizadora</div>\n' +
    '</div>\n</div>\n' +
    '<div class="footer">São Paulo, ' + dataEmissao + ' — Rede D\'Or São Luiz</div>\n' +
    '</div>\n</body>\n</html>';
}

// ── TESTE ───────────────────────────────────────

function testarCertificado() {
  const html = buildHTML(
    'Maria Silva Santos',
    '123.456.789-00',
    'CRM 123456',
    ['14/08', '15/08']
  );
  const blob = HtmlService.createHtmlOutput(html)
      .setWidth(1122).setHeight(793).getBlob().setName('certificado_teste.pdf');
  const folder = getOrCreateFolder(FOLDER_NAME);
  folder.createFile(blob).setName('TESTE_Certificado_Maria_Silva_Santos.pdf');
  Logger.log('Certificado de teste gerado na pasta ' + FOLDER_NAME);
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
