/**
 * Google Apps Script — Gerador de Certificados via Google Slides
 * II Simpósio de Prevenção de IRAS e Stewardship
 *
 * CONFIGURAR (uma vez):
 * 1. Crie um Google Slides com 1 slide modelo (veja instruções abaixo)
 * 2. Substitua TEMPLATE_ID pelo ID da apresentação modelo
 * 3. Execute testarCertificado() para validar
 * 4. Execute gerarTodosCertificados() após o evento
 */

const TEMPLATE_ID = '1dK70dwScRNU2sYa_I-3hD5-C9Ta-l_E1Vknu_abPtK0'; // ID da apresentação Google Slides modelo
const SPREADSHEET_ID = '1phdlEjm__vtHIlLDl_V1AAe21jNjsxmsB61nVDj8Ufk';
const SHEET_NAME = 'Respostas ao formulario 1';
const FOLDER_NAME = 'Certificados_II_Simposio';
const CARGA_HORARIA = '16 horas';

/**
 * COMO CRIAR O SLIDE MODELO:
 * 1. Acesse slides.google.com > Crie uma apresentação em branco
 * 2. Arquivo > Configuração da página > Widescreen 16:9
 * 3. No slide, crie o design do certificado (cores, logos, layout)
 * 4. Insira caixas de texto com estes placeholders EXATOS:
 *    - {{NOME}}
 *    - {{DOCUMENTO}}
 *    - {{DIAS}}
 *    - {{DATA}}
 * 5. Salve. O ID está na URL: slides.google.com/presentation/d/ → ID ← /edit
 */

function gerarTodosCertificados() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Nenhum inscrito.'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
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

    const qualificacao = (row[12] || '').trim() || 'Ouvinte'; // coluna M
    const docInfo = [cpf ? 'CPF ' + cpf : '', conselho].filter(Boolean).join(' — ');
    const diasTexto = formatarDias(dias);
    const dataEmissao = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');

    try {
      // Gera codigo de verificacao unico
      const codigo = gerarCodigoVerificacao();
      const pdfBlob = gerarUmCertificado(nome, docInfo, diasTexto, dataEmissao, codigo, qualificacao);
      const fn = 'Certificado_' + nome.replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s+/g, '_') + '.pdf';
      folder.createFile(pdfBlob).setName(fn);

      // Registra codigo de verificacao na propria linha do participante (colunas K e L)
      const linha = i + 2;
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const mainSheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
      mainSheet.getRange(linha, 11).setValue(codigo); // coluna K
      mainSheet.getRange(linha, 12).setValue(dataEmissao); // coluna L
      gerados++;
    } catch (err) {
      Logger.log('Erro: ' + nome + ' - ' + err);
    }
  }
  Logger.log('Pronto! ' + gerados + ' certificados. ' + pulados + ' sem check-in.');
}

function gerarUmCertificado(nome, documento, dias, dataEmissao, codigo, qualificacao) {
  const tempName = 'temp_cert_' + Utilities.getUuid();
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const tempFile = templateFile.makeCopy(tempName);
  const tempPres = SlidesApp.openById(tempFile.getId());

  const slide = tempPres.getSlides()[0];

  substituirTexto(slide, '{{NOME}}', nome.toUpperCase());
  substituirTexto(slide, '{{DOCUMENTO}}', documento);
  substituirTexto(slide, '{{DIAS}}', dias);
  substituirTexto(slide, '{{DATA}}', dataEmissao);
  substituirTexto(slide, '{{QUALIFICACAO}}', qualificacao || 'Ouvinte');

  // DEBUG: verificar se a substituicao funcionou
  Logger.log('Qualificacao enviada: ' + qualificacao);
  slide.getShapes().forEach(function(s) {
    if (s.getText) {
      var t = s.getText().asString();
      if (t.indexOf('Ouvinte') > -1 || t.indexOf('Palestrante') > -1 || t.indexOf('Comissão') > -1) {
        Logger.log('Texto no shape: ' + t);
      }
    }
  });

  // Codigo de verificacao
  if (codigo) {
    const urlVerif = 'https://itaimscih.github.io/simposio2026/verificar?c=' + codigo;
    substituirTexto(slide, '{{CODIGO}}', codigo);

    // Insere QR Code como imagem (api.qrserver.com — gratuita e estavel)
    const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=' +
        encodeURIComponent(urlVerif);
    try {
      const qrBlob = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true }).getBlob();
      if (qrBlob.getBytes().length > 500) {
        const qrImg = slide.insertImage(qrBlob);
        qrImg.setLeft(35).setTop(466).setWidth(60).setHeight(60);
      }
    } catch(e) {
      Logger.log('QR Code nao inserido: ' + e);
    }

    // Link textual no lugar do placeholder {{QRURL}}
    substituirTexto(slide, '{{QRURL}}', urlVerif);
  }

  tempPres.saveAndClose();

  const pdfBlob = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + tempFile.getId() + '/export?mimeType=application/pdf',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
  ).getBlob();

  DriveApp.removeFile(tempFile);
  return pdfBlob;
}

function substituirTexto(slide, placeholder, valor) {
  slide.getShapes().forEach(function(shape) {
    if (shape.getText) {
      var text = shape.getText();
      if (text) text.replaceAllText(placeholder, valor);
    }
    // Verificar grupos (shapes aninhados)
    if (shape.getType && shape.getType() === SlidesApp.ShapeType.GROUP) {
      shape.getShapes().forEach(function(child) {
        if (child.getText) {
          var t = child.getText();
          if (t) t.replaceAllText(placeholder, valor);
        }
      });
    }
  });
}

// ── Teste ──────────────────────────────────────

function testarOuvinte() { _gerarCertificadoTeste('Ouvinte'); }
function testarPalestrante() { _gerarCertificadoTeste('Palestrante'); }
function testarComissao() { _gerarCertificadoTeste('Comissão Organizadora'); }

function _gerarCertificadoTeste(qualificacao) {
  const codigo = 'CERT-TESTE-' + Utilities.getUuid().substring(0, 4).toUpperCase();
  const dataEmissao = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
  const pdfBlob = gerarUmCertificado(
    'Maria Silva Santos',
    'CPF 123.456.789-00 — CRM 123456',
    'nos dias 14 e 15 de agosto de 2026',
    dataEmissao,
    codigo,
    qualificacao
  );
  const folder = getOrCreateFolder(FOLDER_NAME);
  folder.createFile(pdfBlob).setName('TESTE_' + qualificacao.replace(/\s/g, '_') + '_Maria_Silva_Santos.pdf');
  registrarVerificacao(codigo, 'Maria Silva Santos', dataEmissao);
  Logger.log('Certificado de teste gerado: ' + qualificacao);
}

// ── Verificacao ──────────────────────────────

function gerarCodigoVerificacao() {
  // Codigo totalmente aleatorio — nao deriva de dados pessoais
  return 'CERT-' + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function registrarVerificacao(codigo, nome, dataEmissao) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Verificacao');
  if (!sheet) {
    sheet = ss.insertSheet('Verificacao');
    sheet.appendRow(['Codigo', 'Nome', 'Data de Emissao']);
  }
  sheet.appendRow([codigo, nome, dataEmissao]);
}

// ── Helpers ──────────────────────────────────

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
