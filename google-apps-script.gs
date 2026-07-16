/**
 * Google Apps Script — Backend de Inscrições
 * II Simpósio de Prevenção de IRAS e Stewardship — Regional Sul Rede D'Or
 *
 * COMO CONFIGURAR:
 * 1. Acesse script.google.com > Novo projeto
 * 2. Cole este código e salve (Ctrl+S)
 * 3. Clique em "Implantar" > "Nova implantação" > "Aplicativo da web"
 * 4. Configure:
 *    - Executar como: "Eu"
 *    - Quem tem acesso: "Qualquer pessoa" (ou "Qualquer pessoa com Google")
 * 5. Clique em "Implantar" e copie a URL gerada
 * 6. Substitua SCRIPT_URL no index.html pela URL copiada
 *
 * ESTRUTURA DA PLANILHA:
 * Crie uma planilha no Google Sheets com as colunas:
 * A: Timestamp | B: Nome | C: E-mail | D: Profissao | E: Instituicao
 * F: CPF | G: Conselho de Classe | H: Status
 */

// Configuração
const MAX_VAGAS = 90;
const SPREADSHEET_ID = 'https://docs.google.com/spreadsheets/d/1dqXEYveKt5zd58gokpkFVrEmZOdU5nwFLLeAlIKVDl8/edit?gid=0#gid=0'; // ID da planilha Google Sheets
const SHEET_NAME = 'Inscricoes';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);

    // Verificar vagas disponíveis
    const lastRow = sheet.getLastRow();
    const inscritos = lastRow > 1 ? lastRow - 1 : 0;
    const vagasRestantes = MAX_VAGAS - inscritos;

    if (vagasRestantes <= 0) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, message: 'Vagas esgotadas.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Verificar se e-mail ou CPF ja inscrito
    const emails = sheet.getRange(2, 3, inscritos, 1).getValues().flat();
    if (emails.includes(data.email)) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, message: 'E-mail ja inscrito.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // Verificar CPF duplicado (coluna F)
    if (inscritos > 0) {
      const cpfs = sheet.getRange(2, 6, inscritos, 1).getValues().flat();
      if (cpfs.includes(data.cpf)) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'CPF ja inscrito.' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Registrar inscricao
    sheet.appendRow([
      new Date(),
      data.nome,
      data.email,
      data.profissao,
      data.instituicao,
      data.cpf,
      data.conselho,
      'Confirmado'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'Inscrição confirmada!',
        vagasRestantes: vagasRestantes - 1
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: 'Erro no servidor.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const inscritos = lastRow > 1 ? lastRow - 1 : 0;
  const vagasRestantes = MAX_VAGAS - inscritos;

  return ContentService
    .createTextOutput(JSON.stringify({ inscritos, vagasRestantes, maxVagas: MAX_VAGAS }))
    .setMimeType(ContentService.MimeType.JSON);
}
