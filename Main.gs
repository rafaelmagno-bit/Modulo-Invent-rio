var SPREADSHEET_ID = "1Yh5udFrSAcn0GSlaE54IQVsUAyHjfz-CE0p_YIhD0dI";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Inventário de Bens Móveis 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPlanilha() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function normalizarCodigo(val) {
  if (val === null || val === undefined) return "";
  var str = String(val).trim();
  if (str.endsWith(".0")) {
    str = str.substring(0, str.length - 2);
  }
  return str.replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
}

function formatarDataGS(valor) {
  if (!valor) return '-';
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  }
  return String(valor);
}

function getUserContext() {
  return getCustomUserContext();
}

function getCustomUserContext() {
  try {
    var emailUser = Session.getActiveUser().getEmail().toLowerCase().trim();
    var ss = getPlanilha();
    var sheetUsuarios = ss.getSheetByName("USUARIOS");
    
    // Objeto padrão para usuário NÃO CADASTRADO
    var userContext = {
      email: emailUser || "",
      nome: "Usuário não cadastrado",
      perfil: "NÃO CADASTRADO",
      codigoUg: "",
      siglaUg: "",
      encontrado: false
    };

    if (!sheetUsuarios) return userContext;

    var data = sheetUsuarios.getDataRange().getValues();
    
    // Busca o usuário pelo e-mail na aba USUARIOS
    for (var i = 1; i < data.length; i++) {
      var emailRow = String(data[i][0]).toLowerCase().trim();
      if (emailRow === emailUser && emailUser !== "") {
        userContext.nome = String(data[i][1] || "Usuário").trim();
        userContext.perfil = String(data[i][2] || "").trim();
        userContext.codigoUg = String(data[i][3] || "").trim();
        userContext.siglaUg = data[i][4] ? String(data[i][4]).trim() : "";
        userContext.encontrado = true;
        break;
      }
    }

    // Busca a sigla da UG apenas se o usuário for encontrado e possuir código da UG
    if (userContext.encontrado && !userContext.siglaUg && userContext.codigoUg) {
      userContext.siglaUg = buscarSiglaUg(ss, userContext.codigoUg);
    }

    return userContext;
  } catch (e) {
    Logger.log("Erro em getUserContext: " + e.toString());
    throw new Error("Erro ao identificar o usuário: " + e.message);
  }
}

function buscarSiglaUg(ss, codigoUg) {
  try {
    var sheetUg = ss.getSheetByName("UG") || ss.getSheetByName("UGS");
    if (sheetUg) {
      var dataUgs = sheetUg.getDataRange().getValues();
      for (var j = 1; j < dataUgs.length; j++) {
        if (String(dataUgs[j][0]).trim() === codigoUg) {
          return String(dataUgs[j][2] || dataUgs[j][1]).trim();
        }
      }
    }
  } catch (e) {
    Logger.log("Erro ao buscar sigla da UG: " + e.toString());
  }
  return codigoUg; 
}