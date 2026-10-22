/**
 * SERVIÇO DE BOLSISTAS E CONVERSÃO DE TOMBAMENTOS
 */

// Nome exato da sua aba de bens
var NOME_ABA_INVENTARIO = "INVENTARIO_BENS"; 

/**
 * Busca apenas os tombamentos que possuem valor na Coluna G e estão atribuídos ao bolsista.
 * Função de LEITURA PURA (não distribui mais nada ao ser chamada).
 */
function getTombosBolsista(emailBolsista) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(NOME_ABA_INVENTARIO);
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    var tombosDoBolsista = [];
    var emailProcurado = String(emailBolsista || '').toLowerCase().trim();

    // Loop a partir da linha 2 (i = 1) ignorando o cabeçalho
    for (var i = 1; i < values.length; i++) {
      var linhaAtual = i + 1;
      var tomboOriginal = String(values[i][6] || '').trim();  // Coluna G (índice 6)
      var bolsistaResponsavel = String(values[i][18] || '').trim().toLowerCase(); // Coluna S (índice 18)

      // IGNORA se o tombo original estiver vazio ou for apenas traço '-'
      if (!tomboOriginal || tomboOriginal === '-') {
        continue;
      }

      // Se a linha estiver atribuída a este bolsista
      if (bolsistaResponsavel === emailProcurado) {
        var descricao = String(values[i][9] || '-').trim();      // Coluna J (índice 9)
        var tomboConvertido = String(values[i][19] || '').trim(); // Coluna T (índice 19)

        tombosDoBolsista.push({
          linha: linhaAtual,
          tomboOriginal: tomboOriginal,
          descricao: descricao,
          tomboConvertido: tomboConvertido,
          status: tomboConvertido ? "CONVERTIDO" : "PENDENTE"
        });
      }
    }

    return tombosDoBolsista;
  } catch (e) {
    throw new Error("Erro ao carregar tombos do bolsista: " + e.message);
  }
}

/**
 * Função para rodar UMA VEZ POR DIA via Acionador (Trigger).
 * Distribui apenas os bens que possuem tombo original (Coluna G) e que estão com a Coluna S em branco.
 * Preserva os itens que já foram atribuídos anteriormente.
 */
function distribuirTombosPendentes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUsu = ss.getSheetByName("USUARIOS");
  var sheetInv = ss.getSheetByName(NOME_ABA_INVENTARIO);
  if (!sheetUsu || !sheetInv) return;

  var dataUsu = sheetUsu.getDataRange().getValues();
  var bolsistas = [];

  // Obtém lista de e-mails de usuários cadastrados como BOLSISTA
  for (var u = 1; u < dataUsu.length; u++) {
    var perfil = String(dataUsu[u][2] || '').trim().toUpperCase(); // Coluna C (PERFIL)
    var email = String(dataUsu[u][0] || '').trim().toLowerCase(); // Coluna A (EMAIL)
    
    if (perfil === 'BOLSISTA' && email) {
      bolsistas.push(email);
    }
  }

  if (bolsistas.length === 0) return;

  var lastRow = sheetInv.getLastRow();
  if (lastRow < 2) return;

  // Lê da linha 2 em diante, Colunas A a T
  var dataInv = sheetInv.getRange(2, 1, lastRow - 1, 20).getValues(); 
  var alterou = false;
  var indexBolsista = 0;

  for (var i = 0; i < dataInv.length; i++) {
    var tomboOriginal = String(dataInv[i][6] || '').trim();  // Coluna G (índice 6)
    var bolsistaAtual = String(dataInv[i][18] || '').trim(); // Coluna S (índice 18)
    
    // SÓ atribui bolsista se a Coluna G estiver preenchida E a Coluna S estiver VAZIA
    if (tomboOriginal && tomboOriginal !== '-' && !bolsistaAtual) {
      dataInv[i][18] = bolsistas[indexBolsista % bolsistas.length];
      indexBolsista++;
      alterou = true;
    }
  }

  // Grava de volta na Coluna S (19) apenas se houve novos itens redistribuídos
  if (alterou) {
    var colunaSValores = dataInv.map(function(row) { return [row[18]]; });
    sheetInv.getRange(2, 19, colunaSValores.length, 1).setValues(colunaSValores);
    Logger.log("Distribuição concluída. " + indexBolsista + " item(ns) atribuído(s).");
  } else {
    Logger.log("Nenhum item pendente para redistribuir.");
  }
}

/**
 * Busca apenas os tombamentos que possuem valor na Coluna G e estão atribuídos ao bolsista
 */
function getTombosBolsista(emailBolsista) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Executa a redistribuição equilibrada apenas de linhas com tombo original preenchido
    distribuirTombosPendentes(ss);

    // 2. Busca os tombos atribuídos ao bolsista logado
    var sheet = ss.getSheetByName(NOME_ABA_INVENTARIO);
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    var tombosDoBolsista = [];
    var emailProcurado = String(emailBolsista || '').toLowerCase().trim();

    // Loop a partir da linha 2 (i = 1) ignorando o cabeçalho
    for (var i = 1; i < values.length; i++) {
      var linhaAtual = i + 1;
      var tomboOriginal = String(values[i][6] || '').trim();  // Coluna G (índice 6)
      var bolsistaResponsavel = String(values[i][18] || '').trim().toLowerCase(); // Coluna S (índice 18)

      // IGNORA se o tombo original estiver vazio ou for apenas traço '-'
      if (!tomboOriginal || tomboOriginal === '-') {
        continue;
      }

      // Se a linha estiver atribuída a este bolsista
      if (bolsistaResponsavel === emailProcurado) {
        var descricao = String(values[i][9] || '-').trim();      // Coluna J (índice 9)
        var tomboConvertido = String(values[i][19] || '').trim(); // Coluna T (índice 19)

        tombosDoBolsista.push({
          linha: linhaAtual,
          tomboOriginal: tomboOriginal,
          descricao: descricao,
          tomboConvertido: tomboConvertido,
          status: tomboConvertido ? "CONVERTIDO" : "PENDENTE"
        });
      }
    }

    return tombosDoBolsista;
  } catch (e) {
    throw new Error("Erro ao carregar tombos do bolsista: " + e.message);
  }
}

/**
 * Distribui apenas os bens que possuem tombo original (Coluna G) e sem bolsista atribuído
 */
function distribuirTombosPendentes(ss) {
  var sheetUsu = ss.getSheetByName("USUARIOS");
  var sheetInv = ss.getSheetByName(NOME_ABA_INVENTARIO);
  if (!sheetUsu || !sheetInv) return;

  var dataUsu = sheetUsu.getDataRange().getValues();
  var bolsistas = [];

  for (var u = 1; u < dataUsu.length; u++) {
    var perfil = String(dataUsu[u][2] || '').trim().toUpperCase(); // Coluna C (PERFIL)
    var email = String(dataUsu[u][0] || '').trim().toLowerCase(); // Coluna A (EMAIL)
    
    if (perfil === 'BOLSISTA' && email) {
      bolsistas.push(email);
    }
  }

  if (bolsistas.length === 0) return;

  var lastRow = sheetInv.getLastRow();
  if (lastRow < 2) return;

  var dataInv = sheetInv.getRange(2, 1, lastRow - 1, 20).getValues(); // Lê da Coluna A até T
  var alterou = false;
  var indexBolsista = 0;

  for (var i = 0; i < dataInv.length; i++) {
    var tomboOriginal = String(dataInv[i][6] || '').trim(); // Coluna G (índice 6)
    var bolsistaAtual = String(dataInv[i][18] || '').trim(); // Coluna S (índice 18)
    
    // Só atribui bolsista se a Coluna G estiver preenchida E a Coluna S estiver em branco
    if (tomboOriginal && tomboOriginal !== '-' && !bolsistaAtual) {
      dataInv[i][18] = bolsistas[indexBolsista % bolsistas.length];
      indexBolsista++;
      alterou = true;
    }
  }

  if (alterou) {
    var colunaSValores = dataInv.map(function(row) { return [row[18]]; });
    sheetInv.getRange(2, 19, colunaSValores.length, 1).setValues(colunaSValores);
  }
}

/**
 * Salva o Tombo Convertido na Coluna T
 */
function salvarTomboConvertido(linha, tomboConvertido) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(NOME_ABA_INVENTARIO);
    if (!sheet) throw new Error("Aba " + NOME_ABA_INVENTARIO + " não encontrada.");

    // Grava na Coluna T (Coluna 20)
    sheet.getRange(linha, 20).setValue(String(tomboConvertido).trim());

    return { sucesso: true, linha: linha, tomboConvertido: tomboConvertido };
  } catch (e) {
    throw new Error("Erro ao salvar tombo: " + e.message);
  }
}