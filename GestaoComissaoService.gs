/**
 * Busca UAs da aba ORGANOGRAMA e Membros da aba USUARIOS filtrados pelo código da UG
 */
function getDadosIniciaisGestorComissao(codigoUg) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var codigoUgStr = String(codigoUg || '').trim();

    // 1. Busca Membros da Comissão cadastrados na aba USUARIOS (Col D = CODIGO UG)
    var sheetUsers = ss.getSheetByName("USUARIOS");
    var membros = [];
    if (sheetUsers) {
      var dataUsers = sheetUsers.getDataRange().getValues();
      for (var i = 1; i < dataUsers.length; i++) {
        var ugUser = String(dataUsers[i][3] || '').trim(); // Coluna D: CODIGO UG
        if (ugUser === codigoUgStr) {
          membros.push({
            email: String(dataUsers[i][0] || '').trim(),
            nome: String(dataUsers[i][1] || '').trim(),
            perfil: String(dataUsers[i][2] || '').trim(),
            siape: String(dataUsers[i][5] || '').trim()
          });
        }
      }
    }

    // 2. Busca UAs na aba ORGANOGRAMA (Col D = UG)
    var sheetOrg = ss.getSheetByName("ORGANOGRAMA");
    var uas = [];
    if (sheetOrg) {
      var dataOrg = sheetOrg.getDataRange().getValues();
      for (var j = 1; j < dataOrg.length; j++) {
        var ugOrg = String(dataOrg[j][3] || '').trim(); // Coluna D: UG
        if (ugOrg === codigoUgStr) {
          uas.push({
            codigoSipac: String(dataOrg[j][0] || '').trim(), // Coluna A
            descricao: String(dataOrg[j][1] || '').trim(),   // Coluna B
            status: String(dataOrg[j][5] || '').trim(),      // Coluna F
            responsaveis: String(dataOrg[j][6] || '').trim() // Coluna G
          });
        }
      }
    }

    return { uas: uas, membros: membros };
  } catch (e) {
    throw new Error("Erro ao buscar dados da comissão: " + e.message);
  }
}

/**
 * Grava as alterações na aba ORGANOGRAMA (Coluna F: Status | Coluna G: Responsável)
 */
function salvarEstruturaComissao(codigoUg, atribuicoes) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOrg = ss.getSheetByName("ORGANOGRAMA");
    if (!sheetOrg) throw new Error("Aba ORGANOGRAMA não encontrada.");

    var data = sheetOrg.getDataRange().getValues();
    
    // Mapeia as atribuições enviadas pelo frontend pelo Código SIPAC
    var mapAtribuicoes = {};
    atribuicoes.forEach(function(item) {
      mapAtribuicoes[String(item.codigoSipac).trim()] = item;
    });

    for (var i = 1; i < data.length; i++) {
      var codigoSipac = String(data[i][0] || '').trim();
      
      if (mapAtribuicoes[codigoSipac]) {
        var atrib = mapAtribuicoes[codigoSipac];
        var linha = i + 1;

        if (atrib.semCarga) {
          // Atualiza Coluna F (Status) e limpa Coluna G (Responsáveis)
          sheetOrg.getRange(linha, 6).setValue("Não possui carga patrimonial");
          sheetOrg.getRange(linha, 7).setValue("");
        } else {
          // Se estava marcado como sem carga anteriormente, reseta o status para Pendente
          var statusAtual = String(data[i][5] || '').trim();
          if (statusAtual === "Não possui carga patrimonial" || !statusAtual) {
            sheetOrg.getRange(linha, 6).setValue("Pendente");
          }
          
          // Grava a lista de nomes separados por vírgula na Coluna G
          var nomesConcatenados = (atrib.membrosNomes || []).join(", ");
          sheetOrg.getRange(linha, 7).setValue(nomesConcatenados);
        }
      }
    }

    return { sucesso: true, mensagem: "Atribuições salvas com sucesso!" };
  } catch (e) {
    throw new Error("Erro ao salvar no ORGANOGRAMA: " + e.message);
  }
}