function getDadosIniciaisGestorComissao(codigoUg) {
  try {
    var ss = getPlanilha();
    
    var sheetOrg = ss.getSheetByName("ORGANOGRAMA");
    var dataOrg = sheetOrg ? sheetOrg.getDataRange().getValues() : [];
    var uasUg = [];

    for (var i = 1; i < dataOrg.length; i++) {
      var sipac = String(dataOrg[i][0] || '').trim();
      var desc = String(dataOrg[i][1] || '').trim();
      var ug = String(dataOrg[i][3] || '').trim();
      var status = String(dataOrg[i][5] || 'Pendente').trim();

      if (sipac && ug === String(codigoUg).trim()) {
        uasUg.push({ codigoSipac: sipac, descricao: desc, status: status });
      }
    }

    var sheetMembros = ss.getSheetByName("MEMBROS_COMISSAO");
    var membros = [];
    if (sheetMembros) {
      var dataM = sheetMembros.getDataRange().getValues();
      for (var j = 1; j < dataM.length; j++) {
        if (String(dataM[j][0]).trim() === String(codigoUg).trim()) {
          membros.push({ 
            nome: dataM[j][1], 
            email: dataM[j][2],
            papel: dataM[j][3] || 'Membro'
          });
        }
      }
    }

    var sheetVinc = ss.getSheetByName("VINCULOS_COMISSAO");
    var vinculos = [];
    if (sheetVinc) {
      var dataV = sheetVinc.getDataRange().getValues();
      for (var k = 1; k < dataV.length; k++) {
        if (String(dataV[k][0]).trim() === String(codigoUg).trim()) {
          vinculos.push({ 
            codigoSipac: String(dataV[k][1]).trim(), 
            email: String(dataV[k][2]).trim() 
          });
        }
      }
    }

    return { uas: uasUg, membros: membros, vinculos: vinculos };
  } catch (e) {
    throw new Error("Erro ao carregar dados da comissão: " + e.message);
  }
}

function salvarEstruturaComissao(codigoUg, membros, atribuicoes) {
  try {
    var ss = getPlanilha();
    var hoje = new Date();

    if (membros && membros.length > 0) {
      salvarComissaoSetorial({ codigoUg: codigoUg, membros: membros });
    }

    var sheetVinc = ss.getSheetByName("VINCULOS_COMISSAO");
    if (!sheetVinc) {
      sheetVinc = ss.insertSheet("VINCULOS_COMISSAO");
      sheetVinc.appendRow(["codigoUg", "codigoSipacUa", "emailMembro", "dataVinculo"]);
    }

    var dataV = sheetVinc.getDataRange().getValues();
    for (var j = dataV.length - 1; j >= 1; j--) {
      if (String(dataV[j][0]).trim() === String(codigoUg).trim()) {
        sheetVinc.deleteRow(j + 1);
      }
    }

    if (atribuicoes && Array.isArray(atribuicoes)) {
      atribuicoes.forEach(function(a) {
        if (!a.semCarga && a.membros && a.membros.length > 0) {
          a.membros.forEach(function(email) {
            sheetVinc.appendRow([codigoUg, a.codigoSipac, email, hoje]);
          });
        }
      });
    }

    var sheetOrg = ss.getSheetByName("ORGANOGRAMA");
    if (sheetOrg && atribuicoes && Array.isArray(atribuicoes)) {
      var dataOrg = sheetOrg.getDataRange().getValues();
      atribuicoes.forEach(function(a) {
        for (var r = 1; r < dataOrg.length; r++) {
          if (String(dataOrg[r][0]).trim() === String(a.codigoSipac).trim()) {
            if (a.semCarga) {
              sheetOrg.getRange(r + 1, 6).setValue("Não possui carga patrimonial");
            } else if (String(dataOrg[r][5]).trim() === "Não possui carga patrimonial") {
              sheetOrg.getRange(r + 1, 6).setValue("Pendente");
            }
          }
        }
      });
    }

    return { sucesso: true, mensagem: "Atribuições da comissão salvas com sucesso!" };
  } catch (e) {
    throw new Error("Erro ao salvar estrutura da comissão: " + e.message);
  }
}