/**
 * Obtém dados do relatório setorial aceitando tanto o Código quanto a Sigla da UG
 */
function obterDadosRelatorioUg(codigoUg) {
  try {
    if (!codigoUg) throw new Error("Código da UG não foi informado.");

    var ugsEquivalentes = obterUgsEquivalentes(codigoUg);
    concluirUasEmAndamento(ugsEquivalentes);

    var listaUas = getResumoRelatorioSetorial(ugsEquivalentes);
    var mapaBensPorUa = buscarBensPorUa(ugsEquivalentes);
    var membrosComissao = buscarMembrosComissaoUg(ugsEquivalentes);

    return {
      listaUas: listaUas,
      mapaBensPorUa: mapaBensPorUa,
      membros: membrosComissao
    };
  } catch (e) {
    Logger.log("Erro em obterDadosRelatorioUg: " + e.toString());
    throw new Error(e.message);
  }
}

/**
 * Busca os membros cadastrados para a UG na aba USUARIOS
 */
function buscarMembrosComissaoUg(ugsEquivalentes) {
  var membros = [];
  try {
    var ss = getPlanilha();
    var sheetUsers = ss.getSheetByName("USUARIOS");
    if (!sheetUsers) return membros;

    var dataUsers = sheetUsers.getDataRange().getValues();
    for (var i = 1; i < dataUsers.length; i++) {
      var row = dataUsers[i];
      var codU = normalizarCodigo(row[3]); // Coluna D (CODIGO DA UG)
      var sigU = normalizarCodigo(row[4]); // Coluna E (SIGLA_UG)

      if (ugsEquivalentes.indexOf(codU) !== -1 || ugsEquivalentes.indexOf(sigU) !== -1) {
        var perfil = String(row[2] || '').trim();
        var funcao = perfil === 'PRESIDENTE' ? 'Presidente da Comissão Setorial' :
                     perfil === 'SUBSTITUTO' ? 'Membro Substituto' : 'Membro da Comissão Setorial';

        membros.push({
          nome: String(row[1] || '').trim(),
          funcao: funcao,
          siape: String(row[5] || 'Servidor Ativo').trim()
        });
      }
    }
  } catch (e) {
    Logger.log("Erro em buscarMembrosComissaoUg: " + e.toString());
  }
  return membros;
}

/**
 * Mapeia os identificadores numéricos e siglas correspondentes
 */
function obterUgsEquivalentes(termoUg) {
  var termoNorm = normalizarCodigo(termoUg);
  var lista = [termoNorm];

  try {
    var ss = getPlanilha();
    
    // Busca na aba UG
    var sheetUg = ss.getSheetByName("UG");
    if (sheetUg) {
      var dataUg = sheetUg.getDataRange().getValues();
      for (var i = 1; i < dataUg.length; i++) {
        var cod = normalizarCodigo(dataUg[i][0]);
        var sigla = normalizarCodigo(dataUg[i][2] || dataUg[i][1]);
        if (cod === termoNorm || sigla === termoNorm) {
          if (cod && lista.indexOf(cod) === -1) lista.push(cod);
          if (sigla && lista.indexOf(sigla) === -1) lista.push(sigla);
        }
      }
    }

    // Busca na aba USUARIOS
    var sheetUsers = ss.getSheetByName("USUARIOS");
    if (sheetUsers) {
      var dataUsers = sheetUsers.getDataRange().getValues();
      for (var j = 1; j < dataUsers.length; j++) {
        var codU = normalizarCodigo(dataUsers[j][3]); // Coluna D
        var sigU = normalizarCodigo(dataUsers[j][4]); // Coluna E
        if (codU === termoNorm || sigU === termoNorm) {
          if (codU && lista.indexOf(codU) === -1) lista.push(codU);
          if (sigU && lista.indexOf(sigU) === -1) lista.push(sigU);
        }
      }
    }
  } catch (e) {
    Logger.log("Erro ao buscar UGs equivalentes: " + e.toString());
  }

  return lista;
}

function getResumoRelatorioSetorial(ugsEquivalentes) {
  try {
    var ss = getPlanilha();
    var sheetOrganograma = ss.getSheetByName("ORGANOGRAMA");
    if (!sheetOrganograma) throw new Error("Aba 'ORGANOGRAMA' não encontrada.");

    var sheetBens = ss.getSheetByName("INVENTARIO_BENS");
    var dataUas = sheetOrganograma.getDataRange().getValues();
    var dataBens = sheetBens ? sheetBens.getDataRange().getValues() : [];

    var mapaQtdBens = {};
    for (var j = 1; j < dataBens.length; j++) {
      var rowB = dataBens[j];
      var codUaNorm = normalizarCodigo(rowB[3]);
      var qtd = parseInt(rowB[14]) || 1;
      if (codUaNorm) {
        mapaQtdBens[codUaNorm] = (mapaQtdBens[codUaNorm] || 0) + qtd;
      }
    }

    var uasFiltradas = [];
    for (var i = 1; i < dataUas.length; i++) {
      var row = dataUas[i];
      if (!row[0] && !row[1]) continue;

      var codSipac = String(row[0]).trim();
      var codSipacNorm = normalizarCodigo(codSipac);
      var nomeUa = row[1] ? String(row[1]).trim() : '';
      var ugUaNorm = normalizarCodigo(row[3]); // Coluna D (UG)
      var statusUa = row[5] ? String(row[5]).trim() : 'Pendente';

      // Valida se a UG do organograma corresponde ao Código ou à Sigla
      if (ugsEquivalentes.indexOf(ugUaNorm) !== -1) {
        uasFiltradas.push({
          codigoSipac: codSipac,
          nomeUa: nomeUa,
          status: statusUa || 'Pendente',
          totalBens: mapaQtdBens[codSipacNorm] || 0
        });
      }
    }

    return uasFiltradas;
  } catch (e) {
    Logger.log("Erro em getResumoRelatorioSetorial: " + e.toString());
    throw new Error(e.message);
  }
}

function buscarBensPorUa(ugsEquivalentes) {
  try {
    var ss = getPlanilha();
    var sheetBens = ss.getSheetByName("INVENTARIO_BENS");
    var mapaBensPorUa = {};
    if (!sheetBens) return mapaBensPorUa;

    var dadosBens = sheetBens.getDataRange().getValues();

    for (var i = 1; i < dadosBens.length; i++) {
      var row = dadosBens[i];
      var ugLinhaNorm = normalizarCodigo(row[2]);
      var codUa = String(row[3] || '').trim();

      if (!codUa) continue;
      if (ugsEquivalentes.indexOf(ugLinhaNorm) === -1) continue;

      if (!mapaBensPorUa[codUa]) mapaBensPorUa[codUa] = [];

      mapaBensPorUa[codUa].push({
        tombamento: String(row[6] || 'S/N').trim(),
        descricao: String(row[9] || 'Sem Descrição').trim(),
        localizacao: String(row[11] || 'Não informada').trim(),
        status: String(row[10] || 'Inventariado').trim()
      });
    }

    return mapaBensPorUa;
  } catch (e) {
    return {};
  }
}

function concluirUasEmAndamento(ugsEquivalentes) {
  try {
    var ss = getPlanilha();
    var sheetOrganograma = ss.getSheetByName("ORGANOGRAMA");
    if (!sheetOrganograma) return { status: "success" };

    var range = sheetOrganograma.getDataRange();
    var dataUas = range.getValues();
    var alterou = false;

    for (var i = 1; i < dataUas.length; i++) {
      var ugUaNorm = normalizarCodigo(dataUas[i][3]);
      var statusAtual = dataUas[i][5] ? String(dataUas[i][5]).trim() : '';

      if (ugsEquivalentes.indexOf(ugUaNorm) !== -1 && statusAtual === "Em andamento") {
        dataUas[i][5] = "Concluída";
        alterou = true;
      }
    }

    if (alterou) {
      range.setValues(dataUas);
      SpreadsheetApp.flush();
    }

    return { status: "success" };
  } catch (e) {
    throw new Error(e.message);
  }
}

function atualizarStatusUa(codigoSipac, novoStatus) {
  try {
    var ss = getPlanilha();
    var sheetOrganograma = ss.getSheetByName("ORGANOGRAMA");
    if (!sheetOrganograma) throw new Error("Aba ORGANOGRAMA não encontrada.");

    var dataUas = sheetOrganograma.getDataRange().getValues();
    var targetSipacNorm = normalizarCodigo(codigoSipac);

    for (var i = 1; i < dataUas.length; i++) {
      var codSipacNorm = normalizarCodigo(dataUas[i][0]);
      
      if (codSipacNorm === targetSipacNorm) {
        sheetOrganograma.getRange(i + 1, 6).setValue(novoStatus);
        SpreadsheetApp.flush();
        return { status: "success", codigoSipac: codigoSipac, novoStatus: novoStatus };
      }
    }

    throw new Error("Unidade com código SIPAC " + codigoSipac + " não foi encontrada.");
  } catch (e) {
    Logger.log("Erro em atualizarStatusUa: " + e.toString());
    throw new Error(e.message);
  }
}