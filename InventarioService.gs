function getUasPorUg(codigoUg, nomeOuEmailUsuario) {
  try {
    // Reutiliza o contexto do usuário já existente
    var userCtx = getUserContext();
    var ss = getPlanilha();
    var sheetOrganograma = ss.getSheetByName("ORGANOGRAMA");

    if (!sheetOrganograma) {
      throw new Error("Aba 'ORGANOGRAMA' não foi encontrada.");
    }

    // Define a UG Alvo (ex: "CCJ", "GR" ou "153089")
    var targetUg = String(codigoUg || userCtx.siglaUg || userCtx.codigoUg || "").trim().toUpperCase();
    var activeEmail = userCtx.email.toLowerCase().trim();
    var nomeUser = userCtx.nome.toLowerCase().trim();
    var perfilUser = userCtx.perfil.toUpperCase().trim();

    // 1. Mapeia os Códigos e Siglas correspondentes da aba "UG" para garantir busca exata
    var sheetUg = ss.getSheetByName("UG");
    var dataUg = sheetUg ? sheetUg.getDataRange().getValues() : [];
    var codigosValidosUg = [targetUg];

    if (userCtx.codigoUg) codigosValidosUg.push(String(userCtx.codigoUg).trim().toUpperCase());
    if (userCtx.siglaUg) codigosValidosUg.push(String(userCtx.siglaUg).trim().toUpperCase());

    for (var j = 1; j < dataUg.length; j++) {
      var codUg = String(dataUg[j][0] || '').trim().toUpperCase();  // Coluna A = Código
      var siglaUg = String(dataUg[j][2] || '').trim().toUpperCase(); // Coluna C = Sigla

      if (targetUg === codUg || targetUg === siglaUg) {
        if (codUg) codigosValidosUg.push(codUg);
        if (siglaUg) codigosValidosUg.push(siglaUg);
      }
    }

    var dataOrganograma = sheetOrganograma.getDataRange().getValues();
    var mapaQtdInventariada = getMapaQtdInventariada(ss);
    var listaUas = [];

    for (var i = 1; i < dataOrganograma.length; i++) {
      var codigoSipac      = String(dataOrganograma[i][0] || "").trim(); // Coluna A
      var descricao        = String(dataOrganograma[i][1] || "").trim(); // Coluna B
      var uorgSiads        = String(dataOrganograma[i][2] || "").trim(); // Coluna C
      var ugRow            = String(dataOrganograma[i][3] || "").trim().toUpperCase(); // Coluna D
      var respUnidade      = String(dataOrganograma[i][4] || "").trim(); // Coluna E
      var statusUa         = String(dataOrganograma[i][5] || "Pendente").trim(); // Coluna F
      var respLevantamento = String(dataOrganograma[i][6] || "").trim(); // Coluna G

      var todosResponsaveis = (respUnidade + " " + respLevantamento).toLowerCase();
      var respExibicao = respLevantamento || respUnidade || "Não informado";

      // Validação ESTRITA da UG (Compara apenas o Código/Sigla com a coluna D do ORGANOGRAMA)
      var pertenceUg = (
        targetUg === "" || 
        targetUg === "TODAS" || 
        codigosValidosUg.indexOf(ugRow) !== -1
      );

      // Validação de Perfil/Responsabilidade
      var ehResponsavel = (
        perfilUser === 'PRESIDENTE' ||
        perfilUser === 'GESTOR PATRIMONIAL' ||
        perfilUser === 'GERENCIAL' ||
        perfilUser === 'ADMIN' ||
        (activeEmail !== "" && todosResponsaveis.indexOf(activeEmail) !== -1) ||
        (nomeUser !== "" && todosResponsaveis.indexOf(nomeUser) !== -1)
      );

      if (pertenceUg && ehResponsavel) {
        var codNorm = normalizarCodigo(codigoSipac);
        listaUas.push({
          codigoSipac: codigoSipac,
          descricao: descricao,
          uorgSiads: uorgSiads,
          codigoUg: ugRow,
          responsavel: respExibicao,
          status: statusUa,
          qtdInventariada: mapaQtdInventariada[codNorm] || 0
        });
      }
    }

    return listaUas;
  } catch (e) {
    Logger.log("Erro em getUasPorUg: " + e.toString());
    throw new Error("Erro ao carregar UAs: " + e.message);
  }
}


function getBensPorUa(codigoSipac) {
  try {
    var ss = getPlanilha();
    var sheet = ss.getSheetByName("INVENTARIO_BENS");

    if (!sheet) {
      throw new Error("Aba 'INVENTARIO_BENS' não foi encontrada na planilha.");
    }

    var data = sheet.getDataRange().getValues();
    var bens = [];
    
    // Remove caracteres não alfanuméricos para comparação perfeita (ex: "126944" vs "12.69.44")
    var targetClean = String(codigoSipac || "").replace(/[^0-9a-zA-Z]/g, "").trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var codUaLinha = String(data[i][3] || ""); // Coluna D: CODIGO_UA
      var codUaClean = codUaLinha.replace(/[^0-9a-zA-Z]/g, "").trim().toLowerCase();

      if (codUaClean !== "" && codUaClean === targetClean) {
        var numTombamento = String(data[i][6] || "").trim(); // Coluna G: NUMERO_TOMBAMENTO
        var numProcesso   = String(data[i][8] || "").trim(); // Coluna I: NUMERO_PROCESSO
        var tombamentoOuProcesso = numTombamento || numProcesso || "S/N";

        var rawData = data[i][1]; // Coluna B: DATA_HORA
        var dataHoraFormatada = "";
        if (rawData instanceof Date) {
          dataHoraFormatada = Utilities.formatDate(rawData, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        } else {
          dataHoraFormatada = String(rawData || "").trim();
        }

        bens.push({
          id: String(data[i][0] || "").trim(),
          dataRegistro: dataHoraFormatada,
          codigoUg: String(data[i][2] || "").trim(),
          codigoUa: codUaLinha,
          tombamento: tombamentoOuProcesso,
          descricao: String(data[i][9] || "").trim(),
          situacao: String(data[i][10] || "Pendente").trim(),
          localidade: String(data[i][11] || "").trim(),
          corresponsavel: String(data[i][12] || "").trim(),
          qtd: data[i][14] || 1
        });
      }
    }

    return bens;
  } catch (e) {
    Logger.log("Erro em getBensPorUa: " + e.toString());
    throw new Error("Falha ao carregar bens: " + e.message);
  }
}

function salvarNovoBem(dadosBem) {
  try {
    var ss = getPlanilha();
    var sheetBens = ss.getSheetByName("INVENTARIO_BENS");
    
    if (!sheetBens) {
      sheetBens = ss.insertSheet("INVENTARIO_BENS");
      sheetBens.appendRow([
        "ID", "DATA_HORA", "CODIGO_UG", "CODIGO_UA", 
        "POSSUI_TOMBAMENTO", "FORMATO_TOMBAMENTO", "NUMERO_TOMBAMENTO", 
        "POSSUI_PROCESSO", "NUMERO_PROCESSO", "DESCRICAO", 
        "SITUACAO", "LOCALIDADE", "CORRESPONSAVEL", "USER_EMAIL", "QUANTIDADE",
        "CLASSIFICACAO_SEM_PLAQUETA", "RESERVADO", "DETALHAMENTO_OUTRO"
      ]);
    }

    var rowIndex = dadosBem.rowIndex ? parseInt(dadosBem.rowIndex) : null;
    var userEmail = Session.getActiveUser().getEmail() || "sistema";
    var idBem = dadosBem.id || ("BEM_" + new Date().getTime());

    var novaLinha = [
      idBem,
      new Date(),
      dadosBem.codigoUg || '',
      normalizarCodigo(dadosBem.codigoSipacUa || dadosBem.codigoUa),
      dadosBem.possuiTombamento || 'NÃO',
      dadosBem.formatoTombamento || '',
      dadosBem.numTombamento || dadosBem.numeroTombamento || '',
      dadosBem.possuiProcesso || 'NÃO',
      dadosBem.numProcesso || dadosBem.numeroProcesso || '',
      dadosBem.descricao || '',
      dadosBem.situacao || 'Em uso',
      dadosBem.localidade || '',
      dadosBem.corresponsavel || '',
      userEmail,
      dadosBem.quantidade || 1,
      dadosBem.tipoSemPlaqueta || '',
      '',
      dadosBem.detalhamentoOutro || ''
    ];

    if (rowIndex && rowIndex > 1) {
      sheetBens.getRange(rowIndex, 1, 1, novaLinha.length).setValues([novaLinha]);
    } else {
      sheetBens.appendRow(novaLinha);
    }

    if (!dadosBem.rowIndex) {
      verificarEAtualizarStatusUa(dadosBem.codigoSipacUa || dadosBem.codigoUa);
    }

    return { sucesso: true, mensagem: "Bem salvo com sucesso!" };
  } catch (e) {
    Logger.log("Erro ao salvar bem: " + e.toString());
    throw new Error("Erro ao salvar o bem: " + e.message);
  }
}

function excluirBem(rowIndex) {
  try {
    var ss = getPlanilha();
    var sheetBens = ss.getSheetByName("INVENTARIO_BENS");
    if (!sheetBens) throw new Error("Aba INVENTARIO_BENS não encontrada!");
    
    var index = parseInt(rowIndex);
    if (index > 1 && index <= sheetBens.getLastRow()) {
      sheetBens.deleteRow(index);
      SpreadsheetApp.flush();
      return { status: "success" };
    }
    throw new Error("Linha inválida para exclusão.");
  } catch (erro) {
    Logger.log("Erro em excluirBem: " + erro.toString());
    throw new Error("Erro ao excluir bem: " + erro.message);
  }
}

function verificarEAtualizarStatusUa(codigoSipacUa) {
  var ss = getPlanilha();
  var sheetOrganograma = ss.getSheetByName("ORGANOGRAMA");
  if (!sheetOrganograma) return;

  var data = sheetOrganograma.getDataRange().getValues();
  var targetSipac = normalizarCodigo(codigoSipacUa);

  for (var i = 1; i < data.length; i++) {
    var sipacRow = normalizarCodigo(data[i][0]);
    
    if (sipacRow === targetSipac) {
      var statusAtual = String(data[i][5] || "").trim();
      
      if (statusAtual === "Pendente" || statusAtual === "") {
        sheetOrganograma.getRange(i + 1, 6).setValue("Em andamento");
      }
      break;
    }
  }
}

function getMapaQtdInventariada(ss) {
  var mapa = {};
  var sheet = ss.getSheetByName("INVENTARIO_BENS");
  
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var codNorm = normalizarCodigo(data[i][3]); 
      if (codNorm) mapa[codNorm] = (mapa[codNorm] || 0) + 1;
    }
    return mapa;
  }
  return mapa;
}

function deletarBem(id, rowIndex) {
  try {
    var ss = getPlanilha();
    var sheet = ss.getSheetByName("INVENTARIO_BENS");

    if (!sheet) {
      throw new Error("Aba 'INVENTARIO_BENS' não foi encontrada.");
    }

    var data = sheet.getDataRange().getValues();
    var targetRow = parseInt(rowIndex, 10);
    var targetId = String(id || "").trim();

    // 1. Exclusão direta pelo número exato da linha na planilha
    if (targetRow && targetRow > 1 && targetRow <= data.length) {
      sheet.deleteRow(targetRow);
      return { status: "success" };
    }

    // 2. Busca fallback pela Coluna A ou chave ROW_X
    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0] || "").trim();
      if ((rowId !== "" && rowId === targetId) || ("ROW_" + (i + 1)) === targetId) {
        sheet.deleteRow(i + 1);
        return { status: "success" };
      }
    }

    throw new Error("Registro não localizado para exclusão.");
  } catch (e) {
    Logger.log("Erro em deletarBem: " + e.toString());
    throw new Error(e.message);
  }
}

function getBemPorId(id, rowIndex) {
  try {
    var ss = getPlanilha();
    var sheet = ss.getSheetByName("INVENTARIO_BENS");
    if (!sheet) throw new Error("Aba 'INVENTARIO_BENS' não encontrada.");

    var data = sheet.getDataRange().getValues();
    var targetRow = parseInt(rowIndex, 10);
    var targetId = String(id || "").trim();

    // 1. Busca direta pelo número da linha
    if (targetRow && targetRow > 1 && targetRow <= data.length) {
      return extrairObjetoBem(data[targetRow - 1], targetRow);
    }

    // 2. Busca pela Coluna A ou ID gerado (ROW_X)
    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0] || "").trim();
      if ((rowId !== "" && rowId === targetId) || ("ROW_" + (i + 1)) === targetId) {
        return extrairObjetoBem(data[i], i + 1);
      }
    }

    return null;
  } catch (e) {
    Logger.log("Erro em getBemPorId: " + e.toString());
    throw new Error(e.message);
  }
}

function extrairObjetoBem(row, rowNum) {
  return {
    id: String(row[0] || "").trim() || ("ROW_" + rowNum),
    rowIndex: rowNum,
    codigoUg: String(row[2] || "").trim(),
    codigoUa: String(row[3] || "").trim(),
    possuiTombamento: String(row[4] || "SIM").toUpperCase().trim(),
    formatoTombamento: String(row[5] || "").trim(),
    numeroTombamento: String(row[6] || "").trim(),
    possuiProcesso: String(row[7] || "NÃO").toUpperCase().trim(),
    numeroProcesso: String(row[8] || "").trim(),
    descricao: String(row[9] || "").trim(),
    situacao: String(row[10] || "Em uso").trim(),
    localidade: String(row[11] || "").trim(),
    corresponsavel: String(row[12] || "").trim(),
    qtd: row[14] || 1,
    detalhamentoSemPlaqueta: String(row[15] || "").trim(),
    detalhamentoOutro: String(row[16] || "").trim()
  };
}

function atualizarBem(dados) {
  try {
    var ss = getPlanilha();
    var sheet = ss.getSheetByName("INVENTARIO_BENS");
    if (!sheet) throw new Error("Aba 'INVENTARIO_BENS' não encontrada.");

    var data = sheet.getDataRange().getValues();
    var targetRow = parseInt(dados.rowIndex, 10);
    var rowNum = 0;

    if (targetRow && targetRow > 1 && targetRow <= data.length) {
      rowNum = targetRow;
    } else {
      var targetId = String(dados.id || "").trim();
      for (var i = 1; i < data.length; i++) {
        var rowId = String(data[i][0] || "").trim();
        if ((rowId !== "" && rowId === targetId) || ("ROW_" + (i + 1)) === targetId) {
          rowNum = i + 1;
          break;
        }
      }
    }

    if (rowNum > 1) {
      sheet.getRange(rowNum, 5).setValue(dados.possuiTombamento);
      sheet.getRange(rowNum, 7).setValue(dados.numeroTombamento);
      sheet.getRange(rowNum, 8).setValue(dados.possuiProcesso);
      sheet.getRange(rowNum, 9).setValue(dados.numeroProcesso);
      sheet.getRange(rowNum, 10).setValue(dados.descricao);
      sheet.getRange(rowNum, 11).setValue(dados.situacao);
      sheet.getRange(rowNum, 12).setValue(dados.localidade);
      sheet.getRange(rowNum, 13).setValue(dados.corresponsavel);
      sheet.getRange(rowNum, 15).setValue(dados.qtd);
      sheet.getRange(rowNum, 16).setValue(dados.detalhamentoSemPlaqueta);
      sheet.getRange(rowNum, 17).setValue(dados.detalhamentoOutro);

      return { status: "success" };
    }

    throw new Error("Linha do registro não encontrada para atualização.");
  } catch (e) {
    Logger.log("Erro em atualizarBem: " + e.toString());
    throw new Error(e.message);
  }
}

function obterListaServidores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SERVIDORES');
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // Lê da linha 2, coluna 2 (Coluna B)
  var valores = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  
  // Transforma em lista simples e remove itens em branco
  return valores
    .map(function(row) { return row[0]; })
    .filter(function(nome) { return nome && nome.toString().trim() !== ''; });
}

/**
 * Salva o novo status da UA na aba ORGANOGRAMA (Coluna F / Coluna 6)
 * @param {string|number} codigoUA - Código da UA no SIPAC (Coluna A)
 * @param {string} novoStatus - Novo status a ser gravado (Coluna F)
 * @return {object} Resposta com status de sucesso ou erro
 */
function salvarStatusUA(codigoUA, novoStatus) {
  // Utiliza LockService para evitar inconsistências em gravações simultâneas
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Aguarda até 10s pela liberação de gravação
  } catch (e) {
    throw new Error("O sistema está ocupado. Tente salvar novamente em instantes.");
  }

  try {
    if (!codigoUA || novoStatus === undefined) {
      throw new Error("Código da UA ou Status não informados.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ORGANOGRAMA");
    
    if (!sheet) {
      throw new Error("Aba 'ORGANOGRAMA' não foi encontrada na planilha.");
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      throw new Error("Nenhum dado encontrado na aba ORGANOGRAMA.");
    }

    // Lê a Coluna A (Código da UA) e Coluna F (Status)
    var rangeCodigoUA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var codigoProcurado = String(codigoUA).trim();
    var linhaEncontrada = -1;

    // Procura a linha onde o Código da UA (Coluna A) corresponde
    for (var i = 0; i < rangeCodigoUA.length; i++) {
      var codigoLinha = String(rangeCodigoUA[i][0]).trim();
      if (codigoLinha === codigoProcurado) {
        linhaEncontrada = i + 2; // +2 compensa cabeçalho (linha 1) e índice 0
        break;
      }
    }

    if (linhaEncontrada === -1) {
      throw new Error("UA com o código '" + codigoUA + "' não foi encontrada no Organograma.");
    }

    // Coluna F = Coluna 6 no Apps Script (1-indexed: A=1, B=2, C=3, D=4, E=5, F=6)
    sheet.getRange(linhaEncontrada, 6).setValue(String(novoStatus).trim());

    return {
      sucesso: true,
      mensagem: "Status da UA atualizado com sucesso!"
    };

  } catch (e) {
    throw new Error("Erro ao salvar status da UA: " + e.message);
  } finally {
    lock.releaseLock();
  }
}