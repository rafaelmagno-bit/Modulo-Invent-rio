/**
 * Retorna os dados consolidados do painel gerencial por UG e o progresso dos bolsistas
 */
function getDadosDashboardGerencial() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Mapeia as Siglas da aba UG (Coluna A = Código [0], Coluna C = Sigla [2])
    var sheetUg = ss.getSheetByName("UG");
    var dataUg = sheetUg ? sheetUg.getDataRange().getValues() : [];
    var mapaSiglas = {};

    for (var j = 1; j < dataUg.length; j++) {
      var codUg = String(dataUg[j][0] || '').trim();
      var siglaUg = String(dataUg[j][2] || '').trim();
      if (codUg) {
        mapaSiglas[codUg] = siglaUg || codUg;
      }
    }

    // 2. Mapeia os Ofícios e Nomes dos Bolsistas na aba USUARIOS
    var sheetUsu = ss.getSheetByName("USUARIOS");
    var dataUsu = sheetUsu ? sheetUsu.getDataRange().getValues() : [];
    var mapaOficios = {};
    var bolsistasMap = {};

    for (var u = 1; u < dataUsu.length; u++) {
      var emailUsu = String(dataUsu[u][0] || '').toLowerCase().trim();
      var nomeUsu = String(dataUsu[u][1] || '').trim();
      var perfilUsu = String(dataUsu[u][2] || '').trim().toUpperCase();
      var codUgUsu = String(dataUsu[u][3] || '').trim();
      var oficioUsu = String(dataUsu[u][6] || '').trim();

      if (codUgUsu && oficioUsu) {
        mapaOficios[codUgUsu] = oficioUsu;
      }

      if (perfilUsu === 'BOLSISTA' && emailUsu) {
        bolsistasMap[emailUsu] = {
          email: emailUsu,
          nome: nomeUsu || emailUsu,
          total: 0,
          convertidos: 0,
          pendentes: 0
        };
      }
    }

    // 3. Organograma (Progresso por UG)
    var sheetOrg = ss.getSheetByName("ORGANOGRAMA");
    var dataOrg = sheetOrg ? sheetOrg.getDataRange().getValues() : [];
    var mapaQtd = getMapaQtdInventariada(ss);

    var ugsMap = {};
    var totalGeralBens = 0;
    var totalUas = 0;

    for (var i = 1; i < dataOrg.length; i++) {
      var codigoSipac = String(dataOrg[i][0]).trim();
      var codigoUg = String(dataOrg[i][3]).trim() || "NÃO INFORMADO";
      var statusUa = String(dataOrg[i][5] || "Pendente").trim();

      if (!codigoSipac) continue;

      var codNorm = normalizarCodigo(codigoSipac);
      var qtdBens = mapaQtd[codNorm] || 0;

      if (!ugsMap[codigoUg]) {
        ugsMap[codigoUg] = {
          codigoUg: codigoUg,
          siglaUg: mapaSiglas[codigoUg] || codigoUg,
          oficio: mapaOficios[codigoUg] || '-',
          totalUas: 0,
          uasConcluidas: 0,
          uasEmAndamento: 0,
          uasSemCarga: 0,
          uasPendentes: 0,
          totalBens: 0
        };
      }

      ugsMap[codigoUg].totalUas++;
      ugsMap[codigoUg].totalBens += qtdBens;
      totalGeralBens += qtdBens;
      totalUas++;

      if (statusUa === "Concluída") {
        ugsMap[codigoUg].uasConcluidas++;
      } else if (statusUa === "Em andamento") {
        ugsMap[codigoUg].uasEmAndamento++;
      } else if (statusUa === "Não possui carga patrimonial" || statusUa === "Sem carga") {
        ugsMap[codigoUg].uasSemCarga++;
      } else {
        ugsMap[codigoUg].uasPendentes++;
      }
    }

    // 4. Mapeia a conversão de tombos realizada pelos bolsistas (INVENTARIO_BENS)
    var sheetInv = ss.getSheetByName("INVENTARIO_BENS");
    var totalTombosBolsistas = 0;
    var totalTombosConvertidos = 0;

    if (sheetInv) {
      var dataInv = sheetInv.getDataRange().getValues();
      for (var k = 1; k < dataInv.length; k++) {
        var tomboOriginal = String(dataInv[k][6] || '').trim();  // Coluna G
        var bolsistaEmail = String(dataInv[k][18] || '').toLowerCase().trim(); // Coluna S
        var tomboConvertido = String(dataInv[k][19] || '').trim(); // Coluna T

        if (tomboOriginal && tomboOriginal !== '-') {
          if (bolsistaEmail) {
            if (!bolsistasMap[bolsistaEmail]) {
              bolsistasMap[bolsistaEmail] = {
                email: bolsistaEmail,
                nome: bolsistaEmail,
                total: 0,
                convertidos: 0,
                pendentes: 0
              };
            }

            bolsistasMap[bolsistaEmail].total++;
            totalTombosBolsistas++;

            if (tomboConvertido) {
              bolsistasMap[bolsistaEmail].convertidos++;
              totalTombosConvertidos++;
            } else {
              bolsistasMap[bolsistaEmail].pendentes++;
            }
          }
        }
      }
    }

    var listaUgs = Object.keys(ugsMap).map(function(key) { return ugsMap[key]; });
    var listaBolsistas = Object.keys(bolsistasMap).map(function(key) { return bolsistasMap[key]; });

    return {
      resumoGeral: {
        totalUgs: listaUgs.length,
        totalUas: totalUas,
        totalBens: totalGeralBens,
        totalTombosBolsistas: totalTombosBolsistas,
        totalTombosConvertidos: totalTombosConvertidos
      },
      ugs: listaUgs,
      bolsistas: listaBolsistas
    };
  } catch (e) {
    throw new Error("Erro ao carregar Dashboard Gerencial: " + e.message);
  }
}

/**
 * Salva ou atualiza a Comissão Setorial da UG na aba COMISSOES
 */
function salvarComissaoSetorial(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("COMISSOES");

    // Cria a aba caso ela não exista
    if (!sheet) {
      sheet = ss.insertSheet("COMISSOES");
      sheet.appendRow(["UG", "Número Portaria / Ano", "Presidente", "Membros", "Data Atualização"]);
    }

    var values = sheet.getDataRange().getValues();
    var linhaExistente = -1;

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(dados.codigoUg).trim()) {
        linhaExistente = i + 1;
        break;
      }
    }

    var linhaDados = [
      dados.codigoUg,
      dados.portaria,
      dados.presidente,
      dados.membros,
      new Date().toLocaleString("pt-BR")
    ];

    if (linhaExistente > 0) {
      sheet.getRange(linhaExistente, 1, 1, 5).setValues([linhaDados]);
    } else {
      sheet.appendRow(linhaDados);
    }

    return { sucesso: true };
  } catch (e) {
    throw new Error("Erro ao salvar comissão: " + e.message);
  }
}

/**
 * Lista as comissões cadastradas
 */
function getComissoesSetoriais() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("COMISSOES");
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var comissoes = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      comissoes.push({
        codigoUg: data[i][0],
        portaria: data[i][1],
        presidente: data[i][2],
        membros: data[i][3],
        dataAtualizacao: data[i][4]
      });
    }
  }
  return comissoes;
}

/**
 * Carrega UGs e Servidores de forma otimizada para evitar travamentos
 */
function getDadosAuxiliaresComissao() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Busca UGs (Aba UG: Col A = Código [0], Col C = Sigla [2])
    var sheetUg = ss.getSheetByName("UG");
    var ugs = [];
    if (sheetUg) {
      var dataUg = sheetUg.getDataRange().getValues();
      for (var i = 1; i < dataUg.length; i++) {
        var codigo = String(dataUg[i][0] || '').trim();
        var sigla = String(dataUg[i][2] || '').trim();
        if (codigo) {
          ugs.push({ codigo: codigo, sigla: sigla || codigo });
        }
      }
    }

    // 2. Busca apenas as colunas necessárias (B e C) de SERVIDORES para ganho de performance
    var sheetServ = ss.getSheetByName("SERVIDORES");
    var servidores = [];
    if (sheetServ) {
      var lastRow = sheetServ.getLastRow();
      if (lastRow > 1) {
        // Pega somente Coluna B (Nome) e Coluna C (SIAPE) a partir da linha 2
        var rangeServ = sheetServ.getRange(2, 2, lastRow - 1, 2).getValues();
        for (var j = 0; j < rangeServ.length; j++) {
          var nome = String(rangeServ[j][0] || '').trim();
          var siape = String(rangeServ[j][1] || '').trim();
          if (nome) {
            servidores.push({ nome: nome, siape: siape });
          }
        }
      }
    }

    return { ugs: ugs, servidores: servidores };
  } catch (e) {
    throw new Error("Erro ao carregar dados auxiliares: " + e.message);
  }
}

/**
 * Salva os membros na aba USUARIOS (Colunas A até G)
 */
function salvarMembrosComissao(listaMembros) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("USUARIOS");

    if (!sheet) {
      sheet = ss.insertSheet("USUARIOS");
      sheet.appendRow(["EMAIL", "NOME COMPLETO", "PERFIL", "CODIGO UG", "SIGLA UG", "SIAPE", "OFICIO DE INDICAÇÃO"]);
    }

    var novosRegistros = [];
    listaMembros.forEach(function(m) {
      novosRegistros.push([
        m.email,
        m.nome,
        m.perfil,
        m.codigoUg,
        m.siglaUg,
        m.siape,
        m.oficio // Grava o Ofício de Indicação na Coluna G
      ]);
    });

    if (novosRegistros.length > 0) {
      var ultimaLinha = sheet.getLastRow();
      sheet.getRange(ultimaLinha + 1, 1, novosRegistros.length, 7).setValues(novosRegistros);
    }

    return { sucesso: true, quantidade: novosRegistros.length };
  } catch (e) {
    throw new Error("Erro ao salvar membros na aba USUARIOS: " + e.message);
  }
}