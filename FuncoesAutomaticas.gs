/**
 * Preenche a Coluna C (UG) na aba INVENTARIO_BENS buscando
 * a correspondência da Coluna D (UA) na aba ORGANOGRAMA.
 */
function preencherUgNoInventario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Mapeia UA -> UG a partir da aba ORGANOGRAMA
  var sheetOrg = ss.getSheetByName("ORGANOGRAMA");
  if (!sheetOrg) {
    Logger.log("Aba ORGANOGRAMA não encontrada.");
    return;
  }

  var dataOrg = sheetOrg.getDataRange().getValues();
  var mapaUaUg = {};

  // Pula o cabeçalho (linha 1)
  for (var i = 1; i < dataOrg.length; i++) {
    var codUa = String(dataOrg[i][0] || '').trim(); // Coluna A (UA)
    var codUg = String(dataOrg[i][3] || '').trim(); // Coluna D (UG)

    if (codUa) {
      mapaUaUg[codUa] = codUg;
    }
  }

  // 2. Atualiza a coluna C da aba INVENTARIO_BENS
  var sheetInv = ss.getSheetByName("INVENTARIO_BENS");
  if (!sheetInv) {
    Logger.log("Aba INVENTARIO_BENS não encontrada.");
    return;
  }

  var lastRow = sheetInv.getLastRow();
  if (lastRow <= 1) return; // Retorna se houver apenas o cabeçalho

  // Pega da linha 2 até o fim, pegando as Colunas C (UG) e D (UA)
  var rangeInv = sheetInv.getRange(2, 3, lastRow - 1, 2);
  var valuesInv = rangeInv.getValues();

  var colunaUgAtualizada = [];

  for (var j = 0; j < valuesInv.length; j++) {
    var ugAtual = valuesInv[j][0];                       // Coluna C
    var codUaInv = String(valuesInv[j][1] || '').trim(); // Coluna D

    // Se a UA existir no mapa do Organograma, atualiza. Senão, preserva o valor atual.
    if (mapaUaUg[codUaInv]) {
      colunaUgAtualizada.push([mapaUaUg[codUaInv]]);
    } else {
      colunaUgAtualizada.push([ugAtual]);
    }
  }

  // 3. Grava o resultado na Coluna C (UG) de uma só vez (linha 2 em diante)
  if (colunaUgAtualizada.length > 0) {
    sheetInv.getRange(2, 3, colunaUgAtualizada.length, 1).setValues(colunaUgAtualizada);
  }
}