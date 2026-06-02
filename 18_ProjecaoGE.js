// =============================================================================
// 18_ProjecaoGE.js — Projeção minuto-a-minuto re-ancorada (Gazeta Esportiva)
// =============================================================================

var PROJECAO_AUDIENCIA_GE = {
    DATA_START_ROW: 3,
    DATA_END_ROW: 62,
    AUD_COL: 4,
    TEMA_COL: 6,
    FIRST_HIST_AUD_COL: 8,
    HEADERS: ["MOLDE DIA", "PASSO", "BETA FAIXA", "PROJ"]
};

/**
 * Escreve/atualiza as fórmulas de projeção de audiência re-ancorada
 * na aba GE - Histórico (coluna D, linhas 3-62), criando/atualizando as
 * 4 colunas auxiliares (MOLDE DIA, PASSO, BETA FAIXA, PROJ) à direita.
 */
function instalarProjecaoAudienciaGE_() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) throw Error("Aba não encontrada: " + CFG.GE_OUT_HISTORICO);
    if (sheet.getMaxRows() < PROJECAO_AUDIENCIA_GE.DATA_END_ROW) {
        throw Error("Aba " + CFG.GE_OUT_HISTORICO + " precisa ter ao menos " + PROJECAO_AUDIENCIA_GE.DATA_END_ROW + " linhas para instalar a projeção.");
    }

    var auxStartCol = encontrarInicioAuxProjecaoGE_(sheet);
    if (!auxStartCol) auxStartCol = sheet.getLastColumn() + 2;
    garantirColunasProjecaoGE_(sheet, auxStartCol + PROJECAO_AUDIENCIA_GE.HEADERS.length - 1);

    sheet.getRange(2, auxStartCol, 1, PROJECAO_AUDIENCIA_GE.HEADERS.length)
        .setValues([PROJECAO_AUDIENCIA_GE.HEADERS])
        .setFontWeight("bold")
        .setBackground("#5B2C6F")
        .setFontColor("white")
        .setHorizontalAlignment("center");

    var moldeCol = auxStartCol;
    var passoCol = auxStartCol + 1;
    var betaCol = auxStartCol + 2;
    var projCol = auxStartCol + 3;

    var moldeL = columnToLetter_(moldeCol);
    var betaL = columnToLetter_(betaCol);
    var projL = columnToLetter_(projCol);
    var temaSheetRef = nomeAbaFormulaProjecaoGE_(CFG.GE_OUT_MEDIA_TEMA);

    var audColsHist = colunasAudienciaHistoricasProjecaoGE_(sheet);
    var moldeF = [];
    var passoF = [];
    var betaF = [];
    var projF = [];

    for (var r = PROJECAO_AUDIENCIA_GE.DATA_START_ROW; r <= PROJECAO_AUDIENCIA_GE.DATA_END_ROW; r++) {
        // MOLDE DIA usa a opção preferida: média das audiências históricas fechadas
        // a partir de H (exclui D, dia atual). Assim D pode receber projeções/reais
        // sem realimentar B→MOLDE→PROJ→D durante o programa ao vivo.
        var moldeBase = formulaMoldeBaseHistoricoProjecaoGE_(audColsHist, r);
        moldeF.push(["=IF(OR(" + moldeBase + '="",' + moldeBase + "=0),0.0001," + moldeBase + "*IFERROR(VLOOKUP(F" + r + "," + temaSheetRef + "!$A:$F,6,0),1))"]);

        if (r === PROJECAO_AUDIENCIA_GE.DATA_START_ROW) {
            passoF.push(["=1"]);
            projF.push(['=IF(D' + r + '="","",D' + r + ')']);
        } else {
            passoF.push(["=IF(OR(" + moldeL + (r - 1) + '="",' + moldeL + (r - 1) + "=0),1," + moldeL + r + "/" + moldeL + (r - 1) + ")"]);
            projF.push(["=IF(D" + (r - 1) + '="","",ROUND(' + moldeL + r + "*(1+(D" + (r - 1) + "/" + moldeL + (r - 1) + "-1)*" + betaL + r + "),2))"]);
        }

        // Coeficientes BETA medidos no histórico real (105 dias, AR(1) do desvio
        // relativo à média; meia-vida aproximada de 16 minutos), por faixa horária.
        betaF.push(["=IF(ROW()<=23,0.9478,IF(ROW()<=48,0.9677,0.9552))"]);
    }

    batchSetFormulasSafe_(sheet.getRange(PROJECAO_AUDIENCIA_GE.DATA_START_ROW, moldeCol, 60, 1), moldeF);
    batchSetFormulasSafe_(sheet.getRange(PROJECAO_AUDIENCIA_GE.DATA_START_ROW, passoCol, 60, 1), passoF);
    batchSetFormulasSafe_(sheet.getRange(PROJECAO_AUDIENCIA_GE.DATA_START_ROW, betaCol, 60, 1), betaF);
    batchSetFormulasSafe_(sheet.getRange(PROJECAO_AUDIENCIA_GE.DATA_START_ROW, projCol, 60, 1), projF);

    restaurarFormulasColunaDProjecaoGE_(sheet, projL);

    SpreadsheetApp.flush();
    sheet.getRange(PROJECAO_AUDIENCIA_GE.DATA_START_ROW, PROJECAO_AUDIENCIA_GE.AUD_COL, 60, 1).setNumberFormat("0.00");
    sheet.getRange(PROJECAO_AUDIENCIA_GE.DATA_START_ROW, projCol, 60, 1).setNumberFormat("0.00");
    sheet.getRange(PROJECAO_AUDIENCIA_GE.DATA_START_ROW, moldeCol, 60, 3).setNumberFormat("0.0000");

    [100, 80, 90, 80].forEach(function(w, i) {
        sheet.setColumnWidth(auxStartCol + i, w);
    });

    return {
        auxStartCol: auxStartCol,
        auxProjCol: projCol,
        audColsHistoricas: audColsHist.length
    };
}

/** Entrada pública via menu (com alert de confirmação/sucesso). */
function instalarProjecaoAudienciaGE() {
    var ui = SpreadsheetApp.getUi();
    try {
        var result = instalarProjecaoAudienciaGE_();
        ui.alert(
            "✅ Projeção instalada/atualizada",
            "As fórmulas de projeção de audiência foram atualizadas em " + CFG.GE_OUT_HISTORICO + ".\n\n" +
            "Auxiliares a partir da coluna " + columnToLetter_(result.auxStartCol) + ".\n" +
            "Dias históricos usados no molde: " + result.audColsHistoricas + ".",
            ui.ButtonSet.OK
        );
    } catch (err) {
        ui.alert("❌ Erro ao instalar projeção", err && err.message ? err.message : String(err), ui.ButtonSet.OK);
        throw err;
    }
}

function encontrarInicioAuxProjecaoGE_(sheet) {
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
        if ((headers[i] || "").toString().trim() === PROJECAO_AUDIENCIA_GE.HEADERS[0]) return i + 1;
    }
    return null;
}

function encontrarColunaProjProjecaoGE_(sheet) {
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
        if ((headers[i] || "").toString().trim() === "PROJ") return i + 1;
    }
    return null;
}

function garantirColunasProjecaoGE_(sheet, requiredCol) {
    if (sheet.getMaxColumns() < requiredCol) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCol - sheet.getMaxColumns());
    }
}

function colunasAudienciaHistoricasProjecaoGE_(sheet) {
    var audCols = (typeof getAudColsGE_ === "function") ? getAudColsGE_(sheet) : [];
    if (!audCols || audCols.length === 0) {
        var lastCol = sheet.getLastColumn();
        for (var c = PROJECAO_AUDIENCIA_GE.AUD_COL; c <= lastCol; c += 4) audCols.push(c);
    }
    return audCols.filter(function(c) { return c >= PROJECAO_AUDIENCIA_GE.FIRST_HIST_AUD_COL; });
}

function formulaMoldeBaseHistoricoProjecaoGE_(audColsHist, row) {
    if (!audColsHist || audColsHist.length === 0) return "B" + row;
    var refs = audColsHist.map(function(c) { return columnToLetter_(c) + row; });
    return "(" + refs.join("+") + ")/" + audColsHist.length;
}

function nomeAbaFormulaProjecaoGE_(sheetName) {
    return "'" + (sheetName || "").toString().replace(/'/g, "''") + "'";
}

function restaurarFormulasColunaDProjecaoGE_(sheet, projL) {
    var dRange = sheet.getRange(PROJECAO_AUDIENCIA_GE.DATA_START_ROW, PROJECAO_AUDIENCIA_GE.AUD_COL, 60, 1);
    var values = dRange.getValues();
    var currentFormulas = dRange.getFormulas();
    var formulas = [];
    var ranges = [];
    for (var i = 0; i < values.length; i++) {
        var row = PROJECAO_AUDIENCIA_GE.DATA_START_ROW + i;
        var v = values[i][0];
        var hasFormula = !!currentFormulas[i][0];
        var isBlankOrZero = v === "" || v === null || v === 0 || v === "0";
        if (!hasFormula && !isBlankOrZero) continue;

        if (row === PROJECAO_AUDIENCIA_GE.DATA_START_ROW) {
            if (isBlankOrZero || hasFormula) sheet.getRange(row, PROJECAO_AUDIENCIA_GE.AUD_COL).clearContent();
            continue;
        }
        ranges.push(row);
        formulas.push(['=IF(D' + (row - 1) + '="","",' + projL + row + ')']);
    }

    for (var j = 0; j < ranges.length; j++) {
        batchSetFormulasSafe_(sheet.getRange(ranges[j], PROJECAO_AUDIENCIA_GE.AUD_COL, 1, 1), [formulas[j]]);
    }
}

function restaurarCadeiaProjecaoAudienciaGEOnEdit_(sheet, editedRow) {
    var projCol = encontrarColunaProjProjecaoGE_(sheet);
    if (!projCol) return;
    var projL = columnToLetter_(projCol);
    var startRow = Math.max(editedRow + 1, PROJECAO_AUDIENCIA_GE.DATA_START_ROW + 1);
    if (startRow > PROJECAO_AUDIENCIA_GE.DATA_END_ROW) return;

    var numRows = PROJECAO_AUDIENCIA_GE.DATA_END_ROW - startRow + 1;
    var range = sheet.getRange(startRow, PROJECAO_AUDIENCIA_GE.AUD_COL, numRows, 1);
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
        if (values[i][0] !== "" && values[i][0] !== null) continue;
        var row = startRow + i;
        // onEdit simples pode não ter autorização para Sheets API v4. Somente aqui
        // usamos SpreadsheetApp.setFormula em pt-BR para reconstruir células D vazias;
        // a instalação/recalibragem manual continua usando USER_ENTERED via helpers safe_.
        sheet.getRange(row, PROJECAO_AUDIENCIA_GE.AUD_COL).setFormula('=SE(D' + (row - 1) + '="";"";' + projL + row + ')');
    }
}
