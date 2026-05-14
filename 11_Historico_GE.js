// =============================================================================
// 11_Historico_GE.js — Histórico, dashboards e médias (Gazeta Esportiva)
// =============================================================================

// --- Helpers de data e detecção de colunas ---

function parseDateStr_(s) {
    if (!s || typeof s !== "string") return null;
    var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    return m ? new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])) : null;
}

// Retorna apenas as colunas de audiência reais (com data no row 1), excluindo CALIBRAGEM
function getAudColsGE_(sheet) {
    var lastCol = sheet.getLastColumn();
    var audCols = [];
    for (var c = 4; c <= lastCol; c += 4) {
        var h = sheet.getRange(1, c).getValue();
        var hasDate = (h instanceof Date) ||
            (typeof h === "string" && /^\d{2}\/\d{2}\/\d{4}$/.test(h.trim()));
        if (hasDate) audCols.push(c);
        else break;
    }
    return audCols;
}

// --- Arquivar programa GE (entrada pública via menu) ---

function arquivarProgramaAtualGazeta() {
    const e = SpreadsheetApp.getUi();
    if (e.alert("\u{1f4c1} Arquivar Programa", "Deseja arquivar o programa atual do Gazeta Esportiva no hist\xf3rico?", e.ButtonSet.YES_NO) === e.Button.YES) try {
        const t = extrairDataProgramaGE_() || new Date;
        if (arquivarProgramaGE_(t)) {
            const o = Utilities.formatDate(t, "America/Sao_Paulo", "dd/MM/yyyy");
            e.alert("\u2705 Arquivado!", `Programa de ${o} arquivado no hist\xf3rico do Gazeta.`, e.ButtonSet.OK)
        } else e.alert("\u2139\ufe0f J\xe1 existe", "Este programa j\xe1 est\xe1 arquivado no hist\xf3rico.", e.ButtonSet.OK)
    } catch (t) {
        e.alert("\u274c Erro", "Erro ao arquivar: " + t.message, e.ButtonSet.OK)
    }
}

// --- Criação / inicialização da aba Histórico GE ---

function ensureHistoricoSheetGE_() {
    const e = SpreadsheetApp.getActive();
    let t = e.getSheetByName(CFG.GE_OUT_HISTORICO);
    const o = !t;
    t || (t = e.insertSheet(CFG.GE_OUT_HISTORICO)), t.getRange(1, 1).setValue("").setBackground("#E74C3C"), t.getRange(1, 2).setValue("").setBackground("#E74C3C"), t.getRange(1, 3).setValue("").setBackground("#E74C3C"), t.getRange(2, 1).setValue("HORA").setFontWeight("bold").setBackground("#E74C3C").setFontColor("white"), t.getRange(2, 2).setValue("M\xc9DIA").setFontWeight("bold").setBackground("#E74C3C").setFontColor("white"), t.getRange(2, 3).setValue("% A M\xc9DIA").setFontWeight("bold").setBackground("#E74C3C").setFontColor("white");
    const a = [];
    for (let e = 0; e <= 59; e++) a.push(["18:" + (e + "").padStart(2, "0")]);
    return t.getRange(3, 1, 60, 1).setValues(a), t.setColumnWidth(1, 60), t.setColumnWidth(2, 70), t.setColumnWidth(3, 90), t.setFrozenColumns(3), t.setFrozenRows(2), o && Logger.log(`Aba "${CFG.GE_OUT_HISTORICO}" criada com estrutura correta.`), t
}

function localizarBlocoHistoricoGEPorData_(programDate) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet || !programDate) return null;
    const target = Utilities.formatDate(programDate, CFG.TIMEZONE, "dd/MM/yyyy");
    const lastCol = sheet.getLastColumn();
    if (lastCol <= 3) return null;
    const headers = sheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
    for (let i = 0; i < headers.length; i += 4) {
        const value = headers[i];
        if (!value) continue;
        const dateStr = value instanceof Date ? Utilities.formatDate(value, CFG.TIMEZONE, "dd/MM/yyyy") : value.toString().trim();
        if (dateStr === target) return {
            sheet: sheet,
            dateStr: dateStr,
            startCol: 4 + i,
            audienciaCol: 4 + i,
            shareCol: 5 + i,
            temaCol: 6 + i,
            termosCol: 7 + i
        };
    }
    return null;
}

function obterTermosMinutoAMinutoGE_() {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_TIME);
    if (!sheet || sheet.getLastRow() < 61) throw Error('A aba "' + CFG.GE_OUT_TIME + '" não possui o intervalo E2:E61 completo.');
    return sheet.getRange(2, 5, 60, 1).getValues();
}

function sincronizarTermosMinutoAMinutoParaHistoricoGE_(programDate) {
    const block = localizarBlocoHistoricoGEPorData_(programDate);
    if (!block) return Logger.log("[Speech] Histórico GE sem bloco para " + Utilities.formatDate(programDate, CFG.TIMEZONE, "dd/MM/yyyy")), {
        status: "missing_block"
    };
    const termos = obterTermosMinutoAMinutoGE_();
    block.sheet.getRange(3, block.termosCol, 60, 1).setValues(termos);
    Logger.log("[Speech] TERMOS sincronizados para GE - Histórico em " + block.dateStr + " (coluna " + block.termosCol + ").");
    return {
        status: "updated",
        dateStr: block.dateStr,
        termosCol: block.termosCol
    };
}

// --- Arquivar programa GE (interno, chamado automaticamente pelo runAllGazeta) ---

function arquivarProgramaGE_(programDate) {
    var ss = SpreadsheetApp.getActive();
    var hist = ensureHistoricoSheetGE_();
    var timeSheet = ss.getSheetByName(CFG.GE_OUT_TIME);
    if (!timeSheet || timeSheet.getLastRow() < 2) throw Error("Aba GE - Minuto a Minuto está vazia");

    var dateStr = Utilities.formatDate(programDate, "America/Sao_Paulo", "dd/MM/yyyy");

    // Verificar duplicata
    var lastCol = hist.getLastColumn();
    if (lastCol > 3) {
        var headers = hist.getRange(1, 4, 1, lastCol - 3).getValues()[0];
        for (var i = 0; i < headers.length; i++) {
            var h = headers[i];
            if (!h) continue;
            var hStr = h instanceof Date
                ? Utilities.formatDate(h, "America/Sao_Paulo", "dd/MM/yyyy")
                : h.toString().trim();
            if (hStr === dateStr) {
                Logger.log("Programa GE de " + dateStr + " já existe - pulando");
                return false;
            }
        }
    }

    // Inserir 4 colunas
    hist.insertColumns(4, 4);

    var numExisting = Math.floor((lastCol - 3) / 4);
    var bgColor = (numExisting % 2 === 0) ? "#FDF2F2" : "#FFFFFF";

    hist.getRange(1, 4, 1, 4).merge();
    hist.getRange(1, 4).setValue(dateStr)
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setBackground("#FADBD8")
        .setFontColor("#000000");

    hist.getRange(2, 4, 1, 4).setValues([["AUDIÊNCIA", "&/MIN", "TEMA", "TERMOS"]]);
    hist.getRange(2, 4, 1, 4)
        .setFontWeight("bold")
        .setBackground("#E74C3C")
        .setFontColor("white")
        .setHorizontalAlignment("center");

    hist.setColumnWidth(4, 90);
    hist.setColumnWidth(5, 70);
    hist.setColumnWidth(6, 120);
    hist.setColumnWidth(7, 200);

    hist.getRange(3, 4, 60, 4).setBackground(bgColor);

    // Ler dados do Minuto a Minuto
    var lastRow = timeSheet.getLastRow();
    var numCols = Math.min(6, timeSheet.getLastColumn());
    var data = timeSheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var minuteMap = new Map();
    for (var r = 0; r < data.length; r++) {
        var rowData = data[r];
        var timeStr = "";
        if (rowData[0] instanceof Date) {
            var dt = rowData[0];
            timeStr = (dt.getHours() + "").padStart(2, "0") + ":" + (dt.getMinutes() + "").padStart(2, "0");
        } else if (rowData[0]) {
            var s = rowData[0].toString().trim();
            var m = s.match(/(\d{1,2}):(\d{2})/);
            timeStr = m ? (parseInt(m[1]) + "").padStart(2, "0") + ":" + m[2] : s;
        }
        var tema = rowData[1] ? rowData[1].toString().trim() : "";
        var termos = rowData[4] ? rowData[4].toString().trim() : "";
        if (timeStr) minuteMap.set(timeStr, { tema: tema, termos: termos });
    }

    var values = [];
    for (var min = 0; min <= 59; min++) {
        var t = "18:" + (min + "").padStart(2, "0");
        var entry = minuteMap.get(t);
        values.push(["", "", entry ? entry.tema : "", entry ? entry.termos : ""]);
    }
    hist.getRange(3, 4, values.length, 4).setValues(values);

    hist.getRange(3, 4, 60, 1).setNumberFormat("0.00");
    hist.getRange(3, 5, 60, 1).setNumberFormat("0.00");

    var shareFormulas = [];
    for (var row = 3; row <= 62; row++) {
        shareFormulas.push(['=SE($B' + row + '=0;"";D' + row + '/$B' + row + ')']);
    }
    batchSetFormulasSafe_(hist.getRange(3, 5, 60, 1), shareFormulas);

    hist.getRange(63, 4, 1, 4).setBackground("#F5B7B1");
    hist.getRange(63, 4).setNumberFormat("0.00").setFontWeight("bold");

    atualizarFormulasHistoricoGE_(hist);
    Utilities.sleep(2000);
    garantirCamposHistoricoGE_();
    atualizarMediaSemanaGE_();
    Utilities.sleep(1000);
    atualizarGraficoMediasGE_();
    Utilities.sleep(1000);
    atualizarMediaPorComentaristaGE_();
    atualizarMediaPorDiaSemanaGE_();
    atualizarMediaPorTemaSeExistir_();

    Logger.log("Programa GE de " + dateStr + " arquivado com sucesso");
    return true;
}

// --- Atualizar fórmulas do histórico GE ---

function atualizarFormulasHistoricoGE_(sheet) {
    var audCols = getAudColsGE_(sheet);
    if (audCols.length === 0) return;
    var numPrograms = audCols.length;

    var bcFormulas = [];
    for (var row = 3; row <= 62; row++) {
        var sumParts = audCols.map(function(c) { return columnToLetter_(c) + row; }).join("+");
        var fB = "=(" + sumParts + ")/" + numPrograms;
        var fC = '=SE(OU(B' + row + '="";B' + row + '=0;D' + row + '="");"";D' + row + '/B' + row + '-1)';
        bcFormulas.push([fB, fC]);
    }
    batchSetFormulasSafe_(sheet.getRange(3, 2, 60, 2), bcFormulas);

    var row63Formulas = [];
    for (var i = 0; i < audCols.length; i++) {
        var colL = columnToLetter_(audCols[i]);
        row63Formulas.push(["=(SOMA(" + colL + "3:" + colL + "62))/60"]);
    }
    batchSetFormulasSafe_(sheet.getRange(63, 4, row63Formulas.length, 1), row63Formulas);
    for (var i = 0; i < audCols.length; i++) {
        sheet.getRange(63, audCols[i]).setNumberFormat("0.00");
    }

    setFormulaSafe_(sheet.getRange(63, 2), "=(SOMA(B3:B62))/60");
    sheet.getRange(63, 2).setNumberFormat("0.00");
    setFormulaSafe_(sheet.getRange(63, 3), '=SE(OU(B63="";B63=0;D63="");"";D63/B63-1)');
    sheet.getRange(63, 3).setNumberFormat("0.0%");

    SpreadsheetApp.flush();
    sheet.getRange(3, 2, 60, 1).setNumberFormat("0.00");
    sheet.getRange(3, 3, 60, 1).setNumberFormat("0.0%");

    atualizarCalibragemGE_(sheet);
    atualizarFatorTemaGE_(sheet);
}

// --- Garantir campos de rodapé do histórico GE ---

function garantirCamposHistoricoGE_() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    var maxRows = sheet.getMaxRows();
    if (maxRows < 89) sheet.insertRowsAfter(maxRows, 89 - maxRows);
    var labels = [["MÉDIA"], ["COMENT. 1"], ["COMENT. 2"], ["MÉDIA SEM."], ["% vs ANT."]];
    sheet.getRange(63, 1, 5, 1).setValues(labels);
    sheet.getRange(63, 1, 5, 1).setFontWeight("bold").setFontSize(8).setFontColor("#666666");
    preencherBlocoA66E85HistoricoGE_(sheet);
}

function preencherBlocoA66E85HistoricoGE_(sheet) {
    // Área-alvo oficial: A66:NT89 (layout/fórmulas conforme docs/ge-historico/*_A66_NT89*)
    var fullTarget = sheet.getRange("A66:NT89");
    fullTarget.clearContent();
    fullTarget.clearFormat();
    fullTarget.clearDataValidations();
    fullTarget.clearNote();

    var fixedValues = [
        ["SEMANA 19", "", "", "", ""],
        ["SEMANA 18", "", "", "", ""],
        ["SEMANA 17", "", "", "", ""],
        ["SEMANA 16", "", "", "", ""],
        ["SEMANA 15", "", "", "", ""],
        ["SEMANA 14", "", "", "", ""],
        ["SEMANA 13", "", "", "", ""],
        ["SEMANA 12", "", "", "", ""],
        ["SEMANA 11", "", "", "", ""],
        ["SEMANA 10", "", "", "", ""],
        ["SEMANA 09", "", "", "", ""],
        ["SEMANA 08", "", "", "", ""],
        ["SEMANA 07", "", "", "", ""],
        ["SEMANA 06", "", "", "", ""],
        ["SEMANA 05", "", "", "", ""],
        ["SEMANA 04", "", "", "", ""],
        ["SEMANA 03", "", "", "", ""],
        ["SEMANA 02", "", "", "", ""],
        ["SEMANA 01", "", "-", "MENSAL", ""],
        ["MAIO", "", "", 0.65, ""],
        ["ABRIL", "", "", 0.70, ""],
        ["MARÇO", "", "", 0.71, ""],
        ["FEVEREIRO", "", "", 0.63, ""],
        ["JANEIRO", "", "-", 0.56, ""]
    ];

    var panel = sheet.getRange("A66:E89");
    panel.setValues(fixedValues);

    // Fórmulas dinâmicas oficiais (docs/ge-historico/formulas_A66_NT89.md)
    sheet.getRange("B66").setFormula("=D63");
    sheet.getRange("C66").setFormula("=(B66/B67)-1");

    sheet.getRange("B67").setFormula("=MÉDIA(H63;L63;P63;T63;X63;J65)");
    sheet.getRange("B68").setFormula("=MÉDIA(AB63;AF63;AJ63;AN63;AR63)");
    sheet.getRange("B69").setFormula("=MÉDIA(AV63;AZ63;BD63;BH63;BL63)");
    sheet.getRange("B70").setFormula("=MÉDIA(BP63;BT63;BX63;CB63;CF63)");
    sheet.getRange("B71").setFormula("=MÉDIA(CJ63;CN63;CR63;CV63;CZ63)");
    sheet.getRange("B72").setFormula("=MÉDIA(DD63;DH63;DL63;DP63;DT63)");
    sheet.getRange("B73").setFormula("=MÉDIA(DX63;EB63;EF63;EJ63;EN63)");
    sheet.getRange("B74").setFormula("=MÉDIA(ER63;EV63;EZ63;FD63;FH63)");
    sheet.getRange("B75").setFormula("=MÉDIA(FL63;FP63;FT63;FX63;GB63)");
    sheet.getRange("B76").setFormula("=MÉDIA(GF63;GJ63;GN63;GR63;GV63)");
    sheet.getRange("B77").setFormula("=MÉDIA(GZ63;HD63;HH63;HL63;HP63)");
    sheet.getRange("B78").setFormula("=MÉDIA(HT63;HX63;IB63;IF63;IJ63)");
    sheet.getRange("B79").setFormula("=MÉDIA(IN63;IR63;IV63;IZ63;JD63)");
    sheet.getRange("B80").setFormula("=MÉDIA(JH63;JL63;JP63;JT63;JX63)");
    sheet.getRange("B81").setFormula("=MÉDIA(KB63;KF63;KJ63;KN63;KR63)");
    sheet.getRange("B82").setFormula("=MÉDIA(KV63;KZ63;LD63;LH63;LL63)");
    sheet.getRange("B83").setFormula("=MÉDIA(LP63;LT63;LX63;MB63;MF63)");
    sheet.getRange("B84").setFormula("=MÉDIA(MJ63;MN63;MR63;MV63;MZ63)");

    sheet.getRange("B85").setFormula("=MÉDIA(D63;H63;L63;P63;T63;X63;AB63)");
    sheet.getRange("C85").setFormula("=(B85/B86)-1");
    sheet.getRange("E85").setFormula("=(B85/D85)-1");

    sheet.getRange("B86").setFormula("=MÉDIA(DL63;DH63;DD63;CZ63;CV63;CR63;CN63;CJ63;CF63;CB63;BX63;BT63;BP63;BL63;BH63;BD63;AZ63;AV63;AR63;AN63;AJ63;AF63;AB63)");
    sheet.getRange("B87").setFormula("=MÉDIA(GV63;GR63;GN63;GJ63;GF63;GB63;FX63;FT63;FP63;FL63;FH63;FD63;EZ63;EV63;ER63;EN63;EJ63;EF63;EB63;DX63;DT63;DP63)");
    sheet.getRange("B88").setFormula("=MÉDIA(JX63;JT63;JP63;JL63;JH63;JD63;IZ63;IV63;IR63;IN63;IJ63;IF63;IB63;HX63;HT63;HP63;HL63;HH63;HD63;GZ63)");
    sheet.getRange("B89").setFormula("=MÉDIA(MZ63;MV63;MR63;MN63;MJ63;MF63;MB63;LX63;LT63;LP63;LL63;LH63;LD63;KZ63;KV63;KR63;KN63;KJ63;KF63;KB63)");

    // Aparência: base no estilo do bloco vizinho e alinhamento do modelo oficial
    var model = sheet.getRange(63, 1, 1, 5);
    model.copyTo(panel, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    panel.setVerticalAlignment("middle");
    sheet.getRange("A66:A89").setHorizontalAlignment("left");
    sheet.getRange("B66:E89").setHorizontalAlignment("center");

    // Formatação numérica alinhada ao modelo A66:E89
    sheet.getRange("B66:B89").setNumberFormat("0.00");
    sheet.getRange("C66:C89").setNumberFormat("0.00%");
    sheet.getRange("D85:D89").setNumberFormat("0.00");
    sheet.getRange("E85:E89").setNumberFormat("0.00%");

    panel.setBorder(true, true, true, true, true, true);
}

// --- Média por semana GE ---

function atualizarMediaSemanaGE_() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    var audCols = getAudColsGE_(sheet);
    if (audCols.length === 0) return;

    var programas = [];
    for (var i = 0; i < audCols.length; i++) {
        var col = audCols[i];
        var dateVal = sheet.getRange(1, col).getValue();
        var d = dateVal instanceof Date ? dateVal : parseDateStr_(dateVal);
        if (d) {
            programas.push({ col: col, date: d, dow: d.getDay() });
        }
    }

    var sextas = programas.filter(function(p) { return p.dow === 5; });
    for (var s = 0; s < sextas.length; s++) {
        var sexta = sextas[s];
        var sextaDate = sexta.date;
        var segDate = new Date(sextaDate);
        segDate.setDate(segDate.getDate() - 4);
        segDate.setHours(0, 0, 0, 0);
        var sexEndDate = new Date(sextaDate);
        sexEndDate.setHours(23, 59, 59, 999);

        var weekCols = programas.filter(function(p) {
            var d = new Date(p.date); d.setHours(0, 0, 0, 0);
            return d >= segDate && d <= sexEndDate;
        }).map(function(p) { return p.col; });
        if (weekCols.length === 0) continue;

        var avgParts = weekCols.map(function(c) { return columnToLetter_(c) + "63"; }).join(";");
        sheet.getRange(66, sexta.col).setValue("=MÉDIA(" + avgParts + ")");
        sheet.getRange(66, sexta.col).setNumberFormat("0.00").setFontWeight("bold");

        var sextaAnterior = null;
        for (var j = s + 1; j < sextas.length; j++) { sextaAnterior = sextas[j]; break; }

        if (sextaAnterior) {
            var colAtual = columnToLetter_(sexta.col);
            var colAnterior = columnToLetter_(sextaAnterior.col);
            sheet.getRange(67, sexta.col).setValue(
                '=SE(OU(' + colAnterior + '66="";' + colAnterior + '66=0);"";' +
                colAtual + '66/' + colAnterior + '66-1)');
            sheet.getRange(67, sexta.col).setNumberFormat("+0.0%;-0.0%").setFontWeight("bold");
        } else {
            sheet.getRange(67, sexta.col).setValue("—");
        }
    }
}

// --- Gráfico de médias GE ---

function atualizarGraficoMediasGE_() {
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;

    // FIX: usar detecção baseada em data (exclui colunas CALIBRAGEM)
    var audCols = getAudColsGE_(sheet);
    var numPrograms = audCols.length;

    var charts = sheet.getCharts();
    for (var i = 0; i < charts.length; i++) {
        var info = charts[i].getContainerInfo();
        if (info && info.getAnchorRow() >= 1 && info.getAnchorColumn() >= 8) {
            sheet.removeChart(charts[i]);
        }
    }
    if (numPrograms === 0) return;

    // Série 3: média da semana vigente (Seg→hoje)
    var hoje = new Date();
    var dow = hoje.getDay(); // 0=Dom, 1=Seg, ..., 6=Sab
    var segundaDate = new Date(hoje);
    segundaDate.setDate(hoje.getDate() - (dow === 0 ? 6 : dow - 1));
    segundaDate.setHours(0, 0, 0, 0);
    var hojeEnd = new Date(hoje);
    hojeEnd.setHours(23, 59, 59, 999);

    var colsSemana = audCols.filter(function(c) {
        var val = sheet.getRange(1, c).getValue();
        var d = val instanceof Date ? val : parseDateStr_(val);
        if (!d) return false;
        var dn = new Date(d.getFullYear(), d.getMonth(), d.getDate());
        return dn >= segundaDate && dn <= hojeEnd;
    });

    var numSeries = colsSemana.length > 0 ? 4 : 3;
    var DATA_START_ROW = 70;
    sheet.getRange(DATA_START_ROW, 1, 62, 4).clearContent();

    var headerRow = [["HORA", "MÉDIA HISTÓRICA", "ÚLTIMO PROGRAMA"]];
    if (colsSemana.length > 0) headerRow[0].push("MÉDIA SEMANA");
    sheet.getRange(DATA_START_ROW, 1, 1, numSeries).setValues(headerRow);

    var graphFormulas = [];
    for (var row = 3; row <= 62; row++) {
        var rowF = ["=A" + row, "=B" + row, "=D" + row];
        if (colsSemana.length > 0) {
            var parts = colsSemana.map(function(c) { return columnToLetter_(c) + row; }).join(";");
            rowF.push("=MÉDIA(" + parts + ")");
        }
        graphFormulas.push(rowF);
    }
    batchSetFormulasSafe_(sheet.getRange(DATA_START_ROW + 1, 1, 60, numSeries), graphFormulas);
    sheet.getRange(DATA_START_ROW + 1, 2, 60, numSeries - 1).setNumberFormat("0.00");

    SpreadsheetApp.flush();
    Utilities.sleep(500);

    // FIX: data salva como string "dd/MM/yyyy" — usar fallback de string
    var lastDate = sheet.getRange(1, 4).getValue();
    var dateStr = "";
    if (lastDate instanceof Date) {
        dateStr = Utilities.formatDate(lastDate, "America/Sao_Paulo", "dd/MM/yyyy");
    } else if (typeof lastDate === "string" && lastDate.trim()) {
        dateStr = lastDate.trim();
    }

    // FIX: cores corretas — vermelho=histórica, azul=último
    var seriesOptions = {
        0: { color: "#E74C3C", lineWidth: 3 },  // MÉDIA HISTÓRICA = vermelho
        1: { color: "#2980B9", lineWidth: 2 }   // ÚLTIMO PROGRAMA = azul
    };
    if (colsSemana.length > 0) {
        seriesOptions[2] = { color: "#27AE60", lineWidth: 2, lineDashStyle: "SHORT_DASH" };
    }

    var chartBuilder = sheet.newChart()
        .asLineChart()
        .addRange(sheet.getRange(DATA_START_ROW, 1, 61, 1))
        .addRange(sheet.getRange(DATA_START_ROW, 2, 61, 1))
        .addRange(sheet.getRange(DATA_START_ROW, 3, 61, 1));
    if (colsSemana.length > 0) {
        chartBuilder.addRange(sheet.getRange(DATA_START_ROW, 4, 61, 1));
    }
    var chart = chartBuilder
        .setPosition(2, 9, 0, 0)
        .setOption("title", "Média Histórica vs Último" + (dateStr ? " (" + dateStr + ")" : ""))
        .setOption("titleTextStyle", { fontSize: 12, bold: true })
        .setOption("legend", { position: "bottom" })
        .setOption("hAxis", { title: "Horário", slantedText: true, slantedTextAngle: 45 })
        .setOption("vAxis", { title: "Audiência" })
        .setOption("series", seriesOptions)
        .setOption("width", 800)
        .setOption("height", 400)
        .build();
    sheet.insertChart(chart);
}

// --- Comentaristas fixos da casa (edite para adicionar/remover) ---
var COMENTARISTAS_DA_CASA_GE = ["ALEX", "GARRAFFA", "MULLER", "PAULO SÉRGIO", "SALAZAR", "SILVESTRE"];

// --- Média por comentarista GE ---

function atualizarMediaPorComentaristaGE_() {
    var ss = SpreadsheetApp.getActive();
    var hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!hist) return;
    var tabName = "GE - Média por Comentarista";
    var dest = ss.getSheetByName(tabName);
    if (dest) { dest.clear(); } else { dest = ss.insertSheet(tabName); }

    var audColsList = getAudColsGE_(hist);
    var numPrograms = audColsList.length;

    // === Leitura batch (1 chamada de API) ===
    var programData = [];
    var globalSum = 0, globalCount = 0;
    var dayTotals = {};

    if (numPrograms > 0) {
        var maxAudCol = audColsList[audColsList.length - 1];
        var allData = hist.getRange(1, 1, 65, maxAudCol + 3).getValues();

        for (var i = 0; i < audColsList.length; i++) {
            var audCol = audColsList[i];
            var ci = audCol - 1; // índice 0-based

            // Data (row 1)
            var dateVal = allData[0][ci];
            var dDate = dateVal instanceof Date ? dateVal
                : parseDateStr_(typeof dateVal === "string" ? dateVal.trim() : "");

            // Auto-calcular média das rows 3-62 (índices 2-61) — não depende da fórmula da row 63
            var sum = 0, cnt = 0;
            for (var ri = 2; ri <= 61; ri++) {
                var v = allData[ri][ci];
                if (typeof v === "number" && v > 0) { sum += v; cnt++; }
            }
            if (cnt === 0) continue;
            var avg = sum / cnt;

            var dow = dDate ? dDate.getDay() : -1;
            var dayName = (dow >= 1 && dow <= 5)
                ? ["", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA"][dow] : "";

            // Comentaristas (rows 64-65, índices 63-64)
            var names = [];
            for (var ri2 = 63; ri2 <= 64; ri2++) {
                var raw = allData[ri2] ? allData[ri2][ci] : "";
                if (!raw || typeof raw === "number" || raw instanceof Date) continue;
                var nm = raw.toString().trim().toUpperCase();
                if (nm && !/^\d+([.,]\d+)?$/.test(nm)) names.push(nm);
            }

            programData.push({ avg: avg, dDate: dDate, day: dayName, names: names });
            globalSum += avg; globalCount++;

            if (dayName) {
                if (!dayTotals[dayName]) dayTotals[dayName] = { sum: 0, count: 0 };
                dayTotals[dayName].sum += avg;
                dayTotals[dayName].count++;
            }
        }
    }

    var globalAvg = globalCount > 0 ? globalSum / globalCount : 0;
    var dayMeans = {};
    for (var d in dayTotals) {
        dayMeans[d] = dayTotals[d].count > 0 ? dayTotals[d].sum / dayTotals[d].count : globalAvg;
    }

    // Agrupar aparições por comentarista
    var coment = {};
    for (var i = 0; i < programData.length; i++) {
        var prog = programData[i];
        var seen = {};
        for (var n = 0; n < prog.names.length; n++) {
            var nome = prog.names[n];
            if (seen[nome]) continue;
            seen[nome] = true;
            if (!coment[nome]) coment[nome] = [];
            coment[nome].push({ avg: prog.avg, day: prog.day, dDate: prog.dDate });
        }
    }

    // Calcular estatísticas de um comentarista
    function calcStats(arr) {
        arr.sort(function(a, b) {
            return (a.dDate ? a.dDate.getTime() : 0) - (b.dDate ? b.dDate.getTime() : 0);
        });
        var avgs = arr.map(function(x) { return x.avg; });
        var mean = avgs.reduce(function(s, v) { return s + v; }, 0) / avgs.length;
        var l5 = avgs.slice(-5), l3 = avgs.slice(-3);
        var m5 = l5.reduce(function(s, v) { return s + v; }, 0) / l5.length;
        var m3 = l3.reduce(function(s, v) { return s + v; }, 0) / l3.length;
        // Média ponderada: corrige viés de dia da semana
        // adjusted = avg * (globalAvg / dayMean) → normaliza cada programa pela força do seu dia
        var adj = arr.map(function(x) {
            var dm = dayMeans[x.day] || globalAvg;
            return dm > 0 ? x.avg * (globalAvg / dm) : x.avg;
        });
        var wAvg = adj.reduce(function(s, v) { return s + v; }, 0) / adj.length;
        return [arr.length, mean, m5, m3, wAvg];
    }

    var CASA = COMENTARISTAS_DA_CASA_GE;
    var convidados = Object.keys(coment).filter(function(n) {
        return CASA.indexOf(n) === -1;
    }).sort();

    var header = ["COMENTARISTA", "QTD", "MÉDIA GERAL", "ÚLTIMAS 5", "ÚLTIMAS 3", "MÉD. PONDERADA"];
    var nCols = header.length;
    var currentRow = 1;

    // Escrever seção (DA CASA ou CONVIDADOS)
    function writeSection(title, bgTitle, bgHeader, members, altBg) {
        // Título da seção
        dest.getRange(currentRow, 1).setValue(title);
        dest.getRange(currentRow, 1, 1, nCols)
            .setBackground(bgTitle).setFontColor("white").setFontWeight("bold");
        currentRow++;

        // Cabeçalho de colunas
        dest.getRange(currentRow, 1, 1, nCols).setValues([header]);
        dest.getRange(currentRow, 1, 1, nCols)
            .setBackground(bgHeader).setFontColor("white").setFontWeight("bold")
            .setHorizontalAlignment("center");
        currentRow++;

        if (members.length === 0) {
            dest.getRange(currentRow, 1).setValue("(nenhum)");
            currentRow++;
            return;
        }

        var startRow = currentRow;
        var rows = [];
        for (var k = 0; k < members.length; k++) {
            var nomeK = members[k];
            var arr = coment[nomeK];
            if (arr && arr.length > 0) {
                var s = calcStats(arr);
                rows.push([nomeK, s[0], s[1], s[2], s[3], s[4]]);
            } else {
                rows.push([nomeK, 0, "", "", "", ""]);
            }
        }

        dest.getRange(startRow, 1, rows.length, nCols).setValues(rows);

        // Cores alternadas
        for (var r = 0; r < rows.length; r++) {
            dest.getRange(startRow + r, 1, 1, nCols).setBackground(r % 2 === 0 ? altBg : "#FFFFFF");
        }

        // Formatos numéricos
        dest.getRange(startRow, 2, rows.length, 1).setNumberFormat("0").setHorizontalAlignment("center");
        dest.getRange(startRow, 3, rows.length, 4).setNumberFormat("0.00").setHorizontalAlignment("center");

        currentRow += rows.length;
    }

    writeSection("⭐ COMENTARISTAS DA CASA", "#922B21", "#E74C3C", CASA, "#FDF2F2");
    currentRow++; // linha vazia entre seções
    writeSection("🎙️ CONVIDADOS", "#1A5276", "#2980B9", convidados, "#EBF5FB");

    // Larguras das colunas
    dest.setColumnWidth(1, 185);
    dest.setColumnWidth(2, 55);
    dest.setColumnWidth(3, 115);
    dest.setColumnWidth(4, 100);
    dest.setColumnWidth(5, 100);
    dest.setColumnWidth(6, 135);

    SpreadsheetApp.flush();
}

// --- Média por dia da semana GE ---

function atualizarMediaPorDiaSemanaGE_() {
    var ss = SpreadsheetApp.getActive();
    var hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!hist) return;
    var tabName = "GE - Média por Dia da Semana";
    var dest = ss.getSheetByName(tabName);
    if (dest) { dest.clear(); } else { dest = ss.insertSheet(tabName); }

    var dias = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA"];
    var histName = CFG.GE_OUT_HISTORICO;

    // Agrupar colunas de audiência por dia da semana (newest first = audCols[0] = col 4)
    var audColsDia = getAudColsGE_(hist);
    var map = { "SEGUNDA": [], "TERÇA": [], "QUARTA": [], "QUINTA": [], "SEXTA": [] };

    // Batch-read: lê todas as datas da row 1 em uma única chamada de API
    if (audColsDia.length > 0) {
        var firstAudCol = audColsDia[0];
        var lastAudCol = audColsDia[audColsDia.length - 1];
        var row1Vals = hist.getRange(1, firstAudCol, 1, lastAudCol - firstAudCol + 1).getValues()[0];
        for (var i = 0; i < audColsDia.length; i++) {
            var col = audColsDia[i];
            var dateVal = row1Vals[col - firstAudCol];
            var dDate = dateVal instanceof Date ? dateVal : parseDateStr_(dateVal);
            if (!dDate) continue;
            var dow = dDate.getDay();
            var nomeDia = ["DOMINGO", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"][dow];
            if (map[nomeDia] !== undefined) map[nomeDia].push(col);
        }
    }

    // Helper declarado fora do loop para evitar function declaration dentro de bloco (GAS V8)
    function makeAvgFormula(colList) {
        var refs = colList.map(function(c) {
            return "'" + histName + "'!" + columnToLetter_(c) + "63";
        }).join(";");
        return '=SEERRO(MÉDIA(' + refs + ');0)';
    }

    var header = ["DIA", "MÉDIA HISTÓRICA", "ÚLTIMAS 5", "ÚLTIMAS 3"];
    dest.getRange(1, 1, 1, header.length).setValues([header]);
    dest.getRange(1, 1, 1, header.length)
        .setFontWeight("bold").setBackground("#E74C3C").setFontColor("white").setHorizontalAlignment("center");

    for (var d = 0; d < dias.length; d++) {
        var destRow = d + 2;
        var cols = map[dias[d]];
        dest.getRange(destRow, 1).setValue(dias[d]);

        if (cols.length === 0) {
            dest.getRange(destRow, 2, 1, 3).setValues([["—", "—", "—"]]);
            continue;
        }

        dest.getRange(destRow, 2).setValue(makeAvgFormula(cols));
        dest.getRange(destRow, 3).setValue(makeAvgFormula(cols.slice(0, Math.min(5, cols.length))));
        dest.getRange(destRow, 4).setValue(makeAvgFormula(cols.slice(0, Math.min(3, cols.length))));
    }

    // Aplicar formatação correta (0.00 = padrão Apps Script; exibe 0,45 no locale BR)
    dest.getRange(2, 2, dias.length, 3).setNumberFormat("0.00");
    dest.setFrozenRows(1);
    dest.setColumnWidth(1, 140); dest.setColumnWidth(2, 140);
    dest.setColumnWidth(3, 140); dest.setColumnWidth(4, 140);
}

// --- Atualizar todos os dashboards GE (entrada pública via menu) ---

function atualizarDashboardsGE() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (sheet) {
        atualizarCalibragemGE_(sheet);
        atualizarFatorTemaGE_(sheet);
    }
    garantirCamposHistoricoGE_();
    atualizarMediaSemanaGE_();
    atualizarGraficoMediasGE_();
    atualizarMediaPorComentaristaGE_();
    atualizarMediaPorDiaSemanaGE_();
    atualizarMediaPorTemaSeExistir_();
}

// --- Helpers de fórmulas para a aba Média por Tema GE ---

function gerarFormulaCountTema_(e, t, o, a) {
    const r = t.replace(/"/g, '""'),
        n = [];
    for (let t = 0; t < a; t++) {
        const o = columnToLetter_(6 + 4 * t);
        n.push(`CONT.SE('${e}'!${o}3:${o}62;"${r}")`)
    }
    return "=" + n.join("+")
}

function gerarFormulaMediaTema_(e, t, o, a) {
    const r = t.replace(/"/g, '""'),
        n = [],
        s = [];
    for (let t = 0; t < a; t++) {
        const o = columnToLetter_(6 + 4 * t),
            a = columnToLetter_(5 + 4 * t);
        n.push(`SOMASE('${e}'!${o}3:${o}62;"${r}";'${e}'!${a}3:${a}62)`), s.push(`CONT.SE('${e}'!${o}3:${o}62;"${r}")`)
    }
    const i = n.join("+"),
        c = s.join("+");
    return `=SE((${c})=0;0;(${i})/(${c})*'${e}'!$B$63)`
}

// --- Criar aba Média por Tema GE (entrada pública via menu) ---

function criarMediaPorTemaGE() {
    normalizarTemasNoHistoricoGE_();
    const e = SpreadsheetApp.getUi(),
        t = SpreadsheetApp.getActive(),
        o = t.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!o) return void e.alert("\u274c Erro", 'A aba "GE - Hist\xf3rico" n\xe3o existe.', e.ButtonSet.OK);
    const r = getAudColsGE_(o).length;
    if (0 === r) return void e.alert("\u274c Erro", "N\xe3o h\xe1 programas arquivados no hist\xf3rico.", e.ButtonSet.OK);
    const n = coletarTemasUnicosGE_(o, r);
    if (0 === n.length) return void e.alert("\u274c Erro", "N\xe3o foram encontrados TEMAS no hist\xf3rico.", e.ButtonSet.OK);
    const s = "GE - M\xe9dia por Minuto TEMA";
    let i = t.getSheetByName(s);
    i ? i.clear() : i = t.insertSheet(s), configurarAbaMediaPorTema_(i, o, n, r), e.alert("\u2705 Sucesso!", `Aba "${s}" criada/atualizada com sucesso!\n\n\u{1f4ca} ${n.length} TEMAS encontrados\n\u{1f4c5} ${r} programas no hist\xf3rico\n\nAs m\xe9dias s\xe3o calculadas automaticamente via f\xf3rmulas.`, e.ButtonSet.OK)
}

// --- Atualizar fórmulas da aba Média por Tema GE (entrada pública via menu) ---

function atualizarFormulasTemaGE() {
    normalizarTemasNoHistoricoGE_();
    const l = LockService.getDocumentLock();
    l.waitLock(30000);
    try {
        const e = SpreadsheetApp.getUi(),
            t = SpreadsheetApp.getActive(),
            o = "GE - M\xe9dia por Minuto TEMA",
            a = t.getSheetByName(o);
        if (!a) return void e.alert("\u274c Erro", `A aba "${o}" n\xe3o existe.\n\nCrie a aba primeiro usando o menu:\nGazeta Esportiva \u2192 An\xe1lise por Tema \u2192 Criar/Atualizar M\xe9dia por Tema`, e.ButtonSet.OK);
        const r = t.getSheetByName(CFG.GE_OUT_HISTORICO);
        if (!r) return void e.alert("\u274c Erro", 'A aba "GE - Hist\xf3rico" n\xe3o existe.', e.ButtonSet.OK);
        const s = getAudColsGE_(r).length;
        if (0 === s) return void e.alert("\u274c Erro", "N\xe3o h\xe1 programas arquivados no hist\xf3rico.", e.ButtonSet.OK);
        const i = coletarTemasUnicosGE_(r, s);
        0 !== i.length ? (a.clear(), configurarAbaMediaPorTema_(a, r, i, s), e.alert("\u2705 F\xf3rmulas Atualizadas!", `As f\xf3rmulas da aba "${o}" foram atualizadas com sucesso!\n\n\u{1f4ca} ${i.length} TEMAS\n\u{1f4c5} ${s} programas`, e.ButtonSet.OK)) : e.alert("\u274c Erro", "N\xe3o foram encontrados TEMAS no hist\xf3rico.", e.ButtonSet.OK)
    } finally {
        l.releaseLock()
    }
}

// --- Normalizar temas no histórico GE ---

function normalizarTemasNoHistoricoGE_() {
    const e = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!e) return;
    const audColsNorm = getAudColsGE_(e);
    for (let a = 0; a < audColsNorm.length; a++) {
        const o = 6 + 4 * a;  // coluna TEMA do programa a (col 4=AUD, 5=SHA, 6=TEMA)
        const tRange = e.getRange(3, o, 60, 1),
            vals = tRange.getValues();
        let changed = false;
        for (let i = 0; i < vals.length; i++) {
            const v = vals[i][0];
            if (!v) continue;
            const norm = normalizarTema_(v);
            if (norm !== v) { vals[i][0] = norm; changed = true; }
        }
        if (changed) tRange.setValues(vals);
    }
}

// --- Coletar temas únicos do histórico GE ---

function coletarTemasUnicosGE_(e, t) {
    const o = new Set;
    for (let a = 0; a < t; a++) {
        const t = 6 + 4 * a,
            r = e.getRange(3, t, 60, 1).getValues();
        for (const e of r) {
            const t = e[0];
            t && typeof t === "string" && /[a-zA-ZÀ-ÿ]{2,}/.test(t.trim()) && o.add(t.trim())
        }
    }
    return Array.from(o).sort()
}

// --- Configurar aba Média por Tema GE ---

function configurarAbaMediaPorTema_(e, t, o, a) {
    const r = CFG.GE_OUT_HISTORICO,
        n = ["TEMA", "Minutos Total", "M\xe9dia &/MIN (Todos)", "M\xe9dia &/MIN (\xdaltimos 5)", "M\xe9dia &/MIN (\xdaltimos 3)", "\xcdndice vs M\xe9dia Geral"];
    e.getRange(1, 1, 1, n.length).setValues([n]), e.getRange(1, 1, 1, n.length).setFontWeight("bold").setBackground("#E74C3C").setFontColor("white").setHorizontalAlignment("center");
    const s = Math.min(5, a),
        i = Math.min(3, a),
        c = [];
    for (let e = 0; e < o.length; e++) c.push([o[e]]);
    c.length > 0 && e.getRange(2, 1, c.length, 1).setValues(c);
    const l2 = o.length,
        d = new Array(l2);
    for (let t = 0; t < l2; t++) {
        const n = o[t],
            c = t + 2,
            m = gerarFormulaCountTema_(r, n, a, a),
            p = gerarFormulaMediaTema_(r, n, a, a),
            u = gerarFormulaMediaTema_(r, n, a, s),
            g = gerarFormulaMediaTema_(r, n, a, i);
        const b = sanitizeFormula_(m),
            h = sanitizeFormula_(p),
            v = sanitizeFormula_(u),
            y = sanitizeFormula_(g);
        d[t] = [(b && "=" !== b[0] ? "=" + b : b), (h && "=" !== h[0] ? "=" + h : h), (v && "=" !== v[0] ? "=" + v : v), (y && "=" !== y[0] ? "=" + y : y), sanitizeFormula_("=SE(C" + c + "=\"\";\"\";C" + c + ")")]
    }
    const A = 50;
    for (let t = 0; t < d.length; t += A) {
        const n = d.slice(t, t + A),
            s = e.getRange(t + 2, 2, n.length, 5);
        let i = 0;
        for (;;) {
            try {
                setFormulasUserEntered_(s, n);
                break
            } catch (e) {
                if (++i >= 5) throw e;
                Utilities.sleep(200 * Math.pow(2, i - 1))
            }
        }
    }
    SpreadsheetApp.flush();

    e.getRange(2, 3, o.length, 3).setNumberFormat("0.00");

    const l = o.length;
    e.setColumnWidth(1, 150), e.setColumnWidth(2, 100), e.setColumnWidth(3, 130), e.setColumnWidth(4, 130), e.setColumnWidth(5, 130), e.setColumnWidth(6, 140), l > 0 && (e.getRange(2, 2, l, 1).setNumberFormat("0"), e.getRange(2, 3, l, 3).setNumberFormat("0.00"), e.getRange(2, 6, l, 1).setNumberFormat("0.00")), e.getRange(2, 2, l, 5).setHorizontalAlignment("center");
    for (let t = 0; t < l; t++) {
        const o = t + 2,
            a = t % 2 == 0 ? "#FDF2F2" : "#FFFFFF";
        e.getRange(o, 1, 1, 6).setBackground(a)
    }
    e.setFrozenRows(1);
    const totalRow = l + 3;
    e.getRange(totalRow, 1).setValue("TOTAL/M\xc9DIA GERAL"), e.getRange(totalRow, 1).setFontWeight("bold"), setFormulaSafe_(e.getRange(totalRow, 2), `=SOMA(B2:B${l+1})`), setFormulaSafe_(e.getRange(totalRow, 3), `='${r}'!B63`), e.getRange(totalRow, 1, 1, 6).setBackground("#E74C3C").setFontColor("white").setFontWeight("bold"), e.getRange(totalRow, 2).setNumberFormat("0"), e.getRange(totalRow, 3).setNumberFormat("0.00");
    const u = totalRow + 2;
    e.getRange(u, 1).setValue("\u{1f4ca} Informa\xe7\xf5es:"), e.getRange(u + 1, 1).setValue("\xdaltima atualiza\xe7\xe3o: " + Utilities.formatDate(new Date, "America/Sao_Paulo", "dd/MM/yyyy HH:mm")), e.getRange(u + 2, 1).setValue("Total de programas: " + a), e.getRange(u + 3, 1).setValue('Programas nos "\xdaltimos 5": ' + s), e.getRange(u + 4, 1).setValue('Programas nos "\xdaltimos 3": ' + i), e.getRange(u, 1, 5, 1).setFontStyle("italic").setFontColor("#666666"), SpreadsheetApp.flush()
}

// --- Modelo de predição: CALIBRAGEM GS–GY ---

function atualizarCalibragemGE_(sheet) {
    var audCols = getAudColsGE_(sheet);
    var numPrograms = audCols.length;
    if (numPrograms === 0) return;

    var calibStartCol = 4 + numPrograms * 4;
    var GSl = columnToLetter_(calibStartCol);
    var GTl = columnToLetter_(calibStartCol + 1);
    var GUl = columnToLetter_(calibStartCol + 2);
    var GVl = columnToLetter_(calibStartCol + 3);
    var GWl = columnToLetter_(calibStartCol + 4);
    var GXl = columnToLetter_(calibStartCol + 5);
    var GYl = columnToLetter_(calibStartCol + 6);

    // A1 e B1
    setFormulaSafe_(sheet.getRange(1, 1), "=D63");
    sheet.getRange(1, 1).setNumberFormat("0.00");
    sheet.getRange(1, 2).setValue("PREV");

    // Cabeçalhos row 1
    sheet.getRange(1, calibStartCol, 1, 7)
        .setValues([["BASE HIST.", "RAZÃO D/HIST", "FATOR DIA", "FATOR RECENTE", "INCLINAÇÃO", "RUPTURA", "FATOR FINAL"]])
        .setFontWeight("bold").setBackground("#2C3E50").setFontColor("white").setHorizontalAlignment("center");
    // Cabeçalhos row 2
    sheet.getRange(2, calibStartCol, 1, 7)
        .setValues([["GS", "GT", "GU", "GV", "GW", "GX", "GY"]])
        .setFontWeight("bold").setBackground("#34495E").setFontColor("white").setHorizontalAlignment("center");

    // Colunas históricas (todas exceto D=col4)
    var audColsHist = audCols.filter(function(c) { return c !== 4; });

    // C1 = previsão dinâmica: parcial_atual × (histórico_total / histórico_parcial_21min)
    setFormulaSafe_(sheet.getRange(1, 3), '=MÉDIA(D3:D23)*MÉDIA(SEERRO(FILTER(H63:LH63;MOD(COL(H63:LH63)-COL(H63);4)=0);""))/MÉDIA(SEERRO(FILTER(H3:LH23;MOD(COL(H3:LH23)-COL(H3);4)=0);""))');
    sheet.getRange(1, 3).setNumberFormat("0.00");

    var gsF = [], gtF = [], guF = [], gvF = [], gwF = [], gxF = [], gyF = [];

    for (var r = 3; r <= 62; r++) {
        // GS: média histórica do minuto (exclui programa mais recente)
        if (audColsHist.length === 0) {
            gsF.push(["1"]);
        } else {
            var parts = audColsHist.map(function(c) { return columnToLetter_(c) + r; }).join(";");
            gsF.push(["=MÉDIA(" + parts + ")"]);
        }

        // GT: razão real vs histórico
        gtF.push(['=SE(OU(D' + r + '=""' + ';' + GSl + r + '=0);"";D' + r + '/' + GSl + r + ')']);

        // GU: fator geral do dia (média acumulada de GT)
        if (r === 3) {
            guF.push(["1"]);
        } else {
            guF.push(['=SEERRO(MÉDIA(' + GTl + '$3:ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1));1)']);
        }

        // GV: fator recente ponderado (EMA 3 minutos)
        if (r <= 5) {
            gvF.push(['=' + GUl + r]);
        } else {
            gvF.push(['=SE(LIN()<6;' + GUl + r + ';0,5*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)+0,3*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-2)+0,2*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-3))']);
        }

        // GW: inclinação (delta GT)
        if (r <= 5) {
            gwF.push(["0"]);
        } else {
            gwF.push(['=SE(LIN()<6;0;ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)-ÍNDICE(' + GTl + ':' + GTl + ';LIN()-3))']);
        }

        // GX: detector de ruptura de regime
        if (r <= 5) {
            gxF.push(["0"]);
        } else {
            gxF.push(['=SE(LIN()<6;0;SE(E(MÉDIA(ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)-1;ÍNDICE(' + GTl + ':' + GTl + ';LIN()-2)-1;ÍNDICE(' + GTl + ':' + GTl + ';LIN()-3)-1)<=-0,08;CONT.SE(ÍNDICE(' + GTl + ':' + GTl + ';LIN()-3):ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1);"<=0,90")>=2);1;0))']);
        }

        // GY: fator final de regime (com/sem ruptura)
        if (r <= 5) {
            gyF.push(['=MÍNIMO(1,4;MÁXIMO(0,3;' + GUl + r + '))']);
        } else {
            gyF.push(['=SE(' + GXl + r + '=1;MÍNIMO(1,4;MÁXIMO(0,3;0,65*' + GVl + r + '+0,25*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)+0,10*(' + GVl + r + '+' + GWl + r + ')));MÍNIMO(1,4;MÁXIMO(0,3;0,20*' + GUl + r + '+0,45*' + GVl + r + '+0,20*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)+0,15*(' + GVl + r + '+' + GWl + r + '))))']);
        }
    }

    batchSetFormulasSafe_(sheet.getRange(3, calibStartCol,     60, 1), gsF);
    batchSetFormulasSafe_(sheet.getRange(3, calibStartCol + 1, 60, 1), gtF);
    batchSetFormulasSafe_(sheet.getRange(3, calibStartCol + 2, 60, 1), guF);
    batchSetFormulasSafe_(sheet.getRange(3, calibStartCol + 3, 60, 1), gvF);
    batchSetFormulasSafe_(sheet.getRange(3, calibStartCol + 4, 60, 1), gwF);
    batchSetFormulasSafe_(sheet.getRange(3, calibStartCol + 5, 60, 1), gxF);
    batchSetFormulasSafe_(sheet.getRange(3, calibStartCol + 6, 60, 1), gyF);

    sheet.getRange(3, calibStartCol,     60, 7).setNumberFormat("0.000");
    sheet.getRange(3, calibStartCol + 5, 60, 1).setNumberFormat("0"); // GX binário

    [80, 70, 80, 80, 70, 70, 80].forEach(function(w, i) {
        sheet.setColumnWidth(calibStartCol + i, w);
    });

    atualizarRankingTemaCalibGE_();
    atualizarDuracaoBlocosCalibGE_();
    atualizarAlertaColunasGE_(sheet);

    SpreadsheetApp.flush();
}

// --- Modelo de predição: Fator TEMA GZ–HC ---

function atualizarFatorTemaGE_(sheet) {
    var ss = SpreadsheetApp.getActive();
    var temaSheet = ss.getSheetByName("GE - Média por Minuto TEMA");
    if (!temaSheet) return;

    var audCols = getAudColsGE_(sheet);
    var numPrograms = audCols.length;
    if (numPrograms === 0) return;

    var temaStartCol = 4 + numPrograms * 4 + 7; // logo após GS–GY
    var GZl = columnToLetter_(temaStartCol);
    var HAl = columnToLetter_(temaStartCol + 1);
    var HBl = columnToLetter_(temaStartCol + 2);
    // HC = temaStartCol + 3

    // Cabeçalhos row 1
    sheet.getRange(1, temaStartCol, 1, 4)
        .setValues([["PESO TEMA", "TROCA TEMA", "DUR. BLOCO", "FATOR TEMA"]])
        .setFontWeight("bold").setBackground("#1A5276").setFontColor("white").setHorizontalAlignment("center");
    sheet.getRange(2, temaStartCol, 1, 4)
        .setValues([["GZ", "HA", "HB", "HC"]])
        .setFontWeight("bold").setBackground("#21618C").setFontColor("white").setHorizontalAlignment("center");

    var gzF = [], haF = [], hbF = [], hcF = [];
    for (var r = 3; r <= 62; r++) {
        // GZ: peso do tema via PROCV na aba de temas
        gzF.push(['=SEERRO(PROCV(F' + r + ";'GE - Média por Minuto TEMA'!$A:$F;6;0);1)"]);

        // HA: troca de tema em relação ao minuto anterior
        if (r === 3) {
            haF.push(["0"]);
        } else {
            haF.push(['=SE(F' + r + '=F' + (r - 1) + ';0;1)']);
        }

        // HB: duração consecutiva do bloco do tema
        if (r === 3) {
            hbF.push(["1"]);
        } else {
            hbF.push(['=SE(' + HAl + r + '=1;1;' + HBl + (r - 1) + '+1)']);
        }

        // HC: fator de persistência com ramp-up (70% → 85% → 100%)
        hcF.push(['=SE(F' + r + '="";1;SE(' + HBl + r + '<=1;MÁXIMO(0,92;' + GZl + r + '*0,92);SE(' + HBl + r + '<=3;MÁXIMO(0,96;' + GZl + r + '*0,96);' + GZl + r + ')))']);
    }

    batchSetFormulasSafe_(sheet.getRange(3, temaStartCol,     60, 1), gzF);
    batchSetFormulasSafe_(sheet.getRange(3, temaStartCol + 1, 60, 1), haF);
    batchSetFormulasSafe_(sheet.getRange(3, temaStartCol + 2, 60, 1), hbF);
    batchSetFormulasSafe_(sheet.getRange(3, temaStartCol + 3, 60, 1), hcF);

    sheet.getRange(3, temaStartCol,     60, 4).setNumberFormat("0.000");
    sheet.getRange(3, temaStartCol + 1, 60, 1).setNumberFormat("0"); // HA binário
    sheet.getRange(3, temaStartCol + 2, 60, 1).setNumberFormat("0"); // HB inteiro

    [90, 80, 80, 90].forEach(function(w, i) {
        sheet.setColumnWidth(temaStartCol + i, w);
    });

    SpreadsheetApp.flush();
}

// --- Atualizar calibragem manualmente (entrada pública via menu) ---

function atualizarCalibragemGE() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    atualizarCalibragemGE_(sheet);
    atualizarFatorTemaGE_(sheet);
    SpreadsheetApp.getUi().alert(
        "✅ Calibragem Atualizada!",
        "Colunas GS–GY (modelo de predição) regeneradas com sucesso.",
        SpreadsheetApp.getUi().ButtonSet.OK
    );
}

// --- Adicionar novo dia manual GE (entrada pública via menu) ---

function adicionarNovoDiaManualGE() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) {
        ui.alert("❌ Erro", 'A aba "GE - Histórico" não existe.', ui.ButtonSet.OK);
        return;
    }

    var resp = ui.prompt(
        "📅 Adicionar Novo Dia",
        "Digite a data do programa (dd/mm/aaaa):\n\nColunas serão inseridas com previsão de audiência.",
        ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    var dateStr = resp.getResponseText().trim();
    var match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!match) {
        ui.alert("❌ Formato inválido", "Use dd/mm/aaaa", ui.ButtonSet.OK);
        return;
    }
    var programDate = new Date(parseInt(match[3], 10), parseInt(match[2], 10) - 1, parseInt(match[1], 10));

    var lastCol = sheet.getLastColumn();
    var numPrograms = Math.floor((lastCol - 3) / 4);
    if (numPrograms > 0) {
        var existingDates = sheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
        for (var d = 0; d < existingDates.length; d += 4) {
            var existing = existingDates[d];
            if (!existing) continue;
            var existStr = existing instanceof Date
                ? Utilities.formatDate(existing, "America/Sao_Paulo", "dd/MM/yyyy")
                : existing.toString().trim();
            if (existStr === dateStr) {
                ui.alert("⚠️ Duplicada", dateStr + " já existe.", ui.ButtonSet.OK);
                return;
            }
        }
    }

    sheet.insertColumns(4, 4);

    var bgColor = (numPrograms % 2 === 0) ? "#FDF2F2" : "#FFFFFF";

    sheet.getRange(1, 4, 1, 4).merge();
    sheet.getRange(1, 4).setValue(dateStr)
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setBackground("#FADBD8")
        .setFontColor("#000000");

    sheet.getRange(2, 4, 1, 4).setValues([["AUDIÊNCIA", "&/MIN", "TEMA", "TERMOS"]]);
    sheet.getRange(2, 4, 1, 4)
        .setFontWeight("bold")
        .setBackground("#E74C3C")
        .setFontColor("white")
        .setHorizontalAlignment("center");

    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 70);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 200);

    sheet.getRange(3, 4, 60, 4).setBackground(bgColor);

    sheet.getRange(3, 4, 60, 1).setNumberFormat("0.00");
    sheet.getRange(3, 5, 60, 1).setNumberFormat("0.00");

    // Atualizar médias B/C + gerar CALIBRAGEM GS–GY (D ainda vazio → GT="" → GY≈1)
    atualizarFormulasHistoricoGE_(sheet);
    Utilities.sleep(1500);

    // Preencher D com predição =ARRED(GS*GY*HC, 2) se houver histórico
    var audColsAfter = getAudColsGE_(sheet);
    if (audColsAfter.length > 1) {
        var calibStart = 4 + audColsAfter.length * 4;
        var GSl = columnToLetter_(calibStart);
        var GYl = columnToLetter_(calibStart + 6);
        var ss2 = SpreadsheetApp.getActive();
        var hasTema = !!ss2.getSheetByName("GE - Média por Minuto TEMA");
        var HCl = hasTema ? columnToLetter_(calibStart + 7 + 3) : null;

        var predFormulas = [];
        for (var row = 3; row <= 62; row++) {
            var f = hasTema && HCl
                ? '=ARRED(' + GSl + row + '*' + GYl + row + '*' + HCl + row + ';2)'
                : '=ARRED(' + GSl + row + '*' + GYl + row + ';2)';
            predFormulas.push([f]);
        }
        batchSetFormulasSafe_(sheet.getRange(3, 4, 60, 1), predFormulas);
        Utilities.sleep(1000);
    }

    // Fórmula de share para col E (coluna &/MIN do novo dia)
    var shareFormulas = [];
    for (var row = 3; row <= 62; row++) {
        shareFormulas.push(['=SE($B' + row + '=0;"";D' + row + '/$B' + row + ')']);
    }
    batchSetFormulasSafe_(sheet.getRange(3, 5, 60, 1), shareFormulas);

    Utilities.sleep(500);

    // Bloco novo D63:G65 espelha o bloco anterior H63:K65
    sheet.getRange(63, 8, 3, 4).copyTo(sheet.getRange(63, 4, 3, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    // Textos fixos do bloco novo
    sheet.getRange(64, 4, 2, 1).setValues([["COMENTARISTA 1"], ["COMENTARISTA 2"]]);
    sheet.getRange(64, 6).setValue("DIFERENÇA");

    // Mediana por dia útil (terça=3ª, quarta=4ª, quinta=5ª, sexta=6ª)
    var weekday = programDate.getDay();
    var medianaLabel = {
        2: "MEDIANA 3ª",
        3: "MEDIANA 4ª",
        4: "MEDIANA 5ª",
        5: "MEDIANA 6ª"
    } [weekday] || "MEDIANA";
    sheet.getRange(64, 5).setValue(medianaLabel);

    var audColsForMedian = getAudColsGE_(sheet).filter(function(c) { return c !== 4; });
    var row1Values = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var medianRefs = [];
    for (var iMedian = 0; iMedian < audColsForMedian.length; iMedian++) {
        var cMedian = audColsForMedian[iMedian];
        var h = row1Values[cMedian - 1];
        var dMedian = h instanceof Date ? h : parseDateStr_(h);
        if (!dMedian) continue;
        if (dMedian.getDay() === weekday) medianRefs.push(columnToLetter_(cMedian) + "63");
    }
    if (medianRefs.length > 0) {
        setFormulaSafe_(sheet.getRange(63, 5), "=MED(" + medianRefs.join(";") + ")");
    } else {
        sheet.getRange(63, 5).setValue("");
    }
    setFormulaSafe_(sheet.getRange(63, 6), "=(D63/E63)-1");
    sheet.getRange(63, 6).setNumberFormat("+0.0%;-0.0%");

    // Herdar formatação condicional da coluna equivalente (&/MIN) do bloco anterior: I3:I62 -> E3:E62
    var rules = sheet.getConditionalFormatRules();
    var clonedRules = [];
    var sourceCol = 9, targetCol = 5, startRow = 3, numRows = 60;
    var shouldUpdateRules = false;
    for (var rIdx = 0; rIdx < rules.length; rIdx++) {
        var rule = rules[rIdx];
        var ranges = rule.getRanges();
        var hasSource = false, hasTarget = false;
        for (var rgIdx = 0; rgIdx < ranges.length; rgIdx++) {
            var rg = ranges[rgIdx];
            if (rg.getColumn() === sourceCol && rg.getRow() === startRow && rg.getNumColumns() === 1 && rg.getNumRows() === numRows) hasSource = true;
            if (rg.getColumn() === targetCol && rg.getRow() === startRow && rg.getNumColumns() === 1 && rg.getNumRows() === numRows) hasTarget = true;
        }
        if (hasSource && !hasTarget) {
            var newRanges = ranges.slice();
            newRanges.push(sheet.getRange(startRow, targetCol, numRows, 1));
            clonedRules.push(rule.copy().setRanges(newRanges).build());
            shouldUpdateRules = true;
        } else {
            clonedRules.push(rule);
        }
    }
    if (shouldUpdateRules) sheet.setConditionalFormatRules(clonedRules);

    garantirCamposHistoricoGE_();
    atualizarMediaSemanaGE_();
    Utilities.sleep(1000);
    atualizarGraficoMediasGE_();
    Utilities.sleep(1000);
    atualizarMediaPorComentaristaGE_();
    atualizarMediaPorDiaSemanaGE_();
    atualizarMediaPorTemaSeExistir_();

    SpreadsheetApp.flush();

    ui.alert("✅ Novo dia adicionado!",
        "Programa de " + dateStr + " inserido!\n\n" +
        "📊 Previsão GS×GY aplicada (linhas 3-62)\n" +
        "📈 Gráfico atualizado\n\n" +
        "⚠️ Preencha:\n" +
        "• Linha 64 (col D): Comentarista 1\n" +
        "• Linha 65 (col D): Comentarista 2",
        ui.ButtonSet.OK
    );
}

// --- Atualizar Média por Tema se aba existir ---

function atualizarMediaPorTemaSeExistir_() {
    var ss = SpreadsheetApp.getActive();
    var temaSheet = ss.getSheetByName("GE - Média por Minuto TEMA");
    if (!temaSheet) return;
    var histSheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!histSheet) return;
    var numPrograms = getAudColsGE_(histSheet).length;
    if (numPrograms === 0) return;
    normalizarTemasNoHistoricoGE_();
    var temas = coletarTemasUnicosGE_(histSheet, numPrograms);
    if (temas.length === 0) return;
    temaSheet.clear();
    configurarAbaMediaPorTema_(temaSheet, histSheet, temas, numPrograms);
}

// --- Ver estrutura do histórico GE (entrada pública via menu) ---

function verEstruturaHistoricoGazeta() {
    const e = SpreadsheetApp.getUi();
    e.alert("\u{1f4ca} Estrutura do Hist\xf3rico GE", "\n\u{1f4ca} ESTRUTURA DA PLANILHA GE - HIST\xd3RICO\n\n\u250c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510\n\u2502 Linha 1 \u2502     \u2502       \u2502        \u2502   17/01/2026    \u2502   16/01/2026    \u2502  ... \u2502\n\u251c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2524\n\u2502 Linha 2 \u2502HORA \u2502 M\xc9DIA \u2502%A M\xc9DIA\u2502 AUD\u2502SHA\u2502TEMA\u2502TER\u2502 AUD\u2502SHA\u2502TEMA\u2502TER\u2502  ... \u2502\n\u251c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2524\n\u2502 Linha 3 \u250218:00\u2502 0.45  \u2502 +2.3%  \u25020.46\u25021.2\u2502    \u2502   \u25020.44\u25021.1\u2502    \u2502   \u2502      \u2502\n\u2502 Linha 4 \u250218:01\u2502 0.47  \u2502 +1.4%  \u25020.48\u25021.3\u2502SPFC\u2502...\u25020.46\u25021.2\u2502PALM\u2502...\u2502      \u2502\n\u2502   ...   \u2502 ... \u2502  ...  \u2502  ...   \u2502 ...\u2502...\u2502 ...\u2502...\u2502 ...\u2502...\u2502 ...\u2502...\u2502      \u2502\n\u2502Linha 62 \u250218:59\u2502 0.55  \u2502 -0.5%  \u25020.54\u25021.5\u2502CORI\u2502...\u25020.56\u25021.6\u2502FLAM\u2502...\u2502      \u2502\n\u2514\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2518\n\nCOLUNAS FIXAS:\n\u2022 A: HORA (18:00 a 18:59)\n\u2022 B: M\xc9DIA de todas as audi\xeancias hist\xf3ricas daquele minuto\n\u2022 C: % A M\xc9DIA = (Audi\xeancia mais recente - M\xe9dia) / M\xe9dia\n\nBLOCOS DE 4 COLUNAS (por programa):\n\u2022 AUDI\xcaNCIA: valor em formato 0.00\n\u2022 SHARE: valor em formato 0.0\n\u2022 TEMA: tema classificado do minuto\n\u2022 TERMOS: top 5 termos do minuto\n\nCORES:\n\u2022 Cabe\xe7alhos: vermelho (#E74C3C) - tema do Gazeta\n\u2022 Programas: altern\xe2ncia vermelho claro / branco\n", e.ButtonSet.OK)
}
