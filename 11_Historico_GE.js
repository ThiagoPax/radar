// =============================================================================
// 11_Historico_GE.js — Histórico, dashboards e médias (Gazeta Esportiva)
// =============================================================================
//
// Mudanças nesta versão:
//
//   PREVISÃO INICIAL
//   • gerarPrevisaoInicialGE_ — previsão isolada a partir de histórico real,
//     sem a circularidade GS×GY×HC do modelo anterior (D vazio → GY≈1).
//   • atualizarFormulasHistoricoGE_ refatorada via atualizarFormulasBaseGE_.
//   • Ordem de execução em adicionarNovoDiaManualGE corrigida.
//
//   ÁREA PROTEGIDA A66:NT89 — três correções estruturais
//   • preencherBlocoA66E85HistoricoGE_ reescrita: agora dinâmica (baseada em
//     datas reais da row 1), sem fórmulas hardcoded, sem clearFormat/
//     clearDataValidations/clearNote destrutivos. Limpa F66:NT89 (polução
//     lateral) e preserva D86:D89 (valores externos do user).
//   • atualizarMediaSemanaGE_ não escreve mais em (66, sexta.col) e
//     (67, sexta.col) — média semanal vive integralmente em A66:E89.
//   • atualizarGraficoMediasGE_ usa DATA_START_ROW = 95 (fora da área
//     protegida); limpa preventivamente a área legada A70:D130.
//   • garantirCamposHistoricoGE_ garante mínimo de 160 rows para acomodar
//     dados auxiliares do gráfico em 95-155.
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

// =============================================================================
// MODELO DE PREVISÃO INICIAL — constantes e helpers
// =============================================================================
// Estas constantes governam o modelo isolado que gera a previsão de audiência
// do novo programa antes de existirem dados reais em D. O modelo é construído
// puramente a partir das colunas históricas (H em diante) e NÃO usa GS-GY,
// evitando a circularidade do modelo anterior.
// -----------------------------------------------------------------------------
var PREV_INICIAL_GE = {
    PESO_DIA: 0.40,            // peso do fator dia-da-semana no regime normal
    PESO_TREND: 0.35,          // peso do fator tendência (últimos 5)
    PESO_ACCEL: 0.25,          // peso do fator aceleração (últimos 3 / últimos 5)
    PESO_DIA_RUPT: 0.20,       // peso do fator dia em regime de ruptura
    PESO_TREND_RUPT: 0.55,     // peso do fator tendência em regime de ruptura
    PESO_ACCEL_RUPT: 0.25,     // peso do fator aceleração em regime de ruptura
    BLEND_RUPTURA: 0.50,       // mistura 50/50 regime normal × ruptura
    CLAMP_TREND: [0.75, 1.30], // limites do fator tendência
    CLAMP_ACCEL: [0.85, 1.18], // limites do fator aceleração
    CLAMP_DIA:   [0.78, 1.28], // limites do fator dia da semana
    MIN_AMOSTRAS_TREND: 3,
    MIN_AMOSTRAS_ACCEL: 2,
    MIN_AMOSTRAS_DIA:   2,
    MIN_AMOSTRAS_RUPTURA: 6,
    THRESHOLD_RUPTURA: 0.15,
    N_RECENT_5: 5,
    N_RECENT_3: 3
};

function clampGE_(v, lo, hi) {
    if (typeof v !== "number" || isNaN(v) || !isFinite(v)) return 1;
    return Math.max(lo, Math.min(hi, v));
}

function mediaArrGE_(arr) {
    if (!arr || arr.length === 0) return 0;
    var s = 0;
    for (var i = 0; i < arr.length; i++) s += arr[i];
    return s / arr.length;
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

// =============================================================================
// atualizarFormulasBaseGE_ — fórmulas-base SEM calibragem
// =============================================================================
// Atualiza B (média histórica), C (% à média), row 63 (médias do programa) e
// a coluna de share (&/MIN) — sem tocar nas auxiliares GS-GY ou GZ-HC.
// Permite que adicionarNovoDiaManualGE rode a previsão inicial antes da
// calibragem, evitando a circularidade GY≈1 quando D está vazio.
// -----------------------------------------------------------------------------
function atualizarFormulasBaseGE_(sheet) {
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
    for (var j = 0; j < audCols.length; j++) {
        sheet.getRange(63, audCols[j]).setNumberFormat("0.00");
    }

    setFormulaSafe_(sheet.getRange(63, 2), "=(SOMA(B3:B62))/60");
    sheet.getRange(63, 2).setNumberFormat("0.00");
    setFormulaSafe_(sheet.getRange(63, 3), '=SE(OU(B63="";B63=0;D63="");"";D63/B63-1)');
    sheet.getRange(63, 3).setNumberFormat("0.0%");

    SpreadsheetApp.flush();
    sheet.getRange(3, 2, 60, 1).setNumberFormat("0.00");
    sheet.getRange(3, 3, 60, 1).setNumberFormat("0.0%");
}

// --- Atualizar fórmulas do histórico GE ---
// Wrapper que mantém comportamento histórico (base + calibragem + fator tema).
// Preservado para compatibilidade total com arquivarProgramaGE_ e demais chamadas.
function atualizarFormulasHistoricoGE_(sheet) {
    atualizarFormulasBaseGE_(sheet);
    atualizarCalibragemGE_(sheet);
    atualizarFatorTemaGE_(sheet);
}

// --- Garantir campos de rodapé do histórico GE ---

function garantirCamposHistoricoGE_() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    // Mínimo 160 rows: 1-65 cabeçalho+dados, 66-89 painel A66:NT89 (área protegida),
    // 95-155 dados auxiliares do gráfico (fora da área protegida).
    var maxRows = sheet.getMaxRows();
    if (maxRows < 160) sheet.insertRowsAfter(maxRows, 160 - maxRows);
    var labels = [["MÉDIA"], ["COMENT. 1"], ["COMENT. 2"], ["MÉDIA SEM."], ["% vs ANT."]];
    sheet.getRange(63, 1, 5, 1).setValues(labels);
    sheet.getRange(63, 1, 5, 1).setFontWeight("bold").setFontSize(8).setFontColor("#666666");
    preencherBlocoA66E85HistoricoGE_(sheet);
}

// =============================================================================
// preencherBlocoA66E85HistoricoGE_ — painel A66:E89 DINÂMICO
// =============================================================================
//
// REESCRITO. Antes era hardcoded (fórmulas fixas tipo MÉDIA(H63;L63;P63;T63;X63)
// em ranges rígidos) e destrutivo (clearFormat/clearDataValidations/clearNote
// em toda A66:NT89).
//
// Agora:
//   • Lê datas reais da row 1 e agrupa programas por semana (Seg-Sex) e mês.
//   • Gera labels e fórmulas dinâmicas que referenciam as colunas reais.
//   • Limpa F66:NT89 para varrer polução lateral (gráfico legado, média semanal
//     espalhada em F66:NT67).
//   • Preserva formatação, validações e notas.
//   • Preserva D85:D89 por NOME DE MÊS (não por posição) — sobrevive a shifts
//     mensais do painel.
//
// Layout final:
//   Row 66:    SEMANA ATUAL (do programa mais recente, Seg→data mais recente)
//   Rows 67-84: últimas 18 semanas (Seg-Sex), mais recente primeiro
//   Row 84 D:  header "MENSAL" (etiqueta do bloco abaixo)
//   Rows 85-89: mês atual + últimos 4 meses anteriores
//
// Colunas:
//   A: label dinâmico (datas ou nome do mês)
//   B: média (=SEERRO(MÉDIA(refs);0))
//   C: variação % vs linha abaixo (semanas e meses)
//   D: refs mensais comparativas em D85:D89 (preservadas ou padrão JAN-MAIO);
//      header "MENSAL" em D84
//   E: diferença (B/D)-1 em E85:E89 (variação da média vs referência mensal)
// -----------------------------------------------------------------------------
function preencherBlocoA66E85HistoricoGE_(sheet) {
    // --- 0. Preservar valores existentes em D85:D89 (refs mensais comparativas) ---
    //     Mapear {nome do mês → valor} para sobreviver a shifts mensais do painel
    //     (quando o mês corrente muda, as posições deslocam, mas os valores
    //     devem seguir o mês, não a posição). Só preserva valores numéricos
    //     — ignora textos legados como "MENSAL" que ficavam em D85.
    var oldA85_89 = sheet.getRange("A85:A89").getValues();
    var oldD85_89 = sheet.getRange("D85:D89").getValues();
    var preservedByMonth = {};
    for (var pi = 0; pi < 5; pi++) {
        var oldLabel = (oldA85_89[pi][0] || "").toString().trim();
        var oldValue = oldD85_89[pi][0];
        if (typeof oldValue !== "number") continue;
        // Extrair nome do mês (remove sufixo " (mês atual)" se presente)
        var monthName = oldLabel.replace(/\s*\(.*\)\s*$/, "").trim().toUpperCase();
        if (monthName) preservedByMonth[monthName] = oldValue;
    }

    // --- 1. Limpar APENAS conteúdo das áreas que vamos regenerar ---
    //     Não usar clearFormat/clearDataValidations/clearNote — preserva customizações.
    sheet.getRange("A66:E89").clearContent();

    // --- 2. Limpar polução lateral em F66:NT89 (resíduos do gráfico legado e
    //         de versões antigas de atualizarMediaSemanaGE_ que escreviam em
    //         F66:NT67) ---
    var lastCol = sheet.getLastColumn();
    if (lastCol > 5) {
        sheet.getRange(66, 6, 24, lastCol - 5).clearContent();
    }

    // --- 3. Coletar programas reais (col, data) ---
    var audCols = getAudColsGE_(sheet);
    if (audCols.length === 0) {
        // Sem programas: deixar painel vazio
        return;
    }

    var row1Vals = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var programas = [];
    for (var i = 0; i < audCols.length; i++) {
        var c = audCols[i];
        var v = row1Vals[c - 1];
        var dParsed = v instanceof Date ? v : parseDateStr_(typeof v === "string" ? v.trim() : "");
        if (dParsed) programas.push({ col: c, date: dParsed });
    }
    if (programas.length === 0) return;
    programas.sort(function(a, b) { return b.date.getTime() - a.date.getTime(); });

    // --- 4. Helpers de agrupamento ---
    function mondayOf(date) {
        var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
        var dow = d.getDay();
        var diff = dow === 0 ? 6 : dow - 1;
        d.setDate(d.getDate() - diff);
        return d;
    }
    function weekKey(date) {
        var m = mondayOf(date);
        var mm = (m.getMonth() + 1) < 10 ? "0" + (m.getMonth() + 1) : (m.getMonth() + 1);
        var dd = m.getDate() < 10 ? "0" + m.getDate() : m.getDate();
        return m.getFullYear() + "-" + mm + "-" + dd;
    }
    function monthKey(date) {
        var mm = (date.getMonth() + 1) < 10 ? "0" + (date.getMonth() + 1) : (date.getMonth() + 1);
        return date.getFullYear() + "-" + mm;
    }

    // --- 5. Agrupar por semana e por mês ---
    var semanasMap = {};
    var mesesMap = {};
    for (var k = 0; k < programas.length; k++) {
        var p = programas[k];
        var wk = weekKey(p.date);
        if (!semanasMap[wk]) semanasMap[wk] = { startDate: mondayOf(p.date), cols: [] };
        semanasMap[wk].cols.push(p.col);

        var mk = monthKey(p.date);
        if (!mesesMap[mk]) mesesMap[mk] = { date: new Date(p.date.getFullYear(), p.date.getMonth(), 1), cols: [] };
        mesesMap[mk].cols.push(p.col);
    }

    var semanas = Object.keys(semanasMap).map(function(k) {
        return { key: k, startDate: semanasMap[k].startDate, cols: semanasMap[k].cols };
    });
    semanas.sort(function(a, b) { return b.startDate.getTime() - a.startDate.getTime(); });

    var meses = Object.keys(mesesMap).map(function(k) {
        return { key: k, date: mesesMap[k].date, cols: mesesMap[k].cols };
    });
    meses.sort(function(a, b) { return b.date.getTime() - a.date.getTime(); });

    // --- 6. Identificar semana e mês "atuais" (do programa mais recente) ---
    var maisRecente = programas[0];
    var semanaAtualKey = weekKey(maisRecente.date);
    var mesAtualKey = monthKey(maisRecente.date);

    var semanaAtual = null;
    var semanasAnteriores = [];
    for (var s = 0; s < semanas.length; s++) {
        if (semanas[s].key === semanaAtualKey && !semanaAtual) semanaAtual = semanas[s];
        else semanasAnteriores.push(semanas[s]);
    }

    var mesAtual = null;
    var mesesAnteriores = [];
    for (var m2 = 0; m2 < meses.length; m2++) {
        if (meses[m2].key === mesAtualKey && !mesAtual) mesAtual = meses[m2];
        else mesesAnteriores.push(meses[m2]);
    }

    var MONTH_NAMES_PT = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
                          "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];

    // --- 7. Montar matriz 24×5 (rows 66-89) ---
    var labels = [];
    var bFormulas = [];
    var cFormulas = [];
    var TOTAL = 24;
    for (var rInit = 0; rInit < TOTAL; rInit++) {
        labels.push([""]);
        bFormulas.push([""]);
        cFormulas.push([""]);
    }

    function refs63(cols) {
        return cols.map(function(c) { return columnToLetter_(c) + "63"; }).join(";");
    }
    function fmtRange(start) {
        var fim = new Date(start);
        fim.setDate(fim.getDate() + 4);
        return Utilities.formatDate(start, CFG.TIMEZONE, "dd/MM") + "-" +
               Utilities.formatDate(fim, CFG.TIMEZONE, "dd/MM");
    }

    // Row 0 (sheet 66): semana corrente
    if (semanaAtual) {
        labels[0][0] = "SEM. ATUAL (" + fmtRange(semanaAtual.startDate) + ")";
        bFormulas[0][0] = "=SEERRO(MÉDIA(" + refs63(semanaAtual.cols) + ");0)";
        cFormulas[0][0] = '=SE(OU(B67="";B67=0);"";(B66/B67)-1)';
    } else {
        labels[0][0] = "SEM. ATUAL";
    }

    // Rows 1-18 (sheet 67-84): últimas 18 semanas anteriores
    for (var sa = 0; sa < Math.min(18, semanasAnteriores.length); sa++) {
        var sm = semanasAnteriores[sa];
        labels[1 + sa][0] = fmtRange(sm.startDate);
        bFormulas[1 + sa][0] = "=SEERRO(MÉDIA(" + refs63(sm.cols) + ");0)";
        if (sa < 17) {
            var thisR = 66 + 1 + sa;
            var nextR = thisR + 1;
            cFormulas[1 + sa][0] = '=SE(OU(B' + nextR + '="";B' + nextR + '=0);"";(B' + thisR + '/B' + nextR + ')-1)';
        }
    }

    // Row 19 (sheet 85): mês corrente
    if (mesAtual) {
        labels[19][0] = MONTH_NAMES_PT[mesAtual.date.getMonth()] + " (mês atual)";
        bFormulas[19][0] = "=SEERRO(MÉDIA(" + refs63(mesAtual.cols) + ");0)";
        cFormulas[19][0] = '=SE(OU(B86="";B86=0);"";(B85/B86)-1)';
    } else {
        labels[19][0] = "MÊS ATUAL";
    }

    // Rows 20-23 (sheet 86-89): últimos 4 meses anteriores
    for (var mp = 0; mp < Math.min(4, mesesAnteriores.length); mp++) {
        var mn = mesesAnteriores[mp];
        labels[20 + mp][0] = MONTH_NAMES_PT[mn.date.getMonth()];
        bFormulas[20 + mp][0] = "=SEERRO(MÉDIA(" + refs63(mn.cols) + ");0)";
        if (mp < 3) {
            var thisM = 86 + mp;
            var nextM = thisM + 1;
            cFormulas[20 + mp][0] = '=SE(OU(B' + nextM + '="";B' + nextM + '=0);"";(B' + thisM + '/B' + nextM + ')-1)';
        }
    }

    // --- 8. Aplicar valores e fórmulas (sem clearFormat) ---
    sheet.getRange("A66:A89").setValues(labels);

    // Aplicar fórmulas B (filtra strings vazias para evitar #ERROR! em células sem cálculo)
    var bRange = sheet.getRange("B66:B89");
    var bValuesFinal = [];
    for (var bi = 0; bi < bFormulas.length; bi++) {
        bValuesFinal.push([bFormulas[bi][0] === "" ? "" : bFormulas[bi][0]]);
    }
    bRange.setValues(bValuesFinal);

    var cRange = sheet.getRange("C66:C89");
    var cValuesFinal = [];
    for (var ci = 0; ci < cFormulas.length; ci++) {
        cValuesFinal.push([cFormulas[ci][0] === "" ? "" : cFormulas[ci][0]]);
    }
    cRange.setValues(cValuesFinal);

    // D84: header "MENSAL" (label estático, fora da área de dados mensais)
    sheet.getRange("D84").setValue("MENSAL");

    // D85:D89: refs mensais comparativas
    //   Prioridade por linha: (1) valor preservado por NOME DO MÊS, (2) padrão
    //   do modelo para meses conhecidos (JAN-MAIO), (3) vazio.
    // E85:E89: diferença (B/D)-1 entre média calculada e referência mensal.
    var MONTHLY_DEFAULTS = {
        "JANEIRO": 0.56,
        "FEVEREIRO": 0.63,
        "MARÇO": 0.71,
        "ABRIL": 0.70,
        "MAIO": 0.65
    };
    var dValues = [];
    var eFormulas = [];
    for (var mr = 0; mr < 5; mr++) {
        var labelIdx = 19 + mr; // labels[19] = row 85, ..., labels[23] = row 89
        var labelStr = (labels[labelIdx][0] || "").toString();
        var monthName = labelStr.replace(/\s*\(.*\)\s*$/, "").trim().toUpperCase();
        var dv = "";
        if (monthName) {
            if (preservedByMonth.hasOwnProperty(monthName)) {
                dv = preservedByMonth[monthName];
            } else if (MONTHLY_DEFAULTS.hasOwnProperty(monthName)) {
                dv = MONTHLY_DEFAULTS[monthName];
            }
        }
        dValues.push([dv]);
        var rNum = 85 + mr;
        eFormulas.push(['=SE(OU(D' + rNum + '="";D' + rNum + '=0);"";(B' + rNum + '/D' + rNum + ')-1)']);
    }
    sheet.getRange("D85:D89").setValues(dValues);
    sheet.getRange("E85:E89").setValues(eFormulas);

    // --- 9. Formatação numérica (sem destruir formatação base existente) ---
    sheet.getRange("B66:B89").setNumberFormat("0.00");
    sheet.getRange("C66:C89").setNumberFormat("+0.0%;-0.0%");
    sheet.getRange("D85:D89").setNumberFormat("0.00");
    sheet.getRange("E85:E89").setNumberFormat("+0.0%;-0.0%");
    sheet.getRange("A66:A89").setHorizontalAlignment("left").setFontWeight("bold");
    sheet.getRange("B66:E89").setHorizontalAlignment("center");

    Logger.log("[Painel A66:E89] Reconstruído dinamicamente: " +
        (semanaAtual ? 1 : 0) + " sem atual + " +
        Math.min(18, semanasAnteriores.length) + " sem anteriores + " +
        (mesAtual ? 1 : 0) + " mês atual + " +
        Math.min(4, mesesAnteriores.length) + " meses anteriores. " +
        "Polução em F66:NT89 limpa.");
}

// --- Média por semana GE ---
//
// REESCRITO. Antes pulverizava valores em (66, sexta.col) e (67, sexta.col) de
// cada coluna de sexta-feira histórica — espalhando 80+ valores avulsos pelas
// colunas F:NT das linhas 66 e 67, dentro da área protegida A66:NT89.
//
// Agora a média semanal vive integralmente no painel dinâmico A66:E89
// (preencherBlocoA66E85HistoricoGE_). Esta função apenas limpa eventuais
// resíduos de versões anteriores nas rows 66-67 das colunas F+ —
// safety-net se algo escreveu lá entre execuções.
// -----------------------------------------------------------------------------
function atualizarMediaSemanaGE_() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    var lastCol = sheet.getLastColumn();
    if (lastCol > 5) {
        // Limpa F66:lastCol67 (resíduo de versões anteriores)
        sheet.getRange(66, 6, 2, lastCol - 5).clearContent();
    }
    Logger.log("[atualizarMediaSemanaGE_] Resíduos em F66:NT67 limpos. Cálculo semanal vive em A66:E89.");
}

// --- Gráfico de médias GE ---
//
// REESCRITO. Antes usava DATA_START_ROW = 70, escrevendo 60 linhas de dados
// auxiliares em A70:D130 — DENTRO da área protegida A66:NT89.
//
// Agora DATA_START_ROW = 95 (fora da área protegida). A função limpa
// preventivamente:
//   • o gap A90:D94 entre o painel e a área auxiliar (cobre resíduo legado
//     de versões antigas que escreviam até a row 130);
//   • a nova área A95:D155 (onde os dados serão escritos).
//
// IMPORTANTE: a função NUNCA toca rows 70-89 — está dentro do painel
// A66:NT89 mantido por preencherBlocoA66E85HistoricoGE_.
// -----------------------------------------------------------------------------
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

    // FIX CRÍTICO: row de início = 95, fora da área protegida A66:NT89.
    var DATA_START_ROW = 95;

    // Limpeza preventiva APENAS FORA da área protegida A66:NT89.
    // NÃO tocar A70:D89: está dentro do painel semanal/mensal recém-construído
    // por preencherBlocoA66E85HistoricoGE_ — limpar aqui destruiria o painel.
    // O gap entre o painel e a nova área auxiliar (A90:D94) cobre eventual
    // resíduo de versões antigas (que escreviam até row 130 a partir de row 70).
    // A nova área (A95:D155) é coberta pela limpeza imediatamente abaixo.
    sheet.getRange(90, 1, 5, 4).clearContent(); // A90:D94 (gap fora da área protegida)
    // Limpar área NOVA (A95:D155) preventivamente
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

// =============================================================================
// gerarPrevisaoInicialGE_ — NOVO modelo isolado de previsão inicial
// =============================================================================
//
// Constrói a previsão de audiência do novo programa a partir SOMENTE dos
// programas históricos reais (colunas H em diante), sem depender das colunas
// auxiliares de calibragem GS-GY (que seriam circulares enquanto D está vazio).
//
// Para cada minuto r ∈ [3,62] computa:
//   base(r)    = média de todas as audiências históricas no minuto
//   f_trend(r) = clamp(média últimos 5 / base, [0.75, 1.30])
//   f_accel(r) = clamp(média últimos 3 / média últimos 5, [0.85, 1.18])
//   f_dia(r)   = clamp(média mesmo dia-da-semana / base, [0.78, 1.28])
//
// Combina pelo regime aplicável:
//   • Normal: 0.40·f_dia + 0.35·f_trend + 0.25·f_accel
//   • Ruptura (|últimos_3 / restante − 1| ≥ 15%): blend 50/50 com
//                0.20·f_dia + 0.55·f_trend + 0.25·f_accel
//
// Escreve em D3:D62: `=ARRED(<base*fator_combinado> * HC{r}, 2)`
// → HC (fator tema) permanece dinâmico, reagindo quando F é preenchido.
//
// Retorna `true` se preencheu previsão, `false` se não há histórico suficiente.
// -----------------------------------------------------------------------------
function gerarPrevisaoInicialGE_(sheet, programDate) {
    var audCols = getAudColsGE_(sheet);
    // audCols[0]=4 é o novo dia (vazio); demais são histórico real
    var audColsHist = audCols.filter(function(c) { return c !== 4; });

    if (audColsHist.length === 0) {
        Logger.log("[Prev Inicial] Sem programas históricos — D3:D62 fica vazio");
        return false;
    }

    // Identificar programas do mesmo dia da semana
    var weekday = programDate.getDay();
    var lastCol = sheet.getLastColumn();
    var row1Vals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    var colsMesmoDia = [];
    for (var i = 0; i < audColsHist.length; i++) {
        var c = audColsHist[i];
        var v = row1Vals[c - 1];
        var dParsed = v instanceof Date
            ? v
            : parseDateStr_(typeof v === "string" ? v.trim() : "");
        if (dParsed && dParsed.getDay() === weekday) colsMesmoDia.push(c);
    }

    // Últimos N programas (audColsHist está em ordem newest-first, pois col 8 = mais recente)
    var colsRecent5 = audColsHist.slice(0, Math.min(PREV_INICIAL_GE.N_RECENT_5, audColsHist.length));
    var colsRecent3 = audColsHist.slice(0, Math.min(PREV_INICIAL_GE.N_RECENT_3, audColsHist.length));

    // Sets para lookup O(1)
    var recent5Set = {};
    for (var k1 = 0; k1 < colsRecent5.length; k1++) recent5Set[colsRecent5[k1]] = true;
    var recent3Set = {};
    for (var k2 = 0; k2 < colsRecent3.length; k2++) recent3Set[colsRecent3[k2]] = true;
    var diaSet = {};
    for (var k3 = 0; k3 < colsMesmoDia.length; k3++) diaSet[colsMesmoDia[k3]] = true;

    // Leitura em batch da matriz histórica (rows 3-62 × colunas audColsHist)
    var minCol = audColsHist[0];                          // col 8 (mais recente)
    var maxCol = audColsHist[audColsHist.length - 1];     // col mais antiga (mais à direita)
    var blockData = sheet.getRange(3, minCol, 60, maxCol - minCol + 1).getValues();

    // Localizar posição de HC (fator tema) — fórmula precisa apontar dinamicamente
    var calibStartCol = 4 + audCols.length * 4;
    var temaSheetExists = !!SpreadsheetApp.getActive().getSheetByName("GE - Média por Minuto TEMA");
    var hcColLetter = temaSheetExists ? columnToLetter_(calibStartCol + 7 + 3) : null;

    // Telemetria
    var stats = {
        cols_historicas: audColsHist.length,
        cols_mesmo_dia: colsMesmoDia.length,
        cols_recent5: colsRecent5.length,
        cols_recent3: colsRecent3.length,
        weekday: weekday,
        minutos_preenchidos: 0,
        minutos_vazios: 0,
        minutos_ruptura: 0,
        minutos_base_only: 0,
        com_hc: temaSheetExists
    };

    var formulas = [];
    for (var r = 3; r <= 62; r++) {
        var rowIdx = r - 3;

        var allVals = [], recent5Vals = [], recent3Vals = [], diaVals = [];

        for (var ii = 0; ii < audColsHist.length; ii++) {
            var colH = audColsHist[ii];
            var colIdx = colH - minCol;
            var val = blockData[rowIdx][colIdx];
            if (typeof val !== "number" || val <= 0) continue;
            allVals.push(val);
            if (recent5Set[colH]) recent5Vals.push(val);
            if (recent3Set[colH]) recent3Vals.push(val);
            if (diaSet[colH]) diaVals.push(val);
        }

        if (allVals.length === 0) {
            formulas.push([""]);
            stats.minutos_vazios++;
            continue;
        }

        var base = mediaArrGE_(allVals);

        // f_trend: média recente 5 / base
        var fTrend = 1;
        if (recent5Vals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_TREND && base > 0) {
            fTrend = clampGE_(mediaArrGE_(recent5Vals) / base,
                PREV_INICIAL_GE.CLAMP_TREND[0], PREV_INICIAL_GE.CLAMP_TREND[1]);
        }

        // f_accel: média recente 3 / média recente 5 (captura aceleração)
        var fAccel = 1;
        if (recent3Vals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_ACCEL &&
            recent5Vals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_TREND) {
            var mr5 = mediaArrGE_(recent5Vals);
            if (mr5 > 0) {
                fAccel = clampGE_(mediaArrGE_(recent3Vals) / mr5,
                    PREV_INICIAL_GE.CLAMP_ACCEL[0], PREV_INICIAL_GE.CLAMP_ACCEL[1]);
            }
        }

        // f_dia: média mesmo dia / base
        var fDia = 1;
        if (diaVals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_DIA && base > 0) {
            fDia = clampGE_(mediaArrGE_(diaVals) / base,
                PREV_INICIAL_GE.CLAMP_DIA[0], PREV_INICIAL_GE.CLAMP_DIA[1]);
        }

        // Combinação ponderada regime normal
        var fatorNormal = PREV_INICIAL_GE.PESO_DIA   * fDia
                        + PREV_INICIAL_GE.PESO_TREND * fTrend
                        + PREV_INICIAL_GE.PESO_ACCEL * fAccel;

        var fatorCombinado = fatorNormal;

        // Pouco histórico: só usa base (sem fator)
        if (recent5Vals.length < PREV_INICIAL_GE.MIN_AMOSTRAS_TREND &&
            diaVals.length < PREV_INICIAL_GE.MIN_AMOSTRAS_DIA) {
            fatorCombinado = 1;
            stats.minutos_base_only++;
        } else {
            // Detector de ruptura: últimos 3 vs restante
            if (recent3Vals.length >= 3 && allVals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_RUPTURA) {
                var sumAll = 0, sumRec3 = 0;
                for (var sa = 0; sa < allVals.length; sa++) sumAll += allVals[sa];
                for (var sb = 0; sb < recent3Vals.length; sb++) sumRec3 += recent3Vals[sb];
                var nRest = allVals.length - recent3Vals.length;
                if (nRest > 0) {
                    var meanRest = (sumAll - sumRec3) / nRest;
                    if (meanRest > 0) {
                        var ratioRupt = (sumRec3 / recent3Vals.length) / meanRest;
                        if (ratioRupt <= 1 - PREV_INICIAL_GE.THRESHOLD_RUPTURA ||
                            ratioRupt >= 1 + PREV_INICIAL_GE.THRESHOLD_RUPTURA) {
                            var fatorRupt = PREV_INICIAL_GE.PESO_DIA_RUPT   * fDia
                                          + PREV_INICIAL_GE.PESO_TREND_RUPT * fTrend
                                          + PREV_INICIAL_GE.PESO_ACCEL_RUPT * fAccel;
                            fatorCombinado = PREV_INICIAL_GE.BLEND_RUPTURA * fatorNormal
                                           + (1 - PREV_INICIAL_GE.BLEND_RUPTURA) * fatorRupt;
                            stats.minutos_ruptura++;
                        }
                    }
                }
            }
        }

        var predValue = base * fatorCombinado;

        // Construir fórmula (locale BR: vírgula decimal, ponto-e-vírgula separador)
        var predStr = predValue.toFixed(4).replace(".", ",");
        var formula = (temaSheetExists && hcColLetter)
            ? '=ARRED(' + predStr + '*' + hcColLetter + r + ';2)'
            : '=ARRED(' + predStr + ';2)';
        formulas.push([formula]);
        stats.minutos_preenchidos++;
    }

    batchSetFormulasSafe_(sheet.getRange(3, 4, 60, 1), formulas);
    SpreadsheetApp.flush();

    Logger.log("[Prev Inicial GE] " + JSON.stringify(stats));
    return true;
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

// =============================================================================
// adicionarNovoDiaManualGE — REORGANIZADO (entrada pública via menu)
// =============================================================================
//
// Nova ordem de execução:
//   1. Inserir bloco D:G vazio
//   2. atualizarFormulasBaseGE_ (apenas B/C/row 63/share — SEM calibragem)
//   3. Aplicar fórmula de share em E
//   4. gerarPrevisaoInicialGE_ — preenche D3:D62 com modelo isolado
//   5. atualizarCalibragemGE_ + atualizarFatorTemaGE_ — agora com D preenchido
//   6. Bloco row 63-65 (mediana, comentaristas, diferença)
//   7. Formatação condicional herdada
//   8. Atualizar dashboards
//
// Isto elimina a circularidade GS×GY×HC que ocorria quando D estava vazio.
// -----------------------------------------------------------------------------
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
        "Digite a data do programa (dd/mm/aaaa):\n\nA previsão inicial será calculada a partir do histórico real, sem depender de dados ao vivo.",
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

    // Validar duplicata
    var lastCol = sheet.getLastColumn();
    var numProgramsBefore = Math.floor((lastCol - 3) / 4);
    if (numProgramsBefore > 0) {
        var existingDates = sheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
        for (var dPos = 0; dPos < existingDates.length; dPos += 4) {
            var existing = existingDates[dPos];
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

    // ❶ Inserir bloco vazio D:G
    sheet.insertColumns(4, 4);

    var bgColor = (numProgramsBefore % 2 === 0) ? "#FDF2F2" : "#FFFFFF";

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

    // ❷ Atualizar SOMENTE fórmulas-base (B/C/row63) — SEM calibragem
    //    Importante: nesta etapa NÃO chamamos atualizarCalibragemGE_, pois com
    //    D ainda vazio teríamos GT="" e GY≈1 (modelo degenerado).
    atualizarFormulasBaseGE_(sheet);
    Utilities.sleep(800);

    // ❸ Definir fórmula de share em E (=D/B) — compatível com a previsão futura
    var shareFormulas = [];
    for (var row = 3; row <= 62; row++) {
        shareFormulas.push(['=SE($B' + row + '=0;"";D' + row + '/$B' + row + ')']);
    }
    batchSetFormulasSafe_(sheet.getRange(3, 5, 60, 1), shareFormulas);

    // ❹ PREVISÃO INICIAL ROBUSTA usando modelo isolado de histórico
    var temPrevisao = gerarPrevisaoInicialGE_(sheet, programDate);
    Utilities.sleep(800);

    // ❺ Agora SIM chamamos calibragem (D preenchido com previsão → GT-GY válidos)
    //    Atenção arquitetural: com D contendo previsão, GT=D/GS produzirá valores
    //    consistentes com a previsão. Quando dados reais começarem a entrar em
    //    D ao vivo (sobrescrevendo célula a célula), GT-GY automaticamente
    //    passam a refletir os dados reais — sem mudança estrutural necessária.
    atualizarCalibragemGE_(sheet);
    atualizarFatorTemaGE_(sheet);
    Utilities.sleep(500);

    // ❻ Bloco row 63-65: espelha bloco anterior H63:K65, define labels, mediana
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
    }[weekday] || "MEDIANA";
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

    // ❼ Herdar formatação condicional da coluna equivalente (&/MIN) do bloco anterior: I3:I62 -> E3:E62
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

    // ❽ Dashboards
    garantirCamposHistoricoGE_();
    atualizarMediaSemanaGE_();
    Utilities.sleep(1000);
    atualizarGraficoMediasGE_();
    Utilities.sleep(1000);
    atualizarMediaPorComentaristaGE_();
    atualizarMediaPorDiaSemanaGE_();
    atualizarMediaPorTemaSeExistir_();

    SpreadsheetApp.flush();

    var msgPrev = temPrevisao
        ? "📊 Previsão inicial aplicada (modelo isolado de histórico)\n"
        : "⚠️ Sem histórico suficiente — D3:D62 vazio\n";

    ui.alert("✅ Novo dia adicionado!",
        "Programa de " + dateStr + " inserido!\n\n" +
        msgPrev +
        "📈 Calibragem GS–GY recalculada\n" +
        "📈 Gráfico atualizado\n\n" +
        "⚠️ Preencha:\n" +
        "• Linha 64 (col D): Comentarista 1\n" +
        "• Linha 65 (col D): Comentarista 2\n" +
        "• Coluna F (TEMA): preenchimento atualiza HC dinamicamente",
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
