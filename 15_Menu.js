// =============================================================================
// 15_Menu.js — onOpen, entry points JA + GE, e utilitários de menu
// =============================================================================

// --- Função de teste de fórmulas pt-BR (via menu) ---

function TESTE_formula_ptBR() {
    const e = SpreadsheetApp.getActiveSheet();
    setFormulaSafe_(e.getRange("A1"), '=SE(1=1;"OK";"ERRO")')
}

// =============================================================================
// onOpen — constrói os 4 menus customizados
// =============================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // --- Menu Jogo Aberto ---
  ui.createMenu("📺 Jogo Aberto")
    .addItem("▶️ Rodar Análise Completa + Email", "runAll")
    .addSeparator()
    .addSubMenu(
      ui.createMenu("🎙️ Importar Transcrição")
        .addItem("🔄 Automático (programa de hoje)", "importarTranscricaoAutomatica")
        .addItem("🔗 Por URL do YouTube", "importarTranscricaoPorUrl")
        .addItem("📅 Por Data", "importarTranscricaoPorData")
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu("📊 Histórico")
        .addItem("📥 Arquivar programa atual", "arquivarProgramaAtual")
        .addItem("📋 Ver estrutura do histórico", "verEstruturaHistorico")
    )
    .addSeparator()
    .addItem("🔄 Atualizar Minuto a Minuto", "runLinhaDoTempo")
    .addItem("📊 Atualizar Top 20 Termos", "runTop20")
    .addItem("📧 Reenviar Relatório por Email", "runRelatorioEmail")
    .addSeparator()
    .addItem("🌐 Abrir YouTube Jogo Aberto", "abrirYouTubeJogoAberto")
    .addToUi();

  // --- Menu Gazeta Esportiva ---
  ui.createMenu("📺 Gazeta Esportiva")
    .addItem("▶️ Rodar Análise Completa + Email", "runAllGazeta")
    .addItem("🚦 Atualizar Alerta de Tema", "atualizarAlertaTemaGE")
    .addItem("⚽ Atualizar Calendário Clubes", "atualizarCalendarioClubes")
    .addItem("📊 Gerar Análise Contexto Jogos", "gerarAnaliseContexto")
    .addSeparator()
    .addSubMenu(
      ui.createMenu("🎙️ Importar Transcrição")
        .addItem("🔄 Automático (programa de hoje)", "importarTranscricaoAutomaticaGazeta")
        .addItem("🔗 Por URL do YouTube", "importarTranscricaoPorUrlGazeta")
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu("📊 Histórico")
        .addItem("📥 Arquivar programa atual", "arquivarProgramaAtualGazeta")
        .addItem("➕ Adicionar Novo Dia (Manual)", "adicionarNovoDiaManualGE")
        .addItem("📋 Ver estrutura do histórico", "verEstruturaHistoricoGazeta")
    )
    .addSubMenu(
      ui.createMenu("🔮 Calibragem / Predição")
        .addItem("🔄 Atualizar Calibragem (GS–GY)", "atualizarCalibragemGE")
        .addItem("🎯 Instalar/Atualizar Projeção de Audiência", "instalarProjecaoAudienciaGE")
        .addItem("➕🎯 Adicionar Novo Dia + Projeção", "arquivarEProjetarGE")
        .addItem("🚀 Atualização Completa (Tema → Dia → Calibrar → Projeção)", "executarAtualizacaoHistoricoGECompleta")
    )
    .addSubMenu(
      ui.createMenu("📈 Análise por Tema")
        .addItem("📊 Criar/Atualizar Média por Tema", "criarMediaPorTemaGE")
        .addItem("🔄 Atualizar Fórmulas", "atualizarFormulasTemaGE")
    )
    .addSeparator()
    .addItem("🔄 Atualizar Minuto a Minuto", "runLinhaDoTempoGazeta")
    .addItem("📊 Atualizar Top 20 Termos", "runTop20Gazeta")
    .addItem("📧 Reenviar Relatório por Email", "runRelatorioEmailGazeta")
    .addSeparator()
    .addItem("🌐 Abrir YouTube Gazeta Esportiva", "abrirYouTubeGazetaEsportiva")
    .addToUi();

  // --- Menu Comparativo ---
  ui.createMenu("📊 Comparativo")
    .addItem("📝 Gerar Comparativo JA x GE", "gerarComparativoJAxGE")
    .addItem("📧 Enviar Comparativo por Email", "enviarComparativoEmail")
    .addToUi();

  // --- Menu Configurações ---
  ui.createMenu("⚙️ Configurações")
    .addSubMenu(
      ui.createMenu("⏰ Agendamento Jogo Aberto")
        .addItem("✅ Ativar (15:00)", "ativarAgendamentoJA")
        .addItem("❌ Desativar", "desativarAgendamentoJA")
        .addItem("📋 Ver status", "verStatusAgendamentoJA")
    )
    .addSubMenu(
      ui.createMenu("⏰ Agendamento Gazeta")
        .addItem("✅ Ativar (21:00)", "ativarAgendamentoGE")
        .addItem("❌ Desativar", "desativarAgendamentoGE")
        .addItem("📋 Ver status", "verStatusAgendamentoGE")
    )
    .addSeparator()
    .addItem("⚙️ Configurar Servidor", "configurarServidor")
    .addSeparator()
    .addItem("🗑️ Limpar Cache LLM", "clearLLMCache")
    .addItem("🗑️ Limpar Stopwords", "clearUserStopwords")
    .addToUi();
}

// =============================================================================
// onEdit — auto-atualizar "GE - Média por Comentarista" ao digitar nome
//          + reconstruir cadeia de projeção de audiência (coluna D)
// =============================================================================

function onEdit(e) {
    if (!e || !e.range) return;
    try {
        var sheet = e.range.getSheet();
        var histName = (typeof CFG !== "undefined" && CFG.GE_OUT_HISTORICO)
            ? CFG.GE_OUT_HISTORICO : "GE - Histórico";
        if (sheet.getName() === histName) {
            var row = e.range.getRow();
            var col = e.range.getColumn();

            // Dispara quando comentarista 1 (row 64) ou comentarista 2 (row 65) é editado
            if (row === 64 || row === 65) {
                atualizarMediaPorComentaristaGE_();
            }

            // Projeção de audiência: ao editar a coluna D (4) nas linhas 3..62,
            // reconstrói a cadeia de projeção nas células D vazias abaixo.
            // As células com real digitado são preservadas; as fórmulas abaixo
            // recalculam sozinhas pelo motor da planilha.
            if (col === 4 && row >= 3 && row <= 62 &&
                typeof reconstruirCadeiaProjecaoGE_ === "function") {
                reconstruirCadeiaProjecaoGE_(sheet, row);
            }
        }
    } catch (err) {
        // Silencioso para não interromper a edição do usuário
    }
    // Protegido: um erro aqui não deve travar a digitação do usuário.
    try {
        if (typeof onEditAlertaTemaGE === "function") onEditAlertaTemaGE(e);
    } catch (err2) {
        // silencioso
    }
}

// =============================================================================
// Skeleton compartilhado JA + GE
// =============================================================================

function runAnalysisSkeleton_(e) {
    const t = e.loadInput();
    if (0 === t.rows.length) return e.onEmptyInput(), null;
    const o = e.processInput(t);
    if (!1 !== e.checkProcessed && 0 === o.rows.length) return e.onEmptyProcessed(), null;
    e.ensureSheets();
    const a = getStopwords_(),
        r = computeTopTermsPerMinute_(o.rows, a);
    return {
        rawInput: t,
        input: o,
        stop: a,
        minuteTopTerms: r,
        timeline: e.buildTimeline ? e.buildTimeline(o.rows, r, a) : null
    }
}

function ensureSheetHasData_(e, t) {
    const o = SpreadsheetApp.getActive().getSheetByName(e);
    return !o || o.getLastRow() < 2 ? (t(), null) : o
}

// =============================================================================
// Entry points — Jogo Aberto
// =============================================================================

function runAll() {
    const e = SpreadsheetApp.getUi(),
        t = Date.now(),
        o = runAnalysisSkeleton_({
            loadInput: loadInputFromSheet_,
            onEmptyInput: () => e.alert('Aba "Input" est\xe1 vazia ou n\xe3o cont\xe9m dados v\xe1lidos.'),
            processInput: processAndMapTimestamps_,
            onEmptyProcessed: () => e.alert("N\xe3o foi poss\xedvel detectar o in\xedcio do Jogo Aberto na transcri\xe7\xe3o."),
            ensureSheets: ensureOutputSheets_,
            buildTimeline: buildTimeline_
        });
    if (!o) return;
    const {
        rawInput: a,
        input: r,
        stop: n,
        minuteTopTerms: s,
        timeline: i
    } = o;
    writeAllOutputs_(i, s, n, r.rows);
    const c = ((Date.now() - t) / 1e3).toFixed(1),
        l = {};
    for (const e of i) l[e.tema] = (l[e.tema] || 0) + 1;
    const d = l.MERCHAN || 0;
    let u = "";
    try {
        const e = extrairDataPrograma_() || new Date;
        u = arquivarPrograma_(e) ? "\n\u{1f4c1} Arquivado no Hist\xf3rico: " + Utilities.formatDate(e, "America/Sao_Paulo", "dd/MM/yyyy") : "\n\u{1f4c1} Hist\xf3rico: programa j\xe1 arquivado"
    } catch (e) {
        u = "\n\u26a0\ufe0f Erro ao arquivar: " + e.message, Logger.log("Erro ao arquivar no hist\xf3rico: " + e.message)
    }
    try {
        enviarEmailRelatorio_(gerarRelatorio_(i)), e.alert(`\u2705 An\xe1lise conclu\xedda em ${c}s!\n\n\u{1f4ca} Transcri\xe7\xe3o: ${a.rows.length} minutos brutos\n\u{1f3ac} Jogo Aberto detectado: minuto ${r.jogoAbertoStartIndex+1}\n\u{1f4fa} Sa\xedda: ${i.length} minutos (11:06-12:59)\n\u{1f4b0} MERCHAN: ${d} minutos (${(d/i.length*100).toFixed(1)}%)` + u + "\n\n\u{1f4e7} Relat\xf3rio enviado para " + CFG.EMAIL_DESTINO)
    } catch (t) {
        e.alert(`\u26a0\ufe0f An\xe1lise conclu\xedda em ${c}s, mas houve erro no email:\n${t.message}\n\n\u{1f4ca} Transcri\xe7\xe3o: ${a.rows.length} minutos brutos\n\u{1f3ac} Jogo Aberto detectado: minuto ${r.jogoAbertoStartIndex+1}\n\u{1f4fa} Sa\xedda: ${i.length} minutos (11:06-12:59)\n\u{1f4b0} MERCHAN: ${d} minutos (${(d/i.length*100).toFixed(1)}%)` + u)
    }
}

function runLinhaDoTempo() {
    const e = runAnalysisSkeleton_({
        loadInput: loadInputFromSheet_,
        onEmptyInput: () => SpreadsheetApp.getUi().alert("Aba INPUT est\xe1 vazia."),
        processInput: processAndMapTimestamps_,
        onEmptyProcessed: () => SpreadsheetApp.getUi().alert("N\xe3o foi poss\xedvel detectar o in\xedcio do Jogo Aberto."),
        ensureSheets: ensureOutputSheets_,
        buildTimeline: buildTimeline_
    });
    if (!e) return;
    const {
        timeline: t
    } = e;
    writeLinhaDoTempo_(t), writeMacro_(t), writeSubtemas_(t), SpreadsheetApp.getUi().alert(`Linha do Tempo gerada: ${t.length} entradas.`)
}

function runTop20() {
    const e = runAnalysisSkeleton_({
        loadInput: loadInputFromSheet_,
        onEmptyInput: () => SpreadsheetApp.getUi().alert("Aba INPUT est\xe1 vazia."),
        processInput: processAndMapTimestamps_,
        onEmptyProcessed: () => SpreadsheetApp.getUi().alert("N\xe3o foi poss\xedvel detectar o in\xedcio do Jogo Aberto."),
        ensureSheets: ensureOutputSheets_
    });
    if (!e) return;
    const {
        input: t,
        stop: o,
        minuteTopTerms: a
    } = e;
    writeTop20_(computeTop20_(t.rows, a, o)), SpreadsheetApp.getUi().alert("Top 20 Termos atualizado.")
}

function runRelatorioEmail() {
    const e = SpreadsheetApp.getUi();
    if (ensureSheetHasData_(CFG.OUT_TIME, () => e.alert("\u274c Erro", 'A aba "Minuto a Minuto" est\xe1 vazia. Execute a an\xe1lise completa primeiro.', e.ButtonSet.OK))) try {
        enviarEmailRelatorio_(gerarRelatorio_(null)), e.alert("\u2705 Sucesso", "Relat\xf3rio enviado para " + CFG.EMAIL_DESTINO, e.ButtonSet.OK)
    } catch (t) {
        e.alert("\u274c Erro", "Falha ao enviar email: " + t.message, e.ButtonSet.OK)
    }
}

function abrirYouTubeJogoAberto() {
    const e = HtmlService.createHtmlOutput('\n    <html>\n      <body>\n        <p>Clique no link abaixo para abrir:</p>\n        <a href="https://www.youtube.com/@JogoAberto/streams" target="_blank" \n           onclick="google.script.host.close();" \n           style="font-size: 16px; color: blue;">\n          \u{1f4fa} Abrir YouTube Jogo Aberto\n        </a>\n        <br><br>\n        <button onclick="google.script.host.close();">Fechar</button>\n      </body>\n    </html>\n  ').setWidth(300).setHeight(120);
    SpreadsheetApp.getUi().showModalDialog(e, "YouTube Jogo Aberto")
}

// =============================================================================
// Entry points — Gazeta Esportiva
// =============================================================================

function runAllGazeta() {
    const e = SpreadsheetApp.getUi(),
        t = Date.now(),
        o = runAnalysisSkeleton_({
            loadInput: loadGEInputFromSheet_,
            onEmptyInput: () => e.alert('Aba "' + CFG.GE_INPUT_SHEET + '" est\xe1 vazia ou n\xe3o cont\xe9m dados v\xe1lidos.'),
            processInput: processAndMapTimestampsGE_,
            onEmptyProcessed: () => e.alert("N\xe3o foi poss\xedvel processar a transcri\xe7\xe3o do Gazeta Esportiva."),
            ensureSheets: ensureGESheets_,
            buildTimeline: buildTimelineGE_
        });
    if (!o) return;
    const {
        rawInput: a,
        input: r,
        stop: n,
        minuteTopTerms: s,
        timeline: i
    } = o;
    writeAllOutputsGE_(i, s, n, r.rows);
    const c = ((Date.now() - t) / 1e3).toFixed(1),
        l = {};
    for (const e of i) l[e.tema] = (l[e.tema] || 0) + 1;
    const d = l.MERCHAN || 0;
    let u = "";
    if (r.breakInfo && r.breakInfo.hasBreak) {
        const e = r.breakInfo.breakStart,
            t = r.breakInfo.breakEnd,
            o = r.breakInfo.breakDuration;
        u = `\n\u{1f507} Break detectado: 18:${(e+"").padStart(2,"0")} a 18:${(t+"").padStart(2,"0")} (${o} min)`
    }
    let m = "";
    try {
        const e = extrairDataProgramaGE_() || new Date;
        m = arquivarProgramaGE_(e) ? "\n\u{1f4c1} Arquivado no Hist\xf3rico: " + Utilities.formatDate(e, "America/Sao_Paulo", "dd/MM/yyyy") : "\n\u{1f4c1} Hist\xf3rico: programa j\xe1 arquivado"
    } catch (e) {
        m = "\n\u26a0\ufe0f Erro ao arquivar: " + e.message, Logger.log("Erro ao arquivar no hist\xf3rico GE: " + e.message)
    }
    try {
        enviarEmailRelatorioGE_(gerarRelatorioGE_(i)), e.alert(`\u2705 An\xe1lise Gazeta Esportiva conclu\xedda em ${c}s!\n\n\u{1f4ca} Transcri\xe7\xe3o: ${a.rows.length} minutos brutos\n\u{1f4fa} Sa\xedda: ${i.length} minutos (18:00-18:59)\n\u{1f4b0} MERCHAN/Break: ${d} minutos (${(d/i.length*100).toFixed(1)}%)` + u + m + "\n\n\u{1f4e7} Relat\xf3rio enviado para " + CFG.EMAIL_DESTINO)
    } catch (t) {
        e.alert(`\u26a0\ufe0f An\xe1lise conclu\xedda em ${c}s, mas houve erro no email:\n${t.message}` + u + m)
    }
}

function runLinhaDoTempoGazeta() {
    const e = runAnalysisSkeleton_({
        loadInput: loadGEInputFromSheet_,
        onEmptyInput: () => SpreadsheetApp.getUi().alert('Aba "' + CFG.GE_INPUT_SHEET + '" est\xe1 vazia.'),
        processInput: processAndMapTimestampsGE_,
        onEmptyProcessed: () => SpreadsheetApp.getUi().alert("N\xe3o foi poss\xedvel processar a transcri\xe7\xe3o."),
        ensureSheets: ensureGESheets_,
        buildTimeline: buildTimelineGE_
    });
    if (!e) return;
    const {
        timeline: t
    } = e;
    writeLinhaDoTempoGE_(t), SpreadsheetApp.getUi().alert(`\u2705 Minuto a Minuto atualizado!\n\n${t.length} minutos processados.`)
}

function runTop20Gazeta() {
    const e = runAnalysisSkeleton_({
        loadInput: loadGEInputFromSheet_,
        onEmptyInput: () => SpreadsheetApp.getUi().alert('Aba "' + CFG.GE_INPUT_SHEET + '" est\xe1 vazia.'),
        processInput: processAndMapTimestampsGE_,
        ensureSheets: ensureGESheets_,
        checkProcessed: !1
    });
    if (!e) return;
    const {
        input: t,
        stop: o,
        minuteTopTerms: a
    } = e;
    writeTop20GE_(computeTop20_(t.rows, a, o)), SpreadsheetApp.getUi().alert("\u2705 Top 20 Termos atualizado!")
}

function runRelatorioEmailGazeta() {
    ensureSheetHasData_(CFG.GE_OUT_TIME, () => SpreadsheetApp.getUi().alert('\u274c A aba "' + CFG.GE_OUT_TIME + '" est\xe1 vazia. Execute a an\xe1lise primeiro.')) && (enviarEmailRelatorioGE_(gerarRelatorioGE_(null)), SpreadsheetApp.getUi().alert("\u2705 Relat\xf3rio enviado para " + CFG.EMAIL_DESTINO))
}

function abrirYouTubeGazetaEsportiva() {
    const e = HtmlService.createHtmlOutput('<script>window.open("https://www.youtube.com/@gazetaesportiva/streams", "_blank");google.script.host.close();<\/script>').setWidth(100).setHeight(50);
    SpreadsheetApp.getUi().showModalDialog(e, "Abrindo...")
}
