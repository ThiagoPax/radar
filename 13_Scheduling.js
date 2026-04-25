// =============================================================================
// 13_Scheduling.js — Agendamento de triggers e execução automática (JA + GE)
// =============================================================================

// --- Helper genérico de gerenciamento de triggers ---

function manageTriggers_(e) {
    const t = ScriptApp.getProjectTriggers(),
        o = e.handlers || [];
    if ("remove" !== e.action) {
        if ("list" === e.action) return t.filter(e => o.includes(e.getHandlerFunction()));
        if ("create" === e.action) {
            const t = [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY];
            for (const o of t) ScriptApp.newTrigger(e.handler).timeBased().onWeekDay(o).atHour(e.hour).nearMinute(e.minute).create()
        }
    } else
        for (const e of t) o.includes(e.getHandlerFunction()) && ScriptApp.deleteTrigger(e)
}

// =============================================================================
// Agendamento — Jogo Aberto (JA)
// =============================================================================

function ativarAgendamentoJA() {
    const e = SpreadsheetApp.getUi();
    manageTriggers_({
        action: "remove",
        handlers: ["executarAnaliseAgendadaJA", "executarAnaliseAgendada"]
    }), manageTriggers_({
        action: "create",
        handler: "executarAnaliseAgendadaJA",
        hour: CFG.JA_AUTO_RUN_HOUR,
        minute: CFG.JA_AUTO_RUN_MINUTE
    }), e.alert("\u2705 Agendamento Jogo Aberto ativado!", `O script ser\xe1 executado automaticamente:\n\n\u{1f4c5} Segunda a Sexta\n\u23f0 \xc0s ${CFG.JA_AUTO_RUN_HOUR}:${(CFG.JA_AUTO_RUN_MINUTE+"").padStart(2,"0")}\n\nO sistema ir\xe1:\n1. Buscar o programa do dia no YouTube\n2. Transcrever automaticamente\n3. Analisar e classificar\n4. Enviar relat\xf3rio por email`, e.ButtonSet.OK)
}

function desativarAgendamentoJA() {
    manageTriggers_({
        action: "remove",
        handlers: ["executarAnaliseAgendadaJA", "executarAnaliseAgendada"]
    }), SpreadsheetApp.getUi().alert("\u274c Agendamento Jogo Aberto desativado")
}

function removerTriggersJA_() {
    manageTriggers_({
        action: "remove",
        handlers: ["executarAnaliseAgendadaJA", "executarAnaliseAgendada"]
    })
}

function verStatusAgendamentoJA() {
    const e = SpreadsheetApp.getUi(),
        t = manageTriggers_({
            action: "list",
            handlers: ["executarAnaliseAgendadaJA", "executarAnaliseAgendada"]
        });
    0 === t.length ? e.alert("\u{1f4cb} Status Jogo Aberto", "Nenhum agendamento ativo.", e.ButtonSet.OK) : e.alert("\u{1f4cb} Status Jogo Aberto", `Agendamentos ativos: ${t.length}\nHor\xe1rio: ${CFG.JA_AUTO_RUN_HOUR}:${(CFG.JA_AUTO_RUN_MINUTE+"").padStart(2,"0")}\nDias: Segunda a Sexta`, e.ButtonSet.OK)
}

// Trigger handler JA (delega para executarAnaliseAgendada)
function executarAnaliseAgendadaJA() {
    executarAnaliseAgendada()
}

// Execução agendada JA (lógica real — chamada pelo trigger)
function executarAnaliseAgendada() {
    const e = getTranscriptionServerUrl_(),
        t = (new Date).getDay();
    if (0 !== t && 6 !== t)
        if (e) try {
            console.log("Iniciando importa\xe7\xe3o autom\xe1tica..."), console.log("URL do servidor: " + e);
            const t = chamarServidorTranscricao_(e, {
                autoFind: !0
            });
            t.success ? (salvarTranscricaoNaAbaInput_(t), console.log("Transcri\xe7\xe3o importada, iniciando an\xe1lise..."), runAllSemUI_(), console.log("An\xe1lise conclu\xedda e email enviado!")) : (console.error("Falha ao importar transcri\xe7\xe3o:", t.error), enviarEmailErro_("Falha ao importar transcri\xe7\xe3o: " + t.error + "\nURL do servidor: " + e))
        } catch (e) {
            console.error("Erro na execu\xe7\xe3o agendada:", e), enviarEmailErro_("Erro na execu\xe7\xe3o agendada: " + e.message)
        } else {
            const e = SpreadsheetApp.getActive().getSheetByName(CFG.INPUT_SHEET);
            e && e.getLastRow() > 5 ? (console.log("Transcri\xe7\xe3o manual encontrada, iniciando an\xe1lise..."), runAllSemUI_()) : console.log("Servidor n\xe3o configurado e aba Input vazia - pulando execu\xe7\xe3o")
        } else console.log("Fim de semana - pulando execu\xe7\xe3o")
}

// Execução completa JA sem UI (para uso por triggers)
function runAllSemUI_() {
    const e = loadInputFromSheet_();
    if (0 === e.rows.length) return void console.error("Aba Input vazia");
    const t = processAndMapTimestamps_(e);
    if (0 === t.rows.length) return void console.error("N\xe3o foi poss\xedvel detectar in\xedcio do Jogo Aberto");
    ensureOutputSheets_();
    const o = getStopwords_(),
        a = computeTopTermsPerMinute_(t.rows, o),
        r = buildTimeline_(t.rows, a, o);
    writeAllOutputs_(r, a, o, t.rows);
    try {
        const e = extrairDataPrograma_() || new Date;
        if (arquivarPrograma_(e)) {
            const t = Utilities.formatDate(e, "America/Sao_Paulo", "dd/MM/yyyy");
            console.log("Arquivado no Hist\xf3rico: " + t)
        } else console.log("Hist\xf3rico: programa j\xe1 arquivado")
    } catch (e) {
        console.error("Erro ao arquivar no hist\xf3rico:", e)
    }
    try {
        enviarEmailRelatorio_(gerarRelatorio_(r)), console.log("Relat\xf3rio enviado para " + CFG.EMAIL_DESTINO)
    } catch (e) {
        console.error("Erro ao enviar email:", e)
    }
}

// =============================================================================
// Agendamento — Gazeta Esportiva (GE)
// =============================================================================

function ativarAgendamentoGE() {
    const e = SpreadsheetApp.getUi();
    manageTriggers_({
        action: "remove",
        handlers: ["executarAnaliseAgendadaGE"]
    }), manageTriggers_({
        action: "create",
        handler: "executarAnaliseAgendadaGE",
        hour: CFG.GE_AUTO_RUN_HOUR,
        minute: CFG.GE_AUTO_RUN_MINUTE
    }), e.alert("\u2705 Agendamento Gazeta Esportiva ativado!", `O script ser\xe1 executado automaticamente:\n\n\u{1f4c5} Segunda a Sexta\n\u23f0 \xc0s ${CFG.GE_AUTO_RUN_HOUR}:${(CFG.GE_AUTO_RUN_MINUTE+"").padStart(2,"0")}\n\nO sistema ir\xe1:\n1. Buscar o programa do dia no YouTube\n2. Transcrever automaticamente\n3. Analisar e classificar\n4. Enviar relat\xf3rio por email`, e.ButtonSet.OK)
}

function desativarAgendamentoGE() {
    manageTriggers_({
        action: "remove",
        handlers: ["executarAnaliseAgendadaGE"]
    }), SpreadsheetApp.getUi().alert("\u274c Agendamento Gazeta Esportiva desativado")
}

function removerTriggersGE_() {
    manageTriggers_({
        action: "remove",
        handlers: ["executarAnaliseAgendadaGE"]
    })
}

function verStatusAgendamentoGE() {
    const e = SpreadsheetApp.getUi(),
        t = manageTriggers_({
            action: "list",
            handlers: ["executarAnaliseAgendadaGE"]
        });
    0 === t.length ? e.alert("\u{1f4cb} Status Gazeta Esportiva", "Nenhum agendamento ativo.", e.ButtonSet.OK) : e.alert("\u{1f4cb} Status Gazeta Esportiva", `Agendamentos ativos: ${t.length}\nHor\xe1rio: ${CFG.GE_AUTO_RUN_HOUR}:${(CFG.GE_AUTO_RUN_MINUTE+"").padStart(2,"0")}\nDias: Segunda a Sexta`, e.ButtonSet.OK)
}

// Trigger handler GE (lógica real — chamada pelo trigger)
function executarAnaliseAgendadaGE() {
    const e = getTranscriptionServerUrl_(),
        t = (new Date).getDay();
    if (0 !== t && 6 !== t) {
        if (!e) return Logger.log("Gazeta Esportiva: Servidor n\xe3o configurado."), void MailApp.sendEmail({
            to: CFG.EMAIL_DESTINO,
            subject: "\u26a0\ufe0f Erro Gazeta Esportiva - Servidor n\xe3o configurado",
            body: "O servidor de transcri\xe7\xe3o n\xe3o est\xe1 configurado. Configure em Configura\xe7\xf5es \u2192 Configurar Servidor."
        });
        try {
            const t = buscarGazetaEsportivaHoje_();
            if (!t) return Logger.log("Gazeta Esportiva: Programa n\xe3o encontrado automaticamente."), void MailApp.sendEmail({
                to: CFG.EMAIL_DESTINO,
                subject: "\u26a0\ufe0f Gazeta Esportiva - Programa n\xe3o encontrado",
                body: "N\xe3o foi poss\xedvel encontrar o Gazeta Esportiva de hoje automaticamente.\n\nImporte manualmente usando o menu."
            });
            const o = fetchJsonSafe_(e + "/transcribe", {
                method: "POST",
                contentType: "application/json",
                payload: JSON.stringify({
                    url: t
                }),
                muteHttpExceptions: !0
            });
            if (!o.success) throw Error(o.error || "Erro na transcri\xe7\xe3o");
            ensureGESheets_();
            const a = SpreadsheetApp.getActive().getSheetByName(CFG.GE_INPUT_SHEET);
            a.clear(), a.getRange(1, 1).setValue(t);
            const r = o.transcription.split("\n");
            for (let e = 0; e < r.length; e++) a.getRange(e + 2, 1).setValue(r[e]);
            const n = processAndMapTimestampsGE_(loadGEInputFromSheet_()),
                s = getStopwords_(),
                i = computeTopTermsPerMinute_(n.rows, s),
                c = buildTimelineGE_(n.rows, i, s);
            writeAllOutputsGE_(c, i, s, n.rows), enviarEmailRelatorioGE_(gerarRelatorioGE_(c)), Logger.log("Gazeta Esportiva processado com sucesso: " + o.videoTitle)
        } catch (e) {
            Logger.log("Erro no Gazeta Esportiva agendado: " + e.message), MailApp.sendEmail({
                to: CFG.EMAIL_DESTINO,
                subject: "\u26a0\ufe0f Erro Gazeta Esportiva - " + Utilities.formatDate(new Date, "America/Sao_Paulo", "dd/MM"),
                body: "Erro ao processar: " + e.message
            })
        }
    } else Logger.log("Gazeta Esportiva: Fim de semana, n\xe3o executa.")
}

// =============================================================================
// Stubs legados — delegam para funções JA (compatibilidade com menus antigos)
// =============================================================================

function ativarAgendamento() {
    ativarAgendamentoJA()
}

function desativarAgendamento() {
    desativarAgendamentoJA()
}

function verStatusAgendamento() {
    verStatusAgendamentoJA()
}

function ativarAgendamentoSpeechRadar() {
    instalarTriggersRadar_(), SpreadsheetApp.getUi().alert("✅ Agendamento Speech Radar ativado", `Importação automática configurada para ${CFG.IMPORT_TRIGGER_HOUR}:${(CFG.IMPORT_TRIGGER_MINUTE+"").padStart(2,"0")} de segunda a sexta.`, SpreadsheetApp.getUi().ButtonSet.OK)
}

function desativarAgendamentoSpeechRadar() {
    manageTriggers_({
        action: "remove",
        handlers: ["processarSpeechDoDia_"]
    }), SpreadsheetApp.getUi().alert("❌ Agendamento Speech Radar desativado")
}