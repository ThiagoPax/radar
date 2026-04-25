// =============================================================================
// 10_Historico_JA.js — Histórico de programas (Jogo Aberto)
// =============================================================================

// --- Criação / inicialização da aba Histórico ---

function ensureHistoricoSheet_() {
    const e = SpreadsheetApp.getActive();
    let t = e.getSheetByName(CFG.OUT_HISTORICO);
    return t || (t = e.insertSheet(CFG.OUT_HISTORICO), inicializarEstruturaHistorico_(t)), t
}

function inicializarEstruturaHistorico_(e) {
    e.getRange("A1").setValue(""), e.getRange("A2").setValue("HORA");
    const t = [];
    for (let e = 11; e <= 12; e++)
        for (let o = 0; o <= 59; o++) t.push([`${(e+"").padStart(2,"0")}:${(o+"").padStart(2,"0")}`]);
    e.getRange(3, 1, t.length, 1).setValues(t), e.getRange("B1").setValue(""), e.getRange("B2").setValue("M\xc9DIA"), e.getRange("C1").setValue(""), e.getRange("C2").setValue("% A M\xc9DIA"), e.setColumnWidth(1, 60), e.setColumnWidth(2, 70), e.setColumnWidth(3, 85), e.setFrozenRows(2), e.setFrozenColumns(1), e.getRange(2, 1, 1, 3).setFontWeight("bold").setBackground("#4285F4").setFontColor("white")
}

// --- Arquivar programa atual (entrada pública via menu) ---

function arquivarProgramaAtual() {
    const e = SpreadsheetApp.getUi(),
        t = SpreadsheetApp.getActive().getSheetByName(CFG.OUT_TIME);
    if (!t || t.getLastRow() < 2) return void e.alert("\u274c Erro", 'A aba "Minuto a Minuto" est\xe1 vazia. Execute a an\xe1lise completa primeiro.', e.ButtonSet.OK);
    const o = extrairDataPrograma_() || new Date,
        a = Utilities.formatDate(o, "America/Sao_Paulo", "dd/MM/yyyy");
    if (e.alert("\u{1f4e5} Arquivar Programa", `Deseja arquivar o programa de ${a} no hist\xf3rico?\n\nO programa atual ser\xe1 inserido como o mais recente (colunas D-G) e os programas anteriores ser\xe3o deslocados para a direita.`, e.ButtonSet.YES_NO) === e.Button.YES) try {
        arquivarPrograma_(o), e.alert("\u2705 Sucesso", `Programa de ${a} arquivado com sucesso!`, e.ButtonSet.OK)
    } catch (t) {
        e.alert("\u274c Erro", "Falha ao arquivar: " + t.message, e.ButtonSet.OK)
    }
}

// --- Arquivar programa (interno, chamado automaticamente pelo runAll) ---

function arquivarPrograma_(e) {
    const t = SpreadsheetApp.getActive(),
        o = ensureHistoricoSheet_(),
        a = t.getSheetByName(CFG.OUT_TIME);
    if (!a || a.getLastRow() < 2) throw Error("Aba Minuto a Minuto est\xe1 vazia");
    const r = Utilities.formatDate(e, "America/Sao_Paulo", "dd/MM/yyyy"),
        n = o.getLastColumn();
    if (n > 3) {
        const e = n - 3,
            t = o.getRange(1, 4, 1, e).getValues()[0];
        for (const e of t) {
            if (!e) continue;
            let t = "";
            if (t = e instanceof Date ? Utilities.formatDate(e, "America/Sao_Paulo", "dd/MM/yyyy") : e.toString().trim(), t === r) return Logger.log(`Programa de ${r} j\xe1 existe no hist\xf3rico - pulando`), !1
        }
    }
    o.insertColumns(4, 4);
    const s = Math.floor((n - 3) / 4) % 2 == 0 ? "#F5F5F5" : "#FFFFFF";
    o.getRange(1, 4, 1, 4).merge(), o.getRange(1, 4).setValue(r).setHorizontalAlignment("center").setFontWeight("bold").setBackground("#E8F0FE"), o.getRange(2, 4, 1, 4).setValues([
        ["AUDI\xcaNCIA", "SHARE", "TEMA", "TERMOS"]
    ]), o.getRange(2, 4, 1, 4).setFontWeight("bold").setBackground("#4285F4").setFontColor("white"), o.setColumnWidth(4, 90), o.setColumnWidth(5, 70), o.setColumnWidth(6, 120), o.setColumnWidth(7, 200), o.getRange(3, 4, 120, 4).setBackground(s);
    const i = a.getLastRow();
    if (i < 2) throw Error("Aba Minuto a Minuto n\xe3o tem dados");
    const c = Math.min(6, a.getLastColumn()),
        l = a.getRange(2, 1, i - 1, c).getValues();
    Logger.log(`Lendo ${l.length} linhas do Minuto a Minuto (${c} colunas)`);
    const d = new Map;
    for (const e of l) {
        let t = "";
        if (e[0] instanceof Date) {
            const o = e[0];
            t = `${(o.getHours()+"").padStart(2,"0")}:${(o.getMinutes()+"").padStart(2,"0")}`
        } else if (e[0]) {
            const o = e[0].toString().trim(),
                a = o.match(/(\d{1,2}):(\d{2})/);
            t = a ? `${(parseInt(a[1])+"").padStart(2,"0")}:${a[2]}` : o
        }
        const o = e[1] ? e[1].toString().trim() : "",
            a = e[4] ? e[4].toString().trim() : "";
        if (t) {
            d.set(t, {
                tema: o,
                termos: a
            });
            const e = t.replace(/^0/, "");
            e !== t && d.set(e, {
                tema: o,
                termos: a
            })
        }
    }
    Logger.log(`Mapeados ${d.size} hor\xe1rios \xfanicos`);
    const u = Array.from(d.entries()).slice(0, 5);
    for (const [e, t] of u) Logger.log(`Exemplo: "${e}" -> TEMA: "${t.tema}", TERMOS: "${t.termos.substring(0,30)}..."`);
    const m = [];
    let p = 0,
        g = 0;
    for (let e = 11; e <= 12; e++)
        for (let t = 0; t <= 59; t++) {
            const o = `${(e+"").padStart(2,"0")}:${(t+"").padStart(2,"0")}`;
            let a = "",
                r = "";
            if (11 === e && t <= 5) a = "MERCHAN", r = "comercial, intervalo", p++;
            else {
                let n = d.get(o);
                n || (n = d.get(`${e}:${(t+"").padStart(2,"0")}`)), n || (n = d.get(o + ":00")), n ? (a = n.tema, r = n.termos, p++) : g++
            }
            m.push([0, 0, a, r])
        }
    return Logger.log(`Hor\xe1rios encontrados: ${p}, n\xe3o encontrados: ${g}`), o.getRange(3, 4, m.length, 4).setValues(m), o.getRange(3, 4, m.length, 1).setNumberFormat("0,00"), o.getRange(3, 5, m.length, 1).setNumberFormat("0.0"), atualizarFormulasHistorico_(o), Logger.log(`Programa de ${r} arquivado com sucesso`), !0
}

// --- Atualizar fórmulas de média do histórico ---

function atualizarFormulasHistorico_(e) {
    const t = e.getLastColumn(),
        o = [];
    for (let e = 4; e <= t; e += 4) o.push(e);
    if (0 === o.length) return;
    const a = o.length;
    for (let t = 3; t <= 122; t++) {
        const r = o.map(e => columnToLetter_(e) + t).join("+");
        setFormulaSafe_(e.getRange(t, 2), `(${r})/${a}`), setFormulaSafe_(e.getRange(t, 3), `(D${t}/B${t})-1`)
    }
    SpreadsheetApp.flush(), e.getRange(3, 2, 120, 1).setNumberFormat("0,00"), e.getRange(3, 3, 120, 1).setNumberFormat("0.0%")
}

// --- Ver estrutura do histórico (entrada pública via menu) ---

function verEstruturaHistorico() {
    const e = SpreadsheetApp.getUi();
    e.alert("\u{1f4ca} Estrutura do Hist\xf3rico", "\n\u{1f4ca} ESTRUTURA DA PLANILHA HIST\xd3RICO\n\n\u250c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510\n\u2502 Linha 1 \u2502     \u2502       \u2502        \u2502   27/12/2025    \u2502   26/12/2025    \u2502  ... \u2502\n\u251c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2524\n\u2502 Linha 2 \u2502HORA \u2502 M\xc9DIA \u2502%A M\xc9DIA\u2502 AUD\u2502SHA\u2502TEMA\u2502TER\u2502 AUD\u2502SHA\u2502TEMA\u2502TER\u2502  ... \u2502\n\u251c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2524\n\u2502 Linha 3 \u250211:00\u2502 0.85  \u2502 +4.7%  \u25020.89\u25022.1\u2502    \u2502   \u25020.81\u25021.9\u2502    \u2502   \u2502      \u2502\n\u2502 Linha 4 \u250211:01\u2502 0.87  \u2502 +3.4%  \u25020.90\u25022.2\u2502SPFC\u2502...\u25020.84\u25022.0\u2502PALM\u2502...\u2502      \u2502\n\u2502   ...   \u2502 ... \u2502  ...  \u2502  ...   \u2502 ...\u2502...\u2502 ...\u2502...\u2502 ...\u2502...\u2502 ...\u2502...\u2502      \u2502\n\u2502Linha 122\u250212:59\u2502 1.05  \u2502 -1.9%  \u25021.03\u25022.8\u2502CORI\u2502...\u25021.07\u25022.9\u2502FLAM\u2502...\u2502      \u2502\n\u2514\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2518\n\nCOLUNAS FIXAS:\n\u2022 A: HORA (11:00 a 12:59)\n\u2022 B: M\xc9DIA de todas as audi\xeancias hist\xf3ricas daquele minuto\n\u2022 C: % A M\xc9DIA = (Audi\xeancia mais recente - M\xe9dia) / M\xe9dia\n\nBLOCOS DE 4 COLUNAS (por programa):\n\u2022 AUDI\xcaNCIA: valor em formato 0.00 (ex: 0.89)\n\u2022 SHARE: valor em formato 0.0 (ex: 2.1)\n\u2022 TEMA: tema classificado do minuto\n\u2022 TERMOS: top 5 termos do minuto\n\nCORES ALTERNADAS:\n\u2022 Programas \xedmpares (1\xba, 3\xba, 5\xba...): fundo cinza claro\n\u2022 Programas pares (2\xba, 4\xba, 6\xba...): fundo branco\n\nORDEM CRONOL\xd3GICA:\n\u2022 Programa mais recente: colunas D-G\n\u2022 Segundo mais recente: colunas H-K\n\u2022 Terceiro mais recente: colunas L-O\n\u2022 ... e assim por diante\n", e.ButtonSet.OK)
}