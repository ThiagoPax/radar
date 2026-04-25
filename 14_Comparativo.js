// =============================================================================
// 14_Comparativo.js — Análise comparativa Jogo Aberto × Gazeta Esportiva
// =============================================================================

function gerarComparativoJAxGE() {
    const e = SpreadsheetApp.getUi(),
        t = SpreadsheetApp.getActive(),
        o = t.getSheetByName(CFG.OUT_TIME),
        a = t.getSheetByName(CFG.GE_OUT_TIME);
    if (!o || o.getLastRow() < 2) return void e.alert("\u26a0\ufe0f Dados insuficientes", "Execute primeiro a an\xe1lise do Jogo Aberto.", e.ButtonSet.OK);
    if (!a || a.getLastRow() < 2) return void e.alert("\u26a0\ufe0f Dados insuficientes", "Execute primeiro a an\xe1lise do Gazeta Esportiva.", e.ButtonSet.OK);
    const r = o.getRange(2, 1, o.getLastRow() - 1, 6).getValues(),
        n = {};
    for (const e of r) {
        const t = e[1];
        t && "MERCHAN" !== t && (n[t] = (n[t] || 0) + 1)
    }
    const s = a.getRange(2, 1, a.getLastRow() - 1, 6).getValues(),
        i = {};
    for (const e of s) {
        const t = e[1];
        t && "MERCHAN" !== t && (i[t] = (i[t] || 0) + 1)
    }
    const c = r.length,
        l = s.length,
        d = [],
        u = [],
        m = [],
        p = new Set([...Object.keys(n), ...Object.keys(i)]);
    for (const e of p) {
        const t = n[e] || 0,
            o = i[e] || 0,
            a = c > 0 ? (t / c * 100).toFixed(1) : 0,
            r = l > 0 ? (o / l * 100).toFixed(1) : 0;
        t > 0 && o > 0 ? d.push({
            tema: e,
            jaMin: t,
            geMin: o,
            jaPct: a,
            gePct: r
        }) : t > 0 ? u.push({
            tema: e,
            min: t,
            pct: a
        }) : m.push({
            tema: e,
            min: o,
            pct: r
        })
    }
    d.sort((e, t) => t.jaMin + t.geMin - (e.jaMin + e.geMin)), u.sort((e, t) => t.min - e.min), m.sort((e, t) => t.min - e.min);
    const g = gerarParagrafoComparativo_(d, u, m, c, l),
        A = new Date,
        E = Utilities.formatDate(A, "America/Sao_Paulo", "dd/MM/yyyy");
    let S = `\u{1f4ca} COMPARATIVO - ${E}\n\n`;
    if (S += `Jogo Aberto: ${c} minutos | Gazeta Esportiva: ${l} minutos\n\n`, S += `\u{1f4dd} AN\xc1LISE:\n${g}\n\n`, d.length > 0) {
        S += `\u{1f504} TEMAS EM COMUM (${d.length}):\n`;
        for (const e of d.slice(0, 5)) S += `  \u2022 ${e.tema}: JA ${e.jaPct}% | GE ${e.gePct}%\n`;
        S += "\n"
    }
    if (u.length > 0) {
        S += `\u{1f535} EXCLUSIVOS JOGO ABERTO (${u.length}):\n`;
        for (const e of u.slice(0, 3)) S += `  \u2022 ${e.tema} (${e.pct}%)\n`;
        S += "\n"
    }
    if (m.length > 0) {
        S += `\u{1f534} EXCLUSIVOS GAZETA (${m.length}):\n`;
        for (const e of m.slice(0, 3)) S += `  \u2022 ${e.tema} (${e.pct}%)\n`
    }
    e.alert("\u{1f4ca} Comparativo JA x GE", S, e.ButtonSet.OK), PropertiesService.getScriptProperties().setProperty("ULTIMO_COMPARATIVO", JSON.stringify({
        data: E,
        paragrafo: g,
        temasComuns: d,
        temasExclusivosJA: u,
        temasExclusivosGE: m,
        jaTotalMin: c,
        geTotalMin: l
    }))
}

function gerarParagrafoComparativo_(e, t, o, a, r) {
    if (!hasOpenAIKey_()) {
        let a = "";
        return e.length > 0 && (a += `Ambos os programas abordaram ${e.slice(0,3).map(e=>e.tema).join(", ")}. `), t.length > 0 && (a += `O Jogo Aberto deu destaque exclusivo para ${t[0].tema}. `), o.length > 0 && (a += `O Gazeta Esportiva focou exclusivamente em ${o[0].tema}. `), a || "N\xe3o foi poss\xedvel gerar an\xe1lise comparativa."
    }
    const n = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    let s = "Temas em comum: ";
    e.length > 0 ? s += e.slice(0, 5).map(e => `${e.tema} (JA: ${e.jaPct}%, GE: ${e.gePct}%)`).join(", ") : s += "nenhum", s += ". Exclusivos Jogo Aberto: ", t.length > 0 ? s += t.slice(0, 3).map(e => `${e.tema} (${e.pct}%)`).join(", ") : s += "nenhum", s += ". Exclusivos Gazeta: ", o.length > 0 ? s += o.slice(0, 3).map(e => `${e.tema} (${e.pct}%)`).join(", ") : s += "nenhum";
    const i = `Compare os programas esportivos Jogo Aberto (2h, ${a} min analisados) e Gazeta Esportiva (1h, ${r} min analisados) com base nos dados abaixo. Escreva UM \xdaNICO PAR\xc1GRAFO conciso (3-4 frases) destacando: 1) O que ambos priorizaram, 2) Diferen\xe7as de foco entre eles, considerando a propor\xe7\xe3o de tempo de cada programa.\n\nDADOS: ${s}\n\nREGRAS:\n- Use linguagem jornal\xedstica e objetiva\n- Considere que Jogo Aberto tem o dobro do tempo\n- N\xe3o use bullet points, apenas texto corrido\n- M\xe1ximo 100 palavras`;
    try {
        const e = fetchJsonSafe_(CFG.OPENAI_ENDPOINT, {
            method: "POST",
            headers: {
                Authorization: "Bearer " + n,
                "Content-Type": "application/json"
            },
            payload: JSON.stringify({
                model: CFG.OPENAI_MODEL,
                messages: [{
                    role: "user",
                    content: i
                }],
                temperature: .3,
                max_tokens: 300
            }),
            muteHttpExceptions: !0
        });
        if (e.choices && e.choices[0] && e.choices[0].message) return e.choices[0].message.content.trim()
    } catch (e) {
        Logger.log("Erro ao gerar comparativo com IA: " + e.message)
    }
    return `Os programas compartilharam foco em ${e.length>0?e[0].tema:"temas gerais"}. Proporcionalmente, considerando que o Jogo Aberto tem o dobro de dura\xe7\xe3o, ambos dedicaram tempo similar aos principais clubes do dia.`
}

function enviarComparativoEmail() {
    const e = SpreadsheetApp.getUi(),
        t = PropertiesService.getScriptProperties().getProperty("ULTIMO_COMPARATIVO");
    if (!t) return void e.alert("\u26a0\ufe0f Nenhum comparativo", "Gere primeiro um comparativo usando o menu.", e.ButtonSet.OK);
    const o = JSON.parse(t),
        a = "\u{1f4ca} Comparativo JA x GE - " + o.data;
    let r = `\n    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">\n      <div style="background: linear-gradient(135deg, #1a73e8, #e74c3c); color: white; padding: 20px; text-align: center;">\n        <h1 style="margin: 0;">\u{1f4ca} Comparativo</h1>\n        <p style="margin: 5px 0 0 0;">Jogo Aberto x Gazeta Esportiva | ${o.data}</p>\n      </div>\n      \n      <div style="padding: 20px;">\n        <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px;">\n          <p style="margin: 0; line-height: 1.6;">${o.paragrafo}</p>\n        </div>\n        \n        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">\n          <tr>\n            <td style="width: 50%; padding: 10px; background: #1a73e8; color: white; text-align: center;">\n              <strong>Jogo Aberto</strong><br>${o.jaTotalMin} min\n            </td>\n            <td style="width: 50%; padding: 10px; background: #e74c3c; color: white; text-align: center;">\n              <strong>Gazeta Esportiva</strong><br>${o.geTotalMin} min\n            </td>\n          </tr>\n        </table>\n  `;
    if (o.temasComuns && o.temasComuns.length > 0) {
        r += '\n        <h3 style="color: #333;">\u{1f504} Temas em Comum</h3>\n        <table style="width: 100%; border-collapse: collapse;">\n          <tr style="background: #f5f5f5;">\n            <th style="padding: 8px; text-align: left;">Tema</th>\n            <th style="padding: 8px; text-align: center;">JA</th>\n            <th style="padding: 8px; text-align: center;">GE</th>\n          </tr>\n    ';
        for (const e of o.temasComuns.slice(0, 5)) r += `\n          <tr>\n            <td style="padding: 8px; border-bottom: 1px solid #ddd;">${e.tema}</td>\n            <td style="padding: 8px; text-align: center; border-bottom: 1px solid #ddd;">${e.jaPct}%</td>\n            <td style="padding: 8px; text-align: center; border-bottom: 1px solid #ddd;">${e.gePct}%</td>\n          </tr>\n      `;
        r += "</table>"
    }
    r += '\n      </div>\n      \n      <div style="background: #f5f5f5; padding: 15px; text-align: center; font-size: 12px; color: #666;">\n        Relat\xf3rio comparativo gerado automaticamente\n      </div>\n    </div>\n  ', MailApp.sendEmail({
        to: CFG.EMAIL_DESTINO,
        subject: a,
        htmlBody: r
    }), e.alert("\u2705 Comparativo enviado para " + CFG.EMAIL_DESTINO)
}