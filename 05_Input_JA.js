// =============================================================================
// 05_Input_JA.js — Carregamento e mapeamento de timestamps do Jogo Aberto
// =============================================================================

// --- Carregamento da aba Input ---

function loadInputFromSheet_() {
    const e = SpreadsheetApp.getActive().getSheetByName(CFG.INPUT_SHEET);
    if (!e) throw Error(`Aba "${CFG.INPUT_SHEET}" não encontrada.`);
    const t = e.getLastRow(),
        o = e.getLastColumn();
    if (t < 1 || o < 1) return {
        rows: []
    };
    const a = e.getRange(1, 1, t, Math.min(o, 10)).getValues(),
        r = [],
        n = /^\((\d{1,2}):(\d{2})(?::(\d{2}))?\)\s*/;
    for (let e = 0; e < a.length; e++) {
        const t = a[e][0];
        if (!t) continue;
        const o = t.toString().trim();
        if (!o) continue;
        const s = o.match(n);
        if (s) {
            let e;
            e = s[3] ? 60 * parseInt(s[1], 10) + parseInt(s[2], 10) : parseInt(s[1], 10);
            const t = o.replace(n, "").trim();
            t && r.push({
                originalMinute: e,
                originalTimestamp: s[0].replace(/[()]/g, "").trim(),
                texto: t,
                horaStr: null
            })
        } else {
            if (o.startsWith("http") || o.includes("youtube.com") || "Transcript:" === o) continue;
            r.length > 0 && (r[r.length - 1].texto += " " + o)
        }
    }
    return {
        rows: r
    }
}

// --- Processamento e mapeamento de timestamps ---

function processAndMapTimestamps_(e) {
    const t = e.rows;
    if (0 === t.length) return {
        rows: [],
        jogoAbertoStartIndex: -1,
        scaleFactor: 1
    };
    const o = detectJogoAbertoStart_(t);
    return o < 0 ? (Logger.log("AVISO: Não detectou início do Jogo Aberto, usando minuto 10 como fallback"), processFromIndex_(t, Math.min(10, t.length - 1))) : processFromIndex_(t, o)
}

function detectJogoAbertoStart_(e) {
    const t = {
        jogoAbertoStart: ["bom dia", "jogo aberto", "renata fan", "renata fã", "continuo na elite", "estou na serie a", "estou na série a", "muito bom dia"],
        bandBetIndicators: ["band placar", "bandplacar", "oferecimento band bet", "oferecimento bandbet", "melhores oportunidades do mercado", "qr code", "qrcode", "aponte sua camera"]
    };
    let o = -1,
        a = !1,
        r = -1;
    for (let n = 0; n < e.length; n++) {
        const s = normalizeText_(e[n].texto);
        if (t.bandBetIndicators.some(e => s.includes(normalizeText_(e))) && (o = n), o >= 0 && !a) {
            const i = t.jogoAbertoStart.some(e => s.includes(normalizeText_(e))),
                c = containsAny_(s, ["inter", "colorado", "flamengo", "palmeiras", "corinthians", "serie a", "série a", "rebaixamento", "campeonato"]);
            if (i && (c || n > o) && (e[n].originalMinute >= 8 || n >= 8)) {
                a = !0, r = n;
                break
            }
        }
    }
    if (r < 0)
        for (let t = 0; t < e.length; t++)
            if (!(e[t].originalMinute < 10 && t < 10) && containsAny_(normalizeText_(e[t].texto), ["bom dia", "muito bom dia", "jogo aberto"])) {
                r = t;
                break
            } return r
}

function processFromIndex_(e, t) {
    const o = e.slice(t);
    if (0 === o.length) return {
        rows: [],
        jogoAbertoStartIndex: t,
        scaleFactor: 1
    };
    const a = 60 * CFG.PROGRAM_START_HOUR + CFG.PROGRAM_START_MINUTE,
        r = 60 * CFG.PROGRAM_END_HOUR + CFG.PROGRAM_END_MINUTE,
        n = r - a + 1,
        s = detectarDivisaoRede_(o);
    let i;
    return s >= 0 ? (Logger.log("Divisão de rede detectada no índice " + s), i = mapearComAncoraRede_(o, s, a, r)) : (Logger.log("Divisão de rede não detectada, usando mapeamento linear"), i = mapearLinear_(o, a, r)), {
        rows: fillMissingMinutes_(i, a, r),
        jogoAbertoStartIndex: t,
        scaleFactor: n / o.length,
        usouAncoraRede: s >= 0
    }
}

function detectarDivisaoRede_(e) {
    const t = ["divisao de rede", "divisão de rede", "dividindo a rede", "divide a rede", "divisao da rede", "divisão da rede", "12 horas", "meio dia", "meio-dia", "doze horas"],
        o = Math.floor(.35 * e.length),
        a = Math.floor(.65 * e.length);
    for (let r = o; r <= a && r < e.length; r++) {
        const o = normalizeText_(e[r].texto);
        for (const e of t)
            if (o.includes(normalizeText_(e))) return Logger.log(`Divisão de rede encontrada: "${e}" no índice ${r}`), r
    }
    return -1
}

function mapearComAncoraRede_(e, t, o, a) {
    const r = [],
        n = o,
        s = 719 - n + 1,
        i = e.slice(0, t + 1),
        c = s / i.length;
    Logger.log(`Bloco 1: ${i.length} linhas → ${s} minutos (escala: ${c.toFixed(3)})`);
    for (let e = 0; e < i.length; e++) {
        const t = Math.min(n + Math.round(e * c), 719),
            a = t % 60,
            s = `${(Math.floor(t/60)+"").padStart(2,"0")}:${(a+"").padStart(2,"0")}`;
        r.push({
            horaStr: s,
            horaValue: s,
            texto: i[e].texto,
            originalMinute: i[e].originalMinute,
            originalTimestamp: i[e].originalTimestamp,
            mappedMinuteOffset: t - o,
            bloco: 1
        })
    }
    const l = a,
        d = l - 720 + 1,
        u = e.slice(t + 1);
    if (u.length > 0) {
        const e = d / u.length;
        Logger.log(`Bloco 2: ${u.length} linhas → ${d} minutos (escala: ${e.toFixed(3)})`);
        for (let t = 0; t < u.length; t++) {
            const a = Math.min(720 + Math.round(t * e), l),
                n = a % 60,
                s = `${(Math.floor(a/60)+"").padStart(2,"0")}:${(n+"").padStart(2,"0")}`;
            r.push({
                horaStr: s,
                horaValue: s,
                texto: u[t].texto,
                originalMinute: u[t].originalMinute,
                originalTimestamp: u[t].originalTimestamp,
                mappedMinuteOffset: a - o,
                bloco: 2
            })
        }
    }
    return r
}

function mapearLinear_(e, t, o) {
    const a = o - t + 1,
        r = a / e.length;
    Logger.log(`Mapeamento linear: ${e.length} min → ${a} min (escala: ${r.toFixed(3)})`);
    const n = [];
    for (let o = 0; o < e.length; o++) {
        const a = Math.round(o * r),
            s = t + a,
            i = Math.floor(s / 60),
            c = s % 60,
            l = Math.min(i, CFG.PROGRAM_END_HOUR),
            d = i > CFG.PROGRAM_END_HOUR ? CFG.PROGRAM_END_MINUTE : c,
            u = `${(l+"").padStart(2,"0")}:${(d+"").padStart(2,"0")}`;
        n.push({
            horaStr: u,
            horaValue: u,
            texto: e[o].texto,
            originalMinute: e[o].originalMinute,
            originalTimestamp: e[o].originalTimestamp,
            mappedMinuteOffset: a
        })
    }
    return n
}

function fillMissingMinutes_(e, t, o) {
    if (0 === e.length) return [];
    const a = new Map;
    for (const t of e) {
        const [e, o] = t.horaStr.split(":").map(Number), r = 60 * e + o;
        a.has(r) ? a.get(r).texto += " " + t.texto : a.set(r, t)
    }
    const r = [];
    for (let e = t; e <= o; e++) {
        const o = e % 60,
            n = `${(Math.floor(e/60)+"").padStart(2,"0")}:${(o+"").padStart(2,"0")}`;
        a.has(e) ? r.push(a.get(e)) : r.push({
            horaStr: n,
            horaValue: n,
            texto: "[MERCHAN - conteúdo suprimido]",
            originalMinute: null,
            originalTimestamp: null,
            mappedMinuteOffset: e - t,
            isSuppressedMerchan: !0
        })
    }
    return r
}