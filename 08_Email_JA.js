// =============================================================================
// 08_Email_JA.js — Geração de relatório e envio de email (Jogo Aberto)
// =============================================================================

// --- Extração de data ---

function extrairDataPrograma_() {
    const e = SpreadsheetApp.getActive().getSheetByName(CFG.INPUT_SHEET);
    if (!e || e.getLastRow() < 1) return null;
    const t = e.getRange(1, 1, Math.min(e.getLastRow(), 20), 1).getValues().map(e => e[0]).join(" "),
        o = [/(\d{1,2})\/(\d{1,2})\/(\d{4})/, /(\d{1,2})-(\d{1,2})-(\d{4})/, /(\d{1,2})\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*\s+(\d{4})/i];
    for (const e of o) {
        const o = t.match(e);
        if (o) {
            let t, a, r;
            return e.source.includes("jan|fev") ? (t = parseInt(o[1]), a = {
                jan: 1,
                fev: 2,
                mar: 3,
                abr: 4,
                mai: 5,
                jun: 6,
                jul: 7,
                ago: 8,
                set: 9,
                out: 10,
                nov: 11,
                dez: 12
            } [o[2].toLowerCase().substring(0, 3)] || 1, r = parseInt(o[3])) : (t = parseInt(o[1]), a = parseInt(o[2]), r = parseInt(o[3])), new Date(r, a - 1, t)
        }
    }
    const a = new Date,
        r = a.getDay();
    return 0 === r ? a.setDate(a.getDate() - 2) : 6 === r && a.setDate(a.getDate() - 1), a
}

// --- Geração do relatório ---

function gerarRelatorio_(e) {
    const t = SpreadsheetApp.getActive(),
        o = extrairDataPrograma_() || new Date,
        a = Utilities.formatDate(o, "America/Sao_Paulo", "dd/MM/yyyy"),
        r = ["Domingo", "Segunda-feira", "Ter\xe7a-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "S\xe1bado"][o.getDay()],
        n = t.getSheetByName(CFG.OUT_MACRO),
        s = t.getSheetByName(CFG.OUT_SUB),
        i = t.getSheetByName(CFG.OUT_TOP),
        c = t.getSheetByName(CFG.OUT_TIME),
        l = [];
    if (n && n.getLastRow() > 1) {
        const e = n.getRange(2, 1, n.getLastRow() - 1, 5).getValues();
        for (const t of e) t[0] && "Merchan" !== t[0] && "" !== t[0] && l.length < 5 && l.push({
            tema: t[0],
            minutos: t[1],
            percent: t[2],
            primeira: t[3],
            ultima: t[4]
        })
    }
    const d = [];
    if (s && s.getLastRow() > 1) {
        const e = s.getRange(2, 1, Math.min(s.getLastRow() - 1, 10), 6).getValues();
        for (const t of e) t[0] && "" !== t[0] && d.push({
            tema: t[0],
            subtema: t[1],
            minutos: t[2],
            percent: t[3]
        })
    }
    const u = [];
    if (i && i.getLastRow() > 1) {
        const e = i.getRange(2, 1, Math.min(i.getLastRow() - 1, 5), 3).getValues();
        for (const t of e) t[0] && "" !== t[0] && u.push({
            termo: t[0],
            contagem: t[1],
            minutos: t[2]
        })
    }
    let m = [],
        p = [];
    if (c && c.getLastRow() > 1) {
        const e = c.getRange(2, 1, Math.min(c.getLastRow() - 1, 114), 6).getValues();
        for (const t of e)
            if (t[1] && "MERCHAN" !== t[1] && t[5]) {
                const e = {
                    hora: t[0],
                    tema: t[1],
                    subtema: t[2] || "",
                    trecho: t[5].toString().substring(0, 200)
                };
                m.push(e);
                let o = -1;
                if (t[0] instanceof Date) o = t[0].getHours();
                else if ("string" == typeof t[0]) {
                    const e = t[0].match(/^(\d{1,2}):/);
                    e && (o = parseInt(e[1]))
                } else "number" == typeof t[0] && (o = Math.floor(24 * t[0]));
                12 === o && p.push(e)
            }
    }
    return console.log(`Trechos para resumo: ${m.length}, Trechos do debate (12h): ${p.length}`), {
        data: a,
        diaSemana: r,
        resumo: gerarResumoComIA_(l, d, m, a, u),
        macrotemas: l,
        subtemas: d,
        topTermos: u,
        debate: gerarDadosDebate_(p, a, u)
    }
}

// --- Análise do debate (12h) ---

function gerarDadosDebate_(e, t, o) {
    if (!e || 0 === e.length) return null;
    const a = {},
        r = {},
        n = {};
    for (const t of e) {
        if (t.tema && "MERCHAN" !== t.tema && (a[t.tema] = (a[t.tema] || 0) + 1), t.tema && t.subtema) {
            const e = `${t.tema}|${t.subtema}`;
            r[e] = (r[e] || 0) + 1
        }
        if (t.trecho) {
            const e = t.trecho.toUpperCase().split(/\s+/);
            for (const t of e) {
                const e = t.replace(/[^A-Z\xc1\xc9\xcd\xd3\xda\xc2\xca\xce\xd4\xdb\xc3\xd5\xc7]/g, "");
                e.length > 3 && (n[e] = (n[e] || 0) + 1)
            }
        }
    }
    const s = Object.entries(a).sort((e, t) => t[1] - e[1]).slice(0, 5).map(([t, o]) => ({
            tema: t,
            minutos: o,
            percent: Math.round(o / e.length * 100) + "%"
        })),
        i = Object.entries(r).sort((e, t) => t[1] - e[1]).slice(0, 5).map(([t, o]) => {
            const [a, r] = t.split("|");
            return {
                tema: a,
                subtema: r,
                minutos: o,
                percent: Math.round(o / e.length * 100) + "%"
            }
        }),
        c = new Set(["A", "AI", "AINDA", "ALEM", "ALGUM", "ALGUMA", "ALGUMAS", "ALGUNS", "ALI", "AMBOS", "ANO", "ANOS", "ANTES", "AO", "AONDE", "APENAS", "APOS", "AQUELA", "AQUELAS", "AQUELE", "AQUELES", "AQUI", "AQUILO", "AS", "ASSIM", "ATE", "ATRAVES", "BEM", "BOA", "BOAS", "BOM", "BONS", "CADA", "CARA", "CASA", "CERTO", "COISA", "COISAS", "COM", "COMO", "CONTRA", "DA", "DAQUI", "DAS", "DE", "DELA", "DELAS", "DELE", "DELES", "DEMAIS", "DEPOIS", "DESDE", "DESSA", "DESSE", "DESTA", "DESTE", "DEUS", "DEZ", "DIZ", "DIZER", "DO", "DOIS", "DOS", "DUAS", "E", "EH", "ELA", "ELAS", "ELE", "ELES", "EM", "EMBORA", "ENQUANTO", "ENTAO", "ENTRE", "ERA", "ERAM", "ERAMOS", "ESSA", "ESSAS", "ESSE", "ESSES", "ESTA", "ESTAVA", "ESTAVAM", "ESTAVAMOS", "ESTE", "ESTEJA", "ESTEJAM", "ESTEJAMOS", "ESTES", "ESTEVE", "ESTIVE", "ESTIVEMOS", "ESTIVER", "ESTIVERA", "ESTIVERAM", "ESTIVERAMOS", "ESTIVEREM", "ESTIVERMOS", "ESTIVESSE", "ESTIVESSEM", "ESTIVESSEMOS", "ESTOU", "EU", "FALA", "FALANDO", "FALAR", "FAZ", "FAZENDO", "FAZER", "FEITO", "FEZ", "FICA", "FICAM", "FICAR", "FICOU", "FIM", "FOI", "FOMOS", "FOR", "FORA", "FORAM", "FORMA", "FOSSE", "FUI", "GENTE", "GRANDE", "GRANDES", "HA", "HAVER", "HAJA", "HAJAM", "HAJAMOS", "HAO", "HAVEMOS", "HAVIA", "HOJE", "HORA", "HOUVE", "HOUVEMOS", "ISSO", "ISTO", "JA", "LA", "LHE", "LHES", "LO", "LOGO", "MAIS", "MAS", "ME", "MELHOR", "MENOS", "MESMA", "MESMAS", "MESMO", "MESMOS", "MEU", "MEUS", "MINHA", "MINHAS", "MUITA", "MUITAS", "MUITO", "MUITOS", "NA", "NAO", "NAS", "NEM", "NENHUM", "NENHUMA", "NISSO", "NO", "NOS", "NOSSA", "NOSSAS", "NOSSO", "NOSSOS", "NUM", "NUMA", "NUNCA", "O", "OI", "OLHA", "ONDE", "OS", "OU", "OUTRA", "OUTRAS", "OUTRO", "OUTROS", "PARA", "PARTE", "PELA", "PELAS", "PELO", "PELOS", "PEQUENO", "PODE", "PODEMOS", "PODER", "POIS", "POR", "PORQUE", "POUCO", "PRA", "PRAS", "PRO", "PROS", "PRIMEIRO", "PROPRIA", "PROPRIO", "QUAL", "QUANDO", "QUANTO", "QUE", "QUEM", "QUER", "QUERO", "SABE", "SE", "SEJA", "SEJAM", "SEJAMOS", "SEM", "SEMPRE", "SENDO", "SER", "SERA", "SERAO", "SEREMOS", "SERIA", "SERIAM", "SERIAMOS", "SEU", "SEUS", "SI", "SIDO", "SIM", "SO", "SOB", "SOBRE", "SOMOS", "SOU", "SUA", "SUAS", "TA", "TAL", "TALVEZ", "TAMBEM", "TAO", "TAVA", "TE", "TEM", "TEMOS", "TENDO", "TENHO", "TER", "TERA", "TERAO", "TEREMOS", "TERIA", "TERIAM", "TERIAMOS", "TEUS", "TEVE", "TI", "TIDO", "TINHA", "TINHAM", "TINHAMOS", "TIPO", "TIVE", "TIVEMOS", "TIVER", "TIVERA", "TIVERAM", "TIVERAMOS", "TIVEREM", "TIVERMOS", "TIVESSE", "TIVESSEM", "TIVESSEMOS", "TO", "TODA", "TODAS", "TODO", "TODOS", "TUA", "TUAS", "TUDO", "TU", "UM", "UMA", "UMAS", "UNS", "VAI", "VAMOS", "VEM", "VER", "VERA", "VEZ", "VEZES", "VOCE", "VOCES", "VOU", "JOGO", "ABERTO", "RENATA", "FAN", "PROGRAMA", "BAND", "ONTEM", "AGORA", "ENTENDEU", "VEJA", "GENTE", "PESSOAL", "GALERA", "APLAUSOS", "RISOS", "INAUDIVEL", "INAUDIVEL", "MUSICA", "MUSICA"]),
        l = e => e ? e.toLowerCase().split(/[\s-]+/).map(e => e.charAt(0).toUpperCase() + e.slice(1)).join(" ") : e,
        d = Object.entries(n).filter(([e]) => !c.has((e => e.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""))(e))).sort((e, t) => t[1] - e[1]).slice(0, 5).map(([e, t]) => ({
            termo: l(e),
            contagem: t,
            minutos: Math.ceil(t / 3)
        })),
        u = s.map(e => ({
            ...e,
            tema: l(e.tema)
        })),
        m = i.map(e => ({
            ...e,
            tema: l(e.tema),
            subtema: l(e.subtema)
        }));
    return {
        resumo: gerarResumoDebate_(e, u, t),
        topTemas: u,
        topSubtemas: m,
        topTermos: d,
        totalMinutos: e.length
    }
}

// --- Resumo do debate via IA ---

function gerarResumoDebate_(e, t, o) {
    if (!hasOpenAIKey_() || 0 === e.length) return ["O debate das 12h abordou diversos temas esportivos."];
    const a = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"),
        r = t.map(e => e.tema).join(", "),
        n = `Voc\xea \xe9 um analista esportivo. Resuma em 3 frases o DEBATE do programa Jogo Aberto (12h-13h) de ${o}.\n\nTEMAS MAIS DISCUTIDOS NO DEBATE: ${r}\n\nTRECHOS DO DEBATE:\n${e.slice(0,10).map(e=>`[${e.tema}] ${e.trecho}`).join(" | ")}\n\nREGRAS:\n1. Foque nos debates e opini\xf5es dos comentaristas\n2. Use frases como "Os comentaristas debateram...", "Houve discuss\xe3o sobre..."\n3. Seja espec\xedfico sobre os pontos de vista apresentados\n4. M\xe1ximo 3 frases curtas\n\nResponda APENAS com as 3 frases, nada mais.`;
    try {
        const e = fetchJsonSafe_(CFG.OPENAI_ENDPOINT, {
            method: "post",
            contentType: "application/json",
            headers: {
                Authorization: "Bearer " + a
            },
            payload: JSON.stringify({
                model: CFG.OPENAI_MODEL,
                messages: [{
                    role: "user",
                    content: n
                }],
                max_tokens: 300,
                temperature: .5
            }),
            muteHttpExceptions: !0
        });
        if (e.choices && e.choices[0] && e.choices[0].message) return e.choices[0].message.content.trim().split("\n").filter(e => e.trim().length > 0).slice(0, 3)
    } catch (e) {
        console.error("Erro ao gerar resumo do debate:", e)
    }
    return ["O debate das 12h abordou " + r + "."]
}

// --- Resumo do programa via IA ---

function gerarResumoComIA_(e, t, o, a, r) {
    if (!hasOpenAIKey_()) return gerarResumoSimples_(e);
    const n = e.map(e => `${e.tema}: ${e.minutos} minutos (${e.percent})`).join("; "),
        s = t.slice(0, 5).map(e => `${e.tema}/${e.subtema}: ${e.minutos}min`).join("; "),
        i = r ? r.slice(0, 3).map(e => e.termo) : [],
        c = [],
        l = new Set;
    for (const e of o) !l.has(e.tema) && c.length < 7 && (c.push(`[${e.tema}] ${e.trecho}`), l.add(e.tema));
    const d = o.map(e => e.trecho).join(" "),
        u = `Voc\xea \xe9 um analista esportivo brasileiro. Com base nos dados abaixo do programa "Jogo Aberto" (Band) de ${a}, escreva um resumo de EXATAMENTE 5 frases curtas e objetivas.\n\nDADOS DO PROGRAMA:\n- Macrotemas: ${n}\n- Subtemas principais: ${s}\n- Exemplos de discuss\xf5es: ${c.join(" | ")}\n\nREGRAS:\n1. Descreva os TEMAS DISCUTIDOS de forma objetiva\n2. Use frases como "Foi discutido...", "O programa abordou...", "Houve debate sobre..."\n3. N\xc3O afirme resultados de jogos, contrata\xe7\xf5es ou fatos espec\xedficos ainda\n4. Mencione os times pelo nome\n5. Use verbos no passado\n6. Cada frase deve ter no m\xe1ximo 25 palavras\n7. N\xc3O use bullet points, apenas 5 frases separadas por quebra de linha\n\nResponda APENAS com as 5 frases, nada mais.`;
    try {
        const t = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"),
            o = fetchJsonSafe_(CFG.OPENAI_ENDPOINT, {
                method: "post",
                contentType: "application/json",
                headers: {
                    Authorization: "Bearer " + t
                },
                payload: JSON.stringify({
                    model: CFG.OPENAI_MODEL,
                    messages: [{
                        role: "user",
                        content: u
                    }],
                    max_tokens: 500,
                    temperature: .5
                }),
                muteHttpExceptions: !0
            });
        if (!o.choices || !o.choices[0] || !o.choices[0].message) return gerarResumoSimples_(e);
        const r = o.choices[0].message.content.trim();
        Logger.log("Etapa 1 - Resumo base gerado");
        return enriquecerResumoComFatos_(r, a, c, d, t, e.slice(0, 10).map(e => e.tema), i).split("\n").filter(e => e.trim().length > 0).slice(0, 5)
    } catch (e) {
        console.error("Erro ao gerar resumo com IA:", e)
    }
    return gerarResumoSimples_(e)
}

// --- Enriquecimento do resumo com nomes específicos ---

function enriquecerResumoComFatos_(e, t, o, a, r, n, s) {
    const i = `Analise a transcri\xe7\xe3o de um programa esportivo e extraia TODOS os nomes mencionados.\n\nTRANSCRI\xc7\xc3O:\n${a.substring(0,8e3)}\n\nEXTRAIA DA TRANSCRI\xc7\xc3O:\n1. Nomes de JOGADORES de futebol (com clube e contexto do que foi dito)\n2. Nomes de T\xc9CNICOS/TREINADORES (com clube e contexto)\n3. TRANSFER\xcaNCIAS mencionadas na transcri\xe7\xe3o (jogador, clubes envolvidos, valores se citados)\n4. LES\xd5ES/RECUPERA\xc7\xd5ES mencionadas (jogador, time, situa\xe7\xe3o)\n5. Ex-jogadores famosos mencionados\n\nN\xc3O EXTRAIA:\n- Apresentadores (Renata Fan, Den\xedlson, Chico Garcia, Ronaldo Giovanelli, etc.)\n- Comentaristas e jornalistas\n- Informa\xe7\xf5es que N\xc3O est\xe3o na transcri\xe7\xe3o\n\nFORMATO:\nJOGADOR: [Nome] | CLUBE: [Clube] | CONTEXTO: [resumo do que foi dito na transcri\xe7\xe3o]\nT\xc9CNICO: [Nome] | CLUBE: [Clube] | CONTEXTO: [resumo do que foi dito]\nTRANSFER\xcaNCIA: [Jogador] | DE: [Clube origem] | PARA: [Clube destino] | VALOR: [se mencionado]\n\nIMPORTANTE: \n- Extraia EXATAMENTE como aparece na transcri\xe7\xe3o (mesmo com erros de \xe1udio)\n- N\xc3O invente informa\xe7\xf5es que n\xe3o est\xe3o no texto\n- Extraia APENAS o que foi realmente mencionado`;
    let c = "";
    try {
        const e = fetchJsonSafe_(CFG.OPENAI_ENDPOINT, {
            method: "post",
            contentType: "application/json",
            headers: {
                Authorization: "Bearer " + r
            },
            payload: JSON.stringify({
                model: CFG.OPENAI_MODEL,
                messages: [{
                    role: "user",
                    content: i
                }],
                max_tokens: 1200,
                temperature: .1
            }),
            muteHttpExceptions: !0
        });
        e.choices && e.choices[0] && e.choices[0].message && (c = e.choices[0].message.content.trim(), Logger.log("=== ETAPA 2A: NOMES EXTRA\xcdDOS (GPT) ==="), Logger.log(c))
    } catch (t) {
        return Logger.log("Erro na extra\xe7\xe3o de nomes: " + t), e
    }
    if (!c) return e;
    let l = [];
    if (n && n.length > 0) {
        Logger.log("Temas recebidos para busca: " + n.join(", "));
        for (const e of n) {
            const t = normalizarNomeClube_(e.toUpperCase());
            t && !l.includes(t) && (l.push(t), Logger.log("Adicionado dos temas: " + t))
        }
    }
    const d = extrairClubesMencionados_(c);
    for (const e of d) l.includes(e) || l.push(e);
    Logger.log("Clubes para buscar elencos: " + l.join(", "));
    let u = {};
    const m = PropertiesService.getScriptProperties().getProperty("API_FOOTBALL_KEY");
    if (m && l.length > 0) {
        u = buscarElencosApiFootball_(l, m), Logger.log("=== ETAPA 2B: ELENCOS API-FOOTBALL ===");
        for (const e in u) Logger.log(e + ": " + u[e].slice(0, 10).join(", ") + "...")
    } else Logger.log("API-Football n\xe3o configurada ou sem clubes - pulando busca de elencos");
    let p = corrigirNomesComGemini_(c, u, t, a);
    p || (Logger.log("Gemini n\xe3o dispon\xedvel, usando nomes originais"), p = c);
    const g = n ? n.slice(0, 3).join(", ") : "",
        A = s ? s.join(", ") : "";
    Logger.log("Top 3 Termos para incluir no resumo: " + A);
    const E = `Reescreva este resumo incluindo NOMES ESPEC\xcdFICOS de jogadores e t\xe9cnicos.\n\nDATA DO PROGRAMA: ${t}\n\nTOP 3 TEMAS MAIS DISCUTIDOS: ${g}\nTOP 3 TERMOS MAIS CITADOS: ${A}\n\n\u26a0\ufe0f VERIFICA\xc7\xc3O OBRIGAT\xd3RIA:\n1. Cada um dos TOP 3 TEMAS deve aparecer em pelo menos 1 frase\n2. Cada um dos TOP 3 TERMOS deve estar contextualizado no resumo\n   - Se "Corinthians" \xe9 o termo mais citado, DEVE haver uma frase explicando POR QUE falaram tanto dele\n   - Se "Cruzeiro" aparece nos termos, contextualizar o motivo da discuss\xe3o\n\nRESUMO ATUAL (muito gen\xe9rico, precisa de nomes):\n${e}\n\nNOMES CORRIGIDOS E VALIDADOS (USE ESTES):\n${p}\n\nTRECHOS DA TRANSCRI\xc7\xc3O PARA CONTEXTO:\n${o.slice(0,5).join("\n")}\n\nREGRAS OBRIGAT\xd3RIAS:\n\n1. USE APENAS INFORMA\xc7\xd5ES DA TRANSCRI\xc7\xc3O E DOS NOMES CORRIGIDOS:\n   - N\xc3O invente transfer\xeancias, resultados ou fatos que n\xe3o foram mencionados\n   - Baseie-se APENAS no que est\xe1 nos trechos e nos nomes extra\xeddos/corrigidos\n\n2. PARA CADA TEMA DO TOP 3, inclua uma frase COM NOME ESPEC\xcdFICO:\n   - Use os nomes que aparecem na lista "NOMES CORRIGIDOS E VALIDADOS"\n   - Se n\xe3o h\xe1 nomes espec\xedficos para um tema, descreva o contexto geral da discuss\xe3o\n\n3. VERIFICA\xc7\xc3O DOS TOP 3 TERMOS:\n   \u26a0\ufe0f CADA TERMO DO TOP 3 TERMOS DEVE ESTAR CONTEXTUALIZADO NO RESUMO\n   Exemplo: Se "Corinthians" \xe9 o termo mais citado, o resumo DEVE explicar POR QUE falaram tanto dele\n   (baseado no que est\xe1 na transcri\xe7\xe3o)\n\n4. CORRE\xc7\xd5ES DE GRAFIA:\n   - "Felipe Luiz" \u2192 "Filipe Lu\xeds"\n   - "Memphis de Pai" \u2192 "Memphis Depay"\n   - "Anibal" \u2192 "An\xedbal"\n\n5. USE OS NOMES DA LISTA "CORRIGIDOS":\n   - Se diz "CORRIGIDO: X \u2192 Y", use Y\n   - Se um nome foi marcado como REMOVER, n\xe3o use\n\n6. PROIBIDO:\n   \u274c Express\xf5es gen\xe9ricas como "um jogador", "um t\xe9cnico"\n   \u274c Inventar fatos que n\xe3o est\xe3o na transcri\xe7\xe3o\n   \u2705 Use apenas informa\xe7\xf5es presentes nos dados fornecidos\n\n7. INCLUA T\xc9CNICOS quando s\xe3o assunto na transcri\xe7\xe3o\n\n8. FORMATO: 5-6 frases em par\xe1grafo fluido\n\nANTES DE RESPONDER, VERIFIQUE:\n\u2705 Todos os Top 3 Temas est\xe3o mencionados?\n\u2705 Todos os Top 3 Termos est\xe3o contextualizados?\n\u2705 Todas as informa\xe7\xf5es v\xeam da transcri\xe7\xe3o (n\xe3o inventei nada)?\n\nResponda APENAS com o par\xe1grafo final.`;
    try {
        const e = fetchJsonSafe_(CFG.OPENAI_ENDPOINT, {
            method: "post",
            contentType: "application/json",
            headers: {
                Authorization: "Bearer " + r
            },
            payload: JSON.stringify({
                model: CFG.OPENAI_MODEL,
                messages: [{
                    role: "user",
                    content: E
                }],
                max_tokens: 900,
                temperature: .3
            }),
            muteHttpExceptions: !0
        });
        if (e.choices && e.choices[0] && e.choices[0].message) {
            let t = e.choices[0].message.content.trim();
            return Logger.log("=== ETAPA 2D: RESUMO FINAL ==="), Logger.log(t), t
        }
    } catch (e) {
        console.error("Erro ao reescrever resumo:", e)
    }
    return e
}

// --- Extração de clubes e normalização ---

function extrairClubesMencionados_(e) {
    const t = [],
        o = e.split("\n");
    for (const e of o) {
        const o = e.match(/CLUBE:\s*([^|]+)/i);
        if (o) {
            const e = normalizarNomeClube_(o[1].trim().toUpperCase());
            e && !t.includes(e) && t.push(e)
        }
    }
    return t
}

function normalizarNomeClube_(e) {
    return {
        CORINTHIANS: "CORINTHIANS",
        PALMEIRAS: "PALMEIRAS",
        "S\xc3O PAULO": "S\xc3O PAULO",
        "SAO PAULO": "S\xc3O PAULO",
        SANTOS: "SANTOS",
        FLAMENGO: "FLAMENGO",
        FLUMINENSE: "FLUMINENSE",
        VASCO: "VASCO",
        "VASCO DA GAMA": "VASCO",
        BOTAFOGO: "BOTAFOGO",
        INTERNACIONAL: "INTERNACIONAL",
        INTER: "INTERNACIONAL",
        GR\u00caMIO: "GR\xcaMIO",
        GREMIO: "GR\xcaMIO",
        "ATL\xc9TICO-MG": "ATL\xc9TICO-MG",
        "ATLETICO-MG": "ATL\xc9TICO-MG",
        "ATL\xc9TICO MINEIRO": "ATL\xc9TICO-MG",
        GALO: "ATL\xc9TICO-MG",
        CRUZEIRO: "CRUZEIRO",
        BAHIA: "BAHIA",
        FORTALEZA: "FORTALEZA",
        "ATHLETICO-PR": "ATHLETICO-PR",
        ATHLETICO: "ATHLETICO-PR",
        BRAGANTINO: "BRAGANTINO",
        "RED BULL BRAGANTINO": "BRAGANTINO"
    } [e] || null
}

// --- API-Football: busca de elencos ---

function buscarElencosApiFootball_(e, t) {
    const o = {};
    for (const a of e) {
        const e = CFG.TEAM_IDS[a];
        if (e) {
            try {
                const r = fetchJsonSafe_(`${CFG.API_FOOTBALL_ENDPOINT}/players/squads?team=${e}`, {
                    method: "get",
                    headers: {
                        "x-apisports-key": t
                    },
                    muteHttpExceptions: !0
                });
                if (r.response && r.response.length > 0 && r.response[0].players) {
                    const e = r.response[0].players.map(e => e.name);
                    o[a] = e, Logger.log(`${a}: ${e.length} jogadores encontrados`)
                } else r.errors && Object.keys(r.errors).length > 0 && Logger.log(`Erro API-Football para ${a}: ${JSON.stringify(r.errors)}`)
            } catch (e) {
                Logger.log(`Erro ao buscar elenco de ${a}: ${e}`)
            }
            Utilities.sleep(200)
        } else Logger.log("ID n\xe3o encontrado para " + a)
    }
    return o
}

function hasApiFootballKey_() {
    const e = PropertiesService.getScriptProperties().getProperty("API_FOOTBALL_KEY");
    return e && e.length > 0
}

// --- Resumo simples (fallback sem IA) ---

function gerarResumoSimples_(e) {
    const t = [],
        o = e.reduce((e, t) => e + (parseInt(t.minutos) || 0), 0);
    e.length > 0 && t.push(`O programa teve como destaque principal ${e[0].tema}, com ${e[0].minutos} minutos de cobertura.`), e.length > 1 && t.push(`${e[1].tema} foi o segundo tema mais discutido, ocupando ${e[1].minutos} minutos.`), e.length > 2 && t.push(`${e[2].tema} tamb\xe9m ganhou destaque com ${e[2].minutos} minutos de an\xe1lise.`), t.push(`Foram ${o} minutos de conte\xfado esportivo analisado no programa.`);
    const a = e.slice(3).map(e => e.tema).join(", ");
    return a && t.push(`Outros assuntos abordados: ${a}.`), t
}

// --- Envio de email de relatório ---

function enviarEmailRelatorio_(e) {
    const t = `\u{1f4fa} Relat\xf3rio Jogo Aberto - ${e.diaSemana}, ${e.data}`;
    let o = `\n    <div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0 auto;">\n      <h1 style="color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 10px;">\n        \u{1f4fa} Jogo Aberto - ${e.diaSemana}, ${e.data}\n      </h1>\n      \n      <h2 style="color: #333; margin-top: 30px;">\u{1f4dd} Resumo do Programa</h2>\n      <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; line-height: 1.8;">\n  `;
    for (const t of e.resumo) o += `<p style="margin: 8px 0; color: #333;">${t}</p>`;
    o += '\n      </div>\n      \n      <h2 style="color: #333; margin-top: 30px;">\u{1f3c6} Top 5 Termos Mais Falados</h2>\n      <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">\n        <tr style="background: #1a73e8; color: white;">\n          <th style="padding: 10px; text-align: left;">Termo</th>\n          <th style="padding: 10px; text-align: center;">Men\xe7\xf5es</th>\n          <th style="padding: 10px; text-align: center;">Minutos</th>\n        </tr>\n  ';
    for (let t = 0; t < e.topTermos.length; t++) {
        const a = e.topTermos[t];
        o += `\n        <tr style="background: ${t%2==0?"#f8f9fa":"#ffffff"};">\n          <td style="padding: 10px; border-bottom: 1px solid #ddd;"><strong>${a.termo}</strong></td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.contagem}</td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.minutos}</td>\n        </tr>\n    `
    }
    o += '\n      </table>\n      \n      <h2 style="color: #333; margin-top: 30px;">\u{1f4ca} Top 5 Macrotemas</h2>\n      <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">\n        <tr style="background: #1a73e8; color: white;">\n          <th style="padding: 10px; text-align: left;">Tema</th>\n          <th style="padding: 10px; text-align: center;">Minutos</th>\n          <th style="padding: 10px; text-align: center;">%</th>\n        </tr>\n  ';
    for (let t = 0; t < e.macrotemas.length; t++) {
        const a = e.macrotemas[t];
        o += `\n        <tr style="background: ${t%2==0?"#f8f9fa":"#ffffff"};">\n          <td style="padding: 10px; border-bottom: 1px solid #ddd;"><strong>${a.tema}</strong></td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.minutos}</td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.percent}</td>\n        </tr>\n    `
    }
    o += '\n      </table>\n      \n      <h2 style="color: #333; margin-top: 30px;">\u{1f3af} Top 10 Subtemas</h2>\n      <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">\n        <tr style="background: #1a73e8; color: white;">\n          <th style="padding: 10px; text-align: left;">Tema</th>\n          <th style="padding: 10px; text-align: left;">Subtema</th>\n          <th style="padding: 10px; text-align: center;">Min</th>\n          <th style="padding: 10px; text-align: center;">%</th>\n        </tr>\n  ';
    for (let t = 0; t < e.subtemas.length; t++) {
        const a = e.subtemas[t];
        o += `\n        <tr style="background: ${t%2==0?"#f8f9fa":"#ffffff"};">\n          <td style="padding: 10px; border-bottom: 1px solid #ddd;">${a.tema}</td>\n          <td style="padding: 10px; border-bottom: 1px solid #ddd;">${a.subtema}</td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.minutos}</td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.percent}</td>\n        </tr>\n    `
    }
    if (o += "</table>", e.debate && e.debate.resumo) {
        o += '\n      <div style="margin-top: 50px; border-top: 3px solid #e74c3c; padding-top: 20px;">\n        <h1 style="color: #e74c3c; margin-bottom: 20px;">\n          \u{1f399}\ufe0f Debate (12h - 13h)\n        </h1>\n        \n        <h2 style="color: #333; margin-top: 20px;">\u{1f4dd} Resumo do Debate</h2>\n        <div style="background: #fdf2f2; padding: 15px; border-radius: 8px; border-left: 4px solid #e74c3c; line-height: 1.8;">\n    ';
        for (const t of e.debate.resumo) o += `<p style="margin: 8px 0; color: #333;">${t}</p>`;
        o += '\n        </div>\n        \n        <h2 style="color: #333; margin-top: 30px;">\u{1f3c6} Top 5 Termos do Debate</h2>\n        <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">\n          <tr style="background: #e74c3c; color: white;">\n            <th style="padding: 10px; text-align: left;">Termo</th>\n            <th style="padding: 10px; text-align: center;">Men\xe7\xf5es</th>\n          </tr>\n    ';
        for (let t = 0; t < (e.debate.topTermos || []).length; t++) {
            const a = e.debate.topTermos[t];
            o += `\n          <tr style="background: ${t%2==0?"#fdf2f2":"#ffffff"};">\n            <td style="padding: 10px; border-bottom: 1px solid #ddd;"><strong>${a.termo}</strong></td>\n            <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.contagem}</td>\n          </tr>\n      `
        }
        o += '\n        </table>\n        \n        <h2 style="color: #333; margin-top: 30px;">\u{1f4ca} Top 5 Temas do Debate</h2>\n        <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">\n          <tr style="background: #e74c3c; color: white;">\n            <th style="padding: 10px; text-align: left;">Tema</th>\n            <th style="padding: 10px; text-align: center;">Minutos</th>\n            <th style="padding: 10px; text-align: center;">%</th>\n          </tr>\n    ';
        for (let t = 0; t < (e.debate.topTemas || []).length; t++) {
            const a = e.debate.topTemas[t];
            o += `\n          <tr style="background: ${t%2==0?"#fdf2f2":"#ffffff"};">\n            <td style="padding: 10px; border-bottom: 1px solid #ddd;"><strong>${a.tema}</strong></td>\n            <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.minutos}</td>\n            <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.percent}</td>\n          </tr>\n      `
        }
        o += '\n        </table>\n        \n        <h2 style="color: #333; margin-top: 30px;">\u{1f3af} Top 5 Subtemas do Debate</h2>\n        <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">\n          <tr style="background: #e74c3c; color: white;">\n            <th style="padding: 10px; text-align: left;">Tema</th>\n            <th style="padding: 10px; text-align: left;">Subtema</th>\n            <th style="padding: 10px; text-align: center;">Min</th>\n            <th style="padding: 10px; text-align: center;">%</th>\n          </tr>\n    ';
        for (let t = 0; t < (e.debate.topSubtemas || []).length; t++) {
            const a = e.debate.topSubtemas[t];
            o += `\n          <tr style="background: ${t%2==0?"#fdf2f2":"#ffffff"};">\n            <td style="padding: 10px; border-bottom: 1px solid #ddd;">${a.tema}</td>\n            <td style="padding: 10px; border-bottom: 1px solid #ddd;">${a.subtema}</td>\n            <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.minutos}</td>\n            <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${a.percent}</td>\n          </tr>\n      `
        }
        o += "</table></div>"
    }
    o += `\n      <hr style="margin-top: 40px; border: none; border-top: 1px solid #ddd;">\n      <p style="color: #999; font-size: 12px; text-align: center;">\n        Relat\xf3rio gerado automaticamente pelo Analisador Jogo Aberto<br>\n        Emitido em ${(new Date).toLocaleString("pt-BR",{timeZone:"America/Sao_Paulo"})}\n      </p>\n    </div>\n  `, MailApp.sendEmail({
        to: CFG.EMAIL_DESTINO,
        subject: t,
        htmlBody: o
    })
}

// --- Envio de email de erro ---

function enviarEmailErro_(e) {
    try {
        MailApp.sendEmail({
            to: CFG.EMAIL_DESTINO,
            subject: "\u26a0\ufe0f Erro no Analisador Jogo Aberto",
            body: `Ocorreu um erro na execu\xe7\xe3o autom\xe1tica:\n\n${e}\n\nData/hora: ${(new Date).toLocaleString("pt-BR",{timeZone:"America/Sao_Paulo"})}`
        })
    } catch (e) {
        console.error("Falha ao enviar email de erro:", e)
    }
}