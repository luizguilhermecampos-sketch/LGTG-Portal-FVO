"use strict";

// ======================================================
// PRODUÇÃO — BRITAGEM FVO
// Base: aba "Hora - Hora"
// ======================================================

let baseProducao = [];
let colunasProducao = null;
let periodoProdutividade = "mes";
let graficoDiario = null;
let graficoMensal = null;
let graficoTurmas = null;
let nomeArquivoProducao = "";

const MESES_PRODUCAO = [
    "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
];

const MESES_CURTOS_PRODUCAO = [
    "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
    "Jul", "Ago", "Set", "Out", "Nov", "Dez"
];

if (typeof Chart !== "undefined" && typeof ChartDataLabels !== "undefined") {
    Chart.register(ChartDataLabels);
}

// ======================================================
// INICIALIZAÇÃO
// ======================================================

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("arquivoProducao")?.addEventListener("change", carregarArquivoProducao);
    document.getElementById("dataReferenciaProducao")?.addEventListener("change", atualizarDashboardProducao);
    document.getElementById("btnAtualizarProducao")?.addEventListener("click", atualizarDashboardProducao);
    document.getElementById("btnGerarPDFProducao")?.addEventListener("click", gerarPDFProducao);

    document.querySelectorAll(".periodo-producao").forEach(botao => {
        botao.addEventListener("click", () => {
            if (!baseProducao.length) return;
            periodoProdutividade = botao.dataset.periodo || "mes";
            document.querySelectorAll(".periodo-producao").forEach(item => item.classList.remove("ativo"));
            botao.classList.add("ativo");
            atualizarProdutividadeTurmas();
        });
    });

    window.addEventListener("afterprint", restaurarGraficosDepoisImpressao);
    limparDashboardProducao();
});

// ======================================================
// CARREGAMENTO DA PLANILHA
// ======================================================

function carregarArquivoProducao(event) {
    const arquivo = event.target.files?.[0];
    if (!arquivo) return;

    if (typeof XLSX === "undefined") {
        atualizarStatusBaseProducao("Biblioteca não carregada");
        alert("A biblioteca de leitura Excel (SheetJS) não foi carregada.");
        return;
    }

    nomeArquivoProducao = arquivo.name;
    mostrarCarregamentoProducao("Lendo arquivo Hora - Hora...");
    atualizarStatusBaseProducao("Carregando...");

    const reader = new FileReader();

    reader.onload = e => {
        setTimeout(() => {
            try {
                processarWorkbookProducao(e.target.result);
            } catch (erro) {
                console.error("Erro ao processar produção:", erro);
                ocultarCarregamentoProducao();
                atualizarStatusBaseProducao("Erro ao carregar");
                bloquearControlesProducao(true);
                alert("Não foi possível processar a planilha. Verifique se o arquivo é válido.");
            }
        }, 30);
    };

    reader.onerror = () => {
        ocultarCarregamentoProducao();
        atualizarStatusBaseProducao("Erro ao ler arquivo");
        alert("Não foi possível ler o arquivo selecionado.");
    };

    reader.readAsArrayBuffer(arquivo);
}

function processarWorkbookProducao(arrayBuffer) {
    mostrarCarregamentoProducao("Abrindo a aba Hora - Hora...");

    const workbook = XLSX.read(new Uint8Array(arrayBuffer), {
        type: "array",
        cellStyles: false,
        cellHTML: false,
        cellNF: false,
        cellText: false
    });

    const planilha = workbook.Sheets["Hora - Hora"];

    if (!planilha) {
        ocultarCarregamentoProducao();
        atualizarStatusBaseProducao("Aba não encontrada");
        bloquearControlesProducao(true);
        alert("A aba 'Hora - Hora' não foi encontrada na planilha.");
        return;
    }

    mostrarCarregamentoProducao("Identificando Data, Hora, Produção e Turma...");
    const mapeamento = detectarColunasProducao(planilha);

    if (!mapeamento.ok) {
        ocultarCarregamentoProducao();
        atualizarStatusBaseProducao("Colunas não encontradas");
        bloquearControlesProducao(true);
        alert(mapeamento.mensagem);
        return;
    }

    colunasProducao = mapeamento;
    mostrarCarregamentoProducao("Consolidando produção e turmas...");
    baseProducao = extrairBaseProducao(planilha, mapeamento);

    if (!baseProducao.length) {
        ocultarCarregamentoProducao();
        atualizarStatusBaseProducao("Nenhum dado válido");
        bloquearControlesProducao(true);
        limparDashboardProducao();
        alert("Nenhum registro válido de produção foi encontrado na aba 'Hora - Hora'.");
        return;
    }

    configurarDataInicialProducao();
    bloquearControlesProducao(false);
    atualizarStatusBaseProducao(`${formatarInteiroProducao(baseProducao.length)} registros válidos`);
    definirTextoProducao("ultimaAtualizacaoProducao", new Date().toLocaleString("pt-BR"));
    ocultarCarregamentoProducao();
    atualizarDashboardProducao();
}

// ======================================================
// DETECÇÃO DE CABEÇALHOS
// Cabeçalhos historicamente ficam na linha 4, mas o código
// também verifica as primeiras linhas para manter robustez.
// ======================================================

function detectarColunasProducao(planilha) {
    if (!planilha?.["!ref"]) {
        return { ok: false, mensagem: "A aba 'Hora - Hora' está vazia." };
    }

    const faixa = XLSX.utils.decode_range(planilha["!ref"]);
    const limiteLinha = Math.min(faixa.e.r, Math.max(8, faixa.s.r + 10));
    let melhor = null;

    for (let linha = faixa.s.r; linha <= limiteLinha; linha++) {
        const candidatos = { data: null, hora: null, producao: null, turma: null };

        for (let coluna = faixa.s.c; coluna <= faixa.e.c; coluna++) {
            const bruto = obterValorCelulaProducao(planilha, linha, coluna);
            const cabecalho = normalizarTextoProducao(bruto);
            if (!cabecalho) continue;

            atualizarCandidatoColuna(candidatos, "data", coluna, pontuarCabecalhoData(cabecalho));
            atualizarCandidatoColuna(candidatos, "hora", coluna, pontuarCabecalhoHora(cabecalho));
            atualizarCandidatoColuna(candidatos, "producao", coluna, pontuarCabecalhoProducao(cabecalho));
            atualizarCandidatoColuna(candidatos, "turma", coluna, pontuarCabecalhoTurma(cabecalho));
        }

        const obrigatorios = ["data", "producao", "turma"].filter(chave => candidatos[chave]?.score > 0).length;
        const score = Object.values(candidatos).reduce((soma, item) => soma + (item?.score || 0), 0);

        if (obrigatorios === 3 && (!melhor || score > melhor.score)) {
            melhor = { linhaCabecalho: linha, candidatos, score };
        }
    }

    if (!melhor) {
        return {
            ok: false,
            mensagem:
                "Não foi possível identificar todos os campos necessários na aba 'Hora - Hora'.\n\n" +
                "Campos esperados: Data, Produção Horária e Turma/Turno/Equipe.\n" +
                "O campo Hora é recomendado para calcular as horas registradas com maior precisão."
        };
    }

    return {
        ok: true,
        linhaCabecalho: melhor.linhaCabecalho,
        data: melhor.candidatos.data.coluna,
        hora: melhor.candidatos.hora?.score > 0 ? melhor.candidatos.hora.coluna : null,
        producao: melhor.candidatos.producao.coluna,
        turma: melhor.candidatos.turma.coluna
    };
}

function atualizarCandidatoColuna(objeto, chave, coluna, score) {
    if (score <= 0) return;
    if (!objeto[chave] || score > objeto[chave].score) objeto[chave] = { coluna, score };
}

function pontuarCabecalhoData(texto) {
    if (texto === "data") return 100;
    if (texto === "data hora" || texto === "data/hora") return 90;
    return 0;
}

function pontuarCabecalhoHora(texto) {
    if (texto === "hora" || texto === "horario") return 100;
    if (texto.includes("hora") && !texto.includes("parad")) return 70;
    return 0;
}

function pontuarCabecalhoProducao(texto) {
    if (texto === "producao horaria") return 120;
    if (texto === "producao hora") return 115;
    if (texto === "producao t" || texto === "producao") return 100;
    if (texto.includes("producao") && texto.includes("hora")) return 110;
    if (texto.includes("producao") && !/(meta|dia|mes|turno|acumul)/.test(texto)) return 85;
    if (texto === "realizado" || texto === "realizado t") return 60;
    return 0;
}

function pontuarCabecalhoTurma(texto) {
    if (texto === "turma") return 120;
    if (texto === "equipe") return 110;
    if (texto === "turno") return 100;
    if (texto.includes("turma")) return 90;
    if (texto.includes("equipe")) return 85;
    if (texto.includes("turno")) return 80;
    return 0;
}

// ======================================================
// EXTRAÇÃO DA BASE
// ======================================================

function extrairBaseProducao(planilha, mapa) {
    const base = [];
    const faixa = XLSX.utils.decode_range(planilha["!ref"]);
    const linhaInicial = mapa.linhaCabecalho + 1;

    for (let linha = linhaInicial; linha <= faixa.e.r; linha++) {
        const data = converterDataProducao(obterValorCelulaProducao(planilha, linha, mapa.data));
        if (!data) continue;

        const producao = numeroProducao(obterValorCelulaProducao(planilha, linha, mapa.producao));
        if (producao === null || producao < 0) continue;

        const turmaBruta = textoProducao(obterValorCelulaProducao(planilha, linha, mapa.turma));
        const turma = turmaBruta || "Sem turma";
        const hora = mapa.hora === null ? "" : normalizarHoraProducao(obterValorCelulaProducao(planilha, linha, mapa.hora));

        base.push({
            idLinha: linha,
            data: data.texto,
            ano: data.ano,
            mes: data.mes,
            dia: data.dia,
            hora,
            turma,
            producao
        });
    }

    return base.sort((a, b) => a.data.localeCompare(b.data) || a.hora.localeCompare(b.hora));
}

function obterValorCelulaProducao(planilha, linha, coluna) {
    if (coluna === null || coluna === undefined) return null;
    const endereco = XLSX.utils.encode_cell({ r: linha, c: coluna });
    return planilha[endereco]?.v ?? null;
}

// ======================================================
// FILTRO / DASHBOARD
// ======================================================

function configurarDataInicialProducao() {
    const datas = [...new Set(baseProducao.map(item => item.data))].sort();
    const campo = document.getElementById("dataReferenciaProducao");
    if (!campo || !datas.length) return;
    campo.min = datas[0];
    campo.max = datas[datas.length - 1];
    campo.value = datas[datas.length - 1];
}

function atualizarDashboardProducao() {
    if (!baseProducao.length) return;

    const referencia = obterReferenciaProducao();
    if (!referencia) {
        alert("Selecione uma data de referência válida.");
        return;
    }

    const dadosDia = baseProducao.filter(item => item.data === referencia.texto);
    const dadosMes = baseProducao.filter(item => item.ano === referencia.ano && item.mes === referencia.mes);
    const dadosAno = baseProducao.filter(item => item.ano === referencia.ano);

    definirTextoProducao("producaoDia", formatarToneladasProducao(somarProducao(dadosDia)));
    definirTextoProducao("producaoMes", formatarToneladasProducao(somarProducao(dadosMes)));
    definirTextoProducao("producaoAno", formatarToneladasProducao(somarProducao(dadosAno)));

    definirTextoProducao("labelProducaoDia", formatarDataBRProducao(referencia.texto));
    definirTextoProducao("labelProducaoMes", `${MESES_PRODUCAO[referencia.mes]} / ${referencia.ano}`);
    definirTextoProducao("labelProducaoAno", `${referencia.ano} • acumulado da base`);
    definirTextoProducao(
        "subtituloReferenciaProducao",
        `Referência: ${formatarDataBRProducao(referencia.texto)} • ${formatarInteiroProducao(dadosDia.length)} registro(s) no dia.`
    );

    atualizarGraficosProducao(referencia, dadosMes, dadosAno);
    atualizarProdutividadeTurmas();
    atualizarMetadadosPDFProducao(referencia);
}

function atualizarGraficosProducao(referencia, dadosMes, dadosAno) {
    const porDia = agruparSoma(dadosMes, item => item.data);
    const datas = Object.keys(porDia).sort();
    const valoresDia = datas.map(data => porDia[data]);
    const labelsDia = datas.map(data => data.slice(8, 10));

    const mesesAno = Array.from({ length: 12 }, (_, indice) => indice + 1);
    const valoresMes = mesesAno.map(mes => {
        const registros = dadosAno.filter(item => item.mes === mes);
        return registros.length ? somarProducao(registros) : null;
    });

    definirTextoProducao("descricaoGraficoDiario", `${MESES_PRODUCAO[referencia.mes]} de ${referencia.ano} • produção por dia.`);
    definirTextoProducao("descricaoGraficoMensal", `${referencia.ano} • produção consolidada por mês.`);

    graficoDiario = criarOuAtualizarGraficoBarras(
        graficoDiario,
        "graficoProducaoDiaria",
        labelsDia,
        valoresDia,
        "Produção diária (t)",
        "dia",
        referencia.dia
    );

    graficoMensal = criarOuAtualizarGraficoBarras(
        graficoMensal,
        "graficoProducaoMensal",
        MESES_CURTOS_PRODUCAO,
        valoresMes,
        "Produção mensal (t)",
        "mes",
        referencia.mes
    );
}

function atualizarProdutividadeTurmas() {
    if (!baseProducao.length) return;
    const referencia = obterReferenciaProducao();
    if (!referencia) return;

    const dados = filtrarPeriodoProdutividade(referencia);
    const resumo = calcularProdutividadePorTurma(dados);
    const totalPeriodo = resumo.reduce((soma, item) => soma + item.producao, 0);
    const totalHoras = resumo.reduce((soma, item) => soma + item.horas, 0);

    const labels = resumo.map(item => item.turma);
    const valores = resumo.map(item => item.produtividade);

    if (graficoTurmas) graficoTurmas.destroy();
    const canvas = document.getElementById("graficoProdutividadeTurma");

    if (canvas && typeof Chart !== "undefined") {
        graficoTurmas = new Chart(canvas, {
            type: "bar",
            data: {
                labels,
                datasets: [{
                    label: "Produtividade (t/h)",
                    data: valores,
                    backgroundColor: labels.map((_, i) => i === 0 ? "#3d0764" : "rgba(111, 44, 153, .74)"),
                    borderColor: labels.map((_, i) => i === 0 ? "#3d0764" : "#6f2c99"),
                    borderWidth: 1,
                    borderRadius: 7,
                    maxBarThickness: 72
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                devicePixelRatio: Math.max(window.devicePixelRatio || 1, 2),
                animation: false,
                layout: { padding: { top: 24, right: 8, left: 4, bottom: 0 } },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: context => ` Produtividade: ${formatarNumeroProducao(context.raw, 1)} t/h`
                        }
                    },
                    datalabels: {
                        anchor: "end",
                        align: "top",
                        color: "#2d014d",
                        font: { size: 11, weight: "bold" },
                        formatter: valor => `${formatarNumeroProducao(valor, 1)} t/h`
                    }
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { color: "#544b58", font: { size: 11, weight: "bold" } }
                    },
                    y: {
                        beginAtZero: true,
                        grace: "15%",
                        grid: { color: "rgba(80, 67, 87, .11)" },
                        ticks: { color: "#665d69", font: { size: 10 } },
                        title: { display: true, text: "Produtividade (t/h)", color: "#544b58", font: { size: 10, weight: "bold" } }
                    }
                }
            }
        });
    }

    const melhor = resumo[0] || null;
    definirTextoProducao("melhorProdutividade", melhor ? `${formatarNumeroProducao(melhor.produtividade, 1)} t/h` : "-");
    definirTextoProducao("melhorTurma", melhor ? `${melhor.turma} • ${formatarToneladasProducao(melhor.producao)} t` : "Sem dados no período");
    definirTextoProducao("horasRegistradasTotal", resumo.length ? `${formatarInteiroProducao(totalHoras)} h` : "-");

    const periodoNome = periodoProdutividade === "dia" ? "Dia" : periodoProdutividade === "ano" ? "Ano" : "Mês";
    definirTextoProducao("periodoProdutividadeLabel", periodoNome);
    definirTextoProducao("descricaoProdutividade", `Produção total ÷ horas com registro • ${descricaoPeriodoProdutividade(referencia)}.`);
    definirTextoProducao("periodoPDFProdutividade", `Período: ${descricaoPeriodoProdutividade(referencia)}`);

    montarTabelaTurmas(resumo, totalPeriodo);
}

function filtrarPeriodoProdutividade(referencia) {
    if (periodoProdutividade === "dia") return baseProducao.filter(item => item.data === referencia.texto);
    if (periodoProdutividade === "ano") return baseProducao.filter(item => item.ano === referencia.ano);
    return baseProducao.filter(item => item.ano === referencia.ano && item.mes === referencia.mes);
}

function calcularProdutividadePorTurma(registros) {
    const grupos = new Map();

    registros.forEach(item => {
        if (!grupos.has(item.turma)) grupos.set(item.turma, { turma: item.turma, producao: 0, horas: new Set() });
        const grupo = grupos.get(item.turma);
        grupo.producao += item.producao;
        const chaveHora = item.hora ? `${item.data}|${item.hora}` : `${item.data}|linha-${item.idLinha}`;
        grupo.horas.add(chaveHora);
    });

    return [...grupos.values()]
        .map(grupo => ({
            turma: grupo.turma,
            producao: grupo.producao,
            horas: grupo.horas.size,
            produtividade: grupo.horas.size ? grupo.producao / grupo.horas.size : 0
        }))
        .filter(item => item.horas > 0)
        .sort((a, b) => b.produtividade - a.produtividade || b.producao - a.producao);
}

function montarTabelaTurmas(resumo, totalPeriodo) {
    const tbody = document.getElementById("tabelaTurmasBody");
    if (!tbody) return;

    if (!resumo.length) {
        tbody.innerHTML = '<tr class="linha-vazia-producao"><td colspan="5">Sem registros de turma no período selecionado.</td></tr>';
        return;
    }

    tbody.innerHTML = resumo.map(item => {
        const participacao = totalPeriodo > 0 ? item.producao / totalPeriodo * 100 : 0;
        return `<tr>
            <td>${escaparHTMLProducao(item.turma)}</td>
            <td>${formatarToneladasProducao(item.producao)}</td>
            <td>${formatarInteiroProducao(item.horas)} h</td>
            <td>${formatarNumeroProducao(item.produtividade, 1)} t/h</td>
            <td>${formatarNumeroProducao(participacao, 1)}%</td>
        </tr>`;
    }).join("");
}

// ======================================================
// GRÁFICOS DE PRODUÇÃO
// ======================================================

function criarOuAtualizarGraficoBarras(instancia, canvasId, labels, valores, label, tipo, destaque) {
    if (instancia) instancia.destroy();
    const canvas = document.getElementById(canvasId);
    if (!canvas || typeof Chart === "undefined") return null;

    const cores = labels.map((_, indice) => {
        const atual = tipo === "dia" ? Number(labels[indice]) === Number(destaque) : indice + 1 === Number(destaque);
        return atual ? "#92d050" : "#4c0a72";
    });

    return new Chart(canvas, {
        type: "bar",
        data: {
            labels,
            datasets: [{
                label,
                data: valores,
                backgroundColor: cores,
                borderColor: cores,
                borderWidth: 1,
                borderRadius: 5,
                maxBarThickness: 44
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            devicePixelRatio: Math.max(window.devicePixelRatio || 1, 2),
            animation: false,
            layout: { padding: { top: 22, right: 4, left: 2, bottom: 0 } },
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: { label: context => context.raw === null ? " Sem dados" : ` Produção: ${formatarToneladasProducao(context.raw)} t` }
                },
                datalabels: {
                    display: context => context.raw !== null && context.raw > 0,
                    anchor: "end",
                    align: "top",
                    color: "#2d014d",
                    font: { size: 9, weight: "bold" },
                    formatter: valor => formatarNumeroCompactoProducao(valor)
                }
            },
            scales: {
                x: {
                    grid: { display: false },
                    ticks: { color: "#5f5662", font: { size: 9, weight: "bold" }, maxRotation: 0, autoSkip: tipo === "dia", maxTicksLimit: tipo === "dia" ? 16 : 12 },
                    title: { display: true, text: tipo === "dia" ? "Dia" : "Mês", color: "#6d6470", font: { size: 9, weight: "bold" } }
                },
                y: {
                    beginAtZero: true,
                    grace: "12%",
                    grid: { color: "rgba(80, 67, 87, .11)" },
                    ticks: { color: "#675e6b", font: { size: 9 }, callback: value => formatarNumeroCompactoProducao(value) },
                    title: { display: true, text: "Produção (t)", color: "#5f5662", font: { size: 9, weight: "bold" } }
                }
            }
        }
    });
}

// ======================================================
// PDF
// ======================================================

function gerarPDFProducao() {
    if (!baseProducao.length) {
        alert("Carregue a planilha antes de gerar o PDF.");
        return;
    }

    const referencia = obterReferenciaProducao();
    atualizarMetadadosPDFProducao(referencia);
    prepararGraficosParaImpressao();

    const tituloAnterior = document.title;
    document.title = `Producao_Britagem_FVO_${referencia?.texto || "Relatorio"}`;

    setTimeout(() => {
        window.print();
        setTimeout(() => { document.title = tituloAnterior; }, 400);
    }, 180);
}

function atualizarMetadadosPDFProducao(referencia) {
    if (!referencia) return;
    const agora = new Date().toLocaleString("pt-BR");
    definirTextoProducao("periodoPDFProducao", `Referência: ${formatarDataBRProducao(referencia.texto)}`);
    definirTextoProducao("geradoPDFProducao", `Gerado em: ${agora}`);
    definirTextoProducao("geradoPDFProducao2", `Gerado em: ${agora}`);
}

function prepararGraficosParaImpressao() {
    [graficoDiario, graficoMensal, graficoTurmas].forEach(grafico => {
        if (!grafico) return;
        grafico.$dprTela = grafico.options.devicePixelRatio;
        grafico.options.devicePixelRatio = 3;
        if (grafico.options.plugins?.datalabels?.font) grafico.options.plugins.datalabels.font.size = 12;
        grafico.resize();
        grafico.update("none");
    });
}

function restaurarGraficosDepoisImpressao() {
    [graficoDiario, graficoMensal, graficoTurmas].forEach(grafico => {
        if (!grafico) return;
        grafico.options.devicePixelRatio = grafico.$dprTela || Math.max(window.devicePixelRatio || 1, 2);
        if (grafico.options.plugins?.datalabels?.font) grafico.options.plugins.datalabels.font.size = grafico === graficoTurmas ? 11 : 9;
        grafico.resize();
        grafico.update("none");
    });
}

// ======================================================
// UTILITÁRIOS
// ======================================================

function obterReferenciaProducao() {
    const valor = document.getElementById("dataReferenciaProducao")?.value || "";
    if (!valor) return null;
    const partes = valor.split("-").map(Number);
    if (partes.length !== 3 || !partes[0] || !partes[1] || !partes[2]) return null;
    return { texto: valor, ano: partes[0], mes: partes[1], dia: partes[2] };
}

function somarProducao(registros) {
    return registros.reduce((soma, item) => soma + item.producao, 0);
}

function agruparSoma(registros, chaveFn) {
    return registros.reduce((acc, item) => {
        const chave = chaveFn(item);
        acc[chave] = (acc[chave] || 0) + item.producao;
        return acc;
    }, {});
}

function descricaoPeriodoProdutividade(referencia) {
    if (periodoProdutividade === "dia") return formatarDataBRProducao(referencia.texto);
    if (periodoProdutividade === "ano") return String(referencia.ano);
    return `${MESES_PRODUCAO[referencia.mes]} de ${referencia.ano}`;
}

function converterDataProducao(valor) {
    if (valor === null || valor === undefined || valor === "") return null;

    if (typeof valor === "number" && Number.isFinite(valor) && typeof XLSX !== "undefined") {
        const data = XLSX.SSF.parse_date_code(valor);
        if (data) return montarObjetoDataProducao(data.y, data.m, data.d);
    }

    if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
        return montarObjetoDataProducao(valor.getFullYear(), valor.getMonth() + 1, valor.getDate());
    }

    const texto = String(valor).trim().split(" ")[0];
    let partes;

    if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(texto)) {
        partes = texto.split(/[-/]/).map(Number);
        return montarObjetoDataProducao(partes[0], partes[1], partes[2]);
    }

    if (/^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}$/.test(texto)) {
        partes = texto.split(/[/.-]/).map(Number);
        let ano = partes[2];
        if (ano < 100) ano += 2000;
        return montarObjetoDataProducao(ano, partes[1], partes[0]);
    }

    return null;
}

function montarObjetoDataProducao(ano, mes, dia) {
    if (!ano || !mes || !dia || mes < 1 || mes > 12 || dia < 1 || dia > 31) return null;
    const dataTeste = new Date(ano, mes - 1, dia);
    if (dataTeste.getFullYear() !== ano || dataTeste.getMonth() !== mes - 1 || dataTeste.getDate() !== dia) return null;
    return { ano, mes, dia, texto: `${ano}-${String(mes).padStart(2, "0")}-${String(dia).padStart(2, "0")}` };
}

function normalizarHoraProducao(valor) {
    if (valor === null || valor === undefined || valor === "") return "";

    if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
        return `${String(valor.getHours()).padStart(2, "0")}:${String(valor.getMinutes()).padStart(2, "0")}`;
    }

    if (typeof valor === "number" && Number.isFinite(valor)) {
        const fracao = ((valor % 1) + 1) % 1;
        const totalMinutos = Math.round(fracao * 24 * 60) % (24 * 60);
        return `${String(Math.floor(totalMinutos / 60)).padStart(2, "0")}:${String(totalMinutos % 60).padStart(2, "0")}`;
    }

    const texto = String(valor).trim();
    const match = texto.match(/(\d{1,2}):(\d{2})/);
    if (match) return `${String(Number(match[1])).padStart(2, "0")}:${match[2]}`;
    if (/^\d{1,2}$/.test(texto)) return `${String(Number(texto)).padStart(2, "0")}:00`;
    return texto;
}

function numeroProducao(valor) {
    if (valor === null || valor === undefined || valor === "") return null;
    if (typeof valor === "number") return Number.isFinite(valor) ? valor : null;

    let texto = String(valor).trim().replace(/\s/g, "");
    if (!texto || /#(DIV\/0|N\/A|VALUE|REF|NUM)/i.test(texto)) return null;

    if (texto.includes(",") && texto.includes(".")) {
        if (texto.lastIndexOf(",") > texto.lastIndexOf(".")) texto = texto.replace(/\./g, "").replace(",", ".");
        else texto = texto.replace(/,/g, "");
    } else if (texto.includes(",")) {
        texto = texto.replace(/\./g, "").replace(",", ".");
    }

    texto = texto.replace(/[^0-9+\-.]/g, "");
    const numero = Number(texto);
    return Number.isFinite(numero) ? numero : null;
}

function textoProducao(valor) {
    if (valor === null || valor === undefined) return "";
    return String(valor).trim();
}

function normalizarTextoProducao(valor) {
    return textoProducao(valor)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[()\[\]_%]/g, " ")
        .replace(/[^a-z0-9/]+/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function formatarToneladasProducao(valor) {
    return Number(valor || 0).toLocaleString("pt-BR", { maximumFractionDigits: 0 });
}

function formatarNumeroProducao(valor, casas = 1) {
    return Number(valor || 0).toLocaleString("pt-BR", { minimumFractionDigits: casas, maximumFractionDigits: casas });
}

function formatarInteiroProducao(valor) {
    return Number(valor || 0).toLocaleString("pt-BR", { maximumFractionDigits: 0 });
}

function formatarNumeroCompactoProducao(valor) {
    const numero = Number(valor || 0);
    if (Math.abs(numero) >= 1_000_000) return `${formatarNumeroProducao(numero / 1_000_000, 1)} mi`;
    if (Math.abs(numero) >= 1_000) return `${formatarNumeroProducao(numero / 1_000, 0)} mil`;
    return formatarInteiroProducao(numero);
}

function formatarDataBRProducao(valor) {
    const partes = String(valor || "").split("-");
    return partes.length === 3 ? `${partes[2]}/${partes[1]}/${partes[0]}` : valor;
}

function definirTextoProducao(id, valor) {
    const elemento = document.getElementById(id);
    if (elemento) elemento.textContent = valor;
}

function atualizarStatusBaseProducao(texto) {
    definirTextoProducao("statusBaseProducao", texto);
}

function bloquearControlesProducao(bloquear) {
    ["dataReferenciaProducao", "btnAtualizarProducao", "btnGerarPDFProducao"].forEach(id => {
        const elemento = document.getElementById(id);
        if (elemento) elemento.disabled = bloquear;
    });
    document.querySelectorAll(".periodo-producao").forEach(botao => { botao.disabled = bloquear; });
}

function mostrarCarregamentoProducao(mensagem) {
    if (window.FVOCarregamento?.exibir) window.FVOCarregamento.exibir(mensagem);
}

function ocultarCarregamentoProducao() {
    if (window.FVOCarregamento?.ocultar) window.FVOCarregamento.ocultar();
}

function escaparHTMLProducao(valor) {
    return String(valor ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function limparDashboardProducao() {
    ["producaoDia", "producaoMes", "producaoAno", "melhorProdutividade"].forEach(id => definirTextoProducao(id, "-"));
    definirTextoProducao("melhorTurma", "Aguardando dados");
    definirTextoProducao("horasRegistradasTotal", "-");
    bloquearControlesProducao(true);
}