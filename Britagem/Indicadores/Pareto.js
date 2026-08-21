// ======================================================
// PARETO - PORTAL BRITAGEM FVO
// Leitura da aba "Hora - Hora" e análise por Dia/Mês/Pilha
// ======================================================

let graficoAtual = null;
let filtroAtivo = "dia";
let basePareto = [];
let paretoGerado = false;
let ultimoValorFiltro = null;
let rankingAtual = [];
let ocorrenciasAtual = 0;
let perdaTotalAtual = 0;
let estadoGraficoAntesImpressao = null;

const emailsPareto = [
    "luiz.campos@br.cmoc.com",
    "thais.gilvana@br.cmoc.com"
];

const nomesMesesPareto = [
    "",
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro"
];

// Mapeamento validado da aba "Hora - Hora"
const COL_DATA = 0;   // A - Data
const COL_PERDA = 4;  // E - Perda / Status
const COL_MODO = 6;   // G - Modo de Falha
const COL_MES = 10;   // K - Mês
const COL_PILHA = 11; // L - Pilha

if (typeof Chart !== "undefined" && typeof ChartDataLabels !== "undefined") {
    Chart.register(ChartDataLabels);
}

// ======================================================
// INICIALIZAÇÃO E EVENTOS
// ======================================================

document.addEventListener("DOMContentLoaded", () => {
    const arquivo = document.getElementById("arquivoExcel");
    const gerar = document.getElementById("btnGerarPareto");
    const pdf = document.getElementById("btnGerarPDFPareto");
    const email = document.getElementById("btnEnviarEmailPareto");

    arquivo?.addEventListener("change", carregarBasePareto);
    gerar?.addEventListener("click", gerarPareto);
    pdf?.addEventListener("click", gerarPDFPareto);
    email?.addEventListener("click", enviarEmailPareto);

    document.getElementById("dataFiltro")?.addEventListener("change", () => {
        if (basePareto.length && filtroAtivo === "dia") gerarPareto();
    });

    document.getElementById("mesFiltro")?.addEventListener("change", () => {
        if (basePareto.length && filtroAtivo === "mes") gerarPareto();
    });

    document.getElementById("pilhaFiltro")?.addEventListener("keydown", event => {
        if (event.key === "Enter" && basePareto.length && filtroAtivo === "pilha") gerarPareto();
    });

    window.addEventListener("beforeprint", prepararGraficoParaImpressaoPareto);
    window.addEventListener("afterprint", restaurarGraficoAposImpressaoPareto);
});

// ======================================================
// CARREGAMENTO VISUAL
// ======================================================

function mostrarCarregamentoPareto(mostrar, texto = "Preparando planilha") {
    const overlay = document.getElementById("overlayCarregamentoPareto");
    const textoElemento = document.getElementById("textoCarregamentoPareto");

    if (textoElemento) textoElemento.textContent = texto;
    if (!overlay) return;

    overlay.classList.toggle("visivel", Boolean(mostrar));
    overlay.setAttribute("aria-hidden", mostrar ? "false" : "true");
}

// ======================================================
// CARREGAR / PROCESSAR EXCEL
// ======================================================

function carregarBasePareto(event) {
    const arquivo = event.target.files?.[0];
    if (!arquivo) return;

    if (typeof XLSX === "undefined") {
        atualizarStatusBase("Biblioteca não carregada");
        alert("A biblioteca de leitura Excel (SheetJS) não foi carregada.");
        return;
    }

    mostrarCarregamentoPareto(true, "Lendo arquivo...");
    atualizarStatusBase("Carregando...");

    const reader = new FileReader();

    reader.onload = e => {
        setTimeout(() => {
            try {
                processarArquivoPareto(e.target.result);
            } catch (erro) {
                console.error("Erro ao processar Pareto:", erro);
                mostrarCarregamentoPareto(false);
                atualizarStatusBase("Erro ao carregar");
                alert("Não foi possível processar a planilha. Verifique se o arquivo é válido.");
            }
        }, 30);
    };

    reader.onerror = () => {
        mostrarCarregamentoPareto(false);
        atualizarStatusBase("Erro ao carregar");
        alert("Não foi possível ler o arquivo selecionado.");
    };

    reader.readAsArrayBuffer(arquivo);
}

function processarArquivoPareto(arrayBuffer) {
    mostrarCarregamentoPareto(true, "Abrindo planilha...");

    const dados = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(dados, {
        type: "array",
        cellStyles: false,
        cellHTML: false,
        cellNF: false,
        cellText: false
    });

    const aba = workbook.Sheets["Hora - Hora"];

    if (!aba) {
        mostrarCarregamentoPareto(false);
        atualizarStatusBase("Aba não encontrada");
        alert("A aba 'Hora - Hora' não foi encontrada na planilha.");
        return;
    }

    mostrarCarregamentoPareto(true, "Preparando base do Pareto...");
    basePareto = extrairBaseCompactaPareto(aba);

    if (!basePareto.length) {
        mostrarCarregamentoPareto(false);
        atualizarStatusBase("Nenhum dado válido");
        limparDashboard();
        alert("Nenhum registro de perda válido foi encontrado na aba 'Hora - Hora'.");
        return;
    }

    definirFiltrosIniciais();
    atualizarStatusBase(`${basePareto.length.toLocaleString("pt-BR")} registros carregados`);
    mostrarCarregamentoPareto(false);
    gerarPareto();
}

function extrairBaseCompactaPareto(planilha) {
    const base = [];
    if (!planilha?.["!ref"]) return base;

    const faixa = XLSX.utils.decode_range(planilha["!ref"]);
    const linhaInicial = Math.max(4, faixa.s.r);

    for (let linha = linhaInicial; linha <= faixa.e.r; linha++) {
        const data = converterDataPareto(obterValorCelulaPareto(planilha, linha, COL_DATA));
        if (!data) continue;

        const status = numeroPareto(obterValorCelulaPareto(planilha, linha, COL_PERDA));
        if (status === null || status >= 0) continue;

        const modo = textoPareto(obterValorCelulaPareto(planilha, linha, COL_MODO));
        if (!modo) continue;

        const mesPlanilha = obterValorCelulaPareto(planilha, linha, COL_MES);
        const mes = normalizarMesPareto(mesPlanilha) || String(data.mes);
        const pilha = textoPareto(obterValorCelulaPareto(planilha, linha, COL_PILHA));

        base.push({
            data: data.texto,
            ano: data.ano,
            mes,
            pilha,
            modo,
            perda: Math.abs(status)
        });
    }

    return base;
}

function obterValorCelulaPareto(planilha, linha, coluna) {
    const endereco = XLSX.utils.encode_cell({ r: linha, c: coluna });
    return planilha[endereco]?.v ?? null;
}

function normalizarMesPareto(valor) {
    if (valor === null || valor === undefined || valor === "") return "";

    if (typeof valor === "number" && Number.isFinite(valor)) {
        const numero = Math.trunc(valor);
        return numero >= 1 && numero <= 12 ? String(numero) : "";
    }

    const bruto = removerAcentosPareto(String(valor).trim().toLowerCase());
    const numeroDireto = Number(bruto.replace(/[^0-9]/g, ""));
    if (numeroDireto >= 1 && numeroDireto <= 12) return String(numeroDireto);

    const mapa = {
        janeiro: 1, jan: 1,
        fevereiro: 2, fev: 2,
        marco: 3, mar: 3,
        abril: 4, abr: 4,
        maio: 5, mai: 5,
        junho: 6, jun: 6,
        julho: 7, jul: 7,
        agosto: 8, ago: 8,
        setembro: 9, set: 9,
        outubro: 10, out: 10,
        novembro: 11, nov: 11,
        dezembro: 12, dez: 12
    };

    return mapa[bruto] ? String(mapa[bruto]) : "";
}

function removerAcentosPareto(texto) {
    return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// ======================================================
// FILTROS
// ======================================================

function definirFiltrosIniciais() {
    if (!basePareto.length) return;

    const datas = [...new Set(basePareto.map(item => item.data))].sort();
    const ultimaData = datas[datas.length - 1];
    const primeiraData = datas[0];
    const dataFiltro = document.getElementById("dataFiltro");

    if (dataFiltro) {
        dataFiltro.value = ultimaData;
        dataFiltro.max = ultimaData;
        dataFiltro.min = primeiraData;
    }

    const partes = ultimaData.split("-");
    const mesFiltro = document.getElementById("mesFiltro");
    if (mesFiltro) mesFiltro.value = String(Number(partes[1]));
}

function mostrarFiltro(tipo, botao) {
    filtroAtivo = tipo;

    document.querySelectorAll(".filtro-pareto").forEach(elemento => elemento.classList.add("oculto"));
    document.getElementById(`filtro-${tipo}`)?.classList.remove("oculto");

    document.querySelectorAll(".painel-pareto .aba").forEach(elemento => elemento.classList.remove("ativa"));
    botao?.classList.add("ativa");

    if (basePareto.length) {
        const valor = obterValorFiltroAtual(false);
        if (valor !== null) gerarPareto();
    }
}

function obterValorFiltroAtual(mostrarAlerta = true) {
    if (filtroAtivo === "dia") {
        const valor = document.getElementById("dataFiltro")?.value || "";
        if (!valor && mostrarAlerta) alert("Selecione uma data.");
        return valor || null;
    }

    if (filtroAtivo === "mes") {
        return document.getElementById("mesFiltro")?.value || null;
    }

    if (filtroAtivo === "pilha") {
        const valor = (document.getElementById("pilhaFiltro")?.value || "").trim();
        if (!valor && mostrarAlerta) alert("Digite uma pilha.");
        return valor || null;
    }

    return null;
}

// ======================================================
// GERAR ANÁLISE
// ======================================================

function gerarPareto() {
    if (!basePareto.length) {
        alert("Selecione e carregue uma planilha primeiro.");
        return;
    }

    const valorFiltro = obterValorFiltroAtual(true);
    if (valorFiltro === null) return;

    ultimoValorFiltro = valorFiltro;

    const perdas = {};
    const ocorrenciasPorModo = {};
    let perdaTotal = 0;
    let ocorrenciasTotal = 0;

    for (const item of basePareto) {
        if (!registroAtendeFiltroPareto(item, valorFiltro)) continue;

        perdas[item.modo] = (perdas[item.modo] || 0) + item.perda;
        ocorrenciasPorModo[item.modo] = (ocorrenciasPorModo[item.modo] || 0) + 1;
        perdaTotal += item.perda;
        ocorrenciasTotal += 1;
    }

    montarPareto(perdas, perdaTotal, valorFiltro, ocorrenciasPorModo, ocorrenciasTotal);
}

function registroAtendeFiltroPareto(item, valorFiltro) {
    if (filtroAtivo === "dia") return item.data === valorFiltro;
    if (filtroAtivo === "mes") return item.mes === String(valorFiltro);
    if (filtroAtivo === "pilha") return normalizarPilhaPareto(item.pilha) === normalizarPilhaPareto(valorFiltro);
    return false;
}

function normalizarPilhaPareto(valor) {
    return String(valor ?? "").trim().toUpperCase().replace(/\s+/g, " ");
}

function montarPareto(perdas, perdaTotal, valorFiltro, ocorrenciasPorModo, ocorrenciasTotal) {
    const ranking = Object.entries(perdas)
        .map(([modo, perda]) => ({
            modo,
            perda,
            ocorrencias: ocorrenciasPorModo[modo] || 0
        }))
        .sort((a, b) => b.perda - a.perda);

    if (!ranking.length || perdaTotal <= 0) {
        limparDashboard();
        alert("Nenhum dado de perda foi encontrado para o filtro selecionado.");
        return;
    }

    rankingAtual = ranking;
    ocorrenciasAtual = ocorrenciasTotal;
    perdaTotalAtual = perdaTotal;

    atualizarCards(ranking, perdaTotal, ocorrenciasTotal);
    atualizarDetalhamentoPareto(ranking, perdaTotal, ocorrenciasTotal);

    const rankingGrafico = prepararRankingGraficoPareto(ranking);
    let acumulado = 0;

    const labels = rankingGrafico.map(item => item.modo);
    const valores = rankingGrafico.map(item => item.perda);
    const percentuais = rankingGrafico.map(item => {
        acumulado += item.perda;
        return (acumulado / perdaTotal) * 100;
    });

    desenharGrafico(labels, valores, percentuais);
    atualizarPeriodo(valorFiltro);
    atualizarCabecalhoPDFPareto();
    paretoGerado = true;
}

function prepararRankingGraficoPareto(ranking) {
    // Exibe somente os 10 maiores ofensores.
    // As perdas restantes continuam compondo o total usado nos cards e
    // no percentual acumulado, mas não são agrupadas em uma barra "Outros".
    return ranking.slice(0, 10).map(item => ({ ...item }));
}

// ======================================================
// CARDS E RESUMO
// ======================================================

function atualizarCards(ranking, perdaTotal, ocorrenciasTotal) {
    const principal = ranking[0];
    const percentualPrincipal = (principal.perda / perdaTotal) * 100;
    const somaTop3 = ranking.slice(0, 3).reduce((total, item) => total + item.perda, 0);
    const percentualTop3 = (somaTop3 / perdaTotal) * 100;

    definirTexto("perdaTotal", formatarToneladas(perdaTotal));
    definirTexto("principalOfensor", principal.modo);
    definirTexto("impactoPrincipal", formatarPercentualPareto(percentualPrincipal));
    definirTexto("impactoTop3", formatarPercentualPareto(percentualTop3));
    definirTexto("totalOcorrencias", ocorrenciasTotal.toLocaleString("pt-BR"));
}

function atualizarDetalhamentoPareto(ranking, perdaTotal, ocorrenciasTotal) {
    const corpo = document.getElementById("rankingParetoBody");
    if (!corpo) return;

    let acumulado = 0;
    const linhas = ranking.slice(0, 10).map((item, indice) => {
        const participacao = (item.perda / perdaTotal) * 100;
        acumulado += participacao;

        return `
            <tr>
                <td><span class="posicao-ranking-pareto">${indice + 1}</span></td>
                <td>${escaparHTMLPareto(item.modo)}</td>
                <td>${formatarToneladas(item.perda)}</td>
                <td>
                    <span class="barra-participacao-pareto">
                        <i style="--largura:${Math.min(participacao, 100).toFixed(1)}%"></i>
                        <b>${formatarPercentualPareto(participacao)}</b>
                    </span>
                </td>
                <td>${formatarPercentualPareto(acumulado)}</td>
                <td>${item.ocorrencias.toLocaleString("pt-BR")}</td>
            </tr>`;
    });

    corpo.innerHTML = linhas.join("");

    const principal = ranking[0];
    const percentualTop3 = ranking.slice(0, 3).reduce((total, item) => total + item.perda, 0) / perdaTotal * 100;
    const perdaMedia = ocorrenciasTotal ? perdaTotal / ocorrenciasTotal : 0;
    const terceiroAcumulado = ranking.slice(0, 3).reduce((total, item) => total + item.perda, 0) / perdaTotal * 100;

    let titulo;
    let leitura;

    if (percentualTop3 >= 80) {
        titulo = "Perdas altamente concentradas";
        leitura = `Os três maiores ofensores concentram ${formatarPercentualPareto(percentualTop3)} da perda total. ${principal.modo} lidera o período com ${formatarToneladas(principal.perda)} t e ${principal.ocorrencias} ocorrência(s). A priorização das ações nesses ofensores tende a capturar a maior parcela do impacto.`;
    } else if (percentualTop3 >= 60) {
        titulo = "Concentração moderada das perdas";
        leitura = `Os três maiores ofensores representam ${formatarPercentualPareto(percentualTop3)} da perda total. ${principal.modo} é o principal impacto, com ${formatarToneladas(principal.perda)} t. O Pareto mostra uma concentração relevante, mas ainda há contribuição dos demais modos de falha.`;
    } else {
        titulo = "Perdas distribuídas entre ofensores";
        leitura = `Os três maiores ofensores somam ${formatarPercentualPareto(percentualTop3)} da perda total. ${principal.modo} aparece na liderança com ${formatarToneladas(principal.perda)} t, porém a perda está mais distribuída entre diferentes modos de falha.`;
    }

    definirTexto("tituloResumoPareto", titulo);
    definirTexto("textoResumoPareto", leitura);
    definirTexto("perdaMediaOcorrencia", `${formatarToneladas(perdaMedia)} t`);
    definirTexto("quantidadeOfensores", ranking.length.toLocaleString("pt-BR"));
    definirTexto("acumuladoTerceiro", formatarPercentualPareto(terceiroAcumulado));
    definirTexto("leituraParetoGrafico", `Top 3 concentram ${formatarPercentualPareto(percentualTop3)} da perda`);
}

// ======================================================
// GRÁFICO
// ======================================================

function desenharGrafico(labels, valores, percentuais) {
    if (typeof Chart === "undefined") {
        alert("A biblioteca Chart.js não foi carregada.");
        return;
    }

    const canvas = document.getElementById("graficoPareto");
    if (!canvas) return;

    if (graficoAtual) {
        graficoAtual.destroy();
        graficoAtual = null;
    }

    const coresBarras = labels.map((label, indice) => {
        const paleta = ["#3b005f", "#59117d", "#742699", "#8b40aa", "#9d59b6", "#aa6ec0", "#b681ca", "#c296d1", "#cdb0da", "#d8c6e1"];
        return paleta[Math.min(indice, paleta.length - 1)];
    });

    graficoAtual = new Chart(canvas, {
        data: {
            labels,
            datasets: [
                {
                    type: "bar",
                    label: "Perda (t)",
                    data: valores,
                    backgroundColor: coresBarras,
                    borderColor: coresBarras,
                    borderWidth: 1,
                    borderRadius: 7,
                    borderSkipped: false,
                    barPercentage: 0.68,
                    categoryPercentage: 0.78,
                    maxBarThickness: 62,
                    yAxisID: "y",
                    order: 2
                },
                {
                    type: "line",
                    label: "% Acumulado",
                    data: percentuais,
                    borderColor: "#ee9622",
                    backgroundColor: "#ee9622",
                    borderWidth: 3.5,
                    tension: 0.22,
                    pointRadius: 5,
                    pointHoverRadius: 7.5,
                    clip: false,
                    pointBackgroundColor: "#ffffff",
                    pointBorderColor: "#ee9622",
                    pointBorderWidth: 2.5,
                    yAxisID: "y1",
                    order: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            devicePixelRatio: Math.max(2, Math.min(window.devicePixelRatio || 1, 3)),
            layout: {
                padding: { top: 34, right: 18, left: 8, bottom: 6 }
            },
            animation: { duration: 350 },
            interaction: { mode: "index", intersect: false },
            plugins: {
                legend: { display: false },
                datalabels: {
                    display: context => context.dataset.label === "Perda (t)",
                    anchor: "end",
                    align: "top",
                    offset: 5,
                    clamp: true,
                    clip: false,
                    color: "#2d014d",
                    font: { weight: "800", size: 11 },
                    formatter: value => `${formatarToneladas(value)} t`
                },
                tooltip: {
                    backgroundColor: "rgba(45,1,77,.96)",
                    titleFont: { size: 12, weight: "800" },
                    bodyFont: { size: 11, weight: "700" },
                    padding: 11,
                    callbacks: {
                        label: context => {
                            if (context.dataset.label === "Perda (t)") return ` Perda: ${formatarToneladas(context.raw)} t`;
                            return ` Acumulado: ${formatarPercentualPareto(context.raw)}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: { display: false },
                    ticks: {
                        color: "#554c58",
                        autoSkip: false,
                        maxRotation: 24,
                        minRotation: 0,
                        padding: 6,
                        font: { size: 10.5, weight: "700" },
                        callback: function(value) {
                            const label = this.getLabelForValue(value);
                            return quebrarRotuloPareto(label, 18);
                        }
                    }
                },
                y: {
                    beginAtZero: true,
                    grace: "15%",
                    title: {
                        display: true,
                        text: "Perda (t)",
                        color: "#514a55",
                        font: { size: 11, weight: "800" }
                    },
                    grid: { color: "rgba(81,74,85,.12)" },
                    ticks: {
                        color: "#625a65",
                        font: { size: 10, weight: "700" },
                        callback: valor => Number(valor).toLocaleString("pt-BR")
                    }
                },
                y1: {
                    position: "right",
                    min: 0,
                    // Pequena folga acima de 100% evita cortar a linha/pontos
                    // quando o acumulado atinge o topo do gráfico.
                    max: 105,
                    title: {
                        display: true,
                        text: "% Acumulado",
                        color: "#b4690b",
                        font: { size: 11, weight: "800" }
                    },
                    grid: { drawOnChartArea: false },
                    ticks: {
                        color: "#9a600e",
                        stepSize: 20,
                        font: { size: 10, weight: "800" },
                        callback: valor => `${valor}%`
                    }
                }
            }
        }
    });
}

function quebrarRotuloPareto(texto, limite = 18) {
    const palavras = String(texto ?? "").split(/\s+/).filter(Boolean);
    if (!palavras.length) return "";

    const linhas = [];
    let atual = "";

    palavras.forEach(palavra => {
        const candidato = atual ? `${atual} ${palavra}` : palavra;
        if (candidato.length <= limite || !atual) {
            atual = candidato;
        } else {
            linhas.push(atual);
            atual = palavra;
        }
    });

    if (atual) linhas.push(atual);
    return linhas.slice(0, 3);
}

// ======================================================
// PERÍODO / CABEÇALHO PDF
// ======================================================

function atualizarPeriodo(valorFiltro) {
    const elemento = document.getElementById("periodoPareto");
    if (!elemento) return;

    if (filtroAtivo === "dia") elemento.textContent = `Análise do dia ${formatarDataPareto(valorFiltro)}`;
    if (filtroAtivo === "mes") elemento.textContent = `Análise mensal - ${nomesMesesPareto[Number(valorFiltro)] || valorFiltro}`;
    if (filtroAtivo === "pilha") elemento.textContent = `Análise da Pilha ${valorFiltro}`;
}

function atualizarCabecalhoPDFPareto() {
    const periodo = obterTextoPareto("periodoPareto");
    const gerado = `Gerado em: ${new Date().toLocaleString("pt-BR")}`;

    definirTexto("periodoParetoPDF", `Período: ${periodo}`);
    definirTexto("geradoParetoPDF", gerado);
    definirTexto("periodoParetoPDF2", `Período: ${periodo}`);
    definirTexto("geradoParetoPDF2", gerado);
}

// ======================================================
// PDF / IMPRESSÃO
// ======================================================

function gerarPDFPareto() {
    if (!paretoGerado) {
        alert("Gere um Pareto antes de criar o PDF.");
        return;
    }

    atualizarCabecalhoPDFPareto();

    const tituloAnterior = document.title;
    const periodo = obterTextoPareto("periodoPareto");
    const nomeSeguro = periodo
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-zA-Z0-9]+/g, "_")
        .replace(/^_+|_+$/g, "");

    document.title = `Pareto_Britagem_FVO_${nomeSeguro || "Analise"}`;
    document.body.classList.add("modo-impressao-pareto");

    setTimeout(() => {
        window.print();
        setTimeout(() => {
            document.body.classList.remove("modo-impressao-pareto");
            document.title = tituloAnterior;
        }, 400);
    }, 350);
}

function prepararGraficoParaImpressaoPareto() {
    if (!graficoAtual) return;

    estadoGraficoAntesImpressao = {
        devicePixelRatio: graficoAtual.options.devicePixelRatio,
        xFont: graficoAtual.options.scales?.x?.ticks?.font?.size,
        yFont: graficoAtual.options.scales?.y?.ticks?.font?.size,
        y1Font: graficoAtual.options.scales?.y1?.ticks?.font?.size,
        datalabelFont: graficoAtual.options.plugins?.datalabels?.font?.size,
        linha: graficoAtual.data.datasets?.[1]?.borderWidth,
        ponto: graficoAtual.data.datasets?.[1]?.pointRadius,
        pontoBorda: graficoAtual.data.datasets?.[1]?.pointBorderWidth
    };

    graficoAtual.options.animation = false;
    // Canvas em alta densidade para evitar gráfico borrado no PDF.
    graficoAtual.options.devicePixelRatio = 4;
    graficoAtual.options.scales.x.ticks.font.size = 13;
    graficoAtual.options.scales.y.ticks.font.size = 11.5;
    graficoAtual.options.scales.y1.ticks.font.size = 11.5;
    graficoAtual.options.plugins.datalabels.font.size = 13;
    graficoAtual.data.datasets[1].borderWidth = 4.5;
    graficoAtual.data.datasets[1].pointRadius = 6;
    graficoAtual.data.datasets[1].pointBorderWidth = 2.8;
    graficoAtual.resize();
    graficoAtual.update("none");
}

function restaurarGraficoAposImpressaoPareto() {
    if (!graficoAtual || !estadoGraficoAntesImpressao) return;

    graficoAtual.options.devicePixelRatio = estadoGraficoAntesImpressao.devicePixelRatio || Math.max(2, window.devicePixelRatio || 1);
    graficoAtual.options.scales.x.ticks.font.size = estadoGraficoAntesImpressao.xFont || 10.5;
    graficoAtual.options.scales.y.ticks.font.size = estadoGraficoAntesImpressao.yFont || 10;
    graficoAtual.options.scales.y1.ticks.font.size = estadoGraficoAntesImpressao.y1Font || 10;
    graficoAtual.options.plugins.datalabels.font.size = estadoGraficoAntesImpressao.datalabelFont || 11;
    graficoAtual.data.datasets[1].borderWidth = estadoGraficoAntesImpressao.linha || 3.5;
    graficoAtual.data.datasets[1].pointRadius = estadoGraficoAntesImpressao.ponto || 5;
    graficoAtual.data.datasets[1].pointBorderWidth = estadoGraficoAntesImpressao.pontoBorda || 2.5;
    graficoAtual.resize();
    graficoAtual.update("none");
    estadoGraficoAntesImpressao = null;
}

// ======================================================
// E-MAIL
// ======================================================

function enviarEmailPareto() {
    if (!paretoGerado) {
        alert("Gere um Pareto antes de enviar.");
        return;
    }

    const periodo = obterTextoPareto("periodoPareto");
    const perda = obterTextoPareto("perdaTotal");
    const principal = obterTextoPareto("principalOfensor");
    const impacto = obterTextoPareto("impactoPrincipal");
    const top3 = obterTextoPareto("impactoTop3");
    const ocorrencias = obterTextoPareto("totalOcorrencias");
    const tipo = filtroAtivo === "dia" ? "Dia" : filtroAtivo === "mes" ? "Mês" : "Pilha";

    const assunto = `Pareto de Perdas - Britagem FVO - ${periodo}`;
    const mensagem = `Prezados,

Segue o acompanhamento do Pareto de Perdas da Britagem FVO.

Tipo de análise: ${tipo}
Período: ${periodo}

Perda Total: ${perda} t
Principal Ofensor: ${principal}
Impacto do Principal Ofensor: ${impacto}
Impacto dos Top 3 Ofensores: ${top3}
Ocorrências consideradas: ${ocorrencias}

O relatório completo pode ser gerado pelo Portal FVO.

Atenciosamente.`;

    const destinatarios = emailsPareto.join(";");
    window.location.href = `mailto:${destinatarios}?subject=${encodeURIComponent(assunto)}&body=${encodeURIComponent(mensagem)}`;
}

// ======================================================
// LIMPEZA / STATUS
// ======================================================

function limparDashboard() {
    [
        ["perdaTotal", "-"],
        ["principalOfensor", "-"],
        ["impactoPrincipal", "-"],
        ["impactoTop3", "-"],
        ["totalOcorrencias", "-"],
        ["perdaMediaOcorrencia", "-"],
        ["quantidadeOfensores", "-"],
        ["acumuladoTerceiro", "-"],
        ["leituraParetoGrafico", "Aguardando análise"],
        ["tituloResumoPareto", "Aguardando análise"],
        ["textoResumoPareto", "Nenhum dado encontrado para o filtro selecionado."],
        ["periodoPareto", "Nenhum dado encontrado."]
    ].forEach(([id, texto]) => definirTexto(id, texto));

    const corpo = document.getElementById("rankingParetoBody");
    if (corpo) {
        corpo.innerHTML = `<tr class="linha-vazia-pareto"><td colspan="6">Nenhum dado encontrado para o filtro selecionado.</td></tr>`;
    }

    if (graficoAtual) {
        graficoAtual.destroy();
        graficoAtual = null;
    }

    rankingAtual = [];
    ocorrenciasAtual = 0;
    perdaTotalAtual = 0;
    paretoGerado = false;
}

function atualizarStatusBase(texto) {
    definirTexto("statusBasePareto", texto);
}

// ======================================================
// UTILITÁRIOS
// ======================================================

function converterDataPareto(valor) {
    if (valor === null || valor === undefined || valor === "") return null;

    if (typeof valor === "number") {
        const data = XLSX.SSF.parse_date_code(valor);
        if (!data) return null;
        return montarObjetoDataPareto(data.y, data.m, data.d);
    }

    if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
        return montarObjetoDataPareto(valor.getFullYear(), valor.getMonth() + 1, valor.getDate());
    }

    const texto = String(valor).trim().split(/[ T]/)[0];
    const partes = texto.split(/[\/\-.]/);
    if (partes.length !== 3) return null;

    let ano;
    let mes;
    let dia;

    if (partes[0].length === 4) {
        ano = Number(partes[0]);
        mes = Number(partes[1]);
        dia = Number(partes[2]);
    } else {
        dia = Number(partes[0]);
        mes = Number(partes[1]);
        ano = Number(partes[2]);
    }

    if (ano < 100) ano += 2000;
    if (!ano || !mes || !dia || mes < 1 || mes > 12 || dia < 1 || dia > 31) return null;

    return montarObjetoDataPareto(ano, mes, dia);
}

function montarObjetoDataPareto(ano, mes, dia) {
    return {
        ano,
        mes,
        dia,
        texto: `${ano}-${String(mes).padStart(2, "0")}-${String(dia).padStart(2, "0")}`
    };
}

function numeroPareto(valor) {
    if (valor === null || valor === undefined || valor === "") return null;
    if (typeof valor === "number") return Number.isFinite(valor) ? valor : null;

    let texto = String(valor).trim().replace(/\s/g, "");

    if (texto.includes(",") && texto.includes(".")) {
        if (texto.lastIndexOf(",") > texto.lastIndexOf(".")) {
            texto = texto.replace(/\./g, "").replace(",", ".");
        } else {
            texto = texto.replace(/,/g, "");
        }
    } else if (texto.includes(",")) {
        texto = texto.replace(",", ".");
    }

    const numero = Number(texto);
    return Number.isFinite(numero) ? numero : null;
}

function textoPareto(valor) {
    return valor === null || valor === undefined ? "" : String(valor).trim();
}

function formatarDataPareto(valor) {
    const partes = String(valor).split("-");
    return partes.length === 3 ? `${partes[2]}/${partes[1]}/${partes[0]}` : valor;
}

function formatarToneladas(valor) {
    return Number(valor || 0).toLocaleString("pt-BR", {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    });
}

function formatarPercentualPareto(valor) {
    return `${Number(valor || 0).toFixed(1).replace(".", ",")}%`;
}

function obterTextoPareto(id) {
    return document.getElementById(id)?.textContent?.trim() || "-";
}

function definirTexto(id, texto) {
    const elemento = document.getElementById(id);
    if (elemento) elemento.textContent = texto;
}

function escaparHTMLPareto(valor) {
    return String(valor ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}