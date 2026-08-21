"use strict";

/* ============================================================
   CONCILIAÇÃO DE MASSA — BRITAGEM FVO
   Massa da Britagem: Análise de Pilhas / aba "Pilha"
   Viagens: Relatório de Turno / primeiro bloco de formação da pilha
============================================================ */

let dadosMassaPilhas = [];
let mapaMassaPilhas = new Map();
let mapaRelatorioTurno = new Map();
let workbookRelatorioTurno = null;
let pilhaAtualConciliacao = null;
let dadosAtuaisConciliacao = null;
let graficoComparacaoMassa = null;
let graficoFatorCacamba = null;
let configuracaoGraficoConciliacao = null;

const FATOR_CACAMBA_PADRAO = 50;
const CHAVE_FATOR_LOCAL = "fvo_fator_cacamba_referencia";

/* ============================================================
   INICIALIZAÇÃO
============================================================ */

document.addEventListener("DOMContentLoaded", () => {
    const fatorSalvo = Number(localStorage.getItem(CHAVE_FATOR_LOCAL));
    const inputFator = document.getElementById("fatorCacamba");

    if(inputFator && Number.isFinite(fatorSalvo) && fatorSalvo > 0){
        inputFator.value = fatorSalvo;
    }

    document.getElementById("btnCarregarBases")?.addEventListener("click", carregarBasesConciliacao);
    document.getElementById("btnAnalisarConciliacao")?.addEventListener("click", consultarConciliacao);
    document.getElementById("btnGerarPDF")?.addEventListener("click", gerarPDFConciliacao);

    inputFator?.addEventListener("input", () => {
        const fator = obterFatorReferencia();
        if(fator !== null){
            localStorage.setItem(CHAVE_FATOR_LOCAL, String(fator));
            atualizarFatorNaTela(fator);

            if(dadosAtuaisConciliacao){
                recalcularComNovoFator();
            }else if(dadosMassaPilhas.length){
                renderizarGraficoHistoricoFator(null);
            }
        }
    });

    document.getElementById("numeroPilhaConciliacao")?.addEventListener("keydown", (evento) => {
        if(evento.key === "Enter"){
            consultarConciliacao();
        }
    });

    atualizarFatorNaTela(obterFatorReferencia() ?? FATOR_CACAMBA_PADRAO);
});

/* ============================================================
   CARREGAMENTO DOS ARQUIVOS
============================================================ */

async function carregarBasesConciliacao(){
    if(typeof XLSX === "undefined"){
        alert("A biblioteca de leitura do Excel não foi carregada.");
        return;
    }

    const arquivoMassa = document.getElementById("arquivoAnalisePilhas")?.files?.[0];
    const arquivoTurno = document.getElementById("arquivoRelatorioTurno")?.files?.[0];

    if(!arquivoMassa || !arquivoTurno){
        atualizarStatusBases("Selecione as duas planilhas", "erro");
        alert("Selecione a planilha Análise de Pilhas e o Relatório de Turno.");
        return;
    }

    atualizarStatusBases("Lendo arquivos...", "carregando");
    window.FVOCarregamento?.exibir("Lendo massa da Britagem e viagens das pilhas...");

    try{
        const [workbookMassa, workbookTurno] = await Promise.all([
            lerArquivoExcelConciliacao(arquivoMassa),
            lerArquivoExcelConciliacao(arquivoTurno)
        ]);

        prepararBaseMassa(workbookMassa);
        prepararBaseRelatorioTurno(workbookTurno);
        preencherListaPilhas();

        atualizarStatusBases(`${mapaMassaPilhas.size} pilhas com massa • ${mapaRelatorioTurno.size} com viagens`, "ok");
        atualizarDataSistemaConciliacao();
        renderizarGraficoHistoricoFator(null);

        window.FVOCarregamento?.ocultar();
    }catch(erro){
        console.error("Erro ao carregar bases da conciliação:", erro);
        atualizarStatusBases("Erro ao ler as bases", "erro");
        window.FVOCarregamento?.ocultar();
        alert(erro?.message || "Não foi possível ler as planilhas.");
    }
}

function lerArquivoExcelConciliacao(arquivo){
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (evento) => {
            try{
                const dados = new Uint8Array(evento.target.result);
                resolve(XLSX.read(dados, {type:"array", cellDates:true}));
            }catch(erro){
                reject(erro);
            }
        };
        reader.onerror = () => reject(new Error(`Falha ao ler ${arquivo.name}.`));
        reader.readAsArrayBuffer(arquivo);
    });
}

/* ============================================================
   BASE 1 — ANÁLISE DE PILHAS / MASSA DA BRITAGEM
============================================================ */

function prepararBaseMassa(workbook){
    const abaPilha = workbook.Sheets?.["Pilha"];
    if(!abaPilha){
        throw new Error("A aba 'Pilha' não foi encontrada na planilha Análise de Pilhas.");
    }

    const linhas = XLSX.utils.sheet_to_json(abaPilha, {
        header:1,
        defval:null,
        raw:true
    });

    dadosMassaPilhas = [];
    mapaMassaPilhas = new Map();

    for(let i = 0; i < linhas.length; i++){
        const linha = linhas[i];
        if(!linha){
            continue;
        }

        const numero = normalizarNumeroPilha(linha[0]);
        const massa = valorNumeroConciliacao(linha[18]);

        if(numero === null || massa === null || massa <= 0){
            continue;
        }

        const registro = {
            numero,
            massa,
            inicio: converterDataConciliacao(linha[21]),
            fim: converterDataConciliacao(linha[22]),
            indice:i,
            valorOriginal:linha[0]
        };

        dadosMassaPilhas.push(registro);
        mapaMassaPilhas.set(numero, registro);
    }

    if(!dadosMassaPilhas.length){
        throw new Error("Nenhuma pilha com massa válida foi encontrada na aba 'Pilha'.");
    }
}

/* ============================================================
   BASE 2 — RELATÓRIO DE TURNO / VIAGENS
============================================================ */

function prepararBaseRelatorioTurno(workbook){
    workbookRelatorioTurno = workbook;
    mapaRelatorioTurno = new Map();

    workbook.SheetNames.forEach((nomeAba, ordem) => {
        if(!/pilha/i.test(nomeAba)){
            return;
        }

        const numero = normalizarNumeroPilha(nomeAba);
        if(numero === null){
            return;
        }

        const sheet = workbook.Sheets[nomeAba];
        const linhas = XLSX.utils.sheet_to_json(sheet, {
            header:1,
            defval:null,
            raw:true
        });

        const formacao = extrairFormacaoPrincipal(linhas, nomeAba);
        if(!formacao){
            return;
        }

        formacao.ordemAba = ordem;

        // Se houver pilha antiga e uma aba mais recente com o mesmo número,
        // a última aba do arquivo substitui a anterior.
        mapaRelatorioTurno.set(numero, formacao);
    });

    if(!mapaRelatorioTurno.size){
        throw new Error("Não foram encontrados blocos de formação de pilhas no Relatório de Turno.");
    }
}

function extrairFormacaoPrincipal(linhas, nomeAba){
    let linhaTitulo = -1;

    for(let i = 0; i < linhas.length; i++){
        const textoLinha = (linhas[i] || []).map(valor => normalizarTexto(valor)).join(" ");

        if(
            textoLinha.includes("controle total da formacao da pilha") ||
            textoLinha.includes("controle total da formacao pilha")
        ){
            linhaTitulo = i;
            break;
        }
    }

    if(linhaTitulo < 0){
        return null;
    }

    let linhaCabecalho = -1;
    for(let i = linhaTitulo + 1; i < Math.min(linhas.length, linhaTitulo + 5); i++){
        const normalizados = (linhas[i] || []).map(normalizarTexto);
        if(normalizados.some(v => v === "sequencia") && normalizados.some(v => v.includes("viagens"))){
            linhaCabecalho = i;
            break;
        }
    }

    if(linhaCabecalho < 0){
        return null;
    }

    const cabecalho = (linhas[linhaCabecalho] || []).map(normalizarTexto);

    const colunas = {
        sequencia: indiceCabecalho(cabecalho, ["sequencia"]),
        poligono: indiceCabecalho(cabecalho, ["poligono"]),
        area: indiceCabecalho(cabecalho, ["area"]),
        banco: indiceCabecalho(cabecalho, ["banco"]),
        programadas: indiceCabecalho(cabecalho, ["n viagens programadas", "viagens programadas"]),
        planejadas: indiceCabecalho(cabecalho, ["n viagens planejadas", "viagens planejadas"]),
        realizadas: indiceCabecalho(cabecalho, ["n viagens realizadas", "viagens realizadas"]),
        aderencia: indiceCabecalho(cabecalho, ["aderencia", "cumprimento"]),
        saldo: indiceCabecalho(cabecalho, ["saldo"]),
        comentario: indiceCabecalho(cabecalho, ["comentario"]),
        causa: indiceCabecalho(cabecalho, ["causa desvio do plano", "causa"])
    };

    const colunaPrograma = colunas.programadas >= 0 ? colunas.programadas : colunas.planejadas;
    if(colunaPrograma < 0 || colunas.realizadas < 0){
        return null;
    }

    const detalhes = [];
    let totalProgramadas = null;
    let totalRealizadas = null;
    let aderencia = null;
    let saldo = null;

    for(let i = linhaCabecalho + 1; i < linhas.length; i++){
        const linha = linhas[i] || [];
        const sequenciaTexto = normalizarTexto(linha[colunas.sequencia]);

        if(sequenciaTexto === "total"){
            totalProgramadas = valorNumeroConciliacao(linha[colunaPrograma]);
            totalRealizadas = valorNumeroConciliacao(linha[colunas.realizadas]);
            aderencia = colunas.aderencia >= 0 ? valorNumeroConciliacao(linha[colunas.aderencia]) : null;
            saldo = colunas.saldo >= 0 ? valorNumeroConciliacao(linha[colunas.saldo]) : null;
            break;
        }

        if(!sequenciaTexto){
            // Uma linha em branco normalmente encerra o primeiro bloco.
            const proximaTemConteudo = (linhas[i + 1] || []).some(v => v !== null && v !== "");
            if(!proximaTemConteudo){
                break;
            }
            continue;
        }

        if(sequenciaTexto === "sequencia"){
            break;
        }

        const realizadas = valorNumeroConciliacao(linha[colunas.realizadas]);
        const programadas = valorNumeroConciliacao(linha[colunaPrograma]);

        if(realizadas === null && programadas === null){
            continue;
        }

        detalhes.push({
            sequencia: valorTexto(linha[colunas.sequencia]),
            poligono: colunas.poligono >= 0 ? valorTexto(linha[colunas.poligono]) : "-",
            area: colunas.area >= 0 ? valorTexto(linha[colunas.area]) : "-",
            banco: colunas.banco >= 0 ? valorTexto(linha[colunas.banco]) : "-",
            programadas: programadas ?? 0,
            realizadas: realizadas ?? 0,
            aderencia: colunas.aderencia >= 0 ? valorNumeroConciliacao(linha[colunas.aderencia]) : null,
            comentario: colunas.comentario >= 0 ? valorTexto(linha[colunas.comentario]) : "",
            causa: colunas.causa >= 0 ? valorTexto(linha[colunas.causa]) : ""
        });
    }

    if(totalProgramadas === null){
        totalProgramadas = detalhes.reduce((soma, item) => soma + (item.programadas || 0), 0);
    }
    if(totalRealizadas === null){
        totalRealizadas = detalhes.reduce((soma, item) => soma + (item.realizadas || 0), 0);
    }
    if(aderencia === null && totalProgramadas > 0){
        aderencia = totalRealizadas / totalProgramadas;
    }
    if(saldo === null){
        saldo = totalProgramadas - totalRealizadas;
    }

    if(!Number.isFinite(totalRealizadas) || totalRealizadas <= 0){
        return null;
    }

    return {
        nomeAba,
        titulo: valorTexto((linhas[linhaTitulo] || []).find(v => v !== null && v !== "")),
        programadas:totalProgramadas,
        realizadas:totalRealizadas,
        aderencia,
        saldo,
        detalhes
    };
}

function indiceCabecalho(cabecalho, candidatos){
    for(let i = 0; i < cabecalho.length; i++){
        const atual = cabecalho[i];
        if(!atual){
            continue;
        }

        if(candidatos.some(candidato => atual === candidato || atual.includes(candidato))){
            return i;
        }
    }
    return -1;
}

/* ============================================================
   CONSULTA
============================================================ */

function consultarConciliacao(){
    if(!mapaMassaPilhas.size || !mapaRelatorioTurno.size){
        alert("Carregue as duas bases antes de consultar a pilha.");
        return;
    }

    const numero = normalizarNumeroPilha(document.getElementById("numeroPilhaConciliacao")?.value);
    if(numero === null){
        alert("Informe um número de pilha válido.");
        return;
    }

    const massa = mapaMassaPilhas.get(numero);
    const viagens = mapaRelatorioTurno.get(numero);

    if(!massa){
        alert(`A Pilha ${numero} não foi encontrada na planilha Análise de Pilhas.`);
        return;
    }

    if(!viagens){
        alert(`A Pilha ${numero} não possui um bloco principal de formação válido no Relatório de Turno.`);
        return;
    }

    const fator = obterFatorReferencia();
    if(fator === null){
        alert("Informe um Fator Caçamba maior que zero.");
        return;
    }

    pilhaAtualConciliacao = numero;
    dadosAtuaisConciliacao = {massa, viagens};
    recalcularComNovoFator();
    atualizarDataSistemaConciliacao();
}

function recalcularComNovoFator(){
    if(!dadosAtuaisConciliacao){
        return;
    }

    const fator = obterFatorReferencia();
    if(fator === null){
        return;
    }

    const {massa, viagens} = dadosAtuaisConciliacao;
    const massaEstimada = viagens.realizadas * fator;
    const cargaMediaReal = viagens.realizadas > 0 ? massa.massa / viagens.realizadas : null;
    const diferenca = massa.massa - massaEstimada;
    const diferencaPercentual = massaEstimada !== 0 ? (diferenca / massaEstimada) * 100 : null;

    const resultado = {
        numero:pilhaAtualConciliacao,
        massa,
        viagens,
        fator,
        massaEstimada,
        cargaMediaReal,
        diferenca,
        diferencaPercentual
    };

    preencherCards(resultado);
    preencherResumo(resultado);
    preencherTabelaFormacao(viagens.detalhes);
    preencherCriterios(resultado);
    renderizarGraficoComparacao(resultado);
    renderizarGraficoHistoricoFator(pilhaAtualConciliacao);
}

/* ============================================================
   PREENCHIMENTO DA TELA
============================================================ */

function preencherCards(resultado){
    setTexto("cardPilha", resultado.numero);
    setTexto("cardLadoPilha", resultado.viagens.nomeAba || "Pilha selecionada");
    setTexto("cardMassaBalanca", `${formatarToneladas(resultado.massa.massa)} t`);
    setTexto("cardViagensRealizadas", formatarInteiro(resultado.viagens.realizadas));
    setTexto("cardViagensProgramadas", `Programadas: ${formatarInteiro(resultado.viagens.programadas)}`);
    setTexto("cardCargaMediaReal", `${formatarDecimal(resultado.cargaMediaReal, 2)} t/viagem`);
    const desvioFator = resultado.cargaMediaReal - resultado.fator;
    setTexto(
        "cardCargaMediaContexto",
        `Referência: ${formatarDecimal(resultado.fator,1)} • ${desvioFator >= 0 ? "+" : ""}${formatarDecimal(desvioFator,2)} t/viagem`
    );
    setTexto("cardMassaEstimada", `${formatarToneladas(resultado.massaEstimada)} t`);
    setTexto("cardFatorReferencia", `Fator: ${formatarDecimal(resultado.fator, 1)} t/viagem`);
    setTexto("cardDiferenca", `${resultado.diferenca >= 0 ? "+" : ""}${formatarToneladas(resultado.diferenca)} t`);
    setTexto("cardDiferencaPercentual", `${resultado.diferencaPercentual >= 0 ? "+" : ""}${formatarDecimal(resultado.diferencaPercentual, 2)}% vs. viagens`);

    const cardDiferenca = document.querySelector(".card-diferenca");
    if(cardDiferenca){
        cardDiferenca.dataset.sinal = resultado.diferenca >= 0 ? "positivo" : "negativo";
    }

    setTexto("pdfPilhaConciliacao", `Pilha: ${resultado.numero}`);
}

function preencherResumo(resultado){
    const aderencia = resultado.viagens.aderencia;
    setTexto("resumoAderencia", aderencia === null ? "N/D" : formatarPercentualFracao(aderencia));
    const saldo = resultado.viagens.saldo;
    const textoSaldo = saldo === 0
        ? "0 • conforme programado"
        : saldo < 0
            ? `${formatarInteiro(Math.abs(saldo))} acima do programado`
            : `${formatarInteiro(saldo)} abaixo do programado`;
    setTexto("resumoSaldo", textoSaldo);

    const periodo = montarPeriodo(resultado.massa, resultado.viagens);
    setTexto("resumoPeriodo", periodo);

    const direcao = resultado.diferenca >= 0 ? "acima" : "abaixo";
    setTexto(
        "resumoLeitura",
        `Balança ${Math.abs(resultado.diferencaPercentual).toLocaleString("pt-BR", {minimumFractionDigits:2, maximumFractionDigits:2})}% ${direcao} da estimativa por viagens.`
    );
}

function preencherTabelaFormacao(detalhes){
    const corpo = document.getElementById("corpoTabelaFormacao");
    if(!corpo){
        return;
    }

    if(!detalhes?.length){
        corpo.innerHTML = '<tr><td colspan="7" class="sem-dados-tabela">Sem detalhamento de frentes para a pilha selecionada.</td></tr>';
        return;
    }

    corpo.innerHTML = detalhes.map(item => {
        const aderencia = item.aderencia !== null
            ? formatarPercentualFracao(item.aderencia)
            : (item.programadas > 0 ? formatarPercentualFracao(item.realizadas / item.programadas) : "N/D");

        return `
            <tr>
                <td>${escaparHTML(item.sequencia)}</td>
                <td>${escaparHTML(item.poligono)}</td>
                <td>${escaparHTML(item.area)}</td>
                <td>${escaparHTML(item.banco)}</td>
                <td>${formatarInteiro(item.programadas)}</td>
                <td><strong>${formatarInteiro(item.realizadas)}</strong></td>
                <td>${aderencia}</td>
            </tr>
        `;
    }).join("");
}

function preencherCriterios(resultado){
    setTexto("criterioAbaTurno", resultado.viagens.nomeAba);
    setTexto("criterioFator", `${formatarDecimal(resultado.fator, 1)} t/viagem`);
}

function atualizarFatorNaTela(fator){
    setTexto("badgeFatorComparacao", `${formatarDecimal(fator, 1)} t/viagem`);
    setTexto("badgeFatorHistorico", `Referência: ${formatarDecimal(fator, 1)}`);
    setTexto("criterioFator", `${formatarDecimal(fator, 1)} t/viagem`);
}

/* ============================================================
   GRÁFICO 1 — MASSA BALANÇA × VIAGENS
============================================================ */

function renderizarGraficoComparacao(resultado){
    if(typeof Chart === "undefined"){
        return;
    }

    const canvas = document.getElementById("graficoComparacaoMassa");
    if(!canvas){
        return;
    }

    graficoComparacaoMassa?.destroy();

    graficoComparacaoMassa = new Chart(canvas, {
        type:"bar",
        data:{
            labels:["Balança Britagem", "Viagens × Fator"],
            datasets:[{
                label:"Massa (t)",
                data:[resultado.massa.massa, resultado.massaEstimada],
                backgroundColor:["#3B005F", "#92D050"],
                borderRadius:10,
                borderSkipped:false,
                maxBarThickness:100
            }]
        },
        options:{
            responsive:true,
            maintainAspectRatio:false,
            devicePixelRatio:Math.max(window.devicePixelRatio || 1, 2),
            layout:{padding:{top:24,right:18,left:8,bottom:0}},
            plugins:{
                legend:{display:false},
                tooltip:{
                    callbacks:{
                        label:(ctx) => ` ${formatarToneladas(ctx.raw)} t`
                    }
                },
                datalabels:{
                    anchor:"end",
                    align:"end",
                    color:"#2A063D",
                    font:{size:13,weight:"800"},
                    formatter:(valor) => `${formatarToneladas(valor)} t`
                }
            },
            scales:{
                x:{
                    grid:{display:false},
                    ticks:{color:"#514658",font:{size:12,weight:"700"}}
                },
                y:{
                    beginAtZero:true,
                    grace:"12%",
                    grid:{color:"rgba(59,0,95,.08)"},
                    ticks:{
                        color:"#6B5D70",
                        font:{size:10},
                        callback:(valor) => formatarEixoToneladas(valor)
                    },
                    title:{display:true,text:"Massa (t)",color:"#6B5D70",font:{size:10,weight:"700"}}
                }
            }
        },
        plugins: typeof ChartDataLabels !== "undefined" ? [ChartDataLabels] : []
    });
}

/* ============================================================
   GRÁFICO 2 — HISTÓRICO DO FATOR CAÇAMBA
============================================================ */

function renderizarGraficoHistoricoFator(numeroSelecionado){
    if(typeof Chart === "undefined" || !dadosMassaPilhas.length || !mapaRelatorioTurno.size){
        return;
    }

    const fatorReferencia = obterFatorReferencia() ?? FATOR_CACAMBA_PADRAO;

    let historico = dadosMassaPilhas
        .map(itemMassa => {
            const turno = mapaRelatorioTurno.get(itemMassa.numero);
            if(!turno || !turno.realizadas || turno.realizadas <= 0){
                return null;
            }

            return {
                numero:itemMassa.numero,
                massa:itemMassa.massa,
                viagens:turno.realizadas,
                fator:itemMassa.massa / turno.realizadas,
                data:itemMassa.fim || itemMassa.inicio
            };
        })
        .filter(Boolean)
        .sort((a,b) => a.numero - b.numero);

    if(numeroSelecionado !== null && numeroSelecionado !== undefined){
        const indice = historico.findIndex(item => item.numero === numeroSelecionado);
        if(indice >= 0){
            const inicio = Math.max(0, indice - 29);
            historico = historico.slice(inicio, indice + 1);
        }else{
            historico = historico.slice(-30);
        }
    }else{
        historico = historico.slice(-30);
    }

    const canvas = document.getElementById("graficoFatorCacamba");
    if(!canvas){
        return;
    }

    graficoFatorCacamba?.destroy();

    const labels = historico.map(item => `P${item.numero}`);
    const fatores = historico.map(item => item.fator);
    const referencia = historico.map(() => fatorReferencia);

    const coresPontos = historico.map(item => item.numero === numeroSelecionado ? "#92D050" : "#3B005F");
    const raiosPontos = historico.map(item => item.numero === numeroSelecionado ? 6 : 3.5);

    graficoFatorCacamba = new Chart(canvas, {
        type:"line",
        data:{
            labels,
            datasets:[
                {
                    label:"Carga média real",
                    data:fatores,
                    borderColor:"#3B005F",
                    backgroundColor:"rgba(59,0,95,.08)",
                    borderWidth:2.5,
                    tension:.18,
                    pointBackgroundColor:coresPontos,
                    pointBorderColor:coresPontos,
                    pointRadius:raiosPontos,
                    pointHoverRadius:7,
                    fill:false,
                    order:1
                },
                {
                    label:`Referência (${formatarDecimal(fatorReferencia,1)} t/viagem)`,
                    data:referencia,
                    borderColor:"#E7463F",
                    backgroundColor:"#E7463F",
                    borderWidth:2,
                    pointRadius:0,
                    tension:0,
                    borderDash:[],
                    order:2
                }
            ]
        },
        options:{
            responsive:true,
            maintainAspectRatio:false,
            devicePixelRatio:Math.max(window.devicePixelRatio || 1, 2),
            interaction:{mode:"index",intersect:false},
            layout:{padding:{top:8,right:10,left:5,bottom:0}},
            plugins:{
                legend:{
                    position:"top",
                    align:"end",
                    labels:{usePointStyle:true,boxWidth:9,color:"#564A5C",font:{size:10,weight:"700"}}
                },
                tooltip:{
                    callbacks:{
                        title:(itens) => {
                            const indice = itens?.[0]?.dataIndex ?? 0;
                            const item = historico[indice];
                            const data = item?.data ? ` • ${formatarDataCurtaConciliacao(item.data)}` : "";
                            return `Pilha ${item?.numero ?? "-"}${data}`;
                        },
                        label:(ctx) => {
                            if(ctx.datasetIndex === 1){
                                return ` Referência: ${formatarDecimal(ctx.raw,2)} t/viagem`;
                            }
                            const item = historico[ctx.dataIndex];
                            return [
                                ` Carga média: ${formatarDecimal(ctx.raw,2)} t/viagem`,
                                ` Massa: ${formatarToneladas(item.massa)} t`,
                                ` Viagens: ${formatarInteiro(item.viagens)}`
                            ];
                        }
                    }
                },
                datalabels:{
                    display:(ctx) => ctx.datasetIndex === 0 && historico[ctx.dataIndex]?.numero === numeroSelecionado,
                    align:"top",
                    anchor:"end",
                    color:"#3B005F",
                    font:{size:10,weight:"800"},
                    formatter:(valor) => formatarDecimal(valor,2)
                }
            },
            scales:{
                x:{
                    grid:{display:false},
                    ticks:{color:"#695C6F",font:{size:9,weight:"700"},maxRotation:0,autoSkip:true,maxTicksLimit:15},
                    title:{display:true,text:"Pilha",color:"#74687A",font:{size:9,weight:"700"}}
                },
                y:{
                    grace:"12%",
                    grid:{color:"rgba(59,0,95,.08)"},
                    ticks:{color:"#695C6F",font:{size:9},callback:(valor) => Number(valor).toLocaleString("pt-BR")},
                    title:{display:true,text:"Fator Caçamba (t/viagem)",color:"#74687A",font:{size:9,weight:"700"}}
                }
            }
        },
        plugins: typeof ChartDataLabels !== "undefined" ? [ChartDataLabels] : []
    });

    setTexto("amostraHistoricoFator", `${historico.length} pilha${historico.length === 1 ? "" : "s"}`);
}

/* ============================================================
   PDF
============================================================ */

function gerarPDFConciliacao(){
    if(!dadosAtuaisConciliacao){
        alert("Consulte uma pilha antes de gerar o PDF.");
        return;
    }

    prepararGraficosImpressao();
    const agora = new Date();
    const texto = `Gerado em: ${agora.toLocaleString("pt-BR")}`;
    setTexto("pdfGeradoConciliacao", texto);
    setTexto("rodapeGeracaoConciliacao", texto);

    setTimeout(() => window.print(), 120);
}

function prepararGraficosImpressao(){
    configuracaoGraficoConciliacao = [];

    [graficoComparacaoMassa, graficoFatorCacamba].forEach(grafico => {
        if(!grafico){
            return;
        }

        configuracaoGraficoConciliacao.push({
            grafico,
            devicePixelRatio:grafico.options.devicePixelRatio
        });

        grafico.options.animation = false;
        grafico.options.devicePixelRatio = 4;
        grafico.resize();
        grafico.update("none");
    });
}

window.addEventListener("afterprint", () => {
    (configuracaoGraficoConciliacao || []).forEach(item => {
        item.grafico.options.devicePixelRatio = item.devicePixelRatio || 2;
        item.grafico.resize();
        item.grafico.update("none");
    });
    configuracaoGraficoConciliacao = null;
});

/* ============================================================
   UTILITÁRIOS
============================================================ */

function preencherListaPilhas(){
    const lista = document.getElementById("listaPilhasConciliacao");
    if(!lista){
        return;
    }

    const numeros = [...mapaMassaPilhas.keys()]
        .filter(numero => mapaRelatorioTurno.has(numero))
        .sort((a,b) => a - b);

    lista.innerHTML = numeros.map(numero => `<option value="${numero}"></option>`).join("");
}

function obterFatorReferencia(){
    const valor = valorNumeroConciliacao(document.getElementById("fatorCacamba")?.value);
    return valor !== null && valor > 0 ? valor : null;
}

function atualizarStatusBases(texto, estado="neutro"){
    const elemento = document.getElementById("statusBasesConciliacao");
    if(!elemento){
        return;
    }
    elemento.textContent = texto;
    elemento.dataset.estado = estado;
}

function atualizarDataSistemaConciliacao(){
    const agora = new Date();
    setTexto("ultimaAtualizacaoConciliacao", agora.toLocaleString("pt-BR", {dateStyle:"short", timeStyle:"short"}));
}

function normalizarNumeroPilha(valor){
    if(valor === null || valor === undefined || valor === ""){
        return null;
    }
    const encontrado = String(valor).match(/\d+/);
    if(!encontrado){
        return null;
    }
    const numero = Number.parseInt(encontrado[0], 10);
    return Number.isFinite(numero) ? numero : null;
}

function normalizarTexto(valor){
    if(valor === null || valor === undefined){
        return "";
    }
    return String(valor)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[º°]/g, "")
        .replace(/[^a-zA-Z0-9]+/g, " ")
        .trim()
        .toLowerCase();
}

function valorNumeroConciliacao(valor){
    if(valor === null || valor === undefined || valor === ""){
        return null;
    }
    if(typeof valor === "number"){
        return Number.isFinite(valor) ? valor : null;
    }
    const texto = String(valor).trim();
    if(!texto || texto.startsWith("#")){
        return null;
    }
    const normalizado = texto
        .replace(/\s/g, "")
        .replace(/\.(?=\d{3}(?:\D|$))/g, "")
        .replace(",", ".");
    const numero = Number(normalizado);
    return Number.isFinite(numero) ? numero : null;
}

function valorTexto(valor){
    if(valor === null || valor === undefined || valor === ""){
        return "-";
    }
    return String(valor).trim();
}

function converterDataConciliacao(valor){
    if(!valor){
        return null;
    }
    if(valor instanceof Date && !Number.isNaN(valor.getTime())){
        return valor;
    }
    if(typeof valor === "number" && typeof XLSX !== "undefined"){
        const partes = XLSX.SSF?.parse_date_code?.(valor);
        if(partes){
            return new Date(partes.y, partes.m - 1, partes.d, partes.H || 0, partes.M || 0, partes.S || 0);
        }
    }
    const data = new Date(valor);
    return Number.isNaN(data.getTime()) ? null : data;
}

function montarPeriodo(massa, viagens){
    const inicio = massa.inicio ? formatarDataCurtaConciliacao(massa.inicio) : "";
    const fim = massa.fim ? formatarDataCurtaConciliacao(massa.fim) : "";
    if(inicio && fim){
        return `${inicio} → ${fim}`;
    }
    if(viagens.titulo){
        return viagens.titulo.replace(/controle total da formação da pilha/i, "Formação");
    }
    return viagens.nomeAba || "-";
}

function formatarDataCurtaConciliacao(data){
    if(!(data instanceof Date) || Number.isNaN(data.getTime())){
        return "-";
    }
    return data.toLocaleDateString("pt-BR");
}

function formatarToneladas(valor){
    if(valor === null || valor === undefined || !Number.isFinite(Number(valor))){
        return "-";
    }
    return Number(valor).toLocaleString("pt-BR", {minimumFractionDigits:0, maximumFractionDigits:0});
}

function formatarInteiro(valor){
    if(valor === null || valor === undefined || !Number.isFinite(Number(valor))){
        return "-";
    }
    return Math.round(Number(valor)).toLocaleString("pt-BR");
}

function formatarDecimal(valor, casas=2){
    if(valor === null || valor === undefined || !Number.isFinite(Number(valor))){
        return "-";
    }
    return Number(valor).toLocaleString("pt-BR", {minimumFractionDigits:casas, maximumFractionDigits:casas});
}

function formatarPercentualFracao(valor){
    if(valor === null || valor === undefined || !Number.isFinite(Number(valor))){
        return "N/D";
    }
    return (Number(valor) * 100).toLocaleString("pt-BR", {minimumFractionDigits:1, maximumFractionDigits:1}) + "%";
}

function formatarEixoToneladas(valor){
    const numero = Number(valor);
    if(Math.abs(numero) >= 1000){
        return `${(numero/1000).toLocaleString("pt-BR", {maximumFractionDigits:0})} mil`;
    }
    return numero.toLocaleString("pt-BR");
}

function setTexto(id, texto){
    const elemento = document.getElementById(id);
    if(elemento){
        elemento.textContent = texto ?? "-";
    }
}

function escaparHTML(valor){
    return String(valor ?? "-")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
