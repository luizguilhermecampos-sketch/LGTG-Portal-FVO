let dadosOEE = [];

let graficoEquipamentos = null;
let graficoClassificacao = null;
let graficoNatureza = null;
let graficoTurnos = null;
let graficoMotivos = null;

let analiseOEEGerada = false;


const emailsOEE = [

    "luiz.campos@br.cmoc.com",
    "thais.gilvana@br.cmoc.com"

];


const inputArquivoOEE =
    document.getElementById(
        "arquivoOEE"
    );


const btnGerarOEE =
    document.getElementById(
        "btnGerarOEE"
    );


const btnGerarPDFOEE =
    document.getElementById(
        "btnGerarPDFOEE"
    );


const btnEnviarEmailOEE =
    document.getElementById(
        "btnEnviarEmailOEE"
    );


const filtroDiaOEE =
    document.getElementById(
        "filtroDia"
    );


const filtroTurnoOEE =
    document.getElementById(
        "filtroTurno"
    );


const filtroEquipamentoOEE =
    document.getElementById(
        "filtroEquipamento"
    );


Chart.register(
    ChartDataLabels
);


/* Qualidade base dos gráficos na tela e na impressão */
Chart.defaults.font.family =
    '"Segoe UI", Arial, sans-serif';

Chart.defaults.color =
    "#4F4554";

Chart.defaults.devicePixelRatio =
    Math.max(
        2,
        window.devicePixelRatio || 1
    );


inputArquivoOEE.addEventListener(
    "change",
    carregarPlanilhaOEE
);


btnGerarOEE.addEventListener(
    "click",
    gerarAnaliseOEE
);


btnGerarPDFOEE.addEventListener(
    "click",
    gerarPDFOEE
);


btnEnviarEmailOEE.addEventListener(
    "click",
    enviarEmailOEE
);


[
    filtroDiaOEE,
    filtroTurnoOEE,
    filtroEquipamentoOEE
]
.filter(Boolean)
.forEach(
    function(campo){

        campo.addEventListener(
            "change",
            function(){

                if(
                    dadosOEE.length > 0
                ){

                    gerarAnaliseOEE();

                }

            }
        );

    }
);


/* ============================================================
   CARREGAMENTO
============================================================ */

function mostrarCarregamentoOEE(
    mostrar,
    texto = "Preparando planilha"
){

    const overlay =
        document.getElementById(
            "overlayCarregamentoOEE"
        );


    const textoElemento =
        document.getElementById(
            "textoCarregamentoOEE"
        );


    textoElemento.textContent =
        texto;


    if(mostrar){

        overlay.classList.add(
            "visivel"
        );

    }else{

        overlay.classList.remove(
            "visivel"
        );

    }

}


function atualizarStatusBaseOEE(
    texto
){

    document
    .getElementById(
        "statusBaseOEE"
    )
    .textContent =
        texto;

}


/* ============================================================
   CARREGAR PLANILHA
============================================================ */

function carregarPlanilhaOEE(
    event
){

    const arquivo =
        event.target.files[0];


    if(!arquivo){

        return;

    }


    mostrarCarregamentoOEE(
        true,
        "Lendo arquivo..."
    );


    atualizarStatusBaseOEE(
        "Carregando..."
    );


    const reader =
        new FileReader();


    reader.onload =
        function(e){

            setTimeout(
                function(){

                    try{

                        processarArquivoOEE(
                            e.target.result
                        );

                    }catch(erro){

                        console.error(
                            erro
                        );


                        mostrarCarregamentoOEE(
                            false
                        );


                        atualizarStatusBaseOEE(
                            "Erro ao carregar"
                        );


                        alert(
                            "Não foi possível ler a planilha."
                        );

                    }

                },
                30
            );

        };


    reader.onerror =
        function(){

            mostrarCarregamentoOEE(
                false
            );


            atualizarStatusBaseOEE(
                "Erro ao carregar"
            );


            alert(
                "Não foi possível ler o arquivo."
            );

        };


    reader.readAsArrayBuffer(
        arquivo
    );

}


/* ============================================================
   PROCESSAR PLANILHA
============================================================ */

function processarArquivoOEE(
    arrayBuffer
){

    mostrarCarregamentoOEE(
        true,
        "Abrindo planilha..."
    );


    const dados =
        new Uint8Array(
            arrayBuffer
        );


    const workbook =
        XLSX.read(
            dados,
            {

                type:"array",

                cellDates:true,

                cellStyles:false,

                cellHTML:false,

                cellNF:false,

                cellText:false

            }
        );


    const aba =
        workbook.Sheets[
            "Lista"
        ];


    if(!aba){

        mostrarCarregamentoOEE(
            false
        );


        atualizarStatusBaseOEE(
            "Aba não encontrada"
        );


        alert(
            "A aba 'Lista' não foi encontrada na planilha."
        );

        return;

    }


    mostrarCarregamentoOEE(
        true,
        "Preparando base do OEE..."
    );


    dadosOEE =
        extrairListaOEEOtimizada(
            aba
        );


    if(
        dadosOEE.length === 0
    ){

        mostrarCarregamentoOEE(
            false
        );


        atualizarStatusBaseOEE(
            "Nenhum dado válido"
        );


        alert(
            "Não foram encontrados registros válidos na aba Lista."
        );

        return;

    }


    mostrarCarregamentoOEE(
        true,
        "Preparando filtros..."
    );


    preencherFiltrosOEE();


    atualizarStatusBaseOEE(

        dadosOEE.length
        .toLocaleString(
            "pt-BR"
        ) +

        " registros carregados"

    );


    mostrarCarregamentoOEE(
        true,
        "Gerando análise..."
    );


    gerarAnaliseOEE();


    setTimeout(
        function(){

            mostrarCarregamentoOEE(
                false
            );

        },
        80
    );

}


/* ============================================================
   EXTRAÇÃO OTIMIZADA
============================================================ */

function extrairListaOEEOtimizada(
    planilha
){

    const registros = [];


    if(
        !planilha ||
        !planilha["!ref"]
    ){

        return registros;

    }


    const faixa =
        XLSX.utils.decode_range(
            planilha["!ref"]
        );


    let linhaCabecalho =
        -1;


    let colunas =
        {};


    const limiteCabecalho =
        Math.min(
            faixa.e.r,
            faixa.s.r + 100
        );


    for(
        let linha = faixa.s.r;
        linha <= limiteCabecalho;
        linha++
    ){

        const mapaLinha =
            {};


        for(
            let coluna = faixa.s.c;
            coluna <= faixa.e.c;
            coluna++
        ){

            const valor =
                obterValorPlanilhaOEE(

                    planilha,
                    linha,
                    coluna

                );


            const chave =
                normalizarTexto(
                    valor
                );


            if(chave){

                mapaLinha[
                    chave
                ] =
                    coluna;

            }

        }


        if(
            mapaLinha["data"] !== undefined &&
            mapaLinha["tempo"] !== undefined &&
            mapaLinha["oee"] !== undefined &&
            mapaLinha["equipamento"] !== undefined
        ){

            linhaCabecalho =
                linha;


            colunas =
                mapaLinha;


            break;

        }

    }


    if(
        linhaCabecalho === -1
    ){

        alert(
            "Não foi possível identificar o cabeçalho da aba Lista."
        );


        return [];

    }


    for(
        let linha =
            linhaCabecalho + 1;

        linha <= faixa.e.r;

        linha++
    ){

        const data =
            converterDataExcel(

                obterValorPorNomeOEE(
                    planilha,
                    linha,
                    colunas,
                    "data"
                )

            );


        const tempo =
            converterNumero(

                obterValorPorNomeOEE(
                    planilha,
                    linha,
                    colunas,
                    "tempo"
                )

            );


        if(
            !data ||
            !Number.isFinite(
                tempo
            ) ||
            tempo <= 0
        ){

            continue;

        }


        const registro =
        {

            data:
                data,

            tempo:
                tempo,

            oee:
                limparTexto(
                    obterValorPorNomeOEE(
                        planilha,
                        linha,
                        colunas,
                        "oee"
                    )
                ),

            oeeInfo:
                limparTexto(
                    obterValorPorNomeOEE(
                        planilha,
                        linha,
                        colunas,
                        "oee info"
                    )
                ),

            equipamento:
                limparTexto(
                    obterValorPorNomeOEE(
                        planilha,
                        linha,
                        colunas,
                        "equipamento"
                    )
                ),

            tag:
                limparTexto(
                    obterValorPorNomeOEE(
                        planilha,
                        linha,
                        colunas,
                        "tag"
                    )
                ),

            motivo:
                limparTexto(
                    obterValorPorNomeOEE(
                        planilha,
                        linha,
                        colunas,
                        "motivo"
                    )
                ),

            natureza:
                limparTexto(
                    obterValorPorNomeOEE(
                        planilha,
                        linha,
                        colunas,
                        "natureza da falha"
                    )
                ),

            componente:
                limparTexto(
                    obterValorPorNomeOEE(
                        planilha,
                        linha,
                        colunas,
                        "componente"
                    )
                ),

            turno:
                limparTexto(
                    obterValorPorNomeOEE(
                        planilha,
                        linha,
                        colunas,
                        "turno"
                    )
                ),

            observacao:
                limparTexto(
                    obterValorPorNomeOEE(
                        planilha,
                        linha,
                        colunas,
                        "observacao"
                    )
                )

        };


        registros.push(
            registro
        );

    }


    return registros;

}


/* ============================================================
   LEITURA DIRETA DAS CÉLULAS
============================================================ */

function obterValorPlanilhaOEE(
    planilha,
    linha,
    coluna
){

    const endereco =
        XLSX.utils.encode_cell(
            {

                r:linha,

                c:coluna

            }
        );


    const celula =
        planilha[
            endereco
        ];


    return celula
        ?
        celula.v
        :
        null;

}


function obterValorPorNomeOEE(
    planilha,
    linha,
    colunas,
    nome
){

    const indice =
        colunas[
            normalizarTexto(
                nome
            )
        ];


    if(
        indice === undefined
    ){

        return null;

    }


    return obterValorPlanilhaOEE(

        planilha,

        linha,

        indice

    );

}


/* ============================================================
   TEXTO
============================================================ */

function normalizarTexto(
    valor
){

    if(
        valor === null ||
        valor === undefined
    ){

        return "";

    }


    return String(
        valor
    )
    .trim()
    .toLowerCase()
    .normalize(
        "NFD"
    )
    .replace(
        /[\u0300-\u036f]/g,
        ""
    );

}


function limparTexto(
    valor
){

    if(
        valor === null ||
        valor === undefined
    ){

        return "";

    }


    return String(
        valor
    )
    .trim()
    .replace(
        /\s+/g,
        " "
    );

}


/* ============================================================
   NÚMERO
============================================================ */

function converterNumero(
    valor
){

    if(
        typeof valor ===
        "number"
    ){

        return valor;

    }


    if(
        typeof valor ===
        "string"
    ){

        return Number(

            valor
            .replace(
                ".",
                ""
            )
            .replace(
                ",",
                "."
            )

        );

    }


    return NaN;

}


/* ============================================================
   DATA
============================================================ */

function converterDataExcel(
    valor
){

    if(!valor){

        return null;

    }


    if(
        valor instanceof
        Date
    ){

        if(
            !isNaN(
                valor.getTime()
            )
        ){

            return valor;

        }

    }


    if(
        typeof valor ===
        "number"
    ){

        const dataExcel =
            XLSX.SSF
            .parse_date_code(
                valor
            );


        if(!dataExcel){

            return null;

        }


        return new Date(

            dataExcel.y,

            dataExcel.m - 1,

            dataExcel.d

        );

    }


    if(
        typeof valor ===
        "string"
    ){

        const partes =
            valor
            .trim()
            .split(
                /[\/\-]/
            );


        if(
            partes.length === 3
        ){

            let dia;
            let mes;
            let ano;


            if(
                partes[0].length === 4
            ){

                ano =
                    Number(
                        partes[0]
                    );

                mes =
                    Number(
                        partes[1]
                    );

                dia =
                    Number(
                        partes[2]
                    );

            }else{

                dia =
                    Number(
                        partes[0]
                    );

                mes =
                    Number(
                        partes[1]
                    );

                ano =
                    Number(
                        partes[2]
                    );

            }


            const data =
                new Date(

                    ano,

                    mes - 1,

                    dia

                );


            if(
                !isNaN(
                    data.getTime()
                )
            ){

                return data;

            }

        }

    }


    return null;

}


/* ============================================================
   FILTROS
============================================================ */

function preencherFiltrosOEE(){

    preencherFiltroDia();

    preencherFiltroTurno();

    preencherFiltroEquipamento();

}


function preencherFiltroDia(){

    const input =
        document.getElementById(
            "filtroDia"
        );


    if(!input){

        return;

    }


    const datasValidas =
        dadosOEE
        .map(
            item =>
                item.data
        )
        .filter(
            data =>
                data instanceof Date &&
                !isNaN(
                    data.getTime()
                )
        )
        .sort(
            (a,b) =>
                a - b
        );


    if(
        datasValidas.length === 0
    ){

        input.value = "";

        input.removeAttribute(
            "min"
        );

        input.removeAttribute(
            "max"
        );

        return;

    }


    const primeiraData =
        formatarDataISOOEE(
            datasValidas[0]
        );


    const ultimaData =
        formatarDataISOOEE(
            datasValidas[
                datasValidas.length - 1
            ]
        );


    input.min =
        primeiraData;


    input.max =
        ultimaData;


    /*
       A análise abre no dia mais recente disponível.
       Se o usuário limpar o campo, toda a base fica disponível.
    */
    input.value =
        ultimaData;

}


function formatarDataISOOEE(
    data
){

    if(
        !(data instanceof Date) ||
        isNaN(
            data.getTime()
        )
    ){

        return "";

    }


    const ano =
        data.getFullYear();


    const mes =
        String(
            data.getMonth() + 1
        )
        .padStart(
            2,
            "0"
        );


    const dia =
        String(
            data.getDate()
        )
        .padStart(
            2,
            "0"
        );


    return (
        `${ano}-${mes}-${dia}`
    );

}


function preencherFiltroTurno(){

    const select =
        document.getElementById(
            "filtroTurno"
        );


    select.innerHTML =
        `<option value="todos">Todos</option>`;


    const turnos =
        valoresUnicos(

            dadosOEE.map(
                item =>
                    item.turno
            )

        );


    turnos.sort();


    turnos.forEach(
        function(turno){

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                turno;


            option.textContent =
                `Turno ${turno}`;


            select.appendChild(
                option
            );

        }
    );

}


function preencherFiltroEquipamento(){

    const select =
        document.getElementById(
            "filtroEquipamento"
        );


    select.innerHTML =
        `<option value="todos">Todos</option>`;


    const equipamentos =
        valoresUnicos(

            dadosOEE.map(
                item =>
                    item.equipamento
            )

        );


    equipamentos.sort(
        function(a,b){

            return a.localeCompare(
                b,
                "pt-BR"
            );

        }
    );


    equipamentos.forEach(
        function(equipamento){

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                equipamento;


            option.textContent =
                equipamento;


            select.appendChild(
                option
            );

        }
    );

}


/* ============================================================
   GERAR ANÁLISE
============================================================ */

function gerarAnaliseOEE(){

    if(
        dadosOEE.length === 0
    ){

        alert(
            "Carregue a planilha OEE primeiro."
        );

        return;

    }


    const dadosFiltrados =
        aplicarFiltrosOEE();


    if(
        dadosFiltrados.length === 0
    ){

        alert(
            "Não existem registros para os filtros selecionados."
        );


        limparPainelOEE();


        return;

    }


    atualizarCardsOEE(
        dadosFiltrados
    );


    gerarGraficoEquipamentos(
        dadosFiltrados
    );


    gerarGraficoClassificacao(
        dadosFiltrados
    );


    gerarGraficoNatureza(
        dadosFiltrados
    );


    gerarGraficoTurnos(
        dadosFiltrados
    );


    gerarGraficoMotivos(
        dadosFiltrados
    );


    atualizarResumoOEE(
        dadosFiltrados
    );


    atualizarMetadadosRelatorioOEE(
        dadosFiltrados
    );


    analiseOEEGerada =
        true;

}


/* ============================================================
   FILTRAR
============================================================ */

function aplicarFiltrosOEE(){

    const dia =
        document
        .getElementById(
            "filtroDia"
        )
        ?.value
        .trim() ||
        "";


    const turno =
        document
        .getElementById(
            "filtroTurno"
        )
        .value;


    const equipamento =
        document
        .getElementById(
            "filtroEquipamento"
        )
        .value;


    return dadosOEE.filter(
        function(item){

            if(
                dia
            ){

                const chaveDia =
                    formatarDataISOOEE(
                        item.data
                    );


                if(
                    chaveDia !== dia
                ){

                    return false;

                }

            }


            if(
                turno !== "todos" &&
                item.turno !== turno
            ){

                return false;

            }


            if(
                equipamento !== "todos" &&
                item.equipamento !== equipamento
            ){

                return false;

            }


            return true;

        }
    );

}


/* ============================================================
   CARDS
============================================================ */

function atualizarCardsOEE(
    dados
){

    const total =
        somarTempo(
            dados
        );


    const pmc =
        somarTempo(

            dados.filter(
                item =>
                    normalizarTexto(
                        item.oee
                    ) === "pmc"
            )

        );


    const pfe =
        somarTempo(

            dados.filter(
                item =>
                    normalizarTexto(
                        item.oee
                    ) === "pfe"
            )

        );


    const pfp =
        somarTempo(

            dados.filter(
                item =>
                    normalizarTexto(
                        item.oee
                    ) === "pfp"
            )

        );


    const reducao =
        somarTempo(

            dados.filter(
                item =>
                    normalizarTexto(
                        item.oee
                    )
                    .includes(
                        "reducao"
                    )
            )

        );


    document
    .getElementById(
        "cardHorasTotais"
    )
    .textContent =
        formatarHoras(
            total
        );


    document
    .getElementById(
        "cardPMC"
    )
    .textContent =
        formatarHoras(
            pmc
        );


    document
    .getElementById(
        "cardPFE"
    )
    .textContent =
        formatarHoras(
            pfe
        );


    document
    .getElementById(
        "cardPFP"
    )
    .textContent =
        formatarHoras(
            pfp
        );


    document
    .getElementById(
        "cardReducao"
    )
    .textContent =
        formatarHoras(
            reducao
        );


    document
    .getElementById(
        "percentualPMC"
    )
    .textContent =
        formatarPercentual(
            pmc,
            total
        );


    document
    .getElementById(
        "percentualPFE"
    )
    .textContent =
        formatarPercentual(
            pfe,
            total
        );


    document
    .getElementById(
        "percentualPFP"
    )
    .textContent =
        formatarPercentual(
            pfp,
            total
        );


    document
    .getElementById(
        "percentualReducao"
    )
    .textContent =
        formatarPercentual(
            reducao,
            total
        );


    const classificacao =
        agruparTempo(

            dados,

            item =>
                item.oeeInfo ||
                item.oee ||
                "Não informado"

        );


    const principal =
        maiorGrupo(
            classificacao
        );


    document
    .getElementById(
        "cardPrincipalImpacto"
    )
    .textContent =
        principal
        ?
        principal.nome
        :
        "-";

}


/* ============================================================
   GRÁFICO EQUIPAMENTOS
============================================================ */

function gerarGraficoEquipamentos(
    dados
){

    const grupos =
        agruparTempo(

            dados,

            item =>
                item.equipamento ||
                "Não informado"

        );


    const ranking =
        ordenarGrupos(
            grupos
        );


    destruirGrafico(
        graficoEquipamentos
    );


    graficoEquipamentos =
        new Chart(

            document.getElementById(
                "graficoEquipamentosOEE"
            ),

            {

                type:"bar",

                data:{

                    labels:

                        ranking.map(
                            item =>
                                item.nome
                        ),

                    datasets:[
                        {

                            label:
                                "Horas de perda",

                            data:

                                ranking.map(
                                    item =>
                                        item.valor
                                ),

                            backgroundColor:
                                "#5A1180",

                            borderRadius:
                                5

                        }
                    ]

                },


                options:{

                    indexAxis:
                        "y",

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,


                    layout:{

                        padding:{

                            right:
                                35

                        }

                    },


                    plugins:{

                        legend:{

                            display:
                                false

                        },


                        datalabels:{

                            anchor:
                                "end",

                            align:
                                "right",

                            clamp:
                                true,

                            clip:
                                false,

                            color:
                                "#2D014D",

                            font:{

                                weight:
                                    "bold",

                                size:
                                    11

                            },

                            formatter:
                                function(valor){

                                    return (
                                        valor
                                        .toLocaleString(
                                            "pt-BR",
                                            {
                                                minimumFractionDigits:1,
                                                maximumFractionDigits:1
                                            }
                                        ) +
                                        " h"
                                    );

                                }

                        }

                    },


                    scales:{

                        x:{

                            beginAtZero:
                                true,

                            grace:
                                "10%",

                            title:{

                                display:
                                    true,

                                text:
                                    "Horas"

                            }

                        },


                        y:{

                            grid:{

                                display:
                                    false

                            }

                        }

                    }

                },


                plugins:[
                    ChartDataLabels
                ]

            }

        );

}


/* ============================================================
   GRÁFICO CLASSIFICAÇÃO
============================================================ */

function gerarGraficoClassificacao(
    dados
){

    const grupos =
        agruparTempo(

            dados,

            item =>
                item.oee
                ?
                item.oee
                :
                "Não informado"

        );


    const ranking =
        ordenarGrupos(
            grupos
        );


    const total =
        somarTempo(
            dados
        );


    destruirGrafico(
        graficoClassificacao
    );


    graficoClassificacao =
        new Chart(

            document.getElementById(
                "graficoClassificacaoOEE"
            ),

            {

                type:
                    "doughnut",


                data:{

                    labels:

                        ranking.map(
                            item =>
                                item.nome
                        ),

                    datasets:[
                        {

                            data:

                                ranking.map(
                                    item =>
                                        item.valor
                                ),

                            backgroundColor:[

                                "#2D014D",
                                "#D83A52",
                                "#EE9622",
                                "#92D050",
                                "#2878B8",
                                "#00A0A8",
                                "#7A3E9D",
                                "#9B929F"

                            ],

                            borderWidth:
                                2,

                            borderColor:
                                "#FFFFFF"

                        }
                    ]

                },


                options:{

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    cutout:
                        "55%",


                    plugins:{

                        legend:{

                            position:
                                "right",

                            labels:{

                                boxWidth:
                                    12,

                                font:{

                                    size:
                                        11

                                }

                            }

                        },


                        datalabels:{

                            color:
                                "#FFFFFF",

                            font:{

                                weight:
                                    "bold",

                                size:
                                    11

                            },

                            formatter:
                                function(valor){

                                    const percentual =
                                        total > 0
                                        ?
                                        (
                                            valor /
                                            total
                                        ) * 100
                                        :
                                        0;


                                    if(
                                        percentual < 4
                                    ){

                                        return "";

                                    }


                                    return (

                                        percentual
                                        .toFixed(
                                            1
                                        ) +

                                        "%"

                                    );

                                }

                        }

                    }

                },


                plugins:[
                    ChartDataLabels
                ]

            }

        );

}


/* ============================================================
   GRÁFICO NATUREZA
============================================================ */

function gerarGraficoNatureza(
    dados
){

    const grupos =
        agruparTempo(

            dados,

            item =>
                item.natureza ||
                "Não informado"

        );


    const ranking =
        ordenarGrupos(
            grupos
        );


    const total =
        somarTempo(
            dados
        );


    destruirGrafico(
        graficoNatureza
    );


    graficoNatureza =
        new Chart(

            document.getElementById(
                "graficoNaturezaOEE"
            ),

            {

                type:
                    "doughnut",


                data:{

                    labels:

                        ranking.map(
                            item =>
                                item.nome
                        ),

                    datasets:[
                        {

                            data:

                                ranking.map(
                                    item =>
                                        item.valor
                                ),

                            backgroundColor:[

                                "#2878B8",
                                "#92D050",
                                "#EE9622",
                                "#D83A52",
                                "#7A3E9D",
                                "#00A0A8",
                                "#9B929F"

                            ],

                            borderWidth:
                                2,

                            borderColor:
                                "#FFFFFF"

                        }
                    ]

                },


                options:{

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    cutout:
                        "55%",


                    plugins:{

                        legend:{

                            position:
                                "right",

                            labels:{

                                boxWidth:
                                    12,

                                font:{

                                    size:
                                        11

                                }

                            }

                        },


                        datalabels:{

                            color:
                                "#FFFFFF",

                            font:{

                                weight:
                                    "bold",

                                size:
                                    11

                            },

                            formatter:
                                function(valor){

                                    const percentual =
                                        total > 0
                                        ?
                                        (
                                            valor /
                                            total
                                        ) * 100
                                        :
                                        0;


                                    if(
                                        percentual < 4
                                    ){

                                        return "";

                                    }


                                    return (
                                        percentual
                                        .toFixed(
                                            1
                                        ) +
                                        "%"
                                    );

                                }

                        }

                    }

                },


                plugins:[
                    ChartDataLabels
                ]

            }

        );

}


/* ============================================================
   TURNOS
============================================================ */

function gerarGraficoTurnos(
    dados
){

    const grupos =
        agruparTempo(

            dados,

            item =>
                item.turno ||
                "Não informado"

        );


    const ranking =
        Array.from(

            grupos,

            ([nome,valor]) =>
                ({
                    nome,
                    valor
                })

        )
        .sort(
            (a,b) =>
                a.nome.localeCompare(
                    b.nome,
                    "pt-BR"
                )
        );


    destruirGrafico(
        graficoTurnos
    );


    graficoTurnos =
        new Chart(

            document.getElementById(
                "graficoTurnosOEE"
            ),

            {

                type:
                    "bar",


                data:{

                    labels:

                        ranking.map(
                            item =>
                                item.nome ===
                                "Não informado"
                                ?
                                item.nome
                                :
                                "Turno " +
                                item.nome
                        ),

                    datasets:[
                        {

                            label:
                                "Horas",

                            data:

                                ranking.map(
                                    item =>
                                        item.valor
                                ),

                            backgroundColor:
                                "#92D050",

                            borderRadius:
                                6

                        }
                    ]

                },


                options:{

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,


                    layout:{

                        padding:{

                            top:
                                20

                        }

                    },


                    plugins:{

                        legend:{

                            display:
                                false

                        },


                        datalabels:{

                            anchor:
                                "end",

                            align:
                                "top",

                            clamp:
                                true,

                            clip:
                                false,

                            color:
                                "#2D014D",

                            font:{

                                weight:
                                    "bold"

                            },

                            formatter:
                                function(valor){

                                    return (

                                        valor
                                        .toFixed(
                                            1
                                        ) +

                                        " h"

                                    );

                                }

                        }

                    },


                    scales:{

                        y:{

                            beginAtZero:
                                true,

                            grace:
                                "15%"

                        },


                        x:{

                            grid:{

                                display:
                                    false

                            }

                        }

                    }

                },


                plugins:[
                    ChartDataLabels
                ]

            }

        );

}


/* ============================================================
   TOP 10 MOTIVOS
============================================================ */

function gerarGraficoMotivos(
    dados
){

    const grupos =
        agruparTempo(

            dados,

            item =>
                item.motivo ||
                "Não informado"

        );


    const ranking =
        ordenarGrupos(
            grupos
        )
        .slice(
            0,
            10
        );


    destruirGrafico(
        graficoMotivos
    );


    graficoMotivos =
        new Chart(

            document.getElementById(
                "graficoMotivosOEE"
            ),

            {

                type:
                    "bar",


                data:{

                    labels:

                        ranking.map(
                            item =>
                                item.nome
                        ),

                    datasets:[
                        {

                            label:
                                "Horas",

                            data:

                                ranking.map(
                                    item =>
                                        item.valor
                                ),

                            backgroundColor:
                                "#2878B8",

                            borderRadius:
                                5

                        }
                    ]

                },


                options:{

                    indexAxis:
                        "y",

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,


                    layout:{

                        padding:{

                            right:
                                35

                        }

                    },


                    plugins:{

                        legend:{

                            display:
                                false

                        },


                        datalabels:{

                            anchor:
                                "end",

                            align:
                                "right",

                            clamp:
                                true,

                            clip:
                                false,

                            color:
                                "#2D014D",

                            font:{

                                size:
                                    10,

                                weight:
                                    "bold"

                            },

                            formatter:
                                function(valor){

                                    return (

                                        valor
                                        .toFixed(
                                            1
                                        ) +

                                        " h"

                                    );

                                }

                        }

                    },


                    scales:{

                        x:{

                            beginAtZero:
                                true,

                            grace:
                                "12%"

                        },


                        y:{

                            grid:{

                                display:
                                    false

                            }

                        }

                    }

                },


                plugins:[
                    ChartDataLabels
                ]

            }

        );

}


/* ============================================================
   RESUMO
============================================================ */

function atualizarResumoOEE(
    dados
){

    const classificacoes =
        agruparTempo(

            dados,

            item =>
                item.oeeInfo ||
                item.oee ||
                "Não informado"

        );


    const impacto =
        maiorGrupo(
            classificacoes
        );


    document
    .getElementById(
        "resumoImpacto"
    )
    .textContent =
        impacto
        ?
        (
            impacto.nome +

            " (" +

            formatarHoras(
                impacto.valor
            ) +

            " h)"
        )
        :
        "-";


    const equipamentos =
        agruparTempo(

            dados,

            item =>
                item.equipamento ||
                "Não informado"

        );


    const equipamento =
        maiorGrupo(
            equipamentos
        );


    document
    .getElementById(
        "resumoEquipamento"
    )
    .textContent =
        equipamento
        ?
        (
            equipamento.nome +

            " (" +

            formatarHoras(
                equipamento.valor
            ) +

            " h)"
        )
        :
        "-";


    const naturezas =
        agruparTempo(

            dados,

            item =>
                item.natureza ||
                "Não informado"

        );


    const natureza =
        maiorGrupo(
            naturezas
        );


    document
    .getElementById(
        "resumoNatureza"
    )
    .textContent =
        natureza
        ?
        (
            natureza.nome +

            " (" +

            formatarHoras(
                natureza.valor
            ) +

            " h)"
        )
        :
        "-";


    const turnos =
        agruparTempo(

            dados,

            item =>
                item.turno ||
                "Não informado"

        );


    const turno =
        maiorGrupo(
            turnos
        );


    document
    .getElementById(
        "resumoTurno"
    )
    .textContent =
        turno
        ?
        (
            (
                turno.nome ===
                "Não informado"
                ?
                turno.nome
                :
                "Turno " +
                turno.nome
            ) +

            " (" +

            formatarHoras(
                turno.valor
            ) +

            " h)"
        )
        :
        "-";


    const datas =
        dados
        .map(
            item =>
                item.data
        )
        .sort(
            (a,b) =>
                a - b
        );


    if(
        datas.length > 0
    ){

        const inicio =
            formatarData(
                datas[0]
            );


        const fim =
            formatarData(

                datas[
                    datas.length - 1
                ]

            );


        document
        .getElementById(
            "resumoPeriodo"
        )
        .textContent =
            inicio === fim
            ?
            inicio
            :
            inicio +
            " a " +
            fim;

    }

}


/* ============================================================
   METADADOS DO RELATÓRIO / PDF
============================================================ */

function atualizarMetadadosRelatorioOEE(
    dados
){

    const periodoElemento =
        document.getElementById(
            "periodoRelatorioPDFOEE"
        );


    const filtrosElemento =
        document.getElementById(
            "filtrosRelatorioPDFOEE"
        );


    const geradoElemento =
        document.getElementById(
            "geradoEmRelatorioPDFOEE"
        );


    const periodoPagina2Elemento =
        document.getElementById(
            "periodoPagina2PDFOEE"
        );


    const periodoPagina3Elemento =
        document.getElementById(
            "periodoPagina3PDFOEE"
        );


    if(
        periodoElemento
    ){

        periodoElemento.textContent =
            "Dia analisado: " +
            obterTextoOEE(
                "resumoPeriodo"
            );

    }


    if(
        periodoPagina2Elemento
    ){

        periodoPagina2Elemento.textContent =
            "Dia analisado: " +
            obterTextoOEE(
                "resumoPeriodo"
            );

    }


    if(
        periodoPagina3Elemento
    ){

        periodoPagina3Elemento.textContent =
            "Dia analisado: " +
            obterTextoOEE(
                "resumoPeriodo"
            );

    }


    if(
        filtrosElemento
    ){

        const turno =
            document
            .getElementById(
                "filtroTurno"
            )
            ?.value ||
            "todos";


        const equipamento =
            document
            .getElementById(
                "filtroEquipamento"
            )
            ?.value ||
            "todos";


        const partes = [];


        partes.push(
            turno === "todos"
            ?
            "Todos os turnos"
            :
            "Turno " + turno
        );


        partes.push(
            equipamento === "todos"
            ?
            "Todos os equipamentos"
            :
            equipamento
        );


        filtrosElemento.textContent =
            "Filtros: " +
            partes.join(
                " • "
            );

    }


    if(
        geradoElemento
    ){

        const agora =
            new Date();


        geradoElemento.textContent =
            "Gerado em: " +
            agora.toLocaleString(
                "pt-BR"
            );

    }

}


/* ============================================================
   AGRUPAMENTO
============================================================ */

function agruparTempo(
    dados,
    seletor
){

    const mapa =
        new Map();


    dados.forEach(
        function(item){

            let nome =
                seletor(
                    item
                );


            if(!nome){

                nome =
                    "Não informado";

            }


            nome =
                limparTexto(
                    nome
                );


            if(!nome){

                nome =
                    "Não informado";

            }


            mapa.set(

                nome,

                (
                    mapa.get(
                        nome
                    ) ||
                    0
                ) +

                item.tempo

            );

        }
    );


    return mapa;

}


function ordenarGrupos(
    mapa
){

    return Array.from(

        mapa,

        ([nome,valor]) =>
            ({
                nome,
                valor
            })

    )
    .sort(
        (a,b) =>
            b.valor -
            a.valor
    );

}


function maiorGrupo(
    mapa
){

    const grupos =
        ordenarGrupos(
            mapa
        );


    return grupos.length > 0
        ?
        grupos[0]
        :
        null;

}


function somarTempo(
    dados
){

    return dados.reduce(
        function(total,item){

            return (

                total +

                (
                    Number(
                        item.tempo
                    ) ||
                    0
                )

            );

        },
        0
    );

}


function valoresUnicos(
    valores
){

    return Array.from(

        new Set(

            valores
            .map(
                valor =>
                    limparTexto(
                        valor
                    )
            )
            .filter(
                valor =>
                    valor !== ""
            )

        )

    );

}


/* ============================================================
   FORMATAÇÃO
============================================================ */

function formatarHoras(
    valor
){

    return Number(
        valor || 0
    )
    .toLocaleString(
        "pt-BR",
        {

            minimumFractionDigits:
                1,

            maximumFractionDigits:
                1

        }
    );

}


function formatarPercentual(
    valor,
    total
){

    if(!total){

        return "0,0% do total";

    }


    const percentual =
        (
            valor /
            total
        ) *
        100;


    return (

        percentual
        .toLocaleString(
            "pt-BR",
            {

                minimumFractionDigits:
                    1,

                maximumFractionDigits:
                    1

            }
        ) +

        "% do total"

    );

}


function formatarData(
    data
){

    return new Intl
    .DateTimeFormat(
        "pt-BR"
    )
    .format(
        data
    );

}


function primeiraMaiuscula(
    texto
){

    if(!texto){

        return "";

    }


    return (

        texto.charAt(
            0
        )
        .toUpperCase() +

        texto.slice(
            1
        )

    );

}


function destruirGrafico(
    grafico
){

    if(grafico){

        grafico.destroy();

    }

}


/* ============================================================
   QUALIDADE DOS GRÁFICOS NA IMPRESSÃO / PDF
============================================================ */

function ajustarGraficoParaPDFOEE(
    grafico,
    tamanhoRotulo = 13,
    tamanhoEixo = 12,
    tamanhoLegenda = 12
){

    if(
        !grafico
    ){

        return;

    }


    grafico.options.devicePixelRatio =
        3;


    grafico.options.animation =
        false;


    if(
        grafico.options.plugins
        ?.datalabels
    ){

        grafico.options.plugins.datalabels.font = {

            ...(
                grafico.options.plugins.datalabels.font || {}
            ),

            size:
                tamanhoRotulo,

            weight:
                "700"

        };

    }


    if(
        grafico.options.plugins
        ?.legend
        ?.labels
    ){

        grafico.options.plugins.legend.labels.font = {

            ...(
                grafico.options.plugins.legend.labels.font || {}
            ),

            size:
                tamanhoLegenda,

            weight:
                "600"

        };


        grafico.options.plugins.legend.labels.boxWidth =
            14;


        grafico.options.plugins.legend.labels.padding =
            12;

    }


    Object.values(
        grafico.options.scales || {}
    )
    .forEach(
        function(escala){

            escala.ticks = {

                ...(escala.ticks || {}),

                color:
                    "#4F4554",

                font:{

                    ...(escala.ticks?.font || {}),

                    size:
                        tamanhoEixo,

                    weight:
                        "600"

                }

            };


            if(
                escala.title
                ?.display
            ){

                escala.title.font = {

                    ...(escala.title.font || {}),

                    size:
                        tamanhoEixo,

                    weight:
                        "700"

                };

            }

        }
    );


    grafico.resize();


    grafico.update(
        "none"
    );

}


function prepararGraficosParaPDFOEE(){

    ajustarGraficoParaPDFOEE(
        graficoEquipamentos,
        13,
        12,
        12
    );


    ajustarGraficoParaPDFOEE(
        graficoClassificacao,
        13,
        12,
        12
    );


    ajustarGraficoParaPDFOEE(
        graficoNatureza,
        13,
        12,
        12
    );


    ajustarGraficoParaPDFOEE(
        graficoTurnos,
        14,
        12,
        12
    );


    ajustarGraficoParaPDFOEE(
        graficoMotivos,
        13,
        12,
        12
    );

}


function restaurarGraficosAposPDFOEE(){

    if(
        !analiseOEEGerada ||
        dadosOEE.length === 0
    ){

        return;

    }


    setTimeout(
        function(){

            gerarAnaliseOEE();

        },
        120
    );

}


window.addEventListener(
    "beforeprint",
    prepararGraficosParaPDFOEE
);


window.addEventListener(
    "afterprint",
    restaurarGraficosAposPDFOEE
);


/* ============================================================
   PDF
============================================================ */

function gerarPDFOEE(){

    if(
        !analiseOEEGerada
    ){

        alert(
            "Gere uma análise OEE antes de criar o PDF."
        );

        return;

    }


    atualizarMetadadosRelatorioOEE(
        aplicarFiltrosOEE()
    );


    const tituloAnterior =
        document.title;


    const periodo =
        obterTextoOEE(
            "resumoPeriodo"
        );


    const nomeSeguro =
        periodo
        .replace(
            /[^a-zA-Z0-9]+/g,
            "_"
        );


    document.title =

        "OEE_Britagem_FVO_" +

        nomeSeguro;


    document.body
    .classList
    .add(
        "modo-impressao-oee"
    );


    setTimeout(
        function(){

            window.print();


            setTimeout(
                function(){

                    document.body
                    .classList
                    .remove(
                        "modo-impressao-oee"
                    );


                    document.title =
                        tituloAnterior;

                },
                500
            );

        },
        200
    );

}


/* ============================================================
   EMAIL
============================================================ */

function enviarEmailOEE(){

    if(
        !analiseOEEGerada
    ){

        alert(
            "Gere uma análise OEE antes de enviar."
        );

        return;

    }


    const periodo =
        obterTextoOEE(
            "resumoPeriodo"
        );


    const horasTotais =
        obterTextoOEE(
            "cardHorasTotais"
        );


    const principal =
        obterTextoOEE(
            "resumoImpacto"
        );


    const equipamento =
        obterTextoOEE(
            "resumoEquipamento"
        );


    const natureza =
        obterTextoOEE(
            "resumoNatureza"
        );


    const turno =
        obterTextoOEE(
            "resumoTurno"
        );


    const assunto =

        "OEE - Britagem FVO - " +

        periodo;


    const mensagem =

`Prezados,

Segue o acompanhamento do OEE da Britagem FVO.

Período analisado: ${periodo}

Horas totais de perda: ${horasTotais} h

Principal impacto: ${principal}

Equipamento mais impactado: ${equipamento}

Natureza predominante: ${natureza}

Turno com maior perda: ${turno}

O relatório completo pode ser gerado em PDF pelo Portal FVO.

Atenciosamente.`;


    const destinatarios =
        emailsOEE.join(
            ";"
        );


    const link =

        "mailto:" +

        destinatarios +

        "?subject=" +

        encodeURIComponent(
            assunto
        ) +

        "&body=" +

        encodeURIComponent(
            mensagem
        );


    window.location.href =
        link;

}


function obterTextoOEE(
    id
){

    const elemento =
        document.getElementById(
            id
        );


    return elemento
        ?
        elemento.textContent.trim()
        :
        "-";

}


/* ============================================================
   LIMPAR
============================================================ */

function limparPainelOEE(){

    analiseOEEGerada =
        false;


    const ids = [

        "cardHorasTotais",

        "cardPMC",

        "cardPFE",

        "cardPFP",

        "cardReducao",

        "cardPrincipalImpacto",

        "percentualPMC",

        "percentualPFE",

        "percentualPFP",

        "percentualReducao",

        "resumoImpacto",

        "resumoEquipamento",

        "resumoNatureza",

        "resumoTurno",

        "resumoPeriodo"

    ];


    ids.forEach(
        function(id){

            const elemento =
                document
                .getElementById(
                    id
                );


            if(elemento){

                elemento.textContent =
                    "-";

            }

        }
    );


    if(
        graficoEquipamentos
    ){

        graficoEquipamentos.destroy();

        graficoEquipamentos =
            null;

    }


    if(
        graficoClassificacao
    ){

        graficoClassificacao.destroy();

        graficoClassificacao =
            null;

    }


    if(
        graficoNatureza
    ){

        graficoNatureza.destroy();

        graficoNatureza =
            null;

    }


    if(
        graficoTurnos
    ){

        graficoTurnos.destroy();

        graficoTurnos =
            null;

    }


    if(
        graficoMotivos
    ){

        graficoMotivos.destroy();

        graficoMotivos =
            null;

    }

}