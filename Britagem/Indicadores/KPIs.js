/* ============================================================
   VARIÁVEIS GERAIS
============================================================ */

let abaFechamento = [];

/*
    Agora NÃO armazenamos a aba Indicadores Britagem inteira.

    Guardamos somente:

    Data
    BAP
    S&OP
    Real

    Isso reduz bastante o processamento depois do carregamento.
*/

let baseProducaoDiaria = [];

let graficoProducao = null;
let graficoMTDYTD = null;

let tipoVisualizacao = "mensal";
let tipoProducaoAcumulada = "mtd";


/* ============================================================
   MESES
============================================================ */

let meses = {

    "Janeiro":3,
    "Fevereiro":4,
    "Março":5,
    "Abril":6,
    "Maio":7,
    "Junho":8,
    "Julho":9,
    "Agosto":10,
    "Setembro":11,
    "Outubro":12,
    "Novembro":13,
    "Dezembro":14

};


/* ============================================================
   MAPEAMENTO DA PRODUÇÃO EM BASE SECA

   Os valores abaixo são apenas alternativas de segurança.
   Após carregar a planilha, o sistema procura automaticamente
   os cabeçalhos de Realizado, BAP e S&OP em base seca.
============================================================ */

const linhasProducaoFechamento = {
    bap:7,
    sop:8,
    real:9
};


const nomesMeses = [

    "Jan",
    "Fev",
    "Mar",
    "Abr",
    "Mai",
    "Jun",
    "Jul",
    "Ago",
    "Set",
    "Out",
    "Nov",
    "Dez"

];


/* ============================================================
   COLUNAS DE SEGURANÇA - INDICADORES BRITAGEM

   O sistema identifica as colunas pelos cabeçalhos e aceita
   somente os campos que contenham "seca". As posições abaixo
   servem para validar a estrutura atual da planilha:

   A = DATA
   B = BAP SECO
   C = S&OP SECO
   E = REALIZADO SECO

   Índices base 0:
============================================================ */

const COLUNAS_SECAS_ESPERADAS = {
    data:0,
    bap:1,
    sop:2,
    real:4
};


/* ============================================================
   EMAILS
============================================================ */

const emailsKPI = [

    "luiz.campos@br.cmoc.com",
    "thais.gilvana@br.cmoc.com"

];


/* ============================================================
   EVENTOS
============================================================ */

document
.getElementById(
    "arquivoExcel"
)
.addEventListener(
    "change",
    carregarPlanilha
);


document
.getElementById(
    "mesSelecionado"
)
.addEventListener(
    "change",
    function(){

        if(
            abaFechamento.length > 0
        ){

            gerarIndicadores();

        }

    }
);


document
.getElementById(
    "dataReferencia"
)
.addEventListener(
    "change",
    function(){

        if(
            baseProducaoDiaria.length > 0
        ){

            gerarMTDYTD();

        }

    }
);


/* ============================================================
   MENSAL
============================================================ */

document
.getElementById(
    "btnMensal"
)
.addEventListener(
    "click",
    function(){

        tipoVisualizacao =
            "mensal";


        ativarBotao(
            "btnMensal",
            "btnAcumulado"
        );


        if(
            abaFechamento.length > 0
        ){

            gerarIndicadores();

        }

    }
);


/* ============================================================
   ACUMULADO
============================================================ */

document
.getElementById(
    "btnAcumulado"
)
.addEventListener(
    "click",
    function(){

        tipoVisualizacao =
            "acumulado";


        ativarBotao(
            "btnAcumulado",
            "btnMensal"
        );


        if(
            abaFechamento.length > 0
        ){

            gerarIndicadores();

        }

    }
);


/* ============================================================
   MTD
============================================================ */

document
.getElementById(
    "btnMTD"
)
.addEventListener(
    "click",
    function(){

        tipoProducaoAcumulada =
            "mtd";


        ativarBotao(
            "btnMTD",
            "btnYTD"
        );


        if(
            baseProducaoDiaria.length > 0
        ){

            gerarMTDYTD();

        }

    }
);


/* ============================================================
   YTD
============================================================ */

document
.getElementById(
    "btnYTD"
)
.addEventListener(
    "click",
    function(){

        tipoProducaoAcumulada =
            "ytd";


        ativarBotao(
            "btnYTD",
            "btnMTD"
        );


        if(
            baseProducaoDiaria.length > 0
        ){

            gerarMTDYTD();

        }

    }
);


/* ============================================================
   PDF
============================================================ */

document
.getElementById(
    "btnGerarPDF"
)
.addEventListener(
    "click",
    gerarPDFKPI
);


/* ============================================================
   EMAIL
============================================================ */

document
.getElementById(
    "btnEnviarEmail"
)
.addEventListener(
    "click",
    enviarEmailKPI
);


/* ============================================================
   BOTÃO ATIVO
============================================================ */

function ativarBotao(
    ativo,
    inativo
){

    document
    .getElementById(
        ativo
    )
    .classList
    .add(
        "ativo"
    );


    document
    .getElementById(
        inativo
    )
    .classList
    .remove(
        "ativo"
    );

}


/* ============================================================
   CARREGAMENTO
============================================================ */

function mostrarCarregamento(
    mostrar,
    texto = "Preparando planilha"
){

    const overlay =
        document
        .getElementById(
            "overlayCarregamento"
        );


    const textoElemento =
        document
        .getElementById(
            "textoCarregamento"
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


/* ============================================================
   CARREGAR PLANILHA
============================================================ */

function carregarPlanilha(
    event
){

    const arquivo =
        event.target.files[0];


    if(!arquivo){

        return;

    }


    mostrarCarregamento(
        true,
        "Lendo arquivo..."
    );


    const reader =
        new FileReader();


    reader.onload =
        function(e){

            /*
                Pequeno intervalo para permitir que
                o navegador mostre o overlay antes
                do processamento pesado.
            */

            setTimeout(
                function(){

                    try{

                        processarArquivoExcel(
                            e.target.result
                        );

                    }catch(erro){

                        console.error(
                            erro
                        );


                        mostrarCarregamento(
                            false
                        );


                        alert(
                            "Erro ao processar a planilha."
                        );

                    }

                },
                30
            );

        };


    reader.onerror =
        function(){

            mostrarCarregamento(
                false
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
   PROCESSAR EXCEL
============================================================ */

function processarArquivoExcel(
    arrayBuffer
){

    mostrarCarregamento(
        true,
        "Abrindo planilha..."
    );


    const dados =
        new Uint8Array(
            arrayBuffer
        );


    /*
        Mantemos somente o necessário para leitura.

        O XLSX ainda precisa abrir o arquivo,
        mas reduzimos bastante o processamento
        posterior.
    */

    const workbook =
        XLSX.read(
            dados,
            {

                type:"array",

                cellStyles:false,

                cellHTML:false,

                cellNF:false,

                cellText:false

            }
        );


    /* ========================================================
       FECHAMENTO BRITAGEM
    ======================================================== */

    mostrarCarregamento(
        true,
        "Carregando indicadores mensais..."
    );


    const fechamento =
        workbook.Sheets[
            "Fechamento Britagem"
        ];


    if(!fechamento){

        mostrarCarregamento(
            false
        );


        alert(
            "Aba 'Fechamento Britagem' não encontrada."
        );

        return;

    }


    /*
        Somente A1:O30, incluindo dezembro.

        Antes o SheetJS convertia a aba inteira.
    */

    abaFechamento =
        XLSX.utils.sheet_to_json(
            fechamento,
            {

                header:1,

                defval:null,

                raw:true,

                range:"A1:O30"

            }
        );


    /* ========================================================
       INDICADORES BRITAGEM
    ======================================================== */

    mostrarCarregamento(
        true,
        "Preparando produção diária..."
    );


    const indicadores =
        workbook.Sheets[
            "Indicadores Britagem"
        ];


    if(!indicadores){

        mostrarCarregamento(
            false
        );


        alert(
            "Aba 'Indicadores Britagem' não encontrada."
        );

        return;

    }


    /*
        Aqui está a principal otimização.

        NÃO fazemos:

        sheet_to_json(indicadores)

        porque isso transformaria centenas de
        colunas que nem utilizamos.

        Pegamos somente as colunas cujos cabeçalhos indicam:
        Data
        BAP em tonelada seca
        S&OP em tonelada seca
        Realizado em tonelada seca
    */

    let colunasBaseSeca;


    try{

        colunasBaseSeca =
            identificarColunasBaseSeca(
                indicadores
            );

    }catch(erro){

        mostrarCarregamento(
            false
        );


        console.error(
            erro
        );


        alert(
            erro.message
        );

        return;

    }

    baseProducaoDiaria =
        extrairBaseProducaoDiaria(
            indicadores,
            colunasBaseSeca
        );


    if(
        baseProducaoDiaria.length === 0
    ){

        console.warn(
            "Nenhum registro diário válido encontrado."
        );

    }


    /* ========================================================
       FINALIZAÇÃO
    ======================================================== */

    mostrarCarregamento(
        true,
        "Gerando gráficos..."
    );


    definirUltimaDataReal();


    atualizarData();


    gerarIndicadores();


    gerarMTDYTD();


    /*
        Libera referência grande do workbook
        assim que terminamos de extrair dados.

        Depois disso filtros trabalham somente
        com arrays pequenos.
    */

    setTimeout(
        function(){

            mostrarCarregamento(
                false
            );


            alert(
                "Planilha carregada com sucesso!"
            );

        },
        100
    );

}


function obterMesesAteColuna(
    colunaFinal
){

    const ordem = [
        "Janeiro","Fevereiro","Março","Abril",
        "Maio","Junho","Julho","Agosto",
        "Setembro","Outubro","Novembro","Dezembro"
    ];


    const indiceFinal =
        ordem.findIndex(
            function(nome){
                return meses[nome] === colunaFinal;
            }
        );


    const limite =
        indiceFinal >= 0
        ? indiceFinal
        : 11;


    return ordem
        .slice(0,limite + 1)
        .filter(function(nome){
            return Number.isInteger(meses[nome]);
        })
        .map(function(nome){
            return {
                nome:nome,
                abreviado:nome.slice(0,3),
                coluna:meses[nome]
            };
        });

}


/* ============================================================
   EXTRAÇÃO OTIMIZADA

   IDENTIFICAÇÃO DAS COLUNAS EM BASE SECA
============================================================ */

function normalizarCabecalhoProducao(
    valor
){

    return String(
        valor ?? ""
    )
    .normalize(
        "NFD"
    )
    .replace(
        /[\u0300-\u036f]/g,
        ""
    )
    .toLowerCase()
    .replace(
        /&/g,
        ""
    )
    .replace(
        /[^a-z0-9]+/g,
        " "
    )
    .trim();

}


function identificarColunasBaseSeca(
    planilha
){

    if(
        !planilha ||
        !planilha["!ref"]
    ){

        throw new Error(
            "A aba 'Indicadores Britagem' está vazia."
        );

    }


    /*
        MAPEAMENTO VALIDADO DA BASE DIÁRIA

        A = Data
        B = BAP seco
        C = S&OP seco
        E = Realizado seco

        IMPORTANTE:
        A aba também possui colunas de MTD/YTD mais à direita.
        Não devemos procurar o último cabeçalho parecido na planilha,
        porque isso pode fazer o sistema ler uma série já acumulada e
        acumulá-la novamente.

        Por isso, validamos SOMENTE as colunas A/B/C/E.
    */

    const colunas = {
        data:COLUNAS_SECAS_ESPERADAS.data,
        bap:COLUNAS_SECAS_ESPERADAS.bap,
        sop:COLUNAS_SECAS_ESPERADAS.sop,
        real:COLUNAS_SECAS_ESPERADAS.real
    };


    const faixa =
        XLSX.utils.decode_range(
            planilha["!ref"]
        );


    const ultimaLinhaCabecalho =
        Math.min(
            faixa.e.r,
            faixa.s.r + 10
        );


    for(
        let linha = faixa.s.r;
        linha <= ultimaLinhaCabecalho;
        linha++
    ){

        const cabecalhos = {};


        Object.keys(
            colunas
        )
        .forEach(
            function(chave){

                const celula =
                    planilha[
                        XLSX.utils.encode_cell(
                            {
                                r:linha,
                                c:colunas[chave]
                            }
                        )
                    ];


                cabecalhos[chave] =
                    normalizarCabecalhoProducao(
                        celula ? celula.v : ""
                    );

            }
        );


        const dataValida =
            cabecalhos.data === "data" ||
            cabecalhos.data.startsWith("data ");


        const bapValido =
            /\bbap\b/.test(
                cabecalhos.bap
            ) &&
            /\bseca\b/.test(
                cabecalhos.bap
            ) &&
            !/\bumida\b/.test(
                cabecalhos.bap
            );


        const sopValido =
            /\bsop\b/.test(
                cabecalhos.sop
            ) &&
            /\bseca\b/.test(
                cabecalhos.sop
            ) &&
            !/\bumida\b/.test(
                cabecalhos.sop
            );


        const realValido =
            /\breal(?:izada|izado)?\b/.test(
                cabecalhos.real
            ) &&
            /\bseca\b/.test(
                cabecalhos.real
            ) &&
            !/\bumida\b/.test(
                cabecalhos.real
            );


        if(
            dataValida &&
            bapValido &&
            sopValido &&
            realValido
        ){

            console.info(
                "Colunas MTD/YTD validadas na base diária:",
                {
                    data:"A",
                    bap:"B",
                    sop:"C",
                    real:"E"
                }
            );


            return {
                ...colunas
            };

        }

    }


    throw new Error(
        "Não foi possível validar a base diária da aba 'Indicadores Britagem'. O portal espera: A = Data, B = BAP seco, C = S&OP seco e E = Realizado seco. O cálculo MTD/YTD foi interrompido para evitar utilizar colunas acumuladas ou valores em base úmida."
    );

}

/* ============================================================
   EXTRAÇÃO OTIMIZADA

   Lê exclusivamente as quatro colunas identificadas acima.
============================================================ */

function extrairBaseProducaoDiaria(
    planilha,
    colunas
){

    const base = [];


    if(
        !planilha ||
        !planilha["!ref"]
    ){

        return base;

    }


    const faixa =
        XLSX.utils.decode_range(
            planilha["!ref"]
        );


    /*
        Começamos em 0 porque a função
        automaticamente ignora cabeçalhos
        que não forem datas.
    */

    for(
        let linha = faixa.s.r;
        linha <= faixa.e.r;
        linha++
    ){

        const celulaData =
            planilha[
                XLSX.utils.encode_cell(
                    {
                        r:linha,
                        c:colunas.data
                    }
                )
            ];


        if(!celulaData){

            continue;

        }


        const data =
            converterDataExcel(
                celulaData.v
            );


        if(!data){

            continue;

        }


        const bap =
            obterValorCelula(
                planilha,
                linha,
                colunas.bap
            );


        const sop =
            obterValorCelula(
                planilha,
                linha,
                colunas.sop
            );


        const real =
            obterValorCelula(
                planilha,
                linha,
                colunas.real
            );


        /*
            Ignorar linhas sem qualquer
            dado de produção/meta.
        */

        if(
            bap === null &&
            sop === null &&
            real === null
        ){

            continue;

        }


        zerarHorario(
            data
        );


        base.push({

            data:data,

            bap:bap,

            sop:sop,

            real:real

        });

    }


    base.sort(
        function(a,b){

            return (
                a.data -
                b.data
            );

        }
    );


    return base;

}


/* ============================================================
   VALOR DE UMA CÉLULA
============================================================ */

function obterValorCelula(
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


    if(!celula){

        return null;

    }


    return valorNumero(
        celula.v
    );

}


/* ============================================================
   DEFINIR ÚLTIMA DATA REAL
============================================================ */

function definirUltimaDataReal(){

    let ultimaData =
        null;


    for(
        const item
        of baseProducaoDiaria
    ){

        if(
            item.real !== null
        ){

            ultimaData =
                item.data;

        }

    }


    if(!ultimaData){

        return;

    }


    const campo =
        document
        .getElementById(
            "dataReferencia"
        );


    campo.value =
        formatarDataInput(
            ultimaData
        );


    campo.max =
        formatarDataInput(
            ultimaData
        );

}


/* ============================================================
   INDICADORES ORIGINAIS
============================================================ */

function gerarIndicadores(){

    if(
        abaFechamento.length === 0
    ){

        alert(
            "Carregue uma planilha primeiro."
        );

        return;

    }


    const mesTexto =
        document
        .getElementById(
            "mesSelecionado"
        )
        .value;


    const coluna =
        meses[
            mesTexto
        ];


    if(
        tipoVisualizacao ===
        "mensal"
    ){

        gerarMensal(
            coluna,
            mesTexto
        );

    }else{

        gerarAcumulado(
            coluna,
            mesTexto
        );

    }


    gerarGraficoProducao(
        coluna
    );


    gerarResumo(
        coluna,
        mesTexto
    );


    gerarInsights(
        coluna,
        mesTexto
    );

}


/* ============================================================
   MENSAL
============================================================ */

function gerarMensal(
    coluna,
    mesTexto
){

    const real =
        valorNumero(
            abaFechamento[linhasProducaoFechamento.real][coluna]
        );


    const bap =
        valorNumero(
            abaFechamento[linhasProducaoFechamento.bap][coluna]
        );


    const sop =
        valorNumero(
            abaFechamento[linhasProducaoFechamento.sop][coluna]
        );


    preencherProducao(
        real,
        bap,
        sop
    );


    preencherQualidade(
        coluna,
        false
    );


    preencherPerformance(
        coluna,
        false
    );


    setTexto(
        "periodoGrafico",
        "Mensal até " +
        mesTexto
    );


    setTexto(
        "tituloResumoMes",
        "Resumo - " +
        mesTexto
    );

}


/* ============================================================
   ACUMULADO
============================================================ */

function gerarAcumulado(
    coluna,
    mesTexto
){

    const real =
        somarLinha(
            abaFechamento[linhasProducaoFechamento.real],
            coluna
        );


    const bap =
        somarLinha(
            abaFechamento[linhasProducaoFechamento.bap],
            coluna
        );


    const sop =
        somarLinha(
            abaFechamento[linhasProducaoFechamento.sop],
            coluna
        );


    preencherProducao(
        real,
        bap,
        sop
    );


    preencherQualidade(
        coluna,
        true
    );


    preencherPerformance(
        coluna,
        true
    );


    setTexto(
        "periodoGrafico",
        "Acumulado até " +
        mesTexto
    );


    setTexto(
        "tituloResumoMes",
        "Resumo Acumulado - " +
        mesTexto
    );

}


/* ============================================================
   PRODUÇÃO
============================================================ */

function preencherProducao(
    real,
    bap,
    sop
){

    setTexto(
        "producaoReal",
        formatarToneladas(
            real
        )
    );


    setTexto(
        "metaBap",
        formatarToneladas(
            bap
        )
    );


    setTexto(
        "metaSop",
        formatarToneladas(
            sop
        )
    );


    const atingBAP =
        calcularAtingimento(
            real,
            bap
        );


    const atingSOP =
        calcularAtingimento(
            real,
            sop
        );


    setTexto(
        "atingimentoBAP",
        formatarAtingimento(
            atingBAP
        )
    );


    setTexto(
        "atingimentoSOP",
        formatarAtingimento(
            atingSOP
        )
    );


    setTexto(
        "atingimentoBAPAcumulado",

        calcularAtingimentoAcumulado(
            abaFechamento[linhasProducaoFechamento.real],
            abaFechamento[linhasProducaoFechamento.bap]
        )
    );


    setTexto(
        "atingimentoSOPAcumulado",

        calcularAtingimentoAcumulado(
            abaFechamento[linhasProducaoFechamento.real],
            abaFechamento[linhasProducaoFechamento.sop]
        )
    );


    aplicarStatus(
        "cardAtingimentoBAP",
        atingBAP
    );


    aplicarStatus(
        "cardAtingimentoSOP",
        atingSOP
    );


    alterarCorStatus(
        "statusBAP",
        atingBAP
    );


    alterarCorStatus(
        "statusSOP",
        atingSOP
    );

}


/* ============================================================
   QUALIDADE
============================================================ */

function preencherQualidade(
    coluna,
    acumulado
){

    let p205;
    let nb;
    let bao;
    let gran;


    if(acumulado){

        p205 =
            mediaLinha(
                abaFechamento[6],
                coluna
            );


        nb =
            mediaLinha(
                abaFechamento[22],
                coluna
            );


        bao =
            mediaLinha(
                abaFechamento[24],
                coluna
            );


        gran =
            mediaLinha(
                abaFechamento[11],
                coluna
            );

    }else{

        p205 =
            valorNumero(
                abaFechamento[6][coluna]
            );


        nb =
            valorNumero(
                abaFechamento[22][coluna]
            );


        bao =
            valorNumero(
                abaFechamento[24][coluna]
            );


        gran =
            valorNumero(
                abaFechamento[11][coluna]
            );

    }


    const metaP205 =
        valorNumero(
            abaFechamento[5][coluna]
        );


    const metaNb =
        valorNumero(
            abaFechamento[21][coluna]
        );


    const metaBao =
        valorNumero(
            abaFechamento[23][coluna]
        );


    const metaGran =
        valorNumero(
            abaFechamento[10][coluna]
        );


    setTexto(
        "p205",
        formatarPercentual(
            p205
        )
    );


    setTexto(
        "nb",
        formatarPercentual(
            nb
        )
    );


    setTexto(
        "bao",
        formatarPercentual(
            bao
        )
    );


    setTexto(
        "granulo",
        formatarPercentual(
            gran
        )
    );


    setTexto(
        "metaP205",
        formatarPercentual(
            metaP205
        )
    );


    setTexto(
        "metaNb",
        formatarPercentual(
            metaNb
        )
    );


    setTexto(
        "metaBao",
        formatarPercentual(
            metaBao
        )
    );


    setTexto(
        "metaGranulo",
        formatarPercentual(
            metaGran
        )
    );


    aplicarStatusComparacao(
        "cardP205",
        p205,
        metaP205,
        "maior"
    );


    aplicarStatusComparacao(
        "cardNb",
        nb,
        metaNb,
        "maior"
    );


    aplicarStatusComparacao(
        "cardBao",
        bao,
        metaBao,
        "menor"
    );


    aplicarStatusComparacao(
        "cardGranulo",
        gran,
        metaGran,
        "menor"
    );

}


/* ============================================================
   PERFORMANCE
============================================================ */

function preencherPerformance(
    coluna,
    acumulado
){

    let disponibilidade;
    let utilizacao;
    let produtividade;


    if(acumulado){

        disponibilidade =
            mediaLinha(
                abaFechamento[14],
                coluna
            );


        utilizacao =
            mediaLinha(
                abaFechamento[17],
                coluna
            );


        produtividade =
            mediaLinha(
                abaFechamento[20],
                coluna
            );

    }else{

        disponibilidade =
            valorNumero(
                abaFechamento[14][coluna]
            );


        utilizacao =
            valorNumero(
                abaFechamento[17][coluna]
            );


        produtividade =
            valorNumero(
                abaFechamento[20][coluna]
            );

    }


    const metaDisp =
        valorNumero(
            abaFechamento[12][coluna]
        );


    const metaUtil =
        valorNumero(
            abaFechamento[15][coluna]
        );


    const metaProd =
        valorNumero(
            abaFechamento[18][coluna]
        );


    setTexto(
        "disponibilidade",
        formatarPercentual(
            disponibilidade
        )
    );


    setTexto(
        "utilizacao",
        formatarPercentual(
            utilizacao
        )
    );


    setTexto(
        "produtividade",
        formatarNumero(
            produtividade
        )
    );


    setTexto(
        "metaDisponibilidade",
        formatarPercentual(
            metaDisp
        )
    );


    setTexto(
        "metaUtilizacao",
        formatarPercentual(
            metaUtil
        )
    );


    setTexto(
        "metaProdutividade",
        formatarNumero(
            metaProd
        )
    );


    atualizarBarra(
        "barraDisponibilidade",
        disponibilidade,
        metaDisp
    );


    atualizarBarra(
        "barraUtilizacao",
        utilizacao,
        metaUtil
    );


    atualizarBarra(
        "barraProdutividade",
        produtividade,
        metaProd
    );


    aplicarStatusComparacao(
        "cardDisponibilidade",
        disponibilidade,
        metaDisp,
        "maior"
    );


    aplicarStatusComparacao(
        "cardUtilizacao",
        utilizacao,
        metaUtil,
        "maior"
    );


    aplicarStatusComparacao(
        "cardProdutividade",
        produtividade,
        metaProd,
        "maior"
    );

}


/* ============================================================
   GRÁFICO ORIGINAL
============================================================ */

function gerarGraficoProducao(
    colunaSelecionada
){

    const mesesGrafico =
        obterMesesAteColuna(
            colunaSelecionada
        );


    const labels =
        mesesGrafico.map(
            function(item){
                return item.abreviado;
            }
        );


    const real = [];
    const bap = [];
    const sop = [];


    let acumReal = 0;
    let acumBAP = 0;
    let acumSOP = 0;


    for(
        const itemMes
        of mesesGrafico
    ){

        const coluna =
            itemMes.coluna;

        const valorReal =
            valorNumero(
                abaFechamento[linhasProducaoFechamento.real][coluna]
            );


        const valorBAP =
            valorNumero(
                abaFechamento[linhasProducaoFechamento.bap][coluna]
            );


        const valorSOP =
            valorNumero(
                abaFechamento[linhasProducaoFechamento.sop][coluna]
            );


        if(
            tipoVisualizacao ===
            "mensal"
        ){

            real.push(
                valorReal
            );


            bap.push(
                valorBAP
            );


            sop.push(
                valorSOP
            );

        }else{

            acumReal +=
                valorReal || 0;


            acumBAP +=
                valorBAP || 0;


            acumSOP +=
                valorSOP || 0;


            real.push(
                acumReal
            );


            bap.push(
                acumBAP
            );


            sop.push(
                acumSOP
            );

        }

    }


    if(
        graficoProducao
    ){

        graficoProducao.destroy();

    }


    graficoProducao =
        new Chart(
            document.getElementById(
                "graficoProducao"
            ),
            {

                data:{

                    labels:labels,

                    datasets:[

                        {

                            type:"bar",

                            label:
                                "Produção Real",

                            data:real,

                            backgroundColor:
                                "#4B076C",

                            borderColor:
                                "#2D014D",

                            borderWidth:1,

                            borderRadius:7,

                            borderSkipped:false,

                            order:3,

                            maxBarThickness:52,

                            barPercentage:
                                0.62

                        },

                        {

                            type:"line",

                            label:"BAP",

                            data:bap,

                            borderColor:
                                "#79B832",

                            backgroundColor:
                                "#92D050",

                            borderWidth:3.5,

                            pointRadius:3.5,

                            pointHoverRadius:6,

                            pointBackgroundColor:
                                "#FFFFFF",

                            pointBorderColor:
                                "#79B832",

                            pointBorderWidth:2,

                            order:1,

                            tension:0.25

                        },

                        {

                            type:"line",

                            label:"S&OP",

                            data:sop,

                            borderColor:
                                "#EE9622",

                            backgroundColor:
                                "#F39C12",

                            borderWidth:3.5,

                            pointRadius:3.5,

                            pointHoverRadius:6,

                            pointBackgroundColor:
                                "#FFFFFF",

                            pointBorderColor:
                                "#EE9622",

                            pointBorderWidth:2,

                            order:1,

                            tension:0.25

                        }

                    ]

                },


                options:{

                    responsive:true,

                    maintainAspectRatio:false,


                    interaction:{

                        mode:"index",

                        intersect:false

                    },


                    plugins:{

                        legend:{

                            position:"top",

                            align:"end",

                            labels:{

                                usePointStyle:true,

                                pointStyle:"circle",

                                padding:18,

                                boxWidth:10,

                                boxHeight:10,

                                font:{
                                    size:12,
                                    weight:"700"
                                }

                            }

                        },


                        tooltip:{

                            callbacks:{

                                label:
                                function(context){

                                    return (

                                        context.dataset.label +

                                        ": " +

                                        formatarToneladas(
                                            context.raw
                                        ) +

                                        " t"

                                    );

                                }

                            }

                        }

                    },


                    scales:{

                        x:{

                            grid:{

                                display:false

                            },

                            ticks:{
                                color:"#5F5663",
                                font:{
                                    size:12,
                                    weight:"700"
                                }
                            }

                        },


                        y:{

                            beginAtZero:false,

                            grace:"10%",


                            grid:{
                                color:"rgba(45,1,77,0.08)",
                                drawBorder:false
                            },

                            ticks:{

                                color:"#6B626E",

                                font:{
                                    size:11,
                                    weight:"600"
                                },

                                callback:
                                function(valor){

                                    return Number(
                                        valor
                                    )
                                    .toLocaleString(
                                        "pt-BR"
                                    );

                                }

                            }

                        }

                    }

                }

            }
        );

}


/* ============================================================
   RESUMO ORIGINAL
============================================================ */

function gerarResumo(
    coluna,
    mesTexto
){

    let real;
    let bap;
    let sop;


    if(
        tipoVisualizacao ===
        "mensal"
    ){

        real =
            valorNumero(
                abaFechamento[linhasProducaoFechamento.real][coluna]
            );


        bap =
            valorNumero(
                abaFechamento[linhasProducaoFechamento.bap][coluna]
            );


        sop =
            valorNumero(
                abaFechamento[linhasProducaoFechamento.sop][coluna]
            );

    }else{

        real =
            somarLinha(
                abaFechamento[linhasProducaoFechamento.real],
                coluna
            );


        bap =
            somarLinha(
                abaFechamento[linhasProducaoFechamento.bap],
                coluna
            );


        sop =
            somarLinha(
                abaFechamento[linhasProducaoFechamento.sop],
                coluna
            );

    }


    setTexto(
        "desvioBAP",

        formatarDesvio(
            diferenca(
                real,
                bap
            )
        )
    );


    setTexto(
        "desvioSOP",

        formatarDesvio(
            diferenca(
                real,
                sop
            )
        )
    );


    const melhor =
        localizarMelhorMes(
            coluna
        );


    setTexto(
        "melhorMes",
        melhor.nome
    );


    setTexto(
        "recordeAno",

        formatarToneladas(
            melhor.valor
        ) +
        " t"
    );

}


/* ============================================================
   INSIGHTS
============================================================ */

function gerarInsights(
    coluna,
    mesTexto
){

    const real =
        valorNumero(
            abaFechamento[linhasProducaoFechamento.real][coluna]
        );


    const bap =
        valorNumero(
            abaFechamento[linhasProducaoFechamento.bap][coluna]
        );


    const ating =
        calcularAtingimento(
            real,
            bap
        );


    setTexto(
        "textoInsightProducao",

        ating !== null
        ?
        (
            "Produção atingiu " +
            ating.toFixed(1) +
            "% do BAP em " +
            mesTexto +
            "."
        )
        :
        "-"
    );


    setTexto(
        "textoInsightQualidade",

        compararTexto(

            valorNumero(
                abaFechamento[6][coluna]
            ),

            valorNumero(
                abaFechamento[5][coluna]
            ),

            "P2O5",

            "maior"

        )
    );


    setTexto(
        "textoInsightDisponibilidade",

        compararTexto(

            valorNumero(
                abaFechamento[14][coluna]
            ),

            valorNumero(
                abaFechamento[12][coluna]
            ),

            "Disponibilidade",

            "maior"

        )
    );


    setTexto(
        "textoInsightProdutividade",

        compararTexto(

            valorNumero(
                abaFechamento[20][coluna]
            ),

            valorNumero(
                abaFechamento[18][coluna]
            ),

            "Produtividade",

            "maior"

        )
    );

}


/* ============================================================
   MTD / YTD
============================================================ */

function gerarMTDYTD(){

    if(
        baseProducaoDiaria.length === 0
    ){

        return;

    }


    const dataReferencia =
        obterDataReferencia();


    if(!dataReferencia){

        return;

    }


    let dataInicio;


    if(
        tipoProducaoAcumulada ===
        "mtd"
    ){

        dataInicio =
            new Date(

                dataReferencia
                .getFullYear(),

                dataReferencia
                .getMonth(),

                1

            );

    }else{

        dataInicio =
            new Date(

                dataReferencia
                .getFullYear(),

                0,

                1

            );

    }


    zerarHorario(
        dataInicio
    );


    /*
        Agora FILTRAMOS o array compacto
        em memória.

        Não existe mais leitura do Excel aqui.
    */

    const registros =
        baseProducaoDiaria
        .filter(
            function(item){

                return (

                    item.data >=
                    dataInicio

                    &&

                    item.data <=
                    dataReferencia

                );

            }
        );


    if(
        registros.length === 0
    ){

        limparMTDYTD();

        return;

    }


    const acumulados =
        gerarAcumuladosDiarios(
            registros
        );


    const ultimo =
        acumulados[
            acumulados.length - 1
        ];


    preencherCardsMTDYTD(

        ultimo.real,

        ultimo.sop,

        ultimo.bap

    );


    gerarGraficoMTDYTD(
        acumulados
    );


    atualizarResumoMTDYTD(

        dataInicio,

        dataReferencia,

        ultimo

    );

}


/* ============================================================
   ACUMULADOS
============================================================ */

function gerarAcumuladosDiarios(
    registros
){

    let real =
        0;


    let bap =
        0;


    let sop =
        0;


    return registros.map(
        function(item){

            real +=
                item.real || 0;


            bap +=
                item.bap || 0;


            sop +=
                item.sop || 0;


            return {

                data:
                    item.data,

                real:
                    real,

                bap:
                    bap,

                sop:
                    sop

            };

        }
    );

}


/* ============================================================
   CARDS MTD/YTD
============================================================ */

function preencherCardsMTDYTD(
    real,
    sop,
    bap
){

    setTexto(
        "producaoAcumuladaReal",

        formatarToneladas(
            real
        )
    );


    setTexto(
        "producaoAcumuladaSOP",

        formatarToneladas(
            sop
        )
    );


    setTexto(
        "producaoAcumuladaBAP",

        formatarToneladas(
            bap
        )
    );


    const atingSOP =
        calcularAtingimento(
            real,
            sop
        );


    const atingBAP =
        calcularAtingimento(
            real,
            bap
        );


    setTexto(
        "atingimentoAcumuladoSOP",

        formatarAtingimento(
            atingSOP
        )
    );


    setTexto(
        "atingimentoAcumuladoBAP",

        formatarAtingimento(
            atingBAP
        )
    );


    setTexto(
        "gapAcumuladoSOP",

        formatarDesvio(
            diferenca(
                real,
                sop
            )
        )
    );


    setTexto(
        "gapAcumuladoBAP",

        formatarDesvio(
            diferenca(
                real,
                bap
            )
        )
    );


    aplicarStatus(
        "cardAtingimentoAcumuladoSOP",
        atingSOP
    );


    aplicarStatus(
        "cardAtingimentoAcumuladoBAP",
        atingBAP
    );


    atualizarStatusTexto(
        "statusAcumuladoSOP",
        atingSOP
    );


    atualizarStatusTexto(
        "statusAcumuladoBAP",
        atingBAP
    );

}


/* ============================================================
   GRÁFICO MTD/YTD
============================================================ */

function gerarGraficoMTDYTD(
    dados
){

    const labels =
        dados.map(
            function(item){

                return item.data
                .toLocaleDateString(
                    "pt-BR",
                    {
                        day:"2-digit",
                        month:"2-digit"
                    }
                );

            }
        );


    const real =
        dados.map(
            item =>
                item.real
        );


    const sop =
        dados.map(
            item =>
                item.sop
        );


    const bap =
        dados.map(
            item =>
                item.bap
        );


    if(
        graficoMTDYTD
    ){

        graficoMTDYTD.destroy();

    }


    graficoMTDYTD =
        new Chart(
            document.getElementById(
                "graficoMTDYTD"
            ),
            {

                type:"line",


                data:{

                    labels:labels,


                    datasets:[

                        {

                            label:
                                "Realizado seco",

                            data:
                                real,

                            borderColor:
                                "#3B005F",

                            backgroundColor:
                                "#3B005F",

                            pointBorderColor:
                                "#3B005F",

                            borderWidth:
                                4,

                            pointRadius:
                                2.2,

                            pointHoverRadius:
                                6,

                            pointBackgroundColor:
                                "#FFFFFF",

                            pointBorderWidth:
                                1.8,

                            order:
                                1,

                            tension:
                                0.20

                        },

                        {

                            label:
                                "S&OP seco",

                            data:
                                sop,

                            borderColor:
                                "#F39C12",

                            backgroundColor:
                                "#EE9622",

                            pointBorderColor:
                                "#EE9622",

                            borderWidth:
                                3.2,

                            pointRadius:
                                2,

                            pointHoverRadius:
                                5,

                            pointBackgroundColor:
                                "#FFFFFF",

                            pointBorderWidth:
                                1.8,

                            order:
                                1,

                            tension:
                                0.20

                        },

                        {

                            label:
                                "BAP seco",

                            data:
                                bap,

                            borderColor:
                                "#92D050",

                            backgroundColor:
                                "#79B832",

                            pointBorderColor:
                                "#79B832",

                            borderWidth:
                                3.2,

                            pointRadius:
                                2,

                            pointHoverRadius:
                                5,

                            pointBackgroundColor:
                                "#FFFFFF",

                            pointBorderWidth:
                                1.8,

                            order:
                                1,

                            tension:
                                0.20

                        }

                    ]

                },


                options:{

                    responsive:true,

                    maintainAspectRatio:false,


                    interaction:{

                        mode:"index",

                        intersect:false

                    },


                    animation:{

                        duration:350

                    },


                    plugins:{

                        legend:{

                            position:
                                "top",

                            labels:{

                                usePointStyle:
                                    true,

                                pointStyle:
                                    "circle",

                                padding:
                                    18,

                                boxWidth:
                                    10,

                                boxHeight:
                                    10,

                                font:{
                                    size:12,
                                    weight:"700"
                                }

                            }

                        },


                        tooltip:{

                            callbacks:{

                                title:
                                function(context){

                                    return (
                                        "Data: " +
                                        context[0].label
                                    );

                                },


                                label:
                                function(context){

                                    return (

                                        context.dataset.label +

                                        ": " +

                                        formatarToneladas(
                                            context.raw
                                        ) +

                                        " t seca"

                                    );

                                }

                            }

                        }

                    },


                    scales:{

                        x:{

                            grid:{

                                display:false

                            },


                            ticks:{

                                color:"#625966",

                                font:{
                                    size:10,
                                    weight:"600"
                                },

                                autoSkip:true,

                                maxTicksLimit:
                                    tipoProducaoAcumulada ===
                                    "ytd"
                                    ?
                                    15
                                    :
                                    31,

                                maxRotation:
                                    0

                            }

                        },


                        y:{

                            beginAtZero:true,


                            title:{

                                display:true,

                                text:
                                    "Produção acumulada (tbs)"

                            },


                            ticks:{

                                callback:
                                function(valor){

                                    return Number(
                                        valor
                                    )
                                    .toLocaleString(
                                        "pt-BR"
                                    );

                                }

                            }

                        }

                    }

                }

            }
        );

}


/* ============================================================
   RESUMO MTD/YTD
============================================================ */

function atualizarResumoMTDYTD(
    inicio,
    fim,
    ultimo
){

    const nome =
        tipoProducaoAcumulada ===
        "mtd"
        ?
        "MTD"
        :
        "YTD";


    setTexto(
        "tituloGraficoMTDYTD",

        "Produção " +
        nome
    );


    setTexto(
        "periodoGraficoMTDYTD",

        formatarDataBR(
            inicio
        ) +

        " até " +

        formatarDataBR(
            fim
        )
    );


    const atingSOP =
        calcularAtingimento(

            ultimo.real,

            ultimo.sop

        );


    const atingBAP =
        calcularAtingimento(

            ultimo.real,

            ultimo.bap

        );


    const status =
        document
        .getElementById(
            "statusGeralProducao"
        );


    if(
        atingSOP !== null &&
        atingSOP >= 100
    ){

        status.textContent =
            "META S&OP ATINGIDA";


        status.className =
            "status-producao-ok";


        setTexto(
            "textoResumoMTDYTD",

            "A produção acumulada está acima da meta S&OP. " +

            "O atingimento do S&OP é de " +

            formatarAtingimento(
                atingSOP
            ) +

            " e o atingimento do BAP é de " +

            formatarAtingimento(
                atingBAP
            ) +

            "."

        );

    }else if(
        atingSOP !== null &&
        atingSOP >= 95
    ){

        status.textContent =
            "PRÓXIMO DA META";


        status.className =
            "status-producao-atencao";


        setTexto(
            "textoResumoMTDYTD",

            "A produção acumulada está próxima da meta S&OP. " +

            "O atingimento do S&OP é de " +

            formatarAtingimento(
                atingSOP
            ) +

            " e o atingimento do BAP é de " +

            formatarAtingimento(
                atingBAP
            ) +

            "."

        );

    }else{

        status.textContent =
            "ABAIXO DA META";


        status.className =
            "status-producao-critico";


        setTexto(
            "textoResumoMTDYTD",

            "A produção acumulada está abaixo da meta S&OP. " +

            "O atingimento do S&OP é de " +

            formatarAtingimento(
                atingSOP
            ) +

            " e o atingimento do BAP é de " +

            formatarAtingimento(
                atingBAP
            ) +

            "."

        );

    }

}


/* ============================================================
   STATUS
============================================================ */

function atualizarStatusTexto(
    id,
    percentual
){

    if(
        percentual === null
    ){

        setTexto(
            id,
            "-"
        );

        return;

    }


    if(
        percentual >= 100
    ){

        setTexto(
            id,
            "Meta atingida"
        );

    }else if(
        percentual >= 95
    ){

        setTexto(
            id,
            "Próximo da meta"
        );

    }else{

        setTexto(
            id,
            "Abaixo da meta"
        );

    }

}


/* ============================================================
   PDF
============================================================ */

function gerarPDFKPI(){

    if(
        abaFechamento.length === 0
    ){

        alert(
            "Carregue a planilha antes de gerar o PDF."
        );

        return;

    }


    const tituloAnterior =
        document.title;


    const data =
        document
        .getElementById(
            "dataReferencia"
        )
        .value;


    const mes =
        document
        .getElementById(
            "mesSelecionado"
        )
        .value;


    const modoMensal =
        tipoVisualizacao === "acumulado"
        ?
        "Acumulado até " + mes
        :
        mes;


    const referencia =
        data
        ?
        formatarDataInputBR(
            data
        )
        :
        "Não selecionada";


    const agora =
        new Date();


    setTexto(
        "periodoRelatorioPDF",
        "Período: " + modoMensal
    );


    setTexto(
        "referenciaRelatorioPDF",
        tipoProducaoAcumulada.toUpperCase() +
        " • Referência: " +
        referencia
    );


    setTexto(
        "geradoEmRelatorioPDF",
        "Gerado em: " +
        agora.toLocaleString(
            "pt-BR",
            {
                day:"2-digit",
                month:"2-digit",
                year:"numeric",
                hour:"2-digit",
                minute:"2-digit"
            }
        )
    );


    document.title =
        "KPIs_Britagem_FVO_" +
        (
            data ||
            mes ||
            "Relatorio"
        );


    document.body
    .classList
    .add(
        "modo-impressao-kpi"
    );


    /*
       Atualiza os canvases imediatamente antes da impressão.
       Isso melhora a nitidez quando o navegador adapta o layout
       para A4 paisagem.
    */

    [
        graficoProducao,
        graficoMTDYTD
    ]
    .forEach(
        function(grafico){

            if(!grafico){
                return;
            }

            grafico.options.animation = false;
            grafico.resize();
            grafico.update(
                "none"
            );

        }
    );


    setTimeout(
        function(){

            window.print();


            setTimeout(
                function(){

                    document.body
                    .classList
                    .remove(
                        "modo-impressao-kpi"
                    );


                    document.title =
                        tituloAnterior;

                },
                350
            );

        },
        280
    );

}



/* ============================================================
   EMAIL
============================================================ */

function enviarEmailKPI(){

    if(
        abaFechamento.length === 0
    ){

        alert(
            "Carregue a planilha antes de enviar."
        );

        return;

    }


    const dataReferencia =
        document
        .getElementById(
            "dataReferencia"
        )
        .value;


    const periodo =
        tipoProducaoAcumulada ===
        "mtd"
        ?
        "MTD"
        :
        "YTD";


    const assunto =

        "KPIs Operacionais - Britagem FVO - " +

        periodo +

        " - " +

        formatarDataInputBR(
            dataReferencia
        );


    const mensagem =

`Prezados,

Segue o acompanhamento dos KPIs Operacionais da Britagem FVO.

Período: ${periodo}
Data de referência: ${formatarDataInputBR(dataReferencia)}

Produção Real: ${obterTexto("producaoAcumuladaReal")} t seca
S&OP: ${obterTexto("producaoAcumuladaSOP")} t seca
BAP: ${obterTexto("producaoAcumuladaBAP")} t seca

Atingimento S&OP: ${obterTexto("atingimentoAcumuladoSOP")}
Atingimento BAP: ${obterTexto("atingimentoAcumuladoBAP")}

Gap S&OP: ${obterTexto("gapAcumuladoSOP")}
Gap BAP: ${obterTexto("gapAcumuladoBAP")}

Status: ${obterTexto("statusGeralProducao")}

Atenciosamente.`;


    const destinatarios =
        emailsKPI
        .join(
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


/* ============================================================
   AUXILIARES NUMÉRICOS
============================================================ */

function valorNumero(
    valor
){

    if(
        valor === null ||
        valor === undefined ||
        valor === "" ||
        valor === "#DIV/0!"
    ){

        return null;

    }


    if(
        typeof valor ===
        "number"
    ){

        return Number.isFinite(
            valor
        )
        ?
        valor
        :
        null;

    }


    let texto =
        String(
            valor
        )
        .trim();


    if(
        texto.includes(",") &&
        !texto.includes(".")
    ){

        texto =
            texto.replace(
                ",",
                "."
            );

    }


    const numero =
        Number(
            texto
        );


    return Number.isFinite(
        numero
    )
    ?
    numero
    :
    null;

}


/* ============================================================
   SOMAR LINHA
============================================================ */

function somarLinha(
    linha,
    colunaFinal
){

    let total =
        0;


    let encontrou =
        false;


    for(
        const itemMes
        of obterMesesAteColuna(colunaFinal)
    ){

        const c =
            itemMes.coluna;

        const valor =
            valorNumero(
                linha[c]
            );


        if(
            valor !== null
        ){

            total +=
                valor;


            encontrou =
                true;

        }

    }


    return encontrou
    ?
    total
    :
    null;

}


/* ============================================================
   MÉDIA
============================================================ */

function mediaLinha(
    linha,
    colunaFinal
){

    let total =
        0;


    let quantidade =
        0;


    for(
        const itemMes
        of obterMesesAteColuna(colunaFinal)
    ){

        const c =
            itemMes.coluna;

        const valor =
            valorNumero(
                linha[c]
            );


        if(
            valor !== null
        ){

            total +=
                valor;


            quantidade++;

        }

    }


    return quantidade > 0
    ?
    total /
    quantidade
    :
    null;

}


/* ============================================================
   ATINGIMENTO
============================================================ */

function calcularAtingimento(
    real,
    meta
){

    if(
        real === null ||
        meta === null ||
        meta === 0
    ){

        return null;

    }


    return (

        real /
        meta

    ) * 100;

}


/* ============================================================
   ATINGIMENTO ACUMULADO
============================================================ */

function calcularAtingimentoAcumulado(
    linhaReal,
    linhaMeta
){

    let ultimaColuna =
        null;


    for(
        const itemMes
        of obterMesesAteColuna(meses.Dezembro)
    ){

        const c =
            itemMes.coluna;

        if(
            valorNumero(
                linhaReal[c]
            ) !== null
        ){

            ultimaColuna =
                c;

        }

    }


    if(
        ultimaColuna === null
    ){

        return "-";

    }


    return formatarAtingimento(

        calcularAtingimento(

            somarLinha(
                linhaReal,
                ultimaColuna
            ),

            somarLinha(
                linhaMeta,
                ultimaColuna
            )

        )

    );

}


/* ============================================================
   DIFERENÇA
============================================================ */

function diferenca(
    real,
    meta
){

    if(
        real === null ||
        meta === null
    ){

        return null;

    }


    return (
        real -
        meta
    );

}


/* ============================================================
   FORMATAÇÕES
============================================================ */

function formatarToneladas(
    valor
){

    const numero =
        valorNumero(
            valor
        );


    if(
        numero === null
    ){

        return "-";

    }


    return numero
    .toLocaleString(
        "pt-BR",
        {

            maximumFractionDigits:
                0

        }
    );

}


function formatarNumero(
    valor
){

    return formatarToneladas(
        valor
    );

}


function formatarPercentual(
    valor,
    casas = 2
){

    const numero =
        valorNumero(
            valor
        );


    if(
        numero === null
    ){

        return "-";

    }


    return (
        numero *
        100
    )
    .toFixed(
        casas
    ) +
    "%";

}


function formatarAtingimento(
    valor
){

    if(
        valor === null ||
        !Number.isFinite(
            valor
        )
    ){

        return "-";

    }


    return valor
    .toFixed(
        1
    )
    .replace(
        ".",
        ","
    ) +
    "%";

}


function formatarDesvio(
    valor
){

    if(
        valor === null
    ){

        return "-";

    }


    const sinal =
        valor >= 0
        ?
        "+"
        :
        "";


    return (

        sinal +

        formatarToneladas(
            valor
        ) +

        " t"

    );

}


/* ============================================================
   DATA
============================================================ */

function converterDataExcel(
    valor
){

    if(
        valor === null ||
        valor === undefined ||
        valor === ""
    ){

        return null;

    }


    if(
        valor instanceof
        Date
    ){

        return new Date(
            valor
        );

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


        if(
            !dataExcel
        ){

            return null;

        }


        return new Date(

            dataExcel.y,

            dataExcel.m - 1,

            dataExcel.d

        );

    }


    const texto =
        String(
            valor
        )
        .trim();


    const partes =
        texto.split(
            /[\/\-]/
        );


    if(
        partes.length !==
        3
    ){

        return null;

    }


    if(
        partes[0]
        .length ===
        4
    ){

        return new Date(

            Number(
                partes[0]
            ),

            Number(
                partes[1]
            ) - 1,

            Number(
                partes[2]
            )

        );

    }


    return new Date(

        Number(
            partes[2]
        ),

        Number(
            partes[1]
        ) - 1,

        Number(
            partes[0]
        )

    );

}


/* ============================================================
   DATA REFERÊNCIA
============================================================ */

function obterDataReferencia(){

    const valor =
        document
        .getElementById(
            "dataReferencia"
        )
        .value;


    if(!valor){

        return null;

    }


    const partes =
        valor
        .split(
            "-"
        );


    const data =
        new Date(

            Number(
                partes[0]
            ),

            Number(
                partes[1]
            ) - 1,

            Number(
                partes[2]
            )

        );


    zerarHorario(
        data
    );


    return data;

}


function zerarHorario(
    data
){

    data.setHours(
        0,
        0,
        0,
        0
    );


    return data;

}


function formatarDataInput(
    data
){

    const ano =
        data
        .getFullYear();


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

        ano +
        "-" +
        mes +
        "-" +
        dia

    );

}


function formatarDataBR(
    data
){

    return data
    .toLocaleDateString(
        "pt-BR"
    );

}


function formatarDataInputBR(
    valor
){

    if(!valor){

        return "-";

    }


    const partes =
        valor
        .split(
            "-"
        );


    return (

        partes[2] +
        "/" +
        partes[1] +
        "/" +
        partes[0]

    );

}


/* ============================================================
   STATUS CARD
============================================================ */

function aplicarStatus(
    idCard,
    percentual
){

    const card =
        document
        .getElementById(
            idCard
        );


    if(!card){

        return;

    }


    card
    .classList
    .remove(

        "status-ok",

        "status-atencao",

        "status-critico"

    );


    if(
        percentual === null
    ){

        return;

    }


    if(
        percentual >= 100
    ){

        card
        .classList
        .add(
            "status-ok"
        );

    }else if(
        percentual >= 95
    ){

        card
        .classList
        .add(
            "status-atencao"
        );

    }else{

        card
        .classList
        .add(
            "status-critico"
        );

    }

}


/* ============================================================
   STATUS COMPARAÇÃO
============================================================ */

function aplicarStatusComparacao(
    idCard,
    real,
    meta,
    regra
){

    const card =
        document
        .getElementById(
            idCard
        );


    if(!card){

        return;

    }


    card
    .classList
    .remove(

        "status-ok",

        "status-atencao",

        "status-critico"

    );


    if(
        real === null ||
        meta === null
    ){

        return;

    }


    const sucesso =
        regra ===
        "menor"
        ?
        real <= meta
        :
        real >= meta;


    if(sucesso){

        card
        .classList
        .add(
            "status-ok"
        );


        return;

    }


    const diferencaRelativa =
        Math.abs(
            real -
            meta
        )
        /
        Math.abs(
            meta
        );


    card
    .classList
    .add(

        diferencaRelativa <= 0.05
        ?
        "status-atencao"
        :
        "status-critico"

    );

}


/* ============================================================
   COR STATUS
============================================================ */

function alterarCorStatus(
    id,
    percentual
){

    const elemento =
        document
        .getElementById(
            id
        );


    if(!elemento){

        return;

    }


    if(
        percentual === null
    ){

        elemento.style.color =
            "#999";

    }else if(
        percentual >= 100
    ){

        elemento.style.color =
            "#92D050";

    }else if(
        percentual >= 95
    ){

        elemento.style.color =
            "#F4B400";

    }else{

        elemento.style.color =
            "#D9534F";

    }

}


/* ============================================================
   BARRA
============================================================ */

function atualizarBarra(
    id,
    real,
    meta
){

    const barra =
        document
        .getElementById(
            id
        );


    if(!barra){

        return;

    }


    if(
        real === null ||
        meta === null ||
        meta === 0
    ){

        barra.style.width =
            "0%";


        return;

    }


    const percentual =
        (
            real /
            meta
        ) *
        100;


    barra.style.width =
        Math.min(
            percentual,
            100
        ) +
        "%";


    if(
        percentual >= 100
    ){

        barra.style.background =
            "#92D050";

    }else if(
        percentual >= 95
    ){

        barra.style.background =
            "#F4B400";

    }else{

        barra.style.background =
            "#D9534F";

    }

}


/* ============================================================
   MELHOR MÊS
============================================================ */

function localizarMelhorMes(
    colunaFinal
){

    let maiorValor =
        null;


    let melhorMes =
        "-";


    for(
        const itemMes
        of obterMesesAteColuna(colunaFinal)
    ){

        const coluna =
            itemMes.coluna;

        const valor =
            valorNumero(
                abaFechamento[linhasProducaoFechamento.real][coluna]
            );


        if(
            valor !== null &&
            (
                maiorValor === null ||
                valor > maiorValor
            )
        ){

            maiorValor =
                valor;


            melhorMes =
                itemMes.abreviado;

        }

    }


    return {

        nome:
            melhorMes,

        valor:
            maiorValor

    };

}


/* ============================================================
   TEXTO COMPARAÇÃO
============================================================ */

function compararTexto(
    real,
    meta,
    nome,
    regra
){

    if(
        real === null ||
        meta === null
    ){

        return "-";

    }


    const bom =
        regra ===
        "menor"
        ?
        real <= meta
        :
        real >= meta;


    return (

        nome +

        (
            bom
            ?
            " está dentro ou acima da meta."
            :
            " está abaixo da meta definida."
        )

    );

}


/* ============================================================
   LIMPAR MTD/YTD
============================================================ */

function limparMTDYTD(){

    const ids = [

        "producaoAcumuladaReal",

        "producaoAcumuladaSOP",

        "producaoAcumuladaBAP",

        "atingimentoAcumuladoSOP",

        "atingimentoAcumuladoBAP",

        "gapAcumuladoSOP",

        "gapAcumuladoBAP"

    ];


    ids
    .forEach(
        function(id){

            setTexto(
                id,
                "-"
            );

        }
    );


    setTexto(
        "statusGeralProducao",
        "SEM DADOS"
    );


    setTexto(
        "textoResumoMTDYTD",

        "Não foram encontrados dados para o período selecionado."
    );


    if(
        graficoMTDYTD
    ){

        graficoMTDYTD.destroy();


        graficoMTDYTD =
            null;

    }

}


/* ============================================================
   TEXTO ELEMENTO
============================================================ */

function setTexto(
    id,
    texto
){

    const elemento =
        document
        .getElementById(
            id
        );


    if(elemento){

        elemento.textContent =
            texto;

    }

}


function obterTexto(
    id
){

    const elemento =
        document
        .getElementById(
            id
        );


    return elemento
    ?
    elemento
    .textContent
    .trim()
    :
    "-";

}


/* ============================================================
   DATA ATUALIZAÇÃO
============================================================ */

function atualizarData(){

    const agora =
        new Date();


    setTexto(

        "ultimaAtualizacao",

        agora
        .toLocaleString(
            "pt-BR",
            {

                day:
                    "2-digit",

                month:
                    "2-digit",

                year:
                    "numeric",

                hour:
                    "2-digit",

                minute:
                    "2-digit"

            }
        )

    );

}