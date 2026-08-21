/* ============================================================
   RELATÓRIO DE PILHAS
   MPU - BRITAGEM × AMG - USINA
============================================================ */


/* ============================================================
   INICIALIZAÇÃO SEGURA
============================================================ */

document.addEventListener(
    "DOMContentLoaded",
    inicializarRelatorioPilhas
);


function inicializarRelatorioPilhas(){

    conectarEvento(
        "btnCarregar",
        "click",
        carregarPlanilhas
    );


    conectarEvento(
        "btnConsultarPilha",
        "click",
        consultarPilha
    );


    conectarEvento(
        "btnGerarPDF",
        "click",
        gerarPDFRelatorioPilha
    );


    conectarEvento(
        "btnEnviarEmail",
        "click",
        enviarRelatorioPorEmail
    );


    atualizarDataSistema();

}


function conectarEvento(
    id,
    evento,
    funcao
){

    const elemento =
        document.getElementById(
            id
        );


    if(!elemento){

        console.warn(
            "Elemento não encontrado no Relatório de Pilhas:",
            id
        );

        return;

    }


    elemento.addEventListener(
        evento,
        funcao
    );

}


/* ============================================================
   VARIÁVEIS
============================================================ */

let graficoPilhas = null;

let dadosPilhas = [];

let workbookAMG = null;

let dadosAMG = {};

let pilhaSelecionadaAtual = null;


/* ============================================================
   STATUS VISUAL DAS BASES
============================================================ */

function atualizarStatusBasePilhas(texto, estado = "neutro"){

    const elemento = document.getElementById("statusBasePilhas");

    if(!elemento){
        return;
    }

    elemento.textContent = texto;
    elemento.dataset.estado = estado;
}


/* ============================================================
   PARÂMETROS
============================================================ */

const parametros = [

    {
        chave:"p205",
        nome:"P2O5",
        colunaMPU:24,
        abas:[
            "P2O5"
        ]
    },

    {
        chave:"mgo",
        nome:"MgO",
        colunaMPU:25,
        abas:[
            "MgO"
        ]
    },

    {
        chave:"nb",
        nome:"Nb2O5",
        colunaMPU:26,
        abas:[
            "Nb2O5"
        ]
    },

    {
        chave:"cap",
        nome:"CaO/P2O5",
        colunaMPU:27,
        abas:[
            "CaO P2O5",
            "CaO/P2O5",
            "CaO-P2O5"
        ]
    },

    {
        chave:"fe",
        nome:"Fe2O3",
        colunaMPU:29,
        abas:[
            "Fe2O3"
        ]
    },

    {
        chave:"si",
        nome:"SiO2",
        colunaMPU:30,
        abas:[
            "SiO2"
        ]
    },

    {
        chave:"al",
        nome:"Al2O3",
        colunaMPU:31,
        abas:[
            "Al2O3"
        ]
    },

    {
        chave:"ti",
        nome:"TiO2",
        colunaMPU:32,
        abas:[
            "TiO2"
        ]
    }

];


/*
    COLUNAS NA PLANILHA AMG

    DX = PILHA ATUAL
    EU = RESULTADO AMG

    índices base 0:
*/

const COLUNA_PILHA_AMG = 127;

const COLUNA_VALOR_AMG = 150;


/* ============================================================
   CARREGAR AS DUAS PLANILHAS
============================================================ */

async function carregarPlanilhas(){

    if(
        typeof XLSX === "undefined"
    ){

        alert(
            "A biblioteca de leitura do Excel não foi carregada. Verifique a conexão com a internet e atualize a página."
        );

        return;

    }

    atualizarStatusBasePilhas(
        "Preparando leitura...",
        "carregando"
    );

    const arquivoMPU =
        document
        .getElementById(
            "arquivoPilha"
        )
        .files[0];


    const arquivoAMG =
        document
        .getElementById(
            "arquivoAMG"
        )
        .files[0];


    if(!arquivoMPU){

        atualizarStatusBasePilhas("Selecione a base MPU", "erro");

        alert(
            "Selecione a planilha MPU / Britagem."
        );

        return;

    }


    if(!arquivoAMG){

        atualizarStatusBasePilhas("Selecione a base AMG", "erro");

        alert(
            "Selecione a planilha AMG / Usina."
        );

        return;

    }


    window.FVOCarregamento?.exibir("Lendo as planilhas MPU e AMG...");

    try{

        const resultados =
            await Promise.all([

                lerArquivoExcel(
                    arquivoMPU
                ),

                lerArquivoExcel(
                    arquivoAMG
                )

            ]);


        /* ============================
           MPU
        ============================ */

        const workbookMPU =
            resultados[0];


        const abaPilha =
            workbookMPU
            .Sheets[
                "Pilha"
            ];


        if(!abaPilha){

            alert(
                "A aba 'Pilha' não foi encontrada na planilha MPU."
            );

            return;

        }


        dadosPilhas =
            XLSX.utils
            .sheet_to_json(
                abaPilha,
                {
                    header:1,
                    defval:null,
                    raw:true
                }
            );


        /* ============================
           AMG
        ============================ */

        workbookAMG =
            resultados[1];


        dadosAMG =
            prepararDadosAMG(
                workbookAMG
            );


        /* ============================
           FINAL
        ============================ */

        gerarGraficoPilhas();

        atualizarDataSistema();

        atualizarStatusBasePilhas(
            "MPU + AMG carregadas",
            "ok"
        );


        alert(
            "Planilhas MPU e AMG carregadas com sucesso!"
        );

    }

    catch(erro){

        atualizarStatusBasePilhas(
            "Erro ao ler as bases",
            "erro"
        );

        console.error(
            erro
        );


        alert(
            "Não foi possível carregar uma das planilhas."
        );

    }finally{

        window.FVOCarregamento?.ocultar();

    }

}


/* ============================================================
   LER ARQUIVO EXCEL
============================================================ */

function lerArquivoExcel(
    arquivo
){

    return new Promise(
        function(resolve,reject){

            const reader =
                new FileReader();


            reader.onload =
                function(event){

                    try{

                        const dados =
                            new Uint8Array(
                                event.target.result
                            );


                        const workbook =
                            XLSX.read(
                                dados,
                                {
                                    type:"array",
                                    cellDates:true
                                }
                            );


                        resolve(
                            workbook
                        );

                    }

                    catch(erro){

                        reject(
                            erro
                        );

                    }

                };


            reader.onerror =
                reject;


            reader
            .readAsArrayBuffer(
                arquivo
            );

        }
    );

}


/* ============================================================
   PREPARAR DADOS AMG
============================================================ */

function prepararDadosAMG(
    workbook
){

    const resultado =
        {};


    parametros
    .forEach(
        function(parametro){

            const nomeAba =
                encontrarAba(
                    workbook,
                    parametro.abas
                );


            if(!nomeAba){

                console.warn(
                    "Aba AMG não encontrada:",
                    parametro.nome
                );


                resultado[
                    parametro.chave
                ] = [];


                return;

            }


            resultado[
                parametro.chave
            ] =
                extrairDadosAbaAMG(
                    workbook
                    .Sheets[
                        nomeAba
                    ]
                );

        }
    );


    return resultado;

}


/* ============================================================
   ENCONTRAR ABA
============================================================ */

function encontrarAba(
    workbook,
    nomesPossiveis
){

    const abas =
        workbook
        .SheetNames;


    for(
        const procurado
        of nomesPossiveis
    ){

        const normalizadoProcurado =
            normalizarTexto(
                procurado
            );


        const encontrado =
            abas.find(
                nome =>
                    normalizarTexto(
                        nome
                    ) ===
                    normalizadoProcurado
            );


        if(encontrado){

            return encontrado;

        }

    }


    /*
        Segunda tentativa:
        procura nome parcialmente.
    */

    for(
        const procurado
        of nomesPossiveis
    ){

        const normalizadoProcurado =
            normalizarTexto(
                procurado
            );


        const encontrado =
            abas.find(
                nome => {

                    const n =
                        normalizarTexto(
                            nome
                        );


                    return (
                        n.includes(
                            normalizadoProcurado
                        ) ||
                        normalizadoProcurado
                        .includes(
                            n
                        )
                    );

                }
            );


        if(encontrado){

            return encontrado;

        }

    }


    return null;

}


/* ============================================================
   EXTRAIR DX + EU
============================================================ */

function extrairDadosAbaAMG(
    aba
){

    if(
        !aba ||
        !aba["!ref"]
    ){

        return [];

    }


    const faixa =
        XLSX.utils
        .decode_range(
            aba["!ref"]
        );


    /*
        Tentamos localizar:

        EU = Acumulado Pilha - Usina 47+76
        EU = AMG

        Depois disso começam os resultados.
    */

    let linhaTitulo =
        -1;


    let linhaAMG =
        -1;


    const limiteCabecalho =
        Math.min(
            faixa.e.r,
            50
        );


    for(
        let r = 0;
        r <= limiteCabecalho;
        r++
    ){

        const celulaEU =
            aba[
                XLSX.utils
                .encode_cell(
                    {
                        c:
                            COLUNA_VALOR_AMG,

                        r:
                            r
                    }
                )
            ];


        if(!celulaEU){

            continue;

        }


        const texto =
            normalizarTexto(
                celulaEU.v
            );


        if(
            texto.includes(
                "acumulado pilha"
            ) &&
            texto.includes(
                "usina"
            )
        ){

            linhaTitulo =
                r;

        }


        if(
            linhaTitulo !== -1 &&
            texto === "amg"
        ){

            linhaAMG =
                r;

            break;

        }

    }


    /*
        Caso não ache o texto AMG,
        começa após o título.
    */

    let primeiraLinha =
        0;


    if(
        linhaAMG !== -1
    ){

        primeiraLinha =
            linhaAMG + 1;

    }

    else if(
        linhaTitulo !== -1
    ){

        primeiraLinha =
            linhaTitulo + 1;

    }


    const registros =
        [];


    for(
        let r =
            primeiraLinha;

        r <= faixa.e.r;

        r++
    ){

        const enderecoPilha =
            XLSX.utils
            .encode_cell(
                {
                    c:
                        COLUNA_PILHA_AMG,

                    r:
                        r
                }
            );


        const enderecoValor =
            XLSX.utils
            .encode_cell(
                {
                    c:
                        COLUNA_VALOR_AMG,

                    r:
                        r
                }
            );


        const celulaPilha =
            aba[
                enderecoPilha
            ];


        const celulaValor =
            aba[
                enderecoValor
            ];


        if(
            !celulaPilha ||
            !celulaValor
        ){

            continue;

        }


        const pilhaTexto =
            limparTexto(
                celulaPilha.v
            );


        const valor =
            valorNumeroFlexivel(
                celulaValor.v
            );


        if(
            !pilhaTexto ||
            valor === null
        ){

            continue;

        }


        const numeros =
            extrairNumerosPilha(
                pilhaTexto
            );


        if(
            numeros.length === 0
        ){

            continue;

        }


        registros.push({

            linha:
                r + 1,

            pilhaTexto:
                pilhaTexto,

            pilhas:
                numeros,

            valor:
                valor

        });

    }


    return registros;

}


/* ============================================================
   EXTRAIR NÚMEROS DA PILHA
============================================================ */

function extrairNumerosPilha(
    texto
){

    if(!texto){

        return [];

    }


    const encontrados =
        String(
            texto
        )
        .match(
            /\d+/g
        );


    if(!encontrados){

        return [];

    }


    return Array.from(

        new Set(

            encontrados
            .map(
                numero =>
                    Number(
                        numero
                    )
            )
            .filter(
                numero =>
                    Number.isFinite(
                        numero
                    )
            )

        )

    );

}


/* ============================================================
   OBTER AMG DE UMA PILHA
============================================================ */

function obterAMGPilha(
    numeroPilha
){

    const numero =
        Number(
            numeroPilha
        );


    const resultado =
        {};


    const fontesGerais =
        new Set();


    const fontesCompartilhadas =
        new Set();


    parametros
    .forEach(
        function(parametro){

            const registros =
                dadosAMG[
                    parametro.chave
                ] ||
                [];


            const encontrados =
                registros
                .filter(
                    registro =>
                        registro
                        .pilhas
                        .includes(
                            numero
                        )
                );


            encontrados
            .forEach(
                function(registro){

                    fontesGerais
                    .add(
                        registro.pilhaTexto
                    );


                    if(
                        Array.isArray(
                            registro.pilhas
                        ) &&
                        registro.pilhas.length > 1
                    ){

                        fontesCompartilhadas
                        .add(
                            registro.pilhaTexto
                        );

                    }

                }
            );


            const valores =
                encontrados
                .map(
                    registro =>
                        registro.valor
                )
                .filter(
                    valor =>
                        valor !== null
                );


            resultado[
                parametro.chave
            ] = {

                valor:
                    calcularMedia(
                        valores
                    ),

                quantidade:
                    valores.length,

                fontes:
                    Array.from(
                        new Set(
                            encontrados
                            .map(
                                registro =>
                                    registro.pilhaTexto
                            )
                        )
                    ),

                linhas:
                    encontrados
                    .map(
                        registro =>
                            registro.linha
                    )

            };

        }
    );


    resultado._meta = {

        pilha:
            numero,

        fontes:
            Array.from(
                fontesGerais
            )
            .sort(
                function(a,b){

                    return String(a)
                    .localeCompare(
                        String(b),
                        "pt-BR",
                        {
                            numeric:true,
                            sensitivity:"base"
                        }
                    );

                }
            ),

        fontesCompartilhadas:
            Array.from(
                fontesCompartilhadas
            )
            .sort(
                function(a,b){

                    return String(a)
                    .localeCompare(
                        String(b),
                        "pt-BR",
                        {
                            numeric:true,
                            sensitivity:"base"
                        }
                    );

                }
            ),

        temRetomadaCompartilhada:
            fontesCompartilhadas.size > 0

    };


    return resultado;

}



/* ============================================================
   CONSULTAR PILHA
============================================================ */

function consultarPilha(){

    if(
        dadosPilhas.length === 0
    ){

        alert(
            "Carregue as planilhas primeiro."
        );

        return;

    }


    const numero =
        document
        .getElementById(
            "numeroPilha"
        )
        .value;


    if(!numero){

        alert(
            "Digite o número da pilha."
        );

        return;

    }


    const localizacao =
        encontrarPilhaMPU(
            numero
        );


    if(!localizacao){

        alert(
            "Pilha não encontrada na planilha MPU."
        );

        return;

    }


    pilhaSelecionadaAtual =
        numero;


    const linha =
        localizacao.linha;


    const indice =
        localizacao.indice;


    /* ============================
       MPU
    ============================ */

    const mpu =
        obterDadosMPU(
            linha
        );


    /* ============================
       AMG
    ============================ */

    const amg =
        obterAMGPilha(
            numero
        );


    /* ============================
       PREENCHER TELA
    ============================ */

    preencherDadosGerais(
        linha
    );


    preencherMPU(
        mpu
    );


    preencherAMG(
        amg
    );


    preencherOrigemAMG(
        amg
    );


    preencherTabelaComparacao(
        mpu,
        amg
    );


    gerarResumoRapido(
        mpu,
        amg
    );


    gerarHistorico5Pilhas(
        indice
    );


    gerarInterpretacaoAMG(
        mpu,
        amg
    );


    gerarGraficoPilhas();

}


/* ============================================================
   ENCONTRAR PILHA NA MPU
============================================================ */

function encontrarPilhaMPU(
    numero
){

    for(
        let i = 1;
        i < dadosPilhas.length;
        i++
    ){

        const linha =
            dadosPilhas[i];


        if(
            linha &&
            String(
                linha[0]
            ) ===
            String(
                numero
            )
        ){

            return {

                linha:
                    linha,

                indice:
                    i

            };

        }

    }


    return null;

}


/* ============================================================
   DADOS MPU
============================================================ */

function obterDadosMPU(
    linha
){

    return {

        p205:
            valorNumero(
                linha[24]
            ),

        mgo:
            valorNumero(
                linha[25]
            ),

        nb:
            valorNumero(
                linha[26]
            ),

        cap:
            valorNumero(
                linha[27]
            ),

        fe:
            valorNumero(
                linha[29]
            ),

        si:
            valorNumero(
                linha[30]
            ),

        al:
            valorNumero(
                linha[31]
            ),

        ti:
            valorNumero(
                linha[32]
            )

    };

}


/* ============================================================
   PREENCHER DADOS GERAIS
============================================================ */

function preencherDadosGerais(
    linha
){

    const numero =
        linha[0];


    const massa =
        valorNumero(
            linha[18]
        );


    const inicio =
        converterDataExcel(
            linha[21]
        );


    const fim =
        converterDataExcel(
            linha[22]
        );


    const duracao =
        calcularDuracao(
            inicio,
            fim
        );


    document
    .getElementById(
        "pilhaSelecionada"
    )
    .textContent =
        numero;


    document
    .getElementById(
        "massaSelecionada"
    )
    .textContent =
        formatarToneladas(
            massa
        ) +
        " t";


    document
    .getElementById(
        "periodoFormacao"
    )
    .textContent =
        formatarDataCurta(
            inicio
        ) +
        " → " +
        formatarDataCurta(
            fim
        );


    document
    .getElementById(
        "duracaoFormacao"
    )
    .textContent =
        duracao;


    /* CONSULTA LATERAL */

    document
    .getElementById(
        "consultaNumeroPilha"
    )
    .textContent =
        numero;


    document
    .getElementById(
        "consultaDataInicio"
    )
    .textContent =
        formatarDataHora(
            inicio
        );


    document
    .getElementById(
        "consultaDataFim"
    )
    .textContent =
        formatarDataHora(
            fim
        );

}


/* ============================================================
   PREENCHER MPU
============================================================ */

function preencherMPU(
    mpu
){

    const mapa = {

        mpuP205:
            mpu.p205,

        mpuMgO:
            mpu.mgo,

        mpuNb:
            mpu.nb,

        mpuCaP:
            mpu.cap,

        mpuFe:
            mpu.fe,

        mpuSi:
            mpu.si,

        mpuAl:
            mpu.al,

        mpuTi:
            mpu.ti

    };


    Object
    .entries(
        mapa
    )
    .forEach(
        function(
            [id,valor]
        ){

            document
            .getElementById(
                id
            )
            .textContent =
                formatarQuimica(
                    valor
                );

        }
    );

}


/* ============================================================
   PREENCHER AMG
============================================================ */

function preencherAMG(
    amg
){

    const mapa = {

        amgP205:
            amg.p205.valor,

        amgMgO:
            amg.mgo.valor,

        amgNb:
            amg.nb.valor,

        amgCaP:
            amg.cap.valor,

        amgFe:
            amg.fe.valor,

        amgSi:
            amg.si.valor,

        amgAl:
            amg.al.valor,

        amgTi:
            amg.ti.valor

    };


    Object
    .entries(
        mapa
    )
    .forEach(
        function(
            [id,valor]
        ){

            document
            .getElementById(
                id
            )
            .textContent =
                formatarQuimica(
                    valor
                );

        }
    );

}


/* ============================================================
   ORIGEM / RASTREABILIDADE DOS DADOS AMG
============================================================ */

function preencherOrigemAMG(
    amg
){

    const meta =
        amg &&
        amg._meta
            ? amg._meta
            : {};


    const pilha =
        meta.pilha ??
        pilhaSelecionadaAtual ??
        "-";


    const fontes =
        Array.isArray(
            meta.fontes
        )
            ? meta.fontes
            : [];


    const compartilhadas =
        Array.isArray(
            meta.fontesCompartilhadas
        )
            ? meta.fontesCompartilhadas
            : [];


    const elementoPilha =
        document
        .getElementById(
            "amgPilhaConsultada"
        );


    const elementoFontes =
        document
        .getElementById(
            "amgFontesConsideradas"
        );


    const elementoTipo =
        document
        .getElementById(
            "amgTipoComposicao"
        );


    const elementoCriterio =
        document
        .getElementById(
            "amgCriterioTexto"
        );


    if(elementoPilha){

        elementoPilha.textContent =
            String(pilha);

    }


    if(elementoFontes){

        elementoFontes.textContent =
            fontes.length > 0
                ? fontes.join(" • ")
                : "Nenhum registro AMG associado";

    }


    if(elementoTipo){

        elementoTipo
        .classList
        .remove(
            "origem-individual",
            "origem-compartilhada",
            "origem-sem-dados"
        );


        if(fontes.length === 0){

            elementoTipo.textContent =
                "Sem dados AMG";

            elementoTipo
            .classList
            .add(
                "origem-sem-dados"
            );

        }

        else if(
            meta.temRetomadaCompartilhada
        ){

            elementoTipo.textContent =
                "Retomada compartilhada";

            elementoTipo
            .classList
            .add(
                "origem-compartilhada"
            );

        }

        else{

            elementoTipo.textContent =
                "Pilha individual";

            elementoTipo
            .classList
            .add(
                "origem-individual"
            );

        }

    }


    if(elementoCriterio){

        if(fontes.length === 0){

            elementoCriterio.textContent =
                `Não foram encontrados resultados AMG válidos associados à Pilha ${pilha}.`;

        }

        else if(
            meta.temRetomadaCompartilhada
        ){

            const listaCompartilhadas =
                compartilhadas.join(", ");


            elementoCriterio.textContent =
                `Os valores AMG da Pilha ${pilha} são a média simples dos resultados válidos associados a essa pilha. ` +
                `Foram incluídos registros de retomada compartilhada (${listaCompartilhadas}) porque essas identificações também contêm a Pilha ${pilha}.`;

        }

        else{

            elementoCriterio.textContent =
                `Os valores AMG da Pilha ${pilha} são a média simples dos resultados válidos identificados na base AMG para essa pilha.`;

        }

    }


    const mapaContagens = {

        amgQtdP205:
            amg?.p205?.quantidade ?? 0,

        amgQtdMgO:
            amg?.mgo?.quantidade ?? 0,

        amgQtdNb:
            amg?.nb?.quantidade ?? 0,

        amgQtdCaP:
            amg?.cap?.quantidade ?? 0,

        amgQtdFe:
            amg?.fe?.quantidade ?? 0,

        amgQtdSi:
            amg?.si?.quantidade ?? 0,

        amgQtdAl:
            amg?.al?.quantidade ?? 0,

        amgQtdTi:
            amg?.ti?.quantidade ?? 0

    };


    Object
    .entries(
        mapaContagens
    )
    .forEach(
        function([id,quantidade]){

            const elemento =
                document
                .getElementById(
                    id
                );


            if(elemento){

                elemento.textContent =
                    `${quantidade} registro${quantidade === 1 ? "" : "s"}`;

            }

        }
    );

}


/* ============================================================
   TABELA MPU × AMG
============================================================ */

function preencherTabelaComparacao(
    mpu,
    amg
){

    const tbody =
        document
        .getElementById(
            "corpoTabelaMpuAmg"
        );


    tbody.innerHTML =
        "";


    parametros
    .forEach(
        function(parametro){

            const valorMPU =
                mpu[
                    parametro.chave
                ];


            const valorAMG =
                amg[
                    parametro.chave
                ]
                .valor;


            const diferenca =
                calcularDiferenca(
                    valorAMG,
                    valorMPU
                );


            const variacao =
                calcularVariacao(
                    valorAMG,
                    valorMPU
                );


            const classe =
                classeDiferenca(
                    diferenca
                );


            const seta =
                setaDiferenca(
                    diferenca
                );


            const tr =
                document
                .createElement(
                    "tr"
                );


            tr.innerHTML = `

                <td>
                    <strong>
                        ${parametro.nome}
                    </strong>
                </td>

                <td>
                    ${formatarQuimica(valorMPU)}
                </td>

                <td>
                    ${formatarQuimica(valorAMG)}
                </td>

                <td class="${classe}">
                    ${formatarDiferenca(diferenca)}
                    ${seta}
                </td>

                <td class="${classe}">
                    ${formatarVariacao(variacao)}
                </td>

            `;


            tbody
            .appendChild(
                tr
            );

        }
    );

}


/* ============================================================
   RESUMO RÁPIDO
============================================================ */

function gerarResumoRapido(
    mpu,
    amg
){

    preencherResumoIndividual(

        "resumoComparacao1",

        "P2O5",

        mpu.p205,

        amg.p205.valor

    );


    preencherResumoIndividual(

        "resumoComparacao2",

        "MgO",

        mpu.mgo,

        amg.mgo.valor

    );


    preencherResumoIndividual(

        "resumoComparacao3",

        "CaO/P2O5",

        mpu.cap,

        amg.cap.valor

    );


    /* ============================
       CONTAMINANTES
    ============================ */

    const impurezas = [

        {
            nome:"Fe2O3",
            mpu:mpu.fe,
            amg:amg.fe.valor
        },

        {
            nome:"SiO2",
            mpu:mpu.si,
            amg:amg.si.valor
        },

        {
            nome:"Al2O3",
            mpu:mpu.al,
            amg:amg.al.valor
        },

        {
            nome:"TiO2",
            mpu:mpu.ti,
            amg:amg.ti.valor
        }

    ];


    const validas =
        impurezas.filter(
            item =>
                item.mpu !== null &&
                item.amg !== null
        );


    const bloco =
        document
        .getElementById(
            "resumoComparacao4"
        );


    const p =
        bloco
        .querySelector(
            "p"
        );


    if(
        validas.length === 0
    ){

        p.textContent =
            "Sem dados suficientes para comparação.";


        return;

    }


    const abaixo =
        validas.filter(
            item =>
                item.amg <
                item.mpu
        )
        .length;


    const acima =
        validas.filter(
            item =>
                item.amg >
                item.mpu
        )
        .length;


    if(
        abaixo ===
        validas.length
    ){

        p.textContent =
            "Todos os contaminantes analisados estão abaixo do MPU.";

    }

    else if(
        acima ===
        validas.length
    ){

        p.textContent =
            "Todos os contaminantes analisados estão acima do MPU.";

    }

    else{

        p.textContent =
            `${abaixo} contaminante(s) abaixo e ${acima} acima do MPU.`;

    }

}


/* ============================================================
   RESUMO INDIVIDUAL
============================================================ */

function preencherResumoIndividual(
    id,
    nome,
    valorMPU,
    valorAMG
){

    const bloco =
        document
        .getElementById(
            id
        );


    const titulo =
        bloco
        .querySelector(
            "strong"
        );


    const texto =
        bloco
        .querySelector(
            "p"
        );


    titulo.textContent =
        nome;


    if(
        valorMPU === null ||
        valorAMG === null
    ){

        texto.textContent =
            "Dados insuficientes para comparação.";


        return;

    }


    const diferenca =
        valorAMG -
        valorMPU;


    const variacao =
        valorMPU !== 0
        ?
        (
            diferenca /
            valorMPU
        ) *
        100
        :
        null;


    let direcao =
        "igual ao";


    if(
        diferenca > 0
    ){

        direcao =
            "acima do";

    }

    else if(
        diferenca < 0
    ){

        direcao =
            "abaixo do";

    }


    texto.textContent =

        `AMG está ${formatarQuimica(Math.abs(diferenca))} ` +
        `${direcao} MPU` +
        (
            variacao !== null
            ?
            ` (${formatarVariacao(variacao)}).`
            :
            "."
        );

}


/* ============================================================
   HISTÓRICO DAS 5 PILHAS
============================================================ */

function gerarHistorico5Pilhas(
    indiceAtual
){

    const linhas =
        obterCincoPilhas(
            indiceAtual
        );


    const tbody =
        document
        .getElementById(
            "corpoHistoricoAmg"
        );


    tbody.innerHTML =
        "";


    linhas
    .forEach(
        function(linha){

            const numero =
                linha[0];


            const mpu =
                obterDadosMPU(
                    linha
                );


            const amg =
                obterAMGPilha(
                    numero
                );


            const tr =
                document
                .createElement(
                    "tr"
                );


            if(
                String(numero) ===
                String(
                    pilhaSelecionadaAtual
                )
            ){

                tr.classList
                .add(
                    "pilha-atual-historico"
                );

            }


            tr.innerHTML = `

                <td>
                    <strong>
                        ${numero}
                    </strong>
                </td>


                ${criarCelulasHistorico(
                    mpu.p205,
                    amg.p205.valor
                )}


                ${criarCelulasHistorico(
                    mpu.mgo,
                    amg.mgo.valor
                )}


                ${criarCelulasHistorico(
                    mpu.fe,
                    amg.fe.valor
                )}


                ${criarCelulasHistorico(
                    mpu.si,
                    amg.si.valor
                )}

            `;


            tbody
            .appendChild(
                tr
            );

        }
    );

}


/* ============================================================
   CINCO PILHAS
============================================================ */

function obterCincoPilhas(
    indiceAtual
){

    const resultado =
        [];


    for(
        let i =
            indiceAtual;

        i >= 1 &&
        resultado.length < 5;

        i--
    ){

        const linha =
            dadosPilhas[i];


        if(
            linha &&
            linha[0] !== null &&
            linha[0] !== undefined
        ){

            resultado.push(
                linha
            );

        }

    }


    return resultado;

}


/* ============================================================
   CÉLULAS DO HISTÓRICO
============================================================ */

function criarCelulasHistorico(
    valorMPU,
    valorAMG
){

    const diferenca =
        calcularDiferenca(
            valorAMG,
            valorMPU
        );


    const classe =
        classeDiferenca(
            diferenca
        );


    return `

        <td>
            ${formatarQuimica(valorMPU)}
        </td>

        <td>
            ${formatarQuimica(valorAMG)}
        </td>

        <td class="${classe}">
            ${formatarDiferenca(diferenca)}
        </td>

    `;

}


/* ============================================================
   INTERPRETAÇÃO AUTOMÁTICA
============================================================ */

function gerarInterpretacaoAMG(
    mpu,
    amg
){

    const comparacoes =
        parametros
        .map(
            function(parametro){

                const valorMPU =
                    mpu[
                        parametro.chave
                    ];


                const valorAMG =
                    amg[
                        parametro.chave
                    ]
                    .valor;


                if(
                    valorMPU === null ||
                    valorAMG === null
                ){

                    return null;

                }


                const diferenca =
                    valorAMG -
                    valorMPU;


                const variacao =
                    valorMPU !== 0
                    ?
                    (
                        diferenca /
                        valorMPU
                    ) *
                    100
                    :
                    null;


                return {

                    nome:
                        parametro.nome,

                    mpu:
                        valorMPU,

                    amg:
                        valorAMG,

                    diferenca:
                        diferenca,

                    variacao:
                        variacao,

                    abs:
                        Math.abs(
                            variacao !== null
                            ?
                            variacao
                            :
                            diferenca
                        )

                };

            }
        )
        .filter(
            item =>
                item !== null
        );


    comparacoes.sort(
        function(a,b){

            return (
                b.abs -
                a.abs
            );

        }
    );


    /* ============================
       TEXTO 1 - P2O5
    ============================ */

    const p205 =
        comparacoes.find(
            item =>
                item.nome ===
                "P2O5"
        );


    if(p205){

        const direcao =
            p205.diferenca > 0
            ?
            "acima"
            :
            (
                p205.diferenca < 0
                ?
                "abaixo"
                :
                "igual"
            );


        definirInterpretacao(

            "interpretacaoAMG1",

            `O P2O5 do AMG está ${direcao} do MPU em ` +
            `${formatarQuimica(Math.abs(p205.diferenca))} p.p.`

        );

    }


    /* ============================
       TEXTO 2 - MAIOR VARIAÇÃO
    ============================ */

    const maior =
        comparacoes[0];


    if(maior){

        const direcao =
            maior.diferenca > 0
            ?
            "aumento"
            :
            (
                maior.diferenca < 0
                ?
                "redução"
                :
                "estabilidade"
            );


        definirInterpretacao(

            "interpretacaoAMG2",

            `${maior.nome} apresentou a maior variação entre MPU e AMG, ` +
            `com ${direcao} de ` +
            `${formatarQuimica(Math.abs(maior.diferenca))}.`

        );

    }


    /* ============================
       TEXTO 3 - GERAL
    ============================ */

    if(
        comparacoes.length > 0
    ){

        const acima =
            comparacoes
            .filter(
                item =>
                    item.diferenca >
                    0
            )
            .length;


        const abaixo =
            comparacoes
            .filter(
                item =>
                    item.diferenca <
                    0
            )
            .length;


        definirInterpretacao(

            "interpretacaoAMG3",

            `Na comparação geral, ${abaixo} parâmetro(s) apresentam valor ` +
            `AMG abaixo do MPU e ${acima} acima.`

        );

    }

}


/* ============================================================
   DEFINIR INTERPRETAÇÃO
============================================================ */

function definirInterpretacao(
    id,
    texto
){

    const bloco =
        document
        .getElementById(
            id
        );


    bloco
    .querySelector(
        "p"
    )
    .textContent =
        texto;

}


/* ============================================================
   GRÁFICO DAS PILHAS
============================================================ */

function gerarGraficoPilhas(){

    if(
        typeof Chart === "undefined"
    ){

        alert(
            "A biblioteca dos gráficos não foi carregada. Verifique a conexão com a internet e atualize a página."
        );

        return;

    }

    const labels =
        [];


    const massas =
        [];


    const cores =
        [];


    for(
        let i = 1;
        i < dadosPilhas.length;
        i++
    ){

        const linha =
            dadosPilhas[i];


        if(!linha){

            continue;

        }


        const pilha =
            linha[0];


        const massa =
            valorNumero(
                linha[18]
            );


        if(
            pilha === null ||
            pilha === undefined ||
            massa === null
        ){

            continue;

        }


        labels.push(
            String(
                pilha
            )
        );


        massas.push(
            massa
        );


        if(
            pilhaSelecionadaAtual !== null &&
            String(pilha) ===
            String(
                pilhaSelecionadaAtual
            )
        ){

            cores.push(
                "#92D050"
            );

        }

        else{

            cores.push(
                "#3B005F"
            );

        }

    }


    desenharGrafico(
        labels,
        massas,
        cores
    );

}


/* ============================================================
   DESENHAR GRÁFICO
============================================================ */

function desenharGrafico(
    labels,
    massas,
    cores
){

    const canvas =
        document
        .getElementById(
            "graficoPilhas"
        );


    if(
        graficoPilhas
    ){

        graficoPilhas
        .destroy();

    }


    graficoPilhas =
        new Chart(
            canvas,
            {

                type:
                    "bar",


                data:{

                    labels:
                        labels,


                    datasets:[
                        {

                            label:
                                "Massa da Pilha (t)",

                            data:
                                massas,

                            backgroundColor:
                                cores,

                            borderRadius:
                                4,

                            barPercentage:
                                0.72,

                            categoryPercentage:
                                0.82

                        }
                    ]

                },


                options:{

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    devicePixelRatio:
                        Math.max(
                            window.devicePixelRatio || 1,
                            2
                        ),

                    layout:{
                        padding:{
                            top:8,
                            right:14,
                            bottom:2,
                            left:4
                        }
                    },


                    interaction:{

                        mode:
                            "index",

                        intersect:
                            false

                    },


                    plugins:{

                        legend:{

                            position:
                                "top",

                            labels:{

                                color:
                                    "#2D014D",

                                usePointStyle:
                                    true,

                                font:{

                                    size:11,

                                    weight:
                                        "bold"

                                },

                                boxWidth:10,
                                boxHeight:10,
                                padding:16

                            }

                        },


                        tooltip:{

                            callbacks:{

                                title:
                                    function(context){

                                        return (
                                            "Pilha " +
                                            context[0]
                                            .label
                                        );

                                    },


                                label:
                                    function(context){

                                        return (
                                            "Massa: " +
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

                            title:{

                                display:
                                    true,

                                text:
                                    "Número da Pilha",

                                color:
                                    "#666"

                            },


                            grid:{

                                display:
                                    false

                            },


                            ticks:{

                                color:
                                    "#555",

                                autoSkip:
                                    false,

                                maxRotation:
                                    0,

                                minRotation:
                                    0,

                                font:{

                                    size:
                                        10

                                }

                            }

                        },


                        y:{

                            beginAtZero:
                                true,


                            title:{

                                display:
                                    true,

                                text:
                                    "Massa (t)",

                                color:
                                    "#666"

                            },


                            grid:{

                                color:
                                    "rgba(0,0,0,0.06)"

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
   CÁLCULOS
============================================================ */

function calcularMedia(
    valores
){

    const validos =
        valores
        .filter(
            valor =>
                valor !== null &&
                Number.isFinite(
                    Number(valor)
                )
        );


    if(
        validos.length === 0
    ){

        return null;

    }


    const soma =
        validos
        .reduce(
            function(total,valor){

                return (
                    total +
                    Number(valor)
                );

            },
            0
        );


    return (
        soma /
        validos.length
    );

}


/* ============================================================
   DIFERENÇA
============================================================ */

function calcularDiferenca(
    amg,
    mpu
){

    if(
        amg === null ||
        mpu === null
    ){

        return null;

    }


    return (
        amg -
        mpu
    );

}


/* ============================================================
   VARIAÇÃO
============================================================ */

function calcularVariacao(
    amg,
    mpu
){

    if(
        amg === null ||
        mpu === null ||
        mpu === 0
    ){

        return null;

    }


    return (
        (
            amg -
            mpu
        ) /
        mpu
    ) *
    100;

}


/* ============================================================
   CLASSE DA DIFERENÇA
============================================================ */

function classeDiferenca(
    valor
){

    if(
        valor === null
    ){

        return "";

    }


    if(
        valor > 0
    ){

        return "diferenca-positiva";

    }


    if(
        valor < 0
    ){

        return "diferenca-negativa";

    }


    return "diferenca-neutra";

}


/* ============================================================
   SETA
============================================================ */

function setaDiferenca(
    valor
){

    if(
        valor === null
    ){

        return "";

    }


    if(
        valor > 0
    ){

        return " ↑";

    }


    if(
        valor < 0
    ){

        return " ↓";

    }


    return "";

}


/* ============================================================
   FORMATAR DIFERENÇA
============================================================ */

function formatarDiferenca(
    valor
){

    if(
        valor === null
    ){

        return "-";

    }


    const sinal =
        valor > 0
        ?
        "+"
        :
        "";


    return (
        sinal +
        Number(valor)
        .toFixed(2)
        .replace(
            ".",
            ","
        )
    );

}


/* ============================================================
   FORMATAR VARIAÇÃO
============================================================ */

function formatarVariacao(
    valor
){

    if(
        valor === null ||
        valor === undefined
    ){

        return "-";

    }


    const sinal =
        valor > 0
        ?
        "+"
        :
        "";


    return (
        sinal +
        Number(valor)
        .toFixed(2)
        .replace(
            ".",
            ","
        ) +
        "%"
    );

}


/* ============================================================
   DATAS
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
        valor instanceof Date
    ){

        return valor;

    }


    if(
        typeof valor ===
        "number"
    ){

        const data =
            XLSX.SSF
            .parse_date_code(
                valor
            );


        if(!data){

            return null;

        }


        return new Date(

            data.y,

            data.m - 1,

            data.d,

            data.H || 0,

            data.M || 0,

            data.S || 0

        );

    }


    const data =
        new Date(
            valor
        );


    if(
        isNaN(
            data.getTime()
        )
    ){

        return null;

    }


    return data;

}


/* ============================================================
   FORMATAÇÕES DE DATA
============================================================ */

function formatarDataCurta(
    data
){

    if(!data){

        return "-";

    }


    return data
    .toLocaleDateString(
        "pt-BR"
    );

}


function formatarDataHora(
    data
){

    if(!data){

        return "-";

    }


    return data
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
    );

}


/* ============================================================
   DURAÇÃO
============================================================ */

function calcularDuracao(
    inicio,
    fim
){

    if(
        !inicio ||
        !fim
    ){

        return "-";

    }


    const diferenca =
        fim.getTime() -
        inicio.getTime();


    if(
        diferenca <
        0
    ){

        return "-";

    }


    const horas =
        diferenca /
        (
            1000 *
            60 *
            60
        );


    const dias =
        Math.ceil(
            horas /
            24
        );


    if(
        dias <= 1
    ){

        return (
            horas
            .toFixed(0) +
            " horas"
        );

    }


    return (
        dias +
        " dias"
    );

}


/* ============================================================
   NÚMERO
============================================================ */

function valorNumero(
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


    const texto =
        String(valor)
        .trim()
        .replace(
            ",",
            "."
        );


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
   NÚMERO FLEXÍVEL
============================================================ */

function valorNumeroFlexivel(
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
        String(valor)
        .trim();


    /*
        Exemplo:
        12,35
    */

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


    /*
        Exemplo:
        1.234,56
    */

    else if(
        texto.includes(",") &&
        texto.includes(".")
    ){

        texto =
            texto
            .replace(
                /\./g,
                ""
            )
            .replace(
                ",",
                "."
            );

    }


    texto =
        texto.replace(
            /%/g,
            ""
        );


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
   TEXTO
============================================================ */

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
   NORMALIZAR
============================================================ */

function normalizarTexto(
    valor
){

    return limparTexto(
        valor
    )
    .toLowerCase()
    .normalize(
        "NFD"
    )
    .replace(
        /[\u0300-\u036f]/g,
        ""
    )
    .replace(
        /\s+/g,
        " "
    );

}


/* ============================================================
   TONELADAS
============================================================ */

function formatarToneladas(
    valor
){

    if(
        valor === null ||
        valor === undefined
    ){

        return "-";

    }


    return Number(
        valor
    )
    .toLocaleString(
        "pt-BR",
        {
            maximumFractionDigits:
                0
        }
    );

}


/* ============================================================
   QUÍMICA
============================================================ */

function formatarQuimica(
    valor
){

    if(
        valor === null ||
        valor === undefined
    ){

        return "-";

    }


    return Number(
        valor
    )
    .toFixed(
        2
    )
    .replace(
        ".",
        ","
    );

}


/* ============================================================
   ATUALIZAÇÃO
============================================================ */

function atualizarDataSistema(){

    const agora =
        new Date();


    document
    .getElementById(
        "ultimaAtualizacaoPilha"
    )
    .textContent =
        formatarDataHora(
            agora
        );

}
/* ============================================================
   QUALIDADE DO GRAFICO NA IMPRESSAO
============================================================ */

let configuracaoGraficoTelaPilhas = null;

function prepararGraficoImpressaoPilhas(){

    if(!graficoPilhas){
        return;
    }

    configuracaoGraficoTelaPilhas = {
        devicePixelRatio: graficoPilhas.options.devicePixelRatio,
        fonteX: graficoPilhas.options.scales?.x?.ticks?.font?.size,
        fonteY: graficoPilhas.options.scales?.y?.ticks?.font?.size,
        fonteLegenda: graficoPilhas.options.plugins?.legend?.labels?.font?.size
    };

    graficoPilhas.options.animation = false;
    graficoPilhas.options.devicePixelRatio = 4;

    if(graficoPilhas.options.scales?.x?.ticks){
        graficoPilhas.options.scales.x.ticks.font = {
            ...graficoPilhas.options.scales.x.ticks.font,
            size: 12,
            weight: "600"
        };
    }

    if(graficoPilhas.options.scales?.y?.ticks){
        graficoPilhas.options.scales.y.ticks.font = {
            ...graficoPilhas.options.scales.y.ticks.font,
            size: 11
        };
    }

    if(graficoPilhas.options.plugins?.legend?.labels){
        graficoPilhas.options.plugins.legend.labels.font = {
            ...graficoPilhas.options.plugins.legend.labels.font,
            size: 12,
            weight: "700"
        };
    }

    graficoPilhas.resize();
    graficoPilhas.update("none");
}

function restaurarGraficoTelaPilhas(){

    if(!graficoPilhas || !configuracaoGraficoTelaPilhas){
        return;
    }

    graficoPilhas.options.devicePixelRatio =
        configuracaoGraficoTelaPilhas.devicePixelRatio ||
        Math.max(window.devicePixelRatio || 1, 2);

    if(graficoPilhas.options.scales?.x?.ticks){
        graficoPilhas.options.scales.x.ticks.font = {
            ...graficoPilhas.options.scales.x.ticks.font,
            size: configuracaoGraficoTelaPilhas.fonteX || 10
        };
    }

    if(graficoPilhas.options.scales?.y?.ticks){
        graficoPilhas.options.scales.y.ticks.font = {
            ...graficoPilhas.options.scales.y.ticks.font,
            size: configuracaoGraficoTelaPilhas.fonteY || 10
        };
    }

    if(graficoPilhas.options.plugins?.legend?.labels){
        graficoPilhas.options.plugins.legend.labels.font = {
            ...graficoPilhas.options.plugins.legend.labels.font,
            size: configuracaoGraficoTelaPilhas.fonteLegenda || 11
        };
    }

    graficoPilhas.resize();
    graficoPilhas.update("none");
    configuracaoGraficoTelaPilhas = null;
}

window.addEventListener("afterprint", function(){
    document.body.classList.remove("modo-impressao-pilha");
    restaurarGraficoTelaPilhas();
});


/* ============================================================
   PDF E ENVIO POR E-MAIL
============================================================ */


/* ============================================================
   DESTINATÁRIOS
============================================================ */

const emailsRelatorioPilha = [

    "luiz.campos@br.cmoc.com",

    "thais.gilvana@br.cmoc.com"

];


/* ============================================================
   GERAR PDF
============================================================ */

function gerarPDFRelatorioPilha(){

    if(
        pilhaSelecionadaAtual === null
    ){

        alert(
            "Consulte uma pilha antes de gerar o PDF."
        );

        return;

    }

    const tituloAnterior =
        document.title;

    document.body
    .classList
    .add(
        "modo-impressao-pilha"
    );

    document.title =
        "Relatorio_Pilha_" +
        pilhaSelecionadaAtual +
        "_FVO";

    prepararGraficoImpressaoPilhas();

    setTimeout(
        function(){

            window.print();

            setTimeout(
                function(){

                    document.title =
                        tituloAnterior;

                    document.body
                    .classList
                    .remove(
                        "modo-impressao-pilha"
                    );

                    restaurarGraficoTelaPilhas();

                },
                250
            );

        },
        500
    );

}


/* ============================================================
   ENVIAR POR E-MAIL
============================================================ */

function enviarRelatorioPorEmail(){

    if(
        pilhaSelecionadaAtual === null
    ){

        alert(
            "Consulte uma pilha antes de enviar o relatório."
        );

        return;

    }


    const numero =
        pilhaSelecionadaAtual;


    /*
        PEGAR DADOS QUE ESTÃO NA TELA
    */

    const massa =
        obterTextoElemento(
            "massaSelecionada"
        );


    const periodo =
        obterTextoElemento(
            "periodoFormacao"
        );


    const p205MPU =
        obterTextoElemento(
            "mpuP205"
        );


    const p205AMG =
        obterTextoElemento(
            "amgP205"
        );


    const mgoMPU =
        obterTextoElemento(
            "mpuMgO"
        );


    const mgoAMG =
        obterTextoElemento(
            "amgMgO"
        );


    const nbMPU =
        obterTextoElemento(
            "mpuNb"
        );


    const nbAMG =
        obterTextoElemento(
            "amgNb"
        );


    const capMPU =
        obterTextoElemento(
            "mpuCaP"
        );


    const capAMG =
        obterTextoElemento(
            "amgCaP"
        );


    /*
        ASSUNTO
    */

    const assunto =
        "Relatório de Análise da Pilha " +
        numero +
        " - FVO";


    /*
        CORPO DO E-MAIL
    */

    const mensagem =

`Prezados,

Segue o Relatório de Análise da Pilha ${numero} - FVO.

Resumo:

Pilha: ${numero}
Massa: ${massa}
Período de formação: ${periodo}

MPU - Britagem
P2O5: ${p205MPU}
MgO: ${mgoMPU}
Nb2O5: ${nbMPU}
CaO/P2O5: ${capMPU}

AMG - Usina
P2O5: ${p205AMG}
MgO: ${mgoAMG}
Nb2O5: ${nbAMG}
CaO/P2O5: ${capAMG}

O relatório completo pode ser gerado pelo Portal FVO.

Atenciosamente.`;


    /*
        MONTA DESTINATÁRIOS

        O ; costuma funcionar melhor
        no Outlook Desktop.
    */

    const destinatarios =
        emailsRelatorioPilha
        .join(
            ";"
        );


    /*
        MONTA MAILTO
    */

    const linkEmail =

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


    /*
        ABRIR CLIENTE DE E-MAIL
    */

    window.location.href =
        linkEmail;

}


/* ============================================================
   PEGAR TEXTO DE UM ELEMENTO
============================================================ */

function obterTextoElemento(
    id
){

    const elemento =
        document
        .getElementById(
            id
        );


    if(!elemento){

        return "-";

    }


    const texto =
        elemento
        .textContent
        .trim();


    return texto || "-";

}