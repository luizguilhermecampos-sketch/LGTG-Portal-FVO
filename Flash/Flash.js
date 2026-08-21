/* ============================================================
   FLASH OPERACIONAL - BRITAGEM FVO
============================================================ */


/* ============================================================
   VARIÁVEIS
============================================================ */

let workbookFlash = null;

let flashGerado = false;

let quantidadeImpactos = 0;

let dadosPilhaAtual = null;


/* ============================================================
   EMAILS
============================================================ */

const emailsFlash = [

    "luiz.campos@br.cmoc.com",

    "thais.gilvana@br.cmoc.com"

];


/* ============================================================
   HORÁRIOS
============================================================ */

const horariosFlashPadrao = [

    "06:00",

    "12:00",

    "18:00",

    "23:00"

];


/* ============================================================
   ELEMENTOS
============================================================ */

const arquivoControleEstocagem =
    document.getElementById(
        "arquivoControleEstocagem"
    );


const pilhaFlash =
    document.getElementById(
        "pilhaFlash"
    );


/* ============================================================
   EVENTOS
============================================================ */

arquivoControleEstocagem
.addEventListener(
    "change",
    carregarControleEstocagem
);


pilhaFlash
.addEventListener(
    "change",
    carregarDadosPilhaSelecionada
);


document
.getElementById(
    "btnAdicionarImpacto"
)
.addEventListener(
    "click",
    adicionarImpacto
);


document
.getElementById(
    "btnGerarFlash"
)
.addEventListener(
    "click",
    gerarFlash
);


document
.getElementById(
    "btnGerarFlashFinal"
)
.addEventListener(
    "click",
    gerarFlash
);


document
.getElementById(
    "btnGerarPDFFlash"
)
.addEventListener(
    "click",
    gerarPDFFlash
);


document
.getElementById(
    "btnEnviarEmailFlash"
)
.addEventListener(
    "click",
    enviarEmailFlash
);


document
.getElementById(
    "btnLimparFlash"
)
.addEventListener(
    "click",
    limparFlash
);


/* ============================================================
   MASSA / CÁLCULOS
============================================================ */

const camposCalculoFlash = [

    "massaAcumuladaManualFlash",

    "massaPlanejadaFlash",

    "metaHorariaFlash",

    "mediaHorariaFlash",

    "manutencaoPlanejadaFlash",

    "dataFlash",

    "horarioFlash",

    "producaoFlash"

];


camposCalculoFlash.forEach(
    function(id){

        const elemento =
            document.getElementById(id);


        if(!elemento){

            return;

        }


        elemento.addEventListener(
            "input",
            function(){

                calcularIndicadoresFlash();

                atualizarPreviewSeGerado();

            }
        );


        elemento.addEventListener(
            "change",
            function(){

                calcularIndicadoresFlash();

                atualizarPreviewSeGerado();

            }
        );

    }
);


/* ============================================================
   CAMPOS DA PRÉVIA
============================================================ */

const camposPreview = [

    "dataFlash",

    "horarioFlash",

    "turnoFlash",

    "turmaFlash",

    "operadorFlash",

    "coordenadorFlash",

    "producaoFlash",

    "massaAcumuladaManualFlash",

    "massaPlanejadaFlash",

    "metaHorariaFlash",

    "mediaHorariaFlash",

    "manutencaoPlanejadaFlash",

    "frentesManualFlash",

    "observacoesFlash"

];


camposPreview.forEach(
    function(id){

        const elemento =
            document.getElementById(id);


        if(!elemento){

            return;

        }


        elemento.addEventListener(
            "input",
            atualizarPreviewSeGerado
        );


        elemento.addEventListener(
            "change",
            atualizarPreviewSeGerado
        );

    }
);


/* ============================================================
   FRENTES MANUAIS
============================================================ */

document
.getElementById(
    "frentesManualFlash"
)
.addEventListener(
    "input",
    function(){

        atualizarTagsFrentes();

        atualizarPreviewSeGerado();

    }
);


/* ============================================================
   CARREGAR EXCEL
============================================================ */

async function carregarControleEstocagem(
    event
){

    const arquivo =
        event.target.files[0];


    if(!arquivo){

        return;

    }


    if(typeof XLSX === "undefined"){

        alert(
            "A biblioteca de leitura do Excel não foi carregada. Verifique a conexão com a internet e atualize a página."
        );

        return;

    }


    if(window.FVOCarregamento){

        window.FVOCarregamento.exibir(
            "Lendo arquivo..."
        );

    }


    atualizarStatusArquivo(
        "Carregando...",
        "processando"
    );


    try{

        const buffer =
            await arquivo.arrayBuffer();


        workbookFlash =
            XLSX.read(
                buffer,
                {

                    type:
                        "array",

                    cellDates:
                        true,

                    cellFormula:
                        true

                }
            );


        const pilhas =
            identificarAbasPilhas(
                workbookFlash.SheetNames
            );


        montarListaPilhas(
            pilhas
        );


        document
        .getElementById(
            "statusArquivoFlash"
        )
        .textContent =
            arquivo.name;


        atualizarStatusArquivo(

            pilhas.length +

            " pilha(s) encontrada(s)",

            "ok"

        );


        if(
            pilhas.length === 1
        ){

            pilhaFlash.value =
                pilhas[0];


            carregarDadosPilhaSelecionada();

        }

    }catch(erro){

        console.error(
            erro
        );


        workbookFlash =
            null;


        atualizarStatusArquivo(
            "Erro ao ler arquivo",
            "erro"
        );


        alert(
            "Não foi possível ler a planilha."
        );

    }finally{

        if(window.FVOCarregamento){

            window.FVOCarregamento.ocultar();

        }

    }

}


/* ============================================================
   IDENTIFICAR ABAS
============================================================ */

function identificarAbasPilhas(
    nomesAbas
){

    const regex =
        /(?:PILHA\s*)?(\d{1,3})\s*[-\/]?\s*([AB])$/i;


    return nomesAbas
        .filter(
            function(nome){

                return regex.test(
                    String(nome)
                    .trim()
                );

            }
        )
        .sort(
            ordenarPilhas
        );

}


/* ============================================================
   ORDENAR PILHAS
============================================================ */

function ordenarPilhas(
    a,
    b
){

    const extrair =
        function(valor){

            const match =
                String(valor)
                .toUpperCase()
                .match(
                    /(\d{1,3}).*?([AB])$/
                );


            if(!match){

                return {

                    numero:
                        9999,

                    lado:
                        "Z"

                };

            }


            return {

                numero:
                    Number(
                        match[1]
                    ),

                lado:
                    match[2]

            };

        };


    const pa =
        extrair(a);


    const pb =
        extrair(b);


    if(
        pa.numero !==
        pb.numero
    ){

        return (
            pa.numero -
            pb.numero
        );

    }


    return pa.lado.localeCompare(
        pb.lado
    );

}


/* ============================================================
   LISTA DE PILHAS
============================================================ */

function montarListaPilhas(
    pilhas
){

    pilhaFlash.innerHTML =
        "";


    const inicial =
        document.createElement(
            "option"
        );


    inicial.value =
        "";


    inicial.textContent =
        pilhas.length
        ?
        "Selecione a pilha"
        :
        "Nenhuma pilha encontrada";


    pilhaFlash.appendChild(
        inicial
    );


    pilhas.forEach(
        function(nome){

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                nome;


            option.textContent =
                normalizarNomePilha(
                    nome
                );


            pilhaFlash.appendChild(
                option
            );

        }
    );

}


/* ============================================================
   PILHA SELECIONADA
============================================================ */

function carregarDadosPilhaSelecionada(){

    if(
        !workbookFlash
    ){

        return;

    }


    const nomeAba =
        pilhaFlash.value;


    if(!nomeAba){

        limparDadosAutomaticosPilha();

        return;

    }


    const worksheet =
        workbookFlash.Sheets[
            nomeAba
        ];


    if(!worksheet){

        return;

    }


    try{

        dadosPilhaAtual =
            extrairDadosDaPilha(

                worksheet,

                nomeAba

            );


        preencherDadosAutomaticosPilha(
            dadosPilhaAtual
        );


        /*
            SE CONSEGUIR ENCONTRAR FRENTES,
            PREENCHE O CAMPO MANUAL.
        */

        if(
            dadosPilhaAtual.frentes &&
            dadosPilhaAtual.frentes.length > 0
        ){

            document
            .getElementById(
                "frentesManualFlash"
            )
            .value =
                dadosPilhaAtual.frentes
                .join(
                    ", "
                );

        }


        atualizarTagsFrentes();

        calcularIndicadoresFlash();

        atualizarPreviewSeGerado();

    }catch(erro){

        console.error(
            erro
        );


        alert(
            "A aba foi encontrada, mas ocorreu erro na leitura."
        );

    }

}


/* ============================================================
   EXTRAÇÃO DA PILHA
============================================================ */

function extrairDadosDaPilha(
    worksheet,
    nomeAba
){

    const matriz =
        XLSX.utils.sheet_to_json(
            worksheet,
            {

                header:
                    1,

                defval:
                    null,

                raw:
                    true

            }
        );


    const mapa =
        identificarCabecalhos(
            matriz
        );


    const fechamento =
        localizarLinhaFechamento(
            matriz
        );


    const linhaDados =
        fechamento !== -1
        ?
        fechamento
        :
        localizarUltimaLinhaComDados(
            matriz,
            mapa
        );


    const p2o5 =
        obterValorPorCabecalho(

            matriz,

            linhaDados,

            mapa,

            [
                "P2O5"
            ]

        );


    const mgo =
        obterValorPorCabecalho(

            matriz,

            linhaDados,

            mapa,

            [
                "MGO"
            ]

        );


    const nb2o5 =
        obterValorPorCabecalho(

            matriz,

            linhaDados,

            mapa,

            [
                "NB2O5"
            ]

        );


    const inicio =
        localizarPrimeiraDataPilha(
            matriz
        );


    const frentes =
        extrairFrentes(
            matriz,
            mapa
        );


    return {

        nomeAba:
            nomeAba,

        nomePilha:
            normalizarNomePilha(
                nomeAba
            ),

        p2o5:
            normalizarNumeroExcel(
                p2o5
            ),

        mgo:
            normalizarNumeroExcel(
                mgo
            ),

        nb2o5:
            normalizarNumeroExcel(
                nb2o5
            ),

        inicio:
            inicio,

        frentes:
            frentes

    };

}


/* ============================================================
   CABEÇALHOS
============================================================ */

function identificarCabecalhos(
    matriz
){

    const mapa =
        {};


    const limite =
        Math.min(
            matriz.length,
            35
        );


    for(
        let linha = 0;
        linha < limite;
        linha++
    ){

        const dados =
            matriz[linha] || [];


        dados.forEach(
            function(valor,coluna){

                if(
                    valor === null ||
                    valor === undefined
                ){

                    return;

                }


                const texto =
                    normalizarTexto(
                        valor
                    );


                if(
                    !texto ||
                    texto.length > 80
                ){

                    return;

                }


                if(
                    mapa[texto] ===
                    undefined
                ){

                    mapa[texto] = {

                        coluna:
                            coluna,

                        linha:
                            linha

                    };

                }

            }
        );

    }


    return mapa;

}


/* ============================================================
   FECHAMENTO
============================================================ */

function localizarLinhaFechamento(
    matriz
){

    for(
        let linha = 0;
        linha < matriz.length;
        linha++
    ){

        const encontrou =
            (
                matriz[linha] ||
                []
            )
            .some(
                function(valor){

                    return normalizarTexto(
                        valor
                    )
                    .includes(
                        "FECHAMENTO"
                    );

                }
            );


        if(encontrou){

            return linha;

        }

    }


    return -1;

}


/* ============================================================
   ÚLTIMA LINHA QUÍMICA
============================================================ */

function localizarUltimaLinhaComDados(
    matriz,
    mapa
){

    const colunas =
        obterColunasPorPossiveisCabecalhos(

            mapa,

            [
                "P2O5",
                "MGO",
                "NB2O5"
            ]

        );


    for(
        let linha =
            matriz.length - 1;
        linha >= 0;
        linha--
    ){

        const dados =
            matriz[linha] || [];


        const possui =
            colunas.some(
                function(coluna){

                    return ehNumero(
                        dados[coluna]
                    );

                }
            );


        if(possui){

            return linha;

        }

    }


    return (
        matriz.length -
        1
    );

}


/* ============================================================
   OBTER VALOR
============================================================ */

function obterValorPorCabecalho(
    matriz,
    linhaDados,
    mapa,
    nomes
){

    const coluna =
        localizarColunaCabecalho(
            mapa,
            nomes
        );


    if(
        coluna === -1
    ){

        return null;

    }


    const direto =
        matriz[linhaDados]
        ?
        matriz[linhaDados][coluna]
        :
        null;


    if(
        direto !== null &&
        direto !== undefined &&
        direto !== ""
    ){

        return direto;

    }


    for(
        let linha =
            linhaDados - 1;
        linha >= 0;
        linha--
    ){

        const candidato =
            matriz[linha]
            ?
            matriz[linha][coluna]
            :
            null;


        if(
            ehNumero(
                candidato
            )
        ){

            return candidato;

        }

    }


    return null;

}


/* ============================================================
   LOCALIZAR COLUNA
============================================================ */

function localizarColunaCabecalho(
    mapa,
    nomes
){

    const chaves =
        Object.keys(
            mapa
        );


    for(
        const nome of nomes
    ){

        const alvo =
            normalizarTexto(
                nome
            );


        if(
            mapa[alvo] !==
            undefined
        ){

            return mapa[
                alvo
            ].coluna;

        }


        const encontrada =
            chaves.find(
                function(chave){

                    return chave.includes(
                        alvo
                    );

                }
            );


        if(encontrada){

            return mapa[
                encontrada
            ].coluna;

        }

    }


    return -1;

}


/* ============================================================
   COLUNAS
============================================================ */

function obterColunasPorPossiveisCabecalhos(
    mapa,
    nomes
){

    const resultado =
        [];


    nomes.forEach(
        function(nome){

            const coluna =
                localizarColunaCabecalho(
                    mapa,
                    [nome]
                );


            if(
                coluna !== -1 &&
                !resultado.includes(
                    coluna
                )
            ){

                resultado.push(
                    coluna
                );

            }

        }
    );


    return resultado;

}


/* ============================================================
   TENTAR EXTRAIR FRENTES
============================================================ */

function extrairFrentes(
    matriz,
    mapa
){

    const colunas =
        [];


    Object.keys(
        mapa
    )
    .forEach(
        function(chave){

            if(
                chave.includes(
                    "NIVEL"
                )
                ||
                chave.includes(
                    "FRENTE"
                )
                ||
                chave.includes(
                    "ORIGEM"
                )
            ){

                const coluna =
                    mapa[chave].coluna;


                if(
                    !colunas.includes(
                        coluna
                    )
                ){

                    colunas.push(
                        coluna
                    );

                }

            }

        }
    );


    const resultado =
        new Set();


    colunas.forEach(
        function(coluna){

            matriz.forEach(
                function(linha){

                    const valor =
                        linha
                        ?
                        linha[coluna]
                        :
                        null;


                    if(
                        valor === null ||
                        valor === undefined
                    ){

                        return;

                    }


                    const texto =
                        String(valor)
                        .trim();


                    if(
                        pareceFrenteOuOrigem(
                            texto
                        )
                    ){

                        resultado.add(
                            texto
                        );

                    }

                }
            );

        }
    );


    return Array.from(
        resultado
    )
    .sort(
        ordenarFrentes
    );

}


/* ============================================================
   FRENTE VÁLIDA
============================================================ */

function pareceFrenteOuOrigem(
    texto
){

    if(
        /^M\d+/i.test(
            texto
        )
    ){

        return true;

    }


    const permitido = [

        "BAIA",

        "PULMAO",

        "PULMÃO",

        "ROM",

        "ESTOQUE",

        "PATIO",

        "PÁTIO"

    ];


    const normalizado =
        normalizarTexto(
            texto
        );


    return permitido.some(
        function(item){

            return normalizado.includes(
                normalizarTexto(
                    item
                )
            );

        }
    );

}


/* ============================================================
   ORDENAR FRENTES
============================================================ */

function ordenarFrentes(
    a,
    b
){

    const ma =
        String(a)
        .match(
            /M(\d+)/i
        );


    const mb =
        String(b)
        .match(
            /M(\d+)/i
        );


    if(
        ma &&
        mb
    ){

        return (
            Number(ma[1]) -
            Number(mb[1])
        );

    }


    if(ma){

        return -1;

    }


    if(mb){

        return 1;

    }


    return String(a)
        .localeCompare(
            String(b),
            "pt-BR"
        );

}


/* ============================================================
   INÍCIO DA PILHA
============================================================ */

function localizarPrimeiraDataPilha(
    matriz
){

    for(
        let linha = 0;
        linha < matriz.length;
        linha++
    ){

        for(
            let coluna = 0;
            coluna <
            matriz[linha].length;
            coluna++
        ){

            const data =
                converterParaData(
                    matriz[linha][coluna]
                );


            if(data){

                return data;

            }

        }

    }


    return null;

}


/* ============================================================
   PREENCHER AUTOMÁTICOS
============================================================ */

function preencherDadosAutomaticosPilha(
    dados
){

    setTexto(

        "p2o5Flash",

        dados.p2o5 !== null
        ?
        formatarNumero(
            dados.p2o5,
            2
        )
        :
        "-"

    );


    setTexto(

        "mgoFlash",

        dados.mgo !== null
        ?
        formatarNumero(
            dados.mgo,
            2
        )
        :
        "-"

    );


    setTexto(

        "nb2o5Flash",

        dados.nb2o5 !== null
        ?
        formatarNumero(
            dados.nb2o5,
            2
        )
        :
        "-"

    );


    document
    .getElementById(
        "inicioPilhaFlash"
    )
    .value =
        dados.inicio
        ?
        formatarDataHora(
            dados.inicio
        )
        :
        "-";

}


/* ============================================================
   FRENTES DO CAMPO MANUAL
============================================================ */

function obterFrentesFlash(){

    const texto =
        valorCampo(
            "frentesManualFlash"
        );


    if(!texto){

        return [];

    }


    return Array.from(
        new Set(

            texto
            .split(
                /[,;|]/g
            )
            .map(
                function(item){

                    return item.trim();

                }
            )
            .filter(
                function(item){

                    return item.length > 0;

                }
            )

        )
    );

}


/* ============================================================
   TAGS DE FRENTES
============================================================ */

function atualizarTagsFrentes(){

    const container =
        document.getElementById(
            "listaFrentesFlash"
        );


    const frentes =
        obterFrentesFlash();


    container.innerHTML =
        "";


    if(
        frentes.length === 0
    ){

        container.innerHTML = `

            <span class="tag-frente-vazia">

                Nenhuma frente informada

            </span>

        `;


        return;

    }


    frentes.forEach(
        function(frente){

            const tag =
                document.createElement(
                    "span"
                );


            tag.className =
                "tag-frente-flash";


            tag.textContent =
                frente;


            container.appendChild(
                tag
            );

        }
    );

}


/* ============================================================
   LIMPAR DADOS DE PILHA
============================================================ */

function limparDadosAutomaticosPilha(){

    dadosPilhaAtual =
        null;


    setTexto(
        "p2o5Flash",
        "-"
    );


    setTexto(
        "mgoFlash",
        "-"
    );


    setTexto(
        "nb2o5Flash",
        "-"
    );


    document
    .getElementById(
        "inicioPilhaFlash"
    )
    .value =
        "";

}


/* ============================================================
   CÁLCULOS
============================================================ */

function calcularIndicadoresFlash(){

    const massaAcumulada =
        numeroValor(
            "massaAcumuladaManualFlash"
        );


    const massaPlanejada =
        numeroValor(
            "massaPlanejadaFlash"
        );


    const meta =
        numeroValor(
            "metaHorariaFlash"
        );


    const media =
        numeroValor(
            "mediaHorariaFlash"
        );


    const manutencao =
        numeroValor(
            "manutencaoPlanejadaFlash"
        );


    /* FALTA BRITAR */

    let faltaBritar =
        null;


    if(
        massaPlanejada > 0
    ){

        faltaBritar =
            Math.max(

                massaPlanejada -
                massaAcumulada,

                0

            );


        setTexto(

            "faltaBritarFlash",

            formatarToneladas(
                faltaBritar
            )

        );

    }else{

        setTexto(
            "faltaBritarFlash",
            "-"
        );

    }


    /* % FORMAÇÃO */

    if(
        massaPlanejada > 0
    ){

        const percentual =
            limitarPercentual(

                (
                    massaAcumulada /
                    massaPlanejada
                ) *
                100

            );


        setTexto(

            "percentualFormacaoFlash",

            formatarNumero(
                percentual,
                1
            )

        );

    }else{

        setTexto(
            "percentualFormacaoFlash",
            "-"
        );

    }


    /* ATINGIMENTO */

    if(
        meta > 0 &&
        media > 0
    ){

        setTexto(

            "atingimentoMetaFlash",

            formatarNumero(

                (
                    media /
                    meta
                ) *
                100,

                1

            )

        );

    }else{

        setTexto(
            "atingimentoMetaFlash",
            "-"
        );

    }


    /* HORAS */

    let horasRestantes =
        null;


    if(
        faltaBritar !== null &&
        media > 0
    ){

        horasRestantes =

            (
                faltaBritar /
                media
            )

            +

            manutencao;


        setTexto(

            "horasRestantesFlash",

            formatarNumero(
                horasRestantes,
                1
            )

        );

    }else{

        setTexto(
            "horasRestantesFlash",
            "-"
        );

    }


    /* PREVISÃO */

    if(
        horasRestantes !== null
    ){

        const referencia =
            obterDataHoraFlashAtual();


        const previsao =
            adicionarHoras(

                referencia,

                horasRestantes

            );


        setTexto(

            "previsaoConclusaoFlash",

            formatarDataHora(
                previsao
            )

        );

    }else{

        setTexto(
            "previsaoConclusaoFlash",
            "-"
        );

    }

}


/* ============================================================
   DATA/HORA DO FLASH
============================================================ */

function obterDataHoraFlashAtual(){

    const data =
        valorCampo(
            "dataFlash"
        );


    const hora =
        valorCampo(
            "horarioFlash"
        );


    if(
        data &&
        hora
    ){

        const dataHora =
            new Date(

                data +
                "T" +
                hora

            );


        if(
            !isNaN(
                dataHora.getTime()
            )
        ){

            return dataHora;

        }

    }


    return new Date();

}


/* ============================================================
   IMPACTOS
============================================================ */

function adicionarImpacto(){

    quantidadeImpactos++;


    const numero =
        quantidadeImpactos;


    const container =
        document.getElementById(
            "listaImpactosFlash"
        );


    const item =
        document.createElement(
            "div"
        );


    item.className =
        "item-impacto-flash";


    item.id =
        "impactoFlash-" +
        numero;


    item.innerHTML = `

        <div class="numero-impacto-flash">

            ${String(numero).padStart(2,"0")}

        </div>


        <div class="campo-flash">

            <label>
                Área / Motivo
            </label>

            <select
                id="motivoImpacto-${numero}"
                class="campo-impacto-preview">

                <option value="">
                    Selecione
                </option>

                <option value="Mina">
                    Mina
                </option>

                <option value="Manutenção Mecânica">
                    Manutenção Mecânica
                </option>

                <option value="Manutenção Elétrica">
                    Manutenção Elétrica
                </option>

                <option value="Operação">
                    Operação
                </option>

                <option value="Processo">
                    Processo
                </option>

                <option value="Utilidades">
                    Utilidades
                </option>

                <option value="Outros">
                    Outros
                </option>

            </select>

        </div>


        <div class="campo-flash campo-impacto-descricao">

            <label>
                Impacto / Ocorrência
            </label>

            <input
                type="text"
                id="descricaoImpacto-${numero}"
                class="campo-impacto-preview"
                placeholder="Ex.: Mudança de frente de lavra">

        </div>


        <div class="campo-flash">

            <label>
                Tempo
            </label>

            <input
                type="text"
                id="tempoImpacto-${numero}"
                class="campo-impacto-preview"
                placeholder="Ex.: 00:45">

        </div>


        <button
            type="button"
            class="btn-remover-impacto-flash"
            onclick="removerImpacto(${numero})">

            ×

        </button>

    `;


    container.appendChild(
        item
    );


    item
    .querySelectorAll(
        ".campo-impacto-preview"
    )
    .forEach(
        function(campo){

            campo.addEventListener(
                "input",
                atualizarPreviewSeGerado
            );


            campo.addEventListener(
                "change",
                atualizarPreviewSeGerado
            );

        }
    );

}


/* ============================================================
   REMOVER IMPACTO
============================================================ */

function removerImpacto(
    numero
){

    const elemento =
        document.getElementById(

            "impactoFlash-" +
            numero

        );


    if(elemento){

        elemento.remove();

    }


    atualizarPreviewSeGerado();

}


/* ============================================================
   OBTER IMPACTOS
============================================================ */

function obterImpactosFlash(){

    const impactos =
        [];


    document
    .querySelectorAll(
        ".item-impacto-flash"
    )
    .forEach(
        function(elemento){

            const numero =
                Number(

                    elemento.id.replace(
                        "impactoFlash-",
                        ""
                    )

                );


            const motivo =
                valorCampo(
                    "motivoImpacto-" +
                    numero
                );


            const descricao =
                valorCampo(
                    "descricaoImpacto-" +
                    numero
                );


            const tempo =
                valorCampo(
                    "tempoImpacto-" +
                    numero
                );


            if(
                motivo ||
                descricao ||
                tempo
            ){

                impactos.push({

                    motivo:
                        motivo,

                    descricao:
                        descricao,

                    tempo:
                        tempo

                });

            }

        }
    );


    return impactos;

}


/* ============================================================
   GERAR FLASH
============================================================ */

function gerarFlash(){

    calcularIndicadoresFlash();


    if(
        !pilhaFlash.value
    ){

        alert(
            "Selecione a pilha em formação."
        );


        return;

    }


    if(
        !valorCampo(
            "dataFlash"
        )
    ){

        alert(
            "Informe a data do Flash."
        );


        return;

    }


    if(
        !valorCampo(
            "horarioFlash"
        )
    ){

        alert(
            "Selecione o horário do Flash."
        );


        return;

    }


    montarPreviewFlash();


    flashGerado =
        true;


    document
    .getElementById(
        "previewFlash"
    )
    .scrollIntoView({

        behavior:
            "smooth",

        block:
            "start"

    });

}


/* ============================================================
   PRÉVIA
============================================================ */

function montarPreviewFlash(){

    const preview =
        document.getElementById(
            "previewFlash"
        );


    const data =
        valorCampo(
            "dataFlash"
        );


    const horario =
        valorCampo(
            "horarioFlash"
        );


    const turno =
        valorCampo(
            "turnoFlash"
        );


    const turma =
        valorCampo(
            "turmaFlash"
        );


    const operador =
        valorCampo(
            "operadorFlash"
        );


    const coordenador =
        valorCampo(
            "coordenadorFlash"
        );


    const producao =
        numeroValor(
            "producaoFlash"
        );


    const massaAcumulada =
        numeroValor(
            "massaAcumuladaManualFlash"
        );


    const massaPlanejada =
        numeroValor(
            "massaPlanejadaFlash"
        );


    const meta =
        numeroValor(
            "metaHorariaFlash"
        );


    const media =
        numeroValor(
            "mediaHorariaFlash"
        );


    const manutencao =
        numeroValor(
            "manutencaoPlanejadaFlash"
        );


    const observacoes =
        valorCampo(
            "observacoesFlash"
        );


    const impactos =
        obterImpactosFlash();


    const frentes =
        obterFrentesFlash();


    const percentual =
        numeroSeguro(

            obterTexto(
                "percentualFormacaoFlash"
            )

        );


    const geradoEmFlash =
        new Intl.DateTimeFormat(
            "pt-BR",
            {
                dateStyle: "short",
                timeStyle: "short"
            }
        ).format(new Date());


    let htmlFrentes =
        "";


    if(
        frentes.length
    ){

        htmlFrentes =
            frentes
            .map(
                function(frente){

                    return `

                        <span class="flash-frente-tag">

                            ${escaparHTML(frente)}

                        </span>

                    `;

                }
            )
            .join("");

    }else{

        htmlFrentes = `

            <span class="flash-frente-tag">
                -
            </span>

        `;

    }


    let htmlImpactos =
        "";


    if(
        impactos.length === 0
    ){

        htmlImpactos = `

            <div class="flash-sem-impacto">

                Sem impactos relevantes registrados no período.

            </div>

        `;

    }else{

        htmlImpactos =
            impactos
            .map(
                function(item){

                    return `

                        <div class="flash-impacto-item">

                            <div>

                                <strong>

                                    ${
                                        escaparHTML(

                                            item.descricao
                                            ||
                                            item.motivo
                                            ||
                                            "Impacto"

                                        )
                                    }

                                </strong>


                                ${
                                    item.motivo &&
                                    item.descricao
                                    ?

                                    `<span>${escaparHTML(item.motivo)}</span>`

                                    :

                                    ""
                                }

                            </div>


                            ${
                                item.tempo
                                ?

                                `<b>${escaparHTML(item.tempo)}</b>`

                                :

                                ""
                            }

                        </div>

                    `;

                }
            )
            .join("");

    }


    preview.innerHTML = `

        <div class="flash-documento">


            <!-- CABEÇALHO -->

            <div class="flash-documento-cabecalho">


                <div class="flash-documento-marca">

                    <div class="flash-documento-logo">
                        <img src="../Britagem/Imagens/Logo_CMOC.png" alt="CMOC">
                    </div>

                    <div class="flash-documento-titulo">

                        <span>
                            PORTAL FVO • BRITAGEM
                        </span>

                        <h2>
                            FLASH OPERACIONAL
                        </h2>

                    </div>

                </div>


                <div class="flash-documento-horario">

                    <span>
                        FLASH
                    </span>

                    <strong>
                        ${escaparHTML(horario || "-")}
                    </strong>

                </div>

            </div>


            <!-- IDENTIFICAÇÃO -->

            <div class="flash-identificacao">


                <div>

                    <span>
                        Data
                    </span>

                    <strong>
                        ${formatarDataInput(data)}
                    </strong>

                </div>


                <div>

                    <span>
                        Turno
                    </span>

                    <strong>
                        ${escaparHTML(turno || "-")}
                    </strong>

                </div>


                <div>

                    <span>
                        Turma
                    </span>

                    <strong>
                        ${escaparHTML(turma || "-")}
                    </strong>

                </div>


                <div>

                    <span>
                        Operador
                    </span>

                    <strong>
                        ${escaparHTML(operador || "-")}
                    </strong>

                </div>


                <div>

                    <span>
                        Coordenador
                    </span>

                    <strong>
                        ${escaparHTML(coordenador || "-")}
                    </strong>

                </div>

            </div>


            <!-- PILHA / PRODUÇÃO -->

            <div class="flash-linha-principal">


                <div class="flash-card-pilha-principal">

                    <span>
                        PILHA EM FORMAÇÃO
                    </span>

                    <strong>

                        ${
                            dadosPilhaAtual
                            ?
                            escaparHTML(
                                dadosPilhaAtual.nomePilha
                            )
                            :
                            "-"
                        }

                    </strong>

                </div>


                <div class="flash-card-producao-principal">

                    <span>
                        PRODUÇÃO DO PERÍODO
                    </span>

                    <strong>

                        ${
                            producao > 0
                            ?
                            formatarToneladas(
                                producao
                            )
                            :
                            "-"
                        }

                        <small>
                            t
                        </small>

                    </strong>

                </div>

            </div>


            <!-- QUALIDADE -->

            <div class="flash-secao-titulo">

                <strong>
                    QUALIDADE ATUAL DA PILHA
                </strong>

            </div>


            <div class="flash-grid-qualidade">


                <div>

                    <span>
                        Massa Acumulada
                    </span>

                    <strong>

                        ${
                            massaAcumulada > 0
                            ?
                            formatarToneladas(
                                massaAcumulada
                            )
                            :
                            "-"
                        }

                        <small>
                            t
                        </small>

                    </strong>

                </div>


                <div>

                    <span>
                        P₂O₅
                    </span>

                    <strong>

                        ${obterTexto("p2o5Flash")}

                        <small>%</small>

                    </strong>

                </div>


                <div>

                    <span>
                        MgO
                    </span>

                    <strong>

                        ${obterTexto("mgoFlash")}

                        <small>%</small>

                    </strong>

                </div>


                <div>

                    <span>
                        Nb₂O₅
                    </span>

                    <strong>

                        ${obterTexto("nb2o5Flash")}

                        <small>%</small>

                    </strong>

                </div>

            </div>


            <!-- ACOMPANHAMENTO -->

            <div class="flash-secao-titulo">

                <strong>
                    ACOMPANHAMENTO DA FORMAÇÃO
                </strong>

            </div>


            <div class="flash-grid-formacao">


                <div>

                    <span>
                        Massa Planejada
                    </span>

                    <strong>

                        ${
                            massaPlanejada > 0
                            ?
                            formatarToneladas(
                                massaPlanejada
                            )
                            :
                            "-"
                        }

                        <small>t</small>

                    </strong>

                </div>


                <div>

                    <span>
                        Falta Britar
                    </span>

                    <strong>

                        ${obterTexto("faltaBritarFlash")}

                        <small>t</small>

                    </strong>

                </div>


                <div>

                    <span>
                        Formação
                    </span>

                    <strong>

                        ${obterTexto("percentualFormacaoFlash")}

                        <small>%</small>

                    </strong>

                </div>


                <div>

                    <span>
                        Média Horária
                    </span>

                    <strong>

                        ${
                            media > 0
                            ?
                            formatarNumero(
                                media,
                                0
                            )
                            :
                            "-"
                        }

                        <small>
                            t/h
                        </small>

                    </strong>

                </div>


                <div>

                    <span>
                        Meta Horária
                    </span>

                    <strong>

                        ${
                            meta > 0
                            ?
                            formatarNumero(
                                meta,
                                0
                            )
                            :
                            "-"
                        }

                        <small>
                            t/h
                        </small>

                    </strong>

                </div>


                <div>

                    <span>
                        Atingimento
                    </span>

                    <strong>

                        ${obterTexto("atingimentoMetaFlash")}

                        <small>%</small>

                    </strong>

                </div>

            </div>


            <!-- PROGRESSO -->

            <div class="flash-progresso-container">


                <div class="flash-progresso-topo">

                    <span>
                        Progresso da Formação
                    </span>

                    <strong>

                        ${obterTexto("percentualFormacaoFlash")}%

                    </strong>

                </div>


                <div class="flash-progresso-trilho">

                    <div
                        class="flash-progresso-barra"
                        style="
                            width:${limitarPercentual(percentual)}%;
                        ">
                    </div>

                </div>

            </div>


            <!-- PREVISÃO -->

            <div class="flash-previsao-destaque">


                <div>

                    <span>
                        HORAS RESTANTES
                    </span>

                    <strong>

                        ${obterTexto("horasRestantesFlash")}

                        <small>
                            h
                        </small>

                    </strong>

                </div>


                <div>

                    <span>
                        MANUTENÇÃO PLANEJADA
                    </span>

                    <strong>

                        ${formatarNumero(manutencao,1)}

                        <small>
                            h
                        </small>

                    </strong>

                </div>


                <div class="flash-previsao-principal">

                    <span>
                        PREVISÃO DE CONCLUSÃO
                    </span>

                    <strong>

                        ${obterTexto("previsaoConclusaoFlash")}

                    </strong>

                </div>

            </div>


            <!-- FECHAMENTO EXECUTIVO -->

            <div class="flash-faixa-final">

                <section class="flash-painel-final flash-painel-frentes">
                    <div class="flash-painel-final-titulo">
                        FRENTES / ORIGENS
                    </div>
                    <div class="flash-frentes-preview">
                        ${htmlFrentes}
                    </div>
                </section>

                <section class="flash-painel-final flash-painel-impactos">
                    <div class="flash-painel-final-titulo">
                        PRINCIPAIS IMPACTOS
                    </div>
                    <div class="flash-impactos-preview">
                        ${htmlImpactos}
                    </div>
                </section>

                <section class="flash-painel-final flash-painel-observacoes">
                    <div class="flash-painel-final-titulo">
                        OBSERVAÇÕES
                    </div>
                    <div class="flash-observacoes-preview">
                        <p>
                            ${
                                observacoes
                                ? escaparHTML(observacoes)
                                : "Sem observações adicionais."
                            }
                        </p>
                    </div>
                </section>

            </div>


            <!-- RODAPÉ -->

            <div class="flash-rodape">

                <span>
                    Flash Operacional • Britagem FVO
                </span>

                <span>
                    Qualidade baseada no Controle de Estocagem
                </span>

                <span>
                    Gerado em ${escaparHTML(geradoEmFlash)}
                </span>

            </div>

        </div>

    `;

}


/* ============================================================
   ATUALIZAR PRÉVIA
============================================================ */

function atualizarPreviewSeGerado(){

    if(
        flashGerado
    ){

        calcularIndicadoresFlash();

        montarPreviewFlash();

    }

}


/* ============================================================
   PDF
============================================================ */

function gerarPDFFlash(){

    if(
        !flashGerado
    ){

        gerarFlash();


        if(
            !flashGerado
        ){

            return;

        }

    }


    const tituloAnterior =
        document.title;


    const pilha =
        dadosPilhaAtual
        ?
        dadosPilhaAtual.nomePilha
        :
        "FVO";


    document.title =

        "Flash_Britagem_" +

        pilha +

        "_" +

        (
            valorCampo(
                "horarioFlash"
            )
            .replace(
                ":",
                "h"
            )
            ||
            "Flash"
        );


    document.body
    .classList
    .add(
        "modo-impressao-flash"
    );


    setTimeout(
        function(){

            window.print();


            setTimeout(
                function(){

                    document.body
                    .classList
                    .remove(
                        "modo-impressao-flash"
                    );


                    document.title =
                        tituloAnterior;

                },
                500
            );

        },
        150
    );

}


/* ============================================================
   EMAIL
============================================================ */

function enviarEmailFlash(){

    if(
        !flashGerado
    ){

        gerarFlash();


        if(
            !flashGerado
        ){

            return;

        }

    }


    const pilha =
        dadosPilhaAtual
        ?
        dadosPilhaAtual.nomePilha
        :
        "-";


    const impactos =
        obterImpactosFlash();


    const textoImpactos =
        impactos.length
        ?
        impactos
        .map(
            function(item,index){

                let texto =

                    (
                        index +
                        1
                    ) +

                    ". ";


                if(item.motivo){

                    texto +=
                        item.motivo +
                        " - ";

                }


                texto +=
                    item.descricao
                    ||
                    "Impacto";


                if(item.tempo){

                    texto +=

                        " (" +

                        item.tempo +

                        ")";

                }


                return texto;

            }
        )
        .join(
            "\n"
        )
        :
        "Sem impactos relevantes registrados.";


    const assunto =

        "Flash Operacional - Britagem FVO - " +

        pilha +

        " - " +

        valorCampo(
            "horarioFlash"
        );


    const mensagem =

`Prezados,

Segue o Flash Operacional da Britagem FVO.

Data: ${formatarDataInput(valorCampo("dataFlash"))}
Flash: ${valorCampo("horarioFlash")}
Turno: ${valorCampo("turnoFlash")}
Turma: ${valorCampo("turmaFlash")}
Operador: ${valorCampo("operadorFlash")}
Coordenador: ${valorCampo("coordenadorFlash")}

PILHA EM FORMAÇÃO: ${pilha}

Massa acumulada: ${formatarToneladas(numeroValor("massaAcumuladaManualFlash"))} t
Massa planejada: ${formatarToneladas(numeroValor("massaPlanejadaFlash"))} t
Falta britar: ${obterTexto("faltaBritarFlash")} t
Formação: ${obterTexto("percentualFormacaoFlash")}%

P2O5: ${obterTexto("p2o5Flash")}%
MgO: ${obterTexto("mgoFlash")}%
Nb2O5: ${obterTexto("nb2o5Flash")}%

Frentes / Origens:
${obterFrentesFlash().join(", ") || "-"}

Produção do período: ${formatarToneladas(numeroValor("producaoFlash"))} t
Média horária: ${formatarNumero(numeroValor("mediaHorariaFlash"),0)} t/h
Meta horária: ${formatarNumero(numeroValor("metaHorariaFlash"),0)} t/h
Atingimento: ${obterTexto("atingimentoMetaFlash")}%

Horas restantes: ${obterTexto("horasRestantesFlash")} h
Manutenção planejada: ${formatarNumero(numeroValor("manutencaoPlanejadaFlash"),1)} h
Previsão de conclusão: ${obterTexto("previsaoConclusaoFlash")}

PRINCIPAIS IMPACTOS:
${textoImpactos}

Observações:
${valorCampo("observacoesFlash") || "-"}

Atenciosamente.`;


    window.location.href =

        "mailto:" +

        emailsFlash.join(
            ";"
        ) +

        "?subject=" +

        encodeURIComponent(
            assunto
        ) +

        "&body=" +

        encodeURIComponent(
            mensagem
        );

}


/* ============================================================
   STATUS
============================================================ */

function atualizarStatusArquivo(
    texto,
    status
){

    const elemento =
        document.getElementById(
            "statusLeituraControle"
        );


    elemento.textContent =
        texto;


    elemento.classList.remove(

        "status-arquivo-ok",

        "status-arquivo-erro",

        "status-arquivo-processando"

    );


    if(status === "ok"){

        elemento.classList.add(
            "status-arquivo-ok"
        );

    }


    if(status === "erro"){

        elemento.classList.add(
            "status-arquivo-erro"
        );

    }


    if(
        status ===
        "processando"
    ){

        elemento.classList.add(
            "status-arquivo-processando"
        );

    }

}


/* ============================================================
   DATA ATUAL
============================================================ */

function preencherDataAtual(){

    const hoje =
        new Date();


    const ano =
        hoje.getFullYear();


    const mes =
        String(
            hoje.getMonth() +
            1
        )
        .padStart(
            2,
            "0"
        );


    const dia =
        String(
            hoje.getDate()
        )
        .padStart(
            2,
            "0"
        );


    document
    .getElementById(
        "dataFlash"
    )
    .value =

        ano +
        "-" +
        mes +
        "-" +
        dia;

}


/* ============================================================
   HORÁRIO AUTOMÁTICO
============================================================ */

function selecionarHorarioFlashAutomaticamente(){

    const agora =
        new Date();


    const minutos =
        agora.getHours() *
        60 +
        agora.getMinutes();


    let melhor =
        horariosFlashPadrao[0];


    let distancia =
        Infinity;


    horariosFlashPadrao.forEach(
        function(horario){

            const partes =
                horario.split(
                    ":"
                );


            const valor =

                Number(
                    partes[0]
                ) *
                60

                +

                Number(
                    partes[1]
                );


            const diferenca =
                Math.abs(
                    minutos -
                    valor
                );


            if(
                diferenca <
                distancia
            ){

                distancia =
                    diferenca;


                melhor =
                    horario;

            }

        }
    );


    document
    .getElementById(
        "horarioFlash"
    )
    .value =
        melhor;

}


/* ============================================================
   LIMPAR
============================================================ */

function limparFlash(){

    if(
        !confirm(
            "Deseja limpar os dados do Flash?"
        )
    ){

        return;

    }


    location.reload();

}


/* ============================================================
   NOME PILHA
============================================================ */

function normalizarNomePilha(
    nome
){

    const texto =
        String(nome)
        .trim()
        .toUpperCase();


    const match =
        texto.match(

            /(\d{1,3})\s*[-\/]?\s*([AB])$/

        );


    if(match){

        return (

            match[1] +
            "-" +
            match[2]

        );

    }


    return texto;

}


/* ============================================================
   NORMALIZAR TEXTO
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


    return String(valor)
        .normalize(
            "NFD"
        )
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .trim()
        .replace(
            /\s+/g,
            " "
        )
        .toUpperCase();

}


/* ============================================================
   NÚMERO DO EXCEL
============================================================ */

function normalizarNumeroExcel(
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
        .trim()
        .replace(
            /\s/g,
            ""
        );


    if(
        texto.includes(
            ","
        )
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
   É NÚMERO
============================================================ */

function ehNumero(
    valor
){

    return (
        normalizarNumeroExcel(
            valor
        )
        !== null
    );

}


/* ============================================================
   CONVERTER DATA
============================================================ */

function converterParaData(
    valor
){

    if(
        valor instanceof Date &&
        !isNaN(
            valor.getTime()
        )
    ){

        /*
            EVITA DATAS FALSAS / HORÁRIOS ISOLADOS.
        */

        if(
            valor.getFullYear() <
            2000
        ){

            return null;

        }


        return valor;

    }


    if(
        typeof valor ===
        "number" &&
        valor > 36526 &&
        valor < 80000
    ){

        const partes =
            XLSX.SSF.parse_date_code(
                valor
            );


        if(
            partes &&
            partes.y >= 2000
        ){

            return new Date(

                partes.y,

                partes.m - 1,

                partes.d,

                partes.H || 0,

                partes.M || 0,

                Math.floor(
                    partes.S ||
                    0
                )

            );

        }

    }


    return null;

}


/* ============================================================
   CAMPOS
============================================================ */

function numeroValor(
    id
){

    const elemento =
        document.getElementById(id);


    if(
        !elemento ||
        elemento.value === ""
    ){

        return 0;

    }


    const numero =
        Number(
            elemento.value
        );


    return Number.isFinite(
        numero
    )
    ?
    numero
    :
    0;

}


function valorCampo(
    id
){

    const elemento =
        document.getElementById(id);


    return elemento
    ?
    String(
        elemento.value ||
        ""
    )
    .trim()
    :
    "";

}


function setTexto(
    id,
    texto
){

    const elemento =
        document.getElementById(id);


    if(elemento){

        elemento.textContent =
            texto;

    }

}


function obterTexto(
    id
){

    const elemento =
        document.getElementById(id);


    return elemento
    ?
    elemento.textContent.trim()
    :
    "-";

}


/* ============================================================
   FORMATAR
============================================================ */

function formatarNumero(
    valor,
    casas = 1
){

    if(
        valor === null ||
        valor === undefined ||
        valor === "" ||
        !Number.isFinite(
            Number(valor)
        )
    ){

        return "-";

    }


    return Number(
        valor
    )
    .toLocaleString(
        "pt-BR",
        {

            minimumFractionDigits:
                casas,

            maximumFractionDigits:
                casas

        }
    );

}


function formatarToneladas(
    valor
){

    if(
        valor === null ||
        valor === undefined ||
        valor === "" ||
        !Number.isFinite(
            Number(valor)
        )
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


function formatarDataInput(
    data
){

    if(!data){

        return "-";

    }


    const partes =
        data.split(
            "-"
        );


    if(
        partes.length !== 3
    ){

        return data;

    }


    return (

        partes[2] +
        "/" +
        partes[1] +
        "/" +
        partes[0]

    );

}


/* ============================================================
   HORAS
============================================================ */

function adicionarHoras(
    data,
    horas
){

    return new Date(

        data.getTime()

        +

        horas *
        3600000

    );

}


/* ============================================================
   %
============================================================ */

function limitarPercentual(
    valor
){

    return Math.min(

        Math.max(
            Number(valor) ||
            0,
            0
        ),

        100

    );

}


/* ============================================================
   STRING → NÚMERO
============================================================ */

function numeroSeguro(
    valor
){

    const texto =
        String(
            valor ||
            ""
        )
        .replace(
            "%",
            ""
        )
        .replace(
            /\./g,
            ""
        )
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
    0;

}


/* ============================================================
   SEGURANÇA HTML
============================================================ */

function escaparHTML(
    texto
){

    return String(
        texto ||
        ""
    )
    .replace(
        /&/g,
        "&amp;"
    )
    .replace(
        /</g,
        "&lt;"
    )
    .replace(
        />/g,
        "&gt;"
    )
    .replace(
        /"/g,
        "&quot;"
    )
    .replace(
        /'/g,
        "&#039;"
    );

}


/* ============================================================
   INICIALIZAÇÃO
============================================================ */

document.addEventListener(
    "DOMContentLoaded",
    function(){

        preencherDataAtual();

        selecionarHorarioFlashAutomaticamente();

        adicionarImpacto();

        calcularIndicadoresFlash();

        atualizarTagsFrentes();

    }
);