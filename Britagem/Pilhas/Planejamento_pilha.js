let quantidadeEncabecamentos = 0;

let planejamentoCalculado = false;


/* ============================================================
   EMAILS
============================================================ */

const emailsPlanejamento = [

    "luiz.campos@br.cmoc.com",

    "thais.gilvana@br.cmoc.com"

];


/* ============================================================
   EVENTOS
============================================================ */

function conectarEventoPlanejamento(id,evento,funcao){

    const elemento = document.getElementById(id);

    if(!elemento){
        return;
    }

    elemento.addEventListener(evento,funcao);

}

conectarEventoPlanejamento("btnAdicionarEncabecamento","click",adicionarEncabecamento);
conectarEventoPlanejamento("btnCalcularPlanejamento","click",calcularPlanejamento);
conectarEventoPlanejamento("btnLimparPlanejamento","click",limparPlanejamento);
conectarEventoPlanejamento("btnGerarPDFPlanejamento","click",gerarPDFPlanejamento);
conectarEventoPlanejamento("btnEnviarEmailPlanejamento","click",enviarEmailPlanejamento);


/* ============================================================
   CAMPOS DE ATUALIZAÇÃO AUTOMÁTICA
============================================================ */

const camposAutomaticos = [

    "horasTrabalhadas",

    "metrosConsumidos",

    "pilhaRetomada",

    "metragemTotalRetomada",

    "metragemConsumidaRetomada",

    "ritmoRetomador",

    "manutencaoUsina",

    "inicioRetomada",

    "pilhaPlanejada",

    "massaPlanejada",

    "produtividadePlanejada",

    "manutencaoBritagem",

    "inicioFormacao",

    "comprimentoPatio",

    "massaBase",

    "inicioBase",

    "fimBase"

];


camposAutomaticos.forEach(
    function(id){

        const elemento =
            document.getElementById(id);


        if(!elemento){

            return;

        }


        elemento.addEventListener(
            "input",
            atualizacaoRapida
        );


        elemento.addEventListener(
            "change",
            atualizacaoRapida
        );

    }
);


/* ============================================================
   ATUALIZAÇÃO RÁPIDA
============================================================ */

function atualizacaoRapida(){

    atualizarAvanco();

    atualizarRetomada();

    atualizarBase();

    atualizarDistribuicao();

    atualizarDesenhoPilha();

    atualizarDesenhoRetomada();

}


/* ============================================================
   AVANÇO
============================================================ */

function atualizarAvanco(){

    const horas =
        numeroValor(
            "horasTrabalhadas"
        );


    const metros =
        numeroValor(
            "metrosConsumidos"
        );


    if(
        horas > 0 &&
        metros > 0
    ){

        const avanco =
            metros /
            horas;


        setTexto(
            "avancoCalculado",

            formatarNumero(
                avanco,
                2
            )
        );


        const ritmo =
            document.getElementById(
                "ritmoRetomador"
            );


        if(
            ritmo &&
            ritmo.value === ""
        ){

            ritmo.value =
                avanco.toFixed(
                    2
                );

        }

    }else{

        setTexto(
            "avancoCalculado",
            "-"
        );

    }

}


/* ============================================================
   RETOMADA
============================================================ */

function atualizarRetomada(){

    const total =
        numeroValor(
            "metragemTotalRetomada"
        );


    const consumido =
        Math.min(

            numeroValor(
                "metragemConsumidaRetomada"
            ),

            total > 0
            ?
            total
            :
            Infinity

        );


    const ritmo =
        numeroValor(
            "ritmoRetomador"
        );


    const manutencao =
        numeroValor(
            "manutencaoUsina"
        );


    const inicio =
        dataValor(
            "inicioRetomada"
        );


    if(
        total <= 0
    ){

        setTexto(
            "percentualRetomada",
            "-"
        );


        setTexto(
            "metragemRestante",
            "-"
        );


        setTexto(
            "resumoMetrosRestantes",
            "-"
        );


        return;

    }


    const percentual =
        limitarPercentual(
            (
                consumido /
                total
            ) *
            100
        );


    const restante =
        Math.max(
            total -
            consumido,
            0
        );


    setTexto(
        "percentualRetomada",

        formatarNumero(
            percentual,
            1
        ) +

        "%"
    );


    setTexto(
        "metragemRestante",

        formatarNumero(
            restante,
            1
        )
    );


    setTexto(
        "resumoMetrosRestantes",

        formatarNumero(
            restante,
            0
        ) +

        " m"
    );


    if(
        ritmo <= 0
    ){

        setTexto(
            "horasRestantesRetomada",
            "-"
        );


        setTexto(
            "fimPrevistoRetomada",
            "-"
        );


        atualizarDesenhoRetomada();

        return;

    }


    const horas =
        (
            restante /
            ritmo
        ) +
        manutencao;


    setTexto(
        "horasRestantesRetomada",

        formatarNumero(
            horas,
            1
        )
    );


    if(inicio){

        const fim =
            adicionarHoras(
                inicio,
                horas
            );


        setTexto(
            "fimPrevistoRetomada",

            formatarDataHora(
                fim
            )
        );

    }else{

        setTexto(
            "fimPrevistoRetomada",
            "-"
        );

    }


    atualizarDesenhoRetomada();

}


/* ============================================================
   BASE
============================================================ */

function atualizarBase(){

    const massa =
        numeroValor(
            "massaBase"
        );


    const inicio =
        numeroValor(
            "inicioBase"
        );


    const fim =
        numeroValor(
            "fimBase"
        );


    const comprimento =
        fim -
        inicio;


    if(
        comprimento <= 0
    ){

        setTexto(
            "comprimentoBase",
            "-"
        );


        setTexto(
            "toneladasMetroBase",
            "-"
        );


        return;

    }


    setTexto(
        "comprimentoBase",

        formatarNumero(
            comprimento,
            0
        )
    );


    setTexto(
        "toneladasMetroBase",

        massa > 0
        ?
        formatarNumero(
            massa /
            comprimento,
            1
        )
        :
        "-"
    );

}


/* ============================================================
   ADICIONAR ENCABEÇAMENTO
============================================================ */

function adicionarEncabecamento(){

    quantidadeEncabecamentos++;


    const numero =
        quantidadeEncabecamentos;


    const container =
        document.getElementById(
            "listaEncabecamentos"
        );


    if(!container){
        return;
    }


    const artigo =
        document.createElement(
            "article"
        );


    artigo.className =
        "etapa-formacao etapa-encabecamento";


    artigo.id =
        "encabecamento-" +
        numero;


    artigo.innerHTML = `

        <div class="cabecalho-etapa-formacao">

            <div>

                <span class="numero-etapa">

                    ENC. ${numero}

                </span>

                <h3>

                    Encabeçamento ${numero}

                </h3>

            </div>


            <button
                type="button"
                class="btn-excluir-encabecamento"
                onclick="removerEncabecamento(${numero})">

                ×

            </button>

        </div>


        <div class="grid-etapa-formacao">


            <div class="campo-planejamento">

                <label>
                    Massa
                </label>

                <div class="input-unidade">

                    <input
                        type="number"
                        id="massaEnc-${numero}"
                        class="campo-encabecamento"
                        placeholder="Ex.: 15000">

                    <span>t</span>

                </div>

            </div>


            <div class="campo-planejamento">

                <label>
                    Início
                </label>

                <div class="input-unidade">

                    <input
                        type="number"
                        id="inicioEnc-${numero}"
                        class="campo-encabecamento">

                    <span>m</span>

                </div>

            </div>


            <div class="campo-planejamento">

                <label>
                    Fim
                </label>

                <div class="input-unidade">

                    <input
                        type="number"
                        id="fimEnc-${numero}"
                        class="campo-encabecamento">

                    <span>m</span>

                </div>

            </div>


            <div class="campo-planejamento campo-calculado">

                <label>
                    Comprimento
                </label>

                <strong id="comprimentoEnc-${numero}">
                    -
                </strong>

                <small>m</small>

            </div>


            <div class="campo-planejamento campo-calculado">

                <label>
                    t/m
                </label>

                <strong id="toneladasMetroEnc-${numero}">
                    -
                </strong>

                <small>t/m</small>

            </div>


        </div>

    `;


    container.appendChild(
        artigo
    );


    artigo
    .querySelectorAll(
        ".campo-encabecamento"
    )
    .forEach(
        function(campo){

            campo.addEventListener(
                "input",
                function(){

                    atualizarEncabecamento(
                        numero
                    );


                    atualizarDistribuicao();


                    atualizarDesenhoPilha();

                }
            );

        }
    );


    sugerirPosicaoEncabecamento(
        numero
    );

}


/* ============================================================
   POSIÇÃO SUGERIDA
============================================================ */

function sugerirPosicaoEncabecamento(
    numero
){

    const comprimento =
        numeroValor(
            "comprimentoPatio"
        ) || 420;


    if(numero === 1){

        setValor(
            "inicioEnc-1",
            0
        );


        setValor(
            "fimEnc-1",
            comprimento / 2
        );

    }


    if(numero === 2){

        setValor(
            "inicioEnc-2",
            comprimento / 2
        );


        setValor(
            "fimEnc-2",
            comprimento
        );

    }


    atualizarEncabecamento(
        numero
    );


    atualizarDistribuicao();


    atualizarDesenhoPilha();

}


/* ============================================================
   ATUALIZAR ENCABEÇAMENTO
============================================================ */

function atualizarEncabecamento(
    numero
){

    const massa =
        numeroValor(
            "massaEnc-" +
            numero
        );


    const inicio =
        numeroValor(
            "inicioEnc-" +
            numero
        );


    const fim =
        numeroValor(
            "fimEnc-" +
            numero
        );


    const comprimento =
        fim -
        inicio;


    if(
        comprimento <= 0
    ){

        setTexto(
            "comprimentoEnc-" +
            numero,
            "-"
        );


        setTexto(
            "toneladasMetroEnc-" +
            numero,
            "-"
        );


        return;

    }


    setTexto(
        "comprimentoEnc-" +
        numero,

        formatarNumero(
            comprimento,
            0
        )
    );


    setTexto(
        "toneladasMetroEnc-" +
        numero,

        massa > 0
        ?
        formatarNumero(
            massa /
            comprimento,
            1
        )
        :
        "-"
    );

}


/* ============================================================
   REMOVER ENCABEÇAMENTO
============================================================ */

function removerEncabecamento(
    numero
){

    const elemento =
        document.getElementById(
            "encabecamento-" +
            numero
        );


    if(elemento){

        elemento.remove();

    }


    atualizarDistribuicao();

    atualizarDesenhoPilha();

}


/* ============================================================
   OBTER ENCABEÇAMENTOS
============================================================ */

function obterEncabecamentos(){

    const lista =
        [];


    document
    .querySelectorAll(
        ".etapa-encabecamento"
    )
    .forEach(
        function(elemento){

            const numero =
                Number(
                    elemento.id.replace(
                        "encabecamento-",
                        ""
                    )
                );


            lista.push({

                numero:
                    numero,

                massa:
                    numeroValor(
                        "massaEnc-" +
                        numero
                    ),

                inicio:
                    numeroValor(
                        "inicioEnc-" +
                        numero
                    ),

                fim:
                    numeroValor(
                        "fimEnc-" +
                        numero
                    )

            });

        }
    );


    return lista;

}


/* ============================================================
   DISTRIBUIÇÃO DE MASSA
============================================================ */

function atualizarDistribuicao(){

    const planejada =
        numeroValor(
            "massaPlanejada"
        );


    const base =
        numeroValor(
            "massaBase"
        );


    const encabecamentos =
        obterEncabecamentos();


    const massaEncabecamentos =
        encabecamentos.reduce(
            function(total,item){

                return (
                    total +
                    item.massa
                );

            },
            0
        );


    const distribuida =
        base +
        massaEncabecamentos;


    const diferenca =
        planejada -
        distribuida;


    setTexto(
        "massaPlanejadaResumo",

        formatarToneladas(
            planejada
        ) +

        " t"
    );


    setTexto(
        "massaDistribuida",

        formatarToneladas(
            distribuida
        ) +

        " t"
    );


    setTexto(
        "diferencaDistribuicao",

        formatarToneladasComSinal(
            diferenca
        )
    );


    const campo =
        document.getElementById(
            "diferencaDistribuicao"
        );


    if(!campo){
        return;
    }


    campo.classList.remove(
        "massa-ok",
        "massa-atencao",
        "massa-erro"
    );


    if(
        planejada <= 0
    ){

        return;

    }


    if(
        Math.abs(
            diferenca
        ) < 1
    ){

        campo.classList.add(
            "massa-ok"
        );

    }else if(
        Math.abs(
            diferenca
        ) <=
        planejada *
        0.05
    ){

        campo.classList.add(
            "massa-atencao"
        );

    }else{

        campo.classList.add(
            "massa-erro"
        );

    }

}


/* ============================================================
   CÁLCULO PRINCIPAL
============================================================ */

function calcularPlanejamento(){

    atualizarAvanco();

    atualizarRetomada();

    atualizarBase();

    atualizarDistribuicao();


    const massa =
        numeroValor(
            "massaPlanejada"
        );


    const produtividade =
        numeroValor(
            "produtividadePlanejada"
        );


    const manutencaoBritagem =
        numeroValor(
            "manutencaoBritagem"
        );


    const inicioFormacao =
        dataValor(
            "inicioFormacao"
        );


    const inicioRetomada =
        dataValor(
            "inicioRetomada"
        );


    const totalRetomada =
        numeroValor(
            "metragemTotalRetomada"
        );


    const consumida =
        numeroValor(
            "metragemConsumidaRetomada"
        );


    const ritmo =
        numeroValor(
            "ritmoRetomador"
        );


    const manutencaoUsina =
        numeroValor(
            "manutencaoUsina"
        );


    if(
        massa <= 0
    ){

        alert(
            "Informe a massa planejada."
        );

        return;

    }


    if(
        produtividade <= 0
    ){

        alert(
            "Informe a produtividade planejada."
        );

        return;

    }


    if(!inicioFormacao){

        alert(
            "Informe o início da formação."
        );

        return;

    }


    if(!inicioRetomada){

        alert(
            "Informe o início da retomada."
        );

        return;

    }


    if(
        ritmo <= 0
    ){

        alert(
            "Informe o ritmo do retomador."
        );

        return;

    }


    /* ========================================================
       FORMAÇÃO
    ======================================================== */

    const horasProducao =
        massa /
        produtividade;


    const tempoNecessario =
        horasProducao +
        manutencaoBritagem;


    const fimFormacao =
        adicionarHoras(
            inicioFormacao,
            tempoNecessario
        );


    /* ========================================================
       RETOMADA
    ======================================================== */

    const restante =
        Math.max(
            totalRetomada -
            consumida,
            0
        );


    const horasRetomada =
        (
            restante /
            ritmo
        ) +
        manutencaoUsina;


    const fimRetomada =
        adicionarHoras(
            inicioRetomada,
            horasRetomada
        );


    /* ========================================================
       JANELA
    ======================================================== */

    const janelaBruta =
        diferencaHoras(
            inicioFormacao,
            fimRetomada
        );


    const janelaEfetiva =
        Math.max(
            janelaBruta -
            manutencaoBritagem,
            0
        );


    const massaPossivel =
        janelaEfetiva *
        produtividade;


    const diferenca =
        massaPossivel -
        massa;


    const folga =
        diferencaHoras(
            fimFormacao,
            fimRetomada
        );


    /* ========================================================
       RESULTADOS
    ======================================================== */

    setTexto(
        "tempoNecessarioFormacao",

        formatarNumero(
            tempoNecessario,
            1
        ) +

        " h"
    );


    setTexto(
        "fimPrevistoFormacao",

        formatarDataHora(
            fimFormacao
        )
    );


    setTexto(
        "janelaDisponivel",

        formatarNumero(
            janelaEfetiva,
            1
        ) +

        " h"
    );


    setTexto(
        "resultadoMassaPlanejada",

        formatarToneladas(
            massa
        )
    );


    setTexto(
        "resultadoMassaPossivel",

        formatarToneladas(
            massaPossivel
        )
    );


    setTexto(
        "resultadoDiferencaMassa",

        formatarToneladasComSinal(
            diferenca
        )
    );


    setTexto(
        "resultadoTempoNecessario",

        formatarNumero(
            tempoNecessario,
            1
        )
    );


    setTexto(
        "resultadoJanela",

        formatarNumero(
            janelaEfetiva,
            1
        )
    );


    atualizarStatusPlanejamento(

        massa,

        massaPossivel,

        folga

    );


    atualizarRecomendacoes(

        massa,

        produtividade,

        manutencaoBritagem,

        janelaBruta,

        janelaEfetiva,

        massaPossivel

    );


    atualizarLinhaTempo(

        inicioFormacao,

        fimFormacao,

        inicioRetomada,

        fimRetomada

    );


    atualizarDesenhoPilha();

    atualizarDesenhoRetomada();


    planejamentoCalculado =
        true;


    atualizarResumoPDF();

}


/* ============================================================
   STATUS
============================================================ */

function atualizarStatusPlanejamento(
    planejada,
    possivel,
    folga
){

    const card =
        document.getElementById(
            "cardStatusPlanejamento"
        );


    card.classList.remove(
        "status-planejamento-ok",
        "status-planejamento-atencao",
        "status-planejamento-critico"
    );


    const percentual =
        possivel /
        planejada;


    if(
        percentual >= 1 &&
        folga >= 0
    ){

        card.classList.add(
            "status-planejamento-ok"
        );


        setTexto(
            "statusPlanejamento",
            "VIÁVEL"
        );


        setTexto(
            "detalheStatusPlanejamento",

            "Folga de " +

            formatarNumero(
                folga,
                1
            ) +

            " h"
        );

    }else if(
        percentual >= 0.95
    ){

        card.classList.add(
            "status-planejamento-atencao"
        );


        setTexto(
            "statusPlanejamento",
            "ATENÇÃO"
        );


        setTexto(
            "detalheStatusPlanejamento",
            "Próximo da meta"
        );

    }else{

        card.classList.add(
            "status-planejamento-critico"
        );


        setTexto(
            "statusPlanejamento",
            "INVIÁVEL"
        );


        setTexto(
            "detalheStatusPlanejamento",

            "Déficit de " +

            formatarToneladas(
                Math.max(
                    planejada -
                    possivel,
                    0
                )
            ) +

            " t"
        );

    }

}


/* ============================================================
   RECOMENDAÇÕES
============================================================ */

function atualizarRecomendacoes(
    massa,
    produtividade,
    manutencao,
    janelaBruta,
    janelaEfetiva,
    massaPossivel
){

    let produtividadeNecessaria =
        null;


    if(
        janelaEfetiva > 0
    ){

        produtividadeNecessaria =
            massa /
            janelaEfetiva;


        setTexto(
            "produtividadeNecessaria",

            formatarNumero(
                produtividadeNecessaria,
                0
            )
        );


        setTexto(
            "detalheProdutividadeNecessaria",
            "t/h"
        );

    }else{

        setTexto(
            "produtividadeNecessaria",
            "—"
        );


        setTexto(
            "detalheProdutividadeNecessaria",
            "Sem janela disponível"
        );

    }


    const horasNecessarias =
        (
            massa /
            produtividade
        ) +
        manutencao;


    const horasAdicionais =
        Math.max(
            horasNecessarias -
            janelaBruta,
            0
        );


    setTexto(
        "horasAdicionaisNecessarias",

        formatarNumero(
            horasAdicionais,
            1
        )
    );


    const reducaoManutencao =
        Math.min(
            manutencao,
            horasAdicionais
        );


    setTexto(
        "reducaoManutencaoNecessaria",

        formatarNumero(
            reducaoManutencao,
            1
        )
    );


    if(
        massaPossivel >=
        massa
    ){

        setTexto(
            "textoRecomendacaoPlanejamento",

            "A pilha pode ser concluída dentro da janela operacional disponível."

        );


        return;

    }


    if(
        janelaEfetiva <= 0
    ){

        setTexto(
            "textoRecomendacaoPlanejamento",

            "Não existe janela operacional suficiente antes do término previsto da pilha em retomada. " +

            "É necessário antecipar a formação, ampliar a janela ou revisar o planejamento da retomada."

        );


        return;

    }


    const deficit =
        massa -
        massaPossivel;


    setTexto(
        "textoRecomendacaoPlanejamento",

        "Déficit estimado de " +

        formatarToneladas(
            deficit
        ) +

        " t. A produtividade necessária para atingir a massa planejada dentro da janela é de aproximadamente " +

        formatarNumero(
            produtividadeNecessaria,
            0
        ) +

        " t/h."

    );

}


/* ============================================================
   DESENHO PILHA PLANEJADA
============================================================ */

function atualizarDesenhoPilha(){

    const container = document.getElementById("desenhoPilha");

    if(!container){
        return;
    }

    const comprimento = Math.max(numeroValor("comprimentoPatio") || 420,1);
    const massaBase = Math.max(numeroValor("massaBase"),0);
    const pilha = valorTexto("pilhaPlanejada");

    const inicioBase = Math.min(Math.max(numeroValor("inicioBase"),0),comprimento);
    const fimBase = Math.min(Math.max(numeroValor("fimBase"),0),comprimento);

    const encs = obterEncabecamentos()
        .map(function(item){
            return {
                numero:item.numero,
                massa:Math.max(item.massa,0),
                inicio:Math.min(Math.max(item.inicio,0),comprimento),
                fim:Math.min(Math.max(item.fim,0),comprimento)
            };
        })
        .filter(function(item){
            return item.massa > 0 && item.fim > item.inicio;
        })
        .sort(function(a,b){return a.inicio - b.inicio;});

    let massaTotal = massaBase;
    encs.forEach(function(item){massaTotal += item.massa;});

    setTexto("tituloDesenhoPilha",pilha || "-");
    setTexto("massaTotalDesenho",formatarToneladas(massaTotal) + " t");

    if(massaBase <= 0 || fimBase <= inicioBase){
        container.innerHTML = '<div class="mensagem-desenho-vazio">Preencha a estratégia de formação.</div>';
        return;
    }

    const larguraSVG = 840;
    const alturaSVG = 390;
    const margem = 62;
    const larguraUtil = larguraSVG - margem * 2;
    const escala = larguraUtil / comprimento;
    const solo = 275;
    const trilho = 306;
    const alturaBase = 66;
    const topoBase = solo - alturaBase;

    function x(metro){return margem + metro * escala;}

    const bx1 = x(inicioBase);
    const bx2 = x(fimBase);
    const inclBase = Math.min(48,(bx2-bx1)*.14);

    let svgEnc = "";
    let maiorTopo = topoBase;

    encs.forEach(function(item,index){
        const x1=x(item.inicio), x2=x(item.fim), largura=x2-x1;
        const densidade=item.massa/Math.max(item.fim-item.inicio,1);
        const altura=Math.min(Math.max(densidade*.48,42),92);
        const topo=topoBase-altura;
        maiorTopo=Math.min(maiorTopo,topo);
        const incl=Math.min(36,largura*.15);
        const centro=(x1+x2)/2;

        svgEnc += `
            <g>
                <path d="M ${x1} ${topoBase} L ${x1+incl} ${topo} L ${x2-incl} ${topo} L ${x2} ${topoBase} Z"
                      fill="url(#gradEncPilha)" stroke="#5e9725" stroke-width="1.2"></path>
                <path d="M ${x1+incl} ${topo} L ${x2-incl} ${topo} L ${x2-incl+11} ${topo+9} L ${x1+incl+11} ${topo+9} Z"
                      class="pilha-face-verde-svg"></path>
                <line x1="${x1+incl}" y1="${topo}" x2="${x2-incl}" y2="${topo}" class="linha-crista-svg"></line>
                <text x="${centro}" y="${topo+27}" text-anchor="middle" class="texto-pilha-enc">ENC. ${index+1}</text>
                <text x="${centro}" y="${topo+42}" text-anchor="middle" class="texto-pilha-enc-massa">${formatarToneladas(item.massa)} t</text>
                <text x="${centro}" y="${solo+20}" text-anchor="middle" class="texto-metragem-pilha">${formatarNumero(item.inicio,0)} -> ${formatarNumero(item.fim,0)} m</text>
            </g>`;
    });

    const segmentoAtivo = encs.length ? encs[encs.length-1] : {inicio:inicioBase,fim:fimBase};
    const equipamentoX = x((segmentoAtivo.inicio + segmentoAtivo.fim)/2);

    let regua="";
    for(let i=0;i<=4;i++){
        const metro=comprimento*i/4, px=x(metro);
        regua += `<line x1="${px}" y1="338" x2="${px}" y2="347" class="marca-regua-pilha"></line>
                  <text x="${px}" y="366" text-anchor="middle" class="texto-regua-pilha">${formatarNumero(metro,0)} m</text>`;
    }

    container.innerHTML = `
        <svg viewBox="0 0 ${larguraSVG} ${alturaSVG}" class="svg-pilha-planejamento" role="img" aria-label="Pilha planejada com empilhadeira de duas lanças">
            <style>
                .fundo-patio-svg{fill:#faf8fb;stroke:#e7dfea;stroke-width:1.2}.linha-solo-svg{stroke:#817887;stroke-width:2}.faixa-trilho-svg{fill:#e6e0e9}.linha-trilho-svg{stroke:#8d8492;stroke-width:2;stroke-dasharray:5 5}.pilha-sombra-svg{fill:#cfc7d4;opacity:.45}.pilha-face-lateral-svg{fill:#693b80;opacity:.86}.pilha-face-verde-svg{fill:#6da92b;opacity:.92}.linha-crista-svg{fill:none;stroke:#fff;stroke-opacity:.72;stroke-width:2}.rotulo-equipamento-svg{fill:#2d014d;font-size:10px;font-weight:900}.equipamento-base{fill:#34203f;stroke:#210038;stroke-width:1}.equipamento-roda{fill:#fff;stroke:#2d014d;stroke-width:3}.equipamento-torre{fill:#4d1a68;stroke:#2d014d;stroke-width:1}.equipamento-cabine{fill:#92d050;stroke:#5e9725;stroke-width:1}.equipamento-lanca,.equipamento-lanca-retomadora{fill:#57306a;stroke:#2d014d;stroke-width:1.2}.correia-equipamento{stroke:#d9f0bc;stroke-width:2}.ponto-descarga-equipamento{fill:#92d050;stroke:#fff;stroke-width:2}.roda-retomadora{fill:#fff;stroke:#2d014d;stroke-width:4}.dentes-retomadora line{stroke:#2d014d;stroke-width:3;stroke-linecap:round}.etiqueta-posicao-retomador{fill:#2d014d}
            </style>
            <defs>
                <linearGradient id="gradBasePilha" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stop-color="#63347a"></stop>
                    <stop offset="100%" stop-color="#351748"></stop>
                </linearGradient>
                <linearGradient id="gradEncPilha" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stop-color="#a5df59"></stop>
                    <stop offset="100%" stop-color="#68a724"></stop>
                </linearGradient>
</defs>

            <rect x="18" y="42" width="804" height="280" rx="18" class="fundo-patio-svg"></rect>
            <text x="${larguraSVG/2}" y="28" text-anchor="middle" class="titulo-desenho-pilha">${pilha ? "Pilha "+pilha : "Pilha Planejada"}</text>

            <rect x="${margem-10}" y="${trilho-7}" width="${larguraUtil+20}" height="14" rx="7" class="faixa-trilho-svg"></rect>
            <line x1="${margem-10}" y1="${trilho}" x2="${larguraSVG-margem+10}" y2="${trilho}" class="linha-trilho-svg"></line>
            <line x1="${margem-10}" y1="${solo}" x2="${larguraSVG-margem+10}" y2="${solo}" class="linha-solo-svg"></line>

            <path d="M ${bx1+9} ${solo+8} L ${bx1+inclBase+9} ${topoBase+8} L ${bx2-inclBase+9} ${topoBase+8} L ${bx2+9} ${solo+8} Z" class="pilha-sombra-svg"></path>
            <path d="M ${bx1} ${solo} L ${bx1+inclBase} ${topoBase} L ${bx2-inclBase} ${topoBase} L ${bx2} ${solo} Z" fill="url(#gradBasePilha)" stroke="#351748" stroke-width="1.3"></path>
            <path d="M ${bx1+inclBase} ${topoBase} L ${bx2-inclBase} ${topoBase} L ${bx2-inclBase+13} ${topoBase+10} L ${bx1+inclBase+13} ${topoBase+10} Z" class="pilha-face-lateral-svg"></path>
            <line x1="${bx1+inclBase}" y1="${topoBase}" x2="${bx2-inclBase}" y2="${topoBase}" class="linha-crista-svg"></line>
            <text x="${(bx1+bx2)/2}" y="${topoBase+38}" text-anchor="middle" class="texto-pilha-base">BASE • ${formatarToneladas(massaBase)} t</text>

            ${svgEnc}

            <g transform="translate(${equipamentoX} 0)">
                <text x="0" y="68" text-anchor="middle" class="rotulo-equipamento-svg">EMPILHADEIRA - 2 LANÇAS</text>
                <rect x="-33" y="${trilho-15}" width="66" height="13" rx="5" class="equipamento-base"></rect>
                <circle cx="-22" cy="${trilho+1}" r="7" class="equipamento-roda"></circle>
                <circle cx="22" cy="${trilho+1}" r="7" class="equipamento-roda"></circle>
                <rect x="-9" y="126" width="18" height="${trilho-141}" rx="5" class="equipamento-torre"></rect>
                <rect x="10" y="162" width="29" height="21" rx="4" class="equipamento-cabine"></rect>
                <path d="M -3 139 L -146 102 L -153 113 L -7 153 Z" class="equipamento-lanca"></path>
                <path d="M 3 139 L 146 102 L 153 113 L 7 153 Z" class="equipamento-lanca"></path>
                <line x1="-12" y1="140" x2="-143" y2="106" class="correia-equipamento"></line>
                <line x1="12" y1="140" x2="143" y2="106" class="correia-equipamento"></line>
                <circle cx="-149" cy="108" r="6" class="ponto-descarga-equipamento"></circle>
                <circle cx="149" cy="108" r="6" class="ponto-descarga-equipamento"></circle>
            </g>

            <line x1="${margem-10}" y1="338" x2="${larguraSVG-margem+10}" y2="338" class="linha-regua-pilha"></line>
            ${regua}
        </svg>`;
}


/* ============================================================
   DESENHO RETOMADA
============================================================ */

function atualizarDesenhoRetomada(){

    const container = document.getElementById("desenhoRetomada");

    if(!container){
        return;
    }

    const pilha = valorTexto("pilhaRetomada");
    const total = Math.max(numeroValor("metragemTotalRetomada"),0);
    const consumido = Math.min(Math.max(numeroValor("metragemConsumidaRetomada"),0),total);

    if(total <= 0){
        container.innerHTML = '<div class="mensagem-desenho-vazio">Informe a metragem da pilha.</div>';
        return;
    }

    setTexto("tituloDesenhoRetomada",pilha || "-");

    const restante=Math.max(total-consumido,0);
    setTexto("resumoMetrosRestantes",formatarNumero(restante,0)+" m");

    const largura=840, altura=390, esquerda=62, direita=62;
    const larguraUtil=largura-esquerda-direita;
    const escala=larguraUtil/total;
    const xConsumido=esquerda+consumido*escala;
    const final=esquerda+larguraUtil;
    const solo=275, trilho=306, topo=145;
    const percentual=total>0 ? consumido/total*100 : 0;

    let restantePath="";
    if(restante>0){
        const larguraRestante=final-xConsumido;
        const incl=Math.min(52,Math.max(10,larguraRestante*.16));
        const inclDireita=Math.min(52,Math.max(10,larguraRestante*.16));
        restantePath=`<path d="M ${xConsumido} ${solo} L ${Math.min(xConsumido+incl,final)} ${topo} L ${Math.max(final-inclDireita,xConsumido)} ${topo} L ${final} ${solo} Z"
                             fill="url(#gradRetomadaPilha)" stroke="#5e9725" stroke-width="1.4"></path>
                      <line x1="${Math.min(xConsumido+incl,final)}" y1="${topo}" x2="${Math.max(final-inclDireita,xConsumido)}" y2="${topo}" class="linha-crista-svg"></line>`;
    }

    let regua="";
    for(let i=0;i<=4;i++){
        const metro=total*i/4, px=esquerda+metro*escala;
        regua += `<line x1="${px}" y1="338" x2="${px}" y2="347" class="marca-regua-pilha"></line>
                  <text x="${px}" y="366" text-anchor="middle" class="texto-regua-pilha">${formatarNumero(metro,0)} m</text>`;
    }

    container.innerHTML=`
        <svg viewBox="0 0 ${largura} ${altura}" class="svg-pilha-retomada" role="img" aria-label="Pilha em retomada com retomadora">
            <style>
                .fundo-patio-svg{fill:#faf8fb;stroke:#e7dfea;stroke-width:1.2}.linha-solo-svg{stroke:#817887;stroke-width:2}.faixa-trilho-svg{fill:#e6e0e9}.linha-trilho-svg{stroke:#8d8492;stroke-width:2;stroke-dasharray:5 5}.pilha-sombra-svg{fill:#cfc7d4;opacity:.45}.pilha-face-lateral-svg{fill:#693b80;opacity:.86}.pilha-face-verde-svg{fill:#6da92b;opacity:.92}.linha-crista-svg{fill:none;stroke:#fff;stroke-opacity:.72;stroke-width:2}.rotulo-equipamento-svg{fill:#2d014d;font-size:10px;font-weight:900}.equipamento-base{fill:#34203f;stroke:#210038;stroke-width:1}.equipamento-roda{fill:#fff;stroke:#2d014d;stroke-width:3}.equipamento-torre{fill:#4d1a68;stroke:#2d014d;stroke-width:1}.equipamento-cabine{fill:#92d050;stroke:#5e9725;stroke-width:1}.equipamento-lanca,.equipamento-lanca-retomadora{fill:#57306a;stroke:#2d014d;stroke-width:1.2}.correia-equipamento{stroke:#d9f0bc;stroke-width:2}.ponto-descarga-equipamento{fill:#92d050;stroke:#fff;stroke-width:2}.roda-retomadora{fill:#fff;stroke:#2d014d;stroke-width:4}.dentes-retomadora line{stroke:#2d014d;stroke-width:3;stroke-linecap:round}.etiqueta-posicao-retomador{fill:#2d014d}
            </style>
            <defs>
                <linearGradient id="gradRetomadaPilha" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stop-color="#a4df58"></stop>
                    <stop offset="100%" stop-color="#68a724"></stop>
                </linearGradient>
</defs>

            <rect x="18" y="42" width="804" height="280" rx="18" class="fundo-patio-svg"></rect>
            <text x="${largura/2}" y="28" text-anchor="middle" class="titulo-desenho-pilha">${pilha ? "Pilha "+pilha : "Pilha em Retomada"}</text>

            <rect x="${esquerda-10}" y="${trilho-7}" width="${larguraUtil+20}" height="14" rx="7" class="faixa-trilho-svg"></rect>
            <line x1="${esquerda-10}" y1="${trilho}" x2="${final+10}" y2="${trilho}" class="linha-trilho-svg"></line>
            <line x1="${esquerda-10}" y1="${solo}" x2="${final+10}" y2="${solo}" class="linha-solo-svg"></line>

            <rect x="${esquerda}" y="${solo-18}" width="${Math.max(xConsumido-esquerda,0)}" height="18" rx="3" class="faixa-retomada-consumida"></rect>
            ${restantePath}

            <g transform="translate(${xConsumido} 0)">
                <text x="0" y="72" text-anchor="middle" class="rotulo-equipamento-svg">RETOMADORA</text>
                <line x1="0" y1="86" x2="0" y2="${trilho-16}" class="linha-posicao-retomador"></line>
                <rect x="-38" y="${trilho-15}" width="76" height="13" rx="5" class="equipamento-base"></rect>
                <circle cx="-26" cy="${trilho+1}" r="7" class="equipamento-roda"></circle>
                <circle cx="26" cy="${trilho+1}" r="7" class="equipamento-roda"></circle>
                <rect x="-8" y="177" width="16" height="${trilho-192}" rx="4" class="equipamento-torre"></rect>
                <path d="M -2 188 L 100 149 L 107 160 L 4 205 Z" class="equipamento-lanca-retomadora"></path>
                <circle cx="104" cy="154" r="22" class="roda-retomadora"></circle>
                <g class="dentes-retomadora">
                    <line x1="104" y1="126" x2="104" y2="137"></line>
                    <line x1="104" y1="171" x2="104" y2="182"></line>
                    <line x1="76" y1="154" x2="87" y2="154"></line>
                    <line x1="121" y1="154" x2="132" y2="154"></line>
                </g>
                <rect x="-43" y="91" width="86" height="28" rx="14" class="etiqueta-posicao-retomador"></rect>
                <text x="0" y="110" text-anchor="middle" class="texto-posicao-retomador" style="fill:#fff;font-size:10px">${formatarNumero(consumido,0)} m</text>
            </g>

            <g>
                <rect x="${largura/2-72}" y="63" width="144" height="31" rx="15" fill="#eef7e5" stroke="#92d050"></rect>
                <text x="${largura/2}" y="84" text-anchor="middle" class="texto-percentual-retomada">${formatarNumero(percentual,1)}% RETOMADO</text>
            </g>

            <line x1="${esquerda-10}" y1="338" x2="${final+10}" y2="338" class="linha-regua-pilha"></line>
            ${regua}
        </svg>`;
}


/* ============================================================
   LINHA DO TEMPO
============================================================ */

function atualizarLinhaTempo(
    inicioFormacao,
    fimFormacao,
    inicioRetomada,
    fimRetomada
){

    const menor =
        new Date(
            Math.min(
                inicioFormacao.getTime(),
                inicioRetomada.getTime()
            )
        );


    const maior =
        new Date(
            Math.max(
                fimFormacao.getTime(),
                fimRetomada.getTime()
            )
        );


    const intervalo =
        maior.getTime() -
        menor.getTime();


    if(
        intervalo <= 0
    ){

        return;

    }


    configurarBarraTempo(

        "barraLinhaFormacao",

        inicioFormacao,

        fimFormacao,

        menor,

        intervalo

    );


    configurarBarraTempo(

        "barraLinhaRetomada",

        inicioRetomada,

        fimRetomada,

        menor,

        intervalo

    );


    const pilhaFormacao =
        valorTexto(
            "pilhaPlanejada"
        );


    const pilhaRetomada =
        valorTexto(
            "pilhaRetomada"
        );


    setTexto(
        "nomeLinhaFormacao",

        pilhaFormacao
        ?
        "Formação " +
        pilhaFormacao
        :
        "Formação"
    );


    setTexto(
        "nomeLinhaRetomada",

        pilhaRetomada
        ?
        "Retomada " +
        pilhaRetomada
        :
        "Retomada"
    );


    setTexto(
        "periodoLinhaFormacao",

        formatarDataHora(
            inicioFormacao
        ) +

        " → " +

        formatarDataHora(
            fimFormacao
        )
    );


    setTexto(
        "periodoLinhaRetomada",

        formatarDataHora(
            inicioRetomada
        ) +

        " → " +

        formatarDataHora(
            fimRetomada
        )
    );

}


/* ============================================================
   BARRA DE TEMPO
============================================================ */

function configurarBarraTempo(
    id,
    inicio,
    fim,
    menor,
    intervalo
){

    const barra =
        document.getElementById(id);


    const esquerda =
        (
            (
                inicio.getTime() -
                menor.getTime()
            ) /
            intervalo
        ) *
        100;


    const largura =
        (
            (
                fim.getTime() -
                inicio.getTime()
            ) /
            intervalo
        ) *
        100;


    barra.style.left =
        limitarPercentual(
            esquerda
        ) +
        "%";


    barra.style.width =
        Math.max(
            limitarPercentual(
                largura
            ),
            2
        ) +
        "%";

}


/* ============================================================
   RESUMO PARA PDF
============================================================ */

function atualizarResumoPDF(){

    const pilhaPlanejada =
        valorTexto(
            "pilhaPlanejada"
        ) || "-";


    const pilhaRetomada =
        valorTexto(
            "pilhaRetomada"
        ) || "-";


    const metrosRestantes =
        obterTexto(
            "resumoMetrosRestantes"
        );


    const ritmo =
        numeroValor(
            "ritmoRetomador"
        );


    const produtividade =
        numeroValor(
            "produtividadePlanejada"
        );


    setTexto(
        "pdfNomePilha",
        pilhaPlanejada
    );


    setTexto(
        "pdfPilhaRetomada",
        pilhaRetomada
    );


    setTexto(
        "pdfRetomadaDetalhe",

        metrosRestantes +

        " restantes • " +

        (
            ritmo > 0
            ?
            formatarNumero(
                ritmo,
                1
            ) +
            " m/h"
            :
            "-"
        )
    );


    setTexto(
        "pdfPilhaPlanejada",
        pilhaPlanejada
    );


    setTexto(
        "pdfPilhaDetalhe",

        "Conclusão prevista: " +

        obterTexto(
            "fimPrevistoFormacao"
        )
    );


    setTexto(
        "pdfMassaPlanejada",

        obterTexto(
            "resultadoMassaPlanejada"
        )
    );


    setTexto(
        "pdfMassaPossivel",

        obterTexto(
            "resultadoMassaPossivel"
        )
    );


    setTexto(
        "pdfDiferencaMassa",

        obterTexto(
            "resultadoDiferencaMassa"
        )
    );


    setTexto(
        "pdfStatus",

        obterTexto(
            "statusPlanejamento"
        )
    );


    setTexto(
        "pdfStatusDetalhe",

        obterTexto(
            "detalheStatusPlanejamento"
        )
    );


    setTexto(
        "pdfFimRetomada",

        obterTexto(
            "fimPrevistoRetomada"
        )
    );


    setTexto(
        "pdfFimFormacao",

        obterTexto(
            "fimPrevistoFormacao"
        )
    );


    setTexto(
        "pdfProdutividade",

        produtividade > 0
        ?
        formatarNumero(
            produtividade,
            0
        ) +
        " t/h"
        :
        "-"
    );


    const produtividadeNecessaria =
        obterTexto(
            "produtividadeNecessaria"
        );


    setTexto(
        "pdfProdutividadeNecessaria",

        produtividadeNecessaria === "—"
        ||
        produtividadeNecessaria === "-"
        ?
        "Não aplicável"
        :
        produtividadeNecessaria +
        " t/h"
    );


    const card =
        document.getElementById(
            "pdfCardStatus"
        );


    card.classList.remove(
        "pdf-status-ok",
        "pdf-status-atencao",
        "pdf-status-critico"
    );


    const status =
        obterTexto(
            "statusPlanejamento"
        );


    if(
        status ===
        "VIÁVEL"
    ){

        card.classList.add(
            "pdf-status-ok"
        );

    }else if(
        status ===
        "ATENÇÃO"
    ){

        card.classList.add(
            "pdf-status-atencao"
        );

    }else{

        card.classList.add(
            "pdf-status-critico"
        );

    }

}


/* ============================================================
   GERAR PDF
============================================================ */

function gerarPDFPlanejamento(){

    if(!planejamentoCalculado){
        alert("Calcule o planejamento antes de gerar o PDF.");
        return;
    }

    atualizarResumoPDF();

    const tituloAnterior = document.title;
    document.title = "Planejamento_pilha_" + (valorTexto("pilhaPlanejada") || "FVO");
    document.body.classList.add("modo-pdf-planejamento");

    let restaurado = false;

    function restaurarImpressao(){
        if(restaurado){return;}
        restaurado = true;
        document.body.classList.remove("modo-pdf-planejamento");
        document.title = tituloAnterior;
        window.removeEventListener("afterprint",restaurarImpressao);
    }

    window.addEventListener("afterprint",restaurarImpressao);

    requestAnimationFrame(function(){
        requestAnimationFrame(function(){
            setTimeout(function(){
                window.print();
                setTimeout(restaurarImpressao,1800);
            },180);
        });
    });
}


/* ============================================================
   EMAIL
============================================================ */

function enviarEmailPlanejamento(){

    if(
        !planejamentoCalculado
    ){

        alert(
            "Calcule o planejamento antes de enviar."
        );


        return;

    }


    const assunto =

        "Planejamento de Pilhas - " +

        valorTexto(
            "pilhaPlanejada"
        );


    const mensagem =

`Prezados,

Segue o planejamento operacional das pilhas da Britagem FVO.

Pilha em retomada: ${valorTexto("pilhaRetomada")}
Metros restantes: ${obterTexto("resumoMetrosRestantes")}
Fim previsto da retomada: ${obterTexto("fimPrevistoRetomada")}

Pilha planejada: ${valorTexto("pilhaPlanejada")}
Massa planejada: ${obterTexto("resultadoMassaPlanejada")} t
Massa possível: ${obterTexto("resultadoMassaPossivel")} t
Fim previsto da formação: ${obterTexto("fimPrevistoFormacao")}

Status: ${obterTexto("statusPlanejamento")}
Produtividade necessária: ${obterTexto("produtividadeNecessaria")}

Atenciosamente.`;


    window.location.href =

        "mailto:" +

        emailsPlanejamento.join(
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
   LIMPAR
============================================================ */

function limparPlanejamento(){

    if(
        !confirm(
            "Deseja limpar o planejamento?"
        )
    ){

        return;

    }


    location.reload();

}


/* ============================================================
   UTILIDADES
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


function valorTexto(
    id
){

    const elemento =
        document.getElementById(id);


    return elemento
    ?
    elemento.value.trim()
    :
    "";

}


function dataValor(
    id
){

    const elemento =
        document.getElementById(id);


    if(
        !elemento ||
        !elemento.value
    ){

        return null;

    }


    const data =
        new Date(
            elemento.value
        );


    return isNaN(
        data.getTime()
    )
    ?
    null
    :
    data;

}


function adicionarHoras(
    data,
    horas
){

    return new Date(

        data.getTime() +

        horas *
        3600000

    );

}


function diferencaHoras(
    inicio,
    fim
){

    return (

        fim.getTime() -
        inicio.getTime()

    ) /
    3600000;

}


function formatarNumero(
    valor,
    casas = 1
){

    if(
        valor === null ||
        valor === undefined ||
        valor === "" ||
        !Number.isFinite(
            Number(
                valor
            )
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
            Number(
                valor
            )
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


function formatarToneladasComSinal(
    valor
){

    if(
        valor === null ||
        valor === undefined ||
        !Number.isFinite(
            Number(
                valor
            )
        )
    ){

        return "-";

    }


    const numero =
        Number(
            valor
        );


    return (

        (
            numero > 0
            ?
            "+"
            :
            ""
        ) +

        formatarToneladas(
            numero
        ) +

        " t"

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


function limitarPercentual(
    valor
){

    return Math.min(
        Math.max(
            valor,
            0
        ),
        100
    );

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


function setValor(
    id,
    valor
){

    const elemento =
        document.getElementById(id);


    if(elemento){

        elemento.value =
            valor;

    }

}


/* ============================================================
   INICIALIZAÇÃO
============================================================ */

document.addEventListener(
    "DOMContentLoaded",
    function(){

        atualizarAvanco();

        atualizarRetomada();

        atualizarBase();

        atualizarDistribuicao();

        atualizarDesenhoPilha();

        atualizarDesenhoRetomada();

    }
);