/* ============================================================
   PORTAL FVO - EQUIPAMENTOS DA BRITAGEM PRIMÁRIA
============================================================ */

"use strict";

const equipamentosPrimaria = {
    grelha: {
        ordem: "01",
        categoria: "RECEPÇÃO",
        etapa: "Recepção do ROM",
        nome: "Grelha Fixa",
        descricao: "Equipamento responsável pela retenção inicial de blocos de maior dimensão antes da alimentação da britagem primária.",
        capacidade: "Conforme operação",
        fabricante: "-",
        modelo: "Grelha fixa",
        funcionamento: "O minério proveniente da mina é descarregado sobre a grelha fixa. Os fragmentos compatíveis com a abertura passam para o circuito, enquanto blocos de maior dimensão ficam retidos para tratamento antes de seguirem para a etapa seguinte.",
        caracteristicas: [
            "Abertura aproximada de 160 mm",
            "Primeira etapa de controle granulométrico",
            "Estrutura fixa",
            "Retenção de blocos de grandes dimensões",
            "Proteção dos equipamentos posteriores"
        ],
        imagem: "../Imagens/Primaria/Grelha.png"
    },

    alimentador: {
        ordem: "02",
        categoria: "ALIMENTAÇÃO",
        etapa: "Alimentação do circuito",
        nome: "Alimentador de Sapatas",
        descricao: "Equipamento utilizado para receber o minério após a grelha e promover alimentação contínua e controlada para o circuito de britagem primária.",
        capacidade: "Conforme operação",
        fabricante: "Metso",
        modelo: "Alimentador de Sapatas",
        funcionamento: "O equipamento utiliza uma sequência de sapatas metálicas movimentadas por acionamento mecânico. As sapatas transportam o minério de maneira contínua, permitindo uma alimentação mais estável para os equipamentos seguintes.",
        caracteristicas: [
            "Construção robusta para minério de grande granulometria",
            "Alimentação contínua do circuito",
            "Sapatas metálicas de alta resistência",
            "Controle da taxa de alimentação",
            "Equipamento submetido a elevado esforço mecânico"
        ],
        imagem: "../Imagens/Primaria/Alimentador de sapatas.jpg"
    },

    divisor: {
        ordem: "03",
        categoria: "DISTRIBUIÇÃO",
        etapa: "Distribuição entre linhas",
        nome: "Divisor de Fluxo",
        descricao: "Responsável pela distribuição do minério entre as linhas A e B da britagem primária.",
        capacidade: "Conforme alimentação",
        fabricante: "-",
        modelo: "Divisor A/B",
        funcionamento: "Após a alimentação, o minério chega ao divisor de fluxo. O equipamento direciona o material para uma das duas linhas de processamento ou distribui a alimentação entre ambas, conforme a condição operacional da planta.",
        caracteristicas: [
            "Distribuição de minério entre Linha A e Linha B",
            "Permite operação com uma ou duas linhas",
            "Elemento importante para estabilidade da alimentação",
            "Possibilita isolamento operacional de uma das linhas",
            "Atua antes do peneiramento primário"
        ],
        imagem: "../Imagens/Primaria/Divisor de fluxo.png"
    },

    peneiraA: {
        ordem: "04",
        categoria: "CLASSIFICAÇÃO",
        etapa: "Classificação - Linha A",
        nome: "Peneira A",
        descricao: "Peneira vibratória responsável pela classificação granulométrica do minério na Linha A da britagem primária.",
        capacidade: "660 t/h",
        fabricante: "Metso",
        modelo: "CBS 8' × 20' DD",
        funcionamento: "O minério é distribuído sobre as telas da peneira e submetido à vibração. As partículas menores que as aberturas atravessam as telas, enquanto o material de maior granulometria segue para a britagem por impacto.",
        caracteristicas: [
            "Peneira vibratória de dois decks",
            "Dimensões aproximadas de 8' × 20'",
            "Capacidade aproximada de 660 t/h",
            "Classificação granulométrica antes da britagem",
            "Instalada na Linha A"
        ],
        imagem: "../Imagens/Primaria/Peneira A.jpg"
    },

    peneiraB: {
        ordem: "05",
        categoria: "CLASSIFICAÇÃO",
        etapa: "Classificação - Linha B",
        nome: "Peneira B",
        descricao: "Peneira vibratória responsável pela classificação granulométrica do minério na Linha B da britagem primária.",
        capacidade: "660 t/h",
        fabricante: "Metso",
        modelo: "CBS 8' × 20' DD",
        funcionamento: "O minério é distribuído sobre as telas da peneira e submetido à vibração. As partículas menores que as aberturas atravessam as telas, enquanto o material de maior granulometria segue para a britagem por impacto.",
        caracteristicas: [
            "Peneira vibratória de dois decks",
            "Dimensões aproximadas de 8' × 20'",
            "Capacidade aproximada de 660 t/h",
            "Classificação granulométrica antes da britagem",
            "Instalada na Linha B"
        ],
        imagem: "../Imagens/Primaria/Peneira B.jpg"
    },

    britadorA: {
        ordem: "06",
        categoria: "COMINUIÇÃO",
        etapa: "Britagem por impacto - Linha A",
        nome: "Britador de Impacto A",
        descricao: "Equipamento responsável pela redução granulométrica do material retido no peneiramento da Linha A.",
        capacidade: "400 t/h",
        fabricante: "HAZEMAG",
        modelo: "AP-M 1615",
        funcionamento: "O material alimentado no britador entra na região de atuação do rotor. As barras de impacto transferem energia ao minério e promovem sucessivos impactos contra os elementos internos do equipamento, reduzindo a granulometria.",
        caracteristicas: [
            "Britador de impacto horizontal",
            "Modelo HAZEMAG AP-M 1615",
            "Capacidade aproximada de 400 t/h",
            "Razão de redução aproximada de 8:1",
            "Rotor com barras substituíveis e reversíveis",
            "Utilizado na Linha A da britagem primária"
        ],
        imagem: "../Imagens/Primaria/Britador A.png"
    },

    britadorB: {
        ordem: "07",
        categoria: "COMINUIÇÃO",
        etapa: "Britagem por impacto - Linha B",
        nome: "Britador de Impacto B",
        descricao: "Equipamento responsável pela redução granulométrica do material retido no peneiramento da Linha B.",
        capacidade: "400 t/h",
        fabricante: "HAZEMAG",
        modelo: "AP-M 1615",
        funcionamento: "O material alimentado no britador entra na região de atuação do rotor. As barras de impacto transferem energia ao minério e promovem sucessivos impactos contra os elementos internos do equipamento, reduzindo a granulometria.",
        caracteristicas: [
            "Britador de impacto horizontal",
            "Modelo HAZEMAG AP-M 1615",
            "Capacidade aproximada de 400 t/h",
            "Razão de redução aproximada de 8:1",
            "Rotor com barras substituíveis e reversíveis",
            "Utilizado na Linha B da britagem primária"
        ],
        imagem: "../Imagens/Primaria/Britador B.jpeg"
    },

    correias: {
        ordem: "08",
        categoria: "TRANSPORTE",
        etapa: "Transporte de minério",
        nome: "Correias Transportadoras",
        descricao: "Sistema responsável pelo transporte contínuo do minério entre os equipamentos e etapas da britagem.",
        capacidade: "Variável por correia",
        fabricante: "Diversos",
        modelo: "Transportadores de Correia",
        funcionamento: "O minério é transportado sobre uma correia contínua movimentada por tambores de acionamento. Os transportadores conectam os diversos equipamentos do circuito, garantindo o fluxo contínuo de material.",
        caracteristicas: [
            "Transporte contínuo de minério",
            "Interligação entre os equipamentos",
            "Compostos por correia, tambores, roletes e acionamentos",
            "Possuem sistemas de proteção e emergência",
            "Elemento crítico para continuidade operacional"
        ],
        imagem: "../Imagens/Primaria/Correia transportadora.jpeg"
    }
};

const modal = document.getElementById("modalEquipamento");
const btnFecharModal = document.getElementById("fecharModalEquipamento");
const campoBusca = document.getElementById("buscaEquipamentos");
const botoesFiltro = Array.from(document.querySelectorAll(".filtro-categoria"));
const cardsEquipamentos = Array.from(document.querySelectorAll(".card-equipamento"));
const contadorEquipamentos = document.getElementById("contadorEquipamentos");
const estadoVazio = document.getElementById("estadoVazioEquipamentos");

let filtroCategoriaAtual = "todos";
let ultimoElementoFocado = null;

function normalizarTexto(valor) {
    return String(valor || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim();
}

function atualizarCatalogo() {
    const termo = normalizarTexto(campoBusca?.value);
    let totalVisivel = 0;

    cardsEquipamentos.forEach((card) => {
        const categoria = normalizarTexto(card.dataset.categoria);
        const nome = normalizarTexto(card.dataset.nome);
        const combinaCategoria = filtroCategoriaAtual === "todos" || categoria === filtroCategoriaAtual;
        const combinaBusca = !termo || nome.includes(termo);
        const visivel = combinaCategoria && combinaBusca;

        card.hidden = !visivel;
        card.classList.toggle("filtrado", !visivel);

        if (visivel) totalVisivel += 1;
    });

    if (contadorEquipamentos) {
        contadorEquipamentos.textContent = `${totalVisivel} ${totalVisivel === 1 ? "equipamento" : "equipamentos"}`;
    }

    if (estadoVazio) {
        estadoVazio.hidden = totalVisivel !== 0;
    }
}

botoesFiltro.forEach((botao) => {
    botao.addEventListener("click", () => {
        filtroCategoriaAtual = normalizarTexto(botao.dataset.filtro || "todos");
        botoesFiltro.forEach((item) => item.classList.toggle("ativo", item === botao));
        atualizarCatalogo();
    });
});

campoBusca?.addEventListener("input", atualizarCatalogo);

document.querySelectorAll(".btn-ver-equipamento").forEach((botao) => {
    botao.addEventListener("click", function () {
        abrirEquipamento(this.dataset.equipamento);
    });
});

function definirTexto(id, valor) {
    const elemento = document.getElementById(id);
    if (elemento) elemento.textContent = valor ?? "-";
}

function criarPlaceholderImagem(nomeEquipamento) {
    return `
        <div class="modal-placeholder">
            <svg viewBox="0 0 24 24" aria-hidden="true">
                <rect x="4" y="5" width="16" height="14" rx="2" fill="none" stroke="currentColor" stroke-width="2"/>
                <path d="M6 16l4-4 3 3 2-2 3 3" fill="none" stroke="currentColor" stroke-width="2"/>
            </svg>
            <span>${nomeEquipamento || "Imagem indisponível"}</span>
        </div>
    `;
}

function abrirEquipamento(idEquipamento) {
    const equipamento = equipamentosPrimaria[idEquipamento];
    if (!equipamento || !modal) return;

    ultimoElementoFocado = document.activeElement;

    definirTexto("modalCategoria", equipamento.categoria);
    definirTexto("modalNome", equipamento.nome);
    definirTexto("modalDescricao", equipamento.descricao);
    definirTexto("modalCapacidade", equipamento.capacidade);
    definirTexto("modalFabricante", equipamento.fabricante);
    definirTexto("modalModelo", equipamento.modelo);
    definirTexto("modalFuncionamento", equipamento.funcionamento);
    definirTexto("modalOrdem", equipamento.ordem);
    definirTexto("modalEtapa", equipamento.etapa);

    const lista = document.getElementById("modalCaracteristicas");
    if (lista) {
        lista.innerHTML = "";
        equipamento.caracteristicas.forEach((item) => {
            const li = document.createElement("li");
            li.textContent = item;
            lista.appendChild(li);
        });
    }

    const areaImagem = modal.querySelector(".modal-imagem-equipamento");
    if (areaImagem) {
        areaImagem.innerHTML = "";
        const imagem = document.createElement("img");
        imagem.src = equipamento.imagem;
        imagem.alt = equipamento.nome;
        imagem.addEventListener("error", () => {
            areaImagem.innerHTML = criarPlaceholderImagem(equipamento.nome);
        }, { once: true });
        areaImagem.appendChild(imagem);
    }

    modal.classList.add("aberto");
    modal.setAttribute("aria-hidden", "false");
    document.body.classList.add("modal-equipamento-aberto");
    btnFecharModal?.focus();
}

function fecharModal() {
    if (!modal) return;

    modal.classList.remove("aberto");
    modal.setAttribute("aria-hidden", "true");
    document.body.classList.remove("modal-equipamento-aberto");

    if (ultimoElementoFocado instanceof HTMLElement) {
        ultimoElementoFocado.focus();
    }
}

btnFecharModal?.addEventListener("click", fecharModal);

modal?.addEventListener("click", (event) => {
    if (event.target === modal) fecharModal();
});

document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && modal?.classList.contains("aberto")) {
        fecharModal();
    }
});

atualizarCatalogo();