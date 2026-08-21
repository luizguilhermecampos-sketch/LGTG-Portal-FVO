(function iniciarQualidadeFVO() {
  "use strict";

  const ELEMENTOS = [
    { id: "p2o5", nome: "P2O5", abas: ["P2O5"], percentual: true, fallbackPilha: 24 },
    { id: "fe2o3", nome: "Fe2O3", abas: ["Fe2O3"], percentual: true, fallbackPilha: 29 },
    { id: "sio2", nome: "SiO2", abas: ["SiO2"], percentual: true, fallbackPilha: 30 },
    { id: "bao", nome: "BaO", abas: ["BaO"], percentual: true, fallbackPilha: 28 },
    { id: "mgo", nome: "MgO", abas: ["MgO"], percentual: true, fallbackPilha: 25 },
    { id: "al2o3", nome: "Al2O3", abas: ["Al2O3"], percentual: true, fallbackPilha: 31 },
    { id: "nb2o5", nome: "Nb2O5", abas: ["Nb2O5"], percentual: true, fallbackPilha: 26 },
    { id: "cao", nome: "CaO", abas: ["CaO"], percentual: true, fallbackPilha: null },
    { id: "caop2o5", nome: "CaO/P2O5", abas: ["CaO P2O5", "CaO/P2O5", "CaO-P2O5"], percentual: false, fallbackPilha: 27 }
  ];

  const GRUPOS = ["u47", "u76", "consolidado"];
  const MEDIDAS = ["amg", "cf", "rf"];
  const CORES = {
    roxoEscuro: "#3d0764",
    roxoClaro: "#a987db",
    verdeEscuro: "#3b861a",
    verdeClaro: "#94d65c",
    laranja: "#ee9622",
    cinza: "#766f79"
  };

  const COLUNAS_INDICADORES = {
    data: "A",
    pilha: ["DX", "DU"],
    diario: {
      u47: { amg: "B", cf: "I", rf: "J" },
      u76: { amg: "P", cf: "W", rf: "X" },
      consolidado: { amg: "CH", cf: "CO", rf: "CP" }
    },
    acumuladoPilha: {
      u47: { amg: "DY", cf: "EF", rf: "EG" },
      u76: { amg: "EJ", cf: "EQ", rf: "ER" },
      consolidado: { amg: "EU", cf: "FB", rf: "FC" }
    }
  };

  const estado = {
    indicadores: new Map(),
    pilhasArquivo: [],
    resumos: new Map(),
    graficosMini: new Map(),
    graficosEvolucao: [],
    graficoMassa: null,
    arquivoIndicadores: "",
    arquivoPilhas: "",
    painelAtivo: "visaoGeral"
  };

  const $ = (id) => document.getElementById(id);

  document.addEventListener("DOMContentLoaded", () => {
    $("arquivoIndicadoresQualidade")?.addEventListener("change", carregarIndicadores);
    $("arquivoAnalisePilha")?.addEventListener("change", carregarAnalisePilha);
    $("btnAplicarFiltroQualidade")?.addEventListener("click", () => atualizarDashboard(true));
    $("btnLimparFiltroQualidade")?.addEventListener("click", limparFiltros);
    $("btnGerarPdfQualidade")?.addEventListener("click", gerarPDF);
    $("filtroPilhaQualidade")?.addEventListener("change", () => atualizarDashboard(false));
    $("btnFecharModalEvolucao")?.addEventListener("click", fecharModalEvolucao);
    $("modalEvolucaoQualidade")?.addEventListener("click", (evento) => {
      if (evento.target === $("modalEvolucaoQualidade")) fecharModalEvolucao();
    });
    document.addEventListener("keydown", (evento) => {
      if (evento.key === "Escape") fecharModalEvolucao();
    });
    $("cardsElementosQualidade")?.addEventListener("click", (evento) => {
      const botao = evento.target.closest("[data-evolucao]");
      if (botao) abrirModalEvolucao(botao.dataset.evolucao);
    });
    document.querySelectorAll(".aba-qualidade").forEach((botao) => {
      botao.addEventListener("click", () => ativarAba(botao.dataset.painel));
    });
    renderizarEstadoInicial();
  });

  function normalizar(texto) {
    return String(texto ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[₂]/g, "2")
      .replace(/[₃]/g, "3")
      .replace(/[^a-zA-Z0-9/]+/g, " ")
      .trim()
      .toLowerCase();
  }

  function escaparHTML(valor) {
    return String(valor ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function numero(valor) {
    if (typeof valor === "number") return Number.isFinite(valor) ? valor : null;
    if (valor == null || valor === "") return null;
    let texto = String(valor).trim().replace(/\s/g, "").replace(/%/g, "");
    if (!texto || texto.startsWith("#") || texto === "-") return null;
    if (texto.includes(",")) texto = texto.replace(/\./g, "").replace(",", ".");
    const convertido = Number(texto);
    return Number.isFinite(convertido) ? convertido : null;
  }

  function teor(valor) {
    const convertido = numero(valor);
    return convertido != null && convertido > 0 ? convertido : null;
  }

  function dataExcel(valor) {
    if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
      return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
    }
    if (typeof valor === "number" && window.XLSX?.SSF) {
      const partes = XLSX.SSF.parse_date_code(valor);
      return partes ? new Date(partes.y, partes.m - 1, partes.d) : null;
    }
    const texto = String(valor ?? "").trim();
    let partes = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (partes) {
      const ano = Number(partes[3]) < 100 ? 2000 + Number(partes[3]) : Number(partes[3]);
      return new Date(ano, Number(partes[2]) - 1, Number(partes[1]));
    }
    partes = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (partes) return new Date(Number(partes[1]), Number(partes[2]) - 1, Number(partes[3]));
    return null;
  }

  function iso(data) {
    if (!(data instanceof Date) || Number.isNaN(data.getTime())) return "";
    return [data.getFullYear(), String(data.getMonth() + 1).padStart(2, "0"), String(data.getDate()).padStart(2, "0")].join("-");
  }

  function dataDoInput(id) {
    const valor = $(id)?.value;
    if (!valor) return null;
    const partes = valor.split("-").map(Number);
    return partes.length === 3 ? new Date(partes[0], partes[1] - 1, partes[2]) : null;
  }

  function formatarData(data) {
    return data instanceof Date && !Number.isNaN(data.getTime()) ? data.toLocaleDateString("pt-BR") : "N/D";
  }

  function formatarDataCurta(data) {
    return data instanceof Date && !Number.isNaN(data.getTime())
      ? data.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" })
      : "N/D";
  }

  function formatarNumero(valor, casas = 2) {
    if (valor == null || !Number.isFinite(valor)) return "N/D";
    return new Intl.NumberFormat("pt-BR", { minimumFractionDigits: casas, maximumFractionDigits: casas }).format(valor);
  }

  function formatarTeor(valor, elemento) {
    if (valor == null || !Number.isFinite(valor)) return "N/D";
    return `${formatarNumero(valor, elemento.percentual ? 2 : 3)}${elemento.percentual ? "%" : ""}`;
  }

  function formatarMassa(valor) {
    return valor == null || !Number.isFinite(valor) ? "N/D" : formatarNumero(valor, 0);
  }

  function formatarMetragem(valor) {
    return valor == null || !Number.isFinite(valor) ? "N/D" : formatarNumero(valor, 0);
  }

  function texto(id, valor) {
    const elemento = $(id);
    if (elemento) elemento.textContent = valor;
  }

  function exibirCarregamento(mensagem) {
    if (window.FVOCarregamento) window.FVOCarregamento.exibir(mensagem);
  }

  function ocultarCarregamento() {
    window.FVOCarregamento?.ocultar();
  }

  function aviso(mensagem, tipo = "info") {
    const elemento = $("avisoQualidade");
    if (!elemento) return;
    elemento.textContent = mensagem;
    elemento.className = `aviso-qualidade aviso-${tipo}`;
  }

  function celula(planilha, coluna, linha) {
    return planilha?.[`${coluna}${linha}`] || null;
  }

  function valorCelula(planilha, coluna, linha) {
    return celula(planilha, coluna, linha)?.v ?? null;
  }

  function numeroCelula(planilha, coluna, linha) {
    return teor(valorCelula(planilha, coluna, linha));
  }

  function buscarAba(workbook, nomes) {
    const procurados = (Array.isArray(nomes) ? nomes : [nomes]).map(normalizar);
    let nome = workbook.SheetNames.find((item) => procurados.includes(normalizar(item)));
    if (!nome) {
      nome = workbook.SheetNames.find((item) => procurados.some((procurado) => normalizar(item).includes(procurado) || procurado.includes(normalizar(item))));
    }
    return nome ? { nome, planilha: workbook.Sheets[nome] } : null;
  }

  async function lerWorkbook(arquivo, abas) {
    if (typeof XLSX === "undefined") throw new Error("A biblioteca de leitura do Excel não foi carregada.");
    const buffer = await arquivo.arrayBuffer();
    const opcoes = { type: "array", cellDates: true, cellFormula: true, cellStyles: false, dense: false };
    if (abas?.length) opcoes.sheets = abas;
    return XLSX.read(new Uint8Array(buffer), opcoes);
  }

  async function carregarIndicadores(evento) {
    const arquivo = evento.target.files?.[0];
    if (!arquivo) return;
    exibirCarregamento(`Lendo ${arquivo.name}...`);
    try {
      const nomesAbas = [...new Set(ELEMENTOS.flatMap((item) => item.abas))];
      const workbook = await lerWorkbook(arquivo, nomesAbas);
      const mapeados = mapearIndicadores(workbook);
      if (!mapeados.size) throw new Error("Nenhuma das nove abas químicas foi encontrada.");
      estado.indicadores = mapeados;
      estado.arquivoIndicadores = arquivo.name;
      texto("nomeArquivoIndicadoresQualidade", arquivo.name);
      configurarPeriodoCompleto(true);
      atualizarDashboard(false);
      aviso(`Indicadores carregados: ${arquivo.name}`, "sucesso");
    } catch (erro) {
      console.error(erro);
      aviso(`Erro ao ler Indicadores 2026: ${erro.message}`, "erro");
    } finally {
      ocultarCarregamento();
    }
  }

  async function carregarAnalisePilha(evento) {
    const arquivo = evento.target.files?.[0];
    if (!arquivo) return;
    exibirCarregamento(`Lendo ${arquivo.name}...`);
    try {
      const workbook = await lerWorkbook(arquivo);
      const pilhas = mapearAnalisePilha(workbook);
      if (!pilhas.length) throw new Error("Nenhuma pilha válida foi localizada na aba de pilhas.");
      estado.pilhasArquivo = pilhas;
      estado.arquivoPilhas = arquivo.name;
      texto("nomeArquivoAnalisePilha", arquivo.name);
      configurarPeriodoCompleto(!estado.indicadores.size);
      atualizarDashboard(false);
      aviso(`Análise de Pilha carregada: ${arquivo.name}`, "sucesso");
    } catch (erro) {
      console.error(erro);
      aviso(`Erro ao ler Análise de Pilha: ${erro.message}`, "erro");
    } finally {
      ocultarCarregamento();
    }
  }

  function mapearIndicadores(workbook) {
    const resultado = new Map();
    ELEMENTOS.forEach((elemento) => {
      const aba = buscarAba(workbook, elemento.abas);
      if (!aba?.planilha) return;
      resultado.set(elemento.id, extrairRegistrosIndicadores(aba.planilha));
    });
    return resultado;
  }

  function extrairRegistrosIndicadores(planilha) {
    if (!planilha?.["!ref"]) return [];
    const faixa = XLSX.utils.decode_range(planilha["!ref"]);
    const ultimaLinha = Math.min(faixa.e.r + 1, 5000);
    const registros = [];
    let iniciou = false;
    let vazias = 0;
    for (let linha = 3; linha <= ultimaLinha; linha += 1) {
      const data = dataExcel(valorCelula(planilha, COLUNAS_INDICADORES.data, linha));
      if (!data) {
        if (iniciou) vazias += 1;
        if (vazias >= 50) break;
        continue;
      }
      iniciou = true;
      vazias = 0;
      const pilha = obterPilhaIndicadores(planilha, linha);
      registros.push({
        linha,
        data,
        pilha,
        diario: lerGrupos(planilha, linha, COLUNAS_INDICADORES.diario),
        acumuladoPilha: lerGrupos(planilha, linha, COLUNAS_INDICADORES.acumuladoPilha)
      });
    }
    return registros;
  }

  function lerGrupos(planilha, linha, mapa) {
    return Object.fromEntries(GRUPOS.map((grupo) => [grupo, Object.fromEntries(MEDIDAS.map((medida) => [medida, numeroCelula(planilha, mapa[grupo][medida], linha)]))]));
  }

  function obterPilhaIndicadores(planilha, linha) {
    const principal = String(valorCelula(planilha, COLUNAS_INDICADORES.pilha[0], linha) ?? "").trim();
    if (principal && /\d/.test(principal)) return principal.replace(/\s+/g, " ");
    const alternativa = String(valorCelula(planilha, COLUNAS_INDICADORES.pilha[1], linha) ?? "").trim();
    if (parecePilha(alternativa)) return alternativa.replace(/\s+/g, " ");
    return "";
  }

  function parecePilha(valor) {
    const textoNormalizado = normalizar(valor);
    return Boolean(textoNormalizado && /\d/.test(textoNormalizado) && !/^\d+(?:[.,]\d+)?$/.test(textoNormalizado));
  }

  function mapearAnalisePilha(workbook) {
    const aba = buscarAba(workbook, ["Pilha", "Análise de Pilha", "Analise de Pilha", "MPU"]);
    if (!aba?.planilha) throw new Error('A aba "Pilha" não foi encontrada.');
    const linhas = XLSX.utils.sheet_to_json(aba.planilha, { header: 1, defval: null, raw: true });
    if (!linhas.length) return [];
    const indiceCabecalho = detectarLinhaCabecalho(linhas);
    const cabecalhos = cabecalhosCompostos(linhas, indiceCabecalho);
    const colunas = {
      pilha: localizarColuna(cabecalhos, ["pilha", "numero pilha", "n pilha", "identificacao pilha"], 0),
      metragem: localizarColuna(cabecalhos, ["metragem", "comprimento", "metros"], 5),
      massa: localizarColuna(cabecalhos, ["massa total", "massa pilha", "massa"], 18),
      inicio: localizarColuna(cabecalhos, ["data inicio", "inicio formacao", "data inicial", "inicio"], 21),
      fim: localizarColuna(cabecalhos, ["data fim", "fim formacao", "data final", "termino", "fim"], 22),
      teores: {}
    };
    ELEMENTOS.forEach((elemento) => {
      colunas.teores[elemento.id] = localizarColunaElemento(cabecalhos, elemento);
    });

    const pilhas = [];
    for (let indice = indiceCabecalho + 1; indice < linhas.length; indice += 1) {
      const linha = linhas[indice] || [];
      const pilha = String(linha[colunas.pilha] ?? "").trim();
      if (!pilha || normalizar(pilha) === "pilha") continue;
      const massa = numero(linha[colunas.massa]);
      const metragem = numero(linha[colunas.metragem]);
      const inicio = dataExcel(linha[colunas.inicio]);
      const fim = dataExcel(linha[colunas.fim]);
      const teores = {};
      ELEMENTOS.forEach((elemento) => {
        const coluna = colunas.teores[elemento.id];
        teores[elemento.id] = coluna == null ? null : teor(linha[coluna]);
      });
      if (massa == null && metragem == null && !inicio && !fim && !Object.values(teores).some((valor) => valor != null)) continue;
      pilhas.push({ id: `${normalizarPilha(pilha)}-${indice + 1}`, linha: indice + 1, pilha: pilha.replace(/\s+/g, " "), massa, metragem, inicio, fim, teores });
    }
    return pilhas;
  }

  function detectarLinhaCabecalho(linhas) {
    let melhor = { indice: 0, pontos: -1 };
    linhas.slice(0, 25).forEach((linha, indice) => {
      const textos = (linha || []).map(normalizar);
      let pontos = 0;
      if (textos.some((item) => item === "pilha" || item.includes("numero pilha"))) pontos += 4;
      if (textos.some((item) => item.includes("massa"))) pontos += 2;
      if (textos.some((item) => item.includes("metragem") || item.includes("comprimento"))) pontos += 2;
      if (textos.some((item) => item.includes("inicio"))) pontos += 1;
      if (textos.some((item) => item.includes("fim") || item.includes("termino"))) pontos += 1;
      if (pontos > melhor.pontos) melhor = { indice, pontos };
    });
    return melhor.indice;
  }

  function cabecalhosCompostos(linhas, indice) {
    const totalColunas = Math.max(...linhas.slice(Math.max(0, indice - 1), indice + 2).map((linha) => linha?.length || 0), 0);
    return Array.from({ length: totalColunas }, (_, coluna) => {
      const partes = [];
      for (let linha = Math.max(0, indice - 1); linha <= Math.min(linhas.length - 1, indice + 1); linha += 1) {
        const valor = linhas[linha]?.[coluna];
        if (valor != null && String(valor).trim()) partes.push(String(valor));
      }
      return normalizar(partes.join(" "));
    });
  }

  function localizarColuna(cabecalhos, aliases, fallback) {
    const normalizados = aliases.map(normalizar);
    let indice = cabecalhos.findIndex((item) => normalizados.includes(item));
    if (indice < 0) indice = cabecalhos.findIndex((item) => normalizados.some((alias) => item.includes(alias)));
    return indice >= 0 ? indice : fallback;
  }

  function localizarColunaElemento(cabecalhos, elemento) {
    const aliases = {
      p2o5: ["p2o5", "teor p2o5"], fe2o3: ["fe2o3", "teor fe", "ferro"], sio2: ["sio2", "silica"],
      bao: ["bao", "bario"], mgo: ["mgo", "magnesio"], al2o3: ["al2o3", "aluminio"],
      nb2o5: ["nb2o5", "teor nb", "niobio"], cao: ["cao", "teor cao"],
      caop2o5: ["cao p2o5", "cao/p2o5", "ca p", "relacao ca p"]
    }[elemento.id] || [elemento.nome];
    const normalizados = aliases.map(normalizar);
    let indice = cabecalhos.findIndex((item) => normalizados.includes(item));
    if (indice < 0) {
      indice = cabecalhos.findIndex((item) => normalizados.some((alias) => item.includes(alias)) && (elemento.id !== "cao" || !item.includes("p2o5")));
    }
    return indice >= 0 ? indice : elemento.fallbackPilha;
  }

  function configurarPeriodoCompleto(forcar = false) {
    const datas = todasAsDatasDisponiveis();
    if (!datas.length) return;
    datas.sort((a, b) => a - b);
    if (forcar || !$("dataInicialQualidade").value) $("dataInicialQualidade").value = iso(datas[0]);
    if (forcar || !$("dataFinalQualidade").value) $("dataFinalQualidade").value = iso(datas[datas.length - 1]);
  }

  function todasAsDatasDisponiveis() {
    const datas = [];
    estado.indicadores.forEach((registros) => registros.forEach((registro) => datas.push(registro.data)));
    estado.pilhasArquivo.forEach((pilha) => {
      if (pilha.inicio) datas.push(pilha.inicio);
      if (pilha.fim) datas.push(pilha.fim);
    });
    return datas.filter((data) => data instanceof Date && !Number.isNaN(data.getTime()));
  }

  function limparFiltros() {
    $("filtroPilhaQualidade").value = "";
    configurarPeriodoCompleto(true);
    atualizarDashboard(false);
  }

  function periodoAtual() {
    return { inicio: dataDoInput("dataInicialQualidade"), fim: dataDoInput("dataFinalQualidade") };
  }

  function validarPeriodo(exibirErro = true) {
    const periodo = periodoAtual();
    if (!periodo.inicio || !periodo.fim || periodo.inicio > periodo.fim) {
      if (exibirErro) aviso("Verifique as datas inicial e final.", "erro");
      return null;
    }
    return periodo;
  }

  function normalizarPilha(valor) {
    return normalizar(String(valor ?? "").replace(/\bpilha\b/gi, "")).replace(/\s+/g, "");
  }

  function numerosDaPilha(valor) {
    return [...String(valor ?? "").matchAll(/\d+/g)].map((item) => item[0]);
  }

  function pilhasRelacionadas(a, b) {
    const na = normalizarPilha(a);
    const nb = normalizarPilha(b);
    if (!na || !nb) return false;
    if (na === nb) return true;
    const numerosA = numerosDaPilha(a);
    const numerosB = numerosDaPilha(b);
    return numerosA.some((numeroA) => numerosB.includes(numeroA));
  }

  function registroNoPeriodo(registro, periodo) {
    return registro.data >= periodo.inicio && registro.data <= periodo.fim;
  }

  function pilhaNoPeriodo(pilha, periodo) {
    if (pilha.inicio && pilha.fim) return pilha.inicio <= periodo.fim && pilha.fim >= periodo.inicio;
    if (pilha.inicio) return pilha.inicio >= periodo.inicio && pilha.inicio <= periodo.fim;
    if (pilha.fim) return pilha.fim >= periodo.inicio && pilha.fim <= periodo.fim;
    return true;
  }

  function registrosFiltrados(elementoId, considerarPilha = true) {
    const periodo = validarPeriodo(false);
    if (!periodo) return [];
    const pilhaSelecionada = considerarPilha ? $("filtroPilhaQualidade").value : "";
    return (estado.indicadores.get(elementoId) || []).filter((registro) => {
      if (!registroNoPeriodo(registro, periodo)) return false;
      return !pilhaSelecionada || normalizarPilha(registro.pilha) === normalizarPilha(pilhaSelecionada);
    });
  }

  function pilhasFiltradas() {
    const periodo = validarPeriodo(false);
    if (!periodo) return [];
    const selecionada = $("filtroPilhaQualidade").value;
    return estado.pilhasArquivo.filter((pilha) => pilhaNoPeriodo(pilha, periodo) && (!selecionada || pilhasRelacionadas(pilha.pilha, selecionada)));
  }

  function valorCaminho(objeto, caminho) {
    return caminho.reduce((atual, chave) => atual?.[chave], objeto);
  }

  function mediaValores(registros, caminho) {
    const valores = registros.map((registro) => valorCaminho(registro, caminho)).filter((valor) => valor != null && Number.isFinite(valor));
    return valores.length ? valores.reduce((total, valor) => total + valor, 0) / valores.length : null;
  }

  function ultimoValor(registros, caminho) {
    for (let indice = registros.length - 1; indice >= 0; indice -= 1) {
      const valor = valorCaminho(registros[indice], caminho);
      if (valor != null && Number.isFinite(valor)) return valor;
    }
    return null;
  }

  function resumirElemento(elemento) {
    const registros = registrosFiltrados(elemento.id, true).sort((a, b) => a.data - b.data);
    const porPilha = Boolean($("filtroPilhaQualidade").value);
    const origem = porPilha ? "acumuladoPilha" : "diario";
    const resumo = { elemento, registros };
    GRUPOS.forEach((grupo) => {
      resumo[grupo] = {};
      MEDIDAS.forEach((medida) => {
        const caminho = [origem, grupo, medida];
        resumo[grupo][medida] = porPilha ? ultimoValor(registros, caminho) : mediaValores(registros, caminho);
      });
    });
    return resumo;
  }

  function atualizarDashboard(exibirErroPeriodo) {
    const periodo = validarPeriodo(exibirErroPeriodo);
    if (!periodo) return;
    popularFiltroPilhas(periodo);
    estado.resumos = new Map(ELEMENTOS.map((elemento) => [elemento.id, resumirElemento(elemento)]));
    renderizarKPIsIndicadores(periodo);
    renderizarCardsElementos();
    renderizarTabelasResumo();
    renderizarPilhas(periodo);
    renderizarRastreabilidade();
    renderizarDados();
    atualizarTextosGerais(periodo);
    if (estado.indicadores.size) aviso(`Dashboard atualizado para ${formatarData(periodo.inicio)} a ${formatarData(periodo.fim)}.`, "sucesso");
    else aviso("Carregue a planilha Indicadores 2026 para visualizar AMG, CF e RF.", "info");
  }

  function popularFiltroPilhas(periodo) {
    const seletor = $("filtroPilhaQualidade");
    const anterior = seletor.value;
    const nomes = new Map();
    estado.indicadores.forEach((registros) => registros.filter((registro) => registroNoPeriodo(registro, periodo)).forEach((registro) => {
      if (registro.pilha) nomes.set(normalizarPilha(registro.pilha), registro.pilha);
    }));
    estado.pilhasArquivo.filter((pilha) => pilhaNoPeriodo(pilha, periodo)).forEach((pilha) => nomes.set(normalizarPilha(pilha.pilha), pilha.pilha));
    const ordenadas = [...nomes.values()].sort(ordenarPilhas);
    seletor.innerHTML = '<option value="">Todas as pilhas</option>' + ordenadas.map((pilha) => `<option value="${escaparHTML(pilha)}">${escaparHTML(pilha)}</option>`).join("");
    if (ordenadas.some((pilha) => normalizarPilha(pilha) === normalizarPilha(anterior))) {
      seletor.value = ordenadas.find((pilha) => normalizarPilha(pilha) === normalizarPilha(anterior));
    }
  }

  function ordenarPilhas(a, b) {
    const primeiroA = Number(numerosDaPilha(a)[0] || Number.MAX_SAFE_INTEGER);
    const primeiroB = Number(numerosDaPilha(b)[0] || Number.MAX_SAFE_INTEGER);
    return primeiroA - primeiroB || String(a).localeCompare(String(b), "pt-BR", { numeric: true });
  }

  function renderizarKPIsIndicadores() {
    const datas = new Set();
    const pilhas = new Set();
    let preenchidos = 0;
    let total = 0;
    estado.resumos.forEach((resumo) => {
      resumo.registros.forEach((registro) => {
        datas.add(iso(registro.data));
        if (registro.pilha) pilhas.add(normalizarPilha(registro.pilha));
      });
      GRUPOS.forEach((grupo) => MEDIDAS.forEach((medida) => {
        total += 1;
        if (resumo[grupo][medida] != null) preenchidos += 1;
      }));
    });
    const elementosDisponiveis = [...estado.resumos.values()].filter((resumo) => GRUPOS.some((grupo) => MEDIDAS.some((medida) => resumo[grupo][medida] != null))).length;
    texto("kpiRegistrosQualidade", formatarNumero(datas.size, 0));
    texto("kpiPilhasIndicadores", formatarNumero(pilhas.size, 0));
    texto("kpiElementosQualidade", `${elementosDisponiveis}/9`);
    texto("kpiCoberturaQualidade", `${total ? formatarNumero(preenchidos / total * 100, 0) : "0"}%`);
  }

  function renderizarCardsElementos() {
    const container = $("cardsElementosQualidade");
    container.innerHTML = ELEMENTOS.map((elemento) => {
      const resumo = estado.resumos.get(elemento.id) || resumoVazio(elemento);
      return `
        <article class="card-elemento-qualidade">
          <div class="cabecalho-card-elemento"><h3>${elemento.nome}</h3><span>${elemento.percentual ? "%" : "relação"}</span></div>
          <div class="conteudo-card-elemento">
            <table class="mini-tabela-elemento">
              <thead><tr><th>Origem</th><th>AMG</th><th>CF</th><th>RF</th></tr></thead>
              <tbody>
                ${linhaMiniTabela("Usina 47", resumo.u47, elemento)}
                ${linhaMiniTabela("Usina 76", resumo.u76, elemento)}
                ${linhaMiniTabela("Consolidado", resumo.consolidado, elemento)}
              </tbody>
            </table>
            <div class="mini-grafico-elemento"><canvas id="miniGrafico-${elemento.id}" aria-label="AMG e CF de ${elemento.nome}"></canvas></div>
          </div>
          <div class="rodape-card-elemento"><button type="button" class="btn-evolucao-elemento" data-evolucao="${elemento.id}">Ver evolução →</button></div>
        </article>`;
    }).join("");
    renderizarMiniGraficos();
  }

  function resumoVazio(elemento) {
    const vazio = { amg: null, cf: null, rf: null };
    return { elemento, registros: [], u47: { ...vazio }, u76: { ...vazio }, consolidado: { ...vazio } };
  }

  function linhaMiniTabela(rotulo, dados, elemento) {
    return `<tr><td>${rotulo}</td><td>${formatarTeor(dados.amg, elemento)}</td><td>${formatarTeor(dados.cf, elemento)}</td><td>${formatarTeor(dados.rf, elemento)}</td></tr>`;
  }

  function renderizarMiniGraficos() {
    estado.graficosMini.forEach((grafico) => grafico.destroy());
    estado.graficosMini.clear();
    if (typeof Chart === "undefined") return;
    ELEMENTOS.forEach((elemento) => {
      const canvas = $(`miniGrafico-${elemento.id}`);
      if (!canvas) return;
      const resumo = estado.resumos.get(elemento.id) || resumoVazio(elemento);
      const valores = [resumo.u47.amg, resumo.u47.cf, resumo.u76.amg, resumo.u76.cf];
      const grafico = new Chart(canvas, {
        type: "bar",
        data: {
          labels: ["AMG 47", "CF 47", "AMG 76", "CF 76"],
          datasets: [{
            label: "AMG × CF",
            data: valores,
            backgroundColor: [CORES.roxoEscuro, CORES.roxoClaro, CORES.verdeEscuro, CORES.verdeClaro],
            borderRadius: 4,
            maxBarThickness: 24
          }]
        },
        options: opcoesMiniGrafico(elemento)
      });
      estado.graficosMini.set(elemento.id, grafico);
    });
  }

  function opcoesMiniGrafico(elemento) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 250 },
      plugins: {
        legend: { display: false },
        title: { display: true, text: "AMG × CF", color: CORES.roxoEscuro, font: { size: 10, weight: "700" }, padding: { bottom: 4 } },
        tooltip: { callbacks: { label: (item) => `${item.label}: ${formatarTeor(item.parsed.y, elemento)}` } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 7, weight: "700" }, maxRotation: 0, minRotation: 0 } },
        y: {
          beginAtZero: true,
          grid: { color: "rgba(45,1,77,.07)" },
          ticks: { maxTicksLimit: 4, font: { size: 7 }, callback: (valor) => elemento.percentual ? `${formatarNumero(valor, 0)}%` : formatarNumero(valor, 2) }
        }
      }
    };
  }

  function renderizarTabelasResumo() {
    $("corpoTabelaUsina47").innerHTML = linhasTabelaGrupo("u47");
    $("corpoTabelaUsina76").innerHTML = linhasTabelaGrupo("u76");
    $("corpoTabelaComparacaoQualidade").innerHTML = ELEMENTOS.map((elemento) => {
      const resumo = estado.resumos.get(elemento.id) || resumoVazio(elemento);
      return `<tr><td>${elemento.nome}</td><td>${formatarTeor(resumo.u47.amg, elemento)}</td><td>${formatarTeor(resumo.u47.cf, elemento)}</td><td>${formatarTeor(resumo.u76.amg, elemento)}</td><td>${formatarTeor(resumo.u76.cf, elemento)}</td><td>${formatarTeor(resumo.consolidado.amg, elemento)}</td><td>${formatarTeor(resumo.consolidado.cf, elemento)}</td><td>${formatarTeor(resumo.consolidado.rf, elemento)}</td></tr>`;
    }).join("");
    $("corpoTabelaTeoresQualidade").innerHTML = linhasTabelaMatriz();
    $("corpoTabelaRastreabilidade").innerHTML = linhasTabelaMatriz();
  }

  function linhasTabelaGrupo(grupo) {
    return ELEMENTOS.map((elemento) => {
      const dados = (estado.resumos.get(elemento.id) || resumoVazio(elemento))[grupo];
      return `<tr><td>${elemento.nome}</td><td>${formatarTeor(dados.amg, elemento)}</td><td>${formatarTeor(dados.cf, elemento)}</td><td>${formatarTeor(dados.rf, elemento)}</td></tr>`;
    }).join("");
  }

  function linhasTabelaMatriz() {
    return ELEMENTOS.map((elemento) => {
      const resumo = estado.resumos.get(elemento.id) || resumoVazio(elemento);
      return `<tr><td>${elemento.nome}</td>${GRUPOS.map((grupo) => MEDIDAS.map((medida) => `<td>${formatarTeor(resumo[grupo][medida], elemento)}</td>`).join("")).join("")}</tr>`;
    }).join("");
  }

  function renderizarPilhas() {
    const pilhas = pilhasFiltradas();
    const massas = pilhas.map((pilha) => pilha.massa).filter((valor) => valor != null && Number.isFinite(valor));
    const metragens = pilhas.map((pilha) => pilha.metragem).filter((valor) => valor != null && Number.isFinite(valor));
    const massaTotal = massas.reduce((total, valor) => total + valor, 0);
    texto("kpiTotalPilhas", formatarNumero(pilhas.length, 0));
    texto("kpiMassaTotalPilhas", massas.length ? formatarMassa(massaTotal) : "N/D");
    texto("kpiMassaMediaPilhas", massas.length ? formatarMassa(massaTotal / massas.length) : "N/D");
    texto("kpiMetragemMediaPilhas", metragens.length ? formatarMetragem(metragens.reduce((total, valor) => total + valor, 0) / metragens.length) : "N/D");
    $("corpoTabelaPilhasQualidade").innerHTML = pilhas.length ? pilhas.map((pilha) => linhaTabelaPilha(pilha)).join("") : '<tr><td colspan="13">Carregue a planilha Análise de Pilha ou ajuste o período selecionado.</td></tr>';
    renderizarGraficoMassa(pilhas);
    renderizarResumoPilhaDestaque();
  }

  function linhaTabelaPilha(pilha) {
    const t = pilha.teores;
    const valor = (id) => formatarTeor(t[id], ELEMENTOS.find((item) => item.id === id));
    return `<tr><td>${escaparHTML(pilha.pilha)}</td><td>${formatarMassa(pilha.massa)}</td><td>${formatarMetragem(pilha.metragem)}</td><td>${formatarData(pilha.inicio)}</td><td>${formatarData(pilha.fim)}</td><td>${valor("p2o5")}</td><td>${valor("mgo")}</td><td>${valor("fe2o3")}</td><td>${valor("sio2")}</td><td>${valor("bao")}</td><td>${valor("nb2o5")}</td><td>${valor("cao")}</td><td>${valor("caop2o5")}</td></tr>`;
  }

  function renderizarGraficoMassa(pilhas) {
    estado.graficoMassa?.destroy();
    estado.graficoMassa = null;
    const canvas = $("graficoMassaPilhasQualidade");
    if (!canvas || typeof Chart === "undefined") return;
    const validas = pilhas.filter((pilha) => pilha.massa != null).slice(-40);
    estado.graficoMassa = new Chart(canvas, {
      type: "bar",
      data: { labels: validas.map((pilha) => pilha.pilha), datasets: [{ label: "Massa", data: validas.map((pilha) => pilha.massa), backgroundColor: CORES.roxoEscuro, borderRadius: 5, maxBarThickness: 42 }] },
      options: {
        responsive: true, maintainAspectRatio: false,
        onClick: (_evento, pontos) => {
          if (!pontos.length) return;
          selecionarPilha(validas[pontos[0].index].pilha);
          ativarAba("rastreabilidade");
        },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: (item) => `Massa: ${formatarMassa(item.parsed.y)} t` } } },
        scales: { x: { grid: { display: false }, ticks: { font: { size: 9, weight: "700" } } }, y: { beginAtZero: true, grid: { color: "rgba(45,1,77,.08)" }, ticks: { callback: (valor) => formatarNumero(valor, 0) } } }
      }
    });
  }

  function selecionarPilha(nome) {
    const seletor = $("filtroPilhaQualidade");
    let opcao = [...seletor.options].find((item) => normalizarPilha(item.value) === normalizarPilha(nome));
    if (!opcao) {
      opcao = new Option(nome, nome);
      seletor.add(opcao);
    }
    seletor.value = opcao.value;
    atualizarDashboard(false);
  }

  function encontrarPilhaSelecionada() {
    const selecionada = $("filtroPilhaQualidade").value;
    if (!selecionada) return null;
    const candidatas = estado.pilhasArquivo.filter((pilha) => pilhasRelacionadas(pilha.pilha, selecionada));
    candidatas.sort((a, b) => (b.fim || b.inicio || new Date(0)) - (a.fim || a.inicio || new Date(0)));
    return candidatas[0] || null;
  }

  function renderizarResumoPilhaDestaque() {
    const selecionada = $("filtroPilhaQualidade").value;
    const pilha = encontrarPilhaSelecionada();
    if (!selecionada) {
      $("resumoPilhaDestaque").innerHTML = "Selecione uma pilha no filtro global para visualizar os detalhes.";
      return;
    }
    $("resumoPilhaDestaque").innerHTML = pilha ? `<dl><div><dt>Pilha</dt><dd>${escaparHTML(pilha.pilha)}</dd></div><div><dt>Massa</dt><dd>${formatarMassa(pilha.massa)} t</dd></div><div><dt>Metragem</dt><dd>${formatarMetragem(pilha.metragem)} m</dd></div><div><dt>Início</dt><dd>${formatarData(pilha.inicio)}</dd></div><div><dt>Fim</dt><dd>${formatarData(pilha.fim)}</dd></div></dl>` : `A pilha ${escaparHTML(selecionada)} está nos indicadores, mas não foi localizada na Análise de Pilha.`;
  }

  function renderizarRastreabilidade() {
    const selecionada = $("filtroPilhaQualidade").value;
    const pilha = encontrarPilhaSelecionada();
    const registrosRelacionados = selecionada ? [...estado.resumos.values()].reduce((total, resumo) => total + resumo.registros.length, 0) : 0;
    texto("rastroPilhaNumero", selecionada || "—");
    texto("rastroPilhaMassa", pilha ? formatarMassa(pilha.massa) : "N/D");
    texto("rastroPilhaMetragem", pilha ? formatarMetragem(pilha.metragem) : "N/D");
    texto("rastroPilhaPeriodo", pilha ? `${formatarData(pilha.inicio)} a ${formatarData(pilha.fim)}` : "N/D");
    const status = $("statusRastreabilidadeQualidade");
    if (!selecionada) {
      status.textContent = "Selecione uma pilha no filtro global.";
      status.className = "status-rastreabilidade-qualidade sem-dados";
    } else if (registrosRelacionados) {
      status.textContent = `Dados relacionados encontrados para a pilha ${selecionada}. Os valores exibidos são os acumulados por pilha já existentes no Indicadores 2026.`;
      status.className = "status-rastreabilidade-qualidade encontrado";
    } else {
      status.textContent = "Sem dados suficientes para correlação.";
      status.className = "status-rastreabilidade-qualidade sem-dados";
    }
  }

  function renderizarDados() {
    const linhas = [];
    ELEMENTOS.forEach((elemento) => {
      registrosFiltrados(elemento.id, true).forEach((registro) => linhas.push({ elemento, registro }));
    });
    linhas.sort((a, b) => a.registro.data - b.registro.data || a.elemento.nome.localeCompare(b.elemento.nome));
    texto("contadorLinhasDadosQualidade", `${linhas.length} linha${linhas.length === 1 ? "" : "s"}`);
    const limite = 600;
    $("corpoTabelaDadosQualidade").innerHTML = linhas.length ? linhas.slice(0, limite).map(({ elemento, registro }) => {
      const celulas = GRUPOS.flatMap((grupo) => MEDIDAS.map((medida) => formatarTeor(registro.diario[grupo][medida], elemento)));
      return `<tr><td>${formatarData(registro.data)}</td><td>${escaparHTML(registro.pilha || "N/D")}</td><td>${elemento.nome}</td>${celulas.map((valor) => `<td>${valor}</td>`).join("")}</tr>`;
    }).join("") : '<tr><td colspan="12">Nenhum dado encontrado no período.</td></tr>';
  }

  function atualizarTextosGerais(periodo) {
    const selecionada = $("filtroPilhaQualidade").value;
    const rotulo = `Período analisado: ${formatarData(periodo.inicio)} — ${formatarData(periodo.fim)}${selecionada ? ` · Pilha ${selecionada}` : ""}`;
    texto("periodoAnalisadoQualidade", rotulo);
    texto("rodapePeriodoQualidade", rotulo);
    texto("rodapeArquivosQualidade", origemArquivos());
  }

  function origemArquivos() {
    const itens = [];
    if (estado.arquivoPilhas) itens.push(`Pilhas: ${estado.arquivoPilhas}`);
    if (estado.arquivoIndicadores) itens.push(`Indicadores: ${estado.arquivoIndicadores}`);
    return itens.length ? itens.join(" · ") : "Aguardando planilhas";
  }

  function ativarAba(nome) {
    estado.painelAtivo = nome;
    document.querySelectorAll(".aba-qualidade").forEach((botao) => botao.classList.toggle("ativa", botao.dataset.painel === nome));
    document.querySelectorAll(".painel-aba-qualidade").forEach((painel) => {
      const ativo = painel.id === `painel-${nome}`;
      painel.hidden = !ativo;
      painel.classList.toggle("ativo", ativo);
    });
    if (nome === "pilhas") setTimeout(() => estado.graficoMassa?.resize(), 30);
    if (nome === "visaoGeral") setTimeout(() => estado.graficosMini.forEach((grafico) => grafico.resize()), 30);
  }

  function abrirModalEvolucao(elementoId) {
    const elemento = ELEMENTOS.find((item) => item.id === elementoId);
    if (!elemento) return;
    const registros = registrosFiltrados(elementoId, true).sort((a, b) => a.data - b.data);
    texto("tituloModalEvolucao", elemento.nome);
    texto("subtituloModalEvolucao", `${formatarData(dataDoInput("dataInicialQualidade"))} a ${formatarData(dataDoInput("dataFinalQualidade"))} · AMG, CF e RF`);
    $("modalEvolucaoQualidade").classList.remove("oculto");
    document.body.classList.add("modal-aberto");
    estado.graficosEvolucao.forEach((grafico) => grafico.destroy());
    estado.graficosEvolucao = [];
    if (typeof Chart === "undefined") return;
    estado.graficosEvolucao.push(
      criarGraficoEvolucao($("graficoEvolucaoU47"), registros, "u47", elemento),
      criarGraficoEvolucao($("graficoEvolucaoU76"), registros, "u76", elemento),
      criarGraficoEvolucao($("graficoEvolucaoConsolidado"), registros, "consolidado", elemento)
    );
  }

  function criarGraficoEvolucao(canvas, registros, grupo, elemento) {
    return new Chart(canvas, {
      type: "line",
      data: {
        labels: registros.map((registro) => formatarDataCurta(registro.data)),
        datasets: [
          { label: "AMG", data: registros.map((registro) => registro.diario[grupo].amg), borderColor: CORES.roxoEscuro, backgroundColor: CORES.roxoEscuro, borderWidth: 2.4, pointRadius: 2, tension: .2 },
          { label: "CF", data: registros.map((registro) => registro.diario[grupo].cf), borderColor: CORES.verdeEscuro, backgroundColor: CORES.verdeEscuro, borderWidth: 2.4, pointRadius: 2, tension: .2 },
          { label: "RF", data: registros.map((registro) => registro.diario[grupo].rf), borderColor: CORES.laranja, backgroundColor: CORES.laranja, borderWidth: 2.2, pointRadius: 2, tension: .2 }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false, interaction: { mode: "index", intersect: false },
        plugins: { legend: { position: "top", align: "start" }, tooltip: { callbacks: { label: (item) => `${item.dataset.label}: ${formatarTeor(item.parsed.y, elemento)}` } } },
        scales: { x: { grid: { display: false }, ticks: { maxRotation: 45, minRotation: 0, font: { size: 9 } } }, y: { beginAtZero: true, grid: { color: "rgba(45,1,77,.08)" }, ticks: { callback: (valor) => elemento.percentual ? `${formatarNumero(valor, 1)}%` : formatarNumero(valor, 2) } } }
      }
    });
  }

  function fecharModalEvolucao() {
    const modal = $("modalEvolucaoQualidade");
    if (!modal || modal.classList.contains("oculto")) return;
    modal.classList.add("oculto");
    document.body.classList.remove("modal-aberto");
    estado.graficosEvolucao.forEach((grafico) => grafico.destroy());
    estado.graficosEvolucao = [];
  }

  async function gerarPDF() {
    if (!estado.indicadores.size) return aviso("Carregue a planilha Indicadores 2026 antes de gerar o PDF.", "erro");
    if (!window.jspdf || typeof Chart === "undefined") return aviso("As bibliotecas necessárias para o PDF não foram carregadas.", "erro");
    const periodo = validarPeriodo(true);
    if (!periodo) return;
    const abaAnterior = estado.painelAtivo;
    exibirCarregamento("Gerando relatório — página 1 de 2...");
    try {
      ativarAba("visaoGeral");
      await aguardar(450);
      estado.graficosMini.forEach((grafico) => grafico.update("none"));
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4", compress: true });
      const pilhas = pilhasFiltradas();
      desenharPagina1PDF(pdf, periodo, pilhas);
      exibirCarregamento("Gerando relatório — página 2 de 2...");
      pdf.addPage("a4", "landscape");
      desenharPagina2PDF(pdf, periodo);
      const nome = `Relatorio_Qualidade_${iso(periodo.inicio)}_a_${iso(periodo.fim)}.pdf`;
      pdf.save(nome);
      aviso("Relatório gerado com sucesso em duas páginas A4.", "sucesso");
    } catch (erro) {
      console.error(erro);
      aviso(`Erro ao gerar o PDF: ${erro.message}`, "erro");
    } finally {
      ativarAba(abaAnterior);
      ocultarCarregamento();
    }
  }

  function cabecalhoPDF(pdf, titulo, pagina, periodo) {
    const largura = pdf.internal.pageSize.getWidth();
    pdf.setFillColor(45, 1, 77);
    pdf.rect(0, 0, largura, 22, "F");
    pdf.setTextColor(255, 255, 255);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(15);
    pdf.text(titulo, 9, 9);
    pdf.setFontSize(8);
    pdf.setFont("helvetica", "normal");
    pdf.text(`${formatarData(periodo.inicio)} a ${formatarData(periodo.fim)}${$("filtroPilhaQualidade").value ? ` · Pilha ${$("filtroPilhaQualidade").value}` : ""}`, 9, 16);
    pdf.text(`Página ${pagina}/2`, largura - 9, 9, { align: "right" });
    pdf.text(new Date().toLocaleString("pt-BR"), largura - 9, 16, { align: "right" });
    pdf.setTextColor(45, 1, 77);
  }

  function desenharPagina1PDF(pdf, periodo, pilhas) {
    cabecalhoPDF(pdf, "RELATÓRIO DE QUALIDADE E RASTREABILIDADE", 1, periodo);
    const massas = pilhas.map((pilha) => pilha.massa).filter((valor) => valor != null);
    const metragens = pilhas.map((pilha) => pilha.metragem).filter((valor) => valor != null);
    const massaTotal = massas.reduce((total, valor) => total + valor, 0);
    const indicadores = [
      ["Total de pilhas", String(pilhas.length), "no período"],
      ["Massa total", massas.length ? `${formatarMassa(massaTotal)} t` : "N/D", "Análise de Pilha"],
      ["Massa média", massas.length ? `${formatarMassa(massaTotal / massas.length)} t` : "N/D", "por pilha"],
      ["Metragem média", metragens.length ? `${formatarMetragem(metragens.reduce((a, b) => a + b, 0) / metragens.length)} m` : "N/D", "no período"]
    ];
    indicadores.forEach((item, indice) => desenharKpiPDF(pdf, 9 + indice * 70, 28, 65, 24, item[0], item[1], item[2]));

    pdf.setDrawColor(220, 214, 224);
    pdf.setFillColor(255, 255, 255);
    pdf.roundedRect(9, 57, 182, 73, 2, 2, "FD");
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(10);
    pdf.setTextColor(45, 1, 77);
    pdf.text("Massa das pilhas no período", 14, 65);
    desenharBarrasMassaPDF(pdf, pilhas.filter((pilha) => pilha.massa != null).slice(-18), 15, 70, 170, 52);

    // O gráfico altera a cor de preenchimento para roxo. Antes de desenhar o
    // quadro seguinte é obrigatório restaurar o fundo branco; sem isso, os
    // títulos e valores roxos ficam invisíveis no PDF.
    pdf.setFillColor(255, 255, 255);
    pdf.setDrawColor(220, 214, 224);
    pdf.roundedRect(196, 57, 92, 73, 2, 2, "FD");
    pdf.setTextColor(45, 1, 77);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(10);
    pdf.text("Pilha em destaque", 201, 65);
    const pilha = encontrarPilhaSelecionada();
    const selecionada = $("filtroPilhaQualidade").value || "Todas";
    const detalhes = [
      ["Pilha", pilha?.pilha || selecionada],
      ["Massa", pilha ? `${formatarMassa(pilha.massa)} t` : "N/D"],
      ["Metragem", pilha ? `${formatarMetragem(pilha.metragem)} m` : "N/D"],
      ["Início", pilha ? formatarData(pilha.inicio) : "N/D"],
      ["Fim", pilha ? formatarData(pilha.fim) : "N/D"]
    ];
    detalhes.forEach((item, indice) => {
      const y = 75 + indice * 10;
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(7.5); pdf.setTextColor(81, 74, 85); pdf.text(item[0], 201, y);
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(8); pdf.setTextColor(45, 1, 77); pdf.text(String(item[1]), 283, y, { align: "right" });
      pdf.setDrawColor(238, 234, 240); pdf.line(201, y + 2, 283, y + 2);
    });

    pdf.setFont("helvetica", "bold"); pdf.setFontSize(10); pdf.setTextColor(45, 1, 77);
    pdf.text("Resumo das pilhas", 9, 140);
    desenharTabelaPilhasPDF(pdf, pilhas.slice(0, 8), 9, 145);
    rodapePDF(pdf);
  }

  function desenharKpiPDF(pdf, x, y, w, h, rotulo, valor, detalhe) {
    pdf.setFillColor(255, 255, 255); pdf.setDrawColor(221, 215, 225); pdf.roundedRect(x, y, w, h, 2, 2, "FD");
    pdf.setFillColor(146, 208, 80); pdf.rect(x, y, w, 1.4, "F");
    pdf.setFont("helvetica", "bold"); pdf.setFontSize(7.5); pdf.setTextColor(81, 74, 85); pdf.text(rotulo.toUpperCase(), x + 4, y + 7);
    pdf.setFontSize(14); pdf.setTextColor(45, 1, 77); pdf.text(valor, x + 4, y + 16);
    pdf.setFont("helvetica", "normal"); pdf.setFontSize(7); pdf.setTextColor(100, 93, 104); pdf.text(detalhe, x + 4, y + 21);
  }

  function desenharBarrasMassaPDF(pdf, pilhas, x, y, w, h) {
    if (!pilhas.length) {
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(8); pdf.setTextColor(118, 111, 121); pdf.text("Sem dados de massa no período.", x + w / 2, y + h / 2, { align: "center" });
      return;
    }
    const maximo = Math.max(...pilhas.map((pilha) => pilha.massa), 1);
    const espaco = w / pilhas.length;
    pilhas.forEach((pilha, indice) => {
      const altura = pilha.massa / maximo * (h - 13);
      const larguraBarra = Math.max(2, espaco * .58);
      const bx = x + indice * espaco + (espaco - larguraBarra) / 2;
      const by = y + h - 8 - altura;
      pdf.setFillColor(61, 7, 100); pdf.roundedRect(bx, by, larguraBarra, altura, 1, 1, "F");
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(6.2); pdf.setTextColor(81, 74, 85); pdf.text(String(pilha.pilha).slice(0, 8), bx + larguraBarra / 2, y + h - 3, { align: "center" });
      if (pilhas.length <= 12) {
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(5.8);
        pdf.setTextColor(45, 1, 77);
        pdf.text(`${formatarMassa(pilha.massa)} t`, bx + larguraBarra / 2, Math.max(y + 6, by - 2), { align: "center" });
      }
    });
    pdf.setDrawColor(225, 220, 228); pdf.line(x, y + h - 8, x + w, y + h - 8);
  }

  function desenharTabelaPilhasPDF(pdf, pilhas, x, y) {
    const colunas = [44, 38, 34, 44, 44, 34, 34];
    const cabecalhos = ["Pilha", "Massa (t)", "Metragem", "Início", "Fim", "P2O5", "MgO"];
    let cursor = x;
    cabecalhos.forEach((cabecalho, indice) => {
      pdf.setFillColor(238, 231, 242); pdf.setDrawColor(221, 215, 225); pdf.rect(cursor, y, colunas[indice], 8, "FD");
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(7); pdf.setTextColor(45, 1, 77); pdf.text(cabecalho, cursor + 2, y + 5.2);
      cursor += colunas[indice];
    });
    if (!pilhas.length) {
      pdf.setFont("helvetica", "normal"); pdf.setTextColor(118, 111, 121); pdf.text("Sem dados de pilhas.", x + 2, y + 14);
      return;
    }
    pilhas.forEach((pilha, linha) => {
      const valores = [pilha.pilha, formatarMassa(pilha.massa), formatarMetragem(pilha.metragem), formatarData(pilha.inicio), formatarData(pilha.fim), formatarTeor(pilha.teores.p2o5, ELEMENTOS[0]), formatarTeor(pilha.teores.mgo, ELEMENTOS[4])];
      cursor = x;
      valores.forEach((valor, indice) => {
        pdf.setFillColor(linha % 2 ? 250 : 255, linha % 2 ? 249 : 255, linha % 2 ? 251 : 255); pdf.setDrawColor(236, 232, 238); pdf.rect(cursor, y + 8 + linha * 7, colunas[indice], 7, "FD");
        pdf.setFont("helvetica", indice === 0 ? "bold" : "normal"); pdf.setFontSize(6.5); pdf.setTextColor(63, 55, 68); pdf.text(String(valor).slice(0, 18), cursor + 2, y + 12.7 + linha * 7);
        cursor += colunas[indice];
      });
    });
  }

  function desenharPagina2PDF(pdf, periodo) {
    cabecalhoPDF(pdf, "ANÁLISE DE QUALIDADE - AMG, CF E RF", 2, periodo);
    const legenda = [[CORES.roxoEscuro, "AMG U47"], [CORES.roxoClaro, "CF U47"], [CORES.verdeEscuro, "AMG U76"], [CORES.verdeClaro, "CF U76"]];
    let lx = 9;
    legenda.forEach(([cor, nome]) => {
      const rgb = hexParaRgb(cor); pdf.setFillColor(...rgb); pdf.rect(lx, 26, 4, 4, "F"); pdf.setFont("helvetica", "bold"); pdf.setFontSize(7); pdf.setTextColor(63, 55, 68); pdf.text(nome, lx + 6, 29.2); lx += 33;
    });
    const margem = 9;
    const gapX = 4;
    const gapY = 3;
    const larguraCard = (297 - margem * 2 - gapX * 2) / 3;
    const alturaCard = 52;
    ELEMENTOS.forEach((elemento, indice) => {
      const coluna = indice % 3;
      const linha = Math.floor(indice / 3);
      desenharCardElementoPDF(pdf, elemento, margem + coluna * (larguraCard + gapX), 33 + linha * (alturaCard + gapY), larguraCard, alturaCard);
    });
    rodapePDF(pdf);
  }

  function desenharCardElementoPDF(pdf, elemento, x, y, w, h) {
    const resumo = estado.resumos.get(elemento.id) || resumoVazio(elemento);
    pdf.setFillColor(255, 255, 255); pdf.setDrawColor(220, 214, 224); pdf.roundedRect(x, y, w, h, 2, 2, "FD");
    pdf.setFillColor(146, 208, 80); pdf.rect(x, y, w, 1.2, "F");
    pdf.setFont("helvetica", "bold"); pdf.setFontSize(10.5); pdf.setTextColor(45, 1, 77); pdf.text(elemento.nome, x + 3, y + 8);
    const tabelaX = x + 3;
    const tabelaY = y + 15;
    pdf.setFont("helvetica", "bold"); pdf.setFontSize(6.3); pdf.setTextColor(63, 55, 68);
    pdf.text("Origem", tabelaX, tabelaY); pdf.text("AMG", tabelaX + 23, tabelaY, { align: "right" }); pdf.text("CF", tabelaX + 35, tabelaY, { align: "right" }); pdf.text("RF", tabelaX + 47, tabelaY, { align: "right" });
    [["U47", resumo.u47], ["U76", resumo.u76], ["Cons.", resumo.consolidado]].forEach(([nome, dados], indice) => {
      const yy = tabelaY + 7 + indice * 8.5;
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(6.1); pdf.setTextColor(indice === 2 ? 45 : 63, indice === 2 ? 1 : 55, indice === 2 ? 77 : 68); pdf.text(nome, tabelaX, yy);
      pdf.text(formatarTeor(dados.amg, elemento), tabelaX + 23, yy, { align: "right" }); pdf.text(formatarTeor(dados.cf, elemento), tabelaX + 35, yy, { align: "right" }); pdf.text(formatarTeor(dados.rf, elemento), tabelaX + 47, yy, { align: "right" });
      pdf.setDrawColor(239, 235, 241); pdf.line(tabelaX, yy + 2, tabelaX + 48, yy + 2);
    });
    pdf.setFont("helvetica", "bold"); pdf.setFontSize(6.2); pdf.setTextColor(81, 74, 85); pdf.text("AMG x CF", x + 57, y + 15);
    desenharMiniBarrasPDF(pdf, [resumo.u47.amg, resumo.u47.cf, resumo.u76.amg, resumo.u76.cf], x + 55, y + 17, w - 59, h - 20, elemento);
  }

  function desenharMiniBarrasPDF(pdf, valores, x, y, w, h, elemento) {
    const validos = valores.filter((valor) => valor != null && Number.isFinite(valor));
    if (!validos.length) {
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(5.5); pdf.setTextColor(118, 111, 121); pdf.text("Sem dados", x + w / 2, y + h / 2, { align: "center" });
      return;
    }
    const maximo = Math.max(...validos, 0.0001);
    const cores = [CORES.roxoEscuro, CORES.roxoClaro, CORES.verdeEscuro, CORES.verdeClaro];
    const largura = Math.max(2.5, Math.min(7, w / 6));
    const gap = (w - largura * 4) / 5;
    valores.forEach((valor, indice) => {
      const altura = valor == null ? 0 : valor / maximo * (h - 10);
      const bx = x + gap + indice * (largura + gap);
      const by = y + h - 6 - altura;
      const rgb = hexParaRgb(cores[indice]); pdf.setFillColor(...rgb); if (altura > 0) pdf.roundedRect(bx, by, largura, altura, .7, .7, "F");
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(5.7); pdf.setTextColor(63, 55, 68); pdf.text(["A47", "C47", "A76", "C76"][indice], bx + largura / 2, y + h - 1.5, { align: "center" });
      if (valor != null) {
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(5.6);
        pdf.setTextColor(45, 1, 77);
        pdf.text(formatarTeor(valor, elemento), bx + largura / 2, Math.max(y + 5, by - 1.2), { align: "center" });
      }
    });
  }

  function hexParaRgb(hex) {
    const valor = hex.replace("#", "");
    return [parseInt(valor.slice(0, 2), 16), parseInt(valor.slice(2, 4), 16), parseInt(valor.slice(4, 6), 16)];
  }

  function rodapePDF(pdf) {
    const largura = pdf.internal.pageSize.getWidth();
    const altura = pdf.internal.pageSize.getHeight();
    pdf.setDrawColor(221, 215, 225); pdf.line(9, altura - 7, largura - 9, altura - 7);
    pdf.setFont("helvetica", "normal"); pdf.setFontSize(6); pdf.setTextColor(100, 93, 104);
    pdf.text("Fonte: planilhas carregadas no Portal FVO. Valores N/D correspondem a células sem informação.", 9, altura - 3.2);
    pdf.text(origemArquivos().slice(0, 115), largura - 9, altura - 3.2, { align: "right" });
  }

  function renderizarEstadoInicial() {
    const periodo = { inicio: new Date(), fim: new Date() };
    estado.resumos = new Map(ELEMENTOS.map((elemento) => [elemento.id, resumoVazio(elemento)]));
    renderizarCardsElementos();
    renderizarTabelasResumo();
    renderizarPilhas(periodo);
    renderizarRastreabilidade();
    renderizarDados();
  }

  function aguardar(ms) {
    return new Promise((resolver) => setTimeout(resolver, ms));
  }

  if (window.__QUALIDADE_FVO_TESTE__) {
    window.__QualidadeFVOTeste = {
      ELEMENTOS,
      COLUNAS_INDICADORES,
      estado,
      normalizar,
      numero,
      dataExcel,
      mapearIndicadores,
      mapearAnalisePilha,
      resumirElemento,
      pilhasRelacionadas,
      atualizarDashboard,
      gerarPDF
    };
  }
})();