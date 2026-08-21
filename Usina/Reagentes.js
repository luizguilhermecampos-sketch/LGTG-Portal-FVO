(function iniciarDashboardReagentes() {
  "use strict";

  const ABA_REAGENTES = "Agosto 2026 (2)";
  const CORES = {
    roxo: "#3d0764",
    verde: "#2f9e73",
    amarelo: "#d49516",
    vermelho: "#c93c31",
    cinza: "#766f79"
  };

  const REAGENTES = [
    {
      id: "soda", nome: "Soda", tipo: "direto", modelo: "estoque",
      estoque: "D", carga: "E", erro: "F", diario: "G", meta: "H", turno: "B"
    },
    {
      id: "lupromin715", nome: "Lupromin 715", tipo: "GRD", modelo: "rateado",
      estoque: "P", carga: "Q", erro: "R", tanque: "S", direcionador: "U", diario: "W", meta: "X", turno: "N"
    },
    {
      id: "flotanol", nome: "Flotanol", tipo: "direto", modelo: "estoque",
      estoque: "AF", carga: "AG", erro: "AH", diario: "AI", meta: "AJ", turno: "AD"
    },
    {
      id: "lupromin1168", nome: "Lupromin 1168", tipo: "direto", modelo: "estoque",
      estoque: "AR", carga: "AS", erro: "AT", diario: "AU", meta: "AV", turno: "AP"
    },
    {
      id: "agem", nome: "Agem", tipo: "GRD", modelo: "rateado",
      estoque: "BD", carga: "BF", erro: "BE", tanque: "BG", direcionador: "BI", diario: "BK", meta: "BL", turno: "BB"
    }
  ];

  const estado = {
    reagentes: new Map(),
    indicadores: null,
    pilhas: [],
    arquivoReagentes: "",
    arquivoIndicadores: "",
    graficos: { diario: null, aderencia: null, pilha: null },
    diasSemIdentificacao: 0
  };

  const $ = (id) => document.getElementById(id);

  document.addEventListener("DOMContentLoaded", () => {
    $("arquivoReagentes")?.addEventListener("change", carregarPlanilhaReagentes);
    $("arquivoIndicadoresReagentes")?.addEventListener("change", carregarPlanilhaIndicadores);
    $("btnAtualizarReagentes")?.addEventListener("click", atualizarDashboard);
    $("filtroReagenteGrafico")?.addEventListener("change", atualizarGraficosPrincipais);
    $("filtroPilhaReagentes")?.addEventListener("change", atualizarPilhaSelecionada);
    $("btnGerarPdfReagentes")?.addEventListener("click", gerarPDF);
    $("btnSalvarTratativasReagentes")?.addEventListener("click", salvarTratativas);
    $("btnAbrirConsumoPilha")?.addEventListener("click", abrirConsumoPorPilha);
    $("justificativasReagentes")?.addEventListener("input", atualizarNotasPDF);
    $("acoesTratativasReagentes")?.addEventListener("input", atualizarNotasPDF);
  });

  function normalizar(texto) {
    return String(texto || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }

  function numero(valor) {
    if (typeof valor === "number") return Number.isFinite(valor) ? valor : null;
    if (valor == null || valor === "") return null;
    let texto = String(valor).trim().replace(/\s/g, "");
    if (!texto || texto.startsWith("#")) return null;
    if (texto.includes(",")) texto = texto.replace(/\./g, "").replace(",", ".");
    const convertido = Number(texto);
    return Number.isFinite(convertido) ? convertido : null;
  }

  function dataExcel(valor) {
    if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
      return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
    }
    if (typeof valor === "number" && window.XLSX?.SSF) {
      const partes = XLSX.SSF.parse_date_code(valor);
      return partes ? new Date(partes.y, partes.m - 1, partes.d) : null;
    }
    const texto = String(valor || "").trim();
    let partes = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (partes) return new Date(Number(partes[3]), Number(partes[2]) - 1, Number(partes[1]));
    partes = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (partes) return new Date(Number(partes[1]), Number(partes[2]) - 1, Number(partes[3]));
    return null;
  }

  function iso(data) {
    if (!(data instanceof Date)) return "";
    return [data.getFullYear(), String(data.getMonth() + 1).padStart(2, "0"), String(data.getDate()).padStart(2, "0")].join("-");
  }

  function lerDataFiltro(id) {
    const valor = $(id)?.value;
    if (!valor) return null;
    const partes = valor.split("-").map(Number);
    return new Date(partes[0], partes[1] - 1, partes[2]);
  }

  function celula(planilha, coluna, linha) {
    return planilha?.[`${coluna}${linha}`] || null;
  }

  function valorCelula(planilha, coluna, linha) {
    return celula(planilha, coluna, linha)?.v ?? null;
  }

  function numeroCelula(planilha, coluna, linha) {
    return numero(valorCelula(planilha, coluna, linha));
  }

  function valorManual(planilha, coluna, linha) {
    const item = celula(planilha, coluna, linha);
    if (!item || item.f) return null;
    return numero(item.v);
  }

  function linhasDoMes(planilha) {
    const linhas = [];
    let referencia = null;
    for (let linha = 7; linha <= 60; linha += 1) {
      const data = dataExcel(valorCelula(planilha, "A", linha));
      if (!data) continue;
      if (!referencia) referencia = { ano: data.getFullYear(), mes: data.getMonth() };
      if (data.getFullYear() !== referencia.ano || data.getMonth() !== referencia.mes) {
        if (linhas.length) break;
        continue;
      }
      linhas.push({ linha, data });
    }
    return linhas;
  }

  function consumoPorEstoque(planilha, definicao, linha) {
    const manual = valorManual(planilha, definicao.diario, linha);
    if (manual != null) return { valor: manual, valido: true };

    const estoqueAtual = numeroCelula(planilha, definicao.estoque, linha);
    const estoqueSeguinte = numeroCelula(planilha, definicao.estoque, linha + 1);
    if (estoqueAtual == null || estoqueSeguinte == null) return { valor: 0, valido: false };

    const carga = numeroCelula(planilha, definicao.carga, linha) || 0;
    const erro = numeroCelula(planilha, definicao.erro, linha) || 0;
    return { valor: estoqueAtual - estoqueSeguinte + carga + erro, valido: true };
  }

  function consumoTanque(planilha, definicao, linha) {
    const manual = valorManual(planilha, definicao.tanque, linha);
    if (manual != null) return { valor: manual, valido: true };

    const estoqueAtual = numeroCelula(planilha, definicao.estoque, linha);
    const estoqueSeguinte = numeroCelula(planilha, definicao.estoque, linha + 1);
    if (estoqueAtual == null || estoqueSeguinte == null) return { valor: 0, valido: false };

    const carga = numeroCelula(planilha, definicao.carga, linha) || 0;
    const erro = numeroCelula(planilha, definicao.erro, linha) || 0;
    return { valor: estoqueAtual - estoqueSeguinte + carga + erro, valido: true };
  }

  function mapearReagenteEstoque(planilha, definicao, linhas) {
    return linhas.map(({ linha, data }) => {
      const consumo = consumoPorEstoque(planilha, definicao, linha);
      return {
        data,
        valor: consumo.valor,
        valido: consumo.valido,
        meta: numeroCelula(planilha, definicao.meta, linha) || 0,
        turno: String(valorCelula(planilha, definicao.turno, linha) || "").trim()
      };
    });
  }

  function mapearReagenteRateado(planilha, definicao, linhas) {
    const consumosTanque = new Map();
    let totalTanque = 0;
    let totalDirecionador = 0;

    linhas.forEach(({ linha }) => {
      const tanque = consumoTanque(planilha, definicao, linha);
      const direcionador = numeroCelula(planilha, definicao.direcionador, linha);
      consumosTanque.set(linha, tanque);
      if (tanque.valido) totalTanque += tanque.valor;
      if (direcionador != null) totalDirecionador += direcionador;
    });

    return linhas.map(({ linha, data }) => {
      const finalManual = valorManual(planilha, definicao.diario, linha);
      const direcionador = numeroCelula(planilha, definicao.direcionador, linha);
      const valido = finalManual != null || (direcionador != null && totalDirecionador !== 0 && consumosTanque.get(linha));
      const valor = finalManual != null
        ? finalManual
        : valido ? totalTanque * direcionador / totalDirecionador : 0;

      return {
        data,
        valor,
        valido: Boolean(valido),
        meta: numeroCelula(planilha, definicao.meta, linha) || 0,
        turno: String(valorCelula(planilha, definicao.turno, linha) || "").trim()
      };
    });
  }

  function mapearPlanilhaReagentes(workbook) {
    const nomeAba = workbook.SheetNames.find((nome) => normalizar(nome) === normalizar(ABA_REAGENTES));
    if (!nomeAba || !workbook.Sheets[nomeAba]) {
      throw new Error(`A aba "${ABA_REAGENTES}" não foi encontrada.`);
    }

    const planilha = workbook.Sheets[nomeAba];
    const linhas = linhasDoMes(planilha);
    if (!linhas.length) throw new Error(`Nenhuma data válida foi encontrada na aba "${ABA_REAGENTES}".`);

    const resultado = new Map();
    REAGENTES.forEach((definicao) => {
      const dados = definicao.modelo === "rateado"
        ? mapearReagenteRateado(planilha, definicao, linhas)
        : mapearReagenteEstoque(planilha, definicao, linhas);
      resultado.set(definicao.id, dados);
    });
    return resultado;
  }

  async function lerWorkbook(arquivo, abas) {
    if (typeof XLSX === "undefined") throw new Error("A biblioteca de leitura do Excel não foi carregada.");
    const buffer = await arquivo.arrayBuffer();
    return XLSX.read(new Uint8Array(buffer), {
      type: "array",
      cellDates: true,
      cellFormula: true,
      cellStyles: false,
      dense: false,
      sheets: abas
    });
  }

  async function carregarPlanilhaReagentes(evento) {
    const arquivo = evento.target.files?.[0];
    if (!arquivo) return;
    exibirCarregamento(`Lendo ${arquivo.name}...`);
    try {
      const workbook = await lerWorkbook(arquivo, [ABA_REAGENTES]);
      estado.reagentes = mapearPlanilhaReagentes(workbook);
      estado.arquivoReagentes = arquivo.name;
      texto("nomeArquivoReagentes", arquivo.name);
      configurarPeriodoInicial();
      atualizarDashboard();
      aviso(`Planilha de reagentes carregada: ${arquivo.name}`, "sucesso");
    } catch (erro) {
      console.error(erro);
      aviso(`Erro na planilha de reagentes: ${erro.message}`, "erro");
    } finally {
      ocultarCarregamento();
    }
  }

  function buscarAba(workbook, nome) {
    const encontrada = workbook.SheetNames.find((item) => normalizar(item) === normalizar(nome));
    return encontrada ? workbook.Sheets[encontrada] : null;
  }

  function percorrerDatas(planilha, linhaInicial, limite, callback) {
    let encontrou = false;
    let vazias = 0;
    for (let linha = linhaInicial; linha <= limite; linha += 1) {
      const data = dataExcel(valorCelula(planilha, "A", linha));
      if (!data) {
        if (encontrou) vazias += 1;
        if (vazias >= 45) break;
        continue;
      }
      encontrou = true;
      vazias = 0;
      callback(linha, data);
    }
  }

  function mapearIndicadores(workbook) {
    const alimentacao = buscarAba(workbook, "Alimentação");
    if (!alimentacao) throw new Error('A aba "Alimentação" não foi encontrada.');

    const producao = buscarAba(workbook, "Controle Produção");
    const recuperacoes = buscarAba(workbook, "Recuperações");
    const p2o5 = buscarAba(workbook, "P2O5");
    const mgo = buscarAba(workbook, "MgO");
    const resultado = {
      alimentacao: [], producao: new Map(), recuperacao: new Map(), p2o5: new Map(), mgo: new Map()
    };

    percorrerDatas(alimentacao, 4, 1000, (linha, data) => {
      resultado.alimentacao.push({
        data,
        pilha: String(valorCelula(alimentacao, "B", linha) || "").trim(),
        lado: String(valorCelula(alimentacao, "C", linha) || "").trim(),
        massaA: numeroCelula(alimentacao, "S", linha) || 0,
        massaB: numeroCelula(alimentacao, "T", linha) || 0,
        massaOficial: numeroCelula(alimentacao, "AC", linha) || 0
      });
    });

    if (producao) {
      percorrerDatas(producao, 4, 1600, (linha, data) => {
        resultado.producao.set(iso(data), numeroCelula(producao, "D", linha) || 0);
      });
    }

    if (recuperacoes) {
      percorrerDatas(recuperacoes, 3, 1000, (linha, data) => {
        const percentualPronto = numeroCelula(recuperacoes, "DQ", linha);
        const fracao = numeroCelula(recuperacoes, "AK", linha);
        resultado.recuperacao.set(iso(data), percentualPronto != null ? percentualPronto : (fracao != null ? fracao * 100 : null));
      });
    }

    if (p2o5) {
      percorrerDatas(p2o5, 3, 1000, (linha, data) => {
        resultado.p2o5.set(iso(data), {
          alimentacao: numeroCelula(p2o5, "CH", linha),
          concentrado: numeroCelula(p2o5, "CO", linha)
        });
      });
    }

    if (mgo) {
      percorrerDatas(mgo, 3, 1000, (linha, data) => {
        resultado.mgo.set(iso(data), {
          alimentacao: numeroCelula(mgo, "CH", linha),
          concentrado: numeroCelula(mgo, "CO", linha)
        });
      });
    }
    return resultado;
  }

  async function carregarPlanilhaIndicadores(evento) {
    const arquivo = evento.target.files?.[0];
    if (!arquivo) return;
    exibirCarregamento(`Lendo indicadores de ${arquivo.name}...`);
    try {
      const abas = ["Alimentação", "Controle Produção", "Recuperações", "P2O5", "MgO"];
      const workbook = await lerWorkbook(arquivo, abas);
      estado.indicadores = mapearIndicadores(workbook);
      estado.arquivoIndicadores = arquivo.name;
      texto("nomeArquivoIndicadoresReagentes", arquivo.name);
      atualizarDashboard();
      aviso(`Planilha de indicadores carregada: ${arquivo.name}`, "sucesso");
    } catch (erro) {
      console.error(erro);
      aviso(`Erro na planilha de indicadores: ${erro.message}`, "erro");
    } finally {
      ocultarCarregamento();
    }
  }

  function configurarPeriodoInicial() {
    const datas = [];
    estado.reagentes.forEach((itens) => itens.filter((item) => item.valido).forEach((item) => datas.push(item.data)));
    if (!datas.length) return;
    datas.sort((a, b) => a - b);
    $("dataInicialReagentes").value = iso(datas[0]);
    $("dataFinalReagentes").value = iso(datas[datas.length - 1]);
  }

  function filtrarItens(itens, inicio, fim) {
    return (itens || []).filter((item) => item.valido && item.data >= inicio && item.data <= fim);
  }

  function resumirReagente(definicao, inicio, fim) {
    const itens = filtrarItens(estado.reagentes.get(definicao.id), inicio, fim);
    const total = soma(itens.map((item) => item.valor));
    const metaAcumulada = soma(itens.map((item) => item.meta));
    const media = itens.length ? total / itens.length : 0;
    const metaDiaria = itens.length ? metaAcumulada / itens.length : 0;
    const aderencia = metaAcumulada ? total / metaAcumulada * 100 : 0;
    const ultimo = itens.at(-1) || null;
    const dataBase = ultimo?.data || fim;
    const diasMes = new Date(dataBase.getFullYear(), dataBase.getMonth() + 1, 0).getDate();
    const diasAno = new Date(dataBase.getFullYear(), 1, 29).getMonth() === 1 ? 366 : 365;
    const projecaoMes = media * diasMes;
    const projecaoAno = media * diasAno;
    const metaMes = metaDiaria * diasMes;
    const status = statusAderencia(aderencia, itens.length);
    return {
      definicao, itens, total, metaAcumulada, media, metaDiaria, aderencia, ultimo,
      dias: itens.length, diasMes, projecaoMes, projecaoAno,
      desvioAtual: total - metaAcumulada, desvioProjetado: projecaoMes - metaMes, status
    };
  }

  function statusAderencia(aderencia, quantidade) {
    if (!quantidade) return { chave: "sem-dados", texto: "Sem dados", cor: CORES.cinza };
    if (aderencia < 95) return { chave: "abaixo", texto: "Abaixo da meta", cor: CORES.verde };
    if (aderencia <= 105) return { chave: "meta", texto: "Na meta", cor: CORES.amarelo };
    return { chave: "acima", texto: "Acima da meta", cor: CORES.vermelho };
  }

  function atualizarDashboard() {
    if (!estado.reagentes.size) {
      aviso("Carregue a planilha de reagentes para visualizar o dashboard.", "info");
      return;
    }
    const inicio = lerDataFiltro("dataInicialReagentes");
    const fim = lerDataFiltro("dataFinalReagentes");
    if (!inicio || !fim || inicio > fim) {
      aviso("Verifique as datas inicial e final.", "erro");
      return;
    }

    const resumos = REAGENTES.map((item) => resumirReagente(item, inicio, fim));
    renderizarFaixas(resumos);
    renderizarCards(resumos);
    atualizarGraficosPrincipais();
    montarAnalisePorPilha(inicio, fim);
    carregarTratativas();
    texto("periodoDashboardReagentes", `Período: ${formatarData(inicio)} a ${formatarData(fim)}`);
    texto("origemDashboardReagentes", origemArquivos());
  }

  function renderizarFaixas(resumos) {
    texto("qtdAcimaMeta", resumos.filter((item) => item.status.chave === "acima").length);
    texto("qtdNaMeta", resumos.filter((item) => item.status.chave === "meta").length);
    texto("qtdAbaixoMeta", resumos.filter((item) => item.status.chave === "abaixo").length);
  }

  function renderizarCards(resumos) {
    $("cardsReagentes").innerHTML = resumos.map((resumo) => {
      const largura = Math.min(Math.max(resumo.aderencia, 0), 100);
      const diferencaPct = resumo.aderencia - 100;
      const ultimoTexto = resumo.ultimo
        ? `Último dia: ${formatarKg(resumo.ultimo.valor)} em ${formatarDataCurta(resumo.ultimo.data)}`
        : "Sem lançamento no período";
      return `
        <article class="card-reagente status-${resumo.status.chave}">
          <header>
            <div><h2>${resumo.definicao.nome}</h2><span class="tag-tipo-reagente">${resumo.definicao.tipo}</span></div>
            <strong class="badge-status-reagente">${resumo.status.texto}</strong>
          </header>
          <div class="metricas-card-reagente">
            <section><span>Acum. real</span><strong>${formatarKg(resumo.total)}</strong></section>
            <section><span>Acum. meta</span><strong>${formatarKg(resumo.metaAcumulada)}</strong></section>
            <section><span>Média/dia</span><strong>${formatarKg(resumo.media)}</strong></section>
            <section><span>Meta diária</span><strong>${formatarKg(resumo.metaDiaria)}</strong></section>
          </div>
          <div class="aderencia-card-reagente">
            <strong>${formatarPercentual(resumo.aderencia)}</strong>
            <span class="desvio-percentual">${formatarSinalPercentual(diferencaPct)} vs meta acumulada</span>
          </div>
          <p class="ultimo-consumo-reagente">${ultimoTexto} · ${resumo.dias} dias lançados</p>
          <div class="projecoes-card-reagente">
            <div><span>Projeção fim do mês</span><strong>${formatarKg(resumo.projecaoMes)}</strong></div>
            <div><span>Projeção anual</span><strong>${formatarKg(resumo.projecaoAno)}</strong></div>
            <div><span>Desvio mensal projetado</span><strong class="${resumo.desvioProjetado <= 0 ? "positivo" : "negativo"}">${formatarSinalKg(resumo.desvioProjetado)}</strong></div>
          </div>
          <div class="barra-aderencia-reagente"><span style="width:${largura}%;background:${resumo.status.cor}"></span></div>
        </article>`;
    }).join("");
  }

  function atualizarGraficosPrincipais() {
    if (!estado.reagentes.size || typeof Chart === "undefined") return;
    const inicio = lerDataFiltro("dataInicialReagentes");
    const fim = lerDataFiltro("dataFinalReagentes");
    if (!inicio || !fim || inicio > fim) return;
    const id = $("filtroReagenteGrafico").value;
    const definicao = REAGENTES.find((item) => item.id === id) || REAGENTES[0];
    const itens = filtrarItens(estado.reagentes.get(definicao.id), inicio, fim);

    estado.graficos.diario?.destroy();
    estado.graficos.aderencia?.destroy();

    const labels = itens.map((item) => formatarDataCurta(item.data));
    const coresBarras = itens.map((item) => statusAderencia(item.meta ? item.valor / item.meta * 100 : 0, 1).cor);
    estado.graficos.diario = new Chart($("graficoConsumoDiarioReagentes"), {
      data: {
        labels,
        datasets: [
          { type: "bar", label: "Consumo realizado", data: itens.map((item) => item.valor), backgroundColor: coresBarras, borderRadius: 5, order: 2 },
          { type: "line", label: "Meta diária", data: itens.map((item) => item.meta), borderColor: CORES.roxo, backgroundColor: CORES.roxo, borderWidth: 2.5, pointRadius: 2, tension: 0.18, order: 1 }
        ]
      },
      options: opcoesGraficoKg()
    });

    let realAcumulado = 0;
    let metaAcumulada = 0;
    const aderencias = itens.map((item) => {
      realAcumulado += item.valor;
      metaAcumulada += item.meta;
      return metaAcumulada ? realAcumulado / metaAcumulada * 100 : 0;
    });
    estado.graficos.aderencia = new Chart($("graficoAderenciaReagentes"), {
      type: "line",
      data: {
        labels,
        datasets: [
          { label: "Aderência acumulada", data: aderencias, borderColor: CORES.verde, backgroundColor: "rgba(47,158,115,.12)", fill: true, borderWidth: 3, pointRadius: 2.5, tension: 0.22 },
          { label: "Referência 100%", data: labels.map(() => 100), borderColor: CORES.roxo, borderDash: [7, 5], borderWidth: 2, pointRadius: 0 }
        ]
      },
      options: opcoesGraficoPercentual()
    });
    texto("subtituloConsumoDiario", `${definicao.nome}: realizado e meta diária em kg.`);
    texto("subtituloAderenciaReagentes", `${definicao.nome}: consumo acumulado dividido pela meta diária acumulada.`);
  }

  function opcoesBaseGrafico() {
    return {
      responsive: true,
      maintainAspectRatio: false,
      devicePixelRatio: 2,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "top", align: "start", labels: { usePointStyle: true, boxWidth: 16, boxHeight: 3, padding: 17 } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { maxRotation: 45, minRotation: 0, font: { size: 10, weight: "600" } } },
        y: { beginAtZero: true, grid: { color: "rgba(45,1,77,.08)" } }
      }
    };
  }

  function opcoesGraficoKg() {
    const opcoes = opcoesBaseGrafico();
    opcoes.plugins.tooltip = { callbacks: { label: (item) => `${item.dataset.label}: ${formatarKg(item.parsed.y)}` } };
    opcoes.scales.y.ticks = { callback: (valor) => formatarNumero(valor) };
    return opcoes;
  }

  function opcoesGraficoPercentual() {
    const opcoes = opcoesBaseGrafico();
    opcoes.plugins.tooltip = { callbacks: { label: (item) => `${item.dataset.label}: ${formatarPercentual(item.parsed.y)}` } };
    opcoes.scales.y.suggestedMin = 80;
    opcoes.scales.y.suggestedMax = 120;
    opcoes.scales.y.ticks = { callback: (valor) => `${formatarNumero(valor, 0)}%` };
    return opcoes;
  }

  function partesPilha(pilha, lado) {
    const pilhas = String(pilha || "").split("/").map((item) => item.trim()).filter(Boolean);
    const lados = String(lado || "").split("/").map((item) => item.trim().toUpperCase()).filter(Boolean);
    const porLado = {};
    pilhas.forEach((numeroPilha, indice) => {
      const ladoPilha = lados[indice];
      if (ladoPilha === "A" || ladoPilha === "B") porLado[ladoPilha] = `${numeroPilha} ${ladoPilha}`;
    });
    if (pilhas.length === 1) {
      const ladoUnico = lados[0] === "A" || lados[0] === "B" ? lados[0] : "";
      const nome = `${pilhas[0]}${ladoUnico ? ` ${ladoUnico}` : ""}`;
      porLado.A ||= nome;
      porLado.B ||= nome;
    }
    return { pilhas, porLado };
  }

  function alocacoesAlimentacao(item) {
    const { pilhas, porLado } = partesPilha(item.pilha, item.lado);
    if (!pilhas.length) return [];
    const massaA = Math.max(item.massaA || 0, 0);
    const massaB = Math.max(item.massaB || 0, 0);
    const total = massaA + massaB;
    if (total <= 0) return [];

    const agrupadas = new Map();
    if (massaA > 0) {
      const nome = porLado.A || pilhas[0];
      agrupadas.set(nome, (agrupadas.get(nome) || 0) + massaA / total);
    }
    if (massaB > 0) {
      const nome = porLado.B || pilhas[1] || pilhas[0];
      agrupadas.set(nome, (agrupadas.get(nome) || 0) + massaB / total);
    }
    return [...agrupadas.entries()].map(([nome, fracao]) => ({ nome, fracao }));
  }

  function valoresReagentesNaData(chaveData) {
    const resultado = {};
    REAGENTES.forEach((definicao) => {
      const item = (estado.reagentes.get(definicao.id) || []).find((linha) => linha.valido && iso(linha.data) === chaveData);
      resultado[definicao.id] = item?.valor || 0;
    });
    return resultado;
  }

  function criarPilha(nome, data) {
    return {
      nome, inicio: data, fim: data, alimentacao: 0, producao: 0,
      reagentes: Object.fromEntries(REAGENTES.map((item) => [item.id, 0])),
      ponderados: {
        recuperacao: 0, recuperacaoPeso: 0,
        p2o5Alimentacao: 0, p2o5AlimentacaoPeso: 0,
        p2o5Concentrado: 0, p2o5ConcentradoPeso: 0,
        mgoAlimentacao: 0, mgoAlimentacaoPeso: 0,
        mgoConcentrado: 0, mgoConcentradoPeso: 0
      }
    };
  }

  function adicionarPonderado(pilha, chave, valor, peso) {
    if (valor == null || !Number.isFinite(valor) || peso <= 0) return;
    pilha.ponderados[chave] += valor * peso;
    pilha.ponderados[`${chave}Peso`] += peso;
  }

  function mediaPonderada(pilha, chave) {
    const peso = pilha.ponderados[`${chave}Peso`];
    return peso ? pilha.ponderados[chave] / peso : null;
  }

  function montarAnalisePorPilha(inicio, fim) {
    if (!estado.indicadores) {
      estado.pilhas = [];
      texto("contadorPilhasTopo", "Carregue os indicadores");
      return;
    }

    const mapaPilhas = new Map();
    estado.diasSemIdentificacao = 0;
    estado.indicadores.alimentacao
      .filter((item) => item.data >= inicio && item.data <= fim && item.massaOficial > 0)
      .forEach((item) => {
        const alocacoes = alocacoesAlimentacao(item);
        if (!alocacoes.length) {
          estado.diasSemIdentificacao += 1;
          return;
        }
        const chaveData = iso(item.data);
        const producao = estado.indicadores.producao.get(chaveData) || 0;
        const recuperacao = estado.indicadores.recuperacao.get(chaveData);
        const p2o5 = estado.indicadores.p2o5.get(chaveData) || {};
        const mgo = estado.indicadores.mgo.get(chaveData) || {};
        const reagentesDia = valoresReagentesNaData(chaveData);

        alocacoes.forEach(({ nome, fracao }) => {
          const pilha = mapaPilhas.get(nome) || criarPilha(nome, item.data);
          pilha.inicio = pilha.inicio < item.data ? pilha.inicio : item.data;
          pilha.fim = pilha.fim > item.data ? pilha.fim : item.data;
          const alimentacaoAlocada = item.massaOficial * fracao;
          pilha.alimentacao += alimentacaoAlocada;
          pilha.producao += producao * fracao;
          REAGENTES.forEach((definicao) => {
            pilha.reagentes[definicao.id] += (reagentesDia[definicao.id] || 0) * fracao;
          });
          adicionarPonderado(pilha, "recuperacao", recuperacao, alimentacaoAlocada);
          adicionarPonderado(pilha, "p2o5Alimentacao", p2o5.alimentacao, alimentacaoAlocada);
          adicionarPonderado(pilha, "p2o5Concentrado", p2o5.concentrado, alimentacaoAlocada);
          adicionarPonderado(pilha, "mgoAlimentacao", mgo.alimentacao, alimentacaoAlocada);
          adicionarPonderado(pilha, "mgoConcentrado", mgo.concentrado, alimentacaoAlocada);
          mapaPilhas.set(nome, pilha);
        });
      });

    estado.pilhas = [...mapaPilhas.values()]
      .map((pilha) => {
        pilha.totalReagentes = soma(Object.values(pilha.reagentes));
        pilha.consumoEspecifico = pilha.alimentacao ? pilha.totalReagentes / pilha.alimentacao : 0;
        pilha.recuperacao = mediaPonderada(pilha, "recuperacao");
        pilha.p2o5Alimentacao = mediaPonderada(pilha, "p2o5Alimentacao");
        pilha.p2o5Concentrado = mediaPonderada(pilha, "p2o5Concentrado");
        pilha.mgoAlimentacao = mediaPonderada(pilha, "mgoAlimentacao");
        pilha.mgoConcentrado = mediaPonderada(pilha, "mgoConcentrado");
        return pilha;
      })
      .sort((a, b) => numeroPilha(a.nome) - numeroPilha(b.nome) || a.nome.localeCompare(b.nome, "pt-BR"));
    renderizarAnalisePorPilha();
  }

  function renderizarAnalisePorPilha() {
    const seletor = $("filtroPilhaReagentes");
    const anterior = seletor.value;
    if (!estado.pilhas.length) {
      seletor.innerHTML = '<option value="">Nenhuma pilha no período</option>';
      texto("contadorPilhasTopo", "Nenhuma pilha no período");
      texto("avisoPilhaReagentes", "Nenhuma pilha com massa identificada foi encontrada no período selecionado.");
      preencherTabelasVazias("Nenhuma pilha encontrada no período.");
      limparPilhaSelecionada();
      return;
    }

    seletor.innerHTML = estado.pilhas.map((pilha) => `<option value="${pilha.nome}">${pilha.nome}</option>`).join("");
    seletor.value = estado.pilhas.some((pilha) => pilha.nome === anterior) ? anterior : estado.pilhas.at(-1).nome;
    texto("contadorPilhasTopo", `${estado.pilhas.length} pilha${estado.pilhas.length > 1 ? "s" : ""} no período`);
    texto("avisoPilhaReagentes", estado.diasSemIdentificacao
      ? `${estado.diasSemIdentificacao} dia(s) com massa foram desconsiderados por não possuírem número de pilha identificado.`
      : "Dados associados pela data. As transições foram divididas proporcionalmente entre Pilha A e Pilha B.");

    $("corpoTabelaDesempenhoPilha").innerHTML = estado.pilhas.map((pilha) => `
      <tr>
        <td><strong>${pilha.nome}</strong></td><td>${formatarDataCurta(pilha.inicio)} a ${formatarDataCurta(pilha.fim)}</td>
        <td>${formatarToneladas(pilha.alimentacao)}</td><td>${formatarToneladas(pilha.producao)}</td><td>${formatarPercentualOuTraco(pilha.recuperacao)}</td>
        <td>${formatarPercentualOuTraco(pilha.p2o5Alimentacao)}</td><td>${formatarPercentualOuTraco(pilha.p2o5Concentrado)}</td>
        <td>${formatarPercentualOuTraco(pilha.mgoAlimentacao)}</td><td>${formatarPercentualOuTraco(pilha.mgoConcentrado)}</td>
      </tr>`).join("");

    $("corpoTabelaConsumoPilha").innerHTML = estado.pilhas.map((pilha) => `
      <tr>
        <td><strong>${pilha.nome}</strong></td>
        ${REAGENTES.map((item) => `<td>${formatarKg(pilha.reagentes[item.id])}</td>`).join("")}
        <td><strong>${formatarKg(pilha.totalReagentes)}</strong></td><td>${formatarNumero(pilha.consumoEspecifico, 3)} kg/t</td>
      </tr>`).join("");
    atualizarPilhaSelecionada();
  }

  function atualizarPilhaSelecionada() {
    const nome = $("filtroPilhaReagentes")?.value;
    const pilha = estado.pilhas.find((item) => item.nome === nome);
    if (!pilha) return limparPilhaSelecionada();

    texto("cardNumeroPilha", pilha.nome);
    texto("cardPeriodoPilha", `${formatarDataCurta(pilha.inicio)} a ${formatarDataCurta(pilha.fim)}`);
    texto("cardAlimentacaoPilha", formatarNumero(pilha.alimentacao));
    texto("cardProducaoPilha", formatarNumero(pilha.producao));
    texto("cardRecuperacaoPilha", formatarNumeroOuTraco(pilha.recuperacao, 1));
    texto("cardConsumoEspecificoPilha", formatarNumero(pilha.consumoEspecifico, 3));
    texto("teorP2O5AlimentacaoPilha", formatarPercentualOuTraco(pilha.p2o5Alimentacao));
    texto("teorP2O5ConcentradoPilha", formatarPercentualOuTraco(pilha.p2o5Concentrado));
    texto("teorMgOAlimentacaoPilha", formatarPercentualOuTraco(pilha.mgoAlimentacao));
    texto("teorMgOConcentradoPilha", formatarPercentualOuTraco(pilha.mgoConcentrado));

    if (typeof Chart !== "undefined") {
      estado.graficos.pilha?.destroy();
      estado.graficos.pilha = new Chart($("graficoConsumoPilhaReagentes"), {
        type: "bar",
        data: {
          labels: REAGENTES.map((item) => item.nome),
          datasets: [{ label: `Consumo — Pilha ${pilha.nome}`, data: REAGENTES.map((item) => pilha.reagentes[item.id]), backgroundColor: ["#3d0764", "#2f9e73", "#68a724", "#ee9622", "#2878b8"], borderRadius: 6 }]
        },
        options: opcoesGraficoKg()
      });
    }
  }

  function limparPilhaSelecionada() {
    ["cardNumeroPilha", "cardAlimentacaoPilha", "cardProducaoPilha", "cardRecuperacaoPilha", "cardConsumoEspecificoPilha",
      "teorP2O5AlimentacaoPilha", "teorP2O5ConcentradoPilha", "teorMgOAlimentacaoPilha", "teorMgOConcentradoPilha"].forEach((id) => texto(id, "—"));
    texto("cardPeriodoPilha", "Sem período");
    estado.graficos.pilha?.destroy();
    estado.graficos.pilha = null;
  }

  function preencherTabelasVazias(mensagem) {
    $("corpoTabelaDesempenhoPilha").innerHTML = `<tr><td colspan="9">${mensagem}</td></tr>`;
    $("corpoTabelaConsumoPilha").innerHTML = `<tr><td colspan="8">${mensagem}</td></tr>`;
  }

  function abrirConsumoPorPilha() {
    $("secaoConsumoPilha")?.scrollIntoView({ behavior: "smooth", block: "start" });
    if (!estado.indicadores) aviso("Carregue a planilha de indicadores para liberar a análise por pilha.", "info");
  }

  function chaveTratativas() {
    return `portal-fvo-reagentes-${$("dataInicialReagentes")?.value || "inicio"}-${$("dataFinalReagentes")?.value || "fim"}`;
  }

  function carregarTratativas() {
    try {
      const salvo = JSON.parse(localStorage.getItem(chaveTratativas()) || "{}");
      $("justificativasReagentes").value = salvo.justificativas || "";
      $("acoesTratativasReagentes").value = salvo.acoes || "";
    } catch (erro) {
      console.warn("Não foi possível recuperar as tratativas.", erro);
    }
    atualizarNotasPDF();
  }

  function salvarTratativas() {
    const conteudo = {
      justificativas: $("justificativasReagentes").value.trim(),
      acoes: $("acoesTratativasReagentes").value.trim()
    };
    try {
      localStorage.setItem(chaveTratativas(), JSON.stringify(conteudo));
      atualizarNotasPDF();
      aviso("Justificativas e ações salvas para o período selecionado.", "sucesso");
    } catch (erro) {
      aviso("O navegador não permitiu salvar os textos.", "erro");
    }
  }

  function atualizarNotasPDF() {
    texto("textoJustificativasPdf", $("justificativasReagentes")?.value.trim() || "Não preenchido.");
    texto("textoAcoesPdf", $("acoesTratativasReagentes")?.value.trim() || "Não preenchido.");
  }

  async function gerarPDF() {
    if (!estado.reagentes.size) return aviso("Carregue a planilha de reagentes antes de gerar o PDF.", "erro");
    if (!window.jspdf || typeof html2canvas === "undefined") return aviso("As bibliotecas de PDF não foram carregadas.", "erro");

    exibirCarregamento("Preparando página 1 de 2...");
    document.body.classList.add("gerando-pdf-reagentes");
    atualizarNotasPDF();
    try {
      await aguardar(500);
      const pagina1 = await capturarElemento($("pagina1PdfReagentes"));
      exibirCarregamento("Preparando página 2 de 2...");
      const pagina2 = await capturarElemento($("pagina2PdfReagentes"));
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4", compress: true });
      adicionarPaginaPDF(pdf, pagina1, "Resumo executivo, consumo diário e aderência", 1);
      pdf.addPage("a4", "landscape");
      adicionarPaginaPDF(pdf, pagina2, "Consumo por pilha, desempenho, justificativas e ações", 2);
      pdf.save(`Relatorio_Reagentes_${$("dataInicialReagentes").value}_a_${$("dataFinalReagentes").value}.pdf`);
      aviso("Relatório gerado em duas páginas A4.", "sucesso");
    } catch (erro) {
      console.error(erro);
      aviso(`Erro ao gerar o PDF: ${erro.message}`, "erro");
    } finally {
      document.body.classList.remove("gerando-pdf-reagentes");
      ocultarCarregamento();
    }
  }

  async function capturarElemento(elemento) {
    const largura = Math.max(elemento.scrollWidth, elemento.clientWidth);
    const altura = Math.max(elemento.scrollHeight, elemento.clientHeight);
    return html2canvas(elemento, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff",
      logging: false,
      width: largura,
      height: altura,
      windowWidth: largura,
      windowHeight: altura
    });
  }

  function adicionarPaginaPDF(pdf, canvas, titulo, numeroPagina) {
    const larguraPagina = pdf.internal.pageSize.getWidth();
    const alturaPagina = pdf.internal.pageSize.getHeight();
    pdf.setFillColor(45, 1, 77);
    pdf.rect(0, 0, larguraPagina, 24, "F");
    pdf.setTextColor(255, 255, 255);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(15);
    pdf.text("PORTAL FVO · CONSUMO DE REAGENTES", 10, 10);
    pdf.setFontSize(9);
    pdf.text(titulo, 10, 18);
    pdf.setFont("helvetica", "normal");
    pdf.text(`${periodoAtual()} · Página ${numeroPagina}/2`, larguraPagina - 10, 10, { align: "right" });
    pdf.text(origemArquivos(), larguraPagina - 10, 17, { align: "right" });

    const margemX = 8;
    const topo = 28;
    const margemInferior = 7;
    const larguraUtil = larguraPagina - margemX * 2;
    const alturaUtil = alturaPagina - topo - margemInferior;
    const escala = Math.min(larguraUtil / canvas.width, alturaUtil / canvas.height);
    const larguraImagem = canvas.width * escala;
    const alturaImagem = canvas.height * escala;
    pdf.addImage(canvas.toDataURL("image/png", 1), "PNG", (larguraPagina - larguraImagem) / 2, topo, larguraImagem, alturaImagem, undefined, "FAST");
  }

  function origemArquivos() {
    const nomes = [];
    if (estado.arquivoReagentes) nomes.push(`Reagentes: ${estado.arquivoReagentes}`);
    if (estado.arquivoIndicadores) nomes.push(`Indicadores: ${estado.arquivoIndicadores}`);
    return nomes.length ? nomes.join(" · ") : "Aguardando planilhas";
  }

  function periodoAtual() {
    const inicio = lerDataFiltro("dataInicialReagentes");
    const fim = lerDataFiltro("dataFinalReagentes");
    return inicio && fim ? `${formatarData(inicio)} a ${formatarData(fim)}` : "Período não informado";
  }

  function exibirCarregamento(mensagem) {
    if (window.FVOCarregamento?.exibir) window.FVOCarregamento.exibir(mensagem);
  }

  function ocultarCarregamento() {
    if (window.FVOCarregamento?.ocultar) window.FVOCarregamento.ocultar();
  }

  function aviso(mensagem, tipo = "info") {
    const elemento = $("avisoReagentes");
    if (!elemento) return;
    elemento.textContent = mensagem;
    elemento.className = `aviso-reagentes aviso-${tipo}`;
  }

  function texto(id, valor) {
    const elemento = $(id);
    if (elemento) elemento.textContent = valor;
  }

  function soma(valores) {
    return valores.reduce((total, valor) => total + (Number.isFinite(valor) ? valor : 0), 0);
  }

  function numeroPilha(textoPilha) {
    return Number(String(textoPilha).match(/\d+/)?.[0] || Number.MAX_SAFE_INTEGER);
  }

  function formatarNumero(valor, casas = 0) {
    return new Intl.NumberFormat("pt-BR", { minimumFractionDigits: casas, maximumFractionDigits: casas }).format(valor || 0);
  }

  function formatarNumeroOuTraco(valor, casas = 1) {
    return valor == null || !Number.isFinite(valor) ? "—" : formatarNumero(valor, casas);
  }

  function formatarKg(valor) {
    return `${formatarNumero(valor)} kg`;
  }

  function formatarToneladas(valor) {
    return `${formatarNumero(valor)} t`;
  }

  function formatarPercentual(valor) {
    return `${formatarNumero(valor, 1)}%`;
  }

  function formatarPercentualOuTraco(valor) {
    return valor == null || !Number.isFinite(valor) ? "—" : `${formatarNumero(valor, 2)}%`;
  }

  function formatarSinalKg(valor) {
    return `${valor > 0 ? "+" : ""}${formatarNumero(valor)} kg`;
  }

  function formatarSinalPercentual(valor) {
    return `${valor > 0 ? "+" : ""}${formatarNumero(valor, 1)}%`;
  }

  function formatarData(data) {
    return data.toLocaleDateString("pt-BR");
  }

  function formatarDataCurta(data) {
    return data.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
  }

  function aguardar(ms) {
    return new Promise((resolver) => setTimeout(resolver, ms));
  }

  if (window.__REAGENTES_FVO_TESTE__) {
    window.__ReagentesFVOTeste = {
      mapearPlanilhaReagentes,
      mapearIndicadores,
      partesPilha,
      alocacoesAlimentacao,
      statusAderencia,
      estado,
      reagentes: REAGENTES
    };
  }
})();