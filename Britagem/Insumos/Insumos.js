(() => {
    "use strict";

    const ANO_REFERENCIA = 2026;
    const COR = {
        roxo: "#4A006F",
        roxo2: "#7B2FA0",
        verde: "#79C943",
        verdeEscuro: "#4D9118",
        laranja: "#F08A18",
        azul: "#2F7EBB",
        turquesa: "#07969A",
        vermelho: "#D83A54",
        cinza: "#A89EAC",
        grade: "rgba(91,76,99,.14)",
        texto: "#4B4050"
    };

    const estado = {
        workbook: null,
        ciclos: [],
        perfil: [],
        mensal: [],
        requisicoes: [],
        estoque: [],
        metaConsumo: 12,
        charts: Object.create(null),
        carregadoEm: null
    };

    const $ = (id) => document.getElementById(id);

    function numero(valor) {
        if (typeof valor === "number" && Number.isFinite(valor)) return valor;
        if (valor === null || valor === undefined || valor === "") return null;
        if (typeof valor === "string") {
            const limpo = valor.trim().replace(/\s/g, "");
            if (!limpo || limpo.startsWith("#")) return null;
            const normalizado = limpo.includes(",") && !limpo.includes(".")
                ? limpo.replace(",", ".")
                : limpo.replace(/\.(?=\d{3}(?:\D|$))/g, "").replace(",", ".");
            const n = Number(normalizado);
            return Number.isFinite(n) ? n : null;
        }
        return null;
    }

    function texto(valor) {
        return String(valor ?? "").trim();
    }

    function normalizarBritador(valor) {
        const v = texto(valor).toUpperCase();
        return v === "A" || v === "B" ? v : null;
    }

    function dataValida(data) {
        return data instanceof Date && !Number.isNaN(data.getTime()) && data.getFullYear() >= 2000 && data.getFullYear() <= 2100;
    }

    function dataExcel(valor, anoPadrao = ANO_REFERENCIA) {
        if (valor instanceof Date) return dataValida(valor) ? valor : null;

        if (typeof valor === "number" && Number.isFinite(valor)) {
            if (valor <= 0) return null;
            const d = XLSX?.SSF?.parse_date_code?.(valor);
            if (!d || !d.y || d.y < 2000 || d.y > 2100) return null;
            const data = new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, Math.floor(d.S || 0));
            return dataValida(data) ? data : null;
        }

        let s = texto(valor);
        if (!s) return null;
        s = s.replace(/\s+/g, "");

        let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\/?(\d{1,2}):(\d{2})$/);
        if (m) {
            let ano = Number(m[3]);
            if (ano < 100) ano += 2000;
            const data = new Date(ano, Number(m[2]) - 1, Number(m[1]), Number(m[4]), Number(m[5]));
            return dataValida(data) ? data : null;
        }

        m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
        if (m) {
            let ano = Number(m[3]);
            if (ano < 100) ano += 2000;
            const data = new Date(ano, Number(m[2]) - 1, Number(m[1]));
            return dataValida(data) ? data : null;
        }

        // Alguns registros do arquivo 2026 foram digitados como dd/mm/hh:mm, sem o ano.
        m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{1,2}):(\d{2})$/);
        if (m) {
            const data = new Date(anoPadrao, Number(m[2]) - 1, Number(m[1]), Number(m[3]), Number(m[4]));
            return dataValida(data) ? data : null;
        }

        return null;
    }

    function formatarData(data, comHora = false) {
        if (!dataValida(data)) return "-";
        const base = new Intl.DateTimeFormat("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric" }).format(data);
        if (!comHora) return base;
        return `${base} ${String(data.getHours()).padStart(2, "0")}:${String(data.getMinutes()).padStart(2, "0")}`;
    }

    function formatarNumero(valor, casas = 1) {
        const n = numero(valor);
        if (n === null) return "-";
        return new Intl.NumberFormat("pt-BR", { minimumFractionDigits: casas, maximumFractionDigits: casas }).format(n);
    }

    function formatarInteiro(valor) {
        const n = numero(valor);
        if (n === null) return "-";
        return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 }).format(n);
    }

    function formatarToneladas(valor) {
        const n = numero(valor);
        return n === null ? "-" : `${formatarNumero(n, 0)} t`;
    }

    function formatarKg(valor) {
        const n = numero(valor);
        return n === null ? "-" : `${formatarNumero(n, 0)} kg`;
    }

    function formatarGt(valor) {
        const n = numero(valor);
        return n === null ? "-" : `${formatarNumero(n, 2)} g/t`;
    }

    function chaveMes(data) {
        if (!dataValida(data)) return null;
        return `${data.getFullYear()}-${String(data.getMonth() + 1).padStart(2, "0")}`;
    }

    function rotuloMes(chave) {
        if (!/^\d{4}-\d{2}$/.test(chave || "")) return chave || "-";
        const [ano, mes] = chave.split("-").map(Number);
        const nome = new Intl.DateTimeFormat("pt-BR", { month: "short" })
            .format(new Date(ano, mes - 1, 1))
            .replace(".", "");
        return `${nome.charAt(0).toUpperCase()}${nome.slice(1)} / ${ano}`;
    }

    function linhasDaAba(nome) {
        const ws = estado.workbook?.Sheets?.[nome];
        if (!ws) return [];
        return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    }

    function parseCiclos(nomeAba, britadorEsperado) {
        const linhas = linhasDaAba(nomeAba);
        const ciclos = [];
        for (let i = 3; i < linhas.length; i++) {
            const r = linhas[i] || [];
            const britador = normalizarBritador(r[0]) || britadorEsperado;
            const fornecedor = texto(r[1]).toUpperCase() || "N/D";
            const dataTroca = dataExcel(r[2]);
            const pesoInicial = numero(r[4]);
            const horimetroInicial = numero(r[5]);
            const dataInversao = dataExcel(r[9]);
            const horimetroInversao = numero(r[11]);
            const dataRetirada = dataExcel(r[15]);
            const pesoFinal = numero(r[17]);
            const massaPlanilha = numero(r[19]);
            const desgastePlanilha = numero(r[20]);
            const consumoPlanilha = numero(r[21]);
            const horasPlanilha = numero(r[22]);

            if (!dataTroca && pesoInicial === null && !texto(r[1])) continue;

            const fechado = Boolean(dataRetirada && pesoInicial !== null && pesoFinal !== null && massaPlanilha !== null && massaPlanilha > 0);
            const desgasteCalculadoG = fechado ? Math.max(0, (pesoInicial - pesoFinal) * 1000) : null;
            const desgasteG = desgasteCalculadoG !== null && desgasteCalculadoG > 0
                ? desgasteCalculadoG
                : (desgastePlanilha !== null && desgastePlanilha > 0 ? desgastePlanilha : null);
            const consumoCalculado = fechado && desgasteG !== null && massaPlanilha > 0 ? desgasteG / massaPlanilha : null;
            const consumo = consumoCalculado !== null && Number.isFinite(consumoCalculado)
                ? consumoCalculado
                : (consumoPlanilha !== null && consumoPlanilha > 0 ? consumoPlanilha : null);
            const horas = horasPlanilha !== null && horasPlanilha > 0 ? horasPlanilha : null;

            ciclos.push({
                britador,
                fornecedor,
                dataTroca,
                pesoInicial,
                horimetroInicial,
                dataInversao,
                horimetroInversao,
                dataRetirada,
                pesoFinal,
                massaProcessada: fechado ? massaPlanilha : null,
                desgasteG: fechado ? desgasteG : null,
                consumo: fechado ? consumo : null,
                horas,
                fechado,
                turmaTroca: texto(r[7]),
                turmaInversao: texto(r[13]),
                linhaExcel: i + 1,
                mes: chaveMes(dataTroca)
            });
        }
        return ciclos;
    }

    function parsePerfil() {
        const nome = estado.workbook?.Sheets?.["Perfil Desgaste"] ? "Perfil Desgaste" : "Perfil de Desgaste";
        const linhas = linhasDaAba(nome);
        const inicio = nome === "Perfil Desgaste" ? 3 : 2;
        const registros = [];
        for (let i = inicio; i < linhas.length; i++) {
            const r = linhas[i] || [];
            const data = dataExcel(r[0]);
            const britador = normalizarBritador(r[1]);
            const medida = numero(r[3]);
            if (!data || !britador || medida === null || medida <= 0 || medida > 50) continue;
            registros.push({
                data,
                britador,
                turma: texto(r[2]),
                medida,
                atividade: texto(r[4]) || "N/D",
                mes: chaveMes(data),
                linhaExcel: i + 1
            });
        }
        return registros.sort((a, b) => a.data - b.data);
    }

    function parseMensal() {
        const linhas = linhasDaAba("Consumo Mensal");
        const registros = [];
        for (let i = 2; i < linhas.length; i++) {
            const r = linhas[i] || [];
            const data = dataExcel(r[0]);
            const producao = numero(r[1]);
            const pesoG = numero(r[2]);
            const meta = numero(r[3]);
            const consumo = numero(r[4]);
            if (!data) continue;
            if (meta !== null && meta > 0) estado.metaConsumo = meta;
            registros.push({
                data,
                producao,
                pesoG,
                meta: meta !== null && meta > 0 ? meta : estado.metaConsumo,
                consumo: consumo !== null && consumo > 0 ? consumo : null,
                mes: chaveMes(data)
            });
        }
        return registros.sort((a, b) => a.data - b.data);
    }

    function parseRequisicoes() {
        const linhas = linhasDaAba("Requisições");
        const registros = [];
        for (let i = 2; i < linhas.length; i++) {
            const r = linhas[i] || [];
            const britador = normalizarBritador(r[0]);
            const data = dataExcel(r[1]);
            const quantidade = numero(r[2]);
            if (!britador || !data || quantidade === null || quantidade <= 0) continue;
            registros.push({
                britador,
                data,
                quantidade,
                requisitante: texto(r[3]) || "N/D",
                reserva: texto(r[4]) || "N/D",
                mes: chaveMes(data)
            });
        }
        return registros.sort((a, b) => a.data - b.data);
    }

    function parseEstoque() {
        const linhas = linhasDaAba("Estoque Sucata");
        const registros = [];
        for (let i = 2; i < linhas.length; i++) {
            const r = linhas[i] || [];
            const data = dataExcel(r[0]);
            if (!data) continue;
            const a = numero(r[1]);
            const b = numero(r[2]);
            if (a !== null) registros.push({ data, britador: "A", quantidade: a });
            if (b !== null) registros.push({ data, britador: "B", quantidade: b });
        }
        return registros.sort((a, b) => a.data - b.data);
    }

    function abasAusentes() {
        const obrigatorias = ["Controle de Barras Britador A", "Controle de Barras Britador B"];
        const opcionais = ["Perfil Desgaste", "Requisições", "Consumo Mensal", "Estoque Sucata"];
        return {
            obrigatorias: obrigatorias.filter((n) => !estado.workbook?.Sheets?.[n]),
            opcionais: opcionais.filter((n) => !estado.workbook?.Sheets?.[n])
        };
    }

    function exibirCarregamento(msg) {
        window.FVOCarregamento?.exibir?.(msg);
    }

    function ocultarCarregamento() {
        window.FVOCarregamento?.ocultar?.();
    }

    function statusBase(msg, estadoStatus = "") {
        const el = $("statusBaseInsumos");
        if (!el) return;
        el.textContent = msg;
        if (estadoStatus) el.dataset.estado = estadoStatus;
        else delete el.dataset.estado;
    }

    async function carregarPlanilha() {
        const input = $("arquivoBarras");
        const arquivo = input?.files?.[0];
        if (!arquivo) {
            statusBase("Selecione a planilha", "erro");
            return;
        }
        if (typeof XLSX === "undefined") {
            statusBase("Biblioteca SheetJS não carregada", "erro");
            return;
        }

        try {
            statusBase("Lendo arquivo...", "carregando");
            exibirCarregamento("Lendo Controle de Barras");
            await new Promise((resolve) => setTimeout(resolve, 30));
            const buffer = await arquivo.arrayBuffer();
            estado.workbook = XLSX.read(buffer, { type: "array", cellDates: false });

            const faltas = abasAusentes();
            if (faltas.obrigatorias.length) {
                throw new Error(`Aba obrigatória não encontrada: ${faltas.obrigatorias.join(", ")}`);
            }

            exibirCarregamento("Validando ciclos dos Britadores A e B");
            estado.metaConsumo = 12;
            estado.ciclos = [
                ...parseCiclos("Controle de Barras Britador A", "A"),
                ...parseCiclos("Controle de Barras Britador B", "B")
            ];
            estado.perfil = parsePerfil();
            estado.mensal = parseMensal();
            estado.requisicoes = parseRequisicoes();
            estado.estoque = parseEstoque();
            estado.carregadoEm = new Date();

            preencherPeriodos();
            atualizarTudo();

            const avisos = faltas.opcionais.length ? ` • sem ${faltas.opcionais.join(", ")}` : "";
            statusBase(`Dados carregados${avisos}`, faltas.opcionais.length ? "carregando" : "ok");
            $("ultimaAtualizacaoInsumos").textContent = new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "short" }).format(estado.carregadoEm);
        } catch (erro) {
            console.error("Erro ao carregar Insumos:", erro);
            estado.workbook = null;
            statusBase(erro?.message || "Erro ao carregar planilha", "erro");
            alert(`Não foi possível carregar a planilha de barras.\n\n${erro?.message || erro}`);
        } finally {
            ocultarCarregamento();
        }
    }

    function preencherPeriodos() {
        const select = $("filtroPeriodoInsumos");
        if (!select) return;
        const atual = select.value;
        const meses = new Set();
        [...estado.ciclos, ...estado.perfil, ...estado.requisicoes, ...estado.mensal].forEach((r) => {
            if (r.mes) meses.add(r.mes);
        });
        const lista = [...meses].sort().reverse();
        select.innerHTML = '<option value="TODOS">Histórico completo</option>' +
            lista.map((m) => `<option value="${m}">${rotuloMes(m)}</option>`).join("");
        if (lista.includes(atual)) select.value = atual;
    }

    function filtros() {
        return {
            britador: $("filtroBritadorInsumos")?.value || "TODOS",
            mes: $("filtroPeriodoInsumos")?.value || "TODOS"
        };
    }

    function filtrar(lista, { britador, mes }, usarBritador = true) {
        return lista.filter((r) => {
            if (usarBritador && britador !== "TODOS" && r.britador !== britador) return false;
            if (mes !== "TODOS" && r.mes !== mes) return false;
            return true;
        });
    }

    function statusMedida(medida) {
        const n = numero(medida);
        if (n === null) return { texto: "SEM MEDIÇÃO", classe: "status-neutro", tabela: "neutro" };
        if (n < 2.49) return { texto: "REALIZAR TROCA", classe: "status-critico", tabela: "critico" };
        if (n >= 3 && n <= 5) return { texto: "INSPEÇÃO 2×/TURNO", classe: "status-atencao", tabela: "atencao" };
        if (n > 5 && n <= 11) return { texto: "INSPEÇÃO 1×/TURNO", classe: "status-ok", tabela: "ok" };
        return { texto: "SEM CRITÉRIO EXPLÍCITO", classe: "status-neutro", tabela: "neutro" };
    }

    function ultimaPorData(lista, campo = "data") {
        return [...lista].filter((r) => dataValida(r[campo])).sort((a, b) => b[campo] - a[campo])[0] || null;
    }

    function situacaoBritador(britador) {
        const ciclos = estado.ciclos.filter((c) => c.britador === britador && dataValida(c.dataTroca));
        const atual = [...ciclos].sort((a, b) => b.dataTroca - a.dataTroca)[0] || null;
        const ultimoFechado = [...ciclos].filter((c) => c.fechado && c.consumo !== null)
            .sort((a, b) => (b.dataRetirada || b.dataTroca) - (a.dataRetirada || a.dataTroca))[0] || null;
        const medicao = ultimaPorData(estado.perfil.filter((p) => p.britador === britador));
        const status = statusMedida(medicao?.medida);

        const set = (id, val) => { const el = $(id); if (el) el.textContent = val; };
        set(`fornecedorAtual${britador}`, atual?.fornecedor || "-");
        set(`inicioCampanha${britador}`, formatarData(atual?.dataTroca));
        set(`ultimaInversao${britador}`, formatarData(atual?.dataInversao));
        set(`ultimaMedicao${britador}`, medicao ? formatarData(medicao.data, true) : "-");
        set(`medidaAtual${britador}`, medicao ? `${formatarNumero(medicao.medida, 1)} cm` : "-");
        set(`ultimoConsumo${britador}`, ultimoFechado ? formatarGt(ultimoFechado.consumo) : "-");

        const badge = $(`statusBarra${britador}`);
        if (badge) {
            badge.textContent = status.texto;
            badge.className = `status-barra-insumos ${status.classe}`;
        }
    }

    function atualizarKpis(ciclos, requisicoes) {
        const fechados = ciclos.filter((c) => c.fechado && c.massaProcessada > 0 && c.desgasteG > 0);
        const massa = fechados.reduce((s, c) => s + c.massaProcessada, 0);
        const desgasteG = fechados.reduce((s, c) => s + c.desgasteG, 0);
        const consumoPonderado = massa > 0 ? desgasteG / massa : null;
        const req = requisicoes.reduce((s, r) => s + r.quantidade, 0);

        $("kpiConsumoPonderado").textContent = formatarGt(consumoPonderado);
        $("kpiMassaProcessada").textContent = formatarToneladas(massa);
        $("kpiMassaPerdida").textContent = formatarKg(desgasteG / 1000);
        $("kpiCiclos").textContent = formatarInteiro(fechados.length);
        $("kpiRequisicoes").textContent = formatarInteiro(req);
        $("kpiMeta").textContent = formatarGt(estado.metaConsumo);
        $("metaConsumoTopo").textContent = formatarGt(estado.metaConsumo);
        $("badgeMetaCiclos").textContent = `Meta ${formatarGt(estado.metaConsumo)}`;
    }

    function destruirChart(nome) {
        if (estado.charts[nome]) {
            estado.charts[nome].destroy();
            delete estado.charts[nome];
        }
    }

    function opcoesBaseChart(extra = {}) {
        return {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,
            devicePixelRatio: 2.2,
            interaction: { mode: "nearest", intersect: false },
            plugins: {
                legend: {
                    position: "top",
                    labels: { color: COR.texto, boxWidth: 12, usePointStyle: true, pointStyle: "circle", font: { size: 10, weight: "700" } }
                },
                tooltip: { backgroundColor: "rgba(48,13,64,.94)", titleFont: { size: 11 }, bodyFont: { size: 10 }, padding: 10 },
                datalabels: { display: false }
            },
            scales: {
                x: { grid: { display: false }, ticks: { color: COR.texto, font: { size: 9, weight: "600" }, maxRotation: 55, minRotation: 0 } },
                y: { beginAtZero: true, grid: { color: COR.grade }, ticks: { color: COR.texto, font: { size: 9, weight: "600" } } }
            },
            ...extra
        };
    }

    function graficoConsumoCiclos(ciclos) {
        destruirChart("consumoCiclos");
        const dados = ciclos.filter((c) => c.fechado && c.consumo !== null).sort((a, b) => a.dataTroca - b.dataTroca);
        const labels = dados.map((c) => `${c.britador} • ${formatarData(c.dataTroca).slice(0, 5)}`);
        const valores = dados.map((c) => c.consumo);
        const cores = dados.map((c) => c.britador === "A" ? COR.roxo : COR.verdeEscuro);
        const ctx = $("graficoConsumoCiclos")?.getContext("2d");
        if (!ctx) return;
        estado.charts.consumoCiclos = new Chart(ctx, {
            type: "bar",
            data: {
                labels,
                datasets: [
                    { label: "Consumo do ciclo", data: valores, backgroundColor: cores, borderRadius: 6, maxBarThickness: 30 },
                    { label: "Meta", data: labels.map(() => estado.metaConsumo), type: "line", borderColor: COR.laranja, backgroundColor: COR.laranja, borderWidth: 2.5, pointRadius: 0, tension: 0, order: 0 }
                ]
            },
            options: opcoesBaseChart({
                plugins: {
                    ...opcoesBaseChart().plugins,
                    datalabels: {
                        display: (ctx) => ctx.datasetIndex === 0 && labels.length <= 15,
                        color: COR.roxo,
                        anchor: "end",
                        align: "top",
                        formatter: (v) => `${formatarNumero(v, 1)} g/t`,
                        font: { size: 8, weight: "800" }
                    }
                },
                scales: {
                    ...opcoesBaseChart().scales,
                    y: { beginAtZero: true, suggestedMax: Math.max(15, ...valores, estado.metaConsumo) * 1.15, title: { display: true, text: "Consumo (g/t)", color: COR.texto, font: { size: 10, weight: "700" } }, grid: { color: COR.grade }, ticks: { color: COR.texto, font: { size: 9, weight: "600" } } }
                }
            })
        });
    }

    function graficoMassaPerdida(ciclos) {
        destruirChart("massaPerdida");
        const dados = ciclos.filter((c) => c.fechado && c.desgasteG !== null).sort((a, b) => a.dataTroca - b.dataTroca);
        const labels = dados.map((c) => `${c.britador} • ${formatarData(c.dataTroca).slice(0, 5)}`);
        const valores = dados.map((c) => c.desgasteG / 1000);
        const ctx = $("graficoMassaPerdida")?.getContext("2d");
        if (!ctx) return;
        estado.charts.massaPerdida = new Chart(ctx, {
            type: "bar",
            data: { labels, datasets: [{ label: "Massa perdida (kg)", data: valores, backgroundColor: dados.map((c) => c.britador === "A" ? "#7B2FA0" : "#79C943"), borderRadius: 6, maxBarThickness: 30 }] },
            options: opcoesBaseChart({
                plugins: {
                    ...opcoesBaseChart().plugins,
                    datalabels: { display: (ctx) => labels.length <= 14, anchor: "end", align: "top", color: COR.roxo, formatter: (v) => `${formatarNumero(v, 0)} kg`, font: { size: 8, weight: "800" } }
                },
                scales: {
                    ...opcoesBaseChart().scales,
                    y: { beginAtZero: true, title: { display: true, text: "Massa perdida (kg)", color: COR.texto, font: { size: 10, weight: "700" } }, grid: { color: COR.grade }, ticks: { color: COR.texto, font: { size: 9, weight: "600" } } }
                }
            })
        });
    }

    function graficoPerfil(perfil) {
        destruirChart("perfil");
        const dados = [...perfil].sort((a, b) => a.data - b.data);
        const labels = dados.map((r) => `${formatarData(r.data).slice(0, 5)} ${String(r.data.getHours()).padStart(2, "0")}:${String(r.data.getMinutes()).padStart(2, "0")}`);
        const a = dados.map((r) => r.britador === "A" ? r.medida : null);
        const b = dados.map((r) => r.britador === "B" ? r.medida : null);
        const ctx = $("graficoPerfilDesgaste")?.getContext("2d");
        if (!ctx) return;
        estado.charts.perfil = new Chart(ctx, {
            type: "line",
            data: {
                labels,
                datasets: [
                    { label: "Britador A", data: a, borderColor: COR.roxo, backgroundColor: COR.roxo, borderWidth: 2.4, pointRadius: 2.2, pointHoverRadius: 5, spanGaps: true, tension: .18 },
                    { label: "Britador B", data: b, borderColor: COR.verdeEscuro, backgroundColor: COR.verdeEscuro, borderWidth: 2.4, pointRadius: 2.2, pointHoverRadius: 5, spanGaps: true, tension: .18 },
                    { label: "Troca 2,49 cm", data: labels.map(() => 2.49), borderColor: COR.vermelho, borderWidth: 1.5, borderDash: [7, 5], pointRadius: 0, tension: 0 },
                    { label: "Inspeção 2× 5 cm", data: labels.map(() => 5), borderColor: COR.laranja, borderWidth: 1.5, borderDash: [7, 5], pointRadius: 0, tension: 0 },
                    { label: "Limite ref. 11 cm", data: labels.map(() => 11), borderColor: COR.azul, borderWidth: 1.2, borderDash: [4, 5], pointRadius: 0, tension: 0 }
                ]
            },
            options: opcoesBaseChart({
                plugins: {
                    ...opcoesBaseChart().plugins,
                    legend: { position: "top", labels: { color: COR.texto, boxWidth: 10, usePointStyle: true, pointStyle: "circle", font: { size: 9, weight: "700" } } },
                    datalabels: { display: false }
                },
                scales: {
                    x: { grid: { display: false }, ticks: { color: COR.texto, autoSkip: true, maxTicksLimit: 18, maxRotation: 55, font: { size: 8, weight: "600" } } },
                    y: { beginAtZero: true, suggestedMax: 13, title: { display: true, text: "Espessura (cm)", color: COR.texto, font: { size: 10, weight: "700" } }, grid: { color: COR.grade }, ticks: { color: COR.texto, font: { size: 9, weight: "600" } } }
                }
            })
        });
    }

    function graficoMensal(mensal) {
        destruirChart("mensal");
        const dados = mensal.filter((r) => r.consumo !== null).sort((a, b) => a.data - b.data);
        const labels = dados.map((r) => rotuloMes(r.mes).replace(" / ", "/"));
        const ctx = $("graficoConsumoMensal")?.getContext("2d");
        if (!ctx) return;
        estado.charts.mensal = new Chart(ctx, {
            type: "bar",
            data: {
                labels,
                datasets: [
                    { label: "Consumo mensal", data: dados.map((r) => r.consumo), backgroundColor: COR.roxo2, borderRadius: 7, maxBarThickness: 38 },
                    { label: "Meta", data: dados.map((r) => r.meta || estado.metaConsumo), type: "line", borderColor: COR.laranja, backgroundColor: COR.laranja, borderWidth: 2.4, pointRadius: 2.5, tension: .1 }
                ]
            },
            options: opcoesBaseChart({
                plugins: {
                    ...opcoesBaseChart().plugins,
                    datalabels: { display: (ctx) => ctx.datasetIndex === 0 && labels.length <= 14, anchor: "end", align: "top", color: COR.roxo, formatter: (v) => formatarNumero(v, 1), font: { size: 8, weight: "800" } }
                },
                scales: {
                    ...opcoesBaseChart().scales,
                    y: { beginAtZero: true, title: { display: true, text: "Consumo (g/t)", color: COR.texto, font: { size: 10, weight: "700" } }, grid: { color: COR.grade }, ticks: { color: COR.texto, font: { size: 9, weight: "600" } } }
                }
            })
        });
    }

    function agruparFornecedores(ciclos) {
        const mapa = new Map();
        ciclos.filter((c) => c.fechado && c.consumo !== null && c.massaProcessada > 0 && c.desgasteG > 0).forEach((c) => {
            const chave = `${c.britador}|${c.fornecedor}`;
            if (!mapa.has(chave)) mapa.set(chave, { britador: c.britador, fornecedor: c.fornecedor, ciclos: 0, massa: 0, desgasteG: 0, somaConsumo: 0 });
            const g = mapa.get(chave);
            g.ciclos += 1;
            g.massa += c.massaProcessada;
            g.desgasteG += c.desgasteG;
            g.somaConsumo += c.consumo;
        });
        return [...mapa.values()].map((g) => ({
            ...g,
            media: g.ciclos ? g.somaConsumo / g.ciclos : null,
            ponderado: g.massa ? g.desgasteG / g.massa : null
        })).sort((a, b) => (a.britador + a.fornecedor).localeCompare(b.britador + b.fornecedor));
    }

    function graficoFornecedor(ciclos) {
        destruirChart("fornecedor");
        const grupos = agruparFornecedores(ciclos);
        const labels = grupos.map((g) => `${g.fornecedor} • ${g.britador}`);
        const ctx = $("graficoFornecedor")?.getContext("2d");
        if (!ctx) return;
        estado.charts.fornecedor = new Chart(ctx, {
            type: "bar",
            data: { labels, datasets: [{ label: "Consumo médio (g/t)", data: grupos.map((g) => g.media), backgroundColor: grupos.map((g) => g.britador === "A" ? COR.roxo : COR.verdeEscuro), borderRadius: 7, maxBarThickness: 46 }] },
            options: opcoesBaseChart({
                indexAxis: "y",
                plugins: {
                    ...opcoesBaseChart().plugins,
                    datalabels: { display: true, anchor: "end", align: "right", color: COR.roxo, formatter: (v) => `${formatarNumero(v, 1)} g/t`, font: { size: 8, weight: "800" } }
                },
                scales: {
                    x: { beginAtZero: true, grid: { color: COR.grade }, ticks: { color: COR.texto, font: { size: 9, weight: "600" } }, title: { display: true, text: "Consumo médio (g/t)", color: COR.texto, font: { size: 10, weight: "700" } } },
                    y: { grid: { display: false }, ticks: { color: COR.texto, font: { size: 9, weight: "700" } } }
                }
            })
        });
    }

    function graficoRequisicoes(requisicoes) {
        destruirChart("requisicoes");
        const meses = [...new Set(requisicoes.map((r) => r.mes).filter(Boolean))].sort();
        const soma = (mes, britador) => requisicoes.filter((r) => r.mes === mes && r.britador === britador).reduce((s, r) => s + r.quantidade, 0);
        const ctx = $("graficoRequisicoesMensais")?.getContext("2d");
        if (!ctx) return;
        estado.charts.requisicoes = new Chart(ctx, {
            type: "bar",
            data: {
                labels: meses.map((m) => rotuloMes(m).replace(" / ", "/")),
                datasets: [
                    { label: "Britador A", data: meses.map((m) => soma(m, "A")), backgroundColor: COR.roxo, borderRadius: 6 },
                    { label: "Britador B", data: meses.map((m) => soma(m, "B")), backgroundColor: COR.verde, borderRadius: 6 }
                ]
            },
            options: opcoesBaseChart({
                plugins: { ...opcoesBaseChart().plugins, datalabels: { display: (ctx) => (ctx.dataset.data[ctx.dataIndex] || 0) > 0, color: "#fff", anchor: "center", align: "center", formatter: (v) => v || "", font: { size: 8, weight: "900" } } },
                scales: {
                    x: { stacked: true, grid: { display: false }, ticks: { color: COR.texto, font: { size: 9, weight: "600" } } },
                    y: { stacked: true, beginAtZero: true, title: { display: true, text: "Barras requisitadas", color: COR.texto, font: { size: 10, weight: "700" } }, grid: { color: COR.grade }, ticks: { precision: 0, color: COR.texto, font: { size: 9, weight: "600" } } }
                }
            })
        });
    }

    function atualizarTabelaCiclos(ciclos) {
        const tbody = $("tbodyCiclosInsumos");
        if (!tbody) return;
        const dados = ciclos.filter((c) => c.fechado).sort((a, b) => (b.dataRetirada || b.dataTroca) - (a.dataRetirada || a.dataTroca)).slice(0, 14);
        if (!dados.length) {
            tbody.innerHTML = '<tr><td colspan="8">Sem ciclos concluídos para os filtros selecionados.</td></tr>';
            return;
        }
        tbody.innerHTML = dados.map((c) => `
            <tr>
                <td><span class="tag-britador ${c.britador === "B" ? "b" : ""}">${c.britador}</span></td>
                <td><strong>${c.fornecedor}</strong></td>
                <td>${formatarData(c.dataTroca)}</td>
                <td>${formatarData(c.dataRetirada)}</td>
                <td>${formatarToneladas(c.massaProcessada)}</td>
                <td>${formatarKg(c.desgasteG / 1000)}</td>
                <td><strong>${formatarGt(c.consumo)}</strong></td>
                <td>${c.horas !== null ? `${formatarNumero(c.horas, 1)} h` : "-"}</td>
            </tr>`).join("");
    }

    function atualizarTabelaPerfil(perfil) {
        const tbody = $("tbodyDesgasteInsumos");
        if (!tbody) return;
        const dados = [...perfil].sort((a, b) => b.data - a.data).slice(0, 22);
        if (!dados.length) {
            tbody.innerHTML = '<tr><td colspan="6">Sem medições válidas para os filtros selecionados.</td></tr>';
            return;
        }
        tbody.innerHTML = dados.map((r) => {
            const s = statusMedida(r.medida);
            return `<tr>
                <td>${formatarData(r.data, true)}</td>
                <td><span class="tag-britador ${r.britador === "B" ? "b" : ""}">${r.britador}</span></td>
                <td>${r.turma || "-"}</td>
                <td><strong>${formatarNumero(r.medida, 1)} cm</strong></td>
                <td>${r.atividade}</td>
                <td><span class="tag-status-tabela ${s.tabela}">${s.texto}</span></td>
            </tr>`;
        }).join("");
    }

    function atualizarFornecedores(ciclos) {
        const grupos = agruparFornecedores(ciclos);
        const tbody = $("tbodyFornecedoresInsumos");
        if (!tbody) return;
        if (!grupos.length) {
            tbody.innerHTML = '<tr><td colspan="7">Sem dados válidos de fornecedor para o filtro.</td></tr>';
            return;
        }
        tbody.innerHTML = grupos.map((g) => `
            <tr>
                <td><span class="tag-britador ${g.britador === "B" ? "b" : ""}">${g.britador}</span></td>
                <td><strong>${g.fornecedor}</strong></td>
                <td>${g.ciclos}</td>
                <td>${formatarToneladas(g.massa)}</td>
                <td>${formatarKg(g.desgasteG / 1000)}</td>
                <td>${formatarGt(g.media)}</td>
                <td><strong>${formatarGt(g.ponderado)}</strong></td>
            </tr>`).join("");
    }

    function atualizarResumoConsumo(ciclos) {
        const dados = ciclos.filter((c) => c.fechado && c.consumo !== null);
        const melhor = [...dados].sort((a, b) => a.consumo - b.consumo)[0] || null;
        const pior = [...dados].sort((a, b) => b.consumo - a.consumo)[0] || null;
        const horas = dados.reduce((s, c) => s + (c.horas || 0), 0);
        const acima = dados.filter((c) => c.consumo > estado.metaConsumo).length;

        $("resumoMelhorCiclo").textContent = melhor ? formatarGt(melhor.consumo) : "-";
        $("resumoMelhorCicloInfo").textContent = melhor ? `Britador ${melhor.britador} • ${melhor.fornecedor} • ${formatarData(melhor.dataTroca)}` : "-";
        $("resumoPiorCiclo").textContent = pior ? formatarGt(pior.consumo) : "-";
        $("resumoPiorCicloInfo").textContent = pior ? `Britador ${pior.britador} • ${pior.fornecedor} • ${formatarData(pior.dataTroca)}` : "-";
        $("resumoHorasOperadas").textContent = dados.length ? `${formatarNumero(horas, 1)} h` : "-";
        $("resumoAcimaMeta").textContent = dados.length ? `${acima} de ${dados.length}` : "-";
    }

    function atualizarRequisicoes(requisicoesPeriodo) {
        const mesFiltro = filtros().mes;
        const reqPeriodoTodos = estado.requisicoes.filter((r) => mesFiltro === "TODOS" || r.mes === mesFiltro);
        const totalA = reqPeriodoTodos.filter((r) => r.britador === "A").reduce((s, r) => s + r.quantidade, 0);
        const totalB = reqPeriodoTodos.filter((r) => r.britador === "B").reduce((s, r) => s + r.quantidade, 0);
        $("reqTotalA").textContent = formatarInteiro(totalA);
        $("reqTotalB").textContent = formatarInteiro(totalB);
        $("totalReqFiltro").textContent = formatarInteiro(requisicoesPeriodo.reduce((s, r) => s + r.quantidade, 0));

        ["A", "B"].forEach((b) => {
            const estoque = ultimaPorData(estado.estoque.filter((e) => e.britador === b));
            $(`estoqueAtual${b}`).textContent = estoque ? formatarInteiro(estoque.quantidade) : "-";
            $(`estoqueData${b}`).textContent = estoque ? `Último registro: ${formatarData(estoque.data)}` : "Último registro: -";

            const req = ultimaPorData(estado.requisicoes.filter((r) => r.britador === b));
            $(`ultimaReq${b}`).textContent = req ? `${formatarInteiro(req.quantidade)} barras` : "-";
            $(`ultimaReq${b}Info`).textContent = req ? `${formatarData(req.data)} • Reserva ${req.reserva} • ${req.requisitante}` : "-";
        });

        const tbody = $("tbodyRequisicoesInsumos");
        const dados = [...requisicoesPeriodo].sort((a, b) => b.data - a.data).slice(0, 24);
        tbody.innerHTML = dados.length ? dados.map((r) => `
            <tr>
                <td>${formatarData(r.data)}</td>
                <td><span class="tag-britador ${r.britador === "B" ? "b" : ""}">${r.britador}</span></td>
                <td><strong>${formatarInteiro(r.quantidade)}</strong></td>
                <td>${r.requisitante}</td>
                <td>${r.reserva}</td>
            </tr>`).join("") : '<tr><td colspan="5">Sem requisições para os filtros selecionados.</td></tr>';
    }

    function atualizarTudo() {
        if (!estado.workbook) return;
        const f = filtros();
        const ciclos = filtrar(estado.ciclos, f);
        const perfil = filtrar(estado.perfil, f);
        const req = filtrar(estado.requisicoes, f);
        const mensal = f.mes === "TODOS" ? estado.mensal : estado.mensal.filter((r) => r.mes === f.mes);

        atualizarKpis(ciclos, req);
        situacaoBritador("A");
        situacaoBritador("B");
        graficoConsumoCiclos(ciclos);
        graficoMassaPerdida(ciclos);
        graficoPerfil(perfil);
        graficoMensal(mensal);
        graficoFornecedor(ciclos);
        graficoRequisicoes(req);
        atualizarTabelaCiclos(ciclos);
        atualizarTabelaPerfil(perfil);
        atualizarFornecedores(ciclos);
        atualizarResumoConsumo(ciclos);
        atualizarRequisicoes(req);
        atualizarTextoPDF();
    }

    function atualizarTextoPDF() {
        const f = filtros();
        const brit = f.britador === "TODOS" ? "Britadores A × B" : `Britador ${f.britador}`;
        const periodo = f.mes === "TODOS" ? "Histórico completo" : rotuloMes(f.mes);
        $("pdfFiltroInsumos").textContent = `${brit} • ${periodo}`;
    }

    function configurarAbas() {
        document.querySelectorAll(".aba-insumos").forEach((botao) => {
            botao.addEventListener("click", () => {
                const aba = botao.dataset.aba;
                document.querySelectorAll(".aba-insumos").forEach((b) => b.classList.toggle("ativa", b === botao));
                document.querySelectorAll(".conteudo-aba-insumos").forEach((secao) => secao.classList.toggle("ativa", secao.id === `aba-${aba}`));
                setTimeout(() => Object.values(estado.charts).forEach((c) => c.resize()), 40);
            });
        });
    }

    function gerarPDF() {
        if (!estado.workbook) {
            alert("Carregue a planilha Controle de Barras antes de gerar o PDF.");
            return;
        }
        const agora = new Date();
        const textoGeracao = `Gerado em: ${new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "medium" }).format(agora)}`;
        $("pdfGeradoInsumos").textContent = textoGeracao;
        $("rodapeGeracaoInsumos").textContent = textoGeracao;
        atualizarTextoPDF();
        Object.values(estado.charts).forEach((chart) => {
            chart.options.devicePixelRatio = 3.2;
            chart.resize();
            chart.update("none");
        });
        setTimeout(() => window.print(), 180);
    }

    window.addEventListener("afterprint", () => {
        Object.values(estado.charts).forEach((chart) => {
            chart.options.devicePixelRatio = 2.2;
            chart.resize();
            chart.update("none");
        });
    });

    document.addEventListener("DOMContentLoaded", () => {
        if (typeof Chart !== "undefined" && typeof ChartDataLabels !== "undefined") {
            try { Chart.register(ChartDataLabels); } catch (_) { /* já registrado */ }
            Chart.defaults.font.family = 'Inter, "Segoe UI", Arial, sans-serif';
            Chart.defaults.color = COR.texto;
        }

        configurarAbas();
        $("btnCarregarInsumos")?.addEventListener("click", carregarPlanilha);
        $("btnAtualizarInsumos")?.addEventListener("click", atualizarTudo);
        $("filtroBritadorInsumos")?.addEventListener("change", atualizarTudo);
        $("filtroPeriodoInsumos")?.addEventListener("change", atualizarTudo);
        $("btnGerarPDFInsumos")?.addEventListener("click", gerarPDF);
    });
})();