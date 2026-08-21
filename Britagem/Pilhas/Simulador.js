(function(){
    "use strict";

    const estado = {
        cenario:"provavel",
        horarios:[],
        fila:[]
    };

    const fatores = {
        conservador:{producao:0.85,ritmo:0.90,parada:1.25},
        provavel:{producao:1,ritmo:1,parada:1},
        otimista:{producao:1.10,ritmo:1.05,parada:0.75}
    };

    document.addEventListener("DOMContentLoaded",inicializarSimulador);

    function inicializarSimulador(){
        conectar("abaPlanejamentoCompleto","click",function(){alternarAba(false);});
        conectar("abaSimuladorPilha","click",function(){alternarAba(true);});
        conectar("btnCalcularSimulador","click",calcularTudo);
        conectar("btnPDFSimulador","click",gerarPDFSimulador);
        conectar("btnAdicionarHorarioSimulador","click",adicionarHorario);
        conectar("btnAdicionarPilhaFila","click",adicionarPilhaFila);

        document.querySelectorAll(".btn-cenario-simulador").forEach(function(botao){
            botao.addEventListener("click",function(){
                estado.cenario = botao.dataset.cenario;
                document.querySelectorAll(".btn-cenario-simulador").forEach(function(item){
                    item.classList.toggle("ativo",item === botao);
                });
                calcularTudo();
            });
        });

        const campos = document.querySelectorAll(
            "#painelSimuladorPilha input, #painelSimuladorPilha select"
        );

        campos.forEach(function(campo){
            if(!campo.closest(".formulario-fila-simulador") && campo.id !== "simNovoHorario"){
                campo.addEventListener("input",calcularTudo);
                campo.addEventListener("change",calcularTudo);
            }
        });

        definirDatasIniciais();
        recuperarEstadoLocal();

        if(estado.horarios.length === 0){
            estado.horarios = criarHorariosPadrao();
        }

        calcularTudo();
    }

    function conectar(id,evento,funcao){
        const elemento = document.getElementById(id);
        if(elemento){elemento.addEventListener(evento,funcao);}
    }

    function alternarAba(simuladorAtivo){
        const container = document.querySelector(".container-planejamento-pilha");
        const abaPlanejamento = document.getElementById("abaPlanejamentoCompleto");
        const abaSimulador = document.getElementById("abaSimuladorPilha");

        if(!container || !abaPlanejamento || !abaSimulador){return;}

        container.classList.toggle("modo-simulador-pilha",simuladorAtivo);
        abaPlanejamento.classList.toggle("ativo",!simuladorAtivo);
        abaSimulador.classList.toggle("ativo",simuladorAtivo);
        abaPlanejamento.setAttribute("aria-selected",String(!simuladorAtivo));
        abaSimulador.setAttribute("aria-selected",String(simuladorAtivo));

        if(simuladorAtivo){calcularTudo();}
    }

    function definirDatasIniciais(){
        const agora = new Date();
        agora.setMinutes(0,0,0);

        const alvo = new Date(agora.getTime() + 12 * 3600000);

        definirValorSeVazio("simDataBaseFormacao",dataParaInput(agora));
        definirValorSeVazio("simDataBaseRetomada",dataParaInput(agora));
        definirValorSeVazio("simDataAlvoFormacao",dataParaInput(alvo));
        definirValorSeVazio("simDataAlvoRetomada",dataParaInput(alvo));
        definirValorSeVazio("simNovoHorario",dataParaInput(alvo));
        definirValorSeVazio("filaInicioPilha",dataParaInput(alvo));
    }

    function definirValorSeVazio(id,valor){
        const elemento = document.getElementById(id);
        if(elemento && !elemento.value){elemento.value = valor;}
    }

    function calcularTudo(){
        const alvoFormacao = obterData("simDataAlvoFormacao");
        const alvoRetomada = obterData("simDataAlvoRetomada");

        const formacao = calcularFormacao(alvoFormacao);
        const retomada = calcularRetomada(alvoRetomada);

        preencherResultadoPrincipal(formacao,retomada);
        atualizarDesenhos(formacao,retomada);
        atualizarTabelaProjecoes();
        atualizarFilaPilhas();
        salvarEstadoLocal();
    }

    function calcularFormacao(alvo){
        const base = obterData("simDataBaseFormacao");
        const massaAtual = numero("simMassaAtual");
        const massaPlanejada = numero("simMassaPlanejada");
        const produtividadeInformada = numero("simProdutividade");
        const paradaInformada = numero("simParadaBritagem");
        const fator = fatores[estado.cenario];
        const produtividade = produtividadeInformada * fator.producao;
        const parada = paradaInformada * fator.parada;
        const horas = horasEntre(base,alvo);
        const horasEfetivas = Math.max(0,horas - Math.min(parada,horas));
        const massaPossivel = massaAtual + horasEfetivas * produtividade;
        const massaProjetada = massaPlanejada > 0
            ? Math.min(massaPlanejada,massaPossivel)
            : massaPossivel;
        const restante = Math.max(0,massaPlanejada - massaProjetada);
        const percentual = massaPlanejada > 0
            ? limitar(massaProjetada / massaPlanejada * 100,0,100)
            : 0;

        let fim = null;
        if(base && produtividade > 0 && massaPlanejada > 0){
            const massaFaltanteInicial = Math.max(0,massaPlanejada - massaAtual);
            fim = adicionarHoras(base,massaFaltanteInicial / produtividade + parada);
        }

        return {base,alvo,massaAtual,massaPlanejada,produtividade,parada,massaProjetada,restante,percentual,fim};
    }

    function calcularRetomada(alvo){
        const base = obterData("simDataBaseRetomada");
        const comprimento = Math.max(1,numero("simComprimentoPilha"));
        const posicaoAtual = limitar(numero("simPosicaoAtual"),0,comprimento);
        const ritmoInformado = numero("simRitmoRetomador");
        const paradaInformada = numero("simParadaUsina");
        const sentido = valor("simSentidoRetomada") || "crescente";
        const fator = fatores[estado.cenario];
        const ritmo = ritmoInformado * fator.ritmo;
        const parada = paradaInformada * fator.parada;
        const horas = horasEntre(base,alvo);
        const horasEfetivas = Math.max(0,horas - Math.min(parada,horas));
        const avanco = horasEfetivas * ritmo;

        let posicao = sentido === "crescente"
            ? posicaoAtual + avanco
            : posicaoAtual - avanco;

        posicao = limitar(posicao,0,comprimento);

        const metrosRetomados = sentido === "crescente"
            ? posicao
            : comprimento - posicao;
        const restante = Math.max(0,comprimento - metrosRetomados);
        const percentual = limitar(metrosRetomados / comprimento * 100,0,100);
        const distanciaInicial = sentido === "crescente"
            ? comprimento - posicaoAtual
            : posicaoAtual;

        let fim = null;
        if(base && ritmo > 0){
            fim = adicionarHoras(base,Math.max(0,distanciaInicial) / ritmo + parada);
        }

        return {base,alvo,posicaoAtual,posicao,comprimento,ritmo,parada,sentido,restante,percentual,fim};
    }

    function preencherResultadoPrincipal(formacao,retomada){
        texto("simMassaProjetada",formatarNumero(formacao.massaProjetada,0));
        texto("simPercentualFormacao",formatarNumero(formacao.percentual,1));
        texto("simMassaRestante",formatarNumero(formacao.restante,0));
        texto("simFimFormacao",formatarDataHora(formacao.fim));

        texto("simPosicaoProjetada",formatarNumero(retomada.posicao,1));
        texto("simPercentualRetomado",formatarNumero(retomada.percentual,1));
        texto("simMetrosRestantes",formatarNumero(retomada.restante,1));
        texto("simFimRetomada",formatarDataHora(retomada.fim));
    }

    function atualizarDesenhos(formacao,retomada){
        const larguraTotal = 660;
        const inicioRegua = 40;

        const progresso = document.getElementById("simProgressoPilha");
        if(progresso){
            progresso.setAttribute("x",String(inicioRegua));
            progresso.setAttribute("width",String(larguraTotal * formacao.percentual / 100));
        }

        texto("simTextoPilhaSVG",formatarNumero(formacao.percentual,1) + "% FORMADA");

        const proporcaoPosicao = limitar(retomada.posicao / retomada.comprimento,0,1);
        const xMarcador = inicioRegua + larguraTotal * proporcaoPosicao;
        const marcador = document.getElementById("simMarcadorRetomador");
        if(marcador){
            marcador.setAttribute("transform","translate(" + xMarcador + " 0)");
        }

        texto("simTextoRetomadorSVG",formatarNumero(retomada.posicao,0) + " m");

        const trecho = document.getElementById("simTrechoRetomado");
        if(trecho){
            if(retomada.sentido === "crescente"){
                trecho.setAttribute("x",String(inicioRegua));
                trecho.setAttribute("width",String(larguraTotal * proporcaoPosicao));
            }else{
                trecho.setAttribute("x",String(xMarcador));
                trecho.setAttribute("width",String(Math.max(0,inicioRegua + larguraTotal - xMarcador)));
            }
        }
    }

    function criarHorariosPadrao(){
        const base = obterData("simDataBaseFormacao") || new Date();
        return [7,13,19,1].map(function(hora){
            const data = new Date(base);
            data.setHours(hora,0,0,0);
            if(data <= base){data.setDate(data.getDate() + 1);}
            return data.toISOString();
        }).sort();
    }

    function adicionarHorario(){
        const data = obterData("simNovoHorario");
        if(!data){alert("Informe uma data e um horário para a projeção.");return;}

        const iso = data.toISOString();
        if(!estado.horarios.includes(iso)){estado.horarios.push(iso);}
        estado.horarios.sort();
        atualizarTabelaProjecoes();
        salvarEstadoLocal();
    }

    function atualizarTabelaProjecoes(){
        const corpo = document.getElementById("simTabelaProjecoes");
        if(!corpo){return;}
        corpo.innerHTML = "";

        const formacaoFinal = calcularFormacao(obterData("simDataAlvoFormacao"));
        const retomadaFinal = calcularRetomada(obterData("simDataAlvoRetomada"));

        estado.horarios.forEach(function(iso,indice){
            const alvo = new Date(iso);
            const formacao = calcularFormacao(alvo);
            const retomada = calcularRetomada(alvo);
            const situacao = classificarSituacao(formacaoFinal.fim,retomadaFinal.fim);
            const linha = document.createElement("tr");
            linha.innerHTML = `
                <td>${formatarDataHora(alvo)}</td>
                <td><strong>${formatarNumero(formacao.massaProjetada,0)} t</strong></td>
                <td>${formatarNumero(formacao.percentual,1)}%</td>
                <td>${formatarNumero(retomada.posicao,1)} m</td>
                <td>${formatarNumero(retomada.restante,1)} m</td>
                <td><span class="status-simulador ${situacao.classe}">${situacao.texto}</span></td>
                <td><button type="button" class="btn-excluir-simulador" data-indice-horario="${indice}" aria-label="Excluir horário">×</button></td>
            `;
            corpo.appendChild(linha);
        });

        corpo.querySelectorAll("[data-indice-horario]").forEach(function(botao){
            botao.addEventListener("click",function(){
                estado.horarios.splice(Number(botao.dataset.indiceHorario),1);
                atualizarTabelaProjecoes();
                salvarEstadoLocal();
            });
        });
    }

    function classificarSituacao(fimFormacao,fimRetomada){
        if(!fimFormacao || !fimRetomada){return {texto:"Pendente",classe:"neutro"};}
        const diferencaHoras = (fimFormacao - fimRetomada) / 3600000;
        if(diferencaHoras <= 0){return {texto:"No plano",classe:"ok"};}
        if(diferencaHoras <= 6){return {texto:"Atenção",classe:"atencao"};}
        return {texto:"Risco",classe:"critico"};
    }

    function adicionarPilhaFila(){
        const nome = valor("filaNomePilha").trim();
        const lado = valor("filaLadoPilha");
        const massa = numero("filaMassaPilha");
        const produtividade = numero("filaProdutividadePilha");
        const manutencao = numero("filaManutencaoPilha");

        if(!nome){alert("Informe o nome ou número da pilha.");return;}
        if(massa <= 0 || produtividade <= 0){alert("Informe massa e produtividade maiores que zero.");return;}

        estado.fila.push({nome,lado,massa,produtividade,manutencao});
        document.getElementById("filaNomePilha").value = "";
        atualizarFilaPilhas();
        salvarEstadoLocal();
    }

    function atualizarFilaPilhas(){
        const corpo = document.getElementById("simTabelaFilaPilhas");
        if(!corpo){return;}
        corpo.innerHTML = "";

        let inicio = obterData("filaInicioPilha") || obterData("simDataAlvoFormacao") || new Date();
        const fimRetomada = calcularRetomada(obterData("simDataAlvoRetomada")).fim;
        const ladoRetomada = valor("simLadoRetomada");
        const fator = fatores[estado.cenario];
        let quantidadeAtencao = 0;

        estado.fila.forEach(function(item,indice){
            const produtividade = item.produtividade * fator.producao;
            const manutencao = item.manutencao * fator.parada;
            const fim = adicionarHoras(inicio,item.massa / produtividade + manutencao);
            const agora = new Date();
            let status = {texto:"Planejada",classe:"neutro"};

            if(fim <= agora){status = {texto:"Concluída",classe:"ok"};}
            else if(inicio <= agora && fim > agora){status = {texto:"Em formação",classe:"ok"};}

            if(item.lado === ladoRetomada && fimRetomada && inicio < fimRetomada){
                status = {texto:"Atenção",classe:"atencao"};
                quantidadeAtencao++;
            }

            const linha = document.createElement("tr");
            linha.innerHTML = `
                <td>${indice + 1}</td>
                <td><strong>${escaparHTML(item.nome)}</strong></td>
                <td><span class="lado-pilha-simulador">${item.lado}</span></td>
                <td>${formatarNumero(item.massa,0)} t</td>
                <td>${formatarNumero(produtividade,0)} t/h</td>
                <td>${formatarDataHora(inicio)}</td>
                <td>${formatarDataHora(fim)}</td>
                <td><span class="status-simulador ${status.classe}">${status.texto}</span></td>
                <td><button type="button" class="btn-excluir-simulador" data-indice-fila="${indice}" aria-label="Excluir pilha">×</button></td>
            `;
            corpo.appendChild(linha);
            inicio = fim;
        });

        corpo.querySelectorAll("[data-indice-fila]").forEach(function(botao){
            botao.addEventListener("click",function(){
                estado.fila.splice(Number(botao.dataset.indiceFila),1);
                atualizarFilaPilhas();
                salvarEstadoLocal();
            });
        });

        const resumo = document.getElementById("simResumoFila");
        if(resumo){
            if(estado.fila.length === 0){
                resumo.textContent = "Adicione pilhas para iniciar a simulação da sequência do pátio.";
                resumo.className = "resumo-fila-simulador";
            }else if(quantidadeAtencao > 0){
                resumo.textContent = quantidadeAtencao + " pilha(s) apresentam conflito com a retomada prevista no mesmo lado.";
                resumo.className = "resumo-fila-simulador atencao";
            }else{
                resumo.textContent = "Sequência calculada sem conflito com a retomada informada.";
                resumo.className = "resumo-fila-simulador ok";
            }
        }
    }

    function gerarPDFSimulador(){
        calcularTudo();

        const tituloAnterior = document.title;
        document.title = "Simulador_de_Pilhas_FVO_" + dataArquivo(new Date());
        document.body.classList.add("modo-impressao-simulador");

        let restaurado = false;

        function restaurar(){
            if(restaurado){return;}
            restaurado = true;
            document.body.classList.remove("modo-impressao-simulador");
            document.title = tituloAnterior;
            window.removeEventListener("afterprint",restaurar);
        }

        window.addEventListener("afterprint",restaurar);

        requestAnimationFrame(function(){
            requestAnimationFrame(function(){
                setTimeout(function(){
                    window.print();
                    setTimeout(restaurar,1800);
                },180);
            });
        });
    }

    function salvarEstadoLocal(){
        try{
            localStorage.setItem("fvoSimuladorPilha",JSON.stringify({
                cenario:estado.cenario,
                horarios:estado.horarios,
                fila:estado.fila
            }));
        }catch(erro){/* armazenamento indisponível */}
    }

    function recuperarEstadoLocal(){
        try{
            const salvo = JSON.parse(localStorage.getItem("fvoSimuladorPilha") || "null");
            if(!salvo){return;}
            if(fatores[salvo.cenario]){estado.cenario = salvo.cenario;}
            if(Array.isArray(salvo.horarios)){estado.horarios = salvo.horarios;}
            if(Array.isArray(salvo.fila)){estado.fila = salvo.fila;}

            document.querySelectorAll(".btn-cenario-simulador").forEach(function(botao){
                botao.classList.toggle("ativo",botao.dataset.cenario === estado.cenario);
            });
        }catch(erro){/* estado anterior inválido */}
    }

    function numero(id){
        const n = Number(valor(id).replace(",","."));
        return Number.isFinite(n) ? n : 0;
    }

    function valor(id){
        const elemento = document.getElementById(id);
        return elemento ? String(elemento.value || "") : "";
    }

    function texto(id,conteudo){
        const elemento = document.getElementById(id);
        if(elemento){elemento.textContent = conteudo;}
    }

    function obterData(id){
        const v = valor(id);
        if(!v){return null;}
        const data = new Date(v);
        return Number.isNaN(data.getTime()) ? null : data;
    }

    function horasEntre(inicio,fim){
        if(!inicio || !fim || fim <= inicio){return 0;}
        return (fim - inicio) / 3600000;
    }

    function adicionarHoras(data,horas){
        if(!data || !Number.isFinite(horas)){return null;}
        return new Date(data.getTime() + horas * 3600000);
    }

    function limitar(valor,minimo,maximo){
        return Math.min(maximo,Math.max(minimo,valor));
    }

    function formatarNumero(valor,casas){
        if(!Number.isFinite(valor)){return "-";}
        return valor.toLocaleString("pt-BR",{
            minimumFractionDigits:casas,
            maximumFractionDigits:casas
        });
    }

    function formatarDataHora(data){
        if(!data || Number.isNaN(data.getTime())){return "-";}
        return data.toLocaleString("pt-BR",{
            day:"2-digit",month:"2-digit",year:"2-digit",
            hour:"2-digit",minute:"2-digit"
        });
    }

    function dataParaInput(data){
        const local = new Date(data.getTime() - data.getTimezoneOffset() * 60000);
        return local.toISOString().slice(0,16);
    }

    function dataArquivo(data){
        return data.toISOString().slice(0,16).replace(/[-:T]/g,"");
    }

    function escaparHTML(texto){
        return String(texto).replace(/[&<>'"]/g,function(caractere){
            return ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"})[caractere];
        });
    }
})();