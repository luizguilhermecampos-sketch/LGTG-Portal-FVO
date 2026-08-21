(function iniciarPortalFVO() {
  "use strict";

  function fecharMenus(excecao) {
    document.querySelectorAll(".menu-portal .dropdown-aberto").forEach((item) => {
      if (item !== excecao) {
        item.classList.remove("dropdown-aberto");
        item.querySelector(".dropdown-toggle")?.setAttribute("aria-expanded", "false");
      }
    });
  }

  function configurarMenus() {
    document.querySelectorAll(".menu-portal .dropdown").forEach((dropdown) => {
      const botao = dropdown.querySelector(".dropdown-toggle");
      const submenu = dropdown.querySelector(".submenu");
      if (!botao || !submenu) return;

      botao.setAttribute("aria-haspopup", "true");
      botao.setAttribute("aria-expanded", "false");

      botao.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        const abrir = !dropdown.classList.contains("dropdown-aberto");
        fecharMenus(dropdown);
        dropdown.classList.toggle("dropdown-aberto", abrir);
        botao.setAttribute("aria-expanded", String(abrir));
      });

      submenu.addEventListener("click", (event) => event.stopPropagation());
    });

    document.addEventListener("click", () => fecharMenus());
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") fecharMenus();
    });
  }

  function configurarLinksIndisponiveis() {
    document.querySelectorAll('a[href="#"]:not(.dropdown-toggle)').forEach((link) => {
      link.setAttribute("aria-disabled", "true");
      link.setAttribute("title", link.title || "Página em preparação");
      link.addEventListener("click", (event) => event.preventDefault());
    });
  }

  function configurarVideo() {
    const video = document.querySelector(".hero-video");
    if (!video) return;
    video.muted = true;
    video.loop = true;
    video.playsInline = true;
    video.play()?.catch?.(() => {});
  }

  function destacarPaginaAtual() {
    const atual = new URL(window.location.href);
    document.querySelectorAll('.menu-portal a[href]:not([href="#"]), .menu-dashboard-fvo a[href]:not([href="#"])').forEach((link) => {
      const destino = new URL(link.getAttribute("href"), atual);
      if (destino.pathname !== atual.pathname) return;
      link.setAttribute("aria-current", "page");
      link.closest("li")?.classList.add("menu-secao-ativa");
      link.closest(".dropdown")?.classList.add("menu-secao-ativa");
    });
  }

  function criarCarregamentoCompartilhado() {
    let overlay = document.getElementById("carregamentoCompartilhadoFVO");
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.id = "carregamentoCompartilhadoFVO";
      overlay.className = "overlay-carregamento-fvo oculto";
      overlay.innerHTML = '<div class="caixa-carregamento-fvo"><div class="spinner-carregamento-fvo"></div><strong>Carregando dados...</strong><span id="textoCarregamentoCompartilhadoFVO">Preparando planilha</span></div>';
      document.body.appendChild(overlay);
    }

    const texto = overlay.querySelector("#textoCarregamentoCompartilhadoFVO");
    window.FVOCarregamento = {
      exibir(mensagem = "Preparando planilha") {
        if (texto) texto.textContent = mensagem;
        overlay.classList.remove("oculto");
        document.body.classList.add("carregando-dados");
      },
      ocultar() {
        overlay.classList.add("oculto");
        document.body.classList.remove("carregando-dados");
      }
    };
  }

  document.addEventListener("DOMContentLoaded", () => {
    configurarMenus();
    configurarLinksIndisponiveis();
    configurarVideo();
    destacarPaginaAtual();
    criarCarregamentoCompartilhado();
  });
})();
