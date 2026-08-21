(function configurarPortalUsina() {
  "use strict";

  document.addEventListener("DOMContentLoaded", () => {
    document.body.classList.add("portal-usina");

    document.querySelectorAll(".portal-opcao").forEach((link) => {
      const ehUsina = link.textContent.toLowerCase().includes("usina");
      link.classList.toggle("portal-opcao-ativa", ehUsina);
      if (ehUsina) link.setAttribute("aria-current", "page");
    });

    // Recuperação Metalúrgica e Rendimento Mássico passam a fazer parte
    // da página de Qualidade. Remove as antigas opções separadas dos menus.
    const itensIncorporadosQualidade = [
      "recuperação metalúrgica",
      "rendimento mássico"
    ];

    document.querySelectorAll(".menu-usina .submenu li").forEach((item) => {
      const texto = item.textContent.trim().toLowerCase();
      if (itensIncorporadosQualidade.includes(texto)) {
        item.remove();
      }
    });
  });
})();