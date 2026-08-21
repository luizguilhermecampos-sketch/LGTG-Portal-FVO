(function configurarPortalUsina() {
  "use strict";

  document.addEventListener("DOMContentLoaded", () => {
    document.body.classList.add("portal-usina");
    document.querySelectorAll(".portal-opcao").forEach((link) => {
      const ehUsina = link.textContent.toLowerCase().includes("usina");
      link.classList.toggle("portal-opcao-ativa", ehUsina);
      if (ehUsina) link.setAttribute("aria-current", "page");
    });
  });
})();