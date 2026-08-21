(function configurarPortalBritagem() {
  "use strict";

  document.addEventListener("DOMContentLoaded", () => {
    document.body.classList.add("portal-britagem");
    document.querySelectorAll(".portal-opcao").forEach((link) => {
      const ehBritagem = link.textContent.toLowerCase().includes("britagem");
      link.classList.toggle("portal-opcao-ativa", ehBritagem);
      if (ehBritagem) link.setAttribute("aria-current", "page");
    });
  });
})();