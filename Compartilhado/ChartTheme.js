(function configurarGraficosFVO() {
  "use strict";

  if (typeof Chart === "undefined") return;

  Chart.defaults.color = "#514a55";
  Chart.defaults.font.family = "Arial, Helvetica, sans-serif";
  Chart.defaults.font.size = 12;
  Chart.defaults.font.weight = "600";
  Chart.defaults.devicePixelRatio = Math.min(window.devicePixelRatio || 1, 2);
  Chart.defaults.animation.duration = 300;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 16;
  Chart.defaults.plugins.legend.labels.boxHeight = 3;
  Chart.defaults.plugins.legend.labels.padding = 16;
  Chart.defaults.plugins.tooltip.padding = 11;

  const camadasFVO = {
    id: "camadasFVO",
    beforeInit(chart) {
      (chart.data.datasets || []).forEach((dataset) => {
        if (dataset.type === "line") {
          dataset.order = dataset.order ?? 1;
          dataset.borderWidth = dataset.borderWidth || 2.5;
          dataset.pointRadius = dataset.pointRadius ?? 2;
          dataset.pointHoverRadius = dataset.pointHoverRadius ?? 5;
          dataset.tension = dataset.tension ?? 0.2;
        } else {
          dataset.order = dataset.order ?? 2;
          dataset.borderRadius = dataset.borderRadius ?? 4;
          dataset.maxBarThickness = dataset.maxBarThickness ?? 48;
        }
      });
    }
  };

  Chart.register(camadasFVO);

  function atualizarImpressao(ativar) {
    Chart.defaults.font.size = ativar ? 15 : 12;
    Object.values(Chart.instances || {}).forEach((grafico) => {
      grafico.options.animation = false;
      grafico.resize();
      grafico.update("none");
    });
  }

  window.addEventListener("beforeprint", () => atualizarImpressao(true));
  window.addEventListener("afterprint", () => atualizarImpressao(false));
})();
