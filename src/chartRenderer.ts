import { Canvas } from "skia-canvas";
import { ChartSection, ChartType } from "./types";

// ── Colour palette ────────────────────────────────────────────
const PALETTE = [
  "#1B3A6B", "#2E6DB4", "#4A9FD4",
  "#E8522A", "#F5A623", "#4CAF50",
  "#9C27B0", "#00BCD4", "#FF5722",
  "#607D8B", "#795548", "#009688",
];

function hexToRgba(hex: string, alpha = 1): string {
  const h = hex.replace("#", "");
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function lighten(hex: string, amount = 0.4): string {
  const h = hex.replace("#", "");
  const r = Math.min(255, Math.round(parseInt(h.slice(0, 2), 16) + (255 - parseInt(h.slice(0, 2), 16)) * amount));
  const g = Math.min(255, Math.round(parseInt(h.slice(2, 4), 16) + (255 - parseInt(h.slice(2, 4), 16)) * amount));
  const b = Math.min(255, Math.round(parseInt(h.slice(4, 6), 16) + (255 - parseInt(h.slice(4, 6), 16)) * amount));
  return `#${r.toString(16).padStart(2,"0")}${g.toString(16).padStart(2,"0")}${b.toString(16).padStart(2,"0")}`;
}

/**
 * Renders a ChartSection to a PNG Buffer using skia-canvas + Chart.js
 */
export async function renderChartToBuffer(
  section: ChartSection,
  widthPx = 620,
  heightPx = 360
): Promise<Buffer> {
  const {
    Chart,
    CategoryScale,
    LinearScale,
    BarElement,
    LineElement,
    PointElement,
    ArcElement,
    Title,
    Tooltip,
    Legend,
    Filler,
  } = await import("chart.js");

  Chart.register(
    CategoryScale, LinearScale,
    BarElement, LineElement, PointElement,
    ArcElement, Title, Tooltip, Legend, Filler
  );

  const canvas = new Canvas(widthPx, heightPx);
  const ctx = canvas.getContext("2d");

  // White background
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, widthPx, heightPx);

  const isPie       = section.chartType === "pie" || section.chartType === "doughnut";
  const isStacked   = section.chartType === "stackedBar";
  const isHorizontal = section.chartType === "horizontalBar";
  const isArea      = section.chartType === "area";
  const isLine      = section.chartType === "line";

  // ── Map to Chart.js type ──────────────────────────────────
  let chartJsType: string = section.chartType as string;
  if (isStacked || isHorizontal) chartJsType = "bar";
  if (isArea) chartJsType = "line";

  // ── Build datasets ────────────────────────────────────────
  const datasets = section.datasets.map((ds, i) => {
    const color = ds.color ?? PALETTE[i % PALETTE.length];

    if (isPie) {
      const bgColors = section.labels.map((_, li) =>
        hexToRgba(PALETTE[li % PALETTE.length], 0.88)
      );
      return {
        label: ds.label,
        data: ds.data,
        backgroundColor: bgColors,
        borderColor: "#FFFFFF",
        borderWidth: 2,
        hoverOffset: 8,
      };
    }

    if (isLine || isArea) {
      const fill = isArea ? (ds.fill !== false) : false;
      return {
        label: ds.label,
        data: ds.data,
        borderColor: color,
        backgroundColor: fill ? hexToRgba(color, 0.12) : "transparent",
        fill,
        tension: 0.35,
        pointRadius: 4,
        pointBackgroundColor: color,
        pointBorderColor: "#FFFFFF",
        pointBorderWidth: 1.5,
        borderWidth: 2.5,
      };
    }

    // bar / stacked / horizontal
    return {
      label: ds.label,
      data: ds.data,
      backgroundColor: hexToRgba(color, 0.85),
      borderColor: color,
      borderWidth: 0,
      borderRadius: isStacked ? 0 : 3,
      borderSkipped: false,
    };
  });

  // ── Scales config ─────────────────────────────────────────
  const tickStyle = { font: { size: 10, family: "Arial" }, color: "#666666" };
  const gridStyle = { color: "#F0F0F0", lineWidth: 1 };

  let scales: any = {};
  if (!isPie) {
    const yAxisOpts: any = {
      ticks: {
        ...tickStyle,
        callback: (v: number) =>
          `${section.yAxisPrefix ?? ""}${v}${section.yAxisSuffix ?? ""}`,
      },
      grid: gridStyle,
      beginAtZero: true,
      title: section.yAxisLabel
        ? { display: true, text: section.yAxisLabel, font: { size: 10, family: "Arial" }, color: "#888888" }
        : { display: false },
    };

    const xAxisOpts: any = {
      ticks: { ...tickStyle, maxRotation: section.labels.length > 8 ? 45 : 0 },
      grid: { color: "transparent" },
      title: section.xAxisLabel
        ? { display: true, text: section.xAxisLabel, font: { size: 10, family: "Arial" }, color: "#888888" }
        : { display: false },
    };

    if (isStacked) {
      yAxisOpts.stacked = true;
      xAxisOpts.stacked = true;
    }

    if (isHorizontal) {
      scales = { x: yAxisOpts, y: xAxisOpts };
    } else {
      scales = { x: xAxisOpts, y: yAxisOpts };
    }
  }

  // ── Data labels plugin (basic, inline) ───────────────────
  // We draw value labels manually after chart render
  const showValues = section.showValues === true;

  const config: any = {
    type: chartJsType,
    data: { labels: section.labels, datasets },
    options: {
      indexAxis: isHorizontal ? "y" : "x",
      responsive: false,
      animation: false,
      plugins: {
        legend: {
          display: section.showLegend !== false && (datasets.length > 1 || isPie),
          position: isPie ? "right" : "bottom",
          labels: {
            font: { size: 10, family: "Arial" },
            padding: 14,
            boxWidth: 12,
            boxHeight: 12,
            color: "#444444",
          },
        },
        title: { display: false },
        tooltip: { enabled: false },
      },
      scales,
      layout: { padding: { top: 16, bottom: 16, left: 12, right: 24 } },
    },
  };

  const chart = new Chart(ctx as any, config);

  // Give Chart.js a tick to render
  await new Promise((r) => setTimeout(r, 80));

  // ── Draw value labels on bar/line charts ─────────────────
  if (showValues && !isPie) {
    ctx.font = "bold 9px Arial";
    ctx.fillStyle = "#333333";
    ctx.textAlign = "center";

    chart.data.datasets.forEach((ds: any, di: number) => {
      const meta = chart.getDatasetMeta(di);
      meta.data.forEach((bar: any, bi: number) => {
        const val = ds.data[bi];
        if (val == null) return;
        const label = `${section.yAxisPrefix ?? ""}${val}${section.yAxisSuffix ?? ""}`;
        const x = bar.x;
        const y = bar.y - 5;
        ctx.fillText(label, x, y);
      });
    });
  }

  return canvas.toBuffer("png");
}
