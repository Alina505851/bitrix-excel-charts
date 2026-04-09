import { Chart, registerables } from "chart.js";
import {
  CHART_EXPORT_COLOR_CYCLE,
  chartPaletteForEntity,
} from "@/lib/chart/chartPalettes";
import type { EntityBlockId } from "@/lib/predefinedCharts";
import type { ChartType } from "@/lib/types";

Chart.register(...registerables);

function stripDataUrlPrefix(dataUrl: string): string {
  const i = dataUrl.indexOf(",");
  return i >= 0 ? dataUrl.slice(i + 1) : dataUrl;
}

function hexToRgb(hex: string): { r: number; g: number; b: number } | null {
  const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex.trim());
  if (!m) {
    return null;
  }
  return {
    r: parseInt(m[1]!, 16),
    g: parseInt(m[2]!, 16),
    b: parseInt(m[3]!, 16),
  };
}

function rgbaFromHex(hex: string, alpha: number): string {
  const rgb = hexToRgb(hex);
  if (!rgb) {
    return hex;
  }
  return `rgba(${rgb.r},${rgb.g},${rgb.b},${alpha})`;
}

function barGradient(
  ctx: CanvasRenderingContext2D,
  height: number,
  hex: string,
): CanvasGradient {
  const g = ctx.createLinearGradient(0, 0, 0, height);
  g.addColorStop(0, rgbaFromHex(hex, 1));
  g.addColorStop(0.55, rgbaFromHex(hex, 0.92));
  g.addColorStop(1, rgbaFromHex(hex, 0.62));
  return g;
}

function areaGradient(
  ctx: CanvasRenderingContext2D,
  height: number,
  hex: string,
): CanvasGradient {
  const g = ctx.createLinearGradient(0, 0, 0, height);
  g.addColorStop(0, rgbaFromHex(hex, 0.55));
  g.addColorStop(0.45, rgbaFromHex(hex, 0.22));
  g.addColorStop(1, rgbaFromHex(hex, 0.02));
  return g;
}

/**
 * Растеризация тех же данных, что на графике в UI, в PNG (base64 без префикса) для вставки в Excel.
 * В среде без DOM (тесты) возвращает null.
 */
export async function renderAggregatedChartPng(options: {
  chartType: ChartType;
  labels: string[];
  series: { label: string; values: number[] }[];
  /** Палитра как в UI по разделу CRM */
  entity?: EntityBlockId;
}): Promise<string | null> {
  if (typeof document === "undefined") {
    return null;
  }

  const { chartType, labels, series, entity } = options;
  if (labels.length === 0 || series.length === 0) {
    return null;
  }

  const canvas = document.createElement("canvas");
  canvas.width = 920;
  canvas.height = 500;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return null;
  }

  const colors = entity
    ? chartPaletteForEntity(entity)
    : CHART_EXPORT_COLOR_CYCLE;

  try {
    if (chartType === "pie") {
      const s0 = series[0];
      const ch = new Chart(ctx, {
        type: "pie",
        data: {
          labels,
          datasets: [
            {
              label: s0?.label ?? "",
              data: s0?.values ?? [],
              backgroundColor: labels.map(
                (_, i) => colors[i % colors.length]!,
              ),
              borderColor: "#ffffff",
              borderWidth: 2,
              hoverBorderColor: "#ffffff",
            },
          ],
        },
        options: {
          responsive: false,
          animation: false,
          plugins: {
            legend: {
              position: "right",
              labels: {
                boxWidth: 12,
                padding: 10,
                font: { size: 11, weight: "bold" },
              },
            },
          },
        },
      });
      ch.update();
      const out = stripDataUrlPrefix(canvas.toDataURL("image/png"));
      ch.destroy();
      return out;
    }

    const h = canvas.height;

    const datasets = series.map((s, i) => {
      const c = colors[i % colors.length]!;
      if (chartType === "bar") {
        return {
          label: s.label,
          data: s.values,
          borderColor: "rgba(255,255,255,0.95)",
          borderWidth: 1.5,
          borderSkipped: false,
          borderRadius: 6,
          backgroundColor: barGradient(ctx, h, c),
        };
      }
      if (chartType === "area") {
        return {
          label: s.label,
          data: s.values,
          borderColor: c,
          backgroundColor: areaGradient(ctx, h, c),
          fill: true,
          tension: 0.25,
          pointRadius: 4,
          pointHoverRadius: 6,
          pointBackgroundColor: c,
          pointBorderColor: "#ffffff",
          pointBorderWidth: 2,
        };
      }
      return {
        label: s.label,
        data: s.values,
        borderColor: c,
        backgroundColor: "transparent",
        fill: false,
        tension: 0.25,
        borderWidth: 3,
        pointRadius: 4,
        pointHoverRadius: 6,
        pointBackgroundColor: c,
        pointBorderColor: "#ffffff",
        pointBorderWidth: 2,
      };
    });

    const type = chartType === "bar" ? "bar" : "line";

    const ch = new Chart(ctx, {
      type,
      data: { labels, datasets },
      options: {
        responsive: false,
        animation: false,
        interaction: { mode: "index", intersect: false },
        scales: {
          x: {
            grid: {
              color: "rgba(148, 163, 184, 0.2)",
            },
            ticks: {
              maxRotation: labels.length > 10 ? 48 : 0,
              autoSkip: labels.length > 36,
              maxTicksLimit: 28,
              color: "#475569",
              font: { size: 11 },
            },
          },
          y: {
            beginAtZero: true,
            grid: {
              color: "rgba(148, 163, 184, 0.22)",
            },
            ticks: {
              color: "#475569",
              font: { size: 11 },
            },
          },
        },
        plugins: {
          legend: {
            position: "top",
            labels: {
              boxWidth: 14,
              padding: 14,
              font: { size: 12, weight: "bold" },
              usePointStyle: true,
            },
          },
        },
      },
    });
    ch.update();
    const out = stripDataUrlPrefix(canvas.toDataURL("image/png"));
    ch.destroy();
    return out;
  } catch {
    return null;
  }
}
