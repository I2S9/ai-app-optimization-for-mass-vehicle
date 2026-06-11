/**
 * Weight Tax — bar + line chart (mass kg left, malus € right).
 */

export const WT_CHART_BAR_COLOR = '#9BBB59';
export const WT_CHART_LINE_COLOR = '#C00000';

const FONT = '"Segoe UI", Arial, sans-serif';

export function truncateChartLabel(text, maxLen = 32) {
  const s = String(text || '').trim();
  if (s.length <= maxLen) return s;
  return `${s.slice(0, maxLen - 1)}…`;
}

export function massAxisRange(series) {
  const vals = series.map((s) => s.massKg).filter((n) => n != null && Number.isFinite(n));
  if (!vals.length) return { min: 1200, max: 2100 };
  let min = Math.min(...vals);
  let max = Math.max(...vals);
  min = Math.floor((min - 80) / 100) * 100;
  max = Math.ceil((max + 80) / 100) * 100;
  if (max - min < 500) max = min + 500;
  return { min: Math.max(0, min), max };
}

export function euroAxisRange(series) {
  const vals = series.map((s) => s.malusEuro ?? 0);
  const maxVal = Math.max(...vals, 0);
  if (maxVal <= 0) return { min: 0, max: 5000 };
  const max = Math.ceil(maxVal / 500) * 500;
  return { min: 0, max: Math.max(500, max) };
}

function niceTicks(min, max, count = 6) {
  const span = max - min;
  if (span <= 0) return [min, max];
  const raw = span / Math.max(1, count - 1);
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = Math.ceil(raw / mag) * mag;
  const ticks = [];
  let v = Math.floor(min / step) * step;
  while (v <= max + step * 0.01) {
    if (v >= min - step * 0.01) ticks.push(v);
    v += step;
  }
  if (ticks.length < 2) return [min, max];
  return ticks;
}

/**
 * @param {HTMLCanvasElement} canvas
 * @param {{ series: Array<{ label: string, massKg: number|null, malusEuro: number }>, title?: string, includeTitle?: boolean }} opts
 */
export function drawWeightTaxChart(canvas, opts = {}) {
  const series = Array.isArray(opts.series) ? opts.series : [];
  const dpr = opts.dpr || window.devicePixelRatio || 1;
  const includeTitle = opts.includeTitle === true;
  const title = opts.title || '';

  const rect = canvas.getBoundingClientRect();
  const cssW = Math.max(rect.width || canvas.clientWidth || 800, 480);
  const cssH = Math.max(rect.height || canvas.clientHeight || 560, 560);
  canvas.width = Math.round(cssW * dpr);
  canvas.height = Math.round(cssH * dpr);

  const ctx = canvas.getContext('2d');
  if (!ctx) return;

  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cssW, cssH);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cssW, cssH);

  const titleH = includeTitle ? 44 : 0;
  const margin = {
    top: Math.round(Math.max(32, cssH * 0.08)),
    right: Math.round(Math.max(64, cssW * 0.075)),
    bottom: Math.round(Math.max(64, cssH * 0.16)),
    left: Math.round(Math.max(56, cssW * 0.065)),
  };
  const plotTop = titleH + margin.top;
  const plotBottom = cssH - margin.bottom;
  const plotLeft = margin.left;
  const plotRight = cssW - margin.right;
  const plotW = plotRight - plotLeft;
  const plotH = plotBottom - plotTop;

  if (includeTitle && title) {
    ctx.fillStyle = '#000000';
    ctx.font = `700 14px ${FONT}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(title, cssW / 2, titleH / 2);
  }

  if (!series.length || plotW <= 0 || plotH <= 0) {
    ctx.fillStyle = '#888888';
    ctx.font = `13px ${FONT}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Sélectionnez des véhicules pour afficher le graphe.', cssW / 2, (plotTop + plotBottom) / 2);
    return;
  }

  const massRange = massAxisRange(series);
  const euroRange = euroAxisRange(series);
  const massTicks = niceTicks(massRange.min, massRange.max, 7);
  const euroTicks = niceTicks(euroRange.min, euroRange.max, 6);

  const yMass = (kg) =>
    plotBottom - ((kg - massRange.min) / (massRange.max - massRange.min)) * plotH;
  const yEuro = (eur) =>
    plotBottom - ((eur - euroRange.min) / (euroRange.max - euroRange.min)) * plotH;

  // Grid + left axis (kg)
  ctx.strokeStyle = '#d9d9d9';
  ctx.lineWidth = 1;
  ctx.fillStyle = '#000000';
  ctx.font = `11px ${FONT}`;
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  for (const tick of massTicks) {
    const y = yMass(tick);
    ctx.beginPath();
    ctx.moveTo(plotLeft, y);
    ctx.lineTo(plotRight, y);
    ctx.stroke();
    ctx.fillText(`${tick} kg`, plotLeft - 8, y);
  }

  // Right axis (€)
  ctx.fillStyle = WT_CHART_LINE_COLOR;
  ctx.textAlign = 'left';
  for (const tick of euroTicks) {
    const y = yEuro(tick);
    ctx.fillText(`${tick} €`, plotRight + 8, y);
  }

  const n = series.length;
  const slotW = plotW / n;
  const barW = Math.max(28, Math.min(slotW * 0.78, slotW - 16));
  const labelMaxLen = Math.max(18, Math.min(48, Math.floor(slotW / 7)));
  const points = [];

  series.forEach((item, i) => {
    const cx = plotLeft + slotW * i + slotW / 2;
    const mass = item.massKg;
    if (mass == null || !Number.isFinite(mass)) return;

    const yTop = yMass(mass);
    const yBase = plotBottom;
    const x0 = cx - barW / 2;

    ctx.fillStyle = WT_CHART_BAR_COLOR;
    ctx.fillRect(x0, yTop, barW, yBase - yTop);
    ctx.strokeStyle = '#7a9344';
    ctx.lineWidth = 1;
    ctx.strokeRect(x0, yTop, barW, yBase - yTop);

    // Curb mass label (black, above bar)
    ctx.fillStyle = '#000000';
    ctx.font = `700 11px ${FONT}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';
    ctx.fillText(`${Math.round(mass)} kg`, cx, yTop - 4);

    // X label (gray, below axis)
    ctx.fillStyle = '#808080';
    ctx.font = `10px ${FONT}`;
    ctx.textBaseline = 'top';
    const label = truncateChartLabel(item.label, labelMaxLen);
    ctx.fillText(label, cx, plotBottom + 10, slotW - 8);

    const malus = item.malusEuro ?? 0;
    points.push({ cx, cy: yEuro(malus), malus });
  });

  // Red malus line + markers
  if (points.length) {
    ctx.strokeStyle = WT_CHART_LINE_COLOR;
    ctx.fillStyle = WT_CHART_LINE_COLOR;
    ctx.lineWidth = 2;
    ctx.beginPath();
    points.forEach((p, idx) => {
      if (idx === 0) ctx.moveTo(p.cx, p.cy);
      else ctx.lineTo(p.cx, p.cy);
    });
    ctx.stroke();

    points.forEach((p) => {
      ctx.beginPath();
      ctx.arc(p.cx, p.cy, 4, 0, Math.PI * 2);
      ctx.fill();
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 1.5;
      ctx.stroke();
      ctx.strokeStyle = WT_CHART_LINE_COLOR;

      ctx.fillStyle = WT_CHART_LINE_COLOR;
      ctx.font = `700 11px ${FONT}`;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'bottom';
      const rounded = Math.round(p.malus * 100) / 100;
      const txt = Number.isInteger(rounded) ? String(rounded) : String(rounded);
      ctx.fillText(`${txt} €`, p.cx, p.cy - 8);
    });
  }

  // Plot border
  ctx.strokeStyle = '#000000';
  ctx.lineWidth = 1;
  ctx.strokeRect(plotLeft, plotTop, plotW, plotH);
}

/**
 * Export chart zone (title bar + canvas) to PNG data URL.
 */
export function exportWeightTaxChartPng({ zoneEl, canvasEl, title, variant }) {
  if (!zoneEl || !canvasEl) return null;
  const headerH = 64;
  const titleFontSize = 24;
  const plotEl = zoneEl.querySelector('.wt-chart-plot') || canvasEl;
  const width = Math.max(Math.round(zoneEl.getBoundingClientRect().width), 1);
  const chartH = Math.max(Math.round(plotEl.getBoundingClientRect().height), 1);
  const height = headerH + chartH;
  const dpr = 2;

  const out = document.createElement('canvas');
  out.width = width * dpr;
  out.height = height * dpr;
  const ctx = out.getContext('2d');
  if (!ctx) return null;

  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = '#f3f3f3';
  ctx.fillRect(0, 0, width, headerH);
  ctx.strokeStyle = '#d9d9d9';
  ctx.beginPath();
  ctx.moveTo(0, headerH);
  ctx.lineTo(width, headerH);
  ctx.stroke();

  ctx.fillStyle = '#000000';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.font = `700 ${titleFontSize}px ${FONT}`;
  ctx.fillText(title || '', width / 2, headerH / 2);

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, headerH, width, chartH);

  const snap = document.createElement('canvas');
  snap.width = canvasEl.width;
  snap.height = canvasEl.height;
  const snapCtx = snap.getContext('2d');
  if (snapCtx) snapCtx.drawImage(canvasEl, 0, 0);
  ctx.drawImage(snap, 0, headerH, width, chartH);

  ctx.strokeStyle = '#000000';
  ctx.lineWidth = 2;
  ctx.strokeRect(1, 1, width - 2, height - 2);

  const titlePart = String(title || '')
    .replace(/\s+/g, '-')
    .replace(/[^\w-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
  const filename = `weight-tax-${variant || 'bev'}${titlePart ? `-${titlePart}` : ''}.png`;
  return { dataUrl: out.toDataURL('image/png'), filename };
}
