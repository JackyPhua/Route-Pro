import { useState, useMemo, useEffect, useRef, useLayoutEffect, useCallback } from 'react';
import { Upload } from 'lucide-react';
import * as XLSX from 'xlsx';
import packageJson from '../package.json';
import routeProLogo from './assets/hero.png';
import {
  saveCachedSpreadsheet,
  loadCachedSpreadsheet,
  clearCachedSpreadsheet,
  patchCachedSpreadsheetSheetIndex,
} from './importWorkbookCache.js';
import { sabahHolidayNameForCalendarDate, advanceIsoPastSundaysAndPublicHolidays } from './malaysiaSabahCalendarHolidays.js';

const MONTH_SHORT = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

const WEEKDAY_ABBREV_SUN0 = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

const YEAR_RANGE = Array.from({ length: 2032 - 2023 + 1 }, (_, i) => 2023 + i);

const SELECT_CHEVRON = `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E")`;

/** Native <select> for dark toolbar */
const selectToolbarStyle = {
  padding: '6px 28px 6px 10px',
  borderRadius: '6px',
  border: '1px solid rgba(255,255,255,0.25)',
  background: '#334155',
  color: 'white',
  fontSize: '13px',
  fontWeight: 600,
  cursor: 'pointer',
  appearance: 'none',
  WebkitAppearance: 'none',
  MozAppearance: 'none',
  backgroundImage: SELECT_CHEVRON,
  backgroundRepeat: 'no-repeat',
  backgroundPosition: 'right 8px center',
};

/** Full-width sidebar selects (gradient + chevron) — 紧凑字号留给 Channel / Category 等筛选项 */
const selectSidebarStyle = {
  width: '100%',
  padding: '3px 22px 3px 6px',
  borderRadius: '5px',
  border: '1px solid rgba(148, 163, 184, 0.35)',
  backgroundColor: '#1e293b',
  backgroundImage: `linear-gradient(180deg, #3d4f63 0%, #1e293b 100%), ${SELECT_CHEVRON}`,
  backgroundRepeat: 'no-repeat, no-repeat',
  backgroundSize: '100% 100%, 9px 9px',
  backgroundPosition: '0 0, right 5px center',
  color: '#f8fafc',
  fontSize: '9px',
  fontWeight: 500,
  cursor: 'pointer',
  appearance: 'none',
  WebkitAppearance: 'none',
  MozAppearance: 'none',
  boxShadow: '0 1px 3px rgba(0,0,0,0.35)',
  outline: 'none',
  boxSizing: 'border-box',
};

/** 日历视图内联筛选（浅色底，与侧栏同一套状态） */
const CAL_SELECT_CHEVRON = `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E")`;
const selectCalendarToolbarStyle = {
  minWidth: '110px',
  maxWidth: '100%',
  padding: '6px 28px 6px 10px',
  borderRadius: '6px',
  border: '1px solid #cbd5e1',
  background: '#fff',
  color: '#1e293b',
  fontSize: '12px',
  fontWeight: 600,
  cursor: 'pointer',
  appearance: 'none',
  WebkitAppearance: 'none',
  MozAppearance: 'none',
  backgroundImage: CAL_SELECT_CHEVRON,
  backgroundRepeat: 'no-repeat',
  backgroundPosition: 'right 8px center',
  outline: 'none',
  boxSizing: 'border-box',
};

/** 颜色桶：1 蓝 → 2 灰 → 3 琥珀 → 4 青绿 → 5 紫；第 6 次及以后均为第 6 桶（深红，便于识别「超频」拜访） */
const VISIT_BUCKET_PALETTE = {
  1: {
    chipBg: '#2563eb',
    chipFg: '#eff6ff',
    sidebarBg: 'rgba(37, 99, 235, 0.38)',
    sidebarBorder: '#60a5fa',
    badgeBg: '#1d4ed8',
  },
  2: {
    chipBg: '#64748b',
    chipFg: '#f8fafc',
    sidebarBg: 'rgba(71, 85, 105, 0.55)',
    sidebarBorder: '#94a3b8',
    badgeBg: '#475569',
  },
  3: {
    chipBg: '#d97706',
    chipFg: '#fffbeb',
    sidebarBg: 'rgba(180, 83, 9, 0.35)',
    sidebarBorder: '#fb923c',
    badgeBg: '#b45309',
  },
  4: {
    chipBg: '#0d9488',
    chipFg: '#ecfdf5',
    sidebarBg: 'rgba(13, 148, 136, 0.32)',
    sidebarBorder: '#2dd4bf',
    badgeBg: '#0f766e',
  },
  5: {
    chipBg: '#7c3aed',
    chipFg: '#faf5ff',
    sidebarBg: 'rgba(109, 40, 217, 0.38)',
    sidebarBorder: '#a78bfa',
    badgeBg: '#6d28d9',
  },
  6: {
    chipBg: '#be123c',
    chipFg: '#fff1f2',
    sidebarBg: 'rgba(190, 18, 60, 0.36)',
    sidebarBorder: '#fb7185',
    badgeBg: '#9f1239',
  },
};
const UNSCHED_SIDEBAR = { sidebarBg: '#334155', sidebarBorder: '#3b82f6', badgeBg: null };

const CAL_CHIP_FONT_MIN = 5;
const CAL_CHIP_FONT_MAX = 12;

const CAL_CHIP_BADGE_W = 20;
const CAL_CHIP_BTN_SIZE = 22;

/** Estimate chip layout before row measure runs (avoids flash on first paint). */
function estimateRowChipMetrics(maxTaskCount) {
  if (maxTaskCount <= 0) {
    return { fontPx: CAL_CHIP_FONT_MAX, gap: 4, padV: 4, numPx: 10, badgeW: CAL_CHIP_BADGE_W, btnSize: CAL_CHIP_BTN_SIZE, slotH: 26 };
  }
  if (maxTaskCount <= 3) {
    return { fontPx: 11, gap: 4, padV: 4, numPx: 10, badgeW: CAL_CHIP_BADGE_W, btnSize: CAL_CHIP_BTN_SIZE, slotH: 26 };
  }
  if (maxTaskCount <= 6) {
    return { fontPx: 9, gap: 3, padV: 3, numPx: 9, badgeW: CAL_CHIP_BADGE_W, btnSize: CAL_CHIP_BTN_SIZE, slotH: 22 };
  }
  if (maxTaskCount <= 10) {
    return { fontPx: 8, gap: 2, padV: 2, numPx: 7, badgeW: CAL_CHIP_BADGE_W, btnSize: CAL_CHIP_BTN_SIZE, slotH: 18 };
  }
  return { fontPx: 7, gap: 2, padV: 2, numPx: 6, badgeW: CAL_CHIP_BADGE_W, btnSize: CAL_CHIP_BTN_SIZE, slotH: 14 };
}

/** 整周行统一 slot 高度 + 左右规格；各格条目横向贴齐、纵向按行对齐 */
function computeRowChipMetrics(maxTaskCount, availHeight, textWidth, tasks, measureEl) {
  if (maxTaskCount <= 0) return estimateRowChipMetrics(0);

  const n = maxTaskCount;
  const gap = n > 12 ? 2 : n > 8 ? 3 : n > 4 ? 4 : 4;
  const availForSlots = Math.max(20, availHeight - gap * Math.max(0, n - 1));
  const slotH = Math.max(14, Math.min(28, Math.round((availForSlots / n) * 10) / 10));

  let fontPx = Math.min(CAL_CHIP_FONT_MAX, Math.max(CAL_CHIP_FONT_MIN, slotH * 0.42));
  const padV = slotH >= 24 ? 4 : slotH >= 18 ? 3 : 2;

  const w = Math.max(24, textWidth);
  measureEl.style.width = `${w}px`;
  for (const t of tasks) {
    let size = fontPx;
    measureEl.textContent = t.name ?? '';
    measureEl.style.fontSize = `${size}px`;
    while (size > CAL_CHIP_FONT_MIN && measureEl.scrollWidth > w + 0.5) {
      size -= 0.5;
      measureEl.style.fontSize = `${size}px`;
    }
    fontPx = Math.min(fontPx, size);
  }

  fontPx = Math.max(CAL_CHIP_FONT_MIN, Math.round(fontPx * 2) / 2);
  const numPx = Math.max(CAL_CHIP_FONT_MIN, Math.min(10, fontPx - 0.5));

  return {
    fontPx,
    gap,
    padV,
    numPx,
    badgeW: CAL_CHIP_BADGE_W,
    btnSize: CAL_CHIP_BTN_SIZE,
    slotH,
  };
}

const PRINT_SCALE_SELECTORS = '.summary-print-wrap, .summary-print-page, .cal-print-sheet';

function printPageCellMetrics(orientation = 'portrait') {
  const weekCount = Math.max(1, document.querySelectorAll('.cal-week-block').length);
  const pageMarginMm = 10; // @page margin 5mm per side
  const headerMm = 12; // month title + weekday header row
  const pageW = orientation === 'portrait' ? 210 : 297;
  const pageH = orientation === 'portrait' ? 297 : 210;
  const usableW = pageW - pageMarginMm;
  const usableH = pageH - pageMarginMm - headerMm;
  return {
    weekCount,
    orientation,
    cellH: ((usableH * 96) / 25.4) / weekCount,
    cellW: ((usableW * 96) / 25.4) / 7,
  };
}

function resetCalendarPrintFonts() {
  document.querySelectorAll('.cal-day-tasks').forEach((el) => {
    el.style.removeProperty('--cal-print-font-size');
    el.style.removeProperty('--cal-print-num-font-size');
    el.style.removeProperty('--cal-print-line-height');
    el.style.removeProperty('--cal-print-slot-h');
    el.style.removeProperty('--cal-print-gap');
    el.style.removeProperty('gap');
    el.style.removeProperty('height');
    el.style.removeProperty('flex');
  });
  document.querySelectorAll('.cal-day-task-chip [data-cal-chip-text], .cal-day-task-chip span[aria-hidden]').forEach((el) => {
    if (el instanceof HTMLElement) {
      el.style.removeProperty('font-size');
      el.style.removeProperty('line-height');
    }
  });
  document.querySelectorAll('.cal-day-task-chip span[aria-hidden]').forEach((el) => {
    if (el instanceof HTMLElement) {
      el.style.removeProperty('width');
      el.style.removeProperty('min-width');
      el.style.removeProperty('height');
    }
  });
  // Restore the on-screen single-line ellipsis look that print-fit overrode
  document.querySelectorAll('.cal-day-task-chip [data-cal-chip-text]').forEach((el) => {
    if (el instanceof HTMLElement) {
      el.style.removeProperty('overflow-wrap');
      el.style.setProperty('white-space', 'nowrap');
      el.style.setProperty('text-overflow', 'ellipsis');
    }
  });
  document.querySelectorAll('.cal-day-task-chip').forEach((el) => {
    if (el instanceof HTMLElement) {
      el.style.removeProperty('min-height');
      el.style.removeProperty('max-height');
      el.style.removeProperty('height');
      el.style.removeProperty('flex');
      el.style.removeProperty('padding');
      el.style.removeProperty('gap');
      el.style.removeProperty('align-items');
    }
  });
}

function resetPrintScale() {
  document.body.classList.remove('print-fit-one-page', 'print-calendar-view', 'print-summary-view');
  document.querySelectorAll(PRINT_SCALE_SELECTORS).forEach((el) => {
    el.style.transform = '';
    el.style.zoom = '';
    el.style.width = '';
    el.style.height = '';
    el.style.marginBottom = '';
  });
  resetCalendarPrintFonts();
}

/** Print auto-fit sizes are computed from the real A4 paper size (not the on-screen window). */
const CAL_PRINT_FONT_MAX = 12; // ≈9pt on paper
const CAL_PRINT_FONT_MIN = 4;
const CAL_PRINT_LINE_H = 1.06;

const PRINT_FIT_VERSION = 'print-fit v7 · uniform-global · true-text-width';

/* Print chip geometry: the on-screen chip keeps inline width:20px on the
 * round badge and gap:4px, which the print CSS never overrode — so on paper
 * the name column was ~15% narrower than the fit math assumed and long names
 * wrapped one line more than measured, clipping dense days. In print we force
 * these compact values AND use the same numbers in the width math. */
const CAL_PRINT_BADGE_W = 8;
const CAL_PRINT_CHIP_GAP = 1.5;

/**
 * Measure total height of all names in one cell at a candidate font size.
 * Names may wrap to AT MOST 2 lines; returns Infinity when any name needs more.
 */
function measureCellTextHeight(measureEl, texts, w, fontPx) {
  measureEl.style.width = `${Math.max(12, w)}px`;
  measureEl.style.fontSize = `${fontPx}px`;
  const maxNameH = 2 * fontPx * CAL_PRINT_LINE_H + 2;
  let total = 0;
  for (const text of texts) {
    measureEl.textContent = text;
    const h = measureEl.scrollHeight;
    if (h > maxNameH) return Infinity;
    total += h;
  }
  return total;
}

/**
 * Per-day cell auto-fit from A4 paper metrics (window size irrelevant).
 * ONE font size per cell (tidy), as large as fits: every name ≤2 wrapped
 * lines and the stack of natural-height rows fits the cell height.
 * Rows keep natural height (no stretching), leftover space becomes even gaps.
 */
function fitCalendarPrintCellFonts(orientation = 'portrait') {
  const measureEl = document.createElement('span');
  measureEl.style.cssText = [
    'position:absolute',
    'left:-99999px',
    'top:0',
    'visibility:hidden',
    'display:block',
    'white-space:normal',
    'overflow-wrap:break-word',
    `line-height:${CAL_PRINT_LINE_H}`,
    'font-family:system-ui,-apple-system,"Segoe UI",sans-serif',
    'font-weight:700',
  ].join(';');
  document.body.appendChild(measureEl);

  const cellStats = [];

  try {
    const paper = printPageCellMetrics(orientation);
    const labelW = 11.4; // 3mm printed date column
    const cellChromeW = 6; // 0.4mm cell padding ×2 + 0.2mm gap + 1px borders ×2
    const safetyW = 3; // absorbs badge right-padding + subpixel wrap rounding
    const textW = Math.max(
      14,
      paper.cellW - labelW - cellChromeW - CAL_PRINT_BADGE_W - CAL_PRINT_CHIP_GAP - safetyW,
    );
    const availH = Math.max(12, paper.cellH - 5);

    // Pass 1: fitted font size per cell (largest where every name ≤2 lines
    // and the row stack fits the cell height).
    const fittedCells = [];
    document.querySelectorAll('.cal-day-cell').forEach((cell) => {
      const tasksCol = cell.querySelector('.cal-day-tasks');
      if (!tasksCol) return;
      const chips = [...tasksCol.querySelectorAll('.cal-day-task-chip')];
      const n = chips.length;
      if (n === 0) return;

      const texts = chips.map(
        (chip) => chip.querySelector('[data-cal-chip-text]')?.textContent ?? '',
      );
      const minGap = n > 14 ? 0.4 : n > 8 ? 0.6 : 1;

      let lo = CAL_PRINT_FONT_MIN;
      let hi = CAL_PRINT_FONT_MAX;
      const fitsAt = (f) =>
        measureCellTextHeight(measureEl, texts, textW, f) + minGap * (n - 1) <= availH;
      let fontPx;
      if (fitsAt(hi)) {
        fontPx = hi;
      } else if (!fitsAt(lo)) {
        fontPx = lo;
      } else {
        for (let i = 0; i < 9; i++) {
          const mid = (lo + hi) / 2;
          if (fitsAt(mid)) lo = mid;
          else hi = mid;
        }
        fontPx = Math.max(CAL_PRINT_FONT_MIN, Math.floor(lo * 4) / 4);
      }

      fittedCells.push({ cell, tasksCol, chips, texts, n, minGap, fontPx });
    });

    // One font size for the whole sheet: the densest cell's size, so sparse
    // days never print bigger than busy days.
    const globalFontPx = fittedCells.length
      ? Math.max(CAL_PRINT_FONT_MIN, Math.min(...fittedCells.map((c) => c.fontPx)))
      : CAL_PRINT_FONT_MIN;

    // Pass 2: apply the uniform font; per-cell leftover height still becomes
    // even gaps (capped so sparse cells stay top-aligned instead of scattering
    // 1–2 names oddly).
    fittedCells.forEach(({ cell, tasksCol, chips, texts, n, minGap }) => {
      const fontPx = globalFontPx;
      const textH = Math.min(availH, measureCellTextHeight(measureEl, texts, textW, fontPx));
      const leftover = Math.max(0, availH - textH);
      const gapPx = Math.min(
        n > 1 ? leftover / (n - 1) : 0,
        Math.max(minGap, fontPx * 0.9),
      );

      // Capped so 1–2 digits fit inside the compact print badge
      const numPx = Math.max(3, Math.min(fontPx - 0.5, 6.5));

      tasksCol.setAttribute('data-task-count', String(n));
      tasksCol.style.setProperty('height', '100%', 'important');
      tasksCol.style.setProperty('flex', '1 1 auto', 'important');
      tasksCol.style.setProperty('gap', `${Math.max(minGap, gapPx)}px`, 'important');
      tasksCol.style.setProperty('--cal-print-font-size', `${fontPx}px`);
      tasksCol.style.setProperty('--cal-print-num-font-size', `${numPx}px`);
      tasksCol.style.setProperty('--cal-print-line-height', String(CAL_PRINT_LINE_H));

      chips.forEach((chip) => {
        if (!(chip instanceof HTMLElement)) return;
        chip.style.setProperty('flex', '0 0 auto', 'important');
        chip.style.setProperty('min-height', '0', 'important');
        chip.style.setProperty('height', 'auto', 'important');
        chip.style.setProperty('max-height', 'none', 'important');
        chip.style.setProperty('padding', '0', 'important');
        chip.style.setProperty('align-items', 'flex-start', 'important');
        chip.style.setProperty('gap', `${CAL_PRINT_CHIP_GAP}px`, 'important');
        const textEl = chip.querySelector('[data-cal-chip-text]');
        const numEl = chip.querySelector('span[aria-hidden]');
        if (textEl instanceof HTMLElement) {
          textEl.style.setProperty('font-size', `${fontPx}px`, 'important');
          textEl.style.setProperty('line-height', String(CAL_PRINT_LINE_H), 'important');
          textEl.style.setProperty('white-space', 'normal', 'important');
          textEl.style.setProperty('overflow-wrap', 'break-word', 'important');
          textEl.style.setProperty('text-overflow', 'clip', 'important');
        }
        if (numEl instanceof HTMLElement) {
          numEl.style.setProperty('font-size', `${numPx}px`, 'important');
          numEl.style.setProperty('line-height', String(CAL_PRINT_LINE_H), 'important');
          numEl.style.setProperty('width', `${CAL_PRINT_BADGE_W}px`, 'important');
          numEl.style.setProperty('min-width', `${CAL_PRINT_BADGE_W}px`, 'important');
          numEl.style.setProperty('height', 'auto', 'important');
        }
      });

      const dayNum = cell.querySelector('b')?.textContent ?? '?';
      cellStats.push(`day ${dayNum}: n=${n} font=${fontPx}px gap=${gapPx.toFixed(1)}`);
    });

    console.log(
      `[${PRINT_FIT_VERSION}] weeks=${paper.weekCount} cell=${paper.cellW.toFixed(0)}×${paper.cellH.toFixed(0)}px textW=${textW.toFixed(0)} globalFont=${globalFontPx}px\n` +
        cellStats.join('\n'),
    );
  } finally {
    measureEl.remove();
  }
}

function fitElementToOnePage(el, orientation = 'landscape', minScale = 0.35) {
  el.style.transform = '';
  el.style.zoom = '';
  el.style.width = '';
  el.style.height = '';
  el.style.marginBottom = '';
  const marginMm = 10;
  const pageW = orientation === 'portrait' ? 210 : 297;
  const pageH = orientation === 'portrait' ? 297 : 210;
  const maxW = ((pageW - marginMm) * 96) / 25.4;
  const maxH = ((pageH - marginMm) * 96) / 25.4;
  const h = el.scrollHeight;
  const w = el.scrollWidth;
  const scaleH = h > 0 ? maxH / h : 1;
  const scaleW = w > 0 ? maxW / w : 1;
  const scale = Math.min(1, Math.max(minScale, Math.min(scaleH, scaleW)));
  if (scale >= 0.999) return;
  el.style.zoom = String(scale);
  el.style.width = `${100 / scale}%`;
}

function fitElementToOnePortraitPage(el, minScale = 0.35) {
  fitElementToOnePage(el, 'portrait', minScale);
}

function refitPrintContent() {
  const summary = document.querySelector('.summary-print-wrap:not(.summary-print-empty)');
  if (summary) {
    // Each two-week block prints on its own A4 page; shrink any block that is
    // taller than one page so it never spills onto a third page.
    const pages = summary.querySelectorAll('.summary-print-page');
    if (pages.length > 0) {
      pages.forEach((pg) => fitElementToOnePortraitPage(pg));
    } else {
      fitElementToOnePortraitPage(summary);
    }
    return;
  }
  if (document.querySelector('.cal-print-sheet')) {
    fitCalendarPrintCellFonts('portrait');
  }
}

function preparePrintLayout() {
  resetPrintScale();
  document.body.classList.add('print-fit-one-page');

  const summary = document.querySelector('.summary-print-wrap:not(.summary-print-empty)');
  if (summary) {
    document.body.classList.add('print-summary-view');
    return;
  }

  if (document.querySelector('.cal-print-sheet')) {
    document.body.classList.add('print-calendar-view');
  }
}

function runPrintWithLayout(onAfterPrint) {
  preparePrintLayout();
  window.requestAnimationFrame(() => {
    refitPrintContent();
    window.requestAnimationFrame(() => {
      window.print();
      // window.print() blocks while the dialog is open; always restore the
      // screen afterwards even if the browser never fires `afterprint`.
      window.setTimeout(() => {
        resetPrintScale();
        if (typeof onAfterPrint === 'function') onAfterPrint();
      }, 300);
    });
  });
}

/** 有效拜访序号，至少为 1 */
function normalizeVisitRound(r) {
  const n = Math.floor(Number(r));
  return Number.isFinite(n) && n >= 1 ? n : 1;
}

/** ISO yyyy-mm-dd → 可读标题（用于双击日历弹层） */
function formatIsoDateHeading(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso).trim());
  if (!m) return String(iso);
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  if (!y || !mo || !d) return String(iso);
  try {
    return new Date(y, mo - 1, d).toLocaleDateString(undefined, {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
  } catch {
    return String(iso);
  }
}

/**
 * ISO yyyy-mm-dd interpreted in the **local** calendar (avoid `new Date('yyyy-mm-dd')`, which is UTC midnight
 * and can show the wrong weekday west of Greenwich).
 */
function localWeekdayIndexFromIso(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso).trim());
  if (!m) return NaN;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  return new Date(y, mo - 1, d).getDay();
}

/** Sun … Sat abbreviated label for yyyy-mm-dd in local TZ */
function weekdayAbbrevFromIso(dateStr) {
  const wd = localWeekdayIndexFromIso(dateStr);
  if (!Number.isFinite(wd) || wd < 0 || wd > 6) return '—';
  return WEEKDAY_ABBREV_SUN0[wd];
}

/** ISO yyyy-mm-dd plus one calendar month in local timezone
 * (e.g. Jan 31 → Feb 28/29 per JS Date rollover rules).
 */
function addOneCalendarMonthToIsoDate(dateStr) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateStr).trim());
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  if (!Number.isFinite(y) || !Number.isFinite(mo) || !Number.isFinite(d)) return null;
  const dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
  dt.setMonth(dt.getMonth() + 1);
  const ny = dt.getFullYear();
  const nmo = String(dt.getMonth() + 1).padStart(2, '0');
  const nd = String(dt.getDate()).padStart(2, '0');
  return `${ny}-${nmo}-${nd}`;
}

function formatYyyyMmDdLocal(dt) {
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  const d = String(dt.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * Company policy: skip Sundays when placing copied visits → move target to Monday (+1 calendar day locally).
 * If that Monday falls outside the intended month (e.g. Jun 29 Sun→Jun 30 Mon OK; theoretical Jun 30 Sun→Jul 1), callers may still persist that date (see duplicate fn).
 */
function bumpSundayCopiedVisitToMondayYyyyMmDd(dateStr) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateStr).trim());
  if (!m) return dateStr;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return dateStr;
  if (dt.getDay() !== 0) return dateStr;
  dt.setDate(dt.getDate() + 1);
  return formatYyyyMmDdLocal(dt);
}

/** True if yyyy-mm-dd is a Sunday in local timezone. */
function isoYyyyMmDdIsSundayLocal(dateStr) {
  return localWeekdayIndexFromIso(dateStr) === 0;
}

/**
 * For one agent+year+month: every visit scheduled on Sunday moves to Monday (merge per day; skip duplicate acc).
 * Mondays may spill to the next calendar month (rare month-end Sunday).
 */
function moveSundayVisitsToMondayForAgentMonth(prevSchedules, agentKey, yearVal, month0) {
  if (!prevSchedules || typeof prevSchedules !== 'object') {
    return {
      next: prevSchedules,
      changed: false,
      sundayDatesTouched: 0,
      movedVisits: 0,
      skippedDup: 0,
    };
  }
  const byDate = prevSchedules[agentKey];
  if (!byDate || typeof byDate !== 'object') {
    return {
      next: prevSchedules,
      changed: false,
      sundayDatesTouched: 0,
      movedVisits: 0,
      skippedDup: 0,
    };
  }

  const prefix = `${yearVal}-${String(month0 + 1).padStart(2, '0')}-`;
  const sundayKeys = Object.keys(byDate).filter(
    (k) => typeof k === 'string' && k.startsWith(prefix) && isoYyyyMmDdIsSundayLocal(k),
  ).sort();

  if (sundayKeys.length === 0) {
    return {
      next: prevSchedules,
      changed: false,
      sundayDatesTouched: 0,
      movedVisits: 0,
      skippedDup: 0,
    };
  }

  const agentSched = {};
  for (const k of Object.keys(byDate)) agentSched[k] = [...byDate[k]];

  let movedVisits = 0;
  let skippedDup = 0;
  let sundayDatesTouched = 0;
  let mutated = false;

  for (const sunKey of sundayKeys) {
    const visits = agentSched[sunKey];
    if (!Array.isArray(visits) || visits.length === 0) {
      delete agentSched[sunKey];
      mutated = true;
      continue;
    }
    const monKey = bumpSundayCopiedVisitToMondayYyyyMmDd(sunKey);
    if (monKey === sunKey) continue;

    let monList = [...(agentSched[monKey] || [])];
    const seen = new Set(monList.map((t) => t.acc));

    for (const v of visits) {
      if (!v || v.acc == null || String(v.acc).length === 0) continue;
      const acc = String(v.acc);
      if (seen.has(acc)) {
        skippedDup += 1;
        continue;
      }
      monList.push({
        acc,
        name: String(v.name ?? ''),
        visitRound: normalizeVisitRound(v.visitRound ?? 1),
      });
      seen.add(acc);
      movedVisits += 1;
    }
    agentSched[monKey] = monList;
    delete agentSched[sunKey];
    sundayDatesTouched += 1;
    mutated = true;
  }

  if (!mutated) {
    return {
      next: prevSchedules,
      changed: false,
      sundayDatesTouched: 0,
      movedVisits: 0,
      skippedDup: 0,
    };
  }

  const nextSchedules = normalizePersistedSchedules({ ...prevSchedules, [agentKey]: agentSched });
  return {
    next: nextSchedules,
    changed: true,
    sundayDatesTouched,
    movedVisits,
    skippedDup,
  };
}

/** Copy one agent's visits from source month onto the following month (same day-of-month; bump past Sundays & yellow public holidays day-by-day; merges; skips dup acc per day). */
function duplicateAgentMonthVisitsOneMonthAhead(prevSchedules, agentKey, yearVal, month0) {
  if (!prevSchedules || typeof prevSchedules !== 'object') {
    return {
      next: prevSchedules,
      changed: false,
      added: 0,
      skippedDup: 0,
      sourceDatesShiftedFromBlockedDays: 0,
      daysSpilledPastTargetMonth: 0,
    };
  }
  const byDate = prevSchedules[agentKey];
  if (!byDate || typeof byDate !== 'object') {
    return {
      next: prevSchedules,
      changed: false,
      added: 0,
      skippedDup: 0,
      sourceDatesShiftedFromBlockedDays: 0,
      daysSpilledPastTargetMonth: 0,
    };
  }

  const srcIsoPad = String(month0 + 1).padStart(2, '0');
  const srcPrefix = `${yearVal}-${srcIsoPad}-`;
  const nextMonthIdx = month0 === 11 ? 0 : month0 + 1;
  const nextYearVal = month0 === 11 ? yearVal + 1 : yearVal;
  const expectedYm = `${nextYearVal}-${String(nextMonthIdx + 1).padStart(2, '0')}`;

  const agentSched = {};
  for (const k of Object.keys(byDate)) agentSched[k] = [...byDate[k]];
  let added = 0;
  let skippedDup = 0;
  let changed = false;
  let sourceDatesShiftedFromBlockedDays = 0;
  let daysSpilledPastTargetMonth = 0;

  for (const [dateStr, visits] of Object.entries(byDate)) {
    if (typeof dateStr !== 'string' || !dateStr.startsWith(srcPrefix)) continue;
    if (!Array.isArray(visits) || visits.length === 0) continue;
    const rawTarget = addOneCalendarMonthToIsoDate(dateStr);
    if (!rawTarget) continue;
    if (rawTarget.slice(0, 7) !== expectedYm) {
      continue;
    }
    const targetStr = advanceIsoPastSundaysAndPublicHolidays(rawTarget);
    if (targetStr !== rawTarget) sourceDatesShiftedFromBlockedDays += 1;
    if (targetStr.slice(0, 7) !== expectedYm) daysSpilledPastTargetMonth += 1;

    let tgtList = [...(agentSched[targetStr] || [])];
    const seen = new Set(tgtList.map((t) => t.acc));

    for (const v of visits) {
      if (!v || v.acc == null || String(v.acc).length === 0) continue;
      const acc = String(v.acc);
      if (seen.has(acc)) {
        skippedDup += 1;
        continue;
      }
      tgtList.push({
        acc,
        name: String(v.name ?? ''),
        visitRound: normalizeVisitRound(v.visitRound ?? 1),
      });
      seen.add(acc);
      added += 1;
      changed = true;
    }
    agentSched[targetStr] = tgtList;
  }

  if (!changed) {
    return {
      next: prevSchedules,
      changed: false,
      added: 0,
      skippedDup,
      sourceDatesShiftedFromBlockedDays: 0,
      daysSpilledPastTargetMonth: 0,
    };
  }
  const nextSchedules = normalizePersistedSchedules({ ...prevSchedules, [agentKey]: agentSched });
  return {
    next: nextSchedules,
    changed: true,
    added,
    skippedDup,
    sourceDatesShiftedFromBlockedDays,
    daysSpilledPastTargetMonth,
  };
}

/** 将序号映射到 1–6 配色桶（≥6 次均为桶 6） */
function colorBucketFromRound(round) {
  return Math.min(normalizeVisitRound(round), 6);
}

/** 单日格拜访条列表；整行 slot 等高、左右贴齐，顾客少时下方留白 */
function CalDayTasksColumn({ dateStr, tasks, removeScheduledVisit, metrics }) {
  const { fontPx, gap, padV, numPx, badgeW, btnSize, slotH } =
    metrics ?? estimateRowChipMetrics(tasks.length);

  return (
    <div
      className="cal-day-tasks"
      data-task-count={tasks.length || undefined}
      style={{
        marginTop: '6px',
        display: 'flex',
        flexDirection: 'column',
        gap: `${gap}px`,
        flex: '1 1 auto',
        minHeight: 0,
        minWidth: 0,
        width: '100%',
        alignSelf: 'stretch',
        overflowX: 'hidden',
        overflowY: tasks.length > 0 ? 'auto' : 'visible',
        WebkitOverflowScrolling: 'touch',
        overscrollBehavior: 'contain',
        position: 'relative',
      }}
    >
      {tasks.map((t) => {
        const r = normalizeVisitRound(t.visitRound ?? 1);
        const bucket = colorBucketFromRound(r);
        const tp = VISIT_BUCKET_PALETTE[bucket];
        const chipBg = tp.chipBg;
        const chipFg = tp.chipFg;
        const numLabel = String(r);
        return (
          <div
            key={`${dateStr}-${t.acc}-${r}`}
            className="cal-day-task-chip"
            draggable
            title={`${t.name} · ACC ${t.acc}`}
            onDragStart={(e) =>
              e.dataTransfer.setData('text', JSON.stringify({ ...t, oldDate: dateStr }))
            }
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '4px',
              background: chipBg,
              color: chipFg,
              padding: `${padV}px 4px`,
              borderRadius: '4px',
              cursor: 'move',
              overflow: 'hidden',
              flex: `0 0 ${slotH}px`,
              height: `${slotH}px`,
              minHeight: `${slotH}px`,
              maxHeight: `${slotH}px`,
              width: '100%',
              alignSelf: 'stretch',
              boxSizing: 'border-box',
              border: `1px solid rgba(255,255,255,0.22)`,
              ['--visit-round-color']: chipBg,
            }}
          >
            <span
              style={{
                flexShrink: 0,
                width: `${badgeW}px`,
                minWidth: `${badgeW}px`,
                height: `${badgeW}px`,
                padding: 0,
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                borderRadius: '3px',
                fontSize: `${numPx}px`,
                fontWeight: 800,
                background: 'rgba(0,0,0,0.2)',
                color: chipFg,
              }}
              aria-hidden
            >
              {numLabel}
            </span>
            <span
              data-cal-chip-text
              style={{
                flex: 1,
                minWidth: 0,
                color: chipFg,
                fontSize: `${fontPx}px`,
                lineHeight: 1.1,
                whiteSpace: 'nowrap',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
                fontFamily: 'system-ui, -apple-system, "Segoe UI", sans-serif',
              }}
            >
              {t.name}
            </span>
            <button
              type="button"
              className="no-print"
              draggable={false}
              onMouseDown={(e) => e.stopPropagation()}
              onClick={(e) => {
                e.stopPropagation();
                e.preventDefault();
                removeScheduledVisit(dateStr, t.acc);
              }}
              aria-label={`Remove visit ${t.acc} (${t.name})`}
              title="Remove from this day"
              style={{
                flexShrink: 0,
                width: `${btnSize}px`,
                minWidth: `${btnSize}px`,
                height: `${btnSize}px`,
                padding: 0,
                margin: 0,
                border: 'none',
                borderRadius: '3px',
                cursor: 'pointer',
                fontSize: `${Math.max(10, fontPx + 1)}px`,
                lineHeight: 1,
                fontWeight: 700,
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: chipFg,
                background: 'rgba(0,0,0,0.25)',
              }}
            >
              ×
            </button>
          </div>
        );
      })}
    </div>
  );
}

/** 将月历铺平为周行，每行 7 格（首末补空占位） */
function monthWeekSlots(firstDay, totalDays) {
  const slots = [];
  for (let i = 0; i < firstDay; i++) slots.push({ kind: 'padding', key: `lead-${i}` });
  for (let d = 1; d <= totalDays; d++) slots.push({ kind: 'day', key: `day-${d}`, dayNum: d });
  const trailing = (7 - (slots.length % 7)) % 7;
  for (let t = 0; t < trailing; t++) slots.push({ kind: 'padding', key: `trail-${t}` });
  const weeks = [];
  for (let i = 0; i < slots.length; i += 7) weeks.push(slots.slice(i, i + 7));
  return weeks;
}

/** 整周行统一左右规格；格高由 grid stretch，条在格内纵向拉伸 */
function CalWeekRow({
  weekCells,
  monthIsoPad,
  year,
  agent,
  schedules,
  highlightedDayNum,
  dropOnDate,
  removeScheduledVisit,
  calendarFilterAcc,
  calendarChipRoundFilter,
  onDayDoubleClick,
}) {
  const rowHostRef = useRef(null);
  const [rowMetrics, setRowMetrics] = useState(null);

  const enriched = useMemo(
    () =>
      weekCells.map((slot) => {
        if (slot.kind !== 'day') return slot;
        const dateStr = `${year}-${monthIsoPad}-${String(slot.dayNum).padStart(2, '0')}`;
        const rawTasks = (agent && schedules[agent] && schedules[agent][dateStr]) || [];
        const filteredByCust =
          typeof calendarFilterAcc === 'function'
            ? rawTasks.filter((t) => calendarFilterAcc(t.acc))
            : rawTasks;
        const tasks = filterCalendarTasksByChipRound(filteredByCust, calendarChipRoundFilter);
        return {
          ...slot,
          dateStr,
          isSun: new Date(year, Number(monthIsoPad) - 1, slot.dayNum).getDay() === 0,
          tasks,
        };
      }),
    [weekCells, monthIsoPad, year, agent, schedules, calendarFilterAcc, calendarChipRoundFilter],
  );

  const rowMaxTasks = useMemo(
    () =>
      enriched.reduce(
        (max, c) => (c.kind === 'day' ? Math.max(max, c.tasks.length) : max),
        0,
      ),
    [enriched],
  );

  const rowTaskKey = useMemo(
    () =>
      `${agent}\u0003${rowMaxTasks}\u0003` +
      enriched
        .filter((c) => c.kind === 'day')
        .flatMap((c) => c.tasks)
        .map((t) => `${t.acc}\u0001${t.name}\u0001${normalizeVisitRound(t.visitRound ?? 1)}`)
        .join('\u0002'),
    [enriched, agent, rowMaxTasks],
  );

  const metricsForRow = rowMetrics ?? estimateRowChipMetrics(rowMaxTasks);

  useLayoutEffect(() => {
    const host = rowHostRef.current;
    if (!host || rowMaxTasks === 0) {
      setRowMetrics(null);
      return;
    }

    // While print layout classes are on <body>, cell sizes reflect the A4 print
    // sheet, not the screen. Measuring then would poison rowMetrics with huge
    // values (Chrome fires resize events when the print preview opens/closes).
    const isPrintLayoutActive = () =>
      document.body.classList.contains('print-fit-one-page');

    const measureRow = () => {
      if (isPrintLayoutActive()) return;
      const sampleCell = host.querySelector('.cal-day-cell[data-cal-date]');
      if (!(sampleCell instanceof HTMLElement)) return;

      const dayLabel = sampleCell.querySelector('b');
      const labelH = dayLabel instanceof HTMLElement ? dayLabel.offsetHeight : 22;
      const availH = Math.max(20, sampleCell.clientHeight - labelH - 6 - 4);

      const textW = Math.max(
        40,
        sampleCell.clientWidth - 8 - CAL_CHIP_BADGE_W - CAL_CHIP_BTN_SIZE - 12,
      );

      const allTasks = enriched.filter((c) => c.kind === 'day').flatMap((c) => c.tasks);

      const measureEl = document.createElement('span');
      measureEl.style.cssText = [
        'position:absolute',
        'left:-99999px',
        'top:0',
        'visibility:hidden',
        'pointer-events:none',
        'white-space:nowrap',
        'overflow:hidden',
        'display:block',
        'line-height:1.1',
        'margin:0',
        'padding:0',
        'box-sizing:border-box',
        'font-family:system-ui,-apple-system,"Segoe UI",sans-serif',
      ].join(';');
      document.body.appendChild(measureEl);

      let next;
      try {
        next = computeRowChipMetrics(rowMaxTasks, availH, textW, allTasks, measureEl);
      } finally {
        try {
          measureEl.remove();
        } catch {
          /* ignore detach errors */
        }
      }

      setRowMetrics((prev) => {
        if (
          prev &&
          prev.fontPx === next.fontPx &&
          prev.gap === next.gap &&
          prev.slotH === next.slotH &&
          prev.badgeW === next.badgeW
        ) {
          return prev;
        }
        return next;
      });
    };

    if (typeof document === 'undefined') return;

    let retryTid = 0;
    const onResize = () => {
      if (isPrintLayoutActive()) {
        // A resize fired while the print sheet is up; re-measure once the
        // screen layout is back instead of recording print-sheet sizes.
        window.clearTimeout(retryTid);
        retryTid = window.setTimeout(onResize, 300);
        return;
      }
      try {
        measureRow();
      } catch (err) {
        console.warn('route-planner: row chip measure on resize failed', err);
      }
    };
    onResize();
    window.addEventListener('resize', onResize);
    return () => {
      window.removeEventListener('resize', onResize);
      window.clearTimeout(retryTid);
    };
  }, [rowMaxTasks, rowTaskKey, enriched]);

  return (
    <div className="cal-week-block" style={{ width: '100%', boxSizing: 'border-box' }}>
      <div
        ref={rowHostRef}
        className="cal-week-row"
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(7, minmax(0, 1fr))',
          alignItems: 'stretch',
          width: '100%',
          boxSizing: 'border-box',
        }}
      >
        {enriched.map((cell) => {
          if (cell.kind === 'padding') {
            return (
              <div
                key={cell.key}
                className="cal-pad-cell"
                style={{
                  background: '#f8fafc',
                  border: '1px solid #e2e8f0',
                  minHeight: '144px',
                  height: '100%',
                  boxSizing: 'border-box',
                }}
              />
            );
          }
          const { dayNum, dateStr, isSun, tasks } = cell;
          const isFocused = highlightedDayNum != null && highlightedDayNum === dayNum;
          const holidayName = sabahHolidayNameForCalendarDate(dateStr);
          const cellTitle =
            holidayName != null
              ? `${holidayName} (public holiday) · Double-click this day for the visit list`
              : isSun
                ? 'Sunday · often a weekly rest day or non-working day · Double-click for the visit list'
                : 'Double-click this day for full customer names and accounts';
          return (
            <div
              key={cell.key}
              className="cal-day-cell"
              data-cal-date={dateStr}
              data-cal-focus-day={dayNum}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const raw = e.dataTransfer.getData('text') || e.dataTransfer.getData('text/plain');
                dropOnDate(dateStr, raw);
              }}
              onDoubleClick={() => onDayDoubleClick?.(dateStr)}
              title={cellTitle}
              style={{
                minHeight: '144px',
                height: '100%',
                minWidth: 0,
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'stretch',
                border: '1px solid #cbd5e1',
                padding: '6px 4px',
                background: holidayName ? '#fef08a' : (isSun ? '#ffe4e6' : 'white'),
                position: 'relative',
                boxSizing: 'border-box',
                boxShadow: isFocused ? 'inset 0 0 0 2px #2563eb' : undefined,
                zIndex: isFocused ? 1 : 0,
              }}
            >
              <b style={{ flexShrink: 0, color: isSun ? '#e11d48' : '#1e293b', fontSize: '15px', fontWeight: 800 }}>{dayNum}</b>
              <CalDayTasksColumn
                dateStr={dateStr}
                tasks={tasks}
                removeScheduledVisit={removeScheduledVisit}
                metrics={metricsForRow}
              />
            </div>
          );
        })}
      </div>
    </div>
  );
}

function paletteForColorBucket(bucket) {
  const b = Math.min(Math.max(Math.floor(bucket), 1), 6);
  return { ...UNSCHED_SIDEBAR, ...VISIT_BUCKET_PALETTE[b] };
}

/** Sidebar visit badge color: bucket from a non-negative integer (typically month visit count). */
function paletteForVisitCountDisplay(n) {
  if (n <= 0) return UNSCHED_SIDEBAR;
  return paletteForColorBucket(Math.min(n, 6));
}

/** 侧栏「已排期次数」筛选 */
function matchesVisitScheduleFilter(scheduledTotal, filter) {
  const n = scheduledTotal ?? 0;
  if (filter === 'all') return true;
  if (filter === '0') return n === 0;
  if (filter === '6+') return n >= 6;
  const k = Number(filter);
  if (Number.isFinite(k)) return n === k;
  return true;
}

/** 日历条形角标 visitRound — 筛选「只看第几次拜访」，不是客户当月总条数 */
function filterCalendarTasksByChipRound(tasks, filter) {
  if (filter == null || filter === 'all') return tasks;
  if (filter === '0') return [];
  if (filter === '6+') {
    return tasks.filter((t) => normalizeVisitRound(t.visitRound ?? 1) >= 6);
  }
  const k = Number(filter);
  if (Number.isFinite(k)) {
    return tasks.filter((t) => normalizeVisitRound(t.visitRound ?? 1) === k);
  }
  return tasks;
}

/** Sync：侧栏 Visits（当月每户拜访条数）与日历 Visits（条形 #）共用 all·1–5·6+；侧栏「0× 尚未排期」在日历无同义项 → 设为 all（仍显示条码） */
function mapSidebarVisitsFilterToCalendarChipFilter(sidebarFilter) {
  if (sidebarFilter === '0') return 'all';
  return sidebarFilter;
}

/** 侧栏 Area / Channel / Category：空选项表示不限制 */
function matchesSidebarListFilters(c, area, channelFilter, categoryFilter) {
  return (
    (!area || c.area === area) &&
    (!channelFilter || String(c.channel ?? '—') === channelFilter) &&
    (!categoryFilter || String(c.category ?? '—') === categoryFilter)
  );
}

/** 导入校验：非空 ACC 在表内不得重复（空格子不参与，避免因多行空 ACC 整块导入被拒） */
function findDuplicateImportAccs(customerRows) {
  const tally = Object.create(null);
  for (const r of customerRows) {
    const raw = String(r.acc ?? '').trim();
    if (raw === '') continue;
    tally[raw] = (tally[raw] || 0) + 1;
  }
  return Object.keys(tally)
    .filter((k) => tally[k] > 1)
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
}

/** 表格导入：空白 → 显示为 —（与下拉筛选用同一套值） */
function normalizeImportField(raw) {
  if (raw === undefined || raw === null || String(raw).trim() === '') return '—';
  return String(raw).trim();
}

/** 欄標題比對用（中英混排、空格差異） */
function headerKeyNormalized(k) {
  return String(k ?? '')
    .normalize('NFKC')
    .replace(/\u00a0|\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

/**
 * sheet_to_json 欄キー入面認「店名 / 客户名」，唔好用單一 substring 「NAME」，否則會先撞上 SALES AGENT NAME 等標題，
 * kkk、xxx 之流會存入錯欄位而搜唔到。
 */
function pickDebtorNameColumnKey(keys) {
  if (!keys?.length) return null;
  const list = keys.filter((k) => k != null && String(k).trim() !== '');
  const nk = headerKeyNormalized;
  /** 常見全称 */
  for (const needle of ['DEBTOR NAME', 'CUSTOMER NAME', 'COMPANY NAME', 'SHIP TO NAME', 'BILL TO NAME', 'CUS NAME']) {
    const hit = list.find((k) => nk(k) === nk(needle));
    if (hit) return hit;
  }
  for (const needle of ['DEBTOR NAME', 'CUSTOMER NAME', 'COMPANY NAME', 'SHIP TO NAME', 'CUS NAME']) {
    const hit = list.find((k) => nk(k).includes(needle.replace(/\s+/g, ' ')));
    if (hit) return hit;
  }
  const exactNameOnly = list.find((k) => nk(k) === 'NAME');
  if (exactNameOnly) return exactNameOnly;
  return list.find((k) => {
    const u = nk(k);
    if (!/\bNAME\b/.test(u)) return false;
    if (/SALES\s*AGENT|AGENT\s*NAME|NAME\s*\(\s*sales|NAME\s*\(\s*s\.?\s*a|NAME\s*OF\s*AGENTS?|SALES\s+(MAN|REP|NAME)|SALESMAN/i.test(u)) return false;
    return true;
  });
}

/** 搜尋欄匹配：統一標點空格、唔吃 Excel 嘅零寬字 */
function normalizeSidebarSearchPiece(s) {
  return String(s ?? '')
    .normalize('NFKC')
    .replace(/\u00a0|\u3000/g, ' ')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * 重新导入客户表后合并旧排期：账号仍在新表中则保留拜访（日期、# 不变），姓名随新表更新；
 * 业务员以新表为准（账号若改派，拜访会跟到新业务员下）。新表中不存在的账号之拜访会丢弃。
 */
function mergeSchedulesAfterCustomerReimport(prevSchedules, newCustomers) {
  const byAcc = Object.fromEntries(
    newCustomers.map((c) => [String(c.acc ?? ''), c]),
  );
  let droppedVisits = 0;
  const prev =
    prevSchedules && typeof prevSchedules === 'object' && !Array.isArray(prevSchedules)
      ? prevSchedules
      : {};

  const next = {};

  for (const byDate of Object.values(prev)) {
    if (!byDate || typeof byDate !== 'object') continue;
    for (const [dateStr, visits] of Object.entries(byDate)) {
      if (typeof dateStr !== 'string' || !Array.isArray(visits)) continue;
      for (const v of visits) {
        if (!v || v.acc == null || String(v.acc).length === 0) continue;
        const acc = String(v.acc);
        const c = byAcc[acc];
        if (!c) {
          droppedVisits += 1;
          continue;
        }
        const agent = String(c.agent ?? 'UNASSIGNED').trim() || 'UNASSIGNED';
        if (!next[agent]) next[agent] = {};
        if (!next[agent][dateStr]) next[agent][dateStr] = [];
        next[agent][dateStr].push({
          acc,
          name: String(c.name ?? v.name ?? ''),
          visitRound: normalizeVisitRound(v.visitRound ?? 1),
        });
      }
    }
  }

  return { schedules: normalizePersistedSchedules(next), droppedVisits };
}

/** 从 localStorage 恢复的 schedules 做形状清洗 */
function normalizePersistedSchedules(raw) {
  if (raw == null || typeof raw !== 'object' || Array.isArray(raw)) return {};
  const out = {};
  for (const [agentKey, byDate] of Object.entries(raw)) {
    if (!agentKey || typeof byDate !== 'object' || byDate == null || Array.isArray(byDate)) continue;
    const dates = {};
    for (const [dateStr, visits] of Object.entries(byDate)) {
      if (typeof dateStr !== 'string' || !Array.isArray(visits)) continue;
      const clean = visits
        .filter((v) => v && v.acc != null && String(v.acc).length > 0 && v.name != null && String(v.name).length > 0)
        .map((v) => ({
          acc: String(v.acc),
          name: String(v.name),
          visitRound: normalizeVisitRound(v.visitRound ?? 1),
        }));
      if (clean.length > 0) dates[dateStr] = clean;
    }
    if (Object.keys(dates).length > 0) out[String(agentKey)] = dates;
  }
  return out;
}

function buildPersistPayload(customersList, sourceFileName, schedulesObj) {
  const schedulesSafe =
    schedulesObj && typeof schedulesObj === 'object' && !Array.isArray(schedulesObj) ? schedulesObj : {};
  return {
    v: 1,
    savedAt: new Date().toISOString(),
    importFileName: sourceFileName || '',
    customers: customersList,
    schedules: schedulesSafe,
  };
}

/** 与 localStorage / Electron JSON 相同的解析 */
function hydrateStateFromStoredObject(o) {
  if (!o || typeof o !== 'object') return null;
  if (Array.isArray(o)) return { customers: o, importFileName: '', schedules: {} };
  if (Array.isArray(o.customers))
    return {
      customers: o.customers,
      importFileName: typeof o.importFileName === 'string' ? o.importFileName : '',
      schedules: normalizePersistedSchedules(o.schedules),
    };
  return null;
}

function hydrateStateFromStoredJson(raw) {
  try {
    if (raw == null || String(raw).trim() === '') return null;
    return hydrateStateFromStoredObject(JSON.parse(raw));
  } catch {
    return null;
  }
}

function hasElectronDataBridge() {
  return typeof window !== 'undefined' && window.routePlannerElectron != null;
}

/** 导入的顾客与日历排期存为浏览器内 JSON（localStorage）；Electron 下同构写入磁盘 */
const IMPORT_STORE_KEY = 'route-planner:v1:import-data';
/** 「Clear saved import」出廠預設密碼：使用者首次成功驗證後會改存於 localStorage（不清除 import 時仍保留） */
const CLEAR_SAVED_IMPORT_PASSWORD = '1234';
const CLEAR_CUSTOM_PASSWORD_KEY = 'route-planner:v1:clear-import-password-custom';
const COMPANY_NAME_STORAGE_KEY = 'route-planner:v1:company-name';

function loadStoredCompanyName() {
  try {
    if (typeof localStorage === 'undefined') return '';
    const v = localStorage.getItem(COMPANY_NAME_STORAGE_KEY);
    return v != null && String(v).trim().length > 0 ? String(v).trim() : '';
  } catch {
    return '';
  }
}

function saveStoredCompanyName(name) {
  try {
    if (typeof localStorage === 'undefined') {
      window.alert('Could not save settings (storage unavailable).');
      return false;
    }
    const trimmed = String(name ?? '').trim();
    if (trimmed === '') localStorage.removeItem(COMPANY_NAME_STORAGE_KEY);
    else localStorage.setItem(COMPANY_NAME_STORAGE_KEY, trimmed);
    return true;
  } catch {
    window.alert('Could not save settings (storage blocked or full).');
    return false;
  }
}

function loadClearImportCustomPassword() {
  try {
    if (typeof localStorage === 'undefined') return null;
    const v = localStorage.getItem(CLEAR_CUSTOM_PASSWORD_KEY);
    return v != null && String(v).trim().length > 0 ? String(v).trim() : null;
  } catch {
    return null;
  }
}

function saveClearImportCustomPassword(pwd) {
  try {
    if (typeof localStorage === 'undefined') {
      window.alert('Could not save your new password (storage unavailable).');
      return false;
    }
    localStorage.setItem(CLEAR_CUSTOM_PASSWORD_KEY, String(pwd ?? '').trim());
    return true;
  } catch {
    window.alert('Could not save your new password (storage blocked or full).');
    return false;
  }
}

/** Hidden reset (logo pin taps): removes custom password so factory default CLEAR_SAVED_IMPORT_PASSWORD applies again */
function resetClearImportPasswordToFactoryDefault() {
  try {
    if (typeof localStorage === 'undefined') return false;
    localStorage.removeItem(CLEAR_CUSTOM_PASSWORD_KEY);
    return true;
  } catch {
    return false;
  }
}

/** Electron：原生 alert 後鍵盤焦點常留在 OS；拉回視窗才能鍵入（等同使用者 minimize 再開） */
function restoreElectronWindowKeyboardFocus() {
  try {
    window.routePlannerElectron?.focusWindow?.();
  } catch {
    /* ignore */
  }
  try {
    window.focus?.();
  } catch {
    /* ignore */
  }
}

function getEffectiveClearImportPassword() {
  return loadClearImportCustomPassword() ?? CLEAR_SAVED_IMPORT_PASSWORD;
}

function loadImportStore() {
  if (typeof localStorage === 'undefined') return { customers: [], importFileName: '', schedules: {} };
  try {
    const raw = localStorage.getItem(IMPORT_STORE_KEY);
    if (!raw) return { customers: [], importFileName: '', schedules: {} };
    return (
      hydrateStateFromStoredJson(raw) ?? {
        customers: [],
        importFileName: '',
        schedules: {},
      }
    );
  } catch {
    /* ignore corrupted store */
  }
  return { customers: [], importFileName: '', schedules: {} };
}

function saveRoutePlannerPersisted(customersList, sourceFileName, schedulesObj) {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(
      IMPORT_STORE_KEY,
      JSON.stringify(buildPersistPayload(customersList, sourceFileName, schedulesObj)),
    );
  } catch (err) {
    console.warn('route-planner: persist route planner failed', err);
    window.alert('Could not save to browser storage (quota full or private mode).');
  }
}

const sidebarSectionLabel = {
  display: 'block',
  fontSize: '7px',
  fontWeight: 700,
  letterSpacing: '0.1em',
  textTransform: 'uppercase',
  color: 'rgba(148, 163, 184, 0.95)',
  marginBottom: '1px',
};

/** Clear-import 對話框內 password：強制淺色（background/color/caret），避免系統深色或 :root color-scheme 令輸入框黑底黑字「像無法輸入」 */
const clearImportModalPasswordInputStyle = {
  width: '100%',
  padding: '8px 10px',
  borderRadius: '8px',
  border: '1px solid #cbd5e1',
  fontSize: '14px',
  boxSizing: 'border-box',
  backgroundColor: '#ffffff',
  color: '#0f172a',
  caretColor: '#0f172a',
};

/** 侧栏一行一物：同名 ACC（不同店）在 React / 勾选里不能只靠 acc（会并成一笔） */
function sidebarRowStemKey(row) {
  return [
    String(row.acc ?? ''),
    String(row.name ?? ''),
    String(row.area ?? ''),
    String(row.channel ?? ''),
    String(row.category ?? ''),
  ].join('\u001f');
}

function sidebarRowListKey(row, listIndex) {
  return `${sidebarRowStemKey(row)}\u001f${listIndex}`;
}

function App() {
  const persistedOnMount = useMemo(() => {
    if (hasElectronDataBridge()) {
      return { customers: [], importFileName: '', schedules: {} };
    }
    return loadImportStore();
  }, []);
  const [customers, setCustomers] = useState(persistedOnMount.customers);
  const [agent, setAgent] = useState('');
  /** 侧栏名单筛选（与日历互不干扰） */
  const [sidebarArea, setSidebarArea] = useState('');
  const [sidebarChannelFilter, setSidebarChannelFilter] = useState('');
  const [sidebarCategoryFilter, setSidebarCategoryFilter] = useState('');
  /** 日历格子内拜访条筛选 */
  const [calendarArea, setCalendarArea] = useState('');
  const [calendarChannelFilter, setCalendarChannelFilter] = useState('');
  const [calendarCategoryFilter, setCalendarCategoryFilter] = useState('');
  /** 侧栏：勾选后仅显示该业务员名下尚未拖进日历的客户（用 ACC，对 schedules[agent] 全部日期累计） */
  const [sidebarShowUnassignedOnly, setSidebarShowUnassignedOnly] = useState(false);
  /** 勾选 Show → 弹出卡片列出各分区未上色历店数，数字旁 Select 即到侧栏 */
  const [unassignedBrowseModalOpen, setUnassignedBrowseModalOpen] = useState(false);
  /** 勿用 window.prompt：部分 Electron / 浏览器策略下 prompt 立刻返回 null，表现为按钮「无反应」 */
  const [clearImportDialogOpen, setClearImportDialogOpen] = useState(false);
  /** unlock → 驗證目前密碼；setNewPassword → 首次改密後清除；confirmErase → 已有自訂密碼時勾選並清除 */
  const [clearImportStep, setClearImportStep] = useState('unlock');
  const [clearImportPassword, setClearImportPassword] = useState('');
  const [clearImportNewPassword, setClearImportNewPassword] = useState('');
  const [clearImportNewPasswordConfirm, setClearImportNewPasswordConfirm] = useState('');
  const [clearImportAckDanger, setClearImportAckDanger] = useState(false);
  const [clearImportError, setClearImportError] = useState('');
  const dismissClearImportDialog = useCallback(() => {
    setClearImportDialogOpen(false);
    setClearImportStep('unlock');
    setClearImportPassword('');
    setClearImportNewPassword('');
    setClearImportNewPasswordConfirm('');
    setClearImportAckDanger(false);
    setClearImportError('');
  }, []);
  const [schedules, setSchedules] = useState(persistedOnMount.schedules);
  const [year, setYear] = useState(2026);
  /** 0-indexed month (aligned with Date — 4 = May) */
  const [month, setMonth] = useState(4);
  /** Highlights a calendar cell (1 … days in month); only used in Calendar view */
  const [focusedDay, setFocusedDay] = useState(null);
  /** 双击日历某日：弹层显示该日已排客户（店名 + Account） */
  const [calendarDayDetailDateStr, setCalendarDayDetailDateStr] = useState(null);
  const [viewMode, setViewMode] = useState('calendar');
  /** 打印后 +1，强制日历重新挂载，恢复被打印布局覆盖的屏幕样式 */
  const [printEpoch, setPrintEpoch] = useState(0);
  const [importFileName, setImportFileName] = useState(persistedOnMount.importFileName);
  /** 目前記憶體中的 XLSX workbook；與 IndexedDB 快取併用，重開程式可不重新選檔即可換工作表 */
  const importWorkbookRef = useRef(null);
  const [importSheetNames, setImportSheetNames] = useState([]);
  const [importSheetIndex, setImportSheetIndex] = useState(0);
  const [sidebarVisitCountFilter, setSidebarVisitCountFilter] = useState('all');
  /** When true, sidebar list hides customers with zero visits on the calendar in the toolbar Year–Month */
  const [sidebarMonthScheduledOnly, setSidebarMonthScheduledOnly] = useState(false);
  /** 日历条形「#1/#2/#3」筛选；与侧栏「当月拜访总次数」不同 */
  const [calendarChipRoundFilter, setCalendarChipRoundFilter] = useState('all');
  /** 侧栏店名 / 账号即时过滤（仅影响侧栏列表，与日历工具栏 Find 无关） */
  const [sidebarNameQuery, setSidebarNameQuery] = useState('');
  /** 日历格内拜访条的 Find（仅影响日历上已排客户，与侧栏列表无关） */
  const [calendarNameQuery, setCalendarNameQuery] = useState('');
  /** 侧栏多选：每行单独 key（同 ACC 可并存） */
  const [sidebarSelectedRowKeys, setSidebarSelectedRowKeys] = useState(() => new Set());
  /** 勾选时日历与侧栏 Area / Channel / Category 保持一致；未勾选则两套筛选独立（Find / Visits 仍各自独立） */
  const [syncSidebarCalendarFilters, setSyncSidebarCalendarFilters] = useState(false);
  /** Summary / print: day-of-month range within the selected Year–Month */
  const [summaryDayFrom, setSummaryDayFrom] = useState(1);
  const [summaryDayTo, setSummaryDayTo] = useState(31);

  const electronHydrationDoneRef = useRef(!hasElectronDataBridge());
  /** 左上角 📍：連點 9 次內將 Clear saved import 自訂密碼還原為出廠 1234（僅移除 localStorage 自訂鍵） */
  const logoPinEasterEggRef = useRef({ count: 0, idleResetId: null });
  /** ROUTE PRO 最後的 O：連點 9 次開啟隱藏 Settings（公司名稱等） */
  const logoOSettingsEasterEggRef = useRef({ count: 0, idleResetId: null });
  const [companyName, setCompanyName] = useState(() => loadStoredCompanyName());
  const [settingsDialogOpen, setSettingsDialogOpen] = useState(false);
  const [settingsDraftCompanyName, setSettingsDraftCompanyName] = useState('');
  const [settingsError, setSettingsError] = useState('');
  const persistDebouncedTimerRef = useRef(null);
  const lastPersistJsonRef = useRef('');

  useEffect(() => {
    const api = window.routePlannerElectron;
    if (!api?.load) return;
    let cancelled = false;
    api
      .load()
      .then((raw) => {
        if (cancelled) return;
        const h = hydrateStateFromStoredJson(raw);
        if (h && Array.isArray(h.customers)) {
          setCustomers(h.customers);
          setSchedules(h.schedules && typeof h.schedules === 'object' && !Array.isArray(h.schedules) ? h.schedules : {});
          const fromDisk = typeof h.importFileName === 'string' ? h.importFileName : '';
          setImportFileName((prev) => (fromDisk.trim() !== '' ? fromDisk : prev));
        }
      })
      .catch((err) => {
        console.warn('route-planner: electron datastore load failed', err);
      })
      .finally(() => {
        if (!cancelled) electronHydrationDoneRef.current = true;
      });
    return () => {
      cancelled = true;
    };
  }, []);

  /** 從 IndexedDB 還原上次選取的試算簿（不重跑匯入，只恢復 workbook + sheet 下拉） */
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      const row = await loadCachedSpreadsheet();
      if (!row || cancelled) return;
      try {
        let workbook;
        if (row.kind === 'csv') {
          let text = String(row.payload);
          if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
          workbook = XLSX.read(text, { type: 'string', raw: false });
        } else {
          workbook = XLSX.read(new Uint8Array(row.payload), { type: 'array' });
        }
        const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
        if (names.length === 0) return;
        importWorkbookRef.current = workbook;
        setImportSheetNames(names);
        const idx = Math.max(0, Math.min(Math.floor(Number(row.sheetIndex)) || 0, names.length - 1));
        setImportSheetIndex(idx);
        setImportFileName((prev) =>
          typeof prev === 'string' && prev.trim() !== '' ? prev : row.fileName || '',
        );
      } catch (e) {
        console.warn('route-planner: hydrate import workbook cache failed', e);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    lastPersistJsonRef.current = JSON.stringify(
      buildPersistPayload(customers, importFileName, schedules),
    );
  }, [customers, importFileName, schedules]);

  useEffect(() => {
    const flushSyncDisk = window.routePlannerElectron?.saveSync;
    if (!flushSyncDisk) return;
    const onHide = () => {
      clearTimeout(persistDebouncedTimerRef.current);
      try {
        flushSyncDisk(lastPersistJsonRef.current);
      } catch {
        /* ignore */
      }
    };
    const onVisibility = () => {
      if (document.visibilityState === 'hidden') onHide();
    };
    document.addEventListener('visibilitychange', onVisibility);
    window.addEventListener('beforeunload', onHide);
    return () => {
      document.removeEventListener('visibilitychange', onVisibility);
      window.removeEventListener('beforeunload', onHide);
    };
  }, []);

  useEffect(() => {
    if (!electronHydrationDoneRef.current) return;
    const sched =
      schedules && typeof schedules === 'object' && !Array.isArray(schedules) ? schedules : {};
    const json = JSON.stringify(buildPersistPayload(customers, importFileName, sched));

    if (window.routePlannerElectron?.save) {
      clearTimeout(persistDebouncedTimerRef.current);
      persistDebouncedTimerRef.current = setTimeout(() => {
        window.routePlannerElectron.save(json).catch(() => {
          window.alert('Could not save Route Planner data file.');
        });
      }, 380);
      return () => clearTimeout(persistDebouncedTimerRef.current);
    }

    saveRoutePlannerPersisted(customers, importFileName, sched);
    return undefined;
  }, [customers, importFileName, schedules]);

  const salesAgentOptions = useMemo(() => {
    const s = new Set();
    for (const c of customers) {
      const a = String(c.agent ?? '').trim();
      if (a !== '') s.add(a);
    }
    for (const key of Object.keys(schedules ?? {})) {
      const k = String(key).trim();
      if (k !== '') s.add(k);
    }
    return [...s].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
  }, [customers, schedules]);

  useEffect(() => {
    setAgent((cur) => (!cur ? cur : salesAgentOptions.includes(cur) ? cur : ''));
  }, [salesAgentOptions]);

  const filteredAreas = useMemo(() => {
    if (!agent) return [];
    return [...new Set(customers.filter(c => c.agent === agent).map(c => c.area))].sort();
  }, [agent, customers]);

  const visitCountsTotal = useMemo(() => {
    const counts = {};
    if (agent && schedules[agent]) {
      Object.values(schedules[agent]).forEach((dayList) => {
        dayList.forEach((c) => {
          counts[c.acc] = (counts[c.acc] || 0) + 1;
        });
      });
    }
    return counts;
  }, [agent, schedules]);

  /** 仅当前工具栏所选年月的排期次数 — Visits 筛选与日历格子一致 */
  const visitCountsMonth = useMemo(() => {
    const counts = {};
    if (agent && schedules[agent]) {
      const prefix = `${year}-${String(month + 1).padStart(2, '0')}-`;
      for (const [d, dayList] of Object.entries(schedules[agent])) {
        if (!d.startsWith(prefix)) continue;
        dayList.forEach((c) => {
          counts[c.acc] = (counts[c.acc] || 0) + 1;
        });
      }
    }
    return counts;
  }, [agent, schedules, year, month]);

  const agentScopeCustomers = useMemo(
    () => (agent ? customers.filter((c) => c.agent === agent) : []),
    [agent, customers],
  );

  /** Unassigned 红条与侧栏名单：同一套 Area / Channel / Category 过滤 */
  const unassignedStripScopePool = useMemo(() => {
    if (!agent) return [];
    return agentScopeCustomers.filter((c) =>
      matchesSidebarListFilters(c, sidebarArea, sidebarChannelFilter, sidebarCategoryFilter),
    );
  }, [agent, agentScopeCustomers, sidebarArea, sidebarChannelFilter, sidebarCategoryFilter]);

  const unassignedStripFiltersActive = Boolean(
    sidebarArea || sidebarChannelFilter || sidebarCategoryFilter,
  );

  /** 同上筛选范围内：日历上任一日都尚无排期的客户数 */
  const unassignedCustomerCount = useMemo(() => {
    if (!agent) return 0;
    return unassignedStripScopePool.filter((c) => (visitCountsTotal[c.acc] ?? 0) === 0).length;
  }, [agent, unassignedStripScopePool, visitCountsTotal]);

  const scheduledOnCalendarCount = useMemo(() => {
    if (!agent) return 0;
    return unassignedStripScopePool.filter((c) => (visitCountsTotal[c.acc] ?? 0) > 0).length;
  }, [agent, unassignedStripScopePool, visitCountsTotal]);

  const agentCustomerTotalCount = unassignedStripScopePool.length;

  /** Browse 弹窗：仅用 Channel / Category（不按侧栏 Area）筛出未上色历客户，再按 region 计数 */
  const unassignedPoolForBrowseModal = useMemo(() => {
    if (!agent) return [];
    return agentScopeCustomers.filter(
      (c) =>
        (visitCountsTotal[c.acc] ?? 0) === 0 &&
        matchesSidebarListFilters(c, '', sidebarChannelFilter, sidebarCategoryFilter),
    );
  }, [agent, agentScopeCustomers, visitCountsTotal, sidebarChannelFilter, sidebarCategoryFilter]);

  const unassignedByAreaRowsBrowse = useMemo(() => {
    const tally = {};
    for (const c of unassignedPoolForBrowseModal) {
      const ar = String(c.area ?? '');
      tally[ar] = (tally[ar] || 0) + 1;
    }
    const rows = Object.entries(tally).map(([areaLabel, count]) => ({ areaLabel, count }));
    rows.sort((a, b) =>
      a.areaLabel.localeCompare(b.areaLabel, undefined, { sensitivity: 'base', numeric: true }),
    );
    return rows;
  }, [unassignedPoolForBrowseModal]);

  const unassignedBrowseModalTotal = unassignedPoolForBrowseModal.length;

  /** Browse 卡片表格：仅列出仍有未上色历店的分区（不再为「当前侧栏 Area」强插一行 0，否则会误以为该区仍有待排） */
  const unassignedBrowseSelectAreaRows = useMemo(() => {
    const base = unassignedByAreaRowsBrowse.map((r) => ({ ...r }));
    const labels = new Set(base.map((r) => r.areaLabel));
    const sid = sidebarArea ?? '';
    if (agent && sid !== '' && !labels.has(String(sid))) {
      let cnt = 0;
      for (const c of unassignedPoolForBrowseModal) {
        if (String(c.area ?? '') === String(sid)) cnt++;
      }
      if (cnt > 0) {
        base.push({ areaLabel: sid, count: cnt });
        base.sort((a, b) =>
          a.areaLabel.localeCompare(b.areaLabel, undefined, { sensitivity: 'base', numeric: true }),
        );
      }
    }
    return base;
  }, [agent, unassignedByAreaRowsBrowse, unassignedPoolForBrowseModal, sidebarArea]);

  /** 侧栏 Area 下拉：勾选 Unassigned Show 时用「尚未上日历」池 */
  const sidebarAreaOptionsForSelect = useMemo(() => {
    const pool =
      sidebarShowUnassignedOnly && agent
        ? agentScopeCustomers.filter((c) => (visitCountsTotal[c.acc] ?? 0) === 0)
        : agent
          ? customers.filter((c) => c.agent === agent)
          : [];
    return [...new Set(pool.map((c) => c.area))].sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }),
    );
  }, [sidebarShowUnassignedOnly, agent, agentScopeCustomers, customers, visitCountsTotal]);

  const sidebarChannelOptions = useMemo(() => {
    const pool =
      sidebarShowUnassignedOnly && agent
        ? agentScopeCustomers.filter((c) => (visitCountsTotal[c.acc] ?? 0) === 0)
        : agent
          ? customers.filter((c) => c.agent === agent)
          : [];
    if (pool.length === 0) return [];
    const set = new Set(
      pool
        .filter((c) => !sidebarArea || c.area === sidebarArea)
        .map((c) => String(c.channel ?? '—')),
    );
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
  }, [sidebarShowUnassignedOnly, agent, agentScopeCustomers, customers, sidebarArea, visitCountsTotal]);

  const calendarChannelOptions = useMemo(() => {
    if (!agent) return [];
    const set = new Set(
      customers
        .filter((c) => c.agent === agent && (!calendarArea || c.area === calendarArea))
        .map((c) => String(c.channel ?? '—')),
    );
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
  }, [agent, calendarArea, customers]);

  const sidebarCategoryOptions = useMemo(() => {
    const pool =
      sidebarShowUnassignedOnly && agent
        ? agentScopeCustomers.filter((c) => (visitCountsTotal[c.acc] ?? 0) === 0)
        : agent
          ? customers.filter((c) => c.agent === agent)
          : [];
    if (pool.length === 0) return [];
    const set = new Set(
      pool
        .filter(
          (c) =>
            (!sidebarArea || c.area === sidebarArea) &&
            (!sidebarChannelFilter || String(c.channel ?? '—') === sidebarChannelFilter),
        )
        .map((c) => String(c.category ?? '—')),
    );
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
  }, [
    sidebarShowUnassignedOnly,
    agent,
    agentScopeCustomers,
    customers,
    sidebarArea,
    sidebarChannelFilter,
    visitCountsTotal,
  ]);

  const calendarCategoryOptions = useMemo(() => {
    if (!agent) return [];
    const set = new Set(
      customers
        .filter(
          (c) =>
            c.agent === agent &&
            (!calendarArea || c.area === calendarArea) &&
            (!calendarChannelFilter || String(c.channel ?? '—') === calendarChannelFilter),
        )
        .map((c) => String(c.category ?? '—')),
    );
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
  }, [agent, calendarArea, calendarChannelFilter, customers]);

  const sidebarBaseCustomers = useMemo(() => {
    if (sidebarShowUnassignedOnly) {
      if (!agent) return [];
      return customers.filter(
        (c) =>
          c.agent === agent &&
          (visitCountsTotal[c.acc] ?? 0) === 0 &&
          matchesSidebarListFilters(c, sidebarArea, sidebarChannelFilter, sidebarCategoryFilter),
      );
    }
    if (!agent) return [];
    return customers.filter(
      (c) =>
        c.agent === agent &&
        matchesSidebarListFilters(c, sidebarArea, sidebarChannelFilter, sidebarCategoryFilter),
    );
  }, [
    sidebarShowUnassignedOnly,
    agent,
    customers,
    visitCountsTotal,
    sidebarArea,
    sidebarChannelFilter,
    sidebarCategoryFilter,
  ]);

  const sidebarListedCustomers = useMemo(() => {
    let list = sidebarBaseCustomers.filter((c) =>
      matchesVisitScheduleFilter(visitCountsMonth[c.acc], sidebarVisitCountFilter),
    );
    if (sidebarMonthScheduledOnly && agent) {
      list = list.filter((c) => (visitCountsMonth[c.acc] ?? 0) > 0);
    }
    return list;
  }, [
    sidebarBaseCustomers,
    visitCountsMonth,
    sidebarVisitCountFilter,
    sidebarMonthScheduledOnly,
    agent,
  ]);

  const sidebarDisplayedCustomers = useMemo(() => {
    const raw = normalizeSidebarSearchPiece(sidebarNameQuery);
    const byNameThenAcc = (a, b) => {
      const o = String(a.name ?? '').localeCompare(String(b.name ?? ''), undefined, {
        sensitivity: 'base',
        numeric: true,
      });
      if (o !== 0) return o;
      return String(a.acc ?? '').localeCompare(String(b.acc ?? ''), undefined, {
        sensitivity: 'base',
        numeric: true,
      });
    };

    /** 输入 Find 时在「整条名单」里搜（仍会受 Agent/Area/Channel/Category 制约），避免因 Visits（如 Exactly 2）把尚无排期的新客先筛掉而导致搜不到 */
    const sourceForFind = raw ? sidebarBaseCustomers : sidebarListedCustomers;

    let list =
      !raw
        ? [...sidebarListedCustomers]
        : sourceForFind.filter((c) => {
            const nm = normalizeSidebarSearchPiece(c.name);
            const ac = normalizeSidebarSearchPiece(c.acc);
            return nm.includes(raw) || ac.includes(raw);
          });

    list.sort(byNameThenAcc);
    return list;
  }, [sidebarBaseCustomers, sidebarListedCustomers, sidebarNameQuery]);

  useEffect(() => {
    setSidebarSelectedRowKeys(new Set());
    setSidebarNameQuery('');
    setCalendarNameQuery('');
    setSidebarArea('');
    setSidebarChannelFilter('');
    setSidebarCategoryFilter('');
    setCalendarArea('');
    setCalendarChannelFilter('');
    setCalendarCategoryFilter('');
    setSyncSidebarCalendarFilters(false);
    setSidebarShowUnassignedOnly(false);
    setSidebarMonthScheduledOnly(false);
    setUnassignedBrowseModalOpen(false);
    setCalendarDayDetailDateStr(null);
  }, [agent]);

  /** Unassigned sidebar entries have no calendar visits → month-only list is meaningless */
  useEffect(() => {
    if (sidebarShowUnassignedOnly) setSidebarMonthScheduledOnly(false);
  }, [sidebarShowUnassignedOnly]);

  useEffect(() => {
    if (!sidebarShowUnassignedOnly || !agent) return;
    if (unassignedCustomerCount !== 0) return;
    setSidebarShowUnassignedOnly(false);
  }, [sidebarShowUnassignedOnly, agent, unassignedCustomerCount]);

  useEffect(() => {
    setSidebarChannelFilter('');
    setSidebarCategoryFilter('');
  }, [sidebarArea]);

  useEffect(() => {
    setSidebarCategoryFilter('');
  }, [sidebarChannelFilter]);

  useEffect(() => {
    /** Sync 开启时日历区筛选由侧栏镜像，勿在 Area 变动时清空 Ch/Cat（否则「same as sidebar」一打勾就会被抹掉） */
    if (syncSidebarCalendarFilters) return;
    setCalendarChannelFilter('');
    setCalendarCategoryFilter('');
  }, [calendarArea, syncSidebarCalendarFilters]);

  useEffect(() => {
    if (syncSidebarCalendarFilters) return;
    setCalendarCategoryFilter('');
  }, [calendarChannelFilter, syncSidebarCalendarFilters]);

  /** Sidebar → Calendar（Sync）：仅以侧栏为真相写入日历筛选；不设日历→侧栏（会与级联清空 effect 冲突） */
  useEffect(() => {
    if (!syncSidebarCalendarFilters) return;
    setCalendarArea(sidebarArea);
    setCalendarChannelFilter(sidebarChannelFilter);
    setCalendarCategoryFilter(sidebarCategoryFilter);
    setCalendarNameQuery(sidebarNameQuery);
    setCalendarChipRoundFilter(mapSidebarVisitsFilterToCalendarChipFilter(sidebarVisitCountFilter));
  }, [
    syncSidebarCalendarFilters,
    sidebarArea,
    sidebarChannelFilter,
    sidebarCategoryFilter,
    sidebarNameQuery,
    sidebarVisitCountFilter,
  ]);

  const sidebarSelectAllRef = useRef(null);
  useEffect(() => {
    const el = sidebarSelectAllRef.current;
    if (!el) return;
    const all =
      sidebarDisplayedCustomers.length > 0 &&
      sidebarDisplayedCustomers.every((row, idx) =>
        sidebarSelectedRowKeys.has(sidebarRowListKey(row, idx)),
      );
    const some = sidebarDisplayedCustomers.some((row, idx) =>
      sidebarSelectedRowKeys.has(sidebarRowListKey(row, idx)),
    );
    el.checked = all;
    el.indeterminate = some && !all;
  }, [sidebarDisplayedCustomers, sidebarSelectedRowKeys]);

  /** 侧栏：toolbar 所选当月排期次数 → 调色桶与角标（与「Visits」筛选、日历月份一致） */
  const sidebarColorBucket = (monthVisits) => {
    if (monthVisits <= 0) return 0;
    return Math.min(monthVisits, 6);
  };
  const monthIsoPad = String(month + 1).padStart(2, '0');
  const monthLabel = `${year}-${MONTH_SHORT[month]}`;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const highlightedDay =
    focusedDay != null && focusedDay >= 1 && focusedDay <= daysInMonth ? focusedDay : null;

  useLayoutEffect(() => {
    if (viewMode !== 'calendar' || highlightedDay == null) return;
    const el = document.querySelector(`[data-cal-focus-day="${highlightedDay}"]`);
    if (!(el instanceof HTMLElement)) return;
    const id = window.requestAnimationFrame(() => {
      el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    });
    return () => cancelAnimationFrame(id);
  }, [viewMode, highlightedDay, year, month]);

  /** 日历双击弹层：该日 raw 排期（不受 Area / Find 筛选） */
  const calendarDayCardTasks = useMemo(() => {
    if (calendarDayDetailDateStr == null || !agent) return [];
    const raw = schedules[agent]?.[calendarDayDetailDateStr];
    if (!Array.isArray(raw)) return [];
    return [...raw].sort((a, b) => {
      const ra = normalizeVisitRound(a.visitRound ?? 1);
      const rb = normalizeVisitRound(b.visitRound ?? 1);
      if (ra !== rb) return ra - rb;
      return String(a.acc ?? '').localeCompare(String(b.acc ?? ''), undefined, {
        sensitivity: 'base',
        numeric: true,
      });
    });
  }, [calendarDayDetailDateStr, agent, schedules]);

  const calendarWeeks = useMemo(() => {
    const fd = new Date(year, month, 1).getDay();
    const td = new Date(year, month + 1, 0).getDate();
    return monthWeekSlots(fd, td);
  }, [year, month]);

  /** Summary 日与年月绑定：若仅就地 clamp「至」某日，到过短月后「至」会一直偏低，长月份就看不到末尾几天 */
  useEffect(() => {
    setSummaryDayFrom(1);
    setSummaryDayTo(daysInMonth);
  }, [year, month, daysInMonth]);

  useEffect(() => {
    if (calendarDayDetailDateStr == null) return;
    const onKey = (e) => {
      if (e.key === 'Escape') setCalendarDayDetailDateStr(null);
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [calendarDayDetailDateStr]);

  useEffect(() => {
    if (!unassignedBrowseModalOpen) return;
    const onKey = (e) => {
      if (e.key === 'Escape') {
        setUnassignedBrowseModalOpen(false);
        setSidebarShowUnassignedOnly(false);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [unassignedBrowseModalOpen]);

  useEffect(() => {
    if (!clearImportDialogOpen) return;
    const onKey = (e) => {
      if (e.key === 'Escape') dismissClearImportDialog();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [clearImportDialogOpen, dismissClearImportDialog]);

  /** Electron／某些環境：對話框開啟後焦點未進輸入框，會像「無法輸入密碼」 */
  useLayoutEffect(() => {
    if (!clearImportDialogOpen) return;
    const pickId =
      clearImportStep === 'unlock'
        ? 'clear-import-password'
        : clearImportStep === 'setNewPassword'
          ? 'clear-import-new-password'
          : null;
    if (!pickId) return;

    let cancelled = false;
    const pullBrowserFocusThenField = () => {
      if (cancelled) return;
      restoreElectronWindowKeyboardFocus();
      /** 主進程 focus 完成後再等兩幀才把 DOM focus 進 input ，避免順序競態 */
      window.requestAnimationFrame(() => {
        window.requestAnimationFrame(() => {
          if (cancelled) return;
          const el = document.getElementById(pickId);
          if (el && typeof el.focus === 'function') {
            try {
              el.focus({ preventScroll: true });
            } catch {
              el.focus();
            }
          }
        });
      });
    };

    let timeout40 = 0;
    let timeout120 = 0;
    const frameId = window.requestAnimationFrame(() => {
      pullBrowserFocusThenField();
      timeout40 = window.setTimeout(pullBrowserFocusThenField, 40);
      timeout120 = window.setTimeout(pullBrowserFocusThenField, 120);
    });

    return () => {
      cancelled = true;
      window.cancelAnimationFrame(frameId);
      window.clearTimeout(timeout40);
      window.clearTimeout(timeout120);
    };
  }, [clearImportDialogOpen, clearImportStep]);

  useEffect(() => {
    return () => {
      const pinId = logoPinEasterEggRef.current.idleResetId;
      if (pinId != null) window.clearTimeout(pinId);
      const oId = logoOSettingsEasterEggRef.current.idleResetId;
      if (oId != null) window.clearTimeout(oId);
    };
  }, []);

  useEffect(() => {
    const base = 'Route Planner';
    document.title = companyName.trim() ? `${companyName.trim()} — ${base}` : base;
  }, [companyName]);

  useEffect(() => {
    const fitPrintToOnePage = () => {
      preparePrintLayout();
      refitPrintContent();
    };
    const restoreAfterPrint = () => {
      resetPrintScale();
      setPrintEpoch((n) => n + 1);
    };

    window.addEventListener('beforeprint', fitPrintToOnePage);
    window.addEventListener('afterprint', restoreAfterPrint);
    return () => {
      window.removeEventListener('beforeprint', fitPrintToOnePage);
      window.removeEventListener('afterprint', restoreAfterPrint);
    };
  }, []);

  const handlePrintPdf = useCallback(() => {
    runPrintWithLayout(() => setPrintEpoch((n) => n + 1));
  }, []);

  const dismissSettingsDialog = useCallback(() => {
    setSettingsDialogOpen(false);
    setSettingsDraftCompanyName('');
    setSettingsError('');
  }, []);

  const handleLogoOSettingsEasterEggClick = useCallback(
    (e) => {
      e.stopPropagation();
      const egg = logoOSettingsEasterEggRef.current;
      if (egg.idleResetId != null) window.clearTimeout(egg.idleResetId);
      egg.count += 1;
      egg.idleResetId = window.setTimeout(() => {
        egg.count = 0;
        egg.idleResetId = null;
      }, 4200);
      if (egg.count < 9) return;
      egg.count = 0;
      if (egg.idleResetId != null) window.clearTimeout(egg.idleResetId);
      egg.idleResetId = null;
      setSettingsDraftCompanyName(companyName);
      setSettingsError('');
      setSettingsDialogOpen(true);
    },
    [companyName],
  );

  const submitSettingsSave = useCallback(() => {
    const trimmed = String(settingsDraftCompanyName).trim();
    if (trimmed.length > 120) {
      setSettingsError('Company name must be 120 characters or fewer.');
      return;
    }
    if (!saveStoredCompanyName(trimmed)) return;
    setCompanyName(trimmed);
    dismissSettingsDialog();
    restoreElectronWindowKeyboardFocus();
  }, [settingsDraftCompanyName, dismissSettingsDialog]);

  /** Unassigned browse 卡片：套用分区并关掉弹层（侧栏只显示未上色历顾客） */
  const applyUnassignedBrowseAreaToSidebar = useCallback((areaLabel) => {
    setSidebarArea(areaLabel);
    setSidebarShowUnassignedOnly(true);
    setUnassignedBrowseModalOpen(false);
  }, []);

  const summaryPrintRange = useMemo(() => {
    let from = Math.min(Math.max(summaryDayFrom, 1), daysInMonth);
    let to = Math.min(Math.max(summaryDayTo, 1), daysInMonth);
    if (from > to) {
      const swap = from;
      from = to;
      to = swap;
    }
    return { from, to };
  }, [summaryDayFrom, summaryDayTo, daysInMonth]);

  const customerByAcc = useMemo(() => {
    const map = {};
    customers.forEach((c) => { map[c.acc] = c; });
    return map;
  }, [customers]);

  /** 日历格内拜访条：日历专用 Area / Channel / Category / Find；条形序号用 calendarChipRoundFilter */
  const calendarFilterAcc = useCallback(
    (acc) => {
      const c = customerByAcc[acc];
      if (!c) return true;
      if (calendarArea && c.area !== calendarArea) return false;
      if (calendarChannelFilter && String(c.channel ?? '—') !== calendarChannelFilter) return false;
      if (calendarCategoryFilter && String(c.category ?? '—') !== calendarCategoryFilter) return false;
      const raw = normalizeSidebarSearchPiece(calendarNameQuery);
      if (raw) {
        const nm = normalizeSidebarSearchPiece(c.name);
        const ac = normalizeSidebarSearchPiece(acc);
        if (!nm.includes(raw) && !ac.includes(raw)) return false;
      }
      return true;
    },
    [
      customerByAcc,
      calendarArea,
      calendarChannelFilter,
      calendarCategoryFilter,
      calendarNameQuery,
    ],
  );

  const summarySections = useMemo(() => {
    if (!agent) return null;
    const prefix = `${year}-${monthIsoPad}-`;
    const agentSched = schedules[agent];
    const entries =
      agentSched == null ? [] :
        Object.entries(agentSched)
          .filter(([d]) => d.startsWith(prefix))
          .flatMap(([dateStr, visits]) =>
            visits.map((v) => ({
              dateStr,
              weekday: weekdayAbbrevFromIso(dateStr),
              acc: v.acc,
              name: v.name,
              area: customerByAcc[v.acc]?.area ?? '—',
              visitRound: normalizeVisitRound(v.visitRound ?? 1),
            })),
          )
          .sort((a, b) => (a.dateStr === b.dateStr ? a.acc.localeCompare(b.acc) : a.dateStr.localeCompare(b.dateStr)));

    const byDate = {};
    entries.forEach((row) => {
      if (!byDate[row.dateStr]) byDate[row.dateStr] = [];
      byDate[row.dateStr].push(row);
    });

    const dateKeys = Object.keys(byDate).sort();
    const uniqueAcc = new Set(entries.map((r) => r.acc));
    return {
      entries,
      byDate,
      dateKeys,
      totalVisits: entries.length,
      uniqueCustomers: uniqueAcc.size,
      daysScheduled: dateKeys.length,
    };
  }, [agent, schedules, year, monthIsoPad, customerByAcc]);

  const summaryFiltered = useMemo(() => {
    if (!summarySections) return null;
    const { from, to } = summaryPrintRange;
    const dateKeys = summarySections.dateKeys.filter((d) => {
      const day = Number(d.slice(-2));
      return Number.isFinite(day) && day >= from && day <= to;
    });
    const byDate = {};
    const entries = [];
    dateKeys.forEach((d) => {
      const rows = summarySections.byDate[d];
      byDate[d] = rows;
      entries.push(...rows);
    });
    const uniqueAcc = new Set(entries.map((r) => r.acc));
    return {
      byDate,
      dateKeys,
      totalVisits: entries.length,
      uniqueCustomers: uniqueAcc.size,
      daysScheduled: dateKeys.length,
    };
  }, [summarySections, summaryPrintRange]);

  const renderSummary = () => {
    if (!agent) {
      return (
        <div className="summary-print-wrap summary-print-empty" style={{ padding: '24px', color: '#64748b', flex: 1 }}>
          Choose a sales agent to see visits scheduled for this month.
        </div>
      );
    }
    if (!summarySections || summarySections.totalVisits === 0) {
      return (
        <div className="summary-print-wrap summary-print-empty" style={{ padding: '24px', color: '#64748b', flex: 1 }}>
          No visits on the calendar for <strong style={{ color: '#1e293b' }}>{agent}</strong> in {monthLabel}. Drag customers from the sidebar onto dates in Calendar.
        </div>
      );
    }

    const { from, to } = summaryPrintRange;
    const filtered = summaryFiltered;
    if (!filtered || filtered.totalVisits === 0) {
      return (
        <div className="summary-print-wrap summary-print-empty" style={{ padding: '24px', color: '#64748b', flex: 1, lineHeight: 1.5 }}>
          No visits between <strong style={{ color: '#1e293b' }}>day {from}</strong> and{' '}
          <strong style={{ color: '#1e293b' }}>day {to}</strong> for{' '}
          <strong style={{ color: '#1e293b' }}>{agent}</strong> in {monthLabel}. Expand the Summary range above or schedule visits on those dates in Calendar.
        </div>
      );
    }

    const { byDate, dateKeys, totalVisits, uniqueCustomers, daysScheduled } = filtered;
    const cell = { padding: '10px 12px', borderBottom: '1px solid #e2e8f0', fontSize: '13px' };
    const summaryRangeTitle =
      from === 1 && to === daysInMonth ? monthLabel : `${monthLabel}, days ${from}–${to}`;

    // Print pagination: group dates into calendar weeks (Sun–Sat, same as the
    // calendar grid), two weeks per printed A4 page.
    const firstDow = new Date(`${year}-${monthIsoPad}-01T00:00:00`).getDay();
    const pageIdxOf = (dateStr) =>
      Math.floor(Math.floor((Number(dateStr.slice(-2)) - 1 + firstDow) / 7) / 2);
    const pageGroups = [];
    dateKeys.forEach((d) => {
      const p = pageIdxOf(d);
      const last = pageGroups[pageGroups.length - 1];
      if (last && last.pageIdx === p) last.dateKeys.push(d);
      else pageGroups.push({ pageIdx: p, dateKeys: [d] });
    });

    const renderVisitTable = (keys) => (
      <div className="summary-print-table-host" style={{ border: '1px solid #e2e8f0', borderRadius: '8px', overflow: 'hidden', background: 'white' }}>
        <table className="summary-print-table" style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr style={{ background: '#1e293b', color: 'white', textAlign: 'left', fontSize: '12px' }}>
              <th style={{ ...cell, fontWeight: 600 }}>Date</th>
              <th style={{ ...cell, fontWeight: 600 }}>#</th>
              <th style={{ ...cell, fontWeight: 600 }}>Account</th>
              <th style={{ ...cell, fontWeight: 600 }}>Customer</th>
              <th style={{ ...cell, fontWeight: 600 }}>Area</th>
            </tr>
          </thead>
          <tbody>
            {keys.flatMap((dateStr) =>
              byDate[dateStr].map((row, idx) => {
                const d = row.weekday ?? weekdayAbbrevFromIso(dateStr);
                const dayNum = Number(dateStr.slice(-2));
                const showingDate = idx === 0;
                return (
                  <tr key={`${dateStr}-${row.acc}-${idx}`} style={{ background: 'white' }}>
                    <td style={{ ...cell, color: '#0f172a', fontWeight: showingDate ? 600 : 400, verticalAlign: 'top', whiteSpace: 'nowrap' }}>
                      {showingDate ? `${d} ${dayNum}` : ''}
                    </td>
                    <td style={{ ...cell, color: '#475569', fontWeight: 600, verticalAlign: 'top' }}>{normalizeVisitRound(row.visitRound)}</td>
                    <td style={{ ...cell, color: '#334155', verticalAlign: 'top' }}>{row.acc}</td>
                    <td style={{ ...cell, color: '#1e293b', verticalAlign: 'top' }}>{row.name}</td>
                    <td style={{ ...cell, color: '#64748b', verticalAlign: 'top' }}>{row.area}</td>
                  </tr>
                );
              }),
            )}
          </tbody>
        </table>
      </div>
    );

    return (
      <div
        className="summary-print-wrap"
        style={{
          width: '100%',
          alignSelf: 'stretch',
          flexShrink: 0,
          boxSizing: 'border-box',
          padding: '20px',
          display: 'flex',
          flexDirection: 'column',
          gap: '20px',
        }}
      >
        {pageGroups.map((pg, pi) => {
          const dayA = Number(pg.dateKeys[0].slice(-2));
          const dayB = Number(pg.dateKeys[pg.dateKeys.length - 1].slice(-2));
          return (
            <div
              key={`sum-page-${pg.pageIdx}`}
              className="summary-print-page"
              style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}
            >
              {pi === 0 ? (
                <>
                  <div className="summary-print-page-title">
                    {companyName.trim() ? (
                      <span className="app-company-display">{companyName.trim()}</span>
                    ) : null}
                    <span className="summary-print-page-title-meta">
                      Schedule summary · {summaryRangeTitle}
                      {agent ? ` · ${agent}` : ''}
                    </span>
                  </div>
                  <div className="summary-print-kpi-row" style={{ display: 'flex', gap: '16px', flexWrap: 'wrap' }}>
                    {[
                      { label: 'Total visits', value: totalVisits },
                      { label: 'Customers', value: uniqueCustomers },
                      { label: 'Days with visits', value: daysScheduled },
                    ].map((kpi) => (
                      <div
                        key={kpi.label}
                        className="summary-print-kpi"
                        style={{
                          flex: '1 1 120px',
                          minWidth: '100px',
                          background: '#f8fafc',
                          border: '1px solid #e2e8f0',
                          borderRadius: '8px',
                          padding: '14px 16px',
                        }}
                      >
                        <div className="summary-print-kpi-label" style={{ fontSize: '11px', textTransform: 'uppercase', color: '#64748b', marginBottom: '4px' }}>{kpi.label}</div>
                        <div className="summary-print-kpi-value" style={{ fontSize: '22px', fontWeight: '700', color: '#1e293b' }}>{kpi.value}</div>
                      </div>
                    ))}
                  </div>
                </>
              ) : (
                <div className="summary-print-page-title">
                  <span className="summary-print-page-title-meta">
                    Schedule summary · {summaryRangeTitle}
                    {agent ? ` · ${agent}` : ''} · days {dayA}–{dayB}
                  </span>
                </div>
              )}
              {renderVisitTable(pg.dateKeys)}
            </div>
          );
        })}
      </div>
    );
  };

  const mapImportedRowsToCustomers = (data) =>
    data.map((r, idx) => {
      const ks = Object.keys(r);
      const f = (ts) => ks.find((k) => ts.some((t) => k.toUpperCase().includes(t.toUpperCase())));
      const nameKey = pickDebtorNameColumnKey(ks);
      const rawName = nameKey != null ? r[nameKey] : undefined;
      const name =
        rawName != null && String(rawName).trim() !== '' ? String(rawName).trim() : 'N/A';
      const rawAcc = r[f(['ACC NO', 'ACCOUNT'])];
      const accTrimmed = rawAcc != null ? String(rawAcc).trim() : '';
      /** 試算表格內留白：舊版一律變「N/A」會變成一樣嘅字串，duplicate 檢查會報錯、整份匯入失敗且無唯一客戶鍵 */
      const acc =
        accTrimmed !== '' ? accTrimmed : `NOACC-${idx + 1}`;
      return {
        acc,
        name,
        area: r[f(['AREA'])] || 'OTHER',
        agent: r[f(['SALES AGENT', 'AGENT'])] || 'UNASSIGNED',
        channel: normalizeImportField(r[f(['CHANNEL'])]),
        category: normalizeImportField(r[f(['CATEGORY', 'PRODUCT CATEGORY'])]),
      };
    });

  /**
   * @param {*} workbook XLSX workbook
   * @param {number} sheetIndex SheetNames index
   * @param {string} displayFileName Filename shown next to picker
   * @param {HTMLInputElement | null} input File input — duplicate / error 時可清空並可選清空暫存 workbook
   * @param {{ clearStoredWorkbookOnDup?: boolean }} opts 僅選新檔失敗時清空暫存，避免切換 sheet 失敗拆掉目前 workbook
   */
  const applyImportedWorkbook = (workbook, sheetIndex, displayFileName, input, opts = {}) => {
    const clearStoredWorkbookOnDup = opts.clearStoredWorkbookOnDup ?? false;
    const names = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
    if (names.length === 0) {
      window.alert('No worksheets found in this file.');
      if (clearStoredWorkbookOnDup) {
        importWorkbookRef.current = null;
        setImportSheetNames([]);
        setImportSheetIndex(0);
        setImportFileName('');
      }
      if (input && typeof input.value === 'string') input.value = '';
      return false;
    }
    const idx = Math.max(0, Math.min(Math.floor(Number(sheetIndex)) || 0, names.length - 1));
    const sheetKey = names[idx];
    if (!sheetKey || !workbook.Sheets[sheetKey]) {
      window.alert('This sheet could not be read.');
      return false;
    }
    try {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetKey]);
      const mapped = mapImportedRowsToCustomers(rows);
      const dupAccs = findDuplicateImportAccs(mapped);
      if (dupAccs.length > 0) {
        const sample = dupAccs.slice(0, 12).join(', ');
        const more = dupAccs.length > 12 ? `\n… and ${dupAccs.length - 12} more.` : '';
        window.alert(
          `Import cancelled: each ACC (account number) must be unique in the spreadsheet.\n\nDuplicate ACC values (${dupAccs.length}): ${sample}${more}\n\nFix the file and import again.`,
        );
        if (clearStoredWorkbookOnDup) {
          importWorkbookRef.current = null;
          setImportSheetNames([]);
          setImportSheetIndex(0);
        }
        if (input && typeof input.value === 'string') input.value = '';
        return false;
      }

      let droppedAfterMerge = 0;
      importWorkbookRef.current = workbook;
      setImportSheetNames(names);
      setImportSheetIndex(idx);

      setCustomers(mapped);
      setImportFileName(displayFileName);
      setSchedules((prev) => {
        const merged = mergeSchedulesAfterCustomerReimport(prev, mapped);
        droppedAfterMerge = merged.droppedVisits;
        return merged.schedules;
      });
      /** 新来的客多为 0 次拜访；沿用「Exactly N visits」会像被吃掉一样 */
      setSidebarVisitCountFilter('all');
      setCalendarChipRoundFilter('all');
      setSidebarArea('');
      setCalendarArea('');
      setSidebarChannelFilter('');
      setSidebarCategoryFilter('');
      setCalendarChannelFilter('');
      setCalendarCategoryFilter('');
      setSidebarNameQuery('');
      setCalendarNameQuery('');
      setSyncSidebarCalendarFilters(false);
      setSidebarShowUnassignedOnly(false);
      setUnassignedBrowseModalOpen(false);
      if (droppedAfterMerge > 0) {
        window.alert(
          `Imported ${mapped.length} customers.\n\n` +
            `Calendar visits were kept for every account still listed in this file (dates and visit # unchanged; names refreshed from the sheet).\n\n` +
            `${droppedAfterMerge} visit${droppedAfterMerge === 1 ? '' : 's'} had to be dropped because ${droppedAfterMerge === 1 ? 'that account is' : 'those accounts are'} no longer in the spreadsheet.`,
        );
      }
      void patchCachedSpreadsheetSheetIndex(idx).catch(() => {});
      if (input && typeof input.value === 'string') input.value = '';
      return true;
    } catch (err) {
      console.error(err);
      window.alert('Could not read this sheet.');
      return false;
    }
  };

  const handleFile = (e) => {
    const input = e.target;
    const file = input?.files?.[0];
    if (!file) {
      setImportFileName('');
      importWorkbookRef.current = null;
      setImportSheetNames([]);
      setImportSheetIndex(0);
      return;
    }

    const ext = (file.name.split('.').pop() ?? '').toLowerCase();
    const isCsv = ext === 'csv';

    const reader = new FileReader();
    reader.onload = () => {
      try {
        let workbook;
        /** @type {{ kind: 'csv'|'binary', payload: string|ArrayBuffer}|null} */
        let cacheBits = null;
        if (isCsv) {
          let text = String(reader.result ?? '');
          if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
          cacheBits = { kind: 'csv', payload: text };
          workbook = XLSX.read(text, { type: 'string', raw: false });
        } else {
          const raw = reader.result;
          if (!(raw instanceof ArrayBuffer)) {
            throw new Error('Expected spreadsheet ArrayBuffer.');
          }
          cacheBits = { kind: 'binary', payload: raw.slice(0) };
          workbook = XLSX.read(new Uint8Array(raw), { type: 'array' });
        }
        const ok = applyImportedWorkbook(workbook, 0, file.name ?? '', input, { clearStoredWorkbookOnDup: true });
        if (ok && cacheBits) {
          void saveCachedSpreadsheet({
            fileName: file.name ?? '',
            kind: cacheBits.kind,
            payload: cacheBits.payload,
            sheetIndex: 0,
          });
        }
      } catch (err) {
        console.error(err);
        window.alert('Could not read this file. Try .xlsx / .xls / .csv, or re-export from Excel.');
      }
    };
    reader.onerror = () => {
      window.alert('Failed to read the file from disk.');
    };

    if (isCsv) reader.readAsText(file, 'UTF-8');
    else reader.readAsArrayBuffer(file);
  };

  const performClearImportDataErase = useCallback(() => {
    dismissClearImportDialog();

    /** 下一個 macrotask 再做同步 IO（saveSync）與大批 setState ，讓遮罩先有機會卸載繪製，避免「卡住像沒關」 */
    window.setTimeout(() => {
      if (typeof localStorage !== 'undefined') {
        try {
          localStorage.removeItem(IMPORT_STORE_KEY);
        } catch (err) {
          console.warn('route-planner: removeItem IMPORT_STORE_KEY failed', err);
          window.alert(
            'Could not delete data from this browser\'s storage (privacy / permissions on this PC).\n\n' +
              'Continuing to reset on screen and rewrite the Electron data file.',
          );
        }
      }

      try {
        const payloadJson = JSON.stringify(buildPersistPayload([], '', {}));
        lastPersistJsonRef.current = payloadJson;
        clearTimeout(persistDebouncedTimerRef.current);
        persistDebouncedTimerRef.current = null;
        try {
          window.routePlannerElectron?.saveSync?.(payloadJson);
        } catch {
          /* ignore */
        }
        void window.routePlannerElectron?.save?.(payloadJson)?.catch(() => {});

        setCustomers([]);
        setImportFileName('');
        importWorkbookRef.current = null;
        setImportSheetNames([]);
        setImportSheetIndex(0);
        setSchedules({});
        setAgent('');
        setSidebarVisitCountFilter('all');
        setCalendarChipRoundFilter('all');
        setSidebarArea('');
        setCalendarArea('');
        setSidebarChannelFilter('');
        setSidebarCategoryFilter('');
        setCalendarChannelFilter('');
        setCalendarCategoryFilter('');
        setSidebarNameQuery('');
        setCalendarNameQuery('');
        setSyncSidebarCalendarFilters(false);
        setSidebarShowUnassignedOnly(false);
        setUnassignedBrowseModalOpen(false);
        setSidebarSelectedRowKeys(new Set());
        setFocusedDay(null);
        void clearCachedSpreadsheet();
        const inp = document.getElementById('route-data-import');
        if (inp && typeof inp.value === 'string') inp.value = '';
      } catch (err) {
        console.error('route-planner: performClearImportDataErase failed', err);
        window.alert(
          'Clear did not finish completely. Close and reopen Route Planner.\n\n' +
            'If data is still there, check antivirus/file permissions or try again.',
        );
      }
    }, 0);
  }, [dismissClearImportDialog]);

  const submitClearImportUnlockStep = () => {
    const pwd = String(clearImportPassword).trim();
    const effective = getEffectiveClearImportPassword();
    if (pwd !== effective) {
      setClearImportError('Incorrect password.');
      return;
    }
    setClearImportError('');
    setClearImportPassword('');
    if (loadClearImportCustomPassword() == null) {
      setClearImportStep('setNewPassword');
    } else {
      setClearImportStep('confirmErase');
    }
  };

  const submitClearImportEraseWithNewPassword = () => {
    const n = String(clearImportNewPassword).trim();
    const c = String(clearImportNewPasswordConfirm).trim();
    if (n.length < 4) {
      setClearImportError('New password must be at least 4 characters.');
      return;
    }
    if (n !== c) {
      setClearImportError('New password and confirmation do not match.');
      return;
    }
    if (!clearImportAckDanger) {
      setClearImportError('Check the box to confirm removing all saved data.');
      return;
    }
    if (!saveClearImportCustomPassword(n)) return;
    performClearImportDataErase();
  };

  const submitClearImportEraseConfirmed = () => {
    if (!clearImportAckDanger) {
      setClearImportError('Check the box to confirm removing all saved data.');
      return;
    }
    performClearImportDataErase();
  };

  const onDrop = (dateStr, rawData) => {
    if (!agent) return;
    if (rawData == null || rawData === '') return;
    let parsed;
    try {
      parsed = JSON.parse(rawData);
    } catch {
      return;
    }

    /** @type {{ acc: string, name: string, oldDate?: string, visitRound?: number }[]} */
    let payloads = [];
    if (parsed && parsed.__batch === true && Array.isArray(parsed.items)) {
      payloads = parsed.items.filter((x) => x && x.acc != null && String(x.acc).length > 0);
    } else if (parsed && parsed.acc) {
      payloads = [parsed];
    } else {
      return;
    }
    if (payloads.length === 0) return;

    const placementDate = advanceIsoPastSundaysAndPublicHolidays(dateStr);
    if (placementDate !== dateStr) {
      restoreElectronWindowKeyboardFocus();
      window.alert(
        `That day is a Sunday or a public holiday highlighted on this calendar (${dateStr}). Visits were placed on ${placementDate} instead.`,
      );
    }

    const isCalendarMove = payloads.length === 1 && Boolean(payloads[0].oldDate);

    const isSidebarBatch = payloads.length > 1 && payloads.every((p) => !p.oldDate);

    setSchedules((prev) => {
      const prevAgent = prev[agent] || {};
      const agentSched = {};
      for (const k of Object.keys(prevAgent)) agentSched[k] = [...prevAgent[k]];

      const countAccFor = (acc) => {
        let n = 0;
        for (const arr of Object.values(agentSched)) {
          for (const t of arr) if (t.acc === acc) n += 1;
        }
        return n;
      };

      if (!isSidebarBatch) {
        const data = payloads[0];

        const countAcc = () => {
          let n = 0;
          for (const arr of Object.values(agentSched)) {
            for (const t of arr) if (t.acc === data.acc) n += 1;
          }
          return n;
        };

        if (data.oldDate) {
          const oldList = [...(agentSched[data.oldDate] || [])];
          const nextOld = oldList.filter((t) => t.acc !== data.acc);
          if (nextOld.length > 0) agentSched[data.oldDate] = nextOld;
          else delete agentSched[data.oldDate];
        }

        const daySched = [...(agentSched[placementDate] || [])];
        if (daySched.some((c) => c.acc === data.acc)) return prev;

        let visitRound;
        if (data.oldDate) {
          visitRound = normalizeVisitRound(data.visitRound ?? 1);
        } else {
          visitRound = countAcc() + 1;
        }

        agentSched[placementDate] = [...daySched, { acc: data.acc, name: data.name, visitRound }];
        return { ...prev, [agent]: agentSched };
      }

      let changed = false;
      for (const data of payloads) {
        if (data.oldDate) continue;
        const daySched = [...(agentSched[placementDate] || [])];
        if (daySched.some((c) => c.acc === data.acc)) continue;
        const visitRound = countAccFor(data.acc) + 1;
        agentSched[placementDate] = [...daySched, { acc: data.acc, name: data.name, visitRound }];
        changed = true;
      }
      if (!changed) return prev;
      return { ...prev, [agent]: agentSched };
    });

    if (!isCalendarMove) setSidebarSelectedRowKeys(new Set());
  };

  const removeScheduledVisit = (dateStr, acc) => {
    if (!agent) return;
    setSchedules((prev) => {
      const prevAgent = prev[agent];
      if (!prevAgent?.[dateStr]) return prev;
      const agentSched = {};
      for (const k of Object.keys(prevAgent)) agentSched[k] = [...prevAgent[k]];
      const nextDay = agentSched[dateStr].filter((t) => t.acc !== acc);
      if (nextDay.length === 0) delete agentSched[dateStr];
      else agentSched[dateStr] = nextDay;
      return { ...prev, [agent]: agentSched };
    });
  };

  /** Copy this agent's calendar visits in the viewed month onto the next month's same calendar dates (merge; skip dup acc same day). */
  const handleCopyMonthVisitsToNextMonth = useCallback(() => {
    if (!agent) {
      window.alert('Please select a sales agent first.');
      return;
    }
    const nextMonthIdx = month === 11 ? 0 : month + 1;
    const nextYearVal = month === 11 ? year + 1 : year;
    const nextMonthLabel = `${nextYearVal}-${MONTH_SHORT[nextMonthIdx]}`;
    const curMonthLabel = `${year}-${MONTH_SHORT[month]}`;
    const ok =
      typeof window !== 'undefined' &&
      window.confirm(
        `Copy ${curMonthLabel} calendar visits for the current agent to matching dates in ${nextMonthLabel}?\n\n` +
          'If a target day is a Sunday or a yellow public holiday on this calendar, visits are bumped forward day by day to the next open day.\n\n' +
          'If a customer is already scheduled on the target day, that visit will be skipped (no duplicates).',
      );
    if (!ok) return;

    let dup = /** @type {ReturnType<typeof duplicateAgentMonthVisitsOneMonthAhead>|null} */ (null);
    setSchedules((prev) => {
      dup = duplicateAgentMonthVisitsOneMonthAhead(prev, agent, year, month);
      return dup.changed ? dup.next : prev;
    });
    if (!dup) return;

    if (!dup.changed) {
      if (dup.added === 0 && dup.skippedDup === 0) {
        window.alert('No calendar visits in this month to copy.');
      } else if (dup.added === 0 && dup.skippedDup > 0) {
        window.alert(
          `Nothing to add — ${dup.skippedDup} visit(s) skipped because that customer is already on the matching day next month.`,
        );
      }
      restoreElectronWindowKeyboardFocus();
      return;
    }

    restoreElectronWindowKeyboardFocus();

    let doneMsg = `Copied ${dup.added} visit(s) to next month.`;
    if (dup.skippedDup > 0) doneMsg += ` (${dup.skippedDup} skipped — customer already on that day.)`;
    if (dup.sourceDatesShiftedFromBlockedDays > 0) {
      doneMsg += `\n\n${dup.sourceDatesShiftedFromBlockedDays} copied source day(s) landed on Sundays or holidays and were shifted to the next open day(s).`;
    }
    if (dup.daysSpilledPastTargetMonth > 0) {
      doneMsg += `\n\n${dup.daysSpilledPastTargetMonth} group(s) moved past the end of the target month—open Year/Month and check adjacent month.`;
    }
    window.alert(doneMsg);

    setYear(nextYearVal);
    setMonth(nextMonthIdx);
    setFocusedDay(null);
  }, [agent, month, year]);

  /** One-click fix for old data / mistakes: Sundays in toolbar month → following Monday */
  const handleMoveSundayVisitsToMondayCurrentMonth = useCallback(() => {
    if (!agent) {
      window.alert('Please select a sales agent first.');
      return;
    }
    const curMonthLabel = `${year}-${MONTH_SHORT[month]}`;
    const ok =
      typeof window !== 'undefined' &&
      window.confirm(
        `Move all visits scheduled on Sundays in ${curMonthLabel} for ${agent} to the following Monday?\n\n` +
          'If Monday already lists the same customer, that duplicate on Sunday will be cleared only.',
      );
    if (!ok) return;

    let out = /** @type {ReturnType<typeof moveSundayVisitsToMondayForAgentMonth>|null} */ (null);
    setSchedules((prev) => {
      out = moveSundayVisitsToMondayForAgentMonth(prev, agent, year, month);
      return out.changed ? out.next : prev;
    });
    if (!out) return;
    restoreElectronWindowKeyboardFocus();
    if (!out.changed) {
      window.alert('No visits on Sundays in this month for this agent.');
      return;
    }
    let msg = `Moved ${out.movedVisits} visit(s) off ${out.sundayDatesTouched} Sunday date(s).`;
    if (out.skippedDup > 0) msg += ` (${out.skippedDup} were duplicate accounts already on Monday and were removed.)`;
    window.alert(msg);
  }, [agent, year, month]);

  const renderCalendar = () => (
    <div className="cal-print-sheet">
      <div className="cal-calendar-body" style={{ width: '100%', minWidth: 0, flex: '1 1 auto', minHeight: 0 }}>
        <div className="cal-calendar-print-header">
          <span>
            {companyName.trim() ? `${companyName.trim()} · ` : ''}
            {agent || 'NOT SELECTED'} · {monthLabel}
          </span>
        </div>
        <div
          className="cal-weekday-header"
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(7, minmax(0, 1fr))',
            width: 'max(100%, 800px)',
            boxSizing: 'border-box',
            background: '#f1f5f9',
            border: '1px solid #e2e8f0',
            borderBottom: 'none',
          }}
        >
          {WEEKDAY_ABBREV_SUN0.map((label) => (
            <div
              key={label}
              style={{
                padding: '6px 4px',
                textAlign: 'center',
                fontSize: '11px',
                fontWeight: 700,
                letterSpacing: '0.04em',
                textTransform: 'uppercase',
                color: label === 'Sun' ? '#e11d48' : '#64748b',
                borderRight: '1px solid #e2e8f0',
                boxSizing: 'border-box',
              }}
            >
              {label}
            </div>
          ))}
        </div>
        <div
          className="cal-weeks-stack"
          style={{
            display: 'flex',
            flexDirection: 'column',
            width: 'max(100%, 800px)',
            boxSizing: 'border-box',
            ['--cal-week-rows']: calendarWeeks.length,
          }}
        >
          {calendarWeeks.map((weekCells, wi) => (
            <CalWeekRow
              key={`week-${wi}-${year}-${month}-p${printEpoch}`}
              weekCells={weekCells}
              monthIsoPad={monthIsoPad}
              year={year}
              agent={agent}
              schedules={schedules}
              highlightedDayNum={highlightedDay}
              dropOnDate={(dateStr, raw) => onDrop(dateStr, raw)}
              removeScheduledVisit={removeScheduledVisit}
              calendarFilterAcc={calendarFilterAcc}
              calendarChipRoundFilter={calendarChipRoundFilter}
              onDayDoubleClick={(dateStr) => {
                if (!agent) return;
                setCalendarDayDetailDateStr(dateStr);
              }}
            />
          ))}
        </div>
      </div>
    </div>
  );

  return (
    <div className="app-root-shell" style={{ display: 'flex', width: '100%', height: '100dvh', minHeight: '100dvh', maxHeight: '100dvh', overflow: 'hidden', background: '#0f172a', boxSizing: 'border-box' }}>
      <style>{`
        @media print {
          .no-print { display: none !important; }
          * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          :root {
            color-scheme: light !important;
          }
          @page {
            margin: 5mm;
            size: A4 portrait;
          }
          @page cal-sheet {
            margin: 5mm;
            size: A4 portrait;
          }
          html, body, #root {
            margin: 0 !important;
            padding: 0 !important;
            width: 100% !important;
            height: 100% !important;
            background: white !important;
            overflow: hidden !important;
          }
          body {
            color: #0f172a !important;
          }
          #root {
            display: block !important;
            min-height: 0 !important;
            flex: none !important;
          }
          /* Summary prints as a flowing multi-page document: undo the
             one-page clip (that clip exists for the fixed calendar sheet)
             on the whole ancestor chain so page breaks can happen. */
          html:has(body.print-summary-view),
          body.print-summary-view,
          body.print-summary-view #root {
            height: auto !important;
            max-height: none !important;
            overflow: visible !important;
          }
          body.print-summary-view .app-root-shell {
            display: block !important;
            height: auto !important;
            max-height: none !important;
            min-height: 0 !important;
            overflow: visible !important;
          }
          body.print-summary-view .main-column {
            display: block !important;
            height: auto !important;
            min-height: 0 !important;
            overflow: visible !important;
          }
          body.print-summary-view .main-scroll-region {
            display: block !important;
            height: auto !important;
            min-height: 0 !important;
            overflow: visible !important;
            flex: none !important;
            padding: 0 !important;
          }
          body.print-summary-view .content-white-panel {
            display: block !important;
            height: auto !important;
            min-height: 0 !important;
            overflow: visible !important;
            border: none !important;
          }
          .cal-print-sheet {
            page: cal-sheet;
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            right: 0 !important;
            bottom: 0 !important;
            width: 100% !important;
            height: 100% !important;
            display: flex !important;
            flex-direction: column !important;
            box-sizing: border-box !important;
            overflow: hidden !important;
            background: white !important;
            page-break-after: avoid !important;
            page-break-inside: avoid !important;
            z-index: 99999 !important;
          }
          .cal-print-sheet .cal-calendar-body {
            flex: 1 1 auto !important;
            min-height: 0 !important;
            height: 100% !important;
            width: 100% !important;
            display: flex !important;
            flex-direction: column !important;
            box-sizing: border-box !important;
          }
          .cal-calendar-print-header {
            display: block !important;
            flex-shrink: 0 !important;
            margin: 0 0 1.5mm 0 !important;
            padding: 0 1mm !important;
            font-size: 9pt !important;
            font-weight: 700 !important;
            color: #0f172a !important;
            line-height: 1.2 !important;
          }
          .cal-print-sheet .cal-weekday-header {
            flex-shrink: 0 !important;
            display: grid !important;
            grid-template-columns: repeat(7, 1fr) !important;
            width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            border: 1px solid #cbd5e1 !important;
            border-bottom: none !important;
            background: #f1f5f9 !important;
          }
          .cal-print-sheet .cal-weekday-header > div {
            padding: 1.2mm 0 !important;
            text-align: center !important;
            font-size: 7.5pt !important;
            font-weight: 700 !important;
            border-right: 1px solid #e2e8f0 !important;
            box-sizing: border-box !important;
          }
          .cal-print-sheet .cal-weeks-stack {
            flex: 1 1 auto !important;
            min-height: 0 !important;
            display: grid !important;
            grid-template-rows: repeat(var(--cal-week-rows, 6), minmax(0, 1fr)) !important;
            width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            gap: 0 !important;
            border: 1px solid #cbd5e1 !important;
            border-top: none !important;
          }
          .cal-week-block {
            width: 100% !important;
            min-height: 0 !important;
            height: 100% !important;
            box-sizing: border-box !important;
            break-inside: avoid !important;
            page-break-inside: avoid !important;
          }
          .app-root-shell {
            display: block !important;
            width: 100% !important;
            background: white !important;
          }
          .main-column {
            display: block !important;
            background: white !important;
          }
          .main-scroll-region {
            padding: 0 !important;
            flex: none !important;
          }
          .content-white-panel {
            border: none !important;
            flex: none !important;
          }
          .schedule-strip-print {
            display: none !important;
          }
          /* Summary prints in normal flow: one page per two-week block */
          body.print-summary-view .summary-print-wrap:not(.summary-print-empty) {
            position: static !important;
            display: block !important;
            width: 100% !important;
            height: auto !important;
            overflow: visible !important;
            box-sizing: border-box !important;
            background: white !important;
          }
          body.print-summary-view .summary-print-page {
            break-after: page;
            page-break-after: always;
            break-inside: avoid;
            page-break-inside: avoid;
            transform-origin: top left;
          }
          body.print-summary-view .summary-print-page:last-child {
            break-after: auto;
            page-break-after: auto;
          }
          .cal-week-row {
            width: 100% !important;
            height: 100% !important;
            min-height: 0 !important;
            display: grid !important;
            grid-template-columns: repeat(7, 1fr) !important;
            align-items: stretch !important;
          }
          .cal-day-cell,
          .cal-pad-cell {
            min-height: 0 !important;
            height: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
          }
          .cal-day-cell {
            padding: 0.4mm !important;
            display: flex !important;
            flex-direction: row !important;
            align-items: stretch !important;
            gap: 0.2mm !important;
          }
          .cal-day-cell > b {
            flex: 0 0 auto !important;
            width: 3mm !important;
            min-width: 3mm !important;
            font-size: 8pt !important;
            line-height: 1.1 !important;
            text-align: left !important;
            padding-top: 0.2mm !important;
          }
          .cal-day-tasks {
            flex: 1 1 auto !important;
            min-width: 0 !important;
            min-height: 0 !important;
            height: 100% !important;
            margin-top: 0 !important;
            gap: var(--cal-print-gap, 0.4px) !important;
            display: flex !important;
            flex-direction: column !important;
            overflow: hidden !important;
          }
          .cal-day-tasks > div,
          .cal-day-tasks .cal-day-task-chip {
            flex: 0 0 auto !important;
            width: 100% !important;
            min-height: 0 !important;
            height: auto !important;
            max-height: none !important;
            padding: 0 !important;
            border-radius: 0 !important;
            gap: 0.3mm !important;
            background: transparent !important;
            border: none !important;
            box-shadow: none !important;
            color: var(--visit-round-color, #1e293b) !important;
            display: flex !important;
            align-items: center !important;
          }
          .cal-day-tasks .cal-day-task-chip span[aria-hidden] {
            flex-shrink: 0 !important;
            min-width: 0 !important;
            min-height: 0 !important;
            height: auto !important;
            padding: 0 0.3mm 0 0 !important;
            background: transparent !important;
            font-weight: 800 !important;
          }
          .cal-day-tasks [data-cal-chip-text] {
            flex: 1 1 auto !important;
            min-width: 0 !important;
            font-size: var(--cal-print-font-size, 7pt) !important;
            line-height: var(--cal-print-line-height, 1.08) !important;
            color: var(--visit-round-color, #1e293b) !important;
            font-weight: 700 !important;
            white-space: normal !important;
            overflow-wrap: break-word !important;
            overflow: hidden !important;
            text-overflow: clip !important;
          }
          .cal-day-tasks span[aria-hidden] {
            font-size: var(--cal-print-num-font-size, 6.5pt) !important;
            line-height: var(--cal-print-line-height, 1.05) !important;
            color: var(--visit-round-color, #1e293b) !important;
          }
          .summary-print-wrap {
            overflow: visible !important;
            max-height: none !important;
            padding: 0 !important;
            display: block !important;
            flex: none !important;
            gap: 0 !important;
            break-inside: avoid;
            page-break-inside: avoid;
            transform-origin: top left;
          }
          .summary-print-page-title {
            display: block !important;
            margin: 0 0 6px 0 !important;
            padding: 0 !important;
          }
          .summary-print-page-title .app-company-display {
            display: block !important;
            font-size: 11pt !important;
            margin-bottom: 2px !important;
            color: #0f172a !important;
            -webkit-text-fill-color: #0f172a !important;
            background: none !important;
          }
          .summary-print-page-title-meta {
            display: block !important;
            font-size: 8pt !important;
            font-weight: 600 !important;
            color: #475569 !important;
            letter-spacing: 0.04em !important;
            text-transform: uppercase !important;
          }
          .summary-print-kpi-row {
            display: flex !important;
            flex-wrap: nowrap !important;
            gap: 8px !important;
            margin-bottom: 6px !important;
          }
          .summary-print-kpi {
            flex: 1 1 auto !important;
            min-width: 0 !important;
            padding: 4px 8px !important;
            border-radius: 4px !important;
          }
          .summary-print-kpi-label {
            font-size: 6pt !important;
            margin-bottom: 1px !important;
          }
          .summary-print-kpi-value {
            font-size: 10pt !important;
          }
          .summary-print-table-host {
            overflow: visible !important;
            border-radius: 0 !important;
          }
          .summary-print-table {
            font-size: 7pt !important;
          }
          .summary-print-table th,
          .summary-print-table td {
            padding: 2px 4px !important;
          }
          .summary-print-table thead {
            display: table-header-group;
          }
          .summary-print-table tr {
            break-inside: avoid;
            page-break-inside: avoid;
          }
        }
        /* Apply print calendar layout before the dialog so auto-fit measures real cell size */
        body.print-calendar-view .no-print {
          display: none !important;
        }
        body.print-calendar-view .cal-print-sheet {
          position: fixed !important;
          top: 0 !important;
          left: 0 !important;
          right: 0 !important;
          bottom: 0 !important;
          width: 100% !important;
          height: 100% !important;
          display: flex !important;
          flex-direction: column !important;
          box-sizing: border-box !important;
          overflow: hidden !important;
          background: white !important;
          z-index: 99999 !important;
        }
        body.print-calendar-view .cal-print-sheet .cal-calendar-body {
          flex: 1 1 auto !important;
          min-height: 0 !important;
          height: 100% !important;
          display: flex !important;
          flex-direction: column !important;
        }
        body.print-calendar-view .cal-print-sheet .cal-weeks-stack {
          flex: 1 1 auto !important;
          min-height: 0 !important;
          display: grid !important;
          grid-template-rows: repeat(var(--cal-week-rows, 6), minmax(0, 1fr)) !important;
        }
        body.print-calendar-view .cal-week-block {
          height: 100% !important;
          min-height: 0 !important;
        }
        body.print-calendar-view .cal-week-row {
          display: grid !important;
          grid-template-columns: repeat(7, minmax(0, 1fr)) !important;
          height: 100% !important;
          align-items: stretch !important;
        }
        body.print-calendar-view .cal-day-cell,
        body.print-calendar-view .cal-pad-cell {
          height: 100% !important;
          min-height: 0 !important;
          overflow: hidden !important;
          box-sizing: border-box !important;
          padding: 0.4mm !important;
          display: flex !important;
          flex-direction: row !important;
          align-items: stretch !important;
          gap: 0.2mm !important;
        }
        body.print-calendar-view .cal-day-cell > b {
          flex: 0 0 auto !important;
          width: 3mm !important;
          min-width: 3mm !important;
          font-size: 8pt !important;
          line-height: 1.1 !important;
        }
        body.print-calendar-view .cal-day-tasks {
          flex: 1 1 auto !important;
          min-width: 0 !important;
          min-height: 0 !important;
          height: 100% !important;
          margin-top: 0 !important;
          display: flex !important;
          flex-direction: column !important;
          overflow: hidden !important;
        }
        body.print-calendar-view .cal-day-task-chip {
          flex: 0 0 auto !important;
          width: 100% !important;
          min-height: 0 !important;
          height: auto !important;
          max-height: none !important;
          background: transparent !important;
          border: none !important;
          border-radius: 0 !important;
          padding: 0 !important;
          box-shadow: none !important;
        }
        body.print-calendar-view .cal-day-tasks [data-cal-chip-text] {
          flex: 1 1 auto !important;
          min-width: 0 !important;
          font-size: var(--cal-print-font-size, 7pt) !important;
          line-height: var(--cal-print-line-height, 1.08) !important;
          color: var(--visit-round-color, #1e293b) !important;
          font-weight: 700 !important;
          white-space: normal !important;
          overflow-wrap: break-word !important;
          overflow: hidden !important;
          text-overflow: clip !important;
        }
        body.print-calendar-view .cal-day-tasks span[aria-hidden] {
          font-size: var(--cal-print-num-font-size, 6.5pt) !important;
          line-height: var(--cal-print-line-height, 1.08) !important;
          color: var(--visit-round-color, #1e293b) !important;
          background: transparent !important;
          font-weight: 800 !important;
        }
        .cal-calendar-print-header {
          display: none;
        }
        .import-sr-only {
          position: absolute;
          width: 1px;
          height: 1px;
          padding: 0;
          margin: -1px;
          overflow: hidden;
          clip: rect(0, 0, 0, 0);
          white-space: nowrap;
          border: 0;
        }
        .sidebar-select:hover {
          border-color: rgba(96, 165, 250, 0.5);
          box-shadow: 0 1px 3px rgba(0,0,0,0.35), 0 0 0 1px rgba(59, 130, 246, 0.2);
        }
        .sidebar-select:disabled {
          opacity: 0.5;
          cursor: not-allowed;
          filter: grayscale(0.2);
        }
        .sidebar-select option { background-color: #1e293b; color: #f8fafc; }
        .file-import-btn {
          transition: filter 0.15s ease, transform 0.12s ease, box-shadow 0.15s ease;
        }
        .file-import-btn:hover {
          filter: brightness(1.08);
          transform: translateY(-1px);
          box-shadow: 0 4px 14px rgba(37, 99, 235, 0.45);
        }
        .file-import-btn:active {
          transform: translateY(0);
          filter: brightness(0.97);
        }
        .sidebar-scroll::-webkit-scrollbar { width: 4px; }
        .sidebar-scroll::-webkit-scrollbar-thumb {
          background: rgba(148, 163, 184, 0.35);
          border-radius: 8px;
        }
        .sidebar-scroll::-webkit-scrollbar-track { background: rgba(15, 23, 42, 0.5); }
        .sidebar-scroll { scrollbar-width: thin; scrollbar-color: rgba(148, 163, 184, 0.45) rgba(15, 23, 42, 0.5); }
        .cal-day-tasks::-webkit-scrollbar { width: 6px; }
        .cal-day-tasks::-webkit-scrollbar-thumb {
          background: rgba(100, 116, 139, 0.55);
          border-radius: 6px;
        }
        .cal-day-tasks::-webkit-scrollbar-track { background: rgba(241, 245, 249, 0.9); }
        .cal-day-tasks {
          scrollbar-width: thin;
          scrollbar-color: rgba(100, 116, 139, 0.55) rgba(241, 245, 249, 0.9);
          align-items: stretch;
        }
        .cal-day-cell {
          min-width: 0;
        }
        .cal-day-task-chip {
          width: 100%;
          max-width: 100%;
          align-self: stretch;
        }
        .content-white-panel::-webkit-scrollbar {
          width: 8px;
          height: 8px;
        }
        .content-white-panel::-webkit-scrollbar-thumb {
          background: rgba(148, 163, 184, 0.55);
          border-radius: 8px;
        }
        .content-white-panel::-webkit-scrollbar-track {
          background: rgba(241, 245, 249, 0.9);
        }
        .content-white-panel {
          scrollbar-width: thin;
          scrollbar-color: rgba(100, 116, 139, 0.5) rgba(241, 245, 249, 0.95);
        }
        .app-company-display {
          font-family: 'DM Serif Display', Georgia, 'Times New Roman', serif;
          font-weight: 400;
          font-style: normal;
          letter-spacing: 0.012em;
          font-feature-settings: 'kern' 1, 'liga' 1;
          text-rendering: optimizeLegibility;
          -webkit-font-smoothing: antialiased;
          -moz-osx-font-smoothing: grayscale;
        }
        .main-toolbar-company-wrap {
          display: flex;
          flex-direction: column;
          justify-content: center;
          gap: 3px;
          padding: 4px 0 4px 16px;
          margin-left: 6px;
          border-left: 2px solid transparent;
          border-image: linear-gradient(180deg, #bae6fd 0%, #2563eb 45%, #7c3aed 100%) 1;
          min-width: 0;
          max-width: min(440px, 44vw);
        }
        .main-toolbar-company-label {
          font-family: system-ui, -apple-system, 'Segoe UI', sans-serif;
          font-size: 8px;
          font-weight: 800;
          letter-spacing: 0.24em;
          text-transform: uppercase;
          line-height: 1;
          margin: 0;
          background: linear-gradient(90deg, #64748b 0%, #2563eb 55%, #7c3aed 100%);
          -webkit-background-clip: text;
          background-clip: text;
          -webkit-text-fill-color: transparent;
          color: transparent;
        }
        .main-toolbar-company-name {
          margin: 0;
          font-size: clamp(1.32rem, 2.6vw, 1.92rem);
          line-height: 1.08;
          font-weight: 400;
          background: linear-gradient(
            165deg,
            #38bdf8 0%,
            #2563eb 22%,
            #6366f1 48%,
            #7c3aed 68%,
            #1d4ed8 88%,
            #0ea5e9 100%
          );
          -webkit-background-clip: text;
          background-clip: text;
          -webkit-text-fill-color: transparent;
          color: transparent;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
          /* Light 3D — fewer, softer layers */
          filter:
            drop-shadow(0 1px 0 rgba(255, 255, 255, 0.75))
            drop-shadow(0 2px 0 rgba(147, 197, 253, 0.45))
            drop-shadow(0 3px 0 rgba(59, 130, 246, 0.35))
            drop-shadow(0 4px 6px rgba(15, 23, 42, 0.12));
        }
        .app-company-display--sidebar {
          display: block;
          font-size: 0.82rem;
          line-height: 1.2;
          text-align: center;
          margin: 0 0 2px;
          padding: 0 4px;
          background: linear-gradient(180deg, #f8fafc 0%, #cbd5e1 100%);
          -webkit-background-clip: text;
          background-clip: text;
          -webkit-text-fill-color: transparent;
          color: transparent;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }
        .app-company-display--on-dark {
          font-size: 1.05rem;
          line-height: 1.15;
          background: linear-gradient(105deg, #ffffff 0%, #e2e8f0 55%, #94a3b8 100%);
          -webkit-background-clip: text;
          background-clip: text;
          -webkit-text-fill-color: transparent;
          color: transparent;
        }
        .schedule-strip-agent-row {
          display: inline-flex;
          align-items: baseline;
          flex-wrap: wrap;
          gap: 0 10px;
          min-width: 0;
        }
        .schedule-strip-agent-meta {
          font-family: system-ui, -apple-system, 'Segoe UI', sans-serif;
          font-size: 13px;
          font-weight: 700;
          letter-spacing: 0.04em;
          color: rgba(255, 255, 255, 0.92);
        }
      `}</style>
      
      {/* Sidebar: top controls fixed, customer list fills remaining viewport and scrolls */}
      <aside className="no-print" style={{ width: '238px', minWidth: '238px', minHeight: 0, alignSelf: 'stretch', background: 'linear-gradient(180deg, #0f172a 0%, #0c1222 100%)', color: 'white', padding: '5px 5px', display: 'flex', flexDirection: 'column', boxSizing: 'border-box', borderRight: '1px solid rgba(148, 163, 184, 0.12)', overflow: 'hidden' }}>
        <div style={{ flexShrink: 0 }}>
        <h2 style={{ color: '#60a5fa', textAlign: 'center', margin: '0', fontSize: '0.75rem', fontWeight: 800, letterSpacing: '0.05em', lineHeight: 1.15, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '3px' }}>
          <span
            aria-hidden
            onClick={(e) => {
              e.stopPropagation();
              const egg = logoPinEasterEggRef.current;
              if (egg.idleResetId != null) window.clearTimeout(egg.idleResetId);
              egg.count += 1;
              egg.idleResetId = window.setTimeout(() => {
                egg.count = 0;
                egg.idleResetId = null;
              }, 4200);
              if (egg.count < 9) return;
              egg.count = 0;
              if (egg.idleResetId != null) window.clearTimeout(egg.idleResetId);
              egg.idleResetId = null;
              const hadCustom = loadClearImportCustomPassword() != null;
              if (!resetClearImportPasswordToFactoryDefault()) {
                window.alert('Could not reset password (storage blocked).');
                restoreElectronWindowKeyboardFocus();
                return;
              }
              window.alert(
                hadCustom
                  ? 'Hidden reset OK: Clear saved import now uses factory password 1234 again.'
                  : 'No custom clear password was saved; unlock already uses factory 1234.',
              );
              restoreElectronWindowKeyboardFocus();
            }}
            title=""
            style={{ cursor: 'inherit', userSelect: 'none', display: 'inline-flex', alignItems: 'center', marginRight: '1px' }}
          >
            <img
              src={routeProLogo}
              alt=""
              draggable={false}
              style={{
                width: '14px',
                height: '14px',
                objectFit: 'cover',
                objectPosition: 'center 18%',
                borderRadius: '50%',
                display: 'block',
              }}
            />
          </span>
          ROUTE PR
          <span
            aria-hidden
            onClick={handleLogoOSettingsEasterEggClick}
            title=""
            style={{ cursor: 'inherit', userSelect: 'none' }}
          >
            O
          </span>
        </h2>
        {companyName.trim() ? (
          <p className="app-company-display app-company-display--sidebar" title={companyName.trim()}>
            {companyName.trim()}
          </p>
        ) : null}
        <p style={{ textAlign: 'center', fontSize: '7px', color: 'rgba(148, 163, 184, 0.85)', margin: companyName.trim() ? '0 0 3px' : '1px 0 3px', lineHeight: 1.15 }}>
          {`Import · schedule visits · v${packageJson.version}`}
        </p>

        <div style={{ marginBottom: '4px' }}>
          <label style={sidebarSectionLabel}>1. Data import</label>
          <div style={{ position: 'relative' }}>
            <input
              id="route-data-import"
              className="import-sr-only"
              type="file"
              accept=".xlsx,.xls,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
              onChange={handleFile}
            />
            <label
              htmlFor="route-data-import"
              className="file-import-btn"
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '4px',
                width: '100%',
                padding: '4px 5px',
                borderRadius: '5px',
                cursor: 'pointer',
                boxSizing: 'border-box',
                fontSize: '9px',
                fontWeight: 600,
                color: 'white',
                border: '1px solid rgba(255,255,255,0.14)',
                background: 'linear-gradient(180deg, #3b82f6 0%, #2563eb 45%, #1d4ed8 100%)',
                boxShadow: '0 2px 10px rgba(37, 99, 235, 0.4)',
              }}
            >
              <Upload size={13} strokeWidth={2.25} aria-hidden />
              Choose spreadsheet
            </label>
          </div>
          <div
            style={{
              marginTop: '2px',
              fontSize: '7px',
              color: importFileName ? 'rgba(226, 232, 240, 0.95)' : 'rgba(148, 163, 184, 0.75)',
              lineHeight: 1.25,
              wordBreak: 'break-all',
              minHeight: '10px',
            }}
            title={importFileName || undefined}
          >
            {importFileName ? importFileName : 'No file chosen · .xlsx / .xls / .csv'}
          </div>
          <div
            style={{
              display: 'flex',
              gap: '4px',
              marginTop: '3px',
              alignItems: 'stretch',
              minWidth: 0,
              width: '100%',
              boxSizing: 'border-box',
            }}
          >
            {importSheetNames.length > 1 ? (
              <select
                id="import-sheet-select"
                className="sidebar-select"
                value={importSheetIndex}
                onChange={(e) => {
                  const next = Number(e.target.value);
                  const wb = importWorkbookRef.current;
                  if (wb == null || !Number.isFinite(next) || next === importSheetIndex) return;
                  applyImportedWorkbook(wb, next, importFileName, null);
                }}
                aria-label="Worksheet for import"
                title="Choose worksheet (.xlsx / .xls)"
                style={{
                  ...selectSidebarStyle,
                  flex: '1 1 50%',
                  minWidth: 0,
                  marginBottom: 0,
                  padding: '3px 14px 3px 4px',
                  fontSize: '8px',
                  fontWeight: 600,
                  borderRadius: '4px',
                  backgroundPosition: '0 0, right 4px center',
                  backgroundSize: '100% 100%, 7px 7px',
                  cursor: 'pointer',
                }}
              >
                {importSheetNames.map((nm, i) => (
                  <option key={`${nm}\u001f${i}`} value={i}>
                    {nm && String(nm).trim() !== '' ? nm : `Sheet ${i + 1}`}
                  </option>
                ))}
              </select>
            ) : null}
            <button
              type="button"
              className="no-print"
              onClick={() => {
                setClearImportStep('unlock');
                setClearImportPassword('');
                setClearImportNewPassword('');
                setClearImportNewPasswordConfirm('');
                setClearImportAckDanger(false);
                setClearImportError('');
                setClearImportDialogOpen(true);
              }}
              style={{
                flex: importSheetNames.length > 1 ? '1 1 50%' : '1 1 100%',
                minWidth: 0,
                padding: '3px 5px',
                borderRadius: '4px',
                border: '1px solid rgba(248, 113, 113, 0.4)',
                background: 'rgba(127, 29, 29, 0.38)',
                color: '#fecaca',
                fontSize: '8px',
                fontWeight: 600,
                cursor: 'pointer',
              }}
            >
              Clear saved import…
            </button>
          </div>
        </div>

        <div style={{ marginBottom: '3px' }}>
          <label style={sidebarSectionLabel} htmlFor="sidebar-agent">2. Sales agent</label>
          <select
            id="sidebar-agent"
            className="sidebar-select"
            value={agent}
            onChange={(e) => {
              setAgent(e.target.value);
              setSidebarVisitCountFilter('all');
              setCalendarChipRoundFilter('all');
            }}
            style={{ ...selectSidebarStyle, marginBottom: 0 }}
          >
            <option value="">— Select agent —</option>
            {salesAgentOptions.map((a) => (
              <option key={a} value={a}>
                {a}
              </option>
            ))}
          </select>
        </div>

        <div style={{ display: 'flex', gap: '3px', marginBottom: '4px', alignItems: 'stretch', minWidth: 0 }}>
          <div style={{ flex: '1 1 50%', minWidth: 0, display: 'flex', flexDirection: 'column' }}>
            <label style={sidebarSectionLabel} htmlFor="sidebar-area">
              3. Area <span style={{ fontWeight: 600, opacity: 0.85, textTransform: 'none', letterSpacing: 'normal' }}>(sidebar)</span>
            </label>
            <select
              id="sidebar-area"
              className="sidebar-select"
              value={sidebarArea}
              onChange={(e) => setSidebarArea(e.target.value)}
              style={{ ...selectSidebarStyle, marginBottom: 0 }}
              disabled={!agent}
            >
              <option value="">
                {agent ? '— All areas —' : '— Agent first —'}
              </option>
              {sidebarAreaOptionsForSelect.map((a) => (
                <option key={a} value={a}>
                  {a}
                </option>
              ))}
            </select>
          </div>
          <div style={{ flex: '1 1 50%', minWidth: 0, display: 'flex', flexDirection: 'column' }}>
            <label style={sidebarSectionLabel} htmlFor="sidebar-visit-filter">
              4. Visits{' '}
              <span style={{ fontWeight: 600, opacity: 0.85, textTransform: 'none', letterSpacing: 'normal' }}>(sidebar)</span>
            </label>
            <select
              id="sidebar-visit-filter"
              className="sidebar-select"
              value={sidebarVisitCountFilter}
              onChange={(e) => setSidebarVisitCountFilter(e.target.value)}
              title={`Sidebar list only · counts use calendar month ${monthLabel}.`}
              style={{ ...selectSidebarStyle, marginBottom: 0 }}
              disabled={!agent}
            >
              <option value="all">{agent ? '— Any count —' : '— Agent first —'}</option>
              <option value="0">Not scheduled yet (0×)</option>
              <option value="1">Exactly 1 visit</option>
              <option value="2">Exactly 2 visits</option>
              <option value="3">Exactly 3 visits</option>
              <option value="4">Exactly 4 visits</option>
              <option value="5">Exactly 5 visits</option>
              <option value="6+">6 or more visits</option>
            </select>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '3px', marginBottom: '4px', alignItems: 'stretch', minWidth: 0 }}>
          <div style={{ flex: '1 1 50%', minWidth: 0, display: 'flex', flexDirection: 'column' }}>
            <label style={sidebarSectionLabel} htmlFor="sidebar-channel">
              5. Channel <span style={{ fontWeight: 600, opacity: 0.85, textTransform: 'none', letterSpacing: 'normal' }}>(sidebar)</span>
            </label>
            <select
              id="sidebar-channel"
              className="sidebar-select"
              value={sidebarChannelFilter}
              onChange={(e) => setSidebarChannelFilter(e.target.value)}
              style={{ ...selectSidebarStyle, marginBottom: 0 }}
              disabled={!agent}
            >
              <option value="">
                {agent ? '— All channels —' : '— Agent first —'}
              </option>
              {sidebarChannelOptions.map((ch) => (
                <option key={ch} value={ch}>{ch}</option>
              ))}
            </select>
          </div>
          <div style={{ flex: '1 1 50%', minWidth: 0, display: 'flex', flexDirection: 'column' }}>
            <label style={sidebarSectionLabel} htmlFor="sidebar-category">
              6. Category <span style={{ fontWeight: 600, opacity: 0.85, textTransform: 'none', letterSpacing: 'normal' }}>(sidebar)</span>
            </label>
            <select
              id="sidebar-category"
              className="sidebar-select"
              value={sidebarCategoryFilter}
              onChange={(e) => setSidebarCategoryFilter(e.target.value)}
              style={{ ...selectSidebarStyle, marginBottom: 0 }}
              disabled={!agent}
            >
              <option value="">
                {agent ? '— All categories —' : '— Agent first —'}
              </option>
              {sidebarCategoryOptions.map((cat) => (
                <option key={cat} value={cat}>{cat}</option>
              ))}
            </select>
          </div>
        </div>

        <div style={{ marginBottom: '4px', minWidth: 0 }}>
          <label style={sidebarSectionLabel} htmlFor="sidebar-name-search">
            7. Find customer <span style={{ fontWeight: 600, opacity: 0.85, textTransform: 'none', letterSpacing: 'normal' }}>(sidebar list)</span>
          </label>
          <input
            id="sidebar-name-search"
            type="search"
            autoCapitalize="off"
            spellCheck={false}
            aria-label="Filter sidebar customer list by name or account"
            value={sidebarNameQuery}
            placeholder={agent ? 'Search by name or account #…' : '— Select sales agent —'}
            disabled={!agent}
            onChange={(e) => setSidebarNameQuery(e.target.value)}
            style={{
              width: '100%',
              padding: '3px 5px',
              borderRadius: '4px',
              border: `1px solid ${agent ? 'rgba(148, 163, 184, 0.35)' : 'rgba(100,116,139,0.35)'}`,
              background: agent ? '#1e293b' : '#0f172a',
              color: agent ? '#f8fafc' : 'rgba(148,163,184,0.45)',
              fontSize: '9px',
              fontWeight: 500,
              outline: 'none',
              boxSizing: 'border-box',
              opacity: agent ? 1 : 0.65,
              boxShadow: agent ? '0 1px 3px rgba(0,0,0,0.35)' : 'none',
            }}
          />
          {agent && sidebarNameQuery.trim().length > 0 && (
            <div style={{ marginTop: '2px', fontSize: '7px', color: 'rgba(148, 163, 184, 0.85)', lineHeight: 1.3 }}>
              {sidebarDisplayedCustomers.length} match{sidebarDisplayedCustomers.length === 1 ? '' : 'es'} · Visit count filter ignored while typing (Area / Channel / Category still apply)
            </div>
          )}
        </div>

        <div
          title={
            !agent
              ? undefined
              : unassignedStripFiltersActive
                ? 'Unassigned numbers use sidebar Area / Channel / Category filters (same as list when Show is on).'
                : `Numbers include every outlet for ${agent}. Choose Area / Channel / Category above to narrow.`
          }
          style={{
            marginBottom: '4px',
            padding: '4px 5px',
            borderRadius: '5px',
            border: '1px solid rgba(239, 68, 68, 0.55)',
            background: 'linear-gradient(145deg, rgba(127, 29, 29, 0.42) 0%, rgba(30, 27, 46, 0.5) 100%)',
            display: 'flex',
            alignItems: 'flex-start',
            justifyContent: 'space-between',
            gap: '4px',
            minWidth: 0,
            boxSizing: 'border-box',
            boxShadow: 'inset 0 1px 0 rgba(255,255,255,0.06)',
          }}
        >
          <div
            style={{
              display: 'flex',
              flexDirection: 'column',
              gap: '2px',
              minWidth: 0,
              flex: 1,
              overflow: 'hidden',
            }}
          >
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '4px',
                lineHeight: 1,
                flexWrap: 'wrap',
              }}
            >
              <span
                style={{
                  fontSize: '7px',
                  fontWeight: 700,
                  color: 'rgba(254, 202, 202, 0.98)',
                  textTransform: 'uppercase',
                  letterSpacing: '0.06em',
                  flexShrink: 0,
                }}
              >
                Unassigned
              </span>
              <span
                style={{
                  fontSize: '12px',
                  fontWeight: 900,
                  color: '#ff1744',
                  textShadow: '0 0 12px rgba(255, 23, 68, 0.5)',
                  fontVariantNumeric: 'tabular-nums',
                  flexShrink: 0,
                }}
              >
                {unassignedCustomerCount}
              </span>
              <span style={{ fontSize: '8px', fontWeight: 800, color: '#ff5252', letterSpacing: '0.02em', flexShrink: 0 }}>
                not on calendar
              </span>
            </div>
            {agent && agentCustomerTotalCount > 0 && (
              <div
                style={{
                  fontSize: '6px',
                  fontWeight: 600,
                  color: 'rgba(254, 202, 202, 0.72)',
                  letterSpacing: '0.02em',
                  lineHeight: 1.22,
                }}
              >
                {unassignedStripFiltersActive ? (
                  <>
                    {scheduledOnCalendarCount} on calendar · {agentCustomerTotalCount} matching sidebar filters
                  </>
                ) : (
                  <>
                    {scheduledOnCalendarCount} on calendar · {agentCustomerTotalCount} total outlets
                  </>
                )}
              </div>
            )}
          </div>
          <label
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '4px',
              cursor: agent ? 'pointer' : 'not-allowed',
              flexShrink: 0,
              userSelect: 'none',
              opacity: agent ? 1 : 0.55,
              marginTop: '0',
            }}
            title="勾选 = 打开分区卡片（沿用侧栏 Channel / Category）。在卡片里按 Select 后侧栏切换到该区未上色历清单，同时勾选自动取消（可再勾选重新打开卡片；Cancel / 该区域全部排上会退出未上色历模式）。须先选择 Sales agent。"
          >
            <input
              type="checkbox"
              checked={unassignedBrowseModalOpen}
              onChange={(e) => {
                if (!agent) return;
                if (!e.target.checked) {
                  setSidebarShowUnassignedOnly(false);
                  setUnassignedBrowseModalOpen(false);
                  return;
                }
                setUnassignedBrowseModalOpen(true);
              }}
              disabled={!agent}
              style={{
                width: '12px',
                height: '12px',
                accentColor: '#f43f5e',
                cursor: agent ? 'pointer' : 'not-allowed',
              }}
              aria-label="Show only outlets not yet on the calendar for this agent"
            />
            <span
              style={{
                fontSize: '8px',
                fontWeight: 800,
                color: 'rgba(248, 250, 252, 0.95)',
                textTransform: 'uppercase',
                letterSpacing: '0.08em',
              }}
            >
              Show
            </span>
          </label>
        </div>
        </div>

        <div
          className="sidebar-scroll"
          style={{
            flex: '1 1 0%',
            minHeight: 0,
            overflowY: 'auto',
            overscrollBehavior: 'contain',
            WebkitOverflowScrolling: 'touch',
            background: 'rgba(0,0,0,0.2)',
            padding: '3px',
            borderRadius: '5px',
          }}
        >
          {!agent ? (
            <div style={{ fontSize: '9px', color: 'rgba(148, 163, 184, 0.9)', textAlign: 'center', padding: '6px 3px' }}>
              Select a sales agent to list customers and to use the Unassigned strip.
            </div>
          ) : sidebarShowUnassignedOnly && unassignedCustomerCount === 0 ? (
            <div style={{ fontSize: '9px', color: 'rgba(148, 163, 184, 0.95)', textAlign: 'center', padding: '6px 3px', lineHeight: 1.45 }}>
              Every outlet for <strong style={{ color: '#e2e8f0' }}>{agent}</strong> is already on the calendar (matched by ACC).
            </div>
          ) : sidebarListedCustomers.length === 0 ? (
            sidebarBaseCustomers.length === 0 ? (
              <div style={{ fontSize: '9px', color: 'rgba(148, 163, 184, 0.95)', textAlign: 'center', padding: '6px 3px', lineHeight: 1.45 }}>
                {sidebarShowUnassignedOnly
                  ? 'No outlets left to schedule match area / channel / category — or every customer here is already on the calendar.'
                  : 'No customers match this agent / area / channel / category. Import data or adjust filters.'}
              </div>
            ) : (
              <div style={{ fontSize: '9px', color: 'rgba(148, 163, 184, 0.95)', textAlign: 'center', padding: '6px 3px', lineHeight: 1.45 }}>
                No customers match this visit filter.
                <div style={{ fontSize: '8px', opacity: 0.8, marginTop: '4px' }}>
                  {sidebarBaseCustomers.length} match agent · area · channel · category — try « Any count ».
                </div>
              </div>
            )
          ) : sidebarDisplayedCustomers.length === 0 ? (
            <div style={{ fontSize: '9px', color: 'rgba(148, 163, 184, 0.95)', textAlign: 'center', padding: '6px 3px', lineHeight: 1.45 }}>
              No customers match <strong style={{ color: '#e2e8f0' }}>「{sidebarNameQuery.trim()}」</strong>.
              <div style={{ fontSize: '8px', opacity: 0.8, marginTop: '4px' }}>
                Clear the search or try another keyword — {sidebarListedCustomers.length} still match agent / area / channel / category / visits.
              </div>
            </div>
          ) : (
            <>
              <div
                className="no-print"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '4px',
                  flexWrap: 'wrap',
                  marginBottom: '3px',
                  padding: '1px 3px',
                  borderRadius: '5px',
                  background: 'rgba(15, 23, 42, 0.65)',
                  border: '1px solid rgba(148, 163, 184, 0.2)',
                }}
              >
                <label style={{ display: 'inline-flex', alignItems: 'center', gap: '3px', fontSize: '7px', fontWeight: 600, color: 'rgba(248,250,252,0.92)', cursor: 'pointer', userSelect: 'none' }}>
                  <input
                    ref={sidebarSelectAllRef}
                    type="checkbox"
                    aria-label="Select all customers in list"
                    onMouseDown={(e) => e.stopPropagation()}
                    onChange={() => {
                      setSidebarSelectedRowKeys((prev) => {
                        const visible = sidebarDisplayedCustomers;
                        const allOn =
                          visible.length > 0 &&
                          visible.every((row, i) => prev.has(sidebarRowListKey(row, i)));
                        if (allOn) return new Set();
                        return new Set(visible.map((row, i) => sidebarRowListKey(row, i)));
                      });
                    }}
                    style={{ width: '10px', height: '10px', cursor: 'pointer', accentColor: '#38bdf8' }}
                  />
                  <span>Select all visible</span>
                </label>
                <button
                  type="button"
                  className="no-print"
                  onMouseDown={(e) => e.stopPropagation()}
                  onClick={() => setSidebarSelectedRowKeys(new Set())}
                  disabled={sidebarSelectedRowKeys.size === 0}
                  style={{
                    fontSize: '7px',
                    fontWeight: 600,
                    padding: '1px 6px',
                    borderRadius: '4px',
                    border: '1px solid rgba(148, 163, 184, 0.35)',
                    background: sidebarSelectedRowKeys.size === 0 ? 'rgba(51, 65, 85, 0.5)' : 'rgba(51, 65, 85, 0.95)',
                    color: sidebarSelectedRowKeys.size === 0 ? 'rgba(148, 163, 184, 0.5)' : '#e2e8f0',
                    cursor: sidebarSelectedRowKeys.size === 0 ? 'not-allowed' : 'pointer',
                  }}
                >
                  Clear selection
                </button>
                {sidebarSelectedRowKeys.size > 0 && (
                  <span style={{ fontSize: '7px', color: 'rgba(148, 163, 184, 0.95)', marginLeft: 'auto' }}>
                    {sidebarSelectedRowKeys.size} selected — drag one checked row; selection clears after drop
                  </span>
                )}
              </div>
              {sidebarDisplayedCustomers.map((c, idx) => {
                const rk = sidebarRowListKey(c, idx);
                const monthVisits = visitCountsMonth[c.acc] ?? 0;
                const bucket = sidebarColorBucket(monthVisits);
                const pal = paletteForVisitCountDisplay(monthVisits);
                const badge = monthVisits > 0 ? String(monthVisits) : null;
                const isSelected = sidebarSelectedRowKeys.has(rk);
                const bulkOnList = sidebarDisplayedCustomers.filter((row, i) =>
                  sidebarSelectedRowKeys.has(sidebarRowListKey(row, i)),
                );
                const useBulk = isSelected && bulkOnList.length >= 2;
                return (
                  <div
                    key={rk}
                    draggable
                    onDragStart={(e) => {
                      const payload = useBulk
                        ? {
                            __batch: true,
                            items: bulkOnList.map((row) => ({ acc: row.acc, name: row.name })),
                          }
                        : c;
                      const s = JSON.stringify(payload);
                      e.dataTransfer.setData('text', s);
                      e.dataTransfer.setData('text/plain', s);
                      e.dataTransfer.effectAllowed = 'copyMove';
                    }}
                    style={{
                      position: 'relative',
                      display: 'flex',
                      alignItems: 'stretch',
                      gap: '3px',
                      background: pal.sidebarBg,
                      padding: '3px 3px 3px 4px',
                      borderRadius: '4px',
                      marginBottom: '3px',
                      borderLeft: `3px solid ${isSelected ? '#38bdf8' : pal.sidebarBorder}`,
                      cursor: 'grab',
                      boxShadow: [
                        bucket > 0 ? '0 1px 3px rgba(0,0,0,0.22)' : null,
                        isSelected ? 'inset 0 0 0 2px rgba(56, 189, 248, 0.55)' : null,
                      ]
                        .filter(Boolean)
                        .join(', ') || undefined,
                    }}
                  >
                    <div
                      style={{
                        flexShrink: 0,
                        width: '18px',
                        minWidth: '18px',
                        display: 'flex',
                        justifyContent: 'center',
                        alignItems: 'flex-start',
                        paddingTop: '1px',
                      }}
                    >
                      {badge != null && (
                        <span
                          title={
                            monthVisits >= 6
                              ? `Scheduled ${monthVisits} time(s) in ${monthLabel} (bucket 6 color)`
                              : monthVisits > 0
                                ? `Scheduled ${monthVisits} time(s) in ${monthLabel}`
                                : ''
                          }
                          style={{
                            minWidth: '16px',
                            height: '16px',
                            padding: '0 3px',
                            borderRadius: '999px',
                            fontSize: '8px',
                            fontWeight: 800,
                            display: 'inline-flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            color: '#fff',
                            background: pal.badgeBg ?? '#3b82f6',
                            border: '1px solid rgba(255,255,255,0.25)',
                          }}
                        >
                          {badge}
                        </span>
                      )}
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '7px', color: bucket > 0 ? 'rgba(248, 250, 252, 0.85)' : '#60a5fa' }}>
                        {c.acc}
                        {monthVisits > 0 ? ` · ${monthVisits}× this month` : ''}
                      </div>
                      <div style={{ fontSize: '9px', fontWeight: 700, color: '#f8fafc', lineHeight: 1.18 }}>{c.name}</div>
                    </div>
                    <label
                      style={{
                        display: 'flex',
                        alignItems: 'flex-start',
                        padding: '1px 0 0 2px',
                        cursor: 'pointer',
                        flexShrink: 0,
                      }}
                      onMouseDown={(e) => e.stopPropagation()}
                    >
                      <input
                        type="checkbox"
                        checked={isSelected}
                        aria-label={`Select ${c.name}`}
                        onChange={() => {
                          setSidebarSelectedRowKeys((prev) => {
                            const next = new Set(prev);
                            if (next.has(rk)) next.delete(rk);
                            else next.add(rk);
                            return next;
                          });
                        }}
                        style={{ width: '12px', height: '12px', marginTop: '1px', cursor: 'pointer', accentColor: '#38bdf8' }}
                      />
                    </label>
                  </div>
                );
              })}
            </>
          )}
        </div>
      </aside>

      {/* Main Panel - 必须 flex:1 撑满剩余空间 */}
      <main className="main-column" style={{ flex: 1, display: 'flex', flexDirection: 'column', minHeight: 0, minWidth: 0, background: '#f1f5f9', height: '100%' }}>
        <div className="no-print" style={{ padding: '15px 25px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '12px', flexWrap: 'wrap', background: 'white', borderBottom: '1px solid #e2e8f0' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '0', flexWrap: 'wrap', minWidth: 0, flex: '1 1 auto' }}>
            <div style={{ display: 'flex', gap: '10px', flexShrink: 0 }}>
              <button onClick={()=>setViewMode('calendar')} style={{ padding: '8px 15px', borderRadius: '5px', cursor: 'pointer', background: viewMode==='calendar'?'#1e293b':'white', color: viewMode==='calendar'?'white':'#1e293b', border: viewMode==='calendar'?'none':'1px solid #e2e8f0' }}>Calendar</button>
              <button onClick={()=>setViewMode('summary')} style={{ padding: '8px 15px', borderRadius: '5px', cursor: 'pointer', background: viewMode==='summary'?'#1e293b':'white', color: viewMode==='summary'?'white':'#1e293b', border: viewMode==='summary'?'none':'1px solid #e2e8f0' }}>Summary</button>
            </div>
            {companyName.trim() ? (
              <div className="main-toolbar-company-wrap" title={companyName.trim()}>
                <p className="main-toolbar-company-label">Company</p>
                <p className="app-company-display main-toolbar-company-name">{companyName.trim()}</p>
              </div>
            ) : null}
          </div>
          <button
            type="button"
            onClick={handlePrintPdf}
            title="Print or save as PDF: Calendar and Summary fit one portrait A4 page each."
            style={{ background: '#2563eb', color: 'white', border: 'none', padding: '10px 25px', borderRadius: '8px', cursor: 'pointer', fontWeight:'bold' }}
          >
            PRINT / PDF
          </button>
        </div>

        <div className="main-scroll-region" style={{ flex: 1, minHeight: 0, minWidth: 0, overflowY: 'auto', overflowX: 'hidden', padding: '14px 16px', display: 'flex', flexDirection: 'column' }}>
          <div className="schedule-strip-print" style={{ flexShrink: 0, background: '#1e293b', color: 'white', padding: '15px 25px', borderRadius: '12px 12px 0 0', fontWeight: 'bold', display:'flex', justifyContent:'space-between', alignItems: 'center', gap: '12px', flexWrap: 'wrap', width: '100%', boxSizing: 'border-box' }}>
            <span className="schedule-strip-agent-row" style={{ minWidth: 0 }}>
              {companyName.trim() ? (
                <>
                  <span className="app-company-display app-company-display--on-dark" title={companyName.trim()}>
                    {companyName.trim()}
                  </span>
                  <span style={{ opacity: 0.35, fontWeight: 300 }} aria-hidden>|</span>
                </>
              ) : null}
              <span className="schedule-strip-agent-meta">AGENT: {agent || 'NOT SELECTED'}</span>
            </span>
            <div className="no-print" style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
              <label style={{ fontSize: '11px', fontWeight: 600, opacity: 0.75, textTransform: 'uppercase', marginRight: '2px' }}>Year</label>
              <select
                aria-label="Year"
                value={year}
                onChange={(e) => { setYear(Number(e.target.value)); setFocusedDay(null); }}
                style={selectToolbarStyle}
              >
                {YEAR_RANGE.map((y) => (
                  <option key={y} value={y}>{y}</option>
                ))}
              </select>
              <label style={{ fontSize: '11px', fontWeight: 600, opacity: 0.75, textTransform: 'uppercase', marginLeft: '4px', marginRight: '2px' }}>Month</label>
              <select
                aria-label="Month"
                value={month}
                onChange={(e) => { setMonth(Number(e.target.value)); setFocusedDay(null); }}
                style={{ ...selectToolbarStyle, minWidth: '110px' }}
              >
                {MONTH_SHORT.map((label, m) => (
                  <option key={label} value={m}>{label}</option>
                ))}
              </select>
              <button
                type="button"
                className="no-print"
                onClick={handleCopyMonthVisitsToNextMonth}
                disabled={!agent}
                title="Copies visits to same day next month. Sundays and yellow-calendar public holidays bump forward day-by-day until an open date. Duplicate customers same day skipped."
                style={{
                  marginLeft: '6px',
                  padding: '6px 10px',
                  borderRadius: '6px',
                  border: '1px solid rgba(255,255,255,0.35)',
                  background: agent ? 'rgba(255,255,255,0.14)' : 'rgba(255,255,255,0.06)',
                  color: 'white',
                  fontSize: '11px',
                  fontWeight: 700,
                  cursor: agent ? 'pointer' : 'not-allowed',
                  opacity: agent ? 1 : 0.45,
                  whiteSpace: 'nowrap',
                }}
              >
                Copy → next month
              </button>
              <button
                type="button"
                className="no-print"
                onClick={handleMoveSundayVisitsToMondayCurrentMonth}
                disabled={!agent}
                title="Moves every visit sitting on Sundays in the selected Year/Month to the following Monday. Use this once to repair older copies—or leave new drops alone; Sundays now redirect drops to Monday automatically."
                style={{
                  marginLeft: '6px',
                  padding: '6px 10px',
                  borderRadius: '6px',
                  border: '1px solid rgba(255,255,255,0.35)',
                  background: agent ? 'rgba(255,255,255,0.14)' : 'rgba(255,255,255,0.06)',
                  color: 'white',
                  fontSize: '11px',
                  fontWeight: 700,
                  cursor: agent ? 'pointer' : 'not-allowed',
                  opacity: agent ? 1 : 0.45,
                  whiteSpace: 'nowrap',
                }}
              >
                Sundays → Mon
              </button>
              {viewMode === 'calendar' && (
                <>
                  <label style={{ fontSize: '11px', fontWeight: 600, opacity: 0.75, textTransform: 'uppercase', marginLeft: '4px', marginRight: '2px' }}>Day</label>
                  <select
                    aria-label="Highlight day"
                    value={highlightedDay ?? ''}
                    onChange={(e) => {
                      const v = e.target.value;
                      setFocusedDay(v === '' ? null : Number(v));
                    }}
                    style={{ ...selectToolbarStyle, minWidth: '72px' }}
                  >
                    <option value="">—</option>
                    {Array.from({ length: daysInMonth }, (_, idx) => idx + 1).map((d) => (
                      <option key={d} value={d}>{d}</option>
                    ))}
                  </select>
                </>
              )}
              {viewMode === 'summary' && agent && summarySections && summarySections.totalVisits > 0 && (
                <>
                  <label style={{ fontSize: '11px', fontWeight: 600, opacity: 0.75, textTransform: 'uppercase', marginLeft: '4px', marginRight: '2px' }}>Summary from</label>
                  <select
                    aria-label="Summary date range from day"
                    value={Math.min(Math.max(summaryDayFrom, 1), daysInMonth)}
                    onChange={(e) => setSummaryDayFrom(Number(e.target.value))}
                    style={{ ...selectToolbarStyle, minWidth: '58px' }}
                  >
                    {Array.from({ length: daysInMonth }, (_, idx) => idx + 1).map((d) => (
                      <option key={d} value={d}>{d}</option>
                    ))}
                  </select>
                  <label style={{ fontSize: '11px', fontWeight: 600, opacity: 0.75, textTransform: 'uppercase', marginLeft: '4px', marginRight: '2px' }}>to</label>
                  <select
                    aria-label="Summary date range to day"
                    value={Math.min(Math.max(summaryDayTo, 1), daysInMonth)}
                    onChange={(e) => setSummaryDayTo(Number(e.target.value))}
                    style={{ ...selectToolbarStyle, minWidth: '58px' }}
                  >
                    {Array.from({ length: daysInMonth }, (_, idx) => idx + 1).map((d) => (
                      <option key={d} value={d}>{d}</option>
                    ))}
                  </select>
                </>
              )}
            </div>
            <span style={{ fontWeight: 700, letterSpacing: '0.02em' }}>{monthLabel}</span>
          </div>
          <div
            className="content-white-panel"
            style={{
              background: 'white',
              flex: 1,
              border: '1px solid #e2e8f0',
              display: 'flex',
              flexDirection: 'column',
              width: '100%',
              minWidth: 0,
              boxSizing: 'border-box',
              minHeight: 0,
              overflow: 'auto',
              overscrollBehavior: 'contain',
              WebkitOverflowScrolling: 'touch',
            }}
          >
            {viewMode === 'calendar' ? (
              <>
                <div
                  className="no-print cal-calendar-filters"
                  style={{
                    flexShrink: 0,
                    padding: '8px 12px',
                    borderBottom: '1px solid #e2e8f0',
                    background: '#f8fafc',
                    width: '100%',
                    boxSizing: 'border-box',
                  }}
                >
                  <div
                    style={{
                      display: 'flex',
                      flexWrap: 'wrap',
                      gap: '10px 12px',
                      alignItems: 'flex-end',
                    }}
                  >
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', minWidth: '0', flex: '1 1 120px' }}>
                      <span style={{ fontSize: '10px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Sales agent</span>
                      <select
                        aria-label="Sales agent"
                        value={agent}
                        onChange={(e) => {
              setAgent(e.target.value);
              setSidebarVisitCountFilter('all');
              setCalendarChipRoundFilter('all');
            }}
                        style={{ ...selectCalendarToolbarStyle, width: '100%' }}
                      >
                        <option value="">— Select —</option>
                        {salesAgentOptions.map((a) => (
                          <option key={a} value={a}>
                            {a}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label
                      htmlFor="cal-filter-sync-checkbox"
                      title={
                        syncSidebarCalendarFilters && agent
                          ? 'Filters follow the sidebar (including Visits). Uncheck Sync to edit the calendar toolbar separately.'
                          : 'When Sync is on, Area / Channel / Category / Find / Visits mirror the sidebar toolbar’s matching controls.'
                      }
                      style={{
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '4px',
                        minWidth: '0',
                        flex: '0 0 auto',
                        justifyContent: 'flex-end',
                        cursor: agent ? 'pointer' : 'not-allowed',
                      }}
                    >
                      <span style={{ fontSize: '10px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', lineHeight: 1 }}>
                        Sync
                      </span>
                      <div
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          gap: '8px',
                          padding: '4px 0',
                          minHeight: '34px',
                          boxSizing: 'border-box',
                        }}
                      >
                        <input
                          id="cal-filter-sync-checkbox"
                          type="checkbox"
                          checked={syncSidebarCalendarFilters}
                          disabled={!agent}
                          onChange={(e) => setSyncSidebarCalendarFilters(e.target.checked)}
                          style={{ width: '15px', height: '15px', accentColor: '#2563eb', cursor: agent ? 'pointer' : 'not-allowed', flexShrink: 0 }}
                        />
                        <span
                          style={{
                            fontSize: '11px',
                            fontWeight: 600,
                            color: '#334155',
                            userSelect: 'none',
                            lineHeight: 1.2,
                            maxWidth: '72px',
                          }}
                        >
                          same as sidebar
                        </span>
                      </div>
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', minWidth: '0', flex: '1 1 110px' }}>
                      <span style={{ fontSize: '10px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>
                        Area <span style={{ fontWeight: 600, textTransform: 'none' }}>(calendar)</span>
                      </span>
                      <select
                        aria-label="Calendar area filter"
                        value={calendarArea}
                        onChange={(e) => setCalendarArea(e.target.value)}
                        style={{ ...selectCalendarToolbarStyle, width: '100%' }}
                        disabled={!agent || syncSidebarCalendarFilters}
                      >
                        <option value="">{agent ? '— All areas —' : '— Agent first —'}</option>
                        {filteredAreas.map((a) => (
                          <option key={a} value={a}>{a}</option>
                        ))}
                      </select>
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', minWidth: '0', flex: '1 1 120px' }}>
                      <span style={{ fontSize: '10px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Visits</span>
                      <select
                        aria-label="Filter calendar visits by number on bar"
                        value={calendarChipRoundFilter}
                        onChange={(e) => setCalendarChipRoundFilter(e.target.value)}
                        title={
                          syncSidebarCalendarFilters
                            ? 'Follows sidebar Visits when Sync is on.'
                            : 'Show bars whose badge # matches. Not the same statistic as sidebar “month total visits” unless Sync copies the sidebar choice.'
                        }
                        style={{ ...selectCalendarToolbarStyle, width: '100%' }}
                        disabled={!agent || syncSidebarCalendarFilters}
                      >
                        <option value="all">{agent ? '— All visits —' : '— Agent first —'}</option>
                        <option value="0">Hide all chips</option>
                        <option value="1">1 visit</option>
                        <option value="2">2 visit</option>
                        <option value="3">3 visit</option>
                        <option value="4">4 visit</option>
                        <option value="5">5 visit</option>
                        <option value="6+">6+ visit</option>
                      </select>
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', minWidth: '0', flex: '1 1 110px' }}>
                      <span style={{ fontSize: '10px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>
                        Channel <span style={{ fontWeight: 600, textTransform: 'none' }}>(calendar)</span>
                      </span>
                      <select
                        aria-label="Calendar channel filter"
                        value={calendarChannelFilter}
                        onChange={(e) => setCalendarChannelFilter(e.target.value)}
                        style={{ ...selectCalendarToolbarStyle, width: '100%' }}
                        disabled={!agent || syncSidebarCalendarFilters}
                      >
                        <option value="">{agent ? '— All channels —' : '— Agent first —'}</option>
                        {calendarChannelOptions.map((ch) => (
                          <option key={ch} value={ch}>{ch}</option>
                        ))}
                      </select>
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', minWidth: '0', flex: '1 1 110px' }}>
                      <span style={{ fontSize: '10px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>
                        Category <span style={{ fontWeight: 600, textTransform: 'none' }}>(calendar)</span>
                      </span>
                      <select
                        aria-label="Calendar category filter"
                        value={calendarCategoryFilter}
                        onChange={(e) => setCalendarCategoryFilter(e.target.value)}
                        style={{ ...selectCalendarToolbarStyle, width: '100%' }}
                        disabled={!agent || syncSidebarCalendarFilters}
                      >
                        <option value="">{agent ? '— All categories —' : '— Agent first —'}</option>
                        {calendarCategoryOptions.map((cat) => (
                          <option key={cat} value={cat}>{cat}</option>
                        ))}
                      </select>
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', minWidth: '0', flex: '1 1 200px' }}>
                      <span style={{ fontSize: '10px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Find on calendar</span>
                      <input
                        type="search"
                        aria-label="Filter calendar visits by name or account (sidebar list uses the left panel Find)"
                        value={calendarNameQuery}
                        onChange={(e) => setCalendarNameQuery(e.target.value)}
                        placeholder={agent ? 'Name or account…' : 'Agent first'}
                        disabled={!agent || syncSidebarCalendarFilters}
                        autoCapitalize="off"
                        spellCheck={false}
                        style={{
                          width: '100%',
                          padding: '6px 10px',
                          borderRadius: '6px',
                          border: `1px solid ${agent ? '#cbd5e1' : '#e2e8f0'}`,
                          background: agent ? '#fff' : '#f1f5f9',
                          color: '#1e293b',
                          fontSize: '12px',
                          fontWeight: 500,
                          outline: 'none',
                          boxSizing: 'border-box',
                          opacity: agent ? 1 : 0.7,
                        }}
                      />
                    </label>
                  </div>
                </div>
                {renderCalendar()}
              </>
            ) : (
              renderSummary()
            )}
          </div>
        </div>
      </main>
      {calendarDayDetailDateStr != null && agent ? (
        <div
          className="no-print"
          role="dialog"
          aria-modal="true"
          aria-labelledby="cal-day-card-title"
          style={{
            position: 'fixed',
            inset: 0,
            zIndex: 10050,
            background: 'rgba(15, 23, 42, 0.55)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px',
            boxSizing: 'border-box',
          }}
          onClick={() => setCalendarDayDetailDateStr(null)}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: '#ffffff',
              borderRadius: '14px',
              maxWidth: '440px',
              width: '100%',
              maxHeight: 'min(85vh, 520px)',
              overflow: 'hidden',
              display: 'flex',
              flexDirection: 'column',
              boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.35)',
              border: '1px solid #e2e8f0',
            }}
          >
            <div
              style={{
                padding: '16px 18px 12px',
                borderBottom: '1px solid #e2e8f0',
                flexShrink: 0,
              }}
            >
              <h2 id="cal-day-card-title" style={{ margin: 0, fontSize: '16px', fontWeight: 800, color: '#0f172a' }}>
                Visits this day
              </h2>
              <div style={{ marginTop: '6px', fontSize: '12px', color: '#64748b', fontWeight: 500, lineHeight: 1.4 }}>
                {formatIsoDateHeading(calendarDayDetailDateStr)} · {agent}
              </div>
            </div>
            <div style={{ padding: '12px 18px', overflowY: 'auto', flex: '1 1 auto', minHeight: 0 }}>
              {calendarDayCardTasks.length === 0 ? (
                <div style={{ fontSize: '13px', color: '#64748b', textAlign: 'center', padding: '24px 8px' }}>
                  No visits scheduled on this day.
                </div>
              ) : (
                <ul style={{ listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  {calendarDayCardTasks.map((t, idx) => (
                    <li
                      key={`${t.acc}-${normalizeVisitRound(t.visitRound ?? 1)}-${idx}`}
                      style={{
                        padding: '12px 14px',
                        borderRadius: '10px',
                        background: '#f8fafc',
                        border: '1px solid #e2e8f0',
                      }}
                    >
                      <div
                        style={{
                          fontSize: '11px',
                          fontWeight: 700,
                          color: '#64748b',
                          textTransform: 'uppercase',
                          letterSpacing: '0.04em',
                          marginBottom: '4px',
                        }}
                      >
                        Account
                      </div>
                      <div
                        style={{
                          fontSize: '14px',
                          fontWeight: 700,
                          color: '#0f172a',
                          fontFamily: 'ui-monospace, "Cascadia Code", monospace',
                        }}
                      >
                        {t.acc}
                      </div>
                      <div
                        style={{
                          fontSize: '11px',
                          fontWeight: 700,
                          color: '#64748b',
                          textTransform: 'uppercase',
                          letterSpacing: '0.04em',
                          marginTop: '10px',
                          marginBottom: '4px',
                        }}
                      >
                        Customer
                      </div>
                      <div style={{ fontSize: '15px', fontWeight: 600, color: '#1e293b', lineHeight: 1.35 }}>{t.name}</div>
                      <div style={{ marginTop: '8px', fontSize: '11px', color: '#94a3b8' }}>
                        Visit #{normalizeVisitRound(t.visitRound ?? 1)}
                      </div>
                    </li>
                  ))}
                </ul>
              )}
            </div>
            <div style={{ padding: '12px 18px', borderTop: '1px solid #e2e8f0', flexShrink: 0 }}>
              <button
                type="button"
                onClick={() => setCalendarDayDetailDateStr(null)}
                style={{
                  width: '100%',
                  padding: '10px 14px',
                  borderRadius: '8px',
                  border: '1px solid #cbd5e1',
                  background: '#f1f5f9',
                  color: '#0f172a',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: 'pointer',
                }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      ) : null}
      {unassignedBrowseModalOpen && agent ? (
        <div
          className="no-print"
          role="dialog"
          aria-modal="true"
          aria-labelledby="unassigned-browse-title"
          style={{
            position: 'fixed',
            inset: 0,
            zIndex: 10050,
            background: 'rgba(15, 23, 42, 0.55)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px',
            boxSizing: 'border-box',
          }}
          onClick={() => {
            setUnassignedBrowseModalOpen(false);
            setSidebarShowUnassignedOnly(false);
          }}
        >
          <div
            role="presentation"
            onClick={(e) => e.stopPropagation()}
            style={{
              background: '#ffffff',
              borderRadius: '14px',
              maxWidth: '480px',
              width: '100%',
              maxHeight: 'min(85vh, 560px)',
              overflow: 'hidden',
              display: 'flex',
              flexDirection: 'column',
              boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.35)',
              border: '1px solid #e2e8f0',
            }}
          >
            <div
              style={{
                padding: '16px 18px 12px',
                borderBottom: '1px solid #e2e8f0',
                flexShrink: 0,
              }}
            >
              <h2 id="unassigned-browse-title" style={{ margin: 0, fontSize: '16px', fontWeight: 800, color: '#0f172a' }}>
                Not yet on calendar (by area)
              </h2>
              <div style={{ marginTop: '6px', fontSize: '12px', color: '#64748b', fontWeight: 500, lineHeight: 1.4 }}>
                Sales agent <strong style={{ fontWeight: 700, color: '#334155' }}>{agent}</strong> — counts use sidebar Channel /
                Category filters (all areas combined). Click <strong style={{ fontWeight: 700, color: '#334155' }}>Select</strong>{' '}
                next to a number to jump the sidebar Area and list only outlets still not scheduled there.
              </div>
            </div>
            <div style={{ padding: '12px 18px', overflowY: 'auto', flex: '1 1 auto', minHeight: 0 }}>
              {unassignedBrowseModalTotal === 0 ? (
                <div style={{ fontSize: '13px', color: '#64748b', textAlign: 'center', padding: '20px 8px', lineHeight: 1.5 }}>
                  Every outlet matching the current sidebar filters already has at least one visit on the calendar.
                </div>
              ) : (
                <>
                  <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', marginBottom: '8px' }}>
                    Sidebar area breakdown
                  </div>
                  <table
                    style={{
                      width: '100%',
                      borderCollapse: 'collapse',
                      fontSize: '13px',
                    }}
                  >
                    <thead>
                      <tr>
                        <th
                          style={{
                            textAlign: 'left',
                            padding: '8px 10px',
                            borderBottom: '1px solid #e2e8f0',
                            color: '#64748b',
                            fontWeight: 700,
                          }}
                        >
                          Area
                        </th>
                        <th
                          style={{
                            textAlign: 'right',
                            padding: '8px 10px',
                            borderBottom: '1px solid #e2e8f0',
                            color: '#64748b',
                            fontWeight: 700,
                            whiteSpace: 'nowrap',
                          }}
                        >
                          Not scheduled · Select
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {unassignedBrowseSelectAreaRows.map((row) => (
                        <tr key={`${row.areaLabel}\u001fcounts`}>
                          <td
                            style={{
                              padding: '8px 10px',
                              borderBottom: '1px solid #f1f5f9',
                              color: '#0f172a',
                              fontWeight: 600,
                            }}
                          >
                            {row.areaLabel === '' ? '(No area in data)' : row.areaLabel}
                          </td>
                          <td
                            style={{
                              padding: '6px 10px',
                              borderBottom: '1px solid #f1f5f9',
                              verticalAlign: 'middle',
                              textAlign: 'right',
                              color: '#0f172a',
                            }}
                          >
                            <div
                              style={{
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'flex-end',
                                gap: '10px',
                                flexWrap: 'wrap',
                              }}
                            >
                              <span
                                style={{
                                  fontVariantNumeric: 'tabular-nums',
                                  fontWeight: 700,
                                  fontSize: '13px',
                                }}
                              >
                                {row.count}
                              </span>
                              <button
                                type="button"
                                className="no-print"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  applyUnassignedBrowseAreaToSidebar(row.areaLabel);
                                }}
                                aria-label={
                                  row.areaLabel === ''
                                    ? 'Show unassigned outlets with no area in data'
                                    : `Show unassigned outlets in ${row.areaLabel}`
                                }
                                title="Sidebar: this area · unassigned only"
                                style={{
                                  flexShrink: 0,
                                  padding: '4px 10px',
                                  borderRadius: '6px',
                                  border: '1px solid #cbd5e1',
                                  background: '#f8fafc',
                                  color: '#0f172a',
                                  fontSize: '11px',
                                  fontWeight: 700,
                                  cursor: 'pointer',
                                }}
                              >
                                Select
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr>
                        <td
                          style={{
                            padding: '10px',
                            borderTop: '1px solid #e2e8f0',
                            fontWeight: 800,
                            color: '#0f172a',
                          }}
                        >
                          Total
                        </td>
                        <td
                          style={{
                            padding: '8px 10px',
                            borderTop: '1px solid #e2e8f0',
                            verticalAlign: 'middle',
                            textAlign: 'right',
                            color: '#0f172a',
                          }}
                        >
                          <div
                            style={{
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'flex-end',
                              gap: '10px',
                              flexWrap: 'wrap',
                            }}
                          >
                            <span
                              style={{
                                fontVariantNumeric: 'tabular-nums',
                                fontWeight: 800,
                              }}
                            >
                              {unassignedBrowseModalTotal}
                            </span>
                            <button
                              type="button"
                              className="no-print"
                              onClick={(e) => {
                                e.stopPropagation();
                                applyUnassignedBrowseAreaToSidebar('');
                              }}
                              aria-label="Show unassigned outlets in all areas"
                              title="Sidebar: all areas · unassigned only"
                              style={{
                                flexShrink: 0,
                                padding: '4px 10px',
                                borderRadius: '6px',
                                border: '1px solid #94a3b8',
                                background: '#f1f5f9',
                                color: '#0f172a',
                                fontSize: '11px',
                                fontWeight: 700,
                                cursor: 'pointer',
                              }}
                            >
                              All areas
                            </button>
                          </div>
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </>
              )}
            </div>
            <div style={{ padding: '12px 18px', borderTop: '1px solid #e2e8f0', flexShrink: 0 }}>
              <button
                type="button"
                onClick={() => {
                  setUnassignedBrowseModalOpen(false);
                  setSidebarShowUnassignedOnly(false);
                }}
                style={{
                  width: '100%',
                  padding: '10px 14px',
                  borderRadius: '8px',
                  border: '1px solid #cbd5e1',
                  background: '#f1f5f9',
                  color: '#0f172a',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: 'pointer',
                }}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      ) : null}
      {clearImportDialogOpen ? (
        <div
          className="no-print"
          role="dialog"
          aria-modal="true"
          aria-labelledby="clear-import-title"
          style={{
            position: 'fixed',
            inset: 0,
            zIndex: 10060,
            background: 'rgba(15, 23, 42, 0.55)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '16px',
            boxSizing: 'border-box',
          }}
          onClick={dismissClearImportDialog}
        >
          <div
            role="presentation"
            onMouseDown={(e) => e.stopPropagation()}
            onClick={(e) => e.stopPropagation()}
            style={{
              background: '#ffffff',
              borderRadius: '12px',
              maxWidth: '400px',
              width: '100%',
              overflow: 'hidden',
              boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.35)',
              border: '1px solid #e2e8f0',
              colorScheme: 'light',
            }}
          >
            <div style={{ padding: '14px 16px 10px', borderBottom: '1px solid #e2e8f0' }}>
              <h2 id="clear-import-title" style={{ margin: 0, fontSize: '15px', fontWeight: 800, color: '#0f172a' }}>
                Clear saved import
              </h2>
              <p style={{ margin: '8px 0 0', fontSize: '12px', color: '#64748b', lineHeight: 1.45 }}>
                This removes customers, visits{hasElectronDataBridge() ? ', and the app data file on this PC' : ', and browser storage'}.
                Does not delete your spreadsheet file on disk.
              </p>
            </div>
            <div style={{ padding: '12px 16px 14px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
              {clearImportStep === 'unlock' && loadClearImportCustomPassword() == null ? (
                <div
                  style={{
                    fontSize: '12px',
                    lineHeight: 1.45,
                    color: '#1e40af',
                    background: '#eff6ff',
                    border: '1px solid #bfdbfe',
                    borderRadius: '8px',
                    padding: '10px 12px',
                  }}
                >
                  <strong style={{ fontWeight: 700 }}>Hint:</strong> no custom password is saved yet — use factory default{' '}
                  <span style={{ fontFamily: 'ui-monospace, monospace', fontWeight: 800 }}>1234</span> below. After Continue,
                  you will set your own password for future clears (or reuse 1234 if you prefer).
                </div>
              ) : null}

              {clearImportStep === 'unlock' ? (
                <>
                  <label style={{ fontSize: '11px', fontWeight: 700, color: '#334155' }} htmlFor="clear-import-password">
                    Current password
                  </label>
                  <input
                    id="clear-import-password"
                    type="password"
                    autoComplete="off"
                    spellCheck={false}
                    autoCapitalize="off"
                    autoCorrect="off"
                    value={clearImportPassword}
                    onChange={(e) => {
                      setClearImportPassword(e.target.value);
                      setClearImportError('');
                    }}
                    style={{
                      ...clearImportModalPasswordInputStyle,
                    }}
                  />
                </>
              ) : null}

              {clearImportStep === 'setNewPassword' ? (
                <>
                  <p style={{ margin: 0, fontSize: '12px', color: '#475569', lineHeight: 1.45 }}>
                    Choose a new password (at least 4 characters). You will need it the next time you clear saved import data.
                  </p>
                  <label style={{ fontSize: '11px', fontWeight: 700, color: '#334155' }} htmlFor="clear-import-new-password">
                    New password
                  </label>
                  <input
                    id="clear-import-new-password"
                    type="password"
                    autoComplete="new-password"
                    spellCheck={false}
                    autoCapitalize="off"
                    autoCorrect="off"
                    value={clearImportNewPassword}
                    onChange={(e) => {
                      setClearImportNewPassword(e.target.value);
                      setClearImportError('');
                    }}
                    style={{
                      ...clearImportModalPasswordInputStyle,
                    }}
                  />
                  <label style={{ fontSize: '11px', fontWeight: 700, color: '#334155' }} htmlFor="clear-import-new-password-confirm">
                    Confirm new password
                  </label>
                  <input
                    id="clear-import-new-password-confirm"
                    type="password"
                    autoComplete="new-password"
                    spellCheck={false}
                    autoCapitalize="off"
                    autoCorrect="off"
                    value={clearImportNewPasswordConfirm}
                    onChange={(e) => {
                      setClearImportNewPasswordConfirm(e.target.value);
                      setClearImportError('');
                    }}
                    style={{
                      ...clearImportModalPasswordInputStyle,
                    }}
                  />
                </>
              ) : null}

              {clearImportStep === 'confirmErase' ? (
                <p style={{ margin: 0, fontSize: '12px', color: '#475569', lineHeight: 1.45 }}>
                  Password verified. Confirm that you want to erase all locally saved Route Planner data.
                </p>
              ) : null}

              {clearImportStep === 'setNewPassword' || clearImportStep === 'confirmErase' ? (
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: '8px', fontSize: '12px', color: '#475569', cursor: 'pointer', userSelect: 'none', lineHeight: 1.35 }}>
                  <input
                    type="checkbox"
                    checked={clearImportAckDanger}
                    onChange={(e) => {
                      setClearImportAckDanger(e.target.checked);
                      setClearImportError('');
                    }}
                    style={{ marginTop: '2px', width: '14px', height: '14px', accentColor: '#dc2626', flexShrink: 0 }}
                  />
                  <span>I understand that all Route Planner data saved locally on this computer will be deleted.</span>
                </label>
              ) : null}
              {clearImportError ? (
                <div style={{ fontSize: '12px', fontWeight: 600, color: '#b91c1c' }}>{clearImportError}</div>
              ) : null}
              <div style={{ display: 'flex', gap: '8px', marginTop: '4px' }}>
                <button
                  type="button"
                  onClick={dismissClearImportDialog}
                  style={{
                    flex: 1,
                    padding: '9px 12px',
                    borderRadius: '8px',
                    border: '1px solid #cbd5e1',
                    background: '#f8fafc',
                    color: '#0f172a',
                    fontSize: '13px',
                    fontWeight: 700,
                    cursor: 'pointer',
                  }}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={
                    clearImportStep === 'unlock'
                      ? submitClearImportUnlockStep
                      : clearImportStep === 'setNewPassword'
                        ? submitClearImportEraseWithNewPassword
                        : submitClearImportEraseConfirmed
                  }
                  style={{
                    flex: 1,
                    padding: '9px 12px',
                    borderRadius: '8px',
                    border: '1px solid rgba(185, 28, 28, 0.45)',
                    background: '#fef2f2',
                    color: '#991b1b',
                    fontSize: '13px',
                    fontWeight: 700,
                    cursor: 'pointer',
                  }}
                >
                  {clearImportStep === 'unlock' ? 'Continue' : 'Erase saved data'}
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : null}
      {settingsDialogOpen ? (
        <div
          className="no-print"
          role="dialog"
          aria-modal="true"
          aria-labelledby="settings-dialog-title"
          style={{
            position: 'fixed',
            inset: 0,
            zIndex: 10055,
            background: 'rgba(15, 23, 42, 0.55)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '16px',
            boxSizing: 'border-box',
          }}
          onClick={dismissSettingsDialog}
        >
          <div
            role="presentation"
            onMouseDown={(e) => e.stopPropagation()}
            onClick={(e) => e.stopPropagation()}
            style={{
              background: '#ffffff',
              borderRadius: '12px',
              maxWidth: '420px',
              width: '100%',
              overflow: 'hidden',
              boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.35)',
              border: '1px solid #e2e8f0',
              colorScheme: 'light',
            }}
          >
            <div style={{ padding: '14px 16px 10px', borderBottom: '1px solid #e2e8f0' }}>
              <h2 id="settings-dialog-title" style={{ margin: 0, fontSize: '15px', fontWeight: 800, color: '#0f172a' }}>
                Settings
              </h2>
              <p style={{ margin: '8px 0 0', fontSize: '12px', color: '#64748b', lineHeight: 1.45 }}>
                Company name appears in the sidebar, calendar bar, Summary / PDF, and window title. Leave blank to hide.
              </p>
            </div>
            <div style={{ padding: '12px 16px 14px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <label style={{ fontSize: '11px', fontWeight: 700, color: '#334155' }} htmlFor="settings-company-name">
                Company name
              </label>
              <input
                id="settings-company-name"
                type="text"
                autoComplete="organization"
                spellCheck={false}
                autoCapitalize="words"
                maxLength={120}
                value={settingsDraftCompanyName}
                onChange={(e) => {
                  setSettingsDraftCompanyName(e.target.value);
                  setSettingsError('');
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') submitSettingsSave();
                }}
                placeholder="e.g. ABC Trading Sdn Bhd"
                style={{ ...clearImportModalPasswordInputStyle }}
              />
              {settingsError ? (
                <div style={{ fontSize: '12px', fontWeight: 600, color: '#b91c1c' }}>{settingsError}</div>
              ) : null}
              <div style={{ display: 'flex', gap: '8px', marginTop: '4px' }}>
                <button
                  type="button"
                  onClick={dismissSettingsDialog}
                  style={{
                    flex: 1,
                    padding: '9px 12px',
                    borderRadius: '8px',
                    border: '1px solid #cbd5e1',
                    background: '#fff',
                    color: '#0f172a',
                    fontSize: '13px',
                    fontWeight: 700,
                    cursor: 'pointer',
                  }}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={submitSettingsSave}
                  style={{
                    flex: 1,
                    padding: '9px 12px',
                    borderRadius: '8px',
                    border: '1px solid #2563eb',
                    background: '#2563eb',
                    color: '#fff',
                    fontSize: '13px',
                    fontWeight: 700,
                    cursor: 'pointer',
                  }}
                >
                  Save
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

export default App;