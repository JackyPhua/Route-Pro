import { useState, useMemo, useEffect, useRef, useLayoutEffect, useCallback } from 'react';
import { Upload } from 'lucide-react';
import * as XLSX from 'xlsx';

const SABAH_HOLIDAYS = { "2026-05-01": "Labour Day", "2026-05-30": "Kaamatan", "2026-05-31": "Kaamatan" };
const MONTH_SHORT = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

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
  padding: '6px 28px 6px 8px',
  borderRadius: '6px',
  border: '1px solid rgba(148, 163, 184, 0.35)',
  backgroundColor: '#1e293b',
  backgroundImage: `linear-gradient(180deg, #3d4f63 0%, #1e293b 100%), ${SELECT_CHEVRON}`,
  backgroundRepeat: 'no-repeat, no-repeat',
  backgroundSize: '100% 100%, 10px 10px',
  backgroundPosition: '0 0, right 7px center',
  color: '#f8fafc',
  fontSize: '10px',
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

const CAL_CHIP_FONT_MIN = 7;
const CAL_CHIP_FONT_MAX = 12;

/** 有效拜访序号，至少为 1 */
function normalizeVisitRound(r) {
  const n = Math.floor(Number(r));
  return Number.isFinite(n) && n >= 1 ? n : 1;
}

/** 将序号映射到 1–6 配色桶（≥6 次均为桶 6） */
function colorBucketFromRound(round) {
  return Math.min(normalizeVisitRound(round), 6);
}

/** 单日格拜访条列表；字号由外层 CalWeekRow 按「整周一排」统一计算后传入 fontPx */
function CalDayTasksColumn({ dateStr, tasks, removeScheduledVisit, fontPx }) {
  return (
    <div
      className="cal-day-tasks"
      style={{
        marginTop: '8px',
        display: 'flex',
        flexDirection: 'column',
        gap: '7px',
        flex: '1 1 auto',
        minHeight: '48px',
        minWidth: 0,
        overflowX: 'hidden',
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
            draggable
            title={t.name}
            onDragStart={(e) =>
              e.dataTransfer.setData('text', JSON.stringify({ ...t, oldDate: dateStr }))
            }
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '7px',
              background: chipBg,
              color: chipFg,
              padding: '5px 7px 5px 9px',
              borderRadius: '7px',
              cursor: 'move',
              overflow: 'hidden',
              minHeight: '26px',
              border: `1px solid rgba(255,255,255,0.22)`,
            }}
          >
            <span
              style={{
                flexShrink: 0,
                minWidth: '20px',
                height: '18px',
                padding: '0 2px',
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                borderRadius: '4px',
                fontSize: '10px',
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
                fontSize: fontPx,
                lineHeight: 1.2,
                whiteSpace: 'nowrap',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
                fontFamily: 'system-ui, -apple-system, "Segoe UI", sans-serif',
              }}
            >
              {`• ${t.name}`}
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
              aria-label={`Remove ${t.name}`}
              title="Remove from this day"
              style={{
                flexShrink: 0,
                width: '22px',
                height: '22px',
                padding: 0,
                margin: 0,
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer',
                fontSize: '15px',
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

/** 一周内所有拜访条共用同一字号（整排行同步缩放）；宽度取自该行任一条的文字区（各列等宽） */
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
}) {
  const rowHostRef = useRef(null);
  const [rowFontPx, setRowFontPx] = useState(CAL_CHIP_FONT_MAX);

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

  const rowTaskCount = useMemo(
    () => enriched.reduce((n, c) => n + (c.kind === 'day' ? c.tasks.length : 0), 0),
    [enriched],
  );

  const rowTaskKey = useMemo(
    () =>
      `${agent}\u0003` +
      enriched
        .filter((c) => c.kind === 'day')
        .flatMap((c) => c.tasks)
        .map((t) => `${t.acc}\u0001${t.name}\u0001${normalizeVisitRound(t.visitRound ?? 1)}`)
        .join('\u0002'),
    [enriched, agent],
  );

  useLayoutEffect(() => {
    const host = rowHostRef.current;
    if (!host || rowTaskCount === 0) {
      setRowFontPx(CAL_CHIP_FONT_MAX);
      return;
    }
    const measureRow = () => {
      const probe = host.querySelector('[data-cal-chip-text]');
      const w = probe instanceof HTMLElement ? probe.clientWidth : 0;
      if (w < 6) return;

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
        `width:${w}px`,
        'line-height:1.2',
        'margin:0',
        'padding:0',
        'box-sizing:border-box',
        'font-family:system-ui,-apple-system,"Segoe UI",sans-serif',
      ].join(';');

      document.body.appendChild(measureEl);
      let synced = CAL_CHIP_FONT_MAX;
      try {
        const allVisits = enriched.filter((c) => c.kind === 'day').flatMap((c) => c.tasks);
        for (const t of allVisits) {
          let size = CAL_CHIP_FONT_MAX;
          measureEl.textContent = `• ${t.name}`;
          measureEl.style.fontSize = `${size}px`;
          while (size > CAL_CHIP_FONT_MIN && measureEl.scrollWidth > w + 0.5) {
            size -= 0.5;
            measureEl.style.fontSize = `${size}px`;
          }
          synced = Math.min(synced, size);
        }
      } finally {
        try {
          measureEl.remove();
        } catch {
          /* ignore detach errors */
        }
      }
      setRowFontPx((prev) => {
        const snapped = Math.round(synced * 2) / 2;
        const prevSnapped = Math.round(prev * 2) / 2;
        return Math.abs(prevSnapped - snapped) < 0.11 ? prev : snapped;
      });
    };

    if (typeof document === 'undefined') return;

    try {
      measureRow();
    } catch (err) {
      console.warn('route-planner: row font measure failed', err);
      return;
    }

    const onResize = () => {
      try {
        measureRow();
      } catch (err) {
        console.warn('route-planner: row font measure on resize failed', err);
      }
    };
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, [rowTaskCount, rowTaskKey, enriched]);

  return (
    <div ref={rowHostRef} style={{ display: 'contents' }}>
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
                boxSizing: 'border-box',
              }}
            />
          );
        }
        const { dayNum, dateStr, isSun, tasks } = cell;
        const isFocused = highlightedDayNum != null && highlightedDayNum === dayNum;
        return (
          <div
            key={cell.key}
            className="cal-day-cell"
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault();
              const raw = e.dataTransfer.getData('text') || e.dataTransfer.getData('text/plain');
              dropOnDate(dateStr, raw);
            }}
            style={{
              minHeight: '144px',
              display: 'flex',
              flexDirection: 'column',
              border: '1px solid #cbd5e1',
              padding: '10px',
              background: SABAH_HOLIDAYS[dateStr] ? '#fef08a' : (isSun ? '#ffe4e6' : 'white'),
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
              fontPx={rowFontPx}
            />
          </div>
        );
      })}
    </div>
  );
}

function paletteForColorBucket(bucket) {
  const b = Math.min(Math.max(Math.floor(bucket), 1), 6);
  return { ...UNSCHED_SIDEBAR, ...VISIT_BUCKET_PALETTE[b] };
}

/** 已排期总次数 → 侧栏用色（与最高桶一致） */
function paletteForTotalVisits(total) {
  if (total <= 0) return UNSCHED_SIDEBAR;
  return paletteForColorBucket(Math.min(total, 6));
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
  fontSize: '8px',
  fontWeight: 700,
  letterSpacing: '0.1em',
  textTransform: 'uppercase',
  color: 'rgba(148, 163, 184, 0.95)',
  marginBottom: '3px',
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
  const [schedules, setSchedules] = useState(persistedOnMount.schedules);
  const [year, setYear] = useState(2026);
  /** 0-indexed month (aligned with Date — 4 = May) */
  const [month, setMonth] = useState(4);
  /** Highlights a calendar cell (1 … days in month); only used in Calendar view */
  const [focusedDay, setFocusedDay] = useState(null);
  const [viewMode, setViewMode] = useState('calendar');
  const [importFileName, setImportFileName] = useState(persistedOnMount.importFileName);
  const [sidebarVisitCountFilter, setSidebarVisitCountFilter] = useState('all');
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
        if (h) {
          setCustomers(h.customers);
          setSchedules(h.schedules);
          setImportFileName(h.importFileName);
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

  const sidebarChannelOptions = useMemo(() => {
    if (!agent) return [];
    const set = new Set(
      customers
        .filter((c) => c.agent === agent && (!sidebarArea || c.area === sidebarArea))
        .map((c) => String(c.channel ?? '—')),
    );
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
  }, [agent, sidebarArea, customers]);

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
    if (!agent) return [];
    const set = new Set(
      customers
        .filter(
          (c) =>
            c.agent === agent &&
            (!sidebarArea || c.area === sidebarArea) &&
            (!sidebarChannelFilter || String(c.channel ?? '—') === sidebarChannelFilter),
        )
        .map((c) => String(c.category ?? '—')),
    );
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
  }, [agent, sidebarArea, sidebarChannelFilter, customers]);

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

  const sidebarBaseCustomers = useMemo(() => {
    if (!agent) return [];
    return customers.filter(
      (c) =>
        c.agent === agent &&
        (!sidebarArea || c.area === sidebarArea) &&
        (!sidebarChannelFilter || String(c.channel ?? '—') === sidebarChannelFilter) &&
        (!sidebarCategoryFilter || String(c.category ?? '—') === sidebarCategoryFilter),
    );
  }, [agent, sidebarArea, sidebarChannelFilter, sidebarCategoryFilter, customers]);

  const sidebarListedCustomers = useMemo(
    () =>
      sidebarBaseCustomers.filter((c) =>
        matchesVisitScheduleFilter(visitCountsMonth[c.acc], sidebarVisitCountFilter),
      ),
    [sidebarBaseCustomers, visitCountsMonth, sidebarVisitCountFilter],
  );

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
  }, [agent]);

  useEffect(() => {
    setSidebarChannelFilter('');
    setSidebarCategoryFilter('');
  }, [sidebarArea]);

  useEffect(() => {
    setSidebarCategoryFilter('');
  }, [sidebarChannelFilter]);

  useEffect(() => {
    setCalendarChannelFilter('');
    setCalendarCategoryFilter('');
  }, [calendarArea]);

  useEffect(() => {
    setCalendarCategoryFilter('');
  }, [calendarChannelFilter]);

  /** Sidebar ⇄ Calendar：Area / Channel / Category + Find 搜尋字串雙向同步（Sync 打勾時） */
  useEffect(() => {
    if (!syncSidebarCalendarFilters) return;
    setCalendarArea(sidebarArea);
  }, [syncSidebarCalendarFilters, sidebarArea]);

  useEffect(() => {
    if (!syncSidebarCalendarFilters) return;
    setSidebarArea(calendarArea);
  }, [syncSidebarCalendarFilters, calendarArea]);

  useEffect(() => {
    if (!syncSidebarCalendarFilters) return;
    setCalendarChannelFilter(sidebarChannelFilter);
    setCalendarCategoryFilter(sidebarCategoryFilter);
  }, [syncSidebarCalendarFilters, sidebarChannelFilter, sidebarCategoryFilter]);

  useEffect(() => {
    if (!syncSidebarCalendarFilters) return;
    setSidebarChannelFilter(calendarChannelFilter);
    setSidebarCategoryFilter(calendarCategoryFilter);
  }, [syncSidebarCalendarFilters, calendarChannelFilter, calendarCategoryFilter]);

  useEffect(() => {
    if (!syncSidebarCalendarFilters) return;
    setCalendarNameQuery(sidebarNameQuery);
  }, [syncSidebarCalendarFilters, sidebarNameQuery]);

  useEffect(() => {
    if (!syncSidebarCalendarFilters) return;
    setSidebarNameQuery(calendarNameQuery);
  }, [syncSidebarCalendarFilters, calendarNameQuery]);

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

  /** 侧栏：已排总数 → 调色桶（≥6 与第 6 色相同）；角标显示真实总数 */
  const sidebarColorBucket = (total) => {
    if (total <= 0) return 0;
    return Math.min(total, 6);
  };
  const monthIsoPad = String(month + 1).padStart(2, '0');
  const monthLabel = `${year}-${MONTH_SHORT[month]}`;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const highlightedDay =
    focusedDay != null && focusedDay >= 1 && focusedDay <= daysInMonth ? focusedDay : null;

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
              weekday: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][new Date(dateStr).getDay()],
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
        <div className="summary-print-wrap" style={{ padding: '24px', color: '#64748b', flex: 1 }}>
          Choose a sales agent to see visits scheduled for this month.
        </div>
      );
    }
    if (!summarySections || summarySections.totalVisits === 0) {
      return (
        <div className="summary-print-wrap" style={{ padding: '24px', color: '#64748b', flex: 1 }}>
          No visits on the calendar for <strong style={{ color: '#1e293b' }}>{agent}</strong> in {monthLabel}. Drag customers from the sidebar onto dates in Calendar.
        </div>
      );
    }

    const { from, to } = summaryPrintRange;
    const filtered = summaryFiltered;
    if (!filtered || filtered.totalVisits === 0) {
      return (
        <div className="summary-print-wrap" style={{ padding: '24px', color: '#64748b', flex: 1, lineHeight: 1.5 }}>
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
        <div className="summary-print-page-title">
          Schedule summary · {summaryRangeTitle}
          {agent ? ` · ${agent}` : ''}
        </div>
        <div className="summary-print-kpi-row" style={{ display: 'flex', gap: '16px', flexWrap: 'wrap' }}>
          {[
            { label: 'Total visits', value: totalVisits },
            { label: 'Customers', value: uniqueCustomers },
            { label: 'Days with visits', value: daysScheduled },
          ].map((kpi) => (
            <div
              key={kpi.label}
              style={{
                flex: '1 1 120px',
                minWidth: '100px',
                background: '#f8fafc',
                border: '1px solid #e2e8f0',
                borderRadius: '8px',
                padding: '14px 16px',
              }}
            >
              <div style={{ fontSize: '11px', textTransform: 'uppercase', color: '#64748b', marginBottom: '4px' }}>{kpi.label}</div>
              <div style={{ fontSize: '22px', fontWeight: '700', color: '#1e293b' }}>{kpi.value}</div>
            </div>
          ))}
        </div>

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
              {dateKeys.flatMap((dateStr) =>
                byDate[dateStr].map((row, idx) => {
                  const d = row.weekday ?? ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][new Date(dateStr).getDay()];
                  const dayNum = Number(dateStr.slice(-2));
                  const showingDate = idx === 0;
                  return (
                    <tr key={`${dateStr}-${row.acc}-${idx}`} style={{ background: idx === 0 && showingDate ? '#fff' : 'white' }}>
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
      </div>
    );
  };

  const parseWorkbookFirstSheet = (workbook) => {
    const name = workbook.SheetNames?.[0];
    if (!name || !workbook.Sheets[name]) return [];
    return XLSX.utils.sheet_to_json(workbook.Sheets[name]);
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

  const handleFile = (e) => {
    const input = e.target;
    const file = input?.files?.[0];
    if (!file) {
      setImportFileName('');
      return;
    }

    const ext = (file.name.split('.').pop() ?? '').toLowerCase();
    const isCsv = ext === 'csv';

    const reader = new FileReader();
    reader.onload = () => {
      try {
        let workbook;
        if (isCsv) {
          let text = String(reader.result ?? '');
          if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
          workbook = XLSX.read(text, { type: 'string', raw: false });
        } else {
          workbook = XLSX.read(new Uint8Array(reader.result), { type: 'array' });
        }
        const rows = parseWorkbookFirstSheet(workbook);
        const mapped = mapImportedRowsToCustomers(rows);
        const dupAccs = findDuplicateImportAccs(mapped);
        if (dupAccs.length > 0) {
          const sample = dupAccs.slice(0, 12).join(', ');
          const more = dupAccs.length > 12 ? `\n… and ${dupAccs.length - 12} more.` : '';
          window.alert(
            `Import cancelled: each ACC (account number) must be unique in the spreadsheet.\n\nDuplicate ACC values (${dupAccs.length}): ${sample}${more}\n\nFix the file and import again.`,
          );
          if (input && typeof input.value === 'string') input.value = '';
          return;
        }
        let droppedAfterMerge = 0;
        setCustomers(mapped);
        setImportFileName(file.name ?? '');
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
        if (droppedAfterMerge > 0) {
          window.alert(
            `Imported ${mapped.length} customers.\n\n` +
              `Calendar visits were kept for every account still listed in this file (dates and visit # unchanged; names refreshed from the sheet).\n\n` +
              `${droppedAfterMerge} visit${droppedAfterMerge === 1 ? '' : 's'} had to be dropped because ${droppedAfterMerge === 1 ? 'that account is' : 'those accounts are'} no longer in the spreadsheet.`,
          );
        }
        if (input && typeof input.value === 'string') input.value = '';
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

  const clearStoredImportData = () => {
    if (
      !window.confirm(
        hasElectronDataBridge()
          ? 'Remove all saved customers and calendar visits from this machine (data file and browser cache)?'
          : 'Remove saved customers, import file name, and calendar visits from this browser?',
      )
    ) {
      return;
    }
    try {
      if (typeof localStorage !== 'undefined') localStorage.removeItem(IMPORT_STORE_KEY);
    } catch {
      window.alert('Could not access browser storage.');
      return;
    }
    setCustomers([]);
    setImportFileName('');
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
    setSidebarSelectedRowKeys(new Set());
    setFocusedDay(null);
    const inp = document.getElementById('route-data-import');
    if (inp && typeof inp.value === 'string') inp.value = '';
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

        const daySched = [...(agentSched[dateStr] || [])];
        if (daySched.some((c) => c.acc === data.acc)) return prev;

        let visitRound;
        if (data.oldDate) {
          visitRound = normalizeVisitRound(data.visitRound ?? 1);
        } else {
          visitRound = countAcc() + 1;
        }

        agentSched[dateStr] = [...daySched, { acc: data.acc, name: data.name, visitRound }];
        return { ...prev, [agent]: agentSched };
      }

      let changed = false;
      for (const data of payloads) {
        if (data.oldDate) continue;
        const daySched = [...(agentSched[dateStr] || [])];
        if (daySched.some((c) => c.acc === data.acc)) continue;
        const visitRound = countAccFor(data.acc) + 1;
        agentSched[dateStr] = [...daySched, { acc: data.acc, name: data.name, visitRound }];
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

  const renderCalendar = () => (
    <div className="cal-calendar-body" style={{ width: '100%', minWidth: 0, flex: '1 1 auto', minHeight: 0 }}>
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(7, 1fr)',
          alignItems: 'stretch',
          width: 'max(100%, 800px)',
          boxSizing: 'border-box',
        }}
      >
        {calendarWeeks.map((weekCells, wi) => (
          <CalWeekRow
            key={`week-${wi}-${year}-${month}`}
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
          />
        ))}
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
          html, body, #root {
            height: auto !important;
            max-height: none !important;
            overflow: visible !important;
            background: white !important;
          }
          body {
            color: #0f172a !important;
          }
          #root {
            display: block !important;
            width: 100% !important;
            min-height: 0 !important;
            flex: none !important;
          }
          @page { margin: 12mm; size: auto; }
          .app-root-shell {
            display: block !important;
            width: 100% !important;
            height: auto !important;
            min-height: 0 !important;
            max-height: none !important;
            overflow: visible !important;
            background: white !important;
          }
          .main-column {
            display: block !important;
            height: auto !important;
            overflow: visible !important;
            background: white !important;
          }
          .main-scroll-region {
            padding: 12px 16px !important;
            overflow: visible !important;
            height: auto !important;
            max-height: none !important;
            flex: none !important;
            min-height: 0 !important;
          }
          .content-white-panel {
            overflow: visible !important;
            height: auto !important;
            max-height: none !important;
            flex: none !important;
            border-left: none !important;
            border-right: none !important;
          }
          .cal-calendar-body > div {
            width: 100% !important;
            min-width: 0 !important;
          }
          .summary-print-wrap {
            overflow: visible !important;
            max-height: none !important;
            padding: 12px 0 !important;
            display: block !important;
            flex: none !important;
          }
          .summary-print-page-title {
            display: block !important;
            font-size: 14pt !important;
            font-weight: 700 !important;
            color: #0f172a !important;
            margin: 0 0 14px 0 !important;
            padding: 0 !important;
          }
          .summary-print-table-host {
            overflow: visible !important;
            border-radius: 0 !important;
          }
          .summary-print-kpi-row > * {
            break-inside: avoid;
            page-break-inside: avoid;
          }
          .summary-print-table {
            font-size: 10pt !important;
          }
          .summary-print-table thead {
            display: table-header-group;
          }
          .summary-print-table tr {
            break-inside: avoid;
            page-break-inside: avoid;
          }
          .schedule-strip-print {
            border-radius: 0 !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          .cal-day-tasks {
            max-height: none !important;
            overflow: visible !important;
          }
        }
        .summary-print-page-title {
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
        .sidebar-scroll::-webkit-scrollbar { width: 6px; }
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
      `}</style>
      
      {/* Sidebar: top controls fixed, customer list fills remaining viewport and scrolls */}
      <aside className="no-print" style={{ width: '268px', minWidth: '268px', minHeight: 0, alignSelf: 'stretch', background: 'linear-gradient(180deg, #0f172a 0%, #0c1222 100%)', color: 'white', padding: '10px 8px', display: 'flex', flexDirection: 'column', boxSizing: 'border-box', borderRight: '1px solid rgba(148, 163, 184, 0.12)', overflow: 'hidden' }}>
        <div style={{ flexShrink: 0 }}>
        <h2 style={{ color: '#60a5fa', textAlign: 'center', margin: '0 0 2px', fontSize: '0.9rem', fontWeight: 800, letterSpacing: '0.05em' }}>📍 ROUTE PRO</h2>
        <p style={{ textAlign: 'center', fontSize: '8px', color: 'rgba(148, 163, 184, 0.85)', margin: '0 0 10px' }}>{'Import · schedule visits'}</p>

        <div style={{ marginBottom: '9px' }}>
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
                gap: '8px',
                width: '100%',
                padding: '7px 8px',
                borderRadius: '7px',
                cursor: 'pointer',
                boxSizing: 'border-box',
                fontSize: '10px',
                fontWeight: 600,
                color: 'white',
                border: '1px solid rgba(255,255,255,0.14)',
                background: 'linear-gradient(180deg, #3b82f6 0%, #2563eb 45%, #1d4ed8 100%)',
                boxShadow: '0 2px 10px rgba(37, 99, 235, 0.4)',
              }}
            >
              <Upload size={14} strokeWidth={2.25} aria-hidden />
              Choose spreadsheet
            </label>
          </div>
          <div
            style={{
              marginTop: '6px',
              fontSize: '8px',
              color: importFileName ? 'rgba(226, 232, 240, 0.95)' : 'rgba(148, 163, 184, 0.75)',
              lineHeight: 1.35,
              wordBreak: 'break-all',
              minHeight: '14px',
            }}
            title={importFileName || undefined}
          >
            {importFileName ? importFileName : 'No file chosen · .xlsx / .xls / .csv'}
          </div>
          <button
            type="button"
            className="no-print"
            onClick={clearStoredImportData}
            style={{
              marginTop: '8px',
              width: '100%',
              padding: '6px 8px',
              borderRadius: '6px',
              border: '1px solid rgba(248, 113, 113, 0.4)',
              background: 'rgba(127, 29, 29, 0.38)',
              color: '#fecaca',
              fontSize: '9px',
              fontWeight: 600,
              cursor: 'pointer',
            }}
          >
            Clear saved import (this browser)
          </button>
        </div>

        <div style={{ marginBottom: '8px' }}>
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

        <div style={{ display: 'flex', gap: '6px', marginBottom: '9px', alignItems: 'stretch', minWidth: 0 }}>
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
              <option value="">{agent ? '— All areas —' : '— Agent first —'}</option>
              {filteredAreas.map(a => <option key={a} value={a}>{a}</option>)}
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

        <div style={{ display: 'flex', gap: '6px', marginBottom: '9px', alignItems: 'stretch', minWidth: 0 }}>
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
              <option value="">{agent ? '— All channels —' : '— Agent first —'}</option>
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
              <option value="">{agent ? '— All categories —' : '— Agent first —'}</option>
              {sidebarCategoryOptions.map((cat) => (
                <option key={cat} value={cat}>{cat}</option>
              ))}
            </select>
          </div>
        </div>

        <div style={{ marginBottom: '9px', minWidth: 0 }}>
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
            placeholder={agent ? 'Search by name or account #…' : '— Select agent first —'}
            disabled={!agent}
            onChange={(e) => setSidebarNameQuery(e.target.value)}
            style={{
              width: '100%',
              padding: '6px 8px',
              borderRadius: '6px',
              border: `1px solid ${agent ? 'rgba(148, 163, 184, 0.35)' : 'rgba(100,116,139,0.35)'}`,
              background: agent ? '#1e293b' : '#0f172a',
              color: agent ? '#f8fafc' : 'rgba(148,163,184,0.45)',
              fontSize: '10px',
              fontWeight: 500,
              outline: 'none',
              boxSizing: 'border-box',
              opacity: agent ? 1 : 0.65,
              boxShadow: agent ? '0 1px 3px rgba(0,0,0,0.35)' : 'none',
            }}
          />
          {agent && sidebarNameQuery.trim().length > 0 && (
            <div style={{ marginTop: '3px', fontSize: '8px', color: 'rgba(148, 163, 184, 0.85)', lineHeight: 1.35 }}>
              {sidebarDisplayedCustomers.length} match{sidebarDisplayedCustomers.length === 1 ? '' : 'es'} · Visit count filter ignored while typing (Area / Channel / Category still apply)
            </div>
          )}
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
            padding: '6px',
            borderRadius: '6px',
          }}
        >
          {!agent ? (
            <div style={{ fontSize: '10px', color: 'rgba(148, 163, 184, 0.9)', textAlign: 'center', padding: '12px 6px' }}>
              Select a sales agent to list customers.
            </div>
          ) : sidebarListedCustomers.length === 0 ? (
            sidebarBaseCustomers.length === 0 ? (
              <div style={{ fontSize: '10px', color: 'rgba(148, 163, 184, 0.95)', textAlign: 'center', padding: '12px 6px', lineHeight: 1.5 }}>
                No customers match this agent / area / channel / category. Import data or adjust filters.
              </div>
            ) : (
              <div style={{ fontSize: '10px', color: 'rgba(148, 163, 184, 0.95)', textAlign: 'center', padding: '12px 6px', lineHeight: 1.5 }}>
                No customers match this visit filter.
                <div style={{ fontSize: '9px', opacity: 0.8, marginTop: '6px' }}>{sidebarBaseCustomers.length} match agent · area · channel · category — try « Any count ».</div>
              </div>
            )
          ) : sidebarDisplayedCustomers.length === 0 ? (
            <div style={{ fontSize: '10px', color: 'rgba(148, 163, 184, 0.95)', textAlign: 'center', padding: '12px 6px', lineHeight: 1.5 }}>
              No customers match <strong style={{ color: '#e2e8f0' }}>「{sidebarNameQuery.trim()}」</strong>.
              <div style={{ fontSize: '9px', opacity: 0.8, marginTop: '6px' }}>
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
                  gap: '8px',
                  flexWrap: 'wrap',
                  marginBottom: '6px',
                  padding: '3px 5px',
                  borderRadius: '6px',
                  background: 'rgba(15, 23, 42, 0.65)',
                  border: '1px solid rgba(148, 163, 184, 0.2)',
                }}
              >
                <label style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '8px', fontWeight: 600, color: 'rgba(248,250,252,0.92)', cursor: 'pointer', userSelect: 'none' }}>
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
                    style={{ width: '11px', height: '11px', cursor: 'pointer', accentColor: '#38bdf8' }}
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
                    fontSize: '8px',
                    fontWeight: 600,
                    padding: '2px 7px',
                    borderRadius: '5px',
                    border: '1px solid rgba(148, 163, 184, 0.35)',
                    background: sidebarSelectedRowKeys.size === 0 ? 'rgba(51, 65, 85, 0.5)' : 'rgba(51, 65, 85, 0.95)',
                    color: sidebarSelectedRowKeys.size === 0 ? 'rgba(148, 163, 184, 0.5)' : '#e2e8f0',
                    cursor: sidebarSelectedRowKeys.size === 0 ? 'not-allowed' : 'pointer',
                  }}
                >
                  Clear selection
                </button>
                {sidebarSelectedRowKeys.size > 0 && (
                  <span style={{ fontSize: '8px', color: 'rgba(148, 163, 184, 0.95)', marginLeft: 'auto' }}>
                    {sidebarSelectedRowKeys.size} selected — drag one checked row; selection clears after drop
                  </span>
                )}
              </div>
              {sidebarDisplayedCustomers.map((c, idx) => {
                const rk = sidebarRowListKey(c, idx);
                const total = visitCountsTotal[c.acc] ?? 0;
                const bucket = sidebarColorBucket(total);
                const pal = paletteForTotalVisits(total);
                const badge = total > 0 ? String(total) : null;
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
                      gap: '4px',
                      background: pal.sidebarBg,
                      padding: '6px 6px 6px 8px',
                      borderRadius: '5px',
                      marginBottom: '6px',
                      borderLeft: `4px solid ${isSelected ? '#38bdf8' : pal.sidebarBorder}`,
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
                        width: '22px',
                        minWidth: '22px',
                        display: 'flex',
                        justifyContent: 'center',
                        alignItems: 'flex-start',
                        paddingTop: '2px',
                      }}
                    >
                      {badge != null && (
                        <span
                          title={total >= 6 ? `已排 ${total} 次（颜色与第 6 档相同）` : `已排 ${total} 次`}
                          style={{
                            minWidth: '18px',
                            height: '18px',
                            padding: '0 4px',
                            borderRadius: '999px',
                            fontSize: '9px',
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
                      <div style={{ fontSize: '8px', color: bucket > 0 ? 'rgba(248, 250, 252, 0.85)' : '#60a5fa' }}>{c.acc}{total > 0 ? ` · ${total}×` : ''}</div>
                      <div style={{ fontSize: '10px', fontWeight: 700, color: '#f8fafc', lineHeight: 1.22 }}>{c.name}</div>
                    </div>
                    <label
                      style={{
                        display: 'flex',
                        alignItems: 'flex-start',
                        padding: '2px 0 0 2px',
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
                        style={{ width: '13px', height: '13px', marginTop: '2px', cursor: 'pointer', accentColor: '#38bdf8' }}
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
        <div className="no-print" style={{ padding: '15px 25px', display: 'flex', justifyContent: 'space-between', background: 'white', borderBottom: '1px solid #e2e8f0' }}>
          <div style={{ display: 'flex', gap: '10px' }}>
            <button onClick={()=>setViewMode('calendar')} style={{ padding: '8px 15px', borderRadius: '5px', cursor: 'pointer', background: viewMode==='calendar'?'#1e293b':'white', color: viewMode==='calendar'?'white':'#1e293b' }}>Calendar</button>
            <button onClick={()=>setViewMode('summary')} style={{ padding: '8px 15px', borderRadius: '5px', cursor: 'pointer', background: viewMode==='summary'?'#1e293b':'white', color: viewMode==='summary'?'white':'#1e293b' }}>Summary</button>
          </div>
          <button
            type="button"
            onClick={() => window.print()}
            title="Open Summary, set the From/To day range in the toolbar, then print or save as PDF."
            style={{ background: '#2563eb', color: 'white', border: 'none', padding: '10px 25px', borderRadius: '8px', cursor: 'pointer', fontWeight:'bold' }}
          >
            PRINT / PDF
          </button>
        </div>

        <div className="main-scroll-region" style={{ flex: 1, minHeight: 0, minWidth: 0, overflowY: 'auto', overflowX: 'hidden', padding: '14px 16px', display: 'flex', flexDirection: 'column' }}>
          <div className="schedule-strip-print" style={{ flexShrink: 0, background: '#1e293b', color: 'white', padding: '15px 25px', borderRadius: '12px 12px 0 0', fontWeight: 'bold', display:'flex', justifyContent:'space-between', alignItems: 'center', gap: '12px', flexWrap: 'wrap', width: '100%', boxSizing: 'border-box' }}>
            <span style={{ minWidth: 0 }}>AGENT: {agent || 'NOT SELECTED'}</span>
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
                      title="When Sync is on, Area / Channel / Category and the Find search text mirror the sidebar (and vice versa). Sidebar Visits count & calendar chip # filters stay separate."
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
                          onChange={(e) => {
                            const on = e.target.checked;
                            if (on && agent) {
                              setCalendarArea(sidebarArea);
                              setCalendarChannelFilter(sidebarChannelFilter);
                              setCalendarCategoryFilter(sidebarCategoryFilter);
                              setCalendarNameQuery(sidebarNameQuery);
                            }
                            setSyncSidebarCalendarFilters(on);
                          }}
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
                        disabled={!agent}
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
                        title={`Show only bars whose left badge matches (independent from sidebar “total visits”). “All visits” shows every bar.`}
                        style={{ ...selectCalendarToolbarStyle, width: '100%' }}
                        disabled={!agent}
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
                        disabled={!agent}
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
                        disabled={!agent}
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
                        disabled={!agent}
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
    </div>
  );
}

export default App;