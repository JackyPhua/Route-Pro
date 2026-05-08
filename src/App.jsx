import { useState, useMemo, useEffect, useRef, useLayoutEffect } from 'react';
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
}) {
  const rowHostRef = useRef(null);
  const [rowFontPx, setRowFontPx] = useState(CAL_CHIP_FONT_MAX);

  const enriched = useMemo(
    () =>
      weekCells.map((slot) => {
        if (slot.kind !== 'day') return slot;
        const dateStr = `${year}-${monthIsoPad}-${String(slot.dayNum).padStart(2, '0')}`;
        return {
          ...slot,
          dateStr,
          isSun: new Date(year, Number(monthIsoPad) - 1, slot.dayNum).getDay() === 0,
          tasks: (agent && schedules[agent] && schedules[agent][dateStr]) || [],
        };
      }),
    [weekCells, monthIsoPad, year, agent, schedules],
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

/** 表格导入：空白 → 显示为 —（与下拉筛选用同一套值） */
function normalizeImportField(raw) {
  if (raw === undefined || raw === null || String(raw).trim() === '') return '—';
  return String(raw).trim();
}

/** 导入的顾客列表存为浏览器内 JSON（localStorage），每次新导入整份覆盖 */
const IMPORT_STORE_KEY = 'route-planner:v1:import-data';

function loadImportStore() {
  if (typeof localStorage === 'undefined') return { customers: [], importFileName: '' };
  try {
    const raw = localStorage.getItem(IMPORT_STORE_KEY);
    if (!raw) return { customers: [], importFileName: '' };
    const o = JSON.parse(raw);
    if (Array.isArray(o)) return { customers: o, importFileName: '' };
    if (o && Array.isArray(o.customers))
      return { customers: o.customers, importFileName: typeof o.importFileName === 'string' ? o.importFileName : '' };
  } catch {
    /* ignore corrupted store */
  }
  return { customers: [], importFileName: '' };
}

function saveImportStore(customersList, sourceFileName) {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(
      IMPORT_STORE_KEY,
      JSON.stringify({
        v: 1,
        savedAt: new Date().toISOString(),
        importFileName: sourceFileName || '',
        customers: customersList,
      }),
    );
  } catch (err) {
    console.warn('route-planner: persist import JSON failed', err);
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

function App() {
  const [customers, setCustomers] = useState(() => loadImportStore().customers);
  const [agent, setAgent] = useState('');
  const [area, setArea] = useState('');
  const [schedules, setSchedules] = useState({});
  const [year, setYear] = useState(2026);
  /** 0-indexed month (aligned with Date — 4 = May) */
  const [month, setMonth] = useState(4);
  /** Highlights a calendar cell (1 … days in month); only used in Calendar view */
  const [focusedDay, setFocusedDay] = useState(null);
  const [viewMode, setViewMode] = useState('calendar');
  const [importFileName, setImportFileName] = useState(() => loadImportStore().importFileName);
  const [visitCountFilter, setVisitCountFilter] = useState('all');
  /** 侧栏 Channel / Category（与表格列 CHANNEL、CATEGORY 对应；空档为 —） */
  const [channelFilter, setChannelFilter] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('');
  /** 侧栏店名 / 账号即时过滤（不区分大小写） */
  const [sidebarNameQuery, setSidebarNameQuery] = useState('');
  /** 侧栏多选：勾选后从任一已选行拖到日历可一次排入多天客户 */
  const [sidebarSelectedAccs, setSidebarSelectedAccs] = useState(() => new Set());
  /** Summary / print: day-of-month range within the selected Year–Month */
  const [summaryDayFrom, setSummaryDayFrom] = useState(1);
  const [summaryDayTo, setSummaryDayTo] = useState(31);

  const filteredAreas = useMemo(() => {
    if (!agent) return [];
    return [...new Set(customers.filter(c => c.agent === agent).map(c => c.area))].sort();
  }, [agent, customers]);

  const sidebarChannelOptions = useMemo(() => {
    if (!agent) return [];
    const set = new Set(
      customers
        .filter((c) => c.agent === agent && (!area || c.area === area))
        .map((c) => String(c.channel ?? '—')),
    );
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
  }, [agent, area, customers]);

  const sidebarCategoryOptions = useMemo(() => {
    if (!agent) return [];
    const set = new Set(
      customers
        .filter(
          (c) =>
            c.agent === agent &&
            (!area || c.area === area) &&
            (!channelFilter || String(c.channel ?? '—') === channelFilter),
        )
        .map((c) => String(c.category ?? '—')),
    );
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
  }, [agent, area, channelFilter, customers]);

  const visitCounts = useMemo(() => {
    const counts = {};
    if (agent && schedules[agent]) {
      Object.values(schedules[agent]).forEach((dayList) => {
        dayList.forEach((c) => { counts[c.acc] = (counts[c.acc] || 0) + 1; });
      });
    }
    return counts;
  }, [agent, schedules]);

  const sidebarBaseCustomers = useMemo(() => {
    if (!agent) return [];
    return customers.filter(
      (c) =>
        c.agent === agent &&
        (!area || c.area === area) &&
        (!channelFilter || String(c.channel ?? '—') === channelFilter) &&
        (!categoryFilter || String(c.category ?? '—') === categoryFilter),
    );
  }, [agent, area, channelFilter, categoryFilter, customers]);

  const sidebarListedCustomers = useMemo(
    () => sidebarBaseCustomers.filter((c) => matchesVisitScheduleFilter(visitCounts[c.acc], visitCountFilter)),
    [sidebarBaseCustomers, visitCounts, visitCountFilter],
  );

  const sidebarDisplayedCustomers = useMemo(() => {
    const raw = sidebarNameQuery.trim().toLowerCase();
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

    let list =
      !raw
        ? [...sidebarListedCustomers]
        : sidebarListedCustomers.filter((c) => {
            const nm = String(c.name ?? '').toLowerCase();
            const ac = String(c.acc ?? '').toLowerCase();
            return nm.includes(raw) || ac.includes(raw);
          });

    list.sort(byNameThenAcc);
    return list;
  }, [sidebarListedCustomers, sidebarNameQuery]);

  useEffect(() => {
    setSidebarSelectedAccs(new Set());
    setSidebarNameQuery('');
    setChannelFilter('');
    setCategoryFilter('');
  }, [agent]);

  useEffect(() => {
    setChannelFilter('');
    setCategoryFilter('');
  }, [area]);

  useEffect(() => {
    setCategoryFilter('');
  }, [channelFilter]);

  const sidebarSelectAllRef = useRef(null);
  useEffect(() => {
    const el = sidebarSelectAllRef.current;
    if (!el) return;
    const all =
      sidebarDisplayedCustomers.length > 0 &&
      sidebarDisplayedCustomers.every((c) => sidebarSelectedAccs.has(c.acc));
    const some = sidebarDisplayedCustomers.some((c) => sidebarSelectedAccs.has(c.acc));
    el.checked = all;
    el.indeterminate = some && !all;
  }, [sidebarDisplayedCustomers, sidebarSelectedAccs]);

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

  useEffect(() => {
    const dm = daysInMonth;
    setSummaryDayFrom((f) => Math.min(Math.max(1, f), dm));
    setSummaryDayTo((t) => Math.min(Math.max(1, t), dm));
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
      <div className="summary-print-wrap" style={{ flex: 1, overflow: 'auto', width: '100%', boxSizing: 'border-box', padding: '20px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
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
    data.map((r) => {
      const ks = Object.keys(r);
      const f = (ts) => ks.find((k) => ts.some((t) => k.toUpperCase().includes(t.toUpperCase())));
      return {
        acc: r[f(['ACC NO', 'ACCOUNT'])] || 'N/A',
        name: r[f(['DEBTOR NAME', 'NAME'])] || 'N/A',
        area: r[f(['AREA'])] || 'OTHER',
        agent: r[f(['SALES AGENT', 'AGENT'])] || 'UNASSIGNED',
        channel: normalizeImportField(r[f(['CHANNEL'])]),
        category: normalizeImportField(r[f(['CATEGORY', 'PRODUCT CATEGORY'])]),
      };
    });

  const handleFile = (e) => {
    const input = e.target;
    const file = input?.files?.[0];
    setImportFileName(file?.name ?? '');
    if (!file) return;

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
        setCustomers(mapped);
        saveImportStore(mapped, file.name ?? '');
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

    if (!isCalendarMove) setSidebarSelectedAccs(new Set());
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
          <div style={{ marginTop: '4px', fontSize: '8px', opacity: 0.55, lineHeight: 1.35, color: 'rgba(148, 163, 184, 0.95)' }}>
            Imported customers persist in this browser as JSON (localStorage); each new import replaces the previous save.
            Optional columns for sidebar filters: CHANNEL, CATEGORY (or PRODUCT CATEGORY).
          </div>
        </div>

        <div style={{ marginBottom: '8px' }}>
          <label style={sidebarSectionLabel} htmlFor="sidebar-agent">2. Sales agent</label>
          <select
            id="sidebar-agent"
            className="sidebar-select"
            value={agent}
            onChange={(e) => { setAgent(e.target.value); setArea(''); setVisitCountFilter('all'); }}
            style={{ ...selectSidebarStyle, marginBottom: 0 }}
          >
            <option value="">— Select agent —</option>
            {[...new Set(customers.map(c => c.agent))].map(a => <option key={a} value={a}>{a}</option>)}
          </select>
        </div>

        <div style={{ display: 'flex', gap: '6px', marginBottom: '9px', alignItems: 'stretch', minWidth: 0 }}>
          <div style={{ flex: '1 1 50%', minWidth: 0, display: 'flex', flexDirection: 'column' }}>
            <label style={sidebarSectionLabel} htmlFor="sidebar-area">3. Area</label>
            <select
              id="sidebar-area"
              className="sidebar-select"
              value={area}
              onChange={(e) => setArea(e.target.value)}
              style={{ ...selectSidebarStyle, marginBottom: 0 }}
              disabled={!agent}
            >
              <option value="">{agent ? '— All areas —' : '— Agent first —'}</option>
              {filteredAreas.map(a => <option key={a} value={a}>{a}</option>)}
            </select>
          </div>
          <div style={{ flex: '1 1 50%', minWidth: 0, display: 'flex', flexDirection: 'column' }}>
            <label style={sidebarSectionLabel} htmlFor="sidebar-visit-filter">4. Visits</label>
            <select
              id="sidebar-visit-filter"
              className="sidebar-select"
              value={visitCountFilter}
              onChange={(e) => setVisitCountFilter(e.target.value)}
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
            <label style={sidebarSectionLabel} htmlFor="sidebar-channel">5. Channel</label>
            <select
              id="sidebar-channel"
              className="sidebar-select"
              value={channelFilter}
              onChange={(e) => setChannelFilter(e.target.value)}
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
            <label style={sidebarSectionLabel} htmlFor="sidebar-category">6. Category</label>
            <select
              id="sidebar-category"
              className="sidebar-select"
              value={categoryFilter}
              onChange={(e) => setCategoryFilter(e.target.value)}
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
            7. Find customer
          </label>
          <input
            id="sidebar-name-search"
            type="search"
            autoCapitalize="off"
            spellCheck={false}
            aria-label="Filter customers by name or account"
            value={sidebarNameQuery}
            placeholder={agent ? 'Type name or account…' : '— Select agent first —'}
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
              Showing {sidebarDisplayedCustomers.length} of {sidebarListedCustomers.length} in list
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
                      setSidebarSelectedAccs((prev) => {
                        const visible = sidebarDisplayedCustomers;
                        const allOn =
                          visible.length > 0 && visible.every((row) => prev.has(row.acc));
                        if (allOn) return new Set();
                        return new Set(visible.map((row) => row.acc));
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
                  onClick={() => setSidebarSelectedAccs(new Set())}
                  disabled={sidebarSelectedAccs.size === 0}
                  style={{
                    fontSize: '8px',
                    fontWeight: 600,
                    padding: '2px 7px',
                    borderRadius: '5px',
                    border: '1px solid rgba(148, 163, 184, 0.35)',
                    background: sidebarSelectedAccs.size === 0 ? 'rgba(51, 65, 85, 0.5)' : 'rgba(51, 65, 85, 0.95)',
                    color: sidebarSelectedAccs.size === 0 ? 'rgba(148, 163, 184, 0.5)' : '#e2e8f0',
                    cursor: sidebarSelectedAccs.size === 0 ? 'not-allowed' : 'pointer',
                  }}
                >
                  Clear selection
                </button>
                {sidebarSelectedAccs.size > 0 && (
                  <span style={{ fontSize: '8px', color: 'rgba(148, 163, 184, 0.95)', marginLeft: 'auto' }}>
                    {sidebarSelectedAccs.size} selected — drag one checked row; selection clears after drop
                  </span>
                )}
              </div>
              {sidebarDisplayedCustomers.map((c) => {
                const total = visitCounts[c.acc] ?? 0;
                const bucket = sidebarColorBucket(total);
                const pal = paletteForTotalVisits(total);
                const badge = total > 0 ? String(total) : null;
                const isSelected = sidebarSelectedAccs.has(c.acc);
                const bulkOnList = sidebarDisplayedCustomers.filter((row) => sidebarSelectedAccs.has(row.acc));
                const useBulk = isSelected && bulkOnList.length >= 2;
                return (
                  <div
                    key={c.acc}
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
                      gap: '0',
                      background: pal.sidebarBg,
                      padding: '6px 30px 6px 4px',
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
                    <label
                      style={{
                        display: 'flex',
                        alignItems: 'flex-start',
                        padding: '2px 4px 0 2px',
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
                          setSidebarSelectedAccs((prev) => {
                            const next = new Set(prev);
                            if (next.has(c.acc)) next.delete(c.acc);
                            else next.add(c.acc);
                            return next;
                          });
                        }}
                        style={{ width: '13px', height: '13px', marginTop: '2px', cursor: 'pointer', accentColor: '#38bdf8' }}
                      />
                    </label>
                    <div style={{ flex: 1, minWidth: 0, paddingLeft: '4px' }}>
                      {badge != null && (
                        <span
                          title={total >= 6 ? `已排 ${total} 次（颜色与第 6 档相同）` : `已排 ${total} 次`}
                          style={{
                            position: 'absolute',
                            top: '5px',
                            right: '5px',
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
                      <div style={{ fontSize: '8px', color: bucket > 0 ? 'rgba(248, 250, 252, 0.85)' : '#60a5fa' }}>{c.acc}{total > 0 ? ` · ${total}×` : ''}</div>
                      <div style={{ fontSize: '10px', fontWeight: 700, color: '#f8fafc', lineHeight: 1.22 }}>{c.name}</div>
                    </div>
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
          <div className="schedule-strip-print" style={{ background: '#1e293b', color: 'white', padding: '15px 25px', borderRadius: '12px 12px 0 0', fontWeight: 'bold', display:'flex', justifyContent:'space-between', alignItems: 'center', gap: '12px', flexWrap: 'wrap', width: '100%', boxSizing: 'border-box' }}>
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
            {viewMode === 'calendar' ? renderCalendar() : renderSummary()}
          </div>
        </div>
      </main>
    </div>
  );
}

export default App;