/**
 * Calendar yellow-highlight sources for Malaysian / Sabah-relevant holidays.
 *
 * Fixed Gregorian dates repeat yearly. Islamic & lunar dates change yearly — keys in
 * SABAH_HOLIDAY_EXTRA_DATES MUST be refreshed when Jabatan/state gazettes publish (~yearly).
 * Good Friday dates for Sabah & Sarawak are listed per year below (Western Easter-derived).
 */

export const SABAH_ANNUAL_FIXED_HOLIDAYS = [
  { month: 1, day: 1, name: "New Year's Day" },
  { month: 2, day: 1, name: 'Federal Territory Day (KL / Labuan / Putrajaya)' },
  { month: 5, day: 1, name: 'Labour Day' },
  { month: 5, day: 30, name: 'Kaamatan (Sabah)' },
  { month: 5, day: 31, name: 'Kaamatan (Sabah)' },
  { month: 8, day: 31, name: 'Merdeka (National Day)' },
  { month: 9, day: 16, name: 'Malaysia Day' },
  { month: 12, day: 25, name: 'Christmas Day' },
];

/** yyyy → [month, day] Good Friday Sabah & Sarawak */
export const GOOD_FRIDAY_SABAH_SARAWAK = {
  2023: [4, 7],
  2024: [3, 29],
  2025: [4, 18],
  2026: [4, 3],
  2027: [3, 26],
  2028: [4, 14],
  2029: [3, 30],
  2030: [4, 19],
  2031: [4, 11],
};

/** Nationwide CNY federal (first two days) yyyy → [[m,d],[m,d]] */
export const CHINESE_NEW_YEAR_MALAYSIA = {
  2025: [
    [1, 29],
    [1, 30],
  ],
  2026: [
    [2, 17],
    [2, 18],
  ],
  2027: [
    [2, 6],
    [2, 7],
  ],
  2028: [
    [1, 26],
    [1, 27],
  ],
};

/** Hari Raya Aidilfitri (2 national days typical) — VERIFY yearly */
export const HARI_RAYA_AIDILFITRI_MALAYSIA = {
  2025: [
    [3, 31],
    [4, 1],
  ],
  2026: [
    [3, 21],
    [3, 22],
  ],
  2027: [
    [3, 10],
    [3, 11],
  ],
  2028: [
    [2, 27],
    [2, 28],
  ],
};

/** Hari Raya Haji (+ day 2) — VERIFY yearly / state differences */
export const HARI_RAYA_HAJI_MALAYSIA = {
  2025: [
    [6, 7],
    [6, 8],
  ],
  2026: [
    [5, 27],
    [5, 28],
  ],
};

/**
 * Manual additions (Islamic tentative, Deepavali, Wesak, state extras).
 * Prefer editing this object once per gazette publication.
 */
export const SABAH_HOLIDAY_EXTRA_DATES = {
  '2027-02-08': 'Chinese New Year substitution (often in lieu) *',

  '2026-03-30': "Sabah Governor's Birthday *",
  '2026-05-31': 'Wesak Day *',
  '2026-06-01': "Yang di-Pertuan Agong's Birthday",

  '2026-06-17': 'Awal Muharram (Islamic New Year) *',
  '2026-08-25': "Prophet Muhammad's Birthday *",

  '2026-11-08': 'Deepavali *',
  '2026-11-09': 'Deepavali observed / substitute *',

  '2026-12-24': 'Christmas Eve (Sabah) *',

  /* 2025 pack (verify gazette) */
  '2025-05-12': 'Wesak Day *',
  '2025-06-02': "Yang di-Pertuan Agong's Birthday",
  '2025-06-27': 'Awal Muharram *',
  '2025-09-05': "Prophet Muhammad's Birthday *",
  '2025-10-21': 'Deepavali *',
  '2025-10-22': 'Deepavali observed / substitute *',
  '2025-12-24': 'Christmas Eve (Sabah) *',
  '2025-03-30': "Sabah Governor's Birthday *",

  /* 2027–2028 illustrative national tentatives (* confirm) */
  '2027-05-07': 'Wesak Day *',
  '2027-06-05': "Yang di-Pertuan Agong's Birthday",
  '2027-06-26': 'Awal Muharram *',
  '2027-08-16': "Prophet Muhammad's Birthday *",
  '2027-10-09': 'Deepavali *',
  '2027-03-29': "Sabah Governor's Birthday *",

  '2028-05-18': 'Wesak Day *',
  '2028-06-05': "Yang di-Pertuan Agong's Birthday",
  '2028-06-15': 'Awal Muharram *',
  '2028-08-04': "Prophet Muhammad's Birthday *",
  '2028-10-28': 'Deepavali *',
  '2028-03-30': "Sabah Governor's Birthday *",
};

function pad2(n) {
  return String(n).padStart(2, '0');
}

function mergeHolidayLabels(intoObj, yyyyMmDd, label) {
  const prev = intoObj[yyyyMmDd];
  if (prev == null || String(prev).trim() === '') {
    intoObj[yyyyMmDd] = label;
    return;
  }
  const p = String(prev);
  if (p.includes(label)) return;
  intoObj[yyyyMmDd] = `${p} · ${label}`;
}

function buildComputedOverrides(minYear = 2023, maxYear = 2032) {
  /** @type {Record<string, string>} */
  const out = {};

  for (let y = minYear; y <= maxYear; y++) {
    const gf = GOOD_FRIDAY_SABAH_SARAWAK[y];
    if (gf) mergeHolidayLabels(out, `${y}-${pad2(gf[0])}-${pad2(gf[1])}`, 'Good Friday (Sabah & Sarawak)');

    const cny = CHINESE_NEW_YEAR_MALAYSIA[y];
    if (Array.isArray(cny) && cny.length >= 2) {
      mergeHolidayLabels(out, `${y}-${pad2(cny[0][0])}-${pad2(cny[0][1])}`, 'Chinese New Year');
      mergeHolidayLabels(out, `${y}-${pad2(cny[1][0])}-${pad2(cny[1][1])}`, 'Chinese New Year (Day 2)');
    }

    const raya = HARI_RAYA_AIDILFITRI_MALAYSIA[y];
    if (Array.isArray(raya) && raya.length >= 2) {
      mergeHolidayLabels(out, `${y}-${pad2(raya[0][0])}-${pad2(raya[0][1])}`, 'Hari Raya Aidilfitri');
      mergeHolidayLabels(out, `${y}-${pad2(raya[1][0])}-${pad2(raya[1][1])}`, 'Hari Raya Aidilfitri (Day 2)');
    }

    const haji = HARI_RAYA_HAJI_MALAYSIA[y];
    if (Array.isArray(haji) && haji.length >= 2) {
      mergeHolidayLabels(out, `${y}-${pad2(haji[0][0])}-${pad2(haji[0][1])}`, 'Hari Raya Aidil Adha');
      mergeHolidayLabels(out, `${y}-${pad2(haji[1][0])}-${pad2(haji[1][1])}`, 'Hari Raya Aidil Adha (Day 2)');
    }
  }

  for (const [k, v] of Object.entries(SABAH_HOLIDAY_EXTRA_DATES)) {
    mergeHolidayLabels(out, k, String(v));
  }

  return out;
}

/** Pre-built map yyyy-mm-dd → label fragments */
export const COMPUTED_MALAYSIA_CALENDAR_LABELS_BY_DATE = buildComputedOverrides();

/**
 * Tooltip / yellow-cell title for yyyy-mm-dd
 */
export function sabahHolidayNameForCalendarDate(dateStr) {
  const trimmed = String(dateStr).trim();
  const computed = COMPUTED_MALAYSIA_CALENDAR_LABELS_BY_DATE[trimmed];
  if (typeof computed === 'string' && computed.trim() !== '') return computed.trim();

  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed);
  if (!m) return undefined;
  const mo = Number(m[2]);
  const day = Number(m[3]);
  const hit = SABAH_ANNUAL_FIXED_HOLIDAYS.find((h) => h.month === mo && h.day === day);
  return hit?.name;
}

function parseStrictLocalYyyyMmDd(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso).trim());
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
  return dt;
}

function formatLocalYyyyMmDd(dt) {
  return `${dt.getFullYear()}-${pad2(dt.getMonth() + 1)}-${pad2(dt.getDate())}`;
}

/**
 * Copy-to-next-month: advance day-by-day past Sundays and any date the calendar marks as a public holiday
 * (same yellow cells as `sabahHolidayNameForCalendarDate`). Stops at first “open” calendar day.
 */
export function advanceIsoPastSundaysAndPublicHolidays(dateStr, maxSteps = 31) {
  const dt = parseStrictLocalYyyyMmDd(dateStr);
  if (!dt) return String(dateStr);
  for (let i = 0; i < maxSteps; i++) {
    const iso = formatLocalYyyyMmDd(dt);
    const isSunday = dt.getDay() === 0;
    const hol = sabahHolidayNameForCalendarDate(iso);
    const isHoliday = hol != null && String(hol).trim() !== '';
    if (!isSunday && !isHoliday) return iso;
    dt.setDate(dt.getDate() + 1);
  }
  return formatLocalYyyyMmDd(dt);
}
