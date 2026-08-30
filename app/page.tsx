'use client';

import { ChangeEvent, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

type Mode = 'forecast' | 'live' | null;
type TeamMode = 'uniform' | 'individual';
type TeamMember = { id: string; name: string; dailyBooks: number };
type HistoricalRow = { id: string; date: string; quantity: number };
type LiveRow = { id: string; date: string; received: number; completed: number };
type ImportedCompany = { id: string; code: string; name: string; announcementDate: string | null; completionDate: string | null };
type DailyComparison = { date: string; availableCodes: string[]; completedCodes: string[]; matchedCodes: string[]; completedOnlyCodes: string[]; endingBacklogCodes: string[] };
type WorkbookSheet = { name: string; headers: string[]; rows: string[][] };
type ParsedWorkbookSheet = WorkbookSheet & {
  codeIndex: number;
  nameIndex: number;
  periodIndex: number;
  quarterIndex: number;
  announcementIndex: number;
  completionIndex: number;
};
type ForecastSettings = {
  defaultsVersion: number;
  historicalAnchor: string;
  currentAnchor: string;
  completionDate: string;
  people: number;
  speed: number;
  efficiency: number;
  teamMode: TeamMode;
  members: TeamMember[];
};
type LiveSettings = {
  defaultsVersion: number;
  seasonStartDate: string;
  anchorDate: string;
  completionDate: string;
  people: number;
  speed: number;
  efficiency: number;
  teamMode: TeamMode;
  members: TeamMember[];
  openingBacklog: number;
  expectedRemaining: number;
  importedCompanies: ImportedCompany[];
};

const DAY = 86_400_000;
const FORECAST_KEY = 'busy-season-forecast-v1';
const LIVE_KEY = 'busy-season-live-v1';

const initialForecastSettings: ForecastSettings = {
  defaultsVersion: 3,
  historicalAnchor: '2025-11-14',
  currentAnchor: '2026-11-16',
  completionDate: '2026-11-18',
  people: 4,
  speed: 10,
  efficiency: 85,
  teamMode: 'uniform',
  members: [],
};

const initialLiveSettings: LiveSettings = {
  defaultsVersion: 3,
  seasonStartDate: '2026-08-24',
  anchorDate: '2026-08-31',
  completionDate: '2026-09-04',
  people: 4,
  speed: 9,
  efficiency: 85,
  teamMode: 'uniform',
  members: [],
  openingBacklog: 0,
  expectedRemaining: 150,
  importedCompanies: [],
};

const forecastExample: HistoricalRow[] = [
  { id: 'f1', date: '2025-11-05', quantity: 30 },
  { id: 'f2', date: '2025-11-06', quantity: 48 },
  { id: 'f3', date: '2025-11-07', quantity: 70 },
  { id: 'f4', date: '2025-11-10', quantity: 105 },
  { id: 'f5', date: '2025-11-11', quantity: 135 },
  { id: 'f6', date: '2025-11-12', quantity: 165 },
  { id: 'f7', date: '2025-11-13', quantity: 210 },
  { id: 'f8', date: '2025-11-14', quantity: 285 },
];

const liveExample: LiveRow[] = [
  { id: 'l1', date: '2026-08-24', received: 45, completed: 38 },
  { id: 'l2', date: '2026-08-25', received: 62, completed: 55 },
  { id: 'l3', date: '2026-08-26', received: 88, completed: 64 },
  { id: 'l4', date: '2026-08-27', received: 105, completed: 81 },
  { id: 'l5', date: '2026-08-28', received: 110, completed: 92 },
];

function makeId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function parseDate(value: string) {
  const [year, month, day] = value.split('-').map(Number);
  return new Date(Date.UTC(year, month - 1, day));
}

function toISO(date: Date) {
  return date.toISOString().slice(0, 10);
}

function addDays(date: Date, days: number) {
  return new Date(date.getTime() + days * DAY);
}

function isWorkday(date: Date) {
  const day = date.getUTCDay();
  return day !== 0 && day !== 6;
}

function businessDayDiff(date: Date, anchor: Date) {
  if (date.getTime() === anchor.getTime()) return 0;
  const direction = date > anchor ? 1 : -1;
  let offset = 0;
  for (let cursor = anchor; cursor.getTime() !== date.getTime();) {
    cursor = addDays(cursor, direction);
    if (isWorkday(cursor)) offset += direction;
  }
  return offset;
}

function addBusinessDays(anchor: Date, offset: number) {
  if (offset === 0) return anchor;
  const direction = offset > 0 ? 1 : -1;
  let remaining = Math.abs(offset);
  let cursor = anchor;
  while (remaining > 0) {
    cursor = addDays(cursor, direction);
    if (isWorkday(cursor)) remaining -= 1;
  }
  return cursor;
}

function formatDate(value: string | Date) {
  const date = typeof value === 'string' ? parseDate(value) : value;
  return new Intl.DateTimeFormat('zh-TW', {
    month: 'numeric',
    day: 'numeric',
    weekday: 'short',
    timeZone: 'UTC',
  }).format(date);
}

function formatWeekday(value: string | Date) {
  const date = typeof value === 'string' ? parseDate(value) : value;
  return new Intl.DateTimeFormat('zh-TW', { weekday: 'long', timeZone: 'UTC' }).format(date);
}

function formatT(offset: number) {
  if (offset === 0) return 'T';
  return offset > 0 ? `T+${offset}` : `T${offset}`;
}

function compactNumber(value: number, digits = 0) {
  return new Intl.NumberFormat('zh-TW', { maximumFractionDigits: digits }).format(value);
}

function getTeamCapacity(settings: Pick<ForecastSettings, 'teamMode' | 'members' | 'people' | 'speed' | 'efficiency'>) {
  if (settings.teamMode === 'individual') {
    const dailyCapacity = settings.members.reduce((sum, member) => sum + Math.max(0, member.dailyBooks), 0);
    const teamCount = settings.members.length;
    return {
      teamCount,
      dailyCapacity,
      perPersonCapacity: teamCount > 0 ? dailyCapacity / teamCount : 0,
    };
  }
  const teamCount = Math.max(0, settings.people);
  const perPersonCapacity = Math.max(0, settings.speed) * (settings.efficiency / 100);
  return { teamCount, perPersonCapacity, dailyCapacity: teamCount * perPersonCapacity };
}

function countWorkdays(start: Date, end: Date) {
  if (start > end) return 0;
  let count = 0;
  for (let cursor = start; cursor <= end; cursor = addDays(cursor, 1)) {
    if (isWorkday(cursor)) count += 1;
  }
  return count;
}

function downloadCSV(filename: string, headers: string[], rows: Array<Array<string | number>>) {
  const escape = (cell: string | number) => `"${String(cell).replaceAll('"', '""')}"`;
  const csv = `\uFEFF${[headers, ...rows].map((row) => row.map(escape).join(',')).join('\n')}`;
  const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function parseCSV(text: string) {
  return text
    .replace(/^\uFEFF/, '')
    .trim()
    .split(/\r?\n/)
    .map((line) => line.split(',').map((cell) => cell.trim().replace(/^"|"$/g, '')));
}

function localTodayISO() {
  const date = new Date();
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function normalizeHeader(value: unknown) {
  return String(value ?? '').toLowerCase().replace(/[\s_\-（）()]/g, '');
}

function findColumn(headers: string[], aliases: string[]) {
  const normalized = headers.map(normalizeHeader);
  for (const alias of aliases) {
    const index = normalized.indexOf(normalizeHeader(alias));
    if (index >= 0) return index;
  }
  return -1;
}

function normalizeDateValue(value: unknown, fallbackYear?: number): string | null {
  const text = String(value ?? '').trim();
  if (!text || /^(#N\/A|NULL|N\/A|X|-|—)$/i.test(text)) return null;
  if (/^\d+(?:\.\d+)?$/.test(text) && Number(text) > 20000 && Number(text) < 80000) {
    const excelDate = XLSX.SSF.parse_date_code(Number(text));
    if (excelDate) return `${excelDate.y}-${String(excelDate.m).padStart(2, '0')}-${String(excelDate.d).padStart(2, '0')}`;
  }
  const iso = text.match(/(20\d{2})[-\/]([0-9]{1,2})[-\/]([0-9]{1,2})/);
  if (iso) return `${iso[1]}-${iso[2].padStart(2, '0')}-${iso[3].padStart(2, '0')}`;
  const compact = text.match(/^(20\d{2})([0-9]{2})([0-9]{2})$/);
  if (compact) return `${compact[1]}-${compact[2]}-${compact[3]}`;
  const mmdd = text.match(/^(?:20\d{2}[-\/]?)?([0-9]{1,2})[-\/]?([0-9]{1,2})(?:[_A-Za-z].*)?$/);
  if (mmdd && fallbackYear) return `${fallbackYear}-${mmdd[1].padStart(2, '0')}-${mmdd[2].padStart(2, '0')}`;
  return null;
}

function parseYear(value: unknown) {
  const match = String(value ?? '').match(/(20\d{2})/);
  return match ? Number(match[1]) : undefined;
}

function parseWorkbookSheet(name: string, sheet: XLSX.WorkSheet, requireAnnouncement = true): ParsedWorkbookSheet | null {
  const raw = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: false, defval: '' }) as unknown[][];
  const headers = (raw[0] ?? []).map((value) => String(value ?? '').trim());
  const codeIndex = findColumn(headers, ['comp_id', '公司碼', '公司代碼', '公司代號']);
  const nameIndex = findColumn(headers, ['comp_sname', '公司名稱', '公司簡稱', '公司名']);
  const periodIndex = findColumn(headers, ['annyymm', '財務年月']);
  const quarterIndex = findColumn(headers, ['quarter', '季別']);
  const announcementIndex = findColumn(headers, ['公告日1', '公告日', '財報日']);
  const completionIndex = findColumn(headers, ['完成日', '完成日期']);
  if ((codeIndex < 0 && nameIndex < 0) || periodIndex < 0 || quarterIndex < 0 || (requireAnnouncement && announcementIndex < 0)) return null;
  return {
    name,
    headers,
    rows: raw.slice(1).map((row) => row.map((value) => String(value ?? '').trim())),
    codeIndex,
    nameIndex,
    periodIndex,
    quarterIndex,
    announcementIndex,
    completionIndex,
  };
}

function createDailyRows(startDate: string, endDate: string): LiveRow[] {
  if (!startDate || !endDate || startDate > endDate) return [];
  const rows: LiveRow[] = [];
  for (let cursor = parseDate(startDate); cursor <= parseDate(endDate); cursor = addDays(cursor, 1)) {
    if (isWorkday(cursor)) rows.push({ id: makeId(), date: toISO(cursor), received: 0, completed: 0 });
  }
  return rows;
}

function syncDailyRows(rows: LiveRow[], startDate: string, endDate: string) {
  const previous = new Map(rows.map((row) => [row.date, row]));
  return createDailyRows(startDate, endDate).map((row) => {
    const old = previous.get(row.date);
    return old ? { ...row, id: old.id, received: old.received, completed: old.completed } : row;
  });
}

function reportWorkDate(announcementDate: string) {
  return toISO(addBusinessDays(parseDate(announcementDate), 1));
}

function rowsFromCompanies(companies: ImportedCompany[], startDate: string, endDate: string) {
  const receivedByDate = new Map<string, number>();
  const completedByDate = new Map<string, number>();
  companies.forEach((company) => {
    if (company.announcementDate) {
      const workDate = reportWorkDate(company.announcementDate);
      if (workDate >= startDate && workDate <= endDate) receivedByDate.set(workDate, (receivedByDate.get(workDate) ?? 0) + 1);
    }
    if (company.completionDate && company.completionDate >= startDate && company.completionDate <= endDate) completedByDate.set(company.completionDate, (completedByDate.get(company.completionDate) ?? 0) + 1);
  });
  const dates = new Set(createDailyRows(startDate, endDate).map((row) => row.date));
  [...receivedByDate.keys(), ...completedByDate.keys()].forEach((date) => dates.add(date));
  return [...dates].sort().map((date) => ({ id: makeId(), date, received: receivedByDate.get(date) ?? 0, completed: completedByDate.get(date) ?? 0 }));
}

function summarizeCompanies(companies: ImportedCompany[], startDate: string, today: string) {
  const previousWorkday = toISO(addBusinessDays(parseDate(startDate), -1));
  let announced = 0;
  let due = 0;
  let pending = 0;
  let beforeSeason = 0;
  companies.forEach((company) => {
    if (!company.announcementDate) {
      pending += 1;
      return;
    }
    const workDate = reportWorkDate(company.announcementDate);
    if (company.announcementDate >= previousWorkday) announced += 1;
    if (workDate < startDate) beforeSeason += 1;
    else if (workDate >= startDate && workDate <= today) due += 1;
  });
  return { announced, due, pending, notDue: Math.max(0, announced - due), beforeSeason };
}

function availableReportsForDate(date: string, companies: ImportedCompany[]) {
  const previousWorkday = toISO(addBusinessDays(parseDate(date), -1));
  return companies.filter((company) => company.announcementDate && company.announcementDate >= previousWorkday && company.announcementDate < date).length;
}

function openingBacklogFromCompanies(companies: ImportedCompany[], startDate: string) {
  return companies.filter((company) => {
    if (!company.completionDate || company.completionDate < startDate) return false;
    const workDate = company.announcementDate ? reportWorkDate(company.announcementDate) : null;
    return !workDate || workDate < startDate;
  }).length;
}

function mergeCompany(existing: ImportedCompany, incoming: ImportedCompany) {
  return {
    ...existing,
    name: existing.name || incoming.name,
    announcementDate: existing.announcementDate ?? incoming.announcementDate,
    completionDate: existing.completionDate ?? incoming.completionDate,
  };
}

function companyIdentity(code: string, name: string) {
  const normalizedCode = code.trim();
  if (normalizedCode) return /^\d+$/.test(normalizedCode) ? normalizedCode.replace(/^0+(?=\d)/, '') : normalizedCode.toLocaleLowerCase();
  return name.trim().toLocaleLowerCase();
}

function Metric({ label, value, note }: { label: string; value: string; note: string }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong>{value}</strong>
      <small>{note}</small>
    </div>
  );
}

function NumberField({
  label,
  value,
  onChange,
  min = 0,
  step = 1,
  suffix,
  hint,
}: {
  label: string;
  value: number;
  onChange: (value: number) => void;
  min?: number;
  step?: number;
  suffix?: string;
  hint?: string;
}) {
  return (
    <label className="field">
      <span className="field-label">{label}</span>
      <span className="input-with-suffix">
        <input
          type="number"
          min={min}
          step={step}
          value={value}
          onChange={(event) => onChange(Number(event.target.value))}
        />
        {suffix && <b>{suffix}</b>}
      </span>
      {hint && <small>{hint}</small>}
    </label>
  );
}

function DateField({ label, value, onChange, hint }: { label: string; value: string; onChange: (value: string) => void; hint?: string }) {
  return (
    <label className="field">
      <span className="field-label">{label}</span>
      <input type="date" value={value} onChange={(event) => onChange(event.target.value)} />
      {hint && <small>{hint}</small>}
    </label>
  );
}

function WeekdayDateField({ label, value, onChange }: { label: string; value: string; onChange: (value: string) => void }) {
  const date = value ? parseDate(value) : null;
  const weekday = date ? formatWeekday(date) : '尚未選擇';
  const weekend = date ? !isWorkday(date) : false;
  return (
    <label className="field weekday-field">
      <span className="field-label">{label}</span>
      <input type="date" value={value} onChange={(event) => onChange(event.target.value)} />
      <span className={`weekday-label ${weekend ? 'is-weekend' : ''}`}>
        {weekday}{weekend ? '・非工作日' : '・工作日'}
      </span>
    </label>
  );
}

type TeamCapacitySettings = Pick<ForecastSettings, 'teamMode' | 'members' | 'people' | 'speed' | 'efficiency'>;

function TeamCapacityEditor({ settings, onChange }: { settings: TeamCapacitySettings; onChange: (patch: Partial<TeamCapacitySettings>) => void }) {
  const team = getTeamCapacity(settings);

  function selectMode(mode: TeamMode) {
    if (mode === 'individual' && settings.members.length === 0) {
      const count = Math.max(1, Math.round(settings.people));
      onChange({
        teamMode: mode,
        members: Array.from({ length: count }, (_, index) => ({
          id: makeId(),
          name: `同仁 ${index + 1}`,
          dailyBooks: settings.speed,
        })),
      });
      return;
    }
    onChange({ teamMode: mode });
  }

  function updateMember(id: string, patch: Partial<TeamMember>) {
    onChange({ members: settings.members.map((member) => (member.id === id ? { ...member, ...patch } : member)) });
  }

  return (
    <div className="team-editor">
      <div className="capacity-mode" role="group" aria-label="人力速度設定方式">
        <button className={settings.teamMode === 'uniform' ? 'active' : ''} onClick={() => selectMode('uniform')}>所有人同一速度</button>
        <button className={settings.teamMode === 'individual' ? 'active' : ''} onClick={() => selectMode('individual')}>逐人設定本數</button>
      </div>

      {settings.teamMode === 'uniform' ? (
        <div className="form-card three-fields team-form-card">
          <NumberField label="投入人數" value={settings.people} min={1} onChange={(value) => onChange({ people: value })} suffix="人" />
          <NumberField label="每人每日速度" value={settings.speed} min={0.1} step={0.1} onChange={(value) => onChange({ speed: value })} suffix="本／日" />
          <NumberField label="有效工時率" value={settings.efficiency} min={1} onChange={(value) => onChange({ efficiency: Math.min(value, 100) })} suffix="%" hint="扣除會議、複核與雜務" />
        </div>
      ) : (
        <div className="member-card">
          <div className="member-list">
            <div className="member-row member-header"><span>人員</span><span>每日可完成本數</span><span></span></div>
            {settings.members.map((member, index) => (
              <div className="member-row" key={member.id}>
                <input aria-label={`第 ${index + 1} 位人員姓名`} type="text" value={member.name} placeholder={`同仁 ${index + 1}`} onChange={(event) => updateMember(member.id, { name: event.target.value })} />
                <span className="input-with-suffix member-speed"><input aria-label={`${member.name || `第 ${index + 1} 位人員`}每日可完成本數`} type="number" min="0" step="0.1" value={member.dailyBooks} onChange={(event) => updateMember(member.id, { dailyBooks: Number(event.target.value) })} /><b>本／日</b></span>
                <button className="delete-row" aria-label={`刪除${member.name || `第 ${index + 1} 位人員`}`} onClick={() => onChange({ members: settings.members.filter((item) => item.id !== member.id) })}>×</button>
              </div>
            ))}
          </div>
          <button className="add-row" onClick={() => onChange({ members: [...settings.members, { id: makeId(), name: `同仁 ${settings.members.length + 1}`, dailyBooks: settings.speed }] })}>＋ 新增人員</button>
        </div>
      )}

      <div className="team-summary">
        <span>團隊人數 <strong>{team.teamCount} 人</strong></span>
        <span>團隊每日產能 <strong>{compactNumber(team.dailyCapacity, 1)} 本</strong></span>
        <small>{settings.teamMode === 'individual' ? '逐人本數視為實際日產能，不再乘有效工時率' : '已套用有效工時率'}</small>
      </div>
    </div>
  );
}

function EmptyResult({ text }: { text: string }) {
  return (
    <div className="empty-result">
      <span aria-hidden="true">⌁</span>
      <strong>還差一點資料</strong>
      <p>{text}</p>
    </div>
  );
}

function ForecastWorkspace() {
  const [settings, setSettings] = useState(initialForecastSettings);
  const [rows, setRows] = useState<HistoricalRow[]>([
    { id: 'blank-1', date: '', quantity: 0 },
    { id: 'blank-2', date: '', quantity: 0 },
    { id: 'blank-3', date: '', quantity: 0 },
  ]);
  const [hydrated, setHydrated] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const frame = window.requestAnimationFrame(() => {
      const saved = localStorage.getItem(FORECAST_KEY);
      if (saved) {
        try {
          const parsed = JSON.parse(saved) as { settings: Partial<ForecastSettings>; rows: HistoricalRow[] };
          const migrateDefaultPeople = parsed.settings.defaultsVersion === undefined && parsed.settings.people === 12;
          const savedMembers = parsed.settings.members ?? [];
          const migrateDefaultMembers = (parsed.settings.defaultsVersion ?? 0) < 3 && savedMembers.length === 8;
          setSettings({
            ...initialForecastSettings,
            ...parsed.settings,
            defaultsVersion: 3,
            people: migrateDefaultPeople ? 4 : (parsed.settings.people ?? 4),
            members: migrateDefaultMembers ? savedMembers.slice(0, 4) : savedMembers,
          });
          setRows(parsed.rows);
        } catch { /* keep defaults */ }
      }
      setHydrated(true);
    });
    return () => window.cancelAnimationFrame(frame);
  }, []);

  useEffect(() => {
    if (hydrated) localStorage.setItem(FORECAST_KEY, JSON.stringify({ settings, rows }));
  }, [hydrated, rows, settings]);

  const validRows = useMemo(
    () => rows.filter((row) => row.date && row.quantity > 0).sort((a, b) => a.date.localeCompare(b.date)),
    [rows],
  );

  const result = useMemo(() => {
    if (!validRows.length || !settings.historicalAnchor || !settings.currentAnchor || !settings.completionDate) return null;

    const oldAnchor = parseDate(settings.historicalAnchor);
    const newAnchor = parseDate(settings.currentAnchor);
    const deadline = parseDate(settings.completionDate);
    const team = getTeamCapacity(settings);
    const dailyCapacity = team.dailyCapacity;
    const mapped = validRows.map((row) => {
      const offset = businessDayDiff(parseDate(row.date), oldAnchor);
      return { ...row, offset, mappedDate: toISO(addBusinessDays(newAnchor, offset)) };
    });
    const demand = new Map<string, number>();
    mapped.forEach((row) => demand.set(row.mappedDate, (demand.get(row.mappedDate) ?? 0) + row.quantity));
    const firstDate = parseDate(mapped[0].mappedDate);
    const lastDate = parseDate(mapped[mapped.length - 1].mappedDate);

    const simulate = (capacity: number, collectPoints = false) => {
      let backlog = 0;
      let backlogAtDeadline = 0;
      let peakBacklog = 0;
      let finishDate: Date | null = null;
      const points: Array<{ date: string; demand: number; capacity: number; backlog: number }> = [];
      const end = addDays(lastDate > deadline ? lastDate : deadline, 365);
      for (let cursor = firstDate; cursor <= end; cursor = addDays(cursor, 1)) {
        const date = toISO(cursor);
        const incoming = demand.get(date) ?? 0;
        const available = isWorkday(cursor) ? capacity : 0;
        backlog = Math.max(0, backlog + incoming - available);
        peakBacklog = Math.max(peakBacklog, backlog);
        if (collectPoints && (incoming > 0 || backlog > 0 || (cursor >= firstDate && cursor <= deadline))) {
          points.push({ date, demand: incoming, capacity: available, backlog });
        }
        if (date === settings.completionDate) backlogAtDeadline = backlog;
        if (cursor >= lastDate && backlog <= 0) {
          finishDate = cursor;
          break;
        }
      }
      const lateDemand = mapped.filter((row) => parseDate(row.mappedDate) > deadline).reduce((sum, row) => sum + row.quantity, 0);
      return { backlogAtDeadline: backlogAtDeadline + lateDemand, peakBacklog, finishDate, points };
    };

    const current = simulate(dailyCapacity, true);
    let requiredPeople: number | null = null;
    if (team.perPersonCapacity > 0 && !mapped.some((row) => parseDate(row.mappedDate) > deadline)) {
      for (let people = 1; people <= 999; people += 1) {
        if (simulate(people * team.perPersonCapacity).backlogAtDeadline <= 0) {
          requiredPeople = people;
          break;
        }
      }
    }
    const total = validRows.reduce((sum, row) => sum + row.quantity, 0);
    const plannedDays = countWorkdays(firstDate, deadline);
    const plannedCapacity = plannedDays * dailyCapacity;
    const load = plannedCapacity > 0 ? (total / plannedCapacity) * 100 : Infinity;
    const mappedWithBacklog = mapped.map((row) => ({
      ...row,
      backlog: current.points.find((point) => point.date === row.mappedDate)?.backlog ?? 0,
    }));
    return { ...current, mapped: mappedWithBacklog, total, dailyCapacity, teamCount: team.teamCount, plannedDays, plannedCapacity, load, requiredPeople };
  }, [settings, validRows]);

  function updateRow(id: string, patch: Partial<HistoricalRow>) {
    setRows((current) => current.map((row) => (row.id === id ? { ...row, ...patch } : row)));
  }

  function importRows(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const lines = parseCSV(String(reader.result));
      const data = lines[0]?.some((cell) => /date|日期/i.test(cell)) ? lines.slice(1) : lines;
      const imported = data
        .filter((line) => /^\d{4}-\d{2}-\d{2}$/.test(line[0]) && Number(line[1]) >= 0)
        .map((line) => ({ id: makeId(), date: line[0], quantity: Number(line[1]) }));
      if (imported.length) setRows(imported);
    };
    reader.readAsText(file);
    event.target.value = '';
  }

  const status = result && result.backlogAtDeadline <= 0;
  const maxChart = result ? Math.max(1, ...result.points.slice(-24).flatMap((point) => [point.demand, point.capacity, point.backlog])) : 1;

  return (
    <div className="workspace-grid">
      <section className="input-column">
        <div className="section-heading"><div><span>STEP 1</span><h2>設定本季人力</h2></div></div>
        <TeamCapacityEditor settings={settings} onChange={(patch) => setSettings((current) => ({ ...current, ...patch }))} />

        <div className="section-heading">
          <div><span>STEP 2</span><h2>設定換算基準</h2></div>
          <details>
            <summary>T 日怎麼算？</summary>
            <p>從歷史 T 日往前或往後數工作日，星期六、星期日直接跳過，不增加 T 日數；再從本季 T 日以相同工作日數回推或前推。若輸入日期本身是週末，會與相鄰工作日共用同一個 T 值。</p>
          </details>
        </div>
        <div className="form-card three-fields">
          <DateField label="歷史 T 日" value={settings.historicalAnchor} onChange={(value) => setSettings({ ...settings, historicalAnchor: value })} hint="用來計算每筆 T±N" />
          <DateField label="本季 T 日" value={settings.currentAnchor} onChange={(value) => setSettings({ ...settings, currentAnchor: value })} hint="同一相對日會映射到這裡" />
          <DateField label="要求清完日" value={settings.completionDate} onChange={(value) => setSettings({ ...settings, completionDate: value })} hint="判斷做不做得完的期限" />
        </div>

        <div className="section-heading table-heading">
          <div><span>STEP 3</span><h2>輸入歷史每日財報量</h2></div>
          <div className="inline-actions">
            <button className="text-button" onClick={() => setRows(forecastExample)}>載入範例</button>
            <button className="text-button" onClick={() => fileRef.current?.click()}>匯入 CSV</button>
            <input ref={fileRef} className="sr-only" type="file" accept=".csv,text/csv" onChange={importRows} />
          </div>
        </div>
        <div className="data-table-card">
          <div className="data-table historical-table">
            <div className="table-row table-header"><span>歷史日期</span><span>T 日</span><span>財報數量</span><span></span></div>
            {rows.map((row) => {
              const offset = row.date && settings.historicalAnchor ? businessDayDiff(parseDate(row.date), parseDate(settings.historicalAnchor)) : null;
              return (
                <div className="table-row" key={row.id}>
                  <input aria-label="歷史日期" type="date" value={row.date} onChange={(event) => updateRow(row.id, { date: event.target.value })} />
                  <span className="t-chip">{offset === null ? '—' : formatT(offset)}</span>
                  <input aria-label="財報數量" type="number" min="0" value={row.quantity || ''} placeholder="0" onChange={(event) => updateRow(row.id, { quantity: Number(event.target.value) })} />
                  <button className="delete-row" aria-label="刪除這一列" onClick={() => setRows((current) => current.filter((item) => item.id !== row.id))}>×</button>
                </div>
              );
            })}
          </div>
          <button className="add-row" onClick={() => setRows((current) => [...current, { id: makeId(), date: '', quantity: 0 }])}>＋ 新增一天</button>
        </div>

      </section>

      <aside className="result-column">
        <div className={`verdict ${result ? (status ? 'is-good' : 'is-risk') : ''}`}>
          {!result ? <EmptyResult text="輸入至少一筆歷史日期與數量，就會自動開始預估。" /> : (
            <>
              <span className="verdict-kicker">人力負荷結論</span>
              <div className="verdict-icon" aria-hidden="true">{status ? '✓' : '!'}</div>
              <h2>{status ? '目前人力可負擔' : '目前人力有缺口'}</h2>
              <p>{status ? `預估可在 ${result.finishDate ? formatDate(result.finishDate) : '期限內'} 清完。` : `清完日仍會剩下約 ${compactNumber(result.backlogAtDeadline)} 件。`}</p>
              <div className="capacity-line"><span>每日有效產能</span><strong>{compactNumber(result.dailyCapacity, 1)} 件</strong></div>
            </>
          )}
        </div>

        {result && (
          <>
            <div className="metrics-grid">
              <Metric label="預估總量" value={`${compactNumber(result.total)} 件`} note={`${validRows.length} 個有量日期`} />
              <Metric label="負荷率" value={Number.isFinite(result.load) ? `${compactNumber(result.load)}%` : '—'} note={`${result.plannedDays} 個工作天`} />
              <Metric label="最少人力" value={result.requiredPeople ? `${result.requiredPeople} 人` : '無法估算'} note={result.requiredPeople ? `目前 ${result.teamCount} 人` : '有財報在清完日後才到'} />
              <Metric label="預估清完" value={result.finishDate ? formatDate(result.finishDate) : '超過一年'} note={`尖峰待辦 ${compactNumber(result.peakBacklog)} 件`} />
            </div>

            <div className="chart-card">
              <div className="card-title"><div><span>每日壓力</span><h3>量能與待辦走勢</h3></div><div className="legend"><i className="demand-dot" />收到 <i className="capacity-dot" />產能 <i className="backlog-dot" />待辦</div></div>
              <div className="bar-chart" aria-label="每日收到量、產能與待辦量圖表">
                {result.points.slice(-24).map((point) => (
                  <div className="chart-day" key={point.date} title={`${formatDate(point.date)}｜收到 ${compactNumber(point.demand)}、產能 ${compactNumber(point.capacity)}、待辦 ${compactNumber(point.backlog)}`}>
                    <div className="bars">
                      <i className="demand-bar" style={{ height: `${(point.demand / maxChart) * 100}%` }} />
                      <i className="capacity-bar" style={{ height: `${(point.capacity / maxChart) * 100}%` }} />
                      <i className="backlog-bar" style={{ height: `${(point.backlog / maxChart) * 100}%` }} />
                    </div>
                    <small>{parseDate(point.date).getUTCDate()}</small>
                  </div>
                ))}
              </div>
            </div>

            <div className="mapping-card">
              <div className="card-title"><div><span>T 日映射</span><h3>歷史量 → 本季日期</h3></div><button className="text-button" onClick={() => downloadCSV('忙季預估結果.csv', ['歷史日期', 'T日', '本季日期', '財報數量', '當日結束待辦'], result.mapped.map((row) => [row.date, formatT(row.offset), row.mappedDate, row.quantity, Math.round(row.backlog)]))}>匯出結果</button></div>
              <div className="mapping-list">
                {result.mapped.map((row) => (
                  <div key={row.id}><span>{formatDate(row.date)}</span><b>{formatT(row.offset)}</b><span>→</span><strong>{formatDate(row.mappedDate)}</strong><em>{compactNumber(row.quantity)} 件</em></div>
                ))}
              </div>
            </div>
          </>
        )}
      </aside>
    </div>
  );
}

function LiveWorkspace() {
  const [settings, setSettings] = useState(initialLiveSettings);
  const [rows, setRows] = useState<LiveRow[]>(() => createDailyRows(initialLiveSettings.seasonStartDate, localTodayISO()));
  const [hydrated, setHydrated] = useState(false);
  const companyFileRef = useRef<HTMLInputElement>(null);
  const manpowerFileRef = useRef<HTMLInputElement>(null);
  const [companyWorkbookName, setCompanyWorkbookName] = useState('');
  const [manpowerWorkbookName, setManpowerWorkbookName] = useState('');
  const [companyWorkbookSheets, setCompanyWorkbookSheets] = useState<ParsedWorkbookSheet[]>([]);
  const [manpowerWorkbookSheets, setManpowerWorkbookSheets] = useState<ParsedWorkbookSheet[]>([]);
  const [companyImportSheet, setCompanyImportSheet] = useState('__all__');
  const [importAnnyymm, setImportAnnyymm] = useState('');
  const [importQuarter, setImportQuarter] = useState('');
  const [importError, setImportError] = useState('');
  const [importSummary, setImportSummary] = useState<{ total: number; announced: number; due: number; pending: number; notDue: number; beforeSeason: number; duplicates: number } | null>(null);

  useEffect(() => {
    const frame = window.requestAnimationFrame(() => {
          const saved = localStorage.getItem(LIVE_KEY);
      if (saved) {
        try {
          const parsed = JSON.parse(saved) as { settings: Partial<LiveSettings>; rows: LiveRow[] };
          const migrateDefaultPeople = parsed.settings.defaultsVersion === undefined && parsed.settings.people === 8;
          const savedMembers = parsed.settings.members ?? [];
          const migrateDefaultMembers = (parsed.settings.defaultsVersion ?? 0) < 3 && savedMembers.length === 8;
          const restoredSettings = {
            ...initialLiveSettings,
            ...parsed.settings,
            defaultsVersion: 3,
            people: migrateDefaultPeople ? 4 : (parsed.settings.people ?? 4),
            members: migrateDefaultMembers ? savedMembers.slice(0, 4) : savedMembers,
          };
          if (restoredSettings.importedCompanies.length) {
            restoredSettings.expectedRemaining = summarizeCompanies(restoredSettings.importedCompanies, restoredSettings.seasonStartDate, localTodayISO()).pending;
            if (!restoredSettings.openingBacklog) restoredSettings.openingBacklog = openingBacklogFromCompanies(restoredSettings.importedCompanies, restoredSettings.seasonStartDate);
          }
          setSettings(restoredSettings);
          setRows(restoredSettings.importedCompanies.length ? rowsFromCompanies(restoredSettings.importedCompanies, restoredSettings.seasonStartDate, localTodayISO()) : (parsed.rows?.some((row) => row.date) ? parsed.rows : createDailyRows(restoredSettings.seasonStartDate, localTodayISO())));
        } catch { /* keep defaults */ }
      }
      setHydrated(true);
    });
    return () => window.cancelAnimationFrame(frame);
  }, []);

  useEffect(() => {
    if (hydrated) localStorage.setItem(LIVE_KEY, JSON.stringify({ settings, rows }));
  }, [hydrated, rows, settings]);

  const validRows = useMemo(() => rows.filter((row) => row.date).sort((a, b) => a.date.localeCompare(b.date)), [rows]);
  const currentImportSummary = useMemo(() => {
    if (!settings.importedCompanies.length) return importSummary;
    const today = localTodayISO();
    const counts = summarizeCompanies(settings.importedCompanies, settings.seasonStartDate, today);
    return { total: settings.importedCompanies.length, ...counts, duplicates: importSummary?.duplicates ?? 0 };
  }, [importSummary, settings.importedCompanies, settings.seasonStartDate]);
  const tableRows = useMemo(() => {
    const calculated = rows.slice().sort((a, b) => a.date.localeCompare(b.date)).reduce<{ backlog: number; rows: Array<LiveRow & { incoming: number; available: number; difference: number; endingBacklog: number }> }>((accumulator, row) => {
      const incoming = settings.importedCompanies.length ? availableReportsForDate(row.date, settings.importedCompanies) : row.received;
      const available = accumulator.backlog + incoming;
      const difference = row.completed - available;
      const endingBacklog = Math.max(0, available - row.completed);
      return { backlog: endingBacklog, rows: [...accumulator.rows, { ...row, incoming, available, difference, endingBacklog }] };
    }, { backlog: settings.openingBacklog, rows: [] });
    return calculated.rows;
  }, [rows, settings.importedCompanies, settings.openingBacklog]);
  const dailyComparisonByDate = useMemo(() => {
    const result = new Map<string, DailyComparison>();
    if (!settings.importedCompanies.length) return result;
    const start = settings.seasonStartDate;
    const previousWorkday = toISO(addBusinessDays(parseDate(start), -1));
    const pending = new Map<string, string>();
    settings.importedCompanies.forEach((company) => {
      const workDate = company.announcementDate ? reportWorkDate(company.announcementDate) : null;
      if (company.completionDate && company.completionDate >= start && (!workDate || workDate < start)) pending.set(companyIdentity(company.code, company.name), company.code || company.name);
    });
    tableRows.forEach((row) => {
      settings.importedCompanies.forEach((company) => {
        if (!company.announcementDate || company.announcementDate < previousWorkday || company.announcementDate >= row.date) return;
        const key = companyIdentity(company.code, company.name);
        if (!pending.has(key)) pending.set(key, company.code || company.name);
      });
      const completedCompanies = settings.importedCompanies.filter((company) => company.completionDate === row.date);
      const availableCodes = [...pending.values()];
      const completedCodes = completedCompanies.map((company) => company.code || company.name);
      const matchedCodes = completedCompanies.filter((company) => pending.has(companyIdentity(company.code, company.name))).map((company) => company.code || company.name);
      const completedOnlyCodes = completedCompanies.filter((company) => !pending.has(companyIdentity(company.code, company.name))).map((company) => company.code || company.name);
      completedCompanies.forEach((company) => pending.delete(companyIdentity(company.code, company.name)));
      result.set(row.date, { date: row.date, availableCodes, completedCodes, matchedCodes, completedOnlyCodes, endingBacklogCodes: [...pending.values()] });
    });
    return result;
  }, [settings.importedCompanies, settings.seasonStartDate, tableRows]);
  const result = useMemo(() => {
    if (!validRows.length || !settings.completionDate) return null;
    const received = validRows.reduce((sum, row) => sum + row.received, 0);
    const completed = validRows.reduce((sum, row) => sum + row.completed, 0);
    const backlog = Math.max(0, settings.openingBacklog + received - completed);
    const futureAnnounced = settings.importedCompanies.filter((company) => {
      if (!company.announcementDate) return false;
      const workDate = reportWorkDate(company.announcementDate);
      return workDate >= settings.seasonStartDate && workDate > localTodayISO();
    }).length;
    const outstanding = backlog + settings.expectedRemaining + futureAnnounced;
    const latestDate = parseDate(validRows[validRows.length - 1].date);
    const nextDate = addDays(latestDate, 1);
    const deadline = parseDate(settings.completionDate);
    const daysLeft = countWorkdays(nextDate, deadline);
    const team = getTeamCapacity(settings);
    const dailyCapacity = team.dailyCapacity;
    const remainingCapacity = daysLeft * dailyCapacity;
    const feasible = outstanding <= remainingCapacity;
    const neededDaily = daysLeft > 0 ? outstanding / daysLeft : Infinity;
    const requiredPeople = daysLeft > 0 && team.perPersonCapacity > 0 ? Math.ceil(outstanding / (daysLeft * team.perPersonCapacity)) : null;
    const actualDays = new Set(validRows.filter((row) => row.completed > 0).map((row) => row.date)).size;
    const observedPerPerson = actualDays > 0 && team.teamCount > 0 ? completed / actualDays / team.teamCount : 0;
    let finishDate: Date | null = outstanding <= 0 ? latestDate : null;
    let remaining = outstanding;
    for (let cursor = nextDate; !finishDate && cursor <= addDays(nextDate, 365); cursor = addDays(cursor, 1)) {
      if (isWorkday(cursor)) remaining -= dailyCapacity;
      if (remaining <= 0) finishDate = cursor;
    }
    let cumulativeReceived = settings.openingBacklog;
    let cumulativeCompleted = 0;
    const points = validRows.map((row) => {
      cumulativeReceived += row.received;
      cumulativeCompleted += row.completed;
      return { ...row, backlog: Math.max(0, cumulativeReceived - cumulativeCompleted) };
    });
    return { received, completed, backlog, outstanding, daysLeft, dailyCapacity, teamCount: team.teamCount, remainingCapacity, feasible, neededDaily, requiredPeople, observedPerPerson, finishDate, points };
  }, [settings, validRows]);

  function updateRow(id: string, patch: Partial<LiveRow>) {
    setRows((current) => current.map((row) => (row.id === id ? { ...row, ...patch } : row)));
  }

  function updateSeasonStart(value: string) {
    setSettings((current) => ({ ...current, seasonStartDate: value, openingBacklog: current.importedCompanies.length ? openingBacklogFromCompanies(current.importedCompanies, value) : current.openingBacklog }));
    setRows((current) => settings.importedCompanies.length ? rowsFromCompanies(settings.importedCompanies, value, localTodayISO()) : syncDailyRows(current, value, localTodayISO()));
  }

  async function readWorkbook(event: ChangeEvent<HTMLInputElement>, kind: 'company' | 'manpower') {
    const file = event.target.files?.[0];
    if (!file) return;
    setImportError('');
    setImportSummary(null);
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array', raw: false, cellDates: false, bookVBA: true });
      const sheets = workbook.SheetNames.map((name) => parseWorkbookSheet(name, workbook.Sheets[name], kind === 'company')).filter((sheet): sheet is ParsedWorkbookSheet => Boolean(sheet));
      if (!sheets.length) throw new Error(kind === 'company' ? '找不到可辨識的公司清單工作表（需要公司代碼、annyymm、quarter、公告日欄位）。' : '找不到可辨識的人力工作表（需要公司代碼、annyymm、quarter、完成日欄位）。');
      if (kind === 'company') {
        setCompanyWorkbookName(file.name);
        setCompanyWorkbookSheets(sheets);
        setCompanyImportSheet('__all__');
      } else {
        setManpowerWorkbookName(file.name);
        setManpowerWorkbookSheets(sheets);
      }
      const periods = Array.from(new Set(sheets.flatMap((sheet) => sheet.rows.map((row) => row[sheet.periodIndex]).filter(Boolean)))).sort();
      const quarters = Array.from(new Set(sheets.flatMap((sheet) => sheet.rows.map((row) => row[sheet.quarterIndex]).filter(Boolean)))).sort();
      if (!importAnnyymm || !periods.includes(importAnnyymm)) setImportAnnyymm(periods.includes('2026-06-01') ? '2026-06-01' : (periods[0] ?? ''));
      if (!importQuarter || !quarters.includes(importQuarter)) setImportQuarter(quarters.includes('2') ? '2' : (quarters[0] ?? ''));
    } catch (error) {
      setImportError(error instanceof Error ? error.message : '檔案讀取失敗，請確認是 Excel 或 XLSM。');
    }
    event.target.value = '';
  }

  function applyCompanyImport() {
    if (!companyWorkbookSheets.length || !importAnnyymm || !importQuarter) return;
    const selectedSheets = companyImportSheet === '__all__' ? companyWorkbookSheets : companyWorkbookSheets.filter((sheet) => sheet.name === companyImportSheet);
    const companyMap = new Map<string, ImportedCompany>();
    let rawCount = 0;
    const fallbackYear = parseYear(importAnnyymm);
    selectedSheets.forEach((sheet) => sheet.rows.forEach((row) => {
      if (String(row[sheet.periodIndex] ?? '').trim() !== importAnnyymm || String(row[sheet.quarterIndex] ?? '').trim() !== importQuarter) return;
      const code = String(row[sheet.codeIndex] ?? '').trim();
      const name = String(row[sheet.nameIndex] ?? '').trim();
      const key = companyIdentity(code, name);
      if (!key) return;
      rawCount += 1;
      const company: ImportedCompany = {
        id: key,
        code,
        name: name || code,
        announcementDate: sheet.announcementIndex >= 0 ? normalizeDateValue(row[sheet.announcementIndex], fallbackYear) : null,
        completionDate: null,
      };
      companyMap.set(key, companyMap.has(key) ? mergeCompany(companyMap.get(key)!, company) : company);
    }));
    const companies = Array.from(companyMap.values());
    if (!companies.length) {
      setImportError('在目前選擇的 annyymm 與 quarter 找不到公司資料。');
      return;
    }
    const today = localTodayISO();
    const start = settings.seasonStartDate;
    const counts = summarizeCompanies(companies, start, today);
    setSettings((current) => ({ ...current, expectedRemaining: counts.pending, importedCompanies: companies }));
    setImportSummary({ total: companies.length, ...counts, duplicates: Math.max(0, rawCount - companies.length) });
    setImportError('');
  }

  function applyManpowerImport() {
    if (!manpowerWorkbookSheets.length || !settings.importedCompanies.length || !importAnnyymm || !importQuarter) {
      setImportError(!settings.importedCompanies.length ? '請先匯入公司財報清單，再匯入人力完成紀錄。' : '請先選擇 annyymm 與 quarter。');
      return;
    }
    const normalizeName = (value: string) => value.trim().toLocaleLowerCase();
    const people = settings.members.filter((person) => person.name.trim());
    if (!people.length) {
      setImportError('請先在 Step 1 選擇「逐人設定本數」並輸入人名，系統會用人名尋找對應工作表。');
      return;
    }
    const selectedSheets = people.flatMap((person) => manpowerWorkbookSheets.filter((sheet) => normalizeName(sheet.name) === normalizeName(person.name)));
    if (!selectedSheets.length) {
      setImportError(`找不到與輸入人名相同的工作表：${people.map((person) => person.name).join('、')}`);
      return;
    }
    const fallbackYear = parseYear(importAnnyymm);
    const completionByCompany = new Map<string, string>();
    selectedSheets.forEach((sheet) => sheet.rows.forEach((row) => {
      if (String(row[sheet.periodIndex] ?? '').trim() !== importAnnyymm || String(row[sheet.quarterIndex] ?? '').trim() !== importQuarter || sheet.completionIndex < 0) return;
      const code = String(row[sheet.codeIndex] ?? '').trim();
      const name = String(row[sheet.nameIndex] ?? '').trim();
      const key = companyIdentity(code, name);
      const completion = normalizeDateValue(row[sheet.completionIndex], fallbackYear);
      if (key && completion && !completionByCompany.has(key)) completionByCompany.set(key, completion);
    }));
    const updatedCompanies = settings.importedCompanies.map((company) => ({ ...company, completionDate: completionByCompany.get(companyIdentity(company.code, company.name)) ?? company.completionDate }));
    const today = localTodayISO();
    const start = settings.seasonStartDate;
    const counts = summarizeCompanies(updatedCompanies, start, today);
    setRows(rowsFromCompanies(updatedCompanies, start, today));
    setSettings((current) => ({ ...current, openingBacklog: openingBacklogFromCompanies(updatedCompanies, start), expectedRemaining: counts.pending, importedCompanies: updatedCompanies }));
    setImportSummary({ total: updatedCompanies.length, ...counts, duplicates: 0 });
    setImportError(`已依 ${selectedSheets.map((sheet) => sheet.name).join('、')} 的完成日更新實績。`);
  }

  const maxTrend = result ? Math.max(1, ...result.points.flatMap((point) => [point.received, point.completed, point.backlog])) : 1;

  return (
    <div className="workspace-grid">
      <section className="input-column">
        <div className="section-heading"><div><span>STEP 1</span><h2>設定目前人力</h2></div></div>
        <TeamCapacityEditor settings={settings} onChange={(patch) => setSettings((current) => ({ ...current, ...patch }))} />

        <div className="section-heading"><div><span>STEP 2</span><h2>設定忙季日期</h2></div></div>
        <div className="form-card three-fields season-date-card">
          <WeekdayDateField label="忙季開始日" value={settings.seasonStartDate} onChange={updateSeasonStart} />
          <WeekdayDateField label="T 日" value={settings.anchorDate} onChange={(value) => setSettings({ ...settings, anchorDate: value })} />
          <WeekdayDateField label="忙季結束日" value={settings.completionDate} onChange={(value) => setSettings({ ...settings, completionDate: value })} />
        </div>
        <div className="section-heading"><div><span>STEP 3</span><h2>匯入公司與完成紀錄</h2></div><div className="inline-actions"><button className="text-button" onClick={() => setRows(liveExample)}>載入範例</button><button className="text-button" onClick={() => setRows((current) => syncDailyRows(current, settings.seasonStartDate, localTodayISO()))}>補齊到今天</button></div></div>
        <div className="import-panel">
          <p>公司清單與人力完成紀錄是兩個檔案。先匯入公司清單，再依上方輸入的人名尋找人力檔案中同名工作表。公告日會往後對應到下一個工作日：忙季第一天會包含忙季開始日前一個工作日公告的財報。每日表格會保留 Excel 實際完成日，即使是假日；若當天完成量大於收到量，代表正在消化前幾日待辦。</p>
          <div className="import-file-row"><span>公司財報清單（公告日）</span><button className="text-button" onClick={() => companyFileRef.current?.click()}>選擇 Excel／XLSM</button><input ref={companyFileRef} className="sr-only" type="file" accept=".xlsx,.xlsm" onChange={(event) => readWorkbook(event, 'company')} />{companyWorkbookName && <strong>{companyWorkbookName}</strong>}</div>
          <div className="import-file-row"><span>人力完成紀錄（依人名找工作表）</span><button className="text-button" onClick={() => manpowerFileRef.current?.click()}>選擇 Excel／XLSM</button><input ref={manpowerFileRef} className="sr-only" type="file" accept=".xlsx,.xlsm" onChange={(event) => readWorkbook(event, 'manpower')} />{manpowerWorkbookName && <strong>{manpowerWorkbookName}</strong>}</div>
          {companyWorkbookSheets.length > 0 && <div className="import-controls">
            <label className="field"><span className="field-label">公司清單工作表</span><select value={companyImportSheet} onChange={(event) => setCompanyImportSheet(event.target.value)}><option value="__all__">全部辨識工作表（自動去重）</option>{companyWorkbookSheets.map((sheet) => <option key={sheet.name} value={sheet.name}>{sheet.name}</option>)}</select></label>
            <label className="field"><span className="field-label">annyymm</span><select value={importAnnyymm} onChange={(event) => setImportAnnyymm(event.target.value)}>{Array.from(new Set(companyWorkbookSheets.flatMap((sheet) => sheet.rows.map((row) => row[sheet.periodIndex]).filter(Boolean)))).sort().map((value) => <option key={value} value={value}>{value}</option>)}</select></label>
            <label className="field"><span className="field-label">quarter</span><select value={importQuarter} onChange={(event) => setImportQuarter(event.target.value)}>{Array.from(new Set(companyWorkbookSheets.flatMap((sheet) => sheet.rows.map((row) => row[sheet.quarterIndex]).filter(Boolean)))).sort().map((value) => <option key={value} value={value}>{value}</option>)}</select></label>
            <button className="primary-button import-button" onClick={applyCompanyImport}>匯入公司財報清單</button>
          </div>}
          {manpowerWorkbookSheets.length > 0 && <button className="primary-button manpower-import-button" onClick={applyManpowerImport}>依人名匯入完成日期並更新實績</button>}
          {importError && <div className="import-error">{importError}</div>}
          {currentImportSummary && <div className="import-summary"><span>唯一公司 <strong>{compactNumber(currentImportSummary.total)}</strong></span><span>已公告 <strong>{compactNumber(currentImportSummary.announced)}</strong></span><span>今日前應處理 <strong>{compactNumber(currentImportSummary.due)}</strong></span><span>尚未公告 <strong>{compactNumber(currentImportSummary.pending)}</strong></span>{currentImportSummary.notDue > 0 && <span>已公告但尚未到工作日 <strong>{compactNumber(currentImportSummary.notDue)}</strong></span>}{currentImportSummary.beforeSeason > 0 && <span>忙季前已公告 <strong>{compactNumber(currentImportSummary.beforeSeason)}</strong></span>}<span>去除重複 <strong>{compactNumber(currentImportSummary.duplicates)}</strong></span></div>}
        </div>
        <div className="data-table-card">
          <div className="data-table live-table">
            <div className="table-row table-header"><span>日期</span><span>當日可做本數</span><span>實際完成</span><span>完成差額</span><span></span></div>
            {tableRows.map((row) => (
              <div className="table-row" key={row.id}>
                <div className="table-date-cell">
                  <input aria-label="實績日期" type="date" value={row.date} onChange={(event) => updateRow(row.id, { date: event.target.value })} />
                  {row.date && <small className={!isWorkday(parseDate(row.date)) ? 'is-weekend' : ''}>{formatWeekday(row.date)}{!isWorkday(parseDate(row.date)) ? '・非工作日' : ''}</small>}
                  {dailyComparisonByDate.get(row.date) && <details className="code-detail"><summary>查看公司碼</summary><div><small>可做：{dailyComparisonByDate.get(row.date)!.availableCodes.join('、') || '—'}</small><small>已做：{dailyComparisonByDate.get(row.date)!.completedCodes.join('、') || '—'}</small><small>未配對：{dailyComparisonByDate.get(row.date)!.completedOnlyCodes.join('、') || '—'}</small></div></details>}
                </div>
                <input aria-label="當日可做本數" type="number" min="0" value={(dailyComparisonByDate.get(row.date)?.availableCodes.length ?? row.available) || ''} placeholder="0" readOnly={settings.importedCompanies.length > 0} onChange={(event) => updateRow(row.id, { received: Number(event.target.value) })} />
                <input aria-label="實際完成量" type="number" min="0" value={row.completed || ''} placeholder="0" onChange={(event) => updateRow(row.id, { completed: Number(event.target.value) })} />
                <span className={`completion-difference ${row.completed - (dailyComparisonByDate.get(row.date)?.availableCodes.length ?? row.available) > 0 ? 'is-over' : ''}`}>{row.completed - (dailyComparisonByDate.get(row.date)?.availableCodes.length ?? row.available) > 0 ? `+${row.completed - (dailyComparisonByDate.get(row.date)?.availableCodes.length ?? row.available)}` : row.completed - (dailyComparisonByDate.get(row.date)?.availableCodes.length ?? row.available)}</span>
                <button className="delete-row" aria-label="刪除這一列" onClick={() => setRows((current) => current.filter((item) => item.id !== row.id))}>×</button>
              </div>
            ))}
          </div>
          <button className="add-row" onClick={() => setRows((current) => [...current, { id: makeId(), date: '', received: 0, completed: 0 }])}>＋ 新增一天</button>
        </div>

        {settings.importedCompanies.length > 0 && <details className="company-list-card"><summary>查看已匯入的 {compactNumber(settings.importedCompanies.length)} 家唯一公司</summary><div className="company-list">{settings.importedCompanies.map((company) => <div key={company.id}><span>{company.code || '—'}</span><strong>{company.name}</strong><small>{company.announcementDate ? `公告 ${company.announcementDate}` : '尚未公告'}{company.completionDate ? `・完成 ${company.completionDate}` : ''}</small></div>)}</div></details>}

        <div className="section-heading"><div><span>STEP 4</span><h2>設定剩餘工作</h2></div></div>
        <div className="form-card two-fields">
          <NumberField label="期初未完成" value={settings.openingBacklog} onChange={(value) => setSettings({ ...settings, openingBacklog: value })} suffix="件" hint="第一筆實績前就存在的待辦" />
          <NumberField label="預計後續還會收到" value={settings.expectedRemaining} onChange={(value) => setSettings({ ...settings, expectedRemaining: value })} suffix="件" hint="不知道可先填 0，再做保守情境" />
        </div>

      </section>

      <aside className="result-column">
        <div className={`verdict live-verdict ${result ? (result.feasible ? 'is-good' : 'is-risk') : ''}`}>
          {!result ? <EmptyResult text="輸入至少一天的收到量與完成量，就會即時重算。" /> : (
            <>
              <span className="verdict-kicker">即時完工結論</span>
              <div className="verdict-icon" aria-hidden="true">{result.feasible ? '✓' : '!'}</div>
              <h2>{result.feasible ? '照目前速度做得完' : '照目前速度會逾期'}</h2>
              <p>{result.feasible ? `估計 ${result.finishDate ? formatDate(result.finishDate) : '期限前'} 清完，仍有緩衝。` : `目前推估要到 ${result.finishDate ? formatDate(result.finishDate) : '一年後'} 才能清完。`}</p>
              <div className="capacity-line"><span>待處理總量</span><strong>{compactNumber(result.outstanding)} 件</strong></div>
            </>
          )}
        </div>

        {result && (
          <>
            <div className="metrics-grid">
              <Metric label="目前待辦" value={`${compactNumber(result.backlog)} 件`} note={`累計收到 ${compactNumber(result.received)} 件`} />
              <Metric label="剩餘工作天" value={`${result.daysLeft} 天`} note={`可產出 ${compactNumber(result.remainingCapacity)} 件`} />
              <Metric label="每日需完成" value={Number.isFinite(result.neededDaily) ? `${compactNumber(result.neededDaily, 1)} 件` : '已無時間'} note={`目前產能 ${compactNumber(result.dailyCapacity, 1)} 件`} />
              <Metric label="最少人力" value={result.requiredPeople ? `${result.requiredPeople} 人` : '無法估算'} note={`目前 ${result.teamCount} 人`} />
            </div>

            <div className="progress-card">
              <div className="card-title"><div><span>到目前為止</span><h3>實際完成進度</h3></div><strong>{result.received > 0 ? compactNumber((result.completed / (settings.openingBacklog + result.received)) * 100) : 0}%</strong></div>
              <div className="progress-track"><i style={{ width: `${Math.min(100, result.received > 0 ? (result.completed / (settings.openingBacklog + result.received)) * 100 : 0)}%` }} /></div>
              <div className="progress-notes"><span>已完成 <b>{compactNumber(result.completed)} 件</b></span><span>實際每人日均 <b>{compactNumber(result.observedPerPerson, 1)} 件</b></span></div>
            </div>

            <div className="chart-card">
              <div className="card-title"><div><span>每日實績</span><h3>收到、完成與待辦</h3></div><button className="text-button" onClick={() => downloadCSV('忙季即時進度.csv', ['日期', '收到', '完成', '日末待辦'], result.points.map((row) => [row.date, row.received, row.completed, row.backlog]))}>匯出結果</button></div>
              <div className="bar-chart trend-chart" aria-label="每日收到、完成與待辦圖表">
                {result.points.slice(-18).map((point) => (
                  <div className="chart-day" key={point.id} title={`${formatDate(point.date)}｜收到 ${point.received}、完成 ${point.completed}、待辦 ${point.backlog}`}>
                    <div className="bars">
                      <i className="demand-bar" style={{ height: `${(point.received / maxTrend) * 100}%` }} />
                      <i className="capacity-bar" style={{ height: `${(point.completed / maxTrend) * 100}%` }} />
                      <i className="backlog-bar" style={{ height: `${(point.backlog / maxTrend) * 100}%` }} />
                    </div>
                    <small>{parseDate(point.date).getUTCDate()}</small>
                  </div>
                ))}
              </div>
              <div className="legend bottom-legend"><i className="demand-dot" />收到 <i className="capacity-dot" />完成 <i className="backlog-dot" />待辦</div>
            </div>
          </>
        )}
      </aside>
    </div>
  );
}

function Landing({ onSelect }: { onSelect: (mode: Exclude<Mode, null>) => void }) {
  return (
    <main className="landing">
      <nav className="topbar" aria-label="主要導覽">
        <a className="brand" href="#top" aria-label="忙季算盤首頁"><span className="brand-mark">季</span><span>忙季算盤</span></a>
        <span className="local-badge">資料只存在這台裝置</span>
      </nav>
      <section className="hero" id="top">
        <div className="mode-grid" aria-label="選擇計算方式">
          <button className="mode-card forecast-card" onClick={() => onSelect('forecast')}><span className="card-number">01</span><span className="card-icon" aria-hidden="true">↗</span><span className="card-copy"><strong>我要預估</strong><small>忙季開始前</small><span>把上一季每日財報量換算到這一季，檢查現有人力是否足夠。</span></span><span className="card-action">開始規劃 <b>→</b></span></button>
          <button className="mode-card live-card" onClick={() => onSelect('live')}><span className="card-number">02</span><span className="card-icon" aria-hidden="true">●</span><span className="card-copy"><strong>即時計算</strong><small>忙季進行中</small><span>輸入每天實際收到與完成的數量，隨時確認是否能在期限前清完。</span></span><span className="card-action">更新進度 <b>→</b></span></button>
        </div>
      </section>
      <footer className="landing-footer"><span>免登入・免上傳・免費使用</span><span>所有計算都在你的瀏覽器中完成</span></footer>
    </main>
  );
}

export default function Home() {
  const [mode, setMode] = useState<Mode>(null);

  if (!mode) return <Landing onSelect={setMode} />;

  return (
    <main className="app-shell">
      <header className="workspace-header">
        <button className="brand brand-button" onClick={() => setMode(null)}><span className="brand-mark">季</span><span>忙季算盤</span></button>
        <div className="mode-tabs" role="tablist" aria-label="切換計算模式">
          <button className={mode === 'forecast' ? 'active' : ''} onClick={() => setMode('forecast')} role="tab" aria-selected={mode === 'forecast'}>人力預估</button>
          <button className={mode === 'live' ? 'active' : ''} onClick={() => setMode('live')} role="tab" aria-selected={mode === 'live'}>即時計算</button>
        </div>
        <button className="back-button" onClick={() => setMode(null)}>← 回首頁</button>
      </header>
      <section className="workspace-intro">
        <span className="eyebrow">{mode === 'forecast' ? 'FORECAST' : 'LIVE TRACKING'}</span>
        <h1>{mode === 'forecast' ? '忙季人力負荷預估' : '忙季即時完工計算'}</h1>
        <p>{mode === 'forecast' ? '用歷史量的 T 日分布，檢查本季人力能不能在期限內消化。' : '每天更新收到與完成的件數，立即重算剩餘量與預估清完日。'}</p>
      </section>
      {mode === 'forecast' ? <ForecastWorkspace /> : <LiveWorkspace />}
      <footer className="app-footer"><span>資料已自動保存在這台裝置</span><button onClick={() => { if (window.confirm('確定要清除兩種模式的所有資料嗎？')) { localStorage.removeItem(FORECAST_KEY); localStorage.removeItem(LIVE_KEY); window.location.reload(); } }}>清除本機資料</button></footer>
    </main>
  );
}
