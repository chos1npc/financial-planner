'use client';

import { ChangeEvent, useEffect, useMemo, useRef, useState } from 'react';

type Mode = 'forecast' | 'live' | null;
type TeamMode = 'uniform' | 'individual';
type TeamMember = { id: string; name: string; dailyBooks: number };
type HistoricalRow = { id: string; date: string; quantity: number };
type LiveRow = { id: string; date: string; received: number; completed: number };
type ForecastSettings = {
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
  completionDate: string;
  people: number;
  speed: number;
  efficiency: number;
  teamMode: TeamMode;
  members: TeamMember[];
  openingBacklog: number;
  expectedRemaining: number;
};

const DAY = 86_400_000;
const FORECAST_KEY = 'busy-season-forecast-v1';
const LIVE_KEY = 'busy-season-live-v1';

const initialForecastSettings: ForecastSettings = {
  historicalAnchor: '2025-11-14',
  currentAnchor: '2026-11-16',
  completionDate: '2026-11-18',
  people: 12,
  speed: 10,
  efficiency: 85,
  teamMode: 'uniform',
  members: [],
};

const initialLiveSettings: LiveSettings = {
  completionDate: '2026-09-04',
  people: 8,
  speed: 9,
  efficiency: 85,
  teamMode: 'uniform',
  members: [],
  openingBacklog: 0,
  expectedRemaining: 150,
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
          setSettings({ ...initialForecastSettings, ...parsed.settings, members: parsed.settings.members ?? [] });
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
        <div className="section-heading">
          <div><span>STEP 1</span><h2>設定換算基準</h2></div>
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
          <div><span>STEP 2</span><h2>輸入歷史每日財報量</h2></div>
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

        <div className="section-heading"><div><span>STEP 3</span><h2>設定本季人力</h2></div></div>
        <TeamCapacityEditor settings={settings} onChange={(patch) => setSettings((current) => ({ ...current, ...patch }))} />
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
  const [rows, setRows] = useState<LiveRow[]>([
    { id: 'live-blank-1', date: '', received: 0, completed: 0 },
    { id: 'live-blank-2', date: '', received: 0, completed: 0 },
    { id: 'live-blank-3', date: '', received: 0, completed: 0 },
  ]);
  const [hydrated, setHydrated] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const frame = window.requestAnimationFrame(() => {
      const saved = localStorage.getItem(LIVE_KEY);
      if (saved) {
        try {
          const parsed = JSON.parse(saved) as { settings: Partial<LiveSettings>; rows: LiveRow[] };
          setSettings({ ...initialLiveSettings, ...parsed.settings, members: parsed.settings.members ?? [] });
          setRows(parsed.rows);
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
  const result = useMemo(() => {
    if (!validRows.length || !settings.completionDate) return null;
    const received = validRows.reduce((sum, row) => sum + row.received, 0);
    const completed = validRows.reduce((sum, row) => sum + row.completed, 0);
    const backlog = Math.max(0, settings.openingBacklog + received - completed);
    const outstanding = backlog + settings.expectedRemaining;
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
    const actualDays = new Set(validRows.filter((row) => isWorkday(parseDate(row.date))).map((row) => row.date)).size;
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

  function importRows(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const lines = parseCSV(String(reader.result));
      const data = lines[0]?.some((cell) => /date|日期/i.test(cell)) ? lines.slice(1) : lines;
      const imported = data
        .filter((line) => /^\d{4}-\d{2}-\d{2}$/.test(line[0]))
        .map((line) => ({ id: makeId(), date: line[0], received: Number(line[1]) || 0, completed: Number(line[2]) || 0 }));
      if (imported.length) setRows(imported);
    };
    reader.readAsText(file);
    event.target.value = '';
  }

  const maxTrend = result ? Math.max(1, ...result.points.flatMap((point) => [point.received, point.completed, point.backlog])) : 1;

  return (
    <div className="workspace-grid">
      <section className="input-column">
        <div className="section-heading"><div><span>STEP 1</span><h2>更新每日實績</h2></div><div className="inline-actions"><button className="text-button" onClick={() => setRows(liveExample)}>載入範例</button><button className="text-button" onClick={() => fileRef.current?.click()}>匯入 CSV</button><input ref={fileRef} className="sr-only" type="file" accept=".csv,text/csv" onChange={importRows} /></div></div>
        <div className="data-table-card">
          <div className="data-table live-table">
            <div className="table-row table-header"><span>日期</span><span>實際收到</span><span>實際完成</span><span></span></div>
            {rows.map((row) => (
              <div className="table-row" key={row.id}>
                <input aria-label="實績日期" type="date" value={row.date} onChange={(event) => updateRow(row.id, { date: event.target.value })} />
                <input aria-label="實際收到量" type="number" min="0" value={row.received || ''} placeholder="0" onChange={(event) => updateRow(row.id, { received: Number(event.target.value) })} />
                <input aria-label="實際完成量" type="number" min="0" value={row.completed || ''} placeholder="0" onChange={(event) => updateRow(row.id, { completed: Number(event.target.value) })} />
                <button className="delete-row" aria-label="刪除這一列" onClick={() => setRows((current) => current.filter((item) => item.id !== row.id))}>×</button>
              </div>
            ))}
          </div>
          <button className="add-row" onClick={() => setRows((current) => [...current, { id: makeId(), date: '', received: 0, completed: 0 }])}>＋ 新增一天</button>
        </div>

        <div className="section-heading"><div><span>STEP 2</span><h2>設定剩餘工作</h2></div></div>
        <div className="form-card three-fields">
          <NumberField label="期初未完成" value={settings.openingBacklog} onChange={(value) => setSettings({ ...settings, openingBacklog: value })} suffix="件" hint="第一筆實績前就存在的待辦" />
          <NumberField label="預計後續還會收到" value={settings.expectedRemaining} onChange={(value) => setSettings({ ...settings, expectedRemaining: value })} suffix="件" hint="不知道可先填 0，再做保守情境" />
          <DateField label="要求清完日" value={settings.completionDate} onChange={(value) => setSettings({ ...settings, completionDate: value })} />
        </div>

        <div className="section-heading"><div><span>STEP 3</span><h2>設定目前人力</h2></div></div>
        <TeamCapacityEditor settings={settings} onChange={(patch) => setSettings((current) => ({ ...current, ...patch }))} />
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
