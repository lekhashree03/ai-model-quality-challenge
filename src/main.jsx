import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import {
  Upload,
  RefreshCw,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  FileSpreadsheet,
  Gauge,
  Users,
  Wrench,
  Trophy,
  LineChart,
  Info,
  SearchCheck,
} from 'lucide-react';
import './styles.css';

const CUSTOMER_TARGETS = { throughput: 250000, ttft: 0.5, rpm: 500 };

const PROFILE_USE_CASES = {
  1: { label: 'High-context chat / RAG Q&A', why: '10k input with short output favors retrieval-heavy assistant traffic.' },
  2: { label: 'Long-form generation', why: '10k input with 4k output stresses generation speed and user wait time.' },
  3: { label: 'Agent/tool loop', why: 'Moderate 3.2k input and 400 output resembles iterative agent calls.' },
  4: { label: 'Balanced RAG summarization', why: '1k input and 1k output tests balanced prompt plus generation.' },
  5: { label: 'Document QA / synthesis', why: '8k input and 1k output reflects document-grounded answer generation.' },
  6: { label: 'Very long-context retrieval', why: '60k input and 200 output is dominated by prefill / prompt processing.' },
  7: { label: 'Report generation from long context', why: '17k input and 3.5k output stresses both context and generation.' },
};

const MODEL_SIZE_BUCKETS = ['Small / fastest', 'Small-mid', 'Mid-size', 'Mid-large', 'Large', 'Very large', 'Largest / slowest'];

const fmt = (n, digits = 0) =>
  n === null || n === undefined || Number.isNaN(Number(n)) || !Number.isFinite(Number(n))
    ? '—'
    : Number(n).toLocaleString(undefined, { maximumFractionDigits: digits });
const sec = n =>
  n === null || n === undefined || !Number.isFinite(Number(n)) ? '—' : `${Number(n).toFixed(2)}s`;
const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const num = v => (v === null || v === undefined || v === '' ? null : Number(v));
const validNumber = v => {
  const n = num(v);
  return Number.isFinite(n) ? n : null;
};
const get = (row, names) => {
  const map = Object.fromEntries(Object.keys(row).map(k => [norm(k), row[k]]));
  for (const name of names) if (map[norm(name)] !== undefined && map[norm(name)] !== '') return map[norm(name)];
  return null;
};

function inferModelProfile(fileName, path = '') {
  const text = `${path}/${fileName}`.replace(/_/g, ' ');
  const model = text.match(/Model\s+([A-Z0-9]+)/i)?.[1]?.toUpperCase() || fileName.replace(/\.xlsx$/i, '');
  const profile = text.match(/profile\s*(\d+)/i)?.[1] || 'unknown';
  return { model: `Model ${model}`, modelKey: model, profile: `Profile ${profile}`, profileNum: Number(profile) || 0 };
}

function rowsFromSheet(sheet, expectedColumns = []) {
  if (!sheet) return [];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
  if (!raw.length) return [];

  let bestIndex = -1;
  let bestScore = -1;
  raw.forEach((row, idx) => {
    const normalizedCells = row.map(norm);
    const score = expectedColumns.reduce((acc, col) => acc + (normalizedCells.includes(norm(col)) ? 1 : 0), 0);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = idx;
    }
  });

  const headerIndex = bestScore > 0 ? bestIndex : raw.findIndex(r => r.some(c => String(c).trim() !== ''));
  if (headerIndex < 0) return [];

  const headers = raw[headerIndex].map((h, i) => String(h || `Column ${i + 1}`).trim());
  return raw
    .slice(headerIndex + 1)
    .filter(row => row.some(c => c !== '' && c !== null && c !== undefined))
    .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ''])));
}

function evaluateSweep(best, lowestTTFT) {
  if (!best) return { verdict: 'No data', checks: [] };
  const checks = [
    { key: 'throughput', label: `Throughput ≥ ${fmt(CUSTOMER_TARGETS.throughput)} tok/s`, pass: (best.throughput ?? 0) >= CUSTOMER_TARGETS.throughput, value: fmt(best.throughput) },
    { key: 'ttft', label: `TTFT ≤ ${CUSTOMER_TARGETS.ttft}s`, pass: (lowestTTFT?.ttft ?? Infinity) <= CUSTOMER_TARGETS.ttft, value: sec(lowestTTFT?.ttft) },
    { key: 'rpm', label: `RPM ≥ ${fmt(CUSTOMER_TARGETS.rpm)}`, pass: (best.rpm ?? 0) >= CUSTOMER_TARGETS.rpm, value: fmt(best.rpm, 1) },
  ];
  const passed = checks.filter(c => c.pass).length;
  const verdict = passed === checks.length ? 'Go' : passed >= 2 || (best.throughput ?? 0) >= CUSTOMER_TARGETS.throughput * 0.75 ? 'Review' : 'No-Go';
  return { verdict, checks };
}

function parseWorkbook(buffer, fileName, path = '') {
  const wb = XLSX.read(buffer, { type: 'array' });
  const meta = inferModelProfile(fileName, path);
  const summaryName = wb.SheetNames.find(s => norm(s) === 'summary') || wb.SheetNames.find(s => norm(s).includes('summary')) || wb.SheetNames[0];

  const summaryRaw = rowsFromSheet(wb.Sheets[summaryName], [
    'Input Length',
    'Output Length',
    'Cache %',
    'G Method',
    'Target Prompt G',
    'Batch Size',
    'Max S',
    'Target Max S',
    'Concurrency',
    'Throughput (t/s)',
    'TTFT (sec)',
    'Gen Speed (t/s/user)',
    'RPM',
  ]);

  let lastInput = null,
    lastOutput = null,
    lastCache = null;
  const summaryRows = summaryRaw
    .map((r, i) => {
      const input = validNumber(get(r, ['Input Length', 'Input Tokens', 'Context Length']));
      const output = validNumber(get(r, ['Output Length', 'Output Tokens', 'Generated Tokens']));
      const cache = validNumber(get(r, ['Cache %', 'Cache']));
      if (input !== null) lastInput = input;
      if (output !== null) lastOutput = output;
      if (cache !== null) lastCache = cache;
      return {
        rowId: i + 1,
        inputLength: input ?? lastInput,
        outputLength: output ?? lastOutput,
        cachePct: cache ?? lastCache,
        gMethod: get(r, ['G Method', 'Method']) || '—',
        targetPromptG: validNumber(get(r, ['Target Prompt G', 'Target G', 'G'])),
        batchSize: validNumber(get(r, ['Batch Size', 'Batch'])),
        maxS: validNumber(get(r, ['Max S', 'Maximum S'])),
        targetMaxS: validNumber(get(r, ['Target Max S'])),
        concurrency: validNumber(get(r, ['Concurrency'])),
        promptOnlyThroughput: validNumber(get(r, ['Prompt only Throughput (t/s)', 'Prompt Throughput'])),
        genOnlyThroughput: validNumber(get(r, ['Gen only Throughput (t/s)', 'Generation Throughput'])),
        throughput: validNumber(get(r, ['Throughput (t/s)', 'Throughput Mean', 'Throughput'])),
        throughputPerBox: validNumber(get(r, ['Throughput / box (t/s/csx)', 'Throughput per box'])),
        uncachedThroughput: validNumber(get(r, ['Uncached Throughput (t/s)', 'Uncached Throughput Mean'])),
        uncachedThroughputPerBox: validNumber(get(r, ['Uncached Throughput / box (t/s/csx)'])),
        cachedThroughput: validNumber(get(r, ['Cached Throughput (t/s)', 'Cached Throughput Mean'])),
        cachedThroughputPerBox: validNumber(get(r, ['Cached Throughput / box (t/s/csx)'])),
        ttft: validNumber(get(r, ['TTFT (sec)', 'TTFT Mean', 'TTFT'])),
        realPromptSpeed: validNumber(get(r, ['Real Prompt Speed (t/s/user)', 'Real Prompt Speed Mean'])),
        promptQueueSpeed: validNumber(get(r, ['Prompt Speed with Queueing (t/s/user)', 'Prompt Speed with Queueing Mean'])),
        genSpeed: validNumber(get(r, ['Gen Speed (t/s/user)', 'Real Gen Speed Mean'])),
        rpm: validNumber(get(r, ['RPM', 'Requests Per Minute'])),
        raw: r,
      };
    })
    .filter(r => r.throughput !== null || r.ttft !== null || r.rpm !== null || r.concurrency !== null);

  const scenarioSheets = wb.SheetNames.filter(n => n !== summaryName && /^sim_/i.test(n)).map(name => {
    const rows = rowsFromSheet(wb.Sheets[name], ['Concurrency', 'Throughput Mean', 'TTFT Mean', 'Real Gen Speed Mean'])
      .map(r => ({
        concurrency: validNumber(get(r, ['Concurrency'])),
        throughputMean: validNumber(get(r, ['Throughput Mean', 'Throughput (t/s)', 'Throughput'])),
        ttftMean: validNumber(get(r, ['TTFT Mean', 'TTFT (sec)', 'TTFT'])),
        genSpeedMean: validNumber(get(r, ['Real Gen Speed Mean', 'Gen Speed (t/s/user)'])),
        raw: r,
      }))
      .filter(r => r.concurrency !== null || r.throughputMean !== null || r.ttftMean !== null);
    return { name, rows };
  });

  const best = [...summaryRows].filter(r => r.throughput !== null).sort((a, b) => (b.throughput || 0) - (a.throughput || 0))[0] || null;
  const lowestTTFT = [...summaryRows].filter(r => r.ttft !== null).sort((a, b) => a.ttft - b.ttft)[0] || null;
  const bestRpm = [...summaryRows].filter(r => r.rpm !== null).sort((a, b) => (b.rpm || 0) - (a.rpm || 0))[0] || null;
  const evaluation = evaluateSweep(best, lowestTTFT);

  return {
    id: `${meta.modelKey}-${meta.profileNum}-${fileName}-${Date.now()}-${Math.random()}`,
    fileName,
    path,
    ...meta,
    summaryName,
    summaryRows,
    scenarioSheets,
    best,
    lowestTTFT,
    bestRpm,
    verdict: evaluation.verdict,
    checks: evaluation.checks,
  };
}

async function loadDefaultFiles() {
  const res = await fetch('/perf_data/manifest.json', { cache: 'no-store' });
  if (!res.ok) return [];
  const manifest = await res.json();
  const loaded = [];
  for (const url of manifest.files || []) {
    try {
      const r = await fetch(url);
      const b = await r.arrayBuffer();
      loaded.push(parseWorkbook(b, url.split('/').pop(), url));
    } catch (e) {
      console.warn('Default file failed', url, e);
    }
  }
  return loaded;
}

function Kpi({ label, value, sub }) {
  return (
    <div className="kpi">
      <div className="kpiLabel">{label}</div>
      <div className="kpiValue">{value}</div>
      {sub && <div className="kpiSub">{sub}</div>}
    </div>
  );
}
function Verdict({ v }) {
  const cls = v === 'Go' ? 'go' : v === 'Review' ? 'review' : 'no';
  const Icon = v === 'Go' ? CheckCircle2 : v === 'Review' ? AlertTriangle : XCircle;
  return (
    <span className={`verdict ${cls}`}>
      <Icon size={16} />
      {v}
    </span>
  );
}
function Bar({ value, max, label }) {
  const w = max ? Math.max(4, Math.min(100, (value / max) * 100)) : 0;
  return (
    <div className="barrow">
      <span>{label}</span>
      <div className="bar">
        <i style={{ width: `${w}%` }} />
      </div>
      <b>{fmt(value)}</b>
    </div>
  );
}
function CheckList({ checks }) {
  return (
    <div className="checks">
      {checks.map(c => (
        <div className={`check ${c.pass ? 'pass' : 'fail'}`} key={c.key}>
          {c.pass ? <CheckCircle2 size={14} /> : <XCircle size={14} />}
          <span>{c.label}</span>
          <b>{c.value}</b>
        </div>
      ))}
    </div>
  );
}

function Leaderboard({ title, rows, metric }) {
  return (
    <div className="leader">
      <h3>{title}</h3>
      {rows.slice(0, 5).map((s, idx) => (
        <div className="leaderRow" key={`${title}-${s.id}`}>
          <b>#{idx + 1}</b>
          <span>{s.model} · {s.profile}</span>
          <strong>{metric(s)}</strong>
        </div>
      ))}
    </div>
  );
}

function InsightPanel({ sweeps }) {
  const withBest = sweeps.filter(s => s.best);
  const topThroughput = [...withBest].sort((a, b) => (b.best?.throughput || 0) - (a.best?.throughput || 0));
  const lowTtft = [...withBest].sort((a, b) => (a.lowestTTFT?.ttft ?? Infinity) - (b.lowestTTFT?.ttft ?? Infinity));
  const topRpm = [...withBest].sort((a, b) => (b.best?.rpm || 0) - (a.best?.rpm || 0));
  return (
    <section className="panel">
      <h2><Trophy /> First-class Comparison Rankings</h2>
      <p className="hint">These cards summarize which uploaded sweeps are strongest for throughput-heavy, latency-sensitive, and request-heavy workloads.</p>
      <div className="leaders">
        <Leaderboard title="Top throughput" rows={topThroughput} metric={s => `${fmt(s.best?.throughput)} tok/s`} />
        <Leaderboard title="Lowest TTFT" rows={lowTtft} metric={s => sec(s.lowestTTFT?.ttft)} />
        <Leaderboard title="Highest RPM" rows={topRpm} metric={s => fmt(s.best?.rpm, 1)} />
      </div>
    </section>
  );
}

function ScatterPlot({ sweeps }) {
  const points = sweeps.filter(s => s.best && s.lowestTTFT?.ttft !== null).map(s => ({
    id: s.id,
    label: `${s.modelKey}${s.profileNum}`,
    model: `${s.model} · ${s.profile}`,
    x: s.lowestTTFT?.ttft ?? 0,
    y: s.best?.throughput ?? 0,
    verdict: s.verdict,
  }));
  if (!points.length) return null;
  const width = 900, height = 320, pad = 52;
  const maxX = Math.max(...points.map(p => p.x), CUSTOMER_TARGETS.ttft, 1);
  const maxY = Math.max(...points.map(p => p.y), CUSTOMER_TARGETS.throughput, 1);
  const sx = x => pad + (x / maxX) * (width - pad * 2);
  const sy = y => height - pad - (y / maxY) * (height - pad * 2);
  const targetX = sx(CUSTOMER_TARGETS.ttft);
  const targetY = sy(CUSTOMER_TARGETS.throughput);
  return (
    <section className="panel">
      <h2><LineChart /> Throughput vs TTFT Decision Map</h2>
      <p className="hint">Upper-left is best: high throughput with low time-to-first-token. Dotted lines show the customer decision thresholds.</p>
      <div className="chartWrap">
        <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Throughput versus TTFT scatter plot">
          <line x1={pad} y1={height - pad} x2={width - pad} y2={height - pad} className="axis" />
          <line x1={pad} y1={pad} x2={pad} y2={height - pad} className="axis" />
          <line x1={targetX} y1={pad} x2={targetX} y2={height - pad} className="target" />
          <line x1={pad} y1={targetY} x2={width - pad} y2={targetY} className="target" />
          <text x={targetX + 6} y={pad + 14} className="axisLabel">TTFT target</text>
          <text x={pad + 8} y={targetY - 8} className="axisLabel">Throughput target</text>
          <text x={width / 2 - 40} y={height - 10} className="axisLabel">TTFT seconds</text>
          <text x={10} y={32} className="axisLabel">tok/s</text>
          {points.map(p => (
            <g key={p.id}>
              <circle cx={sx(p.x)} cy={sy(p.y)} r="6" className={`dot ${p.verdict === 'Go' ? 'dotGo' : p.verdict === 'Review' ? 'dotReview' : 'dotNo'}`}>
                <title>{p.model}: {fmt(p.y)} tok/s, {sec(p.x)}</title>
              </circle>
              <text x={sx(p.x) + 8} y={sy(p.y) + 4} className="pointLabel">{p.label}</text>
            </g>
          ))}
        </svg>
      </div>
    </section>
  );
}

function CustomerView({ sweeps }) {
  const ranked = useMemo(() => [...sweeps].filter(s => s.best).sort((a, b) => (b.best.throughput || 0) - (a.best.throughput || 0)), [sweeps]);
  const maxT = Math.max(...ranked.map(s => s.best?.throughput || 0), 1);
  return (
    <section className="panel">
      <h2><Users /> Customer / PM View</h2>
      <p className="hint">Go/no-go uses explicit targets: {fmt(CUSTOMER_TARGETS.throughput)} tok/s, TTFT ≤ {CUSTOMER_TARGETS.ttft}s, RPM ≥ {CUSTOMER_TARGETS.rpm}. Each card shows the best configuration found inside that sweep.</p>
      <div className="cards">
        {ranked.map(s => (
          <div className="card" key={s.id}>
            <div className="cardTop"><h3>{s.model} · {s.profile}</h3><Verdict v={s.verdict} /></div>
            <div className="grid4">
              <Kpi label="Best throughput" value={`${fmt(s.best?.throughput)} tok/s`} />
              <Kpi label="TTFT" value={sec(s.lowestTTFT?.ttft)} sub="lowest in sweep" />
              <Kpi label="RPM" value={fmt(s.best?.rpm, 1)} />
              <Kpi label="Context" value={`${fmt(s.best?.inputLength)} in / ${fmt(s.best?.outputLength)} out`} />
            </div>
            <CheckList checks={s.checks} />
            <Bar value={s.best?.throughput || 0} max={maxT} label="relative throughput" />
            <p className="recommend">Recommended config: batch {fmt(s.best?.batchSize)}, concurrency {fmt(s.best?.concurrency)}, {s.best?.gMethod} G={fmt(s.best?.targetPromptG)}. Cache assumption: {s.best?.cachePct !== null ? `${Math.round(s.best.cachePct * 100)}%` : '—'}.</p>
          </div>
        ))}
      </div>
    </section>
  );
}

function ComparisonView({ sweeps }) {
  const [sortKey, setSortKey] = useState('model');
  const rows = useMemo(() => {
    const base = [...sweeps].filter(s => s.best);
    const sorters = {
      model: (a, b) => a.modelKey.localeCompare(b.modelKey) || a.profileNum - b.profileNum,
      throughput: (a, b) => (b.best?.throughput || 0) - (a.best?.throughput || 0),
      ttft: (a, b) => (a.lowestTTFT?.ttft ?? Infinity) - (b.lowestTTFT?.ttft ?? Infinity),
      rpm: (a, b) => (b.best?.rpm || 0) - (a.best?.rpm || 0),
    };
    return base.sort(sorters[sortKey]);
  }, [sweeps, sortKey]);
  return (
    <section className="panel">
      <div className="sectionTop">
        <h2><Gauge /> Side-by-Side Comparison</h2>
        <div className="sorts">
          <button className={sortKey === 'model' ? 'active' : ''} onClick={() => setSortKey('model')}>Model order</button>
          <button className={sortKey === 'throughput' ? 'active' : ''} onClick={() => setSortKey('throughput')}>Throughput</button>
          <button className={sortKey === 'ttft' ? 'active' : ''} onClick={() => setSortKey('ttft')}>TTFT</button>
          <button className={sortKey === 'rpm' ? 'active' : ''} onClick={() => setSortKey('rpm')}>RPM</button>
        </div>
      </div>
      <div className="tableWrap"><table><thead><tr><th>Model</th><th>Profile</th><th>Verdict</th><th>Pass/Fail reason</th><th>Throughput tok/s</th><th>TTFT</th><th>RPM</th><th>Batch</th><th>Concurrency</th><th>Input/Output</th></tr></thead><tbody>{rows.map(s => <tr key={s.id}><td>{s.model}</td><td>{s.profile}</td><td><Verdict v={s.verdict} /></td><td>{s.checks.filter(c => c.pass).length}/3 targets pass</td><td>{fmt(s.best?.throughput)}</td><td>{sec(s.lowestTTFT?.ttft)}</td><td>{fmt(s.best?.rpm, 1)}</td><td>{fmt(s.best?.batchSize)}</td><td>{fmt(s.best?.concurrency)}</td><td>{fmt(s.best?.inputLength)} / {fmt(s.best?.outputLength)}</td></tr>)}</tbody></table></div>
    </section>
  );
}

function MiniBars({ rows }) {
  const max = Math.max(...rows.map(r => r.throughput || 0), 1);
  return (
    <div className="miniBars">
      {rows.map(r => (
        <div className="miniBar" key={r.rowId} title={`Batch ${fmt(r.batchSize)}, Conc ${fmt(r.concurrency)}: ${fmt(r.throughput)} tok/s`}>
          <span>{r.gMethod} {fmt(r.targetPromptG)}</span>
          <div className="bar"><i style={{ width: `${Math.max(4, ((r.throughput || 0) / max) * 100)}%` }} /></div>
        </div>
      ))}
    </div>
  );
}

function InternalView({ sweeps }) {
  return (
    <section className="panel">
      <h2><Wrench /> Internal Product / Deployment Engineer View</h2>
      <p className="hint">Engineering sanity checks expose raw Summary rows, simulation sheet count, throughput spread, and TTFT anomalies before numbers reach a customer.</p>
      <div className="cards">
        {sweeps.map(s => {
          const anomalies = [];
          const ttfts = s.summaryRows.map(r => r.ttft).filter(v => v !== null);
          const th = s.summaryRows.map(r => r.throughput).filter(Boolean);
          if (ttfts.some(v => v > 1)) anomalies.push('TTFT > 1s');
          if (th.length > 1 && Math.min(...th) / Math.max(...th) < 0.25) anomalies.push('Large throughput spread');
          if (!s.summaryRows.length) anomalies.push('No Summary rows parsed');
          return (
            <div className="card" key={s.id}>
              <div className="cardTop"><h3>{s.model} · {s.profile}</h3><span className="pill">{s.scenarioSheets.length} sim sheets</span></div>
              <div className="grid4"><Kpi label="Rows" value={s.summaryRows.length} /><Kpi label="Max throughput" value={fmt(Math.max(...th, 0))} /><Kpi label="Min TTFT" value={sec(Math.min(...ttfts, Infinity))} /><Kpi label="Source" value={s.fileName} /></div>
              <div className="anomaly">{anomalies.length ? anomalies.map(a => <span className="warn" key={a}>{a}</span>) : <span className="ok">No obvious anomalies</span>}</div>
              <MiniBars rows={s.summaryRows} />
              <div className="miniTable"><table><thead><tr><th>G</th><th>Batch</th><th>Conc.</th><th>Throughput</th><th>TTFT</th><th>Gen/user</th></tr></thead><tbody>{s.summaryRows.map(r => <tr key={r.rowId}><td>{r.gMethod} {fmt(r.targetPromptG)}</td><td>{fmt(r.batchSize)}</td><td>{fmt(r.concurrency)}</td><td>{fmt(r.throughput)}</td><td>{sec(r.ttft)}</td><td>{fmt(r.genSpeed, 1)}</td></tr>)}</tbody></table></div>
            </div>
          );
        })}
      </div>
    </section>
  );
}

function AnalysisSection({ sweeps }) {
  const profiles = Object.entries(PROFILE_USE_CASES);
  const modelStats = useMemo(() => {
    const grouped = new Map();
    sweeps.filter(s => s.best).forEach(s => {
      if (!grouped.has(s.modelKey)) grouped.set(s.modelKey, []);
      grouped.get(s.modelKey).push(s.best.throughput || 0);
    });
    const rows = [...grouped.entries()].map(([modelKey, vals]) => ({
      model: `Model ${modelKey}`,
      modelKey,
      avgThroughput: vals.reduce((a, b) => a + b, 0) / vals.length,
    })).sort((a, b) => b.avgThroughput - a.avgThroughput);
    return rows.map((r, idx) => ({ ...r, inferredSize: MODEL_SIZE_BUCKETS[Math.min(MODEL_SIZE_BUCKETS.length - 1, Math.floor((idx / Math.max(1, rows.length - 1)) * (MODEL_SIZE_BUCKETS.length - 1)))] }));
  }, [sweeps]);
  return (
    <section className="panel">
      <h2><SearchCheck /> Reviewer Analysis: Profiles and Relative Model Sizes</h2>
      <p className="hint">These are product-facing hypotheses based only on traffic shape and relative throughput. The app avoids hardcoding model letters, so a new Model L still parses and appears automatically.</p>
      <div className="analysisGrid">
        <div>
          <h3>Profile use-case interpretation</h3>
          <div className="smallRows">{profiles.map(([n, p]) => <div className="smallRow" key={n}><b>Profile {n}</b><span>{p.label}</span><em>{p.why}</em></div>)}</div>
        </div>
        <div>
          <h3>Relative model-size inference</h3>
          <div className="smallRows">{modelStats.map(m => <div className="smallRow" key={m.modelKey}><b>{m.model}</b><span>{m.inferredSize}</span><em>Average best throughput: {fmt(m.avgThroughput)} tok/s</em></div>)}</div>
        </div>
      </div>
    </section>
  );
}

function App() {
  const [sweeps, setSweeps] = useState([]);
  const [status, setStatus] = useState('Loading default perf_data manifest...');
  useEffect(() => {
    loadDefaultFiles()
      .then(d => {
        setSweeps(d);
        setStatus(d.length ? `Loaded ${d.length} default sweep(s). You can upload more.` : 'No default sweeps found. Upload .xlsx files to begin.');
      })
      .catch(() => setStatus('No default manifest found. Upload .xlsx files to begin.'));
  }, []);
  const onFiles = async files => {
    const parsed = [];
    for (const f of files) {
      if (!/\.xlsx$/i.test(f.name)) continue;
      parsed.push(parseWorkbook(await f.arrayBuffer(), f.name, f.webkitRelativePath || f.name));
    }
    setSweeps(prev => [...prev, ...parsed]);
    setStatus(`Added ${parsed.length} uploaded sweep(s).`);
  };
  return (
    <main>
      <header>
        <div><h1>Cerebras Performance Projection UI</h1><p>Default shipped sweeps + runtime upload for customer go/no-go and internal sanity checks.</p></div>
        <button onClick={() => location.reload()}><RefreshCw size={16} />Reload defaults</button>
      </header>
      <section className="upload"><FileSpreadsheet /><div><h2>Upload more perf sweeps</h2><p>Upload one or many `.xlsx` files. New models and profiles render automatically.</p><input type="file" multiple accept=".xlsx" onChange={e => onFiles([...e.target.files])} /></div><label className="drop">Choose files<Upload size={18} /><input type="file" multiple accept=".xlsx" onChange={e => onFiles([...e.target.files])} /></label></section>
      <p className="status">{status}</p>
      {sweeps.length ? <><InsightPanel sweeps={sweeps} /><ScatterPlot sweeps={sweeps} /><ComparisonView sweeps={sweeps} /><CustomerView sweeps={sweeps} /><InternalView sweeps={sweeps} /><AnalysisSection sweeps={sweeps} /></> : <section className="empty">No sweeps loaded yet.</section>}
      <footer><Info size={16} /> Built as a dynamic tool: default data loads from the manifest, and runtime uploads parse in-browser without rebuilds.</footer>
    </main>
  );
}

createRoot(document.getElementById('root')).render(<App />);
