import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import {
  AreaChart,
  Area,
  BarChart,
  Bar,
  CartesianGrid,
  Legend,
  ResponsiveContainer,
  ScatterChart,
  Scatter,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { Upload, Users, Wrench, AlertTriangle, CheckCircle2, XCircle, FileSpreadsheet } from 'lucide-react';
import './styles.css';

const METRIC_ALIASES = {
  throughput: ['throughput', 'tok/s', 'tokens/sec', 'tokens per second', 'output tps', 'gen speed', 'generation speed', 'generated tokens'],
  ttft: ['ttft', 'time to first token', 'first token'],
  latency: ['latency', 'end to end', 'e2e', 'total time', 'response time'],
  context: ['context', 'input tokens', 'prompt tokens', 'sequence length', 'seq len', 'ctx'],
  outputTokens: ['output tokens', 'completion tokens', 'generated tokens', 'max new tokens', 'decode tokens'],
  concurrency: ['concurrency', 'parallel', 'num users', 'users', 'batch', 'batch size', 'qps', 'rps', 'requests/sec'],
  cost: ['cost', 'price', '$', 'dollar'],
};

function normalizeKey(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[_\-\/()]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findColumn(headers, aliases) {
  const normalized = headers.map((h) => ({ original: h, norm: normalizeKey(h) }));
  return normalized.find((h) => aliases.some((a) => h.norm.includes(a)))?.original;
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const cleaned = String(value).replace(/[$,% commas]/g, '').replace(/,/g, '').trim();
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function inferModelAndProfile(fileName) {
  const clean = fileName.replace(/\\/g, '/');
  const modelMatch = clean.match(/Model[ _-]?([A-Z0-9]+)/i);
  const profileMatch = clean.match(/profile[ _-]?(\d+)/i);
  return {
    model: modelMatch ? `Model ${modelMatch[1].toUpperCase()}` : clean.replace(/\.xlsx$/i, ''),
    profile: profileMatch ? `Profile ${profileMatch[1]}` : 'Profile unknown',
  };
}

async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null });
  const headers = Object.keys(rows[0] || {});
  const columns = Object.fromEntries(
    Object.entries(METRIC_ALIASES).map(([metric, aliases]) => [metric, findColumn(headers, aliases)])
  );
  const inferred = inferModelAndProfile(file.webkitRelativePath || file.name);

  return rows.map((row, index) => ({
    id: `${file.name}-${index}`,
    sourceFile: file.name,
    sheetName,
    rowIndex: index + 2,
    model: String(row.model || row.Model || inferred.model),
    profile: String(row.profile || row.Profile || inferred.profile),
    throughput: toNumber(row[columns.throughput]),
    ttft: toNumber(row[columns.ttft]),
    latency: toNumber(row[columns.latency]),
    context: toNumber(row[columns.context]),
    outputTokens: toNumber(row[columns.outputTokens]),
    concurrency: toNumber(row[columns.concurrency]),
    cost: toNumber(row[columns.cost]),
    raw: row,
    mappedColumns: columns,
  }));
}

function average(values) {
  const valid = values.filter((v) => typeof v === 'number' && Number.isFinite(v));
  return valid.length ? valid.reduce((a, b) => a + b, 0) / valid.length : null;
}

function percentile(values, p) {
  const valid = values.filter((v) => typeof v === 'number' && Number.isFinite(v)).sort((a, b) => a - b);
  if (!valid.length) return null;
  return valid[Math.min(valid.length - 1, Math.floor((p / 100) * valid.length))];
}

function fmt(value, digits = 2) {
  if (value === null || value === undefined || Number.isNaN(value)) return 'N/A';
  return Number(value).toLocaleString(undefined, { maximumFractionDigits: digits });
}

function summarize(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    const key = `${row.model} | ${row.profile}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });

  return [...groups.entries()].map(([key, items]) => {
    const [model, profile] = key.split(' | ');
    const throughputValues = items.map((r) => r.throughput);
    const ttftValues = items.map((r) => r.ttft);
    const latencyValues = items.map((r) => r.latency);
    const contextValues = items.map((r) => r.context);
    const costValues = items.map((r) => r.cost);
    return {
      key,
      model,
      profile,
      rows: items.length,
      avgThroughput: average(throughputValues),
      p50Throughput: percentile(throughputValues, 50),
      p10Throughput: percentile(throughputValues, 10),
      avgTtft: average(ttftValues),
      p90Ttft: percentile(ttftValues, 90),
      avgLatency: average(latencyValues),
      maxContext: Math.max(...contextValues.filter((v) => v !== null), 0) || null,
      avgCost: average(costValues),
      efficiency: average(throughputValues) && average(costValues) ? average(throughputValues) / average(costValues) : null,
    };
  }).sort((a, b) => (b.avgThroughput || 0) - (a.avgThroughput || 0));
}

function anomalyRows(rows) {
  const tps = rows.map((r) => r.throughput).filter((v) => v !== null);
  const ttft = rows.map((r) => r.ttft).filter((v) => v !== null);
  const avgTps = average(tps);
  const avgTtft = average(ttft);
  return rows.filter((r) => {
    const badThroughput = avgTps && r.throughput !== null && r.throughput < avgTps * 0.35;
    const badTtft = avgTtft && r.ttft !== null && r.ttft > avgTtft * 2.5;
    const missingCore = r.throughput === null && r.ttft === null;
    return badThroughput || badTtft || missingCore;
  });
}

function UploadPanel({ onRows }) {
  const [status, setStatus] = useState('Upload one or many .xlsx perf sweeps.');

  async function handleFiles(files) {
    const xlsxFiles = [...files].filter((file) => file.name.toLowerCase().endsWith('.xlsx'));
    if (!xlsxFiles.length) {
      setStatus('No .xlsx files found.');
      return;
    }
    try {
      const parsed = await Promise.all(xlsxFiles.map(parseWorkbook));
      const rows = parsed.flat();
      onRows(rows);
      setStatus(`Parsed ${fmt(rows.length, 0)} rows from ${xlsxFiles.length} workbook(s).`);
    } catch (err) {
      console.error(err);
      setStatus(`Could not parse upload: ${err.message}`);
    }
  }

  return (
    <section className="upload-card">
      <div>
        <p className="eyebrow">Live upload flow</p>
        <h1>Performance Sweep Decision UI</h1>
        <p className="subtext">Upload multiple model/profile `.xlsx` sweeps and instantly compare customer fit and internal projection sanity.</p>
      </div>
      <label className="dropzone">
        <Upload size={28} />
        <span>Choose `.xlsx` files</span>
        <small>Supports multi-file selection and unseen models like Model L.</small>
        <input type="file" multiple accept=".xlsx" onChange={(e) => handleFiles(e.target.files)} />
      </label>
      <p className="status"><FileSpreadsheet size={16} /> {status}</p>
    </section>
  );
}

function CustomerView({ summaries }) {
  const [minTps, setMinTps] = useState(1000);
  const [maxTtft, setMaxTtft] = useState(1000);
  const [minContext, setMinContext] = useState(8192);

  const decisions = summaries.map((s) => {
    const passTps = s.avgThroughput !== null && s.avgThroughput >= minTps;
    const passTtft = s.avgTtft === null || s.avgTtft <= maxTtft;
    const passContext = s.maxContext === null || s.maxContext >= minContext;
    return { ...s, go: passTps && passTtft && passContext, passTps, passTtft, passContext };
  });

  return (
    <section className="panel">
      <div className="panel-header"><Users /><div><h2>Customer / PM View</h2><p>Workload fit, go/no-go signal, and metrics customers recognize.</p></div></div>
      <div className="controls">
        <label>Required throughput tok/s <input type="number" value={minTps} onChange={(e) => setMinTps(Number(e.target.value))} /></label>
        <label>Max TTFT ms <input type="number" value={maxTtft} onChange={(e) => setMaxTtft(Number(e.target.value))} /></label>
        <label>Min context tokens <input type="number" value={minContext} onChange={(e) => setMinContext(Number(e.target.value))} /></label>
      </div>
      <div className="cards">
        {decisions.map((s) => (
          <article className={`decision-card ${s.go ? 'go' : 'nogo'}`} key={s.key}>
            <div className="decision-title">
              {s.go ? <CheckCircle2 /> : <XCircle />}
              <div><h3>{s.model}</h3><p>{s.profile}</p></div>
            </div>
            <strong>{s.go ? 'GO: likely fast enough' : 'NO-GO / needs review'}</strong>
            <dl>
              <dt>Avg tok/s</dt><dd>{fmt(s.avgThroughput)}</dd>
              <dt>Avg TTFT</dt><dd>{fmt(s.avgTtft)} ms</dd>
              <dt>Max context</dt><dd>{fmt(s.maxContext, 0)}</dd>
              <dt>Efficiency proxy</dt><dd>{s.efficiency ? `${fmt(s.efficiency)} tok/s per cost unit` : 'N/A'}</dd>
            </dl>
          </article>
        ))}
      </div>
    </section>
  );
}

function InternalView({ rows, summaries }) {
  const anomalies = anomalyRows(rows);
  const scatter = rows.filter((r) => r.throughput !== null || r.ttft !== null).slice(0, 1000);

  return (
    <section className="panel">
      <div className="panel-header"><Wrench /><div><h2>Internal Product / Deployment View</h2><p>Projection sanity checks, config sensitivity, and anomaly spotting.</p></div></div>
      <div className="chart-grid">
        <div className="chart-card">
          <h3>Average throughput by sweep</h3>
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={summaries}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="key" hide />
              <YAxis />
              <Tooltip />
              <Legend />
              <Bar dataKey="avgThroughput" name="Avg tok/s" />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="chart-card">
          <h3>TTFT vs throughput</h3>
          <ResponsiveContainer width="100%" height={280}>
            <ScatterChart>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="ttft" name="TTFT" unit="ms" />
              <YAxis dataKey="throughput" name="tok/s" />
              <Tooltip cursor={{ strokeDasharray: '3 3' }} />
              <Scatter name="Rows" data={scatter} />
            </ScatterChart>
          </ResponsiveContainer>
        </div>
      </div>
      <div className="table-card">
        <h3><AlertTriangle size={18} /> Anomaly queue ({anomalies.length})</h3>
        <table>
          <thead><tr><th>Model</th><th>Profile</th><th>File</th><th>Row</th><th>Tok/s</th><th>TTFT</th><th>Latency</th></tr></thead>
          <tbody>
            {anomalies.slice(0, 50).map((r) => (
              <tr key={r.id}><td>{r.model}</td><td>{r.profile}</td><td>{r.sourceFile}</td><td>{r.rowIndex}</td><td>{fmt(r.throughput)}</td><td>{fmt(r.ttft)}</td><td>{fmt(r.latency)}</td></tr>
            ))}
            {!anomalies.length && <tr><td colSpan="7">No obvious anomalies detected by the default rules.</td></tr>}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function ComparisonView({ summaries }) {
  return (
    <section className="panel">
      <h2>Side-by-side comparison</h2>
      <div className="chart-card">
        <ResponsiveContainer width="100%" height={320}>
          <AreaChart data={summaries}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="key" hide />
            <YAxis />
            <Tooltip />
            <Legend />
            <Area dataKey="avgThroughput" name="Avg tok/s" fillOpacity={0.2} />
            <Area dataKey="p50Throughput" name="P50 tok/s" fillOpacity={0.2} />
          </AreaChart>
        </ResponsiveContainer>
      </div>
      <div className="table-card">
        <table>
          <thead><tr><th>Rank</th><th>Model</th><th>Profile</th><th>Rows</th><th>Avg tok/s</th><th>P10 tok/s</th><th>P90 TTFT</th><th>Max context</th></tr></thead>
          <tbody>{summaries.map((s, i) => <tr key={s.key}><td>{i + 1}</td><td>{s.model}</td><td>{s.profile}</td><td>{s.rows}</td><td>{fmt(s.avgThroughput)}</td><td>{fmt(s.p10Throughput)}</td><td>{fmt(s.p90Ttft)} ms</td><td>{fmt(s.maxContext, 0)}</td></tr>)}</tbody>
        </table>
      </div>
    </section>
  );
}

function EmptyState() {
  return <section className="empty"><h2>Upload perf sweeps to begin</h2><p>The app parses files in the browser. No rebuild, no hard-coded model list, and no backend required for Vercel.</p></section>;
}

function App() {
  const [rows, setRows] = useState([]);
  const summaries = useMemo(() => summarize(rows), [rows]);

  return (
    <main>
      <UploadPanel onRows={setRows} />
      {rows.length ? <><ComparisonView summaries={summaries} /><CustomerView summaries={summaries} /><InternalView rows={rows} summaries={summaries} /></> : <EmptyState />}
    </main>
  );
}

createRoot(document.getElementById('root')).render(<App />);
