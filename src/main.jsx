import React, {useEffect, useMemo, useState} from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import { Upload, RefreshCw, CheckCircle2, AlertTriangle, XCircle, FileSpreadsheet, Gauge, Users, Wrench } from 'lucide-react';
import './styles.css';

const CUSTOMER_TARGETS = { throughput: 250000, ttft: 0.5, rpm: 500 };
const fmt = (n, digits=0) => n === null || n === undefined || Number.isNaN(Number(n)) ? '—' : Number(n).toLocaleString(undefined,{maximumFractionDigits:digits});
const sec = n => n === null || n === undefined || Number.isNaN(Number(n)) ? '—' : `${Number(n).toFixed(2)}s`;
const norm = s => String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const get = (row, names) => {
  const map = Object.fromEntries(Object.keys(row).map(k => [norm(k), row[k]]));
  for (const name of names) if (map[norm(name)] !== undefined && map[norm(name)] !== '') return map[norm(name)];
  return null;
};
const num = v => v === null || v === undefined || v === '' ? null : Number(v);

function inferModelProfile(fileName, path='') {
  const text = `${path}/${fileName}`.replace(/_/g,' ');
  const model = text.match(/Model\s+([A-Z0-9]+)/i)?.[1]?.toUpperCase() || fileName.replace(/\.xlsx$/i,'');
  const profile = text.match(/profile\s*(\d+)/i)?.[1] || 'unknown';
  return { model: `Model ${model}`, modelKey: model, profile: `Profile ${profile}`, profileNum: Number(profile) || 0 };
}

function parseWorkbook(buffer, fileName, path='') {
  const wb = XLSX.read(buffer, {type:'array'});
  const meta = inferModelProfile(fileName, path);
  const summaryName = wb.SheetNames.find(s => norm(s) === 'summary') || wb.SheetNames[0];
  const summaryRaw = XLSX.utils.sheet_to_json(wb.Sheets[summaryName], {defval:null, blankrows:false});
  const summaryRows = summaryRaw.map((r, i) => ({
    rowId: i + 1,
    inputLength: num(get(r, ['Input Length','Input Tokens','Context Length'])) ?? null,
    outputLength: num(get(r, ['Output Length','Output Tokens','Generated Tokens'])) ?? null,
    cachePct: num(get(r, ['Cache %','Cache'])) ?? null,
    gMethod: get(r, ['G Method','Method']) || '—',
    targetPromptG: num(get(r, ['Target Prompt G','Target G','G'])) ?? null,
    batchSize: num(get(r, ['Batch Size','Batch'])) ?? null,
    maxS: num(get(r, ['Max S','Maximum S'])) ?? null,
    targetMaxS: num(get(r, ['Target Max S'])) ?? null,
    concurrency: num(get(r, ['Concurrency'])) ?? null,
    promptOnlyThroughput: num(get(r, ['Prompt only Throughput (t/s)','Prompt Throughput'])) ?? null,
    genOnlyThroughput: num(get(r, ['Gen only Throughput (t/s)','Generation Throughput'])) ?? null,
    throughput: num(get(r, ['Throughput (t/s)','Throughput Mean','Throughput'])) ?? null,
    throughputPerBox: num(get(r, ['Throughput / box (t/s/csx)','Throughput per box'])) ?? null,
    uncachedThroughput: num(get(r, ['Uncached Throughput (t/s)','Uncached Throughput Mean'])) ?? null,
    cachedThroughput: num(get(r, ['Cached Throughput (t/s)','Cached Throughput Mean'])) ?? null,
    ttft: num(get(r, ['TTFT (sec)','TTFT Mean','TTFT'])) ?? null,
    realPromptSpeed: num(get(r, ['Real Prompt Speed (t/s/user)','Real Prompt Speed Mean'])) ?? null,
    promptQueueSpeed: num(get(r, ['Prompt Speed with Queueing (t/s/user)','Prompt Speed with Queueing Mean'])) ?? null,
    genSpeed: num(get(r, ['Gen Speed (t/s/user)','Real Gen Speed Mean'])) ?? null,
    rpm: num(get(r, ['RPM','Requests Per Minute'])) ?? null,
    raw: r
  })).filter(r => r.throughput || r.ttft || r.rpm || r.concurrency);

  const scenarioSheets = wb.SheetNames.filter(n => n !== summaryName && /^sim_/i.test(n)).map(name => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], {defval:null}).map(r => ({
      concurrency: num(get(r, ['Concurrency'])),
      throughputMean: num(get(r, ['Throughput Mean','Throughput (t/s)','Throughput'])),
      ttftMean: num(get(r, ['TTFT Mean','TTFT (sec)','TTFT'])),
      genSpeedMean: num(get(r, ['Real Gen Speed Mean','Gen Speed (t/s/user)'])),
      raw: r
    })).filter(r => r.concurrency !== null);
    return {name, rows};
  });

  const best = [...summaryRows].sort((a,b)=>(b.throughput||0)-(a.throughput||0))[0] || null;
  const lowestTTFT = [...summaryRows].filter(r=>r.ttft!==null).sort((a,b)=>a.ttft-b.ttft)[0] || null;
  const verdict = !best ? 'No data' : (best.throughput >= CUSTOMER_TARGETS.throughput && (lowestTTFT?.ttft ?? 99) <= CUSTOMER_TARGETS.ttft ? 'Go' : best.throughput >= CUSTOMER_TARGETS.throughput*0.75 ? 'Review' : 'No-Go');

  return { id: `${meta.modelKey}-${meta.profileNum}-${fileName}-${Date.now()}-${Math.random()}`, fileName, path, ...meta, summaryRows, scenarioSheets, best, lowestTTFT, verdict };
}

async function loadDefaultFiles(){
  const res = await fetch('/perf_data/manifest.json', {cache:'no-store'});
  if(!res.ok) return [];
  const manifest = await res.json();
  const loaded = [];
  for(const url of manifest.files || []){
    try{
      const r = await fetch(url);
      const b = await r.arrayBuffer();
      loaded.push(parseWorkbook(b, url.split('/').pop(), url));
    } catch(e){ console.warn('Default file failed', url, e); }
  }
  return loaded;
}

function Kpi({label,value,sub}){return <div className="kpi"><div className="kpiLabel">{label}</div><div className="kpiValue">{value}</div>{sub&&<div className="kpiSub">{sub}</div>}</div>}
function Verdict({v}){const cls=v==='Go'?'go':v==='Review'?'review':'no'; const Icon=v==='Go'?CheckCircle2:v==='Review'?AlertTriangle:XCircle; return <span className={`verdict ${cls}`}><Icon size={16}/>{v}</span>}
function Bar({value,max,label}){const w=max?Math.max(4, Math.min(100, value/max*100)):0; return <div className="barrow"><span>{label}</span><div className="bar"><i style={{width:`${w}%`}} /></div><b>{fmt(value)}</b></div>}

function CustomerView({sweeps}){
  const ranked = useMemo(()=>[...sweeps].filter(s=>s.best).sort((a,b)=>(b.best.throughput||0)-(a.best.throughput||0)),[sweeps]);
  const maxT = Math.max(...ranked.map(s=>s.best?.throughput||0), 1);
  return <section className="panel"><h2><Users/> Customer / PM View</h2><p className="hint">Decision-oriented summary using the best configuration found in each uploaded sweep. Targets are editable in code: {fmt(CUSTOMER_TARGETS.throughput)} tok/s, TTFT ≤ {CUSTOMER_TARGETS.ttft}s, RPM ≥ {CUSTOMER_TARGETS.rpm}.</p>
    <div className="cards">{ranked.map(s=><div className="card" key={s.id}>
      <div className="cardTop"><h3>{s.model} · {s.profile}</h3><Verdict v={s.verdict}/></div>
      <div className="grid4"><Kpi label="Best throughput" value={`${fmt(s.best?.throughput)} tok/s`}/><Kpi label="TTFT" value={sec(s.lowestTTFT?.ttft)} sub="lowest in sweep"/><Kpi label="RPM" value={fmt(s.best?.rpm,1)}/><Kpi label="Context" value={`${fmt(s.best?.inputLength)} in / ${fmt(s.best?.outputLength)} out`}/></div>
      <Bar value={s.best?.throughput||0} max={maxT} label="relative throughput" />
      <p className="recommend">Recommended config: batch {fmt(s.best?.batchSize)}, concurrency {fmt(s.best?.concurrency)}, {s.best?.gMethod} G={fmt(s.best?.targetPromptG)}. Cache assumption: {s.best?.cachePct!==null?`${Math.round(s.best.cachePct*100)}%`:'—'}.</p>
    </div>)}</div>
  </section>
}

function ComparisonView({sweeps}){
  const rows = [...sweeps].filter(s=>s.best).sort((a,b)=>(a.modelKey.localeCompare(b.modelKey)||a.profileNum-b.profileNum));
  return <section className="panel"><h2><Gauge/> Side-by-Side Comparison</h2><div className="tableWrap"><table><thead><tr><th>Model</th><th>Profile</th><th>Verdict</th><th>Throughput tok/s</th><th>TTFT</th><th>RPM</th><th>Batch</th><th>Concurrency</th><th>Input/Output</th></tr></thead><tbody>{rows.map(s=><tr key={s.id}><td>{s.model}</td><td>{s.profile}</td><td><Verdict v={s.verdict}/></td><td>{fmt(s.best?.throughput)}</td><td>{sec(s.lowestTTFT?.ttft)}</td><td>{fmt(s.best?.rpm,1)}</td><td>{fmt(s.best?.batchSize)}</td><td>{fmt(s.best?.concurrency)}</td><td>{fmt(s.best?.inputLength)} / {fmt(s.best?.outputLength)}</td></tr>)}</tbody></table></div></section>
}

function InternalView({sweeps}){
  return <section className="panel"><h2><Wrench/> Internal Product / Deployment Engineer View</h2><div className="cards">{sweeps.map(s=>{
    const anomalies = [];
    const ttfts = s.summaryRows.map(r=>r.ttft).filter(v=>v!==null);
    const th = s.summaryRows.map(r=>r.throughput).filter(Boolean);
    if(ttfts.some(v=>v>1)) anomalies.push('TTFT > 1s');
    if(th.length>1 && Math.min(...th)/Math.max(...th)<0.25) anomalies.push('Large throughput spread');
    if(!s.summaryRows.length) anomalies.push('No Summary rows parsed');
    return <div className="card" key={s.id}><div className="cardTop"><h3>{s.model} · {s.profile}</h3><span className="pill">{s.scenarioSheets.length} sim sheets</span></div>
      <div className="grid4"><Kpi label="Rows" value={s.summaryRows.length}/><Kpi label="Max throughput" value={fmt(Math.max(...th,0))}/><Kpi label="Min TTFT" value={sec(Math.min(...ttfts,Infinity))}/><Kpi label="Source" value={s.fileName}/></div>
      <div className="anomaly">{anomalies.length?anomalies.map(a=><span className="warn" key={a}>{a}</span>):<span className="ok">No obvious anomalies</span>}</div>
      <div className="miniTable"><table><thead><tr><th>G</th><th>Batch</th><th>Conc.</th><th>Throughput</th><th>TTFT</th><th>Gen/user</th></tr></thead><tbody>{s.summaryRows.map(r=><tr key={r.rowId}><td>{r.gMethod} {fmt(r.targetPromptG)}</td><td>{fmt(r.batchSize)}</td><td>{fmt(r.concurrency)}</td><td>{fmt(r.throughput)}</td><td>{sec(r.ttft)}</td><td>{fmt(r.genSpeed,1)}</td></tr>)}</tbody></table></div>
    </div>})}</div></section>
}

function App(){
  const [sweeps,setSweeps]=useState([]); const [status,setStatus]=useState('Loading default perf_data manifest...');
  useEffect(()=>{loadDefaultFiles().then(d=>{setSweeps(d); setStatus(d.length?`Loaded ${d.length} default sweep(s). You can upload more.`:'No default sweeps found. Upload .xlsx files to begin.');}).catch(()=>setStatus('No default manifest found. Upload .xlsx files to begin.'));},[]);
  const onFiles = async (files) => {
    const parsed=[];
    for(const f of files){ if(!/\.xlsx$/i.test(f.name)) continue; parsed.push(parseWorkbook(await f.arrayBuffer(), f.name, f.webkitRelativePath||f.name)); }
    setSweeps(prev=>[...prev,...parsed]); setStatus(`Added ${parsed.length} uploaded sweep(s).`);
  };
  return <main><header><div><h1>Cerebras Performance Projection UI</h1><p>Default shipped sweeps + runtime upload for customer go/no-go and internal sanity checks.</p></div><button onClick={()=>location.reload()}><RefreshCw size={16}/>Reload defaults</button></header>
    <section className="upload"><FileSpreadsheet/><div><h2>Upload more perf sweeps</h2><p>Upload one or many `.xlsx` files. New models and profiles render automatically.</p><input type="file" multiple accept=".xlsx" onChange={e=>onFiles([...e.target.files])}/></div><label className="drop">Choose files<Upload size={18}/><input type="file" multiple accept=".xlsx" onChange={e=>onFiles([...e.target.files])}/></label></section>
    <p className="status">{status}</p>
    {sweeps.length ? <><ComparisonView sweeps={sweeps}/><CustomerView sweeps={sweeps}/><InternalView sweeps={sweeps}/></> : <section className="empty">No sweeps loaded yet.</section>}
  </main>
}

createRoot(document.getElementById('root')).render(<App/>);
