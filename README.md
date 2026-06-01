<<<<<<< HEAD
# Cerebras Performance Sweep UI

**Live URL:** paste your Vercel deployment link here after deploying.

A Vercel-ready React/Vite UI for uploading one or many Cerebras performance projection `.xlsx` sweeps and comparing models for two audiences:

1. **Customer / PM view** — go/no-go workload fit, tokens/sec, TTFT, context, cost-style efficiency proxy.
2. **Internal engineer view** — projection sanity checks, config-level comparison, anomaly detection, raw normalized data.

## Features

- Upload one or many `.xlsx` files live in the browser.
- Supports unseen models such as `Model_L_profile_3/Model L profile 3.xlsx` with no code edits.
- Parses any conforming sweep by inferring model/profile from file name/path and metrics from column names.
- Side-by-side model comparison.
- Audience-specific views instead of one raw table dump.
- Runs as a static frontend, so Vercel deployment is simple.

## Local install and launch

```bash
npm install
npm run dev
```

Open the local URL printed by Vite, usually:

```bash
http://localhost:5173
```

## Build

```bash
npm run build
npm run preview
```

## Deploy to Vercel

1. Push this folder to GitHub.
2. Go to Vercel and create a new project.
3. Import the GitHub repo.
4. Use these settings:
   - Framework preset: `Vite`
   - Build command: `npm run build`
   - Output directory: `dist`
5. Click **Deploy**.
6. Copy the deployment URL and paste it at the top of this README and in the submission form.

## Data assumptions

The app treats the file/column shape as the contract, not hard-coded model names. It expects `.xlsx` files with rows containing performance projection values. Column names can vary; the parser uses flexible matching for common names such as:

- throughput / output tok/s / generated tokens per second
- TTFT / time to first token
- context length / input tokens
- output tokens / generation length
- batch size / concurrency / requests per second
- latency
- cost / price if available

If a metric is missing, the UI still renders and marks that field as unavailable.

## Video talking points

- Customer problem: quickly decide whether a model can satisfy workload requirements.
- Internal problem: compare sweeps, sanity-check projections, and detect suspicious rows before customer sharing.
- Framework choice: React/Vite because the whole workflow can run client-side and deploy easily on Vercel.
- Assumptions: uploaded `.xlsx` files follow the same general column contract; model/profile can be inferred from filename or workbook metadata.
- With more data: add production-vs-projection calibration, richer anomaly rules, confidence intervals, and saved comparison sessions.
=======
# ai-model-quality-challenge
>>>>>>> 66adcd75954f17de9bdf4fdd0fe4d1cf3a6e5bd8
