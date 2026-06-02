# Cerebras Performance UI

**Live URL:** add your Vercel URL here after deployment.

A Vercel-ready React/Vite tool for uploading, parsing, and comparing Cerebras performance projection sweeps.

## What it does

- Loads default `.xlsx` sweeps from `public/perf_data/` using a generated manifest.
- Lets users upload one or many additional `.xlsx` sweeps at runtime.
- Supports new model names such as `Model_L_profile_1` without code changes.
- Shows two audience views:
  - Customer / PM: go/no-go, throughput, TTFT, RPM, context, cache, recommended configuration.
  - Internal Engineer: raw summary rows, scenario sheets, anomaly flags, throughput/TTFT curves.
- Compares uploaded/default models side by side.

## Expected data layout

Put shipped data folders under:

```text
public/perf_data/
  Model_A_profile_1/
    Model A profile 1.xlsx
  Model_A_profile_2/
    Model A profile 2.xlsx
  ...
```

The app parses the workbook dynamically. It looks first for a `Summary` sheet and also reads scenario sheets such as `sim_...`.

## Install

```bash
npm install
```

## Run locally

```bash
npm run build:data
npm run dev
```

Open the local URL printed by Vite, usually:

```text
http://localhost:5173
```

## Add default data

Copy all folders from the extracted `perf_data.zip` into:

```text
public/perf_data/
```

Then regenerate the manifest:

```bash
npm run build:data
```

## Build

```bash
npm run build
```

## Deploy to Vercel

Push this folder to GitHub, then import the repo in Vercel.

Recommended Vercel settings:

```text
Framework Preset: Vite
Build Command: npm run build
Output Directory: dist
Install Command: npm install
```

## Notes

Because static hosting cannot list folders automatically, the default files are discovered by `scripts/build-data-manifest.js` at build time. Runtime uploads are still first-class and do not require rebuilds.
