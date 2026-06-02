# Cerebras Performance Projection UI

## Live Demo

https://YOUR-VERCEL-URL.vercel.app

---

## Overview

This project transforms Cerebras performance projection spreadsheets into actionable views for two audiences:

### Customer / Customer-Facing PM
Provides a go/no-go assessment for a workload based on:
- Throughput (tokens/sec)
- TTFT (Time To First Token)
- RPM (Requests Per Minute)
- Context size
- Recommended deployment configuration

### Internal Product / Deployment Engineer
Provides visibility into:
- Performance sweep configurations
- Batch size and concurrency behavior
- Throughput and TTFT trends
- Potential anomalies
- Cross-model comparisons

The application supports both the shipped performance sweeps and new user-uploaded sweeps without requiring code changes or rebuilds.

---

## Features

### Dynamic Sweep Loading

The application automatically loads the shipped performance sweeps from:

```text
public/perf_data/