# Cloud Instance Optimizer — Localhost App

Real-time Flask + SocketIO web app that queries MongoDB for the cheapest
AWS EC2 and Azure VM instances for every server in your RVTools export.
Tracks all runs per client with full download history.

---

## Quick Start

```bash
cd cloud_optimizer_app

# 1 — Install dependencies (Python 3.9+)
pip install -r requirements.txt

# 2 — Make sure MongoDB is running at localhost:27017
#     with your EC2 and Azure pricing collections loaded

# 3 — Run the app
python app.py

# 4 — Open in browser
#     http://localhost:5000
```

---

## What it does

| Step | Action |
|------|--------|
| 1 | Enter **Client Name** (tracks all runs per client) |
| 2 | Configure MongoDB URI, AWS DB/collection, Azure DB/collection |
| 3 | Select AWS regions + primary region for queries |
| 4 | Upload your **RVTools Excel** (.xlsx) with VM inventory |
| 5 | Click **Run Optimizer** — watch real-time progress |
| 6 | Dashboard appears with AWS + Azure recommendations |
| 7 | Click **Download XLSX** to get the output file |

---

## MongoDB expected structure

**AWS EC2 database** (default: `EC2`, collection per region e.g. `us-east-1`):
```
{ TermType, LeaseContractLength, PurchaseOption, OfferingClass,
  Tenancy, "Operating System", "License Model", "Pre Installed S/W",
  "Instance Type", "Instance Family", PricePerUnit, vCPU, Memory }
```

**Azure VM database** (default: `Azure`, collection: `virtual_machines`):
```
{ vCPU, Memory, OS, "Instance Type",
  "Pay as you go per hour", "1 yr Savings per hour", "3 yr Savings per hour" }
```

---

## Input Excel required columns

| Column | Example |
|--------|---------|
| vCPU | 8 |
| Memory | 32 |
| Operating System | Windows |
| Production Servers | yes / no |
| Tenancy | Shared |
| License Model | No License required |
| Pre Installed S/W | SQL Ent (or blank) |
| VM | Server 1 |
| Group | APPServers-Prod |

---

## Output Excel (color coded)

- **Green rows** — AWS recommended memory < source memory (downsized)
- **Orange rows** — AWS recommended memory > source memory (no smaller instance)
- **White rows** — exact memory match

20 columns: 9 input + 11 new (AWS instance, pricing, Azure instance, Azure pricing)

---

## Client tracking

All runs are saved to `data/clients.json`. Use the **Clients** tab in the
sidebar to:
- Browse all clients with last-run date and cost
- Search by client name
- View full run history per client
- Re-download any previous output file
- Delete individual run records

---

## File structure

```
cloud_optimizer_app/
├── app.py              ← Flask + SocketIO backend
├── requirements.txt    ← Python dependencies
├── templates/
│   └── index.html      ← Full UI (dark theme, real-time)
├── uploads/            ← Input files (auto-created)
├── outputs/            ← Generated XLSX outputs (auto-created)
└── data/
    └── clients.json    ← Client + run history (auto-created)
```
