"""
Cloud Instance Optimizer — Full-Stack Flask App
Run: python app.py  |  Visit: http://localhost:5000
"""

import os, json, uuid, re, time, threading, sys, argparse
from datetime import datetime
from pathlib import Path
from collections import Counter

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, abort
from flask_socketio import SocketIO, emit
from pymongo import MongoClient, InsertOne
from pymongo.errors import ServerSelectionTimeoutError, ConnectionFailure
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import concurrent.futures

# Azure pricing engine — self-contained, no MongoDB needed. Pricing database,
# SKU sizing, OS/SQL/non-prod detection, and Excel generation all live here.
import azure_engine

try:
    import requests
except ImportError:
    requests = None   # pricing-data loading features will report a clear error if this is missing

# ── App setup ────────────────────────────────────────────────────────────────
BASE_DIR    = Path(__file__).parent.resolve()
app         = Flask(__name__,
                    template_folder=str(BASE_DIR / 'templates'),
                    static_folder=str(BASE_DIR / 'static'))
app.config['SECRET_KEY'] = 'cloud-optimizer-secret-2024'

# Resilient template lookup: search BOTH files/templates/ (the correct spot)
# AND files/ directly (in case UI.html got saved next to this script by
# mistake), so a TemplateNotFound error can't happen just from file placement.
from jinja2 import ChoiceLoader, FileSystemLoader
app.jinja_loader = ChoiceLoader([
    FileSystemLoader(str(BASE_DIR / 'templates')),
    FileSystemLoader(str(BASE_DIR)),
])

socketio    = SocketIO(app, async_mode='threading', cors_allowed_origins='*',
                       logger=False, engineio_logger=False)

UPLOAD_DIR  = BASE_DIR / 'uploads'
OUTPUT_DIR  = BASE_DIR / 'outputs'
DATA_DIR    = BASE_DIR / 'data'
for d in [UPLOAD_DIR, OUTPUT_DIR, DATA_DIR]:
    d.mkdir(exist_ok=True)

CLIENTS_FILE = DATA_DIR / 'clients.json'

# ── Client helpers ────────────────────────────────────────────────────────────
def load_clients():
    if CLIENTS_FILE.exists():
        try:
            return json.loads(CLIENTS_FILE.read_text())
        except Exception:
            return []
    return []

def save_clients(clients):
    CLIENTS_FILE.write_text(json.dumps(clients, indent=2))

def upsert_client_run(client_name, run_meta):
    clients = load_clients()
    client  = next((c for c in clients if c['name'].lower() == client_name.lower()), None)
    if client is None:
        client = {'id': str(uuid.uuid4()), 'name': client_name,
                  'created_at': datetime.now().isoformat(), 'runs': []}
        clients.append(client)
    client['last_run'] = run_meta['timestamp']
    client['runs'].insert(0, run_meta)
    save_clients(clients)
    return client

# ── MongoDB — per-URI cached connections ─────────────────────────────────────
_mongo_lock    = threading.Lock()
_mongo_clients = {}   # uri → MongoClient

def get_mongo(uri='mongodb://localhost:27017/'):
    with _mongo_lock:
        mc = _mongo_clients.get(uri)
        if mc is not None:
            try:
                mc.admin.command('ping', serverSelectionTimeoutMS=2000)
                return mc
            except Exception:
                try: mc.close()
                except Exception: pass
                _mongo_clients.pop(uri, None)
        mc = MongoClient(uri, serverSelectionTimeoutMS=6000,
                         connectTimeoutMS=6000, socketTimeoutMS=30000,
                         maxPoolSize=20)
        mc.admin.command('ping')          # raises if unreachable
        _mongo_clients[uri] = mc
        return mc


# ══════════════════════════════════════════════════════════════════════════════
# AWS EC2 PRICING DATA LOADER
# ══════════════════════════════════════════════════════════════════════════════
# Fetches official AWS EC2 pricing directly from the public AWS Price List API
# and loads it into MongoDB in the exact schema the optimizer queries against.
# No AWS account or API key needed — the Price List API is public and free.
#
# Can be used two ways:
#   1. From the running app — Configure tab has a "Load Region Pricing" panel
#      that calls this over Socket.IO with live progress, same as a normal run.
#   2. From the command line, without starting the web server at all:
#         python "With SMC.py" --load-pricing --regions af-south-1 us-west-1
#
# NOTE: each region's price list is several hundred MB and takes a few
# minutes to download and parse — this needs a real internet connection.

REGION_PRICE_INDEX_URL = "https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonEC2/current/{region}/index.json"
PRICING_HOURS_PER_YEAR = 8760   # used to amortize Reserved upfront costs into an effective hourly rate


def fetch_region_price_list(region, progress_cb=None):
    """Download the full EC2 price list JSON for one region. progress_cb(pct, msg) is optional."""
    if requests is None:
        raise RuntimeError("The 'requests' package is required for pricing data loading. "
                            "Install it with: pip install requests --break-system-packages")
    url = REGION_PRICE_INDEX_URL.format(region=region)
    resp = requests.get(url, stream=True, timeout=120)
    resp.raise_for_status()

    total = int(resp.headers.get('content-length', 0))
    downloaded = 0
    chunks = []
    last_pct = -1
    for chunk in resp.iter_content(chunk_size=1024 * 1024):
        chunks.append(chunk)
        downloaded += len(chunk)
        if total:
            pct = int(downloaded / total * 100)
            if pct != last_pct and pct % 5 == 0:
                last_pct = pct
                if progress_cb:
                    progress_cb(pct, f'Downloading {region}: {downloaded/1e6:.0f}MB / {total/1e6:.0f}MB')

    if progress_cb:
        progress_cb(100, f'Downloaded {downloaded/1e6:.0f}MB — parsing JSON...')
    return json.loads(b"".join(chunks))


def build_pricing_documents(price_list, region, progress_cb=None):
    """
    Walk the AWS Price List JSON and build one MongoDB document per
    (Instance Type, OS, Tenancy, License, SW, pricing plan) combination,
    matching the schema the optimizer's queries expect.
    """
    products        = price_list.get("products", {})
    on_demand_terms = price_list.get("terms", {}).get("OnDemand", {})
    reserved_terms  = price_list.get("terms", {}).get("Reserved", {})

    docs = []
    total_products = len(products)

    for i, (sku, product) in enumerate(products.items()):
        if progress_cb and i % 20000 == 0 and total_products:
            progress_cb(int(i / total_products * 100), f'Parsing {region}: {i:,}/{total_products:,} SKUs')

        if product.get("productFamily") != "Compute Instance":
            continue
        attrs = product.get("attributes", {})
        # Only "Used" capacity SKUs are real purchasable instances — excludes
        # Dedicated Host allocation SKUs and other non-standard rows.
        if attrs.get("capacitystatus") != "Used":
            continue
        instance_type = attrs.get("instanceType")
        if not instance_type:
            continue

        base_doc = {
            "Instance Type":     instance_type,
            "Instance Family":   attrs.get("instanceFamily", ""),
            "vCPU":              attrs.get("vcpu", ""),
            "Memory":            attrs.get("memory", ""),
            "Operating System":  attrs.get("operatingSystem", "Linux"),
            "Tenancy":           attrs.get("tenancy", "Shared"),
            "License Model":     attrs.get("licenseModel", "No License required"),
            "Pre Installed S/W": attrs.get("preInstalledSw", "NA"),
            "Region":            region,
        }

        # ── On-Demand pricing ────────────────────────────────────────────────
        for term_code, term_data in on_demand_terms.get(sku, {}).items():
            for rate_code, dim in term_data.get("priceDimensions", {}).items():
                price = dim.get("pricePerUnit", {}).get("USD")
                if price is None or dim.get("unit") != "Hrs":
                    continue
                try:
                    price_f = float(price)
                except (TypeError, ValueError):
                    continue
                if price_f <= 0:
                    continue
                doc = dict(base_doc)
                doc.update({"TermType": "OnDemand", "PricePerUnit": round(price_f, 6)})
                docs.append(doc)

        # ── Reserved pricing (1yr/3yr × No/Partial/All Upfront) ─────────────
        for term_code, term_data in reserved_terms.get(sku, {}).items():
            term_attrs = term_data.get("termAttributes", {})
            lease    = term_attrs.get("LeaseContractLength", "")
            option   = term_attrs.get("PurchaseOption", "")
            offering = term_attrs.get("OfferingClass", "standard")
            if not lease or not option:
                continue
            years = 1 if lease.startswith("1") else 3 if lease.startswith("3") else None
            if years is None:
                continue
            total_hours = years * PRICING_HOURS_PER_YEAR

            upfront_cost, recurring_hourly = 0.0, 0.0
            for rate_code, dim in term_data.get("priceDimensions", {}).items():
                price = dim.get("pricePerUnit", {}).get("USD")
                if price is None:
                    continue
                try:
                    price_f = float(price)
                except (TypeError, ValueError):
                    continue
                if dim.get("unit") == "Quantity":
                    upfront_cost = price_f
                elif dim.get("unit") == "Hrs":
                    recurring_hourly = price_f

            # Amortized effective hourly rate — one comparable number for
            # "cheapest" sorting, same convention as the existing regions.
            effective_hourly = recurring_hourly + (upfront_cost / total_hours if total_hours else 0)
            if effective_hourly <= 0:
                continue

            doc = dict(base_doc)
            doc.update({
                "TermType": "Reserved", "LeaseContractLength": lease,
                "PurchaseOption": option, "OfferingClass": offering,
                "PricePerUnit": round(effective_hourly, 6),
            })
            docs.append(doc)

    if progress_cb:
        progress_cb(100, f'Built {len(docs):,} pricing documents for {region}')
    return docs


def load_pricing_into_mongo(docs, mongo_uri, db_name, collection_name, replace=True, progress_cb=None):
    """Bulk-insert pricing documents into MongoDB, replacing any existing data for that region."""
    mc   = get_mongo(mongo_uri)
    coll = mc[db_name][collection_name]

    if replace:
        coll.delete_many({})

    batch_size = 2000
    inserted = 0
    for i in range(0, len(docs), batch_size):
        batch = docs[i:i + batch_size]
        result = coll.bulk_write([InsertOne(d) for d in batch], ordered=False)
        inserted += result.inserted_count
        if progress_cb:
            progress_cb(int(inserted / len(docs) * 100) if docs else 100,
                        f'Inserted {inserted:,}/{len(docs):,} into {db_name}.{collection_name}')
    return inserted


def load_region_pricing(region, mongo_uri='mongodb://localhost:27017/', db_name='EC2', progress_cb=None):
    """
    End-to-end: download + parse + load pricing for one AWS region.
    progress_cb(pct, msg) is called throughout for live progress reporting.
    Returns the number of documents inserted.
    """
    def cb(pct, msg):
        if progress_cb: progress_cb(pct, msg)

    cb(0, f'Starting pricing load for {region}...')
    price_list = fetch_region_price_list(region, progress_cb=lambda p, m: cb(int(p * 0.5), m))
    docs = build_pricing_documents(price_list, region, progress_cb=lambda p, m: cb(50 + int(p * 0.25), m))
    if not docs:
        raise RuntimeError(f'No pricing documents could be built for region "{region}" — '
                            f'check the region code is correct.')
    inserted = load_pricing_into_mongo(docs, mongo_uri, db_name, region,
                                        replace=True, progress_cb=lambda p, m: cb(75 + int(p * 0.25), m))
    cb(100, f'Done — {inserted:,} pricing records loaded for {region}')
    return inserted

# ── Instance family classification ───────────────────────────────────────────
def classify_instance_family(vcpu, memory_gb):
    if vcpu == 0: return 'general'
    ratio = memory_gb / vcpu
    if ratio < 4:   return 'compute'
    if ratio <= 8:  return 'general'
    return 'memory'

FAMILY_REGEX = {
    'compute': r'^c[0-9]',
    'general': r'^m[0-9]',
    'memory':  r'^r[0-9]',
}

# Standard AWS vCPU sizes in ascending order
AWS_VCPU_SIZES = [1, 2, 4, 8, 16, 32, 48, 64, 96, 128, 192]

def get_vcpu_candidates(vcpu):
    """Return vCPU candidates: exact first, one step down (n-1), one step up."""
    vcpu = int(vcpu)
    candidates = [vcpu]
    below = [v for v in AWS_VCPU_SIZES if v < vcpu]
    if below: candidates.append(below[-1])
    above = [v for v in AWS_VCPU_SIZES if v > vcpu]
    if above: candidates.append(above[0])
    return candidates

def _excl(is_prod):
    if is_prod: return r'^(a1\.|t|.*g[^.]*\.)|^[mcr][34]\.|^t2\.'
    return r'^(a1\.|.*g[^.]*\.)|^[mcr][34]\.|^t2\.'

def _mem_stage():
    return {'$addFields': {'MemoryNumeric': {
        '$toDouble': {'$substrBytes': ['$Memory', 0, {'$indexOfBytes': ['$Memory',' ']}]}
    }}}

def _tier_stage():
    return {'$addFields': {'Tier': {'$switch': {'branches': [
        {'case': {'$regexMatch': {'input':'$Instance Type','regex':r'^t[0-9]+a\.|^t4g\.','options':'i'}}, 'then':1},
        {'case': {'$regexMatch': {'input':'$Instance Type','regex':r'^t[0-9]+\.','options':'i'}},          'then':2},
        {'case': {'$regexMatch': {'input':'$Instance Type','regex':r'^[a-z]+[0-9]+a\.','options':'i'}},   'then':3},
        {'case': {'$regexMatch': {'input':'$Instance Type','regex':r'^[a-z]+[0-9]+i\.','options':'i'}},   'then':4},
    ], 'default':5}}}}

PROJ = {'$project':{'Instance Type':1,'Instance Family':1,'PricePerUnit':1,'vCPU':1,'Memory':1,'_id':0}}

# ── Purchasing option → MongoDB field mapping ─────────────────────────────────
PURCHASE_MAP = {
    '3yr_no_upfront':      ('3yr','No Upfront','standard'),
    '3yr_partial_upfront': ('3yr','Partial Upfront','standard'),
    '3yr_all_upfront':     ('3yr','All Upfront','standard'),
    '1yr_no_upfront':      ('1yr','No Upfront','standard'),
    '1yr_partial_upfront': ('1yr','Partial Upfront','standard'),
    '1yr_all_upfront':     ('1yr','All Upfront','standard'),
    'on_demand':           (None, None, None),
}

def _build_base_match(tenancy, os_, lic, sw, excl_regex, fam_regex, lease, purchase, offering):
    if lease is None:
        base = {
            'TermType': 'OnDemand', 'Tenancy': tenancy,
            'Operating System': os_, 'License Model': lic, 'Pre Installed S/W': sw,
        }
    else:
        base = {
            'TermType': 'Reserved', 'LeaseContractLength': lease,
            'PurchaseOption': purchase, 'OfferingClass': offering,
            'Tenancy': tenancy, 'Operating System': os_,
            'License Model': lic, 'Pre Installed S/W': sw,
        }
    if fam_regex:
        base['$and'] = [
            {'Instance Type': {'$regex': fam_regex,   '$options':'i'}},
            {'Instance Type': {'$not':{'$regex': excl_regex,'$options':'i'}}}
        ]
    else:
        base['Instance Type'] = {'$not':{'$regex': excl_regex,'$options':'i'}}
    return base

def _run_query(collection, tenancy, os_, lic, sw, excl_regex, fam_regex,
               vcpu_i, mem_filter, lease, purchase, offering):
    """Run one aggregate query, return top 3 cheapest sorted by numeric price ASC."""
    vcpu_i   = int(vcpu_i)
    vcpu_str = str(vcpu_i)
    base     = _build_base_match(tenancy, os_, lic, sw, excl_regex, fam_regex,
                                  lease, purchase, offering)
    try:
        return list(collection.aggregate([
            {'$match': base},
            {'$addFields': {
                'vCPU_int':     {'$toInt': '$vCPU'},
                'PriceNumeric': {'$toDouble': '$PricePerUnit'},   # numeric, not string, sort
            }},
            {'$match': {'$or': [
                {'vCPU': {'$eq': vcpu_i}},
                {'vCPU': {'$eq': vcpu_str}},
                {'vCPU_int': {'$eq': vcpu_i}},
            ]}},
            _mem_stage(),
            {'$match': mem_filter},
            _tier_stage(),
            {'$sort': {'PriceNumeric': 1, 'Tier': 1}},
            {'$limit': 3},
            PROJ,
        ], allowDiskUse=True))
    except Exception:
        return []

def _price(row):
    """Extract float price; return infinity if missing."""
    try:
        return float(row.get('PricePerUnit', 999999) or 999999)
    except Exception:
        return 999999.0

def _parse(row, src_mem):
    """Parse a MongoDB result row into the standard 7-tuple."""
    mem_str = row.get('Memory', '0 GiB')
    try:
        mem_num = float(mem_str.split()[0])
    except Exception:
        mem_num = 0.0
    return (
        row.get('Instance Type'),
        row.get('Instance Family'),
        row.get('PricePerUnit'),
        row.get('vCPU'),
        row.get('Memory'),
        mem_num < src_mem,
        mem_num > src_mem,
    )

# ── CORRECT INSTANCE SELECTION LOGIC ─────────────────────────────────────────
# Priority: exact match → downsize (>5% cheaper + >=50% mem) → upsize (≤2.5x) → vCPU fallback
# Hard floor MIN_MEM_RATIO=0.5 prevents insane downsizes like 24GiB→0.5GiB
MIN_MEM_RATIO = 0.5

def get_cheapest_aws_instance(collection, vcpu, memory, os_, tenancy, lic, sw,
                               is_prod, lease, purchase, offering):
    """Find the cheapest appropriate AWS instance with correct accuracy."""
    excl_r  = _excl(is_prod)
    fam_r   = FAMILY_REGEX[classify_instance_family(vcpu, memory)]
    src_mem = float(memory)
    src_vc  = int(vcpu)
    min_mem = src_mem * MIN_MEM_RATIO

    def q(fam, mem_filter, vc):
        return _run_query(collection, tenancy, os_, lic, sw, excl_r,
                          fam, vc, mem_filter, lease, purchase, offering)

    exact_filter = {'MemoryNumeric': {'$gte': src_mem - 0.5, '$lte': src_mem + 0.5}}
    down_filter  = {'MemoryNumeric': {'$gte': min_mem, '$lt': src_mem - 0.5}}
    up_filter    = {'MemoryNumeric': {'$gte': src_mem + 0.5, '$lte': src_mem * 2.5}}

    exact_rows = q(fam_r, exact_filter, src_vc) or q(None, exact_filter, src_vc)
    exact_best = min(exact_rows, key=_price) if exact_rows else None

    down_rows = q(None, down_filter, src_vc)
    if down_rows:
        down_rows.sort(key=lambda r: (_price(r), -float(r.get('Memory','0 GiB').split()[0])))
        down_best = down_rows[0]
    else:
        down_best = None

    up_rows = q(None, up_filter, src_vc)
    up_best = min(up_rows, key=_price) if up_rows else None

    winner = None
    if exact_best:
        winner = exact_best
        if down_best:
            down_mem = float(down_best.get('Memory','0 GiB').split()[0])
            if down_mem >= min_mem and _price(down_best) < _price(exact_best) * 0.95:
                winner = down_best
    elif down_best:
        down_mem = float(down_best.get('Memory','0 GiB').split()[0])
        if down_mem >= min_mem:
            winner = down_best

    if winner is None and up_best:
        winner = up_best

    if winner is None:
        for alt_vc in get_vcpu_candidates(src_vc)[1:]:
            r = q(None, exact_filter, alt_vc) or q(None, up_filter, alt_vc)
            if r:
                winner = min(r, key=_price)
                break

    if winner is None:
        return (None, None, None, None, None, False, False)
    return _parse(winner, src_mem)


# ══════════════════════════════════════════════════════════════════════════════
# FAST CACHED LIKE-TO-LIKE / N-1 / N-2 LOGIC — mirrors the original proven script
# ══════════════════════════════════════════════════════════════════════════════
# This replaces the old multi-query "smart downsize" approach, which was both
# slower (4+ Mongo round-trips per VM) AND sometimes landed on a pricier
# instance than a simple "cheapest that satisfies the minimum" match.
#
# Simple rule: exclude burstable/graviton/older-gen families, require
# vCPU >= target AND Memory >= target, sort by PricePerUnit ascending, take #1.
# Every unique (vcpu, mem, os, tenancy, lic, sw, is_prod, pricing) combination
# is queried ONCE and cached — since many VMs in a fleet share identical specs,
# this collapses hundreds of duplicate queries into a handful of real ones.

def _like_excl(is_prod):
    """
    Exclusion regex — rules out families that are never valid general-purpose
    rehost targets, so the "cheapest" search only ever competes among
    standard server instance types:
      - a1          (old-gen, phased out)
      - Graviton-suffixed families (m6g, c6gn, r6gd, im4gn, is4gen, etc.)
      - p*/g*       (GPU compute/graphics — P3/P4/P5, G4/G5)
      - vt*         (VT1 — dedicated video transcoding)
      - f*          (F1 — FPGA)
      - inf*/trn*   (Inferentia/Trainium — ML inference/training chips)
      - dl*         (DL1 — deep learning)
      - hpc*        (HPC6a/HPC7a — tightly-coupled HPC clusters)
      - mac*        (Mac1/Mac2 — bare-metal macOS hardware)
      - u-*         (High Memory bare metal, e.g. u-6tb1.metal)
      - x*          (X1/X1e/X2 — extreme-memory, far pricier per vCPU)
      - z1d         (high-frequency + NVMe, niche HPC/EDA workload)
      - t*          (burstable — excluded for production only)
    Tighter than a blanket '.*g.*' match so it can never accidentally
    exclude a legitimate cheap family that merely contains the letter 'g'
    somewhere in its name.
    """
    graviton    = r'^[a-z]+[0-9]+g[a-z]*\.'
    old_gen     = r'^a[0-9]'
    specialized = r'^(p[0-9]|g[0-9]|vt[0-9]|f[0-9]|inf[0-9]|trn[0-9]|dl[0-9]|hpc[0-9]|mac[0-9]|u-|x[0-9e]|z1d)'
    if is_prod:
        return rf'{old_gen}|^t[0-9]|{graviton}|{specialized}'
    return rf'{old_gen}|{graviton}|{specialized}'

FLEX_REGEX = r'-flex\.'

def cheapest_like_to_like_aws(collection, vcpu, memory, os_, tenancy, lic, sw,
                               is_prod, lease, purchase, offering, cache, min_gen_pref=5):
    """
    Cheapest instance where vCPU >= target and Memory >= target, bounded to a
    sane size window so it can never return an absurdly oversized instance
    just because it happened to be the cheapest thing in the DB that
    technically satisfies the floor (e.g. a 34 GiB requirement should never
    resolve to a 128 GiB instance). Specialized/GPU/graviton/old-gen families
    and burstable-for-prod are already ruled out by _like_excl.

    min_gen_pref is a HARD filter: whichever generation you select in the UI
    (5th gen+ or 6th gen+) is enforced — only instances at or above that
    generation are eligible. Within that filtered, bounded, excluded pool,
    selection is simple: sort cheapest-first, take the first result.

    Fallback order (each tier only runs if the previous one found nothing):
      1. Bounded size window + generation >= min_gen_pref   (expected case)
      2. Unbounded size      + generation >= min_gen_pref   (sparse data at that size)
      3. Bounded size window + any generation               (true last resort —
         no instance of the requested generation exists for this spec at all)
      4. Unbounded size      + any generation

    Cached per unique parameter combination (including the generation choice)
    for speed across a whole fleet.

    CRITICAL: PricePerUnit is explicitly cast to a numeric type (PriceNumeric)
    before sorting. If it were sorted as a raw string, MongoDB would compare
    it lexicographically — e.g. "10.2" would sort BEFORE "9.5" because '1' < '9'
    — silently returning a more expensive instance as "cheapest".
    """
    key = (int(vcpu), float(memory), os_, tenancy, lic, sw, is_prod, lease, purchase, offering, min_gen_pref)
    if key in cache:
        return cache[key]

    excl_regex = _like_excl(is_prod)
    base = _build_base_match(tenancy, os_, lic, sw, excl_regex, None, lease, purchase, offering)

    vcpu_f, mem_f = float(vcpu), float(memory)
    # Sane ceiling: never more than 2x the requirement, with a small floor
    # of headroom (+4 vCPU / +8 GiB) so very small sources still get options.
    vcpu_ceiling = max(vcpu_f * 2, vcpu_f + 4)
    mem_ceiling  = max(mem_f * 2, mem_f + 8)

    def run(bounded, min_gen=None):
        match_stage = {'vCPU_int': {'$gte': int(vcpu)}}
        if bounded:
            match_stage['vCPU_int']['$lte'] = int(vcpu_ceiling)
        mem_match = {'MemoryNumeric': {'$gte': mem_f}}
        if bounded:
            mem_match['MemoryNumeric']['$lte'] = mem_ceiling

        pipeline = [
            {'$match': base},
            {'$addFields': {
                'vCPU_int':     {'$toInt': '$vCPU'},
                'PriceNumeric': {'$toDouble': '$PricePerUnit'},
                # Extract the generation number from the family code, e.g.
                # "m7i.xlarge" -> 7, "c6a.large" -> 6. Non-numeric/odd names
                # default to 0 so they never falsely pass a min_gen filter.
                'GenMatch': {'$regexFind': {'input': '$Instance Type', 'regex': '^[a-z]+([0-9]+)'}},
            }},
            {'$addFields': {
                'GenNumber': {'$convert': {
                    'input': {'$arrayElemAt': ['$GenMatch.captures', 0]},
                    'to': 'int', 'onError': 0, 'onNull': 0,
                }},
            }},
            {'$match': match_stage},
            _mem_stage(),
            {'$match': mem_match},
        ]
        if min_gen is not None:
            pipeline.append({'$match': {'GenNumber': {'$gte': min_gen}}})
        pipeline += [
            {'$sort': {'PriceNumeric': 1}},
            {'$limit': 1},
            PROJ,
        ]
        try:
            return list(collection.aggregate(pipeline, allowDiskUse=True))
        except Exception:
            return []

    r = run(bounded=True,  min_gen=min_gen_pref)   # 1. expected case
    if not r:
        r = run(bounded=False, min_gen=min_gen_pref)  # 2. sparse data at that size
    if not r:
        r = run(bounded=True)                          # 3. no instance of that gen exists for this spec
    if not r:
        r = run(bounded=False)                         # 4. absolute last resort

    result = _parse(r[0], mem_f) if r else (None, None, None, None, None, False, False)
    cache[key] = result
    return result


def n1_dims(vcpu, memory):
    """N-1 downsize rule — exact logic from the original script."""
    vcpu, memory = int(vcpu), int(memory)
    n1_vcpu = vcpu // 2 if vcpu >= 8 else vcpu
    n1_mem  = memory // 2 if memory >= 16 else memory
    return n1_vcpu, n1_mem


def n2_dims(vcpu, memory):
    """N-2 downsize rule — continues downsizing from N-1, exact script logic."""
    vcpu, memory = int(vcpu), int(memory)
    n1_vcpu, n1_mem = n1_dims(vcpu, memory)
    n2_vcpu = n1_vcpu // 2 if (vcpu >= 8 and n1_vcpu >= 8) else n1_vcpu
    n2_mem  = n1_mem // 2 if (memory >= 16 and n1_mem >= 16) else n1_mem
    return n2_vcpu, n2_mem


def resolve_target_dims(calc_mode, vcpu, memory):
    """Return (target_vcpu, target_memory) based on selected calculation mode."""
    if calc_mode == 'optimized_n1':
        return n1_dims(vcpu, memory)
    if calc_mode == 'optimized_n2':
        return n2_dims(vcpu, memory)
    return int(vcpu), float(memory)   # like_to_like — no downsizing




# ── Purchasing option → Bulk Import label (Compute Savings Plan format) ─────────
# Exact values from VH Calc Input sheet column A of the official template
BULK_PURCHASE_MAP = {
    "3yr_no_upfront":      "3 Yr No Upfront Compute Savings Plan",
    "3yr_partial_upfront": "3 Yr Partial Upfront Compute Savings Plan",
    "3yr_all_upfront":     "3 Yr All Upfront Compute Savings Plan",
    "1yr_no_upfront":      "1 Yr No Upfront Compute Savings Plan",
    "1yr_partial_upfront": "1 Yr Partial Upfront Compute Savings Plan",
    "1yr_all_upfront":     "1 Yr All Upfront Compute Savings Plan",
    "on_demand":           "On-Demand",
}

# OS mapping — exact values from VH Calc Input sheet column G
# Priority: check SQL variant first, then base OS
def map_os_for_bulk(os_str, sw_str):
    os_l = str(os_str or "").strip().lower()
    sw_l = str(sw_str or "").strip().lower()
    is_win  = "win" in os_l
    is_rhel = "rhel" in os_l or "red hat" in os_l
    is_suse = "suse" in os_l
    # NOTE: check RHEL before Linux since "red hat" doesn't contain "linux"
    # but some inputs might say "Red Hat Linux" - always check rhel first

    has_sql_ent = any(x in sw_l for x in ["sql ent", "sql enterprise", "sql server enterprise"])
    has_sql_std = any(x in sw_l for x in ["sql std", "sql standard", "sql server standard"])
    has_sql_web = "sql web" in sw_l or "sql server web" in sw_l

    if has_sql_ent:
        if is_rhel:  return "Red Hat Enterprise Linux with SQL Server Enterprise"
        if is_win:   return "Windows Server with SQL Server Enterprise"
        return "Linux with SQL Server Enterprise"
    if has_sql_std:
        if is_rhel:  return "Red Hat Enterprise Linux with SQL Server Standard"
        if is_win:   return "Windows Server with SQL Server Standard"
        return "Linux with SQL Server Standard"
    if has_sql_web:
        if is_rhel:  return "Red Hat Enterprise Linux with SQL Server Web"
        if is_win:   return "Windows Server with SQL Server Web"
        return "Linux with SQL Server Web"
    # Base OS (no SQL)
    if is_rhel:  return "Red Hat Enterprise Linux"
    if is_suse:  return "SUSE Linux Enterprise Server"
    if is_win:   return "Windows Server"
    return "Linux"

def sanitize_group_name(name):
    """
    AWS Pricing Calculator Bulk Import 'Group name' rules:
      - Cannot contain '<', '>', or '&'
      - Maximum 60 characters
    This cleans any group/sub-group name so it's always accepted:
      - '&' -> 'and ' (keeps meaning instead of just dropping it)
      - '<' and '>' -> removed entirely
      - Collapses extra whitespace left behind
      - If over 60 chars, truncates at the last whole word that fits
        (a clean "short form" instead of cutting mid-word), falling back
        to a hard cut only if there's no space to break on.
    """
    name = str(name or '').strip()
    name = name.replace('&', 'and ')
    name = re.sub(r'[<>]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    if len(name) > 60:
        truncated = name[:60]
        if ' ' in truncated:
            truncated = truncated.rsplit(' ', 1)[0]
        name = truncated.strip()
    return name or 'Group'


def lease_term_label(purchase_key):
    """Human label matching AWS Pricing Calculator group naming, e.g. '3 Year' / '1 Year' / 'On-Demand'."""
    if purchase_key == 'on_demand':
        return 'On-Demand'
    lt = str(purchase_key or '').split('_')[0]
    return {'3yr': '3 Year', '1yr': '1 Year'}.get(lt, lt)


def generate_bulk_import(records, aws_insts, aws_region, purchase_key, out_path, client_name=None):
    """Generate AWS Pricing Calculator Bulk Import — exact match to official template format."""
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Alignment, PatternFill, Font
    import copy

    tpl_path = BASE_DIR / "templates" / "bulk_template.xlsx"
    if tpl_path.exists():
        wb = load_workbook(str(tpl_path))
        ws = wb["Inputs"]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Inputs"

    purchase_label = BULK_PURCHASE_MAP.get(purchase_key, "3 Yr No Upfront Compute Savings Plan")
    region_code    = str(aws_region or "us-east-1").strip()

    # ── AWS Pricing Calculator group hierarchy ────────────────────────────────
    # You create the top-level group manually in the Calculator UI:
    #   My Estimates > "<Client Name> Rehost ARR <Pricing Option>"
    # Then upload this bulk sheet INTO that group. Every row here shares one
    # uniform Group value — "<N> EC2 Servers" — which becomes the sub-group
    # AWS Pricing Calculator creates under it, e.g.:
    #   My Estimates
    #   └── ABC Corp Rehost ARR 3 Year
    #        └── 102 EC2 Servers
    ec2_count          = sum(1 for inst in aws_insts if inst)
    ec2_group_name     = sanitize_group_name(f"{ec2_count} EC2 Servers")
    lease_label        = lease_term_label(purchase_key)
    suggested_top_group = sanitize_group_name(f"{client_name} Rehost ARR {lease_label}") if client_name else None

    # Clear existing data rows (keep rows 1-3: merged title row + header row)
    for row in ws.iter_rows(min_row=4, max_row=max(ws.max_row, 4)):
        for cell in row:
            cell.value = None

    data_row = 4
    for i, rec in enumerate(records):
        inst = aws_insts[i] if i < len(aws_insts) else None
        if not inst:
            continue

        # Group: read directly from the input file's "Group" column
        # (falls back to Resource Pool / Cluster if Group isn't present).
        # Sanitized against AWS Bulk Import's rules: no '<', '>', '&', and
        # max 60 characters — input files often have long/messy group names
        # that would otherwise get silently rejected on upload.
        group = sanitize_group_name(
            rec.get("Group") or rec.get("Resource Pool") or rec.get("Cluster") or ""
        )

        # VM name from input file
        vm_name = str(
            rec.get("VM") or rec.get("VM Name") or rec.get("Name") or
            rec.get("DNS Name") or rec.get("Description") or f"VM-{i+1}"
        ).strip()

        os_str      = str(rec.get("Operating System", "Linux") or "Linux")
        sw_str      = str(rec.get("Pre Installed S/W", "") or "")
        tenancy_raw = str(rec.get("Tenancy", "Shared") or "Shared").strip().lower()

        # Tenancy must exactly match VH Calc Input valid values
        tenancy_val = "Dedicated Instances" if ("dedicated" in tenancy_raw or "host" in tenancy_raw) else "Shared Instances"

        bulk_os = map_os_for_bulk(os_str, sw_str)

        # Col A: formula (integer result via formula — matches template exactly)
        ws.cell(data_row, 1).value = f"=IF(ISBLANK(H{data_row}),0,1)"
        # Col B: Group (from input file)
        ws.cell(data_row, 2).value = group
        # Col C: Description/VM name (from input file)
        ws.cell(data_row, 3).value = vm_name
        # Col D: AWS Region (must be valid region code e.g. us-east-1)
        ws.cell(data_row, 4).value = region_code
        # Col E: Operating System (exact value from VH Calc Input col G)
        ws.cell(data_row, 5).value = bulk_os
        # Col F: Instance Type (e.g. m5.xlarge)
        ws.cell(data_row, 6).value = inst
        # Col G: Tenancy (exact value from VH Calc Input)
        ws.cell(data_row, 7).value = tenancy_val
        # Col H: Number of Instances — MUST be integer, not string
        ws.cell(data_row, 8).value = 1
        # Col I: Assumed Usage — leave blank (optional)
        ws.cell(data_row, 9).value = None
        # Col J: Usage Type
        ws.cell(data_row, 10).value = "Always On"
        # Col K: Purchasing Option (exact value from VH Calc Input col A)
        ws.cell(data_row, 11).value = purchase_label
        # Col L: Storage Type
        ws.cell(data_row, 12).value = "General Purpose SSD (gp3)"
        # Col M: Storage amount (GB) — blank, customer fills per VM
        ws.cell(data_row, 13).value = None
        # Col N: Provisioning IOPS — MUST be integer, not string
        ws.cell(data_row, 14).value = 3000
        # Col O: EBS Throughput (Mbps) — MUST be integer, not string
        ws.cell(data_row, 15).value = 100
        # Col P: Snapshot Frequency
        ws.cell(data_row, 16).value = "No snapshot storage"
        # Col Q: EBS Snapshot amount — blank
        ws.cell(data_row, 17).value = None

        data_row += 1

    wb.save(str(out_path))
    return {
        'rows_written': data_row - 4,
        'ec2_group_name': ec2_group_name,
        'suggested_top_group': suggested_top_group,
    }


# ══════════════════════════════════════════════════════════════════════════════
# AWS SMC SERVICE DISCOVERY
# ══════════════════════════════════════════════════════════════════════════════
# Pure extension — does not touch pricing, sizing, or the existing XLSX
# outputs. Reads the same raw input `records` already used elsewhere in the
# pipeline and derives a list of AWS services the migration will actually
# need, beyond just EC2 compute. This is the data model the future
# Playwright/Selenium/REST-based SMC uploader will consume — each service is
# a plain dict, so the whole list is already JSON-serializable as-is.

def build_aws_smc_services(records):
    """
    Discover AWS services required for this migration from the raw input
    records (same field names used throughout the AWS pipeline: 'Pre
    Installed S/W', 'Production Servers').

    Returns a list of dicts: {'Service Group', 'Service', 'Quantity',
    'placeholder'} — conversion-ready for JSON (no custom objects, no
    non-serializable fields).
    """
    total_vms  = len(records)
    sql_count  = sum(1 for r in records if str(r.get('Pre Installed S/W', 'NA')).strip()
                      not in ('NA', 'nan', 'None', ''))
    prod_count = sum(1 for r in records if str(r.get('Production Servers', '')).strip().lower() == 'yes')

    services = []

    # Always created
    services.append({'Service Group': 'Compute',      'Service': 'Amazon EC2', 'Quantity': total_vms, 'placeholder': False})
    services.append({'Service Group': 'Monitoring',    'Service': 'CloudWatch', 'Quantity': total_vms, 'placeholder': False})
    services.append({'Service Group': 'Backup & DR',   'Service': 'AWS Backup', 'Quantity': total_vms, 'placeholder': False})

    # Conditional
    if sql_count > 0:
        services.append({'Service Group': 'Database', 'Service': 'Amazon RDS', 'Quantity': sql_count, 'placeholder': False})

    if prod_count > 5:
        services.append({'Service Group': 'Networking', 'Service': 'Application Load Balancer', 'Quantity': 1, 'placeholder': False})
        services.append({'Service Group': 'Security',   'Service': 'AWS WAF',                     'Quantity': 1, 'placeholder': False})

    # Future-ready placeholders — included in the model now (Quantity 0,
    # flagged) so the uploader/discovery logic never has to change shape
    # later when these become real, sized services.
    services.append({'Service Group': 'Storage',    'Service': 'S3',             'Quantity': 0, 'placeholder': True})
    services.append({'Service Group': 'Storage',    'Service': 'EBS',            'Quantity': 0, 'placeholder': True})
    services.append({'Service Group': 'Networking', 'Service': 'Direct Connect', 'Quantity': 0, 'placeholder': True})

    return services


def generate_smc_file(services, output_file):
    """
    Write a discovered-services list to an .xlsx workbook with a single
    'SMC Services' tab — columns: Service Group | Service | Quantity.
    `services` is the plain-dict list from build_aws_smc_services() (or
    azure_engine.build_smc_services()) — this function is platform-agnostic.
    Returns the number of services written.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'SMC Services'

    headers = ['Service Group', 'Service', 'Quantity']
    ws.append(headers)
    hdr_fill = PatternFill(start_color='1F3864', end_color='1F3864', fill_type='solid')
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF', size=11)
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')

    placeholder_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    for svc in services:
        row = [svc.get('Service Group', ''), svc.get('Service', ''), svc.get('Quantity', 0)]
        ws.append(row)
        if svc.get('placeholder'):
            for cell in ws[ws.max_row]:
                cell.fill = placeholder_fill
                cell.font = Font(italic=True, color='808080')

    for col in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = max(16, length + 4)

    wb.save(str(output_file))
    return len(services)


def generate_html_report(client_name, run_id, summary, filename, timestamp, purchase_key):
    """Generate a standalone HTML dashboard report for a run."""
    ts_fmt = datetime.fromisoformat(timestamp).strftime('%d %b %Y %H:%M')
    s = summary
    def fmt(v):
        if v is None: return '—'
        if v >= 1e6: return f'${v/1e6:.2f}M'
        if v >= 1000: return f'${v/1000:.1f}K'
        return f'${v:,.0f}'
    def fmtG(v):
        if v >= 1000: return f'{v/1000:.1f}T GiB'
        return f'{v:,.0f} GiB'

    wf = s.get('wf_breakdown', {})
    fb = s.get('fam_breakdown', {})
    os_b = s.get('os_breakdown', {})
    grp  = s.get('group_breakdown', {})

    # Donut data helpers
    def donut_js(data_dict, colors):
        labels = list(data_dict.keys())
        vals   = list(data_dict.values())
        cols   = colors[:len(labels)]
        return f"labels:{json.dumps(labels)},data:{json.dumps(vals)},colors:{json.dumps(cols)}"

    prod_data = {'Production': s.get('prod_count',0), 'Non-Prod': s.get('nonprod_count',0)}
    wf_data   = {k: v for k,v in wf.items() if v}

    html = f"""<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"/><title>Cloud Optimizer Report — {client_name}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:system-ui,sans-serif;background:#05070f;color:#eef0ff;min-height:100vh}}
.topbar{{height:3px;background:linear-gradient(90deg,#6c63ff,#00e5c3,#ff6b9d)}}
.header{{padding:20px 32px;border-bottom:1px solid #1e2840;background:#080b16}}
.h-title{{font-size:20px;font-weight:700}}
.h-sub{{font-size:12px;color:#4f5a8a;margin-top:3px;font-family:monospace}}
.body{{padding:22px 32px;display:flex;flex-direction:column;gap:16px}}
.metrics{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px}}
.mc{{background:#0c1020;border:1px solid #1e2840;border-radius:9px;padding:11px 13px}}
.mv{{font-size:16px;font-weight:700;font-family:monospace;margin-bottom:2px}}
.ml{{font-size:9px;color:#4f5a8a;letter-spacing:1px;text-transform:uppercase}}
.band{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}}
.bc{{border-radius:9px;padding:11px 14px;border:1px solid}}
.bv{{font-size:16px;font-weight:700;font-family:monospace}}
.bl{{font-size:9px;letter-spacing:1.2px;text-transform:uppercase;margin-bottom:5px;font-family:monospace}}
.bn{{font-size:10px;color:#4f5a8a;margin-top:3px}}
.row2{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}}
.row3{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
.panel{{background:#0c1020;border:1px solid #1e2840;border-radius:10px;overflow:hidden}}
.ph{{padding:9px 13px;border-bottom:1px solid #1e2840;font-size:12px;font-weight:600}}
.pb{{padding:11px 13px}}
.grp-row{{display:flex;align-items:center;gap:8px;margin-bottom:5px}}
.grp-n{{font-size:11px;color:#8d97d4;min-width:140px;flex-shrink:0}}
.grp-bar{{flex:1;height:4px;background:#1e2840;border-radius:2px;overflow:hidden}}
.grp-fill{{height:100%;border-radius:2px;background:linear-gradient(90deg,#6c63ff,#00e5c3)}}
.grp-c{{font-family:monospace;font-size:10px;color:#eef0ff;min-width:20px;text-align:right;flex-shrink:0}}
.grp-p{{font-family:monospace;font-size:10px;color:#ffb347;min-width:65px;text-align:right;flex-shrink:0}}
.footer{{padding:14px 32px;border-top:1px solid #1e2840;font-size:10px;color:#2e3660;font-family:monospace}}
</style>
</head>
<body>
<div class="topbar"></div>
<div class="header">
  <div class="h-title">Cloud Instance Optimizer — {client_name}</div>
  <div class="h-sub">Run ID: {run_id} &nbsp;·&nbsp; {ts_fmt} &nbsp;·&nbsp; {purchase_key.replace('_',' ').upper()} &nbsp;·&nbsp; {filename}</div>
</div>
<div class="body">

<div class="metrics">
  <div class="mc"><div class="mv" style="color:#6c63ff">{s.get('total_rows','—')}</div><div class="ml">Total VMs</div></div>
  <div class="mc"><div class="mv" style="color:#00e5a0">{s.get('prod_count','—')}</div><div class="ml">Production</div></div>
  <div class="mc"><div class="mv" style="color:#2196f3">{s.get('sql_count','—')}</div><div class="ml">SQL Servers</div></div>
  <div class="mc"><div class="mv" style="color:#ff9900">{fmt(s.get('total_aws_pm'))}</div><div class="ml">AWS $/mo</div></div>
</div>

<div class="band">
  <div class="bc" style="background:rgba(255,153,0,.05);border-color:rgba(255,153,0,.18)">
    <div class="bl" style="color:#ff9900">AWS Total $/yr</div>
    <div class="bv" style="color:#ff9900">{fmt(s.get('total_aws_py'))}</div>
    <div class="bn">{s.get('total_rows','—')} instances</div>
  </div>
  <div class="bc" style="background:rgba(255,179,71,.05);border-color:rgba(255,179,71,.18)">
    <div class="bl" style="color:#ffb347">Upsized VMs</div>
    <div class="bv" style="color:#ffb347">{s.get('upsized_count','—')} / {s.get('total_rows','—')}</div>
    <div class="bn">vCPU or mem rounded up</div>
  </div>
</div>

<div class="band">
  <div class="bc" style="background:rgba(108,99,255,.05);border-color:rgba(108,99,255,.18)">
    <div class="bl" style="color:#9b95ff">Source vCPU</div>
    <div class="bv" style="color:#9b95ff">{s.get('src_vcpu_total',0):,}</div>
    <div class="bn">→ Target: {s.get('tgt_vcpu_total',0):,}</div>
  </div>
  <div class="bc" style="background:rgba(0,229,195,.05);border-color:rgba(0,229,195,.18)">
    <div class="bl" style="color:#00e5c3">Source Memory</div>
    <div class="bv" style="color:#00e5c3">{fmtG(s.get('src_mem_total',0))}</div>
    <div class="bn">→ Target: {fmtG(s.get('tgt_mem_total',0))}</div>
  </div>
  <div class="bc" style="background:rgba(255,107,157,.05);border-color:rgba(255,107,157,.18)">
    <div class="bl" style="color:#ff6b9d">Non-Prod VMs</div>
    <div class="bv" style="color:#ff6b9d">{s.get('nonprod_count','—')}</div>
    <div class="bn">{round(s.get('nonprod_count',0)/(s.get('total_rows',1) or 1)*100)}% of fleet</div>
  </div>
  <div class="bc" style="background:rgba(79,195,247,.05);border-color:rgba(79,195,247,.18)">
    <div class="bl" style="color:#4fc3f7">Pricing Plan</div>
    <div class="bv" style="color:#4fc3f7" style="font-size:12px">{purchase_key.replace('_',' ').title()}</div>
    <div class="bn">Selected at run time</div>
  </div>
</div>

<div class="row2">
  <div class="panel">
    <div class="ph">Production Split</div>
    <div class="pb" style="height:180px;position:relative"><canvas id="prodChart" role="img" aria-label="Production donut"></canvas></div>
  </div>
  <div class="panel">
    <div class="ph">OS Distribution</div>
    <div class="pb" style="height:180px;position:relative"><canvas id="osChart" role="img" aria-label="OS donut"></canvas></div>
  </div>
  <div class="panel">
    <div class="ph">Workload Family</div>
    <div class="pb" style="height:180px;position:relative"><canvas id="wfChart" role="img" aria-label="Workload donut"></canvas></div>
  </div>
</div>

<div class="row3">
  <div class="panel">
    <div class="ph">Group Breakdown</div>
    <div class="pb" id="grpList"></div>
  </div>
  <div class="panel">
    <div class="ph">AWS Instance Family</div>
    <div class="pb">
"""
    for fam, cnt in sorted(fb.items(), key=lambda x: -x[1]):
        fam_colors = {'Compute optimized':'#4fc3f7','General purpose':'#9b95ff','Memory optimized':'#ff6b9d'}
        col = fam_colors.get(fam,'#8d97d4')
        pct = round(cnt/(s.get('total_rows',1) or 1)*100)
        html += f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;padding:7px 10px;background:#101528;border:1px solid #1e2840;border-radius:6px"><div style="width:8px;height:8px;border-radius:50%;background:{col};flex-shrink:0"></div><div style="flex:1;font-size:11px;color:{col}">{fam}</div><div style="font-family:monospace;font-weight:700;color:#eef0ff">{cnt}</div><div style="font-size:10px;color:#4f5a8a">{pct}%</div></div>\n'

    html += f"""    </div>
  </div>
</div>

</div>
<div class="footer">Generated by Cloud Instance Optimizer &nbsp;·&nbsp; {ts_fmt} &nbsp;·&nbsp; Run {run_id}</div>

<script>
const COLORS=['#6c63ff','#00e5a0','#ff9900','#2196f3','#ff6b9d','#4fc3f7','#ffb347','#9b95ff'];
function donut(id,labels,data,colors){{
  const el=document.getElementById(id);
  new Chart(el,{{type:'doughnut',data:{{labels,datasets:[{{data,backgroundColor:colors,borderColor:'#0c1020',borderWidth:2}}]}},
    options:{{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{{legend:{{position:'right',labels:{{color:'#8d97d4',font:{{size:10}},padding:8}}}}}}}}
  }});
}}
donut('prodChart',{json.dumps(list(prod_data.keys()))},{json.dumps(list(prod_data.values()))},['#00e5a0','#6c63ff']);
donut('osChart',{json.dumps(list(os_b.keys()))},{json.dumps(list(os_b.values()))},COLORS);
donut('wfChart',{json.dumps(list(wf_data.keys()))},{json.dumps(list(wf_data.values()))},['#4fc3f7','#9b95ff','#ff6b9d']);

// Group breakdown bars
const grp={json.dumps({g:{'count':v['count'],'pm':v['pm']} for g,v in grp.items()})};
const maxC=Math.max(...Object.values(grp).map(v=>v.count),1);
const gl=document.getElementById('grpList');
Object.entries(grp).sort((a,b)=>b[1].count-a[1].count).forEach(([g,v])=>{{
  const pm=v.pm>=1000?'$'+(v.pm/1000).toFixed(1)+'K':'$'+Math.round(v.pm);
  gl.innerHTML+='<div class="grp-row"><span class="grp-n">'+g+'</span><div class="grp-bar"><div class="grp-fill" style="width:'+Math.round(v.count/maxC*100)+'%"></div></div><span class="grp-c">'+v.count+'</span><span class="grp-p">'+pm+'/mo</span></div>';
}});
</script>
</body></html>"""
    return html

# ── Routes ────────────────────────────────────────────────────────────────────

# ── Embedded UI (no template file needed) ─────────────────────────────────────

@app.route('/')
def index():
    return render_template('UI.html')

@app.route('/api/clients')
def api_clients():
    return jsonify(load_clients())

@app.route('/api/clients/<client_id>/runs/<run_id>/download')
def download_run_output(client_id, run_id):
    clients = load_clients()
    client  = next((c for c in clients if c['id'] == client_id), None)
    if not client: abort(404)
    run = next((r for r in client['runs'] if r['id'] == run_id), None)
    if not run: abort(404)
    path = Path(run.get('output_path',''))
    if not path.exists(): abort(404, 'Output file no longer exists on disk')
    return send_file(path, as_attachment=True, download_name=path.name)

@app.route('/api/clients/<client_id>/runs/<run_id>/download_report')
def download_run_report(client_id, run_id):
    clients = load_clients()
    client  = next((c for c in clients if c['id'] == client_id), None)
    if not client: abort(404)
    run = next((r for r in client['runs'] if r['id'] == run_id), None)
    if not run: abort(404)
    path = Path(run.get('report_path',''))
    if not path.exists(): abort(404, 'Report file not found')
    return send_file(path, as_attachment=True, download_name=path.name, mimetype='text/html')

@app.route('/api/clients/<client_id>/runs/<run_id>/download_bulk')
def download_run_bulk(client_id, run_id):
    clients = load_clients()
    client  = next((c for c in clients if c['id'] == client_id), None)
    if not client: abort(404)
    run = next((r for r in client['runs'] if r['id'] == run_id), None)
    if not run: abort(404)
    path = Path(run.get('bulk_path',''))
    if not path.exists(): abort(404, 'Bulk import file not found')
    return send_file(path, as_attachment=True, download_name=path.name)

@app.route('/api/clients/<client_id>/runs/<run_id>/download_bulk_part/<int:part_no>')
def download_run_bulk_part(client_id, run_id, part_no):
    clients = load_clients()
    client  = next((c for c in clients if c['id'] == client_id), None)
    if not client: abort(404)
    run = next((r for r in client['runs'] if r['id'] == run_id), None)
    if not run: abort(404)
    bulk_files = run.get('bulk_files', [])
    idx = part_no - 1
    if idx < 0 or idx >= len(bulk_files): abort(404, 'Part not found')
    path = OUTPUT_DIR / bulk_files[idx]
    if not path.exists(): abort(404, 'Bulk import part file not found on disk')
    return send_file(path, as_attachment=True, download_name=path.name)

@app.route('/api/clients/<client_id>/runs/<run_id>/download_smc')
def download_run_smc(client_id, run_id):
    """Download the SMC service discovery workbook (AWS_SMC_*.xlsx or Azure_SMC_*.xlsx)."""
    clients = load_clients()
    client  = next((c for c in clients if c['id'] == client_id), None)
    if not client: abort(404)
    run = next((r for r in client['runs'] if r['id'] == run_id), None)
    if not run: abort(404)
    path = Path(run.get('smc_path',''))
    if not path.exists(): abort(404, 'SMC workbook not found')
    return send_file(path, as_attachment=True, download_name=path.name)

@app.route('/api/clients/<client_id>/runs/<run_id>/delete', methods=['DELETE'])
def delete_run(client_id, run_id):
    clients = load_clients()
    client  = next((c for c in clients if c['id'] == client_id), None)
    if not client: return jsonify({'error':'Client not found'}),404
    run = next((r for r in client['runs'] if r['id'] == run_id), None)
    if run:
        # Also remove the physical output files on disk, not just the record
        for key in ('output_path', 'bulk_path', 'report_path'):
            p = Path(run.get(key, '') or '')
            try:
                if p.exists(): p.unlink()
            except Exception:
                pass
    client['runs'] = [r for r in client['runs'] if r['id'] != run_id]
    save_clients(clients)
    return jsonify({'ok':True})

@app.route('/api/clients/<client_id>/runs/<run_id>/edit', methods=['PUT'])
def edit_run(client_id, run_id):
    """Rename/label an individual run entry (does not change any file, just the display label)."""
    data  = request.get_json(force=True, silent=True) or {}
    label = str(data.get('label', '')).strip()
    if not label:
        return jsonify({'error': 'Label cannot be empty'}), 400
    clients = load_clients()
    client  = next((c for c in clients if c['id'] == client_id), None)
    if not client: return jsonify({'error': 'Client not found'}), 404
    run = next((r for r in client['runs'] if r['id'] == run_id), None)
    if not run: return jsonify({'error': 'Run not found'}), 404
    run['label'] = label
    save_clients(clients)
    return jsonify({'ok': True, 'label': label})

@app.route('/api/clients/<client_id>/rename', methods=['PUT'])
def rename_client(client_id):
    """Rename a client. All of their run history stays attached under the new name."""
    data     = request.get_json(force=True, silent=True) or {}
    new_name = str(data.get('name', '')).strip()
    if not new_name:
        return jsonify({'error': 'Name cannot be empty'}), 400
    clients = load_clients()
    client  = next((c for c in clients if c['id'] == client_id), None)
    if not client: return jsonify({'error': 'Client not found'}), 404
    # Prevent renaming into a name that already belongs to a different client
    dup = next((c for c in clients if c['id'] != client_id and c['name'].lower() == new_name.lower()), None)
    if dup:
        return jsonify({'error': f'A client named "{new_name}" already exists'}), 400
    client['name'] = new_name
    save_clients(clients)
    return jsonify({'ok': True, 'name': new_name})

@app.route('/api/clients/<client_id>/delete', methods=['DELETE'])
def delete_client(client_id):
    """Delete an entire client and every run/file that belongs to them."""
    clients = load_clients()
    client  = next((c for c in clients if c['id'] == client_id), None)
    if not client: return jsonify({'error': 'Client not found'}), 404
    for run in client.get('runs', []):
        for key in ('output_path', 'bulk_path', 'report_path'):
            p = Path(run.get(key, '') or '')
            try:
                if p.exists(): p.unlink()
            except Exception:
                pass
    clients = [c for c in clients if c['id'] != client_id]
    save_clients(clients)
    return jsonify({'ok': True})

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error':'No file'}),400
    f = request.files['file']
    if not f.filename.endswith(('.xlsx','.xls')):
        return jsonify({'error':'Only .xlsx/.xls accepted'}),400
    uid      = str(uuid.uuid4())
    path     = UPLOAD_DIR / f'input_{uid}.xlsx'
    platform = (request.form.get('platform') or 'aws').strip().lower()
    f.save(path)

    if platform == 'azure':
        # Azure engine parses its own sheets (vInfo/In-Scope for RVTools, first
        # matching sheet for Adhoc) and its own column layout — AWS's pandas
        # column check doesn't apply here, just confirm the file opens.
        try:
            wb = load_workbook(path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
        except Exception as e:
            return jsonify({'error': f'Cannot read file: {e}'}), 400
        return jsonify({'file_id': uid, 'filename': f.filename,
                        'rows': None, 'columns': sheet_names})

    try:
        df2 = pd.read_excel(path, engine='openpyxl', nrows=2)
    except Exception as e:
        return jsonify({'error':f'Cannot read file: {e}'}),400
    required = {
    'vCPU',
    'Memory',
    'CPU Utilization %',
    'Memory Utilization %',
    'Operating System',
    'Production Servers',
    'Tenancy',
    'License Model',
    'Pre Installed S/W'
}
    missing  = required - set(df2.columns)
    if missing:
        return jsonify({'error':f'Missing columns: {", ".join(missing)}'}),400
    df_full = pd.read_excel(path, engine='openpyxl')
    return jsonify({'file_id':uid,'filename':f.filename,
                    'rows':len(df_full),'columns':list(df_full.columns)})


# ── SocketIO: load AWS pricing data for a region ──────────────────────────────
@socketio.on('load_pricing')
def handle_load_pricing(data):
    """
    Triggered from the Configure tab's 'Load Region Pricing' panel.
    Downloads AWS's official EC2 price list for the requested region and
    loads it into MongoDB, with live progress reported back over the socket.
    """
    sid       = request.sid
    region    = str(data.get('region', '')).strip()
    mongo_uri = data.get('mongo_uri', 'mongodb://localhost:27017/')
    db_name   = data.get('db', 'EC2')

    if not region:
        socketio.emit('pricing_error', {'message': 'No region specified.'}, to=sid)
        return

    def progress_cb(pct, msg):
        socketio.emit('pricing_progress', {'pct': pct, 'message': msg, 'region': region}, to=sid)

    def run():
        try:
            inserted = load_region_pricing(region, mongo_uri=mongo_uri, db_name=db_name,
                                            progress_cb=progress_cb)
            socketio.emit('pricing_complete', {
                'region': region, 'inserted': inserted,
                'message': f'Loaded {inserted:,} pricing records for {region}',
            }, to=sid)
        except Exception as e:
            socketio.emit('pricing_error', {'message': f'Failed to load pricing for {region}: {e}'}, to=sid)

    threading.Thread(target=run, daemon=True).start()


# ── SocketIO: real-time optimizer run ─────────────────────────────────────────
# ── SocketIO: Azure pricing run (RVTools / Adhoc) ─────────────────────────────
@socketio.on('start_azure_run')
def handle_start_azure_run(data):
    """
    Runs the Azure pricing engine (azure_engine.py) against an uploaded
    RVTools export or adhoc VM list. No MongoDB involved — all pricing is
    a self-contained lookup table inside azure_engine.

    The summary dict is deliberately built with the SAME field names the
    AWS pipeline uses (total_aws_pm, prod_count, os_breakdown, etc.) so the
    existing dashboard/renderDash() JS works for Azure runs with zero changes.
    'bulk_filename'/'report_filename' are left empty since Azure runs don't
    produce an AWS Pricing Calculator bulk import or a separate HTML report —
    the Azure Excel workbook already contains the Summary/VM Detail/etc. tabs.
    """
    sid = request.sid

    def ep(pct, step, expr, log_msg, color='#8d97d4'):
        socketio.emit('progress', {'pct': round(pct, 1), 'step': step, 'expr': expr,
                                    'log': log_msg, 'color': color}, to=sid)
    def ee(msg):
        socketio.emit('run_error', {'message': msg}, to=sid)

    def run():
        try:
            client_name = (data.get('client_name', '') or 'Unknown').strip()
            file_id     = data.get('file_id', '')
            azure_mode  = data.get('azure_mode', 'rvtools')   # 'rvtools' | 'adhoc'

            # RVTools-specific flags
            like_for_like    = bool(data.get('like_for_like', True))
            downsize_cpu     = int(data.get('downsize_cpu', 8) or 8)
            downsize_mem     = int(data.get('downsize_mem', 32) or 32)
            prefer_higher_ram = bool(data.get('prefer_higher_ram', False))

            # Adhoc-specific flags
            pricing_model = data.get('pricing_model', 'byol')   # 'li' | 'byol' | 'ri'
            include_sql   = bool(data.get('include_sql', False))

            # Shared flag
            b_nonprod_only = bool(data.get('b_nonprod_only', False))

            input_path = UPLOAD_DIR / f'input_{file_id}.xlsx'
            if not input_path.exists():
                ee('Input file not found. Please re-upload.')
                return

            run_id      = str(uuid.uuid4())
            ts          = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_client = re.sub(r'[^\w\-]', '_', client_name)

            if azure_mode == 'adhoc':
                # ── Adhoc mode ──────────────────────────────────────────────
                use_ri = (pricing_model == 'ri')
                azure_engine.B_SERIES_ALL = not b_nonprod_only

                ep(10, 'Reading file', f'parse_adhoc_file("input_{file_id}.xlsx")',
                   'Reading adhoc VM list...', '#9b95ff')
                vms = azure_engine.parse_adhoc_file(str(input_path))
                total_rows = len(vms)
                if total_rows == 0:
                    ee('No VMs found in the uploaded file. Check the column layout.')
                    return

                ep(40, 'Pricing', f'{total_rows} VMs — model={pricing_model}',
                   f'Pricing {total_rows} VMs ({pricing_model.upper()})...', '#ff9900')

                pm = 'li' if pricing_model == 'li' else 'byol'
                out_fn   = f'Azure_Pricing_{safe_client}_{ts}.xlsx'
                out_path = OUTPUT_DIR / out_fn
                azure_engine.generate_adhoc_excel(vms, str(out_path), pricing_mode=pm,
                                                   include_sql=include_sql, use_ri=use_ri)

                ep(80, 'Writing Excel', 'generate_adhoc_excel(...)',
                   'Writing output XLSX...', '#ffb347')

                ep(88, 'SMC discovery', 'azure_engine.build_smc_services(vms)',
                   'Discovering required Azure services...', '#9b95ff')
                smc_fn       = f'Azure_SMC_{safe_client}_{ts}.xlsx'
                smc_path     = OUTPUT_DIR / smc_fn
                smc_services = azure_engine.generate_smc_workbook(vms, str(smc_path))

                # Build AWS-shaped summary from adhoc results
                prod_count = sum(1 for v in vms if not v['is_nonprod'])
                sql_count  = sum(1 for v in vms if v['is_sql'])
                no_match   = sum(1 for v in vms if v.get('azure_sku') == 'NO MATCH')
                # Use the 3yr tier as the headline monthly total (matches the
                # "committed plan" framing used elsewhere in the dashboard)
                total_pm   = sum(v.get('total_3yr', 0) for v in vms)
                total_py = total_pm * 12
                os_ctr   = Counter(v['os_type'] for v in vms)
                fam_ctr  = Counter(v.get('azure_series', '') for v in vms if v.get('azure_series'))
                pfx_ctr  = Counter((v.get('azure_sku') or '')[:3] for v in vms if v.get('azure_sku') and v.get('azure_sku') != 'NO MATCH')

                summary = {
                    'total_rows': total_rows, 'prod_count': prod_count,
                    'nonprod_count': total_rows - prod_count, 'sql_count': sql_count,
                    'upsized_count': 0, 'downsized_count': 0,
                    'total_aws_pm': round(total_pm, 2), 'total_aws_py': round(total_py, 2),
                    'src_vcpu_total': sum(int(v.get('source_cpus', 0) or 0) for v in vms),
                    'src_mem_total': sum(float(v.get('source_mem_gib', 0) or 0) for v in vms),
                    'tgt_vcpu_total': sum(int(v.get('azure_vcpus', 0) or 0) for v in vms),
                    'tgt_mem_total': sum(float(v.get('azure_mem', 0) or 0) for v in vms),
                    'fam_breakdown': dict(fam_ctr), 'prefix_breakdown': dict(pfx_ctr),
                    'wf_breakdown': {}, 'os_breakdown': dict(os_ctr),
                    'sql_breakdown': {'SQL Standard': sql_count} if sql_count else {},
                    'group_breakdown': {'All VMs': {'count': total_rows, 'pm': round(total_pm, 2)}},
                    'purchase_key': f'azure_{pricing_model}',
                    'calc_mode': f'azure_adhoc_{pricing_model}',
                    'platform': 'azure',
                    'smc_service_count': len(smc_services),
                }

            else:
                # ── RVTools mode ────────────────────────────────────────────
                azure_engine.LIKE_FOR_LIKE = like_for_like
                azure_engine.DOWNSIZE_CPU_THRESHOLD = downsize_cpu
                azure_engine.DOWNSIZE_MEM_THRESHOLD = downsize_mem
                azure_engine.PREFER_HIGHER_RAM = prefer_higher_ram
                azure_engine.B_SERIES_ALL = not b_nonprod_only

                ep(10, 'Reading file', f'parse_rvtools("input_{file_id}.xlsx")',
                   'Reading RVTools export...', '#9b95ff')
                all_vms = azure_engine.parse_rvtools(str(input_path))
                desktop_vms = [v for v in all_vms if v.get('is_desktop', False)]
                vms = [v for v in all_vms if not v.get('is_desktop', False)]
                total_rows = len(vms)
                if total_rows == 0:
                    ee('No VMs to price after filtering (all powered-off/templates/desktop OS?).')
                    return

                ep(35, 'Pricing', f'{total_rows} VMs — sizing + SKU lookup',
                   f'Pricing {total_rows} VMs...', '#ff9900')
                for i, vm in enumerate(vms):
                    azure_engine.price_vm(vm)
                    if i % max(1, total_rows // 15) == 0:
                        pct = 35 + (i / total_rows) * 40
                        socketio.emit('progress', {
                            'pct': round(pct, 1), 'step': 'Pricing',
                            'expr': f'VM {i+1}/{total_rows}: {vm["vm_name"]}',
                            'log': f'Priced {i+1}/{total_rows}', 'color': '#ff9900'
                        }, to=sid)

                ep(80, 'Writing Excel', 'generate_excel(...)',
                   'Writing output XLSX...', '#ffb347')
                out_fn   = f'Azure_Pricing_{safe_client}_{ts}.xlsx'
                out_path = OUTPUT_DIR / out_fn
                azure_engine.generate_excel(vms, str(out_path), desktop_vms=desktop_vms)

                ep(88, 'SMC discovery', 'azure_engine.build_smc_services(vms)',
                   'Discovering required Azure services...', '#9b95ff')
                smc_fn       = f'Azure_SMC_{safe_client}_{ts}.xlsx'
                smc_path     = OUTPUT_DIR / smc_fn
                smc_services = azure_engine.generate_smc_workbook(vms, str(smc_path))

                prod_count = sum(1 for v in vms if not v['is_nonprod'])
                sql_count  = sum(1 for v in vms if v['is_sql'])
                downsized  = sum(1 for v in vms if v.get('was_downsized'))
                total_pm   = sum(v.get('compute_3yr_ri_byol', 0) + v.get('disk_monthly', 0) + v.get('sql_monthly', 0) for v in vms)
                total_py   = total_pm * 12
                os_ctr     = Counter(v['os_type'] for v in vms)
                fam_ctr    = Counter(v.get('azure_series', '') for v in vms if v.get('azure_series'))
                pfx_ctr    = Counter((v.get('azure_sku') or '').replace('Standard_', '')[:4] for v in vms if v.get('azure_sku') and v.get('azure_sku') != 'NO MATCH')
                grp_ctr, grp_cost = {}, {}
                for v in vms:
                    g = v.get('folder') or v.get('cluster') or 'Unknown'
                    grp_ctr[g]  = grp_ctr.get(g, 0) + 1
                    grp_cost[g] = grp_cost.get(g, 0) + v.get('compute_3yr_ri_byol', 0)

                summary = {
                    'total_rows': total_rows, 'prod_count': prod_count,
                    'nonprod_count': total_rows - prod_count, 'sql_count': sql_count,
                    'upsized_count': 0, 'downsized_count': downsized,
                    'total_aws_pm': round(total_pm, 2), 'total_aws_py': round(total_py, 2),
                    'src_vcpu_total': sum(int(v.get('source_cpus', 0) or 0) for v in vms),
                    'src_mem_total': sum(float(v.get('source_mem_gib', 0) or 0) for v in vms),
                    'tgt_vcpu_total': sum(int(v.get('azure_vcpus', 0) or 0) for v in vms),
                    'tgt_mem_total': sum(float(v.get('azure_mem_gib', 0) or 0) for v in vms),
                    'fam_breakdown': dict(fam_ctr), 'prefix_breakdown': dict(pfx_ctr),
                    'wf_breakdown': {}, 'os_breakdown': dict(os_ctr),
                    'sql_breakdown': {'SQL Standard': sql_count} if sql_count else {},
                    'group_breakdown': {g: {'count': grp_ctr[g], 'pm': round(grp_cost[g], 2)} for g in grp_ctr},
                    'purchase_key': 'azure_3yr_ri_hybrid',
                    'calc_mode': 'azure_like_to_like' if like_for_like else 'azure_n1',
                    'platform': 'azure',
                    'smc_service_count': len(smc_services),
                }

            run_meta = {
                'id': run_id, 'timestamp': datetime.now().isoformat(),
                'filename': out_fn, 'output_path': str(out_path),
                'bulk_filename': None, 'bulk_path': None,
                'report_filename': None, 'report_path': None,
                'smc_filename': smc_fn, 'smc_path': str(smc_path),
                'input_rows': total_rows, 'collection': f'Azure {azure_mode.upper()}',
                'purchase_key': summary['purchase_key'], 'calc_mode': summary['calc_mode'],
                'platform': 'azure', 'summary': summary,
            }
            client_rec = upsert_client_run(client_name, run_meta)

            ep(100, 'Complete', f'✓ {out_fn} ready', f'Done — {total_rows} VMs priced', '#00e5a0')
            time.sleep(0.05)

            socketio.emit('run_complete', {
                'client_id': client_rec['id'], 'client_name': client_name,
                'run_id': run_id, 'filename': out_fn,
                'smc_filename': smc_fn,
                'bulk_filename': None, 'report_filename': None,
                'summary': summary,
            }, to=sid)

        except Exception as e:
            import traceback
            ee(f'Unexpected error: {e}\n{traceback.format_exc()}')

    threading.Thread(target=run, daemon=True).start()


@socketio.on('start_run')
def handle_start_run(data):
    sid = request.sid

    def ep(pct, step, expr, log_msg, color='#8d97d4'):
        socketio.emit('progress',{'pct':round(pct,1),'step':step,'expr':expr,
                                   'log':log_msg,'color':color}, to=sid)
    def ee(msg):
        socketio.emit('run_error',{'message':msg}, to=sid)

    def run():
        try:
            client_name     = (data.get('client_name','') or 'Unknown').strip()
            file_id         = data.get('file_id','')
            mongo_uri       = data.get('mongo_uri','mongodb://localhost:27017/')
            aws_db_name     = data.get('aws_db','EC2')
            aws_coll_name   = data.get('aws_collection','us-east-1')
            purchase_key    = data.get('purchase_key','3yr_no_upfront')
            calc_mode       = data.get('calc_mode','like_to_like')   # like_to_like | optimized_n1 | optimized_n2
            min_gen_pref    = int(data.get('min_generation', 5))     # 5 = "5th gen and above", 6 = "6th gen and above"

            lease, purchase, offering = PURCHASE_MAP.get(purchase_key, ('3yr','No Upfront','standard'))

            # Per-run cache — collapse duplicate specs across the whole fleet into one query
            aws_cache = {}

            # ── Step 1: Connect MongoDB ───────────────────────────────────────
            ep(2,'Connecting',f'MongoClient("{mongo_uri}")',f'Connecting to {mongo_uri}...','#00e5c3')
            try:
                mc       = get_mongo(mongo_uri)
                aws_coll = mc[aws_db_name][aws_coll_name]
            except Exception as e:
                ee(f'MongoDB connection failed: {str(e)[:200]}')
                return

            # Any region string can be selected in the UI, but pricing only
            # works if that region's collection has actually been populated
            # with AWS EC2 price-list data. Warn early rather than silently
            # returning "no instance found" for every single row.
            try:
                aws_doc_count = aws_coll.estimated_document_count()
            except Exception:
                aws_doc_count = 0
            if aws_doc_count == 0:
                ee(f'No pricing data found in {aws_db_name}.{aws_coll_name}. '
                   f'This region collection is empty — import AWS EC2 price-list '
                   f'data for "{aws_coll_name}" into MongoDB before running this region.')
                return

            ep(8,'Connected',f'EC2.{aws_coll_name}',
               f'Connected — {aws_db_name}.{aws_coll_name} ({aws_doc_count:,} price records)','#00e5a0')

            # ── Step 2: Read input ────────────────────────────────────────────
            ep(12,'Reading file',f'pd.read_excel("input_{file_id}.xlsx")',
               'Loading Excel input...','#9b95ff')
            input_path = UPLOAD_DIR / f'input_{file_id}.xlsx'
            if not input_path.exists():
                ee('Input file not found. Please re-upload.')
                return
            df = pd.read_excel(input_path, engine='openpyxl')
            df['Pre Installed S/W'] = df['Pre Installed S/W'].fillna('NA').astype(str)
            df['Pre Installed S/W'] = df['Pre Installed S/W'].replace({'nan':'NA','None':'NA','':"NA"})
            df['Production Servers'] = df['Production Servers'].astype(str)
            records    = df.to_dict('records')
            total_rows = len(records)
            ep(18,'File loaded',f'{total_rows} rows','File loaded — '+str(total_rows)+' VMs','#9b95ff')

            # ── Step 3: Parallel AWS queries ──────────────────────────────────
            ep(22,'AWS queries',
               f'aggregate([{{$match:vCPU>=,Memory>=}},{{$sort:PricePerUnit}},{{$limit:1}}]) — cached per unique spec',
               f'Querying AWS for {total_rows} VMs ({calc_mode.replace("_"," ")})...','#ff9900')

            results_aws = {}

            # per-row sizing decisions for output columns (utilization_based mode)
            row_cpu_modes = {}
            row_mem_modes = {}

            def do_aws(args):
                idx, row = args
                vc  = int(row.get('vCPU', 0) or 0)
                mem = float(row.get('Memory', 0) or 0)
                os_ = str(row.get('Operating System', 'Linux'))
                ten = str(row.get('Tenancy', 'Shared'))
                lic = str(row.get('License Model', 'No License required'))
                sw  = str(row.get('Pre Installed S/W', 'NA'))
                ip  = str(row.get('Production Servers', '')).strip().lower() == 'yes'

                if calc_mode == "utilization_based":
                    # NaN-safe read — pandas NaN is a float so "or 0" won't catch it
                    cpu_raw = row.get("CPU Utilization %", 0)
                    mem_raw = row.get("Memory Utilization %", 0)
                    cpu_util = 0.0 if (cpu_raw is None or (isinstance(cpu_raw, float) and pd.isna(cpu_raw))) else float(cpu_raw)
                    mem_util = 0.0 if (mem_raw is None or (isinstance(mem_raw, float) and pd.isna(mem_raw))) else float(mem_raw)

                    cpu_mode = "like_to_like" if cpu_util > 70 else ("optimized_n2" if cpu_util <= 30 else "optimized_n1")
                    mem_mode = "like_to_like" if mem_util > 70 else ("optimized_n2" if mem_util <= 30 else "optimized_n1")

                    cpu_target, _ = resolve_target_dims(cpu_mode, vc, mem)
                    _, mem_target = resolve_target_dims(mem_mode, vc, mem)
                    tgt_vc  = cpu_target
                    tgt_mem = mem_target

                    row_cpu_modes[idx] = cpu_mode
                    row_mem_modes[idx] = mem_mode
                else:
                    tgt_vc, tgt_mem = resolve_target_dims(calc_mode, vc, mem)

                res = cheapest_like_to_like_aws(
                    aws_coll, tgt_vc, tgt_mem, os_, ten, lic, sw,
                    ip, lease, purchase, offering, aws_cache, min_gen_pref
                )

                if idx % max(1, total_rows // 15) == 0:
                    pct = 22 + (idx / total_rows) * 34
                    socketio.emit('progress', {
                        'pct': round(pct, 1),
                        'step': 'AWS Query',
                        'expr': f'Row {idx+1}/{total_rows}: {tgt_vc}vCPU {tgt_mem}GiB {os_}',
                        'log':  f'AWS: {idx+1}/{total_rows} processed',
                        'color': '#ff9900'
                    }, to=sid)

                return idx, res

            # Use ThreadPoolExecutor — mongo operations are I/O bound
            with concurrent.futures.ThreadPoolExecutor(max_workers=12) as ex:
                for idx, res in ex.map(do_aws, enumerate(records)):
                    results_aws[idx] = res

            # ── Step 4: Build columns ─────────────────────────────────────────
            ep(72,'Cost projections','PricePerMonth=PricePerHour×730 | Year=×8760',
               'Computing costs...','#00e5a0')

            aws_vcpus, aws_mems, aws_insts, aws_fams, wf_list,   = [],[],[],[],[]
            ph_list, pm_list, py_list = [],[],[]
            sizing_decisions,cpu_utils,mem_utils = [],[],[]
            # Flag columns
            ds_flag, us_flag = [], []

            for i, row in enumerate(records):
                vc  = int(row.get('vCPU',0) or 0)
                mem = float(row.get('Memory',0) or 0)
                wf_list.append(classify_instance_family(vc, mem).capitalize())

                ai,af,ap,av,am,ad,au = results_aws.get(i) or (None,None,None,None,None,False,False)
                aws_vcpus.append(av)
                aws_mems.append(am)
                aws_insts.append(ai)
                aws_fams.append(af)
                ph = float(ap) if ap is not None else None
                ph_list.append(ph)
                pm_list.append(round(ph*730,4) if ph else None)
                py_list.append(round(ph*730*12,4) if ph else None)

                # Downsize / upsize flags
                try:
                    aws_m_num = float(str(am or '0').split()[0])
                    ds_flag.append('Yes' if aws_m_num < mem else 'No')
                    us_flag.append('Yes' if aws_m_num > mem else 'No')
                except Exception:
                    ds_flag.append('No'); us_flag.append('No')

                # Utilization & sizing decision columns
                if calc_mode == 'utilization_based':
                    cpu_raw = row.get('CPU Utilization %', '')
                    mem_raw = row.get('Memory Utilization %', '')
                    cpu_utils.append('' if (cpu_raw is None or (isinstance(cpu_raw, float) and pd.isna(cpu_raw))) else cpu_raw)
                    mem_utils.append('' if (mem_raw is None or (isinstance(mem_raw, float) and pd.isna(mem_raw))) else mem_raw)
                    mode_label = {'like_to_like': 'Like-to-Like', 'optimized_n1': 'N-1', 'optimized_n2': 'N-2'}
                    sizing_decisions.append(
                        f"CPU:{mode_label.get(row_cpu_modes.get(i,''),'')}, MEM:{mode_label.get(row_mem_modes.get(i,''),'')}"
                    )
                else:
                    cpu_utils.append('')
                    mem_utils.append('')
                    sizing_decisions.append('')

            df['AWS Recommended vCPU']          = aws_vcpus
            df['AWS Recommended Memory']         = aws_mems
            df['AWS Recommended Instance Type']  = aws_insts
            df['Instance Family']                = aws_fams
            df['Workload Family Type']           = wf_list
            df['Downsized']                      = ds_flag
            df['Upsized']                        = us_flag
            df['PricePerHour']                   = ph_list
            df['PricePerMonth']                  = pm_list
            df['PricePerYear']                   = py_list
            if calc_mode == 'utilization_based':
                df['CPU Utilization %']    = cpu_utils
                df['Memory Utilization %'] = mem_utils
                df['Sizing Decision']      = sizing_decisions

            # ── Step 5: Write output Excel ────────────────────────────────────
            ep(80,'Writing Excel','df.to_excel + openpyxl color coding',
               'Writing output XLSX...','#ffb347')

            run_id      = str(uuid.uuid4())
            ts          = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_client = re.sub(r'[^\w\-]','_', client_name)
            out_fn      = f'AWS_Sizing_{safe_client}_{ts}.xlsx'
            out_path    = OUTPUT_DIR / out_fn
            df.to_excel(out_path, index=False)

            # Color coding
            wb  = load_workbook(out_path)
            ws  = wb.active
            grn = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
            org = PatternFill(start_color='FFE0B2', end_color='FFE0B2', fill_type='solid')
            hdrs = [c.value for c in ws[1]]

            # Bold + colored header
            hdr_fill = PatternFill(start_color='1F3864', end_color='1F3864', fill_type='solid')
            for cell in ws[1]:
                cell.fill = hdr_fill
                cell.font = Font(bold=True, color='FFFFFF', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws.row_dimensions[1].height = 30

            ds_ci = hdrs.index('Downsized') if 'Downsized' in hdrs else None
            us_ci = hdrs.index('Upsized')   if 'Upsized'   in hdrs else None
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                fill = None
                if ds_ci and str(row[ds_ci].value).strip()=='Yes': fill=grn
                elif us_ci and str(row[us_ci].value).strip()=='Yes': fill=org
                if fill:
                    for cell in row: cell.fill = fill
            wb.save(out_path)

            # ── Step 6: Bulk import file ──────────────────────────────────────
            ep(86,'Bulk import','Generating AWS Pricing Calculator bulk import XLSX',
               'Building bulk import template...','#4fc3f7')
            if len(records) > 4000:
                bulk_files = []
                for start_idx in range(0, len(records), 4000):
                    end_idx = start_idx + 4000
                    part_records = records[start_idx:end_idx]
                    part_aws     = aws_insts[start_idx:end_idx]
                    part_no = (start_idx // 4000) + 1
                    bulk_file_name = f"BulkImport_{safe_client}_{ts}_Part{part_no}.xlsx"
                    bulk_file_path = OUTPUT_DIR / bulk_file_name
                    bulk_info = generate_bulk_import(
                        part_records,
                        part_aws,
                        aws_coll_name,
                        purchase_key,
                        bulk_file_path,
                        client_name=client_name
                    )
                    bulk_files.append(bulk_file_name)
                bulk_fn = ", ".join(bulk_files)
                bulk_path = ""
                suggested_top_group = bulk_info['suggested_top_group']
                ec2_group_name      = bulk_info['ec2_group_name']
            else:
                bulk_fn = f'BulkImport_{safe_client}_{ts}.xlsx'
                bulk_path = OUTPUT_DIR / bulk_fn
                bulk_info = generate_bulk_import(
                    records,
                    aws_insts,
                    aws_coll_name,
                    purchase_key,
                    bulk_path,
                    client_name=client_name
                )
                suggested_top_group = bulk_info['suggested_top_group']
                ec2_group_name      = bulk_info['ec2_group_name']
            # ── Step 6b: SMC service discovery workbook ───────────────────────
            ep(89, 'SMC discovery', 'build_aws_smc_services(records)',
               'Discovering required AWS services...', '#9b95ff')
            smc_fn       = f'AWS_SMC_{safe_client}_{ts}.xlsx'
            smc_path     = OUTPUT_DIR / smc_fn
            smc_services = build_aws_smc_services(records)
            generate_smc_file(smc_services, smc_path)

            # ── Step 7: Stats ─────────────────────────────────────────────────
            ep(92,'Summary stats','Aggregating fleet metrics','Computing dashboard stats...','#9b95ff')
            total_pm   = sum(v for v in pm_list if v)
            total_py   = sum(v for v in py_list if v)
            prod_count = sum(1 for r in records if str(r.get('Production Servers','')).strip().lower()=='yes')
            sql_count  = sum(1 for r in records if str(r.get('Pre Installed S/W','NA')).strip()
                              not in ('NA','nan','None',''))
            upsized    = sum(1 for v in us_flag if v=='Yes')
            downsized  = sum(1 for v in ds_flag if v=='Yes')
            fam_ctr    = Counter(f for f in aws_fams if f)
            pfx_ctr    = Counter(str(t or '').split('.')[0][:3] for t in aws_insts if t)
            wf_ctr     = Counter(wf_list)
            os_ctr     = Counter(str(r.get('Operating System','')) for r in records)
            sql_ctr    = Counter(str(r.get('Pre Installed S/W','')) for r in records
                                  if str(r.get('Pre Installed S/W','')) not in ('NA','nan','None','','NA'))
            grp_ctr, grp_cost = {}, {}
            for i,row in enumerate(records):
                g = str(row.get('Group','Unknown'))
                grp_ctr[g]  = grp_ctr.get(g,0) + 1
                grp_cost[g] = grp_cost.get(g,0) + (pm_list[i] or 0)
            summary = {
                'total_rows': total_rows, 'prod_count': prod_count,
                'nonprod_count': total_rows-prod_count, 'sql_count': sql_count,
                'upsized_count': upsized, 'downsized_count': downsized,
                'total_aws_pm': round(total_pm,2), 'total_aws_py': round(total_py,2),
                'src_vcpu_total': sum(int(r.get('vCPU',0) or 0) for r in records),
                'src_mem_total':  sum(float(r.get('Memory',0) or 0) for r in records),
                'tgt_vcpu_total': sum(int(v) for v in aws_vcpus if v),
                'tgt_mem_total':  sum(float(str(v or '0').split()[0]) for v in aws_mems),
                'fam_breakdown': dict(fam_ctr), 'prefix_breakdown': dict(pfx_ctr),
                'wf_breakdown': dict(wf_ctr), 'os_breakdown': dict(os_ctr),
                'sql_breakdown': dict(sql_ctr),
                'group_breakdown': {g:{'count':grp_ctr[g],'pm':round(grp_cost[g],2)} for g in grp_ctr},
                'purchase_key': purchase_key,
                'calc_mode': calc_mode,
                'min_generation': min_gen_pref,
                'suggested_top_group': suggested_top_group,
                'ec2_group_name': ec2_group_name,
                'smc_service_count': len(smc_services),
            }

            # ── Step 8: HTML report ───────────────────────────────────────────
            ep(96,'HTML report','Generating standalone dashboard HTML report',
               'Rendering HTML report...','#ff6b9d')
            report_fn   = f'Report_{safe_client}_{ts}.html'
            report_path = OUTPUT_DIR / report_fn
            html_report = generate_html_report(client_name, run_id, summary,
                                                out_fn, datetime.now().isoformat(), purchase_key)
            report_path.write_text(html_report, encoding='utf-8')

            # ── Step 9: Save client record ────────────────────────────────────
            _bulk_files_list = bulk_files if len(records) > 4000 else [bulk_fn]
            run_meta = {
                'id': run_id, 'timestamp': datetime.now().isoformat(),
                'filename': out_fn, 'output_path': str(out_path),
                'bulk_filename': bulk_fn, 'bulk_path': str(bulk_path) if bulk_path else '',
                'bulk_files': _bulk_files_list,
                'report_filename': report_fn, 'report_path': str(report_path),
                'smc_filename': smc_fn, 'smc_path': str(smc_path),
                'input_rows': total_rows, 'collection': aws_coll_name,
                'purchase_key': purchase_key, 'calc_mode': calc_mode, 'min_generation': min_gen_pref,
                'suggested_top_group': suggested_top_group, 'ec2_group_name': ec2_group_name,
                'summary': summary,
            }
            client_rec = upsert_client_run(client_name, run_meta)

            ep(100,'Complete',f'✓ {out_fn} ready',f'Done — {total_rows} VMs processed','#00e5a0')
            time.sleep(0.05)

            socketio.emit('run_complete',{
                'client_id': client_rec['id'],
                'client_name': client_name,
                'run_id': run_id,
                'filename': out_fn,
                'smc_filename': smc_fn,
                'bulk_filename': bulk_fn,
                'bulk_files': bulk_files if len(records) > 4000 else [bulk_fn],
                'report_filename': report_fn,
                'suggested_top_group': suggested_top_group,
                'ec2_group_name': ec2_group_name,
                'summary': summary,
            }, to=sid)

        except Exception as e:
            import traceback
            ee(f'Unexpected error: {e}\n{traceback.format_exc()}')

    t = threading.Thread(target=run, daemon=True)
    t.start()

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    cli = argparse.ArgumentParser(add_help=False)
    cli.add_argument('--load-pricing', action='store_true',
                      help='Load AWS EC2 pricing data for one or more regions, then exit (no web server).')
    cli.add_argument('--regions', nargs='+', default=[],
                      help='AWS region codes to load, e.g. af-south-1 us-west-1')
    cli.add_argument('--mongo-uri', default='mongodb://localhost:27017/')
    cli.add_argument('--db', default='EC2')
    cli_args, _unknown = cli.parse_known_args()

    if cli_args.load_pricing:
        if not cli_args.regions:
            print('ERROR: --load-pricing requires --regions, e.g.:')
            print('  python "With SMC.py" --load-pricing --regions af-south-1 us-west-1')
            sys.exit(1)
        print('\n' + '='*70)
        print('  AWS EC2 Pricing Data Loader')
        print('='*70)
        for region in cli_args.regions:
            print(f'\n─── Region: {region} ───')
            def cb(pct, msg):
                print(f'  [{pct:3d}%] {msg}')
            try:
                load_region_pricing(region, mongo_uri=cli_args.mongo_uri, db_name=cli_args.db, progress_cb=cb)
            except Exception as e:
                print(f'  ERROR loading {region}: {e}')
        print('\n' + '='*70)
        print('  Done. These regions are now ready to select in the app.')
        print('='*70 + '\n')
        sys.exit(0)

    print('\n' + '='*60)
    PORT = int(os.environ.get('PORT', 5001))
    print(f'  Cloud Instance Optimizer  —  http://localhost:{PORT}')
    print('='*60)
    print(f'  Base dir:  {BASE_DIR}')
    print(f'  Uploads:   {UPLOAD_DIR}')
    print(f'  Outputs:   {OUTPUT_DIR}')
    print(f'  Clients:   {CLIENTS_FILE}')
    ui_in_templates = (BASE_DIR / 'templates' / 'UI.html').exists()
    ui_in_basedir   = (BASE_DIR / 'UI.html').exists()
    if ui_in_templates:
        print(f'  Template:  FOUND at {BASE_DIR / "templates" / "UI.html"}')
    elif ui_in_basedir:
        print(f'  Template:  FOUND at {BASE_DIR / "UI.html"} (fallback location — consider moving into templates/)')
    else:
        print(f'  Template:  *** NOT FOUND *** — place UI.html in {BASE_DIR / "templates"}')
    print('  Tip: load pricing data for a new region without starting the server:')
    print('       python "With SMC.py" --load-pricing --regions af-south-1 us-west-1')
    print('='*60 + '\n')
    socketio.run(app, host='0.0.0.0', port=PORT, debug=False, allow_unsafe_werkzeug=True)
