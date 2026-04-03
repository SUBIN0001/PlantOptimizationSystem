"""
Ashok Leyland Plant Location Decision System - Backend API
Supports: AHP + Entropy + TOPSIS + Monte Carlo Simulation
"""

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Dict, Any, Optional
from mangum import Mangum          # ✅ was missing
import pandas as pd
import numpy as np
from io import BytesIO, StringIO
import json
import re
from scipy.stats import rankdata
import xlsxwriter
import uvicorn

app = FastAPI(title="Ashok Leyland Plant Location Decision API", version="2.0")

# ✅ Only ONE CORSMiddleware — removed the duplicate
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://plantoptimizationsystem.vercel.app",  # no trailing slash
        "http://localhost:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Mangum handler after app + middleware setup
handler = Mangum(app)
# ═════════════════════════════════════════════════════════════════════════════
# DATA MODELS
# ═════════════════════════════════════════════════════════════════════════════

class LocationData(BaseModel):
    name: str
    region: str
    state: str
    vendorBase: float
    manpowerAvailability: float
    capex: float
    govtNorms: float
    logisticsCost: float
    economiesOfScale: float

class Constraint(BaseModel):
    key: str
    label: str
    operator: str  # "gte", "lte", "eq"
    value: float
    enabled: bool
    isCost: bool

class AnalysisRequest(BaseModel):
    locations: List[Dict[str, Any]]
    pairwiseMatrix: List[List[float]]
    constraints: List[Constraint]

class MonteCarloRequest(BaseModel):
    locations: List[Dict[str, Any]]
    weights: List[Dict[str, Any]]
    iterations: int = 1000
    constraints: List[Constraint]
    regionFilter: Optional[Dict[str, Any]] = None

class ExportRequest(BaseModel):
    locations: List[Dict[str, Any]]
    results: List[Dict[str, Any]]
    weights: List[Dict[str, Any]]
    pairwiseMatrix: List[List[float]]
    constraints: List[Constraint]
    regionFilter: Optional[Dict[str, Any]] = None

# ═════════════════════════════════════════════════════════════════════════════
# SHARED UTILITIES  (single definition — used by both format parsers)
# ═════════════════════════════════════════════════════════════════════════════

def extract_numeric(val, default=0.0):
    """Extract first numeric value from a string; handles ₹, commas, +, 'Cr', 'km'."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return float(default)
    val_str = str(val).strip().replace(',', '').replace('+', '')
    numbers = re.findall(r'\d+\.?\d*', val_str)
    return float(numbers[0]) if numbers else float(default)

def normalize_col(name: str) -> str:
    """Lowercase, strip spaces/newlines/underscores for fuzzy column matching."""
    return str(name).strip().lower().replace(' ', '').replace('\n', '').replace('_', '')

def find_col_normalized(row, target: str):
    """Return the actual key whose normalized form matches *target*, or None."""
    target_norm = normalize_col(target)
    for k in row.keys():
        if normalize_col(k) == target_norm:
            return k
    return None

# ═════════════════════════════════════════════════════════════════════════════
# DETAILED FORMAT — score computation helpers (finaleyy.xlsx)
# ═════════════════════════════════════════════════════════════════════════════

def rating_to_score(rating):
    """Convert text ratings to numeric scores (1-10 scale)."""
    if pd.isna(rating):
        return 5.0
    rating_map = {
        'Very High': 10, 'High': 8, 'Medium': 6, 'Low': 4, 'Very Low': 2,
        'Very Mature': 10, 'Mature': 8, 'Moderate': 6, 'Emerging': 5,
        'Nascent': 3, 'Nascent-Growing': 4,
        'Yes': 10, 'Partial': 5, 'No': 0,
    }
    return float(rating_map.get(str(rating).strip(), 5))

def compute_vendor_base_score(row):
    acma = str(row.get('No. of ACMA Member Units  (State, approx.)', '0'))
    acma_num = 800 if '800+' in acma else extract_numeric(acma)
    tier1     = extract_numeric(row.get('Tier-1 Auto Vendors  within 200 km (nos.)', 0))
    tier2     = extract_numeric(row.get('Tier-2 Auto Vendors  within 200 km (nos.)', 0))
    steel     = extract_numeric(row.get('Steel / Castings Suppliers  within 100 km (nos.)', 0))
    ecosystem = rating_to_score(row.get('Vendor Ecosystem  Rating', 'Medium'))
    score = (
        (min(acma_num / 800, 1) * 0.25) +
        (min(tier1 / 95,  1) * 0.30) +
        (min(tier2 / 145, 1) * 0.25) +
        (min(steel / 62,  1) * 0.10) +
        (ecosystem / 10   * 0.10)
    ) * 10
    return round(score, 2)

def compute_manpower_score(row):
    engg_colleges = extract_numeric(row.get('AICTE Engg Colleges  in 50 km radius (nos.)', 0))
    iti           = extract_numeric(row.get('ITI Institutes  in 50 km radius (nos.)', 0))
    iti_grad      = extract_numeric(row.get('Annual ITI Graduates  (State, 000s)', 0))
    engg_grad     = extract_numeric(row.get('Annual Engg Graduates  (State, 000s)', 0))
    availability  = rating_to_score(row.get('Skilled Labour  Availability Rating', 'Medium'))
    score = (
        (min(engg_colleges / 62,  1) * 0.20) +
        (min(iti / 55,            1) * 0.15) +
        (min(iti_grad / 95,       1) * 0.20) +
        (min(engg_grad / 146,     1) * 0.25) +
        (availability / 10        * 0.20)
    ) * 10
    return round(score, 2)

def compute_govt_score(row):
    subsidy      = extract_numeric(row.get('Capital Subsidy  (% of Fixed Assets)', 0))
    sgst         = extract_numeric(row.get('SGST Exemption /  Refund Period (yrs)', 0))
    env_ease     = extract_numeric(row.get('Env. Clearance  Ease (1-10)', 5))
    approval_days = extract_numeric(row.get('Single Window  Approval Days (est.)', 50))
    approval_score = max(0, 10 - (approval_days / 6))
    score = (
        (min(subsidy / 35, 1) * 0.30) +
        (min(sgst / 7,     1) * 0.20) +
        (env_ease / 10     * 0.25) +
        (approval_score / 10 * 0.25)
    ) * 10
    return round(score, 2)

def compute_logistics_score(row):
    return extract_numeric(row.get('Annual Logistics Cost  (₹ Cr/yr, est.)**', 0))

def compute_economies_score(row):
    maturity       = rating_to_score(row.get('Auto Industry Cluster  Maturity', 'Medium'))
    supplier_park  = rating_to_score(row.get('Supplier Park  Availability', 'No'))
    export_hub     = rating_to_score(row.get('Export Hub  Proximity', 'Low'))
    market_demand  = extract_numeric(row.get('Market Demand  Index (1-10)', 5))
    cluster_benefit = extract_numeric(row.get('Cluster Benefit  Score (1-10)', 5))
    score = (
        (maturity / 10        * 0.30) +
        (supplier_park / 10   * 0.20) +
        (export_hub / 10      * 0.15) +
        (market_demand / 10   * 0.20) +
        (cluster_benefit / 10 * 0.15)
    ) * 10
    return round(score, 2)

# ═════════════════════════════════════════════════════════════════════════════
# SIMPLIFIED FORMAT — parsers & row processor
# ═════════════════════════════════════════════════════════════════════════════

LABEL_MAP = {
    'Very High': 9.0,
    'VeryHigh':  9.0,
    'High':      7.0,
    'Medium':    5.0,
    'Low':       3.0,
    'Very Low':  1.0,
    'VeryLow':   1.0,
}

def label_to_score(value) -> float:
    """Convert text ratings to numeric scores."""
    if pd.isna(value):
        return 0.0
    s = str(value).strip()
    # Try exact match first
    if s in LABEL_MAP:
        return LABEL_MAP[s]
    # Try case-insensitive match
    s_lower = s.lower()
    for key, val in LABEL_MAP.items():
        if key.lower() == s_lower:
            return val
    # Try to extract numeric
    return extract_numeric(s, 0.0)

def parse_capex(value) -> float:
    """'₹ 62 Cr' → 62.0  |  '2.8 Cr' → 2.8  |  '2.8' → 2.8"""
    if pd.isna(value):
        return 0.0
    s = str(value).replace('₹', '').replace(',', '').strip().lower()
    # Remove 'cr' or 'crore' suffix
    s = re.sub(r'\s*cr\s*$', '', s)
    s = re.sub(r'\s*crore\s*$', '', s)
    # Remove any other text
    m = re.search(r'[\d.]+', s)
    return float(m.group()) if m else 0.0

def parse_govtnorms(value) -> float:
    """'20 km' → 20.0  |  'High' → 7.0"""
    if pd.isna(value):
        return 0.0
    s = str(value).strip()
    # Check if it contains 'km' - treat as distance
    if 'km' in s.lower():
        return extract_numeric(s, 0.0)
    # Otherwise treat as rating
    return label_to_score(s)

def parse_logistics(value) -> float:
    """
    'Very High' → 9.0  
    'High' → 7.0  
    'Medium' → 5.0  
    'Low' → 3.0  
    'Very Low' → 1.0
    Numeric string → float
    """
    if pd.isna(value):
        return 0.0
    s = str(value).strip()
    # Check if it's a text rating
    if s in LABEL_MAP:
        return LABEL_MAP[s]
    # Try to extract numeric value
    try:
        # Remove any currency symbols and units
        cleaned = s.replace('₹', '').replace('Cr', '').replace('cr', '').replace('crore', '').replace(',', '').strip()
        numbers = re.findall(r'[\d.]+', cleaned)
        if numbers:
            return float(numbers[0])
    except:
        pass
    return 0.0

def parse_manpower(value) -> float:
    """'~95,000 / yr' → 95000.0"""
    return extract_numeric(value, 0.0)

def _gn_simplified(row, *col_names, parser=None) -> float:
    """
    Try each candidate column name (in order); return the parsed value of the
    first one found.  Falls back to 0.0 if none found.
    """
    for name in col_names:
        key = find_col_normalized(row, name)
        if key is not None:
            raw = row[key]
            if not (isinstance(raw, float) and np.isnan(raw)):
                return parser(raw) if parser else extract_numeric(raw, 0.0)
    return 0.0

def _gs_simplified(row, *col_names) -> str:
    for name in col_names:
        key = find_col_normalized(row, name)
        if key is not None:
            v = row[key]
            if not (isinstance(v, float) and np.isnan(v)):
                return str(v).strip()
    return ''

def process_simplified_row(row) -> dict:
    """Process one row from the simplified / pre-scored Excel format."""

    # Helper to get raw string value
    def get_raw_str(*col_names):
        for name in col_names:
            key = find_col_normalized(row, name)
            if key is not None:
                v = row[key]
                if not (isinstance(v, float) and np.isnan(v)):
                    return str(v).strip()
        return ''

    # Helper to get numeric value with optional parser
    def get_numeric(parser=None, *col_names):
        for name in col_names:
            key = find_col_normalized(row, name)
            if key is not None:
                raw = row[key]
                if not (isinstance(raw, float) and np.isnan(raw)):
                    if parser:
                        return parser(raw)
                    return extract_numeric(raw, 0.0)
        return 0.0

    return {
        'name':   get_raw_str('Location'),
        'region': get_raw_str('Region'),
        'state':  get_raw_str('State'),
        'vendorBase': get_numeric(None, 'vendorBase', 'Vendor base', 'Vendor Base'),
        'manpowerAvailability': get_numeric(None, 'manpowerAvailability', 'manpower availability', 'Manpower Availability'),
        'capex': get_numeric(parse_capex, 'capex', 'CAPEX', 'Capex'),
        'govtNorms': get_numeric(parse_govtnorms, 'govtNorms', 'govtnorms', 'Govt Norms'),
        'logisticsCost': get_numeric(parse_logistics, 'logisticsCost', 'logisticscost', 'Logistics Cost'),
        'economiesOfScale': get_numeric(lambda x: label_to_score(x), 'economiesOfScale', 'economiesofscale', 'Economies of Scale'),
        'raw': {}   # no sub-attributes in simplified format
    }

# ═════════════════════════════════════════════════════════════════════════════
# FORMAT DETECTION
# ═════════════════════════════════════════════════════════════════════════════

def is_simplified_format(df: pd.DataFrame) -> bool:
    """
    Return True when the sheet is the pre-scored simplified format.
    Heuristic: Check for presence of key simplified column names.
    """
    simplified_keys = {
        'vendorbase', 'vendor base', 'capex',
        'govtnorms', 'govt norms', 'logisticscost', 'logistics cost',
        'economiesofscale', 'economies of scale', 'manpoweravailability', 'manpower availability'
    }
    cols_norm = {normalize_col(c) for c in df.columns}

    # Count matches (normalize both sides)
    matches = 0
    for key in simplified_keys:
        key_norm = normalize_col(key)
        if key_norm in cols_norm:
            matches += 1

    # If we have at least 3 of the key fields, treat as simplified format
    return matches >= 3

# ═════════════════════════════════════════════════════════════════════════════
# MAIN DATA PROCESSOR — handles both formats
# ═════════════════════════════════════════════════════════════════════════════

def process_excel_data(df: pd.DataFrame) -> list:
    """
    Convert a DataFrame (from either Excel format) into a list of location dicts.

    Supported formats
    -----------------
    1. Detailed (finaleyy.xlsx)  — many sub-attribute columns; scores computed here.
    2. Simplified (pre-scored)   — one aggregated column per dimension.
    """
    # Normalise column names (remove embedded newlines)
    df.columns = [str(col).replace('\n', ' ').strip() for col in df.columns]

    # Drop rows where Location is blank or is a repeated header
    if 'Location' in df.columns:
        df = df[df['Location'].notna()]
        df = df[df['Location'] != 'Location']

    use_simplified = is_simplified_format(df)

    locations = []
    for _, row in df.iterrows():
        try:
            if use_simplified:
                loc = process_simplified_row(row)
            else:
                # ── detailed format (original logic, unchanged) ───────────────
                def gs(col, default=''):
                    v = row.get(col, default)
                    return str(v).strip() if not pd.isna(v) else str(default)

                def find_col(keywords, default=''):
                    for k in row.keys():
                        k_lower = str(k).lower().replace('\n', ' ')
                        if all(kw.lower() in k_lower for kw in keywords):
                            return row[k]
                    return default

                def gs_find(keywords, fallback_key):
                    v = find_col(keywords, None)
                    if v is not None:
                        return str(v).strip() if not pd.isna(v) else ''
                    return gs(fallback_key)

                def ext_num_find(keywords, fallback_key):
                    v = find_col(keywords, None)
                    if v is not None:
                        return extract_numeric(v)
                    return extract_numeric(row.get(fallback_key, 0))

                loc = {
                    'name':   str(row['Location']),
                    'region': gs('Region'),
                    'state':  gs('State'),
                    'vendorBase':           compute_vendor_base_score(row),
                    'manpowerAvailability': compute_manpower_score(row),
                    'capex':                extract_numeric(row.get('Estimated Total  Project CAPEX (₹ Cr)*', 0)),
                    'govtNorms':            compute_govt_score(row),
                    'logisticsCost':        compute_logistics_score(row),
                    'economiesOfScale':     compute_economies_score(row),
                    'raw': {
                        'industrialPark':      gs('Industrial Park / Zone'),
                        'acmaCluster':         gs('ACMA Auto Component Cluster'),
                        'acmaUnits':           gs('No. of ACMA Member Units  (State, approx.)'),
                        'tier1Vendors':        extract_numeric(row.get('Tier-1 Auto Vendors  within 200 km (nos.)', 0)),
                        'tier2Vendors':        extract_numeric(row.get('Tier-2 Auto Vendors  within 200 km (nos.)', 0)),
                        'steelSuppliers':      extract_numeric(row.get('Steel / Castings Suppliers  within 100 km (nos.)', 0)),
                        'vendorEcosystem':     gs('Vendor Ecosystem  Rating'),
                        'keyOEMs':             gs('Key OEMs / Anchors  in the Cluster'),
                        'enggColleges':        extract_numeric(row.get('AICTE Engg Colleges  in 50 km radius (nos.)', 0)),
                        'itiInstitutes':       extract_numeric(row.get('ITI Institutes  in 50 km radius (nos.)', 0)),
                        'itiGraduates':        extract_numeric(row.get('Annual ITI Graduates  (State, 000s)', 0)),
                        'enggGraduates':       extract_numeric(row.get('Annual Engg Graduates  (State, 000s)', 0)),
                        'skilledLabourRating': gs('Skilled Labour  Availability Rating'),
                        'wageSkilled':         extract_numeric(row.get('Avg Monthly Wage –  Skilled Mfg (₹)', 0)),
                        'wageSemiSkilled':     extract_numeric(row.get('Avg Monthly Wage –  Semi-Skilled (₹)', 0)),
                        'attritionRate':       extract_numeric(row.get('Labour Attrition Rate  (%/yr, est.)', 0)),
                        'landCost':            extract_numeric(row.get('Industrial Land Cost  (₹ Cr / Acre)', 0)),
                        'availableLand':       extract_numeric(row.get('Available Land  (Acres, approx.)', 0)),
                        'constructionIndex':   extract_numeric(row.get('Construction Cost  Index (Base TN=100)', 0)),
                        'powerCapex':          extract_numeric(row.get('Power Connection  Capex (₹ Cr, est.)', 0)),
                        'waterCapex':          extract_numeric(row.get('Water / Utilities  Capex (₹ Cr, est.)', 0)),
                        'totalCapex':          extract_numeric(row.get('Estimated Total  Project CAPEX (₹ Cr)*', 0)),
                        'industrialPolicy':    gs_find(['Industrial', 'Policy'], 'State Industrial Policy  (Current)'),
                        'capitalSubsidy':      extract_numeric(row.get('Capital Subsidy  (% of Fixed Assets)', 0)),
                        'sgstExemption':       extract_numeric(row.get('SGST Exemption /  Refund Period (yrs)', 0)),
                        'stampDuty':           gs('Stamp Duty  Exemption'),
                        'powerTariff':         extract_numeric(row.get('Power Tariff – HT  Industrial (₹/kWh)', 0)),
                        'elecDutyExemption':   gs('Electricity Duty  Exemption'),
                        'envClearanceEase':    extract_numeric(row.get('Env. Clearance  Ease (1-10)', 0)),
                        'approvalDays':        extract_numeric(row.get('Single Window  Approval Days (est.)', 0)),
                        'sezNimz':             gs('SEZ / NIMZ /  Special Zone'),
                        'dfcAccessGovt':       gs('Dedicated Freight  Corridor Access'),
                        'nearestPort':         gs_find(['Nearest', 'Port'], 'Nearest Major Port'),
                        'distanceToPort':      extract_numeric(row.get('Distance to Port  (km)', 0)),
                        'roadConnectivity':    ext_num_find(['Road', 'Connectivity'], 'Road / NH  Connectivity (1-10)'),
                        'railConnectivity':    extract_numeric(row.get('Rail Connectivity  (1-10)', 0)),
                        'dfcLogistics':        gs('DFC Access  (Y/N)'),
                        'distanceKeyMarket':   extract_numeric(row.get('Distance to Nearest  Key Market (km)', 0)),
                        'keyMarketCity':       gs('Key Market City'),
                        'inboundFreight':      ext_num_find(['Inbound Freight'], 'Inbound Freight Rate  (₹/MT)'),
                        'outboundFreight':     ext_num_find(['Outbound Freight'], 'Outbound Freight Rate  (₹/MT)'),
                        'annualLogisticsCost': extract_numeric(row.get('Annual Logistics Cost  (₹ Cr/yr, est.)**', 0)),
                        'clusterMaturity':     gs('Auto Industry Cluster  Maturity'),
                        'existingCVOEMs':      gs_find(['CV', 'OEMs'], 'Existing CV / Commercial  OEMs nearby'),
                        'supplierPark':        gs('Supplier Park  Availability'),
                        'exportHub':           gs('Export Hub  Proximity'),
                        'marketDemandIndex':   extract_numeric(row.get('Market Demand  Index (1-10)', 0)),
                        'clusterBenefitScore': extract_numeric(row.get('Cluster Benefit  Score (1-10)', 0)),
                    }
                }

            # Skip rows that produced an empty name (blank / junk rows)
            if loc.get('name', '').strip() in ('', 'nan'):
                continue

            locations.append(loc)

        except Exception as e:
            loc_name = row.get('Location', 'unknown')
            print(f"Error processing row '{loc_name}': {e}")
            continue

    return locations

# ═════════════════════════════════════════════════════════════════════════════
# MCDM ALGORITHMS: AHP + Entropy + TOPSIS
# ═════════════════════════════════════════════════════════════════════════════

def calculate_ahp_weights(matrix):
    matrix = np.array(matrix)
    n = matrix.shape[0]
    col_sums = np.sum(matrix, axis=0)
    normalized = matrix / col_sums
    weights = np.mean(normalized, axis=1)
    lambda_max = np.sum((matrix @ weights) / weights) / n
    ci = (lambda_max - n) / (n - 1)
    ri = [0, 0, 0.58, 0.90, 1.12, 1.24, 1.32, 1.41, 1.45, 1.49][min(n - 1, 9)]
    cr = ci / ri if ri > 0 else 0
    return weights.tolist(), cr

def calculate_entropy_weights(data_matrix):
    X = np.array(data_matrix)
    n, m = X.shape
    if n <= 1:
        return [1.0 / m] * m
    col_sums = np.where(np.sum(X, axis=0) == 0, 1e-10, np.sum(X, axis=0))
    P = X / col_sums
    P = np.where(P == 0, 1e-10, P)
    k = 1 / np.log(n)
    E = -k * np.sum(P * np.log(P), axis=0)
    D = 1 - E
    d_sum = np.sum(D)
    weights = D / d_sum if d_sum != 0 else [1.0 / m] * m
    return weights.tolist()

def topsis_analysis(locations, hybrid_weights, criteria_keys, is_cost):
    n = len(locations)
    m = len(criteria_keys)
    X = np.zeros((n, m))
    for i, loc in enumerate(locations):
        for j, key in enumerate(criteria_keys):
            X[i, j] = float(loc.get(key, 0))
    col_norms = np.sqrt(np.sum(X ** 2, axis=0))
    col_norms = np.where(col_norms == 0, 1, col_norms)
    R = X / col_norms
    W = np.array(hybrid_weights)
    V = R * W
    A_plus  = np.zeros(m)
    A_minus = np.zeros(m)
    for j, key in enumerate(criteria_keys):
        if is_cost.get(key, False):
            A_plus[j]  = np.min(V[:, j])
            A_minus[j] = np.max(V[:, j])
        else:
            A_plus[j]  = np.max(V[:, j])
            A_minus[j] = np.min(V[:, j])
    S_plus  = np.sqrt(np.sum((V - A_plus)  ** 2, axis=1))
    S_minus = np.sqrt(np.sum((V - A_minus) ** 2, axis=1))
    scores  = S_minus / (S_plus + S_minus + 1e-10)
    return scores.tolist()

def apply_constraints(locations, constraints, region_filter=None):
    feasible   = []
    infeasible = []
    for loc in locations:
        is_feasible = True
        reasons = []
        for c in constraints:
            if not c.get('enabled', False):
                continue
            key       = c['key']
            val       = float(loc.get(key, 0))
            threshold = float(c['value'])
            op        = c['operator']
            passes = (
                val >= threshold if op == 'gte' else
                val <= threshold if op == 'lte' else
                abs(val - threshold) < 0.001
            )
            if not passes:
                is_feasible = False
                reasons.append(f"{c['label']} {op} {threshold}")
        if region_filter and region_filter.get('regionFilterEnabled'):
            selected = region_filter.get('selectedRegions', [])
            if selected:
                if loc.get('region') not in selected and loc.get('state') not in selected:
                    is_feasible = False
                    reasons.append("Region/State filter")
        if is_feasible:
            feasible.append(loc)
        else:
            loc['infeasible_reasons'] = reasons
            infeasible.append(loc)
    return feasible, infeasible

# ═════════════════════════════════════════════════════════════════════════════
# MONTE CARLO SIMULATION
# ═════════════════════════════════════════════════════════════════════════════

def monte_carlo_simulation(locations, base_weights, criteria_keys, is_cost,
                           iterations=1000, weight_perturb=0.15, value_perturb=0.05):
    n = len(locations)
    m = len(criteria_keys)
    all_ranks    = {loc['name']: [] for loc in locations}
    base_weights = np.array(base_weights)

    for _ in range(iterations):
        weight_noise      = 1 + np.random.uniform(-weight_perturb, weight_perturb, m)
        perturbed_weights = base_weights * weight_noise
        perturbed_weights = perturbed_weights / np.sum(perturbed_weights)

        perturbed_locations = []
        for loc in locations:
            new_loc = loc.copy()
            for key in criteria_keys:
                noise = 1 + np.random.uniform(-value_perturb, value_perturb)
                new_loc[key] = float(loc.get(key, 0)) * noise
            perturbed_locations.append(new_loc)

        scores = topsis_analysis(perturbed_locations, perturbed_weights, criteria_keys, is_cost)
        ranks  = rankdata([-s for s in scores], method='min')
        for i, loc in enumerate(locations):
            all_ranks[loc['name']].append(int(ranks[i]))

    results = []
    for loc in locations:
        ranks     = all_ranks[loc['name']]
        avg_rank  = np.mean(ranks)
        std_rank  = np.std(ranks)
        rank_counts = np.bincount(ranks, minlength=n + 1)[1:]
        rank_probs  = rank_counts / iterations
        ci_low  = max(1, int(np.percentile(ranks, 2.5)))
        ci_high = min(n, int(np.percentile(ranks, 97.5)))
        results.append({
            'locationId':          loc.get('name', ''),
            'locationName':        loc.get('name', ''),
            'avgRank':             round(avg_rank, 2),
            'stdRank':             round(std_rank, 2),
            'confidenceInterval':  [ci_low, ci_high],
            'rankProbabilities':   rank_probs.tolist(),
            'bestRank':            int(min(ranks)),
            'worstRank':           int(max(ranks)),
        })

    results.sort(key=lambda x: x['avgRank'])
    return results

# ═════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═════════════════════════════════════════════════════════════════════════════

def create_excel_report(data: Dict, output_buffer: BytesIO):
    wb = xlsxwriter.Workbook(output_buffer, {'in_memory': True})

    colors = {
        'primary_dark':   '#1a237e',
        'primary':        '#283593',
        'primary_light':  '#5c6bc0',
        'accent':         '#00acc1',
        'accent_light':   '#4dd0e1',
        'success':        '#2e7d32',
        'success_light':  '#a5d6a7',
        'warning':        '#f57c00',
        'warning_light':  '#ffcc80',
        'danger':         '#c62828',
        'danger_light':   '#ef9a9a',
        'neutral':        '#455a64',
        'neutral_light':  '#cfd8dc',
        'gold':           '#ffd700',
        'silver':         '#c0c0c0',
        'bronze':         '#cd7f32',
        'white':          '#ffffff',
        'bg_light':       '#f5f5f5',
    }

    def fmt(**kw):
        return wb.add_format(kw)

    title_main = fmt(bold=True, font_size=16, font_color=colors['primary_dark'],
                     align='left', valign='vcenter')
    title_sub  = fmt(bold=True, font_size=12, font_color=colors['primary'],
                     align='left', valign='vcenter')

    def sec_hdr(bg):
        return fmt(bold=True, bg_color=bg, font_color=colors['white'],
                   border=1, align='center', valign='vcenter', text_wrap=True)

    header_loc  = sec_hdr(colors['primary_dark'])
    header_vnd  = sec_hdr('#1565c0')
    header_man  = sec_hdr('#00695c')
    header_cap  = sec_hdr('#ef6c00')
    header_gov  = sec_hdr('#6a1b9a')
    header_log  = sec_hdr('#ad1457')
    header_eco  = sec_hdr(colors['success'])

    cell_c   = fmt(border=1, align='center', valign='vcenter')
    cell_l   = fmt(border=1, align='left',   valign='vcenter')
    cell_num = fmt(border=1, align='center', valign='vcenter', num_format='0.00')
    cell_n4  = fmt(border=1, align='center', valign='vcenter', num_format='0.0000')
    cell_pct = fmt(border=1, align='center', valign='vcenter', num_format='0.00%')
    rank_gold   = fmt(bg_color=colors['gold'],   bold=True, font_size=14,
                      border=1, align='center', valign='vcenter')
    rank_silver = fmt(bg_color=colors['silver'], bold=True, font_size=12,
                      border=1, align='center', valign='vcenter')
    rank_bronze = fmt(bg_color=colors['bronze'], bold=True, font_size=12,
                      font_color=colors['white'], border=1, align='center', valign='vcenter')
    inf_fmt     = fmt(border=1, align='center', valign='vcenter',
                      font_color='#9e9e9e', italic=True, bg_color='#f5f5f5')

    locations    = data.get('locations', [])
    results_all  = data.get('results', [])
    weights      = data.get('weights', [])
    pairwise     = data.get('pairwiseMatrix', [])
    constraints  = data.get('constraints', [])
    region_filter = data.get('regionFilter') or {}

    feasible_r   = [r for r in results_all if r.get('feasible')]
    infeasible_r = [r for r in results_all if not r.get('feasible')]

    criteria_keys   = ['vendorBase', 'manpowerAvailability', 'capex',
                       'govtNorms', 'logisticsCost', 'economiesOfScale']
    criteria_labels = ['Vendor Base', 'Manpower', 'CAPEX',
                       'Govt Norms', 'Logistics Cost', 'Economies']
    is_cost = {
        'vendorBase': False, 'manpowerAvailability': False, 'capex': True,
        'govtNorms': False,  'logisticsCost': True,         'economiesOfScale': False,
    }

    def _get_val(loc, key):
        if '.' in key:
            parts = key.split('.', 1)
            sub = loc.get(parts[0], {})
            return sub.get(parts[1], '') if isinstance(sub, dict) else ''
        return loc.get(key, '')

    # ── Sheet 1: Raw Data ────────────────────────────────────────────────────
    ws1 = wb.add_worksheet('1_Raw_Data_Constraints')
    ws1.freeze_panes(4, 3)
    ws1.set_zoom(85)
    ws1.merge_range('A1:H1', '📊 ASHOK LEYLAND - Plant Location Decision System', title_main)
    ws1.merge_range('A2:H2', 'Raw Data with Constraint Filters Applied', title_sub)
    ws1.set_row(0, 30)
    ws1.set_row(1, 20)

    SECTIONS = [
        ('📍 LOCATION', header_loc, [
            ('Location', 'name', False), ('Region', 'region', False),
            ('State', 'state', False),   ('Industrial Park', 'raw.industrialPark', False),
        ]),
        ('🏭 VENDOR BASE', header_vnd, [
            ('ACMA Cluster', 'raw.acmaCluster', False),
            ('ACMA Units', 'raw.acmaUnits', False),
            ('Tier-1\n(200km)', 'raw.tier1Vendors', True),
            ('Tier-2\n(200km)', 'raw.tier2Vendors', True),
            ('Steel\n(100km)', 'raw.steelSuppliers', True),
            ('Ecosystem', 'raw.vendorEcosystem', False),
            ('Key OEMs', 'raw.keyOEMs', False),
            ('★ Score', 'vendorBase', True),
        ]),
        ('👥 MANPOWER', header_man, [
            ('Engg\nColleges', 'raw.enggColleges', True),
            ('ITI\nInst.', 'raw.itiInstitutes', True),
            ('ITI Grads\n(000s)', 'raw.itiGraduates', True),
            ('Engg Grads\n(000s)', 'raw.enggGraduates', True),
            ('Skill Rating', 'raw.skilledLabourRating', False),
            ('Wage\nSkilled', 'raw.wageSkilled', True),
            ('Wage\nSemi', 'raw.wageSemiSkilled', True),
            ('Attrition\n%', 'raw.attritionRate', True),
            ('★ Score', 'manpowerAvailability', True),
        ]),
        ('💰 CAPEX', header_cap, [
            ('Land Cost\n₹Cr/Ac', 'raw.landCost', True),
            ('Land\nAcres', 'raw.availableLand', True),
            ('Const.\nIndex', 'raw.constructionIndex', True),
            ('Power\nCapex', 'raw.powerCapex', True),
            ('Water\nCapex', 'raw.waterCapex', True),
            ('★ Total\nCAPEX', 'raw.totalCapex', True),
        ]),
        ('🏛️ GOVT / NORMS', header_gov, [
            ('Policy', 'raw.industrialPolicy', False),
            ('Cap Sub\n%', 'raw.capitalSubsidy', True),
            ('SGST\nyrs', 'raw.sgstExemption', True),
            ('Stamp\nDuty', 'raw.stampDuty', False),
            ('Power\nTariff', 'raw.powerTariff', True),
            ('Elec Duty\nExempt', 'raw.elecDutyExemption', False),
            ('Env\n1-10', 'raw.envClearanceEase', True),
            ('Approval\nDays', 'raw.approvalDays', True),
            ('SEZ\nNIMZ', 'raw.sezNimz', False),
            ('DFC\nGovt', 'raw.dfcAccessGovt', False),
            ('★ Score', 'govtNorms', True),
        ]),
        ('🚛 LOGISTICS', header_log, [
            ('Nearest Port', 'raw.nearestPort', False),
            ('Dist Port\nkm', 'raw.distanceToPort', True),
            ('Road\n1-10', 'raw.roadConnectivity', True),
            ('Rail\n1-10', 'raw.railConnectivity', True),
            ('DFC', 'raw.dfcLogistics', False),
            ('Dist Mkt\nkm', 'raw.distanceKeyMarket', True),
            ('Mkt City', 'raw.keyMarketCity', False),
            ('Inb Frt\n₹/MT', 'raw.inboundFreight', True),
            ('Out Frt\n₹/MT', 'raw.outboundFreight', True),
            ('★ Ann Cost\n₹Cr', 'raw.annualLogisticsCost', True),
        ]),
        ('📈 ECONOMIES', header_eco, [
            ('Maturity', 'raw.clusterMaturity', False),
            ('CV OEMs', 'raw.existingCVOEMs', False),
            ('Supplier\nPark', 'raw.supplierPark', False),
            ('Export\nHub', 'raw.exportHub', False),
            ('Demand\n1-10', 'raw.marketDemandIndex', True),
            ('Cluster\nScore', 'raw.clusterBenefitScore', True),
            ('★ Score', 'economiesOfScale', True),
        ]),
    ]

    col_details = []
    sec_spans   = []
    col_idx     = 0
    for sec_name, sec_f, cols in SECTIONS:
        start = col_idx
        for label, key, is_num in cols:
            col_details.append((col_idx, label, key, sec_f, is_num))
            col_idx += 1
        sec_spans.append((start, col_idx - 1, sec_name, sec_f))

    for cs, ce, sn, sf in sec_spans:
        if cs == ce:
            ws1.write(2, cs, sn, sf)
        else:
            ws1.merge_range(2, cs, 2, ce, sn, sf)

    for ci, lbl, key, sf, is_num in col_details:
        ws1.write(3, ci, lbl, sf)
        ws1.set_column(ci, ci, max(12, len(lbl.replace('\n', '')) + 2))

    for ri, loc in enumerate(locations):
        for ci, lbl, key, sf, is_num in col_details:
            val = _get_val(loc, key)
            if is_num:
                try:
                    ws1.write_number(ri + 4, ci, float(val), cell_num)
                except (ValueError, TypeError):
                    ws1.write(ri + 4, ci, val, cell_l)
            else:
                ws1.write(ri + 4, ci, str(val) if val is not None else '', cell_l)

    ws1.set_row(2, 25)
    ws1.set_row(3, 40)

    # Constraint summary
    cr_start = len(locations) + 6
    ws1.write(cr_start, 0, '🔒 CONSTRAINTS', title_sub)
    for j, lbl in enumerate(['Criterion', 'Operator', 'Threshold', 'Type', 'Status']):
        ws1.write(cr_start + 1, j, lbl, header_loc)
    for i, c in enumerate(constraints):
        r = cr_start + 2 + i
        ws1.write(r, 0, c.get('label', ''), cell_l)
        ws1.write(r, 1, c.get('operator', ''), cell_c)
        ws1.write(r, 2, c.get('value', 0), cell_num)
        ws1.write(r, 3, 'Cost ↓' if c.get('isCost') else 'Benefit ↑', cell_c)
        sf = fmt(bg_color=colors['success_light'], font_color=colors['success'],
                 bold=True, border=1, align='center') if c.get('enabled') else \
             fmt(bg_color='#eeeeee', font_color='#9e9e9e', border=1, align='center')
        ws1.write(r, 4, '✓ ACTIVE' if c.get('enabled') else '○ INACTIVE', sf)

    fr_row = cr_start + len(constraints) + 3
    ws1.write(fr_row,     0, '🌍 REGION FILTER', title_sub)
    ws1.write(fr_row + 1, 0, 'Filter Active:', cell_l)
    ws1.write(fr_row + 1, 1,
              'YES' if region_filter.get('regionFilterEnabled') else 'NO',
              fmt(bg_color=colors['success_light'], bold=True, border=1, align='center')
              if region_filter.get('regionFilterEnabled') else
              fmt(bg_color='#eeeeee', font_color='#9e9e9e', border=1, align='center'))
    ws1.write(fr_row + 2, 0, 'Selected Regions:', cell_l)
    selected = region_filter.get('selectedRegions', [])
    ws1.write(fr_row + 2, 1, ', '.join(selected) if selected else 'All Regions', cell_l)

    sm_row = fr_row + 4
    ws1.write(sm_row,     0, '📋 FEASIBILITY SUMMARY', title_sub)
    ws1.write(sm_row + 1, 0, 'Total Locations:', cell_l)
    ws1.write(sm_row + 1, 1, len(locations), cell_num)
    ws1.write(sm_row + 2, 0, 'Feasible:', cell_l)
    ws1.write(sm_row + 2, 1, len(feasible_r),
              fmt(bg_color=colors['success_light'], bold=True, border=1, num_format='0'))
    ws1.write(sm_row + 3, 0, 'Infeasible:', cell_l)
    ws1.write(sm_row + 3, 1, len(infeasible_r),
              fmt(bg_color=colors['danger_light'], bold=True, border=1, num_format='0'))

    # ── Sheet 2: Normalised Matrix ───────────────────────────────────────────
    ws2 = wb.add_worksheet('2_Normalized_Matrix')
    ws2.set_zoom(90)
    ws2.merge_range('A1:H1', '🌡️ NORMALIZED DECISION MATRIX (0 = Worst, 1 = Best)', title_main)
    ws2.write(1, 0, 'Heat Map — green = better, red = worse', title_sub)
    ws2.write(3, 0, 'Location', header_loc)
    ws2.set_column(0, 0, 26)
    for j, label in enumerate(criteria_labels):
        ws2.write(3, j + 1, label, header_vnd if j % 2 == 0 else header_man)
        ws2.set_column(j + 1, j + 1, 16)
    mm = {}
    for k in criteria_keys:
        vals = [float(loc.get(k, 0)) for loc in locations]
        mm[k] = (min(vals), max(vals))
    for ri, loc in enumerate(locations):
        ws2.write(ri + 4, 0, loc.get('name', ''), cell_l)
        for j, k in enumerate(criteria_keys):
            lo, hi = mm[k]
            rv = float(loc.get(k, 0))
            norm = ((rv - lo) / (hi - lo)) if hi != lo else 0.5
            norm = (1 - norm) if is_cost[k] else norm
            ws2.write_number(ri + 4, j + 1, round(norm, 4), cell_n4)
    last_dr = len(locations) + 3
    for j in range(6):
        col_l = chr(ord('B') + j)
        ws2.conditional_format(f'{col_l}5:{col_l}{last_dr + 1}', {
            'type': '3_color_scale',
            'min_color': '#f8696b', 'mid_color': '#ffeb84', 'max_color': '#63be7b'
        })

    # ── Sheet 3: Weights ─────────────────────────────────────────────────────
    ws3 = wb.add_worksheet('3_Weight_Calculation')
    ws3.set_zoom(90)
    ws3.merge_range('A1:H1', '⚖️ WEIGHT CALCULATION - AHP + Entropy + Hybrid', title_main)
    ws3.write(3, 0, '📊 AHP Pairwise Matrix', title_sub)
    ws3.write(4, 0, '', header_loc)
    for j, label in enumerate(criteria_labels):
        ws3.write(4, j + 1, label, header_vnd if j % 2 == 0 else header_man)
    for i, mrow in enumerate(pairwise):
        ws3.write(i + 5, 0, criteria_labels[i] if i < len(criteria_labels) else f'C{i+1}', cell_l)
        for j, v in enumerate(mrow):
            ws3.write_number(i + 5, j + 1, v, cell_num)
    sw = len(pairwise) + 7
    ws3.write(sw, 0, '🎯 WEIGHT SUMMARY', title_sub)
    ws3.write(sw + 1, 0, 'Method', header_loc)
    for j, label in enumerate(criteria_labels):
        ws3.write(sw + 1, j + 1, label, header_vnd if j % 2 == 0 else header_man)
    ahp_w  = [w.get('ahpWeight', 0)      for w in weights]
    ent_w  = [w.get('entropyWeight', 0)  for w in weights]
    comb_w = [w.get('combinedWeight', 0) for w in weights]
    for mi, (method, vals, bg) in enumerate([
        ('AHP Weight (Subjective)',        ahp_w,  '#e3f2fd'),
        ('Entropy Weight (Objective)',     ent_w,  '#e8f5e9'),
        ('Hybrid (60% AHP + 40% Entropy)', comb_w, '#fff3e0'),
    ]):
        r = sw + 2 + mi
        ws3.write(r, 0, method,
                  fmt(bg_color=bg, border=1, bold=True, align='left'))
        for j, v in enumerate(vals):
            ws3.write_number(r, j + 1, v, cell_pct)
    ws3.write(sw + 5, 0, 'Cost / Benefit', header_loc)
    for j, k in enumerate(criteria_keys):
        ind = 'COST ↓' if is_cost[k] else 'BENEFIT ↑'
        ws3.write(sw + 5, j + 1, ind,
                  fmt(bg_color=colors['danger_light']  if is_cost[k] else colors['success_light'],
                      font_color=colors['danger']      if is_cost[k] else colors['success'],
                      bold=True, border=1, align='center'))
    cw = sw + 8
    ws3.write(cw,     0, '📊 VISUAL WEIGHT BARS', title_sub)
    ws3.write(cw + 1, 0, 'Criterion', header_loc)
    ws3.write(cw + 1, 1, 'Weight',    header_loc)
    ws3.write(cw + 1, 2, 'Bar',       header_loc)
    ws3.set_column(2, 2, 40)
    for j, (label, w) in enumerate(zip(criteria_labels, comb_w)):
        ws3.write(cw + 2 + j, 0, label, cell_l)
        ws3.write_number(cw + 2 + j, 1, w, cell_pct)
        ws3.write_number(cw + 2 + j, 2, w, cell_pct)
    ws3.conditional_format(f'C{cw+3}:C{cw+8}', {
        'type': 'data_bar', 'bar_color': colors['accent'], 'bar_solid': True
    })

    # ── Sheet 4: TOPSIS Ranking ──────────────────────────────────────────────
    ws4 = wb.add_worksheet('4_TOPSIS_Ranking')
    ws4.set_zoom(90)
    ws4.merge_range('A1:I1', '🏆 TOPSIS RANKING', title_main)
    ws4.write(1, 0, 'Ranked by Composite Score (Higher is Better)', title_sub)
    hdrs  = ['Rank', 'Location', 'TOPSIS Score'] + criteria_labels + ['Status']
    hfmts = [header_loc, header_vnd, header_man] + \
            [header_vnd if j % 2 == 0 else header_man for j in range(6)] + [header_eco]
    widths = [8, 28, 14] + [13] * 6 + [12]
    for j, (h, hf, w) in enumerate(zip(hdrs, hfmts, widths)):
        ws4.write(3, j, h, hf)
        ws4.set_column(j, j, w)

    for ri, r in enumerate(feasible_r):
        cs   = r.get('criteriaScores', {})
        rank = r.get('rank', 0)
        rf   = rank_gold if rank == 1 else rank_silver if rank == 2 else \
               rank_bronze if rank == 3 else cell_c
        nf   = rank_gold if rank == 1 else rank_silver if rank == 2 else \
               rank_bronze if rank == 3 else cell_n4
        row  = ri + 4
        ws4.write(row, 0, rank, rf)
        ws4.write(row, 1, r.get('locationName', ''), cell_l)
        ws4.write_number(row, 2, round(r.get('compositeScore', 0), 4), nf)
        for j, k in enumerate(criteria_keys):
            ws4.write_number(row, j + 3, round(cs.get(k, 0), 4), cell_n4)
        ws4.write(row, 9, '✓ FEASIBLE',
                  fmt(bg_color=colors['success_light'], font_color=colors['success'],
                      bold=True, border=1, align='center'))

    for ri, r in enumerate(infeasible_r):
        row = len(feasible_r) + ri + 4
        for j in range(10):
            ws4.write(row, j, '—' if j != 1 else r.get('locationName', ''), inf_fmt)
        ws4.write(row, 2, 'N/A', inf_fmt)
        ws4.write(row, 9, '✗ INFEASIBLE', inf_fmt)

    last_f = len(feasible_r) + 3
    ws4.conditional_format(f'C5:C{last_f}', {
        'type': '3_color_scale',
        'min_color': '#f8696b', 'mid_color': '#ffeb84', 'max_color': '#63be7b'
    })

    # ── Sheet 5: Dashboard ───────────────────────────────────────────────────
    ws5 = wb.add_worksheet('5_Dashboard')
    ws5.set_zoom(85)
    ws5.merge_range('A1:L1', '📊 ASHOK LEYLAND - Executive Decision Dashboard',
                    fmt(bold=True, font_size=18, font_color=colors['white'],
                        bg_color=colors['primary_dark'], align='center', valign='vcenter'))
    ws5.set_row(0, 35)
    ws5.merge_range('A3:D3', '📋 EXECUTIVE SUMMARY', title_sub)

    if feasible_r:
        best = feasible_r[0]
        ws5.merge_range('A4:D4', f'🏆 RECOMMENDED: {best.get("locationName", "")}',
                        fmt(bold=True, font_size=14, bg_color=colors['gold'],
                            border=2, border_color=colors['warning'],
                            align='center', valign='vcenter'))
        ws5.set_row(3, 30)
        for i, (metric, value, bg) in enumerate([
            ('TOPSIS Score',       round(best.get('compositeScore', 0), 4), colors['success_light']),
            ('Total Locations',    len(locations),                          colors['neutral_light']),
            ('Feasible',           len(feasible_r),                         colors['success_light']),
            ('Infeasible',         len(infeasible_r),                       colors['danger_light']),
        ]):
            ws5.write(5 + i, 0, metric, fmt(bold=True, align='left'))
            ws5.write(5 + i, 1, value,
                      fmt(bg_color=bg, bold=True, border=1, align='center',
                          num_format='0.0000' if isinstance(value, float) else '0'))

    rs = 10
    ws5.merge_range(f'A{rs}:E{rs}', '🏆 LOCATION RANKINGS', title_sub)
    for j, h in enumerate(['Rank', 'Location', 'TOPSIS Score', 'Region', 'State']):
        ws5.write(rs + 1, j, h, header_loc)
    for ri, r in enumerate(feasible_r[:10]):
        lm   = next((l for l in locations if l.get('name') == r.get('locationName')), {})
        rank = r.get('rank', 0)
        rf   = rank_gold if rank == 1 else rank_silver if rank == 2 else \
               rank_bronze if rank == 3 else cell_c
        row  = rs + 2 + ri
        ws5.write(row, 0, rank, rf)
        ws5.write(row, 1, r.get('locationName', ''), cell_l)
        ws5.write_number(row, 2, round(r.get('compositeScore', 0), 4), cell_n4)
        ws5.write(row, 3, lm.get('region', ''), cell_l)
        ws5.write(row, 4, lm.get('state', ''), cell_l)

    # Chart 1 — TOPSIS bar chart
    chart1 = wb.add_chart({'type': 'column'})
    chart1.add_series({
        'name':       'TOPSIS Score',
        'categories': ['4_TOPSIS_Ranking', 4, 1, min(len(feasible_r) + 3, 13), 1],
        'values':     ['4_TOPSIS_Ranking', 4, 2, min(len(feasible_r) + 3, 13), 2],
        'fill':       {'color': colors['primary']},
        'data_labels': {'value': True, 'num_format': '0.0000'},
    })
    chart1.set_title({'name': 'Top 10 Locations by TOPSIS Score'})
    chart1.set_x_axis({'name': 'Location', 'num_font': {'rotation': -45}})
    chart1.set_y_axis({'name': 'TOPSIS Score', 'min': 0, 'max': 1, 'num_format': '0.00'})
    chart1.set_legend({'none': True})
    chart1.set_size({'width': 600, 'height': 350})
    ws5.insert_chart('G3', chart1)

    # Chart 2 — Weight pie
    chart2 = wb.add_chart({'type': 'pie'})
    chart2.add_series({
        'name':       'Criteria Weights',
        'categories': ['3_Weight_Calculation', len(pairwise) + 8, 1, len(pairwise) + 13, 1],
        'values':     ['3_Weight_Calculation', len(pairwise) + 8, 2, len(pairwise) + 13, 2],
        'data_labels': {'percentage': True, 'category': True},
    })
    chart2.set_title({'name': 'Hybrid Criteria Weight Distribution'})
    chart2.set_size({'width': 450, 'height': 350})
    ws5.insert_chart('G20', chart2)

    # Chart 3 — Top-3 criteria line
    if feasible_r:
        chart3 = wb.add_chart({'type': 'line'})
        for i in range(min(3, len(feasible_r))):
            r = feasible_r[i]
            chart3.add_series({
                'name':       r.get('locationName', ''),
                'categories': ['4_TOPSIS_Ranking', 3, 3, 3, 8],
                'values':     ['4_TOPSIS_Ranking', 4 + i, 3, 4 + i, 8],
                'line':       {'color': [colors['gold'], colors['silver'], colors['bronze']][i],
                               'width': 2.5},
                'marker':     {'type': 'circle', 'size': 8},
            })
        chart3.set_title({'name': 'Top 3 Locations — Criteria Comparison'})
        chart3.set_x_axis({'name': 'Criteria'})
        chart3.set_y_axis({'name': 'Score', 'min': 0, 'max': 1})
        chart3.set_legend({'position': 'bottom'})
        chart3.set_size({'width': 600, 'height': 350})
        ws5.insert_chart('A25', chart3)

    wb.close()
    output_buffer.seek(0)
    return output_buffer

# ═════════════════════════════════════════════════════════════════════════════
# API ENDPOINTS
# ═════════════════════════════════════════════════════════════════════════════
@app.get("/")
def root():
    return {"status": "ok"}
@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload and process Excel/CSV file — auto-detects header row and format."""
    try:
        contents = await file.read()
        buffer   = BytesIO(contents)

        if file.filename.endswith('.csv'):
            df = pd.read_csv(buffer)
        else:
            # --- sniff the right header row ---
            # Try header=0 first; if it looks like a simplified format, use it.
            # Otherwise fall back to header=1 for the detailed finaleyy.xlsx layout.
            df_h0 = pd.read_excel(BytesIO(contents), header=0)
            df_h0.columns = [str(c).replace('\n', ' ').strip() for c in df_h0.columns]

            if is_simplified_format(df_h0):
                df = df_h0
            else:
                # Detailed format has a merged title row at row 0; real headers at row 1
                df = pd.read_excel(BytesIO(contents), header=1)

        locations = process_excel_data(df)
        return {"count": len(locations), "locations": locations, "filename": file.filename}

    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error processing file: {str(e)}")


@app.post("/api/analyze")
async def analyze_locations(request: AnalysisRequest):
    """Run AHP-Entropy-TOPSIS analysis."""
    try:
        locations   = request.locations
        matrix      = request.pairwiseMatrix
        constraints = [c.dict() for c in request.constraints]

        criteria_keys = ['vendorBase', 'manpowerAvailability', 'capex',
                         'govtNorms', 'logisticsCost', 'economiesOfScale']
        is_cost = {
            'vendorBase': False, 'manpowerAvailability': False, 'capex': True,
            'govtNorms': False,  'logisticsCost': True,         'economiesOfScale': False,
        }

        feasible_locs, infeasible_locs = apply_constraints(locations, constraints)
        if not feasible_locs:
            return {"weights": [], "results": [],
                    "error": "No feasible locations match the constraints"}

        ahp_weights, cr    = calculate_ahp_weights(matrix)
        X                  = np.array([[loc.get(k, 0) for k in criteria_keys]
                                        for loc in feasible_locs])
        entropy_weights    = calculate_entropy_weights(X)
        hybrid_weights     = [0.6 * a + 0.4 * e
                               for a, e in zip(ahp_weights, entropy_weights)]
        total              = sum(hybrid_weights)
        hybrid_weights     = [w / total for w in hybrid_weights]
        scores             = topsis_analysis(feasible_locs, hybrid_weights,
                                             criteria_keys, is_cost)

        weight_response = []
        name_map = {
            'vendorBase': 'Vendor Base', 'manpowerAvailability': 'Manpower',
            'capex': 'CAPEX',            'govtNorms': 'Govt Norms',
            'logisticsCost': 'Logistics', 'economiesOfScale': 'Economies',
        }
        for i, key in enumerate(criteria_keys):
            weight_response.append({
                'key':            key,
                'name':           name_map[key],
                'isCost':         is_cost[key],
                'ahpWeight':      round(ahp_weights[i], 4),
                'entropyWeight':  round(entropy_weights[i], 4),
                'combinedWeight': round(hybrid_weights[i], 4),
            })

        results = []
        for loc, score in zip(feasible_locs, scores):
            criteria_scores = {}
            for key in criteria_keys:
                val     = loc.get(key, 0)
                max_val = max(l.get(key, 0) for l in feasible_locs)
                min_val = min(l.get(key, 0) for l in feasible_locs)
                if max_val > min_val:
                    criteria_scores[key] = round(
                        (max_val - val) / (max_val - min_val) if is_cost[key]
                        else (val - min_val) / (max_val - min_val), 3)
                else:
                    criteria_scores[key] = 0.5
            results.append({
                'locationId':    loc.get('name', ''),
                'locationName':  loc.get('name', ''),
                'compositeScore': round(score, 4),
                'criteriaScores': criteria_scores,
                'feasible':      True,
                'rank':          0,
            })

        results.sort(key=lambda x: x['compositeScore'], reverse=True)
        for i, r in enumerate(results):
            r['rank'] = i + 1

        for loc in infeasible_locs:
            results.append({
                'locationId':    loc.get('name', ''),
                'locationName':  loc.get('name', ''),
                'compositeScore': 0,
                'criteriaScores': {},
                'feasible':      False,
                'rank':          999,
            })

        return {"weights": weight_response, "results": results,
                "consistencyRatio": round(cr, 4)}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Analysis error: {str(e)}")


@app.post("/api/monte-carlo")
async def monte_carlo(request: MonteCarloRequest):
    """Run Monte Carlo simulation for robustness analysis."""
    try:
        locations     = request.locations
        weights_data  = request.weights
        iterations    = request.iterations
        constraints   = [c.dict() for c in request.constraints]
        region_filter = request.regionFilter

        feasible_locs, _ = apply_constraints(locations, constraints, region_filter)
        if not feasible_locs:
            return {"monteCarloResults": []}

        criteria_keys = ['vendorBase', 'manpowerAvailability', 'capex',
                         'govtNorms', 'logisticsCost', 'economiesOfScale']
        is_cost = {
            'vendorBase': False, 'manpowerAvailability': False, 'capex': True,
            'govtNorms': False,  'logisticsCost': True,         'economiesOfScale': False,
        }
        base_weights = [w.get('combinedWeight', 0) for w in weights_data]
        mc_results   = monte_carlo_simulation(feasible_locs, base_weights,
                                              criteria_keys, is_cost, iterations)
        return {"monteCarloResults": mc_results}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Simulation error: {str(e)}")


@app.post("/api/export-excel")
async def export_excel(request: ExportRequest):
    """Export comprehensive Excel report."""
    try:
        data = {
            'locations':     request.locations,
            'results':       request.results,
            'weights':       request.weights,
            'pairwiseMatrix': request.pairwiseMatrix,
            'constraints':   [c.dict() for c in request.constraints],
            'regionFilter':  request.regionFilter,
        }
        output = BytesIO()
        create_excel_report(data, output)
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Ashok_Leyland_Analysis.xlsx"},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export error: {str(e)}")


@app.get("/api/health")
async def health_check():
    return {"status": "healthy", "service": "Ashok Leyland Plant Location Decision API"}


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)