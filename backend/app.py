from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Dict, Any, Optional
import pandas as pd
import numpy as np
from io import BytesIO, StringIO
import json
import re
from scipy.stats import rankdata
import xlsxwriter
import uvicorn

app = FastAPI(title="Ashok Leyland Plant Location Decision API", version="2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://plantoptimizationsystem.vercel.app",
        "http://localhost:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Safe import — works both locally and on Vercel
try:
    from mangum import Mangum
    handler = Mangum(app)
except ImportError:
    handler = None  # Local dev — Mangum not needed
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
# EXCEL DATA PROCESSING - Maps finaleyy.xlsx to 6 criteria
# ═════════════════════════════════════════════════════════════════════════════

def extract_numeric(val):
    """Extract numeric value from string with ₹, commas, + suffixes"""
    if pd.isna(val) or val is None:
        return 0.0
    val_str = str(val).strip()
    # Handle 800+ format
    if '+' in val_str:
        val_str = val_str.replace('+', '')
    # Extract all numbers
    numbers = re.findall(r'\d+\.?\d*', val_str.replace(',', ''))
    if numbers:
        return float(numbers[0])
    return 0.0

def rating_to_score(rating):
    """Convert text ratings to numeric scores (1-10 scale)"""
    if pd.isna(rating):
        return 5.0
    rating_map = {
        'Very High': 10, 'High': 8, 'Medium': 6, 'Low': 4, 'Very Low': 2,
        'Very Mature': 10, 'Mature': 8, 'Moderate': 6, 'Emerging': 5,
        'Nascent': 3, 'Nascent-Growing': 4,
        'Yes': 10, 'Partial': 5, 'No': 0,
        'High': 8, 'Medium': 6, 'Low': 4  # For Export Hub Proximity
    }
    return float(rating_map.get(str(rating).strip(), 5))

def compute_vendor_base_score(row):
    """Compute composite vendor base score from Excel columns"""
    acma = str(row.get('No. of ACMA Member Units  (State, approx.)', '0'))
    acma_num = 800 if '800+' in acma else extract_numeric(acma)

    tier1 = extract_numeric(row.get('Tier-1 Auto Vendors  within 200 km (nos.)', 0))
    tier2 = extract_numeric(row.get('Tier-2 Auto Vendors  within 200 km (nos.)', 0))
    steel = extract_numeric(row.get('Steel / Castings Suppliers  within 100 km (nos.)', 0))
    ecosystem = rating_to_score(row.get('Vendor Ecosystem  Rating', 'Medium'))

    # Weighted composite (normalized to 0-10 scale)
    score = (
        (min(acma_num / 800, 1) * 0.25) +
        (min(tier1 / 95, 1) * 0.30) +
        (min(tier2 / 145, 1) * 0.25) +
        (min(steel / 62, 1) * 0.10) +
        (ecosystem / 10 * 0.10)
    ) * 10
    return round(score, 2)

def compute_manpower_score(row):
    """Compute composite manpower availability score"""
    engg_colleges = extract_numeric(row.get('AICTE Engg Colleges  in 50 km radius (nos.)', 0))
    iti = extract_numeric(row.get('ITI Institutes  in 50 km radius (nos.)', 0))
    iti_grad = extract_numeric(row.get('Annual ITI Graduates  (State, 000s)', 0))
    engg_grad = extract_numeric(row.get('Annual Engg Graduates  (State, 000s)', 0))
    availability = rating_to_score(row.get('Skilled Labour  Availability Rating', 'Medium'))

    score = (
        (min(engg_colleges / 62, 1) * 0.20) +
        (min(iti / 55, 1) * 0.15) +
        (min(iti_grad / 95, 1) * 0.20) +
        (min(engg_grad / 146, 1) * 0.25) +
        (availability / 10 * 0.20)
    ) * 10
    return round(score, 2)

def compute_govt_score(row):
    """Compute composite government norms/score"""
    subsidy = extract_numeric(row.get('Capital Subsidy  (% of Fixed Assets)', 0))
    sgst = extract_numeric(row.get('SGST Exemption /  Refund Period (yrs)', 0))
    env_ease = extract_numeric(row.get('Env. Clearance  Ease (1-10)', 5))

    # Approval days - lower is better, so invert (assuming 60 days max)
    approval_days = extract_numeric(row.get('Single Window  Approval Days (est.)', 50))
    approval_score = max(0, 10 - (approval_days / 6))  # 60 days -> 0, 30 days -> 5, 0 days -> 10

    score = (
        (min(subsidy / 35, 1) * 0.30) +
        (min(sgst / 7, 1) * 0.20) +
        (env_ease / 10 * 0.25) +
        (approval_score / 10 * 0.25)
    ) * 10
    return round(score, 2)

def compute_logistics_score(row):
    """Extract logistics cost (lower is better for cost criteria)"""
    return extract_numeric(row.get('Annual Logistics Cost  (₹ Cr/yr, est.)**', 0))

def compute_economies_score(row):
    """Compute composite economies of scale score"""
    maturity = rating_to_score(row.get('Auto Industry Cluster  Maturity', 'Medium'))
    supplier_park = rating_to_score(row.get('Supplier Park  Availability', 'No'))
    export_hub = rating_to_score(row.get('Export Hub  Proximity', 'Low'))
    market_demand = extract_numeric(row.get('Market Demand  Index (1-10)', 5))
    cluster_benefit = extract_numeric(row.get('Cluster Benefit  Score (1-10)', 5))

    score = (
        (maturity / 10 * 0.30) +
        (supplier_park / 10 * 0.20) +
        (export_hub / 10 * 0.15) +
        (market_demand / 10 * 0.20) +
        (cluster_benefit / 10 * 0.15)
    ) * 10
    return round(score, 2)

def process_excel_data(df):
    """Process finaleyy.xlsx format DataFrame into location objects"""
    # Clean column names
    df.columns = [str(col).replace('\n', ' ').strip() for col in df.columns]

    # Skip header rows if needed
    if 'S.No.' not in df.columns and 0 in df.columns:
        df = pd.read_excel(BytesIO(), header=1)  # Will be re-read properly in endpoint

    # Filter valid data rows
    df = df[df['Location'].notna()]
    df = df[df['Location'] != 'Location']  # Skip repeated headers

    locations = []
    for _, row in df.iterrows():
        try:
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
                'name': str(row['Location']),
                'region': gs('Region'),
                'state': gs('State'),
                'vendorBase': compute_vendor_base_score(row),
                'manpowerAvailability': compute_manpower_score(row),
                'capex': extract_numeric(row.get('Estimated Total  Project CAPEX (₹ Cr)*', 0)),
                'govtNorms': compute_govt_score(row),
                'logisticsCost': compute_logistics_score(row),
                'economiesOfScale': compute_economies_score(row),
                # ── FULL sub-attribute snapshot from all 7 Excel sections ──────────
                'raw': {
                    # Sec 1 – Location
                    'industrialPark':       gs('Industrial Park / Zone'),
                    # Sec 2 – Vendor Base
                    'acmaCluster':          gs('ACMA Auto Component Cluster'),
                    'acmaUnits':            gs('No. of ACMA Member Units  (State, approx.)'),
                    'tier1Vendors':         extract_numeric(row.get('Tier-1 Auto Vendors  within 200 km (nos.)', 0)),
                    'tier2Vendors':         extract_numeric(row.get('Tier-2 Auto Vendors  within 200 km (nos.)', 0)),
                    'steelSuppliers':       extract_numeric(row.get('Steel / Castings Suppliers  within 100 km (nos.)', 0)),
                    'vendorEcosystem':      gs('Vendor Ecosystem  Rating'),
                    'keyOEMs':              gs('Key OEMs / Anchors  in the Cluster'),
                    # Sec 3 – Manpower
                    'enggColleges':         extract_numeric(row.get('AICTE Engg Colleges  in 50 km radius (nos.)', 0)),
                    'itiInstitutes':        extract_numeric(row.get('ITI Institutes  in 50 km radius (nos.)', 0)),
                    'itiGraduates':         extract_numeric(row.get('Annual ITI Graduates  (State, 000s)', 0)),
                    'enggGraduates':        extract_numeric(row.get('Annual Engg Graduates  (State, 000s)', 0)),
                    'skilledLabourRating':  gs('Skilled Labour  Availability Rating'),
                    'wageSkilled':          extract_numeric(row.get('Avg Monthly Wage –  Skilled Mfg (₹)', 0)),
                    'wageSemiSkilled':      extract_numeric(row.get('Avg Monthly Wage –  Semi-Skilled (₹)', 0)),
                    'attritionRate':        extract_numeric(row.get('Labour Attrition Rate  (%/yr, est.)', 0)),
                    # Sec 4 – CAPEX
                    'landCost':             extract_numeric(row.get('Industrial Land Cost  (₹ Cr / Acre)', 0)),
                    'availableLand':        extract_numeric(row.get('Available Land  (Acres, approx.)', 0)),
                    'constructionIndex':    extract_numeric(row.get('Construction Cost  Index (Base TN=100)', 0)),
                    'powerCapex':           extract_numeric(row.get('Power Connection  Capex (₹ Cr, est.)', 0)),
                    'waterCapex':           extract_numeric(row.get('Water / Utilities  Capex (₹ Cr, est.)', 0)),
                    'totalCapex':           extract_numeric(row.get('Estimated Total  Project CAPEX (₹ Cr)*', 0)),
                    # Sec 5 – Govt / Norms
                    'industrialPolicy':     gs_find(['Industrial', 'Policy'], 'State Industrial Policy  (Current)'),
                    'capitalSubsidy':       extract_numeric(row.get('Capital Subsidy  (% of Fixed Assets)', 0)),
                    'sgstExemption':        extract_numeric(row.get('SGST Exemption /  Refund Period (yrs)', 0)),
                    'stampDuty':            gs('Stamp Duty  Exemption'),
                    'powerTariff':          extract_numeric(row.get('Power Tariff – HT  Industrial (₹/kWh)', 0)),
                    'elecDutyExemption':    gs('Electricity Duty  Exemption'),
                    'envClearanceEase':     extract_numeric(row.get('Env. Clearance  Ease (1-10)', 0)),
                    'approvalDays':         extract_numeric(row.get('Single Window  Approval Days (est.)', 0)),
                    'sezNimz':              gs('SEZ / NIMZ /  Special Zone'),
                    'dfcAccessGovt':        gs('Dedicated Freight  Corridor Access'),
                    # Sec 6 – Logistics
                    'nearestPort':          gs_find(['Nearest', 'Port'], 'Nearest Major Port'),
                    'distanceToPort':       extract_numeric(row.get('Distance to Port  (km)', 0)),
                    'roadConnectivity':     ext_num_find(['Road', 'Connectivity'], 'Road / NH  Connectivity (1-10)'),
                    'railConnectivity':     extract_numeric(row.get('Rail Connectivity  (1-10)', 0)),
                    'dfcLogistics':         gs('DFC Access  (Y/N)'),
                    'distanceKeyMarket':    extract_numeric(row.get('Distance to Nearest  Key Market (km)', 0)),
                    'keyMarketCity':        gs('Key Market City'),
                    'inboundFreight':       ext_num_find(['Inbound Freight'], 'Inbound Freight Rate  (₹/MT)'),
                    'outboundFreight':      ext_num_find(['Outbound Freight'], 'Outbound Freight Rate  (₹/MT)'),
                    'annualLogisticsCost':  extract_numeric(row.get('Annual Logistics Cost  (₹ Cr/yr, est.)**', 0)),
                    # Sec 7 – Economies of Scale
                    'clusterMaturity':      gs('Auto Industry Cluster  Maturity'),
                    'existingCVOEMs':       gs_find(['CV', 'OEMs'], 'Existing CV / Commercial  OEMs nearby'),
                    'supplierPark':         gs('Supplier Park  Availability'),
                    'exportHub':            gs('Export Hub  Proximity'),
                    'marketDemandIndex':    extract_numeric(row.get('Market Demand  Index (1-10)', 0)),
                    'clusterBenefitScore':  extract_numeric(row.get('Cluster Benefit  Score (1-10)', 0)),
                }
            }
            locations.append(loc)
        except Exception as e:
            print(f"Error processing row {row.get('Location', 'unknown')}: {e}")
            continue

    return locations

# ═════════════════════════════════════════════════════════════════════════════
# MCDM ALGORITHMS: AHP + Entropy + TOPSIS
# ═════════════════════════════════════════════════════════════════════════════

def calculate_ahp_weights(matrix):
    """Calculate AHP weights from pairwise comparison matrix"""
    matrix = np.array(matrix)
    n = matrix.shape[0]

    # Normalize matrix (column-wise)
    col_sums = np.sum(matrix, axis=0)
    normalized = matrix / col_sums

    # Row averages give priority weights
    weights = np.mean(normalized, axis=1)

    # Consistency check
    lambda_max = np.sum((matrix @ weights) / weights) / n
    ci = (lambda_max - n) / (n - 1)
    ri = [0, 0, 0.58, 0.90, 1.12, 1.24, 1.32, 1.41, 1.45, 1.49][min(n-1, 9)]
    cr = ci / ri if ri > 0 else 0

    return weights.tolist(), cr

def calculate_entropy_weights(data_matrix):
    """Calculate Entropy weights from decision matrix"""
    # data_matrix: rows = locations, cols = criteria
    X = np.array(data_matrix)
    n, m = X.shape

    if n <= 1:
        return [1.0 / m] * m

    # Normalize to probability matrix (column-wise)
    col_sums = np.sum(X, axis=0)
    col_sums = np.where(col_sums == 0, 1e-10, col_sums)
    P = X / col_sums

    # Handle zeros for log
    P = np.where(P == 0, 1e-10, P)

    # Calculate entropy for each criterion
    k = 1 / np.log(n)
    E = -k * np.sum(P * np.log(P), axis=0)

    # Calculate diversity and weights
    D = 1 - E
    
    d_sum = np.sum(D)
    if d_sum == 0:
        return [1.0 / m] * m
        
    weights = D / d_sum

    return weights.tolist()

def topsis_analysis(locations, hybrid_weights, criteria_keys, is_cost):
    """
    Perform TOPSIS analysis
    locations: list of dicts with criteria values
    hybrid_weights: combined AHP+Entropy weights
    criteria_keys: list of criteria key names
    is_cost: dict mapping key->bool (True if cost criteria)
    """
    # Build decision matrix
    n = len(locations)
    m = len(criteria_keys)
    X = np.zeros((n, m))

    for i, loc in enumerate(locations):
        for j, key in enumerate(criteria_keys):
            X[i, j] = float(loc.get(key, 0))

    # Step 1: Vector normalization
    col_norms = np.sqrt(np.sum(X**2, axis=0))
    col_norms = np.where(col_norms == 0, 1, col_norms)
    R = X / col_norms

    # Step 2: Weighted normalized matrix
    W = np.array(hybrid_weights)
    V = R * W

    # Step 3: Ideal best and worst
    A_plus = np.zeros(m)
    A_minus = np.zeros(m)

    for j, key in enumerate(criteria_keys):
        if is_cost.get(key, False):
            # For cost: lower is better, so min is best
            A_plus[j] = np.min(V[:, j])
            A_minus[j] = np.max(V[:, j])
        else:
            # For benefit: higher is better
            A_plus[j] = np.max(V[:, j])
            A_minus[j] = np.min(V[:, j])

    # Step 4: Separation measures
    S_plus = np.sqrt(np.sum((V - A_plus)**2, axis=1))
    S_minus = np.sqrt(np.sum((V - A_minus)**2, axis=1))

    # Step 5: Relative closeness to ideal
    scores = S_minus / (S_plus + S_minus + 1e-10)

    return scores.tolist()

def apply_constraints(locations, constraints, region_filter=None):
    """Filter locations based on constraints and region filter"""
    feasible = []
    infeasible = []

    for loc in locations:
        is_feasible = True
        reasons = []

        # Check numeric constraints
        for c in constraints:
            if not c.get('enabled', False):
                continue

            key = c['key']
            val = float(loc.get(key, 0))
            threshold = float(c['value'])
            op = c['operator']

            passes = True
            if op == 'gte':
                passes = val >= threshold
            elif op == 'lte':
                passes = val <= threshold
            elif op == 'eq':
                passes = abs(val - threshold) < 0.001

            if not passes:
                is_feasible = False
                reasons.append(f"{c['label']} {op} {threshold}")

        # Check region/state filter
        if region_filter and region_filter.get('regionFilterEnabled'):
            selected = region_filter.get('selectedRegions', [])
            if selected:
                region_match = loc.get('region') in selected
                state_match = loc.get('state') in selected
                if not (region_match or state_match):
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
    """
    Run Monte Carlo simulation with perturbed weights and values
    Returns rank stability metrics for each location
    """
    n = len(locations)
    m = len(criteria_keys)

    # Store ranks across iterations
    all_ranks = {loc['name']: [] for loc in locations}

    base_weights = np.array(base_weights)

    for _ in range(iterations):
        # Perturb weights (±15%)
        weight_noise = 1 + np.random.uniform(-weight_perturb, weight_perturb, m)
        perturbed_weights = base_weights * weight_noise
        perturbed_weights = perturbed_weights / np.sum(perturbed_weights)

        # Perturb location values (±5%)
        perturbed_locations = []
        for loc in locations:
            new_loc = loc.copy()
            for key in criteria_keys:
                val = float(loc.get(key, 0))
                noise = 1 + np.random.uniform(-value_perturb, value_perturb)
                new_loc[key] = val * noise
            perturbed_locations.append(new_loc)

        # Run TOPSIS with perturbed data
        scores = topsis_analysis(perturbed_locations, perturbed_weights, criteria_keys, is_cost)

        # Convert to ranks (higher score = better rank = lower rank number)
        ranks = rankdata([-s for s in scores], method='min')

        for i, loc in enumerate(locations):
            all_ranks[loc['name']].append(int(ranks[i]))

    # Calculate statistics
    results = []
    for loc in locations:
        ranks = all_ranks[loc['name']]
        avg_rank = np.mean(ranks)
        std_rank = np.std(ranks)

        # Rank probability distribution (positions 1 to n)
        rank_counts = np.bincount(ranks, minlength=n+1)[1:]  # Skip index 0
        rank_probs = rank_counts / iterations

        # Confidence interval (95%)
        ci_low = max(1, int(np.percentile(ranks, 2.5)))
        ci_high = min(n, int(np.percentile(ranks, 97.5)))

        results.append({
            'locationId': loc.get('name', ''),
            'locationName': loc.get('name', ''),
            'avgRank': round(avg_rank, 2),
            'stdRank': round(std_rank, 2),
            'confidenceInterval': [ci_low, ci_high],
            'rankProbabilities': rank_probs.tolist(),
            'bestRank': int(min(ranks)),
            'worstRank': int(max(ranks))
        })

    # Sort by average rank
    results.sort(key=lambda x: x['avgRank'])
    return results

# ═════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT GENERATION
# ═════════════════════════════════════════════════════════════════════════════

def create_excel_report(data: Dict, output_buffer: BytesIO):
    """
    Create comprehensive 5-sheet colorful Excel report with visual representations.
    
    Sheets:
    1. Raw Data & Constraints - Color-coded raw data with constraint filters
    2. Normalized Matrix - Heat map style visualization with color scales
    3. Weight Calculation - AHP + Entropy + Hybrid with visual bars
    4. TOPSIS Ranking - Color-ranked results with conditional formatting
    5. Dashboard - Multiple charts and visual summaries
    """
    wb = xlsxwriter.Workbook(output_buffer, {'in_memory': True})
    
    # COLOR PALETTE - Professional Corporate Theme
    colors = {
        'primary_dark': '#1a237e',      # Deep blue
        'primary': '#283593',          # Indigo
        'primary_light': '#5c6bc0',    # Light indigo
        'accent': '#00acc1',           # Cyan accent
        'accent_light': '#4dd0e1',     # Light cyan
        'success': '#2e7d32',          # Green
        'success_light': '#a5d6a7',      # Light green
        'warning': '#f57c00',          # Orange
        'warning_light': '#ffcc80',      # Light orange
        'danger': '#c62828',           # Red
        'danger_light': '#ef9a9a',       # Light red
        'neutral': '#455a64',          # Blue grey
        'neutral_light': '#cfd8dc',      # Light blue grey
        'gold': '#ffd700',             # Gold for #1 rank
        'silver': '#c0c0c0',           # Silver for #2
        'bronze': '#cd7f32',           # Bronze for #3
        'white': '#ffffff',
        'bg_light': '#f5f5f5',
        'bg_alt': '#fafafa'
    }
    
    # FORMAT DEFINITIONS
    title_main = wb.add_format({
        'bold': True, 
        'font_size': 16, 
        'font_color': colors['primary_dark'],
        'align': 'left',
        'valign': 'vcenter'
    })
    
    title_sub = wb.add_format({
        'bold': True, 
        'font_size': 12, 
        'font_color': colors['primary'],
        'align': 'left',
        'valign': 'vcenter'
    })
    
    # Section-specific colored headers
    header_location = wb.add_format({
        'bold': True, 'bg_color': colors['primary_dark'], 'font_color': colors['white'],
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    header_vendor = wb.add_format({
        'bold': True, 'bg_color': '#1565c0', 'font_color': colors['white'],
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    header_manpower = wb.add_format({
        'bold': True, 'bg_color': '#00695c', 'font_color': colors['white'],
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    header_capex = wb.add_format({
        'bold': True, 'bg_color': '#ef6c00', 'font_color': colors['white'],
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    header_govt = wb.add_format({
        'bold': True, 'bg_color': '#6a1b9a', 'font_color': colors['white'],
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    header_logistics = wb.add_format({
        'bold': True, 'bg_color': '#ad1457', 'font_color': colors['white'],
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    header_economies = wb.add_format({
        'bold': True, 'bg_color': colors['success'], 'font_color': colors['white'],
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    
    # Data cell formats
    cell_center = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    cell_left = wb.add_format({
        'border': 1, 'align': 'left', 'valign': 'vcenter'
    })
    cell_number = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'
    })
    cell_number_4dec = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.0000'
    })
    cell_percent = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00%'
    })
    
    # Special ranking formats with medals
    rank_gold = wb.add_format({
        'bg_color': colors['gold'], 'bold': True, 'font_size': 14,
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    rank_silver = wb.add_format({
        'bg_color': colors['silver'], 'bold': True, 'font_size': 12,
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    rank_bronze = wb.add_format({
        'bg_color': colors['bronze'], 'bold': True, 'font_size': 12,
        'font_color': colors['white'],
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    
    # Infeasible format
    infeasible_format = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_color': '#9e9e9e', 'italic': True, 'bg_color': '#f5f5f5'
    })
    
    # Constraint formats
    constraint_active = wb.add_format({
        'bg_color': colors['success_light'], 'font_color': colors['success'],
        'bold': True, 'border': 1, 'align': 'center'
    })
    constraint_inactive = wb.add_format({
        'bg_color': '#eeeeee', 'font_color': '#9e9e9e',
        'border': 1, 'align': 'center'
    })
    
    # EXTRACT DATA
    locations = data.get('locations', [])
    results_all = data.get('results', [])
    weights = data.get('weights', [])
    pairwise = data.get('pairwiseMatrix', [])
    constraints = data.get('constraints', [])
    region_filter = data.get('regionFilter') or {}
    
    feasible_r = [r for r in results_all if r.get('feasible')]
    infeasible_r = [r for r in results_all if not r.get('feasible')]
    
    criteria_keys = ['vendorBase', 'manpowerAvailability', 'capex', 
                     'govtNorms', 'logisticsCost', 'economiesOfScale']
    criteria_labels = ['Vendor Base', 'Manpower', 'CAPEX', 'Govt Norms', 
                       'Logistics Cost', 'Economies']
    is_cost = {
        'vendorBase': False, 'manpowerAvailability': False, 'capex': True,
        'govtNorms': False, 'logisticsCost': True, 'economiesOfScale': False
    }
    
    # ═════════════════════════════════════════════════════════════════════════════
    # SHEET 1: Raw Data & Constraints (Color-coded by section)
    # ═════════════════════════════════════════════════════════════════════════════
    ws1 = wb.add_worksheet('1_Raw_Data_Constraints')
    ws1.freeze_panes(3, 3)
    ws1.set_zoom(85)
    
    # Title
    ws1.merge_range('A1:H1', '📊 ASHOK LEYLAND - Plant Location Decision System', title_main)
    ws1.merge_range('A2:H2', 'Raw Data with Constraint Filters Applied', title_sub)
    ws1.set_row(0, 30)
    ws1.set_row(1, 20)
    
    # Section definitions with colors
    SECTIONS = [
        ('📍 LOCATION', header_location, [
            ('Location', 'name', False),
            ('Region', 'region', False),
            ('State', 'state', False),
            ('Industrial Park', 'raw.industrialPark', False),
        ]),
        ('🏭 VENDOR BASE', header_vendor, [
            ('ACMA Cluster', 'raw.acmaCluster', False),
            ('ACMA Units', 'raw.acmaUnits', False),
            ('Tier-1 Vendors\n(200km)', 'raw.tier1Vendors', True),
            ('Tier-2 Vendors\n(200km)', 'raw.tier2Vendors', True),
            ('Steel Suppliers\n(100km)', 'raw.steelSuppliers', True),
            ('Vendor Ecosystem', 'raw.vendorEcosystem', False),
            ('Key OEMs', 'raw.keyOEMs', False),
            ('★ Score (0-10)', 'vendorBase', True),
        ]),
        ('👥 MANPOWER', header_manpower, [
            ('Engg Colleges\n(50km)', 'raw.enggColleges', True),
            ('ITI Institutes\n(50km)', 'raw.itiInstitutes', True),
            ('ITI Grads\n(000s)', 'raw.itiGraduates', True),
            ('Engg Grads\n(000s)', 'raw.enggGraduates', True),
            ('Skill Rating', 'raw.skilledLabourRating', False),
            ('Wage Skilled\n(₹)', 'raw.wageSkilled', True),
            ('Wage Semi\n(₹)', 'raw.wageSemiSkilled', True),
            ('Attrition\n(%/yr)', 'raw.attritionRate', True),
            ('★ Score (0-10)', 'manpowerAvailability', True),
        ]),
        ('💰 CAPEX', header_capex, [
            ('Land Cost\n(₹Cr/Ac)', 'raw.landCost', True),
            ('Available Land\n(Ac)', 'raw.availableLand', True),
            ('Const. Index', 'raw.constructionIndex', True),
            ('Power Capex\n(₹Cr)', 'raw.powerCapex', True),
            ('Water Capex\n(₹Cr)', 'raw.waterCapex', True),
            ('★ Total CAPEX\n(₹Cr)', 'raw.totalCapex', True),
        ]),
        ('🏛️ GOVT / NORMS', header_govt, [
            ('Policy', 'raw.industrialPolicy', False),
            ('Capital Sub\n(%)', 'raw.capitalSubsidy', True),
            ('SGST Exempt\n(yrs)', 'raw.sgstExemption', True),
            ('Stamp Duty', 'raw.stampDuty', False),
            ('Power Tariff\n(₹/kWh)', 'raw.powerTariff', True),
            ('Elec Duty\nExempt', 'raw.elecDutyExemption', False),
            ('Env Clear\n(1-10)', 'raw.envClearanceEase', True),
            ('Approval\n(days)', 'raw.approvalDays', True),
            ('SEZ/NIMZ', 'raw.sezNimz', False),
            ('DFC Access\n(Govt)', 'raw.dfcAccessGovt', False),
            ('★ Score (0-10)', 'govtNorms', True),
        ]),
        ('🚛 LOGISTICS', header_logistics, [
            ('Nearest Port', 'raw.nearestPort', False),
            ('Dist Port\n(km)', 'raw.distanceToPort', True),
            ('Road Conn\n(1-10)', 'raw.roadConnectivity', True),
            ('Rail Conn\n(1-10)', 'raw.railConnectivity', True),
            ('DFC Access', 'raw.dfcLogistics', False),
            ('Dist Market\n(km)', 'raw.distanceKeyMarket', True),
            ('Market City', 'raw.keyMarketCity', False),
            ('Inb Freight\n(₹/MT)', 'raw.inboundFreight', True),
            ('Out Freight\n(₹/MT)', 'raw.outboundFreight', True),
            ('★ Annual Cost\n(₹Cr)', 'raw.annualLogisticsCost', True),
        ]),
        ('📈 ECONOMIES OF SCALE', header_economies, [
            ('Cluster Maturity', 'raw.clusterMaturity', False),
            ('CV OEMs Nearby', 'raw.existingCVOEMs', False),
            ('Supplier Park', 'raw.supplierPark', False),
            ('Export Hub', 'raw.exportHub', False),
            ('Demand Index\n(1-10)', 'raw.marketDemandIndex', True),
            ('Cluster Score\n(1-10)', 'raw.clusterBenefitScore', True),
            ('★ Score (0-10)', 'economiesOfScale', True),
        ]),
    ]
    
    def _get_val(loc, key):
        """Resolve dotted key like 'raw.tier1Vendors' or 'vendorBase'."""
        if '.' in key:
            parts = key.split('.', 1)
            sub = loc.get(parts[0], {})
            if isinstance(sub, dict):
                return sub.get(parts[1], '')
            return ''
        return loc.get(key, '')
    
    # Build column structure
    col_details = []
    sec_spans = []
    col_idx = 0
    
    for sec_name, sec_fmt, cols in SECTIONS:
        start = col_idx
        for label, key, is_num in cols:
            col_details.append((col_idx, label, key, sec_fmt, is_num))
            col_idx += 1
        sec_spans.append((start, col_idx - 1, sec_name, sec_fmt))
    
    # Row 2: Section merged headers with colors
    for (cs, ce, sn, fmt) in sec_spans:
        if cs == ce:
            ws1.write(2, cs, sn, fmt)
        else:
            ws1.merge_range(2, cs, 2, ce, sn, fmt)
    
    # Row 3: Column labels
    for (ci, lbl, key, fmt, is_num) in col_details:
        ws1.write(3, ci, lbl, fmt)
        content_len = len(lbl.replace('\n', ''))
        ws1.set_column(ci, ci, max(12, content_len + 2))
    
    # Row 4+: Data
    for ri, loc in enumerate(locations):
        data_row = ri + 4
        for (ci, lbl, key, fmt, is_num) in col_details:
            val = _get_val(loc, key)
            if is_num:
                try:
                    ws1.write_number(data_row, ci, float(val), cell_number)
                except (ValueError, TypeError):
                    ws1.write(data_row, ci, val, cell_left)
            else:
                ws1.write(data_row, ci, str(val) if val is not None else '', cell_left)
    
    ws1.set_row(2, 25)
    ws1.set_row(3, 40)
    
    # Constraint Summary
    constraint_start_row = len(locations) + 6
    ws1.write(constraint_start_row, 0, '🔒 CONSTRAINTS APPLIED', title_sub)
    
    ws1.write(constraint_start_row + 1, 0, 'Criterion', header_location)
    ws1.write(constraint_start_row + 1, 1, 'Operator', header_location)
    ws1.write(constraint_start_row + 1, 2, 'Threshold', header_location)
    ws1.write(constraint_start_row + 1, 3, 'Type', header_location)
    ws1.write(constraint_start_row + 1, 4, 'Status', header_location)
    
    for i, c in enumerate(constraints):
        row = constraint_start_row + 2 + i
        ws1.write(row, 0, c.get('label', ''), cell_left)
        ws1.write(row, 1, c.get('operator', ''), cell_center)
        ws1.write(row, 2, c.get('value', 0), cell_number)
        ws1.write(row, 3, 'Cost ↓' if c.get('isCost') else 'Benefit ↑', cell_center)
        status_fmt = constraint_active if c.get('enabled') else constraint_inactive
        ws1.write(row, 4, '✓ ACTIVE' if c.get('enabled') else '○ INACTIVE', status_fmt)
    
    # Region filter info
    filter_row = constraint_start_row + len(constraints) + 3
    ws1.write(filter_row, 0, '🌍 REGION FILTER', title_sub)
    ws1.write(filter_row + 1, 0, 'Filter Active:', cell_left)
    ws1.write(filter_row + 1, 1, 'YES' if region_filter.get('regionFilterEnabled') else 'NO', 
              constraint_active if region_filter.get('regionFilterEnabled') else constraint_inactive)
    ws1.write(filter_row + 2, 0, 'Selected Regions:', cell_left)
    selected = region_filter.get('selectedRegions', [])
    ws1.write(filter_row + 2, 1, ', '.join(selected) if selected else 'All Regions', cell_left)
    
    # Feasibility summary
    summary_row = filter_row + 4
    ws1.write(summary_row, 0, '📋 FEASIBILITY SUMMARY', title_sub)
    ws1.write(summary_row + 1, 0, 'Total Locations:', cell_left)
    ws1.write(summary_row + 1, 1, len(locations), cell_number)
    ws1.write(summary_row + 2, 0, 'Feasible:', cell_left)
    ws1.write(summary_row + 2, 1, len(feasible_r), wb.add_format({
        'bg_color': colors['success_light'], 'bold': True, 'border': 1, 'num_format': '0'
    }))
    ws1.write(summary_row + 3, 0, 'Infeasible:', cell_left)
    ws1.write(summary_row + 3, 1, len(infeasible_r), wb.add_format({
        'bg_color': colors['danger_light'], 'bold': True, 'border': 1, 'num_format': '0'
    }))
    
    # ═════════════════════════════════════════════════════════════════════════════
    # SHEET 2: Normalized Matrix with Heat Map
    # ═════════════════════════════════════════════════════════════════════════════
    ws2 = wb.add_worksheet('2_Normalized_Matrix')
    ws2.set_zoom(90)
    
    ws2.merge_range('A1:H1', '🌡️ NORMALIZED DECISION MATRIX (0 = Worst, 1 = Best)', title_main)
    ws2.write(1, 0, 'Heat Map Visualization - Higher values are better (green), lower are worse (red)', title_sub)
    
    # Headers
    ws2.write(3, 0, 'Location', header_location)
    ws2.set_column(0, 0, 26)
    for j, label in enumerate(criteria_labels):
        ws2.write(3, j + 1, label, header_vendor if j % 2 == 0 else header_manpower)
        ws2.set_column(j + 1, j + 1, 16)
    
    # Compute normalized values
    mm = {}
    for k in criteria_keys:
        vals = [float(loc.get(k, 0)) for loc in locations]
        lo, hi = min(vals), max(vals)
        mm[k] = (lo, hi)
    
    # Write data
    for ri, loc in enumerate(locations):
        ws2.write(ri + 4, 0, loc.get('name', ''), cell_left)
        for j, k in enumerate(criteria_keys):
            lo, hi = mm[k]
            raw_v = float(loc.get(k, 0))
            if hi != lo:
                norm = (raw_v - lo) / (hi - lo)
                norm = (1 - norm) if is_cost[k] else norm
            else:
                norm = 0.5
            ws2.write_number(ri + 4, j + 1, round(norm, 4), cell_number_4dec)
    
    # Apply 3-color scale conditional formatting
    last_data_row = len(locations) + 3
    for j in range(6):
        col_letter = chr(ord('B') + j)
        ws2.conditional_format(f'{col_letter}5:{col_letter}{last_data_row + 1}', {
            'type': '3_color_scale',
            'min_color': '#f8696b',  # Red
            'mid_color': '#ffeb84',  # Yellow
            'max_color': '#63be7b'   # Green
        })
    
    # Add data bars
    for j in range(6):
        col_letter = chr(ord('B') + j)
        ws2.conditional_format(f'{col_letter}5:{col_letter}{last_data_row + 1}', {
            'type': 'data_bar',
            'bar_color': colors['primary_light'],
            'bar_solid': False
        })
    
    # ═════════════════════════════════════════════════════════════════════════════
    # SHEET 3: Weight Calculation with Visual Bars
    # ═════════════════════════════════════════════════════════════════════════════
    ws3 = wb.add_worksheet('3_Weight_Calculation')
    ws3.set_zoom(90)
    
    ws3.merge_range('A1:H1', '⚖️ WEIGHT CALCULATION - AHP + Entropy + Hybrid', title_main)
    
    # AHP Pairwise Matrix
    ws3.write(3, 0, '📊 AHP Pairwise Comparison Matrix', title_sub)
    ws3.write(4, 0, '', header_location)
    for j, label in enumerate(criteria_labels):
        ws3.write(4, j + 1, label, header_vendor if j % 2 == 0 else header_manpower)
    
    for i, mrow in enumerate(pairwise):
        ws3.write(i + 5, 0, criteria_labels[i] if i < len(criteria_labels) else f'C{i+1}', cell_left)
        for j, v in enumerate(mrow):
            ws3.write_number(i + 5, j + 1, v, cell_number)
    
    # Weight summary
    start_w = len(pairwise) + 7
    ws3.write(start_w, 0, '🎯 WEIGHT SUMMARY', title_sub)
    
    ws3.write(start_w + 1, 0, 'Method', header_location)
    for j, label in enumerate(criteria_labels):
        ws3.write(start_w + 1, j + 1, label, header_vendor if j % 2 == 0 else header_manpower)
    
    ahp_w = [w.get('ahpWeight', 0) for w in weights]
    ent_w = [w.get('entropyWeight', 0) for w in weights]
    comb_w = [w.get('combinedWeight', 0) for w in weights]
    
    methods = [
        ('AHP Weight (Subjective)', ahp_w, '#e3f2fd'),
        ('Entropy Weight (Objective)', ent_w, '#e8f5e9'),
        ('Hybrid (60% AHP + 40% Entropy)', comb_w, '#fff3e0')
    ]
    
    for mi, (method_name, w_values, bg_color) in enumerate(methods):
        row = start_w + 2 + mi
        row_fmt = wb.add_format({'bg_color': bg_color, 'border': 1, 'bold': True, 'align': 'left'})
        ws3.write(row, 0, method_name, row_fmt)
        for j, val in enumerate(w_values):
            ws3.write_number(row, j + 1, val, cell_percent)
    
    # Cost/Benefit indicator
    ws3.write(start_w + 5, 0, 'Cost / Benefit', header_location)
    for j, k in enumerate(criteria_keys):
        indicator = 'COST ↓' if is_cost[k] else 'BENEFIT ↑'
        indicator_fmt = wb.add_format({
            'bg_color': colors['danger_light'] if is_cost[k] else colors['success_light'],
            'font_color': colors['danger'] if is_cost[k] else colors['success'],
            'bold': True, 'border': 1, 'align': 'center'
        })
        ws3.write(start_w + 5, j + 1, indicator, indicator_fmt)
    
    # Visual weight bars
    chart_row = start_w + 8
    ws3.write(chart_row, 0, '📊 VISUAL WEIGHT REPRESENTATION', title_sub)
    ws3.write(chart_row + 1, 0, 'Criterion', header_location)
    ws3.write(chart_row + 1, 1, 'Weight', header_location)
    ws3.write(chart_row + 1, 2, 'Visual Bar', header_location)
    ws3.set_column(2, 2, 40)
    
    for j, (label, weight) in enumerate(zip(criteria_labels, comb_w)):
        row = chart_row + 2 + j
        ws3.write(row, 0, label, cell_left)
        ws3.write_number(row, 1, weight, cell_percent)
        ws3.write(row, 2, weight, cell_percent)
    
    # Apply data bars
    ws3.conditional_format(f'C{chart_row + 3}:C{chart_row + 8}', {
        'type': 'data_bar',
        'bar_color': colors['accent'],
        'bar_solid': True,
        'min_value': 0,
        'max_value': 1
    })
    
    # ═════════════════════════════════════════════════════════════════════════════
    # SHEET 4: TOPSIS Ranking with Color Coding
    # ═════════════════════════════════════════════════════════════════════════════
    ws4 = wb.add_worksheet('4_TOPSIS_Ranking')
    ws4.set_zoom(90)
    
    ws4.merge_range('A1:I1', '🏆 TOPSIS RANKING - Ashok Leyland Plant Location Decision', title_main)
    ws4.write(1, 0, 'Ranked by Composite Score (Higher is Better)', title_sub)
    
    # Headers
    headers = ['Rank', 'Location', 'TOPSIS Score'] + criteria_labels + ['Status']
    header_formats = [header_location, header_vendor, header_manpower] + \
                     [header_vendor if j % 2 == 0 else header_manpower for j in range(6)] + \
                     [header_economies]
    
    for j, (h, hf) in enumerate(zip(headers, header_formats)):
        ws4.write(3, j, h, hf)
        if j == 0:
            ws4.set_column(j, j, 8)
        elif j == 1:
            ws4.set_column(j, j, 28)
        elif j == 2:
            ws4.set_column(j, j, 14)
        else:
            ws4.set_column(j, j, 13)
    
    # Feasible locations
    for ri, r in enumerate(feasible_r):
        cs = r.get('criteriaScores', {})
        row = ri + 4
        rank = r.get('rank', 0)
        
        if rank == 1:
            rank_fmt = rank_gold
            score_fmt = rank_gold
        elif rank == 2:
            rank_fmt = rank_silver
            score_fmt = rank_silver
        elif rank == 3:
            rank_fmt = rank_bronze
            score_fmt = rank_bronze
        else:
            rank_fmt = cell_center
            score_fmt = cell_number_4dec
        
        ws4.write(row, 0, rank, rank_fmt)
        ws4.write(row, 1, r.get('locationName', ''), cell_left)
        ws4.write_number(row, 2, round(r.get('compositeScore', 0), 4), score_fmt)
        
        for j, k in enumerate(criteria_keys):
            ws4.write_number(row, j + 3, round(cs.get(k, 0), 4), cell_number_4dec)
        
        status_fmt = wb.add_format({
            'bg_color': colors['success_light'],
            'font_color': colors['success'],
            'bold': True, 'border': 1, 'align': 'center'
        })
        ws4.write(row, 9, '✓ FEASIBLE', status_fmt)
    
    # Infeasible locations
    for ri, r in enumerate(infeasible_r):
        row = len(feasible_r) + ri + 4
        ws4.write(row, 0, '—', infeasible_format)
        ws4.write(row, 1, r.get('locationName', ''), infeasible_format)
        ws4.write(row, 2, 'N/A', infeasible_format)
        for j in range(6):
            ws4.write(row, j + 3, '—', infeasible_format)
        ws4.write(row, 9, '✗ INFEASIBLE', infeasible_format)
    
    # Conditional formatting
    last_feasible = len(feasible_r) + 3
    ws4.conditional_format(f'C5:C{last_feasible}', {
        'type': '3_color_scale',
        'min_color': '#f8696b',
        'mid_color': '#ffeb84',
        'max_color': '#63be7b'
    })
    
    for j in range(6):
        col = chr(ord('D') + j)
        ws4.conditional_format(f'{col}5:{col}{last_feasible}', {
            'type': 'data_bar',
            'bar_color': colors['primary_light']
        })
    
    # ═════════════════════════════════════════════════════════════════════════════
    # SHEET 5: Dashboard with Multiple Charts
    # ═════════════════════════════════════════════════════════════════════════════
    ws5 = wb.add_worksheet('5_Dashboard')
    ws5.set_zoom(85)
    
    # Main title with background
    title_dash = wb.add_format({
        'bold': True, 'font_size': 18, 'font_color': colors['white'],
        'bg_color': colors['primary_dark'], 'align': 'center', 'valign': 'vcenter'
    })
    ws5.merge_range('A1:L1', '📊 ASHOK LEYLAND - Executive Decision Dashboard', title_dash)
    ws5.set_row(0, 35)
    
    # Executive Summary Box
    ws5.merge_range('A3:D3', '📋 EXECUTIVE SUMMARY', title_sub)
    
    if feasible_r:
        best_loc = feasible_r[0]
        
        # Best location highlight
        highlight_fmt = wb.add_format({
            'bold': True, 'font_size': 14, 'bg_color': colors['gold'],
            'border': 2, 'border_color': colors['warning'],
            'align': 'center', 'valign': 'vcenter'
        })
        ws5.merge_range('A4:D4', f'🏆 RECOMMENDED: {best_loc.get("locationName", "")}', highlight_fmt)
        ws5.set_row(3, 30)
        
        # Metrics
        metrics = [
            ('TOPSIS Score', round(best_loc.get('compositeScore', 0), 4), colors['success_light']),
            ('Total Locations', len(locations), colors['neutral_light']),
            ('Feasible Locations', len(feasible_r), colors['success_light']),
            ('Infeasible Locations', len(infeasible_r), colors['danger_light'])
        ]
        
        for i, (metric, value, color) in enumerate(metrics):
            row = 5 + i
            ws5.write(row, 0, metric, wb.add_format({'bold': True, 'align': 'left'}))
            val_fmt = wb.add_format({
                'bg_color': color, 'bold': True, 'border': 1,
                'align': 'center', 'num_format': '0.0000' if isinstance(value, float) else '0'
            })
            ws5.write(row, 1, value, val_fmt)
    
    # Ranking Table
    rank_start = 10
    ws5.merge_range(f'A{rank_start}:E{rank_start}', '🏆 LOCATION RANKINGS', title_sub)
    
    rank_headers = ['Rank', 'Location', 'TOPSIS Score', 'Region', 'State']
    for j, h in enumerate(rank_headers):
        ws5.write(rank_start + 1, j, h, header_location)
    
    for ri, r in enumerate(feasible_r[:10]):
        row = rank_start + 2 + ri
        loc_match = next((l for l in locations if l.get('name') == r.get('locationName')), {})
        
        rank_val = r.get('rank', 0)
        if rank_val == 1:
            ws5.write(row, 0, rank_val, rank_gold)
        elif rank_val == 2:
            ws5.write(row, 0, rank_val, rank_silver)
        elif rank_val == 3:
            ws5.write(row, 0, rank_val, rank_bronze)
        else:
            ws5.write(row, 0, rank_val, cell_center)
        
        ws5.write(row, 1, r.get('locationName', ''), cell_left)
        ws5.write_number(row, 2, round(r.get('compositeScore', 0), 4), cell_number_4dec)
        ws5.write(row, 3, loc_match.get('region', ''), cell_left)
        ws5.write(row, 4, loc_match.get('state', ''), cell_left)
    
    # CHART 1: Ranking Bar Chart
    chart1 = wb.add_chart({'type': 'column'})
    chart1.add_series({
        'name': 'TOPSIS Score',
        'categories': ['4_TOPSIS_Ranking', 4, 1, min(len(feasible_r) + 3, 13), 1],
        'values': ['4_TOPSIS_Ranking', 4, 2, min(len(feasible_r) + 3, 13), 2],
        'fill': {'color': colors['primary']},
        'border': {'color': colors['primary_dark']},
        'data_labels': {'value': True, 'num_format': '0.0000'}
    })
    chart1.set_title({
        'name': 'Top 10 Location Rankings by TOPSIS Score',
        'name_font': {'size': 12, 'bold': True, 'color': colors['primary_dark']}
    })
    chart1.set_x_axis({
        'name': 'Location',
        'name_font': {'size': 10},
        'num_font': {'rotation': -45}
    })
    chart1.set_y_axis({
        'name': 'TOPSIS Score',
        'min': 0, 'max': 1,
        'num_format': '0.00'
    })
    chart1.set_legend({'none': True})
    chart1.set_size({'width': 600, 'height': 350})
    chart1.set_chartarea({
        'fill': {'color': colors['bg_light']},
        'border': {'color': colors['neutral_light']}
    })
    ws5.insert_chart('G3', chart1)
    
    # CHART 2: Weight Distribution Pie Chart
    chart2 = wb.add_chart({'type': 'pie'})
    chart2.add_series({
        'name': 'Criteria Weights',
        'categories': ['3_Weight_Calculation', len(pairwise) + 8, 1, len(pairwise) + 13, 1],
        'values': ['3_Weight_Calculation', len(pairwise) + 8, 2, len(pairwise) + 13, 2],
        'data_labels': {'percentage': True, 'category': True},
        'points': [
            {'fill': {'color': colors['primary']}},
            {'fill': {'color': colors['success']}},
            {'fill': {'color': colors['warning']}},
            {'fill': {'color': colors['danger']}},
            {'fill': {'color': colors['accent']}},
            {'fill': {'color': colors['neutral']}}
        ]
    })
    chart2.set_title({
        'name': 'Hybrid Criteria Weight Distribution',
        'name_font': {'size': 12, 'bold': True, 'color': colors['primary_dark']}
    })
    chart2.set_size({'width': 450, 'height': 350})
    ws5.insert_chart('G20', chart2)
    
    # CHART 3: Criteria Comparison Line Chart
    if feasible_r:
        chart3 = wb.add_chart({'type': 'line'})
        colors_top3 = [colors['gold'], colors['silver'], colors['bronze']]
        
        for i in range(min(3, len(feasible_r))):
            r = feasible_r[i]
            row = 4 + i
            chart3.add_series({
                'name': r.get('locationName', ''),
                'categories': ['4_TOPSIS_Ranking', 3, 3, 3, 8],
                'values': ['4_TOPSIS_Ranking', row, 3, row, 8],
                'line': {'color': colors_top3[i], 'width': 2.5},
                'marker': {'type': 'circle', 'size': 8, 'fill': {'color': colors_top3[i]}},
                'data_labels': {'value': True, 'num_format': '0.00'}
            })
        
        chart3.set_title({
            'name': 'Top 3 Locations - Criteria Score Comparison',
            'name_font': {'size': 12, 'bold': True, 'color': colors['primary_dark']}
        })
        chart3.set_x_axis({'name': 'Criteria'})
        chart3.set_y_axis({'name': 'Normalized Score', 'min': 0, 'max': 1})
        chart3.set_legend({'position': 'bottom'})
        chart3.set_size({'width': 600, 'height': 350})
        ws5.insert_chart('A25', chart3)
    
    # CHART 4: Feasibility Summary Pie
    chart4 = wb.add_chart({'type': 'pie'})
    chart4.add_series({
        'name': 'Feasibility',
        'categories': ['1_Raw_Data_Constraints', len(locations) + 8, 0, len(locations) + 9, 0],
        'values': ['1_Raw_Data_Constraints', len(locations) + 8, 1, len(locations) + 9, 1],
        'data_labels': {'percentage': True, 'category': True},
        'points': [
            {'fill': {'color': colors['success']}},
            {'fill': {'color': colors['danger']}}
        ]
    })
    chart4.set_title({
        'name': 'Feasibility Distribution',
        'name_font': {'size': 11, 'bold': True}
    })
    chart4.set_size({'width': 350, 'height': 300})
    ws5.insert_chart('M25', chart4)
    
    wb.close()
    output_buffer.seek(0)
    return output_buffer

    def _get_val(loc, key):
        """Resolve dotted key like 'raw.tier1Vendors' or 'vendorBase'."""
        if '.' in key:
            parts = key.split('.', 1)
            sub = loc.get(parts[0], {})
            if isinstance(sub, dict):
                return sub.get(parts[1], '')
            return ''
        return loc.get(key, '')

    # Build flat column list
    flat_cols = []
    for sec_name, cols in SECTIONS:
        flat_cols.append(('__section__', sec_name, False))
        flat_cols.extend(cols)

    # Row 0: section spans; Row 1: column labels; Row 2+ data
    # First pass: compute section start columns
    col_idx = 0
    sec_spans = []   # (col_start, col_end, sec_name)
    col_labels = []  # (col_idx, label, is_numeric)
    for sec_name, cols in SECTIONS:
        start = col_idx
        for label, key, is_num in cols:
            col_labels.append((col_idx, label, key, is_num))
            col_idx += 1
        sec_spans.append((start, col_idx - 1, sec_name))

    total_cols = col_idx

    # Row 0: section merged headers
    for (cs, ce, sn) in sec_spans:
        if cs == ce:
            ws1.write(0, cs, sn, sec_hdr)
        else:
            ws1.merge_range(0, cs, 0, ce, sn, sec_hdr)

    # Row 1: column labels
    for (ci, lbl, key, is_num) in col_labels:
        ws1.write(1, ci, lbl, hdr)
        ws1.set_column(ci, ci, max(12, len(lbl.replace('\n', '')) + 2))

    # Row 2+: data
    for ri, loc in enumerate(locations):
        data_row = ri + 2
        for (ci, lbl, key, is_num) in col_labels:
            val = _get_val(loc, key)
            if is_num:
                try:
                    ws1.write_number(data_row, ci, float(val), num2)
                except (ValueError, TypeError):
                    ws1.write(data_row, ci, val, cell)
            else:
                ws1.write(data_row, ci, str(val) if val is not None else '', cell_l)

    ws1.set_row(0, 30)
    ws1.set_row(1, 40)

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 2 — Normalised Parent-Score Matrix
    # ─────────────────────────────────────────────────────────────────────────
    ws2 = wb.add_worksheet('2_Normalised_Matrix')
    ws2.set_column(0, 0, 26)
    criteria_keys = ['vendorBase', 'manpowerAvailability', 'capex',
                     'govtNorms', 'logisticsCost', 'economiesOfScale']
    criteria_labels = ['Vendor Base', 'Manpower', 'CAPEX', 'Govt Norms',
                       'Logistics Cost', 'Economies']
    is_cost_map = {'vendorBase': False, 'manpowerAvailability': False,
                   'capex': True, 'govtNorms': False,
                   'logisticsCost': True, 'economiesOfScale': False}

    ws2.write(0, 0, 'Min-Max Normalised Decision Matrix (0 = worst, 1 = best)', title)
    ws2.write_row(1, 0, ['Location'] + criteria_labels, hdr)
    [ws2.set_column(j+1, j+1, 15) for j in range(len(criteria_keys))]

    # Compute min/max per criterion
    mm = {}
    for k in criteria_keys:
        vals = [float(loc.get(k, 0)) for loc in locations]
        lo, hi = min(vals), max(vals)
        mm[k] = (lo, hi)

    for ri, loc in enumerate(locations):
        row_data = [loc.get('name', '')]
        for k in criteria_keys:
            lo, hi = mm[k]
            raw_v = float(loc.get(k, 0))
            if hi != lo:
                norm = (raw_v - lo) / (hi - lo)
                norm = (1 - norm) if is_cost_map[k] else norm
            else:
                norm = 0.5
            row_data.append(round(norm, 4))
        ws2.write(ri + 2, 0, loc.get('name', ''), cell_l)
        for ci, v in enumerate(row_data[1:]):
            ws2.write_number(ri + 2, ci + 1, v, num4)

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 3 — Weight Calculation (AHP + Entropy + Hybrid)
    # ─────────────────────────────────────────────────────────────────────────
    ws3 = wb.add_worksheet('3_Weight_Calculation')
    ws3.set_column(0, 0, 30)
    [ws3.set_column(j+1, j+1, 13) for j in range(len(criteria_keys))]

    ws3.write(0, 0, 'AHP Pairwise Comparison Matrix', title)
    ws3.write_row(1, 0, [''] + criteria_labels, hdr)
    for i, mrow in enumerate(pairwise):
        ws3.write(i + 2, 0, criteria_labels[i] if i < len(criteria_labels) else f'C{i+1}', cell_l)
        for j, v in enumerate(mrow):
            ws3.write_number(i + 2, j + 1, v, num2)

    start_w = len(pairwise) + 4
    ws3.write(start_w, 0, 'Weight Summary', title)
    ws3.write_row(start_w + 1, 0, ['Method'] + criteria_labels, hdr)

    ahp_w   = [w.get('ahpWeight', 0)      for w in weights]
    ent_w   = [w.get('entropyWeight', 0)  for w in weights]
    comb_w  = [w.get('combinedWeight', 0) for w in weights]

    ws3.write(start_w + 2, 0, 'AHP Weight', cell_l)
    ws3.write(start_w + 3, 0, 'Entropy Weight', cell_l)
    ws3.write(start_w + 4, 0, 'Hybrid (60% AHP + 40% Entropy)', cell_l)
    for j, (a, e, c) in enumerate(zip(ahp_w, ent_w, comb_w)):
        ws3.write_number(start_w + 2, j + 1, a, pct_f)
        ws3.write_number(start_w + 3, j + 1, e, pct_f)
        ws3.write_number(start_w + 4, j + 1, c, pct_f)

    ws3.write_row(start_w + 5, 0, ['Cost / Benefit'] + [
        'Cost ↓' if is_cost_map[k] else 'Benefit ↑' for k in criteria_keys
    ], hdr)

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 4 — TOPSIS Ranking
    # ─────────────────────────────────────────────────────────────────────────
    ws4 = wb.add_worksheet('4_TOPSIS_Ranking')
    ws4.set_column(0, 0, 8)
    ws4.set_column(1, 1, 28)
    ws4.set_column(2, 2, 14)
    [ws4.set_column(j+3, j+3, 12) for j in range(len(criteria_keys))]
    ws4.set_column(len(criteria_keys)+3, len(criteria_keys)+3, 10)

    ws4.write(0, 0, 'TOPSIS Ranking — Ashok Leyland Plant Location Decision', title)
    ws4.write_row(1, 0,
        ['Rank', 'Location', 'TOPSIS Score'] + criteria_labels + ['Feasible'], hdr)

    for ri, r in enumerate(feasible_r):
        cs = r.get('criteriaScores', {})
        fmt_use = best if r.get('rank') == 1 else cell
        ws4.write(ri + 2, 0, r.get('rank', 0), fmt_use)
        ws4.write(ri + 2, 1, r.get('locationName', ''), cell_l)
        ws4.write_number(ri + 2, 2, round(r.get('compositeScore', 0), 4), best if r.get('rank') == 1 else num4)
        for j, k in enumerate(criteria_keys):
            ws4.write_number(ri + 2, j + 3, round(cs.get(k, 0), 4), fmt_use)
        ws4.write(ri + 2, len(criteria_keys) + 3, 'Yes', cell)

    for ri, r in enumerate(infeasible_r):
        row_i = len(feasible_r) + ri + 2
        inf_fmt = wb.add_format({'border': 1, 'align': 'center',
                                 'font_color': '#888888', 'italic': True})
        ws4.write(row_i, 0, '—', inf_fmt)
        ws4.write(row_i, 1, r.get('locationName', ''), inf_fmt)
        ws4.write(row_i, 2, 'N/A', inf_fmt)
        for j in range(len(criteria_keys)):
            ws4.write(row_i, j + 3, '—', inf_fmt)
        ws4.write(row_i, len(criteria_keys) + 3, 'No', inf_fmt)

    # ─────────────────────────────────────────────────────────────────────────
    # SHEET 5 — Dashboard (Summary + Chart)
    # ─────────────────────────────────────────────────────────────────────────
    ws5 = wb.add_worksheet('5_Dashboard')
    ws5.set_column(0, 0, 30)
    ws5.set_column(1, 1, 20)

    ws5.write(0, 0, 'Ashok Leyland Plant Location Decision System', title)
    ws5.write(1, 0, 'Executive Summary & Rankings Dashboard',
              wb.add_format({'italic': True, 'font_color': '#555555'}))

    if feasible_r:
        best_loc = feasible_r[0]
        ws5.write(3, 0, 'Recommended Location',
                  wb.add_format({'bold': True, 'font_size': 13, 'font_color': '#1e3a5f'}))
        ws5.write(3, 1, best_loc.get('locationName', ''),
                  wb.add_format({'bold': True, 'font_size': 13, 'bg_color': '#c6efce',
                                 'border': 1}))
        ws5.write(4, 0, 'TOPSIS Score', cell_l)
        ws5.write_number(4, 1, round(best_loc.get('compositeScore', 0), 4), num4)
        ws5.write(5, 0, 'Total Locations', cell_l)
        ws5.write(5, 1, len(locations), cell)
        ws5.write(6, 0, 'Feasible Locations', cell_l)
        ws5.write(6, 1, len(feasible_r), cell)

    # Ranking table
    r_start = 9
    ws5.write(r_start, 0, 'Location Rankings', sub_t)
    ws5.write_row(r_start + 1, 0, ['Rank', 'Location', 'TOPSIS Score', 'Region', 'State'], hdr)
    for ri, r in enumerate(feasible_r):
        loc_match = next((l for l in locations if l.get('name') == r.get('locationName')), {})
        ws5.write(r_start + 2 + ri, 0, r.get('rank'), cell)
        ws5.write(r_start + 2 + ri, 1, r.get('locationName', ''), cell_l)
        ws5.write_number(r_start + 2 + ri, 2, round(r.get('compositeScore', 0), 4), num4)
        ws5.write(r_start + 2 + ri, 3, loc_match.get('region', ''), cell_l)
        ws5.write(r_start + 2 + ri, 4, loc_match.get('state', ''), cell_l)

    # Bar chart — TOPSIS scores
    if feasible_r:
        chart = wb.add_chart({'type': 'column'})
        chart.add_series({
            'name':       'TOPSIS Score',
            'categories': ['4_TOPSIS_Ranking', 2, 1, len(feasible_r) + 1, 1],
            'values':     ['4_TOPSIS_Ranking', 2, 2, len(feasible_r) + 1, 2],
            'fill':       {'color': '#2d5a8e'},
        })
        chart.set_title({'name': 'Location Rankings by TOPSIS Score'})
        chart.set_x_axis({'name': 'Location'})
        chart.set_y_axis({'name': 'TOPSIS Score', 'min': 0, 'max': 1})
        chart.set_legend({'none': True})
        chart.set_size({'width': 600, 'height': 320})
        ws5.insert_chart('A' + str(r_start + len(feasible_r) + 5), chart)

        # Weights chart
        chart2 = wb.add_chart({'type': 'bar'})
        chart2.add_series({
            'name':       'Combined Weight',
            'categories': ['3_Weight_Calculation', len(pairwise) + 5, 1,
                           len(pairwise) + 5, len(criteria_keys)],
            'values':     ['3_Weight_Calculation', len(pairwise) + 6, 1,
                           len(pairwise) + 6, len(criteria_keys)],
            'fill':       {'color': '#1d7a8a'},
        })
        chart2.set_title({'name': 'Hybrid Criterion Weights'})
        chart2.set_legend({'none': True})
        chart2.set_size({'width': 480, 'height': 260})
        ws5.insert_chart('G' + str(r_start + len(feasible_r) + 5), chart2)

    wb.close()
    output_buffer.seek(0)
    return output_buffer

# ═════════════════════════════════════════════════════════════════════════════
# API ENDPOINTS
# ═════════════════════════════════════════════════════════════════════════════

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload and process Excel/CSV file"""
    try:
        contents = await file.read()
        buffer = BytesIO(contents)

        # Detect file type and read
        if file.filename.endswith('.csv'):
            df = pd.read_csv(buffer)
        else:
            # For Excel, try to find the right header row
            df = pd.read_excel(buffer, header=1)  # Row 1 contains actual headers

        # Process data
        locations = process_excel_data(df)

        return {
            "count": len(locations),
            "locations": locations,
            "filename": file.filename
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error processing file: {str(e)}")

@app.post("/api/analyze")
async def analyze_locations(request: AnalysisRequest):
    """Run AHP-Entropy-TOPSIS analysis"""
    try:
        locations = request.locations
        matrix = request.pairwiseMatrix
        constraints = [c.dict() for c in request.constraints]

        # Criteria configuration
        criteria_keys = ['vendorBase', 'manpowerAvailability', 'capex', 'govtNorms', 'logisticsCost', 'economiesOfScale']
        is_cost = {
            'vendorBase': False,
            'manpowerAvailability': False,
            'capex': True,
            'govtNorms': False,
            'logisticsCost': True,
            'economiesOfScale': False
        }

        # Step 1: Apply constraints
        feasible_locs, infeasible_locs = apply_constraints(locations, constraints)

        if not feasible_locs:
            return {
                "weights": [],
                "results": [],
                "error": "No feasible locations match the constraints"
            }

        # Step 2: Calculate AHP weights
        ahp_weights, cr = calculate_ahp_weights(matrix)

        # Step 3: Calculate Entropy weights
        # Build decision matrix for entropy calculation
        X = np.array([[loc.get(k, 0) for k in criteria_keys] for loc in feasible_locs])
        entropy_weights = calculate_entropy_weights(X)

        # Step 4: Hybrid weights (60% AHP + 40% Entropy)
        hybrid_weights = [0.6 * a + 0.4 * e for a, e in zip(ahp_weights, entropy_weights)]

        # Normalize hybrid weights
        total = sum(hybrid_weights)
        hybrid_weights = [w / total for w in hybrid_weights]

        # Step 5: TOPSIS analysis
        scores = topsis_analysis(feasible_locs, hybrid_weights, criteria_keys, is_cost)

        # Prepare weight response
        weight_response = []
        for i, key in enumerate(criteria_keys):
            weight_response.append({
                'key': key,
                'name': criteria_keys[i].replace('vendorBase', 'Vendor Base')
                                       .replace('manpowerAvailability', 'Manpower')
                                       .replace('capex', 'CAPEX')
                                       .replace('govtNorms', 'Govt Norms')
                                       .replace('logisticsCost', 'Logistics')
                                       .replace('economiesOfScale', 'Economies'),
                'isCost': is_cost[key],
                'ahpWeight': round(ahp_weights[i], 4),
                'entropyWeight': round(entropy_weights[i], 4),
                'combinedWeight': round(hybrid_weights[i], 4)
            })

        # Prepare results with rankings
        results = []
        for i, (loc, score) in enumerate(zip(feasible_locs, scores)):
            # Calculate individual criterion scores (normalized)
            criteria_scores = {}
            for j, key in enumerate(criteria_keys):
                val = loc.get(key, 0)
                # Simple normalization for display
                max_val = max(l.get(key, 0) for l in feasible_locs)
                min_val = min(l.get(key, 0) for l in feasible_locs)
                if max_val > min_val:
                    if is_cost[key]:
                        criteria_scores[key] = round((max_val - val) / (max_val - min_val), 3)
                    else:
                        criteria_scores[key] = round((val - min_val) / (max_val - min_val), 3)
                else:
                    criteria_scores[key] = 0.5

            results.append({
                'locationId': loc.get('name', ''),
                'locationName': loc.get('name', ''),
                'compositeScore': round(score, 4),
                'criteriaScores': criteria_scores,
                'feasible': True,
                'rank': 0  # Will be assigned after sorting
            })

        # Sort by score and assign ranks
        results.sort(key=lambda x: x['compositeScore'], reverse=True)
        for i, r in enumerate(results):
            r['rank'] = i + 1

        # Add infeasible locations
        for loc in infeasible_locs:
            results.append({
                'locationId': loc.get('name', ''),
                'locationName': loc.get('name', ''),
                'compositeScore': 0,
                'criteriaScores': {},
                'feasible': False,
                'rank': 999
            })

        return {
            "weights": weight_response,
            "results": results,
            "consistencyRatio": round(cr, 4)
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Analysis error: {str(e)}")

@app.post("/api/monte-carlo")
async def monte_carlo(request: MonteCarloRequest):
    """Run Monte Carlo simulation for robustness analysis"""
    try:
        locations = request.locations
        weights_data = request.weights
        iterations = request.iterations
        constraints = [c.dict() for c in request.constraints]
        region_filter = request.regionFilter

        # Apply constraints first
        feasible_locs, _ = apply_constraints(locations, constraints, region_filter)

        if not feasible_locs:
            return {"monteCarloResults": []}

        criteria_keys = ['vendorBase', 'manpowerAvailability', 'capex', 'govtNorms', 'logisticsCost', 'economiesOfScale']
        is_cost = {
            'vendorBase': False,
            'manpowerAvailability': False,
            'capex': True,
            'govtNorms': False,
            'logisticsCost': True,
            'economiesOfScale': False
        }

        # Extract base weights
        base_weights = [w.get('combinedWeight', 0) for w in weights_data]

        # Run simulation
        mc_results = monte_carlo_simulation(
            feasible_locs, base_weights, criteria_keys, is_cost, iterations
        )

        return {"monteCarloResults": mc_results}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Simulation error: {str(e)}")

@app.post("/api/export-excel")
async def export_excel(request: ExportRequest):
    """Export comprehensive Excel report"""
    try:
        data = {
            'locations': request.locations,
            'results': request.results,
            'weights': request.weights,
            'pairwiseMatrix': request.pairwiseMatrix,
            'constraints': [c.dict() for c in request.constraints],
            'regionFilter': request.regionFilter
        }

        output = BytesIO()
        create_excel_report(data, output)

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Ashok_Leyland_Analysis.xlsx"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export error: {str(e)}")

@app.get("/api/health")
async def health_check():
    return {"status": "healthy", "service": "Ashok Leyland Plant Location Decision API"}