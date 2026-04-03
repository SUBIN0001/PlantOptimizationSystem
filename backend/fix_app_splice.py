import traceback

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_content = """def process_flexible_excel(file_content: bytes, filename: str) -> Tuple[List[Dict], Dict]:
    '''
    Unified pipeline that attempts detailed extraction first using robust keyword mapping
    (preserving 'raw' keys expected by frontend), falls back to core-only, validates,
    normalises, and reports mode.
    '''
    buffer = BytesIO(file_content)

    if filename.lower().endswith('.csv'):
        df = pd.read_csv(buffer)
    else:
        try:
            df = pd.read_excel(buffer, header=None)
        except Exception:
            buffer.seek(0)
            df = pd.read_excel(buffer)

    start_row = detect_data_start_row(df)
    logger.info(f"Data starts at row: {start_row}")

    if start_row > 0 and start_row - 1 < len(df):
        df.columns = [str(c).replace('\\n', ' ').strip() for c in df.iloc[start_row - 1].values]
        df = df.iloc[start_row:].copy()
    else:
        df.columns = [str(c).replace('\\n', ' ').strip() for c in df.columns]

    row_strings = [str(c).lower() for c in df.columns]
    detailed_keywords = ['acma', 'tier-1', 'iti', 'subsidy', 'freight', 'construction', 'wage']
    has_detailed = any(any(kw in col for kw in detailed_keywords) for col in row_strings)

    if has_detailed:
        locations = process_excel_data_cleaned(df)
        mode = "detailed"
        logger.info(f"Mode selected: detailed ({len(locations)} locations)")
    else:
        locations = extract_core_attributes(df, 0)
        mode = "core"
        logger.info(f"Mode selected: core ({len(locations)} locations)")

    locations = validate_locations(locations)
    logger.info(f"Valid locations after validation: {len(locations)}")

    locations = normalize_criteria_values(locations)

    core_mapping = detect_column_mapping(df, 0)
    metadata = {
        'total_locations': len(locations),
        'mode': mode,
        'data_start_row': int(start_row),
        'columns_detected': list(core_mapping.keys()),
        'missing_attributes': [
            attr for attr in REQUIRED_ATTRIBUTES if attr not in core_mapping
        ],
    }

    return locations, metadata

def process_excel_data_cleaned(df: pd.DataFrame) -> List[Dict]:
    '''Process detailed finaleyy.xlsx format into exact frontend expected keys.'''
    loc_col = None
    for col in df.columns:
        if str(col).lower() in ['location', 'site', 'city', 'place', 'name']:
            loc_col = col
            break

    if loc_col:
        df = df[df[loc_col].notna()]
        df = df[df[loc_col].astype(str).str.lower() != str(loc_col).lower()]

    locations = []
    for _, row in df.iterrows():
        try:
            def gs(col, default=''):
                v = row.get(col, default)
                return str(v).strip() if not pd.isna(v) else str(default)

            def find_col(keywords, default=''):
                for k in row.keys():
                    k_lower = str(k).lower().replace('\\n', ' ')
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
                'name': str(row[loc_col]) if loc_col else 'Unknown',
                'region': gs('Region'),
                'state': gs('State'),
                'vendorBase': compute_vendor_base_score(row),
                'manpowerAvailability': compute_manpower_score(row),
                'capex': extract_numeric(row.get('Estimated Total  Project CAPEX (₹ Cr)*', 0)),
                'govtNorms': compute_govt_score(row),
                'logisticsCost': compute_logistics_score(row),
                'economiesOfScale': compute_economies_score(row),
                'raw': {
                    'industrialPark':       gs('Industrial Park / Zone'),
                    'acmaCluster':          gs('ACMA Auto Component Cluster'),
                    'acmaUnits':            gs('No. of ACMA Member Units  (State, approx.)'),
                    'tier1Vendors':         ext_num_find(['Tier-1'], 'Tier-1 Auto Vendors  within 200 km (nos.)'),
                    'tier2Vendors':         ext_num_find(['Tier-2'], 'Tier-2 Auto Vendors  within 200 km (nos.)'),
                    'steelSuppliers':       ext_num_find(['Steel', 'Castings'], 'Steel / Castings Suppliers  within 100 km (nos.)'),
                    'vendorEcosystem':      gs('Vendor Ecosystem  Rating'),
                    'keyOEMs':              gs('Key OEMs / Anchors  in the Cluster'),
                    'enggColleges':         ext_num_find(['Engg', 'Colleges'], 'AICTE Engg Colleges  in 50 km radius (nos.)'),
                    'itiInstitutes':        ext_num_find(['ITI'], 'ITI Institutes  in 50 km radius (nos.)'),
                    'itiGraduates':         ext_num_find(['ITI', 'Graduates'], 'Annual ITI Graduates  (State, 000s)'),
                    'enggGraduates':        ext_num_find(['Engg', 'Graduates'], 'Annual Engg Graduates  (State, 000s)'),
                    'skilledLabourRating':  gs('Skilled Labour  Availability Rating'),
                    'wageSkilled':          ext_num_find(['Wage', 'Skilled'], 'Avg Monthly Wage –  Skilled Mfg (₹)'),
                    'wageSemiSkilled':      ext_num_find(['Wage', 'Semi-Skilled'], 'Avg Monthly Wage –  Semi-Skilled (₹)'),
                    'attritionRate':        ext_num_find(['Attrition'], 'Labour Attrition Rate  (%/yr, est.)'),
                    'landCost':             ext_num_find(['Land Cost'], 'Industrial Land Cost  (₹ Cr / Acre)'),
                    'availableLand':        ext_num_find(['Available', 'Land'], 'Available Land  (Acres, approx.)'),
                    'constructionIndex':    ext_num_find(['Construction'], 'Construction Cost  Index (Base TN=100)'),
                    'powerCapex':           ext_num_find(['Power', 'Connection'], 'Power Connection  Capex (₹ Cr, est.)'),
                    'waterCapex':           ext_num_find(['Water', 'Utilities'], 'Water / Utilities  Capex (₹ Cr, est.)'),
                    'totalCapex':           ext_num_find(['Estim', 'Total', 'CAPEX'], 'Estimated Total  Project CAPEX (₹ Cr)*'),
                    'industrialPolicy':     gs_find(['Industrial', 'Policy'], 'State Industrial Policy  (Current)'),
                    'capitalSubsidy':       ext_num_find(['Capital Subsidy'], 'Capital Subsidy  (% of Fixed Assets)'),
                    'sgstExemption':        ext_num_find(['SGST'], 'SGST Exemption /  Refund Period (yrs)'),
                    'stampDuty':            gs_find(['Stamp Duty'], 'Stamp Duty  Exemption'),
                    'powerTariff':          ext_num_find(['Power Tariff'], 'Power Tariff – HT  Industrial (₹/kWh)'),
                    'elecDutyExemption':    gs_find(['Electricity Duty'], 'Electricity Duty  Exemption'),
                    'envClearanceEase':     ext_num_find(['Env. Clearance'], 'Env. Clearance  Ease (1-10)'),
                    'approvalDays':         ext_num_find(['Single Window'], 'Single Window  Approval Days (est.)'),
                    'sezNimz':              gs_find(['SEZ'], 'SEZ / NIMZ /  Special Zone'),
                    'dfcAccessGovt':        gs_find(['Dedicated', 'Freight'], 'Dedicated Freight  Corridor Access'),
                    'nearestPort':          gs_find(['Nearest', 'Port'], 'Nearest Major Port'),
                    'distanceToPort':       ext_num_find(['Distance', 'Port'], 'Distance to Port  (km)'),
                    'roadConnectivity':     ext_num_find(['Road', 'Connectivity'], 'Road / NH  Connectivity (1-10)'),
                    'railConnectivity':     ext_num_find(['Rail', 'Connectivity'], 'Rail Connectivity  (1-10)'),
                    'dfcLogistics':         gs_find(['DFC', 'Access'], 'DFC Access  (Y/N)'),
                    'distanceKeyMarket':    ext_num_find(['Nearest', 'Market'], 'Distance to Nearest  Key Market (km)'),
                    'keyMarketCity':        gs_find(['Key Market City'], 'Key Market City'),
                    'inboundFreight':       ext_num_find(['Inbound', 'Freight'], 'Inbound Freight Rate  (₹/MT)'),
                    'outboundFreight':      ext_num_find(['Outbound', 'Freight'], 'Outbound Freight Rate  (₹/MT)'),
                    'annualLogisticsCost':  ext_num_find(['Annual', 'Logistics'], 'Annual Logistics Cost  (₹ Cr/yr, est.)**'),
                    'clusterMaturity':      gs_find(['Cluster', 'Maturity'], 'Auto Industry Cluster  Maturity'),
                    'existingCVOEMs':       gs_find(['CV', 'OEMs'], 'Existing CV / Commercial  OEMs nearby'),
                    'supplierPark':         gs_find(['Supplier Park'], 'Supplier Park  Availability'),
                    'exportHub':            gs_find(['Export', 'Hub'], 'Export Hub  Proximity'),
                    'marketDemandIndex':    ext_num_find(['Market Demand'], 'Market Demand  Index (1-10)'),
                    'clusterBenefitScore':  ext_num_find(['Cluster Benefit'], 'Cluster Benefit  Score (1-10)'),
                }
            }
            locations.append(loc)
        except Exception as e:
            continue

    return locations
"""

start1 = -1
end1 = -1
start2 = -1
end2 = -1

for i, line in enumerate(lines):
    if line.startswith('def process_excel_data(df):'):
        start1 = i
    if start1 != -1 and line.startswith('# FLEXIBLE DATA PROCESSING HELPERS'):
        end1 = i - 1
        break

for i, line in enumerate(lines):
    if line.startswith('def _detect_detailed_column_mapping('):
        start2 = i
    if start2 != -1 and line.startswith('@app.post("/api/upload")'):
        end2 = i - 1
        break

if start1 != -1 and end1 != -1 and start2 != -1 and end2 != -1:
    del lines[start2:end2]
    # insert new content into start2
    lines.insert(start2, new_content + "\n\n")

    # remove start1 to end1
    del lines[start1:end1+1] # +1 to include end1
    
    # replace 'detailed':             {}, with 'raw':             {},
    for i, line in enumerate(lines):
        if "'detailed':" in line and "}," in line:
            lines[i] = line.replace("'detailed':", "'raw':")

    with open('app.py', 'w', encoding='utf-8') as f:
        f.writelines(lines)
    print("SUCCESS")
else:
    print("FAILED TO FIND MARKERS:", start1, end1, start2, end2)
