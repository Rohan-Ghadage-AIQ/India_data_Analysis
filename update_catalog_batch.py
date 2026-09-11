#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Universal Batch Update for Master Catalog and Child Sheets
=========================================================
Author: Senior Geospatial Analyst & Software Engineer

Usage:
    python update_catalog_batch.py --start 101 --end 200
    python update_catalog_batch.py --start 1 --end 100

Features:
1. Automatically parses and matches Description/*.docx files with ArcGIS URLs.
2. Formulates 5-Point Geospatial Analysis Protocols (How to Use This File) and Target Industries.
3. Updates Master Sheet ('GeoJSON Catalog') rows with 115pt height, links, and formatting.
4. Formats Child Sheets:
   - Row 3 Header: 6 columns with Navy fill & White text:
     ['Sr. No', 'Layer Available Fields', 'Target Industries', 'Potential Analysis', 'Potential Analysis Manual', 'Remark']
   - Col 3 ('Target Industries'): Domain-specific stakeholder listing.
   - Col 4 ('Potential Analysis'): Attribute-level Senior Geospatial Analyst analytical methodology.
   - Col 5 ('Potential Analysis Manual'): STRICTLY VACANT / None across all data rows for user manual entry.
   - Col 6 ('Remark'): Concise, high-value technical context notes.
   - Zebra striping, 48pt row height, freeze panes at A4, back link at A1.
5. Handles Excel process lock and cleans up AutoRecover cache.
"""

import os
import sys
import re
import glob
import json
import difflib
import zipfile
import argparse
import subprocess
import urllib.parse
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

# Palette & Styles
COLOR_PRIMARY_NAVY = "2F5496"
COLOR_PRIMARY_TEXT = "FFFFFF"
COLOR_LINK_BLUE = "0563C1"
COLOR_BORDER_LIGHT = "B4C6E7"
COLOR_BORDER_DARK = "1F3864"
COLOR_ZEBRA_EVEN = "F2F5F9"
COLOR_ZEBRA_ODD = "FFFFFF"
COLOR_BANNER_BG = "D9E1F2"

SIDE_THIN_LIGHT = Side(style="thin", color=COLOR_BORDER_LIGHT)
SIDE_THIN_DARK = Side(style="thin", color=COLOR_BORDER_DARK)
SIDE_MEDIUM_DARK = Side(style="medium", color=COLOR_BORDER_DARK)

BORDER_DATA_CELL = Border(
    left=SIDE_THIN_LIGHT, right=SIDE_THIN_LIGHT,
    top=SIDE_THIN_LIGHT, bottom=SIDE_THIN_LIGHT
)
BORDER_HEADER_CELL = Border(
    left=SIDE_THIN_DARK, right=SIDE_THIN_DARK,
    top=SIDE_THIN_DARK, bottom=SIDE_MEDIUM_DARK
)

FILL_HEADER = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
FILL_ZEBRA_EVEN = PatternFill(start_color=COLOR_ZEBRA_EVEN, end_color=COLOR_ZEBRA_EVEN, fill_type="solid")
FILL_ZEBRA_ODD = PatternFill(start_color=COLOR_ZEBRA_ODD, end_color=COLOR_ZEBRA_ODD, fill_type="solid")

FONT_HEADER = Font(name="Calibri", size=11, bold=True, color=COLOR_PRIMARY_TEXT)
FONT_DATA = Font(name="Calibri", size=10, color="000000")
FONT_LINK = Font(name="Calibri", size=10, color=COLOR_LINK_BLUE, underline="single")

ALIGN_CENTER_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT_TOP_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER_TOP = Alignment(horizontal="center", vertical="top")


def close_running_excel():
    """Silently terminate Excel process if running on Windows to prevent file lock errors."""
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", "Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


def extract_docx_metadata(path):
    """Extract full text, title, and ArcGIS item URL from a docx file."""
    try:
        with zipfile.ZipFile(path) as z:
            xml_data = z.read('word/document.xml')
        tree = ET.fromstring(xml_data)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        paras = []
        for p in tree.iterfind('.//w:p', ns):
            t = ''.join(n.text for n in p.iterfind('.//w:t', ns) if n.text)
            if t.strip():
                paras.append(t.strip())
        full_text = '\n'.join(paras)
        url_match = re.search(r'https://www\.arcgis\.com/home/item\.html\?id=([0-9a-fA-F]{32})', full_text)
        if url_match:
            url = f"https://www.arcgis.com/home/item.html?id={url_match.group(1).lower()}"
        else:
            u2 = re.search(r'https?://[^\s]+', full_text)
            url = u2.group(0).rstrip('.,;)') if u2 else ''
        title = paras[0] if paras else os.path.basename(path).replace('.docx', '')
        return {'path': path, 'title': title, 'url': url, 'text': full_text}
    except Exception:
        return {'path': path, 'title': os.path.basename(path).replace('.docx', ''), 'url': '', 'text': ''}


def unescape_name(s):
    """Decode hex characters like _3A_ -> :, _28_ -> (, _26_ -> &, _2C_ -> , and replace underscores."""
    s = re.sub(r'_([0-9A-Fa-f]{2})_', r'%\1', str(s))
    s = re.sub(r'__', '_', s)
    try:
        s = urllib.parse.unquote(s)
    except Exception:
        pass
    s = s.replace('_', ' ')
    return re.sub(r'\s+', ' ', s).strip()


def normalize_str(s):
    s = unescape_name(s).lower()
    s = re.sub(r'[^a-z0-9]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def load_all_docs(docs_dir="Description"):
    """Load and index all docx documents."""
    docs = []
    for p in glob.glob(os.path.join(docs_dir, "*.docx")):
        docs.append(extract_docx_metadata(p))
    return docs


def match_layer_to_doc(raw_name, docs):
    """Fuzzy match a layer name to the best Description docx file."""
    c_raw = normalize_str(raw_name)
    best_doc = None
    best_score = 0.0

    for d in docs:
        c_title = normalize_str(d['title'])
        c_file = normalize_str(os.path.basename(d['path']).replace('.docx', ''))
        
        # Exact substring priority
        if c_raw in c_title or c_title in c_raw or c_raw in c_file or c_file in c_raw:
            score = 0.90 + 0.10 * (min(len(c_raw), len(c_title)) / max(len(c_raw), len(c_title), 1))
        else:
            score = max(
                difflib.SequenceMatcher(None, c_raw, c_title).ratio(),
                difflib.SequenceMatcher(None, c_raw, c_file).ratio()
            )
        if score > best_score:
            best_score = score
            best_doc = d

    return best_doc if best_score >= 0.40 else None


def classify_domain_and_protocols(sr, name, doc_title):
    """Classify domain and build 5-point Senior Geospatial Analysis Protocol."""
    n_lower = name.lower()
    t_lower = (doc_title or "").lower()
    unescaped = unescape_name(name).lower()
    combined = f"{n_lower} {t_lower} {unescaped}"

    # 1. Road Safety, Police & NCRB Accidents
    if any(k in combined for k in ['accident', 'black spot', 'collision', 'killed', 'ncrb', 'crime', 'police station']):
        domain = 'road_safety'
        target_ind = "Traffic Police & Law Enforcement, Ministry of Road Transport & Highways (MoRTH), National Highways Authority of India (NHAI), Motor Insurance Underwriting, Trauma & Emergency Healthcare"
        p1 = "Spatial Autocorrelation & Black Spot Hotspots: Execute Getis-Ord Gi* and Anselin Local Moran's I to isolate statistically significant High-High fatality and collision clusters across administrative polygons."
        p2 = "Geometric & Attribute Severity Ratios: Compute casualty-to-accident severity ratios across specific roadway features, environmental conditions, or vehicle age brackets to benchmark high-risk factors."
        p3 = "Network Corridor Overlay: Spatially intersect collision hotspots with National Highway (NH) and State Highway (SH) networks to prioritize engineering remediation (rumble strips, guard rails, grade separation)."
        p4 = "Spatio-Temporal Trajectory (2021-2024): Model multi-year directional shifts of collision density to evaluate state road safety council interventions and enforcement campaigns."
        p5 = "Golden Hour Healthcare Allocation: Spatially model distance and travel-time isochrones from high-fatality corridors to Level-1/Level-2 trauma centers to identify emergency medical deficit zones."

    # 2. Transport Infrastructure, Aviation, Maritime, Railways & Tolls
    elif any(k in combined for k in ['airport', 'seaport', 'railway', 'track', 'toll', 'fastag', 'bharatmala', 'highway', 'transport', 'corridor']):
        domain = 'transport_infrastructure'
        target_ind = "Airports Authority of India (AAI), Indian Railways & DFCCIL, Port & Maritime Authorities, National Highways Authority of India (NHAI), Logistics & Multi-Modal Freight Operators"
        p1 = "Multi-Modal Freight Mobility & Terminal Catchment Analysis: Perform network distance and drive-time catchment analysis around airports, seaports, and railway stations to model passenger and cargo throughput."
        p2 = "Corridor Density & Congestion Indexing: Calculate railway track density, toll plaza collection velocity, and terminal capacity utilization per administrative and economic corridor."
        p3 = "Intermodal Connectivity & Bottleneck Overlay: Spatially intersect freight railway lines, national expressways, and port hinterlands to identify freight transfer chokepoints and multimodal logistics gaps."
        p4 = "Spatio-Temporal Infrastructure Expansion Modeling: Track corridor electrification, dedicated freight corridor commissioning, and toll throughput growth trajectories."
        p5 = "Strategic Logistics Asset Location-Allocation: Optimize spatial placement of Multi-Modal Logistics Parks (MMLPs), inland container depots (ICDs), and railway siding facilities."

    # 3. Rural Development, Habitations, PMGSY & Jal Jeevan Mission (JJM)
    elif any(k in combined for k in ['jjm', 'har_ghar_jal', 'mgnrega', 'rural', 'habitation', 'fhtc', 'pmgsy', 'tap water']):
        domain = 'rural_development'
        target_ind = "Ministry of Jal Shakti (National Jal Jeevan Mission), Ministry of Rural Development, State Rural Water & Sanitation Missions (SWSM), Panchayati Raj Institutions"
        p1 = "Rural Infrastructure Saturation & Coverage Hotspot Detection: Execute spatial clustering (Gi*) on Functional Household Tap Connection (FHTC) percentages and MGNREGA employment mandays to identify underserved rural pockets."
        p2 = "Saturation & Disparity Normalization: Calculate percentage coverage of tap water connectivity and MGNREGA job card utilization relative to total rural census households per block/district."
        p3 = "Last-Mile Connectivity & Water Source Overlay: Spatially intersect rural habitations with surface water pipelines, ground water recharge schemes, and PMGSY all-weather road alignments."
        p4 = "Spatio-Temporal Saturation Velocity Modeling (2020-2024): Track quarterly progress toward 100% 'Har Ghar Jal' certification and seasonal shifts in MGNREGA rural labor demand."
        p5 = "Targeted Rural Scheme Capital Allocation: Prioritize infrastructure grant allocations for water-stressed habitations, over-exploited groundwater blocks, and high-demand rural employment clusters."

    # 4. Banking, Finance, Domestic Product & Economy
    elif any(k in combined for k in ['domestic_product', 'net_state_domestic', 'per_capita', 'inflation', 'bank', 'credit', 'deposit', 'wage', 'gva', 'gsdp']):
        domain = 'finance_economy'
        target_ind = "Ministry of Finance, Reserve Bank of India (RBI), Commercial Banks & NBFCs, Regional Economic Development Agencies, Investment & Credit Rating Agencies"
        p1 = "Regional Economic Disparity & GSDP Spatial Clustering: Run Anselin Local Moran's I on Net State Domestic Product and per capita income to isolate high-growth economic corridors versus lagging districts."
        p2 = "Per Capita & Purchasing Power Normalization: Calculate inflation-adjusted per capita economic output, credit-to-deposit ratios, and state revenue contribution indices."
        p3 = "Spatial Macro-Economic Correlation: Perform Geographically Weighted Regression (GWR) correlating state domestic product with infrastructure density, urbanization rates, and workforce participation."
        p4 = "Spatio-Temporal Economic Trajectory Modeling: Track multi-year economic convergence and growth differentials across states to benchmark regional industrial policy impacts."
        p5 = "Capital Investment & Fiscal Resource Allocation: Guide state-level capital expenditure prioritization, commercial banking network expansion, and targeted regional development subsidies."

    # 5. Demographics, Religion, Caste, Workforce & Social Development
    elif any(k in combined for k in ['demographic', 'religion', 'caste', 'socio_economic', 'worker_population', 'labour_force', 'unemployment', 'age_gr', 'population']):
        domain = 'demographics_society'
        target_ind = "Ministry of Social Justice & Empowerment, NITI Aayog, Ministry of Labour & Employment, Census Operations, Academic & Social Policy Research Institutes"
        p1 = "Demographic Composition & Social Vulnerability Hotspots: Execute spatial autocorrelation on religious and caste demographic proportions, worker participation rates, and dependency ratios."
        p2 = "Demographic Parity & Labour Force Normalization: Compute Gender Parity Indices in workforce participation, youth dependency ratios, and socio-economic deprivation indices."
        p3 = "Multi-Tier Spatial Overlay: Intersect demographic and caste distribution with industrial corridors, agricultural belts, and basic civic infrastructure to evaluate equitable development access."
        p4 = "Spatio-Temporal Demographic Transition Modeling: Track decadal shifts in age-cohort structures, urbanization rates, and workforce mobility patterns across states and districts."
        p5 = "Social Welfare Program Resource Allocation: Optimize spatial targeting of affirmative action programs, skill development centres, and targeted minority welfare initiatives."

    # 6. Health, Medical Officers, Facilities & Diseases
    elif any(k in combined for k in ['health', 'chikungunya', 'dengue', 'kala_azar', 'malaria', 'stunted', 'wasted', 'infant_mortality', 'mortality', 'gdmo', 'scs_2c_phcs', 'specialists', 'chcs', 'sncu', 'maternity', 'family_planning', 'life_expectancy']):
        domain = 'health_medicine'
        target_ind = "Ministry of Health & Family Welfare (MoHFW), National Health Authority (NHA), State Public Health Directorates, WHO / Global Health Organizations, Hospital & Diagnostic Networks"
        p1 = "Epidemiological Surveillance & Vector-Borne Disease Hotspots: Execute Getis-Ord Gi* to pinpoint statistically significant spatial clusters of Dengue, Chikungunya, Malaria, and Kala-Azar incidence."
        p2 = "Healthcare Human Resource & Capacity Deficit Indexing: Compute doctor-to-population and specialist-to-CHC ratios across rural districts to identify acute medical manpower deficits."
        p3 = "Facility Catchment & Vulnerability Overlay: Intersect high malnutrition (stunting/wasting) and maternal/infant mortality zones with PHC/CHC locations to model spatial accessibility gaps."
        p4 = "Spatio-Temporal Disease Transmission Trajectory: Track multi-year vector disease outbreaks and child health indicator progress toward Sustainable Development Goals (SDGs)."
        p5 = "Healthcare Infrastructure Location-Allocation: Optimize siting of Special Newborn Care Units (SNCU), blood storage centres, and specialist doctor deployments in underserved rural blocks."

    # 7. Water Resources, River Basins, Floods, Glaciers & Soils
    elif any(k in combined for k in ['river_basin', 'river_sub_basin', 'flood', 'water_depth', 'rainfall_erosivity', 'irrigation', 'inter_basin', 'litholog', 'glacial_lakes', 'soil', 'nutrients', 'monsoon']):
        domain = 'water_climate'
        target_ind = "Central Water Commission (CWC), Central Ground Water Board (CGWB), National Disaster Management Authority (NDMA), Irrigation & Water Resources Departments, Agriculture & Soil Survey Boards"
        p1 = "Hydrological Risk & Flood Inundation Vulnerability Mapping: Perform spatial clustering of historical flood events, river basin discharge points, and glacial lake expansion to demarcate hazard zones."
        p2 = "Hydrogeological Deficit & Soil Nutrient Indexing: Calculate groundwater depth depletion rates, soil macro/micro-nutrient deficiency indices, and rainfall erosivity factors per watershed catchment."
        p3 = "Inter-Basin Water Transfer & Irrigation Network Overlay: Intersect proposed river linking channels with drought-prone agricultural blocks and command areas to evaluate water redistribution efficacy."
        p4 = "Spatio-Temporal Monsoon & Climate Impact Modeling: Model decadal trends in monsoon variability, glacial lake volume fluctuations, and groundwater table fluctuations."
        p5 = "Watershed Management & Disaster Mitigation Siting: Drive spatial multi-criteria siting of artificial groundwater recharge structures, flood retention basins, and micro-irrigation systems."

    # 8. Elections & Governance
    elif any(k in combined for k in ['election', 'vidhan_sabha', 'parliamentary_boundary', 'assembly', 'legislative']):
        domain = 'elections_governance'
        target_ind = "Election Commission of India (ECI), Political Parties & Electoral Strategy Firms, Public Policy Think Tanks, News & Media Networks, Academic Political Scientists"
        p1 = "Electoral Geography & Swing-Constituency Hotspots: Execute spatial autocorrelation (Moran's I) on winning vote shares and victory margins to detect regional party strongholds and swing corridors."
        p2 = "Voter Participation & Turnout Disparity Indexing: Calculate spatial voter turnout ratios, female-to-male voting ratios, and urban voter apathy indices across parliamentary and assembly seats."
        p3 = "Demographic & Electoral Boundary Overlay: Spatially correlate electoral outcomes with demographic composition, literacy rates, and regional infrastructure development levels."
        p4 = "Inter-Election Swing & Volatility Modeling: Track seat flips, vote share swings, and third-party vote distributions across consecutive election cycles (2019 vs 2024)."
        p5 = "Electoral Logistics & Polling Infrastructure Siting: Guide spatial allocation of vulnerable/critical polling booths, central security forces deployment, and EVM strong-room logistics."

    # 9. Smart Cities, Urban Planning & Municipal Wards
    elif any(k in combined for k in ['smart_cities', 'ward_boundaries', 'urban', 'municipal', 'industrial_land', 'industrial_corridor', 'mega_food_park']):
        domain = 'urban_industrial'
        target_ind = "Ministry of Housing & Urban Affairs (MoHUA), Smart Cities Mission, State Industrial Development Corporations (SIDC), Urban Local Bodies (ULBs), Industrial Real Estate Developers"
        p1 = "Urban Infrastructure Saturation & Smart City Catchment Analysis: Perform spatial network analysis and service coverage buffer modeling around designated Smart Cities and industrial land parks."
        p2 = "Ward-Level Civic Density & Amenities Indexing: Compute civic infrastructure density, population per municipal ward, and open space availability indices for localized municipal governance."
        p3 = "Industrial Corridor & Supply Chain Overlay: Intersect industrial park nodes and Mega Food Parks with national freight corridors, highways, and raw material catchments to model supply chain access."
        p4 = "Spatio-Temporal Urban Expansion Modeling: Track urban footprint sprawl, industrial land absorption rates, and infrastructure project completion timelines across municipal jurisdictions."
        p5 = "Civic Facility & Industrial Hub Siting Optimization: Optimize location-allocation of solid waste processing facilities, urban fire stations, and food processing aggregation clusters."

    # 10. Education, Literacy & Human Capital
    elif any(k in combined for k in ['educational', 'enrolment', 'literacy', 'teachers', 'school', 'college']):
        domain = 'education'
        target_ind = "Ministry of Education, State School Education Departments, Higher Education Councils, EdTech & Academic Research Organizations, National Skill Development Corporation (NSDC)"
        p1 = "Educational Access & Gross Enrolment Hotspot Detection: Identify statistically significant spatial clusters of high and low Gross Enrolment Ratios (GER) and literacy rates across districts."
        p2 = "Gender Parity & Pupil-Teacher Ratio Indexing: Compute spatial Gender Parity Index (GPI) and Pupil-Teacher Ratios (PTR) across primary, secondary, and higher secondary institutions."
        p3 = "Institutional Catchment & Network Distance Overlay: Intersect recognized educational institutions with rural habitation centroids to model student travel times and identify school-deficit villages."
        p4 = "Spatio-Temporal Enrolment & Retention Modeling: Track multi-year trends in male/female literacy progression and student retention rates under national educational missions."
        p5 = "Educational Resource & Teacher Allocation: Guide spatial allocation of teacher recruitments, model school construction, and vocational training labs in historically underserved blocks."

    # 11. Administrative Boundaries (Default Fallback)
    else:
        domain = 'admin_boundaries'
        target_ind = "Survey of India (SOI), National Informatics Centre (NIC), Urban & Regional Planning Authorities, Cartographic & GIS Organizations, Census Operations"
        p1 = "Administrative Boundary Hierarchy & Topological Validation: Perform multi-tier spatial topology checks and boundary alignment across State, District, Sub-District, and Ward polygons."
        p2 = "Territorial Compactness & Normalization: Compute geometric compactness indices (Polsby-Popper, Schwartzberg) and area denominators (st_area) for standardized thematic density mapping."
        p3 = "Cross-Sector Administrative Overlay: Serve as the primary spatial boundary scaffolding for aggregating socio-economic, agricultural, infrastructure, and demographic indicators."
        p4 = "Spatio-Temporal Boundary Evolution Modeling: Track administrative reorganizations, district bifurcations, and newly notified territorial units across official gazettes."
        p5 = "Jurisdictional Allocation & Planning Siting: Guide delineation of electoral constituencies, administrative police beats, planning zones, and public service catchments."

    how_to_use = (
        f"Geospatial Analysis & Modeling Protocol:\n"
        f"1. {p1}\n"
        f"2. {p2}\n"
        f"3. {p3}\n"
        f"4. {p4}\n"
        f"5. {p5}"
    )

    return domain, target_ind, how_to_use


def generate_field_analysis(domain, f_clean, base_target_ind):
    """Generate senior-level field analysis, target industries, and concise remarks."""
    f_lower = f_clean.lower()

    # Spatial Key / Identifier
    if f_lower in ['objectid', 'objectid_1', 'fid', 'id', 'pfafstetter id', 'gems_id_w', 'plaza code', 'st_code', 'dt_code', 'gid']:
        t_ind = "GIS & Spatial Data Infrastructure, IT & Spatial Database Services"
        p_an = "Primary spatial feature indexing (R-tree), unique record joins, and topological geometry validation"
        rem = "ESRI primary key / spatial index"
    # Administrative Codes
    elif any(k in f_lower for k in ['state code', 'district code', 'sub district code', 'village code', 'lgd', 'constituency code', 'ward code', 'lgd_dist', 'lgd_statec', 'prjcode']):
        t_ind = "E-Governance, Public Administration, Inter-Departmental Spatial Data Integration"
        p_an = "Standardized relational key joins with Local Government Directory (LGD/NIC) datasets and national administrative spatial layers"
        rem = "Official Government of India administrative code"
    # Census Codes
    elif 'census' in f_lower and 'code' in f_lower:
        t_ind = "Demographic Analytics, Socio-Economic Research, Policy Planning"
        p_an = "Spatial join with Census demographic polygons for population-at-risk normalization and socio-economic vulnerability indexing"
        rem = "Standard Census 2011 administrative identifier"
    # Names / Labels
    elif f_lower in ['name', 'state name', 'district name', 'country name', 'sub district name', 'village name', 'ward name', 'circle name', 'region name', 'division name', 'location', 'state', 'district', 'state_name', 'dist_name', 'subdistrict name']:
        t_ind = "State & Central Governance, Urban & Regional Planning, Cartography"
        p_an = "Thematic choropleth labeling, administrative boundary aggregation, and spatial dissolve operations"
        rem = "Official administrative boundary name"
    # Area & Perimeter
    elif 'area' in f_lower and any(k in f_lower for k in ['shape', 'sq', 'st_area', 'shape_area', 'total', 'percent_flooded']):
        t_ind = "GIS & Geospatial Analytics, Cartography, Regional Planning"
        p_an = "Spatial normalization denominator for computing thematic density per 1,000 km² and area coverage ratios"
        rem = "Projected polygon area in coordinate square units"
    elif 'perimeter' in f_lower or 'length' in f_lower or 'st_perimeter' in f_lower or 'shape_length' in f_lower:
        t_ind = "GIS & Geospatial Analytics, Boundary Mapping, Infrastructure Planning"
        p_an = "Boundary compactness analysis (Polsby-Popper / Schwartzberg index) to evaluate perimeter border effects"
        rem = "Projected polygon perimeter in coordinate linear units"
    # Geographic Coordinates
    elif f_lower in ['latitude', 'longitude', 'x', 'y', 'lat', 'lon', 'latitude of epicentre', 'longitude of epicentre']:
        t_ind = "GIS & GPS Fleet Navigation, Spatial Analytics, Cartography"
        p_an = "Geographic coordinate spatial referencing (WGS84 EPSG:4326), point-in-polygon spatial queries, and proximity buffering"
        rem = "Geographic point coordinate"
    # Thematic Domain-Specific Fields
    else:
        t_ind = base_target_ind
        if domain == 'road_safety':
            p_an = f"Spatial clustering (Getis-Ord Gi*) of {f_clean} to isolate high-risk roadway stretches, correlate with highway infrastructure defects, and prioritize safety interventions"
            rem = "Key road safety and crash outcome indicator"
        elif domain == 'transport_infrastructure':
            p_an = f"Corridor network analysis and throughput volume modeling for {f_clean} to detect logistics bottlenecks and optimize multi-modal freight asset placement"
            rem = "Transportation and logistics infrastructure metric"
        elif domain == 'rural_development':
            p_an = f"Habitation-level infrastructure saturation mapping and last-mile connectivity modeling for {f_clean} under national rural development programs"
            rem = "Rural development and civic infrastructure connectivity metric"
        elif domain == 'finance_economy':
            p_an = f"District-level spatial rate calculation and economic disparity mapping (Gini/Theil index) for {f_clean} to evaluate regional economic performance and capital deepening"
            rem = "Official economic performance and financial inclusion metric"
        elif domain == 'demographics_society':
            p_an = f"Spatial demographic rate normalization and vulnerability clustering for {f_clean} to guide targeted social welfare and affirmative action programs"
            rem = "Socio-demographic development and census indicator"
        elif domain == 'health_medicine':
            p_an = f"Epidemiological hotspot detection and spatial healthcare catchment accessibility modeling (2SFCA) for {f_clean} to identify healthcare deficit clusters"
            rem = "Public health and disease surveillance metric"
        elif domain == 'water_climate':
            p_an = f"Spatial hazard mapping and hydrological catchment vulnerability modeling for {f_clean} to guide water security and disaster mitigation planning"
            rem = "Hydrological and natural resource metric"
        elif domain == 'urban_industrial':
            p_an = f"Urban infrastructure density analysis and industrial corridor accessibility modeling for {f_clean} to optimize civic zoning and logistics routing"
            rem = "Urban planning and industrial development metric"
        elif domain == 'education':
            p_an = f"Spatial enrolment parity mapping and institutional catchment distance modeling for {f_clean} to optimize educational infrastructure allocation"
            rem = "Educational access and human capital metric"
        elif domain == 'elections_governance':
            p_an = f"Spatial voting pattern analysis and electoral swing-cluster detection for {f_clean} across parliamentary and assembly constituencies"
            rem = "Electoral outcome and voter participation metric"
        else:
            p_an = f"Spatial distribution modeling, administrative density normalization, and infrastructure overlay for {f_clean}"
            rem = "Spatial attribute metric"

    return t_ind, p_an, rem


def run_batch_update(start_sr, end_sr, workbook_path="exported_geojson_540andtif_catalog.xlsx"):
    print("=" * 85)
    print(f"BATCH UPDATE FOR LAYERS {start_sr} TO {end_sr}")
    print(f"Target Workbook: {workbook_path}")
    print("=" * 85)

    if not os.path.exists(workbook_path):
        print(f"[!] Error: Workbook '{workbook_path}' not found.")
        sys.exit(1)

    # 1. Close Excel if open to prevent PermissionError
    print("[*] Ensuring workbook is not locked by Excel...")
    close_running_excel()

    # 2. Index Word Docs
    print("[*] Scanning Description/*.docx files...")
    docs = load_all_docs("Description")
    print(f"[+] Loaded {len(docs)} Word description documents.")

    # 3. Load Workbook
    print(f"[*] Loading workbook '{workbook_path}' ...")
    wb = openpyxl.load_workbook(workbook_path)
    sheet_names = wb.sheetnames

    # Map Sr. No to sheet name (e.g. 101 -> '101_India_3A__Growth_of_Net')
    sr_to_sheet = {}
    for s in sheet_names:
        m = re.match(r'^(\d+)_', s)
        if m:
            sr_to_sheet[int(m.group(1))] = s

    cat_ws = wb["GeoJSON Catalog"]

    # Identify master columns
    col_map = {}
    for c in range(1, cat_ws.max_column + 1):
        v = str(cat_ws.cell(1, c).value or "").strip()
        if v:
            col_map[v] = c

    how_col = col_map.get("How to Use This File", 4)
    ind_col = col_map.get("Target Industries", 5)
    src_col = col_map.get("Source", 6)

    # -------------------------------------------------------------------------
    # PART 1: UPDATE MASTER SHEET ('GeoJSON Catalog')
    # -------------------------------------------------------------------------
    print("\n" + "=" * 50)
    print(f"[1/2] Updating Master Sheet Rows {start_sr} to {end_sr}")
    print("=" * 50)

    updated_master = 0
    layer_cache = {}

    for sr in range(start_sr, end_sr + 1):
        row_idx = sr + 1
        name = str(cat_ws.cell(row_idx, 2).value or "").strip()
        if not name:
            continue

        fields_raw = str(cat_ws.cell(row_idx, 3).value or "")
        fields = [f.strip() for f in fields_raw.split(',') if f.strip()]

        # Match to doc
        matched_doc = match_layer_to_doc(name, docs)
        doc_title = matched_doc['title'] if matched_doc else unescape_name(name)
        url = matched_doc['url'] if matched_doc else ""

        # Classify domain and generate protocols
        domain, target_ind, how_to_use = classify_domain_and_protocols(sr, name, doc_title)

        # Store in cache for child sheets
        layer_cache[sr] = {
            'name': name,
            'doc_title': doc_title,
            'url': url,
            'domain': domain,
            'target_ind': target_ind,
            'how_to_use': how_to_use,
            'fields': fields
        }

        # 1. Update How to Use This File (Column 4)
        c_how = cat_ws.cell(row=row_idx, column=how_col, value=how_to_use)
        c_how.font = FONT_DATA
        c_how.alignment = ALIGN_LEFT_TOP_WRAP
        cat_ws.row_dimensions[row_idx].height = 115

        # 2. Update Target Industries (Column 5)
        c_ind = cat_ws.cell(row=row_idx, column=ind_col, value=target_ind)
        c_ind.font = FONT_DATA
        c_ind.alignment = ALIGN_LEFT_TOP_WRAP

        # 3. Update Source URL (Column 6)
        if url:
            c_src = cat_ws.cell(row=row_idx, column=src_col, value=url)
            c_src.hyperlink = url
            c_src.font = FONT_LINK
            c_src.alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)

        updated_master += 1

    print(f"[+] Successfully updated {updated_master} rows in 'GeoJSON Catalog'.")

    # Set master column widths & AutoFilter
    master_widths = {
        "A": 10, "B": 42, "C": 55, "D": 58,
        "E": 35, "F": 50, "G": 20, "H": 20
    }
    for col_letter, w in master_widths.items():
        cat_ws.column_dimensions[col_letter].width = w

    cat_ws.auto_filter.ref = f"A1:H{cat_ws.max_row}"

    # -------------------------------------------------------------------------
    # PART 2: UPDATE CHILD SHEETS
    # -------------------------------------------------------------------------
    print("\n" + "=" * 50)
    print(f"[2/2] Updating Child Sheets ({start_sr} to {end_sr})")
    print("=" * 50)

    child_headers = [
        ("Sr. No", 10),
        ("Layer Available Fields", 45),
        ("Target Industries", 35),
        ("Potential Analysis", 55),
        ("Potential Analysis Manual", 35),  # STRICTLY VACANT FOR USER ENTRY
        ("Remark", 35),
    ]

    updated_sheets_count = 0
    total_fields_updated = 0

    for sr in range(start_sr, end_sr + 1):
        info = layer_cache.get(sr)
        if not info:
            continue

        sheet_name = sr_to_sheet.get(sr)
        if not sheet_name or sheet_name not in wb.sheetnames:
            print(f"[!] Warning: Sheet for Sr {sr} not found in workbook.")
            continue

        ws = wb[sheet_name]

        # 1. Update Row 3 Table Headers (ensure 6 columns)
        ws.row_dimensions[3].height = 26
        for col_idx, (hdr_text, col_w) in enumerate(child_headers, start=1):
            cell = ws.cell(row=3, column=col_idx, value=hdr_text)
            cell.font = FONT_HEADER
            cell.fill = FILL_HEADER
            cell.alignment = ALIGN_CENTER_CENTER
            cell.border = BORDER_HEADER_CELL
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = col_w

        # Clear any old headers in Col 7 or beyond
        for c_extra in range(7, ws.max_column + 1):
            cell_extra = ws.cell(row=3, column=c_extra)
            cell_extra.value = None
            cell_extra.fill = PatternFill(fill_type=None)
            cell_extra.border = Border()

        # 2. Update Data Rows (Row 4 onwards)
        max_r = ws.max_row
        for r in range(4, max_r + 1):
            f_name = str(ws.cell(r, 2).value or "").strip()
            if not f_name or f_name.lower() in ["none", ""]:
                continue

            fill_style = FILL_ZEBRA_ODD if (r % 2 == 0) else FILL_ZEBRA_EVEN
            t_ind, p_an, rem = generate_field_analysis(info['domain'], f_name, info['target_ind'])

            # Col 1: Sr. No
            c1 = ws.cell(row=r, column=1, value=r - 3)
            c1.font = FONT_DATA
            c1.alignment = ALIGN_CENTER_TOP
            c1.border = BORDER_DATA_CELL
            c1.fill = fill_style

            # Col 2: Field Name
            c2 = ws.cell(row=r, column=2, value=f_name)
            c2.font = FONT_DATA
            c2.alignment = ALIGN_LEFT_TOP_WRAP
            c2.border = BORDER_DATA_CELL
            c2.fill = fill_style

            # Col 3: Target Industries
            c3 = ws.cell(row=r, column=3, value=t_ind)
            c3.font = FONT_DATA
            c3.alignment = ALIGN_LEFT_TOP_WRAP
            c3.border = BORDER_DATA_CELL
            c3.fill = fill_style

            # Col 4: Potential Analysis (Senior Geospatial Analyst Depth)
            c4 = ws.cell(row=r, column=4, value=p_an)
            c4.font = FONT_DATA
            c4.alignment = ALIGN_LEFT_TOP_WRAP
            c4.border = BORDER_DATA_CELL
            c4.fill = fill_style

            # Col 5: Potential Analysis Manual (STRICTLY VACANT FOR USER MANUAL ENTRY)
            c5 = ws.cell(row=r, column=5)
            c5.value = None
            c5.font = FONT_DATA
            c5.alignment = ALIGN_LEFT_TOP_WRAP
            c5.border = BORDER_DATA_CELL
            c5.fill = fill_style

            # Col 6: Remark (Clear context note)
            c6 = ws.cell(row=r, column=6, value=rem)
            c6.font = FONT_DATA
            c6.alignment = ALIGN_LEFT_TOP_WRAP
            c6.border = BORDER_DATA_CELL
            c6.fill = fill_style

            # Clear extra columns if any
            for c_extra in range(7, ws.max_column + 1):
                c_ex = ws.cell(row=r, column=c_extra)
                c_ex.value = None

            ws.row_dimensions[r].height = 48
            total_fields_updated += 1

        # AutoFilter
        ws.auto_filter.ref = f"A3:F{max_r}"
        ws.freeze_panes = "A4"

        # A1 Back Link
        back_cell = ws.cell(row=1, column=1)
        back_cell.value = "← Back to Catalog"
        back_cell.font = FONT_LINK
        back_cell.fill = PatternFill(start_color=COLOR_BANNER_BG, end_color=COLOR_BANNER_BG, fill_type="solid")
        back_cell.alignment = ALIGN_CENTER_CENTER
        back_cell.border = BORDER_DATA_CELL
        back_cell.hyperlink = f"#'GeoJSON Catalog'!A{sr + 1}"

        updated_sheets_count += 1
        if updated_sheets_count % 10 == 0 or updated_sheets_count == (end_sr - start_sr + 1):
            print(f"    -> Processed {updated_sheets_count}/{(end_sr - start_sr + 1)} child sheets ({total_fields_updated} fields)...")

    # -------------------------------------------------------------------------
    # PART 3: SAVE WORKBOOK
    # -------------------------------------------------------------------------
    print("\n" + "=" * 50)
    print(f"[*] Saving updated workbook to '{workbook_path}' ...")
    wb.save(workbook_path)
    print(f"[+] Successfully saved '{workbook_path}'!")
    print("=" * 85)
    print("BATCH UPDATE COMPLETE:")
    print(f"  - Master Rows Updated:   {updated_master} (Rows {start_sr + 1} to {end_sr + 1})")
    print(f"  - Child Sheets Updated:  {updated_sheets_count} (Sheets {start_sr} to {end_sr})")
    print(f"  - Total Fields Updated:  {total_fields_updated}")
    print(f"  - 'Potential Analysis Manual' Column Added & Kept Vacant Across All {updated_sheets_count} Sheets")
    print("=" * 85)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Universal Batch Update for Master Catalog & Child Sheets")
    parser.add_argument("--start", type=int, default=101, help="Starting Sr. No (e.g. 101)")
    parser.add_argument("--end", type=int, default=200, help="Ending Sr. No (e.g. 200)")
    parser.add_argument("--workbook", type=str, default="exported_geojson_540andtif_catalog.xlsx", help="Workbook filename")
    args = parser.parse_args()

    run_batch_update(args.start, args.end, args.workbook)
