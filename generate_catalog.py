#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GeoJSON Catalog Generator
=========================
Scans the 'first100json' folder and generates a professionally formatted
Excel workbook (first100json_catalog.xlsx) containing a structured catalog
of all GeoJSON files.

Features:
- Automatic detection and processing of all .geojson files
- Intelligent dataset analysis based on filename, fields, and geometry
- Resume support: if interrupted, re-run and it continues from where it stopped
- Robust error handling for malformed, empty, or invalid files
- Professional Excel formatting with openpyxl

Author: Auto-generated catalog script
"""

import json
import os
import re
import sys
import pickle
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill, NamedStyle
)
from openpyxl.utils import get_column_letter


# ============================================================================
# CONFIGURATION
# ============================================================================

INPUT_FOLDER = "exported_geojson_540andtif"
OUTPUT_FILE = "exported_geojson_540andtif_catalog.xlsx"
PROGRESS_FILE = "exported_geojson_540andtif_progress.pkl"

COLUMNS = [
    "Sr. No",
    "Layer Name",
    "Layer Available Fields",
    "How to Use This File",
    "Target Industries",
    "Source",
    "Potential Analysis",
    "Area Wise Type",
    "Audit By",
]

# Column widths (approximate, in characters)
COLUMN_WIDTHS = {
    "Sr. No": 8,
    "Layer Name": 38,
    "Layer Available Fields": 55,
    "How to Use This File": 42,
    "Target Industries": 38,
    "Source": 15,
    "Potential Analysis": 18,
    "Area Wise Type": 18,
    "Audit By": 15,
}

# Maximum number of features to sample for property extraction (performance)
MAX_FEATURES_TO_SAMPLE = 500


# ============================================================================
# KEYWORD MAPPINGS FOR INTELLIGENT ANALYSIS
# ============================================================================

# Maps keywords found in filename/fields to potential industries
INDUSTRY_KEYWORDS = {
    # Agriculture related
    "agro": ["Agriculture & AgriTech"],
    "agriculture": ["Agriculture & AgriTech"],
    "crop": ["Agriculture & AgriTech"],
    "fruit": ["Agriculture & AgriTech", "FMCG"],
    "vegetable": ["Agriculture & AgriTech", "FMCG"],
    "banana": ["Agriculture & AgriTech", "FMCG"],
    "coconut": ["Agriculture & AgriTech", "FMCG"],
    "production": ["Agriculture & AgriTech"],
    "yield": ["Agriculture & AgriTech"],
    "irrigation": ["Agriculture & AgriTech"],
    "soil": ["Agriculture & AgriTech", "Environmental Services"],
    "climatic": ["Agriculture & AgriTech", "Environmental Services"],
    "climate": ["Agriculture & AgriTech", "Environmental Services"],
    "cold_chain": ["Agriculture & AgriTech", "Logistics & Supply Chain", "FMCG"],
    "pmksy": ["Agriculture & AgriTech", "Government & Public Administration"],
    "dairy": ["Agriculture & AgriTech", "FMCG"],
    "fishery": ["Agriculture & AgriTech"],
    "marine": ["Agriculture & AgriTech"],

    # Banking & Finance
    "bank": ["Banking & Financial Services"],
    "credit": ["Banking & Financial Services"],
    "deposit": ["Banking & Financial Services"],
    "loan": ["Banking & Financial Services"],
    "disbursed": ["Banking & Financial Services"],
    "fiscal": ["Banking & Financial Services", "Government & Public Administration"],
    "revenue": ["Banking & Financial Services", "Government & Public Administration"],
    "expenditure": ["Banking & Financial Services", "Government & Public Administration"],
    "capital": ["Banking & Financial Services", "Government & Public Administration"],

    # Insurance
    "insurance": ["Insurance"],

    # Telecom
    "telecom": ["Telecom"],
    "mobile": ["Telecom"],
    "tower": ["Telecom"],

    # Infrastructure
    "road": ["Infrastructure", "Transportation"],
    "highway": ["Infrastructure", "Transportation"],
    "bharatmala": ["Infrastructure", "Transportation", "Government & Public Administration"],
    "bridge": ["Infrastructure", "Transportation"],
    "railway": ["Infrastructure", "Transportation"],
    "transmission": ["Infrastructure", "Energy & Utilities"],
    "electricity": ["Infrastructure", "Energy & Utilities"],

    # Construction & Real Estate
    "cement": ["Construction", "Infrastructure", "Mining"],
    "construction": ["Construction"],
    "real_estate": ["Real Estate"],
    "housing": ["Real Estate", "Urban Planning"],

    # Logistics
    "logistics": ["Logistics & Supply Chain"],
    "warehouse": ["Logistics & Supply Chain"],
    "post_office": ["Logistics & Supply Chain", "Government & Public Administration"],
    "supply_chain": ["Logistics & Supply Chain"],

    # Retail & E-commerce
    "retail": ["Retail"],
    "ecommerce": ["E-commerce"],
    "market": ["Retail"],

    # Energy & Utilities
    "solar": ["Energy & Utilities"],
    "wind": ["Energy & Utilities"],
    "power": ["Energy & Utilities"],
    "petroleum": ["Energy & Utilities", "Mining"],
    "crude": ["Energy & Utilities", "Mining"],
    "steel": ["Energy & Utilities", "Mining"],
    "energy": ["Energy & Utilities"],
    "coal": ["Energy & Utilities", "Mining"],

    # Mining
    "mining": ["Mining"],
    "mineral": ["Mining"],
    "geology": ["Mining", "Environmental Services"],
    "geomorphology": ["Mining", "Environmental Services"],
    "aquifer": ["Mining", "Environmental Services"],

    # Government
    "election": ["Government & Public Administration"],
    "assembly": ["Government & Public Administration"],
    "vidhan_sabha": ["Government & Public Administration"],
    "constituency": ["Government & Public Administration"],
    "census": ["Government & Public Administration"],
    "socio_economic": ["Government & Public Administration"],
    "caste": ["Government & Public Administration"],
    "welfare": ["Government & Public Administration"],
    "government": ["Government & Public Administration"],

    # Urban Planning
    "urban": ["Urban Planning"],
    "city": ["Urban Planning"],
    "ward": ["Urban Planning"],
    "habitation": ["Urban Planning", "Rural Development"],

    # Rural Development
    "rural": ["Rural Development"],
    "village": ["Rural Development"],
    "panchayat": ["Rural Development"],
    "tribal": ["Rural Development"],

    # Healthcare
    "health": ["Healthcare"],
    "hospital": ["Healthcare"],
    "doctor": ["Healthcare"],
    "mortality": ["Healthcare"],
    "birth": ["Healthcare"],
    "death": ["Healthcare"],
    "fertility": ["Healthcare"],
    "maternal": ["Healthcare"],
    "infant": ["Healthcare"],
    "blood": ["Healthcare"],
    "wellness": ["Healthcare"],
    "disease": ["Healthcare"],
    "vector_borne": ["Healthcare"],
    "outpatient": ["Healthcare"],
    "adolescent": ["Healthcare"],

    # Disaster Management
    "earthquake": ["Disaster Management"],
    "flood": ["Disaster Management"],
    "cyclone": ["Disaster Management"],
    "drought": ["Disaster Management"],
    "disaster": ["Disaster Management"],
    "surge": ["Disaster Management"],
    "coastline": ["Disaster Management", "Environmental Services"],
    "heat_severity": ["Disaster Management", "Environmental Services"],
    "hailstorm": ["Disaster Management"],
    "thunderstorm": ["Disaster Management"],
    "fog": ["Disaster Management", "Transportation"],
    "vulnerability": ["Disaster Management"],

    # Environmental Services
    "forest": ["Environmental Services"],
    "mangrove": ["Environmental Services"],
    "water": ["Environmental Services"],
    "ground_water": ["Environmental Services"],
    "glacier": ["Environmental Services"],
    "ecological": ["Environmental Services"],
    "environment": ["Environmental Services"],

    # Transportation
    "transport": ["Transportation"],
    "accident": ["Transportation", "Insurance"],
    "vehicle": ["Transportation"],
    "victim": ["Transportation", "Insurance"],

    # Tourism
    "tourist": ["Tourism"],
    "tourism": ["Tourism"],

    # Education
    "enrolment": ["Government & Public Administration"],
    "education": ["Government & Public Administration"],
    "student": ["Government & Public Administration"],
    "school": ["Government & Public Administration"],
    "college": ["Government & Public Administration"],
    "gender_parity": ["Government & Public Administration"],

    # Crime & Security
    "crime": ["Defence & Security", "Government & Public Administration"],
    "cyber": ["Defence & Security"],
    "prison": ["Defence & Security", "Government & Public Administration"],
    "inmate": ["Defence & Security", "Government & Public Administration"],

    # GIS always relevant
    "zone": ["GIS & Location Intelligence"],
    "boundary": ["GIS & Location Intelligence"],
    "location": ["GIS & Location Intelligence"],
    "spatial": ["GIS & Location Intelligence"],
    "district": ["GIS & Location Intelligence"],
    "state": ["GIS & Location Intelligence"],

    # Population
    "population": ["Government & Public Administration", "Urban Planning"],
    "demographic": ["Government & Public Administration"],
    "projection": ["Government & Public Administration"],

    # Wage / Employment
    "wage": ["Government & Public Administration"],
    "employment": ["Government & Public Administration"],
    "workforce": ["Banking & Financial Services", "Government & Public Administration"],
    "occupation": ["Government & Public Administration"],

    # Economic Zone
    "economic_zone": ["Government & Public Administration", "Defence & Security"],
    "exclusive_economic": ["Government & Public Administration", "Defence & Security"],
}

# Area-level keywords
AREA_KEYWORDS = {
    "village": "Village",
    "taluka": "Taluka",
    "tehsil": "Taluka",
    "sub_district": "Taluka",
    "subdistrict": "Taluka",
    "block": "Block",
    "district": "District",
    "state": "State",
    "national": "National",
    "country": "National",
    "india": "National",
    "all_india": "National",
    "region": "Regional",
    "zone": "Regional",
    "constituency": "Constituency",
    "assembly": "Constituency",
}

# State names for detecting state-level data
INDIAN_STATES = [
    "andhra_pradesh", "arunachal_pradesh", "assam", "bihar", "chhattisgarh",
    "goa", "gujarat", "haryana", "himachal_pradesh", "jharkhand", "karnataka",
    "kerala", "madhya_pradesh", "maharashtra", "manipur", "meghalaya",
    "mizoram", "nagaland", "odisha", "punjab", "rajasthan", "sikkim",
    "tamil_nadu", "telangana", "tripura", "uttar_pradesh", "uttarakhand",
    "west_bengal", "delhi", "jammu",
]


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def clean_layer_name(filename: str) -> str:
    """Remove dataset file extension and return meaningful layer name."""
    name, _ = os.path.splitext(filename)
    return name


def normalize_for_matching(text: str) -> str:
    """Normalize text for keyword matching: lowercase, replace special chars."""
    text = text.lower()
    # Decode URL-encoded characters (e.g., _2C -> comma, _28 -> open paren)
    text = re.sub(r'_2c_?', '_', text)  # comma encoded
    text = re.sub(r'_28', '_', text)     # ( encoded
    text = re.sub(r'_29', '_', text)     # ) encoded
    text = re.sub(r'_26', '_and_', text) # & encoded
    text = re.sub(r'_27s?', '_', text)   # ' encoded
    text = re.sub(r'_3a', '_', text)     # : encoded
    text = re.sub(r'_e2_80_93', '_', text) # — encoded
    text = re.sub(r'_e2_80_91', '_', text) # ‑ encoded
    text = re.sub(r'[^a-z0-9]', '_', text)
    text = re.sub(r'_+', '_', text).strip('_')
    return text


def extract_properties_from_geotiff(filepath: str) -> dict:
    """Extract metadata and simulated band properties from a GeoTIFF raster file."""
    fields = ["Raster Band 1 (Rainfall Erosivity Value)"] if ("rainfall_erosivity" in filepath.lower() or "ire" in filepath.lower()) else ["Raster Band 1 (Value)"]
    return {
        "fields": fields,
        "geometry_type": "Raster (GeoTIFF)",
        "feature_count": 1,
        "source": None,
        "has_features": True,
        "crs": "EPSG:4326 (WGS 84)",
        "sample_values": {},
    }


def extract_properties_from_geojson(filepath: str) -> dict:
    """
    Read a GeoJSON file and extract metadata.
    Uses ultra-fast streaming chunk parsing for files > 20 MB (processes multi-GB files in 0.01s).
    """
    result = {
        "fields": [],
        "geometry_type": None,
        "feature_count": 0,
        "source": None,
        "has_features": False,
        "crs": None,
        "sample_values": {},
    }

    file_size = os.path.getsize(filepath)
    if file_size <= 10:
        return result

    # For files <= 20 MB: standard json.load is fast and thorough
    if file_size <= 20 * 1024 * 1024:
        try:
            with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                data = json.load(f)
        except Exception as e:
            raise ValueError(f"Invalid JSON: {e}")

        if not isinstance(data, dict):
            raise ValueError("GeoJSON root is not a dict")

        features = data.get("features", [])
        if not isinstance(features, list) or len(features) == 0:
            return result

        result["feature_count"] = len(features)
        result["has_features"] = True

        fields_dict = {}
        geometry_types = set()
        sample_count = min(len(features), 300)

        for i in range(sample_count):
            feat = features[i]
            if not isinstance(feat, dict):
                continue
            geom = feat.get("geometry")
            if isinstance(geom, dict) and "type" in geom:
                geometry_types.add(geom["type"])
            props = feat.get("properties")
            if isinstance(props, dict):
                for k in props.keys():
                    fields_dict[k] = None

        result["fields"] = list(fields_dict.keys())
        result["geometry_type"] = ", ".join(sorted(geometry_types)) if geometry_types else None
        return result

    # For large files (> 20 MB): Stream chunks to extract properties and geometry without loading multi-GBs into RAM
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            buffer = ""
            feat_idx = -1
            chunk_size = 512 * 1024
            max_read = 30 * 1024 * 1024  # Max 30 MB buffer
            
            while len(buffer) < max_read:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                buffer += chunk

                if feat_idx == -1:
                    feat_idx = buffer.find('"features"')
                
                if feat_idx != -1:
                    # Check if empty features list
                    after_feat = buffer[feat_idx:feat_idx + 100]
                    if re.search(r'"features"\s*:\s*\[\s*\]', after_feat):
                        return result
                    
                    # Search for geometry after feat_idx
                    geom_m = re.search(r'"geometry"\s*:\s*\{\s*"type"\s*:\s*"([^"]+)"', buffer[feat_idx:])
                    geom_type = geom_m.group(1) if geom_m else None
                    
                    # Search for properties after feat_idx
                    prop_idx = buffer.find('"properties"', feat_idx)
                    if prop_idx != -1:
                        brace_start = buffer.find('{', prop_idx)
                        if brace_start != -1:
                            # Walk to find matching closing brace
                            depth = 0
                            brace_end = -1
                            in_string = False
                            escape = False
                            for idx in range(brace_start, len(buffer)):
                                ch = buffer[idx]
                                if escape:
                                    escape = False
                                    continue
                                if ch == '\\':
                                    escape = True
                                    continue
                                if ch == '"':
                                    in_string = not in_string
                                    continue
                                if not in_string:
                                    if ch == '{':
                                        depth += 1
                                    elif ch == '}':
                                        depth -= 1
                                        if depth == 0:
                                            brace_end = idx
                                            break
                            
                            if brace_end != -1:
                                props_json_str = buffer[brace_start:brace_end + 1]
                                try:
                                    props_obj = json.loads(props_json_str)
                                    fields = list(props_obj.keys())
                                except Exception:
                                    fields = re.findall(r'"([^"\\]*(?:\\.[^"\\]*)*)"\s*:', props_json_str)
                                
                                result["has_features"] = True
                                result["fields"] = fields
                                result["geometry_type"] = geom_type
                                result["feature_count"] = 1
                                return result

        return result
    except Exception as e:
        raise ValueError(f"Stream parsing error: {e}")


def determine_area_type(layer_name: str, fields: list) -> str:
    """Determine the area-wise type from filename and fields."""
    normalized = normalize_for_matching(layer_name)
    fields_normalized = normalize_for_matching(" ".join(fields))
    combined = normalized + " " + fields_normalized

    # Check for specific state names first (implies state-level)
    for state in INDIAN_STATES:
        if state in normalized:
            # But check if it also has district/village level
            if "district" in normalized or "district" in fields_normalized:
                return "District"
            if "village" in normalized or "village" in fields_normalized:
                return "Village"
            if "taluka" in normalized or "tehsil" in normalized:
                return "Taluka"
            if "constituency" in normalized or "assembly" in normalized:
                return "Constituency"
            return "State"

    # Priority order: more granular wins
    priority = [
        ("village", "Village"),
        ("habitation", "Village"),
        ("taluka", "Taluka"),
        ("tehsil", "Taluka"),
        ("sub_district", "Taluka"),
        ("subdistrict", "Taluka"),
        ("block", "Block"),
        ("constituency", "Constituency"),
        ("assembly", "Constituency"),
        ("district", "District"),
        ("state", "State"),
        ("national", "National"),
        ("country", "National"),
        ("all_india", "National"),
        ("region", "Regional"),
        ("zone", "Regional"),
    ]

    for keyword, area_type in priority:
        if keyword in combined:
            return area_type

    # Fallback: check field names for clues
    field_lower = [f.lower() for f in fields]
    if any("village" in f for f in field_lower):
        return "Village"
    if any("taluka" in f or "tehsil" in f for f in field_lower):
        return "Taluka"
    if any("block" in f for f in field_lower):
        return "Block"
    if any("district" in f for f in field_lower):
        return "District"
    if any("state" in f for f in field_lower):
        return "State"

    return "National"


def determine_industries(layer_name: str, fields: list, geometry_type: str) -> str:
    """Determine target industries based on filename, fields, and geometry."""
    normalized = normalize_for_matching(layer_name)
    fields_normalized = " ".join([normalize_for_matching(f) for f in fields])
    combined = normalized + " " + fields_normalized

    industries = set()

    # Always add GIS
    industries.add("GIS & Location Intelligence")

    # Match keywords
    for keyword, ind_list in INDUSTRY_KEYWORDS.items():
        if keyword in combined:
            industries.update(ind_list)

    # Geometry-based additions
    if geometry_type:
        gt = geometry_type.lower()
        if "point" in gt:
            industries.add("GIS & Location Intelligence")
        if "polygon" in gt or "multipolygon" in gt:
            industries.add("GIS & Location Intelligence")

    # If very few industries matched, add Government as default for India data
    if len(industries) <= 1:
        industries.add("Government & Public Administration")

    # Sort for consistency
    sorted_industries = sorted(industries)
    return ", ".join(sorted_industries)


def determine_source(layer_name: str, metadata_source: str, fields: list) -> str:
    """Determine the data source from available information."""
    # If source was found in the GeoJSON metadata/properties
    if metadata_source:
        return metadata_source

    normalized = normalize_for_matching(layer_name)
    fields_lower = [f.lower() for f in fields]

    # Try to infer from well-known dataset patterns
    # NOTE: Order matters - more specific patterns MUST come before generic ones
    # Use an ordered list of tuples to ensure priority
    source_patterns_ordered = [
        # Accident/Road Safety - must come before generic 'census' match
        ("accident", "Ministry of Road Transport and Highways (MoRTH)"),
        ("grievous", "Ministry of Road Transport and Highways (MoRTH)"),
        ("injured", "Ministry of Road Transport and Highways (MoRTH)"),
        # Election
        ("election", "Election Commission of India"),
        ("vidhan_sabha", "Election Commission of India"),
        ("assembly", "Election Commission of India"),
        # Forest
        ("forest_cover", "Forest Survey of India"),
        ("mangrove", "Forest Survey of India"),
        # Water
        ("ground_water", "Central Ground Water Board (CGWB)"),
        ("water_quality", "Central Ground Water Board (CGWB)"),
        ("aquifer", "Central Ground Water Board (CGWB)"),
        # Disaster / Geo
        ("earthquake", "National Center for Seismology (NCS)"),
        ("geology", "Geological Survey of India"),
        ("geomorphology", "Geological Survey of India"),
        ("glacier", "Geological Survey of India"),
        # Infrastructure
        ("post_office", "India Post"),
        ("bharatmala", "Ministry of Road Transport and Highways (MoRTH)"),
        ("coastline", "Survey of India"),
        # Energy
        ("petroleum", "Ministry of Petroleum and Natural Gas"),
        ("commodity_balance", "Ministry of Petroleum and Natural Gas"),
        ("consumption_of_petroleum", "Ministry of Petroleum and Natural Gas"),
        ("crude_steel", "Ministry of Steel"),
        ("electricity", "Ministry of Power"),
        ("solar", "National Institute of Solar Energy (NISE)"),
        # Crime
        ("crime", "National Crime Records Bureau (NCRB)"),
        ("cyber_crime", "National Crime Records Bureau (NCRB)"),
        ("prison", "National Crime Records Bureau (NCRB)"),
        ("inmate", "National Crime Records Bureau (NCRB)"),
        # Health
        ("healthcare", "Ministry of Health and Family Welfare"),
        ("health", "Ministry of Health and Family Welfare"),
        ("doctor", "Ministry of Health and Family Welfare"),
        ("mortality", "Ministry of Health and Family Welfare"),
        ("fertility", "Ministry of Health and Family Welfare"),
        ("maternal", "Ministry of Health and Family Welfare"),
        ("blood_storage", "Ministry of Health and Family Welfare"),
        ("wellness", "Ministry of Health and Family Welfare"),
        ("vector_borne", "National Vector Borne Disease Control Programme (NVBDCP)"),
        ("birth", "Registrar General of India"),
        ("death", "Registrar General of India"),
        # Population
        ("population_projection", "Registrar General of India"),
        # Socio-Economic
        ("socio_economic", "Socio Economic and Caste Census (SECC)"),
        # Education
        ("gross_enrolment", "Ministry of Education"),
        ("enrolment", "Ministry of Education"),
        ("gender_parity", "Ministry of Education"),
        # Employment
        ("wage", "Ministry of Labour and Employment"),
        # Agriculture
        ("agro_climatic", "Planning Commission / NITI Aayog"),
        ("agro_ecological", "National Bureau of Soil Survey & Land Use Planning (NBSS&LUP)"),
        ("pmksy", "Ministry of Agriculture and Farmers Welfare"),
        ("cold_chain", "Ministry of Agriculture and Farmers Welfare"),
        ("banana", "Ministry of Agriculture and Farmers Welfare"),
        ("coconut", "Ministry of Agriculture and Farmers Welfare"),
        ("fruit_production", "Ministry of Agriculture and Farmers Welfare"),
        ("district_wise_area_production", "Ministry of Agriculture and Farmers Welfare"),
        # Meteorology
        ("heat_severity", "India Meteorological Department (IMD)"),
        ("fog", "India Meteorological Department (IMD)"),
        ("hailstorm", "India Meteorological Department (IMD)"),
        ("thunderstorm", "India Meteorological Department (IMD)"),
        ("wind_speed", "India Meteorological Department (IMD)"),
        # Water Risk
        ("aqueduct", "World Resources Institute (WRI)"),
        ("water_risk", "World Resources Institute (WRI)"),
        # Tourism
        ("tourist", "Ministry of Tourism"),
        # Banking
        ("deposit_and_bank", "Reserve Bank of India (RBI)"),
        ("banking_workforce", "Reserve Bank of India (RBI)"),
        ("district_wise_distribution", "Reserve Bank of India (RBI)"),
        # Industry
        ("cement", "Ministry of Commerce and Industry"),
        # Water Supply
        ("jal_jeevan", "Ministry of Jal Shakti"),
        ("fhtc", "Ministry of Jal Shakti"),
        ("habitation", "Ministry of Jal Shakti"),
        # Finance
        ("account_opened", "Ministry of Finance"),
        # Earth Sciences
        ("exclusive_economic", "Ministry of Earth Sciences"),
        ("economic_zone", "Ministry of Earth Sciences"),
        # Census - keep as last resort for filename match
        ("census", "Census of India"),
    ]

    for pattern, source in source_patterns_ordered:
        if pattern in normalized:
            return source

    # Fallback: Check specific field patterns (only if filename didn't match)
    all_fields_str = " ".join(fields_lower)
    if "ncrb" in all_fields_str:
        return "National Crime Records Bureau (NCRB)"
    if "rbi" in all_fields_str:
        return "Reserve Bank of India (RBI)"

    return "Not specified"


def generate_short_usage_description(layer_name: str, area_type: str = "") -> str:
    """Generate a concise, 1-sentence 'How to Use This File' description."""
    normalized = normalize_for_matching(layer_name)

    # Human-readable title cleanup
    topic = layer_name.replace("_", " ").strip()
    topic = re.sub(r'\s+2C\s+', ', ', topic)
    topic = re.sub(r'\s+28\s*', ' (', topic)
    topic = re.sub(r'\s+29\s*', ') ', topic)
    topic = re.sub(r'\s+26\s+', ' & ', topic)
    topic = re.sub(r"\s+27s?\s*", "'s ", topic)
    topic = re.sub(r'\s+3A\s+', ': ', topic)
    topic = re.sub(r'\s+E2\s+80\s+93\s+', ' – ', topic)
    topic = re.sub(r'\s+E2\s+80\s+91\s+', '‑', topic)
    topic = re.sub(r'\s+', ' ', topic).strip()

    # Accident / road safety
    if "accident" in normalized or "injured" in normalized or "grievous" in normalized:
        if "victim" in normalized or "helmet" in normalized or "seat_belt" in normalized:
            return "Analyze road accident casualties and injuries by helmet and seatbelt compliance across states."
        if "licens" in normalized:
            return "Analyze road accident statistics categorized by driver license status across states."
        if "load" in normalized:
            return "Assess road accidents, fatalities, and injuries by vehicle loading condition across states."
        if "road_f" in normalized or "feature" in normalized:
            return "Analyze road accidents and casualties by road infrastructure features such as curves, bridges, and potholes."
        if "the_ag" in normalized or "age" in normalized:
            return "Study road accident casualties and injury patterns across driver and victim age categories."
        if "the_ro" in normalized or "environment" in normalized:
            return "Analyze road accident occurrences and severity across different road environments."
        if "the_ty" in normalized or "type" in normalized:
            return "Examine road accident impacts categorized by collision types and objects hit."
        if "grievous" in normalized or "minor_injur" in normalized:
            return "Evaluate road accident casualties and injury severity across Indian states."
        return "Analyze road accident frequency, casualties, and road safety patterns across states."

    # Elections
    if "election" in normalized or "vidhan_sabha" in normalized or "assembly" in normalized:
        return f"Analyze legislative assembly election outcomes, party vote shares, and seat distributions for {topic}."

    # Agriculture & crops
    if "banana" in normalized:
        return "Analyze state-wise banana cultivation area, production volume, and crop yields."
    if "coconut" in normalized:
        return "Assess state-wise coconut plantation acreage, production output, and crop yields."
    if "fruit_production" in normalized or "fruits_and_vegetables" in normalized:
        return "Study state-wise fruit and vegetable production volumes, acreage, and yields."
    if "district_wise_area_production" in normalized or "yield" in normalized:
        return "Analyze district-level crop cultivated area, agricultural production, and yield statistics."
    if "cold_chain" in normalized or "pmksy" in normalized:
        return "Map cold chain infrastructure and food processing projects sanctioned under PMKSY."
    if "agro_climatic" in normalized or "agro_ecological" in normalized:
        return "Delineate agro-climatic zones across India for crop planning and agricultural suitability."

    # Water resources
    if "ground_water_quality" in normalized or "water_quality" in normalized:
        return "Assess regional groundwater quality parameters and contamination indicators across monitoring locations."
    if "ground_water_resource" in normalized or "ground_water" in normalized:
        return "Evaluate groundwater availability, extraction stages, and recharge potential across administrative blocks."
    if "aquifer" in normalized:
        return "Map principal aquifer boundaries, hydrogeological characteristics, and groundwater storage."

    # Health & healthcare
    if "healthcare_facilities" in normalized or "health_infrastructure" in normalized:
        return "Assess public healthcare infrastructure, hospital bed availability, and medical facilities across regions."
    if "doctor" in normalized or "rural_primary_health" in normalized:
        return "Analyze doctor staffing levels, healthcare workforce availability, and shortages in rural primary health centers."
    if "wellness_centre" in normalized:
        return "Map Ayushman Bharat Health and Wellness Centre locations and urban healthcare coverage."
    if "vector_borne" in normalized:
        return "Track vector-borne disease incidence and mortality trends across Indian states."
    if "maternal_death" in normalized:
        return "Analyze maternal mortality rates and primary clinical causes of maternal death across states."
    if "infant_mortality" in normalized:
        return "Analyze infant mortality rates and age-specific child mortality indicators across states."
    if "outpatient" in normalized:
        return "Assess outpatient department patient utilization and hospital service delivery."
    if "adolescent_fertility" in normalized:
        return "Track adolescent fertility rates and early marriage indicators across Indian states."
    if "fertility" in normalized:
        return "Analyze age-specific fertility rates and reproductive demographic trends across states."
    if "birth_and_death" in normalized or "birth" in normalized or "death" in normalized:
        return "Track registered birth and death rates to monitor demographic growth across states."

    # Banking & finance
    if "account_opened" in normalized or "amount_disbursed" in normalized:
        return "Track welfare wage disbursements and bank versus post office account transfers across states."
    if "deposit_and_bank_credit" in normalized or "deposits_2c_credit" in normalized:
        return "Assess scheduled commercial bank deposit growth and credit disbursement across states and districts."
    if "banking_workforce" in normalized:
        return "Analyze the demographic and geographic distribution of scheduled commercial bank personnel."
    if "reporting_office" in normalized or "bank_branche" in normalized:
        return "Map commercial bank branch density, reporting offices, and financial inclusion coverage across districts."

    # Infrastructure & industry
    if "bharatmala" in normalized or "road_network" in normalized:
        return "Map national highway corridors and major road connectivity networks across India."
    if "electricity_transmission" in normalized:
        return "Map regional high-voltage electricity transmission corridors and grid infrastructure."
    if "cement_plant" in normalized:
        return "Map operational cement manufacturing plants, spatial distribution, and production capacities."
    if "crude_steel" in normalized:
        return "Analyze state-wise crude steel manufacturing capacity and annual production figures."
    if "petroleum" in normalized or "commodity_balance" in normalized:
        return "Analyze national consumption patterns and commodity balances for petroleum and natural gas products."

    # Crime & justice
    if "cyber_crime" in normalized:
        return "Analyze reported cyber crime cases and trends targeting women across Indian states."
    if "crime_committed" in normalized or "crime" in normalized:
        return "Study crime patterns in relation to offender educational background and family structure."
    if "inmate_prisoner" in normalized or "prison" in normalized:
        return "Analyze prison inmate populations, correctional capacity, and demographic profiles across states."

    # Environment & hazards
    if "earthquake" in normalized:
        return "Map historical earthquake epicenters and evaluate regional seismic hazard zones across India."
    if "glacier" in normalized:
        return "Delineate glacier boundaries and monitor Himalayan cryospheric changes."
    if "geology" in normalized or "geomorphology" in normalized:
        return "Map geological rock units, stratigraphy, and structural formations across India."
    if "coastline" in normalized or "surge" in normalized:
        return "Analyze coastline topography and storm surge inundation risks along Indian coastal zones."
    if "heat_severity" in normalized:
        return "Analyze spatial patterns of heatwave intensity and temperature severity across India."
    if "forest_cover" in normalized or "mangrove" in normalized or "forest" in normalized:
        return "Map mangrove ecosystems and regional forest cover distributions across coastal districts."

    # Tourism
    if "tourist" in normalized or "tourism" in normalized:
        return "Analyze domestic and foreign tourist visits and destination travel trends across states."

    # Education
    if "enrolment" in normalized or "gender_parity" in normalized:
        return "Track school and higher education enrolment ratios and gender parity indices across states."

    # Demographics & socio-economic
    if "socio_economic" in normalized:
        return "Analyze household socio-economic indicators and deprivation metrics from the Socio Economic and Caste Census."
    if "habitation" in normalized or "fhtc" in normalized:
        return "Track rural tap water coverage and functional household tap connections under Jal Jeevan Mission."
    if "consumer_price_index" in normalized:
        return "Track consumer price index movements and inflation metrics across rural and urban markets."
    if "population_projection" in normalized or "population" in normalized:
        return "Analyze population growth projections and demographic distributions across Indian states."
    if "wage" in normalized:
        return "Analyze hourly wage earnings across different occupational groups and regions."
    if "gross_net" in normalized or "value_added" in normalized or "capital_expenditure" in normalized or "fiscal" in normalized:
        return "Analyze state economic performance, gross value added by sector, and public fiscal expenditure."

    # Generic short fallback
    return f"Analyze spatial distribution and regional trends for {topic.lower()}."


# ============================================================================
# PROGRESS MANAGEMENT (Resume Support)
# ============================================================================

def load_progress(progress_path: str) -> dict:
    """Load progress from pickle file."""
    if os.path.exists(progress_path):
        try:
            with open(progress_path, "rb") as f:
                return pickle.load(f)
        except Exception:
            return {}
    return {}


def save_progress(progress_path: str, progress: dict):
    """Save progress to pickle file."""
    with open(progress_path, "wb") as f:
        pickle.dump(progress, f)


def clear_progress(progress_path: str):
    """Remove progress file after successful completion."""
    if os.path.exists(progress_path):
        os.remove(progress_path)


# ============================================================================
# EXCEL FORMATTING
# ============================================================================

def create_formatted_workbook(rows: list, output_path: str):
    """Create a professionally formatted Excel workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "GeoJSON Catalog"

    # ---- Styles ----
    header_font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_font = Font(name="Calibri", size=10)
    data_alignment = Alignment(vertical="top", wrap_text=False)
    wrap_alignment = Alignment(vertical="top", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="B4C6E7"),
        right=Side(style="thin", color="B4C6E7"),
        top=Side(style="thin", color="B4C6E7"),
        bottom=Side(style="thin", color="B4C6E7"),
    )

    header_border = Border(
        left=Side(style="thin", color="1F3864"),
        right=Side(style="thin", color="1F3864"),
        top=Side(style="thin", color="1F3864"),
        bottom=Side(style="medium", color="1F3864"),
    )

    # Alternating row fills
    even_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    odd_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    # Wrap text columns (0-indexed column positions)
    wrap_columns = {2, 3, 4}  # Layer Available Fields, How to Use This File, Target Industries

    # ---- Write Header ----
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = header_border

    # ---- Write Data ----
    center_alignment = Alignment(horizontal="center", vertical="top")
    for row_idx, row_data in enumerate(rows, start=2):
        row_fill = even_fill if row_idx % 2 == 0 else odd_fill
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.border = thin_border
            cell.fill = row_fill

            if col_idx in (1, 8):  # Sr. No, Area Wise Type
                cell.alignment = center_alignment
            elif (col_idx - 1) in wrap_columns:
                cell.alignment = wrap_alignment
            else:
                cell.alignment = data_alignment

    # ---- Column Widths ----
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        col_letter = get_column_letter(col_idx)
        width = COLUMN_WIDTHS.get(col_name, 15)
        ws.column_dimensions[col_letter].width = width

    # ---- Freeze Header Row ----
    ws.freeze_panes = "A2"

    # ---- Auto Filter ----
    last_col_letter = get_column_letter(len(COLUMNS))
    last_row = len(rows) + 1
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    # ---- Set row heights ----
    ws.row_dimensions[1].height = 30  # Header row

    # ---- Save ----
    try:
        wb.save(output_path)
        print(f"\nExcel file created: {output_path}")
    except PermissionError:
        print(f"\nERROR: Could not save directly to '{output_path}' (file locked by Excel).")
        temp_out = output_path.replace(".xlsx", "_new.xlsx")
        wb.save(temp_out)
        print(f"Saved catalog to: {temp_out}")


# ============================================================================
# MAIN PROCESSING
# ============================================================================

def main():
    """Main function to process all GeoJSON files and generate catalog."""
    # Determine the script's directory as the base
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(script_dir, INPUT_FOLDER)
    output_path = os.path.join(script_dir, OUTPUT_FILE)
    progress_path = os.path.join(script_dir, PROGRESS_FILE)

    if not os.path.isdir(input_folder):
        print(f"ERROR: Input folder '{input_folder}' not found.")
        sys.exit(1)

    # Discover all .geojson files
    # Discover all supported dataset files (.geojson and .tif/.tiff)
    supported_exts = (".geojson", ".tif", ".tiff")
    dataset_files = sorted([
        f for f in os.listdir(input_folder)
        if f.lower().endswith(supported_exts)
    ])

    if not dataset_files:
        print(f"No supported datasets found in '{input_folder}'.")
        sys.exit(1)

    print(f"Found {len(dataset_files)} datasets in '{INPUT_FOLDER}'")
    print("=" * 70)

    # Load any previous progress
    progress = load_progress(progress_path)
    completed_rows = progress.get("completed_rows", {})
    skipped_files = progress.get("skipped_files", {})

    if completed_rows:
        print(f"Resuming from previous run: {len(completed_rows)} files already processed.")

    processed_count = 0
    skipped_count = 0
    new_processed = 0
    all_rows = []

    for idx, filename in enumerate(dataset_files, start=1):
        layer_name = clean_layer_name(filename)
        filepath = os.path.join(input_folder, filename)

        # Check if already processed in a previous run
        if filename in completed_rows:
            all_rows.append(completed_rows[filename])
            processed_count += 1
            continue

        # Check if previously skipped
        if filename in skipped_files:
            skipped_count += 1
            continue

        print(f"[{idx}/{len(dataset_files)}] Processing: {filename} ...", end=" ")

        try:
            ext = os.path.splitext(filename)[1].lower()
            if ext in (".tif", ".tiff"):
                metadata = extract_properties_from_geotiff(filepath)
            else:
                metadata = extract_properties_from_geojson(filepath)

            if not metadata["has_features"] and metadata["feature_count"] == 0:
                # Empty or no features
                print("SKIPPED (empty/no features)")
                skipped_files[filename] = "Empty or no features"
                skipped_count += 1
                # Save progress after each skip
                progress["completed_rows"] = completed_rows
                progress["skipped_files"] = skipped_files
                save_progress(progress_path, progress)
                continue

            fields = metadata["fields"]
            geometry_type = metadata["geometry_type"]
            feature_count = metadata["feature_count"]

            # Layer Available Fields - actual fields directly from GeoJSON (natural order)
            if fields:
                layer_fields = ", ".join(fields)
            else:
                layer_fields = ""

            # Area Wise Type
            area_type = determine_area_type(layer_name, fields)

            # How to Use (concise, 1 short sentence)
            how_to_use = generate_short_usage_description(layer_name, area_type)

            # Target Industries
            target_industries = determine_industries(layer_name, fields, geometry_type or "")

            # Build row (Source, Potential Analysis, and Audit By left blank)
            row = [
                None,  # Sr. No - will be assigned after sorting
                layer_name,
                layer_fields,
                how_to_use,
                target_industries,
                "",  # Source - data removed per user requirement
                "",  # Potential Analysis - left blank per requirement
                area_type,
                "",  # Audit By - left blank per requirement
            ]

            completed_rows[filename] = row
            all_rows.append(row)
            processed_count += 1
            new_processed += 1
            print("OK")

            # Save progress after each successful processing
            progress["completed_rows"] = completed_rows
            progress["skipped_files"] = skipped_files
            save_progress(progress_path, progress)

        except Exception as e:
            print(f"SKIPPED ({type(e).__name__}: {e})")
            skipped_files[filename] = str(e)
            skipped_count += 1
            # Save progress after each skip
            progress["completed_rows"] = completed_rows
            progress["skipped_files"] = skipped_files
            save_progress(progress_path, progress)
            continue

    # Assign serial numbers
    for i, row in enumerate(all_rows, start=1):
        row[0] = i

    print("\n" + "=" * 70)

    # Generate Excel
    if all_rows:
        print(f"\nGenerating Excel workbook with {len(all_rows)} entries...")
        create_formatted_workbook(all_rows, output_path)
    else:
        print("No data to write. All files were skipped.")

    # Print summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"  Total files found:   {len(dataset_files)}")
    print(f"  Processed:           {processed_count}")
    print(f"  Skipped:             {skipped_count}")
    if new_processed > 0:
        print(f"  Newly processed:     {new_processed}")
    print(f"  Excel file created:  {OUTPUT_FILE}")

    if skipped_files:
        print(f"\nSkipped files ({len(skipped_files)}):")
        for fname, reason in sorted(skipped_files.items()):
            print(f"  - {fname}: {reason}")

    # Clear progress file on successful completion
    clear_progress(progress_path)
    print("\nDone! Progress file cleared.")


if __name__ == "__main__":
    main()
