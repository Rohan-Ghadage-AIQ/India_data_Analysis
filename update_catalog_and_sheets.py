#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Update Catalog and First 3 Layer Sheets (Senior Geospatial Analyst Edition)
===========================================================================
1. In 'GeoJSON Catalog':
   - Removes 'Potential Analysis' column if present.
   - Pastes official ArcGIS URLs into 'Source' column (Column 6) for rows 2, 3, and 4.
   - Populates 'How to Use This File' (Column 4) with deep Geospatial Analyst protocols.
   - Populates 'Target Industries' (Column 5) with comprehensive domain industry lists.
   - Adjusts column widths, row heights, alignment, and auto-filter (A1:H428).
2. In individual layer sheets:
   - '1_Accidents_Victims_classified' (66 fields)
   - '2_Accidents_classified_accordin' (21 fields)
   - '3_Accidents_classified_accordin' (78 fields)
   Populates 'Target Industries', 'Potential Analysis', and 'Remark' from the
   standpoint of a Senior Geospatial Analyst & Spatial Data Scientist.
"""

import sys
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Styles matching create_layer_sheets.py
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
    left=SIDE_THIN_LIGHT,
    right=SIDE_THIN_LIGHT,
    top=SIDE_THIN_LIGHT,
    bottom=SIDE_THIN_LIGHT,
)
BORDER_HEADER_CELL = Border(
    left=SIDE_THIN_DARK,
    right=SIDE_THIN_DARK,
    top=SIDE_THIN_DARK,
    bottom=SIDE_MEDIUM_DARK,
)

FILL_ZEBRA_EVEN = PatternFill(start_color=COLOR_ZEBRA_EVEN, end_color=COLOR_ZEBRA_EVEN, fill_type="solid")
FILL_ZEBRA_ODD = PatternFill(start_color=COLOR_ZEBRA_ODD, end_color=COLOR_ZEBRA_ODD, fill_type="solid")

FONT_DATA = Font(name="Calibri", size=10, color="000000")
FONT_LINK = Font(name="Calibri", size=10, color=COLOR_LINK_BLUE, underline="single")
ALIGN_LEFT_TOP_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER_TOP = Alignment(horizontal="center", vertical="top")


# =============================================================================
# LAYER 1: Non-Wearing of Helmet & Seat Belt Safety Devices
# =============================================================================
def get_layer1_field_info(field_name: str):
    f = field_name.strip()
    
    # Metadata / Identifiers
    if f == "OBJECTID":
        return (
            "GIS & Spatial Data Infrastructure, IT & Spatial Database Services",
            "R-tree spatial indexing, topological geometry validation, and primary key joins with national spatial data infrastructure (NSDI)",
            "ESRI internal primary spatial key"
        )
    elif f == "Name":
        return (
            "State & Central Governance, Urban & Regional Planning, Transport Administration",
            "Thematic choropleth classification (Natural Breaks/Jenks), administrative boundary aggregation, and spatial dissolve operations",
            "Standard State / Union Territory administrative name"
        )
    elif f == "State Code (LGD)":
        return (
            "E-Governance, Public Administration, Inter-Departmental Spatial Data Integration",
            "Relational foreign key joins with Local Government Directory (LGD/NIC) datasets and national administrative boundary layers",
            "Standard Government of India LGD code for national interoperability"
        )
    elif f == "Census 2011 Code":
        return (
            "Socio-Economic Research, Demographic Analytics, Policy Planning",
            "Spatial join with Census 2011 demographic polygons for population-at-risk normalization and socio-economic vulnerability indexing",
            "Standard Census 2011 code for demographic normalization"
        )
        
    # Helmet Non-Use Fields (Drivers & Passengers / Pillion)
    elif "Non Wearing of Helmet" in f:
        is_driver = "Drivers Persons" in f
        role_str = "two-wheeler driver" if is_driver else "pillion passenger"
        
        if "Killed in Numbers" in f:
            if is_driver:
                return (
                    "Traffic Police & Law Enforcement, MoRTH, Helmet Manufacturers, Motor & Life Insurance",
                    "Hotspot Analysis (Getis-Ord Gi*) to isolate statistically significant High-High fatality clusters; choropleth mapping normalized by registered 2-wheelers (Vahan GIS); spatial overlay with National/State Highway corridors",
                    "Primary indicator for deploying automated camera enforcement (ANPR/e-challan) for helmet detection"
                )
            else:
                return (
                    "Traffic Law Enforcement, Helmet Manufacturers, Motor Insurance, Road Safety NGOs",
                    "Spatial calculation of Pillion-to-Driver casualty ratios to isolate states where pillion helmet enforcement severely lags behind rider compliance; spatial regression against state penalty structures",
                    "Pillion rider compliance index (mandatory under Motor Vehicles Act)"
                )
        elif "Killed in Ranks" in f:
            return (
                "MoRTH, State Road Safety Councils, Transport Policy Formulation",
                f"Spatial rank-order correlation (Spearman's rho on spatial weights matrix) to detect inter-state policy and enforcement divergence for {role_str} fatalities across neighboring jurisdictions",
                f"State benchmark rank (1 = highest {role_str} fatalities)"
            )
        elif "Grievously Injured" in f:
            return (
                "Trauma & Emergency Healthcare Networks, Health & Motor Insurance, Emergency Medical Services (EMS)",
                f"Network service area analysis (isochrone modeling) mapping grievous head/spine injury density of {role_str}s against Level-1/2 Golden Hour trauma care hospital reachability",
                f"Identifies critical trauma care deficit zones for {role_str}s"
            )
        elif "Minor Injury" in f:
            return (
                "Outpatient Healthcare Facilities, General Insurance Claims Underwriting, Road Safety Advocacy",
                f"Spatial injury severity ratio mapping (Minor Injury vs. Grievous/Fatal) to differentiate low-speed urban collisions from high-speed rural highway crashes for {role_str}s",
                "Evaluates low-impact urban collision distribution"
            )
        elif "Total Injured" in f:
            return (
                "Public Health Systems, Motor Insurance Providers, Transport Regulators",
                f"Bivariate spatial autocorrelation (Local Moran's I) mapping total {role_str} injury burden against road network density and urbanization index",
                f"Combined non-fatal {role_str} casualty metric"
            )

    # Seat Belt Non-Use Fields (Drivers & Passengers)
    elif "Non Wearing of Seat Belt" in f:
        is_driver = "Drivers Persons" in f
        role_str = "four-wheeler driver" if is_driver else "passenger (front/rear)"
        
        if "Killed in Numbers" in f:
            if is_driver:
                return (
                    "Automotive OEMs & Safety Systems, Traffic Police, Motor Insurance Underwriting, MoRTH",
                    "Spatial cluster mapping of unrestrained driver fatalities along high-speed expressways and national freight corridors; spatial correlation with vehicle age and commercial fleet density",
                    "Metric for in-cabin restraint enforcement and crashworthiness spatial modeling"
                )
            else:
                return (
                    "Commercial Fleet & Ride-Hailing (Ola/Uber), Traffic Enforcement, Motor Insurance, MoRTH",
                    "Spatial calculation of Passenger-to-Driver seatbelt fatality ratios to detect rear-seat restraint compliance deficits; spatial corridor analysis of taxi and inter-city passenger transit routes",
                    "Rear & co-passenger seatbelt compliance indicator"
                )
        elif "Killed in Ranks" in f:
            return (
                "MoRTH, State Transport Departments, Road Safety Councils",
                f"Spatial rank distribution and year-on-year rank transition tracking across states for unrestrained {role_str} fatalities",
                f"State benchmark rank for unrestrained {role_str} deaths"
            )
        elif "Grievously Injured" in f:
            return (
                "Trauma Healthcare Centers, Health & Motor Insurance, Crash Forensics",
                f"Spatial correlation of blunt thoracic/abdominal injury clusters with high-speed road networks for unrestrained {role_str}s; disability claims forecasting",
                f"Severe trauma metric from absence of seatbelt restraint ({'driver steering-wheel impact' if is_driver else 'passenger cabin ejection'})"
            )
        elif "Minor Injury" in f:
            return (
                "Outpatient Healthcare, Motor Insurance Surveyors, Consumer Safety Forums",
                f"Spatial distribution of low-speed cabin impact injuries in urban agglomerations for unrestrained {role_str}s",
                f"Minor cabin impact injury metric for {role_str}s"
            )
        elif "Total Injured" in f:
            return (
                "Healthcare Networks, Motor Insurance Providers, Transport Regulators",
                f"Cumulative spatial mapping of non-fatal unrestrained {role_str} casualties against state vehicle fleet growth rates",
                f"Total non-fatal {role_str} casualties without seat belt"
            )
            
    # Spatial Geometry Fields
    elif "st_area" in f.lower():
        return (
            "GIS & Geospatial Analytics, Cartography, Regional Planning",
            "Spatial normalization denominator for computing accident, fatality, and injury density per 1,000 km²",
            "Projected polygon area in coordinate square units"
        )
    elif "st_perimeter" in f.lower():
        return (
            "GIS & Geospatial Analytics, Boundary Mapping, Infrastructure Planning",
            "Boundary compactness analysis (Polsby-Popper / Schwartzberg index) to evaluate border effects in interstate traffic spillover",
            "Projected polygon perimeter in coordinate linear units"
        )
    
    return (
        "Transportation & Traffic Safety, General Insurance, Public Policy",
        "Spatial statistical safety analysis, state-level risk assessment",
        "Official MoRTH Road Accident Statistics attribute"
    )


# =============================================================================
# LAYER 2: Accidents Classified According to Driver License Status
# =============================================================================
def get_layer2_field_info(field_name: str):
    f = field_name.strip()
    
    # Metadata / Identifiers
    if f == "OBJECTID":
        return (
            "GIS & Spatial Data Infrastructure, IT & Spatial Database Services",
            "R-tree spatial indexing, topological geometry validation, and primary key joins with national spatial data infrastructure (NSDI)",
            "ESRI internal primary spatial key"
        )
    elif f == "Name":
        return (
            "State & Central Governance, Urban & Regional Planning, Transport Administration",
            "Thematic choropleth classification (Natural Breaks/Jenks), administrative boundary aggregation, and spatial dissolve operations",
            "Standard State / Union Territory administrative name"
        )
    elif f == "State Code (LGD)":
        return (
            "E-Governance, Public Administration, Inter-Departmental Spatial Data Integration",
            "Relational foreign key joins with Local Government Directory (LGD/NIC) datasets and national administrative boundary layers",
            "Standard Government of India LGD code for national interoperability"
        )
    elif f == "Census 2011 Code":
        return (
            "Socio-Economic Research, Demographic Analytics, Policy Planning",
            "Spatial join with Census 2011 demographic polygons for population-at-risk normalization and socio-economic vulnerability indexing",
            "Standard Census 2011 code for demographic normalization"
        )
        
    # License Category Fields
    elif "Valid Permanent License" in f:
        return (
            "Motor Insurance Underwriting, Commercial Driving Schools, RTO Transport Departments, Automotive OEMs",
            "Isolation of infrastructure-induced crashes (road geometry black spots, poor lighting, missing signage) vs. driver competence; spatial correlation with International Roughness Index (IRI) and traffic volume",
            "Certified driver crashes reflect roadway engineering defects rather than incompetence"
        )
    elif "Learner's Licence" in f or "Learner" in f:
        return (
            "Driving Training Schools & Institutes, RTO Licensing Authorities, Road Safety Advocacy NGOs",
            "Spatial correlation between novice driver crash density and the geographic distribution of Automated Driving Test Centers (ADTCs) and RTO testing tracks; curriculum outcome assessment",
            "Measures novice driver safety and licensing curriculum effectiveness"
        )
    elif "Without Licence" in f:
        return (
            "Traffic Police & Law Enforcement, Judicial & Legal Authorities, Motor Insurance Fraud Units, MoRTH",
            "Hotspot analysis (Getis-Ord Gi*) of unlicensed driver crash proportions (Without Licence / Total Accidents); identification of systemic enforcement voids, illegal driving corridors, and underage driving zones",
            "Critical traffic hazard: highlights acute enforcement and compliance voids"
        )
    elif "Others Not known" in f:
        return (
            "Traffic Police & Crime Bureaus, Crash Forensics, Insurance Claims Investigation",
            "Kernel density estimation (KDE) of untraced/unrecorded driver incidents to isolate surveillance blind spots, hit-and-run highway stretches, and cross-border jurisdiction gaps",
            "Indicator of highway surveillance blind spots and hit-and-run frequency"
        )
    elif "Total Accidents" in f:
        return (
            "Ministry of Road Transport & Highways (MoRTH), NHAI, General Insurance Councils, Urban Planning",
            "Multi-year temporal trajectory (2021–2023) and spatial shift of total accident density; baseline benchmark for national Vision Zero road safety targets",
            "State aggregate crash volume benchmark"
        )
        
    # Spatial Geometry Fields
    elif "st_area" in f.lower():
        return (
            "GIS & Geospatial Analytics, Cartography, Regional Planning",
            "Spatial normalization denominator for computing accident density per 1,000 km²",
            "Projected polygon area in coordinate square units"
        )
    elif "st_perimeter" in f.lower():
        return (
            "GIS & Geospatial Analytics, Boundary Mapping, Infrastructure Planning",
            "Boundary compactness analysis (Polsby-Popper / Schwartzberg index) to evaluate border effects in interstate traffic spillover",
            "Projected polygon perimeter in coordinate linear units"
        )
        
    return (
        "Transportation & Traffic Safety, General Insurance, Public Policy",
        "Spatial statistical safety analysis, state-level risk assessment",
        "Official MoRTH Road Accident Statistics attribute"
    )


# =============================================================================
# LAYER 3: Accidents Classified According to Vehicle Load Condition
# =============================================================================
def get_layer3_field_info(field_name: str):
    f = field_name.strip()
    
    # Metadata / Identifiers
    if f == "OBJECTID":
        return (
            "GIS & Spatial Data Infrastructure, IT & Spatial Database Services",
            "R-tree spatial indexing, topological geometry validation, and primary key joins with national spatial data infrastructure (NSDI)",
            "ESRI internal primary spatial key"
        )
    elif f == "Name":
        return (
            "State & Central Governance, Urban & Regional Planning, Transport Administration",
            "Thematic choropleth classification (Natural Breaks/Jenks), administrative boundary aggregation, and spatial dissolve operations",
            "Standard State / Union Territory administrative name"
        )
    elif f == "State Code (LGD)":
        return (
            "E-Governance, Public Administration, Inter-Departmental Spatial Data Integration",
            "Relational foreign key joins with Local Government Directory (LGD/NIC) datasets and national administrative boundary layers",
            "Standard Government of India LGD code for national interoperability"
        )
    elif f == "Census 2011 Code":
        return (
            "Socio-Economic Research, Demographic Analytics, Policy Planning",
            "Spatial join with Census 2011 demographic polygons for population-at-risk normalization and socio-economic vulnerability indexing",
            "Standard Census 2011 code for demographic normalization"
        )
        
    # Normally Loaded Vehicles
    elif f.startswith("Normally Loaded"):
        if "Accidents" in f:
            return (
                "Commercial Logistics & Fleet Operators, National Highways Authority (NHAI), Motor Insurance",
                "Baseline spatial crash density under standard operating limits; bivariate regression against state freight tonnage throughput and national highway network density",
                "Represents the largest volume share of total vehicle accidents (MoRTH 2023)"
            )
        elif "Rank" in f and "Persons" not in f:
            return (
                "MoRTH, State Transport Authorities, Freight Logistics Councils",
                "Spatial rank distribution and inter-state comparative benchmarking of baseline freight movements",
                "State benchmark rank for normally loaded vehicle accidents"
            )
        elif "Persons Killed" in f:
            return (
                "Life & Motor Insurance, Emergency Medical Services (EMS), Automotive Commercial OEMs",
                "Spatial lethality modeling (Killed / Accidents) for standard freight traffic; evaluation of commercial vehicle passive crash protection and road barrier crashworthiness",
                "Fatalities occurring under compliant loading conditions"
            )
        elif "Persons Rank" in f:
            return (
                "MoRTH, State Road Safety Councils",
                "Comparative state ranking of mortality from standard freight/passenger movements",
                "State benchmark rank for normally loaded fatalities"
            )
        elif "Grievously Injured" in f:
            return (
                "Trauma Care Centers, Health & Motor Insurance, Emergency Response",
                "Network accessibility modeling of severe freight crash victims to designated Golden Hour trauma centers",
                "Severe trauma metric from standard vehicle collisions"
            )
        elif "Minor Injured" in f:
            return (
                "Outpatient Healthcare, Insurance Claims Management",
                "Spatial frequency modeling of non-critical commercial collision injuries in urban delivery belts",
                "Minor casualty indicator for compliant loading conditions"
            )

    # Overloaded / Hanging Load Vehicles
    elif f.startswith("Overloaded/Hanging"):
        if "Accidents" in f:
            return (
                "Highway Patrol, Weighbridge & Toll Plaza Operators, Freight Logistics, MoRTH",
                "Spatial identification of high-risk overloaded freight corridors; spatial multi-criteria evaluation (SMCE) to prioritize automated Weigh-in-Motion (WIM) sensors and virtual weighbridges",
                "Critical traffic hazard: severe compromise to vehicle braking distance and roll stability"
            )
        elif "Rank" in f and "Persons" not in f:
            return (
                "MoRTH, State Transport Departments, Traffic Police",
                "Spatial rank tracking to evaluate state-level enforcement efficacy against illegal freight overloading",
                "State benchmark rank for overloading-related crashes"
            )
        elif "Persons Killed" in f:
            return (
                "Motor Insurance Underwriting, Highway Safety Authorities, Forensic Crash Investigators",
                "Overloading Lethality Index calculation (Killed / Accidents) comparing kinetic fatality severity against normal loads; third-party commercial liability risk modeling",
                "Extreme lethality metric driven by massive momentum and rollover tendencies"
            )
        elif "Persons Rank" in f:
            return (
                "MoRTH, State Road Safety Committees",
                "Inter-state mortality benchmarking from commercial vehicle overloading violations",
                "State benchmark rank for overloading fatalities"
            )
        elif "Grievously Injured" in f:
            return (
                "Trauma Care & Emergency Surgery, Health & Disability Insurance",
                "Spatial modeling of severe crush and multi-trauma injuries along heavy transit highways",
                "High-impact crush trauma metric from overloaded vehicle collisions"
            )
        elif "Minor Injured" in f:
            return (
                "Emergency Medical Units, Insurance Claims Surveyors",
                "Spatial incidence tracking of non-fatal casualties in overloaded freight corridors",
                "Minor injury metric in overloading incidents"
            )

    # Empty Vehicles (Unladen return trips)
    elif f.startswith("Empty"):
        if "Accidents" in f:
            return (
                "Commercial Fleet Management, Logistics & Supply Chain, Highway Safety",
                "Spatial mapping of 'deadhead' freight logistics corridors where unladen commercial vehicles traveling at excessive return speeds with reduced wet-tire traction experience loss-of-control crashes",
                "Highlights safety risks of unladen commercial vehicles on return transit trips"
            )
        elif "Persons Killed" in f:
            return (
                "Commercial Fleet Operators, Motor Insurance Underwriting",
                "Fatality rate modeling for unladen high-speed truck collisions with smaller vehicles and two-wheelers",
                "Fatalities resulting from high-speed unladen vehicle collisions"
            )
        elif "Grievously Injured" in f:
            return (
                "Trauma Healthcare Providers, Motor Insurance Underwriters",
                "Spatial analysis of severe collision trauma along empty return freight routes",
                "Severe trauma metric from empty vehicle collisions"
            )
        elif "Minor Injured" in f:
            return (
                "Emergency Clinics, Insurance Claims Assessment",
                "Minor casualty tracking in deadhead logistics corridors",
                "Minor injury metric from empty vehicle collisions"
            )

    # Unknown Load Condition
    elif f.startswith("Not Known"):
        if "Accidents" in f:
            return (
                "Traffic Police, Accident Investigation Agencies, MoRTH",
                "Spatial audit of crash reporting data completeness; identification of police jurisdictions with unclassified accident records and missing technical inspection logs",
                "Indicator of incomplete accident investigation or hit-and-run events"
            )
        elif "Persons Killed" in f:
            return (
                "Forensic Investigators, Insurance Claims Surveyors",
                "Fatality mapping in unclassified or untraced vehicle crashes",
                "Fatalities where loading condition was unrecorded"
            )
        elif "Grievously Injured" in f:
            return (
                "Emergency Healthcare Services, Health Insurance",
                "Trauma demand forecasting for unclassified collision categories",
                "Grievous injuries in unclassified load condition crashes"
            )
        elif "Minor Injured" in f:
            return (
                "Healthcare Facilities, Insurance Surveyors",
                "Minor casualties in unclassified collision categories",
                "Minor injuries in unclassified load condition crashes"
            )

    # Total All India Metrics
    elif f.startswith("Total all India"):
        if "Accidents" in f:
            return (
                "Ministry of Road Transport & Highways (MoRTH), NHAI, IRDAI",
                "Multi-year macro trend and spatial density modeling (crashes per 1,000 km² and per km of National Highway); national road safety target tracking",
                "State aggregate road accident volume"
            )
        elif "Persons Killed" in f:
            return (
                "MoRTH, Life & Health Insurance, National Road Safety Board",
                "State mortality rate calculation and spatial progress monitoring toward Vision Zero fatality reduction targets",
                "Total road accident fatalities in the state"
            )
        elif "Persons Grievously Injured" in f:
            return (
                "National Health Authority, Trauma Care Networks, Insurance Regulators",
                "State trauma healthcare infrastructure load forecasting and long-term economic disability cost estimation",
                "Total grievous road injuries in the state"
            )
        elif "Persons Minor Injured" in f:
            return (
                "Public Health Systems, Motor Insurance Providers",
                "Total non-critical casualty volume modeling and lost-workday economic productivity analysis",
                "Total minor road injuries in the state"
            )

    # Spatial Geometry Fields
    elif "st_area" in f.lower():
        return (
            "GIS & Geospatial Analytics, Cartography, Regional Planning",
            "Spatial normalization denominator for computing accident, fatality, and injury density per 1,000 km²",
            "Projected polygon area in coordinate square units"
        )
    elif "st_perimeter" in f.lower():
        return (
            "GIS & Geospatial Analytics, Boundary Mapping, Infrastructure Planning",
            "Boundary compactness analysis (Polsby-Popper / Schwartzberg index) to evaluate border effects in interstate traffic spillover",
            "Projected polygon perimeter in coordinate linear units"
        )

    return (
        "Transportation & Traffic Safety, Commercial Fleet Logistics, General Insurance",
        "Vehicle load safety analytics, regulatory compliance modeling",
        "Official MoRTH Road Accident Statistics attribute"
    )


# =============================================================================
# WORKBOOK UPDATE ORCHESTRATOR
# =============================================================================
def update_catalog_and_sheets(workbook_path: str):
    print(f"[*] Loading workbook: {workbook_path} ...")
    wb = openpyxl.load_workbook(workbook_path)
    
    # -------------------------------------------------------------------------
    # 1. Update Master Sheet: 'GeoJSON Catalog'
    # -------------------------------------------------------------------------
    cat_ws = wb["GeoJSON Catalog"]
    print(f"[*] Processing 'GeoJSON Catalog' (max_row={cat_ws.max_row}, max_column={cat_ws.max_column}) ...")
    
    # Locate headers
    potential_col_idx = None
    source_col_idx = None
    how_to_use_col_idx = None
    target_ind_col_idx = None
    
    for c in range(1, cat_ws.max_column + 1):
        val = str(cat_ws.cell(1, c).value or "").strip()
        if val == "Potential Analysis":
            potential_col_idx = c
        elif val == "Source":
            source_col_idx = c
        elif val == "How to Use This File":
            how_to_use_col_idx = c
        elif val == "Target Industries":
            target_ind_col_idx = c

    print(f"    Columns detected -> HowToUse: {how_to_use_col_idx}, TargetInd: {target_ind_col_idx}, Source: {source_col_idx}, PotentialAnalysis: {potential_col_idx}")

    # Geospatial Analyst Deep-Dive: How to Use This File (Column 4)
    geospatial_how_to_use = {
        2: (
            "Geospatial Analysis & Modeling Protocol:\n"
            "1. Spatial Autocorrelation & Cluster Hotspots: Execute Getis-Ord Gi* and Anselin Local Moran's I to isolate statistically significant High-High spatial clusters of helmet and seatbelt non-compliance fatalities across state polygons.\n"
            "2. Vulnerability Ratios: Compute spatial ratios of pillion-to-driver fatalities and passenger-to-driver seatbelt casualties to benchmark states where pillion helmet and rear-seatbelt enforcement severely lags.\n"
            "3. Spatio-Temporal Shift (2021-2023): Model mean directional center shifts and percentage trajectory to evaluate the geographic efficacy of state-level MV Act amendments and automated camera (ANPR/e-challan) enforcement networks.\n"
            "4. Risk Normalization & Exposure: Normalize casualties against administrative polygon area (st_area), state road network length, and registered two-wheeler/four-wheeler vehicular density (Vahan linkage).\n"
            "5. Trauma Care Facility Allocation: Spatially intersect high grievous-injury polygons with national highway corridors and tertiary trauma centers to identify Golden Hour emergency medical care deficit zones."
        ),
        3: (
            "Geospatial Analysis & Modeling Protocol:\n"
            "1. Enforcement Void & Unlicensed Hotspot Mapping: Run spatial cluster analysis on the proportion of crashes caused by unlicensed drivers (Without Licence / Total Accidents) to identify regional governance blind spots, illegal driving corridors, and underage driving zones.\n"
            "2. Surveillance & Hit-and-Run Void Analysis: Map the spatial density of 'Others Not Known' driver categories to isolate highway corridors with acute surveillance black spots, missing CCTV coverage, and low police FIR/DAR recording completeness.\n"
            "3. Novice Driver Vulnerability & Testing Efficacy: Correlate Learner's License crash rates with the geographic distribution of Regional Transport Offices (RTOs) and Automated Driving Test Centers (ADTCs) to evaluate driver licensing curriculum outcomes.\n"
            "4. Human Factor vs. Infrastructure Defect Isolation: Contrast high crash rates among 'Valid Permanent License' holders against unlicensed drivers; certified driver crash clusters isolate roadway engineering black spots, geometric defects, and signage deficiencies rather than driver incompetence.\n"
            "5. Predictive Patrol & Checkpoint Allocation: Perform location-allocation network modeling to optimize the spatial deployment of mobile traffic interceptors and digital license verification checkpoints."
        ),
        4: (
            "Geospatial Analysis & Modeling Protocol:\n"
            "1. Freight Corridor Overloading Lethality Index: Calculate and map the spatial Lethality Index (Persons Killed / Accidents) comparing overloaded/hanging loads against normally loaded vehicles to evaluate kinetic impact severity and rollover tendencies along national freight corridors.\n"
            "2. Unladen (Empty) Return-Trip Speed Dynamic Analysis: Spatially isolate crashes involving 'Empty' commercial vehicles to detect high-speed deadhead freight logistics corridors where unladen trucks face reduced wet-tire traction and loss-of-control risks.\n"
            "3. Weigh-in-Motion (WIM) Spatial Site Prioritization: Execute Spatial Multi-Criteria Evaluation (SMCE) overlaying overloaded vehicle crash hotspots with NHAI toll plazas and state border checkpoints to prioritize automated Weigh-in-Motion sensors and virtual weighbridges.\n"
            "4. Freight Volume & Economic Output Normalization: Normalize normally loaded and overloaded vehicle crashes against state industrial manufacturing output (IIP), mineral/mining transit routes, and highway network density (st_area).\n"
            "5. Temporal Axle-Load Policy Impact Modeling: Track 2021-2023 state ranking divergence to evaluate how revised motor vehicle axle-load regulations and interstate check-post removals have influenced freight transit safety."
        ),
    }

    # Master Sheet Target Industries (Column 5)
    master_target_industries = {
        2: "Traffic Police & Law Enforcement, Road Safety Regulators (MoRTH), Automotive OEMs & Safety Systems, Motor & Life Insurance, Emergency Medical & Trauma Healthcare",
        3: "RTO & Licensing Authorities, Traffic Police & Law Enforcement, Driving Schools & Training Institutes, Motor Insurance Underwriting, Judicial & Legal Systems",
        4: "Commercial Logistics & Fleet Operators, National Highways Authority (NHAI), Weighbridge & Toll Plaza Operators, Freight & Commercial Vehicle Insurance, Highway Patrol",
    }

    # Update How to Use This File
    if how_to_use_col_idx:
        print(f"    Updating 'How to Use This File' at column {how_to_use_col_idx}...")
        for row_idx, analysis_text in geospatial_how_to_use.items():
            cell = cat_ws.cell(row=row_idx, column=how_to_use_col_idx, value=analysis_text)
            cell.font = FONT_DATA
            cell.alignment = ALIGN_LEFT_TOP_WRAP
            cat_ws.row_dimensions[row_idx].height = 115
            print(f"    Row {row_idx} 'How to Use This File' updated.")

    # Update Target Industries in master sheet
    if target_ind_col_idx:
        print(f"    Updating 'Target Industries' at column {target_ind_col_idx}...")
        for row_idx, ind_text in master_target_industries.items():
            cell = cat_ws.cell(row=row_idx, column=target_ind_col_idx, value=ind_text)
            cell.font = FONT_DATA
            cell.alignment = ALIGN_LEFT_TOP_WRAP
            print(f"    Row {row_idx} 'Target Industries' updated.")

    # Target URLs from Description docx files
    top3_urls = {
        2: "https://www.arcgis.com/home/item.html?id=29612a43bf074167977b268c4d95871d",
        3: "https://www.arcgis.com/home/item.html?id=5ed657af3ce746c3b3bbc4e623581a1c",
        4: "https://www.arcgis.com/home/item.html?id=49580888f4654af4b24d5b1e43599448",
    }
    
    # Paste URLs in Source column
    if source_col_idx:
        for row_idx, url in top3_urls.items():
            cell = cat_ws.cell(row=row_idx, column=source_col_idx, value=url)
            cell.hyperlink = url
            cell.font = FONT_LINK
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)
            print(f"    Row {row_idx} Source updated: {url}")
        
    # Delete 'Potential Analysis' column if present
    if potential_col_idx is not None:
        print(f"    Deleting column {potential_col_idx} ('Potential Analysis') ...")
        cat_ws.delete_cols(potential_col_idx, 1)
        
    new_headers = [cat_ws.cell(1, c).value for c in range(1, cat_ws.max_column + 1)]
    print(f"    Updated headers ({len(new_headers)} cols): {new_headers}")
    
    # Set tuned column widths for master sheet
    master_col_widths = {
        "A": 10,  # Sr. No
        "B": 42,  # Layer Name
        "C": 55,  # Layer Available Fields
        "D": 58,  # How to Use This File (expanded for geospatial analyst protocols)
        "E": 35,  # Target Industries
        "F": 50,  # Source (URL fits cleanly)
        "G": 20,  # Area Wise Type
        "H": 20,  # Audit By
    }
    for col_letter, w in master_col_widths.items():
        cat_ws.column_dimensions[col_letter].width = w
        
    # Update AutoFilter to cover all rows and new 8 columns
    max_col_letter = get_column_letter(len(new_headers))
    cat_ws.auto_filter.ref = f"A1:{max_col_letter}{cat_ws.max_row}"
    print(f"    AutoFilter updated to: {cat_ws.auto_filter.ref}")

    # -------------------------------------------------------------------------
    # 2. Update Individual Sheets for First 3 Rows
    # -------------------------------------------------------------------------
    layer_configs = [
        ("1_Accidents_Victims_classified", get_layer1_field_info),
        ("2_Accidents_classified_accordin", get_layer2_field_info),
        ("3_Accidents_classified_accordin", get_layer3_field_info),
    ]
    
    for sname, info_func in layer_configs:
        if sname not in wb.sheetnames:
            print(f"[!] Warning: Sheet '{sname}' not found in workbook!")
            continue
            
        ws = wb[sname]
        print(f"[*] Updating Sheet '{sname}' (max_row={ws.max_row}) ...")
        
        # Ensure column widths on individual sheet
        ws.column_dimensions["A"].width = 10  # Sr. No
        ws.column_dimensions["B"].width = 45  # Layer Available Fields
        ws.column_dimensions["C"].width = 38  # Target Industries
        ws.column_dimensions["D"].width = 58  # Potential Analysis (expanded for senior geospatial analyst depth)
        ws.column_dimensions["E"].width = 38  # Remark (clear and concise)
        
        updated_count = 0
        for r in range(4, ws.max_row + 1):
            field_name = str(ws.cell(r, 2).value or "").strip()
            if not field_name:
                continue
                
            target_ind, pot_analysis, remark = info_func(field_name)
            fill_style = FILL_ZEBRA_ODD if (r % 2 == 0) else FILL_ZEBRA_EVEN
            
            # Col 3: Target Industries
            c3 = ws.cell(row=r, column=3, value=target_ind)
            c3.font = FONT_DATA
            c3.alignment = ALIGN_LEFT_TOP_WRAP
            c3.border = BORDER_DATA_CELL
            c3.fill = fill_style
            
            # Col 4: Potential Analysis (Senior Geospatial Analyst Methodology)
            c4 = ws.cell(row=r, column=4, value=pot_analysis)
            c4.font = FONT_DATA
            c4.alignment = ALIGN_LEFT_TOP_WRAP
            c4.border = BORDER_DATA_CELL
            c4.fill = fill_style
            
            # Col 5: Remark (Clear, purposeful remark only where relevant)
            c5 = ws.cell(row=r, column=5, value=remark)
            c5.font = FONT_DATA
            c5.alignment = ALIGN_LEFT_TOP_WRAP
            c5.border = BORDER_DATA_CELL
            c5.fill = fill_style
            
            # Set comfortable row height for multi-line analysis text
            ws.row_dimensions[r].height = 48
            
            updated_count += 1
            
        print(f"    Successfully updated {updated_count} fields in '{sname}'")

    # -------------------------------------------------------------------------
    # 3. Save Workbook
    # -------------------------------------------------------------------------
    print(f"[*] Saving updated workbook to '{workbook_path}' ...")
    wb.save(workbook_path)
    print(f"[+] Workbook saved successfully!")


if __name__ == "__main__":
    target_wb = "exported_geojson_540andtif_catalog.xlsx"
    if not os.path.exists(target_wb):
        print(f"[!] Error: File '{target_wb}' does not exist.")
        sys.exit(1)
    update_catalog_and_sheets(target_wb)
