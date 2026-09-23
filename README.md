# Geospatial Layer Catalog & Analytical Metadata Framework

An automated, senior-grade geospatial metadata extraction, enrichment, and workbook cataloging pipeline. This repository provides automated tools to build, link, and format master multi-sheet Excel catalogs with senior geospatial analysis protocols, target industry listings, and authoritative source URLs.

---

## Repository Structure

```
.
├── Description/                             # 587 Word (.docx) layer specification documents
├── exported_geojson_540andtif_catalog.xlsx  # Master workbook with catalog + 427 linked child sheets
├── update_catalog_batch.py                 # Universal parameterized CLI for batch updating catalog & sheets
├── batch_update_100_layers.py              # Batch update script tailored for layers 1–100
├── update_catalog_and_sheets.py            # Reference field-level enrichment engine
├── create_layer_sheets.py                  # Script to generate individual child sheets per layer
├── generate_catalog.py                     # Initial catalog builder from raw geospatial assets
├── geojson_style_dashboard2.0.html         # Interactive web-based GeoJSON style & visualization dashboard
├── requirements.txt                        # Python dependencies
├── .gitignore                              # Excludes heavy raw GeoJSON/TIF data (66+ GB) and cache
└── README.md                               # Repository documentation
```

> **Note on Geospatial Data**: Raw GeoJSON files and GeoTIFFs (`exported_geojson_540andtif/` and `first100json/`, totaling ~66 GB) are excluded via `.gitignore` to keep the Git repository lightweight and performant.

---

## Features & Standards

### 1. Master Sheet (`GeoJSON Catalog`)
- **How to Use This File (Col 4)**: Tailored 5-point **Senior Geospatial Analysis & Modeling Protocol**:
  1. *Spatial Autocorrelation & Hotspot Detection* (Getis-Ord $G_i^*$, Anselin Local Moran's I).
  2. *Attribute Severity & Normalization Ratios* (Geometric density, area denominators).
  3. *Network Corridor & Multi-Modal Overlay* (Proximity buffering, infrastructure intersects).
  4. *Spatio-Temporal Trajectory Modeling* (Multi-year shifts, convergence/divergence).
  5. *Location-Allocation & Siting Optimization* (Deficit catchments, service allocation).
- **Target Industries (Col 5)**: Explicit stakeholder domain listings (e.g. NHAI, MoRTH, NHA, MoHFW, ECI, CWC, RBI).
- **Source (Col 6)**: Authoritative, verified ArcGIS item URLs as clickable hyperlinks.
- **Visual Presentation**: Row height 115 pt with text wrapping, navy headers, AutoFilter enabled.

### 2. Child Sheets (Individual Layer Sheets)
- **Table Header (Row 3)**:
  ```
  ['Sr. No', 'Layer Available Fields', 'Target Industries', 'Potential Analysis', 'Potential Analysis Manual', 'Remark']
  ```
- **`Potential Analysis Manual` (Col 5)**: **Strictly vacant** across all data rows, styled with borders and zebra striping for analyst manual entry.
- **`Potential Analysis` (Col 4)**: Attribute-level geospatial analytical methodology.
- **`Remark` (Col 6)**: Concise, high-value technical context notes.
- **Navigation & Layout**:
  - Cell `A1`: `← Back to Catalog` hyperlink pointing directly back to the corresponding row in `GeoJSON Catalog`.
  - Row height: 48 pt with text wrapping.
  - Frozen panes at `A4` for header persistence during scrolling.

---

## Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/Rohan-Ghadage-AIQ/India_data_Analysis.git
cd Excel_Sheet-json

# Install dependencies
pip install -r requirements.txt
```

### Running Batch Updates

Use `update_catalog_batch.py` to process any range of layers:

```bash
# Update layers 101 to 200
python update_catalog_batch.py --start 101 --end 200

# Update layers 1 to 100
python update_catalog_batch.py --start 1 --end 100

# Update layers 201 to 300
python update_catalog_batch.py --start 201 --end 300
```

#### CLI Parameters:
- `--start`: Starting serial number (default: `101`)
- `--end`: Ending serial number (default: `200`)
- `--workbook`: Target Excel workbook file (default: `exported_geojson_540andtif_catalog.xlsx`)

---

## Interactive Dashboard

Open `geojson_style_dashboard2.0.html` in any modern web browser to preview, style, and interactively filter GeoJSON layers with custom map legends and vector tiles.
