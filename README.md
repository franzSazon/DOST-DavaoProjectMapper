# DOST-Davao Project Mapper

![Demo](assets/project_demo.gif)

A Streamlit application for the Department of Science and Technology - Davao Region. It reads project tracker Excel workbooks for the CEST, LGIA, and SSCP programs, consolidates the data, and displays projects on an interactive map with filters, KPIs, and export options.

## Features

### Data Ingestion
- Two ingestion modes: **Division-Based Template** (separate CEST/LGIA/SSCP sheets) and **General Template** (auto-detects header rows and project sections in other layouts).
- Maps inconsistent source column headers (for example "Amount of Funding", "Project Cost", "Revised") to a single canonical schema.
- Normalizes division names, project status, and funding amounts (original and revised) across sheets.

### Mapping
- Displays projects on a Folium map centered on Davao, with marker clustering, division-based color coding, and configurable pin styles.
- Resolves missing coordinates using the Mapbox Geocoding API based on project location and beneficiary data.
- Clicking a marker opens a panel with project details: funding, timeline, proponent, and status.

### Filters and KPIs
- Sidebar filters for division, project status, approval date range, and budget range, with options to include or exclude records with missing values.
- KPI scorecards summarizing the currently filtered project set (counts, funding totals, status breakdown).
- Raw data table view of the cleaned dataset.
- CSV export of the filtered results.

### Map Export
- Generates a branded map image with legend and summary stats, using a headless browser screenshot pipeline.
- Optional feature. Requires additional dependencies listed in Installation.

### AI Summary and Q&A
- Generates a text summary of the filtered project data and supports free-form questions about the dataset.
- Uses a local Ollama LLM (`llama3.2:3b`). Data is processed on the local machine, not sent to an external API.
- Optional feature. Requires Ollama running locally.

## Installation

1. Clone the repository.
2. Install the core dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. To enable map export, install the additional dependencies:
   ```bash
   pip install selenium webdriver-manager pillow
   ```
   Without these, the app runs normally and the export feature is disabled.
4. To enable geocoding, add a Mapbox token to `.streamlit/secrets.toml`:
   ```toml
   MAPBOX_API_KEY = "pk.xxxxxxxx"
   ```
5. To enable the AI summary and Q&A features, install [Ollama](https://ollama.com), pull the `llama3.2:3b` model, and set the URL in `.streamlit/secrets.toml` if it is not the default:
   ```toml
   OLLAMA_URL = "http://localhost:11434"
   ```
6. Run the app:
   ```bash
   streamlit run app_refactor.py
   ```

## Usage

1. Select the ingestion mode that matches your workbook (Division-Based Template or General Template).
2. Upload the project Excel file in the sidebar.
3. Use the filters to narrow projects by division, status, date, or budget.
4. Click markers on the map for project details, check the KPI scorecards, and export data as CSV or a map image.
5. Optionally, generate an AI summary or ask questions about the filtered dataset.

## Technology Stack

| Layer | Tools |
|---|---|
| App framework | Streamlit |
| Data processing | Pandas, openpyxl |
| Mapping | Folium, streamlit-folium |
| Geocoding | Mapbox Geocoding API |
| AI summary/Q&A | Ollama (local LLM, `llama3.2:3b`) |
| Map export (optional) | Selenium, webdriver-manager, Pillow |

## Status

Maintained for internal use at DOST-Davao. Issues and pull requests are welcome.
