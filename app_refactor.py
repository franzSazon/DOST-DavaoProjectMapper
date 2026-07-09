"""
DOST-Davao Project Mapper

Streamlit app that ingests a DOST-Davao project Excel workbook, cleans and
consolidates it across divisions, and renders an interactive map with
filterable KPIs and per-project detail views.
"""

import pandas as pd
import streamlit as st
import random
import folium
from folium.features import DivIcon
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
import io
import requests
import time
import tempfile
try:
    from PIL import Image, ImageDraw, ImageFont
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    EXPORT_AVAILABLE = True
except ImportError:
    EXPORT_AVAILABLE = False

# ==========================================
# 1. CONSTANTS
# ==========================================
# Exact column headers as they appear in the source Excel workbook.
STATUS_BOOL_COL = "Project status"
GROUP_KEY_COL = "Project Title"
COORDINATES_COL = "Coordinates"
FUNDING_COL = "Amount of \nfunding provided"
DATE_APPROVED_COL = "Date of\nApproval"
DATE_END_COL = "Date of \nCompletion/\nTerminated\n(if terminated)"

# Sheets in the workbook that contain project data (others are ignored).
DATA_SHEETS = ["CEST", "LGIA", "SSCP"]

# Accepted TRUE/FALSE representations found in the raw sheet.
BOOL_MAP = {
    "TRUE": True, "FALSE": False,
    "True": True, "False": False,
    "1": True, "0": False,
    "1.0": True, "0.0": False,
}

STATUS_LABELS = ["Ongoing", "Completed", "Terminated"]

# Map styling.
DAVAO_CENTER = [7.06750309148034, 125.60060334232874]
DAVAO_ZOOM_START = 13
DIVISION_COLORS = {"CEST": "#2980b9", "LGIA": "#27ae60", "SSCP": "#8e44ad"}
DEFAULT_MARKER_COLOR = "#7f8c8d"
CLICK_TOLERANCE_DEGREES = 0.0001

MISSING_VALUE_TOKENS = {"nan", "nat", "none", "natype"}

# Local LLM API Configuration
OLLAMA_MODEL = "phi3"

# ==========================================
# 2. DATA PROCESSING PIPELINE
# ==========================================
def process_sheet(xls, sheet_name):
    """
    Read one division sheet and collapse its multi-row project blocks into
    one record per project, with one-hot status columns and parsed
    coordinates.

    Returns an empty DataFrame if the sheet doesn't match the expected
    format (e.g. an intro or cover sheet).
    """
    df = pd.read_excel(xls, sheet_name=sheet_name, header=1, dtype=str)
    df.columns = [str(col).strip() for col in df.columns]
    original_cols = list(df.columns)

    if STATUS_BOOL_COL not in original_cols:
        return pd.DataFrame()

    status_idx = original_cols.index(STATUS_BOOL_COL)
    status_label_col = original_cols[status_idx + 1]

    # Each project spans several rows; only the first row of a block has
    # the group key populated. Forward-fill the shared fields within a
    # block so every row carries the full project context.
    is_new_group = df[GROUP_KEY_COL].notna() & (df[GROUP_KEY_COL].astype(str).str.strip() != "")
    df["_group_id"] = is_new_group.cumsum()

    shared_cols = [c for c in original_cols if c not in (STATUS_BOOL_COL, status_label_col)]
    df[shared_cols] = df.groupby("_group_id")[shared_cols].ffill()
    df["_status_bool"] = df[STATUS_BOOL_COL].map(BOOL_MAP)

    records = []
    for _, group in df.groupby("_group_id"):
        base = group.iloc[0][shared_cols].to_dict()
        onehot = {label: 0 for label in STATUS_LABELS}

        for _, row in group.iterrows():
            label = str(row[status_label_col]).strip()
            if label in onehot and row["_status_bool"] is True:
                onehot[label] = 1

        base.update(onehot)
        records.append(base)

    result = pd.DataFrame(records)

    if COORDINATES_COL in result.columns:
        split_coords = result[COORDINATES_COL].astype(str).str.split(",", n=1, expand=True)
        result["Lat"] = pd.to_numeric(split_coords[0].str.strip(), errors="coerce")
        long_values = split_coords[1].str.strip() if split_coords.shape[1] > 1 else pd.NA
        result["Long"] = pd.to_numeric(long_values, errors="coerce")
    else:
        result["Lat"] = pd.NA
        result["Long"] = pd.NA

    result.insert(0, "Division", sheet_name)
    return result


def _assign_map_status(row):
    """Collapse the one-hot status columns into a single display label."""
    for label in STATUS_LABELS:
        if row.get(label) == 1:
            return label
    return "Unknown"


@st.cache_data
def load_and_clean_data(uploaded_file):
    """Load every recognized division sheet and combine them into one clean DataFrame."""
    xls = pd.ExcelFile(uploaded_file)
    cleaned_sheets = {}

    for sheet in xls.sheet_names:
        if sheet in DATA_SHEETS:
            cleaned = process_sheet(xls, sheet)
            if not cleaned.empty:
                cleaned_sheets[sheet] = cleaned

    combined = pd.concat(cleaned_sheets.values(), ignore_index=True)
    combined["Map_Status"] = combined.apply(_assign_map_status, axis=1)
    combined = combined.dropna(subset=["Lat", "Long"])

    return combined


# ==========================================
# 3. MAP GENERATION
# ==========================================
def _build_badge_html(abbrev, hex_color):
    """Build the HTML for a marker's abbreviation badge."""
    return f"""
    <div style="
        display: inline-flex;
        align-items: center;
        width: max-content;
        background-color: white;
        border: 2px solid {hex_color};
        border-radius: 12px;
        padding: 4px 8px;
        font-size: 11px;
        font-family: Arial, sans-serif;
        font-weight: bold;
        color: #2c3e50;
        box-shadow: 0px 3px 6px rgba(0,0,0,0.3);
        white-space: nowrap;
        cursor: pointer;
    ">
        <div style="width: 8px; height: 8px; border-radius: 50%; background-color: {hex_color}; margin-right: 5px;"></div>
        {abbrev}
    </div>
    """


def create_map(df):
    """Render the Davao region map with a clustered marker per project."""
    davao_map = folium.Map(location=DAVAO_CENTER, zoom_start=DAVAO_ZOOM_START)
    marker_cluster = MarkerCluster().add_to(davao_map)

    for _, row in df.iterrows():
        abbrev = str(row.get("Name Abbreviation", "DOST Project"))
        division = str(row.get("Division", "N/A"))
        hex_color = DIVISION_COLORS.get(division, DEFAULT_MARKER_COLOR)

        folium.Marker(
            location=[row["Lat"], row["Long"]],
            tooltip="📍 View project details",
            icon=DivIcon(
                icon_anchor=(0, 0),
                html=_build_badge_html(abbrev, hex_color),
                # Removes Leaflet's default icon size/overflow constraints
                # so the badge can render at its natural size.
                class_name="custom-badge",
            ),
        ).add_to(marker_cluster)

    return davao_map


# ==========================================
# 4. PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="DOST-Davao Project Mapper",
    page_icon="assets/dost_icon.png",
    layout="wide",
)

st.markdown("""
    <style>
 
        /* ==========================================
           1. FONT IMPORT
           ========================================== */
        /* Montserrat weights: 400=Regular, 700=Bold, 900=Black */
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
 
 
        /* ==========================================
           2. TYPOGRAPHY
           ========================================== */
        /* Scale the whole app down ~5%. Streamlit's spacing, padding, and
           font sizes are almost entirely rem-based, so shrinking the root
           font-size scales everything proportionally. This is used instead
           of transform: scale(), which clips iframes (the Folium map) and
           throws off click coordinates. */
        html {
            font-size: 90% !important;
        }

        /* Base font for the whole app */
        html, body, [class*="css"] {
            font-family: 'Montserrat', 'Arial', sans-serif !important;
            font-weight: 400 !important;
        }
 
        /* Sub-headers use Bold */
        h2, h3, h4, h5, h6 {
            font-family: 'Montserrat', 'Arial', sans-serif !important;
            font-weight: 700 !important;
        }
 
        /* Main H1 header uses Black (900 is the max valid CSS weight) */
        h1 {
            font-family: 'Montserrat', 'Arial', sans-serif !important;
            font-weight: 900 !important;
        }
 
 
        /* ==========================================
           3. LAYOUT SPACING
           ========================================== */
        /* Trim default top/bottom padding on the main content area */
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 0rem;
        }
 
        /* Remove top padding above sidebar content */
        [data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
            padding-top: 0rem;
        }
 
        /* Collapse the invisible sidebar header spacing */
        [data-testid="stSidebarHeader"] {
            padding-top: 0rem !important;
            padding-bottom: 0rem !important;
            min-height: 0px !important;
        }
 
        /* Belt-and-suspenders: force sidebar content padding to zero */
        [data-testid="stSidebarUserContent"] {
            padding-top: 0rem !important;
        }
 
        /* Pull sidebar image up to eat leftover space above it */
        [data-testid="stSidebar"] .stImage {
            margin-top: -2.5rem;
        }
 
 
        /* ==========================================
           4. FILE UPLOADER
           ========================================== */
        /* Force all text inside the uploader to white */
        [data-testid="stFileUploader"] div,
        [data-testid="stFileUploader"] span,
        [data-testid="stFileUploader"] small,
        [data-testid="stFileUploader"] label,
        [data-testid="stFileUploader"] p {
            color: #FFFFFF !important;
        }
 
        /* Dropzone as a column. align-items: stretch lets the uploaded-file
           row go full width (fixes filename truncation) — the instructions
           text and upload button opt out of stretching via align-self. */
        [data-testid="stFileUploaderDropzone"] {
            width: 100% !important;
            display: flex !important;
            flex-direction: column !important;
            align-items: stretch !important;
            justify-content: center !important;
        }
 
        /* Center the empty-state instructions text */
        [data-testid="stFileUploaderDropzoneInstructions"] {
            align-self: center !important;
            text-align: center !important;
            margin-bottom: -3.5rem !important;
        }
 
        /* Center the Upload button */
        [data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-secondary"] {
            width: 100% !important;
            align-self: center !important;
        }
 
        /* "Add files" button, present only when accept_multiple_files=True */
        [data-testid="stBaseButton-borderlessIcon"] {
            color: #FFFFFF !important;
            margin-top: -1.7rem !important;
        }
 
        [data-testid="stBaseButton-borderlessIcon"] [data-testid="stIconMaterial"] {
            color: transparent !important;
        }
 
        .st-emotion-cache-oeolxv { /* for svg icon in file uploader */
            background-color: #27343a !important;
            outline: 2px solid #1f292e !important;
        }

        .st-emotion-cache-pedct3 button { /* for 'x' icon in file uploader */   
            color: #27343a !important;
            background-color: #FFFFFF !important;
        }
            
        /* Let the uploaded filename take its natural width instead of truncating */
        [data-testid="stFileUploaderDropzone"] [data-testid="stFileUploaderFileName"] {
            max-width: none !important;
            overflow: visible !important;
            text-overflow: unset !important;
            white-space: normal !important;
            word-break: break-word !important;
        }
 
 
        /* ==========================================
           5. SIDEBAR TEXT COLOR
           ========================================== */
        /* Force all sidebar text elements to white */
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] span {
            color: #FFFFFF !important;
        }
 
        /* Metric values need a separate override */
        [data-testid="stSidebar"] [data-testid="stMetricValue"] {
            color: #FFFFFF !important;
        }
 
 
        /* ==========================================
           6. SIDEBAR ELEMENT SPACING
           ========================================== */
 
        /* Give the sidebar itself some horizontal padding so content
           doesn't sit flush against the edges */
        [data-testid="stSidebarUserContent"] {
            padding-left: 0.5rem !important;
            padding-right: 0.5rem !important;
        }
 
        /* ==========================================
           7. CHROME / UI CLEANUP
           ========================================== */
        header {background-color: transparent !important;}
        #MainMenu {visibility: hidden;}
        .stDeployButton {display: none;}

        .st-e3 {
            background-color: transparent !important;
        }
 
 
        # /* ==========================================
        #    8. LOCK SIDEBAR SCROLLING
        #    ========================================== */
        # /* Force the sidebar container to hide ALL overflowing content */
        # [data-testid="stSidebar"] {
        #     overflow-x: hidden !important;
        #     overflow-y: hidden !important;
        # }
 
        # /* Streamlit sometimes wraps sidebar content in an inner scrolling div */
        # [data-testid="stSidebar"] > div:first-child {
        #     overflow-x: hidden !important;
        #     overflow-y: hidden !important;
        # }
 
        # /* Hide the scrollbar visual across all browsers */
        # [data-testid="stSidebar"] ::-webkit-scrollbar {
        #     display: none !important;
        #     width: 0px !important;
        #     height: 0px !important;
        # }
 
        # [data-testid="stSidebar"] * {
        #     scrollbar-width: none !important; /* Firefox */
        #     -ms-overflow-style: none !important; /* IE and Edge */
        # }
 
 
        /* ==========================================
           9. MAP UI ENHANCEMENTS
           ========================================== */
        /* Target the Folium iframe to round corners and add a soft shadow */
        iframe {
            border-radius: 16px !important;
            box-shadow: -5px 6px 26px 0px rgba(0,0,0,0.1) !important;
            overflow: hidden !important;
        
        }
            
        /* ==========================================
           10. CHAT UI FIXES
           ========================================== */
        /* 1. Force the main chat input container to a crisp white with a DOST blue outline */
        [data-testid="stChatInput"] > div {
            color: #2C3E50 !important; /* Sets base text and caret color to dark gray */
            background-color: #FFFFFF !important;
            outline: 2px solid #00AEEF !important;
        }
            
        /* 2. Force the actual typed text to be readable */
        [data-testid="stChatInput"] textarea {
            color: #000000  !important; /* Dark gray typing text */
            background-color: #FFFFFF !important; /* Lets the white container show through safely */
        }

        /* 3. Style the 'Send' paper airplane icon to match your DOST Blue outline */
        [data-testid="stChatInput"] button svg {
            fill: #FFFFFF !important; /* Blue icon instead of invisible white */
        }
            
        /* 4. Completely un-anchor the chat input wrapper from the bottom */
        [data-testid="stBottomBlock"], 
        [data-testid="stBottom"],
        .stChatFloatingInputContainer,
        [data-testid="stChatInput"] {
            position: static !important;
            bottom: auto !important;
            background-color: transparent !important;
        }


        /* ==========================================
           11. CHAT AVATAR COLORS
           ========================================== */
        /* User messages: dark slate */
        [data-testid="stChatMessageAvatarUser"] {
            background-color: #222D32 !important;
        }

        /* Assistant messages: DOST blue */
        [data-testid="stChatMessageAvatarAssistant"] {
            background-color: #00AEEF !important;
        }

        /* ==========================================
           12. SIDEBAR ELEMENTS COLOR FIX
           ========================================== */
        .st-emotion-cache-yiekhv {
            background-color: #FFFFFF !important;
        }

        .st-f3 {
            background-color: transparent !important;
        }

        .st-fg {
            background-color: #FFFFFF !important;
        }

        /* ==========================================
           13. SLIDER TWEAKS
           ========================================== */
        /* Hide the min/max bounds that fade in and out on the slider */
        [data-testid="stTickBar"] {
            display: none !important;
        }

        /* ==========================================
           14. FULLSCREEN MAP DIALOG OVERRIDE
           ========================================== */
        /* Force the Fullscreen Map View dialog to exceed Streamlit's "large" max-width */
        div[data-testid="stDialog"] [role="dialog"] {
            width: 95vw !important;
            max-width: 95vw !important;
            margin-top: -2.5vh !important; /* Pull the dialog up higher */
            align-self: flex-start !important; /* Anchor it to the top if it's a flex container */
        }

    </style>
""", unsafe_allow_html=True)
 
st.markdown("### DOST-Davao Project Mapper 📍")

# ==========================================
# 5. HELPERS FOR DISPLAY FORMATTING
# ==========================================
def clean_missing(value, fallback="TBA / Not specified"):
    """Replace pandas' NaN/NaT text representations with a readable fallback."""
    text = str(value)
    return fallback if text.lower() in MISSING_VALUE_TOKENS else text


def format_currency(amount):
    """Format a peso amount, abbreviating to millions when large."""
    if amount >= 1_000_000:
        return f"₱{amount / 1_000_000:.2f}M"
    return f"₱{amount:,.0f}"


# ==========================================
# 6. MODAL DIALOGS
# ==========================================
def render_project_details_content(clicked_projects):
    st.success(f"Found {len(clicked_projects)} project(s) at this location.")
    st.divider()

    for _, row in clicked_projects.iterrows():
        title = str(row.get(GROUP_KEY_COL, "Unnamed Project"))
        division = str(row.get("Division", "N/A"))
        status = str(row.get("Map_Status", "Unknown"))
        
        funding_raw = row.get(FUNDING_COL, "Not specified")
        try:
            funding_val = float(str(funding_raw).replace(',', ''))
            display_funding = f"₱{funding_val:,.2f}" if funding_val % 1 != 0 else f"₱{int(funding_val):,}"
        except ValueError:
            display_funding = f"₱{funding_raw}" if funding_raw != "Not specified" else "Not specified"

        agency = row.get("Implementing Agency", "Not specified")
        proponent = str(row.get("Name of Proponent", "Not specified"))
        remarks = str(row.get("Remarks", "No remarks provided."))

        date_approved = clean_missing(row.get(DATE_APPROVED_COL, "Not specified"))
        date_end = clean_missing(row.get(DATE_END_COL, "Not specified"))

        st.header(title)
        st.subheader(f"**Division:** {division} | **Status:** {status}")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**🏢 Agency:** {agency}")
            st.markdown(f"**👤 Proponent:** {proponent}")
            st.markdown(f"**💰 Funding:** {display_funding}")
        with col2:
            # [:10] strips any trailing " 00:00:00" timestamp from date strings.
            st.markdown(f"**📅 Date Approved:** {date_approved[:10]}")
            st.markdown(f"**🏁 End/Target Date:** {date_end[:10]}")

        st.info(f"**Remarks:** {remarks}")
        st.divider()


@st.dialog("📋 Located Project Details", width="large")
def show_project_details(clicked_projects):
    """Display full details for every project at a clicked map location."""
    render_project_details_content(clicked_projects)


@st.dialog("📊 Raw Data Explorer", width="large")
def show_raw_data(df):
    """Display the filtered dataset in a plain, sortable table."""
    st.caption(f"Currently viewing {len(df)} filtered projects.")
    st.dataframe(df, width="stretch")


@st.dialog("🗺️ Fullscreen Map View", width="large")
def show_fullscreen_map(df):
    """Render the map in a large modal window for better visibility."""
    project_map = create_map(df)
    map_data = st_folium(
        project_map,
        use_container_width=True,
        height=570,
        returned_objects=["last_object_clicked"],
        key="fullscreen_map"
    )
    
    clicked = map_data.get("last_object_clicked")
    if clicked:
        click_key = (round(clicked["lat"], 6), round(clicked["lng"], 6))
        
        # Only show the spinner transition if the user clicked a NEW marker
        if st.session_state.get("fullscreen_last_click") != click_key:
            st.session_state["fullscreen_last_click"] = click_key
            with st.spinner("Fetching project details..."):
                time.sleep(3.5)
                
        clicked_projects = df[
            (abs(df["Lat"] - clicked["lat"]) < CLICK_TOLERANCE_DEGREES)
            & (abs(df["Long"] - clicked["lng"]) < CLICK_TOLERANCE_DEGREES)
        ]
        if not clicked_projects.empty:
            st.markdown("### 📋 Clicked Location Details")
            render_project_details_content(clicked_projects)


@st.dialog("🖼️ Map Export Options", width="large")
def show_export_dialog(df):
    """Configure, preview, and download the high-res map export."""
    if not EXPORT_AVAILABLE:
        st.warning("⚠️ High-Res Map Export is currently disabled. Please install `selenium`, `webdriver-manager`, and `pillow` to enable this feature.")
        return

    st.write("Configure your map export settings below:")
    
    col1, col2 = st.columns(2)
    with col1:
        zoom_level = st.slider("Zoom Level", min_value=8, max_value=18, value=13)
    with col2:
        title = st.text_input("Report Title", value="DOST-Davao Regional Project Mapping Report")
        
    if st.button("Generate Preview", width="stretch"):
        with st.spinner("Capturing map and generating report (this may take a few seconds)..."):
            # Create a dedicated map for export with the selected zoom and use the fixed DAVAO_CENTER
            export_map = folium.Map(location=DAVAO_CENTER, zoom_start=zoom_level, zoom_control=False)
            marker_cluster = MarkerCluster().add_to(export_map)

            for _, row in df.iterrows():
                abbrev = str(row.get("Name Abbreviation", "DOST Project"))
                division = str(row.get("Division", "N/A"))
                hex_color = DIVISION_COLORS.get(division, DEFAULT_MARKER_COLOR)

                folium.Marker(
                    location=[row["Lat"], row["Long"]],
                    icon=DivIcon(
                        icon_anchor=(0, 0),
                        html=_build_badge_html(abbrev, hex_color),
                        class_name="custom-badge",
                    ),
                ).add_to(marker_cluster)
            
            # Hide scrollbar in the export map
            export_map.get_root().header.add_child(folium.Element("<style>body, html { margin:0; padding:0; overflow: hidden !important; }</style>"))
            
            raw_image = capture_map_screenshot(export_map)
            
            active_divs = df["Division"].unique().tolist()
            active_stats = df["Map_Status"].unique().tolist()
            
            branded_bytes = create_branded_export(raw_image, active_divs, active_stats, title)
            
            # Store in session_state to persist between interactions in dialog
            st.session_state["export_preview_bytes"] = branded_bytes

    if "export_preview_bytes" in st.session_state:
        st.image(st.session_state["export_preview_bytes"], width="stretch")
        st.download_button(
            label="📥 Download High-Res Map (PNG)",
            data=st.session_state["export_preview_bytes"],
            file_name="DOST_Davao_Map_Export.png",
            mime="image/png",
            width="stretch"
        )


# ==========================================
# 7. SIDEBAR: BRANDING & FILE UPLOAD
# ==========================================
st.sidebar.image("assets/dost_davao_logo.png", width="stretch")

uploaded_file = st.sidebar.file_uploader(
    "Upload DOST-Davao Excel File (.xlsx)",
    type=["xlsx"],
)


# ==========================================
# 8. SIDEBAR: FILTERS & KPI DASHBOARD
# ==========================================
def format_export_data(df):
    """Clean and reorder the dataframe for stakeholder presentation."""
    export_df = df.copy()
    
    # 1. Drop internal logic & redundant columns
    cols_to_drop = [COORDINATES_COL, "_group_id", "_status_bool"] + STATUS_LABELS
    export_df = export_df.drop(columns=[c for c in cols_to_drop if c in export_df.columns])
    
    # 2. Rename columns for professionalism
    rename_map = {
        "Map_Status": "Status",
        "Lat": "Latitude",
        "Long": "Longitude"
    }
    export_df = export_df.rename(columns=rename_map)
    
    # 3. Reorder: Move Status, Remarks, Lat, Long after the Date columns
    current_cols = list(export_df.columns)
    
    # Pull out our target columns so we can safely reposition them
    target_order = ["Status", "Remarks", "Latitude", "Longitude"]
    for col in target_order:
        if col in current_cols:
            current_cols.remove(col)
            
    # Find the index of the Date columns to slice the list
    insert_idx = len(current_cols) # Default to the end
    if DATE_APPROVED_COL in current_cols:
        insert_idx = current_cols.index(DATE_APPROVED_COL) + 1
        
    # If the End Date column exists, make sure we insert *after* it
    if DATE_END_COL in current_cols:
        insert_idx = max(insert_idx, current_cols.index(DATE_END_COL) + 1)
        
    # Reassemble the final column order
    final_cols = current_cols[:insert_idx] + target_order + current_cols[insert_idx:]
    
    # Ensure we only include columns that actually exist in the df to prevent KeyError
    final_cols = [c for c in final_cols if c in export_df.columns]
    
    return export_df[final_cols]


def render_filters(clean_df):
    """Render the sidebar filter controls and return the filtered DataFrame."""
    st.sidebar.header("🔍 Filter Dashboard")

    available_divisions = clean_df["Division"].unique().tolist()
    selected_divisions = st.sidebar.multiselect(
        "Select Project(s):", available_divisions, default=available_divisions
    )

    available_statuses = clean_df["Map_Status"].unique().tolist()
    selected_statuses = st.sidebar.multiselect(
        "Select Project Status:", available_statuses, default=available_statuses
    )

    filtered_df = clean_df[
        clean_df["Division"].isin(selected_divisions)
        & clean_df["Map_Status"].isin(selected_statuses)
    ]

    if DATE_APPROVED_COL in clean_df.columns:
        approval_dates = pd.to_datetime(clean_df[DATE_APPROVED_COL], errors="coerce")
        valid_dates = approval_dates.dropna()
        
        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
        else:
            min_date = pd.Timestamp("2010-01-01").date()
            max_date = pd.Timestamp.today().date()
        
        date_range = st.sidebar.date_input(
            "Filter by Approval Date:",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )
        
        include_missing_dates = st.sidebar.checkbox("Include projects with missing dates", value=True)

        if isinstance(date_range, tuple):
            if len(date_range) == 2:
                start_dt = pd.to_datetime(date_range[0])
                end_dt = pd.to_datetime(date_range[1])
                filtered_approval_dates = pd.to_datetime(filtered_df[DATE_APPROVED_COL], errors="coerce")
                
                date_mask = (filtered_approval_dates >= start_dt) & (filtered_approval_dates <= end_dt)
                if include_missing_dates:
                    date_mask = date_mask | filtered_approval_dates.isna()
                    
                filtered_df = filtered_df[date_mask]
            elif len(date_range) == 1:
                start_dt = pd.to_datetime(date_range[0])
                filtered_approval_dates = pd.to_datetime(filtered_df[DATE_APPROVED_COL], errors="coerce")
                
                date_mask = (filtered_approval_dates >= start_dt)
                if include_missing_dates:
                    date_mask = date_mask | filtered_approval_dates.isna()
                    
                filtered_df = filtered_df[date_mask]

    if FUNDING_COL in clean_df.columns:
        funding_numeric = pd.to_numeric(
            clean_df[FUNDING_COL].astype(str).str.replace(",", "", regex=False),
            errors="coerce"
        )
        valid_funding = funding_numeric.dropna()
        
        if not valid_funding.empty:
            min_budget = int(valid_funding.min()) - 1
            max_budget = int(valid_funding.max()) + 1
            
            if max_budget > min_budget:
                budget_range = st.sidebar.slider(
                    "Filter by Budget (₱):",
                    min_value=min_budget,
                    max_value=max_budget,
                    value=(min_budget, max_budget),
                    step=1000,
                )
                
                include_missing_budget = st.sidebar.checkbox("Include projects with missing budget", value=True)
                
                filtered_funding = pd.to_numeric(
                    filtered_df[FUNDING_COL].astype(str).str.replace(",", "", regex=False),
                    errors="coerce"
                )
                
                budget_mask = (filtered_funding >= budget_range[0]) & (filtered_funding <= budget_range[1])
                if include_missing_budget:
                    budget_mask = budget_mask | filtered_funding.isna()
                    
                filtered_df = filtered_df[budget_mask]

    st.sidebar.markdown("---")
    st.sidebar.metric(label="Projects Displayed", value=len(filtered_df))

    # --- NEW: Generate Presentation-Ready Data ---
    presentation_df = format_export_data(filtered_df)

    if st.sidebar.button("📄 View Raw Data Table", width="stretch"):
        show_raw_data(presentation_df) # Pass the clean data to the modal

    # --- Task 4.2: Filtered CSV Export ---
    csv_data = presentation_df.to_csv(index=False).encode('utf-8')
    
    st.sidebar.download_button(
        label="📥 Download Filtered Data (CSV)",
        data=csv_data,
        file_name="DOST_Davao_Filtered_Projects.csv",
        mime="text/csv",
        width="stretch" 
    )

    if st.sidebar.button("🖼️ Map Export Options", width="stretch"):
        st.session_state.pop("export_preview_bytes", None) # clear old preview
        show_export_dialog(filtered_df)

    return filtered_df


def render_kpi_scorecards(filtered_df):
    """Render the five-column KPI summary above the map."""
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

    with kpi1:
        st.metric(label="Total Projects", value=len(filtered_df))

    with kpi2:
        ongoing_count = len(filtered_df[filtered_df["Map_Status"] == "Ongoing"])
        st.metric(label="🚀 Ongoing", value=ongoing_count)

    with kpi3:
        completed_count = len(filtered_df[filtered_df["Map_Status"] == "Completed"])
        st.metric(label="✅ Completed", value=completed_count)

    with kpi4:
        terminated_count = len(filtered_df[filtered_df["Map_Status"] == "Terminated"])
        st.metric(label="🛑 Terminated", value=terminated_count)

    with kpi5:
        if FUNDING_COL in filtered_df.columns:
            total_funding = pd.to_numeric(
                filtered_df[FUNDING_COL].astype(str).str.replace(",", "", regex=False),
                errors="coerce",
            ).sum()
            formatted_funding = format_currency(total_funding)
        else:
            formatted_funding = "N/A"

        st.metric(label="💰 Total Funding", value=formatted_funding)


def handle_map_click(map_data, filtered_df):
    """
    Open the project detail modal for whatever was clicked on the map.

    st_folium re-returns the same last_object_clicked value on every rerun
    until a new marker is clicked, so the click is tracked in session_state
    to avoid re-opening the dialog on unrelated reruns (e.g. a sidebar
    button press, which would otherwise collide with an already-open
    dialog and raise a StreamlitAPIException).
    """
    clicked = map_data.get("last_object_clicked")
    if not clicked:
        return

    click_key = (round(clicked["lat"], 6), round(clicked["lng"], 6))
    if click_key == st.session_state.get("last_handled_click"):
        return

    st.session_state["last_handled_click"] = click_key

    clicked_projects = filtered_df[
        (abs(filtered_df["Lat"] - clicked["lat"]) < CLICK_TOLERANCE_DEGREES)
        & (abs(filtered_df["Long"] - clicked["lng"]) < CLICK_TOLERANCE_DEGREES)
    ]

    if not clicked_projects.empty:
        show_project_details(clicked_projects)

# ==========================================
# 9. AI ASSISTANT & Q&A LOGIC (LOCAL PRIVACY VIA OLLAMA)
# ==========================================

def get_ollama_url():
    """
    Retrieve the active Ngrok tunnel link from Streamlit secrets.

    Returns the string URL if configured, otherwise returns None.
    This replaces the previous get_gemini_client method to ensure
    absolute data privacy locally on your hardware.
    """
    if "OLLAMA_URL" in st.secrets:
        return st.secrets["OLLAMA_URL"]
    return None


def generate_executive_summary(ollama_url, df):
    """Generate a stakeholder-ready executive summary using the local Phi-3 model."""
    clean_df = format_export_data(df)
    
    # Generate high-level summary statistics to avoid context-overflow
    total_projects = len(clean_df)
    status_counts = clean_df['Status'].value_counts().to_dict() if 'Status' in clean_df.columns else {}
    
    # Attempt to find a funding column to sum
    funding_col = next((c for c in clean_df.columns if 'fund' in c.lower() or 'cost' in c.lower() or 'amount' in c.lower()), None)
    total_funding = 0
    if funding_col:
        total_funding = pd.to_numeric(clean_df[funding_col].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').sum()

    div_col = next((c for c in clean_df.columns if 'div' in c.lower() or 'unit' in c.lower()), None)
    div_counts = clean_df[div_col].value_counts().to_dict() if div_col else {}

    summary_text = f"""
    Total Projects: {total_projects}
    Projects by Status: {status_counts}
    Total Funding: Php {total_funding:,.2f}
    Projects by Division: {div_counts}
    """

    # Limit rows to avoid huge prompt payloads that slow down local models
    MAX_ROWS = 50
    sample_df = clean_df.head(MAX_ROWS)
    csv_data = sample_df.to_csv(index=False)
    
    row_notice = ""
    if len(clean_df) > MAX_ROWS:
        row_notice = f"(Note: The CSV data below has been truncated to the top {MAX_ROWS} rows out of {len(clean_df)} to improve AI speed. Use the Data Summary above for accurate totals.)"

    prompt = f"""
    Please write a 2-paragraph executive summary based ONLY on the following regional project data summary and CSV sample.
    Focus on highlighting key funding allocations, division focus, and overall project health (Ongoing vs Completed vs Terminated).
    Do not use external knowledge.

    Use markdown to format your response neatly using bullets, highlighting, and headers. 
    Do not use h1 (#) or h2 (##) headers, only use h3 (###) if you want to use a header.

    Data Summary (Use these for overall totals):
    {summary_text}

    CSV Data Sample {row_notice}:
    {csv_data}
    """

    payload = {"model": OLLAMA_MODEL, "prompt": prompt, "stream": False}

    # Crucial headers to bypass Ngrok free-tier interceptor page
    headers = {
        "Content-Type": "application/json",
        "ngrok-skip-browser-warning": "true",
    }

    try:
        response = requests.post(
            f"{ollama_url}/api/generate", json=payload, headers=headers, timeout=600
        )
        response.raise_for_status()
        return response.json().get("response", "No response generated.")
    except Exception as e:
        return f"⚠️ Error generating local summary: {str(e)}"


def ask_ai_about_data(ollama_url, df, user_query, chat_history):
    """Answer natural-language data questions using the local Phi-3 instance with conversation history."""
    clean_df = format_export_data(df)
    
    # Limit rows to avoid huge prompt payloads that slow down local models
    MAX_ROWS = 50
    sample_df = clean_df.head(MAX_ROWS)
    csv_data = sample_df.to_csv(index=False)
    
    row_notice = ""
    if len(clean_df) > MAX_ROWS:
        row_notice = f"(Note: The data has been truncated to the top {MAX_ROWS} rows out of {len(clean_df)} to improve AI speed. Summarize based on this sample.)"

    # Format historical turns for context mapping
    history_text = "\n".join(
        [f"{msg['role'].capitalize()}: {msg['content']}" for msg in chat_history]
    )

    prompt = f"""
    You are an AI data assistant for the DOST-Davao Regional Office.
    Here is the currently filtered DOST project data in CSV format {row_notice}:

    {csv_data}

    User Query: {user_query}

    Based ONLY on the provided CSV data and the conversation history, answer the user's query.
    Be concise and professional, avoid using emdashes and just be functional with your tone.

    Use markdown to format your response in a neat way using bullets, 
    highlighting and headers, do not use h1 or h2, only use h3 if you want to use a header.
    """

    payload = {"model": OLLAMA_MODEL, "prompt": prompt, "stream": False}

    headers = {
        "Content-Type": "application/json",
        "ngrok-skip-browser-warning": "true",
    }

    try:
        response = requests.post(
            f"{ollama_url}/api/generate", json=payload, headers=headers, timeout=600
        )
        response.raise_for_status()
        return response.json().get("response", "No response generated.")
    except Exception as e:
        return f"⚠️ Error communicating with Local AI: {str(e)}"

# ==========================================
# 10. MAP EXPORT & BRANDING
# ==========================================
def capture_map_screenshot(folium_map):
    """Save folium map to HTML and capture a screenshot using headless Chrome."""
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--hide-scrollbars')
    
    # Try to use webdriver-manager, fallback to basic Chrome if it fails
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception:
        driver = webdriver.Chrome(options=options)
        
    try:
        with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as tmp:
            html_path = tmp.name
            
            # Hide scrollbars specifically for the screenshot and force full width to avoid gaps
            folium_map.get_root().header.add_child(folium.Element("<style>body, html { margin:0 !important; padding:0 !important; overflow: hidden !important; width: 100vw !important; height: 100vh !important; } .folium-map { width: 100vw !important; height: 100vh !important; }</style>"))
            folium_map.save(html_path)
            
        driver.get(f"file:///{html_path}")
        time.sleep(3) # Wait for map tiles and clusters to load
        
        png_data = driver.get_screenshot_as_png()
        return Image.open(io.BytesIO(png_data))
    finally:
        driver.quit()
        # NEW: Delete the temporary HTML file so the server disk doesn't fill up
        try:
            os.remove(html_path)
        except Exception:
            pass

def create_branded_export(map_image, selected_divisions, selected_statuses, title_text="DOST-Davao Regional Project Mapping Report"):
    """Combine the map screenshot with DOST branding and active filter details."""
    canvas_width = map_image.width
    canvas_height = map_image.height + 250
    
    # Create canvas
    canvas = Image.new('RGB', (canvas_width, canvas_height), '#222d32')
    draw = ImageDraw.Draw(canvas)
    
    # Paste Map at the bottom
    canvas.paste(map_image, (0, 250))
    
    text_x = 250
    # Try to load and paste logo
    try:
        logo = Image.open("assets/dost_davao_logo.png").convert("RGBA")
        # Resize logo to fit nicely in header (e.g. max height 150)
        ratio = 150 / logo.height
        new_width = int(logo.width * ratio)
        logo = logo.resize((new_width, 150))
        # Paste at left, centered vertically in the 250px header
        canvas.paste(logo, (50, 50), logo)
        
        # Shift text to the right of the logo with a 50px buffer
        text_x = 40 + new_width + 40
    except Exception:
        pass
        
    # Attempt to load a default font, fallback to default PIL font
    try:
        font_title = ImageFont.truetype("assets/Montserrat-Bold.ttf", 50)
        font_subtitle = ImageFont.truetype("assets/Montserrat-Regular.ttf", 36)
    except Exception:
        font_title = ImageFont.load_default()
        font_subtitle = ImageFont.load_default()
        
    # Text Placement
    div_str = "All" if len(selected_divisions) > 2 else ", ".join(selected_divisions)
    stat_str = "All" if len(selected_statuses) > 2 else ", ".join(selected_statuses)
    subtitle_text = f"Divisions: {div_str} | Status: {stat_str}"
    
    draw.text((text_x, 60), title_text, font=font_title, fill=(255, 255, 255))
    draw.text((text_x, 135), subtitle_text, font=font_subtitle, fill=(255, 255, 255))
    
    # Export to bytes
    img_byte_arr = io.BytesIO()
    canvas.save(img_byte_arr, format='PNG')
    img_byte_arr.seek(0)
    
    return img_byte_arr.getvalue()

# ==========================================
# 11. MAIN APPLICATION LOGIC
# ==========================================
if uploaded_file is not None:
    with st.spinner("Processing..."):
        clean_df = load_and_clean_data(uploaded_file)

    filtered_df = render_filters(clean_df)
    render_kpi_scorecards(filtered_df)

    project_map = create_map(filtered_df)
    map_data = st_folium(
        project_map,
        use_container_width=True,
        height=460,
        returned_objects=["last_object_clicked"],
    )

    handle_map_click(map_data, filtered_df)

    col1, col2 = st.columns([0.85, 0.15])
    with col2:
        if st.button("🔍 Maximize Map", use_container_width=True):
            show_fullscreen_map(filtered_df)


    # (Removed original high-res export block from main body)

    # ==========================================
    # TASKS 5.2, 5.3, 5.4: AI CHAT INTERFACE
    # ==========================================
    st.markdown("---")
    
    # Use columns to put the Clear Chat button neatly on the right side
    head_col1, head_col2 = st.columns([4, 1])
    with head_col1:
        st.markdown("### 🗾 DOST-Davao Project Mapper AI")
    
    ollama_url = get_ollama_url()
    if ollama_url:
        with head_col2:
            if st.button("🗑️ Clear Chat", width="stretch"):
                st.session_state.chat_messages = [
                    {"role": "assistant", "content": "Hello! I am your DOST-Davao Project Mapper Data Assistant. How can I help you analyze the map data today?"}
                ]
                st.rerun()

        if st.button("📊 Generate Executive Summary for Current Map", width="stretch"):
            with st.spinner("Drafting executive summary..."):
                summary = generate_executive_summary(ollama_url, filtered_df)
                st.info(summary)
        
        st.caption("Ask natural language questions about the projects currently visible on the map.")
        
        # Initialize chat history in Streamlit Session State
        if "chat_messages" not in st.session_state:
            st.session_state.chat_messages = [
                {"role": "assistant", "content": "Hello! I am your DOST-Davao Project Mapper Data Assistant. How can I help you analyze the map data today?"}
            ]

        # Render existing chat messages
        for message in st.session_state.chat_messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        # Chat input and AI response
        phrase_list = ["What is the total funding for ongoing " + random.choice(DATA_SHEETS) + " projects?",
                        "What are the date of completion for ongoing " + random.choice(DATA_SHEETS) + " projects?",
                        "What are the date of approval for ongoing " + random.choice(DATA_SHEETS) + " projects?",
                        "What are the funding amounts for ongoing " + random.choice(DATA_SHEETS) + " projects?",
                        "What are the project titles for " + random.choice(DATA_SHEETS) + " projects?",
                        "What are the project abbreviation for " + random.choice(DATA_SHEETS) + " projects?",
                        "What are the project division for " + random.choice(DATA_SHEETS) + " projects?",
                        "What are the project coordinates for " + random.choice(DATA_SHEETS) + " projects?",
                        ]

        # Initialize random placeholder in session state if it doesn't exist
        if "chat_placeholder" not in st.session_state:
            st.session_state.chat_placeholder = f"Ask a question about the current visible on the map. Ex. {random.choice(phrase_list)}"

        # Use the static placeholder from session state
        if prompt := st.chat_input(st.session_state.chat_placeholder):
            
            # 1. Display user message
            with st.chat_message("user"):
                st.markdown(prompt)
                
            # 2. Get the AI response
            with st.chat_message("assistant"):
                with st.spinner("Analyzing map data..."):
                    ai_response = ask_ai_about_data(ollama_url, filtered_df, prompt, st.session_state.chat_messages)
                    st.markdown(ai_response)
            
            # 3. Append both to session state
            st.session_state.chat_messages.append({"role": "user", "content": prompt})
            st.session_state.chat_messages.append({"role": "assistant", "content": ai_response})

    else:
        st.error("🔑 Ollama URL not found. Please add OLLAMA_URL to your `.streamlit/secrets.toml` file.")

else:
    st.info("Upload the raw Excel file in the sidebar to begin.")