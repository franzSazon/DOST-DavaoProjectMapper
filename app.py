import streamlit as st
import pandas as pd
import folium
from folium.features import DivIcon
from streamlit_folium import st_folium
from folium.plugins import MarkerCluster

# Define constants based on the exact Excel headers
STATUS_BOOL_COL = "Project status"
GROUP_KEY_COL = "Project Title"
COORDINATES_COL = "Coordinates"

# --- 1. DATA PROCESSING PIPELINE ---
def process_sheet(xls, sheet_name):
    # Read directly from the ExcelFile object
    df = pd.read_excel(xls, sheet_name=sheet_name, header=1, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    original_cols = list(df.columns)
    
    # Safely find the status index
    if STATUS_BOOL_COL not in original_cols:
        return pd.DataFrame() # Skip if sheet doesn't match format
        
    status_idx = original_cols.index(STATUS_BOOL_COL)
    status_label_col = original_cols[status_idx + 1] 

    is_new_group = df[GROUP_KEY_COL].notna() & (df[GROUP_KEY_COL].astype(str).str.strip() != "")
    df["_group_id"] = is_new_group.cumsum()

    shared_cols = [c for c in original_cols if c not in (STATUS_BOOL_COL, status_label_col)]
    df[shared_cols] = df.groupby("_group_id")[shared_cols].ffill()

    bool_map = {"TRUE": True, "FALSE": False, "True": True, "False": False, "1": True, "0": False, "1.0": True, "0.0": False}
    df["_status_bool"] = df[STATUS_BOOL_COL].map(bool_map)

    records = []
    for gid, g in df.groupby("_group_id"):
        base = g.iloc[0][shared_cols].to_dict()
        onehot = {"Ongoing": 0, "Completed": 0, "Terminated": 0}
        for _, row in g.iterrows():
            label = str(row[status_label_col]).strip()
            if label in onehot and row["_status_bool"] is True:
                onehot[label] = 1
        base.update(onehot)
        records.append(base)

    result = pd.DataFrame(records)

    has_coords = COORDINATES_COL in result.columns
    if has_coords:
        split_coords = result[COORDINATES_COL].astype(str).str.split(",", n=1, expand=True)
        result["Lat"] = pd.to_numeric(split_coords[0].str.strip(), errors='coerce')
        result["Long"] = pd.to_numeric(split_coords[1].str.strip() if split_coords.shape[1] > 1 else pd.NA, errors='coerce')
    else:
        result["Lat"] = pd.NA
        result["Long"] = pd.NA

    # Tag which program/sheet this project came from
    result.insert(0, "Division", sheet_name) 
    return result

@st.cache_data
def load_and_clean_data(uploaded_file):
    # Load the uploaded file into a pandas ExcelFile object
    xls = pd.ExcelFile(uploaded_file)
    cleaned_sheets = {}

    for sheet in xls.sheet_names:
        # We can dynamically skip sheets that aren't data (like an intro sheet) if needed
        if sheet in ["CEST", "LGIA", "SSCP"]: 
            cleaned = process_sheet(xls, sheet)
            if not cleaned.empty:
                cleaned_sheets[sheet] = cleaned

    # Combine all divisions
    combined = pd.concat(cleaned_sheets.values(), ignore_index=True)
    
    # Create a unified Status column for the map pins
    def get_status(row):
        if row.get('Ongoing') == 1: return 'Ongoing'
        if row.get('Completed') == 1: return 'Completed'
        if row.get('Terminated') == 1: return 'Terminated'
        return 'Unknown'
        
    combined['Map_Status'] = combined.apply(get_status, axis=1)
    
    # Drop rows without valid coordinates for mapping
    combined = combined.dropna(subset=['Lat', 'Long'])
    
    return combined

# --- 2. MAP GENERATION (With Abbreviation Badges FIXED) ---
def create_map(df):
    # Center map on Davao Region
    davao_map = folium.Map(location=[7.06750309148034, 125.60060334232874], zoom_start=13)
    
    # Define colors based on division
    color_map = {'CEST': '#2980b9', 'LGIA': '#27ae60', 'SSCP': '#8e44ad'}
    
    # Initialize the cluster layer
    marker_cluster = MarkerCluster().add_to(davao_map)
    
    for _, row in df.iterrows():
        # Safely fetch the abbreviation
        abbrev = str(row.get('Name Abbreviation', 'DOST Project'))
        division = str(row.get('Division', 'N/A'))
        hex_color = color_map.get(division, '#7f8c8d') 
        
        # --- FIXED CSS ---
        html_badge = f"""
        <div style="
            display: inline-flex; 
            align-items: center;
            width: max-content; /* MAGIC FIX 1: Forces background to stretch */
            background-color: white;
            border: 2px solid {hex_color};
            border-radius: 12px;
            padding: 4px 8px; /* Slightly thicker padding for breathing room */
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
        
        # Place the custom DivIcon
        folium.Marker(
            location=[row['Lat'], row['Long']],
            tooltip="📍 View project details",
            icon=DivIcon(
                icon_anchor=(0, 0), 
                html=html_badge,
                class_name="custom-badge" # MAGIC FIX 2: Removes default Leaflet restrictions
            )
        ).add_to(marker_cluster)
        
    return davao_map

# --- 3. STREAMLIT APP UI ---

# 1. Update page config to include the Favicon (Browser Tab Icon)
st.set_page_config(
    page_title="DOST-Davao Project Mapper", 
    page_icon="assets/dost_icon.png", # Use forward slashes!
    layout="wide"
)

# 2. CSS Hack for Layout, Typography, and Sidebar Spacing
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
            padding-top: 1rem;
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

        /* Dropzone as a column. align-items: stretch lets the uploaded-file row
        go full width (fixes filename truncation) — Upload button and
        instructions text opt out of stretching below via align-self. */
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

        /* Center the Upload button — correct testid has the "st" prefix */
        [data-testid="stFileUploaderDropzone"] [data-testid="stBaseButton-secondary"] {
            width: 100% !important;
            align-self: center !important;
        }

        /* "Add files" button (only present when accept_multiple_files=True) */
        [data-testid="stBaseButton-borderlessIcon"] {
            color: #FFFFFF !important;
            margin-top: 0rem !important;
        }
        [data-testid="stBaseButton-borderlessIcon"] [data-testid="stIconMaterial"] {
            color: #FFFFFF !important;
        }

        /* Let the uploaded filename take its natural width instead of truncating */
        [data-testid="stFileUploaderDropzone"] [data-testid="stFileUploaderFileName"] {
            max-width: none !important;
            overflow: visible !important;
            text-overflow: unset !important;
            white-space: normal !important;
            word-break: break-word !important;
        }
        
        /* "Add files" button (only present when accept_multiple_files=True) */
        [data-testid="stBaseButton-borderlessIcon"] {
            color: #FFFFFF !important;
        }

        /* Optional: give it some breathing room from the file card above it */
        [data-testid="stBaseButton-borderlessIcon"] {
            margin-top: -1.7rem !important;
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
           6. CHROME / UI CLEANUP
           ========================================== */
        header {background-color: transparent !important;}
        #MainMenu {visibility: hidden;}
        .stDeployButton {display: none;}
            
        /* ==========================================
           7. LOCK SIDEBAR SCROLLING
           ========================================== */
        /* Force the sidebar container to hide ALL overflowing content */
        [data-testid="stSidebar"] {
            overflow-x: hidden !important;
            overflow-y: hidden !important;
        }
        
        /* Sometimes Streamlit wraps the sidebar content in an inner div that scrolls */
        [data-testid="stSidebar"] > div:first-child {
            overflow-x: hidden !important;
            overflow-y: hidden !important;
        }

        /* Completely hide the scrollbar visual across all browsers */
        [data-testid="stSidebar"] ::-webkit-scrollbar {
            display: none !important;
            width: 0px !important;
            height: 0px !important;
        }
        
        [data-testid="stSidebar"] * {
            scrollbar-width: 10px !important; /* Firefox (Fixed from 10px) */
            -ms-overflow-style: none !important; /* IE and Edge */
        }

        /* ==========================================
           8. MAP UI ENHANCEMENTS
           ========================================== */
        /* Target the Folium iframe to round corners and add a soft shadow */
        iframe {
            border-radius: 16px !important;
            box-shadow: -5px 6px 26px 0px rgba(0,0,0,0.1) !important;
            overflow: hidden !important; 
        }

    </style>
""", unsafe_allow_html=True)

st.markdown("### DOST-Davao Project Mapper 📍")

# --- 4. Actually Rendering Stuff ---

# ==========================================
# 1. SIDEBAR: STATIC BRANDING (Always visible)
# ==========================================
st.sidebar.image("assets/dost_davao_logo.png", width="stretch")

# ==========================================
# 2. SIDEBAR: FILE UPLOAD
# ==========================================
uploaded_file = st.sidebar.file_uploader(
    "Upload DOST-Davao Excel File (.xlsx)",
    type=["xlsx"], # this removes the "Add files" button entirely
)

# --- POP-UP MODAL FUNCTION ---
# This acts as our "Right Sidebar" alternative. It pops up over the map.
@st.dialog("📋 Located Project Details", width="large")
def show_project_details(clicked_projects):
    st.success(f"Found {len(clicked_projects)} project(s) at this location.")
    st.divider()
    
    for _, row in clicked_projects.iterrows():
        # Safely extract all values
        title = str(row.get(GROUP_KEY_COL, 'Unnamed Project'))
        division = str(row.get('Division', 'N/A'))
        status = str(row.get('Map_Status', 'Unknown'))
        funding = row.get('Amount of \nfunding provided', 'Not specified')
        agency = row.get('Implementing Agency', 'Not specified')
        proponent = str(row.get('Name of Proponent', 'Not specified'))
        remarks = str(row.get('Remarks', 'No remarks provided.'))
        
        # --- NEW: Safe Date Extraction ---
        # Notice the exact placement of the \n characters based on your list
        date_approved = str(row.get('Date of\nApproval', 'Not specified'))
        date_end = str(row.get('Date of \nCompletion/\nTerminated\n(if terminated)', 'Not specified'))
        
        # Clean up Pandas 'nan' or 'NaT' text if the Excel cell was completely blank
        if date_approved.lower() in ['nan', 'nat', 'none', 'natype']: date_approved = 'TBA / Not specified'
        if date_end.lower() in ['nan', 'nat', 'none', 'natype']: date_end = 'TBA / Not specified'
        
        # Display the UI Card
        st.header(f"{title}")
        st.subheader(f"**Division:** {division} | **Status:** {status}")
        
        # Split the details into two neat columns inside the modal
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"**🏢 Agency:** {agency}")
            st.markdown(f"**👤 Proponent:** {proponent}")
            st.markdown(f"**💰 Funding:** ₱{funding}")
            
        with col2:
            st.markdown(f"**📅 Date Approved:** {date_approved[:10]}") # [:10] safely strips off any " 00:00:00" timestamps
            st.markdown(f"**🏁 End/Target Date:** {date_end[:10]}")
            
        # Remarks remain full-width at the bottom
        st.info(f"**Remarks:** {remarks}")
        st.divider()

@st.dialog("📊 Raw Data Explorer", width="large")
def show_raw_data(df):
    st.caption(f"Currently viewing {len(df)} filtered projects.")
    st.dataframe(df, width="stretch")

# ==========================================
# 3. MAIN APP LOGIC (Only runs after upload)
# ==========================================
if uploaded_file is not None:
    with st.spinner("Processing..."):
        clean_df = load_and_clean_data(uploaded_file)
        
    st.sidebar.header("🔍 Filter Dashboard")
    
    available_divisions = clean_df['Division'].unique().tolist()
    selected_divisions = st.sidebar.multiselect("Select Division(s):", available_divisions, default=available_divisions)
    
    available_statuses = clean_df['Map_Status'].unique().tolist()
    selected_statuses = st.sidebar.multiselect("Select Project Status:", available_statuses, default=available_statuses)
    
    filtered_df = clean_df[
        (clean_df['Division'].isin(selected_divisions)) &
        (clean_df['Map_Status'].isin(selected_statuses))
    ]
    
    st.sidebar.markdown("---")
    st.sidebar.metric(label="Projects Displayed", value=len(filtered_df))
    
    # NEW: Button to trigger the raw data modal
    if st.sidebar.button("📄 View Raw Data Table", width="stretch"):
        show_raw_data(filtered_df)
    
    # ==========================================
    # NEW: KPI SCORECARDS (Task 2.4 - Updated)
    # ==========================================
    # Create 5 columns for our metrics to include Terminated
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    
    with kpi1:
        st.metric(label="Total Projects", value=len(filtered_df))
        
    with kpi2:
        ongoing_count = len(filtered_df[filtered_df['Map_Status'] == 'Ongoing'])
        st.metric(label="🚀 Ongoing", value=ongoing_count)
        
    with kpi3:
        completed_count = len(filtered_df[filtered_df['Map_Status'] == 'Completed'])
        st.metric(label="✅ Completed", value=completed_count)
        
    with kpi4:
        terminated_count = len(filtered_df[filtered_df['Map_Status'] == 'Terminated'])
        st.metric(label="🛑 Terminated", value=terminated_count)
        
    with kpi5:
        # Safely convert the string values to floats and sum
        funding_col = 'Amount of \nfunding provided'
        if funding_col in filtered_df.columns:
            total_funding = pd.to_numeric(
                filtered_df[funding_col].astype(str).str.replace(',', '', regex=False), 
                errors='coerce'
            ).sum()
            
            # Format the massive number cleanly
            if total_funding >= 1_000_000:
                formatted_funding = f"₱{total_funding/1_000_000:.2f}M"
            else:
                formatted_funding = f"₱{total_funding:,.0f}"
        else:
            formatted_funding = "N/A"
            
        st.metric(label="💰 Total Funding", value=formatted_funding)
    # ==========================================
    
    # --- MAP RENDERING ---
    project_map = create_map(filtered_df)
    # st.warning(filtered_df.columns.tolist())

    # Notice we reduced the height from 550 to 480 so the map + KPIs fit on one screen!
    map_data = st_folium(
        project_map, 
        use_container_width=True, 
        height=480,
        returned_objects=["last_object_clicked"]
    )
    
    # --- CLICK LISTENER ---
    if map_data.get("last_object_clicked"):
        clicked_lat = map_data["last_object_clicked"]["lat"]
        clicked_lng = map_data["last_object_clicked"]["lng"]
        tolerance = 0.0001 
        
        clicked_projects = filtered_df[
            (abs(filtered_df['Lat'] - clicked_lat) < tolerance) & 
            (abs(filtered_df['Long'] - clicked_lng) < tolerance)
        ]
        
        if not clicked_projects.empty:
            # Trigger the pop-up modal!
            show_project_details(clicked_projects)

else:
    st.info("Upload the raw Excel file in the sidebar to begin.")