import streamlit as st
import pandas as pd
import folium
from folium.features import DivIcon
from streamlit_folium import st_folium

# Define constants based on the exact Excel headers
STATUS_BOOL_COL = "Project status"
GROUP_KEY_COL = "Project Title" # Assuming this is the main identifier, adjust if needed
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

# --- 2. MAP GENERATION ---
def create_map(df):
    # Center map on Davao Region (Using your exact coordinates)
    davao_map = folium.Map(location=[7.06750309148034, 125.60060334232874], zoom_start=13)
    
    # Define CSS Hex colors based on division for the badge backgrounds
    color_map = {'CEST': '#007bff', 'LGIA': '#28a745', 'SSCP': '#6f42c1'}
    
    for _, row in df.iterrows():
        title = str(row.get(GROUP_KEY_COL, 'Unnamed'))
        
        # Truncate the title to 25 characters to prevent massive map clutter
        short_title = (title[:25] + '...') if len(title) > 25 else title
        bg_color = color_map.get(row['Division'], '#6c757d') # Default to gray
        
        # Create a custom HTML/CSS label badge
        html_label = f"""
        <div style="
            background-color: {bg_color};
            color: white;
            border-radius: 4px;
            padding: 3px 6px;
            font-size: 11px;
            font-family: Arial, sans-serif;
            white-space: nowrap;
            font-weight: bold;
            border: 1px solid white;
            box-shadow: 2px 2px 4px rgba(0,0,0,0.4);
        ">
            {short_title}
        </div>
        """
        
        # HTML formatting for the popup (Detailed info when clicked)
        popup_text = f"""
        <b>Title:</b> {title}<br>
        <b>Division:</b> {row.get('Division', 'N/A')}<br>
        <b>Status:</b> {row.get('Map_Status', 'Unknown')}<br>
        """
        
        # Place the DivIcon on the map
        folium.Marker(
            location=[row['Lat'], row['Long']],
            icon=DivIcon(
                icon_size=(150, 36),
                icon_anchor=(0, 0),
                html=html_label,
            ),
            popup=folium.Popup(popup_text, max_width=300)
        ).add_to(davao_map)
        
    return davao_map

# --- 3. STREAMLIT APP UI ---
st.set_page_config(page_title="DOST-Davao Project Mapping", layout="wide")
st.title("📍 DOST-Davao Project Mapper")

uploaded_file = st.file_uploader("Upload DOST-Davao Excel File (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    with st.spinner("Processing sheets and standardizing data..."):
        clean_df = load_and_clean_data(uploaded_file)
        
    st.success(f"Successfully loaded {len(clean_df)} mapped projects across CEST, LGIA, and SSCP.")
    
    tab1, tab2 = st.tabs(["🗺️ Map View", "📊 Raw Data"])
    with tab1:
        project_map = create_map(clean_df)
        st_folium(project_map, width=1200, height=600)
    with tab2:
        st.dataframe(clean_df)
else:
    st.info("Upload the raw Excel file to begin.")