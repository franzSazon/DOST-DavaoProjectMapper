# DOST-Davao Project Mapper

New branch commit test.

## Overview
The DOST-Davao Project Mapper is a Streamlit application designed to ingest, process, and visualize project data from DOST-Davao Excel workbooks. It consolidates data across different divisions and provides an interactive map interface to track and manage regional projects.

## Features
*   **Data Ingestion and Processing:** Automatically reads and cleans data from project Excel workbooks, consolidating records across different DOST divisions (e.g., CEST, LGIA, SSCP).
*   **Interactive Mapping:** Displays project locations on an interactive map using Folium, allowing users to visually navigate project distributions.
*   **Filterable KPIs:** Provides key performance indicators and filtering options to analyze projects by status, division, and other relevant metrics.
*   **Detailed Project Views:** Allows users to select individual projects on the map to view specific details and status updates.
*   **AI Integration:** Utilizes Ollama's llama3.2:3b for advanced data processing or summarization tasks within the application.

## Installation and Setup

1.  Clone the repository.
2.  Install the required dependencies using pip:
    ```bash
    pip install -r requirements.txt
    ```
3.  Set up your Streamlit configuration and secrets (like API keys) in the `.streamlit/secrets.toml` file.
4.  Run the application using Streamlit:
    ```bash
    streamlit run app_refactor.py
    ```

## Technology Stack
*   Python
*   Streamlit
*   Pandas
*   Folium
*   Ollama local LLM (llama3.2:3b)
