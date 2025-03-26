import streamlit as st
import pandas as pd
import openai
import json
import os

# Secure API Key Handling
if "openai" in st.secrets and "api_key" in st.secrets["openai"]:
    openai.api_key = st.secrets["openai"]["api_key"]
else:
    openai_api_key = os.getenv("OPENAI_API_KEY")
    if not openai_api_key:
        st.error("OpenAI API key not found. Please set it in .streamlit/secrets.toml or as an environment variable.")
        st.stop()
    openai.api_key = openai_api_key

# -----------------------
# Load Criteria from Sheet1 (Columns A–M)
# -----------------------
@st.cache_data
def load_criteria(filename):
    try:
        # Read Sheet1: columns A through M using header row 0.
        df = pd.read_excel(filename, sheet_name=0, header=0, usecols="A:M")
        if df.shape[1] < 13:
            st.error(f"Excel file has only {df.shape[1]} column(s) but at least 13 are required. Check file formatting.")
            return None, None, None, None
        # Extract options from header row:
        # Role options: columns B-D (indices 1 to 3)
        role_options = df.columns[1:4].tolist()
        # Remove "Caregiver" from role options
        role_options = [r for r in role_options if r.lower() != "caregiver"]
        # Lifecycle options: columns F-I (indices 5 to 8)
        lifecycle_options = df.columns[5:9].tolist()
        # Journey options: columns J-M (indices 9 to 12)
        journey_options = df.columns[9:13].tolist()
        matrix_df = df.copy()  # The entire sheet is used as the matrix.
        return role_options, lifecycle_options, journey_options, matrix_df
    except Exception as e:
        st.error(f"Error reading the Excel file (Sheet1): {e}")
        return None, None, None, None

role_options, lifecycle_options, journey_options, matrix_df = load_criteria("test.xlsx")
if any(v is None for v in [role_options, lifecycle_options, journey_options, matrix_df]):
    st.stop()

# Prepend placeholder text to the dropdown lists.
role_placeholder = "Audience"
lifecycle_placeholder = "Product Life Cycle"
journey_placeholder = "Customer Journey Focus"

new_role_options = [role_placeholder] + role_options
new_lifecycle_options = [lifecycle_placeholder] + lifecycle_options
new_journey_options = [journey_placeholder] + journey_options

# -----------------------
# Helper: Filter Strategic Imperatives from the Matrix (Sheet1)
# -----------------------
def filter_strategic_imperatives(df, role, lifecycle, journey):
    """
    Filters the matrix (df) for strategic imperatives where the cells in the
    selected role, lifecycle, and journey columns contain an "x" (case-insensitive).
    Assumes there is a column named "Strategic Imperative".
    """
    if role not in df.columns or lifecycle not in df.columns or journey not in df.columns:
        st.error("The Excel file's columns do not match the expected names for filtering.")
        return []
    try:
        filtered = df[
            (df[role].astype(str).str.lower() == 'x') &
            (df[lifecycle].astype(str).str.lower() == 'x') &
            (df[journey].astype(str).str.lower() == 'x')
        ]
        return filtered["Strategic Imperative"].dropna().tolist()
    except Exception as e:
        st.error(f"Error filtering strategic imperatives: {e}")
        return []

# -----------------------
# Helper: Generate Tactical Recommendation Output
# -----------------------
def generate_ai_output(tactic_text, selected_differentiators):
    """
    Uses the OpenAI API (gpt-3.5
