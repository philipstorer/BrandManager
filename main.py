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

# Load Excel Data from Sheet1 (Criteria)
@st.cache_data
def load_criteria(filename):
    try:
        # Read columns A through M from Sheet1 using header row 0
        df = pd.read_excel(filename, sheet_name=0, header=0, usecols="A:M")
        if df.shape[1] < 13:
            st.error(f"Excel file has only {df.shape[1]} column(s) but at least 13 are required. Check file formatting.")
            return None, None, None, None
        # Extract selection options from header row:
        # Role options: columns B to D (indices 1-3)
        role_options = df.columns[1:4].tolist()
        # Lifecycle options: columns F to I (indices 5-8)
        lifecycle_options = df.columns[5:9].tolist()
        # Journey options: columns J to M (indices 9-12)
        journey_options = df.columns[9:13].tolist()
        # The entire sheet is used as the matrix
        matrix_df = df.copy()
        return role_options, lifecycle_options, journey_options, matrix_df
    except Exception as e:
        st.error(f"Error reading the Excel file (Sheet1): {e}")
        return None, None, None, None

role_options, lifecycle_options, journey_options, matrix_df = load_criteria("test.xlsx")
if any(v is None for v in [role_options, lifecycle_options, journey_options, matrix_df]):
    st.stop()

# Helper: Filter Strategic Imperatives from the matrix (Sheet1)
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

# Helper: Generate AI Output using OpenAI API
def generate_ai_output(tactic_text, selected_differentiators):
    """
    Uses the OpenAI API to generate a 2-3 sentence description of the tactic,
    along with an estimated cost range and timeframe.
    Returns a dictionary with keys "description", "cost", and "timeframe".
    """
    differentiators_text = ", ".join(selected_differentiators) if selected_differentiators else "None"
    prompt = f"""
You are an expert pharmaceutical marketing strategist.
Given the following tactic: "{tactic_text}"
and the selected product differentiators: "{differentiators_text}",
please provide a short 2-3 sentence description of the tactic.
Also, provide an estimated cost range in USD and an estimated timeframe in months for implementation.
Return the output as a JSON object with keys "description", "cost", and "timeframe".
    """
    try:
        with st.spinner("Generating AI output..."):
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are an expert pharmaceutical marketing strategist."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
            )
        content = response.choices[0].message.content.strip()
        try:
            output = json.loads(content)
        except json.JSONDecodeError:
            st.error("Error decoding AI response. Please try again.")
            output = {"description": "N/A", "cost": "N/A", "timeframe": "N/A"}
        return output
    except Exception as e:
        st.error(f"Error generating AI output: {e}")
        return {"description": "N/A", "cost": "N/A", "timeframe": "N/A"}

# Build the Streamlit Interface

st.title("Pharma Strategy Online Tool")

# Step 1: Criteria Selection
st.header("Step 1: Select Your Criteria")
role_selected = st.selectbox("Select your role", role_options)
lifecycle_selected = st.selectbox("Select the product lifecycle stage", lifecycle_options)
journey_selected = st.selectbox("Select the customer journey focus", journey_options)

# Step 2: Strategic Imperatives Selection
st.header("Step 2: Select Strategic Imperatives")
strategic_options = filter_strategic_imperatives(matrix_df, role_selected, lifecycle_selected, journey_selected)
if not strategic_options:
    st.warning("No strategic imperatives found for these selections. Please try different options.")
else:
    selected_strategics = st.multiselect("Select up to 3 Strategic Imperatives", options=strategic_options, max_selections=3)

# Step 3: Product Differentiators Selection
st.header("Step 3: Select Product Differentiators")
try:
    sheet2 = pd.read_excel("test.xlsx", sheet_name=1, header=0)
except Exception as e:
    st.error(f"Error reading Sheet2: {e}")
    st.stop()
if "Differentiator" not in sheet2.columns:
    st.error("Sheet2 must have a column named 'Differentiator'.")
    st.stop()
product_diff_options = sheet2["Differentiator"].dropna().unique().tolist()
selected_differentiators = st.multiselect("Select up to 3 Product Differentiators", options=product_diff_options, max_selections=3)

# Step 4: Generate Tactical Recommendations based on Sheet3
if st.button("Generate Strategy"):
    if not selected_strategics:
        st.error("Please select at least one strategic imperative.")
    else:
        st.header("Tactical Recommendations")
        try:
            sheet3 = pd.read_excel("test.xlsx", sheet_name=2, header=0)
        except Exception as e:
            st.error(f"Error reading Sheet3: {e}")
            st.stop()
        # Ensure Sheet3 has the required columns
        required_cols = ["Strategic Imperative", "Patient & Caregiver", "HCP Engagement"]
        if not all(col in sheet3.columns for col in required_cols):
            st.error("Sheet3 must have columns named 'Strategic Imperative', 'Patient & Caregiver', and 'HCP Engagement'.")
            st.stop()
        
        # For each selected strategic imperative, pull the appropriate tactic
        for imperative in selected_strategics:
            row = sheet3[sheet3["Strategic Imperative"] == imperative]
            if row.empty:
                st.info(f"No tactic found for strategic imperative: {imperative}")
                continue
            # Determine the tactic based on role selection
            if role_selected == "HCP":
                tactic = row["HCP Engagement"].iloc[0]
            else:
                tactic = row["Patient & Caregiver"].iloc[0]
            
            # Optionally customize with product differentiators
            if selected_differentiators:
                tactic_customized = f"{tactic} (Customized with: {', '.join(selected_differentiators)})"
            else:
                tactic_customized = tactic
            
            # Generate AI output for this tactic
            ai_output = generate_ai_output(tactic_customized, selected_differentiators)
            st.subheader(f"{imperative}: {tactic_customized}")
            st.write(ai_output.get("description", "No description available."))
            st.write(f"**Estimated Cost:** {ai_output.get('cost', 'N/A')}")
            st.write(f"**Estimated Timeframe:** {ai_output.get('timeframe', 'N/A')}")
