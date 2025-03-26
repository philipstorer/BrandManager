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
# Load Criteria from Sheet1 (A–M)
# -----------------------
@st.cache_data
def load_criteria(filename):
    try:
        # Read columns A through M from Sheet1 using header row 0.
        df = pd.read_excel(filename, sheet_name=0, header=0, usecols="A:M")
        if df.shape[1] < 13:
            st.error(f"Excel file has only {df.shape[1]} column(s) but at least 13 are required. Check file formatting.")
            return None, None, None, None
        # Extract options from the header row:
        # Role options: columns B-D (indices 1-3)
        role_options = df.columns[1:4].tolist()
        # Lifecycle options: columns F-I (indices 5-8)
        lifecycle_options = df.columns[5:9].tolist()
        # Journey options: columns J-M (indices 9-12)
        journey_options = df.columns[9:13].tolist()
        matrix_df = df.copy()  # Entire sheet as matrix
        return role_options, lifecycle_options, journey_options, matrix_df
    except Exception as e:
        st.error(f"Error reading the Excel file (Sheet1): {e}")
        return None, None, None, None

role_options, lifecycle_options, journey_options, matrix_df = load_criteria("test.xlsx")
if any(v is None for v in [role_options, lifecycle_options, journey_options, matrix_df]):
    st.stop()

# Prepend placeholders to dropdown lists.
role_placeholder = "Audience"
lifecycle_placeholder = "Product Life Cycle"
journey_placeholder = "Customer Journey Focus"

# Remove "Caregiver" from role options.
filtered_role_options = [option for option in role_options if option != "Caregiver"]
new_role_options = [role_placeholder] + filtered_role_options
new_lifecycle_options = [lifecycle_placeholder] + lifecycle_options
new_journey_options = [journey_placeholder] + journey_options

# -----------------------
# Helper: Filter Strategic Imperatives from Sheet1 Matrix
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
    Uses the OpenAI API (gpt-3.5-turbo) to generate a 2-3 sentence description of the tactic.
    The prompt instructs the model to explain how implementing this tactic will deliver on the strategic imperative,
    highlighting how it leverages the product differentiators.
    Return ONLY a JSON object with keys "description", "cost", and "timeframe".
    """
    differentiators_text = ", ".join(selected_differentiators) if selected_differentiators else "None"
    prompt = f"""
You are an expert pharmaceutical marketing strategist.
Given the following tactic: "{tactic_text}"
and considering the selected product differentiators: "{differentiators_text}",
explain in 2-3 sentences how implementing this tactic will deliver on the strategic imperative and effectively leverage these differentiators.
Also, provide an estimated cost range in USD and an estimated timeframe in months for implementation.
Return ONLY a JSON object with exactly the following keys: "description", "cost", "timeframe".
"""
    try:
        with st.spinner("Generating tactical recommendation..."):
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
            st.error("Error decoding the response. Please try again.")
            output = {"description": "N/A", "cost": "N/A", "timeframe": "N/A"}
        return output
    except Exception as e:
        st.error(f"Error generating tactical recommendation: {e}")
        return {"description": "N/A", "cost": "N/A", "timeframe": "N/A"}

# -----------------------
# Build the Streamlit Interface
# -----------------------

st.title("Pharma AI Brand Manager")

# Step 1: Criteria Selection
st.header("Step 1: Select Your Criteria")
role_selected = st.selectbox("", new_role_options)
lifecycle_selected = st.selectbox("", new_lifecycle_options)
journey_selected = st.selectbox("", new_journey_options)

# Only show further steps if all selections in Step 1 are made (i.e. not placeholders).
if role_selected != role_placeholder and lifecycle_selected != lifecycle_placeholder and journey_selected != journey_placeholder:
    # Step 2: Strategic Imperatives
    st.header("Step 2: Select Strategic Imperatives")
    strategic_options = filter_strategic_imperatives(matrix_df, role_selected, lifecycle_selected, journey_selected)
    if not strategic_options:
        st.warning("No strategic imperatives found for these selections. Please try different options.")
    else:
        selected_strategics = st.multiselect("Select up to 3 Strategic Imperatives", options=strategic_options, max_selections=3)

    # Only show Step 3 if at least one strategic imperative is selected.
    if selected_strategics and len(selected_strategics) > 0:
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

        # Only show the CTA if at least one product differentiator is selected.
        if selected_differentiators and len(selected_differentiators) > 0:
            if st.button("Generate Strategic Plan"):
                st.header("Tactical Recommendations")
                try:
                    sheet3 = pd.read_excel("test.xlsx", sheet_name=2, header=0)
                except Exception as e:
                    st.error(f"Error reading Sheet3: {e}")
                    st.stop()
                required_cols = ["Strategic Imperative", "Patient & Caregiver", "HCP Engagement"]
                if not all(col in sheet3.columns for col in required_cols):
                    st.error("Sheet3 must have columns named 'Strategic Imperative', 'Patient & Caregiver', and 'HCP Engagement'.")
                    st.stop()
                
                # For each selected strategic imperative, pull the appropriate tactic.
                for imperative in selected_strategics:
                    row = sheet3[sheet3["Strategic Imperative"] == imperative]
                    if row.empty:
                        st.info(f"No tactic found for strategic imperative: {imperative}")
                        continue
                    # Determine the tactic based on user role.
                    if role_selected == "HCP":
                        tactic = row["HCP Engagement"].iloc[0]
                    else:
                        tactic = row["Patient & Caregiver"].iloc[0]
                    
                    # Generate tactical recommendation (the title shows the imperative and tactic, no raw differentiator text).
                    ai_output = generate_ai_output(tactic, selected_differentiators)
                    st.subheader(f"{imperative}: {tactic}")
                    st.write(ai_output.get("description", "No description available."))
                    st.write(f"**Estimated Cost:** {ai_output.get('cost', 'N/A')}")
                    st.write(f"**Estimated Timeframe:** {ai_output.get('timeframe', 'N/A')}")
        else:
            st.info("Please select at least one product differentiator to proceed.")
    else:
        st.info("Please select at least one strategic imperative to proceed.")
else:
    st.info("Please complete all criteria selections in Step 1 to proceed.")
