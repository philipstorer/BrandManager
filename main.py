import streamlit as st
import pandas as pd
import openai
import json
import os

# -----------------------
# Debug: Print Current Working Directory
# -----------------------
st.write("Current working directory:", os.getcwd())

# -----------------------
# Secure API Key Handling
# -----------------------
if "openai" in st.secrets and "api_key" in st.secrets["openai"]:
    openai.api_key = st.secrets["openai"]["api_key"]
else:
    openai_api_key = os.getenv("OPENAI_API_KEY")
    if not openai_api_key:
        st.error("OpenAI API key not found. Please set it in .streamlit/secrets.toml or as an environment variable.")
        st.stop()
    openai.api_key = openai_api_key

# -----------------------
# Load Excel Data from Sheet1 (Criteria) from test.xlsx
# -----------------------
@st.cache_data
def load_excel_data(filename):
    try:
        # Read columns A through M from Sheet1 using header row 0
        raw_df = pd.read_excel(filename, sheet_name=0, header=0, usecols="A:M")
        st.write("Data shape (rows, columns):", raw_df.shape)
        if raw_df.shape[1] < 13:
            st.error(f"Excel file has only {raw_df.shape[1]} column(s) but at least 13 are required. Check file formatting.")
            return None, None, None, None

        # Extract selection options based on header names:
        # Role options: cells B1 to D1 → columns index 1 to 3
        role_options = raw_df.columns[1:4].tolist()
        # Lifecycle options: cells F1 to I1 → columns index 5 to 8
        lifecycle_options = raw_df.columns[5:9].tolist()
        # Journey options: cells J1 to M1 → columns index 9 to 12
        journey_options = raw_df.columns[9:13].tolist()

        # The matrix data (strategic imperatives and "x" marks) are in the entire DataFrame
        matrix_df = raw_df.copy()

        return role_options, lifecycle_options, journey_options, matrix_df
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")
        return None, None, None, None

role_options, lifecycle_options, journey_options, matrix_df = load_excel_data("test.xlsx")
if role_options is None or lifecycle_options is None or journey_options is None or matrix_df is None:
    st.stop()

# Debug: Display the extracted selection options
st.write("Role Options:", role_options)
st.write("Lifecycle Options:", lifecycle_options)
st.write("Journey Options:", journey_options)

# -----------------------
# Helper Functions
# -----------------------
def filter_strategic_imperatives(df, role, lifecycle, journey):
    """
    Filters the matrix (df) for strategic imperatives for which the cells in the
    selected role, lifecycle, and journey columns contain an "x" (case-insensitive).
    Assumes that one of the columns in df is labeled "Strategic Imperative".
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

def generate_ai_output(customized_result, selected_differentiators):
    """
    Calls the OpenAI API to generate a 2-3 sentence description of the strategic recommendation,
    along with an estimated cost range and timeframe.
    Returns a dictionary with keys: "description", "cost", and "timeframe".
    """
    differentiators_text = ", ".join(selected_differentiators) if selected_differentiators else "None"
    prompt = f"""
You are an expert pharmaceutical marketing strategist.
Given the following strategy description: "{customized_result}"
and the selected product differentiators: "{differentiators_text}",
please provide a short 2-3 sentence description of the strategic recommendation.
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

# -----------------------
# Build the Streamlit Interface
# -----------------------
st.title("Pharma Strategy Online Tool")

# Step 1: Criteria Selection
st.header("Step 1: Select Your Criteria")
role_selected = st.selectbox("Select your role", role_options)
lifecycle_selected = st.selectbox("Select the product lifecycle stage", lifecycle_options)
journey_selected = st.selectbox("Select the customer journey focus", journey_options)

# Step 2: Strategic Imperatives
st.header("Step 2: Select Strategic Imperatives")
strategic_options = filter_strategic_imperatives(matrix_df, role_selected, lifecycle_selected, journey_selected)
if not strategic_options:
    st.warning("No strategic imperatives found for these selections. Please try different options.")
else:
    selected_strategics = st.multiselect("Select up to 3 Strategic Imperatives", options=strategic_options, max_selections=3)

# Step 3: Product Differentiators
st.header("Step 3: Select Product Differentiators")
# Load Sheet2 (assumed to have a column "Product Differentiators")
sheet2 = pd.read_excel("test.xlsx", sheet_name=1, header=0)
if "Product Differentiators" not in sheet2.columns:
    st.error("Sheet2 must have a column named 'Product Differentiators'.")
    st.stop()
product_diff_options = sheet2["Product Differentiators"].dropna().unique().tolist()
selected_differentiators = st.multiselect("Select up to 3 Product Differentiators", options=product_diff_options, max_selections=3)

# Generate and Display Results
if st.button("Generate Strategy"):
    if not selected_strategics:
        st.error("Please select at least one strategic imperative.")
    else:
        st.header("Strategic Recommendations")
        # Load Sheet3 (assumed to have columns "Strategic Imperative" and "Result")
        sheet3 = pd.read_excel("test.xlsx", sheet_name=2, header=0)
        if "Strategic Imperative" not in sheet3.columns or "Result" not in sheet3.columns:
            st.error("Sheet3 must have columns named 'Strategic Imperative' and 'Result'.")
            st.stop()
        results_df = sheet3[sheet3["Strategic Imperative"].isin(selected_strategics)]
        if results_df.empty:
            st.info("No results found for the selected strategic imperatives.")
        else:
            for idx, row in results_df.iterrows():
                base_result = row["Result"]
                differentiators_text = ", ".join(selected_differentiators) if selected_differentiators else ""
                customized_result = base_result
                if differentiators_text:
                    customized_result += f" (Customized with: {differentiators_text})"
                ai_output = generate_ai_output(customized_result, selected_differentiators)
                st.subheader(customized_result)
                st.write(ai_output.get("description", "No description available."))
                st.write(f"**Estimated Cost:** {ai_output.get('cost', 'N/A')}")
                st.write(f"**Estimated Timeframe:** {ai_output.get('timeframe', 'N/A')}")
