import pandas as pd
import streamlit as st

# Function to process the file and generate output
def process_file(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        df_students = df.iloc[3:]
        midsem_weight = df.iloc[1, df.columns.get_loc('Mid Sem')]
        endsem_weight = df.iloc[1, df.columns.get_loc('Endsem')]
        quiz1_weight = df.iloc[1, df.columns.get_loc('Quiz 1')]
        quiz2_weight = df.iloc[1, df.columns.get_loc('Quiz 2')]
        max_midsem = df.iloc[0, df.columns.get_loc('Mid Sem')]
        max_endsem = df.iloc[0, df.columns.get_loc('Endsem')]
        max_quiz1 = df.iloc[0, df.columns.get_loc('Quiz 1')]
        max_quiz2 = df.iloc[0, df.columns.get_loc('Quiz 2')]

        def calculate_total_score(row):
            midsem_score = (row['Mid Sem'] / max_midsem) * midsem_weight
            endsem_score = (row['Endsem'] / max_endsem) * endsem_weight
            quiz1_score = (row['Quiz 1'] / max_quiz1) * quiz1_weight
            quiz2_score = (row['Quiz 2'] / max_quiz2) * quiz2_weight
            return midsem_score + endsem_score + quiz1_score + quiz2_score

        df_students['Total Scaled/100'] = df_students.apply(calculate_total_score, axis=1)
        df_sorted = df_students.sort_values(by='Total Scaled/100', ascending=False)

        # Generate output Excel file
        output_file = 'Output.xlsx'
        with pd.ExcelWriter(output_file) as writer:
            df_sorted.to_excel(writer, sheet_name='Sorted by Total Score', index=False)

        return output_file
    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None

# Webpage UI
st.markdown(
    """
    <style>
    body {
        background-color: #306998;
    }
    .title {
        font-size: 8vw;
        font-weight: bold;
        color: green;
        text-align: left;
        margin-bottom: -50px;
    }
    .subtitle {
        font-size: 1.5vw;
        color: green;
        text-align: left;
        margin-top: 0px;
    }
    </style>
    <div class="title">BROWSE</div>
    <div class="subtitle">Python Project with 🐍</div>
    """,
    unsafe_allow_html=True,
)

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    output_file = process_file(uploaded_file)
    if output_file:
        with open(output_file, "rb") as f:
            st.download_button(
                label="Download Processed File",
                data=f,
                file_name="Processed_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
