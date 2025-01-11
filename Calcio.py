import streamlit as st
import pandas as pd
import os

@st.cache_data
def load_excel_files(uploaded_files):
    """Loads data from multiple Excel files and their sheets."""
    sheet_data = {}
    sheet_mapping = {}

    for uploaded_file in uploaded_files:
        excel_file = pd.ExcelFile(uploaded_file)
        for sheet in excel_file.sheet_names:
            df = excel_file.parse(sheet_name=sheet)
            sheet_key = sheet  # Use only the sheet name as key
            sheet_data[sheet_key] = df
            sheet_mapping[sheet] = uploaded_file.name  # Map sheet name to file name

    return sheet_data, sheet_mapping

def combine_selected_sheets(sheet_data, selected_sheets):
    """Combine selected sheets into a single DataFrame, ensuring required columns."""
    dataframes = []
    for sheet in selected_sheets:
        df = sheet_data[sheet]
        df['Sheet Name'] = sheet  # Add sheet name column
        if 'DOMARE' not in df.columns:
            df['DOMARE'] = 0
        if 'ASS. DOMARE' not in df.columns:
            df['ASS. DOMARE'] = 0
        dataframes.append(df)

    combined_df = pd.concat(dataframes, ignore_index=True)
    return combined_df

def add_total_row(dataframe, group_by_column):
    """Add a total row to the DataFrame."""
    total_row = pd.DataFrame({
        group_by_column: ["Totalt"],
        'DOMARE': [dataframe['DOMARE'].sum()],
        'ASS. DOMARE': [dataframe['ASS. DOMARE'].sum()],
        'Totalt': [dataframe['Totalt'].sum()]
    })
    return pd.concat([dataframe, total_row], ignore_index=True)

def display_summary_statistics(summary_df):
    """Display basic statistics about the summary data."""
    # st.markdown("""### Sammanfattning""")
    # st.write(f"- **Totalt antal domare**: {summary_df['DOMARE'].sum()}")
    # st.write(f"- **Totalt antal matcher**: {len(summary_df)}")
    # st.write(f"- **Totalt antal assistentmatcher**: {summary_df['ASS. DOMARE'].sum()}")

def generate_bar_chart(dataframe, group_column):
    """Generate a bar chart for the given data."""
    chart_data = dataframe[dataframe[group_column] != "Totalt"]
    st.bar_chart(data=chart_data.set_index(group_column)[['DOMARE', 'ASS. DOMARE']])

st.set_page_config(page_title="Excel Data Analysis Tool", layout="wide")
st.title("Excel Data Analysis Tool")

st.markdown("""
### Ladda upp Excel-filer
Ladda upp Excel-filer som innehåller information om domare och assisterande. Du kan filtrera och välja specifika serier för analys.
""")

uploaded_files = st.file_uploader("Ladda upp Excel-filer", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    sheet_data, sheet_mapping = load_excel_files(uploaded_files)

    # st.markdown("### Översikt över uppladdade filer")
    # st.write([file.name for file in uploaded_files])

    st.markdown("### Välj serier att analysera")
    all_sheets = list(sheet_data.keys())
    # sheet_filter = st.text_input("Filtrera efter namn på flik:", "")
    filtered_sheets = [sheet for sheet in all_sheets] # if sheet_filter.lower() in sheet.lower()]

    select_all = st.checkbox("Välj alla filtrerade serier")
    if select_all:
        selected_sheets = filtered_sheets
    else:
        selected_sheets = st.multiselect(
            "Välj de serier du vill analysera:",
            options=filtered_sheets,
            default=filtered_sheets
        )

    if selected_sheets:
        combined_df = combine_selected_sheets(sheet_data, selected_sheets)

        if 'NAMN' in combined_df.columns:
            # Ensure NAMN column is string type
            combined_df['NAMN'] = combined_df['NAMN'].astype(str)

            summary_df = combined_df.groupby('NAMN', as_index=False).agg({
                'DOMARE': 'sum',
                'ASS. DOMARE': 'sum'
            })
            summary_df['Totalt'] = summary_df['DOMARE'] + summary_df['ASS. DOMARE']

            st.markdown("### Summerad data")
            st.dataframe(summary_df.reset_index(drop=True), use_container_width=True)

            display_summary_statistics(summary_df)

            csv = summary_df.to_csv(index=False).encode('utf-8')
            # st.download_button(
            #     label="Ladda ner summerad data som CSV",
            #     data=csv,
            #     file_name="summary_data.csv",
            #     mime="text/csv"
            # )

            # st.markdown("### Mappning av serier till filer")
            # st.write(sheet_mapping)

            referee_names = sorted(combined_df['NAMN'].unique())
            selected_referee = st.selectbox("Välj en domare för att visa statistik:", options=referee_names)

            if selected_referee:
                referee_stats = combined_df[combined_df['NAMN'] == selected_referee]
                referee_summary = referee_stats.groupby('Sheet Name', as_index=False).agg({
                    'DOMARE': 'sum',
                    'ASS. DOMARE': 'sum'
                })
                referee_summary['Totalt'] = referee_summary['DOMARE'] + referee_summary['ASS. DOMARE']

                referee_summary = add_total_row(referee_summary, 'Sheet Name')

                st.markdown(f"### Statistik för {selected_referee}")
                st.dataframe(referee_summary.reset_index(drop=True), use_container_width=True)

                # generate_bar_chart(referee_summary, 'Sheet Name')
        else:
            st.error("De valda serierna innehåller inte den obligatoriska kolumnen: 'NAMN'.")
    else:
        st.info("Välj minst en flik för att inkludera i summeringen.")
else:
    st.info("Ladda upp en eller flera Excel-filer för att fortsätta.")
