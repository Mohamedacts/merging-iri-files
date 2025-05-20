import pandas as pd
import streamlit as st
from io import StringIO

def merge_excel_files(file_list):
    frames = []
    for uploaded_file in file_list:
        df = pd.read_excel(uploaded_file)
        frames.append(df)
    merged_df = pd.concat(frames, ignore_index=True)
    return merged_df

def main():
    # Sidebar with credentials
    st.sidebar.title("About the Developer")
    st.sidebar.markdown("""
    **Name:** Mohamed Ali  
    **Phone:** +966581764292  
    [**LinkedIn**](https://www.linkedin.com/in/mohameddalli)
    """)

    st.title("Excel to CSV Merger")
    st.write(
        "Upload multiple Excel files (.xlsx or .xls). "
        "This app will merge them and let you download the combined data as a CSV file."
    )

    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.info(f"{len(uploaded_files)} file(s) uploaded. Processing...")
        try:
            merged_df = merge_excel_files(uploaded_files)
            st.success(f"Successfully merged {len(uploaded_files)} files!")
            st.write("Preview of merged data:")
            st.dataframe(merged_df.head(10), use_container_width=True)

            # Write merged DataFrame to CSV in memory
            csv_buffer = StringIO()
            merged_df.to_csv(csv_buffer, index=False)
            csv_data = csv_buffer.getvalue()

            st.download_button(
                label="Download as CSV",
                data=csv_data,
                file_name="merged_data.csv",
                mime="text/csv"
            )
        except Exception as e:
            st.error(f"An error occurred: {e}")

    else:
        st.warning("Please upload at least two Excel files to merge.")

if __name__ == "__main__":
    main()
