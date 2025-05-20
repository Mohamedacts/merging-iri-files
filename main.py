import pandas as pd
import streamlit as st
from io import BytesIO

def merge_excel_files(file_list):
    """Merge multiple Excel files into a single DataFrame."""
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

    st.title("Excel File Merger")
    st.write(
        "Upload multiple Excel files (.xlsx or .xls). "
        "This app will merge them and let you download the combined file."
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

            # Write merged DataFrame to Excel in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                merged_df.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="Download Merged Excel",
                data=output,
                file_name="merged_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"An error occurred: {e}")

    else:
        st.warning("Please upload at least two Excel files to merge.")

if __name__ == "__main__":
    main()
