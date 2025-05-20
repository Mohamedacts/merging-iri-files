import pandas as pd
import streamlit as st

def merge_excel_files(file_list):
    frames = []
    for uploaded_file in file_list:
        df = pd.read_excel(uploaded_file)
        frames.append(df)
    merged_df = pd.concat(frames, ignore_index=True)
    return merged_df

def main():
    st.title("Excel File Merger")
    st.write("Upload multiple Excel files to merge them into one")

    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.write("Processing files...")
        merged_df = merge_excel_files(uploaded_files)
        st.dataframe(merged_df.head(5), use_container_width=True)

        # Export merged data
        buffer = merged_df.to_excel(index=False)
        st.download_button(
            label="Download Merged Excel",
            data=buffer,
            file_name="merged_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("✅ Data merged successfully!")

if __name__ == "__main__":
    main()
