import pandas as pd
import streamlit as st
from io import BytesIO
import pypyodbc
import tempfile
import os

def merge_excel_files(file_list):
    frames = []
    for uploaded_file in file_list:
        df = pd.read_excel(uploaded_file)
        frames.append(df)
    merged_df = pd.concat(frames, ignore_index=True)
    return merged_df

def write_to_access(df, table_name="MergedData"):
    # Create a temporary Access database file
    temp_db_fd, temp_db_path = tempfile.mkstemp(suffix=".accdb")
    os.close(temp_db_fd)
    # Create Access database
    connection_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={temp_db_path};"
    )
    conn = pypyodbc.connect(connection_str, autocommit=True)
    cursor = conn.cursor()

    # Build CREATE TABLE statement dynamically
    columns = []
    for col in df.columns:
        dtype = "MEMO" if df[col].dtype == object else "DOUBLE"
        columns.append(f"[{col}] {dtype}")
    create_table_sql = f"CREATE TABLE [{table_name}] ({', '.join(columns)});"
    cursor.execute(create_table_sql)

    # Insert data
    for _, row in df.iterrows():
        placeholders = ','.join(['?'] * len(df.columns))
        insert_sql = f"INSERT INTO [{table_name}] VALUES ({placeholders})"
        cursor.execute(insert_sql, tuple(row))
    conn.commit()
    cursor.close()
    conn.close()

    # Read the file into memory
    with open(temp_db_path, "rb") as f:
        db_bytes = f.read()
    os.remove(temp_db_path)
    return db_bytes

def main():
    # Sidebar with credentials
    st.sidebar.title("About the Developer")
    st.sidebar.markdown("""
    **Name:** Mohamed Ali  
    **Phone:** +966581764292  
    [**LinkedIn**](https://www.linkedin.com/in/mohameddalli)
    """)

    st.title("Excel to Access Merger")
    st.write(
        "Upload multiple Excel files (.xlsx or .xls). "
        "This app will merge them and let you download the combined data as an Access (.accdb) database."
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

            # Write merged DataFrame to Access database in memory
            st.info("Converting to Access database (this may take a while for large files)...")
            db_bytes = write_to_access(merged_df)

            st.download_button(
                label="Download as Access (.accdb)",
                data=db_bytes,
                file_name="merged_data.accdb",
                mime="application/octet-stream"
            )
        except Exception as e:
            st.error(f"An error occurred: {e}")

    else:
        st.warning("Please upload at least two Excel files to merge.")

if __name__ == "__main__":
    main()
