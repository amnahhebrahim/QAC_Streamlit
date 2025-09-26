import streamlit as st
import pandas as pd
import tempfile
import zipfile
import io
import pyperclip  # For clipboard copy (optional, depends on backend environment)
import openpyxl
import os
import re
# Replace this with your actual backend function
def run_backend_on_files(extracted_path):
    print("Extracted path is",extracted_path)
    #For now hardcoded:
    start_row = 3
    end_row = 20
    start_column = 1
    end_column = 15
    pilos_dict=[]
    # for p in extracted_path:
    for root, dirs, files in os.walk(extracted_path):
        for f in files:
            p = os.path.join(root, f)
            # To open the workbook
            # workbook object is created
            wb_obj = openpyxl.load_workbook(p, data_only=True)

            # Get workbook active sheet object
            # from the active attribute
            for i in wb_obj.sheetnames:
                # print(i)
                if i =="PILOsReport":
                    sheet_obj=wb_obj[i]
                    data =[]
                    for row in sheet_obj.iter_rows(min_row=start_row, max_row=end_row, min_col=start_column, max_col=end_column):
                        row_data = []
                        for cell in row:
                            row_data.append(cell.value)
                        data.append(row_data)

                    df_temp=pd.DataFrame(data)
                    instructor=df_temp.iloc[0]
                    pilos_asses=df_temp.iloc[17]
                    pilos_list=pilos_asses.values.tolist()
                    d={}
                    columns_list=["Subject", "PILOS","SO1","SO2","SO3","SO4","SO5","SO6","SO7"]
                    for j in columns_list:
                        # for i in pilos_list:
                        # print("I is at",i)
                        if j =="Subject":
                            subject = [splits.replace(".xlsx", "") for splits in p.split('\\') if "EENG" in splits][0]
                            d.update({j:subject})
                            print(j,p)
                        else:
                            d.update({j:pilos_list[columns_list.index(j)]})
                            print(j,pilos_list[columns_list.index(j)])
                    # df_master=df_master.append(pilos_dict,ignore_index=True)
                    pilos_dict.append(d)
                # print(d)

    pilos_df= pd.DataFrame(pilos_dict)

    final_pilos_df=calculate_average_pilos(pilos_df)
    return final_pilos_df

def calculate_average_pilos(df):
    so_columns = [col for col in df.columns if re.match(r'SO\d+', col)]
    averages=[]

    for col in df.columns:
        if col not in so_columns:
            averages.append("")
        else:
            list_of_percentages=df[col].dropna()
            averages.append(list_of_percentages.mean())
            # print("col",col, "list:",list_of_percentages)
    df.loc[len(df)] = averages

    return df


def get_students_info(tmpdir):
    st.title("Student Information")
    return {"Name": "John Doe", "Age": 25, "Major": "Computer Science"}


def pilos(tmpdir):
    
    st.text("Loading the PILOs' Average Calculator")
    df = run_backend_on_files(tmpdir)
    # File uploader

    st.success("Processing complete ✅")

    # Display the dataframe
    st.dataframe(df)

    # Download option
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download as CSV",
        csv,
        "results.csv",
        "text/csv",
        key="download-csv"
    )

    # Copy to clipboard option (browser limitation workaround)
    # Streamlit doesn’t allow direct clipboard copy on server,
    # so we use JS trick:
    st.text_area("Copy results:", df.to_csv(index=False), height=200)
    st.caption("Select all (Ctrl+A) and copy (Ctrl+C)")

    # Optional: If running locally and want automatic copy:
    if st.button("Copy to Clipboard (local only)"):
        try:
            pyperclip.copy(df.to_csv(index=False))
            st.success("Copied to clipboard!")
        except Exception as e:
            st.error(f"Clipboard copy failed: {e}")

st.title("QAC ToolBox")

uploaded_file = st.file_uploader("Upload a ZIP file", type=["zip"])

if uploaded_file is not None:
    # Create a temporary directory to extract zip
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(uploaded_file, "r") as z:
            z.extractall(tmpdir)
        print(tmpdir)
        # Run your backend code
        


page_names_to_funcs = {
    "PiLOs": pilos(tmpdir),
    "Students": get_students_info(tmpdir),
    # "DataFrame Demo": data_frame_demo
}

demo_name = st.sidebar.selectbox("Choose a demo", page_names_to_funcs.keys())
page_names_to_funcs[demo_name]()