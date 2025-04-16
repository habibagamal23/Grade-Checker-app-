import streamlit as st
import pandas as pd
import os
import glob
import tempfile
import zipfile
import io
import base64
import re
import shutil


st.set_page_config(page_title="Grade Checker App", layout="wide")

def extract_and_merge(Zip_path, output_folder):
    """
    1) Creates a temporary folder internally (not passed in).
    2) Extracts 'Zip_path' into that temp folder using patool.
    3) Merges .csv/.xlsx/.xls files that share a base code (ignoring underscore suffix).
    4) Saves final merged CSVs to 'output_folder'.
    
    Example of merging:
      - MATH101_1.csv, MATH101_2.xlsx => MATH101.csv
      - CSAI330_1.xlsx, CSAI330_2.csv => CSAI330.csv
    
    Args:
        Zip_path (str): The path to the .rar archive.
        output_folder (str): The folder where merged CSVs will be placed.
    
    Returns:
        str: The output_folder path (with final merged CSVs inside).
    """
    # Ensure final output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # 1) Create a temporary directory to hold extracted files
    with tempfile.TemporaryDirectory() as temp_dir:
        st.info(f"Extracting {os.path.basename(Zip_path)} into a temporary directory...")
        
        # 2) Extract using patool (which calls an external tool, e.g. unrar or 7z)
        with zipfile.ZipFile(Zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        st.info("Extraction completed. Now merging files by base code...")

        # 3) Prepare dictionary: { base_code: [list_of_dataframes] }
        course_data = {}
        
        # Regex: e.g. "CSAI330_1" => "CSAI330"
        pattern = re.compile(r"^([A-Za-z0-9]+\d+)(?:_.*)?$", re.IGNORECASE)
        
        # Track files processed
        processed_files = []
        
        # Traverse the extracted folder
        for root, _, files in os.walk(temp_dir):
            for filename in files:
                lower_name = filename.lower()
                # Only consider .csv/.xlsx/.xls
                if not (lower_name.endswith(".csv") or 
                        lower_name.endswith(".xlsx") or 
                        lower_name.endswith(".xls")):
                    continue
                
                base_name = os.path.splitext(filename)[0]  # e.g. "MATH101_1"
                m = pattern.match(base_name)
                if m:
                    base_code = m.group(1).strip()
                else:
                    base_code = base_name  # fallback if no match

                file_path = os.path.join(root, filename)
                processed_files.append(f"{filename} → {base_code}")
                
                # Read into a DataFrame
                try:
                    if lower_name.endswith(".csv"):
                        df = pd.read_csv(file_path)
                    else:
                        df = pd.read_excel(file_path)
                    
                    # Accumulate in dictionary
                    course_data.setdefault(base_code, []).append(df)
                except Exception as e:
                    st.warning(f"Couldn't read {filename}: {str(e)}")

        # 4) Merge each base code's DataFrames and save to 'output_folder'
        merged_files = []
        for course, df_list in course_data.items():
            if not df_list:
                continue
            merged_df = pd.concat(df_list, ignore_index=True)
            
            out_name = course + ".csv"
            out_path = os.path.join(output_folder, out_name)
            merged_df.to_csv(out_path, index=False)
            
            merged_files.append(f"{course}: {len(df_list)} file(s) merged")
    
    # Temp folder is automatically cleaned up
    return output_folder, processed_files, merged_files

def normalize_grade(grade):
    """Normalize grade format for comparison"""
    g = str(grade).strip().upper()
    if g in ["P", "PASS"]:
        return "PASS"
    if g in ["ABSENT", "ABS"]:
        return "ABSENT"
    return g



def read_roster_file(file_path, header_keywords=None):
    """
    Reads an Excel roster file that may contain preliminary information
    (e.g., title, course info) before the actual header row.
    
    This function scans the top rows for a row that contains the required keywords.
    When found, that row is used as the header.
    """
    if header_keywords is None:
        header_keywords = {"LETTER GRADE"}
    else:
        # Make sure the keywords are uppercase for matching.
        header_keywords = {kw.upper() for kw in header_keywords}
    
    # Read the file without a header to inspect rows.
    df_raw = pd.read_excel(file_path, header=None)
    
    header_row = None
    id_variants = {"SID", "STUDENT ID"}
    for i, row in df_raw.iterrows():
        # Convert the row's values to strings, uppercase them, and strip whitespace
        row_upper = row.astype(str).str.upper().str.strip()
        if "LETTER GRADE" in set(row_upper) and (set(row_upper) & id_variants):
            header_row = i
            break
    
    if header_row is None:
        raise ValueError("Header row with required keywords not found in " + file_path)
    
    # Read the file again using the found row as header.
    df = pd.read_excel(file_path, header=header_row)
    return df

    """
    Compare grades between roster files and downloaded files.
    Returns results and summary stats.
    """
    results = []
    all_unmatched = []
    summary_stats = {
        "total_courses": 0,
        "courses_with_mismatches": 0,
        "total_mismatches": 0,
        "total_students": 0,
        "withdrawn_students": 0
    }
    
    # Create a dictionary of downloaded files for easy lookup
    downloaded_dict = {os.path.splitext(os.path.basename(f.name))[0]: f for f in downloaded_files}
    
    for roster_file in roster_files:
        base_name = os.path.splitext(os.path.basename(roster_file.name))[0]
        
        # Check if we have a matching downloaded file
        if base_name in downloaded_dict:
            summary_stats["total_courses"] += 1
            downloaded_file = downloaded_dict[base_name]
            
            # Process the files
            try:
                # Create temporary files to save the uploaded files
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_roster:
                    tmp_roster.write(roster_file.getvalue())
                    tmp_roster_path = tmp_roster.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp_downloaded:
                    tmp_downloaded.write(downloaded_file.getvalue())
                    tmp_downloaded_path = tmp_downloaded.name
                
                # Read the files
                try:
                    df_roster = read_roster_file(tmp_roster_path)
                except ValueError as e:
                    results.append({
                        "course": base_name,
                        "status": "error",
                        "message": str(e),
                        "data": None
                    })
                    continue
                
                df_downloaded = pd.read_csv(tmp_downloaded_path)
                
                # Clean up temporary files
                os.unlink(tmp_roster_path)
                os.unlink(tmp_downloaded_path)
                
                # Clean column names
                df_roster.columns = df_roster.columns.str.strip()
                df_downloaded.columns = df_downloaded.columns.str.strip()
                
                # Check for required columns
                if 'SID' in df_roster.columns:
                    sid_col = 'SID'
                elif 'Student ID' in df_roster.columns:
                    sid_col = 'Student ID'
                else:
                    results.append({
                        "course": base_name,
                        "status": "error",
                        "message": "Required student id column is missing.",
                        "data": None
                    })
                    continue
                
                if 'Letter Grade' not in df_roster.columns:
                    results.append({
                        "course": base_name,
                        "status": "error",
                        "message": "'Letter Grade' column is missing.",
                        "data": None
                    })
                    continue
                
                if 'ID' not in df_downloaded.columns or 'Approved final grade' not in df_downloaded.columns:
                    results.append({
                        "course": base_name,
                        "status": "error",
                        "message": "Required columns are missing in downloaded file.",
                        "data": None
                    })
                    continue
                
                # Extract relevant columns
                df_roster_sub = df_roster[[sid_col, 'Letter Grade']].copy()
                if 'Withdrawn' in df_downloaded.columns:
                    df_downloaded_sub = df_downloaded[['ID', 'Approved final grade', 'Withdrawn']].copy()
                else:
                    df_downloaded_sub = df_downloaded[['ID', 'Approved final grade']].copy()
                    df_downloaded_sub['Withdrawn'] = ''
                
                # Clean data
                df_roster_sub[sid_col] = df_roster_sub[sid_col].astype(str).str.strip()
                df_downloaded_sub['ID'] = df_downloaded_sub['ID'].astype(str).str.strip()
                df_roster_sub['Letter Grade'] = df_roster_sub['Letter Grade'].astype(str).str.strip().str.upper()
                df_downloaded_sub['Approved final grade'] = df_downloaded_sub['Approved final grade'].astype(str).str.strip().str.upper()
                df_downloaded_sub['Withdrawn'] = df_downloaded_sub['Withdrawn'].astype(str).str.strip().str.upper()
                
                # Merge datasets
                merged = pd.merge(df_roster_sub, df_downloaded_sub, left_on=sid_col, right_on='ID', how='outer', indicator=True)
                merged['ID_final'] = merged[sid_col].combine_first(merged['ID']).astype(str).str.strip()
                merged = merged[merged['ID_final'].notna() & (merged['ID_final'] != '') & (merged['ID_final'].str.lower() != 'nan')]
                
                # Track withdrawals
                merged['is_withdrawn'] = (merged['Approved final grade'] == 'W') | (merged['Withdrawn'] == 'WITHDRAWN')
                withdrawn_ids = merged.loc[merged['is_withdrawn'], 'ID_final'].tolist()
                summary_stats["withdrawn_students"] += len(withdrawn_ids)
                
                # Check for grade mismatches
                merged['norm_roster_grade'] = merged['Letter Grade'].apply(normalize_grade)
                merged['norm_downloaded_grade'] = merged['Approved final grade'].apply(normalize_grade)
                merged['mismatch'] = (merged['_merge'] == 'both') & (
                    ((merged['norm_roster_grade'] == "ABSENT") & (merged['norm_downloaded_grade'] != "F")) |
                    ((merged['norm_roster_grade'] != "ABSENT") & (merged['norm_roster_grade'] != merged['norm_downloaded_grade']))
                )
                merged['matched'] = ~merged['mismatch']
                merged['course'] = base_name
                
                # Prepare results
                result = merged[['course', 'ID_final', 'Letter Grade', 'Approved final grade', 'matched', 'is_withdrawn']]
                result = result[result['ID_final'].str.isnumeric()]
                
                # Update summary statistics
                summary_stats["total_students"] += len(result)
                unmatched = result[~result['matched']]
                unmatched_count = unmatched.shape[0]
                
                if unmatched_count > 0:
                    summary_stats["courses_with_mismatches"] += 1
                    summary_stats["total_mismatches"] += unmatched_count
                    all_unmatched.append(unmatched)
                
                results.append({
                    "course": base_name,
                    "status": "success",
                    "message": f"Processed {len(result)} students, found {unmatched_count} mismatches, {len(withdrawn_ids)} withdrawn",
                    "data": result,
                    "unmatched": unmatched,
                    "withdrawn": withdrawn_ids
                })
                
            except Exception as e:
                results.append({
                    "course": base_name,
                    "status": "error",
                    "message": f"Error processing files: {str(e)}",
                    "data": None
                })
        else:
            results.append({
                "course": base_name,
                "status": "error",
                "message": "No matching downloaded file found",
                "data": None
            })
    
    return results, summary_stats, all_unmatched


def compare_grades(roster_files, downloaded_files):
    """
    Compare grades between roster files and downloaded files.
    Returns results and summary stats.
    """
    results = []
    all_unmatched = []
    # Track unique student IDs across all courses
    unique_student_ids = set()
    
    summary_stats = {
        "total_courses": 0,
        "courses_with_mismatches": 0,
        "total_mismatches": 0,
        "total_students": 0,  # This will now be calculated at the end
        "unique_students": 0, # New stat for unique students
        "withdrawn_students": 0
    }
    
    # Create a dictionary of downloaded files for easy lookup
    downloaded_dict = {os.path.splitext(os.path.basename(f.name))[0]: f for f in downloaded_files}
    
    for roster_file in roster_files:
        base_name = os.path.splitext(os.path.basename(roster_file.name))[0]
        
        # Check if we have a matching downloaded file
        if base_name in downloaded_dict:
            summary_stats["total_courses"] += 1
            downloaded_file = downloaded_dict[base_name]
            
            # Process the files
            try:
                # Create temporary files to save the uploaded files
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_roster:
                    tmp_roster.write(roster_file.getvalue())
                    tmp_roster_path = tmp_roster.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp_downloaded:
                    tmp_downloaded.write(downloaded_file.getvalue())
                    tmp_downloaded_path = tmp_downloaded.name
                
                # Read the files
                try:
                    df_roster = read_roster_file(tmp_roster_path)
                except ValueError as e:
                    results.append({
                        "course": base_name,
                        "status": "error",
                        "message": str(e),
                        "data": None
                    })
                    continue
                
                df_downloaded = pd.read_csv(tmp_downloaded_path)
                
                # Clean up temporary files
                os.unlink(tmp_roster_path)
                os.unlink(tmp_downloaded_path)
                
                # Clean column names
                df_roster.columns = df_roster.columns.str.strip()
                df_downloaded.columns = df_downloaded.columns.str.strip()
                
                # Check for required columns
                if 'SID' in df_roster.columns:
                    sid_col = 'SID'
                elif 'Student ID' in df_roster.columns:
                    sid_col = 'Student ID'
                else:
                    results.append({
                        "course": base_name,
                        "status": "error",
                        "message": "Required student id column is missing.",
                        "data": None
                    })
                    continue
                
                if 'Letter Grade' not in df_roster.columns:
                    results.append({
                        "course": base_name,
                        "status": "error",
                        "message": "'Letter Grade' column is missing.",
                        "data": None
                    })
                    continue
                
                if 'ID' not in df_downloaded.columns or 'Approved final grade' not in df_downloaded.columns:
                    results.append({
                        "course": base_name,
                        "status": "error",
                        "message": "Required columns are missing in downloaded file.",
                        "data": None
                    })
                    continue
                
                # Extract relevant columns
                df_roster_sub = df_roster[[sid_col, 'Letter Grade']].copy()
                if 'Withdrawn' in df_downloaded.columns:
                    df_downloaded_sub = df_downloaded[['ID', 'Approved final grade', 'Withdrawn']].copy()
                else:
                    df_downloaded_sub = df_downloaded[['ID', 'Approved final grade']].copy()
                    df_downloaded_sub['Withdrawn'] = ''
                
                # Clean data
                df_roster_sub[sid_col] = df_roster_sub[sid_col].astype(str).str.strip()
                df_downloaded_sub['ID'] = df_downloaded_sub['ID'].astype(str).str.strip()
                df_roster_sub['Letter Grade'] = df_roster_sub['Letter Grade'].astype(str).str.strip().str.upper()
                df_downloaded_sub['Approved final grade'] = df_downloaded_sub['Approved final grade'].astype(str).str.strip().str.upper()
                df_downloaded_sub['Withdrawn'] = df_downloaded_sub['Withdrawn'].astype(str).str.strip().str.upper()
                
                # Merge datasets
                merged = pd.merge(df_roster_sub, df_downloaded_sub, left_on=sid_col, right_on='ID', how='outer', indicator=True)
                merged['ID_final'] = merged[sid_col].combine_first(merged['ID']).astype(str).str.strip()
                merged = merged[merged['ID_final'].notna() & (merged['ID_final'] != '') & (merged['ID_final'].str.lower() != 'nan')]
                
                # Track withdrawals
                merged['is_withdrawn'] = (merged['Approved final grade'] == 'W') | (merged['Withdrawn'] == 'WITHDRAWN')
                withdrawn_ids = merged.loc[merged['is_withdrawn'], 'ID_final'].tolist()
                summary_stats["withdrawn_students"] += len(withdrawn_ids)
                
                # Add this course's student IDs to our unique student tracking set
                valid_ids = merged['ID_final'][merged['ID_final'].str.isnumeric()]
                unique_student_ids.update(valid_ids)
                
                # Check for grade mismatches
                merged['norm_roster_grade'] = merged['Letter Grade'].apply(normalize_grade)
                merged['norm_downloaded_grade'] = merged['Approved final grade'].apply(normalize_grade)
                merged['mismatch'] = (merged['_merge'] == 'both') & (
                    ((merged['norm_roster_grade'] == "ABSENT") & (merged['norm_downloaded_grade'] != "F")) |
                    ((merged['norm_roster_grade'] != "ABSENT") & (merged['norm_roster_grade'] != merged['norm_downloaded_grade']))
                )
                merged['matched'] = ~merged['mismatch']
                merged['course'] = base_name
                
                # Prepare results
                result = merged[['course', 'ID_final', 'Letter Grade', 'Approved final grade', 'matched', 'is_withdrawn']]
                result = result[result['ID_final'].str.isnumeric()]
                
                # Update total students (not unique) - keep this for per-course counting
                summary_stats["total_students"] += len(result)
                
                unmatched = result[~result['matched']]
                unmatched_count = unmatched.shape[0]
                
                if unmatched_count > 0:
                    summary_stats["courses_with_mismatches"] += 1
                    summary_stats["total_mismatches"] += unmatched_count
                    all_unmatched.append(unmatched)
                
                results.append({
                    "course": base_name,
                    "status": "success",
                    "message": f"Processed {len(result)} students, found {unmatched_count} mismatches, {len(withdrawn_ids)} withdrawn",
                    "data": result,
                    "unmatched": unmatched,
                    "withdrawn": withdrawn_ids
                })
                
            except Exception as e:
                results.append({
                    "course": base_name,
                    "status": "error",
                    "message": f"Error processing files: {str(e)}",
                    "data": None
                })
        else:
            results.append({
                "course": base_name,
                "status": "error",
                "message": "No matching downloaded file found",
                "data": None
            })
    
    # Set the unique students count
    summary_stats["unique_students"] = len(unique_student_ids)
    
    return results, summary_stats, all_unmatched

def create_download_link(df, filename):
    """
    Create a download link for a dataframe
    """
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
    return href

def get_all_results_df(results):
    """
    Combine all results into a single dataframe
    """
    dfs = []
    for result in results:
        if result["status"] == "success" and result["data"] is not None:
            dfs.append(result["data"])
    
    if dfs:
        return pd.concat(dfs, ignore_index=True)
    return None

def create_zip_file(directory):
    """Create a zipfile from a directory"""
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(directory):
            for file in files:
                zipf.write(os.path.join(root, file), 
                           os.path.relpath(os.path.join(root, file), 
                                         os.path.join(directory, '..')))
    memory_file.seek(0)
    return memory_file

# Main Streamlit app
st.title("Grade Checker App")

# Create a tabbed interface
tab1, tab2 = st.tabs(["Extract & Merge", "Compare Grades"])

# Tab 1: Extract & Merge
with tab1:
    st.header("Extract & Merge Files")
    
    st.markdown("""
    This section extracts files from a RAR archive and merges files that share a base code.
    
    **Example:**
    - MATH101_1.csv, MATH101_2.xlsx → MATH101.csv
    - CSAI330_1.xlsx, CSAI330_2.csv → CSAI330.csv
    """)
    
    rar_file = st.file_uploader("Upload RAR file", type=["rar"])
    
    if rar_file:
        st.success(f"✅ {rar_file.name} uploaded")
        
        if st.button("Extract & Merge Files", key="extract_btn"):
            # Save uploaded file to a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.rar') as tmp_file:
                tmp_file.write(rar_file.getvalue())
                tmp_file_path = tmp_file.name
            
            # Create output directory
            output_dir = tempfile.mkdtemp()
            
            with st.spinner("Extracting and merging files..."):
                try:
                    # Call extract_and_merge function
                    output_folder, processed_files, merged_files = extract_and_merge(tmp_file_path, output_dir)
                    
                    # Cleanup
                    os.unlink(tmp_file_path)
                    
                    st.success(f"Processing complete! {len(merged_files)} merged files created.")
                    
                    # Show processed files
                    with st.expander("Files processed"):
                        for file in processed_files:
                            st.text(file)
                    
                    # Show merged files
                    with st.expander("Files merged"):
                        for file in merged_files:
                            st.text(file)
                    
                    # Create a download link for all merged files
                    merged_files_zip = create_zip_file(output_folder)
                    st.download_button(
                        label="Download All Merged Files",
                        data=merged_files_zip,
                        file_name="merged_files.zip",
                        mime="application/zip"
                    )
                    
                    # Store the output folder path in session state
                    st.session_state.merged_folder = output_folder
                    st.session_state.has_merged_files = True
                    
                    # Add a button to go to the next tab
                    st.success("Files are ready! Click the 'Compare Grades' tab to continue.")
                    
                except Exception as e:
                    st.error(f"Error during extraction: {str(e)}")
                    if os.path.exists(tmp_file_path):
                        os.unlink(tmp_file_path)
                    if os.path.exists(output_dir):
                        shutil.rmtree(output_dir)

# Tab 2: Compare Grades
with tab2:
    st.header("Compare Grades")
    
    st.markdown("""
    This section compares grades between roster files and downloaded grade files.
    
    ### Instructions:
    1. Upload Excel roster files (.xlsx or .xls)
    2. Upload downloaded CSV grade files (.csv)
    3. Click the 'Compare Grades' button
    4. Review the results and download comparison files
    """)
    
    # Check if we have merged files from the first tab
    if 'has_merged_files' in st.session_state and st.session_state.has_merged_files:
        st.info("Merged files from extraction are available. You can use them or upload new files.")
        
        if st.button("Use Merged Files as Downloaded Files"):
            # Get list of CSV files from the merged folder
            merged_files = [f for f in glob.glob(os.path.join(st.session_state.merged_folder, "*.csv"))]
            st.session_state.downloaded_files_paths = merged_files
            st.success(f"✓ Using {len(merged_files)} merged files as downloaded files")
    
    # File upload sections
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Upload Roster Files")
        roster_files = st.file_uploader("Select roster Excel files", type=["xlsx", "xls"], accept_multiple_files=True, key="roster_uploader")
        if roster_files:
            st.success(f"✅ {len(roster_files)} roster files uploaded")
            roster_names = [file.name for file in roster_files]
            st.write("Uploaded files:", ", ".join(roster_names))
    
    with col2:
        st.subheader("Upload Downloaded Files")
        # Only show this uploader if we're not using merged files
        if 'downloaded_files_paths' not in st.session_state:
            downloaded_files = st.file_uploader("Select downloaded CSV files", type=["csv"], accept_multiple_files=True, key="downloaded_uploader")
            if downloaded_files:
                st.success(f"✅ {len(downloaded_files)} downloaded files uploaded")
                downloaded_names = [file.name for file in downloaded_files]
                st.write("Uploaded files:", ", ".join(downloaded_names))
        else:
            # Display the names of the merged files we're using
            merged_files = st.session_state.downloaded_files_paths
            st.success(f"✅ Using {len(merged_files)} merged files")
            merged_names = [os.path.basename(file) for file in merged_files]
            st.write("Files:", ", ".join(merged_names))
            
            # Add option to clear and upload new files
            if st.button("Clear and Upload New Files"):
                del st.session_state.downloaded_files_paths
                st.experimental_rerun()
    
    # Check if we should use merged files as downloaded files
    if 'downloaded_files_paths' in st.session_state:
        # Convert file paths to file-like objects
        file_objects = []
        for file_path in st.session_state.downloaded_files_paths:
            file_name = os.path.basename(file_path)
            with open(file_path, 'rb') as f:
                # Create a file-like object
                file_obj = io.BytesIO(f.read())
                file_obj.name = file_name
                file_objects.append(file_obj)
        
        downloaded_files = file_objects
    
    # Process files when button is clicked
    compare_ready = roster_files and (downloaded_files if 'downloaded_files' in locals() else False)
    
    if st.button("Compare Grades", type="primary", disabled=not compare_ready):
        with st.spinner("Processing files..."):
            results, summary_stats, all_unmatched = compare_grades(roster_files, downloaded_files)
        
        # Display summary
        st.header("Summary")
        st.metric("Total Courses Processed", summary_stats["total_courses"])
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Students", summary_stats["unique_students"])
        with col2:
            st.metric("Total Mismatches", summary_stats["total_mismatches"])
        with col3:
            st.metric("Withdrawn Students", summary_stats["withdrawn_students"])
        
        # Create output directory for comparison results
        output_dir = tempfile.mkdtemp()
        
        # Save all results
        all_results_df = get_all_results_df(results)
        if all_results_df is not None:
            all_results_path = os.path.join(output_dir, "all_results.csv")
            all_results_df.to_csv(all_results_path, index=False)
        
        # Save all mismatches
        if all_unmatched:
            all_unmatched_df = pd.concat(all_unmatched, ignore_index=True)
            all_unmatched_path = os.path.join(output_dir, "all_mismatches.csv")
            all_unmatched_df.to_csv(all_unmatched_path, index=False)
        
        # Save individual course results
        for result in results:
            if result["status"] == "success" and result["data"] is not None:
                course_path = os.path.join(output_dir, f"{result['course']}_comparison.csv")
                result["data"].to_csv(course_path, index=False)
        
        # Display detailed results
        st.header("Course Results")
        
        # Create tabs for each course result
        tabs = st.tabs([result["course"] for result in results])
        
        for i, tab in enumerate(tabs):
            result = results[i]
            with tab:
                if result["status"] == "success":
                    st.success(result["message"])
                    
                    # Display withdrawn students if any
                    if result["withdrawn"]:
                        with st.expander(f"Withdrawn Students ({len(result['withdrawn'])})"):
                            st.write(", ".join(result["withdrawn"]))
                    
                    # Display mismatches if any
                    if len(result["unmatched"]) > 0:
                        st.subheader(f"Grade Mismatches ({len(result['unmatched'])})")
                        st.dataframe(
                            result["unmatched"][['ID_final', 'Letter Grade', 'Approved final grade']],
                            use_container_width=True
                        )
                    else:
                        st.success("No grade mismatches found!")
                    
                    # Display all data
                    with st.expander("View All Data"):
                        st.dataframe(result["data"], use_container_width=True)
                    
                    # Download link for this course
                    st.markdown(create_download_link(result["data"], f"{result['course']}_comparison.csv"), unsafe_allow_html=True)
                else:
                    st.error(result["message"])
        
        # All mismatches section
        if all_unmatched:
            st.header("All Mismatches")
            all_unmatched_df = pd.concat(all_unmatched, ignore_index=True)
            st.dataframe(all_unmatched_df[['course', 'ID_final', 'Letter Grade', 'Approved final grade']], use_container_width=True)
            
            # Download link for all mismatches
            st.markdown(create_download_link(all_unmatched_df, "all_mismatches.csv"), unsafe_allow_html=True)
        
        # Download all results
        if all_results_df is not None:
            st.header("Download Results")
            
            # Download link for all results
            st.markdown(create_download_link(all_results_df, "all_results.csv"), unsafe_allow_html=True)
            
            # Create a download button for all files
            results_zip = create_zip_file(output_dir)
            st.download_button(
                label="Download All Comparison Files (ZIP)",
                data=results_zip,
                file_name="comparison_results.zip",
                mime="application/zip"
            )

# Footer
st.divider()
st.caption("Grade Checker App")