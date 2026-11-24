from pathlib import Path
import pandas as pd
import streamlit as st
from docx import Document
from pptx import Presentation

# =========================================================
# 1. PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Wi-Fi Technical Document Search",
    page_icon="🔎",
    layout="wide" 
)

st.title("Wi-Fi Technical Document Search")

st.write(
    "Search technical documents in the Sample Wi-Fi repository "
    "using exact keyword or phrase matching."
)

# =========================================================
# 2. SUPPORTED FILE FORMATS
# =========================================================

SUPPORTED_EXTENSIONS = {
    ".docx" , 
    ".pptx" ,
    ".xlsx"
}

# =========================================================
# 3. TEXT CLEANING
# =========================================================

def clean_text(text):

    if text is None:
        return ""

    return " ".join(str(text).split())

# =========================================================
# 4. WORD DOCUMENT EXTRACTION
# =========================================================

def extract_word(file_path):

    doc = Document(file_path)

    extracted_text = []

    # Extract paragraphs
    for paragraph in doc.paragraphs:

        text = clean_text(paragraph.text)

        if text:
            extracted_text.append(text)

    for table in doc.tables:         
        for row in table.rows:
            row_text = []

            for cell in row.cells:

                text = clean_text(cell.text)

                if text:

                    row_text.append(text)

            if row_text:
                extracted_text.append(" ".join(row_text))


    return "\n".join(extracted_text)            

# # =========================================================
# # 5. POWERPOINT DOCUMENT EXTRACTION
# # =========================================================

def extract_powerpoint(file_path):

    presentation = Presentation(file_path)

    extracted_text = []

    for slide in presentation.slides:
         
         for shape in slide.shapes:
            # Extract text from text boxes
            if getattr(shape , "has_text_frame" , False):

                text = clean_text(shape.text)

                if text:
                    extracted_text.append(text)

            if getattr(shape, "has_table" , False):
                for row in shape.table.rows:
                    row_text = []

                    for cell in row.cells:

                        text = clean_text(cell.text)    

                        if text:
                            row_text.append(text)   

                    if row_text:
                        extracted_text.append(" ".join(row_text)) 

    return "\n".join(extracted_text)                                            

# # =========================================================
# # 6. EXCEL DOCUMENT EXTRACTION
# # =========================================================

def extract_excel(file_path):

    extracted_text = []

    excel_file = pd.ExcelFile(
        file_path,
        engine = "openpyxl"
    )

    # Extract all worksheets
    for sheet_name in excel_file.sheet_names:

        df = pd.read_excel(
            excel_file, 
            sheet_name = sheet_name,
            header = None,
            dtype = object
        )

        # Extract all rows and columns
        for row in df.itertuples(
            index = False,
            name = None
        ):

            row_text = []

            for value in row:

                if pd.isna(value):
                    continue

                text = clean_text(value)

                if text:
                    row_text.append(text)
 
            if row_text:        
                extracted_text.append( " ".join(row_text))

    return "\n".join(extracted_text)

# # =========================================================
# # 7. EXTRACT TEXT BASED ON FILE EXTENSION
# # =========================================================

def extract_document_text(file_path):

    extension = file_path.suffix.lower()

    if extension == ".pptx":
        return extract_powerpoint(file_path)

    elif extension == ".xlsx":
        return extract_excel(file_path)

    elif extension == ".docx":
        return extract_word(file_path)

    return ""

# =========================================================
# 8. SEARCH ALL DOCUMENTS
# =========================================================

def search_documents(repository_path , search_keyword):

    repository = Path(repository_path)

    results = []
    errors = []

    keyword = search_keyword.casefold().strip()

    total_files = 0
    searched_files = 0

    # Recursively scan all folders
    all_files = [
        file_path
        for file_path in repository.glob("*")
        if file_path.is_file()
        and file_path.suffix.lower() in SUPPORTED_EXTENSIONS
        and not file_path.name.startswith("~$")

    ]

    total_files = len(all_files)
    progress = st.progress(0)
    status = st.empty()

    for index, file_path in enumerate(all_files):

        try:

            status.write(
                 f"Searching: {file_path.name}"
            )

            document_text = extract_document_text(file_path)

            searched_files += 1

            if keyword in document_text.casefold():

                results.append({

                    "Document Name" : file_path.name,
                    "File Type":   file_path.suffix.lower(),
                    "Folder" : str(file_path.parent.relative_to(repository)) ,
                    "Document Path" : str(file_path)
                })


        except Exception as error:

            errors.append
            ({
                "File": str(file_path.relative_to(repository))
                ,

                "Error" : str(error)
            })

        progress.progress(
            (index + 1) / total_files
        )     


    progress.empty()
    status.empty()

    return results, errors, total_files, searched_files

# =========================================================
# 9. SIDEBAR SETTINGS
# =========================================================

st.sidebar.header("Repository Settings")

repository_path = st.sidebar.text_input(
                            "Repository Folder Path",
                             value="Sample Wi-Fi 7 WG Doc/Sample Wi-Fi 7 WG Doc"
)

st.sidebar.info(
      "The application searches all subfolders inside the repository"
)

# =========================================================
# 10. SEARCH INTERFACE
# =========================================================

st.subheader("Seach Technical Documents")

search_keyword = st.text_input(
    "Enter keyword or phrase",
    placeholder = "Example: Multi-Link Operation"
)

search_button = st.button(
    "Search Documents",
    type = "primary",
    use_container_width = True
)

# =========================================================
# 11. SEARCH EXECUTION
# =========================================================

if search_button:

    if not search_keyword.strip():

        st.warning(
            "Please enter a keyword or phrase"
        )

    elif not Path(repository_path).exists():

        st.error("Repository folder not found. "
            "Please enter the correct folder path.")     

    else:

        with st.spinner("Searching documents...."):

            results, errors, total_files, searched_files = (
                search_documents(
                    repository_path,
                    search_keyword
                )
            )

        st.divider()    

        st.subheader("Search Summary")

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Total Documents",
            total_files
        )

        col2.metric(
            "Documents Search" ,
            searched_files)

        col3.metric
        (
            "Matching Documents" ,
            len(results)
        )


        st.write(
            f"**Search keyword:** {search_keyword}"
        )

        # =================================================
        # DISPLAY MATCHING DOCUMENTS
        # =================================================

        st.subheader("Matching Documents")

        if results:

            st.success(
                f"Found {len(results)} documents "
                "containing the searched keyword."
            )

            results_df = pd.DataFrame(results)

            st.dataframe(
                results_df[
                    [
                        "Document Name",
                        "File Type",
                        "Folder",
                        "Document Path"
                    ]
                ],
                use_container_width = True,
                hide_index = True
            )

            csv_data = results_df.to_csv(
                index = False
            )

            st.download_button(
                label = "Download Matching Documents",
                data = csv_data,
                file_name="matching_documents.csv",
                mime="text/csv"
            )


        else:

            st.info(
                "No documents containing the searched keyword were found."
            )

        if errors:

            with st.expander(
                f"Files that could not be searched ({len(errors)})"
            ) :

                 for error in errors:

                    st.write(
                        f"**{error['File']}**"
                    )
                    st.caption(
                        error["Error"]
                    )

# =========================================================
# 12. SEARCH INFORMATION
# =========================================================

with st.expander("Search Information") :

    st.markdown(
         """

         **Supported formats**

        - Word: DOCX
        - PowerPoint: PPTX
        - Excel: XLSX

        **Search methodology**

        - Exact literal keyword and phrase matching.
        - Case-insensitive search.
        - Searches all files recursively across repository folders.
        - Returns the names and paths of matching documents.

        **Limitations**

        - Text inside images and scanned documents is not extracted.
        - Legacy DOC, PPT and XLS formats are not supported.
        - Excel formulas are not recalculated; saved cell values are searched.
        - A phrase must appear contiguously in the extracted text.
"""
    )


st.caption(
         "Wi-Fi Technical Document Search | Keyword-Based Retrieval"
)