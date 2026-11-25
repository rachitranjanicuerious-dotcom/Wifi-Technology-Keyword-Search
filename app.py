# from pathlib import Path
# import pandas as pd
# import streamlit as st
# from docx import Document
# from pptx import Presentation
# import re
# import zipfile
# import io

# # =========================================================
# # 1. PAGE CONFIGURATION
# # =========================================================

# st.set_page_config(
#     page_title="Wi-Fi Technical Document Search",
#     page_icon="🔎",
#     layout="wide" 
# )

# st.title("Wi-Fi Technical Document Search")

# st.write(
#     "Search technical documents in the Sample Wi-Fi repository "
#     "using exact keyword or phrase matching."
# )

# # Function to Zip the retieved files
# def create_zip_file(results):

#     zip_buffer = io.BytesIO()

#     with zipfile.ZipFile(
#         zip_buffer,
#         mode="w",
#         compression=zipfile.ZIP_DEFLATED
#     ) as zip_file:

#         for result in results:

#             file_path = Path(result["Document Path"])

#             if file_path.exists():
#                 zip_file.write(
#                     file_path,
#                     arcname=file_path.name
#                 )

#     zip_buffer.seek(0)
#     return zip_buffer

# # =========================================================
# # 2. SUPPORTED FILE FORMATS
# # =========================================================

# SUPPORTED_EXTENSIONS = {
#     ".docx" , 
#     ".pptx" ,
#     ".xlsx"
# }

# # =========================================================
# # 3. TEXT CLEANING
# # =========================================================

# def clean_text(text):

#     if text is None:
#         return ""

#     return " ".join(str(text).split())

# def normalize_for_search(text):
#     # Convert hyphens and underscores into spaces
#     text = re.sub(r"[-_]+", " ", str(text))

#     # Convert multiple spaces/newlines into a single space
#     text = re.sub(r"\s+", " ", text)

#     # Ignore uppercase and lowercase differences
#     return text.casefold().strip()

# # =========================================================
# # 4. WORD DOCUMENT EXTRACTION
# # =========================================================

# def extract_word(file_path):

#     doc = Document(file_path)

#     extracted_text = []

#     # Extract paragraphs
#     for paragraph in doc.paragraphs:

#         text = clean_text(paragraph.text)

#         if text:
#             extracted_text.append(text)

#     for table in doc.tables:         
#         for row in table.rows:
#             row_text = []

#             for cell in row.cells:

#                 text = clean_text(cell.text)

#                 if text:

#                     row_text.append(text)

#             if row_text:
#                 extracted_text.append(" ".join(row_text))


#     return "\n".join(extracted_text)            

# # =========================================================
# # 5. POWERPOINT DOCUMENT EXTRACTION
# # =========================================================

# def extract_powerpoint(file_path):

#     presentation = Presentation(file_path)

#     extracted_text = []

#     for slide in presentation.slides:
         
#          for shape in slide.shapes:
#             # Extract text from text boxes
#             if getattr(shape , "has_text_frame" , False):

#                 text = clean_text(shape.text)

#                 if text:
#                     extracted_text.append(text)

#             if getattr(shape, "has_table" , False):
#                 for row in shape.table.rows:
#                     row_text = []

#                     for cell in row.cells:

#                         text = clean_text(cell.text)    

#                         if text:
#                             row_text.append(text)   

#                     if row_text:
#                         extracted_text.append(" ".join(row_text)) 

#     return "\n".join(extracted_text)                                            

# =========================================================
# 6. EXCEL DOCUMENT EXTRACTION
# =========================================================

# def extract_excel(file_path):

#     extracted_text = []

#     excel_file = pd.ExcelFile(
#         file_path,
#         engine = "openpyxl"
#     )

#     # Extract all worksheets
#     for sheet_name in excel_file.sheet_names:

#         df = pd.read_excel(
#             excel_file, 
#             sheet_name = sheet_name,
#             header = None,
#             dtype = object
#         )

#         # Extract all rows and columns
#         for row in df.itertuples(
#             index = False,
#             name = None
#         ):

#             row_text = []

#             for value in row:

#                 if pd.isna(value):
#                     continue

#                 text = clean_text(value)

#                 if text:
#                     row_text.append(text)
 
#             if row_text:        
#                 extracted_text.append( " ".join(row_text))

#     return "\n".join(extracted_text)

# # # =========================================================
# # # 7. EXTRACT TEXT BASED ON FILE EXTENSION
# # # =========================================================

# def extract_document_text(file_path):

#     extension = file_path.suffix.lower()

#     if extension == ".pptx":
#         return extract_powerpoint(file_path)

#     elif extension == ".xlsx":
#         return extract_excel(file_path)

#     elif extension == ".docx":
#         return extract_word(file_path)

#     return ""

# # =========================================================
# # 8. SEARCH ALL DOCUMENTS
# # =========================================================

# def search_documents(repository_path , search_keyword):

#     repository = Path(repository_path)

#     results = []
#     errors = []

#     keyword = normalize_for_search(search_keyword)

#     # Create whole-word matching pattern
#     pattern = r"(?<![A-Za-z0-9])" + re.escape(keyword) + r"(?![A-Za-z0-9])"

#     total_files = 0
#     searched_files = 0

#     # Recursively scan all folders
#     all_files = [
#     file_path
#     for file_path in repository.rglob("*")
#     if file_path.is_file()
#     and file_path.suffix.lower() in SUPPORTED_EXTENSIONS
#     and not file_path.name.startswith("~$")
# ]

#     total_files = len(all_files)
#     progress = st.progress(0)
#     status = st.empty()

#     for index, file_path in enumerate(all_files):

#         try:

#             status.write(
#                  f"Searching: {file_path.name}"
#             )

#             document_text = extract_document_text(file_path)
#             normalized_document = normalize_for_search(document_text)
#             searched_files += 1

#             # Check for exact word or phrase
#             if re.search(pattern, normalized_document):
#                     results.append({
#                         "Document Name": file_path.name,
#                         "File Type": file_path.suffix.lower(),
#                         "Document Path": str(file_path)
#                     })

#         except Exception as error:

#             errors.append
#             ({
#                 "File": str(file_path.relative_to(repository))
#                 ,

#                 "Error" : str(error)
#             })

#         progress.progress(
#             (index + 1) / total_files
#         )     


#     progress.empty()
#     status.empty()

#     return results, errors, total_files, searched_files

# # =========================================================
# # 9. SIDEBAR SETTINGS
# # =========================================================

# st.sidebar.header("Repository Settings")

# BASE_DIR = Path(__file__).resolve().parent

# DEFAULT_REPOSITORY = (
#     BASE_DIR
#     / "Sample Wi-Fi 7 WG Doc"
#     / "Sample Wi-Fi 7 WG Doc"
# )

# repository_path = st.sidebar.text_input(
#     "Repository Folder Path",
#     value=str(DEFAULT_REPOSITORY)
# )

# # repository_path = st.sidebar.text_input(
# #                             "Repository Folder Path",
# #                              value="Sample Wi-Fi 7 WG Doc/Sample Wi-Fi 7 WG Doc"
# # )

# st.sidebar.info(
#       "The application searches all subfolders inside the repository"
# )

# # =========================================================
# # 10. SEARCH INTERFACE
# # =========================================================

# st.subheader("Seach Technical Documents")

# search_keyword = st.text_input(
#     "Enter keyword or phrase",
#     placeholder = "Example: Multi-Link Operation"
# )

# search_button = st.button(
#     "Search Documents",
#     type = "primary",
#     use_container_width = True
# )

# # =========================================================
# # 11. SEARCH EXECUTION
# # =========================================================

# if search_button:

#     if not search_keyword.strip():

#         st.warning(
#             "Please enter a keyword or phrase"
#         )

#     elif not Path(repository_path).exists():

#         st.error("Repository folder not found. "
#             "Please enter the correct folder path.")     

#     else:

#         with st.spinner("Searching documents...."):

#             results, errors, total_files, searched_files = (
#                 search_documents(
#                     repository_path,
#                     search_keyword
#                 )
#             )

#         st.divider()    

#         st.subheader("Search Summary")

#         col1, col2, col3 = st.columns(3)

#         col1.metric(
#             "Total Documents",
#             total_files
#         )

#         col2.metric(
#             "Documents Search" ,
#             searched_files)

#         col3.metric
#         (
#             "Matching Documents" ,
#             len(results)
#         )


#         st.write(
#             f"**Search keyword:** {search_keyword}"
#         )

#         # =================================================
#         # DISPLAY MATCHING DOCUMENTS
#         # =================================================

#         st.subheader("Matching Documents")

#         if results:

#             st.success(
#                 f"Found {len(results)} documents "
#                 "containing the searched keyword."
#             )

#             results_df = pd.DataFrame(results)

#             st.dataframe(
#                 results_df[
#                     [
#                         "Document Name",
#                         "File Type",
#                         "Document Path"
#                     ]
#                 ],
#                 use_container_width = True,
#                 hide_index = True
#             )

#             csv_data = results_df.to_csv(
#                 index = False
#             )

#             st.download_button(
#                 label = "Download Matching Documents",
#                 data = csv_data,
#                 file_name="matching_documents.csv",
#                 mime="text/csv"
#             )

#             zip_file = create_zip_file(results)

#             st.download_button(
#             label="Download Retrieved Files (ZIP)",
#             data=zip_file,
#             file_name="Retrieved_Documents.zip",
#             mime="application/zip"
#         )


#         else:

#             st.info(
#                 "No documents containing the searched keyword were found."
#             )

#         if errors:

#             with st.expander(
#                 f"Files that could not be searched ({len(errors)})"
#             ) :

#                  for error in errors:

#                     st.write(
#                         f"**{error['File']}**"
#                     )
#                     st.caption(
#                         error["Error"]
#                     )

# # =========================================================
# # 12. SEARCH INFORMATION
# # =========================================================

# with st.expander("Search Information") :

#     st.markdown(
#          """

#          **Supported formats**

#         - Word: DOCX
#         - PowerPoint: PPTX
#         - Excel: XLSX

# """
#     )


# st.caption(
#          "Wi-Fi Technical Document Search | Keyword-Based Retrieval"
# )




from pathlib import Path
import pandas as pd
import streamlit as st
from docx import Document
from pptx import Presentation
import re
import zipfile
import io
import html


# =========================================================
# 1. PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="iCuerious | Wi-Fi Technical Document Search",
    page_icon="🔎",
    layout="wide"
)

BASE_DIR = Path(__file__).resolve().parent


# =========================================================
# 2. LOAD EXTERNAL CSS AND HTML
# =========================================================

def load_css():

    css_path = BASE_DIR / "styles.css"

    if css_path.exists():

        css = css_path.read_text(encoding="utf-8")

        st.html(f"<style>{css}</style>")

    else:

        st.warning("styles.css file not found.")


def load_html(filename):

    html_path = BASE_DIR / filename

    if html_path.exists():

        content = html_path.read_text(encoding="utf-8")

        st.html(content)

    else:

        st.warning(f"{filename} file not found.")


# Load CSS before rendering application components
load_css()


# =========================================================
# 3. ZIP DOWNLOAD
# =========================================================

def create_zip_file(results):

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED
    ) as zip_file:

        for result in results:

            file_path = Path(result["Document Path"])

            if file_path.exists():

                zip_file.write(
                    file_path,
                    arcname=file_path.name
                )

    zip_buffer.seek(0)

    return zip_buffer


# =========================================================
# 4. SUPPORTED FILE FORMATS
# =========================================================

SUPPORTED_EXTENSIONS = {
    ".docx",
    ".pptx",
    ".xlsx"
}


# =========================================================
# 5. TEXT CLEANING AND NORMALIZATION
# =========================================================

def clean_text(text):

    if text is None:
        return ""

    return " ".join(str(text).split())


def normalize_for_search(text):

    # Convert hyphens and underscores into spaces
    text = re.sub(r"[-_]+", " ", str(text))

    # Remove multiple spaces and newlines
    text = re.sub(r"\s+", " ", text)

    # Ignore uppercase and lowercase differences
    return text.casefold().strip()


# =========================================================
# 6. WORD DOCUMENT EXTRACTION
# =========================================================

def extract_word(file_path):

    doc = Document(file_path)

    extracted_text = []

    # Extract paragraphs
    for paragraph in doc.paragraphs:

        text = clean_text(paragraph.text)

        if text:
            extracted_text.append(text)

    # Extract tables
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


# =========================================================
# 7. POWERPOINT DOCUMENT EXTRACTION
# =========================================================

def extract_powerpoint(file_path):

    presentation = Presentation(file_path)

    extracted_text = []

    for slide in presentation.slides:

        for shape in slide.shapes:

            # Extract text from text boxes
            if getattr(shape, "has_text_frame", False):

                text = clean_text(shape.text)

                if text:
                    extracted_text.append(text)

            # Extract tables
            if getattr(shape, "has_table", False):

                for row in shape.table.rows:

                    row_text = []

                    for cell in row.cells:

                        text = clean_text(cell.text)

                        if text:
                            row_text.append(text)

                    if row_text:
                        extracted_text.append(" ".join(row_text))

    return "\n".join(extracted_text)


# =========================================================
# 8. EXCEL DOCUMENT EXTRACTION
# =========================================================

def extract_excel(file_path):

    extracted_text = []

    excel_file = pd.ExcelFile(
        file_path,
        engine="openpyxl"
    )

    # Extract all worksheets
    for sheet_name in excel_file.sheet_names:

        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=None,
            dtype=object
        )

        # Extract all rows and columns
        for row in df.itertuples(
            index=False,
            name=None
        ):

            row_text = []

            for value in row:

                if pd.isna(value):
                    continue

                text = clean_text(value)

                if text:
                    row_text.append(text)

            if row_text:
                extracted_text.append(" ".join(row_text))

    return "\n".join(extracted_text)


# =========================================================
# 9. EXTRACT TEXT BASED ON FILE EXTENSION
# =========================================================

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
# 10. SEARCH ALL DOCUMENTS
# =========================================================

def search_documents(repository_path, search_keyword):

    repository = Path(repository_path)

    results = []
    errors = []

    keyword = normalize_for_search(search_keyword)

    # Create whole-word matching pattern
    pattern = (
        r"(?<![A-Za-z0-9])"
        + re.escape(keyword)
        + r"(?![A-Za-z0-9])"
    )

    # Recursively scan all folders
    all_files = [

        file_path

        for file_path in repository.rglob("*")

        if file_path.is_file()

        and file_path.suffix.lower() in SUPPORTED_EXTENSIONS

        and not file_path.name.startswith("~$")

    ]

    total_files = len(all_files)
    searched_files = 0

    # Progress indicators
    progress = st.progress(0)
    status = st.empty()

    for index, file_path in enumerate(all_files):

        try:

            status.caption(
                f"Searching: {file_path.name}"
            )

            # Extract document text
            document_text = extract_document_text(file_path)

            # Normalize extracted text
            normalized_document = normalize_for_search(
                document_text
            )

            searched_files += 1

            # Check exact word or phrase
            if re.search(pattern, normalized_document):

                results.append({

                    "Document Name": file_path.name,

                    "File Type": file_path.suffix.lower(),

                    "Document Path": str(file_path)

                })

        except Exception as error:

            errors.append({

                "File": str(
                    file_path.relative_to(repository)
                ),

                "Error": str(error)

            })

        # Update progress
        if total_files:

            progress.progress(
                (index + 1) / total_files
            )

    progress.empty()
    status.empty()

    return (
        results,
        errors,
        total_files,
        searched_files
    )


# =========================================================
# 11. SIDEBAR SETTINGS
# =========================================================

DEFAULT_REPOSITORY = (
    BASE_DIR
    / "Sample Wi-Fi 7 WG Doc"
    / "Sample Wi-Fi 7 WG Doc"
)


# Sidebar branding
with st.sidebar:

    st.html(
        """
        <div class="ic-sidebar-brand">

            <div class="ic-brand-mark">
                iC
            </div>

            <div class="ic-sidebar-brand-text">

                <strong>iCuerious</strong>

                <small>WORKSPACE</small>

            </div>

        </div>
        """
    )

    st.markdown("### Repository Settings")

    repository_path = st.text_input(
        "Repository Folder Path",
        value=str(DEFAULT_REPOSITORY)
    )

    st.info(
        "The application searches all supported files "
        "in every subfolder of the selected repository."
    )

    st.html(
        """
        <div class="ic-sidebar-note">

            <span class="ic-sidebar-dot"></span>

            <div>

                <strong>Local Repository</strong>

                <small>
                    Files are searched from the folder path
                    configured above.
                </small>

            </div>

        </div>
        """
    )


# =========================================================
# 12. MAIN SEARCH INTERFACE
# =========================================================

# Load header and hero section
load_html("header.html")


# Search section heading
st.html(
    """
    <div class="ic-section-heading">
        Search Your Repository
    </div>
    """
)

st.caption(
    "Enter a technical keyword or an exact phrase "
    "to locate matching documents."
)


# Search input
search_keyword = st.text_input(
    "Keyword or phrase",
    placeholder="Example: Multi-Link Operation",
    label_visibility="collapsed"
)


# Search button
search_button = st.button(
    "Search Documents  →",
    type="primary",
    use_container_width=True
)


# =========================================================
# 13. SEARCH EXECUTION AND RESULTS
# =========================================================

if search_button:

    # Validate search keyword
    if not search_keyword.strip():

        st.warning(
            "Please enter a keyword or phrase."
        )

    # Validate repository path
    elif not Path(repository_path).is_dir():

        st.error(
            "Repository folder not found. "
            "Please enter a valid folder path."
        )

    else:

        # Execute search
        with st.spinner(
            "Searching documents in the repository..."
        ):

            (
                results,
                errors,
                total_files,
                searched_files
            ) = search_documents(
                repository_path,
                search_keyword
            )

        st.divider()


        # =================================================
        # SEARCH SUMMARY
        # =================================================

        st.html(
            """
            <div class="ic-section-heading">
                Search Summary
            </div>
            """
        )

        # Three metric cards
        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Total Documents",
            total_files
        )

        col2.metric(
            "Documents Searched",
            searched_files
        )

        col3.metric(
            "Matching Documents",
            len(results)
        )

        # Display search keyword
        safe_keyword = html.escape(search_keyword)

        st.html(
            f"""
            <div class="ic-query-chip">

                SEARCH QUERY &nbsp;&nbsp;

                <strong>{safe_keyword}</strong>

            </div>
            """
        )


        # =================================================
        # MATCHING DOCUMENTS
        # =================================================

        st.html(
            """
            <div class="ic-section-heading">
                Matching Documents
            </div>
            """
        )

        if results:

            st.success(
                f"Found {len(results)} matching document(s)."
            )

            results_df = pd.DataFrame(results)

            # Display results table
            st.dataframe(
                results_df[
                    [
                        "Document Name",
                        "File Type",
                        "Document Path"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )


            # =================================================
            # DOWNLOAD OPTIONS
            # =================================================

            download_col1, download_col2 = st.columns(2)

            # Download CSV
            with download_col1:

                st.download_button(
                    label="Download Results (CSV)",

                    data=results_df.to_csv(index=False),

                    file_name="matching_documents.csv",

                    mime="text/csv",

                    use_container_width=True
                )

            # Download ZIP
            with download_col2:

                zip_file = create_zip_file(results)

                st.download_button(
                    label="Download Retrieved Files (ZIP)",

                    data=zip_file,

                    file_name="Retrieved_Documents.zip",

                    mime="application/zip",

                    use_container_width=True
                )

        else:

            st.info(
                "No documents containing the searched "
                "keyword were found."
            )


        # =================================================
        # FILES THAT COULD NOT BE SEARCHED
        # =================================================

        if errors:

            with st.expander(
                f"Files that could not be searched ({len(errors)})"
            ):

                for error in errors:

                    st.write(
                        f"**{error['File']}**"
                    )

                    st.caption(
                        error["Error"]
                    )


# =========================================================
# 14. SEARCH INFORMATION
# =========================================================

with st.expander("Search Information"):

    st.markdown(
        """
        **Supported File Formats**

        - Microsoft Word: DOCX
        - Microsoft PowerPoint: PPTX
        - Microsoft Excel: XLSX
        """
    )


# =========================================================
# 15. FOOTER
# =========================================================

st.html(
    """
    <div class="ic-footer">

        iCuerious &nbsp; | &nbsp;
        Wi-Fi Technical Document Search &nbsp; | &nbsp;
        Keyword-Based Retrieval

    </div>
    """
)
