from pathlib import Path

import docx
import docx2txt
from langchain_community.document_loaders import UnstructuredWordDocumentLoader
from spire.doc.common import *
from spire.doc import *


def extract_images_from_docx(docx_path, output_folder, return_text: bool = False):
    """Extracts images from a Word document and saves them to a folder."""

    # Extract text and image data from the document
    text = docx2txt.process(docx_path, output_folder)
    if return_text:
        return text


def extract_text_and_tables_from_docx(docx_path):
    doc = docx.Document(docx_path)
    doc.save(".\\test_save.docx")


def load_docx_unstructured(docx_file_path):
    doc = UnstructuredWordDocumentLoader(docx_file_path, mode="elements")
    data = doc.load()
    return data


def docx_to_html(docx_file_path, html_file_path):
    # Create a Document instance
    document = Document()
    # Load a Word document
    document.LoadFromFile(docx_file_path)
    # Export document style to head in HTML
    document.HtmlExportOptions.IsExportDocumentStyles = True
    # Set the type of CSS style sheet as internal
    document.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal
    # Embed images in HTLM code
    document.HtmlExportOptions.ImageEmbedded = True
    # Export form fields as text
    document.HtmlExportOptions.IsTextInputFormFieldAsText = True
    # Save the document as an HTML file
    document.SaveToFile(html_file_path, FileFormat.Html)
    # Dispose resources
    document.Dispose()


if __name__ == "__main__":
    # docx_fpath_test = ".\\docs\\fake_hedge_fund_letter.docx"
    docx_fpath_test = ".\\docs\\Time_Series_Model_Development_Report_old.docx"
    # extract_text_and_tables_from_docx(docx_fpath_test)
    load_docx_unstructured(docx_fpath_test)
