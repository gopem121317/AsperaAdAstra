"""Streamlit app for Gen AI hackathon, team AAA"""

import functools
import hashlib
import os
from collections import defaultdict
from pathlib import Path
from typing import Any

import streamlit as st
from langchain_community.llms import OpenAI
from langchain_community.document_loaders import (DirectoryLoader,
                                                  Docx2txtLoader,
                                                  UnstructuredCSVLoader,
                                                  UnstructuredExcelLoader,
                                                  UnstructuredImageLoader,
                                                  UnstructuredMarkdownLoader,
                                                  UnstructuredPDFLoader,
                                                  UnstructuredWordDocumentLoader,
                                                  UnstructuredHTMLLoader)

import frontend.tmp
from document_updater import DocumentUpdater as doc_updater
from docx_parser import load_docx_unstructured

_TMP_DIR = frontend.tmp.__file__

_SHA1_BUF_SIZE = 65536

_UNSTRUCTURED_LOADERS = {  # TODO: include additional extensions and loaders as needed
    '.csv': UnstructuredCSVLoader,
    '.docx': functools.partial(UnstructuredWordDocumentLoader, mode='elements'),
    '.jpeg': UnstructuredImageLoader,
    '.html': functools.partial(UnstructuredHTMLLoader, mode='elements'),
    '.md': functools.partial(UnstructuredMarkdownLoader, mode='elements'),
    '.xlsx': UnstructuredExcelLoader,  # does this need a sheet name?
}


def get_file_hash(fp):
    sha1 = hashlib.sha1()
    with open(fp, 'rb') as f:
        while True:
            data = f.read(_SHA1_BUF_SIZE)
            if not data:
                break
            sha1.update(data)
    return sha1


def copy_uploaded_files_to_tmp_dir(files):
    file_paths = []
    for uploaded_file in files:
        uploaded_bytes = uploaded_file.getvalue()
        fp = os.path.join(_TMP_DIR, uploaded_file.name)
        file_paths.append(Path(fp).as_posix())
        if os.path.exists(fp):
            sha1_existing = get_file_hash(fp)
            sha1_incoming = hashlib.sha1(uploaded_bytes)
            if sha1_incoming != sha1_existing:
                os.remove(fp)
            else:
                continue
        with open(fp, 'rb') as f:
            f.write(uploaded_bytes)
    return file_paths


def group_uploaded_files_by_ext(files):
    d = defaultdict(list)
    for f in files:
        ext = os.path.splitext(f)[1]
        d[ext].append(f)
    return dict(d)


def process_uploaded_files(files) -> str | None:
    copied_files = copy_uploaded_files_to_tmp_dir(files)
    file_paths_by_ext = group_uploaded_files_by_ext(copied_files)

    # check for a docx file
    if 'docx' not in file_paths_by_ext:
        st.write('Must provide a .docx file')
        return

    # get file path of document that serves as template for updating
    # TODO: if more than one docx file, show selectbox, ask user to select which doc is target?
    target_file_path = file_paths_by_ext['.docx'][0]
    target_file_name = os.path.split(target_file_path)[1]

    # load unstructured data
    docs = []
    for ext, file_paths in file_paths_by_ext.items():
        loader = _UNSTRUCTURED_LOADERS.get(ext)
        if not loader:
            st.write(f'{ext} not supported')
            return
        for fp in file_paths:
            data = loader.load(fp)
            data.metadata["source_document"] = target_file_name
            docs += data

    # TODO: does everything need to go to doc store?

    # generate summary
    summary = doc_updater.summarize_document(docs)

    # upload to vector store
    doc_updater.add_documents_to_vector_store(docs)

    return summary


def generate_llm_response(input_text):
    llm = OpenAI(openai_api_key=st.secrets['openai_api_key'])
    st.info(llm(input_text))


st.title('📝 Doctor')

openai_api_key = st.secrets.get('openai_api_key')
if not openai_api_key:
    st.info('OpenAI API key not found')
    st.stop()

tab1, tab2, tab3 = st.tabs(['Generate', 'Refine', 'Q&A'])

with tab1:
    st.header('Generate updated document')
    uploaded_files = st.file_uploader(
        'Upload a document or data artifact',
        type=('doc', 'docx', 'csv', 'xls', 'xlsx'),
        accept_multiple_files=True,
    )
    if uploaded_files:
        st.divider()
        with st.spinner('Processing uploaded files...'):
            summary = process_uploaded_files(uploaded_files)
        st.caption('Summary')
        st.write_stream(summary)
        if st.button('Generate updated document', type='primary'):
            with st.spinner('Generating document...'):
                updated_doc, changes = None  # TODO: update doc
            if changes:
                with st.expander('Show changes'):
                    col1, col2 = st.columns(2)
                    for change in changes:
                        pass
                    st.write(changes)  # TODO: beautify, maybe in a table
            if updated_doc:
                st.download_button('Download updated document', updated_doc)
            else:
                st.write('Something went wrong')

with tab2:
    st.header('Refine document')

with tab3:
    st.header('Q&A')
    with st.form('chat_form'):
        text = st.text_area('Ask me something about the generated document:', placeholder='Can you give me a short summary?')
        submitted = st.form_submit_button("Submit", disabled=not uploaded_files)
        if submitted:
            generate_llm_response(text)
