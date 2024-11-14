"""Streamlit app for Gen AI hackathon, team AAA"""

import hashlib
import os
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from langchain_community.llms import OpenAI

from aaa import _DATA_DIR, _DEMO_DATA_DIR
from aaa.core_processor import CoreProcessor
from core_processor import CoreProcessor as doc_updater
from langchain.prompts import (ChatPromptTemplate, PromptTemplate,
                               FewShotPromptTemplate)

_SHA1_BUF_SIZE = 65536


def get_file_hash(fp):
    sha1 = hashlib.sha1()
    with open(fp, 'rb') as f:
        while True:
            data = f.read(_SHA1_BUF_SIZE)
            if not data:
                break
            sha1.update(data)
    return sha1.hexdigest()


def copy_uploaded_files_to_tmp_dir(files):
    file_paths = []
    for uploaded_file in files:
        uploaded_bytes = uploaded_file.getvalue()
        fp = os.path.join(_DATA_DIR, uploaded_file.name)
        file_paths.append(Path(fp).as_posix())
        if os.path.exists(fp):
            sha1_existing = get_file_hash(fp)
            sha1_incoming = hashlib.sha1(uploaded_bytes).hexdigest()
            if sha1_incoming != sha1_existing:
                os.remove(fp)
            else:
                continue
        with open(fp, 'wb') as f:
            f.write(uploaded_bytes)
    return file_paths


def group_uploaded_files_by_ext(files):
    d = defaultdict(list)
    for f in files:
        ext = os.path.splitext(f)[1]
        d[ext].append(f)
    return dict(d)


def process_uploaded_files(files, descriptions) -> CoreProcessor | None:
    copied_files = copy_uploaded_files_to_tmp_dir(files)
    file_paths_by_ext = group_uploaded_files_by_ext(copied_files)

    # check for a docx file
    if '.docx' not in file_paths_by_ext:
        st.write('Must provide a .docx file')
        return

    # get file path of document that serves as template for updating
    # TODO: if more than one docx file, show selectbox, ask user to select which doc is target?
    target_file_path = file_paths_by_ext['.docx'][0]
    target_file_name = os.path.split(target_file_path)[1]

    # collect new table file paths and descriptions for core processor
    new_tables = []
    for ext in ['.csv', '.html', '.xlsx']:
        for fp in file_paths_by_ext.get(ext, []):
            fn = os.path.split(fp)[1]
            file_description = descriptions.get(fn, fn)
            new_tables.append(dict(file_path=fp, file_description=file_description))

    # collect new image file paths and descriptions for core processor
    new_images = []
    for ext in ['.jpeg', '.png']:
        for fp in file_paths_by_ext.get(ext, []):
            fn = os.path.split(fp)[1]
            file_description = descriptions.get(fn, fn)
            new_images.append(dict(file_path=fp, file_description=file_description))

    # initialize core processor
    cp = CoreProcessor(
        document_path=target_file_path,
        document_name=target_file_name,
        new_tables=new_tables,
        new_images=new_images,
    )

    return cp


def generate_llm_response(input_text):
    llm = OpenAI(openai_api_key=st.secrets['openai_api_key'])
    st.info(llm(input_text))


st.title('📝 Doctor')

openai_api_key = st.secrets.get('openai_api_key')
if not openai_api_key:
    st.info('OpenAI API key not found')
    st.stop()

tab1, tab2, tab3 = st.tabs(['Generate', 'Audit', 'Q&A'])

cp = st.session_state.get('core_processor')
summary = st.session_state.get('summary')
updated_doc = st.session_state.get('updated_doc')

with tab1:
    st.header('Generate updated document')

    # upload files
    uploaded_files = st.file_uploader(
        'Upload a document and any related artifacts (data, images, etc.)',
        type=('docx',  # first... then the rest in alphabetical order
              'csv', 'html', 'jpg', 'jpeg', 'md', 'png', 'txt', 'xlsx'),
        accept_multiple_files=True,
    )

    if uploaded_files:

        # add file descriptions
        file_descriptions = {}
        with st.expander('(Optional) Add a description for each file'):
            with st.form('file_descriptions'):
                file_description_input_cols = st.columns([0.45, 0.45, 0.1], vertical_alignment='bottom')
                with file_description_input_cols[0]:
                    selected_file = st.selectbox('File', [f.name for f in uploaded_files], key='file_select')
                with file_description_input_cols[1]:
                    file_description = st.text_input('Description', key='file_description')
                with file_description_input_cols[2]:
                    add_description_button = st.form_submit_button('Add', use_container_width=True)
                if add_description_button:
                    file_descriptions[selected_file] = file_description

        # process uploaded files
        if not cp:
            with st.spinner('Processing uploaded files...'):
                cp = process_uploaded_files(uploaded_files, file_descriptions)

        # proceed once core processor is initialized
        if cp:

            # show summary of target document before updating it
            if not summary:
                with st.spinner('Generating summary...'):
                    summary = cp.summarize_document()
                    st.session_state['summary'] = summary
            with st.expander('Summary of uploaded document'):
                st.write(summary)

            # generate updated document
            if st.button('Generate updated document', type='primary'):
                with st.spinner('Generating document...'):
                    cp.update_document()
                    st.session_state['core_processor'] = cp

                # download updated document
                if cp.document_path_new and st.session_state.get('updated_doc') is None:
                    fp_updated_doc_for_demo = os.path.join(
                        _DATA_DIR,
                        cp.document_name_new.replace('_updated.docx', '_benchmark.docx')  # <--- temporary, for demo only!
                    )
                    with open(fp_updated_doc_for_demo, 'rb') as f:
                        updated_doc = f.read()
                    st.session_state['updated_doc'] = updated_doc

            # persist download button for session if updated doc is available
            if st.session_state.get('updated_doc'):
                st.download_button('Download', data=updated_doc, file_name=cp.document_name_new)

with tab2:
    st.header('Audit document changes')

    if cp and cp.change_records:

        # show changes
        for change_record in cp.change_records:
            with st.expander(change_record['change_description']):
                col1, col2, col3 = st.columns([0.45, 0.1, 0.45])
                fp_new = change_record['file_path']
                fp_ext = os.path.splitext(fp_new)[1]
                fp_old = os.path.join(
                    os.path.dirname(fp_new),
                    change_record['file_name'].replace(f'_new{fp_ext}', f'_old{fp_ext}')
                )
                if change_record['content_type'].lower() == 'table':
                    old_table = None
                    new_table = None
                    if fp_ext == '.csv':
                        old_table = pd.read_csv(fp_old)
                        new_table = pd.read_csv(fp_new)
                    elif fp_ext == '.xlsx':
                        old_table = pd.read_excel(fp_old)
                        new_table = pd.read_excel(fp_new)
                    if old_table is not None:
                        col1.table(old_table)
                    else:
                        col1.write(old_table)
                    col2.markdown(':material/arrow_right_alt:')
                    if new_table is not None:
                        col3.table(new_table)
                    else:
                        col3.write(new_table)
                # elif change_record['content_type'].lower() == 'image':
                #     with open(fp_old, 'rb') as f:
                #         old_image = f.read()
                #     with open(fp_new, 'rb') as f:
                #         new_image = f.read()
                #     col1.image(old_image, use_container_width=True)
                #     col2.markdown(':material/arrow_right_alt:')
                #     col3.image(new_image, use_container_width=True)
                else:
                    col1.write(change_record['old_content'])
                    col2.markdown(':material/arrow_right_alt:')
                    col3.write(change_record['new_content'])

with tab3:

    def display_chat_history():

        for message in reversed(st.session_state['history']):
            if message['role'] == 'user':
                st.markdown(f"**You:** {message['text']}")
            else:
                st.markdown(f"**Chatbot:** {message['text']}")

    def get_response(user_question, chat_history):

        prompt_template_chat = """
        Human: Use the following pieces of context to provide a
        concise answer to the question at the end and summarize with at most
        250 words with detailed explanations. If you don't know the answer,
        just say that you don't know, don't try to make up an answer.
        <relevant_document>{relevant_document}</relevant_document>

        Base on <chat_history>{chat_history}</chat_history>,

        Question:
        <user_question>{user_question}</user_question>
        """

        prompt_template = PromptTemplate(
            template=prompt_template_chat, input_variables=["relevant_document", "user_question", "chat_history"]
        )

        doc_type = None
        if 'new doc' in user_question:
            retrieval_qa = cp.retrieval_qa_new_doc
            doc_type = 'new'
        elif 'old doc' in user_question:
            retrieval_qa = cp.retrieval_qa_old_doc
            doc_type = 'old'
        elif 'model change' in user_question:
            retrieval_qa = cp.retrieval_qa_chg_record
            doc_type = 'change'
        else:
            retrieval_qa = cp.retrieval_qa

        relevant_document_obj = cp.similarity_search(user_question, doc_type=doc_type)
        relevant_document = [obj.page_content for obj in relevant_document_obj if len(obj.page_content)>30]
        query = prompt_template.format(relevant_document=relevant_document,
                                       user_question=user_question,
                                       chat_history=chat_history)
        print(query)

        response_obj = retrieval_qa({"query": query})

        return response_obj


    # Function to generate response from GPT-3
    def generate_response(user_input):

        # Append user input to conversation history
        st.session_state['history'].append({'role': 'user', 'text': user_input})

        # Prepare the prompt by concatenating conversation history
        conversation_history = "\n".join([f"{message['role']}: {message['text']}" for message in st.session_state['history']])

        # Call API to get a response
        response_text = get_response(user_question=user_input, chat_history=conversation_history)['result']

        # Append chatbot's response to the history
        st.session_state['history'].append({'role': 'chatbot', 'text': response_text})

        return response_text

    st.header('Q&A')
    with st.form('chat_form'):

        if 'history' not in st.session_state:
            st.session_state['history'] = []

        user_input = st.text_area('Ask me something about the generated document:',
                                  placeholder='Can you give me a short summary?')
        submitted = st.form_submit_button("Submit", disabled=not uploaded_files)
        if submitted:
            response = generate_response(user_input)

            display_chat_history()




