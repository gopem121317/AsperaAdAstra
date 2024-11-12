import json
import os
import win32com.client
from pathlib import Path
from typing import List
from uuid import uuid4

import bs4
import fitz
import mammoth
import pandas as pd
from dotenv import load_dotenv
from langchain.chains import RetrievalQA
from langchain.chains.question_answering import load_qa_chain
from langchain.chains.summarize import load_summarize_chain
from langchain.prompts import (ChatPromptTemplate, PromptTemplate,
                               FewShotPromptTemplate)
from langchain.retrievers.multi_vector import MultiVectorRetriever
from langchain.storage import InMemoryStore
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma
from langchain_community.document_loaders import (DirectoryLoader,
                                                  Docx2txtLoader,
                                                  UnstructuredCSVLoader,
                                                  UnstructuredExcelLoader,
                                                  UnstructuredImageLoader,
                                                  UnstructuredMarkdownLoader,
                                                  UnstructuredPDFLoader,
                                                  UnstructuredWordDocumentLoader,
                                                  UnstructuredHTMLLoader)
from langchain_community.vectorstores.utils import filter_complex_metadata
from langchain_core.documents import Document
from langchain_core.runnables import RunnablePassthrough
from langchain_openai import AzureChatOpenAI, ChatOpenAI, OpenAI, OpenAIEmbeddings
from numba.cuda.testing import test_data_dir

from docx_parser import load_docx_unstructured

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY",
                           "...")


def extract_tables_and_texts_from_pdf(pdf_file_path):
    """Extract tables from the PDF file."""
    doc = fitz.open(pdf_file_path)
    for page in doc:
        tabs = page.find_tables()
        if tabs.tables:
            print(tabs[0].extract())


def extract_tables_and_texts_from_markdown(md_file_path):
    doc = UnstructuredMarkdownLoader(md_file_path, mode="elements")
    data = doc.load()
    return data


def extract_tables_and_texts_from_html(html_file_path):
    doc = UnstructuredHTMLLoader(html_file_path, mode="elements")
    data = doc.load()
    return data


def load_table_from_excel_sheet(excel_file_path):
    tbl = UnstructuredExcelLoader(excel_file_path)
    data = tbl.load()
    return data


def resave_word_as_html(word_file_path):
    doc = win32com.client.GetObject(word_file_path)
    doc.SaveAs(FileName="ToHtml.html", FileFormat=8)
    doc.Close()


def load_csv_file(csv_file_path):
    doc = UnstructuredCSVLoader(csv_file_path)
    data = doc.load()
    return data


class DocumentUpdater:
    def __init__(self, document_name: str):
        """Initialize the PDF Table Updater with OpenAI API key."""
        self.doc_name_old = document_name
        self.doc_name_new = None
        self.doc_old = None
        self.doc_new = None
        self.llm = ChatOpenAI(
            # model="gpt-4",
            model="gpt-4o",
            temperature=0,
            openai_api_key=OPENAI_API_KEY
        )
        self.embeddings = OpenAIEmbeddings(
            # model="text-embedding-3-large",
            openai_api_key=OPENAI_API_KEY
        )
        self.vector_store = Chroma(
            collection_name="collection_doc_updater",
            embedding_function=self.embeddings,
            persist_directory="./chroma_langchain_db",
        )
        self._clear_vector_store()
        self.retriever = self.vector_store.as_retriever(
            # search_type="mmr",
            search_type="similarity",
            search_kwargs={"k": 5}
        )
        self.retriever_old_doc = self.vector_store.as_retriever(
            search_type="mmr",
            search_kwargs={
                "k": 5,
                "filter": {
                    "source_document": self.doc_name_old,
                }
            }
        )
        self.retriever_chg_record = self.vector_store.as_retriever(
            search_type="mmr",
            search_kwargs={
                "filter": {
                    "source": "change_record",
                }
            }
        )
        self.retrieval_qa = RetrievalQA.from_chain_type(
            llm=self.llm,
            chain_type="stuff",
            retriever=self.retriever,
            return_source_documents=True,
            verbose=True
        )
        self.retrieval_qa_old_doc = RetrievalQA.from_chain_type(
            llm=self.llm,
            chain_type="refine",
            # retriever=self.retriever,
            retriever=self.retriever_old_doc,
            return_source_documents=True,
            verbose=True
        )
        self.retrieval_qa_chg_record = RetrievalQA.from_chain_type(
            llm=self.llm,
            chain_type="stuff",
            retriever=self.retriever_chg_record,
            return_source_documents=True,
            verbose=True
        )
        self.change_records = []

    def _clear_vector_store(self):
        coll = self.vector_store.get()

        ids_to_delete = []
        for idx in range(len(coll["ids"])):
            _id = coll["ids"][idx]
            metadata = coll["metadatas"][idx]
            if metadata.get("source") == self.doc_name_old:
                ids_to_delete.append(_id)

        if len(ids_to_delete) > 0:
            self.vector_store.delete(ids=ids_to_delete)

    def _enrich_documents_lc(self, documents: List[Document]):
        for doc in documents:
            doc.metadata["source_document"] = self.doc_name_old

        return documents

    def _record_change(self, change_description: str,
                       new_content: str = None, old_content: str = None):
        data = {
            "change_description": change_description,
            "new_content": new_content,
            "old_content": old_content
        }
        json_string = json.dumps(data)
        self.change_records.append(
            Document(
                page_content=json_string,
                metadata={"source": "change_record"}
            )
        )

    def add_documents_to_vector_store(self, docs, unique_ids: list = None):
        if unique_ids is None:
            unique_ids = [str(uuid4()) for _ in range(len(docs))]
        self.vector_store.add_documents(documents=filter_complex_metadata(docs), ids=unique_ids)

    def summarize_document(self, docs):
        chain = load_summarize_chain(
            self.llm,
            chain_type="stuff",
            # verbose=True
        )
        return chain.run(docs)

    @staticmethod
    def build_prompt_for_search_table_query():
        template = """
        Please identify the table in the document based on the following criteria:
        
        - The table is about {table_description}
        - The table has similar content as {table_string}

        Consider the context of the document to identify the correct table, especially the text before and after the table.
        Mention any captions or labels associated with the table.
        """
        return PromptTemplate(
            template=template,
            input_variables=["table_description", "table_string"]
        )

    @staticmethod
    def build_prompt_for_search_image_query():
        template = """
        Please identify the image in the document based on the following criteria:

        - The image illustrates {image_description}
        - The image is located in the document near the following text: {context}
        
        Consider the context of the document to identify the correct image, especially the text before and after the image.
        Mention any captions or labels associated with the image.
        """
        return PromptTemplate(
            template=template,
            input_variables=["image_description"]
        )

    def find_table_in_document(self, table_string=None, table_description=None):
        prompt = self.build_prompt_for_search_table_query()
        query = prompt.format(table_description=table_description,
                              table_string=table_string)
        response = self.retrieval_qa_old_doc({"query": query})

        for doc in response["source_documents"]:
            if doc.metadata["category"] == "Table":
                return doc

    def find_image_in_document(self, image_description):
        prompt = self.build_prompt_for_search_image_query()
        query = prompt.format(image_description=image_description)
        response = self.retrieval_qa_old_doc({"query": query})

        for doc in response["source_documents"]:
            if doc.metadata["category"] == "Image":
                return doc

    @staticmethod
    def get_few_shot_prompt_for_table_illustration():
        # TODO
        examples = [
            {
                "question": "",
                "prompt": ""
            },
            {
                "question": "",
                "prompt": ""
            },
            {
                "question": "",
                "prompt": ""
            }
        ]
        example_prompt = PromptTemplate.from_template("Question: {question}\n{answer}")
        prompt = FewShotPromptTemplate(
            examples=examples,
            example_prompt=example_prompt,
            suffix="{input}",
            input_variables=["input"]
        )
        return prompt

    def generate_new_content_description(self, new_content: Document) -> Document:
        prompt = self.get_few_shot_prompt_for_table_illustration()
        query = prompt.invoke({"input": new_content}).to_string()
        des = self.llm.predict(query)
        return Document(
            page_content=des,
            metadata={"source_document": self.doc_name_new},
        )

    def similarity_search(self, query):
        return self.vector_store.similarity_search(query)

    def rag_qna(self, query):
        response = self.retrieval_qa({"query": query})
        return response

    def _process_table_update_task(self, new_table, table_description):
        # Find the table in the document
        resp = self.find_table_in_document(table_description)
        _chg_des = (f"Table on {table_description} has been updated with user provided input."
                    f"Old Table: \n"
                    f"New Table: {new_table}")
        self._record_change(_chg_des, old_content=None, new_content=new_table)

    def summarize_changes(self):
        return self.retrieval_qa

    def save_updated_doc(self, updated_tables, updated_images):
        # TODO: placeholder
        doc = win32com.client.GetObject(self.doc_old)
        for idx, table in enumerate(updated_tables):
            doc.Tables(idx).Range.Text = table
        doc.SaveAs(self.doc_name_new)
        doc.Close()


def prepare_test_case_data(data_dir):
    docs = load_docx_unstructured(Path(data_dir, "Time_Series_Model_Development_Report_old.docx").as_posix())
    img_loader = UnstructuredImageLoader(Path(data_dir, "bizA_revenue_ts_old.jpeg").as_posix())
    img_data = img_loader.load()
    docs += img_data
    for doc in docs:
        doc.metadata["source_document"] = "Time_Series_Model_Development_Report_old.docx"
    return docs


if __name__ == "__main__":
    test_input_dir = ".\\synthetic_data"
    documents = prepare_test_case_data(test_input_dir)
    doc_updater = DocumentUpdater(document_name="Time_Series_Model_Development_Report_old.docx")
    summary = doc_updater.summarize_document(documents)
    print(summary)
    doc_updater.add_documents_to_vector_store(documents)
    csv_doc = load_csv_file(".\\docs\\BizB_Actual_vs_Fitted_new.csv")
    found_tbl = doc_updater.find_table_in_document(
        table_string=csv_doc[0].page_content,
        table_description="Fitted vs Actual Values for biz B")
    print(found_tbl)
