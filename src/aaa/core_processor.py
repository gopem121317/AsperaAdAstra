import json
import os
from os import getenv

import win32com.client
from pathlib import Path
from typing import List
from uuid import uuid4

import pandas as pd
from dotenv import load_dotenv
from langchain.chains import RetrievalQA
from langchain.chains.question_answering import load_qa_chain
from langchain.chains.summarize import load_summarize_chain
from langchain.prompts import (ChatPromptTemplate, PromptTemplate,
                               FewShotPromptTemplate)
from langchain_chroma import Chroma
from langchain_community.document_loaders import (UnstructuredCSVLoader,
                                                  UnstructuredExcelLoader,
                                                  UnstructuredImageLoader,
                                                  UnstructuredWordDocumentLoader)
from langchain_community.vectorstores.utils import filter_complex_metadata
from langchain_core.documents import Document
from langchain_openai import AzureChatOpenAI, ChatOpenAI, OpenAI, OpenAIEmbeddings, AzureOpenAIEmbeddings

from aaa.doc_handler import DocHandler, load_docx_unstructured

# Set up your Azure OpenAI credentials
os.environ["OPENAI_API_KEY"] = "fa5095aa40864544b164dbf337f7abe4"
os.environ["OPENAI_API_VERSION"] = "2024-06-01"
os.environ["AZURE_OPENAI_ENDPOINT"] = "https://genai-openai-asperaadastra.openai.azure.com"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")


def load_table_from_csv(csv_file_path: str):
    doc = UnstructuredCSVLoader(csv_file_path)
    data = doc.load()
    return data


def load_tables_from_excel(excel_file_path):
    tbl = UnstructuredExcelLoader(excel_file_path)
    data = tbl.load()
    return data


def load_image_simple(image_path: str, document_name: str = None):
    # img = UnstructuredImageLoader(image_path)  # Need Tesseract!!!
    # data = img.load()
    # return data
    data = os.path.basename(image_path)
    if document_name is None:
        document_name = data

    return [Document(
        page_content=data,
        metadata={"category": "Image",
                  "filename": document_name,
                  "source": image_path}
    )]


class CoreProcessor:

    def __init__(self, document_path, document_name: str, new_tables: List[dict] = None, new_images: List[dict] = None):
        self.document_name_old = document_name
        self.document_path_old = document_path
        self.elements_doc_old = None
        self.document_name_new = self.document_name_old.replace(".docx", "_updated.docx")
        self.document_path_new = Path(os.path.dirname(self.document_path_old), self.document_name_new).as_posix()
        self.new_tables = new_tables
        self.new_images = new_images
        self.doc_handler = DocHandler()
        self.change_records = []
        self.llm = AzureChatOpenAI(
                temperature=0,
                azure_deployment='gpt-4o',
                openai_api_key=OPENAI_API_KEY,
                azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
                api_version=os.getenv("OPENAI_API_VERSION"),
                model='gpt-4o'
            )

        self.embeddings = AzureOpenAIEmbeddings(
            model="text-embedding-3-large",
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
            api_key=OPENAI_API_KEY,
            openai_api_version=os.getenv("OPENAI_API_VERSION")
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
                    "filename": self.document_name_old,
                }
            }
        )
        self.retriever_new_doc = self.vector_store.as_retriever(
            search_type="mmr",
            search_kwargs={
                "k": 5,
                "filter": {
                    "filename": self.document_name_new,
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
        self.retrieval_qa_new_doc = RetrievalQA.from_chain_type(
            llm=self.llm,
            chain_type="stuff",
            retriever=self.retriever_new_doc,
            return_source_documents=True,
            verbose=True
        )

    def _clear_vector_store(self, clear_all=True):
        coll = self.vector_store.get()

        ids_to_delete = []
        for idx in range(len(coll["ids"])):
            _id = coll["ids"][idx]
            metadata = coll["metadatas"][idx]
            if clear_all:
                ids_to_delete.append(_id)
            else:
                if metadata.get("filename") == self.document_name_old:
                    ids_to_delete.append(_id)

        if len(ids_to_delete) > 0:
            self.vector_store.delete(ids=ids_to_delete)

    def add_documents_to_vector_store(self, documents: List[Document], unique_ids: list = None):
        if unique_ids is None:
            unique_ids = [str(uuid4()) for _ in range(len(documents))]

        for i, doc in enumerate(documents):
            doc.metadata["doc_unique_id"] = unique_ids[i]

        self.vector_store.add_documents(documents=filter_complex_metadata(documents), ids=unique_ids)

    def summarize_document(self, documents: List[Document] = None):
        chain = load_summarize_chain(
            self.llm,
            chain_type="stuff",
            # verbose=True
        )
        if documents is None:
            if self.elements_doc_old is None:
                self.process_original_document()
            documents = [x["document"] for x in self.elements_doc_old]
        return chain.run(documents)

    def _record_change(self, change_description: str,
                       new_content: str = None, old_content: str = None,
                       content_type: str = "Table", file_name: str = None,
                       file_path: str = None):
        data = {
            "change_description": change_description,
            "new_content": new_content,
            "old_content": old_content,
            "content_type": content_type,
            "file_name": file_name,
            "file_path": file_path
        }
        json_string = json.dumps(data)
        self.add_documents_to_vector_store(documents=[
            Document(
                page_content=json_string,
                metadata={"source": "change_record"}
            )
        ])
        self.change_records.append(data)

    def process_original_document(self):
        self.doc_handler.split_word_document(word_document_path=self.document_path_old)
        docs_list = self.doc_handler.export_docs()
        self.elements_doc_old = self.doc_handler.list_documents_parsed

        docs = [x["document"] for x in docs_list]
        ids = [str(y["id"]) for y in docs_list]

        self.add_documents_to_vector_store(documents=docs, unique_ids=ids)

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
        - The table has similar content as {image_string}

        Consider the context of the document to identify the correct image, especially the text before and after the image.
        Mention any captions or labels associated with the image.
        """
        return PromptTemplate(
            template=template,
            input_variables=["image_description", "image_string"]
        )

    def find_table_in_document(self, table_string=None, table_description=None):
        prompt = self.build_prompt_for_search_table_query()
        query = prompt.format(table_description=table_description,
                              table_string=table_string)
        response = self.retrieval_qa_old_doc({"query": query})

        for doc in response["source_documents"]:
            if doc.metadata["category"] == "Table":
                return doc

    def find_image_in_document(self, image_string, image_description):
        prompt = self.build_prompt_for_search_image_query()
        query = prompt.format(
            image_description=image_description,
            image_string=image_string
        )
        response = self.retrieval_qa_old_doc({"query": query})

        for doc in response["source_documents"]:
            if doc.metadata["category"] == "Image":
                return doc

    def similarity_search(self, query, k=8, doc_type=None):
        filter_val = None
        if doc_type == 'new':
            filter_val = {"filename": self.document_name_new}
        elif doc_type == 'old':
            filter_val = {"filename": self.document_name_old}
        elif doc_type == 'change':
            filter_val = {"source": "change_record"}
        return self.vector_store.similarity_search(query, k=k, filter=filter_val)

    def rag_qna(self, query):
        response = self.retrieval_qa({"query": query})
        return response

    def _process_table_update_task(self, new_table, table_description, file_name=None, file_path=None):
        # Find the table in the document
        tbl_found = self.find_table_in_document(new_table, table_description)
        page_num = tbl_found.metadata["page_number"]
        _chg_des = (f"Table on {table_description} from Page {page_num} has been updated with user provided input.")
        self._record_change(
            change_description=_chg_des,
            new_content=new_table,
            old_content=tbl_found.page_content,
            content_type="Table",
            file_name=file_name,
            file_path=file_path
        )
        return {
            "id": tbl_found.metadata["doc_unique_id"]
        }

    def _process_image_update_task(self, new_image, image_description, file_name=None, file_path=None):
        # Find the image in the document
        img_found = self.find_image_in_document(new_image, image_description)
        _chg_des = (f"Image on {image_description} has been updated with user provided input.")
        _mock_flag = True
        if _mock_flag:
            img_found = None
        if img_found is None:
            self._record_change(
                change_description=_chg_des,
                new_content=new_image,
                old_content=file_name,
                content_type="Image",
                file_name=file_name,
                file_path=file_path
            )
            return None
        else:
            self._record_change(
                change_description=_chg_des,
                new_content=new_image,
                old_content=img_found.page_content,
                content_type="Image",
                file_name=file_name,
                file_path=file_path
            )
            return {
                "id": img_found.metadata["doc_unique_id"]
            }

    def update_document(self):
        # Parse original document and upload to vector store
        if self.elements_doc_old is None:
            self.process_original_document()

        all2update = []
        for tbl_dict in self.new_tables:
            tbl = load_table_from_csv(tbl_dict["file_path"])
            tbl_fname = os.path.basename(tbl_dict["file_path"])
            tbl_df = pd.read_csv(tbl_dict["file_path"])
            tbl_str = tbl[0].page_content
            tbl_des = tbl_dict["file_description"]
            res_tbl = self._process_table_update_task(tbl_str, tbl_des, file_name=tbl_fname, file_path=tbl_dict["file_path"])
            res_tbl["table"] = tbl_df
            all2update.append(res_tbl)

        for img_dict in self.new_images:
            img = load_image_simple(img_dict["file_path"])
            img_fname = os.path.basename(img_dict["file_path"])
            img_str = img[0].page_content
            img_des = img_dict["file_description"]
            res_img = self._process_image_update_task(img_str, img_des, file_name=img_fname, file_path=img_dict["file_path"])
            if res_img is not None:
                res_img["image"] = img_dict["file_path"]
                all2update.append(res_img)

        self.doc_handler.update_document(all2update, output_doc_path=self.document_path_new)
        print(f"New document has been saved to {self.document_path_new}.")
        docs_new = load_docx_unstructured(self.document_path_new)
        self.add_documents_to_vector_store(docs_new)
        print("New document has been uploaded to the vector store.")


if __name__ == '__main__':
    test_input_dir = r"C:\Users\LXaXu\IdeaProjects\DocUpdaterAAA\raw_data\model_doc"
    # test_doc_name = "NVDA_Internal_Credit_Report.docx"
    test_doc_name = "Time_Series_Model_Development_Report_old.docx"
    test_doc_path = Path(test_input_dir, test_doc_name).as_posix()
    # test_docs = load_docx_unstructured(test_doc_path)
    # test_image_old_path = Path(test_input_dir, "scatter_bizA_old.jpg").as_posix()
    # test_image_old = load_image_simple(test_image_old_path, test_doc_name)
    # test_docs += test_image_old
    test_new_tables = [
        {
            "file_path": Path(test_input_dir, "fitted_vs_actual_biz_b.csv").as_posix(),
            "file_description": "Fitted vs Actual Values for biz B"
        }
    ]
    test_new_images =[
        {
            "file_path": Path(test_input_dir, "scatterplot_biz a_gdp.jpeg").as_posix(),
            "file_description": "Scatter plot for Business A"
        }
    ]
    doc_updater = CoreProcessor(
        document_path=test_doc_path,
        document_name=test_doc_name,
        new_tables=test_new_tables,
        new_images=test_new_images
    )
    # summary = doc_updater.summarize_document()
    # print(summary)

    doc_updater.update_document()
    print(doc_updater.change_records)

    # doc_updater.add_documents_to_vector_store(test_docs)
    #
    # test_csv_file = Path(test_input_dir, "Fitted_vs_Actual_Business_X.csv").as_posix()
    #
    # test_new_tbl = load_table_from_csv(test_csv_file)
    # found_tbl = doc_updater.find_table_in_document(
    #     table_string=test_new_tbl[0].page_content,
    #     table_description="Fitted vs Actual Values for biz B")
    # print(found_tbl)
    #
    # test_image_new_path = Path(test_input_dir, "Business X_scatter.jpeg").as_posix()
    # test_image_new = load_image_simple(test_image_new_path)
    # found_img = doc_updater.find_image_in_document(
    #     image_string=test_image_new[0].page_content,
    #     image_description="Scatter plot for Business A"
    # )
    # print(found_img)
