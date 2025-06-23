# file_id and lastModifiedDateTime are now used to track changes.
# Only new/updated documents get reprocessed.
# Metadata is persisted in processed_files.json.

import os, requests, shutil, tempfile, json
import uuid
from io import BytesIO
from typing import Generator, Tuple, List
from dotenv import load_dotenv
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_core.documents import Document
from langchain_openai import OpenAIEmbeddings
from langchain_chroma import Chroma
from openai import OpenAI

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_SITE = os.getenv("SITE_URL_NEW")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
PROCESSED_META_FILE = "processed_files.json"

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

def fetch_files(token: str) -> Generator[Tuple[str, BytesIO, str, str, str], None, None]:
    headers = {"Authorization": f"Bearer {token}"}
    site_name = SHAREPOINT_SITE.split("/")[-1]

    site_res = requests.get(f"https://graph.microsoft.com/v1.0/sites/root:/sites/{site_name}", headers=headers)
    site_res.raise_for_status()
    site_id = site_res.json()["id"]

    drive_res = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
    drive_res.raise_for_status()
    drive_id = drive_res.json()["id"]

    def traverse_items(folder_id="root"):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children"
        res = requests.get(url, headers=headers)
        res.raise_for_status()
        for item in res.json().get("value", []):
            if item.get("folder"):
                yield from traverse_items(item["id"])
            elif item.get("file"):
                file_name = item["name"]
                file_id = item["id"]
                last_modified = item.get("lastModifiedDateTime")
                web_url = item.get("webUrl", f"https://sharepoint.com/{file_name}")
                content_res = requests.get(
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content", headers=headers)
                content_res.raise_for_status()
                yield file_name, BytesIO(content_res.content), web_url, file_id, last_modified

    yield from traverse_items()

def save_temp_file(file_stream, suffix):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(file_stream.read())
    tmp.flush()
    return tmp.name

def load_document(file_name, file_stream, url: str, file_id: str) -> List[Document]:
    ext = os.path.splitext(file_name)[1].lower()
    print(f"[📂 Loading] File: {file_name}, Extension: {ext}")

    from langchain_community.document_loaders import (
        PyPDFLoader, UnstructuredExcelLoader,
        UnstructuredWordDocumentLoader, UnstructuredPowerPointLoader,
        CSVLoader, TextLoader
    )

    try:
        if ext == ".pdf":
            path = save_temp_file(file_stream, ".pdf")
            docs = PyPDFLoader(path).load()
            print("DOCs:",docs)
        elif ext == ".docx":
            path = save_temp_file(file_stream, ".docx")
            docs = UnstructuredWordDocumentLoader(path).load()
        elif ext == ".pptx":
            path = save_temp_file(file_stream, ".pptx")
            docs = UnstructuredPowerPointLoader(path).load()
        elif ext in [".xls", ".xlsx"]:
            path = save_temp_file(file_stream, ".xlsx")
            docs = UnstructuredExcelLoader(path).load()
        elif ext == ".csv":
            path = save_temp_file(file_stream, ".csv")
            docs = CSVLoader(path).load()
        elif ext == ".txt":
            path = save_temp_file(file_stream, ".txt")
            docs = TextLoader(path).load()
        elif ext in [".mp3", ".mp4"]:
            temp_path = save_temp_file(file_stream, ext)
            client = OpenAI(api_key=OPENAI_API_KEY)
            with open(temp_path, "rb") as f:
                transcription = client.audio.transcriptions.create(model="gpt-4o-mini-transcribe", file=f)
            docs = [Document(page_content=transcription.text)]
        else:
            print(f"[ℹ️ Unsupported file type] Skipping: {file_name}")
            return []

        for d in docs:
            d.metadata["source"] = url
            d.metadata["file_id"] = file_id  # ✅ Add file ID to metadata

        print(f"[✅ Loaded] {file_name} -> {len(docs)} document(s)")
        return docs

    except Exception as e:
        print(f"[❌ ERROR loading {file_name}]: {e}")
        return []

def load_processed_metadata():
    if os.path.exists(PROCESSED_META_FILE):
        with open(PROCESSED_META_FILE, "r") as f:
            return json.load(f)
    return {}

def save_processed_metadata(metadata):
    with open(PROCESSED_META_FILE, "w") as f:
        json.dump(metadata, f, indent=2)

def chunk_documents(docs):
    splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=200)
    return splitter.split_documents(docs)

def embed_and_store(chunks):
    persist_path = "chroma_db"
    print("[💡 Initializing Embedding Model]")
    embedding_model = OpenAIEmbeddings(openai_api_key=OPENAI_API_KEY)

    print("[📂 Loading or Creating Chroma Vector Store]")
    vectorstore = Chroma(persist_directory=persist_path, embedding_function=embedding_model)

    total_deleted = 0
    file_ids_to_update = {chunk.metadata.get("file_id") for chunk in chunks if "file_id" in chunk.metadata}

    for file_id in file_ids_to_update:
        print(f"[🧹 Cleaning] Searching for old chunks with file_id: {file_id}")
        matching_docs = vectorstore.similarity_search("", k=1000, filter={"file_id": file_id})
        ids_to_delete = [doc.metadata.get("uuid") for doc in matching_docs if doc.metadata.get("file_id") == file_id and doc.metadata.get("uuid")]

        if ids_to_delete:
            print(f"[🗑️ Deleting] {len(ids_to_delete)} chunks for file_id: {file_id}")
            vectorstore.delete(ids=ids_to_delete)
            total_deleted += len(ids_to_delete)
        else:
            print(f"[ℹ️ No existing chunks found for file_id: {file_id}]")

    # 🧠 Assign new UUIDs to all new chunks
    for chunk in chunks:
        chunk.metadata["uuid"] = str(uuid.uuid4())

    print(f"\n[➕ Adding Chunks to Vector Store] Total new chunks: {len(chunks)}")
    vectorstore.add_documents(chunks)

    print("[✅ Vector Store Updated and Persisted]")
    print(f"\n[📉 Summary] Total old chunks deleted due to file updates: {total_deleted}")

# def process_new_or_updated_file(file_id: str):
#     print(f"[🔄 Webhook Trigger] Processing file_id: {file_id}")
#     token = get_access_token()

#     for file_name, file_stream, file_url, fid, last_modified in fetch_files(token):
#         if fid == file_id:
#             print(f"[📥 Matched File] {file_name} (ID: {fid})")
#             docs = load_document(file_name, file_stream, file_url, fid)
#             if not docs:
#                 print(f"[⚠️ Empty or failed document load for]: {file_name}")
#                 return
#             chunks = chunk_documents(docs)
#             embed_and_store(chunks)

#             # Update processed metadata
#             metadata = load_processed_metadata()
#             metadata[file_id] = last_modified
#             save_processed_metadata(metadata)
#             print(f"[✅ Done] {file_name} re-processed from webhook event.")
#             return

#     print(f"[❌ File ID {file_id} not found in SharePoint fetch]")

def main():
    token = get_access_token()
    print("[🔑 Access Token Retrieved]")

    all_chunks = []
    previous_metadata = load_processed_metadata()
    current_metadata = {}

    for file_name, file_stream, file_url, file_id, last_modified in fetch_files(token):
        current_metadata[file_id] = last_modified
        if previous_metadata.get(file_id) == last_modified:
            print(f"[⏩ Skipping unchanged file] {file_name}")
            continue

        print(f"\n[📥 Processing File] {file_name}")
        docs = load_document(file_name, file_stream, file_url, file_id)
        if not docs:
            print(f"[⚠️ No Documents Returned] Skipping {file_name}")
            continue
        chunks = chunk_documents(docs)
        print(f"[✂️ Chunked] {file_name}: {len(chunks)} chunks")
        if chunks:
            print(f"[🧾 First Chunk Preview] {chunks[0].page_content[:300]}...")
            print(f"[📎 First Metadata] {chunks[0].metadata}")
        all_chunks.extend(chunks)

    if all_chunks:
        print(f"\n[💾 Storing Embeddings] Total chunks: {len(all_chunks)}")
        embed_and_store(all_chunks)
        print("[✅ Done] Embeddings stored in ChromaDB.")
    else:
        print("[⚠️ No chunks created. Nothing stored.]")

    # ✅ Detect and delete chunks from files removed in SharePoint
    deleted_file_ids = set(previous_metadata.keys()) - set(current_metadata.keys())
    if deleted_file_ids:
        print(f"\n[🗑️ File Deletion Detected] {len(deleted_file_ids)} file(s) removed from SharePoint")

        embedding_model = OpenAIEmbeddings(model="text-embedding-3-small", openai_api_key=OPENAI_API_KEY)
        vectorstore = Chroma(persist_directory="chroma_db", embedding_function=embedding_model)

        for file_id in deleted_file_ids:
            print(f"[🔍] Removing chunks for deleted file_id: {file_id}")
            result = vectorstore._collection.get(where={"file_id": file_id})
            ids_to_delete = result["ids"] if result and "ids" in result else []

            if ids_to_delete:
                vectorstore.delete(ids=ids_to_delete)
                print(f"[✅ Deleted] {len(ids_to_delete)} chunks removed for file_id: {file_id}")
            else:
                print(f"[ℹ️ Skipped] No chunks found for file_id: {file_id}")
    else:
        print("\n[✔️ No file deletions detected]")

    # ✅ Save updated metadata
    save_processed_metadata(current_metadata)

    # ✅ Final chunk count summary
    print("\n[📊 Verifying Total Chunks in ChromaDB]")
    embedding_model = OpenAIEmbeddings(model="text-embedding-3-small", openai_api_key=OPENAI_API_KEY)
    vectorstore = Chroma(persist_directory="chroma_db", embedding_function=embedding_model)
    total_chunks = vectorstore._collection.count()
    print(f"[📦 Total Chunks in ChromaDB]: {total_chunks}")


if __name__ == "__main__":
    main()
