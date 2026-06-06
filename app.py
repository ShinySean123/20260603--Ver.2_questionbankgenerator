import streamlit as st
import pandas as pd
import re
import math
import io
import json
import time      
import random    
import urllib.request
import urllib.parse
import requests  
import base64
import fitz  # PyMuPDF
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

# Word 處理相關
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 網頁基礎配置
st.set_page_config(page_title="AI 醫學共筆智慧題庫系統", page_icon="🧠", layout="centered")

st.title("🧠 AI 醫學共筆全功能智慧題庫系統")
st.markdown("內建【極速全講義圖文出題】與【現成題目自動補詳解】雙模引擎，全面優化 RPD 降耗。")

# ==================== 1. 🔑 API 金鑰設定面板 ====================
env_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets: env_key = st.secrets["GEMINI_API_KEY"]
except Exception: pass

with st.sidebar:
    st.header("🔑 系統安全憑證")
    user_live_key = st.text_input("💡 請輸入 API Key：", value=env_key if env_key else "", type="password")
    api_key = user_live_key.strip() if user_live_key else env_key

if not api_key:
    st.warning("⚠️ 請先在左側邊欄貼入您在 Google AI Studio 申請的 `AIzaSy` 金鑰。")
    st.stop()

# ==================== 🌟 核心底層 HTTP 直連函數 ====================
def generate_content_via_http_with_retry(contents_list, api_key, max_retries=4):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    
    parts = []
    for item in contents_list:
        if isinstance(item, dict) and item.get("mime_type") == "image/jpeg":
            b64_data = base64.b64encode(item["data"]).decode('utf-8')
            parts.append({"inlineData": {"mimeType": "image/jpeg", "data": b64_data}})
        else:
            parts.append({"text": str(item)})

    payload = {"contents": [{"parts": parts}]}
    headers = {"Content-Type": "application/json"}
    
    for attempt in range(max_retries):
        resp = requests.post(url, headers=headers, json=payload)
        if resp.status_code == 200:
            return resp.json()['candidates'][0]['content']['parts'][0]['text']
        elif resp.status_code == 503:
            if attempt < max_retries - 1:
                wait_time = (2 ** attempt) + random.uniform(0, 1)
                st.warning(f"⏳ 門口短暫大塞車 (503) - 正在自動重試第 {attempt+1}/{max_retries} 次...")
                time.sleep(wait_time)
                continue
            else:
                raise Exception(f"Google 門口持續 503 塞車，已重試 {max_retries} 次。")
        else:
            raise Exception(f"Google 門口回應錯誤 ({resp.status_code}): {resp.text}")

# ==================== 🗂️ 講義出題模式專用：GitHub 自動資料夾掃描 ====================
GITHUB_USER = "ShinySean123"
GITHUB_REPO = "20260603--Ver.2_questionbankgenerator"
GITHUB_FOLDER_HIST = "history_db"          
GITHUB_FOLDER_PDF = "current_materials"    

encoded_user = urllib.parse.quote(GITHUB_USER)
encoded_repo = urllib.parse.quote(GITHUB_REPO)

@st.cache_data(ttl=60)
def scan_github_folders():
    all_excel = []
    all_pdf = []
    
    # 掃描歷史題庫
    try:
        url_hist = f"https://api.github.com/repos/{encoded_user}/{encoded_repo}/contents/{urllib.parse.quote(GITHUB_FOLDER_HIST)}"
        req = urllib.request.Request(url_hist, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as resp:
            for item in json.loads(resp.read().decode()):
                if item['type'] == 'file' and item['name'].endswith('.xlsx'): all_excel.append(item['name'])
    except Exception:
        try:
            html_url = f"https://github.com/{encoded_user}/{encoded_repo}/tree/main/{urllib.parse.quote(GITHUB_FOLDER_HIST)}"
            req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: html_text = resp.read().decode('utf-8')
            all_excel = list(set(re.findall(r'title="([^"]+\.xlsx)"', html_text)))
        except: pass

    # 掃描雲端講義
    try:
        url_pdf = f"https://api.github.com/repos/{encoded_user}/{encoded_repo}/contents/{urllib.parse.quote(GITHUB_FOLDER_PDF)}"
        req = urllib.request.Request(url_pdf, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as resp:
            for item in json.loads(resp.read().decode()):
                if item['type'] == 'file' and item['name'].lower().endswith('.pdf'): all_pdf.append(item['name'])
    except Exception:
        try:
            html_url = f"https://github.com/{encoded_user}/{encoded_repo}/tree/main/{urllib.parse.quote(GITHUB_FOLDER_PDF)}"
            req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: html_text = resp.read().decode('utf-8')
            found = re.findall(r'title="([^"]+\.[pP][dD][fF])"', html_text)
            all_pdf = list(set([urllib.parse.unquote(p) for p in found]))
        except: pass
        
    return all_excel, all_pdf

all_excel_files, cloud_pdf_files = scan_github_folders()

def fetch_excel_titles(file_name):
    encoded_name = urllib.parse.quote(file_name)
    raw_url = f"https://raw.githubusercontent.com/{encoded_user}/{encoded_repo}/main/{GITHUB_FOLDER_HIST}/{encoded_name}"
    try:
        req = urllib.request.Request(raw_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as resp: df = pd.read_excel(io.BytesIO(resp.read()))
        q_col = next((c for c in df.columns if any(k in str(c) for k in ["題目", "Question", "內容"])), None)
        if q_col: return df[q_col].dropna().astype(str).tolist()
    except: pass
    return []

def fetch_cloud_pdf_bytes(file_name):
    encoded_name = urllib.parse.quote(file_name)
    raw_url = f"https://raw.githubusercontent.com/{encoded_user}/{encoded_repo}/main/{GITHUB_FOLDER_PDF}/{encoded_name}"
    try:
        req = urllib.request.Request(raw_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as resp: return resp.read()
    except:
        try:
            raw_url_alt = f"https://github.com/{encoded_user}/{encoded_repo}/raw/main/{GITHUB_FOLDER_PDF}/{encoded_name}"
            req = urllib.request.Request(raw_url_alt, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: return resp.read()
        except: return None

# ==================== 🌟 2. 核心大選單：功能模式選擇 ====================
st.markdown("---")
system_mode = st.radio(
    "🎯 **請選擇您本次想要使用的智慧功能：**",
    ["📚 智慧全自動融合出題系統 (看講義/看圖片出題)", "📝 題目原封不動配詳解系統 (讀 Excel 題目生詳解)"],
    index=0,
    horizontal=True
)
st.markdown("---")

# ==================== 🛠️ 模式 A：智慧全自動融合出題系統 ====================
if "智慧全自動融合出題系統" in system_mode:
    
    file_options = ["❌ 不使用歷史資料（全新出題）"]
    if all_excel_files:
        file_options.append("💥 比對資料夾內【所有檔案】（全面防重複）")
        for f in all_excel_files: file_options.append(f)
        
    with st.sidebar:
        st.header("⚙️ 雲端講義與去重設定")
        selected_mode = st.selectbox("請選擇歷史題庫防重複模式：", file_options)
        st.markdown("---")
        st.header("📚 雲端講義書櫃")
        if cloud_pdf_files:
            st.success(f"🟢 偵測到雲端有 {len(cloud_pdf_files)} 份 PDF 講義")
            selected_cloud_pdfs = st.multiselect("請勾選本次想連動出題的雲端講義：", cloud_pdf_files)
        else:
            st.info("ℹ️ 目前雲端未偵測到 PDF。")
            selected_cloud_pdfs = []

    history_titles = []
    if "【所有檔案】" in selected_mode:
        for f in all_excel_files: history_titles.extend(fetch_excel_titles(f))
    elif selected_mode != "❌ 不使用歷史資料（全新出題）":
        history_titles = fetch_excel_titles(selected_mode)

    st.subheader("📂 Step 1: 選取或上傳課程講義 PDF")
    uploaded_pdfs = st.file_uploader("從本機電腦上傳新講義 PDF (可多選)", type=["pdf"], accept_multiple_files=True)
    total_pdf_count = len(uploaded_pdfs) + len(selected_cloud_pdfs)

    if total_pdf_count > 0:
        st.markdown(f"📊 **目前已鎖定講義總數：{total_pdf_count} 份**")
        
        st.subheader("📝 Step 2: 設定出題參數")
        col1, col2, col3 = st.columns(3)
        with col1: page_range = st.text_input("想根據哪幾頁出題？", "整份")
        with col2: topic_name = st.text_input("章節/主題名稱", "醫學綜合領域測驗")
        with col3: num_questions = st.number_input("預計生成題數", min_value=1, max_value=50, value=10)

        st.markdown("---")
        col_num, _ = st.columns([1, 2])
        with col_num: start_q_num = st.number_input("🔢 設定「起始題號」", min_value=1, max_value=999, value=1, step=1)
        st.markdown("---")

        col4, col5 = st.columns(2)
        with col4: exam_title_input = st.text_input("Word 考卷大標題", f"{topic_name}_綜合測驗")
        with col5: excel_filename_input = st.text_input("Excel 輸出檔名", f"{topic_name}_綜合題庫")

        exam_title = str(exam_title_input) if exam_title_input else "測驗題庫"
        excel_filename = str(excel_filename_input) if excel_filename_input else "精修題庫"

        if st.button("⚡ 開始全自動雙模融合出題 ⚡", use_container_width=True):
            try:
                contents_payload = []
                def process_pdf_to_compressed_images(pdf_bytes, pdf_name):
                    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    for i in range(len(doc)):
                        page = doc.load_page(i)
                        pix = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
                        img_data = pix.tobytes("jpeg")
                        contents_payload.append(f"=== 【{pdf_name}】第 {i+1} 頁 ===")
                        contents_payload.append({"mime_type": "image/jpeg", "data": img_data})

                with st.spinner("📷 正在為整包講義進行高比例視覺轉碼與大瘦身..."):
                    for pdf_file in uploaded_pdfs: process_pdf_to_compressed_images(pdf_file.read(), pdf_file.name)
                    for cloud_pdf_name in selected_cloud_pdfs:
                        c_bytes = fetch_cloud_pdf_bytes(cloud_pdf_name)
                        if c_bytes: process_pdf_to_compressed_images(c_bytes, cloud_pdf_name)

                with st.spinner("🧠 輕量包封裝完成！AI 正在一次性極速研讀整份圖文講義... 請稍候"):
                    range_instruction = f"精準鎖定這些影像中的【{page_range}】" if "整份" not in page_range and "全部" not in page_range else "「通盤掃描並融合這幾份講義影像」的完整內容"
                    history_block = "⚠️ 絕對禁止重複、改寫或高度雷同以下這些已經出過的舊題目：\n" + "\n".join([f"- {t}" for t in history_titles]) if history_titles else ""

                    prompt = f"""
                    你現在是一位資深的醫學與生物科學教授。請根據我提供的這份完整講義影像，{range_instruction}，並圍繞核心主題【{topic_name}】一次性出齊 {num_questions} 題五選一的單選題。
                    請特別注意發揮你的視覺辨識能力，若講義中有重要心電圖、病理切片或解剖流程圖，務必將其觀念轉化為考題！
                    {history_block}
                    格式必須為 JSON 列表(Array)，每個物件的 Key 嚴格為："題目內容", "選項A", "選項B", "選項C", "選項D", "選項E", "正確答案", "針對各選項之詳解", "出處"
                    詳解與出處必須用繁體中文詳細解釋，題目與選項及正確答案必須是全英文。請直接輸出 JSON 陣列，不要包含 ```json 等 Markdown 包裝。
                    """
                    contents_payload.append(prompt)
                    clean_response = generate_content_via_http_with_retry(contents_payload, api_key)
                    
                    if clean_response.startswith("
http://googleusercontent.com/immersive_entry_chip/0
http://googleusercontent.com/immersive_entry_chip/1
http://googleusercontent.com/immersive_entry_chip/2
http://googleusercontent.com/immersive_entry_chip/3

---

### 🎨 你會看到的新改變：
1. **網頁首頁上方橫式按鈕選單**：使用者一進首頁，就能在「**📚 智慧全自動融合出題系統**」與「**📝 題目原封不動配詳解系統**」之間自由點選切換。
2. **UI 智慧變形**：點選出題系統時，側邊欄會自動亮起雲端講義書櫃；點選配詳解系統時，畫面會立刻自動變更為 Excel 上傳與欄位防呆對接。
3. **完美隔離暫存**：模式 A 與模式 B 的下載按鈕和狀態完全獨立，互不干擾！

直接把這整份完全體 `app.py` 推上 GitHub 吧！這下子你的醫學共筆出題控制台就正式達到「完全體大滿貫」了！
