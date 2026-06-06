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

# 網頁配置
st.set_page_config(page_title="AI 醫學共筆題庫工作站", page_icon="🧠", layout="centered")

st.title("🧠 AI 醫學共筆題庫雙模工作站")
st.markdown("共筆組長專屬神器：結合【全講義圖文智慧出題】與【現成題目自動配詳解】兩大核心功能！")

# ==================== 1. 🔑 共享 API 金鑰設定面板 ====================
env_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets: env_key = st.secrets["GEMINI_API_KEY"]
except Exception: pass

with st.sidebar:
    st.header("🔑 API 金鑰配置")
    user_live_key = st.text_input("請輸入 API Key：", value=env_key if env_key else "", type="password")
    api_key = user_live_key.strip() if user_live_key else env_key
    
    st.markdown("---")
    st.caption("💡 提示：本工作站全面採用底層 HTTP 直連技術，優化傳輸封裝，確保每次呼叫只扣減最極限的 1 次 RPD 額度。")

if not api_key:
    st.warning("⚠️ 請先在左側邊欄填入您在 Google AI Studio 申請的 `AIzaSy` 金鑰以解鎖系統。")
    st.stop()

# ==================== 🌟 共享的終極單發 HTTP 直連函數 ====================
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
                st.warning(f"⏳ Google 門口短暫大塞車 (503) - 正在重試第 {attempt+1}/{max_retries} 次...")
                time.sleep(wait_time)
                continue
            else:
                raise Exception(f"Google 門口持續 503 塞車，已重試 {max_retries} 次。")
        else:
            raise Exception(f"Google 門口回應錯誤 ({resp.status_code}): {resp.text}")

# ==================== 🗂️ 功能切換導覽器 ====================
# 在首頁最上方渲染切換大按鈕
main_mode = st.radio(
    "🎯 請選擇您目前想要使用的共筆功能模組：",
    ["📚 模組 A：上傳/連動講義 ➡️ 智慧圖文大融合出題", "📝 模組 B：上傳現成題目 ➡️ 專家級醫學詳解補完"],
    index=0,
    horizontal=True
)

st.markdown("---")

# ==============================================================================
# 🌟 模組 A：全自動講義圖文出題系統
# ==============================================================================
if "模組 A" in main_mode:
    st.subheader("📚 模式 A：講義圖文智慧出題（支援心電圖/解剖圖辨識）")
    st.caption("系統將自動將 PDF 講義轉化為極輕量高壓縮影像，讓 AI 直接肉眼看圖出題，並連動 GitHub 歷史資料夾防止題目重複。")

    # --- GitHub 自動資料夾掃描 ---
    GITHUB_USER = "ShinySean123"
    GITHUB_REPO = "20260603--Ver.2_questionbankgenerator"
    GITHUB_FOLDER_HIST = "history_db"          
    GITHUB_FOLDER_PDF = "current_materials"    

    encoded_user = urllib.parse.quote(GITHUB_USER)
    encoded_repo = urllib.parse.quote(GITHUB_REPO)

    # 智慧歷史題庫掃描
    github_api_hist_url = f"https://api.github.com/repos/{encoded_user}/{encoded_repo}/contents/{urllib.parse.quote(GITHUB_FOLDER_HIST)}"
    file_options = ["❌ 不使用歷史資料（全新出題）"]
    all_excel_files = [] 

    try:
        req = urllib.request.Request(github_api_hist_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            api_data = json.loads(response.read().decode())
        for item in api_data:
            if item['type'] == 'file' and item['name'].endswith('.xlsx'): all_excel_files.append(item['name'])
    except Exception:
        try:
            html_url = f"https://github.com/{encoded_user}/{encoded_repo}/tree/main/{urllib.parse.quote(GITHUB_FOLDER_HIST)}"
            req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: html_text = resp.read().decode('utf-8')
            all_excel_files = list(set(re.findall(r'title="([^"]+\.xlsx)"', html_text)))
        except: pass

    if all_excel_files:
        file_options.append("💥 比對資料夾內【所有檔案】（全面防重複）")
        for f in all_excel_files: file_options.append(f)

    # 智慧雲端講義書櫃掃描
    cloud_pdf_files = []
    github_api_pdf_url = f"https://api.github.com/repos/{encoded_user}/{encoded_repo}/contents/{urllib.parse.quote(GITHUB_FOLDER_PDF)}"

    try:
        req = urllib.request.Request(github_api_pdf_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            pdf_api_data = json.loads(response.read().decode())
        for item in pdf_api_data:
            if item['type'] == 'file' and item['name'].lower().endswith('.pdf'): cloud_pdf_files.append(item['name'])
    except Exception:
        try:
            html_url = f"https://github.com/{encoded_user}/{encoded_repo}/tree/main/{urllib.parse.quote(GITHUB_FOLDER_PDF)}"
            req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: html_text = resp.read().decode('utf-8')
            found_pdfs = re.findall(r'title="([^"]+\.[pP][dD][fF])"', html_text)
            cloud_pdf_files = list(set([urllib.parse.unquote(p) for p in found_pdfs]))
        except Exception: pass

    # 側邊欄加載專用 UI
    history_titles = []
    with st.sidebar:
        st.header("⚙️ 雲端書櫃與歷史庫狀態")
        selected_mode = st.selectbox("請選擇歷史題庫防重複模式：", file_options)
        st.markdown("---")
        if cloud_pdf_files:
            st.success(f"🟢 偵測到雲端書櫃有 {len(cloud_pdf_files)} 份 PDF")
            selected_cloud_pdfs = st.multiselect("請勾選本次想連動出題的雲端講義：", cloud_pdf_files)
        else:
            st.info("ℹ️ 目前雲端書櫃內未偵測到 PDF。")
            selected_cloud_pdfs = []

    def fetch_excel_titles(file_name):
        encoded_name = urllib.parse.quote(file_name)
        raw_url = f"https://raw.githubusercontent.com/{encoded_user}/{encoded_repo}/main/{GITHUB_FOLDER_HIST}/{encoded_name}"
        try:
            req = urllib.request.Request(raw_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: df = pd.read_excel(io.BytesIO(resp.read()))
            q_col = next((c for c in df.columns if any(k in str(c) for k in ["題目", "Question"])), None)
            if q_col: return df[q_col].dropna().astype(str).tolist()
        except: pass
        return []

    if "【所有檔案】" in selected_mode:
        for f in all_excel_files: history_titles.extend(fetch_excel_titles(f))
    elif selected_mode != "❌ 不使用歷史資料（全新出題）":
        history_titles = fetch_excel_titles(selected_mode)

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

    # --- 介面渲染與出題參數 ---
    uploaded_pdfs = st.file_uploader("從本機電腦上傳新講義 PDF (可多選)", type=["pdf"], accept_multiple_files=True)
    total_pdf_count = len(uploaded_pdfs) + len(selected_cloud_pdfs)

    if total_pdf_count > 0:
        st.markdown(f"📊 **目前已鎖定講義總數：{total_pdf_count} 份**")
        col1, col2, col3 = st.columns(3)
        with col1: page_range = st.text_input("想根據哪幾頁出題？", "整份")
        with col2: topic_name = st.text_input("章節/主題名稱", "醫學綜合領域測驗")
        with col3: num_questions = st.number_input("預計生成題數", min_value=1, max_value=50, value=10)

        col_num, col_blank = st.columns([1, 2])
        with col_num: start_q_num = st.number_input("🔢 設定「起始題號」", min_value=1, max_value=999, value=1, step=1, key="mode_a_qnum")

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

                with st.spinner("📷 正在為整包講義與醫療圖表進行極速視覺轉碼與大瘦身..."):
                    for pdf_file in uploaded_pdfs: process_pdf_to_compressed_images(pdf_file.read(), pdf_file.name)
                    for cloud_pdf_name in selected_cloud_pdfs:
                        c_bytes = fetch_cloud_pdf_bytes(cloud_pdf_name)
                        if c_bytes: process_pdf_to_compressed_images(c_bytes, cloud_pdf_name)

                with st.spinner("🧠 輕量視覺包封裝完成！AI 正在一次性極速研讀整份圖文講義出題中..."):
                    range_instruction = f"精準鎖定這些影像中的【{page_range}】" if "整份" not in page_range and "全部" not in page_range else "「通盤掃描並融合這幾份講義影像」的完整內容，宏觀地在不同的章節與圖表中提取重點"
                    history_block = ""
                    if history_titles: history_block = "⚠️ 絕對禁止重複、改寫或高度雷同以下這些已經出過的舊題目：\n" + "\n".join([f"- {t}" for t in history_titles])

                    prompt = f"""
                    你現在是一位資深的醫學與生物科學教授。請根據我為你提供的這份完整講義影像（包含文字與所有醫學圖表），{range_instruction}，並圍繞核心主題【{topic_name}】一次性出齊 {num_questions} 題五選一的單選題。
                    請特別發揮你的視覺辨識能力，若講義中有重要的心電圖、流程圖譜或解剖結構，務必將其核心觀念轉化為考題！
                    {history_block}
                    輸出的內容必須嚴格遵守以下規則：
                    1. 格式必須是 JSON 格式的列表(Array)，內含多個物件，每個物件的Key必須嚴格為："題目內容", "選項A", "選項B", "選項C", "選項D", "選項E", "正確答案", "針對各選項之詳解", "出處"
                    2. 只有【針對各選項之詳解】與【出處】欄位必須用繁體中文詳細解釋。
                    3. 【題目內容】與【選項A】~【選項E】還有【正確答案】請使用「全英文 (Full English)」。
                    4. 【出處】請根據我夾帶圖片前方的文字標籤（例如：=== 【檔名】第 X 頁 ===），精準指出這題是出自哪一個檔案的第幾頁！
                    5. 正確答案（A, B, C, D, E）的總體數量分布要稍微平均一些。
                    請直接輸出完整的 JSON 陣列，不要包含 ```json 等任何 Markdown 外包裝字串。
                    """
                    contents_payload.append(prompt)

                    clean_response = generate_content_via_http_with_retry(contents_payload, api_key)
                    if clean_response.startswith("
http://googleusercontent.com/immersive_entry_chip/0
http://googleusercontent.com/immersive_entry_chip/1
http://googleusercontent.com/immersive_entry_chip/2
http://googleusercontent.com/immersive_entry_chip/3

---

### 🎨 這次雙模整合版的優勢：

1. **一鍵橫向切換**：在首頁最上方新增了 `main_mode` 單選按鈕。你只需要輕輕一拉，整個網頁的 UI 與邏輯就會在「講義出題」與「純配詳解」之間一秒切換。
2. **記憶體狀態獨立**：我把兩個功能的 `st.session_state` 下載快取完全切分（`_a` 與 `_b`）。這樣你在用模式 A 出完題後，直接切換到模式 B 幫別的考題配詳解，兩個檔案在後台不會打架或覆蓋。
3. **金鑰一體化**：在左側邊欄填一次 API Key，兩個功能同時啟用，操作極致精簡。

直接把這段終極整合版推上 GitHub 吧！這組二合一工作站，絕對能讓你這個醫三的共筆組長在處理日常考題、複習資料時效率再次翻倍！
