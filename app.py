import streamlit as st
import pandas as pd
import re
import math
import io
import json
import time      # 🌟 核心：引入時間套件來控制慢速發送
import random    
import urllib.request
import urllib.parse
import requests  
import base64
import fitz  # PyMuPDF 高畫質影像渲染引擎
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

# Word 處理相關
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==================== 0. 全域常規樣式與字元定義 ====================
TRIPLE_BACKTICK = chr(96) * 3
BT_JSON = TRIPLE_BACKTICK + "json"
BT_ONLY = TRIPLE_BACKTICK

def sanitize_f(name): 
    """全域共用的檔名非法字元過濾器"""
    return re.sub(r'[\\/:*?"<>|]', '_', str(name))

EXCEL_COL_WIDTHS = {
    'A': 8,   # 題號
    'B': 45,  # 題目內容
    'C': 30,  # 選項A
    'D': 30,  # 選項B
    'E': 30,  # 選項C
    'F': 30,  # 選項D
    'G': 30,  # 選項E
    'H': 15,  # 正確答案
    'I': 60,  # 針對各選項之詳解
    'J': 40   # 出處
}
EXCEL_BORDER = Border(
    top=Side(style='thin'), 
    bottom=Side(style='thin'), 
    left=Side(style='thin'), 
    right=Side(style='thin')
)

# 網頁配置
st.set_page_config(page_title="AI 醫學共筆題庫工作站", page_icon="🧠", layout="centered")

st.title("🧠 AI 醫學共筆題庫三模工作站")
st.markdown("共筆組長專屬完全體：整合【講義智慧出題】、【純題配詳解】與【現成題庫轉 Word】三大核心功能！")

# ==================== 1. 🔑 共享 API 金鑰設定面板 ====================
env_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets: 
        env_key = st.secrets["GEMINI_API_KEY"]
except Exception: 
    pass

with st.sidebar:
    st.header("🔑 API 金鑰配置")
    user_live_key = st.text_input("請輸入 API Key：", value=env_key if env_key else "", type="password")
    api_key = user_live_key.strip() if user_live_key else env_key
    
    st.markdown("---")
    st.caption("💡 提示：本工作站全面採用底層 HTTP 直連技術。全新『模組 C』為純排版引擎，完全不消耗任何 API 額度且無需金鑰驗證。")

# 導覽器放置在金鑰檢查前，確保介面正常渲染
main_mode = st.radio(
    "🎯 請選擇您目前想要使用的共筆功能模組：",
    [
        "📚 模組 A：講義圖文智慧出題", 
        "📝 模組 B：現成題目自動配詳解", 
        "📄 模組 C：既有題庫 Excel ➡️ 轉 Word 考卷"
    ],
    index=0,
    horizontal=True
)

st.markdown("---")

if not api_key and main_mode in ["📚 模組 A：講義圖文智慧出題", "📝 模組 B：現成題目自動配詳解"]:
    st.warning("⚠️ 請先在左側邊欄填入您在 Google AI Studio 申請的 `AIzaSy` 金鑰以解鎖系統。")
    st.stop()

# ==================== 🌟 共享的終極單發 HTTP 直連函數 ====================
def generate_content_via_http_with_retry(contents_list, api_key, max_retries=4):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    
    parts = []
    for item in contents_list:
        if isinstance(item, dict) and item.get("mime_type") == "image/jpeg":
            b64_data = base64.b64encode(item["data"]).decode('utf-8')
            parts.append({
                "inlineData": {
                    "mimeType": "image/jpeg",
                    "data": b64_data
                }
            })
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
                st.warning(f"⏳ 門口短暫大塞車 (503) - 正在重試第 {attempt+1}/{max_retries} 次...")
                time.sleep(wait_time)
                continue
            else:
                raise Exception(f"Google 門口持續 503 塞車，已重試 {max_retries} 次。")
        else:
            raise Exception(f"Google 門口回應錯誤 ({resp.status_code}): {resp.text}")

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

    github_api_hist_url = f"https://api.github.com/repos/{encoded_user}/{encoded_repo}/contents/{urllib.parse.quote(GITHUB_FOLDER_HIST)}"
    file_options = ["❌ 不使用歷史資料（全新出題）"]
    all_excel_files = [] 

    try:
        req = urllib.request.Request(github_api_hist_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            api_data = json.loads(response.read().decode())
        for item in api_data:
            if item['type'] == 'file' and item['name'].endswith('.xlsx'): 
                all_excel_files.append(item['name'])
    except Exception:
        try:
            html_url = f"https://github.com/{encoded_user}/{encoded_repo}/tree/main/{urllib.parse.quote(GITHUB_FOLDER_HIST)}"
            req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: 
                html_text = resp.read().decode('utf-8')
            all_excel_files = list(set(re.findall(r'title="([^"]+\.xlsx)"', html_text)))
        except: 
            pass

    if all_excel_files:
        file_options.append("💥 比對資料夾內【所有檔案】（全面防重複）")
        for f in all_excel_files: 
            file_options.append(f)

    cloud_pdf_files = []
    github_api_pdf_url = f"https://api.github.com/repos/{encoded_user}/{encoded_repo}/contents/{urllib.parse.quote(GITHUB_FOLDER_PDF)}"

    try:
        req = urllib.request.Request(github_api_pdf_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            pdf_api_data = json.loads(response.read().decode())
        for item in pdf_api_data:
            if item['type'] == 'file' and item['name'].lower().endswith('.pdf'): 
                cloud_pdf_files.append(item['name'])
    except Exception:
        try:
            html_url = f"https://github.com/{encoded_user}/{encoded_repo}/tree/main/{urllib.parse.quote(GITHUB_FOLDER_PDF)}"
            req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: 
                html_text = resp.read().decode('utf-8')
            found_pdfs = re.findall(r'title="([^"]+\.[pP][dD][fF])"', html_text)
            cloud_pdf_files = list(set([urllib.parse.unquote(p) for p in found_pdfs]))
        except Exception: 
            pass

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
            with urllib.request.urlopen(req) as resp: 
                df = pd.read_excel(io.BytesIO(resp.read()))
            q_col = next((c for c in df.columns if any(k in str(c) for k in ["題目", "Question"])), None)
            if q_col: 
                return df[q_col].dropna().astype(str).tolist()
        except: 
            pass
        return []

    if "【所有檔案】" in selected_mode:
        for f in all_excel_files: 
            history_titles.extend(fetch_excel_titles(f))
    elif selected_mode != "❌ 不使用歷史資料（全新出題）":
        history_titles = fetch_excel_titles(selected_mode)

    def fetch_cloud_pdf_bytes(file_name):
        encoded_name = urllib.parse.quote(file_name)
        raw_url = f"https://raw.githubusercontent.com/{encoded_user}/{encoded_repo}/main/{GITHUB_FOLDER_PDF}/{encoded_name}"
        try:
            req = urllib.request.Request(raw_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: 
                return resp.read()
        except:
            try:
                raw_url_alt = f"https://github.com/{encoded_user}/{encoded_repo}/raw/main/{GITHUB_FOLDER_PDF}/{encoded_name}"
                req = urllib.request.Request(raw_url_alt, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req) as resp: 
                    return resp.read()
            except: 
                return None

    uploaded_pdfs = st.file_uploader("從本機電腦上傳新講義 PDF (可多選)", type=["pdf"], accept_multiple_files=True)
    total_pdf_count = len(uploaded_pdfs) + len(selected_cloud_pdfs)

    if total_pdf_count > 0:
        st.markdown(f"📊 **目前已鎖定講義總數：{total_pdf_count} 份**")
        
        col_q1, col_q2, col_q3 = st.columns(3)
        with col_q1: 
            page_range = st.text_input("想根據哪幾頁出題？", "整份")
        with col_q2: 
            num_questions = st.number_input("預計生成題數", min_value=1, max_value=100, value=10)
        with col_q3: 
            start_q_num = st.number_input("🔢 設定「起始題號」", min_value=1, max_value=999, value=1, step=1, key="mode_a_qnum")

        st.markdown("---")
        st.subheader("🌐 選擇出題語系樣式")
        lang_style = st.radio(
            "請指定 AI 出題時的題目與選項語言：",
            ["1. 中文出題（專有名詞採『英文 (中文)』雙語標記，貼近國考格式）", "2. 全英文出題（題目與選項皆為 Full English，貼近臨床跑台）"],
            index=0,
            horizontal=True
        )

        st.markdown("---")
        st.subheader("🏷️ 設定大標題與檔名")
        
        end_q_num = start_q_num + num_questions - 1
        default_remarks = f"{start_q_num:02d}~{end_q_num:02d}"

        col_t1, col_t2 = st.columns(2)
        with col_t1: subject_name = st.text_input("科目名稱", "生理學", key="sub_a")
        with col_t2: teacher_name = st.text_input("老師名稱", "王大明", key="tea_a")
            
        col_t3, col_t4 = st.columns(2)
        with col_t3: topic_name = st.text_input("課堂主題", "心血管系統", key="top_a")
        with col_t4: 
            remarks = st.text_input("備註 (預設為題號範圍)", value=default_remarks, key="rem_a")

        final_title_filename = f"{subject_name}_{teacher_name}_{topic_name}_{remarks}"
        st.info(f"📁 系統預覽輸出名稱將為：**{final_title_filename}**")

        if st.button("⚡ 開始全自動雙模融合出題 ⚡", use_container_width=True):
            try:
                all_pages_payload = []
                # 將 PDF 拆解為一頁頁的清單
                def process_pdf_to_flat_list(pdf_bytes, pdf_name):
                    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    for i in range(len(doc)):
                        page = doc.load_page(i)
                        pix = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
                        img_data = pix.tobytes("jpeg")
                        all_pages_payload.append({
                            "label": f"=== 【{pdf_name}】第 {i+1} 頁 ===",
                            "img_object": {"mime_type": "image/jpeg", "data": img_data}
                        })

                with st.spinner("📷 正在啟動後台虛擬掃描機，進行高畫質轉碼..."):
                    for pdf_file in uploaded_pdfs: process_pdf_to_flat_list(pdf_file.read(), pdf_file.name)
                    for cloud_pdf_name in selected_cloud_pdfs:
                        c_bytes = fetch_cloud_pdf_bytes(cloud_pdf_name)
                        if c_bytes: process_pdf_to_flat_list(c_bytes, cloud_pdf_name)

                total_extracted_pages = len(all_pages_payload)
                if total_extracted_pages == 0:
                    st.error("❌ 未成功加載任何講義。")
                    st.stop()

                # 🌟 【慢速溫和分片核心機制】：每 2 頁切成一個極輕量 Chunk，原地發呆 3 秒鐘
                CHUNK_SIZE = 2
                total_chunks = math.ceil(total_extracted_pages / CHUNK_SIZE)
                base_q_per_chunk = max(1, math.floor(num_questions / total_chunks))
                
                raw_questions = []
                accumulated_titles = list(history_titles)

                progress_bar = st.progress(0.0)
                status_text = st.empty()

                for chunk_idx in range(total_chunks):
                    start_p = chunk_idx * CHUNK_SIZE
                    end_p = min(start_p + CHUNK_SIZE, total_extracted_pages)
                    
                    # 畫面提示目前進度
                    status_text.markdown(f"🧠 **AI 太空漫步中... 正在溫和處理第 {chunk_idx+1}/{total_chunks} 講義分片 (第 {start_p+1}~{end_p} 頁)...**")
                    
                    chunk_payload = []
                    for p_idx in range(start_p, end_p):
                        chunk_payload.append(all_pages_payload[p_idx]["label"])
                        chunk_payload.append(all_pages_payload[p_idx]["img_object"])
                    
                    current_chunk_q_target = base_q_per_chunk
                    if chunk_idx == total_chunks - 1:  # 最後一個分片，把剩餘配額補齊
                        current_chunk_q_target = num_questions - len(raw_questions)
                    
                    if current_chunk_q_target <= 0:
                        break

                    # 智慧歷史考點去重
                    history_block = ""
                    if accumulated_titles: 
                        history_block = """
                        【🚨 歷史考點去重指令】：
                        請深入分析以下已出的考點機轉、生理機制、臨床指標，本次新題目『絕對禁止』重複測驗這些已考過的機制！
                        """ + "\n".join([f"- 舊題考點: {t}" for t in accumulated_titles[:60]]) + """
                        請從當前講義中，發掘全新、尚未被覆蓋的生理機轉、病理特徵或臨床指標來進行命題。
                        """

                    if "1. 中文出題" in lang_style:
                        lang_prompt = """
                        【語系要求】：
                        - 題目與選項必須使用繁體中文。遇到醫學專有名詞，請嚴格採取「英文搭配括號中文」的方式呈現（例如：Myocardial Infarction (心肌梗塞)）。
                        """
                    else:
                        lang_prompt = """
                        【語系要求】：
                        - 題目與選項必須完全使用純英文 (Full English) 撰寫。
                        """

                    prompt = f"""
                    你現在是一位資深的醫學教授。請根據我為你提供的這組輕量講義分片影像（包含文字與所有醫學圖表），在主題【{topic_name}】下精心設計 {current_chunk_q_target} 題五選一的單選題。
                    
                    {lang_prompt}
                    {history_block}
                    
                    【詳解與出處恆定要求】：
                    - 針對各選項之詳解必須使用繁體中文進行逐行專家級辨析（換行用 \\n）。正確答案固定輸出大寫字母。
                    - 出處格式固定為：「對應之原始PDF檔名」第 XX 頁。

                    輸出格式：JSON 格式的列表(Array)，內含多個物件，每個物件的Key必須嚴格為："題目內容", "選項A", "選項B", "選項C", "選項D", "選項E", "正確答案", "針對各選項之詳解", "出處"
                    請直接輸出完整的 JSON 陣列，不要包含 ```json 等任何 Markdown 外包裝字串。
                    """
                    chunk_payload.append(prompt)

                    # 發送極輕量請求
                    chunk_response_text = generate_content_via_http_with_retry(chunk_payload, api_key)
                    
                    chunk_response_text = chunk_response_text.strip()
                    if chunk_response_text.startswith(BT_JSON): 
                        chunk_response_text = chunk_response_text.split(BT_JSON)[1].split(BT_ONLY)[0].strip()
                    elif chunk_response_text.startswith(BT_ONLY): 
                        chunk_response_text = chunk_response_text.split(BT_ONLY)[1].split(BT_ONLY)[0].strip()
                    
                    try:
                        chunk_questions = json.loads(chunk_response_text)
                        for q in chunk_questions:
                            raw_questions.append(q)
                            q_title = str(q.get('題目內容', '')).strip()
                            if q_title: accumulated_titles.append(q_title)
                    except Exception:
                        pass
                    
                    progress_bar.progress((chunk_idx + 1) / total_chunks)
                    
                    # 🌟 慢速通關大絕招：每跑完一個分片，後台強制原地有氧發呆 3 秒鐘，給 Google 伺服器喘息時間
                    if chunk_idx < total_chunks - 1:
                        time.sleep(3.0)

                status_text.success("🎉 全講義慢速通關研讀完畢！正在啟動高質感格式排版引擎...")
                raw_questions = raw_questions[:num_questions]

                # --- 排版與產出成果 ---
                processed_rows = []
                opt_labels = ['A', 'B', 'C', 'D', 'E']
                for idx, q in enumerate(raw_questions):
                    current_q_num = int(start_q_num) + idx
                    row_dict = {'題號': current_q_num, '題目內容': str(q.get('題目內容', '')).strip()}
                    for lbl in opt_labels: row_dict[f'選項{lbl}'] = str(q.get(f'選項{lbl}', '')).strip()
                    ans = str(q.get('正確答案', '')).upper().strip()
                    row_dict['正確答案'] = ans if ans in opt_labels else ""
                    row_dict['針對各選項之詳解'] = str(q.get('針對各選項之詳解', '')).strip()
                    row_dict['出處'] = str(q.get('出處', '')).strip()
                    processed_rows.append(row_dict)

                excel_out = io.BytesIO()
                pd.DataFrame(processed_rows).to_excel(excel_out, index=False)
                excel_out.seek(0)
                wb = load_workbook(excel_out)
                ws = wb.active
                for letter, width in EXCEL_COL_WIDTHS.items(): ws.column_dimensions[letter].width = width
                for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
                    for cell in row:
                        cell.border = EXCEL_BORDER
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                        if r_idx == 1:
                            cell.font = Font(bold=True)
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        if cell.column_letter in ['A', 'H']: cell.alignment = Alignment(horizontal='center', vertical='center')
                        if r_idx > 1:
                            cw = EXCEL_COL_WIDTHS.get(cell.column_letter, 20)
                            est = math.ceil((len(str(cell.value)) * 1.8) / cw)
                            if est > 1: ws.row_dimensions[r_idx].height = est * 18
                final_excel_bytes = io.BytesIO()
                wb.save(final_excel_bytes)

                doc = Document()
                sec = doc.sections[0]
                sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1.27)
                doc.styles['Normal'].font.name = 'Times New Roman'
                doc.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                doc.styles['Normal'].font.size = Pt(12)
                PURPLE, BLUE = RGBColor(112, 48, 160), RGBColor(0, 50, 150)
                
                title_p = doc.add_paragraph()
                title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                title_p.add_run(final_title_filename).bold = True
                title_p.runs[-1].font.size = Pt(16)

                for r in processed_rows:
                    doc.add_paragraph(f"{r['題號']}. {r['題目內容']}").paragraph_format.space_after = Pt(6)
                    for lbl in opt_labels:
                        txt = r.get(f'選項{lbl}', '')
                        if txt:
                            op = doc.add_paragraph(f"({lbl}) {txt}")
                            op.paragraph_format.left_indent, op.paragraph_format.space_after = Pt(18), Pt(0)
                    ans_p = doc.add_paragraph()
                    ans_p.paragraph_format.space_before = Pt(6)
                    ans_p.add_run("Ans : ").bold = True
                    ans_p.add_run(f"({r['正確答案']})")
                    expl = str(r['針對各選項之詳解'])
                    if expl and expl.lower() != "nan":
                        h = doc.add_paragraph()
                        h.paragraph_format.space_before, h.paragraph_format.space_after = Pt(4), Pt(0)
                        run = h.add_run("詳解 :"); run.bold, run.font.color.rgb = True, PURPLE
                        for line in expl.split('\n'):
                            if not line.strip(): continue
                            lp = doc.add_paragraph()
                            lp.paragraph_format.left_indent, lp.paragraph_format.space_after = Pt(18), Pt(2)
                            m = re.match(r'^([A-F])\s*([\(（].*?[\)隱]|[:：])', line.strip())
                            if m:
                                lp.add_run(m.group(0)).bold = True
                                lp.runs[-1].font.color.rgb = PURPLE
                                lp.add_run(line.strip()[len(m.group(0)):]).font.color.rgb = PURPLE
                            else: lp.add_run(line.strip()).font.color.rgb = PURPLE
                    src = str(r['出處'])
                    if src and src.lower() != "nan":
                        sp = doc.add_paragraph(); sp.paragraph_format.space_before = Pt(2)
                        sp.add_run("出處 : ").bold = True; sp.runs[-1].font.color.rgb = BLUE
                        sp.add_run(src).font.color.rgb = BLUE
                    doc.add_paragraph("")

                final_word_bytes = io.BytesIO()
                doc.save(final_word_bytes)

                st.session_state["generated_excel_a"] = final_excel_bytes.getvalue()
                st.session_state["generated_word_a"] = final_word_bytes.getvalue()
                st.session_state["saved_exam_title_a"] = final_title_filename

            except Exception as e:
                st.error(f"出題過程出錯：{e}")

        if "generated_excel_a" in st.session_state and "generated_word_a" in st.session_state:
            st.success("🎉 模式 A：講義題庫與試卷皆已設計完成！請下載：")
            s_name = sanitize_f(st.session_state["saved_exam_title_a"])
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1: st.download_button("📊 下載精修 Excel 題庫 (.xlsx)", data=st.session_state["generated_excel_a"], file_name=f"{s_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with dl_col2: st.download_button("📄 下載精修 Word 試卷 (.docx)", data=st.session_state["generated_word_a"], file_name=f"{s_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

# ==============================================================================
# 🌟 模組 B：現成題目自動配詳解系統
# ==============================================================================
elif "模組 B" in main_mode:
    st.subheader("📝 模式 B：現成題目自動配詳解（超輕量・不消耗 Token）")
    st.caption("上傳現有的題目 Excel 檔，AI 將原封不動為您欄位對接，並逐題配上高質感的繁體中文醫學詳解與選項辨析。")

    uploaded_file = st.file_uploader("請選擇包含純題目的 Excel 檔案 (.xlsx)", type=["xlsx"], key="mode_b_uploader")

    if uploaded_file:
        try:
            df_input = pd.read_excel(uploaded_file)
            st.success(f"📊 成功讀取現有考題！共偵測到 **{len(df_input)}** 道題目。")
            
            col_q = next((c for c in df_input.columns if any(k in str(c) for k in ["題目", "Question", "內容"])), None)
            col_a = next((c for c in df_input.columns if "A" in str(c)), None)
            col_b = next((c for c in df_input.columns if "B" in str(c)), None)
            col_c = next((c for c in df_input.columns if "C" in str(c)), None)
            col_d = next((c for c in df_input.columns if "D" in str(c)), None)
            col_e = next((c for c in df_input.columns if "E" in str(c)), None)

            if not (col_q and col_a and col_b):
                st.error("❌ Excel 內找不到對應的『題目內容』或『選項』表頭欄位，請檢查 Excel 架構。")
                st.stop()

            start_q_num_b = st.number_input("🔢 設定「起始題號」", min_value=1, max_value=999, value=1, step=1, key="mode_b_qnum")
            end_q_num_b = start_q_num_b + len(df_input) - 1
            calculated_remarks_b = f"{start_q_num_b:02d}~{end_q_num_b:02d}"

            st.markdown("---")
            st.subheader("🏷️ 設定大標題與檔名")
            col_t1_b, col_t2_b = st.columns(2)
            with col_t1_b: subject_name_b = st.text_input("科目名稱", "生理學", key="sub_b")
            with col_t2_b: teacher_name_b = st.text_input("老師名稱", "王大明", key="tea_b")
                
            col_t3_b, col_t4_b = st.columns(2)
            with col_t3_b: topic_name_b = st.text_input("課堂主題", "心血管系統", key="top_b")
            with col_t4_b: remarks_b = st.text_input("備註 (預設為題號範圍)", value=calculated_remarks_b, key="rem_b")

            final_title_filename_b = f"{subject_name_b}_{teacher_name_b}_{topic_name_b}_{remarks_b}"
            st.info(f"📁 系統預覽輸出名稱將為：**{final_title_filename_b}**")

            cleaned_questions = []
            for idx, row in df_input.iterrows():
                cleaned_questions.append({
                    "orig_idx": idx + 1,
                    "題目內容": str(row[col_q]).strip(),
                    "選項A": str(row[col_a]).strip() if pd.notna(row[col_a]) else "",
                    "選項B": str(row[col_b]).strip() if pd.notna(row[col_b]) else "",
                    "選項C": str(row[col_c]).strip() if pd.notna(row[col_c]) else "",
                    "選項D": str(row[col_d]).strip() if pd.notna(row[col_d]) else "",
                    "選項E": str(row[col_e]).strip() if col_e and pd.notna(row[col_e]) else ""
                })

            if st.button("⚡ 開始全自動配對醫學詳解 ⚡", use_container_width=True):
                try:
                    with st.spinner("🧠 正在啟動醫學核心知識庫，逐題補全詳解中... 請稍候"):
                        input_data_json = json.dumps(cleaned_questions, ensure_ascii=False)
                        
                        prompt = f"""
                        你現在是一位資深的醫學與生物科學教授。請根據我提供給你的 JSON 題目列表，【原封不動】地保留題目內容與選項，並補上最精準的【正確答案】以及極為詳細的【針對各選項之詳解】。
                        嚴格規則：
                        1. 絕對不允許修改我給你的「題目內容」與「選項A」~「選項E」中的任何一個字。
                        2. 【針對各選項之詳解】必須用繁體中文詳細解釋。逐行解釋為什麼該選項正確或錯誤，換行請用 \\n 符號。
                        3. 【正確答案】請固定輸出大寫字母（A, B, C, D 或 E）。
                        4. 輸出格式必須嚴格符合 JSON 列表(Array)，Key 必須為："題目內容", "選項A", "選項B", "選項C", "選項D", "選項E", "正確答案", "針對各選項之詳解"
                        請直接輸出完整的 JSON 陣列，不要包含 ```json 等包裝。

                        原始題目列表：
                        {input_data_json}
                        """
                        
                        ai_response = generate_content_via_http_with_retry([prompt], api_key)
                        
                        ai_response = ai_response.strip()
                        if ai_response.startswith(BT_JSON): ai_response = ai_response.split(BT_JSON)[1].split(BT_ONLY)[0].strip()
                        elif ai_response.startswith(BT_ONLY): ai_response = ai_response.split(BT_ONLY)[1].split(BT_ONLY)[0].strip()
                        raw_results = json.loads(ai_response)

                    with st.spinner("🎨 詳解補全完成！正在套用高質感格式排版引擎..."):
                        processed_rows_b = []
                        opt_labels = ['A', 'B', 'C', 'D', 'E']
                        for idx, q in enumerate(raw_results):
                            current_q_num = int(start_q_num_b) + idx
                            row_dict = {'題號': current_q_num, '題目內容': str(q.get('題目內容', '')).strip()}
                            for lbl in opt_labels: row_dict[f'選項{lbl}'] = str(q.get(f'選項{lbl}', '')).strip()
                            ans = str(q.get('正確答案', '')).upper().strip()
                            row_dict['正確答案'] = ans if ans in opt_labels else ""
                            row_dict['針對各選項之詳解'] = str(q.get('針對各選項之詳解', '')).strip()
                            row_dict['出處'] = "醫學資料庫專家詳解"
                            processed_rows_b.append(row_dict)

                        excel_out_b = io.BytesIO()
                        pd.DataFrame(processed_rows_b).to_excel(excel_out_b, index=False)
                        excel_out_b.seek(0)
                        wb_b = load_workbook(excel_out_b)
                        ws_b = wb_b.active
                        for letter, width in EXCEL_COL_WIDTHS.items(): ws_b.column_dimensions[letter].width = width
                        for r_idx, row in enumerate(ws_b.iter_rows(min_row=1, max_row=ws_b.max_row), 1):
                            for cell in row:
                                cell.border = EXCEL_BORDER
                                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                                if r_idx == 1:
                                    cell.font = Font(bold=True)
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                                if cell.column_letter in ['A', 'H']: cell.alignment = Alignment(horizontal='center', vertical='center')
                                if r_idx > 1:
                                    cw = EXCEL_COL_WIDTHS.get(cell.column_letter, 20)
                                    est = math.ceil((len(str(cell.value)) * 1.8) / cw)
                                    if est > 1: ws_b.row_dimensions[r_idx].height = est * 18
                        final_excel_bytes_b = io.BytesIO()
                        wb_b.save(final_excel_bytes_b)

                        doc_b = Document()
                        sec_b = doc_b.sections[0]
                        sec_b.top_margin = sec_b.bottom_margin = sec_b.left_margin = sec_b.right_margin = Cm(1.27)
                        doc_b.styles['Normal'].font.name = 'Times New Roman'
                        doc_b.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                        doc_b.styles['Normal'].font.size = Pt(12)
                        PURPLE = RGBColor(112, 48, 160)
                        
                        title_p_b = doc_b.add_paragraph()
                        title_p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        title_p_b.add_run(final_title_filename_b).bold = True
                        title_p_b.runs[-1].font.size = Pt(16)

                        for r in processed_rows_b:
                            doc_b.add_paragraph(f"{r['題號']}. {r['題目內容']}").paragraph_format.space_after = Pt(6)
                            for lbl in opt_labels:
                                txt = r.get(f'選項{lbl}', '')
                                if txt:
                                    op = doc_b.add_paragraph(f"({lbl}) {txt}")
                                    op.paragraph_format.left_indent, op.paragraph_format.space_after = Pt(18), Pt(0)
                            ans_p = doc_b.add_paragraph()
                            ans_p.paragraph_format.space_before = Pt(6)
                            ans_p.add_run("Ans : ").bold = True
                            ans_p.add_run(f"({r['正確答案']})")
                            expl = str(r['針對各選項之詳解'])
                            if expl and expl.lower() != "nan":
                                h = doc_b.add_paragraph()
                                h.paragraph_format.space_before, h.paragraph_format.space_after = Pt(4), Pt(0)
                                run = h.add_run("詳解 :"); run.bold, run.font.color.rgb = True, PURPLE
                                for line in expl.split('\n'):
                                    if not line.strip(): continue
                                    lp = doc_b.add_paragraph()
                                    lp.paragraph_format.left_indent, lp.paragraph_format.space_after = Pt(18), Pt(2)
                                    m = re.match(r'^([A-F])\s*([\(（].*?[\)隱]|[:：])', line.strip())
                                    if m:
                                        lp.add_run(m.group(0)).bold = True
                                        lp.runs[-1].font.color.rgb = PURPLE
                                        lp.add_run(line.strip()[len(m.group(0)):]).font.color.rgb = PURPLE
                                    else: lp.add_run(line.strip()).font.color.rgb = PURPLE
                            doc_b.add_paragraph("")

                        final_word_bytes_b = io.BytesIO()
                        doc_b.save(final_word_bytes_b)

                        st.session_state["sol_excel_b"] = final_excel_bytes_b.getvalue()
                        st.session_state["sol_word_b"] = final_word_bytes_b.getvalue()
                        st.session_state["saved_exam_title_b"] = final_title_filename_b

                except Exception as e: st.error(f"分析過程出錯：{e}")

            if "sol_excel_b" in st.session_state and "sol_word_b" in st.session_state:
                st.success("🎉 模式 B：現成題目之專家詳解已全數配對補全！請下載：")
                s_name_b = sanitize_f(st.session_state["saved_exam_title_b"])
                dl_col1_b, dl_col2_b = st.columns(2)
                with dl_col1_b: st.download_button("📊 下載附詳解題庫 (.xlsx)", data=st.session_state["sol_excel_b"], file_name=f"{s_name_b}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with dl_col2_b: st.download_button("📄 下載附詳解試卷 (.docx)", data=st.session_state["sol_word_b"], file_name=f"{s_name_b}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        except Exception as e: st.error(f"讀取 Excel 檔案發生錯誤：{e}")

# ==============================================================================
# 🌟 模組 C：既有題庫已含詳解 Excel ➡️ 直接轉成 Word 考卷
# ==============================================================================
else:
    st.subheader("📄 模式 C：Excel 已含詳解 ➡️ 直接高品質渲染 Word 考卷")
    st.caption("【免金鑰、免額度】直接將現有含有詳解與正確答案的 Excel 資料表，原封不動快速編排為具備標準字型與詳解高亮色的 Word 試卷檔。")

    uploaded_file_c = st.file_uploader("請選擇包含完整題目與詳解的 Excel 檔案 (.xlsx)", type=["xlsx"], key="mode_c_uploader")

    if uploaded_file_c:
        try:
            df_input_c = pd.read_excel(uploaded_file_c)
            st.success(f"📊 成功讀取現有題庫！共偵測到 **{len(df_input_c)}** 道題庫內容。")

            col_q = next((c for c in df_input_c.columns if any(k in str(c) for k in ["題目", "Question", "內容"])), None)
            col_a = next((c for c in df_input_c.columns if "A" in str(c)), None)
            col_b = next((c for c in df_input_c.columns if "B" in str(c)), None)
            col_c = next((c for c in df_input_c.columns if "C" in str(c)), None)
            col_d = next((c for c in df_input_c.columns if "D" in str(c)), None)
            col_e = next((c for c in df_input_c.columns if "E" in str(c)), None)
            col_ans = next((c for c in df_input_c.columns if any(k in str(c) for k in ["答案", "正確", "Answer"])), None)
            col_expl = next((c for c in df_input_c.columns if any(k in str(c) for k in ["詳解", "解析", "Explain", "Explanation"])), None)
            col_src = next((c for c in df_input_c.columns if any(k in str(c) for k in ["出處", "來源", "Source"])), None)

            if not (col_q and col_a and col_b):
                st.error("❌ 找不到基本的『題目內容』或『選項』欄位，請確認 Excel 表頭。")
                st.stop()

            start_q_num_c = st.number_input("🔢 設定「起始題號」", min_value=1, max_value=999, value=1, step=1, key="mode_c_qnum")
            end_q_num_c = start_q_num_c + len(df_input_c) - 1
            calculated_remarks_c = f"{start_q_num_c:02d}~{end_q_num_c:02d}"

            st.markdown("---")
            st.subheader("🏷️ 設定大標題與檔名")
            col_t1_c, col_t2_c = st.columns(2)
            with col_t1_c: subject_name_c = st.text_input("科目名稱", "生理學", key="sub_c")
            with col_t2_c: teacher_name_c = st.text_input("老師名稱", "王大明", key="tea_c")
                
            col_t3_c, col_t4_c = st.columns(2)
            with col_t3_c: topic_name_c = st.text_input("課堂主題", "心血管系統", key="top_c")
            with col_t4_c: remarks_c = st.text_input("備註 (預設為題號範圍)", value=calculated_remarks_c, key="rem_c")

            final_title_filename_c = f"{subject_name_c}_{teacher_name_c}_{topic_name_c}_{remarks_c}"
            st.info(f"📁 系統預覽輸出名稱將為：**{final_title_filename_c}**")

            if st.button("📥 一鍵原封不動轉換為 Word 試卷 📥", use_container_width=True):
                with st.spinner("🎨 正在啟動排版引擎，進行字型美化、段落縮排與高亮著色中..."):
                    doc_c = Document()
                    sec_c = doc_c.sections[0]
                    sec_c.top_margin = sec_c.bottom_margin = sec_c.left_margin = sec_c.right_margin = Cm(1.27)
                    doc_c.styles['Normal'].font.name = 'Times New Roman'
                    doc_c.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                    doc_c.styles['Normal'].font.size = Pt(12)
                    PURPLE, BLUE = RGBColor(112, 48, 160), RGBColor(0, 50, 150)
                    
                    title_p = doc_c.add_paragraph()
                    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    title_p.add_run(final_title_filename_c).bold = True
                    title_p.runs[-1].font.size = Pt(16)

                    opt_labels = ['A', 'B', 'C', 'D', 'E']
                    for idx, row in df_input_c.iterrows():
                        current_q_num = int(start_q_num_c) + idx
                        q_txt = str(row[col_q]).strip()
                        doc_c.add_paragraph(f"{current_q_num}. {q_txt}").paragraph_format.space_after = Pt(6)
                        
                        cols_opts = [col_a, col_b, col_c, col_d, col_e]
                        for lbl, c_opt in zip(opt_labels, cols_opts):
                            if c_opt and pd.notna(row[c_opt]):
                                opt_txt = str(row[c_opt]).strip()
                                if opt_txt:
                                    op = doc_c.add_paragraph(f"({lbl}) {opt_txt}")
                                    op.paragraph_format.left_indent, op.paragraph_format.space_after = Pt(18), Pt(0)
                        
                        if col_ans and pd.notna(row[col_ans]):
                            ans_txt = str(row[col_ans]).upper().strip()
                            ans_p = doc_c.add_paragraph(); ans_p.paragraph_format.space_before = Pt(6)
                            ans_p.add_run("Ans : ").bold = True; ans_p.add_run(f"({ans_txt})")
                        
                        if col_expl and pd.notna(row[col_expl]):
                            expl_txt = str(row[col_expl]).strip()
                            if expl_txt and expl_txt.lower() != "nan":
                                h = doc_c.add_paragraph(); h.paragraph_format.space_before, h.paragraph_format.space_after = Pt(4), Pt(0)
                                run = h.add_run("詳解 :"); run.bold, run.font.color.rgb = True, PURPLE
                                for line in expl_txt.split('\n'):
                                    if not line.strip(): continue
                                    lp = doc_c.add_paragraph(); lp.paragraph_format.left_indent, lp.paragraph_format.space_after = Pt(18), Pt(2)
                                    m = re.match(r'^([A-F])\s*([\(（].*?[\)隱]|[:：])', line.strip())
                                    if m:
                                        lp.add_run(m.group(0)).bold = True; lp.runs[-1].font.color.rgb = PURPLE
                                        lp.add_run(line.strip()[len(m.group(0)):]).font.color.rgb = PURPLE
                                    else: lp.add_run(line.strip()).font.color.rgb = PURPLE
                        
                        if col_src and pd.notna(row[col_src]):
                            src_txt = str(row[col_src]).strip()
                            if src_txt and src_txt.lower() != "nan":
                                sp = doc_c.add_paragraph(); sp.paragraph_format.space_before = Pt(2)
                                sp.add_run("出處 : ").bold = True; sp.runs[-1].font.color.rgb = BLUE
                                sp.add_run(src_txt).font.color.rgb = BLUE
                        doc_c.add_paragraph("")

                    final_word_bytes_c = io.BytesIO()
                    doc_c.save(final_word_bytes_c)
                    st.session_state["sol_word_c"] = final_word_bytes_c.getvalue()
                    st.session_state["saved_exam_title_c"] = final_title_filename_c

            if "sol_word_c" in st.session_state:
                st.success("🎉 模式 C：Word 考卷排版渲染已完美達成！請點擊下載：")
                s_name_c = sanitize_f(st.session_state["saved_exam_title_c"])
                st.download_button(label="📄 下載精修排版 Word 試卷 (.docx)", data=st.session_state["sol_word_c"], file_name=f"{s_name_c}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        except Exception as e: st.error(f"轉換排版過程發生錯誤：{e}")
