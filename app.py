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
# 使用 chr(96) * 3 動態拼接三個反引號，100% 避免 Git 與 Streamlit Cloud 語法截斷 Bug
TRIPLE_BACKTICK = chr(96) * 3
BT_JSON = TRIPLE_BACKTICK + "json"
BT_ONLY = TRIPLE_BACKTICK

# 共享的 Excel 格式化美化參數
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

st.title("🧠 AI 醫學共筆題庫雙模工作站")
st.markdown("共筆組長專屬神器：結合【全講義圖文智慧出題】與【現成題目自動配詳解】兩大核心功能！")

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
    st.caption("💡 提示：本工作站全面採用底層 HTTP 直連技術，優化傳輸封裝，確保每次呼叫只扣減最極限的 1 次 RPD 額度。")

if not api_key:
    st.warning("⚠️ 請先在左側邊欄填入您在 Google AI Studio 申請的 `AIzaSy` 金鑰以解鎖系統。")
    st.stop()

# ==================== 🌟 共享的終極單發 HTTP 直連函數 ====================
def generate_content_via_http_with_retry(contents_list, api_key, max_retries=4):
    url = f"[https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=){api_key}"
    
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
                st.warning(f"⏳ Google 門口短暫大塞車 (503) - 正在重試第 {attempt+1}/{max_retries} 次...")
                time.sleep(wait_time)
                continue
            else:
                raise Exception(f"Google 門口持續 503 塞車，已重試 {max_retries} 次。")
        else:
            raise Exception(f"Google 門口回應錯誤 ({resp.status_code}): {resp.text}")

# ==================== 🗂️ 功能切換導覽器 ====================
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
    github_api_hist_url = f"[https://api.github.com/repos/](https://api.github.com/repos/){encoded_user}/{encoded_repo}/contents/{urllib.parse.quote(GITHUB_FOLDER_HIST)}"
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
            html_url = f"[https://github.com/](https://github.com/){encoded_user}/{encoded_repo}/tree/main/{urllib.parse.quote(GITHUB_FOLDER_HIST)}"
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

    # 智慧雲端講義書櫃掃描
    cloud_pdf_files = []
    github_api_pdf_url = f"[https://api.github.com/repos/](https://api.github.com/repos/){encoded_user}/{encoded_repo}/contents/{urllib.parse.quote(GITHUB_FOLDER_PDF)}"

    try:
        req = urllib.request.Request(github_api_pdf_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            pdf_api_data = json.loads(response.read().decode())
        for item in pdf_api_data:
            if item['type'] == 'file' and item['name'].lower().endswith('.pdf'): 
                cloud_pdf_files.append(item['name'])
    except Exception:
        try:
            html_url = f"[https://github.com/](https://github.com/){encoded_user}/{encoded_repo}/tree/main/{urllib.parse.quote(GITHUB_FOLDER_PDF)}"
            req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: 
                html_text = resp.read().decode('utf-8')
            found_pdfs = re.findall(r'title="([^"]+\.[pP][dD][fF])"', html_text)
            cloud_pdf_files = list(set([urllib.parse.unquote(p) for p in found_pdfs]))
        except Exception: 
            pass

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
        raw_url = f"[https://raw.githubusercontent.com/](https://raw.githubusercontent.com/){encoded_user}/{encoded_repo}/main/{GITHUB_FOLDER_HIST}/{encoded_name}"
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
        raw_url = f"[https://raw.githubusercontent.com/](https://raw.githubusercontent.com/){encoded_user}/{encoded_repo}/main/{GITHUB_FOLDER_PDF}/{encoded_name}"
        try:
            req = urllib.request.Request(raw_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: 
                return resp.read()
        except:
            try:
                raw_url_alt = f"[https://github.com/](https://github.com/){encoded_user}/{encoded_repo}/raw/main/{GITHUB_FOLDER_PDF}/{encoded_name}"
                req = urllib.request.Request(raw_url_alt, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req) as resp: 
                    return resp.read()
            except: 
                return None

    # --- 介面渲染與出題參數 ---
    uploaded_pdfs = st.file_uploader("從本機電腦上傳新講義 PDF (可多選)", type=["pdf"], accept_multiple_files=True)
    total_pdf_count = len(uploaded_pdfs) + len(selected_cloud_pdfs)

    if total_pdf_count > 0:
        st.markdown(f"📊 **目前已鎖定講義總數：{total_pdf_count} 份**")
        col1, col2, col3 = st.columns(3)
        with col1: 
            page_range = st.text_input("想根據哪幾頁出題？", "整份")
        with col2: 
            topic_name = st.text_input("章節/主題名稱", "醫學綜合領域測驗")
        with col3: 
            num_questions = st.number_input("預計生成題數", min_value=1, max_value=50, value=10)

        col_num, col_blank = st.columns([1, 2])
        with col_num: 
            start_q_num = st.number_input("🔢 設定「起始題號」", min_value=1, max_value=999, value=1, step=1, key="mode_a_qnum")

        col4, col5 = st.columns(2)
        with col4: 
            exam_title_input = st.text_input("Word 考卷大標題", f"{topic_name}_綜合測驗")
        with col5: 
            excel_filename_input = st.text_input("Excel 輸出檔名", f"{topic_name}_綜合題庫")

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
                    for pdf_file in uploaded_pdfs: 
                        process_pdf_to_compressed_images(pdf_file.read(), pdf_file.name)
                    for cloud_pdf_name in selected_cloud_pdfs:
                        c_bytes = fetch_cloud_pdf_bytes(cloud_pdf_name)
                        if c_bytes: 
                            process_pdf_to_compressed_images(c_bytes, cloud_pdf_name)

                with st.spinner("🧠 輕量視覺包封裝完成！AI 正在一次性極速研讀整份圖文講義出題中..."):
                    range_instruction = f"精準鎖定這些影像中的【{page_range}】" if "整份" not in page_range and "全部" not in page_range else "「通盤掃描並融合這幾份講義影像」的完整內容，宏觀地在不同的章節與圖表中提取重點"
                    history_block = ""
                    if history_titles: 
                        history_block = "⚠️ 絕對禁止重複、改寫或高度雷同以下這些已經出過的舊題目：\n" + "\n".join([f"- {t}" for t in history_titles])

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
                    
                    # 🚀 使用安全動態字元清理 Markdown 外包裝，避開 401/Syntax 衝突
                    clean_response = clean_response.strip()
                    if clean_response.startswith(BT_JSON): 
                        clean_response = clean_response.split(BT_JSON)[1].split(BT_ONLY)[0].strip()
                    elif clean_response.startswith(BT_ONLY): 
                        clean_response = clean_response.split(BT_ONLY)[1].split(BT_ONLY)[0].strip()
                        
                    raw_questions = json.loads(clean_response)

                with st.spinner("🎨 題目設計完成！正在套用高質感格式排版引擎..."):
                    processed_rows = []
                    opt_labels = ['A', 'B', 'C', 'D', 'E']
                    for idx, q in enumerate(raw_questions):
                        current_q_num = int(start_q_num) + idx
                        row_dict = {'題號': current_q_num, '題目內容': str(q.get('題目內容', '')).strip()}
                        for lbl in opt_labels: 
                            row_dict[f'選項{lbl}'] = str(q.get(f'選項{lbl}', '')).strip()
                        ans = str(q.get('正確答案', '')).upper().strip()
                        row_dict['正確答案'] = ans if ans in opt_labels else ""
                        row_dict['針對各選項之詳解'] = str(q.get('針對各選項之詳解', '')).strip()
                        row_dict['出處'] = str(q.get('出處', '')).strip()
                        processed_rows.append(row_dict)

                    # --- Excel 格式排版 ---
                    excel_out = io.BytesIO()
                    pd.DataFrame(processed_rows).to_excel(excel_out, index=False)
                    excel_out.seek(0)
                    wb = load_workbook(excel_out)
                    ws = wb.active
                    for letter, width in EXCEL_COL_WIDTHS.items(): 
                        ws.column_dimensions[letter].width = width
                    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
                        for cell in row:
                            cell.border = EXCEL_BORDER
                            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                            if r_idx == 1:
                                cell.font = Font(bold=True)
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            if cell.column_letter in ['A', 'H']: 
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            if r_idx > 1:
                                cw = EXCEL_COL_WIDTHS.get(cell.column_letter, 20)
                                est = math.ceil((len(str(cell.value)) * 1.8) / cw)
                                if est > 1: 
                                    ws.row_dimensions[r_idx].height = est * 18
                    final_excel_bytes = io.BytesIO()
                    wb.save(final_excel_bytes)

                    # --- Word 格式排版 ---
                    doc = Document()
                    sec = doc.sections[0]
                    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1.27)
                    doc.styles['Normal'].font.name = 'Times New Roman'
                    doc.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                    doc.styles['Normal'].font.size = Pt(12)
                    PURPLE, BLUE = RGBColor(112, 48, 160), RGBColor(0, 50, 150)
                    title_p = doc.add_paragraph()
                    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    title_p.add_run(exam_title).bold = True
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
                                else: 
                                    lp.add_run(line.strip()).font.color.rgb = PURPLE
                        src = str(r['出處'])
                        if src and src.lower() != "nan":
                            sp = doc.add_paragraph()
                            sp.paragraph_format.space_before = Pt(2)
                            sp.add_run("出處 : ").bold = True
                            sp.runs[-1].font.color.rgb = BLUE
                            sp.add_run(src).font.color.rgb = BLUE
                        doc.add_paragraph("")

                    final_word_bytes = io.BytesIO()
                    doc.save(final_word_bytes)

                    st.session_state["generated_excel_a"] = final_excel_bytes.getvalue()
                    st.session_state["generated_word_a"] = final_word_bytes.getvalue()
                    st.session_state["saved_excel_filename_a"] = excel_filename
                    st.session_state["saved_exam_title_a"] = exam_title

            except Exception as e:
                st.error(f"出題過程出錯：{e}")

        # 下載按鈕
        if "generated_excel_a" in st.session_state and "generated_word_a" in st.session_state:
            st.success("🎉 模式 A：講義題庫與試卷皆已設計完成！請下載：")
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                st.download_button("📊 下載精修 Excel 題庫 (.xlsx)", data=st.session_state["generated_excel_a"], file_name=f"{st.session_state['saved_excel_filename_a']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with dl_col2:
                st.download_button("📄 下載精修 Word 試卷 (.docx)", data=st.session_state["generated_word_a"], file_name=f"{st.session_state['saved_exam_title_a']}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

# ==============================================================================
# 🌟 模組 B：現成題目自動配詳解系統
# ==============================================================================
else:
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
                st.error("❌ Excel 內找不到對應的『題目內容』或『選項』表頭欄位，請確認名稱。")
                st.stop()

            col_num_b, col_blank_b = st.columns([1, 2])
            with col_num_b: 
                start_q_num_b = st.number_input("🔢 設定「起始題號」", min_value=1, max_value=999, value=1, step=1, key="mode_b_qnum")

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
                        
                        # 🚀 使用安全動態字元清理 Markdown 標籤，避開 Syntax 衝突
                        ai_response = ai_response.strip()
                        if ai_response.startswith(BT_JSON): 
                            ai_response = ai_response.split(BT_JSON)[1].split(BT_ONLY)[0].strip()
                        elif ai_response.startswith(BT_ONLY): 
                            ai_response = ai_response.split(BT_ONLY)[1].split(BT_ONLY)[0].strip()
                            
                        raw_results = json.loads(ai_response)

                    with st.spinner("🎨 詳解補全完成！正在套用高質感格式排版引擎..."):
                        processed_rows_b = []
                        opt_labels = ['A', 'B', 'C', 'D', 'E']
                        for idx, q in enumerate(raw_results):
                            current_q_num = int(start_q_num_b) + idx
                            row_dict = {'題號': current_q_num, '題目內容': str(q.get('題目內容', '')).strip()}
                            for lbl in opt_labels: 
                                row_dict[f'選項{lbl}'] = str(q.get(f'選項{lbl}', '')).strip()
                            ans = str(q.get('正確答案', '')).upper().strip()
                            row_dict['正確答案'] = ans if ans in opt_labels else ""
                            row_dict['針對各選項之詳解'] = str(q.get('針對各選項之詳解', '')).strip()
                            row_dict['出處'] = "醫學資料庫專家詳解"
                            processed_rows_b.append(row_dict)

                        # --- Excel 產出 ---
                        excel_out_b = io.BytesIO()
                        pd.DataFrame(processed_rows_b).to_excel(excel_out_b, index=False)
                        excel_out_b.seek(0)
                        wb_b = load_workbook(excel_out_b)
                        ws_b = wb_b.active
                        for letter, width in EXCEL_COL_WIDTHS.items(): 
                            ws_b.column_dimensions[letter].width = width
                        for r_idx, row in enumerate(ws_b.iter_rows(min_row=1, max_row=ws_b.max_row), 1):
                            for cell in row:
                                cell.border = EXCEL_BORDER
                                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                                if r_idx == 1:
                                    cell.font = Font(bold=True)
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                                if cell.column_letter in ['A', 'H']: 
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                                if r_idx > 1:
                                    cw = EXCEL_COL_WIDTHS.get(cell.column_letter, 20)
                                    est = math.ceil((len(str(cell.value)) * 1.8) / cw)
                                    if est > 1: 
                                        ws_b.row_dimensions[r_idx].height = est * 18
                        final_excel_bytes_b = io.BytesIO()
                        wb_b.save(final_excel_bytes_b)

                        # --- Word 產出 ---
                        doc_b = Document()
                        sec_b = doc_b.sections[0]
                        sec_b.top_margin = sec_b.bottom_margin = sec_b.left_margin = sec_b.right_margin = Cm(1.27)
                        doc_b.styles['Normal'].font.name = 'Times New Roman'
                        doc_b.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                        doc_b.styles['Normal'].font.size = Pt(12)
                        title_p_b = doc_b.add_paragraph()
                        title_p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        title_p_b.add_run("精修醫學題庫_含專家詳解解說").bold = True
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
                                    else: 
                                        lp.add_run(line.strip()).font.color.rgb = PURPLE
                            doc_b.add_paragraph("")

                        final_word_bytes_b = io.BytesIO()
                        doc_b.save(final_word_bytes_b)

                        st.session_state["sol_excel_b"] = final_excel_bytes_b.getvalue()
                        st.session_state["sol_word_b"] = final_word_bytes_b.getvalue()

                except Exception as e:
                    st.error(f"分析過程出錯：{e}")

        # 下載按鈕
        if "sol_excel_b" in st.session_state and "sol_word_b" in st.session_state:
            st.success("🎉 模式 B：現成題目之專家詳解已全數配對補全！請下載：")
            dl_col1_b, dl_col2_b = st.columns(2)
            with dl_col1_b: 
                st.download_button("📊 下載附詳解題庫 (.xlsx)", data=st.session_state["sol_excel_b"], file_name="精修醫學詳解題庫.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with dl_col2_b: 
                st.download_button("📄 下載附詳解試卷 (.docx)", data=st.session_state["sol_word_b"], file_name="精修醫學詳解試卷.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
