import streamlit as st
import pandas as pd
import re
import math
import io
import json
import urllib.request
import urllib.parse
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

# Word 處理相關
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Google Gemini API 相關 (改用經典穩定版 SDK)
try:
    import google.generativeai as genai
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

# 網頁配置
st.set_page_config(page_title="AI 雲端講義題庫系統", page_icon="🧠", layout="centered")

st.title("🧠 AI 雲端全自動題庫生成系統 (終極穩定版)")

if not HAS_GEMINI:
    st.error("❌ 缺失 google-generativeai 套件，請確認 requirements.txt 已更新。")
    st.stop()

# ==================== 1. 🔍 系統設定 ====================
env_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        env_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

# 預設使用你在程式碼裡寫死的代鑰
hardcoded_key = "AQ.Ab8RN6IQXEwIP4B08KDv9FvC_KtFY1ARYKB_4IpIl-3pncwSCA"
default_key = env_key if env_key else hardcoded_key

with st.expander("🔑 API 金鑰設定面板", expanded=False):
    st.markdown("如果遇到權限錯誤，請直接在下方貼上 API Key 強制覆寫。")
    user_live_key = st.text_input("💡 請輸入 API Key：", 
                                  value=default_key if "AIzaSy" in default_key and "貼在這裡" not in default_key else "",
                                  type="password")

api_key = user_live_key.strip() if user_live_key else default_key

if not api_key or "貼在這裡" in api_key:
    st.warning("⚠️ 請貼入您在 Google AI Studio 申請的 `AIzaSy` 金鑰。")
    st.stop()

# 🚨 核心切換：強制綁定經典版 SDK
genai.configure(api_key=api_key)
# 初始化 2.5 最新模型
model = genai.GenerativeModel('gemini-2.5-flash')

# ==================== 2. 🗂️ GitHub 自動資料夾雙模掃描 ====================
GITHUB_USER = "ShinySean123"
GITHUB_REPO = "20260603--Ver.2_questionbankgenerator"
GITHUB_FOLDER_HIST = "history_db"          
GITHUB_FOLDER_PDF = "current_materials"    

encoded_user = urllib.parse.quote(GITHUB_USER)
encoded_repo = urllib.parse.quote(GITHUB_REPO)

# --- 2A. 智慧歷史題庫掃描 ---
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
        with urllib.request.urlopen(req) as resp: html_text = resp.read().decode('utf-8')
        all_excel_files = list(set(re.findall(r'title="([^"]+\.xlsx)"', html_text)))
    except: pass

if all_excel_files:
    file_options.append("💥 比對資料夾內【所有檔案】（全面防重複）")
    for f in all_excel_files: file_options.append(f)

# --- 2B. 智慧雲端講義書櫃掃描 ---
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
    if not cloud_pdf_files:
        try:
            html_url = f"https://github.com/{encoded_user}/{encoded_repo}/tree/main/{urllib.parse.quote(GITHUB_FOLDER_PDF)}"
            req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as resp: html_text = resp.read().decode('utf-8')
            found_pdfs = re.findall(r'title="([^"]+\.[pP][dD][fF])"', html_text)
            cloud_pdf_files = list(set([urllib.parse.unquote(p) for p in found_pdfs]))
        except Exception: pass

# --- 2C. 側邊欄 UI 渲染 ---
history_titles = []
with st.sidebar:
    st.header("⚙️ 雲端資料庫狀態")
    st.markdown(f"**目前帳號:** `{GITHUB_USER}`")
    selected_mode = st.selectbox("請選擇歷史題庫防重複模式：", file_options)
    
    st.markdown("---")
    st.header("📚 雲端講義書櫃")
    if cloud_pdf_files:
        st.success(f"🟢 成功解鎖！偵測到雲端有 {len(cloud_pdf_files)} 份 PDF 講義")
        selected_cloud_pdfs = st.multiselect("請勾選本次想連動出題的雲端講義：", cloud_pdf_files)
    else:
        st.info(f"ℹ️ 目前雲端 `{GITHUB_FOLDER_PDF}/` 資料夾內未偵測到 PDF。")
        selected_cloud_pdfs = []

# ==================== 讀取雲端二進位流函數 ====================
def fetch_excel_titles(file_name):
    encoded_name = urllib.parse.quote(file_name)
    raw_url = f"https://raw.githubusercontent.com/{encoded_user}/{encoded_repo}/main/{GITHUB_FOLDER_HIST}/{encoded_name}"
    try:
        req = urllib.request.Request(raw_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as resp:
            df = pd.read_excel(io.BytesIO(resp.read()))
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

# ==================== 3. UI 主介面 ====================
st.subheader("📂 Step 1: 選取或上傳課程講義 PDF")

uploaded_pdfs = st.file_uploader("從本機電腦上傳新講義 PDF (可多選)", type=["pdf"], accept_multiple_files=True)
total_pdf_count = len(uploaded_pdfs) + len(selected_cloud_pdfs)

if total_pdf_count > 0:
    st.markdown(f"📊 **目前已鎖定講義總數：{total_pdf_count} 份**")
    if selected_cloud_pdfs: st.caption(f"☁️ 雲端講義：{', '.join(selected_cloud_pdfs)}")
    if uploaded_pdfs: st.caption(f"💻 本地講義：{', '.join([f.name for f in uploaded_pdfs])}")

    st.subheader("📝 Step 2: 設定出題參數")
    col1, col2, col3 = st.columns(3)
    with col1:
        page_range = st.text_input("想根據哪幾頁出題？(填特定頁數，或填『整份』)", "整份")
    with col2:
        topic_name = st.text_input("章節/主題名稱", "醫學綜合領域測驗")
    with col3:
        num_questions = st.number_input("預計生成題數", min_value=1, max_value=50, value=10)

    st.markdown("---")
    col_num, col_blank = st.columns([1, 2])
    with col_num:
        start_q_num = st.number_input("🔢 設定此份題庫的「起始題號」", min_value=1, max_value=999, value=1, step=1)
    st.markdown("---")

    col4, col5 = st.columns(2)
    with col4:
        exam_title_input = st.text_input("Word 考卷大標題", f"{topic_name}_綜合測驗")
    with col5:
        excel_filename_input = st.text_input("Excel 輸出檔名", f"{topic_name}_綜合題庫")

    exam_title = str(exam_title_input) if exam_title_input else "測驗題庫"
    excel_filename = str(excel_filename_input) if excel_filename_input else "精修題庫"

    # ==================== 4. AI 出題與排版核心 (經典穩定版) ====================
    if st.button("⚡ 開始全自動雙模融合出題 ⚡", use_container_width=True):
        try:
            with st.spinner("🧠 正在準備與同步雲端/本地講義檔案..."):
                contents_payload = []
                
                # 處理本地上傳的 PDF
                for pdf_file in uploaded_pdfs:
                    pdf_bytes = pdf_file.read()
                    contents_payload.append({
                        "mime_type": "application/pdf",
                        "data": pdf_bytes
                    })
                
                # 處理雲端講義
                for cloud_pdf_name in selected_cloud_pdfs:
                    c_bytes = fetch_cloud_pdf_bytes(cloud_pdf_name)
                    if c_bytes:
                        contents_payload.append({
                            "mime_type": "application/pdf",
                            "data": c_bytes
                        })

            with st.spinner("🧠 AI 正在通盤研讀所有講義並為您精心設計題目中... 請稍候"):
                range_instruction = f"精準鎖定這些 PDF 檔案中的【{page_range}】" if "整份" not in page_range and "全部" not in page_range else "「通盤掃描並融合這幾份 PDF 檔案」的完整內容，宏觀地在不同的講義、章節與核心觀念中平均分佈提取核心重點"

                history_block = ""
                if history_titles:
                    history_block = "⚠️ 絕對禁止重複、改寫或高度雷同以下這些已經出過的舊題目：\n" + "\n".join([f"- {t}" for t in history_titles])

                prompt = f"""
                你現在是一位資深的醫學與生物科學教授。請根據使用者夾帶的這多份 PDF 檔案，{range_instruction}，並圍繞核心主題【{topic_name}】出 {num_questions} 題五選一的單選題。
                
                {history_block}
                
                輸出的內容必須嚴格遵守以下規則：
                1. 格式必須是 JSON 格式的列表(Array)，內含多個物件，每個物件的Key必須嚴格為：
                   "題目內容", "選項A", "選項B", "選項C", "選項D", "選項E", "正確答案", "針對各選項之詳解", "出處"
                
                2. 只有【針對各選項之詳解】與【出處】欄位必須用繁體中文詳細解釋。詳解必須非常詳細，逐行解釋為什麼該選項正確或錯誤，換行請用 \\n 符號。
                3. 【題目內容】與【選項A】~【選項E】還有【正確答案】請使用「全英文 (Full English)」。
                4. 【出處】格式固定為：「該題目對應之原始PDF完整真實檔名」第 XX 頁。因為本次有多份講義，你必須精準指出這題到底是出自哪一個檔案的第幾頁！
                5. 正確答案（A, B, C, D, E）的總體數量分布要稍微平均一些，但也不要太過絕對的平均。

                請直接輸出完整的 JSON 陣列，不要包含 ```json 等任何 Markdown 外包裝字串。
                """

                # 把 Prompt 加進去
                contents_payload.append(prompt)

                # 發送請求（經典版引擎）
                response = model.generate_content(contents_payload)

                clean_response = response.text.strip()
                if clean_response.startswith("```json"):
                    clean_response = clean_response.split("```json")[1].split("```")[0].strip()
                elif clean_response.startswith("```"):
                    clean_response = clean_response.split("```")[1].split("```")[0].strip()
                    
                raw_questions = json.loads(clean_response)

            with st.spinner("🎨 題目設計完成！正在套用高質感格式排版引擎..."):
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

                # ==================== 5. 產出 Excel ====================
                excel_out = io.BytesIO()
                pd.DataFrame(processed_rows).to_excel(excel_out, index=False)
                excel_out.seek(0)
                
                wb = load_workbook(excel_out)
                ws = wb.active
                col_widths = {'A': 8, 'B': 45, 'C': 30, 'D': 30, 'E': 30, 'F': 30, 'G': 30, 'H': 15, 'I': 60, 'J': 40}
                for letter, width in col_widths.items(): ws.column_dimensions[letter].width = width
                
                border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
                    max_h_lines = 1
                    for cell in row:
                        cell.border = border
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                        if r_idx == 1:
                            cell.font = Font(bold=True)
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        if cell.column_letter in ['A', 'H']: cell.alignment = Alignment(horizontal='center', vertical='center')
                        if r_idx > 1:
                            cw = col_widths.get(cell.column_letter, 20)
                            est = math.ceil((len(str(cell.value)) * 1.8) / cw)
                            if est > max_h_lines: max_h_lines = est
                    if r_idx > 1: ws.row_dimensions[r_idx].height = max_h_lines * 18
                
                final_excel_bytes = io.BytesIO()
                wb.save(final_excel_bytes)

                # ==================== 6. 產出 Word ====================
                doc = Document()
                sec = doc.sections[0]
                sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1.27)
                doc.styles['Normal'].font.name = 'Times New Roman'
                doc.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                doc.styles['Normal'].font.size = Pt(12)
                
                PURPLE, BLUE = RGBColor(112, 48, 160), RGBColor(0, 50, 150)
                title_p = doc.add_paragraph()
                title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = title_p.add_run(exam_title)
                run.bold, run.font.size = True, Pt(16)

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
                                r1 = lp.add_run(m.group(0)); r1.bold, r1.font.color.rgb = True, PURPLE
                                r2 = lp.add_run(line.strip()[len(m.group(0)):]); r2.font.color.rgb = PURPLE
                            else:
                                lp.add_run(line.strip()).font.color.rgb = PURPLE
                    
                    src = str(r['出處'])
                    if src and src.lower() != "nan":
                        sp = doc.add_paragraph(); sp.paragraph_format.space_before = Pt(2)
                        rl = sp.add_run("出處 : "); rl.bold, rl.font.color.rgb = True, BLUE
                        sp.add_run(src).font.color.rgb = BLUE
                    doc.add_paragraph("")

                final_word_bytes = io.BytesIO()
                doc.save(final_word_bytes)

                st.session_state["generated_excel"] = final_excel_bytes.getvalue()
                st.session_state["generated_word"] = final_word_bytes.getvalue()
                st.session_state["saved_excel_filename"] = excel_filename
                st.session_state["saved_exam_title"] = exam_title

        except Exception as e:
            st.error(f"出題過程出錯：{e}")
            st.exception(e)

    # ==================== 5. 獨立下載按鈕 ====================
    if "generated_excel" in st.session_state and "generated_word" in st.session_state:
        st.success("🎉 題庫與試卷皆已設計完成！請在下方直接點擊下載原始檔案：")
        
        def sanitize_f(name): return re.sub(r'[\\/:*?"<>|]', '_', str(name))
        s_excel_name = sanitize_f(st.session_state["saved_excel_filename"])
        s_word_name = sanitize_f(st.session_state["saved_exam_title"])

        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                label="📊 下載精修 Excel 題庫 (.xlsx)",
                data=st.session_state["generated_excel"],
                file_name=f"{s_excel_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with dl_col2:
            st.download_button(
                label="📄 下載精修 Word 試卷 (.docx)",
                data=st.session_state["generated_word"],
                file_name=f"{s_word_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
