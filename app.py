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

# Google Gemini API 相關
try:
    from google import genai
    from google.genai import types
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

# 網頁配置
st.set_page_config(page_title="AI 雲端全自動題庫系統", page_icon="🧠", layout="centered")

st.title("🧠 AI 雲端全自動題庫生成系統")
st.markdown("支援【自訂起始題號】與【多檔融合】！自動避開 GitHub 歷史重複題目。")

if not HAS_GEMINI:
    st.error("❌ 缺失 google-genai 套件，請在 requirements.txt 中新增。")
    st.stop()

# ==================== 1. 智慧系統設定 ====================
api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if not api_key:
    # 內建本機專用備用金鑰，免手動輸入
    api_key = "AQ.Ab8RN6IIZIJtyv09DngH_BnTiEG1HhsldAVXr_d07K31bykFTw"

client = genai.Client(api_key=api_key)

# ==================== 2. 🗂️ GitHub API 自動資料夾掃描 ====================
# 📝 已經為你更新為正確的 GitHub 資訊！
GITHUB_USER = "ShinySean123"
GITHUB_REPO = "20260603--Ver.2_questionbankgenerator"
GITHUB_FOLDER = "history_db"          

# 對專案名稱與資料夾進行安全的網址編碼，防止特殊符號導致連線失敗
encoded_user = urllib.parse.quote(GITHUB_USER)
encoded_repo = urllib.parse.quote(GITHUB_REPO)
encoded_folder = urllib.parse.quote(GITHUB_FOLDER)

github_api_url = f"https://api.github.com/repos/{encoded_user}/{encoded_repo}/contents/{encoded_folder}"
file_options = ["❌ 不使用歷史資料（全新出題）"]
all_excel_files = [] 

try:
    # 建立一個模擬瀏覽器的 Request，防止被 GitHub API 拒絕
    req = urllib.request.Request(github_api_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as response:
        api_data = json.loads(response.read().decode())
    for item in api_data:
        if item['type'] == 'file' and item['name'].endswith('.xlsx'):
            all_excel_files.append(item['name'])
    if all_excel_files:
        file_options.append("💥 比對資料夾內【所有檔案】（全面防重複）")
        for f in all_excel_files:
            file_options.append(f)
except Exception as e:
    # 後台靜態調試輸出
    pass

history_titles = []
with st.sidebar:
    st.header("⚙️ 雲端資料庫狀態")
    st.markdown(f"**帳號:** `{GITHUB_USER}`")
    st.markdown(f"**專案:** `{GITHUB_REPO}`")
    st.markdown(f"**資料夾:** `{GITHUB_FOLDER}/`")
    
    if len(file_options) == 1:
        st.error("⚠️ 無法讀取雲端題庫！請檢查 GitHub 上是否已建立 `history_db` 資料夾，且裡面至少有一個 .xlsx 檔案。")
    
    selected_mode = st.selectbox("請選擇本次防重複比對模式：", file_options)

def fetch_excel_titles(file_name):
    encoded_name = urllib.parse.quote(file_name)
    raw_url = f"https://raw.githubusercontent.com/{encoded_user}/{encoded_repo}/main/{encoded_folder}/{encoded_name}"
    try:
        req = urllib.request.Request(raw_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as resp:
            df = pd.read_excel(io.BytesIO(resp.read()))
        q_col = next((c for c in df.columns if any(k in str(c) for k in ["題目", "Question"])), None)
        if q_col: return df[q_col].dropna().astype(str).tolist()
    except: pass
    return []

if "【所有檔案】" in selected_mode:
    with st.sidebar:
        with st.spinner("正在打包讀取所有雲端題庫..."):
            for f in all_excel_files: history_titles.extend(fetch_excel_titles(f))
            if history_titles: st.success(f"🔥 終極防護啟動！已鎖定全資料夾共 {len(history_titles)} 題歷史紀錄！")
elif selected_mode != "❌ 不使用歷史資料（全新出題）":
    with st.sidebar:
        with st.spinner(f"正在讀取 {selected_mode}..."):
            history_titles = fetch_excel_titles(selected_mode)
            if history_titles: st.success(f"🟢 成功鎖定《{selected_mode}》中 {len(history_titles)} 題歷史紀錄。")

# ==================== 3. UI 介面 ====================
uploaded_pdfs = st.file_uploader("📂 Step 1: 上傳今天的課程講義 PDF (可多選)", type=["pdf"], accept_multiple_files=True)

if uploaded_pdfs:
    st.subheader("📝 Step 2: 設定出題參數")
    col1, col2, col3 = st.columns(3)
    with col1:
        page_range = st.text_input("想根據哪幾頁出題？(填特定頁數，或填『整份』)", "整份")
    with col2:
        topic_name = st.text_input("章節/主題名稱", "醫學核心領域綜論")
    with col3:
        num_questions = st.number_input("預計生成題數", min_value=1, max_value=50, value=10)

    # 🌟 [全新功能] 讓使用者自訂起始題號
    st.markdown("---")
    col_num, col_blank = st.columns([1, 2])
    with col_num:
        start_q_num = st.number_input("🔢 設定此份題庫的「起始題號」", min_value=1, max_value=999, value=1, step=1, help="例如輸入 51，產出的題目編號就會從 51, 52, 53... 開始往下排。")
    st.markdown("---")

    col4, col5 = st.columns(2)
    with col4:
        exam_title_input = st.text_input("Word 考卷大標題", f"{topic_name}_綜合測驗")
    with col5:
        excel_filename_input = st.text_input("Excel 輸出檔名", f"{topic_name}_綜合題庫")

    exam_title = str(exam_title_input) if exam_title_input else "測驗題庫"
    excel_filename = str(excel_filename_input) if excel_filename_input else "精修題庫"

    # ==================== 4. AI 出題與排版核心 ====================
    if st.button("⚡ 開始全自動多講義融合出題 ⚡", use_container_width=True):
        try:
            with st.spinner(f"🧠 AI 正在同時研讀您上傳的 {len(uploaded_pdfs)} 份講義並精心設計題目中..."):
                
                gemini_file_objects = []
                for pdf_file in uploaded_pdfs:
                    pdf_bytes = pdf_file.read()
                    gemini_file = client.files.upload(
                        file=io.BytesIO(pdf_bytes),
                        config=types.UploadFileConfig(mime_type="application/pdf")
                    )
                    gemini_file_objects.append(gemini_file)

                range_instruction = f"精準鎖定這些 PDF 檔案中的【{page_range}】" if "整份" not in page_range and "全部" not in page_range else "「通盤掃描並融合這幾份 PDF 檔案」的完整內容，宏觀地在不同的講義、章節與核心觀念中平均分佈提取核心重點"

                history_block = ""
                if history_titles:
                    history_block = "⚠️ 絕對禁止重複、改寫或高度雷同以下這些已經出過的舊題目：\n" + "\n".join([f"- {t}" for t in history_titles])

                prompt = f"""
                你現在是一位資深的醫學與生物科學教授。請根據使用者上傳的這多份 PDF 檔案，{range_instruction}，並圍繞核心主題【{topic_name}】出 {num_questions} 題五選一的單選題。
                
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

                contents_payload = []
                for g_file in gemini_file_objects: contents_payload.append(g_file)
                contents_payload.append(prompt)

                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=contents_payload,
                )

                for g_file in gemini_file_objects: client.files.delete(name=g_file.name)

                clean_response = response.text.strip()
                if clean_response.startswith("```json"):
                    clean_response = clean_response.split("```json")[1].split("```")[0].strip()
                elif clean_response.startswith("```"):
                    clean_response = clean_response.split("```")[1].split("```")[0].strip()
                    
                raw_questions = json.loads(clean_response)

            with st.spinner("🎨 題目設計完成！正在套用高質感格式排版引擎..."):
                # --- 5. 資料整理 (加入自訂起始題號邏輯) ---
                processed_rows = []
                opt_labels = ['A', 'B', 'C', 'D', 'E']

                for idx, q in enumerate(raw_questions):
                    # 關鍵：將原本的 idx + 1 改成 start_q_num + idx
                    current_q_num = int(start_q_num) + idx
                    
                    row_dict = {'題號': current_q_num, '題目內容': str(q.get('題目內容', '')).strip()}
                    for lbl in opt_labels:
                        row_dict[f'選項{lbl}'] = str(q.get(f'選項{lbl}', '')).strip()
                    
                    ans = str(q.get('正確答案', '')).upper().strip()
                    row_dict['正確答案'] = ans if ans in opt_labels else ""
                    row_dict['針對各選項之詳解'] = str(q.get('針對各選項之詳解', '')).strip()
                    row_dict['出處'] = str(q.get('出處', '')).strip()
                    processed_rows.append(row_dict)

                # ==================== 6. 產出 Excel ====================
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

                # ==================== 7. 產出 Word ====================
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
                run.bold = True
                run.font.size = Pt(16)

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

                # 將產出的二進位資料存入暫存，防止點擊下載按鈕時消失
                st.session_state["generated_excel"] = final_excel_bytes.getvalue()
                st.session_state["generated_word"] = final_word_bytes.getvalue()
                st.session_state["saved_excel_filename"] = excel_filename
                st.session_state["saved_exam_title"] = exam_title

        except Exception as e:
            st.error(f"出題過程出錯：{e}")
            st.exception(e)

    # ==================== 5. 獨立渲染下載按鈕機制 ====================
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
