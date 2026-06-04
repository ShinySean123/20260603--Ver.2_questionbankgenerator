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

# PDF 文本提取工具
from pypdf import PdfReader

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
st.markdown("內建【本地 PDF 文字解析引擎】，徹底解決 401 上傳憑證報錯！")

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
    api_key = "AQ.Ab8RN6JYf-iaPJ_Ta8FocF8iIrB6b9RoeXvDkB5Rt2Ml1mqCng"

client = genai.Client(api_key=api_key)

# ==================== 2. 🗂️ GitHub API 自動資料夾掃描 ====================
GITHUB_USER = "ShinySean123"
GITHUB_REPO = "20260603--Ver.2_questionbankgenerator"
GITHUB_FOLDER = "history_db"          

encoded_user = urllib.parse.quote(GITHUB_USER)
encoded_repo = urllib.parse.quote(GITHUB_REPO)
encoded_folder = urllib.parse.quote(GITHUB_FOLDER)

github_api_url = f"https://api.github.com/repos/{encoded_user}/{encoded_repo}/contents/{encoded_folder}"
file_options = ["❌ 不使用歷史資料（全新出題）"]
all_excel_files = [] 

try:
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
except Exception:
    pass

history_titles = []
with st.sidebar:
    st.header("⚙️ 雲端資料庫狀態")
    st.markdown(f"**帳號:** `{GITHUB_USER}`\n**專案:** `{GITHUB_REPO}`")
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
            if history_titles: st.success(f"🔥 已鎖定全資料夾共 {len(history_titles)} 題歷史紀錄！")
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
        page_range = st.text_input("想根據哪幾頁出題？(可填特定單一頁數如 64，或填『整份』)", "整份")
    with col2:
        topic_name = st.text_input("章節/主題名稱", "醫學核心領域綜論")
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

    # ==================== 4. AI 出題與排版核心 ====================
    if st.button("⚡ 開始全自動多講義融合出題 ⚡", use_container_width=True):
        try:
            with st.spinner("📖 正在利用本地引擎高速解析 PDF 文本內容..."):
                
                # 建立用來存放所有講義內容的文字容器
                all_pdf_combined_text = ""
                
                for pdf_file in uploaded_pdfs:
                    reader = PdfReader(pdf_file)
                    total_pages = len(reader.pages)
                    
                    # 智慧解析頁數：判斷是要抓單頁還是整份
                    target_pages = []
                    digit_match = re.search(r'\d+', page_range)
                    
                    if digit_match and "整份" not in page_range and "全部" not in page_range:
                        # 使用者指定了特定頁數
                        p_num = int(digit_match.group())
                        if 1 <= p_num <= total_pages:
                            target_pages = [p_num]
                        else:
                            st.warning(f"⚠️ 檔案 {pdf_file.name} 總共只有 {total_pages} 頁，您輸入的第 {p_num} 頁超出範圍，自動切換為讀取整份。")
                            target_pages = list(range(1, total_pages + 1))
                    else:
                        # 讀取整份
                        target_pages = list(range(1, total_pages + 1))
                    
                    # 撈取指定頁面的文字
                    for p_idx in target_pages:
                        page_text = reader.pages[p_idx - 1].extract_text()
                        if page_text:
                            all_pdf_combined_text += f"\n--- 來源檔案: {pdf_file.name} | 頁碼: 第 {p_idx} 頁 ---\n"
                            all_pdf_combined_text += page_text

                if not all_pdf_combined_text.strip():
                    st.error("❌ 無法從您上傳的 PDF 中提取出任何文字（可能是掃描圖片檔），請更換檔案後再試。")
                    st.stop()

            with st.spinner("🧠 文字提取完畢！Gemini AI 正在過濾歷史題庫並精心設計題目中..."):
                
                history_block = ""
                if history_titles:
                    history_block = "⚠️ 絕對禁止重複、改寫或高度雷同以下這些已經出過的舊題目：\n" + "\n".join([f"- {t}" for t in history_titles])

                prompt = f"""
                你現在是一位資深的醫學與生物科學教授。請詳細研讀下方由使用者提供之 PDF 講義的真實文本內容：
                
                【原始講義文本內容開始】
                {all_pdf_combined_text}
                【原始講義文本內容結束】
                
                請圍繞核心主題【{topic_name}】，並針對上述講義內容設計出 {num_questions} 題五選一的單選題。
                
                {history_block}
                
                輸出的內容必須嚴格遵守以下規則：
                1. 格式必須是 JSON 格式的列表(Array)，內含多個物件，每個物件的Key必須嚴格為：
                   "題目內容", "選項A", "選項B", "選項C", "選項D", "選項E", "正確答案", "針對各選項之詳解", "出處"
                
                2. 只有【針對各選項之詳解】與【出處】欄位必須用繁體中文詳細解釋。詳解必須非常詳細，逐行解釋為什麼該選項正確或錯誤，換行請用 \\n 符號。
                3. 【題目內容】與【選項A】~【選項E】還有【正確答案】請使用「全英文 (Full English)」。
                4. 【出處】格式固定為：「對應之原始PDF真實檔名」第 XX 頁。請根據文本上方標註的來源檔案與頁碼，精準指出這題到底是出自哪一個檔案的第幾頁！
                5. 正確答案（A, B, C, D, E）的總體數量分布要稍微平均一些，但也不要太過絕對的平均。

                請直接輸出完整的 JSON 陣列，不要包含 ```json 等任何 Markdown 外包裝字串。
                """

                # 呼叫最新的官方模型，直接用一般的 generate_content 傳遞純文字，完美繞過 401 憑證限制！
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt,
                )

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
                    for lbl in opt_labels:
                        row_dict[f'選項{lbl}'] = str(q.get(f'選項{lbl}', '')).strip()
                    
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

                # 將產出的資料存入暫存狀態，防止刷新消失
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
