import streamlit as st
import pandas as pd
import re
import math
import io
import json
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

# 網頁基礎配置
st.set_page_config(page_title="AI 題庫一體化生成系統", page_icon="🤖", layout="centered")

st.title("🤖 AI 題庫生成與排版一體化系統")
st.markdown("直接上傳講義 PDF，AI 自動研讀出題並輸出完美的 **Word 試卷** 與 **Excel 檔案**！")

if not HAS_GEMINI:
    st.error("❌ 缺失 google-genai 套件，請確認環境或 requirements.txt 中是否有新增。")
    st.stop()

# ==================== 1. 系統設定 (API Key) ====================
# 優先從 Streamlit Secrets 讀取，其次讓使用者在網頁上手動輸入
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    api_key = st.text_input("🔑 請輸入你的 Gemini API Key", type="password")

if not api_key:
    st.warning("請先輸入第一步申請的 Gemini API Key 才能啟用 AI 出題功能。")
    st.stop()

# 初始化新版 Gemini 用戶端
client = genai.Client(api_key=api_key)

# ==================== 2. UI 介面 ====================
uploaded_pdf = st.file_uploader("📂 Step 1: 上傳課程講義 PDF", type=["pdf"])

if uploaded_pdf is not None:
    pdf_name = uploaded_pdf.name
    
    st.subheader("📝 Step 2: 設定出題參數")
    col1, col2, col3 = st.columns(3)
    with col1:
        page_range = st.text_input("想根據哪幾頁出題？", "第 64 頁")
    with col2:
        topic_name = st.text_input("章節/主題名稱", "腦脊髓液CSF")
    with col3:
        num_questions = st.number_input("預計生成題數", min_value=1, max_value=50, value=10)

    col4, col5 = st.columns(2)
    with col4:
        exam_title_input = st.text_input("Word 考卷大標題", f"{topic_name}_練習題")
    with col5:
        excel_filename_input = st.text_input("Excel 輸出檔名", f"{topic_name}_精修題庫")

    exam_title = str(exam_title_input) if exam_title_input else "測驗題庫"
    excel_filename = str(excel_filename_input) if excel_filename_input else "精修題庫"

    # ==================== 3. AI 出題與排版核心 ====================
    if st.button("⚡ 開始全自動生成 (AI 出題 + 自動排版) ⚡", use_container_width=True):
        try:
            with st.spinner("🧠 AI 正在研讀您的 PDF 講義並精心設計題目中... 預計需要 30~60 秒"):
                
                # 讀取 PDF 二進位資料並上傳給 Gemini File API
                pdf_bytes = uploaded_pdf.read()
                uploaded_gemini_file = client.files.upload(
                    file=io.BytesIO(pdf_bytes),
                    config=types.UploadFileConfig(mime_type="application/pdf")
                )

                # 根據您的 4 點核心要求精心設計的 API 級別 Prompt
                prompt = f"""
                你現在是一位資深的醫學與生物科學教授。請根據使用者上傳的 PDF 檔案中【{page_range}】關於【{topic_name}】的內容，
                出 {num_questions} 題五選一的單選題，且不要和前面出過的題目重複。
                
                輸出的內容必須嚴格遵守以下規則：
                1. 格式必須是 JSON 格式的列表(Array)，內含多個物件，每個物件的Key必須嚴格為：
                   "題目內容", "選項A", "選項B", "選項C", "選項D", "選項E", "正確答案", "針對各選項之詳解", "出處"
                
                2. 只有【針對各選項之詳解】與【出處】欄位必須用繁體中文詳細解釋。詳解必須非常詳細，逐行解釋為什麼該選項正確或錯誤，換行請用 \\n 符號。
                3. 【題目內容】與【選項A】~【選項E】還有【正確答案】請使用「全英文 (Full English)」。
                4. 【出處】格式固定為：「{pdf_name}」{page_range}。
                5. 正確答案（A, B, C, D, E）的總體數量分布要稍微平均一些（不要全是C之類的），但也不要太過絕對的平均（保留點隨機性）。

                請直接輸出完整的 JSON 陣列，不要包含 ```json 等任何 Markdown 外包裝字串。
                """

                # 呼叫最新的官方 gemini-2.5-flash 模型
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[uploaded_gemini_file, prompt],
                )

                # 習慣良好：處理完畢立刻刪除雲端暫存檔
                client.files.delete(name=uploaded_gemini_file.name)

                # 解析 JSON 數據並處理防呆
                clean_response = response.text.strip()
                if clean_response.startswith("```json"):
                    clean_response = clean_response.split("```json")[1].split("```")[0].strip()
                elif clean_response.startswith("```"):
                    clean_response = clean_response.split("```")[1].split("```")[0].strip()
                    
                raw_questions = json.loads(clean_response)

            with st.spinner("🎨 題目生成成功！高質感格式排版引擎啟動中..."):
                # --- 4. 資料整理 ---
                processed_rows = []
                opt_labels = ['A', 'B', 'C', 'D', 'E']

                for idx, q in enumerate(raw_questions, 1):
                    row_dict = {
                        '題號': idx,
                        '題目內容': str(q.get('題目內容', '')).strip()
                    }
                    for lbl in opt_labels:
                        row_dict[f'選項{lbl}'] = str(q.get(f'選項{lbl}', '')).strip()
                    
                    ans = str(q.get('正確答案', '')).upper().strip()
                    row_dict['正確答案'] = ans if ans in opt_labels else ""
                    row_dict['針對各選項之詳解'] = str(q.get('針對各選項之詳解', '')).strip()
                    row_dict['出處'] = str(q.get('出處', '')).strip()
                    processed_rows.append(row_dict)

                # ==================== 5. 產出 Excel (含自動列高) ====================
                excel_out = io.BytesIO()
                pd.DataFrame(processed_rows).to_excel(excel_out, index=False)
                excel_out.seek(0)
                
                wb = load_workbook(excel_out)
                ws = wb.active
                
                col_widths = {
                    'A': 8, 'B': 45, 
                    'C': 30, 'D': 30, 'E': 30, 'F': 30, 'G': 30,
                    'H': 15, 'I': 60, 'J': 40
                }
                for letter, width in col_widths.items(): 
                    ws.column_dimensions[letter].width = width
                
                border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
                    max_h_lines = 1
                    for cell in row:
                        cell.border = border
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                        if r_idx == 1:
                            cell.font = Font(bold=True)
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        if cell.column_letter in ['A', 'H']:
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        
                        if r_idx > 1:
                            cw = col_widths.get(cell.column_letter, 20)
                            est = math.ceil((len(str(cell.value)) * 1.8) / cw)
                            if est > max_h_lines: max_h_lines = est
                    if r_idx > 1: ws.row_dimensions[r_idx].height = max_h_lines * 18
                
                final_excel = io.BytesIO()
                wb.save(final_excel)

                # ==================== 6. 產出 Word (全紫詳解無圓點) ====================
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
                            m = re.match(r'^([A-F])\s*([\(（].*?[\)）]|[:：])', line.strip())
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

                word_out = io.BytesIO()
                doc.save(word_out)

                st.success("🎉 全自動生成成功！請點擊下方按鈕下載檔案。")
                
                # 輸出下載按鈕
                def sanitize_f(name): return re.sub(r'[\\/:*?"<>|]', '_', str(name))
                st.download_button("📊 下載 Excel 精修題庫", final_excel.getvalue(), f"{sanitize_f(excel_filename)}.xlsx")
                st.download_button("📄 下載 Word 考卷試卷", word_out.getvalue(), f"{sanitize_f(exam_title)}.docx")

        except Exception as e:
            st.error(f"系統執行出錯，請稍後重試：{e}")