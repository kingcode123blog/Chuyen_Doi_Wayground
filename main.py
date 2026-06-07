from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import os
import uuid
import uvicorn

app = FastAPI(title="Smart Exam Online")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Thư mục lưu file tạm thời
BASE_DIR = "/tmp" if os.path.exists("/tmp") else "."

# --- GIAO DIỆN HTML CẬP NHẬT: HIỂN THỊ 3 NÚT TẢI FILE RIÊNG BIỆT ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Smart Exam Tool Online</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 20px auto; padding: 20px; background-color: #f4f6f9; color: #333;}
        .container { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        textarea { width: 100%; height: 250px; padding: 12px; font-size: 14px; border: 1px solid #ccc; border-radius: 5px; box-sizing: border-box; resize: vertical;}
        .action-btn { background-color: #007bff; color: white; border: none; padding: 12px 20px; font-size: 16px; cursor: pointer; border-radius: 5px; margin-top: 15px; width: 100%; font-weight: bold;}
        .action-btn:hover { background-color: #0056b3; }
        .action-btn:disabled { background-color: #6c757d; }
        #status { margin-top: 15px; font-weight: bold; text-align: center; }
        
        /* Style cho vùng hiển thị 3 file kết quả */
        .result-container { margin-top: 25px; padding: 20px; border: 2px dashed #28a745; border-radius: 8px; background-color: #f8fff9; display: none; }
        .result-title { margin-top: 0; color: #28a745; text-align: center; }
        .file-list { display: flex; flex-direction: column; gap: 10px; margin-top: 15px; }
        .download-link { display: flex; align-items: center; justify-content: space-between; padding: 12px; background: white; border: 1px solid #ddd; border-radius: 5px; text-decoration: none; color: #333; font-weight: bold; transition: all 0.2s; }
        .download-link:hover { border-color: #28a745; background-color: #e8f5e9; transform: translateX(5px); }
        .btn-dl { background-color: #28a745; color: white; padding: 6px 12px; border-radius: 4px; font-size: 13px; }
    </style>
</head>
<body>
    <div class="container">
        <h2>🚀 Tool Xử Lý Đề Thi Trực Tuyến</h2>
        <p>Dán nội dung văn bản (Đề thi + Đáp án ở cuối) vào ô bên dưới:</p>
        <textarea id="examText" placeholder="Câu 1: Thủ đô của Việt Nam là?&#10;A. Hà Nội&#10;B. TP.HCM...&#10;&#10;PHẦN ĐÁP ÁN&#10;Câu 1: Đáp án A"></textarea>
        <button onclick="sendTextToServer()" id="submitBtn" class="action-btn">✨ XỬ LÝ ĐỀ THI</button>
        
        <div id="status"></div>

        <div id="resultZone" class="result-container">
            <h3 class="result-title">🎉 Xử Lý Hoàn Tất! Mời Bạn Tải File:</h3>
            <div class="file-list" id="fileList"></div>
        </div>
    </div>

    <script>
        async function sendTextToServer() {
            const text = document.getElementById("examText").value;
            const btn = document.getElementById("submitBtn");
            const statusText = document.getElementById("status");
            const resultZone = document.getElementById("resultZone");
            const fileList = document.getElementById("fileList");

            if (!text.trim()) {
                alert("Vui lòng nhập nội dung đề thi!");
                return;
            }

            btn.innerText = "☕ Server đang phân tích dữ liệu, đợi tí nhé...";
            btn.disabled = true;
            statusText.innerText = "";
            resultZone.style.display = "none"; // Ẩn vùng kết quả cũ nếu có

            try {
                const response = await fetch("/api/process", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ text: text })
                });

                if (!response.ok) throw new Error("Lỗi kết nối tới Server");

                const data = await response.getReader ? await response.json() : await response.json();
                
                if (data.error) throw new Error(data.error);

                // Tạo giao diện cho 3 file tải riêng biệt
                fileList.innerHTML = `
                    <a class="download-link" href="${data.excel_url}" target="_blank">
                        <span>📗 1. File Excel Import Quizizz</span>
                        <span class="btn-dl">Tải về</span>
                    </a>
                    <a class="download-link" href="${data.word_bold_url}" target="_blank">
                        <span>📕 2. File Word Đề + Đáp Án In Đậm Đỏ</span>
                        <span class="btn-dl">Tải về</span>
                    </a>
                    <a class="download-link" href="${data.word_std_url}" target="_blank">
                        <span>📘 3. File Word Đề Riêng - Giải Chi Tiết Riêng</span>
                        <span class="btn-dl">Tải về</span>
                    </a>
                `;
                
                // Hiện vùng tải file lên
                resultZone.style.display = "block";
                statusText.style.color = "green";
                statusText.innerText = "Đã chuẩn bị xong dữ liệu bên dưới!";
            } catch (error) {
                statusText.style.color = "red";
                statusText.innerText = "❌ Đã xảy ra lỗi: " + error.message;
            } finally {
                btn.innerText = "✨ XỬ LÝ ĐỀ THI";
                btn.disabled = false;
            }
        }
    </script>
</body>
</html>
"""

# --- CÁC HÀM XỬ LÝ LOGIC (Đã đổi đường dẫn động theo file_path) ---
def extract_options_smart(text):
    pattern = r"([A-D]|[a-d])[\.\)\:]\s+"
    parts = re.split(pattern, text)
    options = []
    if len(parts) > 2:
        for i in range(1, len(parts), 2):
            if i + 1 < len(parts):
                options.append({"label": parts[i].upper(), "content": parts[i+1].strip()})
    return options

def parse_answer_zone(paragraphs):
    answer_map, explanation_map = {}, {}
    current_id, current_explanation, is_collecting_explanation = None, [], False
    p_q_id = re.compile(r"^(Câu\s+(\d+))[:\.]?", re.IGNORECASE)
    p_ans_line = re.compile(r"[:\.\·\-\s]*Đáp án\s*[:\.]?\s*(.*)", re.IGNORECASE)
    p_explain_start = re.compile(r"^(Giải thích|Hướng dẫn|Lời giải)[:\.]?", re.IGNORECASE)

    for para in paragraphs:
        text = para.text.strip()
        if not text: continue
        match_id = p_q_id.match(text)
        if match_id:
            if current_id and current_explanation:
                explanation_map[current_id] = "\n".join(current_explanation).strip()
                current_explanation = []
            current_id = match_id.group(2)
            is_collecting_explanation = False
            if "Đáp án" in text: pass
            else: continue
        if current_id:
            match_ans = p_ans_line.search(text)
            if match_ans:
                ans_content = match_ans.group(1).strip()
                checkbox_matches = re.findall(r"([A-Da-d])[\)\.\:]\s*(?:Đ|TRUE|ĐÚNG|S|FALSE|SAI)", ans_content, re.IGNORECASE)
                if checkbox_matches:
                    mapping = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'a': 1, 'b': 2, 'c': 3, 'd': 4}
                    true_matches = re.findall(r"([A-Da-d])[\)\.\:]\s*(?:Đ|TRUE|ĐÚNG)", ans_content, re.IGNORECASE)
                    true_indices = [mapping[char.upper()] for char in true_matches if char.upper() in mapping]
                    answer_map[current_id] = sorted(list(set(true_indices)))
                else:
                    mc_match = re.search(r"\b([A-D])\b", ans_content.upper())
                    if mc_match: answer_map[current_id] = [{'A': 1, 'B': 2, 'C': 3, 'D': 4}.get(mc_match.group(1), 1)]
            if p_explain_start.match(text):
                is_collecting_explanation = True
                content = re.sub(r"^(Giải thích|Hướng dẫn|Lời giải)[:\.]?\s*", "", text, flags=re.IGNORECASE)
                if content: current_explanation.append(content)
            elif is_collecting_explanation and not match_ans:
                current_explanation.append(text)
    if current_id and current_explanation:
        explanation_map[current_id] = "\n".join(current_explanation).strip()
    return answer_map, explanation_map

def parse_docx_split_mode(doc):
    all_paras = doc.paragraphs
    split_index = -1
    for i, para in enumerate(all_paras):
        if any(kw in para.text.strip().upper() for kw in ["ĐÁP ÁN VÀ GIẢI THÍCH", "HƯỚNG DẪN CHẤM", "PHẦN ĐÁP ÁN"]):
            split_index = i
            break
    question_paras, answer_paras = (all_paras[:split_index], all_paras[split_index:]) if split_index != -1 else (all_paras, [])
    key_map, explain_map = parse_answer_zone(answer_paras)
    questions, current_q, current_part = [], {}, 1
    p_header, p_q_start = re.compile(r"^PHẦN\s+(\d+)", re.IGNORECASE), re.compile(r"^(Câu\s+(\d+))[:\.]?\s*(.*)", re.IGNORECASE)
    
    def save_q():
        nonlocal current_q
        if current_q:
            q_id = current_q["id"]
            current_q["correct_indices"] = key_map.get(q_id, [])
            current_q["explanation"] = explain_map.get(q_id, "")
            questions.append(current_q)
        current_q = {}

    for para in question_paras:
        text = para.text.strip()
        if not text: continue
        m_header = p_header.match(text)
        if m_header: save_q(); current_part = int(m_header.group(1)); continue
        m_q = p_q_start.match(text)
        if m_q:
            save_q()
            current_q = {
                "id": m_q.group(2), "text": m_q.group(3).strip(), 
                "type": "Checkbox" if current_part == 2 else ("Open-Ended" if current_part == 3 else "Multiple Choice"),
                "time": 60 if current_part == 2 else (300 if current_part == 3 else 30),
                "options": [], "correct_indices": [], "explanation": ""
            }
            continue
        if not current_q: continue
        if current_part != 3:
            opts_inline = extract_options_smart(text)
            if opts_inline: current_q["options"].extend(opts_inline)
            else:
                m_opt = re.match(r"^([A-D]|[a-d])[\.\)\:]\s*(.*)", text)
                if m_opt: current_q["options"].append({"label": m_opt.group(1).upper(), "content": m_opt.group(2).strip()})
                else: 
                    if len(current_q["options"]) == 0: current_q["text"] += "\n" + text
        else: current_q["text"] += "\n" + text
    save_q()
    return questions

def export_excel(questions, file_path):
    rows = []
    columns = ["Question Text", "Question Type", "Option 1", "Option 2", "Option 3", "Option 4", "Option 5", "Correct Answer", "Time in seconds", "Explanation"]
    for q in questions:
        row = {col: "" for col in columns}
        row["Question Text"], row["Question Type"], row["Time in seconds"], row["Explanation"] = q["text"], q["type"], q["time"], q["explanation"]
        for i, opt in enumerate(q["options"][:5]): row[f"Option {i+1}"] = opt["content"]
        if q["type"] != "Open-Ended": row["Correct Answer"] = ",".join(map(str, q["correct_indices"])) if q["correct_indices"] else "1"
        rows.append(row)
    df = pd.DataFrame(rows, columns=columns)
    df.to_excel(file_path, index=False)
    return file_path

def export_word_bold(questions, file_path):
    doc = Document()
    doc.styles['Normal'].font.name, doc.styles['Normal'].font.size = 'Times New Roman', Pt(12)
    doc.add_heading('ĐỀ THI (ĐÁP ÁN IN ĐẬM ĐỎ)', level=0).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for q in questions:
        p = doc.add_paragraph()
        p.add_run(f"Câu {q['id']}: ").bold = True
        p.add_run(q['text'])
        if q['options']:
            for i, opt in enumerate(q['options']):
                is_correct = (i + 1) in q['correct_indices']
                p_opt = doc.add_paragraph()
                p_opt.paragraph_format.left_indent = Pt(18)
                run_l = p_opt.add_run(f"{chr(97+i)+')' if q['type']=='Checkbox' else opt['label']+'.'} ")
                run_c = p_opt.add_run(f"{opt['content']}")
                if is_correct:
                    run_l.bold = run_c.bold = True
                    run_l.font.color.rgb = run_c.font.color.rgb = RGBColor(255, 0, 0)
    doc.save(file_path)
    return file_path

def export_word_standard(questions, file_path):
    doc = Document()
    doc.styles['Normal'].font.name, doc.styles['Normal'].font.size = 'Times New Roman', Pt(12)
    doc.add_heading('ĐỀ KIỂM TRA', level=0).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for q in questions:
        p = doc.add_paragraph()
        p.add_run(f"Câu {q['id']}: ").bold = True
        p.add_run(q['text'])
        for i, opt in enumerate(q['options']):
            p_o = doc.add_paragraph()
            p_o.paragraph_format.left_indent = Pt(24)
            p_o.add_run(f"{chr(97+i)+')' if q['type']=='Checkbox' else opt['label']+'.'} {opt['content']}")
    doc.add_page_break()
    doc.add_heading('BẢNG ĐÁP ÁN', level=1).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    if len(questions) > 0:
        t = doc.add_table(rows=((len(questions)//5)+(1 if len(questions)%5>0 else 0))*2, cols=5)
        t.style = 'Table Grid'
        for idx, q in enumerate(questions):
            r, c = (idx // 5) * 2, idx % 5
            t.cell(r, c).text = str(q['id'])
            t.cell(r, c).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            ans_cell = t.cell(r+1, c)
            ans_text = ",".join([chr(96+i) for i in q['correct_indices']]) if q['type']=="Checkbox" else (['A','B','C','D'][q['correct_indices'][0]-1] if q['correct_indices'] else "")
            ans_cell.text = ans_text
            ans_cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            if ans_text:
                run = ans_cell.paragraphs[0].runs[0]
                run.bold = True; run.font.color.rgb = RGBColor(255, 0, 0)
    doc.add_page_break()
    doc.add_heading('LỜI GIẢI CHI TIẾT', level=1).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for q in questions:
        if q.get('explanation'):
            p = doc.add_paragraph()
            run = p.add_run(f"Câu {q['id']}: ")
            run.bold = run.italic = True; run.font.color.rgb = RGBColor(0,0,255)
            p.add_run(q['explanation'])
    doc.save(file_path)
    return file_path

# --- ENDPOINTS XỬ LÝ ---
@app.get("/", response_class=HTMLResponse)
async def serve_ui():
    return HTML_TEMPLATE

@app.post("/api/process")
async def process_exam_text(request: Request):
    try:
        data = await request.json()
        raw_text = data.get("text", "")
        if not raw_text: return {"error": "Nội dung trống!"}
        
        doc = Document()
        for line in raw_text.split('\n'):
            if line.strip(): doc.add_paragraph(line.strip())
            
        questions = parse_docx_split_mode(doc)
        
        # Tạo chuỗi ID ngẫu nhiên để không bị trùng file giữa các lượt dùng
        uid = uuid.uuid4().hex[:8]
        
        excel_name = f"Quizizz_{uid}.xlsx"
        word_bold_name = f"WordBold_{uid}.docx"
        word_std_name = f"WordStd_{uid}.docx"
        
        # Lưu file vật lý vào thư mục tạm
        export_excel(questions, os.path.join(BASE_DIR, excel_name))
        export_word_bold(questions, os.path.join(BASE_DIR, word_bold_name))
        export_word_standard(questions, os.path.join(BASE_DIR, word_std_name))
        
        # Trả về cục JSON chứa 3 đường dẫn endpoint để tải file
        return {
            "excel_url": f"/api/download/{excel_name}",
            "word_bold_url": f"/api/download/{word_bold_name}",
            "word_std_url": f"/api/download/{word_std_name}"
        }
    except Exception as e:
        return {"error": str(e)}

# Endpoint phụ trách việc phân phối tải file xuống khi click nút trên giao diện
@app.get("/api/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(BASE_DIR, filename)
    if os.path.exists(file_path):
        # Đặt lại tên hiển thị mặc định lúc tải xuống cho đẹp mắt
        friendly_name = "Quizizz_Import_File.xlsx"
        if "WordBold" in filename: friendly_name = "De_Thi_Bold_Dap_An.docx"
        elif "WordStd" in filename: friendly_name = "De_Thi_Chuan_Format_Full.docx"
        
        return FileResponse(file_path, media_type="application/octet-stream", filename=friendly_name)
    return {"error": "File không tồn tại hoặc đã bị xóa"}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
