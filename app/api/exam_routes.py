import os
import tempfile
import pypandoc
from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional, Dict, Any, List

router = APIRouter(prefix="/api/exams", tags=["Exams"])


# --- Schemas ---

class AnswerExportRequest(BaseModel):
    subject: str
    grade: int
    exam_title: str
    school_name: Optional[str] = "TRƯỜNG THCS LÊ QUÝ ĐÔN"
    part_2_tf: Optional[List[Dict[str, Any]]] = []
    part_3_essay: Optional[List[Dict[str, Any]]] = []
    answer_matrix: Optional[Dict[str, Any]] = {}


class FullExamDocxRequest(BaseModel):
    markdown_content: str
    subject: Optional[str] = "Toan_Hoc"
    grade: Optional[int] = 6


# --- Helper Functions ---

def build_answers_markdown(req: AnswerExportRequest) -> str:
    md = [
        f"# {req.school_name.upper()}\n",
        f"## HƯỚNG DẪN CHẤM & ĐÁP ÁN CHI TIẾT - {req.exam_title.upper()}",
        f"**Môn:** {req.subject.upper()} {req.grade}\n",
        "---\n"
    ]

    # 1. Bảng ma trận 4 mã đề
    if req.answer_matrix:
        md.append("### I. BẢNG ĐÁP ÁN TRẮC NGHIỆM 4 MÃ ĐỀ ĐỐI CHIẾU\n")
        md.append("| Câu hỏi | Mã 101 | Mã 102 | Mã 103 | Mã 104 |")
        md.append("| :---: | :---: | :---: | :---: | :---: |")
        total_q = max(len(req.answer_matrix.get('101', {})), 4)
        for i in range(1, total_q + 1):
            k = f"Câu {i}"
            c101 = req.answer_matrix.get('101', {}).get(k, '-')
            c102 = req.answer_matrix.get('102', {}).get(k, '-')
            c103 = req.answer_matrix.get('103', {}).get(k, '-')
            c104 = req.answer_matrix.get('104', {}).get(k, '-')
            md.append(f"| {k} | {c101} | {c102} | {c103} | {c104} |")
        md.append("\n")

    # 2. Đáp án Đúng / Sai
    if req.part_2_tf:
        md.append("### II. ĐÁP ÁN PHẦN II (ĐÚNG / SAI)\n")
        for idx, q in enumerate(req.part_2_tf, 1):
            ans = q.get('answers', {})
            a_str = "ĐÚNG" if str(ans.get('a', '')).lower() in ['đ', 'true', 'd'] else "SAI"
            b_str = "ĐÚNG" if str(ans.get('b', '')).lower() in ['đ', 'true', 'd'] else "SAI"
            c_str = "ĐÚNG" if str(ans.get('c', '')).lower() in ['đ', 'true', 'd'] else "SAI"
            d_str = "ĐÚNG" if str(ans.get('d', '')).lower() in ['đ', 'true', 'd'] else "SAI"

            md.append(f"**Câu {idx}:** a) **{a_str}** | b) **{b_str}** | c) **{c_str}** | d) **{d_str}**\n")
            if q.get('explanation'):
                md.append(f"- *Giải thích:* {q.get('explanation')}\n")

    # 3. Lời giải tự luận & Barem điểm
    if req.part_3_essay:
        md.append("### III. HƯỚNG DẪN CHẤM & BAREM TỰ LUẬN\n")
        for idx, q in enumerate(req.part_3_essay, 1):
            pts = q.get('points', 1.5)
            sol = q.get('solution') or q.get('answer', '')
            md.append(f"**Bài {idx} ({pts} điểm):**\n")
            md.append(f"- **Lời giải chi tiết:**\n\n{sol}\n")

            guides = q.get('scoring_guide', q.get('guide', []))
            if guides:
                md.append("- **Barem thang điểm:**")
                for g in guides:
                    step_text = g.get('step', g.get('content', ''))
                    score = g.get('points', g.get('score', '0.5 đ'))
                    md.append(f"  - {step_text} (**{score}**)")
            md.append("\n")

    return "\n".join(md)


# --- Endpoints ---

@router.post("/export-answers-pdf")
def export_answers_pdf(req: AnswerExportRequest):
    """Biên dịch riêng phần Hướng dẫn chấm sang PDF chuẩn in ấn qua Pandoc."""
    try:
        full_markdown = build_answers_markdown(req)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            pdf_path = tmp_pdf.name

        pypandoc.convert_text(
            source=full_markdown,
            to='pdf',
            format='markdown+tex_math_dollars+raw_tex',
            outputfile=pdf_path,
            extra_args=[
                '--pdf-engine=xelatex',
                '-V', 'geometry:margin=2cm',
                '-V', 'mainfont=DejaVu Serif'
            ]
        )

        filename = f"Dap_An_Barem_{req.subject}_Lop_{req.grade}.pdf"
        return FileResponse(
            pdf_path,
            filename=filename,
            media_type="application/pdf"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi tạo PDF: {str(e)}")


@router.post("/export-answers-docx")
def export_answers_docx(req: AnswerExportRequest):
    """Biên dịch riêng phần Hướng dẫn chấm sang Word .docx qua Pandoc (Native Word Equation)."""
    try:
        full_markdown = build_answers_markdown(req)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            output_path = tmp_docx.name

        pypandoc.convert_text(
            source=full_markdown,
            to='docx',
            format='markdown+tex_math_dollars+raw_tex',
            outputfile=output_path,
            extra_args=['--mathml']
        )

        filename = f"Dap_An_Barem_{req.subject}_Lop_{req.grade}.docx"
        return FileResponse(
            output_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi tạo Word DOCX: {str(e)}")


@router.post("/export-full-docx")
def export_full_docx(req: FullExamDocxRequest):
    """Biên dịch toàn bộ văn bản Markdown đề thi sang Word .docx qua Pandoc."""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            output_path = tmp_docx.name

        pypandoc.convert_text(
            source=req.markdown_content,
            to='docx',
            format='markdown+tex_math_dollars+raw_tex',
            outputfile=output_path,
            extra_args=['--mathml']
        )

        filename = f"De_Thi_{req.subject}_Lop_{req.grade}.docx"
        return FileResponse(
            output_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xuất đề thi DOCX: {str(e)}")