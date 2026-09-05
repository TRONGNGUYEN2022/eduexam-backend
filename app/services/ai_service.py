import os
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from typing import List, Literal, Optional

client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

# Định nghĩa Schema dữ liệu đầu ra từ AI
class OptionSchema(BaseModel):
    id: Literal["A", "B", "C", "D"]
    text: str
    is_correct: bool

class QuestionSchema(BaseModel):
    subject: str
    grade: int
    topic: str
    cognitive_level: Literal["nhan_biet", "thong_hieu", "van_dung", "van_dung_cao"]
    question_type: Literal["multiple_choice", "true_false", "short_answer", "essay"]
    content: str
    options: Optional[List[OptionSchema]] = None
    solution: str

class ExamMatrixOutput(BaseModel):
    exam_title: str
    duration_minutes: int
    part_1_mcq: List[QuestionSchema]
    part_2_true_false: List[QuestionSchema]
    part_3_essay: List[QuestionSchema]

def generate_exam_with_gemini(subject: str, grade: int, topic: str, duration: int) -> ExamMatrixOutput:
    prompt = f"""
    Bạn là chuyên gia khảo thí THCS tại Việt Nam. 
    Hãy tạo một đề kiểm tra chuẩn chương trình GDPT 2018 với:
    - Môn: {subject}
    - Lớp: {grade}
    - Chủ đề: {topic}
    - Thời gian làm bài: {duration} phút

    Cấu trúc bắt buộc:
    1. Phần I: 4 câu trắc nghiệm nhiều lựa chọn (MCQ), 4 đáp án A, B, C, D (mức độ Nhận biết, Thông hiểu).
    2. Phần II: 1 câu Đúng/Sai gồm 4 ý a, b, c, d.
    3. Phần III: 2 câu Tự luận (Vận dụng, Vận dụng cao) có lời giải chi tiết.
    4. Toàn bộ công thức Toán/KHTN đặt trong cặp dấu $...$.
    """

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ExamMatrixOutput,
            temperature=0.3,
        ),
    )
    return response.parsed