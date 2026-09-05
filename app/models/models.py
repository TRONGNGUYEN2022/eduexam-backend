from sqlalchemy import Column, Integer, String, Text, Boolean, ForeignKey, Float, DateTime, Enum
from sqlalchemy.orm import relationship
import datetime
import enum
from app.core.database import Base

class CognitiveLevelEnum(str, enum.Enum):
    nhan_biet = "nhan_biet"
    thong_hieu = "thong_hieu"
    van_dung = "van_dung"
    van_dung_cao = "van_dung_cao"

class QuestionTypeEnum(str, enum.Enum):
    multiple_choice = "multiple_choice"
    true_false = "true_false"
    short_answer = "short_answer"
    essay = "essay"

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(255), unique=True, nullable=False)
    hashed_password = Column(String(255), nullable=False)
    full_name = Column(String(100))
    school_name = Column(String(255))
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    questions = relationship("Question", back_populates="creator")
    exams = relationship("Exam", back_populates="creator")

class Question(Base):
    __tablename__ = "questions"
    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=True)
    subject = Column(String(100), nullable=False)
    grade = Column(Integer, nullable=False)
    topic = Column(String(255), nullable=False)
    cognitive_level = Column(Enum(CognitiveLevelEnum), nullable=False)
    question_type = Column(Enum(QuestionTypeEnum), nullable=False)
    content = Column(Text, nullable=False)
    solution = Column(Text, nullable=True)
    is_public = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    creator = relationship("User", back_populates="questions")
    options = relationship("QuestionOption", back_populates="question", cascade="all, delete-orphan")
    exam_associations = relationship("ExamQuestion", back_populates="question")

class QuestionOption(Base):
    __tablename__ = "question_options"
    id = Column(Integer, primary_key=True, index=True)
    question_id = Column(Integer, ForeignKey("questions.id", ondelete="CASCADE"))
    option_label = Column(String(5), nullable=False)
    content = Column(Text, nullable=False)
    is_correct = Column(Boolean, default=False)

    question = relationship("Question", back_populates="options")

class Exam(Base):
    __tablename__ = "exams"
    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=True)
    title = Column(String(255), nullable=False)
    subject = Column(String(100), nullable=False)
    grade = Column(Integer, nullable=False)
    duration_minutes = Column(Integer, default=60)
    total_score = Column(Float, default=10.0)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

    creator = relationship("User", back_populates="exams")
    exam_questions = relationship("ExamQuestion", back_populates="exam", cascade="all, delete-orphan")

class ExamQuestion(Base):
    __tablename__ = "exam_questions"
    id = Column(Integer, primary_key=True, index=True)
    exam_id = Column(Integer, ForeignKey("exams.id", ondelete="CASCADE"))
    question_id = Column(Integer, ForeignKey("questions.id"))
    order_index = Column(Integer, nullable=False)
    score = Column(Float, default=0.25)

    exam = relationship("Exam", back_populates="exam_questions")
    question = relationship("Question", back_populates="exam_associations")