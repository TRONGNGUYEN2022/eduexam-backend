# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import io
import json
import re
import base64
import zipfile
import sqlite3
import xml.etree.ElementTree as ET
from typing import Optional, List, Dict, Any, Tuple, Union

from fastapi import FastAPI, HTTPException, Request, UploadFile, File, Form, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import latex2mathml.converter
from lxml import etree

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

app = FastAPI(title="EduExam Generic Deterministic Engine", version="75.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DATA_DIR = "/var/data" if os.path.exists("/var/data") else "."
DB_FILE = os.path.join(DATA_DIR, "exam_bank.db")
DB_DRIVE_FILE_NAME = "exam_bank.db"

GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_ROOT_FOLDER_ID", "")
SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
SERVICE_ACCOUNT_INFO_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")

# XSLT CHUẨN ISO/IEC 29500 MML2OMML
MML2OMML_XSL = b"""<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <xsl:output method="xml" encoding="UTF-8" indent="no"/>

    <xsl:template match="/*">
        <m:oMath>
            <xsl:apply-templates select="*"/>
        </m:oMath>
    </xsl:template>

    <xsl:template match="*[local-name()='mi' or local-name()='mn' or local-name()='mo' or local-name()='mtext']">
        <m:r>
            <m:t><xsl:value-of select="normalize-space(.)"/></m:t>
        </m:r>
    </xsl:template>

    <xsl:template match="*[local-name()='mfrac']">
        <m:f>
            <m:num><xsl:apply-templates select="*[1]"/></m:num>
            <m:den><xsl:apply-templates select="*[2]"/></m:den>
        </m:f>
    </xsl:template>

    <xsl:template match="*[local-name()='msqrt']">
        <m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e><xsl:apply-templates select="*"/></m:e>
        </m:rad>
    </xsl:template>

    <xsl:template match="*[local-name()='mroot']">
        <m:rad>
            <m:radPr><m:degHide m:val="0"/></m:radPr>
            <m:deg><xsl:apply-templates select="*[2]"/></m:deg>
            <m:e><xsl:apply-templates select="*[1]"/></m:e>
        </m:rad>
    </xsl:template>

    <xsl:template match="*[local-name()='msup']">
        <m:sSup>
            <m:e><xsl:apply-templates select="*[1]"/></m:e>
            <m:sup><xsl:apply-templates select="*[2]"/></m:sup>
        </m:sSup>
    </xsl:template>

    <xsl:template match="*[local-name()='msub']">
        <m:sSub>
            <m:e><xsl:apply-templates select="*[1]"/></m:e>
            <m:sub><xsl:apply-templates select="*[2]"/></m:sub>
        </m:sSub>
    </xsl:template>

    <xsl:template match="*[local-name()='msubsup']">
        <m:sSubSup>
            <m:e><xsl:apply-templates select="*[1]"/></m:e>
            <m:sub><xsl:apply-templates select="*[2]"/></m:sub>
            <m:sup><xsl:apply-templates select="*[3]"/></m:sup>
        </m:sSubSup>
    </xsl:template>

    <xsl:template match="*[local-name()='mfenced']">
        <m:d>
            <m:dPr>
                <m:begChr m:val="&#40;"/>
                <m:endChr m:val="&#41;"/>
            </m:dPr>
            <m:e><xsl:apply-templates select="*"/></m:e>
        </m:d>
    </xsl:template>

    <xsl:template match="*[local-name()='mtable']">
        <m:m>
            <xsl:apply-templates select="*"/>
        </m:m>
    </xsl:template>

    <xsl:template match="*[local-name()='mtr']">
        <m:mr>
            <xsl:apply-templates select="*"/>
        </m:mr>
    </xsl:template>

    <xsl:template match="*[local-name()='mtd']">
        <m:e>
            <xsl:apply-templates select="*"/>
        </m:e>
    </xsl:template>

    <xsl:template match="*">
        <xsl:apply-templates select="*"/>
    </xsl:template>
</xsl:stylesheet>
"""

try:
    _xslt_parsed = etree.parse(io.BytesIO(MML2OMML_XSL))
    MATH_XSLT = etree.XSLT(_xslt_parsed)
except Exception:
    MATH_XSLT = None

class GenerateExamRequest(BaseModel):
    level: str = "thpt"
    subject: str = "Toán học"
    grade: int = 12
    topic: str = "Chung"
    lesson: Optional[str] = "Chung"
    cognitive_level: Optional[str] = "M2"
    duration_minutes: int = 90
    exam_title: str = "ĐỀ KIỂM TRA ĐỊNH KỲ"
    include_answers: Optional[bool] = False
    selected_question_ids: Optional[List[int]] = []

class QuestionSave(BaseModel):
    level_stage: str = "thpt"
    grade: int
    subject: str = "Toán học"
    topic: str
    lesson: Optional[str] = "Chung"
    level: str = "Thông hiểu"
    cognitive_level: Optional[str] = "M2"
    q_type: str
    content: str
    options: Optional[List[str]] = []
    correct_answer: Optional[str] = ""
    explanation: Optional[str] = ""
    has_image: Optional[int] = 0
    image_base64: Optional[str] = ""
    tikz_code: Optional[str] = ""

class BulkSaveRequest(BaseModel):
    questions: List[QuestionSave]

def extract_clean_folder_id(raw_id_or_url: str) -> str:
    if not raw_id_or_url:
        return ""
    text = raw_id_or_url.strip()
    match = re.search(r'folders/([a-zA-Z0-9_-]+)', text)
    return match.group(1) if match else text

def get_drive_service() -> Tuple[Any, str]:
    scopes = ['https://www.googleapis.com/auth/drive']
    try:
        if SERVICE_ACCOUNT_INFO_JSON and SERVICE_ACCOUNT_INFO_JSON.strip():
            raw_json = SERVICE_ACCOUNT_INFO_JSON.strip()
            if (raw_json.startswith("'") and raw_json.endswith("'")) or (raw_json.startswith('"') and raw_json.endswith('"')):
                raw_json = raw_json[1:-1]
            try:
                info = json.loads(raw_json)
            except Exception:
                info = json.loads(raw_json.encode('utf-8').decode('unicode_escape'))
            creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
        elif os.path.exists(SERVICE_ACCOUNT_FILE):
            creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
        else:
            return None, "Thiếu cấu hình Service Account trên máy chủ."
        service = build('drive', 'v3', credentials=creds, cache_discovery=False)
        return service, ""
    except Exception as e:
        return None, f"Lỗi xác thực Service Account: {str(e)}"

def find_target_db_file(service, folder_id: str) -> Optional[str]:
    queries = []
    if folder_id and " " not in folder_id:
        queries.append(f"name = '{DB_DRIVE_FILE_NAME}' and '{folder_id}' in parents and trashed = false")
    queries.append(f"name = '{DB_DRIVE_FILE_NAME}' and trashed = false")

    for q in queries:
        try:
            res = service.files().list(
                q=q, spaces='drive', fields="files(id, name, parents)",
                supportsAllDrives=True, includeItemsFromAllDrives=True
            ).execute()
            files = res.get('files', [])
            if files:
                return files[0]['id']
        except Exception:
            continue
    return None

def backup_db_to_drive() -> Tuple[bool, str]:
    service, auth_err = get_drive_service()
    if not service:
        return False, auth_err
    if not os.path.exists(DB_FILE):
        return False, f"Không tìm thấy file SQLite ({DB_FILE})."

    try:
        clean_folder_id = extract_clean_folder_id(GOOGLE_DRIVE_FOLDER_ID)
        target_file_id = find_target_db_file(service, clean_folder_id)
        media_upload = MediaFileUpload(DB_FILE, mimetype='application/x-sqlite3', resumable=True)

        if target_file_id:
            service.files().update(fileId=target_file_id, media_body=media_upload, supportsAllDrives=True).execute()
            return True, f"Sao lưu thành công! Đã ghi đè vào file 'exam_bank.db' (ID: {target_file_id})."
        else:
            metadata = {'name': DB_DRIVE_FILE_NAME}
            if clean_folder_id and " " not in clean_folder_id:
                metadata['parents'] = [clean_folder_id]
            created = service.files().create(body=metadata, media_body=media_upload, fields='id', supportsAllDrives=True).execute()
            return True, f"Tạo mới và sao lưu thành công (ID: {created.get('id')})!"
    except Exception as e:
        return False, f"Lỗi Google Drive: {str(e)}"

def restore_db_from_drive() -> Tuple[bool, str]:
    service, auth_err = get_drive_service()
    if not service:
        return False, auth_err

    try:
        clean_folder_id = extract_clean_folder_id(GOOGLE_DRIVE_FOLDER_ID)
        target_file_id = find_target_db_file(service, clean_folder_id)
        if not target_file_id:
            return False, "Không tìm thấy tệp 'exam_bank.db' trên Google Drive để phục hồi."

        request = service.files().get_media(fileId=target_file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        with open(DB_FILE, "wb") as f:
            f.write(fh.read())
        return True, "Khôi phục cơ sở dữ liệu từ Google Drive thành công!"
    except Exception as e:
        return False, f"Lỗi khôi phục: {str(e)}"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            level_stage TEXT DEFAULT 'thpt',
            grade INTEGER NOT NULL,
            subject TEXT DEFAULT 'Toán học',
            topic TEXT NOT NULL,
            lesson TEXT DEFAULT 'Chung',
            level TEXT NOT NULL,
            cognitive_level TEXT DEFAULT 'M2',
            q_type TEXT NOT NULL,
            content TEXT NOT NULL,
            options_json TEXT,
            correct_answer TEXT,
            explanation TEXT,
            has_image INTEGER DEFAULT 0,
            image_base64 TEXT,
            tikz_code TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

init_db()

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def generic_clean_latex(text: str) -> str:
    if not text:
        return ""
    s = str(text)
    s = re.sub(r'^(?:#+\s*|\*\*\s*)(?:Câu|Bài|Question)\s*\d+[\s.:*]*', '', s, flags=re.MULTILINE | re.IGNORECASE)
    s = re.sub(r'\^{2,}', '^', s)
    s = re.sub(r'\$\$', '$', s)
    s = re.sub(r'\$\s*\$', '', s)
    return s.strip()

def clean_generic_ocr_question(raw_block: str) -> tuple[str, str, List[str]]:
    if not raw_block:
        return "", "essay", []

    text = raw_block.strip()
    text = re.sub(r'\s*(?:Trang\s*\d+\/\d+|Mã\s*đề(?:\s*thi)?\s*[\d\w]+|\*\*PHẦN\s*[I|V|X]+[\s\S]*).*$', '', text, flags=re.IGNORECASE).strip()

    opt_pattern = r'(?:^|\n|\s{2,})([A-D])\s*[\.,:]\s*(.*?)(?=(?:(?:^|\n|\s{2,})[A-D]\s*[\.,:]\s*)|$)'
    matches = list(re.finditer(opt_pattern, text, flags=re.DOTALL))

    if len(matches) >= 3:
        first_opt_idx = matches[0].start()
        body = text[:first_opt_idx].strip()
        options = []
        for m in matches:
            val = re.sub(r'\s+', ' ', m.group(2)).strip()
            options.append(val)
        return body, "single_choice", options

    if re.search(r'(?:^|\n|\s+)-\s*[a-d]\)', text) or "chọn đúng hoặc sai" in text.lower():
        return text, "true_false", []

    return text, "essay", []

def parse_mistral_zip_tree(zip_bytes: bytes) -> Tuple[str, Dict[str, str]]:
    image_assets: Dict[str, str] = {}
    page_markdown_map: Dict[int, str] = {}
    root_markdown_text = ""

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        namelist = zf.namelist()

        for name in namelist:
            lower = name.lower()
            if lower.endswith(('.jpg', '.jpeg', '.png', '.webp', '.gif')):
                base_name = os.path.basename(name)
                img_bytes = zf.read(name)
                mime = "image/png" if lower.endswith('.png') else "image/jpeg"
                b64_str = f"data:{mime};base64,{base64.b64encode(img_bytes).decode('utf-8')}"
                image_assets[base_name] = b64_str
                image_assets[name] = b64_str

        root_md_candidates = [n for n in namelist if n.lower().endswith('.md') and 'pages/' not in n.lower()]
        if root_md_candidates:
            best_root_md = next((n for n in root_md_candidates if os.path.basename(n).lower() == 'markdown.md'), root_md_candidates[0])
            root_markdown_text = zf.read(best_root_md).decode('utf-8', errors='ignore').strip()

        for name in namelist:
            if name.lower().endswith('.md') and 'pages/' in name.lower():
                page_match = re.search(r'page-(\d+)', name, flags=re.IGNORECASE)
                if page_match:
                    page_num = int(page_match.group(1))
                    page_markdown_map[page_num] = zf.read(name).decode('utf-8', errors='ignore').strip()

    final_markdown = root_markdown_text if root_markdown_text else "\n\n".join(page_markdown_map[p] for p in sorted(page_markdown_map.keys()))
    return final_markdown, image_assets

def parse_mistral_zip_stream(content_bytes: bytes, default_grade=12, default_stage="thpt", default_subject="Toán học", default_topic="Chung") -> List[Dict[str, Any]]:
    full_md, images = parse_mistral_zip_tree(content_bytes)

    split_match = re.search(r'(?:^|\n)\s*(?:#+\s*)?(?:HƯỚNG\s+DẪN\s+GIẢI|HƯỚNG\s+DẪN\s+CHẤM|ĐÁP\s+ÁN|LỜI\s+GIẢI)', full_md, flags=re.IGNORECASE)
    exam_raw = full_md[:split_match.start()] if split_match else full_md
    guide_raw = full_md[split_match.start():] if split_match else ""

    guide_dict: Dict[int, Tuple[str, str]] = {}
    if guide_raw:
        g_matches = list(re.finditer(r'(?:^|\n)\s*(?:\|\s*)?(?:#+\s*)?(?:\*\*)?(?:Bài|Câu|Question)\s*(\d+)[\s.:-]*', guide_raw, flags=re.IGNORECASE))
        for idx, gm in enumerate(g_matches):
            q_num = int(gm.group(1))
            start_pos = gm.start()
            end_pos = g_matches[idx + 1].start() if idx + 1 < len(g_matches) else len(guide_raw)
            raw_expl = guide_raw[start_pos:end_pos].strip()

            expl_img = ""
            inline_imgs = re.findall(r'!\[.*?\]\((.*?)\)', raw_expl)
            for img_p in inline_imgs:
                b_name = os.path.basename(img_p.strip())
                if b_name in images:
                    expl_img = images[b_name]
                    break

            clean_expl = re.sub(r'^(?:\|\s*)?(?:#+\s*)?(?:\*\*)?(?:Bài|Câu|Question)\s*\d+[\s.:-]*', '', raw_expl, flags=re.IGNORECASE).strip()
            clean_expl = re.sub(r'!\[.*?\](?:\(.*?\))?', '', clean_expl).strip()
            guide_dict[q_num] = (clean_expl, expl_img)

    q_matches = list(re.finditer(r'(?:^|\n)\s*(?:#+\s*)?(?:\*\*)?(?:Bài|Câu|Question)\s*(\d+)[\s.:-]*', exam_raw, flags=re.IGNORECASE))
    parsed_questions = []

    for idx, qm in enumerate(q_matches):
        q_num = int(qm.group(1))
        start_pos = qm.start()
        end_pos = q_matches[idx + 1].start() if idx + 1 < len(q_matches) else len(exam_raw)
        q_block = exam_raw[start_pos:end_pos].strip()

        # THUẬT TOÁN RÀNG BUỘC BIÊN KHỐI: CHỈ LẤY ẢNH NẰM TRONG ĐÚNG TEXT BLOCK CỦA CÂU NÀY
        primary_img = ""
        inline_imgs = re.findall(r'!\[.*?\]\((.*?)\)', q_block)
        for img_p in inline_imgs:
            b_name = os.path.basename(img_p.strip())
            if b_name in images:
                primary_img = images[b_name]
                break

        q_clean = re.sub(r'^(?:#+\s*)?(?:\*\*)?(?:Bài|Câu|Question)\s*\d+[\s.:-]*', '', q_block, flags=re.IGNORECASE).strip()
        q_clean = re.sub(r'!\[.*?\](?:\(.*?\))?', '', q_clean).strip()

        expl_text, expl_img = guide_dict.get(q_num, ("", ""))
        if not primary_img and expl_img:
            primary_img = expl_img

        clean_body, q_type, options = clean_generic_ocr_question(q_clean)
        
        parsed_questions.append({
            "grade": default_grade,
            "level_stage": default_stage,
            "subject": default_subject,
            "lesson": default_topic,
            "topic": default_topic,
            "cognitive_level": "M2" if q_num <= 4 else "M3",
            "q_type": q_type,
            "content": generic_clean_latex(clean_body),
            "options": [generic_clean_latex(o) for o in options],
            "correct_answer": "",
            "explanation": generic_clean_latex(expl_text),
            "has_image": 1 if primary_img else 0,
            "image_base64": primary_img,
            "tikz_code": ""
        })

    return parsed_questions

def set_cell_shading(cell, color_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="64748B", sz="6", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def mathml_element_to_omml_xml(node: ET.Element) -> str:
    tag = node.tag.split('}')[-1] if '}' in node.tag else node.tag
    text = (node.text or '').strip()

    if tag in ('mi', 'mn', 'mo', 'mtext'):
        return f'<m:r><m:t>{text}</m:t></m:r>'
    elif tag == 'mfrac':
        children = list(node)
        num = "".join(mathml_element_to_omml_xml(c) for c in (children[0] if len(children) > 0 else []))
        den = "".join(mathml_element_to_omml_xml(c) for c in (children[1] if len(children) > 1 else []))
        return f'<m:f><m:num>{num}</m:num><m:den>{den}</m:den></m:f>'
    elif tag in ('msqrt', 'mroot'):
        children = list(node)
        base = "".join(mathml_element_to_omml_xml(c) for c in (children[0] if children else []))
        return f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{base}</m:e></m:rad>'
    elif tag == 'msup':
        children = list(node)
        base = mathml_element_to_omml_xml(children[0]) if len(children) > 0 else ""
        sup = mathml_element_to_omml_xml(children[1]) if len(children) > 1 else ""
        return f'<m:sSup><m:e>{base}</m:e><m:sup>{sup}</m:sup></m:sSup>'
    elif tag == 'msub':
        children = list(node)
        base = mathml_element_to_omml_xml(children[0]) if len(children) > 0 else ""
        sub = mathml_element_to_omml_xml(children[1]) if len(children) > 1 else ""
        return f'<m:sSub><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>'
    elif tag == 'msubsup':
        children = list(node)
        base = mathml_element_to_omml_xml(children[0]) if len(children) > 0 else ""
        sub = mathml_element_to_omml_xml(children[1]) if len(children) > 1 else ""
        sup = mathml_element_to_omml_xml(children[2]) if len(children) > 2 else ""
        return f'<m:sSubSup><m:e>{base}</m:e><m:sub>{sub}</m:sub><m:sup>{sup}</m:sup></m:sSubSup>'
    elif tag == 'mtable':
        rows = "".join(mathml_element_to_omml_xml(c) for c in node)
        return f'<m:m>{rows}</m:m>'
    elif tag == 'mtr':
        cells = "".join(f'<m:e>{mathml_element_to_omml_xml(c)}</m:e>' for c in node)
        return f'<m:mr>{cells}</m:mr>'
    return "".join(mathml_element_to_omml_xml(c) for c in node)

def convert_latex_to_omml_native(latex_code: str):
    clean = latex_code.strip()
    clean = re.sub(r'\\text\s*\{([^}]*)\}', r'\1', clean)
    clean = re.sub(r'\\overline\{([^\}]+)\}', r'\\bar{\1}', clean)
    clean = re.sub(r'\\overrightarrow\{([^\}]+)\}', r'\\vec{\1}', clean)
    
    try:
        mathml = latex2mathml.converter.convert(clean)
        mathml_clean = re.sub(r'\sxmlns="[^"]+"', '', mathml)
        tree = ET.fromstring(mathml_clean.encode('utf-8'))
        inner_omml = mathml_element_to_omml_xml(tree)
        xml_str = f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{inner_omml}</m:oMath>'
        return parse_xml(xml_str)
    except Exception:
        try:
            return parse_xml(f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>{clean}</m:t></m:r></m:oMath>')
        except Exception:
            return None

def write_formatted_text_with_math(paragraph, raw_text: str):
    parts = raw_text.split('$')
    for idx, part in enumerate(parts):
        if not part:
            continue
        if idx % 2 == 1:
            omml = convert_latex_to_omml_native(part)
            if omml is not None:
                paragraph._p.append(omml)
            else:
                run = paragraph.add_run(f" {part} ")
                run.italic = True
        else:
            paragraph.add_run(part)

def render_content_block_to_docx(doc: docx.Document, raw_text: str):
    lines = raw_text.split('\n')
    table_lines = []
    in_table = False

    for line in lines:
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            table_lines.append(stripped)
            continue
        else:
            if in_table and table_lines:
                build_word_table_from_markdown(doc, table_lines)
                table_lines = []
                in_table = False

        if not stripped:
            continue

        sub_items = re.split(r'(?=(?:^|\s+)-\s*[a-d]\))', stripped)
        for item in sub_items:
            item_clean = item.strip()
            if not item_clean:
                continue
            
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            
            m_sub = re.match(r'^(?:-\s*)?([a-d]\))\s*(.*)$', item_clean, flags=re.IGNORECASE)
            if m_sub:
                p.paragraph_format.left_indent = Inches(0.25)
                r_lbl = p.add_run(f"{m_sub.group(1)} ")
                r_lbl.bold = True
                write_formatted_text_with_math(p, m_sub.group(2))
            else:
                write_formatted_text_with_math(p, item_clean)

    if in_table and table_lines:
        build_word_table_from_markdown(doc, table_lines)

def build_word_table_from_markdown(doc: docx.Document, table_lines: List[str]):
    cleaned_rows = []
    for line in table_lines:
        if '---' in line:
            continue
        cells = [c.strip() for c in line.split('|')[1:-1]]
        if cells:
            cleaned_rows.append(cells)

    if not cleaned_rows:
        return

    num_rows = len(cleaned_rows)
    num_cols = max(len(r) for r in cleaned_rows)

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="64748B", sz="6", val="single")

    for r_idx, row in enumerate(cleaned_rows):
        for c_idx, cell_value in enumerate(row):
            if c_idx < num_cols:
                cell = table.cell(r_idx, c_idx)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
                
                if r_idx == 0:
                    set_cell_shading(cell, "F1F5F9")
                
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.space_before = Pt(2)
                
                if r_idx == 0:
                    p.paragraph_format.line_spacing = 1.0
                    r = p.add_run()
                    r.bold = True
                write_formatted_text_with_math(p, cell_value)

    doc.add_paragraph()

def render_options_to_docx_table(doc: docx.Document, opts: List[str]):
    if not opts:
        return

    raw_clean = []
    for o_idx, opt in enumerate(opts):
        letter = ['A', 'B', 'C', 'D'][o_idx] if o_idx < 4 else f"({o_idx+1})"
        clean_text = re.sub(r'^[A-D]\s*[\.,:]\s*', '', str(opt)).strip()
        raw_clean.append((letter, clean_text))

    max_len = max(len(t[1]) for t in raw_clean)

    if max_len < 20 and len(raw_clean) == 4:
        cols_count = 4
        rows_data = [raw_clean]
        col_widths = [Inches(1.7), Inches(1.7), Inches(1.7), Inches(1.7)]
    elif max_len < 45 and len(raw_clean) == 4:
        cols_count = 2
        rows_data = [
            [raw_clean[0], raw_clean[1]],
            [raw_clean[2], raw_clean[3]]
        ]
        col_widths = [Inches(3.4), Inches(3.4)]
    else:
        cols_count = 1
        rows_data = [[item] for item in raw_clean]
        col_widths = [Inches(6.8)]

    table = doc.add_table(rows=len(rows_data), cols=cols_count)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for r_idx, row in enumerate(rows_data):
        for c_idx, (letter, opt_text) in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.width = col_widths[c_idx]
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.line_spacing = 1.15
            
            r_let = p.add_run(f"{letter}. ")
            r_let.bold = True
            r_let.font.name = 'Times New Roman'
            r_let.font.size = Pt(11)
            write_formatted_text_with_math(p, opt_text)

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(3)

def insert_base64_image_to_docx(doc: docx.Document, b64_str: str):
    if not b64_str or not isinstance(b64_str, str):
        return
    try:
        clean_b64 = re.sub(r'^data:image\/[a-zA-Z]+;base64,', '', b64_str.strip())
        img_bytes = base64.b64decode(clean_b64)
        img_stream = io.BytesIO(img_bytes)
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run()
        run.add_picture(img_stream, width=Inches(3.6))
    except Exception as e:
        print(f"[Word Image Error]: {e}")

@app.post("/api/exams/export-word")
async def export_word_docx(request: Request):
    data = await request.json()
    base = data.get("base_exam") or {}
    include_answers = data.get("include_answers", False)
    
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.78)
        section.bottom_margin = Inches(0.78)
        section.left_margin = Inches(0.78)
        section.right_margin = Inches(0.78)

    title_p = doc.add_paragraph()
    t_run = title_p.add_run(base.get('exam_title', 'ĐỀ KIỂM TRA ĐỊNH KỲ').upper())
    t_run.bold = True
    t_run.font.name = 'Times New Roman'
    t_run.font.size = Pt(14)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_after = Pt(2)

    info_p = doc.add_paragraph()
    i_run = info_p.add_run(f"Môn: {base.get('subject', 'Toán học')} | Khối: {base.get('grade', 12)} | Thời gian: 90 phút\n")
    i_run.bold = True
    i_run.font.name = 'Times New Roman'
    i_run.font.size = Pt(11)
    info_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_p.paragraph_format.space_after = Pt(12)

    def add_exam_section(heading: str, q_list: list, q_start_idx: int):
        if not q_list:
            return q_start_idx
        
        sec_table = doc.add_table(rows=1, cols=1)
        sec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = sec_table.cell(0, 0)
        set_cell_shading(cell, "EEF2FF")
        set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        s_run = p.add_run(heading)
        s_run.bold = True
        s_run.font.name = 'Times New Roman'
        s_run.font.size = Pt(11.5)
        s_run.font.color.rgb = RGBColor(49, 46, 129)

        doc.add_paragraph()

        for idx, q in enumerate(q_list):
            num = q_start_idx + idx
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(6)
            p_q.paragraph_format.space_after = Pt(3)
            p_q.paragraph_format.line_spacing = 1.15
            
            lbl_run = p_q.add_run(f"Câu {num} [{q.get('cognitive_level', 'M2')}]: ")
            lbl_run.bold = True
            lbl_run.font.name = 'Times New Roman'
            lbl_run.font.size = Pt(11)
            
            raw_content = q.get('content', '')
            render_content_block_to_docx(doc, raw_content)

            # CHÈN HÌNH MINH HỌA ĐÚNG BLOCK VĂN BẢN
            if q.get('image_base64'):
                insert_base64_image_to_docx(doc, q.get('image_base64'))

            opts = q.get('options') or []
            if opts:
                render_options_to_docx_table(doc, opts)

            if include_answers and q.get('explanation'):
                sol_table = doc.add_table(rows=1, cols=1)
                sol_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                s_cell = sol_table.cell(0, 0)
                set_cell_shading(s_cell, "F0FDF4")
                set_cell_margins(s_cell, top=100, bottom=100, left=150, right=150)
                
                sp = s_cell.paragraphs[0]
                sp.paragraph_format.space_before = Pt(2)
                sp.paragraph_format.space_after = Pt(2)
                sol_title = sp.add_run("Lời giải chi tiết & Barem:\n")
                sol_title.bold = True
                sol_title.font.name = 'Times New Roman'
                sol_title.font.size = Pt(10.5)
                sol_title.font.color.rgb = RGBColor(6, 95, 70)
                write_formatted_text_with_math(sp, q.get('explanation', ''))
                doc.add_paragraph()

        return q_start_idx + len(q_list)

    cur_idx = 1
    cur_idx = add_exam_section("PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN", base.get('part_1_mcq', []), cur_idx)
    cur_idx = add_exam_section("PHẦN II. TRẮC NGHIỆM ĐÚNG / SAI", base.get('part_2_tf', []), cur_idx)
    cur_idx = add_exam_section("PHẦN III. TỰ LUẬN / TRẢ LỜI NGẮN", base.get('part_3_short_ans') or base.get('part_3_essay', []), cur_idx)

    end_p = doc.add_paragraph()
    end_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p.paragraph_format.space_before = Pt(14)
    e_run = end_p.add_run("--- HẾT ---")
    e_run.bold = True
    e_run.font.name = 'Times New Roman'

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="De_Thi_Chuan.docx"'}
    )

@app.get("/health")
@app.get("/")
def health_check():
    return {"status": "ok", "service": "EduExam Deterministic Engine", "version": "75.0.0"}

@app.post("/api/upload-mistral-zip")
async def upload_mistral_zip(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    level_stage: str = Form("thpt"),
    grade: int = Form(12),
    subject: str = Form("Toán học"),
    topic: str = Form("Chung"),
    auto_save: Optional[Union[bool, str]] = Form(False)
):
    content_bytes = await file.read()
    is_auto_save = str(auto_save).lower() in ("true", "1", "yes")

    parsed_questions = parse_mistral_zip_stream(
        content_bytes, default_grade=grade, default_stage=level_stage,
        default_subject=subject, default_topic=topic
    )

    saved_count = 0
    if is_auto_save and parsed_questions:
        conn = get_db()
        c = conn.cursor()
        for q in parsed_questions:
            c.execute("""
                INSERT INTO questions (
                    level_stage, grade, subject, topic, lesson, level, cognitive_level, q_type,
                    content, options_json, correct_answer, explanation, has_image, image_base64, tikz_code
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                q.get("level_stage", level_stage), int(q.get("grade", grade)), q.get("subject", subject),
                q.get("topic", topic), q.get("lesson", topic), "Thông hiểu",
                q.get("cognitive_level", "M2"), q.get("q_type", "essay"), q.get("content", ""),
                json.dumps(q.get("options") or [], ensure_ascii=False) if q.get("options") else None,
                "", q.get("explanation", ""), q.get("has_image", 0), q.get("image_base64", ""), ""
            ))
            saved_count += 1
        conn.commit()
        conn.close()
        background_tasks.add_task(backup_db_to_drive)

    return {
        "status": "success",
        "total_parsed": len(parsed_questions),
        "saved_count": saved_count,
        "questions": parsed_questions
    }

@app.post("/api/questions/bulk-save")
def bulk_save_questions(req: BulkSaveRequest, background_tasks: BackgroundTasks):
    conn = get_db()
    c = conn.cursor()
    saved_ids = []
    for q in req.questions:
        c.execute("""
            INSERT INTO questions (
                level_stage, grade, subject, topic, lesson, level, cognitive_level, q_type,
                content, options_json, correct_answer, explanation, has_image, image_base64, tikz_code
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            q.level_stage, q.grade, q.subject, q.topic, q.lesson or "Chung",
            q.level, q.cognitive_level or "M2", q.q_type,
            generic_clean_latex(q.content),
            json.dumps([generic_clean_latex(o) for o in (q.options or [])], ensure_ascii=False) if q.options else None,
            "", generic_clean_latex(q.explanation), q.has_image, q.image_base64, q.tikz_code
        ))
        saved_ids.append(c.lastrowid)
    conn.commit()
    conn.close()
    background_tasks.add_task(backup_db_to_drive)
    return {"status": "success", "count": len(saved_ids), "ids": saved_ids}

@app.get("/api/questions/filter")
def filter_questions(
    level_stage: Optional[str] = None,
    grade: Optional[int] = None,
    q_type: Optional[str] = None,
    lesson: Optional[str] = None,
    cognitive_level: Optional[str] = None,
    keyword: Optional[str] = None
):
    conn = get_db()
    c = conn.cursor()
    query = "SELECT * FROM questions WHERE 1=1"
    params = []
    if level_stage: query += " AND level_stage = ?"; params.append(level_stage)
    if grade: query += " AND grade = ?"; params.append(grade)
    if q_type: query += " AND q_type = ?"; params.append(q_type)
    if lesson: query += " AND (lesson LIKE ? OR topic LIKE ?)"; params.extend([f"%{lesson}%", f"%{lesson}%"])
    if cognitive_level: query += " AND cognitive_level = ?"; params.append(cognitive_level)
    if keyword: query += " AND content LIKE ?"; params.append(f"%{keyword}%")

    query += " ORDER BY id DESC LIMIT 200"
    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    results = []
    for r in rows:
        item = dict(r)
        item["options"] = json.loads(item["options_json"]) if item.get("options_json") else []
        results.append(item)
    return {"total": len(results), "questions": results}

@app.delete("/api/questions/{question_id}")
def delete_question(question_id: int):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM questions WHERE id = ?", (question_id,))
    deleted = c.rowcount
    conn.commit()
    conn.close()
    if deleted == 0:
        raise HTTPException(status_code=404, detail="Không tìm thấy câu hỏi để xóa.")
    return {"status": "success", "deleted_id": question_id}

@app.post("/api/exams/generate-full")
def generate_full_exam(req: GenerateExamRequest):
    if not req.selected_question_ids:
        raise HTTPException(status_code=400, detail="Vui lòng chọn ít nhất 1 câu hỏi từ kho để tạo đề.")

    conn = get_db()
    c = conn.cursor()
    placeholders = ','.join(['?'] * len(req.selected_question_ids))
    c.execute(f"SELECT * FROM questions WHERE id IN ({placeholders})", req.selected_question_ids)
    rows = c.fetchall()
    conn.close()

    id_map = {r["id"]: dict(r) for r in rows}
    ordered_items = [id_map[qid] for qid in req.selected_question_ids if qid in id_map]

    p1, p2, p3 = [], [], []
    for item in ordered_items:
        item["options"] = json.loads(item["options_json"]) if item.get("options_json") else []
        if item["q_type"] == "single_choice":
            p1.append(item)
        elif item["q_type"] == "true_false":
            p2.append(item)
        else:
            p3.append(item)

    return {
        "status": "success",
        "base_exam": {
            "exam_title": req.exam_title or "ĐỀ KIỂM TRA ĐỊNH KỲ",
            "subject": req.subject,
            "grade": req.grade,
            "level": req.level,
            "part_1_mcq": p1,
            "part_2_tf": p2,
            "part_3_short_ans": p3
        },
        "answer_matrix": {"101": {f"Câu {i+1}": "A" for i in range(len(p1))}}
    }

@app.post("/api/drive/upload-db")
def api_upload_db():
    success, message = backup_db_to_drive()
    return {"status": "success" if success else "error", "message": message}

@app.post("/api/drive/restore-db")
def api_restore_db():
    success, message = restore_db_from_drive()
    if success:
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM questions")
        count = c.fetchone()[0]
        conn.close()
        return {
            "status": "success",
            "message": f"{message} (Hiện có {count} câu hỏi trong ngân hàng)",
            "count": count
        }
    return {"status": "error", "message": message}

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)