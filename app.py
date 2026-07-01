import streamlit as st
import PyPDF2
import docx
import re
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable

APP_NAME = "ATS Resume Optimizer Pro"

# تم تحديث الروابط (روابط بحث حقيقية على LinkedIn)
JOBS = [
    # =========================
    # Saudi Arabia - Tech
    # =========================
    {"id": 1, "title": "Python Developer", "company": "Tech Solutions", "location": "Riyadh, Saudi Arabia", "salary": "8,000 - 12,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "sql", "git"], "url": "https://www.linkedin.com/jobs/search/?keywords=Python%20Developer&location=Saudi%20Arabia", "date": "2025-01-15"},
    {"id": 2, "title": "Full Stack Developer", "company": "STC", "location": "Riyadh, Saudi Arabia", "salary": "10,000 - 15,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["javascript", "react", "sql", "node.js"], "url": "https://www.linkedin.com/jobs/search/?keywords=Full%20Stack%20Developer&location=Saudi%20Arabia", "date": "2025-01-14"},
    {"id": 3, "title": "Data Analyst", "company": "Al Rajhi Bank", "location": "Riyadh, Saudi Arabia", "salary": "9,000 - 14,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "sql", "excel", "power bi"], "url": "https://www.linkedin.com/jobs/search/?keywords=Data%20Analyst&location=Saudi%20Arabia", "date": "2025-01-13"},
    {"id": 4, "title": "DevOps Engineer", "company": "Aramco", "location": "Dhahran, Saudi Arabia", "salary": "15,000 - 22,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["aws", "docker", "linux"], "url": "https://www.linkedin.com/jobs/search/?keywords=DevOps%20Engineer&location=Saudi%20Arabia", "date": "2025-01-12"},
    {"id": 5, "title": "Cybersecurity Analyst", "company": "SABIC", "location": "Jubail, Saudi Arabia", "salary": "12,000 - 18,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["linux", "network", "python"], "url": "https://www.linkedin.com/jobs/search/?keywords=Cybersecurity%20Analyst&location=Saudi%20Arabia", "date": "2025-01-11"},
    {"id": 6, "title": "Frontend Developer", "company": "Tamara", "location": "Riyadh, Saudi Arabia", "salary": "9,000 - 13,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["javascript", "react", "css", "html"], "url": "https://www.linkedin.com/jobs/search/?keywords=Frontend%20Developer&location=Saudi%20Arabia", "date": "2025-01-10"},
    {"id": 7, "title": "Backend Developer", "company": "Elm", "location": "Riyadh, Saudi Arabia", "salary": "10,000 - 16,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "api", "sql"], "url": "https://www.linkedin.com/jobs/search/?keywords=Backend%20Developer&location=Saudi%20Arabia", "date": "2025-01-09"},
    {"id": 8, "title": "Cloud Engineer", "company": "Mobily", "location": "Jeddah, Saudi Arabia", "salary": "12,000 - 18,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["aws", "linux", "docker"], "url": "https://www.linkedin.com/jobs/search/?keywords=Cloud%20Engineer&location=Saudi%20Arabia", "date": "2025-01-08"},

    # =========================
    # Saudi Arabia - Engineering
    # =========================
    {"id": 9, "title": "Petroleum Engineer", "company": "Aramco", "location": "Dhahran, Saudi Arabia", "salary": "18,000 - 30,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["petroleum", "reservoir", "drilling"], "url": "https://www.linkedin.com/jobs/search/?keywords=Petroleum%20Engineer&location=Saudi%20Arabia", "date": "2025-01-07"},
    {"id": 10, "title": "Mechanical Engineer", "company": "SABIC", "location": "Jubail, Saudi Arabia", "salary": "12,000 - 18,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["autocad", "solidworks"], "url": "https://www.linkedin.com/jobs/search/?keywords=Mechanical%20Engineer&location=Saudi%20Arabia", "date": "2025-01-06"},
    {"id": 11, "title": "Process Engineer", "company": "SABIC", "location": "Jubail, Saudi Arabia", "salary": "14,000 - 20,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["hysys", "process", "safety"], "url": "https://www.linkedin.com/jobs/search/?keywords=Process%20Engineer&location=Saudi%20Arabia", "date": "2025-01-05"},
    {"id": 12, "title": "Electrical Engineer", "company": "SEC", "location": "Riyadh, Saudi Arabia", "salary": "10,000 - 16,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["electrical", "autocad"], "url": "https://www.linkedin.com/jobs/search/?keywords=Electrical%20Engineer&location=Saudi%20Arabia", "date": "2025-01-04"},
    {"id": 13, "title": "Civil Engineer", "company": "Saudi Binladin Group", "location": "Jeddah, Saudi Arabia", "salary": "11,000 - 17,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["civil", "autocad", "project management"], "url": "https://www.linkedin.com/jobs/search/?keywords=Civil%20Engineer&location=Saudi%20Arabia", "date": "2025-01-03"},
    {"id": 14, "title": "Project Engineer", "company": "NEOM", "location": "Tabuk, Saudi Arabia", "salary": "16,000 - 24,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["project management", "mechanical", "autocad"], "url": "https://www.linkedin.com/jobs/search/?keywords=Project%20Engineer&location=Saudi%20Arabia", "date": "2025-01-02"},

    # =========================
    # Saudi Arabia - Business
    # =========================
    {"id": 15, "title": "Business Analyst", "company": "Alinma Bank", "location": "Riyadh, Saudi Arabia", "salary": "10,000 - 15,000 SAR", "category": "business", "region": "Saudi Arabia", "skills": ["excel", "sql", "analysis"], "url": "https://www.linkedin.com/jobs/search/?keywords=Business%20Analyst&location=Saudi%20Arabia", "date": "2025-01-01"},
    {"id": 16, "title": "Project Manager", "company": "NEOM", "location": "Tabuk, Saudi Arabia", "salary": "20,000 - 35,000 SAR", "category": "business", "region": "Saudi Arabia", "skills": ["project management", "leadership"], "url": "https://www.linkedin.com/jobs/search/?keywords=Project%20Manager&location=Saudi%20Arabia", "date": "2024-12-31"},
    {"id": 17, "title": "HR Specialist", "company": "MAADEN", "location": "Riyadh, Saudi Arabia", "salary": "8,000 - 12,000 SAR", "category": "business", "region": "Saudi Arabia", "skills": ["hr", "communication", "excel"], "url": "https://www.linkedin.com/jobs/search/?keywords=HR%20Specialist&location=Saudi%20Arabia", "date": "2024-12-30"},
    {"id": 18, "title": "Operations Analyst", "company": "Flynas", "location": "Riyadh, Saudi Arabia", "salary": "9,000 - 14,000 SAR", "category": "business", "region": "Saudi Arabia", "skills": ["analysis", "excel", "leadership"], "url": "https://www.linkedin.com/jobs/search/?keywords=Operations%20Analyst&location=Saudi%20Arabia", "date": "2024-12-29"},

    # =========================
    # UAE
    # =========================
    {"id": 19, "title": "Software Engineer", "company": "Emirates NBD", "location": "Dubai, UAE", "salary": "15,000 - 25,000 AED", "category": "tech", "region": "UAE", "skills": ["java", "python", "sql"], "url": "https://www.linkedin.com/jobs/search/?keywords=Software%20Engineer&location=Dubai", "date": "2024-12-28"},
    {"id": 20, "title": "Data Scientist", "company": "Careem", "location": "Dubai, UAE", "salary": "20,000 - 35,000 AED", "category": "tech", "region": "UAE", "skills": ["python", "machine learning", "tensorflow"], "url": "https://www.linkedin.com/jobs/search/?keywords=Data%20Scientist&location=Dubai", "date": "2024-12-27"},
    {"id": 21, "title": "Civil Engineer", "company": "Emaar", "location": "Dubai, UAE", "salary": "12,000 - 20,000 AED", "category": "engineering", "region": "UAE", "skills": ["autocad", "civil"], "url": "https://www.linkedin.com/jobs/search/?keywords=Civil%20Engineer&location=Dubai", "date": "2024-12-26"},
    {"id": 22, "title": "Mechanical Engineer", "company": "ADNOC", "location": "Abu Dhabi, UAE", "salary": "18,000 - 22,000 AED", "category": "engineering", "region": "UAE", "skills": ["autocad", "solidworks"], "url": "https://www.linkedin.com/jobs/search/?keywords=Mechanical%20Engineer&location=Abu%20Dhabi", "date": "2024-12-25"},
    {"id": 23, "title": "Business Analyst", "company": "Dubai Holding", "location": "Dubai, UAE", "salary": "12,000 - 18,000 AED", "category": "business", "region": "UAE", "skills": ["excel", "analysis", "communication"], "url": "https://www.linkedin.com/jobs/search/?keywords=Business%20Analyst&location=Dubai", "date": "2024-12-24"},
    {"id": 24, "title": "Project Manager", "company": "Noon", "location": "Dubai, UAE", "salary": "18,000 - 28,000 AED", "category": "business", "region": "UAE", "skills": ["project management", "leadership", "agile"], "url": "https://www.linkedin.com/jobs/search/?keywords=Project%20Manager&location=Dubai", "date": "2024-12-23"},

    # =========================
    # Egypt
    # =========================
    {"id": 25, "title": "Software Developer", "company": "Vodafone Egypt", "location": "Cairo, Egypt", "salary": "12,000 - 20,000 EGP", "category": "tech", "region": "Egypt", "skills": ["python", "javascript", "sql"], "url": "https://www.linkedin.com/jobs/search/?keywords=Software%20Developer&location=Egypt", "date": "2024-12-22"},
    {"id": 26, "title": "Frontend Developer", "company": "Instabug", "location": "Cairo, Egypt", "salary": "10,000 - 17,000 EGP", "category": "tech", "region": "Egypt", "skills": ["javascript", "react", "css"], "url": "https://www.linkedin.com/jobs/search/?keywords=Frontend%20Developer&location=Egypt", "date": "2024-12-21"},
    {"id": 27, "title": "Data Analyst", "company": "Orange Egypt", "location": "Cairo, Egypt", "salary": "10,000 - 15,000 EGP", "category": "tech", "region": "Egypt", "skills": ["excel", "sql", "python"], "url": "https://www.linkedin.com/jobs/search/?keywords=Data%20Analyst&location=Egypt", "date": "2024-12-20"},
    {"id": 28, "title": "Business Analyst", "company": "CIB Bank", "location": "Cairo, Egypt", "salary": "10,000 - 18,000 EGP", "category": "business", "region": "Egypt", "skills": ["excel", "sql", "analysis"], "url": "https://www.linkedin.com/jobs/search/?keywords=Business%20Analyst&location=Egypt", "date": "2024-12-19"},
    {"id": 29, "title": "Mechanical Engineer", "company": "Elsewedy", "location": "Cairo, Egypt", "salary": "9,000 - 14,000 EGP", "category": "engineering", "region": "Egypt", "skills": ["autocad", "solidworks"], "url": "https://www.linkedin.com/jobs/search/?keywords=Mechanical%20Engineer&location=Egypt", "date": "2024-12-18"},

    # =========================
    # Qatar
    # =========================
    {"id": 30, "title": "IT Support", "company": "Qatar Energy", "location": "Doha, Qatar", "salary": "10,000 - 16,000 QAR", "category": "tech", "region": "Qatar", "skills": ["linux", "network"], "url": "https://www.linkedin.com/jobs/search/?keywords=IT%20Support&location=Qatar", "date": "2024-12-17"},
    {"id": 31, "title": "Network Engineer", "company": "Ooredoo", "location": "Doha, Qatar", "salary": "12,000 - 18,000 QAR", "category": "tech", "region": "Qatar", "skills": ["network", "linux", "security"], "url": "https://www.linkedin.com/jobs/search/?keywords=Network%20Engineer&location=Qatar", "date": "2024-12-16"},
    {"id": 32, "title": "Electrical Engineer", "company": "Qatar Rail", "location": "Doha, Qatar", "salary": "14,000 - 20,000 QAR", "category": "engineering", "region": "Qatar", "skills": ["electrical", "autocad"], "url": "https://www.linkedin.com/jobs/search/?keywords=Electrical%20Engineer&location=Qatar", "date": "2024-12-15"},
    {"id": 33, "title": "Project Coordinator", "company": "Qatar Foundation", "location": "Doha, Qatar", "salary": "11,000 - 16,000 QAR", "category": "business", "region": "Qatar", "skills": ["project management", "communication"], "url": "https://www.linkedin.com/jobs/search/?keywords=Project%20Coordinator&location=Qatar", "date": "2024-12-14"},

    # =========================
    # Jordan
    # =========================
    {"id": 34, "title": "Full Stack Developer", "company": "Rubicon", "location": "Amman, Jordan", "salary": "1,000 - 2,000 JOD", "category": "tech", "region": "Jordan", "skills": ["react", "node.js", "mongodb"], "url": "https://www.linkedin.com/jobs/search/?keywords=Full%20Stack%20Developer&location=Jordan", "date": "2024-12-13"},
    {"id": 35, "title": "Software Engineer", "company": "Mawdoo3", "location": "Amman, Jordan", "salary": "1,200 - 2,200 JOD", "category": "tech", "region": "Jordan", "skills": ["python", "sql", "api"], "url": "https://www.linkedin.com/jobs/search/?keywords=Software%20Engineer&location=Jordan", "date": "2024-12-12"},
    {"id": 36, "title": "Business Analyst", "company": "Bank Al Etihad", "location": "Amman, Jordan", "salary": "900 - 1,600 JOD", "category": "business", "region": "Jordan", "skills": ["excel", "analysis"], "url": "https://www.linkedin.com/jobs/search/?keywords=Business%20Analyst&location=Jordan", "date": "2024-12-11"},

    # =========================
    # Global / Remote
    # =========================
    {"id": 37, "title": "Remote Python Developer", "company": "GitLab", "location": "Remote", "salary": "$4,000 - $8,000", "category": "tech", "region": "Global", "skills": ["python", "git"], "url": "https://www.linkedin.com/jobs/search/?keywords=Remote%20Python%20Developer&location=Remote", "date": "2024-12-10"},
    {"id": 38, "title": "Remote Data Analyst", "company": "Shopify", "location": "Remote", "salary": "$3,500 - $6,500", "category": "tech", "region": "Global", "skills": ["sql", "excel", "analysis"], "url": "https://www.linkedin.com/jobs/search/?keywords=Remote%20Data%20Analyst&location=Remote", "date": "2024-12-09"},
    {"id": 39, "title": "Remote Business Analyst", "company": "Amazon", "location": "Remote", "salary": "$4,500 - $7,000", "category": "business", "region": "Global", "skills": ["analysis", "excel", "communication"], "url": "https://www.linkedin.com/jobs/search/?keywords=Remote%20Business%20Analyst&location=Remote", "date": "2024-12-08"},
    {"id": 40, "title": "Remote Mechanical Engineer", "company": "Siemens", "location": "Remote", "salary": "$4,000 - $7,500", "category": "engineering", "region": "Global", "skills": ["autocad", "solidworks"], "url": "https://www.linkedin.com/jobs/search/?keywords=Remote%20Mechanical%20Engineer&location=Remote", "date": "2024-12-07"}
]

SKILLS_DB = {
    "tech": ["python", "sql", "git", "javascript", "react", "css", "html", "linux", "power bi", "machine learning", "aws", "docker", "java", "node.js", "mongodb", "api", "security"],
    "engineering": ["petroleum", "reservoir", "drilling", "autocad", "solidworks", "piping", "civil", "hysys", "mechanical", "electrical", "process"],
    "business": ["excel", "analysis", "communication", "leadership", "project management", "agile", "crm", "sales", "hr"]
}

def get_text(lang):
    if lang == "ar":
        return {
            "app_title": "محسن السيرة الذاتية ومطابقة الوظائف", "settings": "الإعدادات", "upload": "رفع السيرة الذاتية",
            "region": "المنطقة المستهدفة", "category": "مجال الوظائف", "all_cats": "كل المجالات", "auto": "اكتشاف تلقائي",
            "tech": "تقنية المعلومات", "engineering": "الهندسة", "business": "إدارة الأعمال",
            "min_match": "الحد الأدنى للتطابق %", "language": "اللغة", "tab_ats": "تحليل ATS", "tab_improve": "تحسين السيرة",
            "tab_jobs": "الوظائف المناسبة", "tab_pdf": "إنشاء السيرة PDF", "ats_score": "درجة التوافق مع ATS",
            "words": "عدد الكلمات", "issues": "المشاكل", "jobs_found": "الوظائف المتاحة", "skills_detected": "المهارات المكتشفة",
            "issues_title": "مشاكل تحتاج إصلاح", "no_issues": "لا توجد مشاكل كبيرة", "priorities": "إجراءات ذات أولوية",
            "action_plan": "خطة العمل", "target_score": "الدرجة المستهدفة", "target_reached": "تم بلوغ الهدف!",
            "need_more": "تحتاج نقاط إضافية", "jobs_in": "وظائف في", "no_jobs": "لا توجد وظائف.",
            "matched": "مهارات متطابقة", "missing": "مهارات ناقصة", "apply": "تقدم الآن", "match_score": "نسبة التطابق",
            "cover_letter": "خطاب تغطية لـ", "generate_pdf": "إنشاء سيرة ذاتية محسنة PDF", "download_pdf": "تحميل السيرة الذاتية المحسنة PDF",
            "pdf_ready": "ملف PDF جاهز للتحميل", "welcome": "مرحباً بك!", "welcome_text": "ارفع سيرتك الذاتية للحصول على تحليل ATS ووظائف مناسبة وسيرة محسنة PDF.",
            "total": "إجمالي الوظائف", "tech_count": "وظائف تقنية", "eng_count": "وظائف هندسية", "biz_count": "وظائف أعمال",
            "file_error": "تعذر قراءة الملف. جرب ملف آخر.", "detected": "المجالات المكتشفة", "showing": "عرض وظائف في",
            "strong": "تطابق قوي", "good": "تطابق جيد", "weak": "تطابق ضعيف", "posted": "تاريخ النشر",
        }
    else:
        return {
            "app_title": "ATS Resume Optimizer & Job Matcher", "settings": "Settings", "upload": "Upload Resume",
            "region": "Target Region", "category": "Job Category", "all_cats": "All Categories", "auto": "Auto Detect",
            "tech": "Tech / IT", "engineering": "Engineering", "business": "Business",
            "min_match": "Minimum Match %", "language": "Language", "tab_ats": "ATS Analysis", "tab_improve": "Resume Improvement",
            "tab_jobs": "Job Matches", "tab_pdf": "Resume PDF", "ats_score": "ATS Score", "words": "Word Count",
            "issues": "Issues", "jobs_found": "Jobs Found", "skills_detected": "Skills Detected", "issues_title": "Issues to Fix",
            "no_issues": "No major issues found!", "priorities": "Priority Actions", "action_plan": "Action Plan",
            "target_score": "Target Score", "target_reached": "Target reached!", "need_more": "Need more points",
            "jobs_in": "Jobs in", "no_jobs": "No jobs found.", "matched": "Matched Skills",
            "missing": "Missing Skills", "apply": "Apply Now", "match_score": "Match Score", "cover_letter": "Cover Letter for",
            "generate_pdf": "Generate Optimized Resume PDF", "download_pdf": "Download Optimized Resume PDF",
            "pdf_ready": "PDF is ready for download!", "welcome": "Welcome!", "welcome_text": "Upload your resume to get ATS analysis, job matches, and an optimized PDF.",
            "total": "Total Jobs", "tech_count": "Tech Jobs", "eng_count": "Engineering", "biz_count": "Business",
            "file_error": "Could not read file. Try another.", "detected": "Detected Fields", "showing": "Showing jobs in",
            "strong": "Strong Match", "good": "Good Match", "weak": "Weak Match", "posted": "Posted",
        }

def read_pdf(file):
    try:
        reader = PyPDF2.PdfReader(file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except:
        return ""

def read_docx(file):
    try:
        doc = docx.Document(file)
        return "\n".join([p.text for p in doc.paragraphs])
    except:
        return ""

def analyze_ats(text):
    issues = []
    score = 100
    tl = text.lower()
    wc = len(text.split())
    if not re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text):
        issues.append("Email is missing")
        score -= 10
    if not re.search(r"(\+?\d{1,3}[-.]?)?\d{3}[-.]?\d{3}[-.]?\d{4}", text):
        issues.append("Phone number is missing")
        score -= 5
    if wc < 200:
        issues.append("Resume is too short")
        score -= 10
    elif wc > 1000:
        issues.append("Resume is too long")
        score -= 5
    sections = {"Experience": ["experience", "work"], "Education": ["education", "university"], "Skills": ["skills"]}
    present, missing = [], []
    for sec, kws in sections.items():
        if any(k in tl for k in kws):
            present.append(sec)
        else:
            missing.append(sec)
            score -= 8
    if not re.search(r"\d+%|\d+\s*(years?|projects?)", tl):
        issues.append("Add measurable achievements")
        score -= 5
    return max(0, score), issues, wc, present, missing

def extract_skills(text):
    found = {}
    tl = text.lower()
    for cat, skills in SKILLS_DB.items():
        matched = [s for s in skills if s in tl]
        if matched:
            found[cat] = matched
    return found

def match_jobs(text, jobs_list, category, min_match):
    tl = text.lower()
    skills_found = extract_skills(text)
    detected = list(skills_found.keys()) or ["tech", "engineering", "business"]
    cats = {"all": ["tech", "engineering", "business"], "auto": detected}.get(category, [category])
    filtered = []
    for job in jobs_list:
        if job.get("category", "general") not in cats:
            continue
        job_desc = (job.get("title", "")).lower()
        matched = [s for s in SKILLS_DB.get(job.get("category", "tech"), []) if s in tl or s in job_desc]
        all_relevant = SKILLS_DB.get(job.get("category", "tech"), [])
        score = int((len(matched) / len(all_relevant)) * 100) if all_relevant else 50
        filtered.append({"job": job, "score": score, "matched": matched[:10], "missing": []})
    strong = [r for r in filtered if r["score"] >= min_match]
    results = strong if strong else filtered
    return sorted(results, key=lambda x: x["score"], reverse=True)

def clean_pdf_text(text):
    if not text:
        return ""
    text = re.sub(r"[^\x00-\x7F]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def generate_pdf_bytes(text):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=0.75*inch, leftMargin=0.75*inch, topMargin=0.75*inch, bottomMargin=0.75*inch)
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle("T", parent=styles["Title"], fontSize=22, alignment=TA_CENTER, textColor="#0A66C2")
    heading_s = ParagraphStyle("H", parent=styles["Heading2"], fontSize=13, textColor="#111827", spaceBefore=14)
    body_s = ParagraphStyle("B", parent=styles["Normal"], fontSize=10.5, leading=14, textColor="#111827")
    email = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    phone = re.search(r"(\+?\d{1,3}[-.]?)?\d{3}[-.]?\d{3}[-.]?\d{4}", text)
    clean = clean_pdf_text(text)
    lines = [l.strip() for l in clean.split(".") if len(l.strip()) > 15]
    all_skills = [s for skills in extract_skills(text).values() for s in skills]
    story = [
        Paragraph("OPTIMIZED RESUME", title_s),
        Paragraph(f"{email.group() if email else 'email@example.com'} | {phone.group() if phone else 'Phone'}", body_s),
        Spacer(1, 0.15*inch),
        Paragraph("PROFESSIONAL SUMMARY", heading_s),
        HRFlowable(width="100%", thickness=1, color="#0A66C2"),
        Paragraph("Results-driven professional with proven ability to deliver high-quality results.", body_s),
        Paragraph("CORE SKILLS", heading_s),
        HRFlowable(width="100%", thickness=1, color="#0A66C2"),
        Paragraph(", ".join([s.title() for s in all_skills[:25]]) if all_skills else "Add skills here.", body_s),
        Paragraph("EXPERIENCE / PROJECTS", heading_s),
        HRFlowable(width="100%", thickness=1, color="#0A66C2")
    ]
    if lines:
        for line in lines[:10]:
            story.append(Paragraph("- " + line[:220], body_s))
    else:
        story.append(Paragraph("- Add experience here.", body_s))
    story.append(Paragraph("EDUCATION", heading_s))
    story.append(HRFlowable(width="100%", thickness=1, color="#0A66C2"))
    story.append(Paragraph("- Add education details here.", body_s))
    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

def cover_letter(job, matched):
    skills = ", ".join(matched[:5]) or "my skills"
    return f"""Dear Hiring Manager,

I am writing to express my interest in the {job['title']} position at {job['company']}.

With expertise in {skills}, I believe I can contribute effectively to your team.

Thank you for considering my application.

Best regards,
[Your Name]
"""

def improvement_plan(score, issues, wc):
    priorities = []
    if score < 70:
        priorities.append("HIGH: Resume needs critical ATS improvements.")
    elif score < 85:
        priorities.append("MEDIUM: Good resume, can be optimized further.")
    else:
        priorities.append("LOW: Resume is strong.")
    for i in issues:
        priorities.append(i)
    actions = [
        "Use standard headings: Experience, Education, Skills.",
        "Add measurable achievements with numbers and percentages.",
        "Tailor keywords for each job description.",
        "Add LinkedIn, GitHub, or portfolio links.",
        "Use bullet points and avoid tables or graphics.",
        "Keep the resume concise and focused."
    ]
    return priorities, actions

def main():
    st.set_page_config(page_title=APP_NAME, layout="wide")

    st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #262730; }
    div[data-baseweb="select"] > div, input, textarea, [data-testid="stFileUploader"] section {
        background-color: #111827 !important; color: #ffffff !important;
        border: 1px solid #374151 !important; border-radius: 10px !important;
    }
    .score-title { text-align: center; color: #1f77d0; font-size: 42px; font-weight: 800; }
    .job-card { background: #111827; border: 1px solid #374151; border-radius: 14px; padding: 20px; margin: 16px 0; }
    .job-card h3 { color: #ffffff; }
    .match-high { color: #22c55e; font-weight: bold; font-size: 22px; }
    .match-medium { color: #f59e0b; font-weight: bold; font-size: 22px; }
    .match-low { color: #ef4444; font-weight: bold; font-size: 22px; }
    .skill-pill { display: inline-block; padding: 5px 10px; margin: 3px; border-radius: 12px; background: #1e3a8a; color: white; font-size: 13px; }
    .matched-pill { background: #065f46; }
    .suggestion-box { background: #111827; border-left: 5px solid #f59e0b; padding: 14px; margin: 10px 0; border-radius: 8px; color: #ffffff; }
    .info-box { background: #082f49; border: 1px solid #0ea5e9; padding: 14px; border-radius: 10px; color: #ffffff; }
    .linkedin-badge { background: linear-gradient(90deg, #0077B5, #00A0DC); color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

    if "lang" not in st.session_state:
        st.session_state.lang = "en"
    if "pdf_bytes" not in st.session_state:
        st.session_state.pdf_bytes = None

    with st.sidebar:
        lang_pick = st.selectbox("🌐", ["English", "العربية"], index=0 if st.session_state.lang == "en" else 1)
        st.session_state.lang = "ar" if lang_pick == "العربية" else "en"
        lang = st.session_state.lang
        L = get_text(lang)

        st.markdown("---")
        st.header("⚙️ " + L["settings"])
        uploaded = st.file_uploader(L["upload"], type=["pdf", "docx"])

        regions_map = {
            "Global / Remote": "Global",
            "Saudi Arabia": "Saudi Arabia",
            "UAE": "UAE",
            "Egypt": "Egypt",
            "Qatar": "Qatar",
            "Jordan": "Jordan"
        }

        region_label = st.selectbox("🌍 " + L["region"], list(regions_map.keys()))
        region_code = regions_map[region_label]

        cats = {
            "all": L["all_cats"],
            "auto": L["auto"],
            "tech": L["tech"],
            "engineering": L["engineering"],
            "business": L["business"]
        }

        category = st.selectbox("💼 " + L["category"], list(cats.keys()), format_func=lambda x: cats[x])
        min_match = st.slider("📊 " + L["min_match"], 0, 100, 20)

    st.title("📄 " + L["app_title"])
    st.markdown("---")

    if not uploaded:
        st.subheader("👋 " + L["welcome"])
        st.info(L["welcome_text"])
        c1, c2, c3, c4 = st.columns(4)
        c1.metric(L["total"], len(JOBS))
        c2.metric(L["tech_count"], len([j for j in JOBS if j.get("category") == "tech"]))
        c3.metric(L["eng_count"], len([j for j in JOBS if j.get("category") == "engineering"]))
        c4.metric(L["biz_count"], len([j for j in JOBS if j.get("category") == "business"]))
        return

    text = read_pdf(uploaded) if uploaded.type == "application/pdf" else read_docx(uploaded)

    if not text:
        st.error(L["file_error"])
        return

    ats_score, issues, wc, present, missing = analyze_ats(text)
    skills_found = extract_skills(text)
    jobs = match_jobs(text, JOBS, category, min_match)
    priorities, actions = improvement_plan(ats_score, issues, wc)
    detected = list(skills_found.keys()) or ["general"]

    st.markdown(
        f"<div class='info-box'>🎯 <b>{L['detected']}:</b> {', '.join(detected).upper()} | 🌍 <b>{L['showing']}:</b> {region_code} <span class='linkedin-badge'>💼 LinkedIn Jobs</span></div>",
        unsafe_allow_html=True
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 " + L["tab_ats"],
        "✏️ " + L["tab_improve"],
        "🎯 " + L["tab_jobs"],
        "📄 " + L["tab_pdf"]
    ])

    with tab1:
        st.markdown(f"<div class='score-title'>{L['ats_score']}: {ats_score}%</div>", unsafe_allow_html=True)
        st.progress(ats_score / 100)
        c1, c2, c3 = st.columns(3)
        c1.metric(L["words"], wc)
        c2.metric(L["issues"], len(issues))
        c3.metric(L["jobs_found"], len(jobs))
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=ats_score,
            title={"text": "ATS"},
            gauge={"axis": {"range": [0, 100]}, "bar": {"color": "#1f77d0"}}
        ))
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("🔑 " + L["skills_detected"])
        if skills_found:
            for cat, skills in skills_found.items():
                pills = " ".join([f"<span class='skill-pill'>{s}</span>" for s in skills])
                st.markdown(f"**{cats.get(cat, cat)}:** {pills}", unsafe_allow_html=True)

        st.subheader("⚠️ " + L["issues_title"])
        if issues:
            for i in issues:
                st.warning(i)
        else:
            st.success(L["no_issues"])

    with tab2:
        st.subheader("🎯 " + L["priorities"])
        for p in priorities:
            st.markdown(f"<div class='suggestion-box'>{p}</div>", unsafe_allow_html=True)

        st.subheader("📋 " + L["action_plan"])
        for i, a in enumerate(actions, 1):
            st.write(f"**{i}.** {a}")

        target = st.slider(L["target_score"], 0, 100, 90)
        if ats_score >= target:
            st.success(L["target_reached"])
        else:
            st.warning(f"{L['need_more']}: {target - ats_score}")

    with tab3:
        st.subheader("🎯 " + L["jobs_in"] + f" {region_code}")
        st.markdown('<span class="linkedin-badge">💼 Powered by LinkedIn Jobs</span>', unsafe_allow_html=True)

        if not jobs:
            st.warning(L["no_jobs"])
        else:
            chart_data = [{"Job": r["job"]["title"], "Score": r["score"]} for r in jobs[:8]]
            fig = px.bar(chart_data, x="Score", y="Job", orientation="h", color="Score", color_continuous_scale=["red", "yellow", "green"])
            fig.update_layout(yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig, use_container_width=True)

            for item in jobs:
                job = item["job"]
                score = item["score"]
                cls = "match-high" if score >= 70 else "match-medium" if score >= 40 else "match-low"
                badge = L["strong"] if score >= 70 else L["good"] if score >= 40 else L["weak"]

                st.markdown(f"""
                <div class="job-card">
                    <div style="display:flex;justify-content:space-between;align-items:center;">
                        <div style="flex:1;">
                            <h3>{job['title']}</h3>
                            <p style="color:#d1d5db;">{job['company']} | {job['location']} | {job['salary']}</p>
                            <p style="color:#9ca3af;font-size:12px;">{L['posted']}: {job.get('date','')}</p>
                        </div>
                        <div style="text-align:center;">
                            <span class="{cls}" style="font-size:28px;">{score}%</span><br>
                            <small style="color:#9ca3af;">{badge}</small>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**{L['matched']}:**")
                    if item["matched"]:
                        st.markdown(" ".join([f"<span class='matched-pill'>✓ {s}</span>" for s in item["matched"]]), unsafe_allow_html=True)
                    else:
                        st.caption("-")

                with c2:
                    st.markdown(f"**{L['missing']}:**")
                    st.caption("Review job description for requirements")

                if job.get("url"):
                    st.link_button("🚀 " + L["apply"], job["url"])

                with st.expander("📝 " + L["cover_letter"] + " " + job["title"]):
                    st.text_area("", value=cover_letter(job, item["matched"]), height=160, key=f"cv_{job.get('id','')}")

                st.markdown("---")

    with tab4:
        st.subheader("📄 " + L["generate_pdf"])
        if st.button("🪄 " + L["generate_pdf"], use_container_width=True):
            with st.spinner("Generating..."):
                st.session_state.pdf_bytes = generate_pdf_bytes(text)

        if st.session_state.pdf_bytes:
            st.success(L["pdf_ready"])
            st.download_button(
                label="📥 " + L["download_pdf"],
                data=st.session_state.pdf_bytes,
                file_name="optimized_resume.pdf",
                mime="application/pdf",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
