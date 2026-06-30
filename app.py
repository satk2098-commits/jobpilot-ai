import streamlit as st
import PyPDF2
import docx
import re
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
import requests
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable

APP_NAME = "ATS Resume Optimizer Pro"

# بيانات الوظائف الافتراضية (fallback)
DEFAULT_JOBS = [
    {"id": 1, "title": "Python Developer", "company": "Tech Solutions", "location": "Riyadh, KSA", "salary": "8000-12000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "sql", "git"], "url": "https://www.linkedin.com/jobs/", "date": "2025-01-15"},
    {"id": 2, "title": "Petroleum Engineer", "company": "Aramco", "location": "Dhahran, KSA", "salary": "15000-25000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["petroleum", "reservoir", "drilling"], "url": "https://www.linkedin.com/jobs/", "date": "2025-01-14"},
    {"id": 3, "title": "Data Analyst", "company": "STC", "location": "Riyadh, KSA", "salary": "9000-14000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "sql", "excel"], "url": "https://www.linkedin.com/jobs/", "date": "2025-01-13"},
    {"id": 4, "title": "Mechanical Engineer", "company": "ADNOC", "location": "Abu Dhabi, UAE", "salary": "18000-22000 AED", "category": "engineering", "region": "UAE", "skills": ["autocad", "solidworks"], "url": "https://www.linkedin.com/jobs/", "date": "2025-01-12"},
    {"id": 5, "title": "Business Analyst", "company": "Consulting Group", "location": "Dubai, UAE", "salary": "12000-18000 AED", "category": "business", "region": "UAE", "skills": ["excel", "sql", "analysis"], "url": "https://www.linkedin.com/jobs/", "date": "2025-01-11"},
]

SKILLS_DB = {
    "tech": ["python", "sql", "git", "javascript", "react", "css", "html", "linux", "network", "power bi", "machine learning", "pandas", "numpy", "api", "django", "flask"],
    "engineering": ["petroleum", "reservoir", "drilling", "simulation", "autocad", "solidworks", "piping", "hysys", "process", "safety", "mechanical", "electrical"],
    "business": ["excel", "analysis", "communication", "leadership", "project management", "agile", "scrum", "crm", "sales", "marketing"]
}

# =====================================================
# TRANSLATION
# =====================================================
def get_text(lang):
    if lang == "ar":
        return {
            "app_title": "محسن السيرة الذاتية ومطابقة الوظائف",
            "settings": "الإعدادات",
            "upload": "رفع السيرة الذاتية",
            "region": "المنطقة المستهدفة",
            "category": "مجال الوظائف",
            "all_cats": "كل المجالات",
            "auto": "اكتشاف تلقائي",
            "tech": "تقنية المعلومات",
            "engineering": "الهندسة والبترول",
            "business": "إدارة الأعمال",
            "min_match": "الحد الأدنى للتطابق %",
            "language": "اللغة",
            "tab_ats": "تحليل ATS",
            "tab_improve": "تحسين السيرة",
            "tab_jobs": "الوظائف المناسبة",
            "tab_pdf": "إنشاء السيرة PDF",
            "ats_score": "درجة التوافق مع ATS",
            "words": "عدد الكلمات",
            "issues": "المشاكل",
            "jobs_found": "الوظائف المتاحة",
            "skills_detected": "المهارات المكتشفة",
            "issues_title": "مشاكل تحتاج إصلاح",
            "no_issues": "لا توجد مشاكل كبيرة",
            "priorities": "إجراءات ذات أولوية",
            "action_plan": "خطة العمل",
            "target_score": "الدرجة المستهدفة",
            "target_reached": "تم بلوغ الهدف!",
            "need_more": "تحتاج نقاط إضافية",
            "jobs_in": "وظائف في",
            "no_jobs": "لا توجد وظائف. جرب تغيير المنطقة أو المجال.",
            "matched": "مهارات متطابقة",
            "missing": "مهارات ناقصة",
            "apply": "تقدم الآن",
            "match_score": "نسبة التطابق",
            "cover_letter": "خطاب تغطية لـ",
            "generate_pdf": "إنشاء سيرة ذاتية محسنة PDF",
            "download_pdf": "تحميل السيرة الذاتية المحسنة PDF",
            "pdf_ready": "ملف PDF جاهز للتحميل",
            "welcome": "مرحباً بك!",
            "welcome_text": "ارفع سيرتك الذاتية للحصول على تحليل ATS ونصائح تحسين ووظائف مناسبة وسيرة محسنة PDF.",
            "total": "إجمالي الوظائف",
            "tech_count": "وظائف تقنية",
            "eng_count": "وظائف هندسية",
            "biz_count": "وظائف أعمال",
            "file_error": "تعذر قراءة الملف. جرب ملف آخر.",
            "detected": "المجالات المكتشفة",
            "showing": "عرض وظائف في",
            "strong": "تطابق قوي",
            "good": "تطابق جيد",
            "weak": "تطابق ضعيف",
            "posted": "تاريخ النشر",
            "real_jobs": "وظائف حقيقية من الإنترنت",
            "loading_jobs": "جاري تحميل الوظائف...",
            "api_error": "تعذر تحميل الوظائف من API. عرض الوظائف الافتراضية.",
        }
    else:
        return {
            "app_title": "ATS Resume Optimizer & Job Matcher",
            "settings": "Settings",
            "upload": "Upload Resume",
            "region": "Target Region",
            "category": "Job Category",
            "all_cats": "All Categories",
            "auto": "Auto Detect",
            "tech": "Tech / IT",
            "engineering": "Engineering & Petroleum",
            "business": "Business & Management",
            "min_match": "Minimum Match %",
            "language": "Language",
            "tab_ats": "ATS Analysis",
            "tab_improve": "Resume Improvement",
            "tab_jobs": "Job Matches",
            "tab_pdf": "Resume PDF",
            "ats_score": "ATS Score",
            "words": "Word Count",
            "issues": "Issues",
            "jobs_found": "Jobs Found",
            "skills_detected": "Skills Detected",
            "issues_title": "Issues to Fix",
            "no_issues": "No major issues found!",
            "priorities": "Priority Actions",
            "action_plan": "Action Plan",
            "target_score": "Target Score",
            "target_reached": "Target reached!",
            "need_more": "Need more points",
            "jobs_in": "Jobs in",
            "no_jobs": "No jobs found. Try changing region or category.",
            "matched": "Matched Skills",
            "missing": "Missing Skills",
            "apply": "Apply Now",
            "match_score": "Match Score",
            "cover_letter": "Cover Letter for",
            "generate_pdf": "Generate Optimized Resume PDF",
            "download_pdf": "Download Optimized Resume PDF",
            "pdf_ready": "PDF is ready for download!",
            "welcome": "Welcome!",
            "welcome_text": "Upload your resume to get ATS analysis, improvement tips, job matches, and an optimized PDF.",
            "total": "Total Jobs",
            "tech_count": "Tech Jobs",
            "eng_count": "Engineering",
            "biz_count": "Business",
            "file_error": "Could not read file. Try another.",
            "detected": "Detected Fields",
            "showing": "Showing jobs in",
            "strong": "Strong Match",
            "good": "Good Match",
            "weak": "Weak Match",
            "posted": "Posted",
            "real_jobs": "Real Jobs from Internet",
            "loading_jobs": "Loading jobs...",
            "api_error": "Could not load jobs from API. Showing default jobs.",
        }

# =====================================================
# JSEARCH API FUNCTION
# =====================================================
def fetch_real_jobs(query, location, num_pages=1):
    """Fetch real jobs from JSearch API"""
    url = "https://jsearch.p.rapidapi.com/search"
    
    api_key = st.secrets.get("RAPIDAPI_KEY", "")
    
    if not api_key:
        st.warning("API Key is not configured. Using default jobs.")
        return None
    
    querystring = {
        "query": f"{query} in {location}",
        "location": location,
        "page": "1",
        "num_pages": str(num_pages)
    }
    
    headers = {
        "X-RapidAPI-Key": api_key,
        "X-RapidAPI-Host": "jsearch.p.rapidapi.com"
    }
    
    try:
        response = requests.get(url, headers=headers, params=querystring, timeout=15)
        response.raise_for_status()
        data = response.json()
        
        jobs = []
        if data.get("status") == "OK" and data.get("data"):
            for job in data["data"][:15]:  # Limit to 15 jobs
                jobs.append({
                    "id": job.get("job_id", f"api_{len(jobs)}"),
                    "title": job.get("job_title", "Unknown Position"),
                    "company": job.get("employer_name", "Unknown Company"),
                    "location": f"{job.get('job_city','')} {job.get('job_country','')}".strip(),
                    "salary": job.get("job_salary_currency", "") + " " + str(job.get("job_min_salary", "N/A")),
                    "category": "tech",
                    "region": location,
                    "skills": [],
                    "url": job.get("job_apply_link", job.get("job_google_link", "#")),
                    "date": job.get("job_posted_at", "")[:10],
                    "description": job.get("job_description", "")[:200]
                })
        return jobs
    except Exception as e:
        st.error(f"API Error: {str(e)}")
        return None

# =====================================================
# FILE READER
# =====================================================
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

# =====================================================
# ATS ANALYSIS
# =====================================================
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
    
    sections = {"Experience": ["experience", "work", "employment"], "Education": ["education", "university", "degree"], "Skills": ["skills", "technologies"]}
    present = []
    missing = []
    for sec, kws in sections.items():
        if any(k in tl for k in kws):
            present.append(sec)
        else:
            missing.append(sec)
            score -= 8
    
    if not re.search(r"\d+%|\d+\s*(years?|months?|projects?)", tl):
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
        
        job_skills = job.get("skills", [])
        job_desc = (job.get("description", "") + " " + job.get("title", "")).lower()
        
        matched = []
        missing = []
        for skill in SKILLS_DB.get(job.get("category", "tech"), []):
            if skill in tl or skill in job_desc:
                matched.append(skill)
            elif skill in job_skills:
                missing.append(skill)
        
        # Calculate score based on matched skills
        all_relevant = SKILLS_DB.get(job.get("category", "tech"), [])
        score = int((len(matched) / len(all_relevant)) * 100) if all_relevant else 50
        
        filtered.append({"job": job, "score": score, "matched": matched[:10], "missing": missing[:10]})
    
    strong = [r for r in filtered if r["score"] >= min_match]
    results = strong if strong else filtered
    return sorted(results, key=lambda x: x["score"], reverse=True)

def clean_pdf_text(text):
    if not text:
        return ""
    for old, new in {"'": "-", "'": "'", "\u201c": '"', "\u201d": '"', "\t": " "}.items():
        text = text.replace(old, new)
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
    
    all_skills = []
    for skills in extract_skills(text).values():
        all_skills.extend(skills)
    
    story = []
    story.append(Paragraph("OPTIMIZED RESUME", title_s))
    story.append(Paragraph(f"{email.group() if email else 'email@example.com'} | {phone.group() if phone else 'Phone'}", body_s))
    story.append(Spacer(1, 0.15*inch))
    story.append(Paragraph("PROFESSIONAL SUMMARY", heading_s))
    story.append(HRFlowable(width="100%", thickness=1, color="#0A66C2"))
    story.append(Paragraph("Results-driven professional with proven ability to deliver high-quality results.", body_s))
    story.append(Paragraph("CORE SKILLS", heading_s))
    story.append(HRFlowable(width="100%", thickness=1, color="#0A66C2"))
    story.append(Paragraph(", ".join([s.title() for s in all_skills[:25]]) if all_skills else "Add skills here.", body_s))
    story.append(Paragraph("EXPERIENCE / PROJECTS", heading_s))
    story.append(HRFlowable(width="100%", thickness=1, color="#0A66C2"))
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

# =====================================================
# MAIN APP
# =====================================================
def main():
    st.set_page_config(page_title=APP_NAME, layout="wide")

    # CSS
    st.markdown("""<style>
    .stApp { background-color: #0e1117; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #262730; }
    div[data-baseweb="select"] > div, input, textarea, [data-testid="stFileUploader"] section {
        background-color: #111827 !important; color: #ffffff !important;
        border: 1px solid #374151 !important; border-radius: 10px !important; }
    .score-title { text-align: center; color: #1f77d0; font-size: 42px; font-weight: 800; }
    .job-card { background: #111827; border: 1px solid #374151; border-radius: 14px; padding: 20px; margin: 16px 0; }
    .job-card h3 { color: #ffffff; }
    .match-high { color: #22c55e; font-weight: bold; font-size: 22px; }
    .match-medium { color: #f59e0b; font-weight: bold; font-size: 22px; }
    .match-low { color: #ef4444; font-weight: bold; font-size: 22px; }
    .skill-pill { display: inline-block; padding: 5px 10px; margin: 3px; border-radius: 12px; background: #1e3a8a; color: white; font-size: 13px; }
    .matched-pill { background: #065f46; }
    .missing-pill { background: #7f1d1d; }
    .suggestion-box { background: #111827; border-left: 5px solid #f59e0b; padding: 14px; margin: 10px 0; border-radius: 8px; color: #ffffff; }
    .info-box { background: #082f49; border: 1px solid #0ea5e9; padding: 14px; border-radius: 10px; color: #ffffff; }
    .api-badge { background: linear-gradient(90deg, #00C853, #64DD17); color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: bold; }
    </style>""", unsafe_allow_html=True)

    if "lang" not in st.session_state:
        st.session_state.lang = "en"
    if "pdf_bytes" not in st.session_state:
        st.session_state.pdf_bytes = None
    if "real_jobs" not in st.session_state:
        st.session_state.real_jobs = []

    # Sidebar
    with st.sidebar:
        lang_pick = st.selectbox("🌐", ["English", "العربية"], index=0 if st.session_state.lang == "en" else 1)
        st.session_state.lang = "ar" if lang_pick == "العربية" else "en"
        lang = st.session_state.lang
        L = get_text(lang)

        st.markdown("---")
        st.header("⚙️ " + L["settings"])
        uploaded = st.file_uploader(L["upload"], type=["pdf", "docx"])
        region_label = st.selectbox("🌍 " + L["region"], list({"Global / Remote": "Global", "Saudi Arabia": "Saudi Arabia", "UAE": "UAE", "Egypt": "Egypt", "Qatar": "Qatar", "Jordan": "Jordan"}.keys()))
        region_code = {"Global / Remote": "Global", "Saudi Arabia": "Saudi Arabia", "UAE": "UAE", "Egypt": "Egypt", "Qatar": "Qatar", "Jordan": "Jordan"}[region_label]
        cats = {"all": L["all_cats"], "auto": L["auto"], "tech": L["tech"], "engineering": L["engineering"], "business": L["business"]}
        category = st.selectbox("💼 " + L["category"], list(cats.keys()), format_func=lambda x: cats[x])
        min_match = st.slider("📊 " + L["min_match"], 0, 100, 20)
        
        # API Toggle
        use_api = st.checkbox("🌐 " + L["real_jobs"], value=True)

    st.title("📄 " + L["app_title"])
    st.markdown("---")

    if not uploaded:
        st.subheader("👋 " + L["welcome"])
        st.info(L["welcome_text"])
        c1, c2, c3, c4 = st.columns(4)
        total_jobs = len(st.session_state.real_jobs) if st.session_state.real_jobs else len(DEFAULT_JOBS)
        c1.metric(L["total"], total_jobs)
        c2.metric(L["tech_count"], len([j for j in (st.session_state.real_jobs or DEFAULT_JOBS) if j.get("category") == "tech"]))
        c3.metric(L["eng_count"], len([j for j in (st.session_state.real_jobs or DEFAULT_JOBS) if j.get("category") == "engineering"]))
        c4.metric(L["biz_count"], len([j for j in (st.session_state.real_jobs or DEFAULT_JOBS) if j.get("category") == "business"]))
        return

    text = read_pdf(uploaded) if uploaded.type == "application/pdf" else read_docx(uploaded)
    if not text:
        st.error(L["file_error"])
        return

    # Fetch real jobs if enabled
    jobs_list = DEFAULT_JOBS
    if use_api:
        with st.spinner(L["loading_jobs"]):
            api_jobs = fetch_real_jobs("developer engineer analyst", region_code, 1)
            if api_jobs:
                jobs_list = api_jobs
                st.session_state.real_jobs = api_jobs
                st.success(f"✅ Loaded {len(api_jobs)} real jobs from API!")
            else:
                st.warning(L["api_error"])
                jobs_list = DEFAULT_JOBS

    ats_score, issues, wc, present, missing = analyze_ats(text)
    skills_found = extract_skills(text)
    jobs = match_jobs(text, jobs_list, category, min_match)
    priorities, actions = improvement_plan(ats_score, issues, wc)
    detected = list(skills_found.keys()) or ["general"]

    st.markdown(f"<div class='info-box'>🎯 <b>{L['detected']}:</b> {', '.join(detected).upper()} | 🌍 <b>{L['showing']}:</b> {region_label} {'🌐 API' if use_api and st.session_state.real_jobs else ''}</div>", unsafe_allow_html=True)

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
        fig = go.Figure(go.Indicator(mode="gauge+number", value=ats_score, title={"text": "ATS"}, gauge={"axis": {"range": [0, 100]}, "bar": {"color": "#1f77d0"}}))
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
        st.subheader("🎯 " + L["jobs_in"] + f" {region_label}")
        if use_api and st.session_state.real_jobs:
            st.markdown('<span class="api-badge">🌐 LIVE API JOBS</span>', unsafe_allow_html=True)
        
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
                st.markdown(f"""<div class="job-card">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                <div><h3>{job['title']}</h3><p style="color:#d1d5db;">{job['company']} | {job['location']} | {job['salary']}</p>
                <p style="color:#9ca3af;font-size:12px;">{L['posted']}: {job.get('date','')}</p></div>
                <div style="text-align:center;"><span class="{cls}" style="font-size:28px;">{score}%</span><br><small style="color:#9ca3af;">{badge}</small></div>
                </div></div>""", unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**{L['matched']}:**")
                    if item["matched"]:
                        st.markdown(" ".join([f"<span class='matched-pill'>{s}</span>" for s in item["matched"]]), unsafe_allow_html=True)
                    else:
                        st.caption("-")
                with c2:
                    st.markdown(f"**{L['missing']}:**")
                    if item["missing"]:
                        st.markdown(" ".join([f"<span class='missing-pill'>{s}</span>" for s in item["missing"]]), unsafe_allow_html=True)
                    else:
                        st.success("All matched!")
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
