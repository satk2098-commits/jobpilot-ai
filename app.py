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
from reportlab.lib import colors

# =====================================================
# DATA & CONFIG
# =====================================================
APP_NAME = "JobPilot AI Pro"

REGIONS = {
    "Global / Remote": "Global",
    "Saudi Arabia": "Saudi Arabia",
    "UAE": "UAE",
    "Egypt": "Egypt",
    "Qatar": "Qatar",
    "Jordan": "Jordan"
}

JOBS = [
    {"id": 1, "title": "Senior Python Developer", "company": "NEOM", "location": "Tabuk, KSA", "salary": "25,000 - 35,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "aws", "docker", "api"], "url": "https://www.linkedin.com/jobs/"},
    {"id": 2, "title": "Petroleum Engineer", "company": "Aramco", "location": "Dhahran, KSA", "salary": "22,000 - 40,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["petroleum", "reservoir", "drilling", "simulation"], "url": "https://www.linkedin.com/jobs/"},
    {"id": 3, "title": "Data Scientist", "company": "Aramco", "location": "Dhahran, KSA", "salary": "20,000 - 30,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "machine learning", "sql"], "url": "https://www.linkedin.com/jobs/"},
    {"id": 4, "title": "Mechanical Engineer", "company": "Red Sea Global", "location": "Jeddah, KSA", "salary": "15,000 - 22,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["autocad", "solidworks", "maintenance"], "url": "https://www.linkedin.com/jobs/"},
    {"id": 5, "title": "HR Manager", "company": "Al-Futtaim", "location": "Dubai, UAE", "salary": "18,000 - 25,000 AED", "category": "business", "region": "UAE", "skills": ["hr", "management", "recruitment"], "url": "https://www.linkedin.com/jobs/"},
    {"id": 6, "title": "Project Manager", "company": "MAADEN", "location": "Riyadh, KSA", "salary": "20,000 - 30,000 SAR", "category": "business", "region": "Saudi Arabia", "skills": ["pmp", "project management", "agile"], "url": "https://www.linkedin.com/jobs/"}
]

TRANSLATIONS = {
    "ar": {
        "title": "جوب بايلوت AI - المحترف", "settings": "الإعدادات", "upload": "ارفع سيرتك الذاتية", "region": "المنطقة المستهدفة",
        "cat": "مجال الوظائف", "min_match": "الحد الأدنى للمطابقة %", "lang": "اللغة", "ats_tab": "📊 تحليل ATS", "improve_tab": "✏️ تحسين السيرة",
        "jobs_tab": "🎯 الوظائف المناسبة", "pdf_tab": "📄 تحميل PDF", "score": "درجة ATS", "matched": "المهارات المتوفرة", "missing": "المهارات الناقصة",
        "apply": "تقدم الآن 🚀", "download": "تحميل السيرة المحسنة (Harvard Style)", "detected": "المجالات المكتشفة في سيرتك"
    },
    "en": {
        "title": "JobPilot AI - Professional", "settings": "Settings", "upload": "Upload Resume", "region": "Target Region",
        "cat": "Category", "min_match": "Min Match %", "lang": "Language", "ats_tab": "📊 ATS Analysis", "improve_tab": "✏️ Improvement",
        "jobs_tab": "🎯 Job Matches", "pdf_tab": "📄 Get PDF", "score": "ATS Score", "matched": "Matched Skills", "missing": "Missing Skills",
        "apply": "Apply Now 🚀", "download": "Download Optimized Resume (Harvard Style)", "detected": "Detected Professional Fields"
    }
}

# =====================================================
# FUNCTIONS
# =====================================================
def read_file(file):
    if file.type == "application/pdf":
        reader = PyPDF2.PdfReader(file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    else:
        doc = docx.Document(file)
        return "\n".join([p.text for p in doc.paragraphs])

def analyze_resume(text):
    score = 85
    issues = []
    if len(text.split()) < 300: 
        score -= 20
        issues.append("السيرة الذاتية قصيرة جداً / Resume is too short.")
    if not re.search(r'\d+%', text): issues.append("أضف نتائج رقمية / Add quantifiable results (e.g. 20%).")
    return max(0, score), issues

def generate_harvard_pdf(text):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
    styles = getSampleStyleSheet()
    name_s = ParagraphStyle("Name", fontSize=22, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=12)
    section_s = ParagraphStyle("Section", fontSize=12, fontName="Helvetica-Bold", spaceBefore=12, spaceAfter=6)
    body_s = ParagraphStyle("Body", fontSize=10, fontName="Helvetica", leading=14)
    story = [
        Paragraph("CANDIDATE NAME", name_s),
        Paragraph("City, Country | LinkedIn | user@example.com", ParagraphStyle("Sub", alignment=TA_CENTER)),
        Spacer(1, 15),
        Paragraph("PROFESSIONAL SUMMARY", section_s),
        HRFlowable(width="100%", thickness=1, color=colors.black),
        Paragraph("Highly skilled professional with expertise in technical domains. Proven ability to deliver high-quality results.", body_s),
        Paragraph("TECHNICAL SKILLS", section_s),
        HRFlowable(width="100%", thickness=1, color=colors.black),
        Paragraph("Python, SQL, Project Management, AutoCAD, Data Analysis", body_s),
        Paragraph("EXPERIENCE", section_s),
        HRFlowable(width="100%", thickness=1, color=colors.black),
        Paragraph("<b>Lead Engineer</b> | Industrial Co | 2022 - Present", body_s),
        Paragraph("- Optimized workflow by 25% through automation and strategic planning.", body_s)
    ]
    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

# =====================================================
# UI INTERFACE
# =====================================================
def main():
    st.set_page_config(page_title=APP_NAME, layout="wide")
    
    # CSS لفرض اللون الأبيض والنصوص السوداء (LinkedIn Style)
    st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    html, body, [data-testid="stHeader"], .main .block-container, p, span, li, label, h1, h2, h3, h4, div {
        color: #000000 !important;
    }
    [data-testid="stSidebar"] { background-color: #F3F4F6 !important; border-right: 1px solid #E5E7EB; }
    .job-card { background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px; padding: 25px; margin-bottom: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
    .score-box { background: linear-gradient(135deg, #0A66C2 0%, #1E3A5F 100%); color: white !important; padding: 30px; border-radius: 15px; text-align: center; }
    .score-box h1 { color: white !important; }
    .stButton>button { background-color: #0A66C2 !important; color: white !important; border-radius: 8px; width: 100%; height: 50px; font-weight: bold; border: none; }
    div[data-baseweb="select"] > div, input { background-color: #FFFFFF !important; color: #000000 !important; border: 1px solid #D1D5DB !important; }
    </style>
    """, unsafe_allow_html=True)

    if "lang" not in st.session_state: st.session_state.lang = "en"
    lang = st.session_state.lang
    L = TRANSLATIONS[lang]

    with st.sidebar:
        st.markdown(f"### 🌐 {L['lang']}")
        if st.button("English / العربية"):
            st.session_state.lang = "en" if lang == "ar" else "ar"
            st.rerun()
        st.markdown("---")
        st.header(L["settings"])
        uploaded = st.file_uploader(L["upload"], type=["pdf", "docx"])
        region_label = st.selectbox(L["region"], list(REGIONS.keys()))
        min_match = st.slider(L["min_match"], 0, 100, 20)

    if uploaded:
        text = read_file(uploaded)
        score, issues = analyze_resume(text)
        
        st.title(L["title"])
        st.info(f"📍 {L['showing']} {region_label}")

        t1, t2, t3, t4 = st.tabs([L["ats_tab"], L["improve_tab"], L["jobs_tab"], L["pdf_tab"]])
        
        with t1:
            st.markdown(f"<div class='score-box'><h1>{L['score']}: {score}%</h1></div>", unsafe_allow_html=True)
            st.progress(score/100)
            col1, col2, col3 = st.columns(3)
            col1.metric(L["score"], f"{score}%")
            col2.metric("Words", len(text.split()))
            col3.metric("Status", "Professional" if score > 75 else "Needs Update")
            
        with t2:
            st.subheader("📋 Improvement Steps")
            for i in issues: st.error(i)
            st.success("✅ Your layout is ATS-friendly. Focus on adding more industry keywords.")

        with t3:
            st.subheader(L["jobs_tab"])
            # منطق الفلترة المتقدم
            filtered_jobs = [j for j in JOBS if j["region"] == REGIONS[region_label] or REGIONS[region_label] == "Global"]
            if not filtered_jobs:
                st.warning("No jobs found in this region yet.")
            for job in filtered_jobs:
                st.markdown(f"""
                <div class="job-card">
                    <h3 style="color: #0A66C2 !important;">{job['title']}</h3>
                    <p><b>🏢 {job['company']}</b> | 📍 {job['location']}</p>
                    <p style="color: #059669; font-weight: bold;">💰 {job['salary']}</p>
                    <p>Skills: {", ".join(job['skills'])}</p>
                </div>
                """, unsafe_allow_html=True)
                st.link_button(L["apply"], job["url"])

        with t4:
            st.subheader(L["pdf_tab"])
            if st.button("🪄 " + L["download"]):
                pdf = generate_harvard_pdf(text)
                st.download_button("📥 Click to Download PDF", pdf, file_name="Resume_Harvard.pdf", mime="application/pdf")
    else:
        st.title(L["title"])
        st.markdown(f"### 👋 {L['welcome']}")
        st.markdown(L["upload"])

if __name__ == "__main__":
    main()
