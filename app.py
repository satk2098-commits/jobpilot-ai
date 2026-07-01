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
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib import colors

# =====================================================
# CONFIGURATION & DATA
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
    # KSA - Tech
    {"id": 1, "title": "Senior Python Developer", "company": "NEOM", "location": "Tabuk, KSA", "salary": "25,000 - 35,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "aws", "docker", "api"], "url": "https://www.linkedin.com/jobs/"},
    {"id": 2, "title": "Data Scientist", "company": "Aramco", "location": "Dhahran, KSA", "salary": "20,000 - 30,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "machine learning", "sql", "statistics"], "url": "https://www.linkedin.com/jobs/"},
    {"id": 3, "title": "Cybersecurity Specialist", "company": "SABIC", "location": "Jubail, KSA", "salary": "18,000 - 25,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["network security", "linux", "siem"], "url": "https://www.linkedin.com/jobs/"},
    # KSA - Engineering
    {"id": 4, "title": "Petroleum Engineer", "company": "Aramco", "location": "Dhahran, KSA", "salary": "22,000 - 40,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["petroleum", "reservoir", "drilling", "simulation"], "url": "https://www.linkedin.com/jobs/"},
    {"id": 5, "title": "Mechanical Engineer", "company": "Red Sea Global", "location": "Jeddah, KSA", "salary": "15,000 - 22,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["autocad", "solidworks", "maintenance"], "url": "https://www.linkedin.com/jobs/"},
    # UAE
    {"id": 6, "title": "AI Research Engineer", "company": "G42", "location": "Abu Dhabi, UAE", "salary": "25,000 - 45,000 AED", "category": "tech", "region": "UAE", "skills": ["pytorch", "tensorflow", "nlp"], "url": "https://www.linkedin.com/jobs/"},
    {"id": 7, "title": "Business Analyst", "company": "Emirates NBD", "location": "Dubai, UAE", "salary": "15,000 - 22,000 AED", "category": "business", "region": "UAE", "skills": ["analysis", "excel", "sql", "tableau"], "url": "https://www.linkedin.com/jobs/"}
]

TRANSLATIONS = {
    "ar": {
        "title": "جوب بايلوت AI - المحترف", "settings": "الإعدادات", "upload": "ارفع سيرتك الذاتية", "region": "المنطقة",
        "cat": "المجال", "min_match": "الحد الأدنى للمطابقة", "lang": "اللغة", "ats_tab": "تحليل ATS", "improve_tab": "تحسين السيرة",
        "jobs_tab": "وظائف مقترحة", "pdf_tab": "تحميل PDF", "score": "درجة ATS", "matched": "مهارات متطابقة", "missing": "مهارات ناقصة",
        "apply": "تقدم الآن", "download": "تحميل السيرة المحسنة", "detected": "المجالات المكتشفة"
    },
    "en": {
        "title": "JobPilot AI - Professional", "settings": "Settings", "upload": "Upload Resume", "region": "Region",
        "cat": "Category", "min_match": "Min Match %", "lang": "Language", "ats_tab": "ATS Analysis", "improve_tab": "Improvement",
        "jobs_tab": "Job Matches", "pdf_tab": "Get PDF", "score": "ATS Score", "matched": "Matched", "missing": "Missing",
        "apply": "Apply Now", "download": "Download Optimized CV", "detected": "Detected Fields"
    }
}

# =====================================================
# CORE FUNCTIONS
# =====================================================
def read_file(file):
    if file.type == "application/pdf":
        reader = PyPDF2.PdfReader(file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    else:
        doc = docx.Document(file)
        return "\n".join([p.text for p in doc.paragraphs])

def analyze_resume(text):
    score = 80
    issues = []
    if len(text.split()) < 300: 
        score -= 20
        issues.append("Resume is too short. Add more professional details.")
    if not re.search(r'\d+%', text): issues.append("Add quantifiable results (e.g. 20% growth).")
    return max(0, score), issues

def generate_harvard_pdf(text):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    
    # Harvard Style Settings
    name_style = ParagraphStyle("Name", fontSize=20, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=12)
    section_style = ParagraphStyle("Section", fontSize=12, fontName="Helvetica-Bold", spaceBefore=12, spaceAfter=6, borderPadding=2)
    body_style = ParagraphStyle("Body", fontSize=10, fontName="Helvetica", leading=12)

    story = []
    story.append(Paragraph("YOUR NAME", name_style))
    story.append(Paragraph("City, Country | Phone: +000 000 000 | Email: user@example.com", ParagraphStyle("Sub", alignment=TA_CENTER)))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("PROFESSIONAL SUMMARY", section_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.black))
    story.append(Paragraph("Highly motivated professional with expertise in technical and analytical domains. Proven ability to deliver results.", body_style))
    
    story.append(Paragraph("TECHNICAL SKILLS", section_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.black))
    story.append(Paragraph("Python, SQL, AWS, Docker, Project Management, Strategic Planning", body_style))

    story.append(Paragraph("EXPERIENCE", section_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.black))
    story.append(Paragraph("<b>Senior Associate</b> | Global Tech Co | 2021 - Present", body_style))
    story.append(Paragraph("- Led team of 5 to develop automated reporting system, reducing work hours by 15%.", body_style))

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

# =====================================================
# INTERFACE
# =====================================================
def main():
    st.set_page_config(page_title=APP_NAME, layout="wide")
    
    # Custom CSS for Professional UI (White background, Black text)
    st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; color: #000000 !important; }
    p, li, label, h1, h2, h3, div { color: #000000 !important; }
    .job-card { background: #F8F9FA; border: 1px solid #DEE2E6; border-radius: 10px; padding: 20px; margin-bottom: 10px; }
    .stSidebar { background-color: #F1F3F5 !important; }
    .score-box { background: #0A66C2; color: white !important; padding: 20px; border-radius: 10px; text-align: center; }
    [data-testid="stMetricValue"] { color: #0A66C2 !important; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

    if "lang" not in st.session_state: st.session_state.lang = "en"
    lang = st.session_state.lang
    L = TRANSLATIONS[lang]

    # Sidebar
    with st.sidebar:
        st.title("⚙️ " + L["settings"])
        if st.button("🔄 Switch to " + ("English" if lang == "ar" else "العربية")):
            st.session_state.lang = "en" if lang == "ar" else "ar"
            st.rerun()
        
        uploaded = st.file_uploader(L["upload"], type=["pdf", "docx"])
        region_label = st.selectbox(L["region"], list(REGIONS.keys()))
        min_match = st.slider(L["min_match"], 0, 100, 20)

    # Main Area
    st.title("🚀 " + L["title"])
    
    if uploaded:
        text = read_file(uploaded)
        score, issues = analyze_resume(text)
        
        t1, t2, t3, t4 = st.tabs([L["ats_tab"], L["improve_tab"], L["jobs_tab"], L["pdf_tab"]])
        
        with t1:
            st.markdown(f"<div class='score-box'><h1>{L['score']}: {score}%</h1></div>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            col1.metric(L["score"], f"{score}%")
            col2.metric("Status", "Good" if score > 70 else "Needs Work")
            
        with t2:
            st.subheader("💡 Tips for Success")
            for issue in issues: st.warning(issue)
            st.info("Tip: Always mention specific tools like Python or AutoCAD if they are in the job description.")

        with t3:
            st.subheader(L["jobs_tab"])
            filtered_jobs = [j for j in JOBS if j["region"] == REGIONS[region_label] or REGIONS[region_label] == "Global"]
            for job in filtered_jobs:
                with st.container():
                    st.markdown(f"""
                    <div class="job-card">
                        <h3>{job['title']}</h3>
                        <p><b>{job['company']}</b> | {job['location']}</p>
                        <p style="color: #059669; font-weight: bold;">Salary: {job['salary']}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    st.link_button(L["apply"], job["url"])

        with t4:
            st.subheader(L["pdf_tab"])
            if st.button("🪄 " + L["download"]):
                pdf = generate_harvard_pdf(text)
                st.download_button("📥 Save PDF", pdf, file_name="Resume_Harvard_Style.pdf", mime="application/pdf")
    else:
        st.markdown("### 👋 Welcome to JobPilot Pro")
        st.info("Please upload your CV in the sidebar to begin your journey to success.")

if __name__ == "__main__":
    main()
