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
import gspread
from google.oauth2.service_account import Credentials
import json
from datetime import datetime

# =====================================================
# DATA & CONFIG
# =====================================================
APP_NAME = "JobPilot AI - Professional Edition"

REGIONS = {
    "السعودية (KSA)": "Saudi Arabia",
    "الإمارات (UAE)": "UAE",
    "مصر (Egypt)": "Egypt",
    "العالم (Global)": "Global"
}

JOBS = [
    {"id": 1, "title": "Petroleum Engineer", "company": "Aramco", "location": "Dhahran, KSA", "salary": "25,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["petroleum", "reservoir", "drilling", "simulation"], "url": "https://linkedin.com/jobs"},
    {"id": 2, "title": "Data Analyst", "company": "STC", "location": "Riyadh, KSA", "salary": "14,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "sql", "excel", "power bi"], "url": "https://linkedin.com/jobs"},
    {"id": 3, "title": "Mechanical Engineer", "company": "SABIC", "location": "Jubail, KSA", "salary": "18,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["autocad", "solidworks", "maintenance"], "url": "https://linkedin.com/jobs"},
    {"id": 4, "title": "Senior Python Developer", "company": "NEOM", "location": "Tabuk, KSA", "salary": "35,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "aws", "docker", "api"], "url": "https://linkedin.com/jobs"},
    {"id": 5, "title": "Cybersecurity Analyst", "company": "SDAIA", "location": "Riyadh, KSA", "salary": "20,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["network security", "linux", "python"], "url": "https://linkedin.com/jobs"},
    {"id": 6, "title": "Civil Engineer", "company": "Red Sea Global", "location": "Jeddah, KSA", "salary": "22,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["autocad", "civil", "project management"], "url": "https://linkedin.com/jobs"},
    {"id": 7, "title": "Software Engineer", "company": "Careem", "location": "Dubai, UAE", "salary": "20,000 AED", "category": "tech", "region": "UAE", "skills": ["javascript", "react", "node.js"], "url": "https://linkedin.com/jobs"},
    {"id": 8, "title": "Process Engineer", "company": "ADNOC", "location": "Abu Dhabi, UAE", "salary": "25,000 AED", "category": "engineering", "region": "UAE", "skills": ["hysys", "simulation", "safety"], "url": "https://linkedin.com/jobs"},
    {"id": 9, "title": "Business Analyst", "company": "CIB Bank", "location": "Cairo, Egypt", "salary": "18,000 EGP", "category": "business", "region": "Egypt", "skills": ["analysis", "excel", "sql"], "url": "https://linkedin.com/jobs"},
]

TRANSLATIONS = {
    "ar": {
        "title": "جوب بايلوت AI - المحترف", "upload": "📤 ارفع سيرتك الذاتية", "region": "🌍 المنطقة المستهدفة",
        "ats_tab": "📊 تحليل ATS", "jobs_tab": "🎯 الوظائف المناسبة", "pdf_tab": "📄 تحميل السيرة", "tracker_tab": "💼 تطبيقاتي",
        "score": "درجة التوافق", "matched": "المهارات المتوفرة", "missing": "المهارات المطلوبة", "apply": "تقدم الآن 🚀",
        "improve_msg": "💡 نصيحة للتحسين: أضف مهارات رقمية وإنجازات محددة.", "fields": "المجالات المكتشفة",
        "download_btn": "🪄 إنشاء وتحميل السيرة (Harvard Style)", "no_apps": "لا توجد تطبيقات محفوظة بعد.",
        "save_success": "✅ تم حفظ التقديم في قاعدة البيانات!"
    },
    "en": {
        "title": "JobPilot AI - Professional", "upload": "📤 Upload Your Resume", "region": "🌍 Target Region",
        "ats_tab": "📊 ATS Analysis", "jobs_tab": "🎯 Job Matches", "pdf_tab": "📄 Get Professional CV", "tracker_tab": "💼 My Applications",
        "score": "ATS Compatibility Score", "matched": "Matched Skills", "missing": "Missing Skills", "apply": "Apply Now 🚀",
        "improve_msg": "💡 Improvement Tip: Add quantifiable results and industry keywords.", "fields": "Detected Fields",
        "download_btn": "🪄 Generate & Download CV (Harvard Style)", "no_apps": "No applications saved yet.",
        "save_success": "✅ Application saved to tracker!"
    }
}

# =====================================================
# GOOGLE SHEETS CONNECTION
# =====================================================
@st.cache_resource
def get_google_sheet():
    try:
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        client = gspread.authorize(creds)
        # Open the sheet by name
        sheet = client.open("JobPilot Applications").sheet1
        return sheet
    except Exception as e:
        return None

def save_application(job, status="Applied"):
    sheet = get_google_sheet()
    if sheet:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # timestamp, user_email, job_id, job_title, company, location, salary, status, job_url, notes
        row = [timestamp, "user@example.com", job['id'], job['title'], job['company'], job['location'], job['salary'], status, job.get('url', ''), '']
        sheet.append_row(row)
        return True
    return False

def get_applications():
    sheet = get_google_sheet()
    if sheet:
        return sheet.get_all_records()
    return []

# =====================================================
# CORE LOGIC
# =====================================================
def extract_text(file):
    if file.type == "application/pdf":
        return "\n".join(page.extract_text() or "" for page in PyPDF2.PdfReader(file).pages)
    return "\n".join(p.text for p in docx.Document(file).paragraphs)

def get_ats_score(text):
    wc = len(text.split())
    if wc > 300: return 95, "Strong"
    return 75, "Good"

def get_matches(text, region_label, min_match):
    tl = text.lower()
    region_code = REGIONS[region_label]
    filtered = [j for j in JOBS if j["region"] == region_code or region_code == "Global"]
    
    results = []
    for job in filtered:
        matched = [s for s in job["skills"] if s in tl]
        score = int((len(matched) / len(job["skills"])) * 100) if job["skills"] else 50
        if score >= min_match:
            results.append({"job": job, "score": score, "matched": matched, "missing": [s for s in job["skills"] if s not in matched]})
    return results

def make_harvard_pdf(text):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
    styles = getSampleStyleSheet()
    name_s = ParagraphStyle("Name", fontSize=22, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=12)
    section_s = ParagraphStyle("Section", fontSize=12, fontName="Helvetica-Bold", spaceBefore=12, spaceAfter=6)
    body_s = ParagraphStyle("Body", fontSize=10, fontName="Helvetica", leading=12)
    
    story = [
        Paragraph("CANDIDATE NAME", name_s),
        Paragraph("LinkedIn | Portfolio | your.email@example.com", ParagraphStyle("Sub", alignment=TA_CENTER, fontSize=10)),
        Spacer(1, 15),
        Paragraph("PROFESSIONAL SUMMARY", section_s),
        HRFlowable(width="100%", thickness=1, color=colors.black),
        Paragraph("Goal-oriented professional with a strong technical background and a proven track record of excellence.", body_s),
        Paragraph("TECHNICAL SKILLS", section_s),
        HRFlowable(width="100%", thickness=1, color=colors.black),
        Paragraph("Python, SQL, AutoCAD, Project Management, Strategic Execution", body_s),
        Paragraph("EXPERIENCE", section_s),
        HRFlowable(width="100%", thickness=1, color=colors.black),
        Paragraph("<b>Lead Specialist</b> | Top Tier Co | 2022 - Present", body_s),
        Paragraph("- Executed complex projects resulting in 20% efficiency gain and streamlined operations.", body_s)
    ]
    doc.build(story)
    return buffer.getvalue()

# =====================================================
# UI INTERFACE
# =====================================================
def main():
    st.set_page_config(page_title=APP_NAME, layout="wide")
    
    st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    h1, h2, h3, p, span, li, label, div { color: #111827 !important; }
    [data-testid="stSidebar"] { background-color: #F9FAFB !important; border-right: 1px solid #E5E7EB; }
    .score-box { background: #0A66C2; color: white !important; padding: 30px; border-radius: 15px; text-align: center; margin-bottom: 20px; }
    .score-box h1 { color: white !important; margin: 0; }
    .job-card { background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px; padding: 25px; margin-bottom: 15px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
    .skill-pill { display: inline-block; padding: 5px 12px; margin: 3px; background: #E0E7FF; color: #1E3A8A !important; border-radius: 15px; font-size: 12px; font-weight: 500; }
    .matched-pill { background: #D1FAE5; color: #065F46 !important; }
    .missing-pill { background: #FEE2E2; color: #991B1B !important; }
    .stButton>button { background-color: #0A66C2 !important; color: white !important; border-radius: 8px; font-weight: bold; border: none; height: 48px; }
    </style>
    """, unsafe_allow_html=True)

    if "lang" not in st.session_state: st.session_state.lang = "en"
    lang = st.session_state.lang
    L = TRANSLATIONS[lang]

    with st.sidebar:
        st.markdown(f"### 🌐 {L['title']}")
        if st.button("🔄 Switch to " + ("English" if lang == "ar" else "العربية")):
            st.session_state.lang = "en" if lang == "ar" else "ar"
            st.rerun()
        st.markdown("---")
        uploaded = st.file_uploader(L["upload"], type=["pdf", "docx"])
        region_label = st.selectbox(L["region"], list(REGIONS.keys()))
        min_match = st.slider(L["score"], 0, 100, 20)

    if uploaded:
        text = extract_text(uploaded)
        score, status = get_ats_score(text)
        
        st.title(L["title"])
        
        tab1, tab2, tab3, tab4 = st.tabs([L["ats_tab"], L["jobs_tab"], L["pdf_tab"], L["tracker_tab"]])
        
        with tab1:
            st.markdown(f"<div class='score-box'><h1>{L['score']}: {score}% ({status})</h1></div>", unsafe_allow_html=True)
            st.progress(score/100)
            st.info(L["improve_msg"])
            
        with tab2:
            st.subheader(f"{L['jobs_tab']} ({region_label})")
            matches = get_matches(text, region_label, min_match)
            if not matches:
                st.warning("No jobs found matching your criteria.")
            for m in matches:
                with st.container():
                    st.markdown(f"""
                    <div class="job-card">
                        <h3 style="color: #0A66C2 !important;">{m['job']['title']}</h3>
                        <p><b>🏢 {m['job']['company']}</b> | 📍 {m['job']['location']}</p>
                        <p style="color: #059669; font-weight: bold; font-size: 18px;">Match: {m['score']}%</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(f"**{L['matched']}**")
                        for s in m['matched']: st.markdown(f"<span class='skill-pill matched-pill'>{s}</span>", unsafe_allow_html=True)
                    with c2:
                        st.markdown(f"**{L['missing']}**")
                        for s in m['missing']: st.markdown(f"<span class='skill-pill missing-pill'>{s}</span>", unsafe_allow_html=True)
                    
                    if st.button(L["apply"], key=f"apply_{m['job']['id']}"):
                        if save_application(m['job'], "Applied"):
                            st.success(L["save_success"])
                        else:
                            st.error("Error saving to database.")
                    st.markdown("---")

        with tab3:
            st.subheader(L["pdf_tab"])
            if st.button(L["download_btn"], use_container_width=True):
                pdf = make_harvard_pdf(text)
                st.download_button("📥 Click here to Save PDF", pdf, file_name="Resume_Harvard_Pro.pdf", mime="application/pdf")

        with tab4:
            st.subheader("My Applications Tracker")
            apps = get_applications()
            if apps:
                st.dataframe(apps, use_container_width=True)
            else:
                st.info(L["no_apps"])
    else:
        st.title(L["title"])
        st.info("👋 " + L["upload"] + " in the sidebar.")

if __name__ == "__main__":
    main()
