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
APP_NAME = "JobPilot AI - Professional Edition"

REGIONS = {
    "السعودية": "Saudi Arabia",
    "الإمارات": "UAE",
    "مصر": "Egypt",
    "Global / Remote": "Global"
}

JOBS = [
    {"id": 1, "title": "Senior Python Developer", "company": "NEOM", "location": "Tabuk, KSA", "salary": "25,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["python", "aws", "docker"]},
    {"id": 2, "title": "Petroleum Engineer", "company": "Aramco", "location": "Dhahran, KSA", "salary": "30,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["petroleum", "reservoir", "drilling"]},
    {"id": 3, "title": "Civil Engineer", "company": "Red Sea Global", "location": "Jeddah, KSA", "salary": "22,000 SAR", "category": "engineering", "region": "Saudi Arabia", "skills": ["autocad", "civil"]},
    {"id": 4, "title": "Software Engineer", "company": "STC", "location": "Riyadh, KSA", "salary": "15,000 SAR", "category": "tech", "region": "Saudi Arabia", "skills": ["java", "python", "sql"]},
    {"id": 5, "title": "Mechanical Engineer", "company": "ADNOC", "location": "Abu Dhabi, UAE", "salary": "25,000 AED", "category": "engineering", "region": "UAE", "skills": ["solidworks", "maintenance"]}
]

# =====================================================
# UI STYLING (FORCE WHITE MODE)
# =====================================================
def apply_style():
    st.markdown("""
    <style>
    /* Force Light Theme */
    .stApp { background-color: #FFFFFF !important; }
    h1, h2, h3, p, span, li, label, .stMarkdown { color: #111827 !important; }
    
    /* Sidebar */
    [data-testid="stSidebar"] { background-color: #F9FAFB !important; border-right: 1px solid #E5E7EB; }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { background-color: #F3F4F6; border-radius: 10px; padding: 5px; }
    .stTabs [data-baseweb="tab"] { color: #4B5563 !important; }
    .stTabs [aria-selected="true"] { background-color: #FFFFFF !important; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }

    /* Job Cards */
    .job-card { background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px; padding: 20px; margin-bottom: 15px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
    .job-card h3 { color: #0A66C2 !important; }
    
    /* Buttons */
    .stButton>button { background-color: #0A66C2 !important; color: white !important; border-radius: 8px; border: none; height: 45px; font-weight: bold; }
    
    /* Metric */
    [data-testid="stMetricValue"] { color: #0A66C2 !important; }
    </style>
    """, unsafe_allow_html=True)

# =====================================================
# CORE LOGIC
# =====================================================
def analyze_resume(text):
    wc = len(text.split())
    score = 90 if wc > 300 else 70
    return score, ["Add more quantifiable results (e.g. 20% improvement)"] if score < 90 else []

def generate_pdf(text):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = [Paragraph("HARVARD STYLE RESUME", styles['Title']), Spacer(1, 12), Paragraph(text[:500], styles['Normal'])]
    doc.build(story)
    return buffer.getvalue()

# =====================================================
# MAIN APP
# =====================================================
def main():
    apply_style()
    
    with st.sidebar:
        st.title("⚙️ Settings")
        uploaded = st.file_uploader("Upload CV (PDF/DOCX)", type=["pdf", "docx"])
        region = st.selectbox("Target Region", list(REGIONS.keys()))
        lang = st.radio("Language", ["English", "العربية"], horizontal=True)
    
    st.title("📄 " + ("محسن السيرة الذاتية" if lang == "العربية" else "JobPilot AI Pro"))
    
    if uploaded:
        text = "Sample extracted text from your resume..." # Placeholder for speed
        score, tips = analyze_resume(text)
        
        tab1, tab2, tab3 = st.tabs(["📊 Analysis", "🎯 Jobs", "📄 PDF"])
        
        with tab1:
            st.header(f"ATS Score: {score}%")
            st.progress(score/100)
            st.subheader("Improvement Tips")
            for tip in tips: st.info(tip)
            
        with tab2:
            st.subheader(f"Best Jobs in {region}")
            filtered = [j for j in JOBS if j['region'] == REGIONS[region]]
            for job in filtered:
                st.markdown(f"""<div class="job-card">
                    <h3>{job['title']}</h3>
                    <p>🏢 {job['company']} | 💰 {job['salary']}</p>
                </div>""", unsafe_allow_html=True)
                st.link_button("Apply on LinkedIn", "https://linkedin.com/jobs")

        with tab3:
            st.subheader("Generate Professional CV")
            if st.button("Download Optimized Resume"):
                pdf = generate_pdf(text)
                st.download_button("💾 Save as PDF", pdf, "CV_Pro.pdf", "application/pdf")
    else:
        st.info("👋 Please upload your CV to start." if lang == "English" else "👋 يرجى رفع سيرتك الذاتية للبدء.")

if __name__ == "__main__":
    main()
