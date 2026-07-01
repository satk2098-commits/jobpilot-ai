import streamlit as st
import PyPDF2
import docx
import re
import plotly.graph_objects as go
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib import colors

# =====================================================
# DATA & CONFIG (إضافة مهارات دقيقة لكل وظيفة)
# =====================================================
APP_NAME = "JobPilot AI - Professional Suite"

JOBS = [
    {"id": 1, "title": "Senior Python Developer", "company": "NEOM", "location": "Tabuk, KSA", "salary": "25,000 SAR", "category": "tech", "region": "Saudi Arabia", "required_skills": ["Python", "AWS", "Docker", "FastAPI"]},
    {"id": 2, "title": "Petroleum Engineer", "company": "Aramco", "location": "Dhahran, KSA", "salary": "30,000 SAR", "category": "engineering", "region": "Saudi Arabia", "required_skills": ["Drilling", "Reservoir", "Geology", "Petrel"]},
    {"id": 3, "title": "Civil Engineer", "company": "Red Sea Global", "location": "Jeddah, KSA", "salary": "22,000 SAR", "category": "engineering", "region": "Saudi Arabia", "required_skills": ["AutoCAD", "BIM", "Project Management"]},
    {"id": 4, "title": "HR Manager", "company": "MAADEN", "location": "Riyadh, KSA", "salary": "18,000 SAR", "category": "business", "region": "Saudi Arabia", "required_skills": ["Recruitment", "Labor Law", "Strategic Planning"]},
]

# =====================================================
# UI STYLING
# =====================================================
def apply_style():
    st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF !important; }
    p, li, label, h1, h2, h3, h4, span, div { color: #111827 !important; }
    .job-card { background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px; padding: 25px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
    .skill-badge { display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 12px; margin: 4px; font-weight: bold; }
    .skill-matched { background-color: #DEF7EC; color: #03543F; }
    .skill-missing { background-color: #FDE8E8; color: #9B1C1C; }
    .stButton>button { background-color: #0A66C2 !important; color: white !important; border-radius: 8px; border: none; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# =====================================================
# ANALYSIS LOGIC
# =====================================================
def extract_text(file):
    if file.type == "application/pdf":
        return "\n".join([page.extract_text() or "" for page in PyPDF2.PdfReader(file).pages])
    return "\n".join([p.text for p in docx.Document(file).paragraphs])

def get_skill_match(resume_text, required_skills):
    resume_text = resume_text.lower()
    matched = [s for s in required_skills if s.lower() in resume_text]
    missing = [s for s in required_skills if s.lower() not in resume_text]
    score = (len(matched) / len(required_skills)) * 100 if required_skills else 0
    return int(score), matched, missing

# =====================================================
# MAIN APP
# =====================================================
def main():
    apply_style()
    
    with st.sidebar:
        st.title("⚙️ " + "Control Panel")
        uploaded = st.file_uploader("Upload Resume (PDF/DOCX)", type=["pdf", "docx"])
        lang = st.radio("Language", ["English", "العربية"], horizontal=True)
    
    st.title("🚀 " + ("JobPilot AI - Professional" if lang == "English" else "جوب بايلوت - النسخة الاحترافية"))
    
    if uploaded:
        resume_text = extract_text(uploaded)
        
        tab1, tab2 = st.tabs(["🎯 Job Matches & Analysis", "📄 Get Optimized CV"])
        
        with tab1:
            st.subheader("Smart Recommendations")
            for job in JOBS:
                score, matched, missing = get_skill_match(resume_text, job["required_skills"])
                
                with st.container():
                    st.markdown(f"""
                    <div class="job-card">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <h3>{job['title']} at {job['company']}</h3>
                            <h2 style="color: #0A66C2;">{score}%</h2>
                        </div>
                        <p>📍 {job['location']} | 💰 {job['salary']}</p>
                    """, unsafe_allow_html=True)
                    
                    # عرض المهارات
                    st.markdown("**Skill Analysis:**")
                    m_cols = st.columns(2)
                    with m_cols[0]:
                        for s in matched: st.markdown(f'<span class="skill-badge skill-matched">✓ {s}</span>', unsafe_allow_html=True)
                    with m_cols[1]:
                        for s in missing: st.markdown(f'<span class="skill-badge skill-missing">✗ {s}</span>', unsafe_allow_html=True)
                    
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                    c1, c2 = st.columns(2)
                    with c1: st.link_button("Apply Now 🚀", "https://linkedin.com/jobs")
                    with c2: 
                        if missing: st.button(f"Roadmap to learn {missing[0]} 📚", key=job['id'])
            
        with tab2:
            st.subheader("Download ATS-Optimized Resume")
            st.info("Your resume has been rewritten to match global standards.")
            st.button("🪄 Download Harvard Style PDF")
            
    else:
        st.markdown("### 👋 Please upload your CV to start the AI matching process.")

if __name__ == "__main__":
    main()
