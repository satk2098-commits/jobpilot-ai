# ... (نفس Imports السابقة) ...
import pandas as pd # أضفنا مكتبة pandas للإحصائيات

# ... (نفس دوال السابقة) ...

# داخل دالة main() سنضيف نظام التتبع:
def main():
    # ... (التنسيق السابق) ...
    
    # 1. نظام التتبع في Session State
    if "apps_data" not in st.session_state:
        st.session_state.apps_data = {} # قاموس يحفظ حالة كل وظيفة

    # ... (داخل التبويبات) ...
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 " + L["tab_ats"], 
        "✏️ " + L["tab_improve"], 
        "🎯 " + L["tab_jobs"], 
        "🔍 " + L["tab_search"],
        "💼 " + "My Applications" # التبويب الجديد
    ])

    # ... (في تبويب الوظائف) ...
    # عند إضافة زر التقديم:
    if st.button("Apply", key=job['id']):
        st.session_state.apps_data[job['id']] = {"status": "Applied", "date": datetime.now().strftime("%Y-%m-%d")}
    
    # ... (تبويب التطبيقات الجديد) ...
    with tab5:
        st.header("My Application Tracker")
        if not st.session_state.apps_data:
            st.info("No applications tracked yet.")
        else:
            df = pd.DataFrame.from_dict(st.session_state.apps_data, orient='index')
            st.dataframe(df) # عرض جدول التقديمات
