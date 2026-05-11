import streamlit as st
import requests
import re

# --- 1. إعدادات الهوية البصرية ---
st.set_page_config(page_title="Innovation Radar Pro", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap');
    html, body, [data-testid="stSidebar"], .stApp {
        direction: rtl !important; text-align: right !important; font-family: 'Tajawal', sans-serif !important;
    }
    .pro-header {
        background: linear-gradient(90deg, #1e3a8a, #10b981);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        font-size: 2.5em; font-weight: 900; text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. إدارة الحالة ---
if "gate_passed" not in st.session_state: st.session_state.gate_passed = False
if "full_report" not in st.session_state: st.session_state.full_report = None
if "final_idea" not in st.session_state: st.session_state.final_idea = ""

# --- 3. المحرك المباشر (REST API) المتوافق مع CURL ---
def call_pro_api(prompt):
    try:
        # تأكد من وضع المفتاح في Secrets باسم PRO_API_KEY
        api_key = st.secrets["PRO_API_KEY"]
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
        
        headers = {'Content-Type': 'application/json'}
        
        # تحويل هيكل الـ CURL إلى قاموس بايثون (JSON)
        payload = {
            "contents": [
                {
                    "parts": [{"text": prompt}]
                }
            ]
        }
        
        response = requests.post(url, json=payload, headers=headers)
        
        if response.status_code == 200:
            result = response.json()
            return result['candidates'][0]['content']['parts'][0]['text']
        else:
            return f"خطأ في الاتصال: {response.status_code} - {response.text}"
    except Exception as e:
        return f"فشل النظام: {str(e)}"

# --- 4. واجهة المستخدم ---
st.markdown("<h1 class='pro-header'>🛡️ رادار الابتكار Pro</h1>", unsafe_allow_html=True)

if not st.session_state.gate_passed:
    idea_input = st.text_area("اشرح فكرتك التقنية هنا...", height=150)
    if st.button("بدء الفحص 🚀"):
        if idea_input:
            st.session_state.final_idea = idea_input
            st.session_state.gate_passed = True
            st.rerun()
else:
    if st.session_state.full_report is None:
        with st.spinner("جاري التحليل المعمق..."):
            pro_prompt = f"""
            بصفتك خبير براءات اختراع، حلل الفكرة: "{st.session_state.final_idea}"
            وقدم تقريراً باللغة العربية كالتالي:
            [===LEVEL1===] التشخيص الاستراتيجي والمنافسين.
            [===LEVEL2===] المطالبات التقنية وأقرب 3 اختراعات مشابهة.
            [===LEVEL3===] الجدوى الاقتصادية وخارطة الطريق.
            [===AUDIT===] نسبة الفرادة المحتملة (نسبة مئوية مع التبرير).
            [===SOVEREIGNTY===] توصية السيادة التسويقية.
            """
            st.session_state.full_report = call_pro_api(pro_prompt)
    
    # عرض النتائج في تبويبات
    parts = re.split(r'\[===LEVEL[1-3]===\]|\[===AUDIT===\]|\[===SOVEREIGNTY===\]', st.session_state.full_report)
    
    tab1, tab2, tab3 = st.tabs(["📊 التشخيص", "🔧 المطالبات والمراجع", "🛣️ الطريق للتنفيذ"])
    
    with tab1: st.write(parts[1] if len(parts) > 1 else st.session_state.full_report)
    with tab2: st.write(parts[2] if len(parts) > 2 else "لا توجد تفاصيل.")
    with tab3: st.write(parts[3] if len(parts) > 3 else "لا توجد تفاصيل.")

    st.divider()
    if len(parts) > 4:
        st.success(f"⭐ مؤشر الفرادة: {parts[4]}")
    if len(parts) > 5:
        st.warning(f"💡 توصية السيادة: {parts[5]}")

    if st.button("فحص جديد"):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()
