import streamlit as st
import requests
import re
from docx import Document
from io import BytesIO
import base64

# --- 1. إعدادات الهوية البصرية وتنسيق RTL (حل مشكلة الاتجاه والتلاصق) ---
st.set_page_config(page_title="Innovation Radar Pro v2", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
    
    /* إجبار التطبيق بالكامل على اتجاه اليمين إلى اليسار */
    html, body, [data-testid="stAppViewContainer"], .stApp {
        direction: rtl !important;
        text-align: right !important;
        font-family: 'Tajawal', sans-serif !important;
    }
    
    /* تنسيق النصوص داخل التبويبات والصناديق لضمان عدم التلاصق */
    p, li, span, div {
        direction: rtl !important;
        text-align: right !important;
        line-height: 1.8 !important; /* زيادة المسافة بين الأسطر لمنع تداخل الحروف */
        letter-spacing: 0.2px !important; /* مسافة بسيطة بين الحروف */
    }

    /* تحسين شكل العناوين */
    h1, h2, h3 {
        color: #1e3a8a !important;
        font-weight: 700 !important;
    }

    /* تنسيق أزرار التبويبات لتظهر بشكل صحيح من اليمين */
    .stTabs [data-baseweb="tab-list"] {
        direction: rtl !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. دالة إنشاء ملف Word (بخطوط عربية منسقة) ---
def create_docx(report_text):
    doc = Document()
    # تنظيف النص من الرموز البرمجية قبل وضعه في الوورد
    clean_text = re.sub(r'\[===.*?===\]', '', report_text)
    doc.add_heading('تقرير رادار الابتكار الاحترافي', 0)
    p = doc.add_paragraph(clean_text)
    p.alignment = 1 # محاذاة للوسط أو اليمين
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. المحرك المباشر المتوافق مع Gemini Flash Latest ---
def call_pro_api(prompt, image_file=None):
    try:
        api_key = st.secrets["PRO_API_KEY"]
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={api_key}"
        
        if image_file:
            img_data = base64.b64encode(image_file.read()).decode('utf-8')
            payload = {
                "contents": [{
                    "parts": [
                        {"text": prompt},
                        {"inline_data": {"mime_type": "image/jpeg", "data": img_data}}
                    ]
                }]
            }
        else:
            payload = {"contents": [{"parts": [{"text": prompt}]}]}

        response = requests.post(url, json=payload, headers={'Content-Type': 'application/json'})
        if response.status_code == 200:
            return response.json()['candidates'][0]['content']['parts'][0]['text']
        return f"خطأ في الاتصال: {response.status_code}"
    except Exception as e:
        return f"فشل المحرك: {str(e)}"

# --- 4. واجهة المستخدم ---
st.markdown("<h1 style='text-align:center;'>🛡️ رادار الابتكار Pro</h1>", unsafe_allow_html=True)

if not st.session_state.get("gate_passed"):
    col1, col2 = st.columns([2, 1])
    with col1:
        idea_input = st.text_area("✍️ اشرح فكرتك التقنية بالتفصيل:", height=200, placeholder="مثال: نظام ذكي لصيانة الجسور باستخدام الدرون...")
    with col2:
        uploaded_file = st.file_uploader("🖼️ ارفع رسم كروكي (اختياري):", type=["jpg", "png", "jpeg"])
    
    if st.button("بدء الفحص الاستراتيجي 🚀"):
        if idea_input:
            st.session_state.final_idea = idea_input
            st.session_state.uploaded_file = uploaded_file
            st.session_state.gate_passed = True
            st.rerun()
else:
    if st.session_state.get("full_report") is None:
        with st.spinner("جاري تحليل البيانات والصور وتوليد التقرير المنسق..."):
            prompt = f"""
            بصفتك خبير براءات اختراع عالمي، حلل الفكرة (والصورة المرفقة إن وجدت): "{st.session_state.final_idea}"
            وقدم تقريراً باللغة العربية مقسماً بوضوح باستخدام الأوسمة التالية حصراً:
            [===LEVEL1===] (للتشخيص الاستراتيجي)
            [===LEVEL2===] (للمطالبات والمراجع)
            [===LEVEL3===] (للجدوى والطريق)
            [===AUDIT===] (للفرادة)
            [===SOVEREIGNTY===] (للسيادة)
            ملاحظة: اجعل المسافات واضحة بين الكلمات والأسطر.
            """
            st.session_state.full_report = call_pro_api(prompt, st.session_state.uploaded_file)

    report = st.session_state.full_report
    parts = re.split(r'\[===LEVEL[1-3]===\]|\[===AUDIT===\]|\[===SOVEREIGNTY===\]', report)
    
    tab1, tab2, tab3 = st.tabs(["📊 التشخيص الاستراتيجي", "🔧 المراجع التقنية", "🛣️ خارطة التنفيذ"])
    
    with tab1:
        st.markdown(f"<div style='direction:rtl; text-align:right;'>{parts[1] if len(parts)>1 else report}</div>", unsafe_allow_html=True)
    with tab2:
        st.markdown(f"<div style='direction:rtl; text-align:right;'>{parts[2] if len(parts)>2 else ''}</div>", unsafe_allow_html=True)
    with tab3:
        st.markdown(f"<div style='direction:rtl; text-align:right;'>{parts[3] if len(parts)>3 else ''}</div>", unsafe_allow_html=True)

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.info(f"**⭐ مؤشر الفرادة**\n\n{parts[4] if len(parts)>4 else 'جاري التقييم...'}")
    with c2:
        st.warning(f"**💡 توصية السيادة**\n\n{parts[5] if len(parts)>5 else 'جاري التحليل...'}")

    st.divider()
    docx_file = create_docx(report)
    st.download_button(
        label="📥 تحميل التقرير كملف Word منسق",
        data=docx_file,
        file_name="Innovation_Report_Pro.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    if st.button("🔄 فحص ابتكار جديد"):
        st.session_state.clear()
        st.rerun()
