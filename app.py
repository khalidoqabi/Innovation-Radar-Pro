import streamlit as st
import requests
import re
from docx import Document
from io import BytesIO
import base64

# --- إعدادات الصفحة ---
st.set_page_config(page_title="Innovation Radar Pro v2", layout="wide")

# --- دالة إنشاء ملف Word مخصص ---
def create_docx(report_text):
    doc = Document()
    doc.add_heading('تقرير رادار الابتكار الاحترافي', 0)
    doc.add_paragraph(report_text)
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- دالة الاتصال المطور (تدعم الصور والنصوص) ---
def call_pro_api(prompt, image_file=None):
    try:
        api_key = st.secrets["PRO_API_KEY"]
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={api_key}"
        
        if image_file:
            # إذا وجدت صورة، يتم إرسالها مع النص (Multimodal)
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
        return f"خطأ: {response.status_code}"
    except Exception as e:
        return f"فشل: {str(e)}"

# --- واجهة المستخدم ---
st.markdown("<h1 style='text-align:center; color:#1e3a8a;'>🛡️ رادار الابتكار Pro (النسخة التنفيذية)</h1>", unsafe_allow_html=True)

if not st.session_state.get("gate_passed"):
    col1, col2 = st.columns([2, 1])
    with col1:
        idea_input = st.text_area("✍️ اشرح فكرتك التقنية:", height=200)
    with col2:
        uploaded_file = st.file_uploader("🖼️ ارفع رسم كروكي أو صورة (اختياري):", type=["jpg", "png", "jpeg"])
    
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
            بصفتك خبير براءات اختراع، حلل الفكرة (والصورة المرفقة إن وجدت): "{st.session_state.final_idea}"
            وقدم تقريراً باللغة العربية مقسماً بوضوح كالتالي:
            [===LEVEL1===] التشخيص الاستراتيجي والمنافسين.
            [===LEVEL2===] المطالبات التقنية وأقرب 3 اختراعات مشابهة.
            [===LEVEL3===] الجدوى الاقتصادية وخارطة الطريق.
            [===AUDIT===] نسبة الفرادة المحتملة.
            [===SOVEREIGNTY===] توصية السيادة التسويقية.
            """
            st.session_state.full_report = call_pro_api(prompt, st.session_state.uploaded_file)

    # --- عرض التقرير بتنسيق احترافي ---
    report = st.session_state.full_report
    parts = re.split(r'\[===LEVEL[1-3]===\]|\[===AUDIT===\]|\[===SOVEREIGNTY===\]', report)
    
    tab1, tab2, tab3 = st.tabs(["📊 التشخيص", "🔧 التقنية والمراجع", "🛣️ التنفيذ"])
    
    with tab1:
        st.info(parts[1] if len(parts)>1 else "جاري التحميل...")
    with tab2:
        st.success(parts[2] if len(parts)>2 else "جاري التحميل...")
    with tab3:
        st.warning(parts[3] if len(parts)>3 else "جاري التحميل...")

    st.divider()
    # عرض النتائج النهائية في صناديق بارزة
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"### ⭐ الفرادة\n{parts[4] if len(parts)>4 else 'N/A'}")
    with c2:
        st.markdown(f"### 💡 السيادة\n{parts[5] if len(parts)>5 else 'N/A'}")

    # --- أزرار التحميل ---
    st.divider()
    docx_file = create_docx(report)
    st.download_button(
        label="📥 تحميل التقرير كملف Word مفسر",
        data=docx_file,
        file_name="Innovation_Report_Pro.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    if st.button("🔄 فحص ابتكار جديد"):
        st.session_state.clear()
        st.rerun()
