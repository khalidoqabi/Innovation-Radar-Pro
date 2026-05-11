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
    
    /* إجبار التطبيق بالكامل على اتجاه اليمين والمحاذاة لليمين */
    html, body, [data-testid="stAppViewContainer"], .stApp {
        direction: rtl !important;
        text-align: right !important;
        font-family: 'Tajawal', sans-serif !important;
    }
    
    /* منع التوسيط نهائياً في النصوص والفقرات */
    p, li, span, div, label {
        direction: rtl !important;
        text-align: right !important;
        line-height: 1.6 !important;
        display: block; /* يضمن عدم التمركز في الوسط */
    }

    /* ضمان محاذاة العناوين لليمين */
    h1, h2, h3, h4, h5, h6 {
        text-align: right !important;
        color: #1e3a8a !important;
        direction: rtl !important;
    }

    /* تحسين شكل صناديق التنبيه لتكون محاذاتها يميناً */
    div[data-testid="stMarkdownContainer"] > div {
        text-align: right !important;
    }
    </style>
""", unsafe_allow_html=True)
# --- 2. دالة إنشاء ملف Word (بخطوط عربية منسقة) ---
def create_docx(report_text):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    from docx.oxml.ns import qn

    doc = Document()

    # تنظيف النص
    clean_text = re.sub(r'\[===.*?===\]', '', report_text)
    clean_text = clean_text.replace('###', '').replace('---', '').replace('**', '').replace('*', '')

    # إضافة العنوان ومحاذاته
    title = doc.add_heading('تقرير رادار الابتكار الاحترافي', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # ضبط اتجاه العنوان لليمين
    title.paragraph_format.right_to_left = True

    for para in clean_text.split('\n'):
        text = para.strip()
        if text:
            p = doc.add_paragraph()
            
            # --- القوة هنا: ضبط اتجاه النص (Reading Order) من اليمين ---
            p.paragraph_format.right_to_left = True
            
            # المحاذاة البصرية لليمين
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            run = p.add_run(text)
            run.font.name = 'Arial'
            
            # إخبار الوورد أن هذا النص يتبع اللغات ذات الاتجاه المعقد (Complex Scripts)
            run._element.rPr.get_or_add_rFonts().set(qn('w:cs'), 'Arial')
            run._element.rPr.get_or_add_rtl().val = True # تفعيل خاصية RTL للكلمات
            
            if any(h in text for h in ["التشخيص", "المطالبات", "الجدوى", "الفرادة", "توصية"]):
                run.bold = True
                run.font.size = Pt(14)
            else:
                run.font.size = Pt(12)

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()
# --- 3. المحرك المباشر المتوافق مع Gemini Flash Latest ---
def call_pro_api(prompt, image_file=None):
    # المفتاح يُستدعى فقط عند تنفيذ الدالة
    api_key = st.secrets["PRO_API_KEY"]
    
    # الرابط المستقر (v1) لضمان عدم حدوث 403 أو 404
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-flash-latest:generateContent?key={api_key}"
    
    headers = {'Content-Type': 'application/json'}
    
    try:
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
            payload = {
                "contents": [{"parts": [{"text": prompt}]}]
            }

        response = requests.post(url, json=payload, headers=headers, timeout=30)
        
        if response.status_code == 200:
            return response.json()['candidates'][0]['content']['parts'][0]['text']
        else:
            return f"خطأ في الاتصال: {response.status_code} - {response.text}"
            
    except Exception as e:
        return f"فشل النظام في الاستجابة: {str(e)}"

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
