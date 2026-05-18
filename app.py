import streamlit as st
import requests
import re
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn
from io import BytesIO
import base64

# --- 1. إعدادات الهوية البصرية وتنسيق RTL ---
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
        display: block; 
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

# ==========================================
# 1. دالة الاتصال بالمحرك (جوجل API) - نسخة بيتا الإقليمية المستقرة
# ==========================================
def call_pro_api(prompt, image_file=None):
    if "PRO_API_KEY" not in st.secrets:
        return "خطأ: مفتاح API غير موجود في إعدادات Secrets"
        
    api_key = st.secrets["PRO_API_KEY"]
    
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={api_key}"
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


# ==========================================
# 2. دالة توليد ملف الوورد المنسق (تم إصلاح الـ try و except)
# ==========================================
def create_docx(report_text, idea_title):
    try:
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt, RGBColor
        from docx.oxml.ns import qn
        import re
        from io import BytesIO

        doc = Document()

        # تنظيف النص تماماً من الأوسمة والنجوم
        clean_text = re.sub(r'\[===.*?===\]', '', report_text)
        clean_text = clean_text.replace('###', '').replace('---', '').replace('**', '').replace('*', '')

        # استخلاص سطر واحد مختصر كعنوان للابتكار (أول 60 حرف أو السطر الأول)
        short_title = idea_title.split('\n')[0].strip()
        if len(short_title) > 60:
            short_title = short_title[:60] + "..."

        # تعريف الألوان الرسمية
        MAIN_COLOR = RGBColor(30, 58, 138)  # أزرق داكن #1e3a8a ليعكس الهوية البصرية
        TEXT_COLOR = RGBColor(0, 0, 0)      # أسود للمتن

        # --- 1. العنوان الرئيسي في الأعلى (في المنتصف) ---
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.right_to_left = True
        
        run_title = p_title.add_run('تقرير رادار الابتكار الاحترافي')
        run_title.bold = True
        run_title.font.size = Pt(18)
        run_title.font.name = 'Arial'
        run_title.font.color.rgb = MAIN_COLOR
        run_title._element.get_or_add_rPr().get_or_add_rtl().val = True

        # --- 2. عنوان الابتكار المخصص (في المنتصف تحت العنوان الرئيسي) ---
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.right_to_left = True
        p_sub.paragraph_format.space_after = Pt(24) # مسافة جمالية قبل بدء التقرير
        
        run_sub = p_sub.add_run(f"عنوان الابتكار: {short_title}")
        run_sub.bold = False
        run_sub.font.size = Pt(14)
        run_sub.font.name = 'Arial'
        run_sub.font.color.rgb = RGBColor(100, 116, 139) # لون رمادي مميز وخاص بالعنوان الفرعي
        run_sub._element.get_or_add_rPr().get_or_add_rtl().val = True

        # --- 3. متن التقرير (الضبط الكامل Justify) ---
        for para in clean_text.split('\n'):
            text = para.strip()
            if text:
                p = doc.add_paragraph()
                # تطبيق الضبط الكامل لتفعيل "الكشيدة" التلقائية وتنسيق الحواف
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.right_to_left = True
                p.paragraph_format.line_spacing = 1.3 # تباعد أسطر مريح للقراءة
                
                run = p.add_run(text)
                run.font.name = 'Arial'
                run._element.get_or_add_rPr().get_or_add_rtl().val = True
                
                # تمييز العناوين الفرعية لتكون عريضة وبحجم 16 وباللون الأزرق الداكن
                keywords = ["التشخيص", "المطالبات", "الجدوى", "الفرادة", "توصية", "المنافسون", "خارطة"]
                if any(h in text for h in keywords):
                    run.bold = True
                    run.font.size = Pt(16)
                    run.font.color.rgb = MAIN_COLOR
                    p.paragraph_format.space_before = Pt(14)
                    p.paragraph_format.space_after = Pt(6)
                else:
                    # الفقرات العادية بحجم 14 ولون أسود
                    run.font.size = Pt(14)
                    run.font.color.rgb = TEXT_COLOR
                    p.paragraph_format.space_after = Pt(8)

        # --- 4. حقوق الإعداد والتطوير في الأسفل تماماً ---
        doc.add_paragraph() 
        p_footer = doc.add_paragraph()
        p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_footer.paragraph_format.right_to_left = True
        
        run_line = p_footer.add_run("________________________________")
        doc.add_paragraph()
        
        p_credit = doc.add_paragraph()
        p_credit.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_credit.paragraph_format.right_to_left = True
        
        run_credit = p_credit.add_run('إعداد وتطوير: أ. خالد العقبي | المسار الرقمي')
        run_credit.bold = True
        run_credit.font.size = Pt(11)
        run_credit.font.name = 'Arial'
        run_credit.font.color.rgb = MAIN_COLOR
        run_credit._element.get_or_add_rPr().get_or_add_rtl().val = True

        bio = BytesIO()
        doc.save(bio)
        return bio.getvalue()
        
    except Exception as e:
        import streamlit as st
        st.error(f"حدث خطأ أثناء توليد ملف Word: {str(e)}")
        return None

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
            وقدم تقريراً احترافياً باللغة العربية. 
            هام جداً: التزم بترك مسافات واضحة بين الكلمات والأسطر، واستخدم الأوسمة التالية حصراً للتقسيم:
            [===LEVEL1===] التشخيص الاستراتيجي والمنافسين
            [===LEVEL2===] المطالبات التقنية والمراجع
            [===LEVEL3===] الجدوى الاقتصادية وخارطة الطريق
            [===AUDIT===] نسبة الفرادة المحتملة
            [===SOVEREIGNTY===] توصية السيادة التسويقية
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
    docx_file = create_docx(report, st.session_state.final_idea)
    
    if docx_file:
        st.download_button(
            label="📥 تحميل التقرير كملف Word منسق",
            data=docx_file,
            file_name="Innovation_Report_Pro.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    if st.button("🔄 فحص ابتكار جديد"):
        st.session_state.clear()
        st.rerun()
