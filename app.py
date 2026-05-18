import streamlit as st
import requests
import re
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
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
# 2. دالة الاتصال بالمحرك (جوجل API) - النسخة المعتمدة لديك
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
# 3. دالة توليد ملف الوورد المنسق (ضبط كامل وضمان عدم الهروب لليسار)
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

        # 1. استخراج الرقم برمجياً قبل تنظيف النص لتنسيقه بشكل مخصص
        score_match = re.search(r'\[===SCORE===\]\s*(\d+)\s*\[===/SCORE===\]', report_text)
        innovation_score = score_match.group(1) if score_match else None

        # 2. تنظيف النص تماماً وبشكل صارم من كل الأوسمة (بما فيها أوسمة السكور والرقم نفسه)
        # هذا السطر يحذف وسام السكور وما بينهما لكي لا يظهر الرقم مجرداً في المتن
        clean_text = re.sub(r'\[===SCORE===\].*?\[===/SCORE===\]', '', report_text, flags=re.DOTALL)
        # تنظيف بقية الأوسمة والنجوم المتبقية
        clean_text = re.sub(r'\[===.*?===\]', '', clean_text)
        clean_text = clean_text.replace('###', '').replace('---', '').replace('**', '').replace('*', '')

        # استخلاص سطر واحد مختصر كعنوان للابتكار
        short_title = idea_title.split('\n')[0].strip()
        if len(short_title) > 60:
            short_title = short_title[:60] + "..."

        # تعريف الألوان الرسمية
        MAIN_COLOR = RGBColor(30, 58, 138)  # أزرق داكن
        TEXT_COLOR = RGBColor(0, 0, 0)      # أسود للمتن

        # --- أ. العنوان الرئيسي في الأعلى (موسط) ---
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.right_to_left = True
        
        run_title = p_title.add_run('تقرير رادار الابتكار الاحترافي')
        run_title.bold = True
        run_title.font.size = Pt(18)
        run_title.font.name = 'Arial'
        run_title.font.color.rgb = MAIN_COLOR
        run_title._element.get_or_add_rPr().get_or_add_rtl().val = True

        # --- ب. عنوان الابتكار المخصص (موسط) ---
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.right_to_left = True
        
        run_sub = p_sub.add_run(f"عنوان الابتكار: {short_title}")
        run_sub.font.size = Pt(14)
        run_sub.font.name = 'Arial'
        run_sub.font.color.rgb = RGBColor(100, 116, 139)
        run_sub._element.get_or_add_rPr().get_or_add_rtl().val = True

        # --- ج. إضافة مؤشر الفرادة بشكل رسمي مخصص داخل ملف الوورد (إذا وجد) ---
        if innovation_score:
            p_score = doc.add_paragraph()
            p_score.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_score.paragraph_format.right_to_left = True
            p_score.paragraph_format.space_before = Pt(6)
            p_score.paragraph_format.space_after = Pt(24) # مسافة جمالية قبل المتن
            
            # تلوين النسبة بناءً على قيمتها لتطابق المتصفح
            score_num = int(innovation_score)
            SCORE_COLOR = RGBColor(16, 185, 129) if score_num >= 85 else (RGBColor(245, 158, 11) if score_num >= 65 else RGBColor(239, 68, 68))
            
            run_score_label = p_score.add_run("مؤشر الفرادة المحتملة للابتكار: ")
            run_score_label.font.size = Pt(13)
            run_score_label.font.name = 'Arial'
            run_score_label.font.color.rgb = RGBColor(71, 85, 105)
            run_score_label._element.get_or_add_rPr().get_or_add_rtl().val = True
            
            run_score_val = p_score.add_run(f"{innovation_score}%")
            run_score_val.bold = True
            run_score_val.font.size = Pt(16)
            run_score_val.font.name = 'Arial'
            run_score_val.font.color.rgb = SCORE_COLOR
            run_score_val._element.get_or_add_rPr().get_or_add_rtl().val = True
        else:
            p_sub.paragraph_format.space_after = Pt(24)

        # --- د. متن التقرير (الضبط الكامل) ---
        for para in clean_text.split('\n'):
            text = para.strip()
            if text:
                p = doc.add_paragraph()
                p.paragraph_format.right_to_left = True
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.line_spacing = 1.3
                
                run = p.add_run(text)
                run.font.name = 'Arial'
                
                rPr = run._element.get_or_add_rPr()
                rPr.get_or_add_rtl().val = True
                
                keywords = ["التشخيص", "المطالبات", "الجدوى", "الفرادة", "توصية", "المنافسون", "خارطة"]
                if any(h in text for h in keywords):
                    run.bold = True
                    run.font.size = Pt(16)
                    run.font.color.rgb = MAIN_COLOR
                    p.paragraph_format.space_before = Pt(14)
                    p.paragraph_format.space_after = Pt(6)
                else:
                    run.font.size = Pt(14)
                    run.font.color.rgb = TEXT_COLOR
                    p.paragraph_format.space_after = Pt(8)

        # --- هـ. حقوق الإعداد والتطوير في الأسفل تماماً ---
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

# ==========================================
# 4. واجهة المستخدم (مع درع حماية الحصص والمؤشرات المرئية)
# ==========================================
st.markdown("<h1 style='text-align:center;'>🛡️ رادار الابتكار Pro</h1>", unsafe_allow_html=True)

# تهيئة متغيرات الجلسة للحماية والدرع
if "gate_passed" not in st.session_state:
    st.session_state.gate_passed = False
if "full_report" not in st.session_state:
    st.session_state.full_report = None

if not st.session_state.gate_passed:
    col1, col2 = st.columns([2, 1])
    with col1:
        idea_input = st.text_area("✍️ اشرح فكرتك التقنية بالتفصيل:", height=200, placeholder="مثال: نظام ذكي لصيانة الجسور باستخدام الدرون...")
    with col2:
        uploaded_file = st.file_uploader("🖼️ ارفع رسم كروكي (اختياري):", type=["jpg", "png", "jpeg"])
    
    if st.button("بدء الفحص الاستراتيجي 🚀"):
        if idea_input:
            # تشغيل السينير والاتصال هنا لحماية الحصص من الـ Reruns العشوائية
            with st.spinner("جاري تحليل البيانات والصور وتوليد التقرير المنسق..."):
                
                # صياغة البرومبت البروتوكولي المطور لاستخراج النصوص، الرقم، والجدول
                prompt = f"""
                بصفتك خبير براءات اختراع عالمي ومستشاراً استراتيجياً لابتكار القيمة وفرض السيادة التسويقية، حلل الفكرة التقنية التالية (والصورة المرفقة إن وجدت):
                " {idea_input} "

                قم بإعداد تقرير استشاري صارم ومفصل باللغة العربية، مع الالتزام التام بالبنية والبروتوكول التالي. لا تخرج عن هذه الأوسمة نهائياً، واجعل الفراغات والمسافات بين الكلمات واضحة وقابلة للقراءة:

                [===LEVEL1===] 
                ### 📊 أولاً: التشخيص الاستراتيجي والجوهر الهندسي
                1. تفكيك المنظومة التقنية: (شرح المكونات، المدخلات، العمليات، والمخرجات).
                2. السيناريو التشغيلي والديناميكي: (وصف حي خطوة بخطوة لكيفية عمل الابتكار على أرض الواقع).

                [===COMPETITORS_TABLE===]
                صغ جدول مقارنة بالمؤشرات (Markdown Table) يقارن هذه الفكرة مع أبرز 3 منافسين (إقليميين أو عالميين) بناءً على المعايير التالية حصراً كأعمدة: (المنافس | الفجوة التقنية لديه | ميزتك التنافسية الصارمة). لا تكتب أي نص خارج حدود الجدول في هذا القسم.

                [===LEVEL2===]
                ### 🔧 ثانياً: المطالبات التقنية والمراجع القانونية
                1. صياغة مسودة المطالبات التقنية (Patent Claims): صياغة هندسية دقيقة تحمي الابتكار قانونياً.
                2. الفن السابق (Prior Art): تحديد الثغرات في براءات الاختراع الحالية وكيف تتجاوزها هذه الفكرة.

                [===LEVEL3===]
                ### 🛣️ ثالثاً: الجدوى الاقتصادية وخارطة السيادة
                1. نموذج فرض السيادة: (كيف يمكن للمبتكر احتكار حصته في السوق ومنع الهندسة العكسية والتقليد السريع).
                2. خارطة الطريق التنفيذية: خطوات عملية تبدأ من النموذج الأولي (MVP) وحتى الإطلاق.

                [===AUDIT===]
                ### ⭐ رابعاً: تفصيل مؤشر الفرادة
                قم بتقييم الفكرة من 100 بناءً على الأوزان التالية واشرح سبب التقييم لكل نقطة باختصار:
                1. الجدة وعدم وجود فن سابق (30%)
                2. الخطوة الابتكارية وعدم البداهة (25%)
                3. تميز البنية الهندسية وتدفق البيانات (15%)
                4. القابلية للتطبيق وحل المشكلة التقنية (15%)
                5. صعوبة الهندسة العكسية وبناء الخندق المائي (15%)

                [===SCORE===]
                كتابة الرقم الإجمالي النهائي فقط لنسبة الفرادة (المعدل الموزون) الذي احتسبته بناءً على المعايير الخمسة أعلاه بين الأوسمة، مثال: 85
                [===/SCORE===]

                [===SOVEREIGNTY===]
                ### 💡 خامساً: توصية السيادة التسويقية المطلقة
                قدم التوصية الذهبية للمبتكر لكي "يتوقف عن التسويق التقليدي ويفرض سيادته التقنية والتجارية" بناءً على نقاط القوة المستخرجة.
                """
                
                # استدعاء الـ API وحفظ التقرير فوراً في الـ session_state
                st.session_state.full_report = call_pro_api(prompt, uploaded_file)
                st.session_state.final_idea = idea_input
                st.session_state.gate_passed = True
                st.rerun()
else:
    # المرحلة الثانية: مرحلة العرض الآمن (لا يوجد أي استدعاء للـ API هنا)
    report = st.session_state.full_report
    
    # 1. استخراج رقم نسبة الفرادة برمجياً عبر الـ Regex لتغذية شريط التقدم المرئي
    score_match = re.search(r'\[===SCORE===\]\s*(\d+)\s*\[===/SCORE===\]', report)
    if score_match:
        innovation_score = int(score_match.group(1))
    else:
        innovation_score = 75  # قيمة احتياطية
        
    # تنظيف نص التقرير المعروض من أوسمة السكور لتفادي ظهور رموز للمستخدم
    clean_report = re.sub(r'\[===SCORE===\].*?\[===/SCORE===\]', '', report, flags=re.DOTALL)
    
    # تفكيك الأجزاء بناءً على بنية البروتوكول الجديد
    parts = re.split(r'\[===LEVEL[1-3]===\]|\[===COMPETITORS_TABLE===\]|\[===AUDIT===\]|\[===SOVEREIGNTY===\]', clean_report)
    
    level1_text = parts[1] if len(parts) > 1 else report
    competitors_table = parts[2] if len(parts) > 2 else ""
    level2_text = parts[3] if len(parts) > 3 else ""
    level3_text = parts[4] if len(parts) > 4 else ""
    audit_text = parts[5] if len(parts) > 5 else ""
    sovereignty_text = parts[6] if len(parts) > 6 else ""

    # 2. عرض بطاقة مؤشر الفرادة المرئية والمتفاعلة لونياً (CSS Card)
    st.markdown("### 🎯 التقييم الفوري للابتكار")
    
    if innovation_score >= 85:
        score_color = "#10b981"  # أخضر
        score_status = "فرادة استثنائية - خندق مائي حصين 🛡️"
    elif innovation_score >= 65:
        score_color = "#f59e0b"  # برتقالي ذهبي
        score_status = "فرادة جيدة - تحتاج لتعزيز السيادة 💡"
    else:
        score_color = "#ef4444"  # أحمر
        score_status = "فرادة منخفضة - السوق مزدحم بالمنافسين ⚠️"

    st.markdown(f"""
        <div style="background-color: #f8fafc; border-right: 6px solid {score_color}; padding: 20px; border-radius: 8px; margin-bottom: 25px; text-align: right; direction: rtl;">
            <span style="font-size: 14pt; color: #64748b; font-weight: 500;">مؤشر الفرادة المحتملة للابتكار:</span>
            <div style="display: block; margin: 10px 0;">
                <span style="font-size: 36pt; font-weight: bold; color: {score_color};">{innovation_score}%</span>
                <span style="font-size: 14pt; font-weight: bold; color: #1e3a8a; margin-right: 15px;">({score_status})</span>
            </div>
            <div style="background-color: #e2e8f0; border-radius: 4px; height: 12px; width: 100%; overflow: hidden;">
                <div style="background-color: {score_color}; height: 100%; width: {innovation_score}%; border-radius: 4px;"></div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # 3. عرض التبويبات الثلاثة المتقدمة مدمج بها جدول المنافسين
    tab1, tab2, tab3 = st.tabs(["📊 التشخيص والجوهر الهندسي", "🔧 المطالبات والتحصين", "🛣️ خطة السيادة والتنفيذ"])
    
    with tab1:
        st.markdown(f"<div style='direction:rtl; text-align:right;'>{level1_text}</div>", unsafe_allow_html=True)
        if competitors_table.strip():
            st.markdown("#### 🔄 مصفوفة المقارنة الديناميكية مع أبرز 3 منافسين")
            st.markdown(f"<div style='direction:rtl; text-align:right;'>{competitors_table}</div>", unsafe_allow_html=True)
            
    with tab2:
        st.markdown(f"<div style='direction:rtl; text-align:right;'>{level2_text}</div>", unsafe_allow_html=True)
        st.divider()
        st.markdown(f"<div style='direction:rtl; text-align:right;'>{audit_text}</div>", unsafe_allow_html=True)
        
    with tab3:
        st.markdown(f"<div style='direction:rtl; text-align:right;'>{level3_text}</div>", unsafe_allow_html=True)

    st.divider()
    
    # 4. عرض توصية السيادة التسويقية المطلقة في صندوق ذهبي بارز
    st.warning(f"**💡 توصية السيادة التسويقية المطلقة (مستشارك الاستراتيجي)**\n\n{sovereignty_text.replace('### خامساً: توصية السيادة التسويقية المطلقة', '')}")

    st.divider()
    
    # 5. زر توليد وتحميل ملف الـ Word المعالج والمضبوط بالكامل
    docx_file = create_docx(report, st.session_state.final_idea)
    if docx_file:
        st.download_button(
            label="📥 تحميل التقرير الاستشاري كملف Word منسق",
            data=docx_file,
            file_name="Innovation_Sovereignty_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    if st.button("🔄 فحص ابتكار جديد"):
        st.session_state.clear()
        st.rerun()
