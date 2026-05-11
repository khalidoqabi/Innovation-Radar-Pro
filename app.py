import streamlit as st
import google.generativeai as genai
import re

# --- 1. إعدادات الهوية البصرية المتطورة ---
st.set_page_config(page_title="Innovation Radar Pro | رادار الابتكار الاحترافي", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700;900&display=swap');
    html, body, [data-testid="stSidebar"], .stApp {
        direction: rtl !important; text-align: right !important; font-family: 'Tajawal', sans-serif !important;
    }
    .pro-header {
        background: linear-gradient(90deg, #1e3a8a, #3b82f6, #10b981);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        font-size: 3em; font-weight: 900; text-align: center; margin-bottom: 10px;
    }
    .score-box {
        background: #f0fdf4; border: 2px solid #10b981; padding: 20px;
        border-radius: 15px; text-align: center; font-size: 1.5em; font-weight: bold;
    }
    .sovereignty-box {
        background: #fff7ed; border-right: 5px solid #ea580c; padding: 15px;
        border-radius: 10px; font-style: italic; color: #9a3412;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. إدارة الحالة ---
if "gate_passed" not in st.session_state: st.session_state.gate_passed = False
if "full_report" not in st.session_state: st.session_state.full_report = None
if "final_idea" not in st.session_state: st.session_state.final_idea = ""

# --- 3. محرك الاتصال بـ API المطور (apifreellm) ---
def call_pro_api(prompt):
    try:
        # تأكد من وضع المفتاح في Secrets باسم PRO_API_KEY
        api_key = st.secrets["PRO_API_KEY"]
        genai.configure(api_key=api_key)
        
        # استخدام موديل flash للسرعة أو pro للدقة العالية
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        response = model.generate_content(prompt)
        
        if response.text:
            return response.text
        else:
            return "عذراً، لم يتمكن المحرك من توليد استجابة."
    except Exception as e:
        return f"حدث خطأ في الاتصال بمحرك جوجل: {str(e)}"

# --- 4. منطق التقرير المطور (المعايير الجديدة) ---
def generate_pro_report(idea):
    pro_prompt = f"""
    أنت مستشار ابتكار وخبير براءات اختراع دولي. حلل الفكرة التالية بعمق: "{idea}"
    
    قم بصياغة تقرير احترافي جداً باللغة العربية مقسماً كالتالي:
    [===LEVEL1===] التشخيص الاستراتيجي: تحليل السوق والمنافسين المحتملين.
    [===LEVEL2===] المطالبات التقنية والمراجع: اذكر (أقرب 3 اختراعات/براءات اختراع عالمية) تشبه الفكرة وروابطها المنطقية.
    [===LEVEL3===] الجدوى الاقتصادية وخارطة الطريق نحو التصنيع.
    
    [===AUDIT===] تحليل الفرادة:
    أعطِ درجة فرادة مئوية (%) بناءً على (الجدة، الخطوة الابتكارية، القابلية للتصنيع).
    
    [===SOVEREIGNTY===] توصية السيادة التسويقية:
    نصيحة استراتيجية من وجهة نظر "السيادة التسويقية" لضمان حماية الحصة السوقية.
    """
    return call_pro_api(pro_prompt)

# --- 5. واجهة المستخدم ---
st.markdown("<h1 class='pro-header'>🛡️ رادار الابتكار Pro</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;'>نسخة السيادة التسويقية والبحث المرجعي</p>", unsafe_allow_html=True)

if not st.session_state.gate_passed:
    st.info("أهلاً بك يا دكتور خالد. هذه النسخة مجهزة لتقديم تحليل 'أقرب الاختراعات' ونسبة الفرادة.")
    idea_input = st.text_area("اشرح فكرتك هنا بالتفصيل التقني...", height=150)
    
    if st.button("بدء الفحص الاستراتيجي 🚀"):
        if idea_input:
            with st.spinner("جاري التحليل المعمق..."):
                # محاكاة المحقق الذكي
                st.session_state.final_idea = idea_input
                st.session_state.gate_passed = True
                st.rerun()
        else:
            st.warning("يرجى إدخال تفاصيل الفكرة أولاً.")

else:
    if st.session_state.full_report is None:
        with st.spinner("جاري التواصل مع قواعد البيانات العالمية وتوليد التقرير..."):
            st.session_state.full_report = generate_pro_report(st.session_state.final_idea)
    
    # تقسيم النتائج
    parts = re.split(r'\[===LEVEL[1-3]===\]|\[===AUDIT===\]|\[===SOVEREIGNTY===\]', st.session_state.full_report)
    
    tab1, tab2, tab3 = st.tabs(["📊 التشخيص", "🔧 المطالبات والمراجع", "🛣️ الطريق للتنفيذ"])
    
    with tab1: st.write(parts[1] if len(parts) > 1 else st.session_state.full_report)
    with tab2: st.write(parts[2] if len(parts) > 2 else "بحث المراجع قيد المراجعة.")
    with tab3: st.write(parts[3] if len(parts) > 3 else "خارطة الطريق قيد التجهيز.")

    # عرض نسبة الفرادة وتوصية السيادة في الأسفل
    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("<div class='score-box'>⭐ مؤشر الفرادة والابتكار<br>" + (parts[4] if len(parts) > 4 else "N/A") + "</div>", unsafe_allow_html=True)
    with col2:
        st.markdown("<div class='sovereignty-box'>💡 توصية السيادة التسويقية:<br>" + (parts[5] if len(parts) > 5 else "تأكد من حماية الملكية الفكرية.") + "</div>", unsafe_allow_html=True)

    st.text_area("انسخ التقرير النهائي:", value=st.session_state.full_report, height=100)
    if st.button("فحص جديد"):
        st.session_state.clear()
        st.rerun()
