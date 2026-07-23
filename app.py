# -*- coding: utf-8 -*-
"""
نظام قياس الأداء واحتساب العمولات لشركات تقنية المعلومات (ERP)
================================================================
تطبيق Streamlit تفاعلي متكامل لـ:
  - إدارة قوالب Excel لسبع وظائف تشغيلية/إدارية
  - استيراد وتدقيق البيانات
  - محرك حسابات مالية يطبّق قاعدة "No Cash, No Commission"
  - لوحة KPIs ورسوم بيانية تفاعلية
  - تصدير تقرير PDF عربي احترافي (RTL) قابل للطباعة

للتشغيل:
    pip install -r requirements.txt
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta
import plotly.graph_objects as go
import plotly.express as px

# ============================================================
# 0) إعدادات الصفحة العامة + الهوية البصرية (Navy / Teal / Gold)
# ============================================================

st.set_page_config(
    page_title="نظام قياس الأداء والعمولات - ERP",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

NAVY = "#0B2545"
TEAL = "#13A89E"
GOLD = "#C9A24B"
LIGHT_BG = "#F4F6F9"
DANGER = "#C0392B"
SUCCESS = "#1E8449"


def inject_css():
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&family=Tajawal:wght@400;500;700&display=swap');

        html, body, [class*="css"]  {{
            font-family: 'Cairo', 'Tajawal', sans-serif !important;
            direction: rtl;
        }}

        .stApp {{
            background-color: {LIGHT_BG};
        }}

        /* شريط علوي */
        .app-header {{
            background: linear-gradient(90deg, {NAVY} 0%, #123a63 100%);
            padding: 22px 30px;
            border-radius: 14px;
            color: white;
            margin-bottom: 18px;
            border-bottom: 4px solid {GOLD};
        }}
        .app-header h1 {{
            margin: 0;
            font-size: 26px;
            font-weight: 800;
        }}
        .app-header p {{
            margin: 4px 0 0 0;
            color: {TEAL};
            font-size: 14px;
        }}

        /* بطاقات KPI */
        .kpi-card {{
            background: white;
            border-radius: 12px;
            padding: 16px 18px;
            box-shadow: 0 2px 8px rgba(11,37,69,0.08);
            border-right: 5px solid {TEAL};
            text-align: right;
        }}
        .kpi-card.gold {{ border-right-color: {GOLD}; }}
        .kpi-card.danger {{ border-right-color: {DANGER}; }}
        .kpi-card .kpi-label {{
            font-size: 13px;
            color: #667;
            font-weight: 600;
        }}
        .kpi-card .kpi-value {{
            font-size: 24px;
            font-weight: 800;
            color: {NAVY};
            margin-top: 4px;
        }}

        .section-title {{
            color: {NAVY};
            font-weight: 800;
            border-right: 5px solid {GOLD};
            padding-right: 10px;
            margin-top: 10px;
        }}

        .stTabs [data-baseweb="tab-list"] {{
            gap: 6px;
        }}
        .stTabs [data-baseweb="tab"] {{
            background-color: white;
            border-radius: 10px 10px 0 0;
            padding: 10px 16px;
            font-weight: 700;
            color: {NAVY};
        }}
        .stTabs [aria-selected="true"] {{
            background-color: {NAVY} !important;
            color: white !important;
        }}

        div.stButton > button {{
            background-color: {TEAL};
            color: white;
            font-weight: 700;
            border-radius: 8px;
            border: none;
        }}
        div.stButton > button:hover {{
            background-color: {NAVY};
            color: white;
        }}

        .frozen-badge {{
            background-color: #FDECEA;
            color: {DANGER};
            padding: 2px 10px;
            border-radius: 20px;
            font-weight: 700;
            font-size: 12px;
        }}
        .paid-badge {{
            background-color: #EAFAF1;
            color: {SUCCESS};
            padding: 2px 10px;
            border-radius: 20px;
            font-weight: 700;
            font-size: 12px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def kpi_card(label, value, css_class=""):
    st.markdown(
        f"""
        <div class="kpi-card {css_class}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# 1) النسب والقواعد القابلة للتعديل (Governance Settings)
# ============================================================

DEFAULT_RATES = {
    # عمولة المبيعات
    "sales_commission_rate": 3.5,      # % من قيمة العقد (3-4%)
    "sales_split_signing": 40,          # % عند التوقيع + التحصيل الأول
    "sales_split_design": 30,           # % بعد اعتماد Design Blueprinting
    "sales_split_golive": 30,           # % بعد Go-Live والتشغيل الفعلي

    # عمولة PMO
    "pmo_commission_rate": 1.25,        # % من صافي ربح المحفظة المغلقة (1-1.5%)

    # عمولة مدير المشروع PM
    "pm_commission_rate": 0.75,         # % من عقد التطبيق (0.5-1%)
    "pm_split_uat": 50,                 # % عند اجتياز UAT
    "pm_split_stable": 50,              # % بعد 30 يوم تشغيل مستقر

    # مكافأة مدير التشغيل OM
    "om_bonus_rate": 0.75,              # % ربع سنوية (0.5-1%)
    "om_rework_threshold": 9.0,         # الحد الأقصى المسموح لـ Rework Rate %
    "om_adherence_threshold": 85.0,     # الحد الأدنى لـ Adherence %

    # حافز التحصيل
    "collect_grace_days": 30,           # فترة السماح (يوم)
    "collect_late_days": 60,            # عتبة التعثر (يوم)
    "collect_current_rate": 0.175,      # % للدفعات الجارية (0.1-0.25%)
    "collect_late_rate": 1.75,          # % للدفعات المتأخرة/المتعثرة (1-2.5%)

    # مضاعف حافز الأداء (CEO Scorecard)
    "ceo_tier1_threshold": 90.0,        # إنجاز >= هذه النسبة => صرف كامل
    "ceo_tier1_payout": 100.0,
    "ceo_tier2_threshold": 80.0,        # إنجاز >= هذه النسبة (وأقل من tier1) => صرف جزئي
    "ceo_tier2_payout": 75.0,
    "ceo_below_payout": 0.0,            # أقل من tier2 => تجميد
}

ROLE_LABELS = {
    "sales": "إدارة المبيعات",
    "pmo": "مدير مكتب إدارة المشاريع (PMO)",
    "pm": "مدراء المشاريع (PM)",
    "om": "مدير التشغيل (OM)",
    "collection": "المساعد الإداري للتحصيل",
    "presales": "مهندس الحلول قبل البيع",
    "consultants": "المستشارون وفريق التنفيذ",
}

ROLE_COLUMNS = {
    "sales": ["اسم موظف المبيعات", "رقم العقد", "اسم العميل", "قيمة العقد",
              "تاريخ التوقيع", "حالة التحصيل", "اعتماد Design Blueprinting",
              "تحقق Go-Live"],
    "pmo": ["اسم مدير PMO", "اسم المحفظة/المشروع", "صافي الربح",
            "الهامش المستهدف %", "الهامش المحقق %", "حالة الإغلاق المالي والفني"],
    "pm": ["اسم مدير المشروع", "رقم المشروع", "قيمة عقد التطبيق",
           "اجتياز UAT", "تاريخ اجتياز UAT", "استقرار 30 يوم بعد التشغيل",
           "حالة التحصيل"],
    "om": ["اسم مدير التشغيل", "الربع المالي", "القيمة الأساس للمكافأة",
           "نسبة Rework %", "نسبة الالتزام بالجدولة Adherence %", "حالة التحصيل"],
    "collection": ["اسم موظف التحصيل", "رقم الفاتورة", "اسم العميل",
                    "المبلغ المحصل", "تاريخ الاستحقاق", "تاريخ التحصيل الفعلي",
                    "عدد أيام التأخير"],
    "presales": ["اسم مهندس الحلول", "عدد العروض المقدمة", "عدد الصفقات الفائزة",
                 "نسبة الاستغلال %", "تقييم رضا العميل CSAT"],
    "consultants": ["اسم المستشار", "عدد المشاريع المنفذة", "نسبة الاستغلال %",
                     "رضا العميل CSAT", "نسبة Rework %"],
}

YES_NO = ["نعم", "لا"]
COLLECT_STATUS = ["محصّل", "غير محصّل"]


def init_state():
    if "rates" not in st.session_state:
        st.session_state.rates = DEFAULT_RATES.copy()
    if "data" not in st.session_state:
        st.session_state.data = {role: None for role in ROLE_COLUMNS}
    if "ceo_achievement" not in st.session_state:
        st.session_state.ceo_achievement = 92.0


init_state()


# ============================================================
# 2) مولّدات البيانات التجريبية (Data Generators)
# ============================================================

def generate_sample_data(role, n=8, seed=42):
    rng = np.random.default_rng(seed)
    today = datetime(2026, 7, 1)

    if role == "sales":
        names = [f"مندوب مبيعات {i+1}" for i in range(n)]
        rows = []
        for i, name in enumerate(names):
            contract_value = rng.integers(150_000, 900_000)
            rows.append({
                "اسم موظف المبيعات": name,
                "رقم العقد": f"CT-{2026}{i+100}",
                "اسم العميل": f"عميل {chr(65+i)}",
                "قيمة العقد": int(contract_value),
                "تاريخ التوقيع": (today - timedelta(days=int(rng.integers(10, 240)))).date(),
                "حالة التحصيل": rng.choice(COLLECT_STATUS, p=[0.7, 0.3]),
                "اعتماد Design Blueprinting": rng.choice(YES_NO, p=[0.65, 0.35]),
                "تحقق Go-Live": rng.choice(YES_NO, p=[0.45, 0.55]),
            })
        return pd.DataFrame(rows)

    if role == "pmo":
        rows = []
        for i in range(n):
            target_margin = rng.integers(18, 28)
            achieved_margin = target_margin + rng.integers(-6, 6)
            rows.append({
                "اسم مدير PMO": "مدير مكتب إدارة المشاريع",
                "اسم المحفظة/المشروع": f"محفظة مشاريع {i+1}",
                "صافي الربح": int(rng.integers(80_000, 400_000)),
                "الهامش المستهدف %": target_margin,
                "الهامش المحقق %": achieved_margin,
                "حالة الإغلاق المالي والفني": rng.choice(["مغلق", "غير مغلق"], p=[0.75, 0.25]),
            })
        return pd.DataFrame(rows)

    if role == "pm":
        rows = []
        for i in range(n):
            uat_pass = rng.choice(YES_NO, p=[0.8, 0.2])
            rows.append({
                "اسم مدير المشروع": f"مدير مشروع {i+1}",
                "رقم المشروع": f"PRJ-{300+i}",
                "قيمة عقد التطبيق": int(rng.integers(100_000, 500_000)),
                "اجتياز UAT": uat_pass,
                "تاريخ اجتياز UAT": (today - timedelta(days=int(rng.integers(5, 90)))).date() if uat_pass == "نعم" else None,
                "استقرار 30 يوم بعد التشغيل": rng.choice(YES_NO, p=[0.6, 0.4]),
                "حالة التحصيل": rng.choice(COLLECT_STATUS, p=[0.75, 0.25]),
            })
        return pd.DataFrame(rows)

    if role == "om":
        rows = []
        for i in range(4):
            rows.append({
                "اسم مدير التشغيل": "مدير التشغيل",
                "الربع المالي": f"Q{i+1} 2026",
                "القيمة الأساس للمكافأة": int(rng.integers(400_000, 900_000)),
                "نسبة Rework %": round(rng.uniform(4, 12), 1),
                "نسبة الالتزام بالجدولة Adherence %": round(rng.uniform(78, 96), 1),
                "حالة التحصيل": rng.choice(COLLECT_STATUS, p=[0.8, 0.2]),
            })
        return pd.DataFrame(rows)

    if role == "collection":
        rows = []
        for i in range(n + 4):
            due = today - timedelta(days=int(rng.integers(10, 150)))
            delay = int(rng.integers(0, 120))
            rows.append({
                "اسم موظف التحصيل": "المساعد الإداري للتحصيل",
                "رقم الفاتورة": f"INV-{5000+i}",
                "اسم العميل": f"عميل {chr(65 + (i % 8))}",
                "المبلغ المحصل": int(rng.integers(20_000, 220_000)),
                "تاريخ الاستحقاق": due.date(),
                "تاريخ التحصيل الفعلي": (due + timedelta(days=delay)).date(),
                "عدد أيام التأخير": delay,
            })
        return pd.DataFrame(rows)

    if role == "presales":
        rows = []
        for i in range(n):
            proposals = int(rng.integers(5, 20))
            won = int(rng.integers(1, proposals))
            rows.append({
                "اسم مهندس الحلول": f"مهندس حلول {i+1}",
                "عدد العروض المقدمة": proposals,
                "عدد الصفقات الفائزة": won,
                "نسبة الاستغلال %": round(rng.uniform(70, 95), 1),
                "تقييم رضا العميل CSAT": round(rng.uniform(3.5, 5.0), 2),
            })
        return pd.DataFrame(rows)

    if role == "consultants":
        rows = []
        for i in range(n):
            rows.append({
                "اسم المستشار": f"مستشار تنفيذ {i+1}",
                "عدد المشاريع المنفذة": int(rng.integers(2, 10)),
                "نسبة الاستغلال %": round(rng.uniform(72, 92), 1),
                "رضا العميل CSAT": round(rng.uniform(3.4, 5.0), 2),
                "نسبة Rework %": round(rng.uniform(3, 11), 1),
            })
        return pd.DataFrame(rows)

    return pd.DataFrame(columns=ROLE_COLUMNS[role])


def empty_template(role):
    return pd.DataFrame(columns=ROLE_COLUMNS[role])


def get_role_df(role, auto_sample=True):
    """يرجع بيانات الدور من session_state، أو يولّد بيانات تجريبية إن لم تتوفر."""
    df = st.session_state.data.get(role)
    if df is None and auto_sample:
        df = generate_sample_data(role)
        st.session_state.data[role] = df
    return df


# ============================================================
# 3) محرك الحسابات المالية (Financial & Logic Engine)
# ============================================================

def calc_sales_commissions(df, rates):
    df = df.copy()
    rate = rates["sales_commission_rate"] / 100.0
    df["إجمالي العمولة المستحقة نظرياً"] = (df["قيمة العقد"] * rate).round(2)

    collected = df["حالة التحصيل"] == "محصّل"
    df["دفعة التوقيع/التحصيل الأول"] = np.where(
        collected, (df["إجمالي العمولة المستحقة نظرياً"] * rates["sales_split_signing"] / 100).round(2), 0.0
    )
    df["دفعة اعتماد Design Blueprinting"] = np.where(
        collected & (df["اعتماد Design Blueprinting"] == "نعم"),
        (df["إجمالي العمولة المستحقة نظرياً"] * rates["sales_split_design"] / 100).round(2), 0.0
    )
    df["دفعة Go-Live"] = np.where(
        collected & (df["تحقق Go-Live"] == "نعم"),
        (df["إجمالي العمولة المستحقة نظرياً"] * rates["sales_split_golive"] / 100).round(2), 0.0
    )
    df["إجمالي العمولة المصروفة"] = (
        df["دفعة التوقيع/التحصيل الأول"] + df["دفعة اعتماد Design Blueprinting"] + df["دفعة Go-Live"]
    )
    df["العمولة المجمّدة (بانتظار التحصيل)"] = np.where(
        ~collected, df["إجمالي العمولة المستحقة نظرياً"], df["إجمالي العمولة المستحقة نظرياً"] - df["إجمالي العمولة المصروفة"]
    ).round(2)
    df["حالة الصرف"] = np.where(~collected, "مجمّدة - بانتظار التحصيل", "قيد الصرف/مصروفة")
    return df


def calc_pmo_commissions(df, rates):
    df = df.copy()
    rate = rates["pmo_commission_rate"] / 100.0
    closed = df["حالة الإغلاق المالي والفني"] == "مغلق"
    margin_ok = df["الهامش المحقق %"] >= df["الهامش المستهدف %"]
    eligible = closed & margin_ok
    df["العمولة المستحقة"] = np.where(eligible, (df["صافي الربح"] * rate).round(2), 0.0)
    df["العمولة المجمّدة"] = np.where(~eligible, (df["صافي الربح"] * rate).round(2), 0.0)
    df["حالة الاستحقاق"] = np.where(
        eligible, "مستحقة", np.where(~closed, "مجمّدة - لم يُغلق مالياً/فنياً", "مجمّدة - لم يُحافظ على الهامش")
    )
    return df


def calc_pm_commissions(df, rates):
    df = df.copy()
    rate = rates["pm_commission_rate"] / 100.0
    df["إجمالي العمولة المستحقة نظرياً"] = (df["قيمة عقد التطبيق"] * rate).round(2)
    collected = df["حالة التحصيل"] == "محصّل"
    df["دفعة اجتياز UAT"] = np.where(
        collected & (df["اجتياز UAT"] == "نعم"),
        (df["إجمالي العمولة المستحقة نظرياً"] * rates["pm_split_uat"] / 100).round(2), 0.0
    )
    df["دفعة استقرار 30 يوم"] = np.where(
        collected & (df["استقرار 30 يوم بعد التشغيل"] == "نعم"),
        (df["إجمالي العمولة المستحقة نظرياً"] * rates["pm_split_stable"] / 100).round(2), 0.0
    )
    df["إجمالي العمولة المصروفة"] = df["دفعة اجتياز UAT"] + df["دفعة استقرار 30 يوم"]
    df["العمولة المجمّدة"] = np.where(
        ~collected, df["إجمالي العمولة المستحقة نظرياً"], df["إجمالي العمولة المستحقة نظرياً"] - df["إجمالي العمولة المصروفة"]
    ).round(2)
    return df


def calc_om_bonus(df, rates):
    df = df.copy()
    rate = rates["om_bonus_rate"] / 100.0
    collected = df["حالة التحصيل"] == "محصّل"
    rework_ok = df["نسبة Rework %"] <= rates["om_rework_threshold"]
    adherence_ok = df["نسبة الالتزام بالجدولة Adherence %"] >= rates["om_adherence_threshold"]
    eligible = collected & rework_ok & adherence_ok
    df["المكافأة المستحقة"] = np.where(eligible, (df["القيمة الأساس للمكافأة"] * rate).round(2), 0.0)
    df["المكافأة المجمّدة"] = np.where(~eligible, (df["القيمة الأساس للمكافأة"] * rate).round(2), 0.0)

    def reason(row):
        if not collected.loc[row.name]:
            return "مجمّدة - بانتظار التحصيل"
        if not rework_ok.loc[row.name]:
            return "مجمّدة - تجاوز حد Rework"
        if not adherence_ok.loc[row.name]:
            return "مجمّدة - Adherence أقل من الحد"
        return "مستحقة"

    df["حالة الاستحقاق"] = df.apply(reason, axis=1)
    return df


def calc_collection_incentive(df, rates):
    df = df.copy()
    grace = rates["collect_grace_days"]
    late_th = rates["collect_late_days"]
    current_rate = rates["collect_current_rate"] / 100.0
    late_rate = rates["collect_late_rate"] / 100.0

    def compute(row):
        d = row["عدد أيام التأخير"]
        if d <= grace:
            return row["المبلغ المحصل"] * current_rate, "دفعة جارية (ضمن السماح)"
        elif d >= late_th:
            return row["المبلغ المحصل"] * late_rate, "دفعة متأخرة/متعثرة"
        else:
            # منطقة بينية بين فترة السماح وعتبة التعثر: تطبق نسبة الجاري كحد أدنى
            return row["المبلغ المحصل"] * current_rate, "دفعة ضمن الفترة الوسيطة"

    results = df.apply(compute, axis=1)
    df["الحافز المستحق"] = results.apply(lambda x: round(x[0], 2))
    df["تصنيف الدفعة"] = results.apply(lambda x: x[1])
    return df


def ceo_multiplier(achievement_pct, rates):
    if achievement_pct >= rates["ceo_tier1_threshold"]:
        return rates["ceo_tier1_payout"], "إنجاز ممتاز - صرف كامل"
    elif achievement_pct >= rates["ceo_tier2_threshold"]:
        return rates["ceo_tier2_payout"], "إنجاز جيد - صرف جزئي"
    else:
        return rates["ceo_below_payout"], "أقل من الحد الأدنى - تجميد الحافز"


# ============================================================
# 4) أدوات Excel (تصدير قوالب / تصدير نهائي)
# ============================================================

def df_to_excel_bytes(sheets: dict):
    """sheets: dict{sheet_name: dataframe} -> bytes لملف Excel."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, d in sheets.items():
            safe_name = name[:31]
            d.to_excel(writer, sheet_name=safe_name, index=False)
    buffer.seek(0)
    return buffer


def validate_columns(df, role):
    expected = set(ROLE_COLUMNS[role])
    actual = set(df.columns)
    missing = expected - actual
    extra = actual - expected
    return missing, extra


# ============================================================
# 5) الواجهة الرئيسية
# ============================================================

inject_css()

st.markdown(
    """
    <div class="app-header">
        <h1>📊 نظام قياس الأداء واحتساب العمولات - شركات ERP</h1>
        <p>محرك حوكمة مالية وتشغيلية يطبّق قاعدة "لا تحصيل = لا عمولة" مع لوحات KPI تفاعلية وتقرير PDF عربي كامل</p>
    </div>
    """,
    unsafe_allow_html=True,
)

tab_home, tab_settings, tab_templates, tab_upload, tab_calc, tab_detail, tab_report = st.tabs(
    ["🏠 الرئيسية", "⚙️ إعدادات النسب", "📥 قوالب Excel", "📤 استيراد البيانات",
     "💰 محرك الحسابات", "📊 التفاصيل والرسوم", "📄 التقرير النهائي"]
)

rates = st.session_state.rates

# ---------------------------------------------------------------
# TAB: الإعدادات — تعديل كل النسب مباشرة
# ---------------------------------------------------------------
with tab_settings:
    st.markdown('<p class="section-title">⚙️ لوحة تحكم النسب والقواعد (قابلة للتعديل بالكامل)</p>', unsafe_allow_html=True)
    st.caption("عدّل أي نسبة أو عتبة هنا، وستنعكس فوراً على كل الحسابات في التطبيق.")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**عمولة المبيعات**")
        rates["sales_commission_rate"] = st.number_input("نسبة عمولة المبيعات %", 0.0, 20.0, rates["sales_commission_rate"], 0.1)
        rates["sales_split_signing"] = st.number_input("دفعة التوقيع/التحصيل الأول %", 0, 100, rates["sales_split_signing"])
        rates["sales_split_design"] = st.number_input("دفعة اعتماد Design Blueprinting %", 0, 100, rates["sales_split_design"])
        rates["sales_split_golive"] = st.number_input("دفعة Go-Live %", 0, 100, rates["sales_split_golive"])
        total_split = rates["sales_split_signing"] + rates["sales_split_design"] + rates["sales_split_golive"]
        if total_split != 100:
            st.warning(f"⚠️ مجموع نسب التجزئة الحالي {total_split}% (يفضّل أن يساوي 100%)")

        st.markdown("**عمولة PMO**")
        rates["pmo_commission_rate"] = st.number_input("نسبة عمولة PMO %", 0.0, 20.0, rates["pmo_commission_rate"], 0.05)

    with c2:
        st.markdown("**عمولة مدير المشروع PM**")
        rates["pm_commission_rate"] = st.number_input("نسبة عمولة PM %", 0.0, 20.0, rates["pm_commission_rate"], 0.05)
        rates["pm_split_uat"] = st.number_input("دفعة اجتياز UAT %", 0, 100, rates["pm_split_uat"])
        rates["pm_split_stable"] = st.number_input("دفعة استقرار 30 يوم %", 0, 100, rates["pm_split_stable"])

        st.markdown("**مكافأة مدير التشغيل OM**")
        rates["om_bonus_rate"] = st.number_input("نسبة مكافأة OM %", 0.0, 20.0, rates["om_bonus_rate"], 0.05)
        rates["om_rework_threshold"] = st.number_input("الحد الأقصى لـ Rework Rate %", 0.0, 100.0, rates["om_rework_threshold"], 0.5)
        rates["om_adherence_threshold"] = st.number_input("الحد الأدنى لـ Adherence %", 0.0, 100.0, rates["om_adherence_threshold"], 0.5)

    with c3:
        st.markdown("**حافز التحصيل**")
        rates["collect_grace_days"] = st.number_input("فترة السماح (يوم)", 0, 365, rates["collect_grace_days"])
        rates["collect_late_days"] = st.number_input("عتبة التعثر (يوم)", 0, 365, rates["collect_late_days"])
        rates["collect_current_rate"] = st.number_input("نسبة الحافز - دفعات جارية %", 0.0, 20.0, rates["collect_current_rate"], 0.01)
        rates["collect_late_rate"] = st.number_input("نسبة الحافز - دفعات متأخرة/متعثرة %", 0.0, 20.0, rates["collect_late_rate"], 0.05)

        st.markdown("**مضاعف بطاقة الأداء (CEO Scorecard)**")
        rates["ceo_tier1_threshold"] = st.number_input("عتبة الصرف الكامل (إنجاز %) ≥", 0.0, 100.0, rates["ceo_tier1_threshold"])
        rates["ceo_tier1_payout"] = st.number_input("نسبة الصرف عند التميز %", 0.0, 100.0, rates["ceo_tier1_payout"])
        rates["ceo_tier2_threshold"] = st.number_input("عتبة الصرف الجزئي (إنجاز %) ≥", 0.0, 100.0, rates["ceo_tier2_threshold"])
        rates["ceo_tier2_payout"] = st.number_input("نسبة الصرف الجزئي %", 0.0, 100.0, rates["ceo_tier2_payout"])

    st.session_state.rates = rates

    if st.button("↺ إعادة تعيين كل النسب للقيم الافتراضية"):
        st.session_state.rates = DEFAULT_RATES.copy()
        st.rerun()

# ---------------------------------------------------------------
# TAB: القوالب — تصدير قوالب Excel فارغة لكل وظيفة
# ---------------------------------------------------------------
with tab_templates:
    st.markdown('<p class="section-title">📥 تصدير قوالب Excel لكل وظيفة</p>', unsafe_allow_html=True)
    st.caption("نزّل القالب الفارغ الخاص بكل وظيفة، عبّئه بالبيانات الفعلية، ثم ارفعه من تبويب «استيراد البيانات».")

    cols = st.columns(4)
    for i, (role, label) in enumerate(ROLE_LABELS.items()):
        with cols[i % 4]:
            st.markdown(f"**{label}**")
            template_df = empty_template(role)
            buf = df_to_excel_bytes({label: template_df})
            st.download_button(
                label="⬇️ تنزيل القالب",
                data=buf,
                file_name=f"قالب_{role}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"tpl_{role}",
            )

    st.divider()
    st.markdown('<p class="section-title">📦 تصدير حزمة كل القوالب دفعة واحدة</p>', unsafe_allow_html=True)
    all_templates = {ROLE_LABELS[r]: empty_template(r) for r in ROLE_COLUMNS}
    buf_all = df_to_excel_bytes(all_templates)
    st.download_button(
        "⬇️ تنزيل جميع القوالب (ملف واحد بـ 7 أوراق عمل)",
        data=buf_all,
        file_name="جميع_القوالب_7_وظائف.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------------------------------------------
# TAB: استيراد البيانات — رفع وتدقيق الملفات
# ---------------------------------------------------------------
with tab_upload:
    st.markdown('<p class="section-title">📤 استيراد ملفات Excel المكتملة</p>', unsafe_allow_html=True)
    st.caption("يمكنك رفع بيانات فعلية لكل وظيفة، أو الاكتفاء بالبيانات التجريبية المولّدة تلقائياً لتجربة النظام.")

    for role, label in ROLE_LABELS.items():
        with st.expander(f"📁 {label}", expanded=False):
            uploaded = st.file_uploader(f"رفع ملف بيانات: {label}", type=["xlsx", "xls"], key=f"up_{role}")
            colA, colB = st.columns([1, 1])
            with colA:
                if uploaded is not None:
                    try:
                        new_df = pd.read_excel(uploaded)
                        missing, extra = validate_columns(new_df, role)
                        if missing:
                            st.error(f"❌ أعمدة ناقصة في الملف: {', '.join(missing)}")
                        else:
                            if extra:
                                st.warning(f"⚠️ أعمدة إضافية غير معروفة سيتم تجاهلها: {', '.join(extra)}")
                            st.session_state.data[role] = new_df[ROLE_COLUMNS[role]]
                            st.success(f"✅ تم استيراد {len(new_df)} صف بنجاح لوظيفة «{label}»")
                    except Exception as e:
                        st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
            with colB:
                if st.button(f"🎲 استخدام بيانات تجريبية لـ «{label}»", key=f"sample_{role}"):
                    st.session_state.data[role] = generate_sample_data(role)
                    st.success("تم توليد بيانات تجريبية.")

            current = get_role_df(role, auto_sample=False)
            if current is not None:
                st.dataframe(current, use_container_width=True, height=200)
            else:
                st.info("لا توجد بيانات محمّلة بعد لهذه الوظيفة.")

# جلب / تجهيز كل الجداول (مع توليد تجريبي إن لزم) لبقية التبويبات
sales_df = get_role_df("sales")
pmo_df = get_role_df("pmo")
pm_df = get_role_df("pm")
om_df = get_role_df("om")
collection_df = get_role_df("collection")
presales_df = get_role_df("presales")
consultants_df = get_role_df("consultants")

# تطبيق محرك الحسابات
sales_calc = calc_sales_commissions(sales_df, rates)
pmo_calc = calc_pmo_commissions(pmo_df, rates)
pm_calc = calc_pm_commissions(pm_df, rates)
om_calc = calc_om_bonus(om_df, rates)
collection_calc = calc_collection_incentive(collection_df, rates)

# ---------------------------------------------------------------
# TAB: محرك الحسابات — عرض الجداول المحسوبة
# ---------------------------------------------------------------
with tab_calc:
    st.markdown('<p class="section-title">💰 نتائج محرك الحسابات (تطبيق فوري لقاعدة "لا تحصيل = لا عمولة")</p>', unsafe_allow_html=True)

    st.markdown("**عمولات المبيعات**")
    st.dataframe(sales_calc, use_container_width=True)

    st.markdown("**عمولة PMO**")
    st.dataframe(pmo_calc, use_container_width=True)

    st.markdown("**عمولة مدراء المشاريع (PM)**")
    st.dataframe(pm_calc, use_container_width=True)

    st.markdown("**مكافأة مدير التشغيل (OM) - ربع سنوية**")
    st.dataframe(om_calc, use_container_width=True)

    st.markdown("**حافز التحصيل**")
    st.dataframe(collection_calc, use_container_width=True)

    st.divider()
    st.markdown('<p class="section-title">🎯 مضاعف بطاقة الأداء (CEO Incentive Multiplier)</p>', unsafe_allow_html=True)
    st.session_state.ceo_achievement = st.slider(
        "نسبة إنجاز المؤشرات الموزونة الإجمالية %", 0.0, 100.0, st.session_state.ceo_achievement, 0.5
    )
    payout_pct, reason = ceo_multiplier(st.session_state.ceo_achievement, rates)
    c1, c2 = st.columns(2)
    with c1:
        kpi_card("نسبة صرف الحافز المستحقة", f"{payout_pct:.0f}%", "gold" if payout_pct == 100 else ("" if payout_pct > 0 else "danger"))
    with c2:
        kpi_card("الحالة", reason, "" if payout_pct > 0 else "danger")

# ============================================================
# حسابات موحّدة لبطاقات KPI الرئيسية
# ============================================================
total_sales_value = sales_calc["قيمة العقد"].sum()
total_collected_cash = collection_calc["المبلغ المحصل"].sum()

total_commission_paid = (
    sales_calc["إجمالي العمولة المصروفة"].sum()
    + pmo_calc["العمولة المستحقة"].sum()
    + pm_calc["إجمالي العمولة المصروفة"].sum()
    + om_calc["المكافأة المستحقة"].sum()
    + collection_calc["الحافز المستحق"].sum()
)
total_commission_frozen = (
    sales_calc["العمولة المجمّدة (بانتظار التحصيل)"].sum()
    + pmo_calc["العمولة المجمّدة"].sum()
    + pm_calc["العمولة المجمّدة"].sum()
    + om_calc["المكافأة المجمّدة"].sum()
)

avg_dso = collection_calc["عدد أيام التأخير"].mean() if len(collection_calc) else 0
avg_rework = pd.concat([om_calc["نسبة Rework %"], consultants_df["نسبة Rework %"]]).mean()
avg_csat = pd.concat([presales_df["تقييم رضا العميل CSAT"], consultants_df["رضا العميل CSAT"]]).mean()
avg_utilization = pd.concat([presales_df["نسبة الاستغلال %"], consultants_df["نسبة الاستغلال %"]]).mean()
avg_adherence = om_calc["نسبة الالتزام بالجدولة Adherence %"].mean()

# ---------------------------------------------------------------
# TAB: الرئيسية — بطاقات KPI + رسوم رئيسية
# ---------------------------------------------------------------
with tab_home:
    st.markdown('<p class="section-title">🔢 المؤشرات الرئيسية (KPIs)</p>', unsafe_allow_html=True)
    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    with r1c1:
        kpi_card("إجمالي المبيعات (قيمة العقود)", f"{total_sales_value:,.0f} ر.س")
    with r1c2:
        kpi_card("التدفقات النقدية المحصلة", f"{total_collected_cash:,.0f} ر.س", "gold")
    with r1c3:
        kpi_card("إجمالي العمولات المستحقة/المصروفة", f"{total_commission_paid:,.0f} ر.س")
    with r1c4:
        kpi_card("العمولات المجمّدة (بانتظار التحصيل)", f"{total_commission_frozen:,.0f} ر.س", "danger")

    r2c1, r2c2, r2c3 = st.columns(3)
    with r2c1:
        kpi_card("متوسط أيام التحصيل (DSO)", f"{avg_dso:,.1f} يوم")
    with r2c2:
        kpi_card("متوسط نسبة Rework", f"{avg_rework:,.1f}%")
    with r2c3:
        kpi_card("متوسط رضا العملاء (CSAT)", f"{avg_csat:,.2f} / 5")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="section-title">📈 نظرة عامة على العمولات حسب الوظيفة</p>', unsafe_allow_html=True)

    overview_df = pd.DataFrame({
        "الوظيفة": ["المبيعات", "PMO", "مدير المشروع PM", "مدير التشغيل OM", "التحصيل"],
        "مصروفة/مستحقة": [
            sales_calc["إجمالي العمولة المصروفة"].sum(),
            pmo_calc["العمولة المستحقة"].sum(),
            pm_calc["إجمالي العمولة المصروفة"].sum(),
            om_calc["المكافأة المستحقة"].sum(),
            collection_calc["الحافز المستحق"].sum(),
        ],
        "مجمّدة": [
            sales_calc["العمولة المجمّدة (بانتظار التحصيل)"].sum(),
            pmo_calc["العمولة المجمّدة"].sum(),
            pm_calc["العمولة المجمّدة"].sum(),
            om_calc["المكافأة المجمّدة"].sum(),
            0,
        ],
    })

    fig_overview = go.Figure()
    fig_overview.add_trace(go.Bar(name="مصروفة/مستحقة", x=overview_df["الوظيفة"], y=overview_df["مصروفة/مستحقة"], marker_color=TEAL))
    fig_overview.add_trace(go.Bar(name="مجمّدة (بانتظار التحصيل)", x=overview_df["الوظيفة"], y=overview_df["مجمّدة"], marker_color=DANGER))
    fig_overview.update_layout(
        barmode="stack", template="plotly_white",
        font=dict(family="Cairo", size=13), legend=dict(orientation="h", y=1.15),
        margin=dict(t=30, b=10),
    )
    st.plotly_chart(fig_overview, use_container_width=True)

# ---------------------------------------------------------------
# TAB: التفاصيل والرسوم — Gauge + توزيعات
# ---------------------------------------------------------------
with tab_detail:
    st.markdown('<p class="section-title">📊 مقاييس الأداء (Gauges)</p>', unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        fig_util = go.Figure(go.Indicator(
            mode="gauge+number",
            value=avg_utilization,
            title={"text": "معدل استغلال الكوادر %"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": TEAL},
                "steps": [
                    {"range": [0, 80], "color": "#FDECEA"},
                    {"range": [80, 88], "color": "#EAFAF1"},
                    {"range": [88, 100], "color": "#FEF5E7"},
                ],
                "threshold": {"line": {"color": GOLD, "width": 4}, "value": 88},
            },
        ))
        fig_util.update_layout(font=dict(family="Cairo"), margin=dict(t=40, b=10))
        st.plotly_chart(fig_util, use_container_width=True)

    with g2:
        fig_spi = go.Figure(go.Indicator(
            mode="gauge+number",
            value=avg_adherence,
            title={"text": "مؤشر أداء الجدولة SPI (تقريبي عبر Adherence) %"},
            gauge={
                "axis": {"range": [0, 120]},
                "bar": {"color": NAVY},
                "steps": [
                    {"range": [0, 85], "color": "#FDECEA"},
                    {"range": [85, 100], "color": "#EAFAF1"},
                    {"range": [100, 120], "color": "#FEF5E7"},
                ],
                "threshold": {"line": {"color": GOLD, "width": 4}, "value": 85},
            },
        ))
        fig_spi.update_layout(font=dict(family="Cairo"), margin=dict(t=40, b=10))
        st.plotly_chart(fig_spi, use_container_width=True)

    st.divider()
    st.markdown('<p class="section-title">📉 مراحل صرف عمولات المبيعات</p>', unsafe_allow_html=True)
    stages_df = pd.DataFrame({
        "المرحلة": ["التوقيع/التحصيل الأول", "اعتماد Design Blueprinting", "Go-Live"],
        "المبلغ": [
            sales_calc["دفعة التوقيع/التحصيل الأول"].sum(),
            sales_calc["دفعة اعتماد Design Blueprinting"].sum(),
            sales_calc["دفعة Go-Live"].sum(),
        ],
    })
    fig_stages = px.bar(stages_df, x="المرحلة", y="المبلغ", color="المرحلة",
                         color_discrete_sequence=[NAVY, TEAL, GOLD], text_auto=".2s")
    fig_stages.update_layout(template="plotly_white", font=dict(family="Cairo"), showlegend=False, margin=dict(t=20))
    st.plotly_chart(fig_stages, use_container_width=True)

    st.divider()
    st.markdown('<p class="section-title">🥧 توزيع العمولات حسب حالة التحصيل</p>', unsafe_allow_html=True)
    dist_df = pd.DataFrame({
        "الحالة": ["مصروفة/مستحقة", "مجمّدة (بانتظار التحصيل)"],
        "القيمة": [total_commission_paid, total_commission_frozen],
    })
    fig_pie = px.pie(dist_df, names="الحالة", values="القيمة", color="الحالة",
                      color_discrete_map={"مصروفة/مستحقة": TEAL, "مجمّدة (بانتظار التحصيل)": DANGER}, hole=0.45)
    fig_pie.update_layout(font=dict(family="Cairo"), margin=dict(t=20))
    st.plotly_chart(fig_pie, use_container_width=True)

    st.divider()
    st.markdown('<p class="section-title">📥 تصدير كل البيانات المحسوبة (Excel)</p>', unsafe_allow_html=True)
    final_sheets = {
        "المبيعات": sales_calc, "PMO": pmo_calc, "مدير المشروع PM": pm_calc,
        "مدير التشغيل OM": om_calc, "التحصيل": collection_calc,
        "قبل البيع": presales_df, "الاستشاريون": consultants_df,
    }
    final_buf = df_to_excel_bytes(final_sheets)
    st.download_button("⬇️ تنزيل التقرير المحاسبي الكامل (Excel)", data=final_buf,
                        file_name="التقرير_المحاسبي_الشامل.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------
# TAB: التقرير النهائي (PDF)
# ---------------------------------------------------------------
with tab_report:
    st.markdown('<p class="section-title">📄 التقرير النهائي - معاينة وتصدير PDF</p>', unsafe_allow_html=True)

    report_date = st.date_input("تاريخ التقرير", datetime(2026, 7, 21))
    company_name = st.text_input("اسم الشركة", "شركة الحلول التقنية المتكاملة (ERP)")
    period_label = st.text_input("الفترة المالية المشمولة بالتقرير", "الربع الثالث 2026")

    def build_report_html():
        rows_html = ""
        role_summaries = [
            ("إدارة المبيعات", sales_calc["إجمالي العمولة المصروفة"].sum(), sales_calc["العمولة المجمّدة (بانتظار التحصيل)"].sum()),
            ("PMO", pmo_calc["العمولة المستحقة"].sum(), pmo_calc["العمولة المجمّدة"].sum()),
            ("مدير المشروع PM", pm_calc["إجمالي العمولة المصروفة"].sum(), pm_calc["العمولة المجمّدة"].sum()),
            ("مدير التشغيل OM", om_calc["المكافأة المستحقة"].sum(), om_calc["المكافأة المجمّدة"].sum()),
            ("التحصيل", collection_calc["الحافز المستحق"].sum(), 0),
        ]
        for name, paid, frozen in role_summaries:
            rows_html += f"""
            <tr>
                <td>{name}</td>
                <td>{paid:,.0f} ر.س</td>
                <td>{frozen:,.0f} ر.س</td>
            </tr>
            """

        payout_pct, reason = ceo_multiplier(st.session_state.ceo_achievement, rates)

        html = f"""
        <html dir="rtl" lang="ar">
        <head>
        <meta charset="utf-8">
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
            body {{ font-family: 'Cairo', sans-serif; direction: rtl; color: #1a1a1a; margin: 30px; }}
            .header {{ background: {NAVY}; color: white; padding: 20px 26px; border-radius: 10px;
                       border-bottom: 5px solid {GOLD}; margin-bottom: 20px; }}
            .header h1 {{ margin: 0; font-size: 22px; }}
            .header p {{ margin: 4px 0 0 0; font-size: 13px; color: {TEAL}; }}
            h2 {{ color: {NAVY}; border-right: 5px solid {GOLD}; padding-right: 10px; font-size: 16px; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 13px; }}
            th {{ background: {NAVY}; color: white; padding: 8px; text-align: center; }}
            td {{ padding: 7px; text-align: center; border-bottom: 1px solid #ddd; }}
            tr:nth-child(even) {{ background: #F4F6F9; }}
            .scorecard {{ background: #FEF9EF; border: 1px solid {GOLD}; border-radius: 8px; padding: 14px; margin-bottom: 20px;}}
            .sign-box {{ display: flex; justify-content: space-between; margin-top: 40px; }}
            .sign-item {{ width: 45%; text-align: center; border-top: 1px solid #333; padding-top: 8px; }}
            @media print {{ body {{ margin: 10mm; }} }}
        </style>
        </head>
        <body>
            <div class="header">
                <h1>{company_name}</h1>
                <p>تقرير الأداء والعمولات التشغيلية | الفترة: {period_label} | تاريخ الإصدار: {report_date}</p>
            </div>

            <h2>ملخص الأداء المالي والعمولات (مستحقة / مصروفة / مجمّدة)</h2>
            <table>
                <tr><th>الوظيفة</th><th>مبالغ مستحقة/مصروفة</th><th>مبالغ مجمّدة (بانتظار التحصيل)</th></tr>
                {rows_html}
                <tr style="font-weight:bold;">
                    <td>الإجمالي</td>
                    <td>{total_commission_paid:,.0f} ر.س</td>
                    <td>{total_commission_frozen:,.0f} ر.س</td>
                </tr>
            </table>

            <h2>بطاقة الأداء المتوازن (Balanced Scorecard)</h2>
            <div class="scorecard">
                <p><b>نسبة إنجاز المؤشرات الموزونة الإجمالية:</b> {st.session_state.ceo_achievement:.1f}%</p>
                <p><b>نسبة صرف الحافز المستحقة:</b> {payout_pct:.0f}% — {reason}</p>
                <p><b>متوسط أيام التحصيل (DSO):</b> {avg_dso:,.1f} يوم &nbsp;|&nbsp;
                   <b>متوسط Rework:</b> {avg_rework:,.1f}% &nbsp;|&nbsp;
                   <b>متوسط CSAT:</b> {avg_csat:,.2f}/5 &nbsp;|&nbsp;
                   <b>معدل الاستغلال:</b> {avg_utilization:,.1f}%</p>
            </div>

            <h2>قسم الاعتمادات والتوقيعات</h2>
            <div class="sign-box">
                <div class="sign-item">إعداد الموارد البشرية<br><br>الاسم / التوقيع / التاريخ</div>
                <div class="sign-item">اعتماد المدير التنفيذي (CEO)<br><br>الاسم / التوقيع / التاريخ</div>
            </div>
        </body>
        </html>
        """
        return html

    report_html = build_report_html()

    st.markdown("**معاينة التقرير:**")
    st.components.v1.html(report_html, height=650, scrolling=True)

    st.divider()
    c1, c2 = st.columns(2)

    with c1:
        st.download_button(
            "⬇️ تنزيل التقرير بصيغة HTML (للطباعة المباشرة عبر المتصفح Ctrl+P)",
            data=report_html.encode("utf-8"),
            file_name="التقرير_النهائي.html",
            mime="text/html",
        )

    with c2:
        try:
            from weasyprint import HTML
            pdf_bytes = HTML(string=report_html).write_pdf()
            st.download_button(
                "⬇️ تنزيل التقرير بصيغة PDF (عربي كامل)",
                data=pdf_bytes,
                file_name="التقرير_النهائي.pdf",
                mime="application/pdf",
            )
        except Exception as e:
            st.warning(
                "⚠️ توليد PDF مباشرة يتطلب تثبيت WeasyPrint ومكتباته الجهازية "
                "(Pango, Cairo, GDK-Pixbuf). استخدم زر تصدير HTML كبديل فوري، "
                "أو راجع ملف README.md لتفاصيل التثبيت.\n\n"
                f"تفاصيل الخطأ: {e}"
            )
