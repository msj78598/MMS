# -*- coding: utf-8 -*-
# mms_premise_tracker.py — Premise key + robust LastDaily + safe open/closed + Excel export

import re
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="MMS | Premise Tracker", layout="wide")

# ===================== Helpers =====================
def norm_col(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip()).lower()

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """يبحث عن العمود بالاسم الدقيق ثم بالاحتواء الجزئي (AR/EN)."""
    if df is None or df.empty:
        return None
    nm = {norm_col(c): c for c in df.columns}
    for c in candidates:
        if norm_col(c) in nm:
            return nm[norm_col(c)]
    for c in df.columns:
        for cand in candidates:
            if norm_col(cand) in norm_col(c):
                return c
    return None

def smart_parse_datetime(series: pd.Series, excel_origin: str = "1899-12-30") -> pd.Series:
    """تحويل مختلط (نصي/سيريال Excel) إلى datetime؛ يدعم dayfirst وExcel 1900/1904."""
    if series is None:
        return pd.Series([], dtype="datetime64[ns]")

    s = series.copy()

    def clean(x):
        if pd.isna(x): return np.nan
        x = str(x).strip()
        if x == "" or x.lower() in {"none", "nan", "null", "-", "—", "0"}: return np.nan
        # مثل 0000-00-00
        if re.fullmatch(r"0{2,}[-/:]0{2,}[-/:]0{2,}.*", x): return np.nan
        return x

    s = s.map(clean)

    # محاولة أولى (dayfirst=True)
    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)

    # محاولة ثانية (dayfirst=False) لما تفشل الأولى
    need2 = parsed.isna()
    if need2.any():
        parsed.loc[need2] = pd.to_datetime(s[need2], errors="coerce", dayfirst=False, infer_datetime_format=True)

    # Excel serial fallback (1900/1904 origin)
    need_excel = parsed.isna()
    if need_excel.any():
        as_num = pd.to_numeric(s.where(need_excel), errors="coerce")
        mask = as_num.notna()
        if mask.any():
            parsed.loc[mask] = pd.to_datetime(as_num[mask], unit="d", origin=excel_origin, errors="coerce")

    return parsed

def as_int0(x):
    """تحويل آمن إلى int؛ يعيد 0 عند NaN/None/غير رقمي."""
    try:
        v = pd.to_numeric(x, errors="coerce")
        if hasattr(v, "iloc"):
            v = v.iloc[0]
        if pd.isna(v):
            return 0
        return int(float(v))
    except Exception:
        return 0

def to_excel_download(df: pd.DataFrame) -> bytes:
    """إنشاء ملف Excel؛ يفضّل xlsxwriter وإن لم يوجد يستخدم openpyxl."""
    bio = BytesIO()
    try:
        with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Results")
    except ModuleNotFoundError:
        with pd.ExcelWriter(bio, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Results")
    return bio.getvalue()

# ===================== UI: Uploads & Settings =====================
st.title("📊 نظام تتبع العدادات غير المتصلة — Premise Tracker")

with st.sidebar:
    st.header("📁 ملفات الإدخال")
    disconnected_file = st.file_uploader("ملف العدادات غير المتصلة", type=["xlsx", "xls"])
    insp_files  = st.file_uploader("ملفات مهام الفحص (0..N)",  type=["xlsx", "xls"], accept_multiple_files=True)
    maint_files = st.file_uploader("ملفات مهام الصيانة (0..N)", type=["xlsx", "xls"], accept_multiple_files=True)

    st.markdown("---")
    st.header("⚙️ إعدادات التحويل")
    excel_origin = st.selectbox("Excel Origin (للتواريخ الرقمية)", ["1899-12-30", "1904-01-01"], index=0)

    st.markdown("—")
    st.header("🔒 حالات تعتبر (مقفلة)")
    default_closed_terms = """
closed, complete, completed, done, resolved,
مغلق, مغلقة, مقفلة, مقفل, منجز, منجزة, منتهية, منتهي, تمت المعالجة
""".strip()
    closed_terms_input = st.text_area("قائمة مفصولة بفواصل (يمكن تعديلها)", value=default_closed_terms, height=90)
    CLOSED_TERMS = {w.strip().lower() for w in closed_terms_input.split(",") if w.strip()}

    st.markdown("---")
    start_btn = st.button("🚀 ابدأ التحليل")

if not start_btn or not disconnected_file:
    st.info("⬆️ ارفع الملفات ثم اضغط **ابدأ التحليل**.")
    st.stop()

# ===================== Read Disconnected (Premise key) =====================
st.subheader("📘 قراءة ملف العدادات غير المتصلة (Premise = المفتاح)")

dis_df = pd.read_excel(disconnected_file)

PREMISE_CANDS_DIS = ["Utility Site Id", "Premise", "رقم المكان"]
LAST_CANDS = [
    "Last Daily", "LastDaily", "Last Communication", "Last Comm",
    "آخر قراءة", "آخر اتصال", "اخر اتصال", "اخر قراءة", "Last Daily Read", "Last Daily Date"
]

premise_col_dis = pick_col(dis_df, PREMISE_CANDS_DIS)
last_daily_col  = pick_col(dis_df, LAST_CANDS)

if not premise_col_dis:
    st.error("لم يتم العثور على عمود رقم المكان في ملف غير المتصلين (توقّع: Utility Site Id).")
    st.stop()

dis_df["_KEY_PREMISE"] = dis_df[premise_col_dis].astype(str).str.strip()

# اختيار يدوي لعمود LastDaily (اختياري)
with st.expander("🔧 اختيار عمود 'آخر اتصال' يدويًا (اختياري)"):
    last_choice = st.selectbox("اختر عمود 'آخر اتصال' (إن رغبت):", options=["(اكتشاف تلقائي)"] + list(dis_df.columns), index=0)
    if last_choice != "(اكتشاف تلقائي)":
        last_daily_col = last_choice

# تحويل LastDaily
if last_daily_col and last_daily_col in dis_df.columns:
    parsed = smart_parse_datetime(dis_df[last_daily_col], excel_origin=excel_origin)
    dis_df["LastDaily"] = parsed
    ok = int(parsed.notna().sum())
    st.success(f"تحويل '{last_daily_col}': {ok}/{len(dis_df)} قيماً صالحة.")
    with st.expander("🧪 أمثلة غير قابلة للتحويل"):
        bad = dis_df.loc[parsed.isna(), [last_daily_col]].head(15)
        st.write("لا توجد أمثلة غير قابلة للتحويل ✅" if bad.empty else bad)
else:
    dis_df["LastDaily"] = pd.NaT
    st.warning("⚠️ لم يُعثر على عمود 'آخر اتصال' — سيُترك فارغًا.")

summary_base_cols = ["_KEY_PREMISE", "LastDaily"]
summary_extra_cols = [c for c in dis_df.columns if c not in summary_base_cols]
summary = dis_df[summary_base_cols + summary_extra_cols].copy()

# ===================== Read Tasks (multi files) =====================
def load_task_files(files, kind_label: str) -> pd.DataFrame:
    """توحيد مخطط ملفات الفحص/الصيانة باستخدام Premise كمفتاح."""
    if not files:
        return pd.DataFrame()

    frames = []
    for f in files:
        df = pd.read_excel(f)

        premise_col = pick_col(df, ["Premise", "Utility Site Id", "رقم المكان"])
        reg_col     = pick_col(df, ["Task Registration Date Time", "Request Registration Date Time",
                                    "تاريخ التسجيل", "تاريخ تسجيل المهمة", "تاريخ تسجيل الطلب"])
        close_col   = pick_col(df, ["Task Closed Time", "Task Completed Time",
                                    "تاريخ الإقفال", "وقت إقفال المهمة"])
        status_col  = pick_col(df, ["Task Status", "Request Status", "الحالة",
                                    "حالة المهمة", "حالة الطلب"])
        result_col  = pick_col(df, ["Technician's Result", "Final Result",
                                    "Final Result (Dispatcher's Result)", "النتيجة النهائية", "نتيجة الفني"])

        if not premise_col:
            st.warning(f"تجاهل الملف '{getattr(f,'name','file')}' لعدم وجود Premise/Utility Site Id.")
            continue

        tmp = pd.DataFrame()
        tmp["_KEY_PREMISE"] = df[premise_col].astype(str).str.strip()
        tmp["reg_time"]   = smart_parse_datetime(df[reg_col],   excel_origin=excel_origin) if (reg_col   and reg_col   in df.columns) else pd.NaT
        tmp["close_time"] = smart_parse_datetime(df[close_col], excel_origin=excel_origin) if (close_col and close_col in df.columns) else pd.NaT
        tmp["status"]     = df[status_col].astype(str)  if (status_col and status_col in df.columns) else np.nan
        tmp["result"]     = df[result_col].astype(str)  if (result_col and result_col in df.columns) else np.nan
        tmp["bucket"]     = kind_label
        tmp["source"]     = getattr(f, "name", kind_label)
        frames.append(tmp)

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)

insp_df  = load_task_files(insp_files,  "فحص")
maint_df = load_task_files(maint_files, "صيانة")

# ===================== Safe Summaries =====================
def summarize_tasks(df: pd.DataFrame, prefix: str, closed_terms: set[str]) -> pd.DataFrame:
    """تجميع آمن لكل Premise: إجمالي، مفتوح، آخر حالة/نتيجة/تاريخ."""
    if df is None or df.empty or "_KEY_PREMISE" not in df.columns:
        return pd.DataFrame(columns=["_KEY_PREMISE"])

    df = df.copy()
    for col in ["reg_time", "close_time", "status", "result"]:
        if col not in df.columns:
            df[col] = pd.NaT if col in ["reg_time", "close_time"] else np.nan

    status_norm = df["status"].astype(str).str.strip().str.lower()
    is_closed_by_status = status_norm.isin(closed_terms)
    is_closed_by_time   = df["close_time"].notna()
    df["_is_open"] = ~(is_closed_by_status | is_closed_by_time)

    df["_latest_date"] = df["close_time"].where(df["close_time"].notna(), df["reg_time"])

    latest = (
        df.sort_values(["_KEY_PREMISE", "_latest_date", "reg_time"], na_position="last")
          .drop_duplicates("_KEY_PREMISE", keep="last")
          .loc[:, ["_KEY_PREMISE", "status", "result", "_latest_date"]]
          .rename(columns={
              "status": f"{prefix}latest_status",
              "result": f"{prefix}latest_result",
              "_latest_date": f"{prefix}latest_date"
          })
    )

    agg = (
        df.groupby("_KEY_PREMISE", dropna=False)
          .agg(**{
              f"{prefix}total": ("_KEY_PREMISE", "count"),
              f"{prefix}open":  ("_is_open", "sum"),
          })
          .reset_index()
    )

    out = agg.merge(latest, on="_KEY_PREMISE", how="left")
    return out

insp_sum  = summarize_tasks(insp_df,  "insp_",  CLOSED_TERMS)
maint_sum = summarize_tasks(maint_df, "maint_", CLOSED_TERMS)

summary = summary.merge(insp_sum,  on="_KEY_PREMISE", how="left")
summary = summary.merge(maint_sum, on="_KEY_PREMISE", how="left")

# تأكيد نوع LastDaily بعد الدمج
if "LastDaily" in summary.columns:
    summary["LastDaily"] = smart_parse_datetime(summary["LastDaily"], excel_origin=excel_origin)

# تحويل الأعمدة الرقمية لأعداد صحيحة
for col in ["insp_open", "maint_open", "insp_total", "maint_total"]:
    if col in summary.columns:
        summary[col] = pd.to_numeric(summary[col], errors="coerce").fillna(0).astype(int)

# ===================== Next Action =====================
def next_action(row):
    insp_open  = as_int0(row.get("insp_open", 0))
    maint_open = as_int0(row.get("maint_open", 0))
    if maint_open > 0:
        return "تسريع صيانة مفتوحة"
    if insp_open > 0:
        return "متابعة فحص مفتوح"
    return "لا إجراء"

summary["Next Action"] = summary.apply(next_action, axis=1)

# ===================== KPIs =====================
st.subheader("📈 مؤشرات عامة")
c1, c2, c3 = st.columns(3)
c1.metric("عدد العدادات غير المتصلة", f"{summary['_KEY_PREMISE'].nunique():,}")
c2.metric("مهام فحص مفتوحة",        int(summary.get("insp_open",  pd.Series()).fillna(0).sum()))
c3.metric("مهام صيانة مفتوحة",       int(summary.get("maint_open", pd.Series()).fillna(0).sum()))

# ===================== Unified Table =====================
st.subheader("📋 الجدول الموحد")
display_cols = ["_KEY_PREMISE", "LastDaily"]
for c in ["Office", "States", "Logistic State", "Gateway Id", "Latitude", "Longitude"]:
    if c in summary.columns:
        display_cols.append(c)
display_cols += [c for c in [
    "insp_total","insp_open","insp_latest_status","insp_latest_result","insp_latest_date",
    "maint_total","maint_open","maint_latest_status","maint_latest_result","maint_latest_date",
    "Next Action"
] if c in summary.columns]

st.dataframe(summary[display_cols] if display_cols else summary, use_container_width=True)

# ===================== Download (Excel) =====================
excel_bytes = to_excel_download(summary)
st.download_button(
    label="⬇️ تنزيل النتائج (Excel)",
    data=excel_bytes,
    file_name=f"premise_tracker_results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
