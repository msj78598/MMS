# -*- coding: utf-8 -*-
# inspection_performance_suite.py
# Unified Streamlit app to showcase Inspection efforts, tie to Maintenance, and use Disconnected list.
# Key = Premise / Utility Site Id. Robust date parsing. Multi-file uploads. Multi-tab analytics. Excel exports.

import re
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

# ------------------------ Page Setup ------------------------
st.set_page_config(page_title="Inspection Performance Suite — Premise Key", layout="wide")

# ------------------------ Helpers ------------------------
def norm_col(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip()).lower()

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Find a column by exact normalized name or by containment (AR/EN)."""
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
    """Parse text + Excel serials. Tries dayfirst then monthfirst. Supports Excel 1900/1904 origins."""
    if series is None:
        return pd.Series([], dtype="datetime64[ns]")
    s = series.copy()

    def clean(x):
        if pd.isna(x): return np.nan
        x = str(x).strip()
        if x == "" or x.lower() in {"none", "nan", "null", "-", "—", "0"}: return np.nan
        if re.fullmatch(r"0{2,}[-/:]0{2,}[-/:]0{2,}.*", x): return np.nan
        return x

    s = s.map(clean)

    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)
    # try again month-first
    need2 = parsed.isna()
    if need2.any():
        parsed.loc[need2] = pd.to_datetime(s[need2], errors="coerce", dayfirst=False, infer_datetime_format=True)

    # Excel serial fallback
    need_excel = parsed.isna()
    if need_excel.any():
        as_num = pd.to_numeric(s.where(need_excel), errors="coerce")
        mask = as_num.notna()
        if mask.any():
            parsed.loc[mask] = pd.to_datetime(as_num[mask], unit="d", origin=excel_origin, errors="coerce")

    return parsed

def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Create a multi-sheet Excel file. Prefer xlsxwriter, fallback to openpyxl."""
    bio = BytesIO()
    try:
        with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, index=False, sheet_name=name[:31] or "Sheet")
    except ModuleNotFoundError:
        with pd.ExcelWriter(bio, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, index=False, sheet_name=name[:31] or "Sheet")
    return bio.getvalue()

def infer_bucket_from_name(name: str, kind_label: str) -> str:
    """Infer maintenance bucket from file name; inspection stays 'فحص'."""
    if kind_label == "فحص":
        return "فحص"
    n = (name or "").lower()
    if any(k in n for k in ["استبدال"]): return "استبدال"
    if any(k in n for k in ["تحسين", "استخراج", "تحسين واستخراج"]): return "تحسين/استخراج"
    if any(k in n for k in ["صيانة"]): return "صيانة"
    if any(k in n for k in ["كشف", "معاينة"]): return "كشف/معاينة"
    return kind_label

def normalize_task_flags(df: pd.DataFrame, closed_terms: set[str]) -> pd.DataFrame:
    """Ensure required cols exist + compute open/last flags."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["_KEY_PREMISE","reg_time","close_time","status","result","bucket","source","_is_open","_last"])
    d = df.copy()
    for col in ["reg_time","close_time"]:
        if col not in d.columns:
            d[col] = pd.NaT
    for col in ["status","result","bucket","source"]:
        if col not in d.columns:
            d[col] = np.nan
    status_norm = d["status"].astype(str).str.strip().str.lower()
    is_closed_by_status = status_norm.isin(closed_terms)
    is_closed_by_time   = d["close_time"].notna()
    d["_is_open"] = ~(is_closed_by_status | is_closed_by_time)
    d["_last"]    = d["close_time"].fillna(d["reg_time"])
    return d

def safe_hist_bar(series, bins=10, title=None):
    """Streamlit-safe histogram (no IntervalIndex)."""
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        st.info("لا توجد بيانات كافية للرسم.")
        return
    counts, edges = np.histogram(s, bins=bins)
    labels = [f"{int(np.floor(edges[i]))}–{int(np.ceil(edges[i+1]))}" for i in range(len(edges)-1)]
    hist_df = pd.DataFrame({"bin": labels, "count": counts})
    if title:
        st.markdown(f"#### {title}")
    st.bar_chart(hist_df.set_index("bin"))

# ------------------------ Sidebar (Inputs & Settings) ------------------------
st.title("📊 Inspection Performance Suite — Premise Key")

with st.sidebar:
    st.header("📁 ملفات الإدخال")
    dis_file   = st.file_uploader("ملف العدادات غير المتصلة", type=["xlsx","xls"])
    insp_files = st.file_uploader("ملفات مهام الفحص (0..N)",  type=["xlsx","xls"], accept_multiple_files=True)
    mnt_files  = st.file_uploader("ملفات مهام الصيانة (0..N)", type=["xlsx","xls"], accept_multiple_files=True)

    st.markdown("---")
    st.header("⚙️ إعدادات التواريخ")
    excel_origin = st.selectbox("Excel Origin (للأرقام التسلسلية)", ["1899-12-30", "1904-01-01"], index=0)

    st.markdown("—")
    st.header("🔒 حالات تعتبر (مقفلة)")
    default_closed_terms = """
closed, complete, completed, done, resolved,
مغلق, مغلقة, مقفلة, مقفل, منجز, منجزة, منتهية, منتهي, تمت المعالجة
""".strip()
    closed_terms_input = st.text_area("قائمة مفصولة بفواصل (يمكن تعديلها)", value=default_closed_terms, height=90)
    CLOSED_TERMS = {w.strip().lower() for w in closed_terms_input.split(",") if w.strip()}

    st.markdown("---")
    st.header("🚀 التنفيذ")
    run_btn = st.button("ابدأ التحليل")

if not run_btn or not dis_file:
    st.info("⬆️ ارفع الملفات ثم اضغط **ابدأ التحليل**.")
    st.stop()

# ------------------------ Read Disconnected ------------------------
dis_df = pd.read_excel(dis_file)

PREMISE_CANDS_DIS = ["Utility Site Id", "Premise", "رقم المكان"]
LAST_CANDS = [
    "Last Daily", "LastDaily", "Last Communication", "Last Comm",
    "Last Daily Read", "Last Daily Date",
    "آخر قراءة", "آخر اتصال", "اخر اتصال", "اخر قراءة"
]

premise_dis = pick_col(dis_df, PREMISE_CANDS_DIS)
last_col    = pick_col(dis_df, LAST_CANDS)

if not premise_dis:
    st.error("لم يتم العثور على عمود Premise/Utility Site Id في ملف غير المتصلين.")
    st.stop()

dis_df["_KEY_PREMISE"] = dis_df[premise_dis].astype(str).str.strip()

with st.expander("🔧 اختيار عمود 'آخر اتصال' يدويًا (اختياري)"):
    last_choice = st.selectbox("اختر عمود 'آخر اتصال':", options=["(اكتشاف تلقائي)"] + list(dis_df.columns), index=0)
    if last_choice != "(اكتشاف تلقائي)":
        last_col = last_choice

if last_col and last_col in dis_df.columns:
    dis_df["LastDaily"] = smart_parse_datetime(dis_df[last_col], excel_origin=excel_origin)
    ok = int(dis_df["LastDaily"].notna().sum())
    st.success(f"تحويل '{last_col}': {ok}/{len(dis_df)} قيماً صالحة.")
    with st.expander("🧪 أمثلة غير قابلة للتحويل"):
        bad = dis_df.loc[dis_df["LastDaily"].isna(), [last_col]].head(12)
        st.write("لا توجد أمثلة غير قابلة للتحويل ✅" if bad.empty else bad)
else:
    dis_df["LastDaily"] = pd.NaT
    st.warning("⚠️ لم يُعثر على عمود 'آخر اتصال' — سيُترك فارغًا.")

# ------------------------ Read Tasks (Inspection / Maintenance) ------------------------
def load_task_files(files, kind_label: str) -> pd.DataFrame:
    """Normalize multiple files to a unified schema using Premise as the key."""
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
            st.warning(f"تجاهل '{getattr(f,'name','file')}' — لا يحتوي Premise/Utility Site Id.")
            continue

        tmp = pd.DataFrame()
        tmp["_KEY_PREMISE"] = df[premise_col].astype(str).str.strip()
        tmp["reg_time"]   = smart_parse_datetime(df[reg_col],   excel_origin=excel_origin) if (reg_col   and reg_col   in df.columns) else pd.NaT
        tmp["close_time"] = smart_parse_datetime(df[close_col], excel_origin=excel_origin) if (close_col and close_col in df.columns) else pd.NaT
        tmp["status"]     = df[status_col].astype(str) if (status_col and status_col in df.columns) else np.nan
        tmp["result"]     = df[result_col].astype(str) if (result_col and result_col in df.columns) else np.nan
        tmp["bucket"]     = infer_bucket_from_name(getattr(f, "name", ""), kind_label)
        tmp["source"]     = getattr(f, "name", kind_label)
        frames.append(tmp)

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)

insp_df = load_task_files(insp_files, "فحص")
mnt_df  = load_task_files(mnt_files,  "صيانة")

# Normalize flags
insp_df = normalize_task_flags(insp_df, CLOSED_TERMS)
mnt_df  = normalize_task_flags(mnt_df,  CLOSED_TERMS)

# ------------------------ Base Join (disconnected) ------------------------
base = dis_df[["_KEY_PREMISE","LastDaily"]].drop_duplicates()

# Latest inspection (by close_time then reg_time)
insp_sorted = insp_df.sort_values(["_KEY_PREMISE","_last","reg_time"], na_position="last")
insp_latest = insp_sorted.drop_duplicates("_KEY_PREMISE", keep="last")

# First/Latest maintenance
mnt_sorted = mnt_df.sort_values(["_KEY_PREMISE","_last","reg_time"], na_position="last")
mnt_first  = mnt_df.sort_values(["_KEY_PREMISE","reg_time"], na_position="last").drop_duplicates("_KEY_PREMISE", keep="first")
mnt_latest = mnt_sorted.drop_duplicates("_KEY_PREMISE", keep="last")
if not mnt_latest.empty:
    status_norm = mnt_latest["status"].astype(str).str.strip().str.lower()
    mnt_latest["mnt_closed"] = mnt_latest["close_time"].notna() | status_norm.isin(CLOSED_TERMS)
else:
    mnt_latest["mnt_closed"] = pd.Series(dtype=bool)

# Join to base
base = base.merge(
    insp_latest[["_KEY_PREMISE","reg_time","close_time","status","result","bucket","source"]]
      .rename(columns={"reg_time":"insp_reg","close_time":"insp_close","status":"insp_status","result":"insp_result","bucket":"insp_bucket","source":"insp_source"}),
    on="_KEY_PREMISE", how="left"
)
base = base.merge(
    mnt_first[["_KEY_PREMISE","reg_time"]].rename(columns={"reg_time":"mnt_first_reg"}),
    on="_KEY_PREMISE", how="left"
)
base = base.merge(
    mnt_latest[["_KEY_PREMISE","reg_time","close_time","status","result","bucket","source","mnt_closed"]]
      .rename(columns={"reg_time":"mnt_last_reg","close_time":"mnt_last_close","status":"mnt_last_status","result":"mnt_last_result","bucket":"mnt_last_bucket","source":"mnt_last_source"}),
    on="_KEY_PREMISE", how="left"
)

base["insp_done"]  = base["insp_close"].notna()
base["has_mnt"]    = base["mnt_first_reg"].notna()
base["mnt_open"]   = base["has_mnt"] & ~base["mnt_closed"].fillna(False)

# Days between last inspection close and first maintenance reg (no negatives)
base["days_from_insp_to_mnt"] = (base["mnt_first_reg"] - base["insp_close"]).dt.days
base.loc[base["days_from_insp_to_mnt"] < 0, "days_from_insp_to_mnt"] = np.nan

# ------------------------ Aggregations (counts per premise) ------------------------
def per_premise_rollup(insp_df: pd.DataFrame, mnt_df: pd.DataFrame, dis_df: pd.DataFrame) -> pd.DataFrame:
    # Inspection agg
    if insp_df.empty:
        insp_by = pd.DataFrame(columns=["_KEY_PREMISE","insp_cnt","insp_open","insp_closed","insp_last_date","insp_buckets"])
    else:
        _i = insp_df.copy()
        _i["_last_date"] = _i["_last"]
        g = _i.groupby("_KEY_PREMISE")
        insp_by = g.agg(
            insp_cnt=(" _KEY_PREMISE".strip(), "count"),
            insp_open=("_is_open", "sum"),
            insp_last_date=("_last_date", "max")
        ).reset_index()
        insp_by["insp_closed"] = insp_by["insp_cnt"] - insp_by["insp_open"]
        insp_buckets = g["bucket"].apply(lambda s: ", ".join(sorted(set(map(str, s))))).reset_index().rename(columns={"bucket":"insp_buckets"})
        insp_by = insp_by.merge(insp_buckets, on="_KEY_PREMISE", how="left")

    # Maintenance agg
    if mnt_df.empty:
        mnt_by = pd.DataFrame(columns=["_KEY_PREMISE","mnt_cnt","mnt_open","mnt_closed","mnt_last_status","mnt_last_date","mnt_buckets"])
    else:
        _m = mnt_df.copy()
        _m["_last_date"] = _m["_last"]
        g = _m.groupby("_KEY_PREMISE")
        mnt_by = g.agg(
            mnt_cnt=(" _KEY_PREMISE".strip(), "count"),
            mnt_open=("_is_open", "sum"),
            mnt_last_date=("_last_date", "max")
        ).reset_index()
        mnt_by["mnt_closed"] = mnt_by["mnt_cnt"] - mnt_by["mnt_open"]
        _m_sorted = _m.sort_values(["_KEY_PREMISE","_last_date","reg_time"], na_position="last")
        last_idx = _m_sorted.groupby("_KEY_PREMISE").tail(1).index
        last_slice = _m_sorted.loc[last_idx, ["_KEY_PREMISE","status"]].rename(columns={"status":"mnt_last_status"})
        mnt_buckets = g["bucket"].apply(lambda s: ", ".join(sorted(set(map(str, s))))).reset_index().rename(columns={"bucket":"mnt_buckets"})
        mnt_by = mnt_by.merge(last_slice, on="_KEY_PREMISE", how="left").merge(mnt_buckets, on="_KEY_PREMISE", how="left")

    per_p = dis_df[["_KEY_PREMISE","LastDaily"]].drop_duplicates()
    per_p = per_p.merge(insp_by, on="_KEY_PREMISE", how="left").merge(mnt_by, on="_KEY_PREMISE", how="left")

    for col in ["insp_cnt","insp_open","insp_closed","mnt_cnt","mnt_open","mnt_closed"]:
        if col in per_p.columns:
            per_p[col] = pd.to_numeric(per_p[col], errors="coerce").fillna(0).astype(int)

    per_p["has_insp"] = per_p["insp_cnt"].gt(0)
    per_p["has_mnt"]  = per_p["mnt_cnt"].gt(0)
    return per_p

per_prem = per_premise_rollup(insp_df, mnt_df, dis_df)

# ------------------------ TABS ------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1) جهود الفحص",
    "2) متابعة ما بعد الفحص",
    "3) أثر الفحص على الاتصال",
    "4) ملخص لكل عداد",
    "5) تنزيل التقارير"
])

# ==================== Tab 1: Inspection Efforts ====================
with tab1:
    st.subheader("🧰 جهود الفحص (Inspection Efforts)")

    total_disconnected = dis_df["_KEY_PREMISE"].nunique()
    insp_done_count    = int(base["insp_done"].sum())
    insp_rate          = (insp_done_count / total_disconnected * 100.0) if total_disconnected else 0.0
    insp_dur = (insp_df["close_time"] - insp_df["reg_time"]).dt.days.dropna()
    avg_insp_days = float(insp_dur.mean()) if not insp_dur.empty else 0.0

    k1, k2, k3 = st.columns(3)
    k1.metric("إجمالي غير المتصلة (Premise)", f"{total_disconnected:,}")
    k2.metric("تم فحصها (Premise)", f"{insp_done_count:,}", f"{insp_rate:.1f}%")
    k3.metric("متوسط مدة الفحص (يوم)", f"{avg_insp_days:.1f}")

    st.markdown("### الاتجاه الزمني للفحوص")
    if not insp_df.empty:
        i_trend = insp_df.copy()
        i_trend["t"] = i_trend["close_time"].fillna(i_trend["reg_time"])
        i_trend = i_trend.dropna(subset=["t"])
        i_trend["day"] = i_trend["t"].dt.date
        daily = i_trend.groupby("day")["_KEY_PREMISE"].count().reset_index().rename(columns={"_KEY_PREMISE":"inspections"})
        st.line_chart(daily.set_index("day"))
        st.dataframe(daily, use_container_width=True)
    else:
        st.info("لا توجد بيانات فحص لعرض الاتجاه الزمني.")

    st.markdown("### أكثر نتائج الفحص تكرارًا")
    if "result" in insp_df.columns and not insp_df.empty:
        top_res = (insp_df["result"].astype(str)
                   .str.strip().str.lower()
                   .replace({"nan":np.nan})
                  ).value_counts(dropna=True).head(20).rename_axis("result").reset_index(name="count")
        st.bar_chart(top_res.set_index("result"))
        st.dataframe(top_res, use_container_width=True)
    else:
        st.info("لا تتوفر أعمدة نتائج فحص لعرضها.")

# ==================== Tab 2: Post-Inspection Follow-up ====================
with tab2:
    st.subheader("🔗 متابعة ما بعد الفحص (مسؤولية الصيانة)")

    r1 = base[(base["insp_done"]) & (~base["has_mnt"])].copy().sort_values(["insp_close"], ascending=[False])
    r2 = base[(base["insp_done"]) & (base["has_mnt"]) & (base["mnt_closed"])].copy().sort_values(["mnt_last_close","mnt_last_reg"], ascending=[False, False])
    r3 = base[(base["insp_done"]) & (base["mnt_open"])].copy().sort_values(["mnt_last_reg"], ascending=[False])

    c1, c2, c3 = st.columns(3)
    c1.metric("مفحوصة ولا يوجد صيانة", len(r1))
    c2.metric("صيانة مُقفلة وما زال غير متصل", len(r2))
    c3.metric("صيانة مفتوحة بعد الفحص", len(r3))

    cols = ["_KEY_PREMISE","LastDaily",
            "insp_reg","insp_close","insp_status","insp_result",
            "mnt_first_reg","mnt_last_reg","mnt_last_close","mnt_last_status","mnt_last_result","mnt_last_bucket",
            "days_from_insp_to_mnt"]

    st.markdown("### 1) مفحوصة ولا يوجد صيانة")
    st.dataframe(r1[cols], use_container_width=True)

    st.markdown("### 2) صيانة مُقفلة — وما زال غير متصل")
    st.dataframe(r2[cols], use_container_width=True)

    st.markdown("### 3) صيانة مفتوحة بعد الفحص")
    st.dataframe(r3[cols], use_container_width=True)

# ==================== Tab 3: Impact on Reconnection ====================
with tab3:
    st.subheader("🔌 أثر الفحص على الاتصال")

    st.info("ملاحظة: بما أن البيانات المتوفرة هي لقائمة (غير المتصلين) الحالية فقط، فالعداد الموجود هنا يُعتبر ما يزال غير متصل. لقياس الرجوع الفعلي للاتصال نحتاج لقطات يومية أو ملف الأحداث المتصلة.")

    tmp = base[base["insp_done"] & base["has_mnt"]].copy()
    if not tmp.empty:
        st.markdown("### الفاصل الزمني بين إقفال الفحص وبداية الصيانة (أيام)")
        days = pd.to_numeric(tmp["days_from_insp_to_mnt"], errors="coerce").dropna()
        safe_hist_bar(days, bins=12)

        st.dataframe(
            tmp[["_KEY_PREMISE","insp_close","mnt_first_reg","days_from_insp_to_mnt"]]
              .sort_values("days_from_insp_to_mnt", ascending=False),
            use_container_width=True
        )
    else:
        st.info("لا توجد حالات بها فحص متبوع بصيانة لحساب الفاصل الزمني.")

# ==================== Tab 4: Per-Premise Summary ====================
with tab4:
    st.subheader("📚 ملخص لكل عداد غير متصل (مرات الفحص + الصيانة + السلال)")

    # Build per-premise
    def per_premise_rollup_for_tab(insp_df, mnt_df, dis_df):
        if insp_df.empty:
            insp_by = pd.DataFrame(columns=["_KEY_PREMISE","insp_cnt","insp_open","insp_closed","insp_last_date","insp_buckets"])
        else:
            _i = insp_df.copy()
            _i["_last_date"] = _i["_last"]
            g = _i.groupby("_KEY_PREMISE")
            insp_by = g.agg(
                insp_cnt=(" _KEY_PREMISE".strip(), "count"),
                insp_open=("_is_open", "sum"),
                insp_last_date=("_last_date", "max")
            ).reset_index()
            insp_by["insp_closed"] = insp_by["insp_cnt"] - insp_by["insp_open"]
            insp_buckets = g["bucket"].apply(lambda s: ", ".join(sorted(set(map(str, s))))).reset_index().rename(columns={"bucket":"insp_buckets"})
            insp_by = insp_by.merge(insp_buckets, on="_KEY_PREMISE", how="left")

        if mnt_df.empty:
            mnt_by = pd.DataFrame(columns=["_KEY_PREMISE","mnt_cnt","mnt_open","mnt_closed","mnt_last_status","mnt_last_date","mnt_buckets"])
        else:
            _m = mnt_df.copy()
            _m["_last_date"] = _m["_last"]
            g = _m.groupby("_KEY_PREMISE")
            mnt_by = g.agg(
                mnt_cnt=(" _KEY_PREMISE".strip(), "count"),
                mnt_open=("_is_open", "sum"),
                mnt_last_date=("_last_date", "max")
            ).reset_index()
            mnt_by["mnt_closed"] = mnt_by["mnt_cnt"] - mnt_by["mnt_open"]
            _m_sorted = _m.sort_values(["_KEY_PREMISE","_last_date","reg_time"], na_position="last")
            last_idx = _m_sorted.groupby("_KEY_PREMISE").tail(1).index
            last_slice = _m_sorted.loc[last_idx, ["_KEY_PREMISE","status"]].rename(columns={"status":"mnt_last_status"})
            mnt_buckets = g["bucket"].apply(lambda s: ", ".join(sorted(set(map(str, s))))).reset_index().rename(columns={"bucket":"mnt_buckets"})
            mnt_by = mnt_by.merge(last_slice, on="_KEY_PREMISE", how="left").merge(mnt_buckets, on="_KEY_PREMISE", how="left")

        per_p = dis_df[["_KEY_PREMISE","LastDaily"]].drop_duplicates()
        per_p = per_p.merge(insp_by, on="_KEY_PREMISE", how="left").merge(mnt_by, on="_KEY_PREMISE", how="left")
        for col in ["insp_cnt","insp_open","insp_closed","mnt_cnt","mnt_open","mnt_closed"]:
            if col in per_p.columns:
                per_p[col] = pd.to_numeric(per_p[col], errors="coerce").fillna(0).astype(int)
        per_p["has_insp"] = per_p["insp_cnt"].gt(0)
        per_p["has_mnt"]  = per_p["mnt_cnt"].gt(0)
        return per_p

    per_prem_tab = per_premise_rollup_for_tab(insp_df, mnt_df, dis_df)

    fc1, fc2, fc3 = st.columns([2,2,2])
    with fc1:
        search_prem = st.text_input("🔎 ابحث عن Premise", value="")
    with fc2:
        f_has_mnt = st.selectbox("فلتر وجود صيانة", ["الكل", "يوجد صيانة", "لا يوجد صيانة"], index=0)
    with fc3:
        all_buckets = sorted(set(", ".join(per_prem_tab["mnt_buckets"].dropna().astype(str)).split(", "))) if "mnt_buckets" in per_prem_tab.columns else []
        sel_buckets = st.multiselect("سلال الصيانة", options=[b for b in all_buckets if b], default=[])

    fdf = per_prem_tab.copy()
    if search_prem.strip():
        s = search_prem.strip()
        fdf = fdf[fdf["_KEY_PREMISE"].astype(str).str.contains(s, case=False, na=False)]

    if f_has_mnt == "يوجد صيانة":
        fdf = fdf[fdf["has_mnt"] == True]
    elif f_has_mnt == "لا يوجد صيانة":
        fdf = fdf[fdf["has_mnt"] == False]

    if sel_buckets:
        fdf = fdf[fdf["mnt_buckets"].fillna("").apply(lambda x: any(b in str(x) for b in sel_buckets))]

    display_cols = [
        "_KEY_PREMISE","LastDaily",
        "insp_cnt","insp_open","insp_closed","insp_last_date","insp_buckets",
        "mnt_cnt","mnt_open","mnt_closed","mnt_last_status","mnt_last_date","mnt_buckets",
        "has_insp","has_mnt"
    ]
    display_cols = [c for c in display_cols if c in fdf.columns]
    st.dataframe(
        fdf[display_cols].sort_values(["has_insp","has_mnt","mnt_open","insp_open"], ascending=[False, False, False, False]),
        use_container_width=True
    )

# ==================== Tab 5: Exports ====================
with tab5:
    st.subheader("⬇️ تنزيل التقارير (Excel متعدد الأوراق)")

    r1 = base[(base["insp_done"]) & (~base["has_mnt"])].copy().sort_values(["insp_close"], ascending=[False])
    r2 = base[(base["insp_done"]) & (base["has_mnt"]) & (base["mnt_closed"])].copy().sort_values(["mnt_last_close","mnt_last_reg"], ascending=[False, False])
    r3 = base[(base["insp_done"]) & (base["mnt_open"])].copy().sort_values(["mnt_last_reg"], ascending=[False])

    cols = ["_KEY_PREMISE","LastDaily",
            "insp_reg","insp_close","insp_status","insp_result","insp_bucket","insp_source",
            "mnt_first_reg","mnt_last_reg","mnt_last_close","mnt_last_status","mnt_last_result","mnt_last_bucket","mnt_last_source",
            "days_from_insp_to_mnt"]

    export_sheets = {
        "per_premise_summary": per_prem,
        "inspected_no_maintenance": r1[cols],
        "maintenance_closed_still_disconnected": r2[cols],
        "maintenance_open_post_inspection": r3[cols],
        "base_join": base
    }
    excel_bytes = to_excel_bytes(export_sheets)
    st.download_button(
        "⬇️ تنزيل كل التقارير (Excel)",
        data=excel_bytes,
        file_name=f"inspection_suite_reports_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.caption("وجود Premise ضمن (غير المتصلين) يعني أنه ما يزال غير متصل عند لحظة إنشاء التقرير. لإثبات العودة للاتصال بدقة نحتاج لقطات يومية/ملف قراءات متصلة.")
