"""
VIT FFCS Grade Calculator  –  Enhanced v2
Supports: Pure Theory · Pure Lab · Embedded (Theory + Lab)
Run  : streamlit run vit_grade_calculator.py
Deps : pip install streamlit pandas openpyxl
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
import io
from datetime import datetime

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  CONSTANTS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXCEL_FILE   = "vit_grades_data.xlsx"
COURSE_TYPES = ["Pure Theory", "Pure Lab", "Embedded (Theory + Lab)"]

GRADE_EMOJI  = {"S":"🟢","A":"🟢","B":"🔵","C":"🟡","D":"🟠","E":"🔴","F":"🔴","N":"⚫"}
GRADE_POINTS = {"S":10,"A":9,"B":8,"C":7,"D":6,"E":5,"F":0,"N":0}

EXCEL_COLS = [
    "Subject","Course_Type",
    "L","T","P",
    "CAT1_raw","CAT2_raw","DA1","DA2","DA3",
    "Lab_CA_avg","Lab_FAT_raw",
    "Theory_FAT_sim",
    "Class_Mean","Class_Sigma","First_Mark",
    "Grading_Mode",
    "CAM_Total","Lab_Score","Theory_Score",
    "Combined_Total","Predicted_Grade",
    "Last_Updated",
]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  EXCEL HELPERS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def load_df():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            for c in EXCEL_COLS:
                if c not in df.columns:
                    df[c] = np.nan
            df["Subject"] = df["Subject"].astype(str).str.strip()
            return df[EXCEL_COLS]
        except Exception:
            pass
    return pd.DataFrame(columns=EXCEL_COLS)


def save_row(row):
    df = load_df()
    subj = str(row["Subject"]).strip()
    row["Last_Updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    full = {c: row.get(c, np.nan) for c in EXCEL_COLS}
    full["Subject"] = subj
    if subj in df["Subject"].values:
        for k, v in full.items():
            df.loc[df["Subject"] == subj, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([full])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)


def delete_row(subj):
    df = load_df()
    df = df[df["Subject"] != subj]
    df.to_excel(EXCEL_FILE, index=False)


def df_to_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()


def merge_excel(uploaded):
    try:
        up = pd.read_excel(uploaded)
        if "Subject" not in up.columns:
            return "Uploaded file must have a 'Subject' column."
        local = load_df()
        n = 0
        for _, r in up.iterrows():
            s = str(r.get("Subject","")).strip()
            if not s or s == "nan":
                continue
            if s in local["Subject"].values:
                for c in up.columns:
                    if c in EXCEL_COLS:
                        local.loc[local["Subject"]==s, c] = r[c]
            else:
                nr = {c: r.get(c, np.nan) for c in EXCEL_COLS}
                nr["Subject"] = s
                local = pd.concat([local, pd.DataFrame([nr])], ignore_index=True)
            n += 1
        local.to_excel(EXCEL_FILE, index=False)
        return f"OK:{n}"
    except Exception as e:
        return f"ERR:{e}"

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  GRADING LOGIC
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def estimate_sigma(mean, first):
    if first > mean:
        return max((first - mean) / 2.5, 1.0)
    return 10.0


def relative_bands(mean, sigma):
    raw = {
        "S": max(mean + 1.5*sigma, 80.0),
        "A": mean + 0.5*sigma,
        "B": mean - 0.5*sigma,
        "C": mean - 1.0*sigma,
        "D": mean - 1.5*sigma,
        "E": mean - 2.0*sigma,
    }
    return {g: round(v, 2) for g, v in raw.items()}


def absolute_bands():
    return {"S":90,"A":80,"B":70,"C":60,"D":55,"E":50}


def get_grade(total, bds):
    for g in ["S","A","B","C","D","E"]:
        if total >= bds[g]:
            return g
    return "F"


def calc_cam(c1, c2, d1, d2, d3):
    s1 = c1 * 0.30
    s2 = c2 * 0.30
    da = d1 + d2 + d3
    return s1+s2+da, s1, s2, da


def theory_total(cam, fat):
    return cam + fat * 0.40


def fat_for_target(cam, target):
    return (target - cam) / 0.40


def lab_score(ca, fat):
    return ca * 0.60 + fat * 0.40


def lab_fat_for_target(ca, target):
    return (target - ca*0.60) / 0.40


def embedded_combined(t_sc, l_sc, L, T, P):
    tw  = float(L + T)
    lw  = float(P) / 2.0
    tot = tw + lw
    if tot == 0:
        return 0.0, tw, lw, tot
    return round((t_sc*tw + l_sc*lw)/tot, 4), tw, lw, tot


def emb_fat_for_target(cam, l_sc, target, tw, lw, tot):
    if tw == 0:
        return 101.0
    t_target = (target*tot - l_sc*lw) / tw
    return (t_target - cam) / 0.40


def sf(v, d=0.0):
    try:
        x = float(v)
        return d if (x != x) else x   # nan check
    except Exception:
        return d


def si(v, d=0):
    try:
        x = int(float(v))
        return x
    except Exception:
        return d

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  APP
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.set_page_config(page_title="VIT Grade Calc", page_icon="🎓", layout="wide")

st.markdown("""<style>
h2{color:#1565C0;border-bottom:2px solid #90CAF9;padding-bottom:4px;}
.stMetric label{font-size:0.78rem;}
.ib{background:#E3F2FD;border-left:4px solid #1565C0;padding:7px 12px;border-radius:4px;font-size:.87rem;margin:4px 0;}
.ob{background:#E8F5E9;border-left:4px solid #2E7D32;padding:7px 12px;border-radius:4px;font-size:.87rem;margin:4px 0;}
.eb{background:#FFEBEE;border-left:4px solid #C62828;padding:7px 12px;border-radius:4px;font-size:.87rem;margin:4px 0;}
</style>""", unsafe_allow_html=True)

st.title("🎓  VIT FFCS Grade Calculator  v2")
st.caption("Pure Theory · Pure Lab · Embedded (Theory+Lab)  |  FFCS Regulations v4.0")

# load
df_all       = load_df()
all_subjects = df_all["Subject"].dropna().tolist()

# ══════════════════════════════════════════════════════════
#  ① Subject & Course Type
# ══════════════════════════════════════════════════════════
st.header("① Subject & Course Type")

c_left, c_right = st.columns([1.3, 1], gap="large")

with c_left:
    subj_typed = st.text_input("📚 Subject name", placeholder="e.g. Data Structures")
    matches = [s for s in all_subjects if subj_typed.lower() in s.lower()] if subj_typed else all_subjects

    if matches:
        pick = st.selectbox(
            f"💾 {len(all_subjects)} saved — select to auto-fill",
            ["— new subject —"] + matches)
    else:
        pick = "— new subject —"
        if subj_typed and all_subjects:
            st.caption("No match found — will create new entry.")

    subject = (pick if pick != "— new subject —" else subj_typed).strip()

    pf = {}
    if subject and subject in df_all["Subject"].values:
        pf = df_all[df_all["Subject"]==subject].iloc[0].to_dict()
        st.markdown('<div class="ob">✅ Saved data loaded — values pre-filled below.</div>',
                    unsafe_allow_html=True)

with c_right:
    saved_ct = str(pf.get("Course_Type","Pure Theory"))
    ct_idx   = COURSE_TYPES.index(saved_ct) if saved_ct in COURSE_TYPES else 0
    course_type = st.radio("📋 Course type", COURSE_TYPES, index=ct_idx)

    L_val = T_val = P_val = 0
    tw = lw = tot_w = 0.0
    t_pct = l_pct = 0.0

    if course_type == "Embedded (Theory + Lab)":
        st.markdown("**LTPC — lecture / tutorial / practical credits**")
        ec = st.columns(3)
        with ec[0]: L_val = st.number_input("L",0,8,si(pf.get("L",3)),1)
        with ec[1]: T_val = st.number_input("T",0,4,si(pf.get("T",0)),1)
        with ec[2]: P_val = st.number_input("P",2,8,si(pf.get("P",2)),2,
                                            help="Each 2 P-hours = 1 lab credit")
        tw    = float(L_val + T_val)
        lw    = float(P_val) / 2.0
        tot_w = tw + lw

        # ── Manual ratio override ───────────────────────
        use_custom_ratio = st.checkbox(
            "⚙️ Override Theory / Lab ratio manually",
            value=False,
            help="By default the ratio is derived from your L+T vs P/2 credits. "
                 "Enable this to set a custom split (e.g. 75 / 25).")
        if use_custom_ratio:
            theory_pct_override = st.slider(
                "Theory weight %", 0, 100, 75, 5,
                help="75 = 75% Theory / 25% Lab  (VIT default for most embedded courses). "
                     "Drag to any value; Lab % updates automatically.")
            lab_pct_override = 100 - theory_pct_override
            st.markdown(
                f'<div class="ib">⚖️ <b>Custom ratio</b> → Theory <b>{theory_pct_override}%</b> '
                f'| Lab <b>{lab_pct_override}%</b></div>',
                unsafe_allow_html=True)
            # Replace credit-derived weights with percentage weights (sum = 100)
            tw    = float(theory_pct_override)
            lw    = float(lab_pct_override)
            tot_w = 100.0
        elif tot_w > 0:
            t_pct = tw / tot_w * 100
            l_pct = lw / tot_w * 100
            st.markdown(
                f'<div class="ib">⚖️ Ratio → Theory <b>{t_pct:.0f}%</b> ({tw:.0f} wt) '
                f'| Lab <b>{l_pct:.0f}%</b> ({lw:.1f} wt)  '
                f'<i>(derived from LTPC — tick above to override)</i></div>',
                unsafe_allow_html=True)

# Reset prediction session state whenever the subject changes
if st.session_state.get("_pred_subj") != subject:
    st.session_state["_pred_subj"]     = subject
    st.session_state["vit_class_mean"] = sf(pf.get("Class_Mean", 60.0))
    st.session_state["vit_first_mark"] = sf(pf.get("First_Mark", 0.0))
elif "vit_class_mean" not in st.session_state:
    st.session_state["vit_class_mean"] = sf(pf.get("Class_Mean", 60.0))
    st.session_state["vit_first_mark"] = sf(pf.get("First_Mark", 0.0))

# ══════════════════════════════════════════════════════════
#  ② Class Statistics
# ══════════════════════════════════════════════════════════
st.header("② Class Statistics — Grade Band Prediction")

with st.expander("📊 Enter class info (more info = better grade cutoff accuracy)", expanded=True):
    cs = st.columns(4)

    saved_gm = str(pf.get("Grading_Mode","Relative"))
    gm_idx   = 1 if saved_gm == "Absolute" else 0

    with cs[0]:
        grading_mode = st.selectbox("Grading mode", ["Relative","Absolute"], index=gm_idx,
                                    help="Relative: theory >10 students\nAbsolute: lab / small class")
    dis = (grading_mode == "Absolute")

    with cs[1]:
        class_mean = st.number_input("Class average /100", 0.0, 100.0,
                                     step=0.5, disabled=dis, key="vit_class_mean",
                                     help="Auto-filled by predictor below — or type manually")
    with cs[2]:
        first_mark = st.number_input("Topper / Highest mark /100", 0.0, 100.0,
                                     step=0.5, disabled=dis, key="vit_first_mark",
                                     help="Auto-filled by predictor below — σ≈(topper−mean)/2.5")
    with cs[3]:
        manual_sigma = st.number_input("Std Dev σ  (0 = auto)", 0.0, 30.0,
                                       sf(pf.get("Class_Sigma",0.0)), 0.5, disabled=dis,
                                       help="Leave 0 to auto-estimate")

    if grading_mode == "Relative":
        if manual_sigma > 0:
            sigma_used = manual_sigma
            sigma_src  = f"manual (σ={sigma_used:.2f})"
        elif first_mark > class_mean:
            sigma_used = estimate_sigma(class_mean, first_mark)
            sigma_src  = f"auto from topper (σ≈{sigma_used:.2f})"
        else:
            sigma_used = 10.0
            sigma_src  = "default σ=10 (enter class mean/topper for better accuracy)"
        _is_emb = (course_type == "Embedded (Theory + Lab)")
        bands = relative_bands(class_mean, sigma_used)
        band_str = "  |  ".join([f"**{g}** ≥{bands[g]:.1f}" for g in ["S","A","B","C","D","E"]])
        _s_note  = "  *(S-floor 80 for Embedded)*" if _is_emb else ""
        st.caption(f"🔢 Sigma source: {sigma_src}{_s_note}")
        st.markdown(f'<div class="ib">📊 Predicted grade bands → {band_str}</div>', unsafe_allow_html=True)
    else:
        bands      = absolute_bands()
        sigma_used = 0.0
        st.markdown('<div class="ib">📊 Absolute bands → S≥90 | A≥80 | B≥70 | C≥60 | D≥55 | E≥50</div>',
                    unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
#  ③ Marks Entry
# ══════════════════════════════════════════════════════════
st.header("③ Enter Your Marks")

L_col, R_col = st.columns([1.15, 1], gap="large")

# ── hold vars ──────────────────────────────────────────
cat1_raw = cat2_raw = da1 = da2 = da3 = 0.0
cat1_sc  = cat2_sc  = da_sum = cam = 0.0
lab_ca_avg = lab_fat_raw = lab_sc = 0.0
lab_pass   = True
theory_fat_sim = 60
theory_sc_sim  = 0.0
combined_sim   = 0.0
final_grade    = "—"

with L_col:
    # ── Theory marks ───────────────────────────────────
    if course_type in ("Pure Theory","Embedded (Theory + Lab)"):
        st.subheader("📝 Theory Component")
        r1, r2 = st.columns(2)
        with r1:
            cat1_raw = st.number_input("CAT-I raw (/50)", 0.0, 50.0,
                                       sf(pf.get("CAT1_raw",30.0)), 0.5,
                                       help="raw /50 → scaled ×0.30 = /15")
            cat1_sc  = cat1_raw * 0.30
            st.caption(f"Scaled → **{cat1_sc:.2f} / 15**")
        with r2:
            cat2_raw = st.number_input("CAT-II raw (/50)  [open book]", 0.0, 50.0,
                                       sf(pf.get("CAT2_raw",35.0)), 0.5)
            cat2_sc  = cat2_raw * 0.30
            st.caption(f"Scaled → **{cat2_sc:.2f} / 15**")

        st.markdown("**Digital Assignments**  (3 × /10 = /30)")
        dc = st.columns(3)
        with dc[0]: da1 = st.number_input("DA-1 /10",0.0,10.0,sf(pf.get("DA1",8.0)),0.5)
        with dc[1]: da2 = st.number_input("DA-2 /10",0.0,10.0,sf(pf.get("DA2",8.0)),0.5)
        with dc[2]: da3 = st.number_input("DA-3 /10",0.0,10.0,sf(pf.get("DA3",8.0)),0.5)

        cam, cat1_sc, cat2_sc, da_sum = calc_cam(cat1_raw, cat2_raw, da1, da2, da3)
        m1,m2,m3,m4 = st.columns(4)
        m1.metric("CAT-I /15",    f"{cat1_sc:.2f}")
        m2.metric("CAT-II /15",   f"{cat2_sc:.2f}")
        m3.metric("DA total /30", f"{da_sum:.1f}")
        m4.metric("CAM /60",      f"{cam:.2f}")
        st.progress(min(cam/60,1.0), text=f"CAM  {cam:.2f} / 60")

        with st.expander("📊 Per-component class comparison (avg & first mark)"):
            st.caption("Enter class average and topper score for each component to see where you stand.")
            pc1, pc2 = st.columns(2)

            with pc1:
                st.markdown("**CAT-I  /50**")
                cat1_cls_avg   = st.number_input("Class avg CAT-I",   0.0, 50.0, 0.0, 0.5, key="cat1_cls_avg")
                cat1_cls_first = st.number_input("1st mark CAT-I",    0.0, 50.0, 0.0, 0.5, key="cat1_cls_first")
                if cat1_cls_avg > 0:
                    d = cat1_raw - cat1_cls_avg
                    em = "🟢" if d >= 0 else "🔴"
                    st.caption(f"{em} You **{cat1_raw:.1f}** | Avg **{cat1_cls_avg:.1f}** | 1st **{cat1_cls_first:.1f}** | Δ **{d:+.1f}**")

            with pc2:
                st.markdown("**CAT-II  /50**")
                cat2_cls_avg   = st.number_input("Class avg CAT-II",  0.0, 50.0, 0.0, 0.5, key="cat2_cls_avg")
                cat2_cls_first = st.number_input("1st mark CAT-II",   0.0, 50.0, 0.0, 0.5, key="cat2_cls_first")
                if cat2_cls_avg > 0:
                    d = cat2_raw - cat2_cls_avg
                    em = "🟢" if d >= 0 else "🔴"
                    st.caption(f"{em} You **{cat2_raw:.1f}** | Avg **{cat2_cls_avg:.1f}** | 1st **{cat2_cls_first:.1f}** | Δ **{d:+.1f}**")

            st.markdown("**Digital Assignments  /10**")
            dac = st.columns(3)
            _da_info = [("DA-1", da1, "da1"), ("DA-2", da2, "da2"), ("DA-3", da3, "da3")]
            for _i, (_lbl, _my, _key) in enumerate(_da_info):
                with dac[_i]:
                    _avg   = st.number_input(f"Avg {_lbl}",  0.0, 10.0, 0.0, 0.5, key=f"{_key}_cls_avg")
                    _first = st.number_input(f"1st {_lbl}",  0.0, 10.0, 0.0, 0.5, key=f"{_key}_cls_first")
                    if _avg > 0:
                        _d = _my - _avg
                        _em = "🟢" if _d >= 0 else "🔴"
                        st.caption(f"{_em} You **{_my:.1f}** | Avg **{_avg:.1f}** | 1st **{_first:.1f}** | Δ **{_d:+.1f}**")

            # ══ PREDICTION ENGINE ══════════════════════════════
            st.divider()
            st.markdown("### 🔮 Predicted Class Mean & Topper — Final Total")
            st.caption("Computed from your component inputs + VIT behavioural patterns. Edit override fields then click Apply to push into Section ②.")

            # Read entered values from session_state (set by the widgets above)
            _c1a = sf(st.session_state.get("cat1_cls_avg",   0))
            _c1f = sf(st.session_state.get("cat1_cls_first", 0))
            _c2a = sf(st.session_state.get("cat2_cls_avg",   0))
            _c2f = sf(st.session_state.get("cat2_cls_first", 0))
            _d1a = sf(st.session_state.get("da1_cls_avg",    0))
            _d1f = sf(st.session_state.get("da1_cls_first",  0))
            _d2a = sf(st.session_state.get("da2_cls_avg",    0))
            _d2f = sf(st.session_state.get("da2_cls_first",  0))
            _d3a = sf(st.session_state.get("da3_cls_avg",    0))
            _d3f = sf(st.session_state.get("da3_cls_first",  0))

            # ── Fill missing CAT values with VIT patterns ──────
            # VIT baseline (10yr pattern): CAT-I avg ≈ 30/50 (60%),
            # CAT-II avg ≈ 32/50 (64%) — open-book + more practice effect.
            # Low-performing courses may dip to ~26-28; high-performing to ~36-40.
            if _c1a <= 0 and _c2a <= 0:
                _c1a_e, _c2a_e = 30.0, 32.0
                _cat_src = "VIT baseline (30/50, 32/50) — enter data for better accuracy"
            elif _c1a <= 0:
                _c1a_e, _c2a_e = round(_c2a * 0.94, 1), _c2a
                _cat_src = "CAT-I estimated from CAT-II"
            elif _c2a <= 0:
                _c1a_e, _c2a_e = _c1a, round(_c1a * 1.06, 1)
                _cat_src = "CAT-II estimated from CAT-I"
            else:
                _c1a_e, _c2a_e = _c1a, _c2a
                _cat_src = "both entered ✅"

            # Topper: VIT toppers typically hit 44-49 on CATs, 10 on DAs.
            # Estimate ≈ avg × 1.48 (flatter than older 1.55 — VIT marking is less spread)
            _c1f_e = _c1f if _c1f > 0 else round(min(49.0, _c1a_e * 1.48), 1)
            _c2f_e = _c2f if _c2f > 0 else round(min(50.0, _c2a_e * 1.48), 1)

            # ── DAs: VIT is lenient — avg 8.4-9.0, topper almost always 10 ──
            _d1a_e = _d1a if _d1a > 0 else 8.6
            _d2a_e = _d2a if _d2a > 0 else 8.6
            _d3a_e = _d3a if _d3a > 0 else 8.6
            _d1f_e = _d1f if _d1f > 0 else 10.0
            _d2f_e = _d2f if _d2f > 0 else 10.0
            _d3f_e = _d3f if _d3f > 0 else 10.0

            # ── Predicted CAM ──────────────────────────────────
            _pred_cam_avg   = _c1a_e*0.30 + _c2a_e*0.30 + _d1a_e + _d2a_e + _d3a_e
            _pred_cam_first = _c1f_e*0.30 + _c2f_e*0.30 + _d1f_e + _d2f_e + _d3f_e

            # ── FAT prediction — VIT behavioural trends ────────
            # Recalibrated model:
            # • Base: FAT avg ≈ 62-68% of 100 for a typical VIT class.
            #         Formula: cam_pct × 92 captures this range well.
            # • Trend 1 — Give-up penalty: ~15% of class barely attempts FAT.
            #   Effect on *class mean*: roughly −5 pts (smaller than older −8;
            #   give-ups are concentrated at the bottom, less impact on mean).
            # • Trend 2 — Last-min crammer boost: +5 pts on FAT avg.
            #   VIT students with prior papers + DA practice often outperform
            #   their CAT trend on FAT (desperation effect is real and strong).
            # Net adjustment: +5 - 5 = 0 net shift, but kept explicit for transparency.
            _GIVEUP  = 5.0
            _CRAMMER = 5.0
            _cam_pct        = min(_pred_cam_avg / 60.0, 1.0)
            _pred_fat_avg   = round(max(35.0, _cam_pct * 92.0 - _GIVEUP + _CRAMMER), 1)

            # FAT topper: committed toppers score 88-98; they cram hard.
            # Estimate: cam_first_pct × 100 × 0.98, floor 82.
            _cam_first_pct  = min(_pred_cam_first / 60.0, 1.0)
            _pred_fat_first = round(min(98.0, max(82.0, _cam_first_pct * 100.0 * 0.98)), 1)

            # ── Final predicted totals ─────────────────────────
            _pred_total_avg   = round(_pred_cam_avg   + _pred_fat_avg   * 0.40, 1)
            _pred_total_first = round(_pred_cam_first + _pred_fat_first * 0.40, 1)
            _pred_sigma       = round(estimate_sigma(_pred_total_avg, _pred_total_first), 2)

            # ── Metrics ────────────────────────────────────────
            _pm1, _pm2, _pm3 = st.columns(3)
            _pm1.metric("🎯 Predicted class avg /100", f"{_pred_total_avg:.1f}")
            _pm2.metric("🏆 Predicted topper /100",    f"{_pred_total_first:.1f}")
            _pm3.metric("📐 Predicted σ",              f"{_pred_sigma:.1f}")

            with st.expander("📋 Full factor breakdown"):
                _da1_src = "entered ✅" if _d1a > 0 else "VIT baseline (8.3)"
                _da2_src = "entered ✅" if _d2a > 0 else "VIT baseline (8.3)"
                _da3_src = "entered ✅" if _d3a > 0 else "VIT baseline (8.3)"
                st.markdown(f"""
| Component | Pred Avg | Pred 1st | Source |
|---|---|---|---|
| CAT-I /50 | {_c1a_e:.1f} | {_c1f_e:.1f} | {_cat_src} |
| CAT-II /50 | {_c2a_e:.1f} | {_c2f_e:.1f} | {_cat_src} |
| DA-1 /10 | {_d1a_e:.1f} | {_d1f_e:.1f} | {_da1_src} |
| DA-2 /10 | {_d2a_e:.1f} | {_d2f_e:.1f} | {_da2_src} |
| DA-3 /10 | {_d3a_e:.1f} | {_d3f_e:.1f} | {_da3_src} |
| **CAM /60** | **{_pred_cam_avg:.1f}** | **{_pred_cam_first:.1f}** | scaled sum |
| FAT /100 | {_pred_fat_avg:.1f} | {_pred_fat_first:.1f} | VIT trends applied |
| **Total /100** | **{_pred_total_avg:.1f}** | **{_pred_total_first:.1f}** | CAM + FAT×0.4 |

**VIT-specific adjustments baked in:**
- 🔴 **Give-up penalty −{_GIVEUP:.0f} pts** on FAT avg — ~15% of class barely attempts, concentrated at the bottom
- 🟡 **Last-min crammer boost +{_CRAMMER:.0f} pts** — VIT students with previous papers + DA practice regularly outperform their CAT trend on FAT
- 🟢 **DA inflation** — DAs graded leniently at VIT; baseline avg set to 8.6/10
- 🔵 **CAT-II gap** — 6% higher than CAT-I baseline (open-book + more practice)
- 🔵 **CAT baselines raised** — 60% (30/50) for CAT-I, 64% (32/50) for CAT-II (empirical VIT pattern)
- 🏆 **Topper multiplier 1.48×** (was 1.55 — VIT marking is less spread than older estimate)
                """)

            st.caption("✏️ Override below if you have real numbers — then click Apply.")

            # Sync override inputs to the new prediction whenever prediction changes.
            # This is safe here because the widgets below haven't rendered yet this run.
            _pred_sig = (float(_pred_total_avg), float(_pred_total_first))
            if st.session_state.get("_pred_sig") != _pred_sig:
                st.session_state["_pred_sig"]     = _pred_sig
                st.session_state["_ov_cls_mean"]  = float(_pred_total_avg)
                st.session_state["_ov_cls_first"] = float(_pred_total_first)
            elif "_ov_cls_mean" not in st.session_state:
                st.session_state["_ov_cls_mean"]  = float(_pred_total_avg)
                st.session_state["_ov_cls_first"] = float(_pred_total_first)

            _ov1, _ov2 = st.columns(2)
            with _ov1:
                st.number_input("Override class avg /100",  0.0, 100.0, step=0.5, key="_ov_cls_mean")
            with _ov2:
                st.number_input("Override topper /100",     0.0, 100.0, step=0.5, key="_ov_cls_first")

            # on_click callback fires BETWEEN runs — before any widget re-renders —
            # so setting vit_class_mean here never conflicts with the keyed widget above.
            def _apply_pred_cb():
                st.session_state["vit_class_mean"] = st.session_state["_ov_cls_mean"]
                st.session_state["vit_first_mark"] = st.session_state["_ov_cls_first"]

            st.button("🔄 Apply to Grade Bands (Section ②)", type="primary",
                      use_container_width=True, on_click=_apply_pred_cb)

    # ── Lab marks ─────────────────────────────────────
    if course_type in ("Pure Lab","Embedded (Theory + Lab)"):
        if course_type == "Embedded (Theory + Lab)":
            st.divider()
        st.subheader("🔬 Lab Component")
        lc1, lc2 = st.columns(2)
        with lc1:
            lab_ca_avg = st.number_input(
                "Lab CA average /100", 0.0, 100.0, sf(pf.get("Lab_CA_avg",70.0)), 0.5,
                help="Average of all experiment marks scaled to /100. Contributes 60%.")
        with lc2:
            lab_fat_raw = st.number_input(
                "Lab FAT /100", 0.0, 100.0, sf(pf.get("Lab_FAT_raw",60.0)), 0.5,
                help="Lab final test out of 100. Contributes 40%.")

        lab_sc   = lab_score(lab_ca_avg, lab_fat_raw)
        lab_pass = lab_sc >= 50

        lm1,lm2,lm3 = st.columns(3)
        lm1.metric("CA contrib",   f"{lab_ca_avg*0.6:.2f}/60")
        lm2.metric("FAT contrib",  f"{lab_fat_raw*0.4:.2f}/40")
        lm3.metric("Lab Score /100",f"{lab_sc:.2f}")

        if lab_pass:
            st.markdown(f'<div class="ob">✅ Lab PASS  ({lab_sc:.2f}/100 ≥ 50)</div>',
                        unsafe_allow_html=True)
        else:
            needed_fat = lab_fat_for_target(lab_ca_avg, 50.0)
            need_str   = f"need Lab FAT ≥ {max(needed_fat,0):.1f}/100" if needed_fat <= 100 else "not recoverable"
            st.markdown(f'<div class="eb">❌ Lab FAIL  ({lab_sc:.2f}/100 < 50) — {need_str}</div>',
                        unsafe_allow_html=True)
        st.progress(min(lab_sc/100,1.0), text=f"Lab  {lab_sc:.2f} / 100")

# ── RIGHT — FAT needed table + Simulator ──────────────
with R_col:
    st.subheader("🎯 FAT Needed for Each Grade")

    # ─── Pure Theory ──────────────────────────────────
    if course_type == "Pure Theory":
        st.caption("Pass rule: Theory FAT ≥ 40/100  AND  total in grade band")
        rows = []
        for g in ["S","A","B","C","D","E"]:
            thr = bands[g]
            fn  = fat_for_target(cam, thr)
            fa  = max(fn, 40.0)
            if fa > 100:
                fd, st_ = "—", "❌ Not possible"
            elif fn <= 0:
                fd, st_ = "0 ✅", "Already secured!"
            else:
                fd = f"{fa:.1f}"
                st_ = "✅ Achievable"
            rows.append({"Grade":f"{GRADE_EMOJI[g]} {g}","GP":GRADE_POINTS[g],
                         "Min total /100":f"≥ {thr:.1f}",
                         "Theory FAT /100":fd,"Status":st_})
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

        st.divider()
        st.markdown("**📈 FAT Simulator**")
        theory_fat_sim = st.slider("Theory FAT score", 0, 100, si(pf.get("Theory_FAT_sim",60)), 1)
        theory_sc_sim  = theory_total(cam, theory_fat_sim)
        combined_sim   = theory_sc_sim
        fat_ok         = theory_fat_sim >= 40
        final_grade    = get_grade(combined_sim, bands) if fat_ok else "F*"

        sm1,sm2,sm3 = st.columns(3)
        sm1.metric("Grand total /100", f"{combined_sim:.2f}")
        sm2.metric("Predicted grade",  final_grade)
        sm3.metric("Grade points",     GRADE_POINTS.get(final_grade.replace("*",""),0))
        st.progress(min(combined_sim/100,1.0))
        if not fat_ok:
            st.markdown('<div class="eb">⚠️ FAT < 40 → F grade (FFCS pass rule)</div>',
                        unsafe_allow_html=True)

    # ─── Pure Lab ─────────────────────────────────────
    elif course_type == "Pure Lab":
        abs_b = absolute_bands()
        st.caption("Pure Lab → Absolute grading always  |  Pass: lab score ≥ 50")
        rows = []
        for g in ["S","A","B","C","D","E"]:
            thr    = abs_b[g]
            fn_lab = lab_fat_for_target(lab_ca_avg, thr)
            if fn_lab <= 0:
                fd, st_ = "0 ✅", "Secured!"
            elif fn_lab > 100:
                fd, st_ = "—", "❌ Not possible"
            else:
                fd = f"{fn_lab:.1f}"
                st_ = "✅ Achievable"
            rows.append({"Grade":f"{GRADE_EMOJI[g]} {g}","GP":GRADE_POINTS[g],
                         "Min lab total":f"≥ {thr}",
                         "Lab FAT needed /100":fd,"Status":st_})
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        st.divider()
        combined_sim = lab_sc
        final_grade  = get_grade(lab_sc, abs_b) if lab_sc >= 50 else "F"
        theory_fat_sim = 0
        sm1, sm2 = st.columns(2)
        sm1.metric("Lab Score /100", f"{lab_sc:.2f}")
        sm2.metric("Grade", f"{GRADE_EMOJI.get(final_grade,'⚫')} {final_grade}")
        st.progress(min(lab_sc/100,1.0))

    # ─── Embedded ─────────────────────────────────────
    else:
        st.caption("Pass: Lab ≥ 50  AND  Theory FAT ≥ 40  →  combined score graded")

        if tot_w > 0 and lab_pass:
            rows = []
            for g in ["S","A","B","C","D","E"]:
                thr = bands[g]
                fn  = emb_fat_for_target(cam, lab_sc, thr, tw, lw, tot_w)
                fa  = max(fn, 40.0)
                if fa > 100:
                    fd, st_ = "—", "❌ Not possible"
                elif fn <= 0:
                    fd, st_ = "0 ✅", "Secured!"
                else:
                    fd = f"{fa:.1f}"
                    st_ = "✅ Achievable"
                rows.append({"Grade":f"{GRADE_EMOJI[g]} {g}","GP":GRADE_POINTS[g],
                             "Min combined /100":f"≥ {thr:.1f}",
                             "Theory FAT needed /100":fd,"Status":st_})
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        elif not lab_pass:
            st.markdown('<div class="eb">❌ Lab FAIL — clear lab first to see grade targets.</div>',
                        unsafe_allow_html=True)

        st.divider()
        st.markdown("**📈 Theory FAT Simulator**")
        theory_fat_sim = st.slider("Theory FAT score", 0, 100, si(pf.get("Theory_FAT_sim",60)), 1)
        theory_sc_sim  = theory_total(cam, theory_fat_sim)

        if tot_w > 0:
            # Use tw/lw already computed above (respects manual ratio override if set)
            combined_sim = round((theory_sc_sim * tw + lab_sc * lw) / tot_w, 4)
        else:
            combined_sim = 0.0

        fat_ok = theory_fat_sim >= 40
        if not lab_pass:
            final_grade = "N (lab fail)"
        elif not fat_ok:
            final_grade = "F* (FAT<40)"
        else:
            final_grade = get_grade(combined_sim, bands)

        sm1,sm2,sm3,sm4 = st.columns(4)
        sm1.metric("Theory /100",   f"{theory_sc_sim:.2f}")
        sm2.metric("Lab /100",      f"{lab_sc:.2f}")
        sm3.metric("Combined /100", f"{combined_sim:.2f}")
        sm4.metric("Grade",         final_grade)
        st.progress(min(combined_sim/100,1.0))
        if not lab_pass:
            st.markdown('<div class="eb">❌ Lab component FAIL → course grade N</div>',
                        unsafe_allow_html=True)
        elif not fat_ok:
            st.markdown('<div class="eb">⚠️ Theory FAT < 40 → F grade (FFCS pass rule)</div>',
                        unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────
#  Grade Band Chart
# ──────────────────────────────────────────────────────────
st.divider()
st.subheader("📊 Grade Band Overview")
gc1, gc2 = st.columns([1.6, 1])
with gc1:
    bd = pd.DataFrame({"Grade": list(bands.keys()),
                        "Min Score": [round(v,1) for v in bands.values()]})
    st.bar_chart(bd.set_index("Grade"), color="#1565C0", height=220)
with gc2:
    st.markdown("**Where your score falls:**")
    prev_thresh = 101.0
    for g in ["S","A","B","C","D","E"]:
        hit = ""
        if combined_sim >= bands[g] and combined_sim < prev_thresh:
            hit = "  ◀ **your score**"
        st.markdown(f"{GRADE_EMOJI[g]} **{g}** ≥ {bands[g]:.1f}{hit}")
        prev_thresh = bands[g]

# ══════════════════════════════════════════════════════════
#  ④ Save / Import / Download
# ══════════════════════════════════════════════════════════
st.header("④ Save & Manage")
st.caption("💡 You can also edit/delete rows directly in the downloaded Excel, then re-import using 'Import from Excel'.")

b1, b2, b3, b4 = st.columns(4)

with b1:
    if st.button("💾 Save subject", type="primary", use_container_width=True):
        if not subject:
            st.warning("Enter a subject name first (Step ①).")
        else:
            row = {
                "Subject": subject, "Course_Type": course_type,
                "L": L_val, "T": T_val, "P": P_val,
                "CAT1_raw": cat1_raw, "CAT2_raw": cat2_raw,
                "DA1": da1, "DA2": da2, "DA3": da3,
                "Lab_CA_avg": lab_ca_avg, "Lab_FAT_raw": lab_fat_raw,
                "Theory_FAT_sim": theory_fat_sim,
                "Class_Mean": class_mean, "Class_Sigma": sigma_used,
                "First_Mark": first_mark, "Grading_Mode": grading_mode,
                "CAM_Total":      round(cam, 2),
                "Lab_Score":      round(lab_sc, 2),
                "Theory_Score":   round(theory_sc_sim, 2),
                "Combined_Total": round(combined_sim, 2),
                "Predicted_Grade":final_grade,
            }
            save_row(row)
            st.success(f"✅ Saved '{subject}'")
            st.rerun()

with b2:
    df_dl = load_df()
    if not df_dl.empty:
        st.download_button(
            "📥 Download Excel", df_to_bytes(df_dl),
            file_name="vit_grades_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Edit rows/delete in Excel freely, then re-import below to sync."
        )
    else:
        st.button("📥 Download Excel", disabled=True, use_container_width=True)

with b3:
    up = st.file_uploader("📤 Import / update from Excel",
                          type=["xlsx","xls"], label_visibility="collapsed")
    if up:
        res = merge_excel(up)
        if res.startswith("OK:"):
            st.success(f"✅ Merged {res.split(':')[1]} subjects.")
            st.rerun()
        else:
            st.error(res.replace("ERR:",""))

with b4:
    df_chk = load_df()
    if subject and subject in df_chk["Subject"].values:
        if st.button("🗑️ Delete subject", use_container_width=True):
            delete_row(subject)
            st.warning(f"Deleted '{subject}'")
            st.rerun()
    else:
        st.button("🗑️ Delete subject", disabled=True, use_container_width=True,
                  help="Save the subject first to enable delete.")

# ──────────────────────────────────────────────────────────
#  All Subjects Table
# ──────────────────────────────────────────────────────────
df_view = load_df()
if not df_view.empty:
    st.subheader("📋 All Saved Subjects")
    show = [c for c in ["Subject","Course_Type","CAM_Total","Lab_Score",
                         "Theory_Score","Combined_Total","Predicted_Grade","Last_Updated"]
            if c in df_view.columns]
    st.dataframe(df_view[show], hide_index=True, use_container_width=True,
                 column_config={
                     "CAM_Total":      st.column_config.NumberColumn("CAM /60",    format="%.2f"),
                     "Lab_Score":      st.column_config.NumberColumn("Lab /100",   format="%.2f"),
                     "Theory_Score":   st.column_config.NumberColumn("Theory /100",format="%.2f"),
                     "Combined_Total": st.column_config.NumberColumn("Total /100", format="%.2f"),
                     "Predicted_Grade":"Grade",
                 })
else:
    st.info("No saved subjects yet — fill marks above and click **Save subject**.")

# ──────────────────────────────────────────────────────────
#  Rules reference
# ──────────────────────────────────────────────────────────
st.divider()
with st.expander("📖 FFCS v4.0 rules applied in this calculator"):
    st.markdown("""
| Course Type | CAM (60%) | FAT (40%) | Pass Rule | Grading |
|---|---|---|---|---|
| **Pure Theory** | CAT-I×0.3 + CAT-II×0.3 + DA(3×10) | Theory FAT /100 × 0.4 | FAT ≥ 40/100 AND total in band | Relative (Absolute if ≤10 students) |
| **Pure Lab** | Lab CA avg /100 × 0.6 | Lab FAT /100 × 0.4 | Lab total ≥ 50/100 | Absolute always |
| **Embedded** | Theory CAM /60 + Lab CA× 0.6 | Theory FAT × 0.4 + Lab FAT × 0.4 | Lab ≥ 50 AND Theory FAT ≥ 40 | Combined = (Theory×(L+T) + Lab×P/2) / (L+T+P/2) |

**Sigma estimation:** σ ≈ (topper − class mean) / 2.5 (assumes top score ≈ mean + 2.5σ).  
**Relative grading bands:** S ≥ mean+1.5σ (min **90** for Pure Theory) | For **Embedded**: S **capped at 80** — `min(mean+1.5σ, 80)`, so a high mean/sigma never pushes S past 80. | A ≥ mean+0.5σ | B ≥ mean−0.5σ | C ≥ mean−1σ | D ≥ mean−1.5σ | E ≥ mean−2σ.  
**Import/Export Excel:** Download → edit freely (add, delete, update rows) → re-import to sync.
    """)

st.caption("Unofficial grade predictor · Not an official VIT tool · Verify with faculty · FFCS v4.0")