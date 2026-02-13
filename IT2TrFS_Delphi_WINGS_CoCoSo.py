import streamlit as st
import numpy as np
import pandas as pd
import graphviz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io, base64
import matplotlib.pyplot as plt

# ============================================================
# 1) IT2TrFS core utilities (used by BOTH WINGS + CoCoSo)
# ============================================================

def format_it2(it2):
    u, l = it2
    return f"(({u[0]:.6f},{u[1]:.6f},{u[2]:.6f},{u[3]:.6f};{u[4]:.2f},{u[5]:.2f}), ({l[0]:.6f},{l[1]:.6f},{l[2]:.6f},{l[3]:.6f};{l[4]:.2f},{l[5]:.2f}))"

def zero_it2():
    return ((0,0,0,0,1,1), (0,0,0,0,0.9,0.9))

def add_it2(A, B):
    Au, Al = A; Bu, Bl = B
    new_u = (Au[0]+Bu[0], Au[1]+Bu[1], Au[2]+Bu[2], Au[3]+Bu[3], min(Au[4],Bu[4]), min(Au[5],Bu[5]))
    new_l = (Al[0]+Bl[0], Al[1]+Bl[1], Al[2]+Bl[2], Al[3]+Bl[3], min(Al[4],Bl[4]), min(Al[5],Bl[5]))
    return (new_u, new_l)

def sub_it2(A, B):
    Au, Al = A; Bu, Bl = B
    new_u = (Au[0]-Bu[0], Au[1]-Bu[1], Au[2]-Bu[2], Au[3]-Bu[3], min(Au[4],Bu[4]), min(Au[5],Bu[5]))
    new_l = (Al[0]-Bl[0], Al[1]-Bl[1], Al[2]-Bl[2], Al[3]-Bl[3], min(Al[4],Bl[4]), min(Al[5],Bl[5]))
    return (new_u, new_l)

def mul_it2(A, B):
    Au, Al = A; Bu, Bl = B
    new_u = (Au[0]*Bu[0], Au[1]*Bu[1], Au[2]*Bu[2], Au[3]*Bu[3], min(Au[4],Bu[4]), min(Au[5],Bu[5]))
    new_l = (Al[0]*Bl[0], Al[1]*Bl[1], Al[2]*Bl[2], Al[3]*Bl[3], min(Al[4],Bl[4]), min(Al[5],Bl[5]))
    return (new_u, new_l)

def scalar_mul_it2(k, A):
    # scale a,b,c,d only; keep heights
    Au, Al = A
    new_u = (k*Au[0], k*Au[1], k*Au[2], k*Au[3], Au[4], Au[5])
    new_l = (k*Al[0], k*Al[1], k*Al[2], k*Al[3], Al[4], Al[5])
    return (new_u, new_l)

def defuzz_it2(A):
    Au, Al = A
    return (Au[0]+Au[1]+Au[2]+Au[3]+Al[0]+Al[1]+Al[2]+Al[3]) / 8

def it2_weighted_avg(it2_list, weights):
    if len(it2_list) != len(weights) or len(it2_list) == 0:
        return None
    if not np.isclose(sum(weights), 1.0):
        return None
    out = zero_it2()
    for it2, w in zip(it2_list, weights):
        out = add_it2(out, scalar_mul_it2(w, it2))
    return out

def it2_weighted_geo(it2_list, weights):
    """
    Weighted geometric aggregation for IT2TrFS (parameter-wise power product).
    Heights are combined conservatively via min.
    """
    if len(it2_list) != len(weights) or len(it2_list) == 0:
        return None
    if not np.isclose(sum(weights), 1.0):
        return None

    eps = 1e-12
    u_a = np.prod([max(it2[0][0], eps)**w for it2, w in zip(it2_list, weights)])
    u_b = np.prod([max(it2[0][1], eps)**w for it2, w in zip(it2_list, weights)])
    u_c = np.prod([max(it2[0][2], eps)**w for it2, w in zip(it2_list, weights)])
    u_d = np.prod([max(it2[0][3], eps)**w for it2, w in zip(it2_list, weights)])

    l_a = np.prod([max(it2[1][0], eps)**w for it2, w in zip(it2_list, weights)])
    l_b = np.prod([max(it2[1][1], eps)**w for it2, w in zip(it2_list, weights)])
    l_c = np.prod([max(it2[1][2], eps)**w for it2, w in zip(it2_list, weights)])
    l_d = np.prod([max(it2[1][3], eps)**w for it2, w in zip(it2_list, weights)])

    uh1 = min([it2[0][4] for it2 in it2_list]); uh2 = min([it2[0][5] for it2 in it2_list])
    lh1 = min([it2[1][4] for it2 in it2_list]); lh2 = min([it2[1][5] for it2 in it2_list])

    return ((u_a, u_b, u_c, u_d, uh1, uh2),
            (l_a, l_b, l_c, l_d, lh1, lh2))


# ============================================================
# 2) IT2TrFS-CoCoSo linguistic scale (YOUR EXACT VALUES)
# ============================================================

COCOSO_SCALE = [
    ("Very Poor",   "VP", ((0,   0,   0,   0.1, 1, 1), (0.05,0,   0,   0.05,0.9,0.9))),
    ("Poor",        "P",  ((0,   0.1, 0.1, 0.3, 1, 1), (0.05,0.1, 0.1, 0.25,0.9,0.9))),
    ("Medium Poor", "MP", ((0.1, 0.3, 0.3, 0.5, 1, 1), (0.15,0.3, 0.3, 0.45,0.9,0.9))),
    ("Fair",        "F",  ((0.3, 0.5, 0.5, 0.7, 1, 1), (0.35,0.5, 0.5, 0.65,0.9,0.9))),
    ("Medium Good", "MG", ((0.5, 0.7, 0.7, 0.9, 1, 1), (0.55,0.7, 0.7, 0.85,0.9,0.9))),
    ("Good",        "G",  ((0.7, 0.9, 0.9, 1.0, 1, 1), (0.75,0.9, 0.9, 0.95,0.9,0.9))),
    ("Very good",   "VG", ((0.9, 1.0, 1.0, 1.0, 1, 1), (0.95,1.0, 1.0, 0.95,0.9,0.9))),
]

COCOSO_CODE_TO_IT2 = {code: it2 for _, code, it2 in COCOSO_SCALE}
COCOSO_CODES = list(COCOSO_CODE_TO_IT2.keys())

def cocoso_scale_table():
    rows = []
    for name, code, it2 in COCOSO_SCALE:
        u, l = it2
        rows.append({
            "Linguistic Attribute": name,
            "Code": code,
            "IT2TrFS": format_it2(it2)
        })
    return pd.DataFrame(rows)


# ============================================================
# 3) IT2TrFS–CoCoSo module (manual input like IVIFN app)
# ============================================================

def module_cocoso():
    st.header("📊 IT2TrFS–CoCoSo (Manual input)")

    with st.expander("Show Linguistic Scale Reference (fixed)", expanded=True):
        st.dataframe(cocoso_scale_table(), hide_index=True, use_container_width=True)

    st.subheader("Step 1: Define Alternatives and Criteria")
    c1, c2 = st.columns(2)
    alts_in = c1.text_input("Alternatives (comma-separated)", "T1, T2, T3", key="cocoso_alts_in")
    crits_in = c2.text_input("Criteria (comma-separated)", "C1, C2, C3", key="cocoso_crits_in")
    alternatives = [a.strip() for a in alts_in.split(",") if a.strip()]
    criteria = [c.strip() for c in crits_in.split(",") if c.strip()]

    if len(alternatives) < 2 or len(criteria) < 1:
        st.warning("Please enter at least 2 alternatives and at least 1 criterion.")
        return

    st.subheader("Step 2: Criterion Type & Weights")
    if "cocoso_crit_df" not in st.session_state or list(st.session_state.cocoso_crit_df["Criterion"]) != criteria:
        w = [round(1/len(criteria), 6)] * len(criteria)
        if len(w) > 1:
            w[-1] = 1.0 - sum(w[:-1])
        st.session_state.cocoso_crit_df = pd.DataFrame({
            "Criterion": criteria,
            "Type": ["Benefit"] * len(criteria),
            "Weight": w
        })

    st.session_state.cocoso_crit_df = st.data_editor(
        st.session_state.cocoso_crit_df,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=["Benefit","Cost"]),
            "Weight": st.column_config.NumberColumn("Weight", min_value=0.0, max_value=1.0, format="%.6f"),
        },
        key="cocoso_crit_editor"
    )

    crit_types = st.session_state.cocoso_crit_df["Type"].tolist()
    crit_w = st.session_state.cocoso_crit_df["Weight"].astype(float).tolist()

    if not np.isclose(sum(crit_w), 1.0):
        st.error(f"Criterion weights must sum to 1. Current sum: {sum(crit_w):.6f}")
        st.stop()

    st.subheader("Step 3: Experts")
    n_exp = st.number_input("Number of experts", min_value=1, max_value=20, value=2, step=1, key="cocoso_nexp")

    st.markdown("**Expert weights** (must sum to 1.0)")
    if n_exp > 1:
        cols = st.columns(n_exp)
        exp_w = []
        for i in range(n_exp):
            with cols[i]:
                exp_w.append(st.number_input(f"E{i+1}", 0.0, 1.0, value=1.0/n_exp, step=0.05, format="%.2f", key=f"cocoso_expw_{i}"))
        if not np.isclose(sum(exp_w), 1.0):
            st.error(f"Expert weights must sum to 1.0. Current sum: {sum(exp_w):.2f}")
            st.stop()
    else:
        exp_w = [1.0]

    st.subheader("Step 4: Expert Evaluation Matrices (dropdown codes)")
    if "cocoso_expert_mats" not in st.session_state:
        st.session_state.cocoso_expert_mats = {}

    # reset if dimensions changed
    if len(st.session_state.cocoso_expert_mats) != n_exp:
        st.session_state.cocoso_expert_mats = {i: pd.DataFrame("F", index=alternatives, columns=criteria) for i in range(n_exp)}
    else:
        # ensure correct index/columns
        for i in range(n_exp):
            if not isinstance(st.session_state.cocoso_expert_mats.get(i), pd.DataFrame) or \
               set(st.session_state.cocoso_expert_mats[i].index) != set(alternatives) or \
               set(st.session_state.cocoso_expert_mats[i].columns) != set(criteria):
                st.session_state.cocoso_expert_mats[i] = pd.DataFrame("F", index=alternatives, columns=criteria)

    tabs = st.tabs([f"Expert {i+1}" for i in range(n_exp)])
    for i, tab in enumerate(tabs):
        with tab:
            st.session_state.cocoso_expert_mats[i] = st.data_editor(
                st.session_state.cocoso_expert_mats[i],
                use_container_width=True,
                column_config={c: st.column_config.SelectboxColumn(c, options=COCOSO_CODES) for c in criteria},
                key=f"cocoso_mat_{i}"
            )

    st.subheader("Step 5: Run CoCoSo Calculation")
    if st.button("✅ Calculate CoCoSo Ranking", type="primary", use_container_width=True, key="cocoso_run"):
        # 5.1 aggregated IT2 decision matrix
        agg_it2 = pd.DataFrame(index=alternatives, columns=criteria, dtype=object)
        for a in alternatives:
            for c in criteria:
                it2s = []
                for ei in range(n_exp):
                    code = str(st.session_state.cocoso_expert_mats[ei].loc[a, c]).strip()
                    it2s.append(COCOSO_CODE_TO_IT2[code])
                agg_it2.loc[a, c] = it2_weighted_avg(it2s, exp_w)

        st.markdown("### 5.1 Aggregated IT2TrFS decision matrix")
        st.dataframe(agg_it2.applymap(format_it2), use_container_width=True)

        # 5.2 crisp defuzz
        crisp = agg_it2.applymap(defuzz_it2)
        st.markdown("### 5.2 Crisp matrix (defuzzified)")
        st.dataframe(crisp.style.format(precision=6), use_container_width=True)

        # 5.3 normalization (crisp CoCoSo normalization)
        norm = crisp.copy()
        for j, c in enumerate(criteria):
            col = crisp[c].astype(float)
            if crit_types[j] == "Benefit":
                mx = np.nanmax(col.values)
                norm[c] = col / mx if mx != 0 else 0.0
            else:
                mn = np.nanmin(col.values)
                norm[c] = mn / col.replace(0, np.nan)
                norm[c] = norm[c].fillna(0.0)

        st.markdown("### 5.3 Normalized crisp matrix")
        st.dataframe(norm.style.format(precision=6), use_container_width=True)

        # 5.4 S and P
        w = np.array(crit_w, dtype=float)
        S = norm.values @ w
        P = np.prod(np.power(np.maximum(norm.values, 1e-12), w), axis=1)

        res = pd.DataFrame({"Alternative": alternatives, "S": S, "P": P})
        st.markdown("### 5.4 CoCoSo S and P")
        st.dataframe(res.style.format(precision=6), use_container_width=True, hide_index=True)

        # 5.5 final scores (Kia/Kib/Kic/K)
        s = res["S"].values
        p = res["P"].values

        Kia = (s + p) / np.sum(s + p) if np.sum(s + p) != 0 else np.zeros_like(s)
        Kib = (s / np.min(s) if np.min(s) != 0 else 0) + (p / np.min(p) if np.min(p) != 0 else 0)

        tau = 0.5
        denom = tau*np.max(s) + (1-tau)*np.max(p)
        Kic = (tau*s + (1-tau)*p) / denom if denom != 0 else np.zeros_like(s)

        K = np.power(Kia*Kib*Kic, 1/3) + (Kia + Kib + Kic)/3

        final = pd.DataFrame({"Alternative": alternatives, "Kia": Kia, "Kib": Kib, "Kic": Kic, "K": K})
        final["Rank"] = final["K"].rank(ascending=False, method="min").astype(int)
        final = final.sort_values("Rank").reset_index(drop=True)

        st.markdown("### 5.5 Final CoCoSo ranking")
        st.dataframe(final.style.format(precision=6), use_container_width=True, hide_index=True)

        with st.expander("Ranking chart", expanded=False):
            fig, ax = plt.subplots(figsize=(9, 5))
            ax.bar(final["Alternative"], final["K"])
            ax.set_ylabel("K (final score)")
            ax.set_title("IT2TrFS–CoCoSo final scores")
            st.pyplot(fig)


# ============================================================
# 4) Minimal WINGS module placeholder
#    (Keep your full WINGS code here; not modified in this snippet)
# ============================================================

def module_wings_placeholder():
    st.warning("Paste your full IT2TrFS–WINGS code here (kept unchanged).")
    st.info("This file shows the revised IT2TrFS–CoCoSo linguistic scale and module.")


# ============================================================
# 5) MAIN APP (sidebar 2 options)
# ============================================================

def main():
    st.set_page_config(page_title="IT2TrFS Toolkit (WINGS + CoCoSo)", layout="wide", page_icon="📊")
    st.title("🧰 IT2TrFS Toolkit")

    st.sidebar.header("Navigation")
    page = st.sidebar.radio("Choose a Model", ["IT2TrFS–WINGS", "IT2TrFS–CoCoSo"], index=1)

    if page == "IT2TrFS–WINGS":
        module_wings_placeholder()
    else:
        module_cocoso()

if __name__ == "__main__":
    main()
