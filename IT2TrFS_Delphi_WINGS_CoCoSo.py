import streamlit as st
import numpy as np
import pandas as pd
import graphviz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import base64
import matplotlib.pyplot as plt

# =========================================================
# IT2TrFS REPRESENTATION
#   IT2 = (UMF, LMF)
#   UMF = (a,b,c,d,uh1,uh2)
#   LMF = (e,f,g,h,lh1,lh2)
# =========================================================

def format_it2(it2):
    u, l = it2
    return f"(({u[0]:.6f},{u[1]:.6f},{u[2]:.6f},{u[3]:.6f};{u[4]:.1f},{u[5]:.1f}), ({l[0]:.6f},{l[1]:.6f},{l[2]:.6f},{l[3]:.6f};{l[4]:.1f},{l[5]:.1f}))"

def zero_it2():
    return ((0,0,0,0,1,1), (0,0,0,0,0.9,0.9))

def add_it2(A, B):
    Au, Al = A
    Bu, Bl = B
    new_u = (Au[0] + Bu[0], Au[1] + Bu[1], Au[2] + Bu[2], Au[3] + Bu[3], min(Au[4], Bu[4]), min(Au[5], Bu[5]))
    new_l = (Al[0] + Bl[0], Al[1] + Bl[1], Al[2] + Bl[2], Al[3] + Bl[3], min(Al[4], Bl[4]), min(Al[5], Bl[5]))
    return (new_u, new_l)

def sub_it2(A, B):
    Au, Al = A
    Bu, Bl = B
    new_u = (Au[0] - Bu[0], Au[1] - Bu[1], Au[2] - Bu[2], Au[3] - Bu[3], min(Au[4], Bu[4]), min(Au[5], Bu[5]))
    new_l = (Al[0] - Bl[0], Al[1] - Bl[1], Al[2] - Bl[2], Al[3] - Bl[3], min(Al[4], Bl[4]), min(Al[5], Bl[5]))
    return (new_u, new_l)

def mul_it2(A, B):
    Au, Al = A
    Bu, Bl = B
    new_u = (Au[0] * Bu[0], Au[1] * Bu[1], Au[2] * Bu[2], Au[3] * Bu[3], min(Au[4], Bu[4]), min(Au[5], Bu[5]))
    new_l = (Al[0] * Bl[0], Al[1] * Bl[1], Al[2] * Bl[2], Al[3] * Bl[3], min(Al[4], Bl[4]), min(Al[5], Bl[5]))
    return (new_u, new_l)

def scalar_mul_it2(k, A):
    Au, Al = A
    new_u = (k * Au[0], k * Au[1], k * Au[2], k * Au[3], Au[4], Au[5])
    new_l = (k * Al[0], k * Al[1], k * Al[2], k * Al[3], Al[4], Al[5])
    return (new_u, new_l)

def it2_pow(A, w):
    """
    Excel-like power applied parameter-wise:
      a' = a^w, b' = b^w, ... h' = h^w
    Heights unchanged.
    """
    Au, Al = A

    def pw(x):
        # Excel handles 0^0 as 1; we mimic safely
        if x == 0 and w == 0:
            return 1.0
        return float(x) ** float(w)

    new_u = (pw(Au[0]), pw(Au[1]), pw(Au[2]), pw(Au[3]), Au[4], Au[5])
    new_l = (pw(Al[0]), pw(Al[1]), pw(Al[2]), pw(Al[3]), Al[4], Al[5])
    return (new_u, new_l)

# =========================================================
# CoCoSo DEFUZZIFICATION (EXCEL MATCH)
#   Crisp = (Score(UMF) + Score(LMF))/2
#   Score(UMF) = a + [(d-a) + (uh2*c-a) + (uh1*b-a)]/4
#   Score(LMF) = e + [(h-e) + (lh2*g-e) + (lh1*f-e)]/4
# =========================================================

def cocoso_crisp_score(it2):
    Au, Al = it2
    a,b,c,d,uh1,uh2 = Au
    e,f,g,h,lh1,lh2 = Al

    score_u = (((d-a) + ((uh2*c) - a) + ((uh1*b) - a)) / 4.0) + a
    score_l = (((h-e) + ((lh2*g) - e) + ((lh1*f) - e)) / 4.0) + e
    return (score_u + score_l) / 2.0

# =========================================================
# IT2TrFS-CoCoSo linguistic scale (YOUR REQUIRED VALUES)
# =========================================================

COCOSO_LINGUISTIC_TERMS = {
    "VP": ((0,0,0,0.1,1,1), (0.05,0,0,0.05,0.9,0.9)),
    "P" : ((0,0.1,0.1,0.3,1,1), (0.05,0.1,0.1,0.25,0.9,0.9)),
    "MP": ((0.1,0.3,0.3,0.5,1,1), (0.15,0.3,0.3,0.45,0.9,0.9)),
    "F" : ((0.3,0.5,0.5,0.7,1,1), (0.35,0.5,0.5,0.65,0.9,0.9)),
    "MG": ((0.5,0.7,0.7,0.9,1,1), (0.55,0.7,0.7,0.85,0.9,0.9)),
    "G" : ((0.7,0.9,0.9,1.0,1,1), (0.75,0.9,0.9,0.95,0.9,0.9)),
    "VG": ((0.9,1.0,1.0,1.0,1,1), (0.95,1.0,1.0,0.95,0.9,0.9)),
}

COCOSO_FULL = {
    "VP":"Very Poor","P":"Poor","MP":"Medium Poor","F":"Fair",
    "MG":"Medium Good","G":"Good","VG":"Very Good"
}

# =========================================================
# CoCoSo NORMALIZATION (EXCEL MATCH: IT2_F_CoCoSo_F)
# =========================================================

def normalize_it2_matrix_excel(agg_matrix, criteria_types, alternatives, criteria):
    """
    agg_matrix: dict[(alt, crit)] -> IT2TrFS
    criteria_types: list of 'Benefit'/'Cost' aligned with criteria

    Benefit normalization (Excel):
      divisor = max over alternatives of max(a,b,c,d) [UMF only]
      (a,b,c,d,e,f,g,h) /= divisor

    Cost normalization (Excel):
      base_min = min over alternatives of min(a,b,c,d) [UMF only]
      then reverse divide:
        UMF: (min/d, min/c, min/b, min/a)
        LMF: (min/h, min/g, min/f, min/e)
      heights unchanged
    """
    norm = {}

    for j, crit in enumerate(criteria):
        # collect UMF parameters across alts for this criterion
        umf_params = []
        for alt in alternatives:
            Au, _ = agg_matrix[(alt, crit)]
            umf_params.extend([Au[0], Au[1], Au[2], Au[3]])

        if criteria_types[j].lower().startswith("b"):  # Benefit
            div = max(umf_params) if len(umf_params) else 1.0
            div = div if div != 0 else 1.0

            for alt in alternatives:
                Au, Al = agg_matrix[(alt, crit)]
                a,b,c,d,uh1,uh2 = Au
                e,f,g,h,lh1,lh2 = Al
                norm[(alt, crit)] = (
                    (a/div, b/div, c/div, d/div, uh1, uh2),
                    (e/div, f/div, g/div, h/div, lh1, lh2)
                )

        else:  # Cost
            base_min = min(umf_params) if len(umf_params) else 0.0
            # Excel uses the min of (a,b,c,d) across all alternatives (UMF)
            m = base_min

            for alt in alternatives:
                Au, Al = agg_matrix[(alt, crit)]
                a,b,c,d,uh1,uh2 = Au
                e,f,g,h,lh1,lh2 = Al

                # reverse divide (match Excel formulas)
                def safe_div(num, den):
                    den = float(den)
                    if den == 0:
                        return 0.0
                    return float(num) / den

                norm_umf = (safe_div(m,d), safe_div(m,c), safe_div(m,b), safe_div(m,a), uh1, uh2)
                norm_lmf = (safe_div(m,h), safe_div(m,g), safe_div(m,f), safe_div(m,e), lh1, lh2)

                norm[(alt, crit)] = (norm_umf, norm_lmf)

    return norm

def it2_to_row(it2):
    Au, Al = it2
    return {
        "a":Au[0],"b":Au[1],"c":Au[2],"d":Au[3],"uh1":Au[4],"uh2":Au[5],
        "e":Al[0],"f":Al[1],"g":Al[2],"h":Al[3],"lh1":Al[4],"lh2":Al[5],
    }

def format_it2_table(matrix_dict, alternatives, criteria, value_formatter=format_it2):
    df = pd.DataFrame(index=alternatives, columns=criteria, dtype=object)
    for alt in alternatives:
        for crit in criteria:
            df.loc[alt, crit] = value_formatter(matrix_dict[(alt, crit)])
    return df

# =========================================================
# IT2TrFS-CoCoSo APP (Excel-matching)
# =========================================================

def cocoso_app():
    st.header("📊 IT2TrFS-CoCoSo (Excel-matching)")
    st.caption("Implements the same workflow as the Excel sheet: IT2_F_CoCoSo_F (defuzzification only at the end).")

    with st.expander("Linguistic scale (VP…VG)"):
        scale_df = pd.DataFrame(
            [{"Abbr":k, "Meaning":COCOSO_FULL[k], "IT2TrFS":format_it2(v)} for k,v in COCOSO_LINGUISTIC_TERMS.items()]
        )
        st.dataframe(scale_df, hide_index=True, use_container_width=True)

    st.subheader("Step 1: Alternatives, Criteria, Types, Weights")
    c1, c2 = st.columns(2)
    alts_in = c1.text_input("Alternatives (comma-separated)", "T1, T2, T3", key="cocoso_alts_in")
    crits_in = c2.text_input("Criteria (comma-separated)", "C1, C2, C3", key="cocoso_crits_in")

    alternatives = [a.strip() for a in alts_in.split(",") if a.strip()]
    criteria = [c.strip() for c in crits_in.split(",") if c.strip()]

    if len(alternatives) < 1 or len(criteria) < 1:
        st.warning("Please provide at least 1 alternative and 1 criterion.")
        return

    # criteria table
    if "cocoso_crit_df_it2" not in st.session_state or set(st.session_state.cocoso_crit_df_it2["Criterion"]) != set(criteria):
        w = [round(1/len(criteria), 6)] * len(criteria)
        if len(criteria) > 0:
            w[-1] = 1.0 - sum(w[:-1])
        st.session_state.cocoso_crit_df_it2 = pd.DataFrame({
            "Criterion": criteria,
            "Type": ["Benefit"] * len(criteria),
            "Weight": w
        })

    edited_crit_df = st.data_editor(
        st.session_state.cocoso_crit_df_it2,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=["Benefit","Cost"]),
            "Weight": st.column_config.NumberColumn("Weight", format="%.6f", min_value=0.0, max_value=1.0, step=0.000001),
        },
        key="cocoso_crit_editor_it2"
    )

    criteria_types = edited_crit_df["Type"].tolist()
    criteria_weights = edited_crit_df["Weight"].astype(float).tolist()

    if not np.isclose(sum(criteria_weights), 1.0):
        st.error(f"Criteria weights must sum to 1.0 (now: {sum(criteria_weights):.6f}).")
        return

    st.subheader("Step 2: Expert evaluations (linguistic)")
    num_experts = st.number_input("Number of experts", min_value=1, max_value=30, value=2, step=1, key="cocoso_ne_it2")

    st.markdown("**Expert weights** (must sum to 1.0)")
    expert_weights = []
    if num_experts == 1:
        expert_weights = [1.0]
        st.info("Single expert → weight = 1.0")
    else:
        cols = st.columns(num_experts)
        for i in range(num_experts):
            with cols[i]:
                expert_weights.append(
                    st.number_input(
                        f"E{i+1}",
                        min_value=0.0, max_value=1.0,
                        value=round(1/num_experts, 6),
                        step=0.01,
                        format="%.6f",
                        key=f"cocoso_ew_{i}"
                    )
                )
        if not np.isclose(sum(expert_weights), 1.0):
            st.error(f"Expert weights must sum to 1.0 (now: {sum(expert_weights):.6f}).")
            return

    # decision matrices per expert: alternatives x criteria with linguistic abbreviations
    if "cocoso_expert_dfs_it2" not in st.session_state:
        st.session_state.cocoso_expert_dfs_it2 = {}

    need_reset = (
        len(st.session_state.cocoso_expert_dfs_it2) != num_experts
        or (num_experts > 0 and (
            set(st.session_state.cocoso_expert_dfs_it2.get(0, pd.DataFrame()).index) != set(alternatives)
            or set(st.session_state.cocoso_expert_dfs_it2.get(0, pd.DataFrame()).columns) != set(criteria)
        ))
    )
    if need_reset:
        st.session_state.cocoso_expert_dfs_it2 = {
            i: pd.DataFrame("F", index=alternatives, columns=criteria) for i in range(num_experts)
        }

    tabs = st.tabs([f"Expert {i+1}" for i in range(num_experts)])
    for i, tab in enumerate(tabs):
        with tab:
            st.session_state.cocoso_expert_dfs_it2[i] = st.data_editor(
                st.session_state.cocoso_expert_dfs_it2[i],
                use_container_width=True,
                column_config={
                    c: st.column_config.SelectboxColumn(c, options=list(COCOSO_LINGUISTIC_TERMS.keys()))
                    for c in criteria
                },
                key=f"cocoso_editor_it2_{i}"
            )

    st.subheader("Step 3: Calculate (Excel workflow)")
    tau = st.number_input("τ (tau)", min_value=0.0, max_value=1.0, value=0.5, step=0.05, key="cocoso_tau")

    if st.button("✅ Run IT2TrFS-CoCoSo", type="primary", use_container_width=True, key="cocoso_run_it2"):
        with st.spinner("Computing..."):

            # -------------------------------------------------
            # 3.1 Aggregate expert matrices -> IT2TrFS matrix
            # (no defuzzification here)
            # -------------------------------------------------
            agg_matrix = {}
            for alt in alternatives:
                for crit in criteria:
                    # weighted average of parameters
                    acc = None
                    for e in range(num_experts):
                        term = st.session_state.cocoso_expert_dfs_it2[e].loc[alt, crit]
                        it2 = COCOSO_LINGUISTIC_TERMS[term]
                        it2w = scalar_mul_it2(expert_weights[e], it2)
                        acc = it2w if acc is None else add_it2(acc, it2w)
                    agg_matrix[(alt, crit)] = acc

            st.markdown("#### 3.1 Aggregated IT2TrFS Decision Matrix (NO defuzz)")
            st.dataframe(format_it2_table(agg_matrix, alternatives, criteria), use_container_width=True)

            # -------------------------------------------------
            # 3.2 Normalize (Excel match)
            # -------------------------------------------------
            norm_matrix = normalize_it2_matrix_excel(
                agg_matrix=agg_matrix,
                criteria_types=criteria_types,
                alternatives=alternatives,
                criteria=criteria
            )

            st.markdown("#### 3.2 Normalized IT2TrFS Matrix (Excel rules)")
            st.dataframe(format_it2_table(norm_matrix, alternatives, criteria), use_container_width=True)

            # -------------------------------------------------
            # 3.3 SBi and PBi in IT2 domain (NO defuzz)
            #   SBi(param) = Σ wj * r_ij(param)
            #   PBi(param) = Σ (r_ij(param) ^ wj)   <-- Excel
            # -------------------------------------------------
            SBi = {}
            PBi = {}
            for alt in alternatives:
                s_acc = zero_it2()
                p_acc = zero_it2()
                for j, crit in enumerate(criteria):
                    r = norm_matrix[(alt, crit)]
                    wj = float(criteria_weights[j])

                    s_acc = add_it2(s_acc, scalar_mul_it2(wj, r))
                    p_acc = add_it2(p_acc, it2_pow(r, wj))

                SBi[alt] = s_acc
                PBi[alt] = p_acc

            sbi_df = pd.DataFrame([{"Alternative":alt, **it2_to_row(SBi[alt])} for alt in alternatives])
            pbi_df = pd.DataFrame([{"Alternative":alt, **it2_to_row(PBi[alt])} for alt in alternatives])

            st.markdown("#### 3.3 SBi (IT2TrFS) — no defuzz")
            st.dataframe(sbi_df.style.format(precision=6), use_container_width=True, hide_index=True)

            st.markdown("#### 3.3 PBi (IT2TrFS) — Excel uses Σ(r^w), no defuzz")
            st.dataframe(pbi_df.style.format(precision=6), use_container_width=True, hide_index=True)

            # -------------------------------------------------
            # 3.4 Defuzzification at the END (Excel match)
            # Crisp SBi, Crisp PBi
            # -------------------------------------------------
            crisp_S = {alt: cocoso_crisp_score(SBi[alt]) for alt in alternatives}
            crisp_P = {alt: cocoso_crisp_score(PBi[alt]) for alt in alternatives}

            df_crisp = pd.DataFrame({
                "Alternative": alternatives,
                "Crisp SBi": [crisp_S[a] for a in alternatives],
                "Crisp PBi": [crisp_P[a] for a in alternatives],
            })

            st.markdown("#### 3.4 Crisp SBi & Crisp PBi (defuzz at end only)")
            st.dataframe(df_crisp.style.format(precision=6), use_container_width=True, hide_index=True)

            # -------------------------------------------------
            # 3.5 Kia, Kib, Kic, K (Excel formulas)
            # -------------------------------------------------
            sumS = sum(crisp_S.values())
            sumP = sum(crisp_P.values())
            minS = min(crisp_S.values())
            minP = min(crisp_P.values())
            maxS = max(crisp_S.values())
            maxP = max(crisp_P.values())

            rows = []
            denom_kic = (tau*maxS + (1.0-tau)*maxP)
            denom_kic = denom_kic if denom_kic != 0 else 1.0

            for alt in alternatives:
                S = crisp_S[alt]
                P = crisp_P[alt]

                Kia = (S + P) / (sumS + sumP) if (sumS + sumP) != 0 else 0.0
                Kib = (S/minS if minS != 0 else 0.0) + (P/minP if minP != 0 else 0.0)
                Kic = ((tau*S) + ((1.0-tau)*P)) / denom_kic

                K = (Kia*Kib*Kic)**(1/3) + ((Kia + Kib + Kic)/3)

                rows.append({
                    "Alternative": alt,
                    "Kia": Kia,
                    "Kib": Kib,
                    "Kic": Kic,
                    "K": K
                })

            dfK = pd.DataFrame(rows)
            # Excel RANK is descending
            dfK["Rank"] = dfK["K"].rank(ascending=False, method="min").astype(int)
            dfK = dfK.sort_values("Rank").reset_index(drop=True)

            st.markdown("#### 3.5 Final CoCoSo indices (Excel: Kia, Kib, Kic, K) & Rank")
            st.dataframe(dfK.style.format(precision=6), use_container_width=True, hide_index=True)

# =========================================================
# ------------------- YOUR WINGS CODE ----------------------
# (kept as-is, only wrapped in a function)
# =========================================================

# Define linguistic terms for IT2TrFS (WINGS)
LINGUISTIC_TERMS = {
    "strength": {
        "VLR": ((0, 0.1, 0.1, 0.1, 1, 1), (0.0, 0.1, 0.1, 0.05, 0.9, 0.9)),
        "LR": ((0.2, 0.3, 0.3, 0.4, 1, 1), (0.25, 0.3, 0.3, 0.35, 0.9, 0.9)),
        "MR": ((0.4, 0.5, 0.5, 0.6, 1, 1), (0.45, 0.5, 0.5, 0.55, 0.9, 0.9)),
        "HR": ((0.6, 0.7, 0.7, 0.8, 1, 1), (0.65, 0.7, 0.7, 0.75, 0.9, 0.9)),
        "VHR": ((0.8, 0.9, 0.9, 1, 1, 1), (0.85, 0.90, 0.90, 0.95, 0.9, 0.9))
    },
    "influence": {
        "ELI": ((0, 0.1, 0.1, 0.2, 1, 1), (0.05, 0.1, 0.1, 0.15, 0.9, 0.9)),
        "VLI": ((0.1, 0.2, 0.2, 0.35, 1, 1), (0.15, 0.2, 0.2, 0.3, 0.9, 0.9)),
        "LI": ((0.2, 0.35, 0.35, 0.5, 1, 1), (0.25, 0.35, 0.35, 0.45, 0.9, 0.9)),
        "MI": ((0.35, 0.5, 0.5, 0.65, 1, 1), (0.40, 0.5, 0.5, 0.6, 0.9, 0.9)),
        "HI": ((0.5, 0.65, 0.65, 0.8, 1, 1), (0.55, 0.65, 0.65, 0.75, 0.9, 0.9)),
        "VHI": ((0.65, 0.80, 0.80, 0.9, 1, 1), (0.7, 0.8, 0.8, 0.85, 0.9, 0.9)),
        "EHI": ((0.8, 0.9, 0.9, 1, 1, 1), (0.85, 0.9, 0.9, 0.95, 0.9, 0.9))
    }
}

FULL_FORMS = {
    "VLR": "Very Low Relevance",
    "LR": "Low Relevance",
    "MR": "Medium Relevance",
    "HR": "High Relevance",
    "VHR": "Very High Relevance",
    "ELI": "Extremely Low Influence",
    "VLI": "Very Low Influence",
    "LI": "Low Influence",
    "MI": "Medium Influence",
    "HI": "High Influence",
    "VHI": "Very High Influence",
    "EHI": "Extremely High Influence"
}

def defuzz_it2(A):
    Au, Al = A
    return (Au[0] + Au[1] + Au[2] + Au[3] + Al[0] + Al[1] + Al[2] + Al[3]) / 8

def identity_it2(n):
    I_mat = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for i in range(n):
        I_mat[i][i] = ((1, 1, 1, 1, 1, 1), (1, 1, 1, 1, 1, 1))
    return I_mat

def compute_total_relation_matrix(normalized_matrix):
    n = len(normalized_matrix)
    I = identity_it2(n)

    # Convert normalized_matrix to 4D array for parameter-wise computation
    Z_4d = np.zeros((2, 2, n, n, 4))
    for i in range(n):
        for j in range(n):
            Au, Al = normalized_matrix[i][j]
            Z_4d[0, 0, i, j, :] = Au[:4]
            Z_4d[0, 1, i, j, :2] = Au[4:]
            Z_4d[1, 0, i, j, :] = Al[:4]
            Z_4d[1, 1, i, j, :2] = Al[4:]

    for i in range(2):
        for j in range(2):
            if j == 0:
                for k in range(4):
                    Z_component = Z_4d[i, j, :, :, k]
                    try:
                        T_component = Z_component @ np.linalg.pinv(np.eye(n) - Z_component)
                    except np.linalg.LinAlgError:
                        T_component = np.zeros((n, n))
                    Z_4d[i, j, :, :, k] = T_component

    T = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for i in range(n):
        for j in range(n):
            T[i][j] = (
                (Z_4d[0, 0, i, j, 0], Z_4d[0, 0, i, j, 1], Z_4d[0, 0, i, j, 2], Z_4d[0, 0, i, j, 3], Z_4d[0, 1, i, j, 0], Z_4d[0, 1, i, j, 1]),
                (Z_4d[1, 0, i, j, 0], Z_4d[1, 0, i, j, 1], Z_4d[1, 0, i, j, 2], Z_4d[1, 0, i, j, 3], Z_4d[1, 1, i, j, 0], Z_4d[1, 1, i, j, 1])
            )
    return T

def calculate_TI_TR(T):
    n = len(T)
    TI = [zero_it2() for _ in range(n)]
    TR = [zero_it2() for _ in range(n)]
    for i in range(n):
        for j in range(n):
            TI[i] = add_it2(TI[i], T[i][j])
            TR[j] = add_it2(TR[j], T[i][j])
    return TI, TR

def wings_method_experts(strengths_list, influence_matrices_list, weights=None):
    n = len(strengths_list[0])
    num_experts = len(strengths_list)
    if weights is None:
        weights = [1.0 / num_experts] * num_experts

    avg_sidrm = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for exp in range(num_experts):
        w = weights[exp]
        for i in range(n):
            str_w = scalar_mul_it2(w, strengths_list[exp][i])
            avg_sidrm[i][i] = add_it2(avg_sidrm[i][i], str_w)
            for j in range(n):
                if i != j:
                    inf_w = scalar_mul_it2(w, influence_matrices_list[exp][i][j])
                    avg_sidrm[i][j] = add_it2(avg_sidrm[i][j], inf_w)

    s1U=s2U=s3U=s4U=s1L=s2L=s3L=s4L=0.0
    for i in range(n):
        for j in range(n):
            Au, Al = avg_sidrm[i][j]
            s1U += Au[0]; s2U += Au[1]; s3U += Au[2]; s4U += Au[3]
            s1L += Al[0]; s2L += Al[1]; s3L += Al[2]; s4L += Al[3]
    s = s1U+s2U+s3U+s4U+s1L+s2L+s3L+s4L

    Z_mat = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for i in range(n):
        for j in range(n):
            Au, Al = avg_sidrm[i][j]
            new_u = (Au[0]/s if s else 0, Au[1]/s if s else 0, Au[2]/s if s else 0, Au[3]/s if s else 0, Au[4], Au[5])
            new_l = (Al[0]/s if s else 0, Al[1]/s if s else 0, Al[2]/s if s else 0, Al[3]/s if s else 0, Al[4], Al[5])
            Z_mat[i][j] = (new_u, new_l)

    T_mat = compute_total_relation_matrix(Z_mat)
    TI, TR = calculate_TI_TR(T_mat)

    engagement = [add_it2(TI[i], TR[i]) for i in range(n)]
    role = [sub_it2(TI[i], TR[i]) for i in range(n)]

    TI_defuzz = np.array([defuzz_it2(TI[i]) for i in range(n)])
    TR_defuzz = np.array([defuzz_it2(TR[i]) for i in range(n)])
    engagement_defuzz = np.array([defuzz_it2(engagement[i]) for i in range(n)])
    role_defuzz = np.array([defuzz_it2(role[i]) for i in range(n)])

    return {
        'average_sidrm': avg_sidrm,
        'scaling_factor': s,
        'normalized_matrix': Z_mat,
        'total_matrix': T_mat,
        'total_impact': TI,
        'total_receptivity': TR,
        'engagement': engagement,
        'role': role,
        'total_impact_defuzz': TI_defuzz,
        'total_receptivity_defuzz': TR_defuzz,
        'engagement_defuzz': engagement_defuzz,
        'role_defuzz': role_defuzz
    }

def format_it2_df(mat, index, columns):
    df = pd.DataFrame(index=index, columns=columns)
    for i in range(len(index)):
        for j in range(len(columns)):
            df.iloc[i, j] = format_it2(mat[i][j])
    return df

def generate_flowchart_for_expert(expert_data, component_names, expert_idx=None):
    n = len(component_names)
    graph = graphviz.Digraph(comment=f'IT2TrFS WINGS Flowchart - Expert {expert_idx+1}' if expert_idx is not None else 'IT2TrFS WINGS Flowchart')
    graph.attr(rankdir='TD', size='8,8')

    for comp_idx, comp_name in enumerate(component_names):
        strength = expert_data['strengths_linguistic'][comp_idx]
        label = f"{comp_name} ({strength})"
        graph.node(comp_name, label=label, shape='box', style='rounded,filled', fillcolor='lightblue', fontsize='12')

    for from_idx, from_comp in enumerate(component_names):
        for to_idx, to_comp in enumerate(component_names):
            if from_idx == to_idx:
                continue
            influence = expert_data['influence_matrix_linguistic'][from_idx][to_idx]
            if influence != "ELI":
                graph.edge(from_comp, to_comp, label=influence)

    return graph

def create_word_report(results, component_names, n_experts=1, expert_weights=None):
    doc = Document()
    title = doc.add_heading('IT2TrFS WINGS Analysis Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    from datetime import datetime
    doc.add_paragraph(f"Report generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph(f"Number of experts: {n_experts}")
    if expert_weights and n_experts > 1:
        weights_text = "Expert weights: " + ", ".join([f"Expert {i+1}: {weight:.2f}" for i, weight in enumerate(expert_weights)])
        doc.add_paragraph(weights_text)

    comp_para = doc.add_paragraph("Components analyzed: ")
    for i, name in enumerate(component_names):
        comp_para.add_run(f"{i+1}. {name}  ")

    doc.add_heading('Impact, Receptivity, Engagement, and Role Results', level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = 'Component'
    hdr[1].text = 'Total Impact (TI)'
    hdr[2].text = 'Total Receptivity (TR)'
    hdr[3].text = 'Engagement (TI+TR)'
    hdr[4].text = 'Role (TI-TR)'

    for i, name in enumerate(component_names):
        row = table.add_row().cells
        row[0].text = name
        row[1].text = f"{results['total_impact_defuzz'][i]:.6f}"
        row[2].text = f"{results['total_receptivity_defuzz'][i]:.6f}"
        row[3].text = f"{results['engagement_defuzz'][i]:.6f}"
        row[4].text = f"{results['role_defuzz'][i]:.6f}"

    return doc

def get_word_download_link(doc):
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    b64 = base64.b64encode(file_stream.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="it2trfs_wings_analysis_report.docx">Download Word Report</a>'
    return href

def wings_app():
    st.title("📊 IT2TrFS WINGS Method Analysis Platform")
    st.write("IT2TrFS-WINGS module (your existing implementation).")

    tab_howto, tab_analysis = st.tabs(["📘 How to Use", "📊 Analysis"])

    with tab_howto:
        st.markdown("Use the sidebar to configure components/experts and run WINGS.")
        with st.expander("Linguistic Terms Reference"):
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Strength/Relevance Terms**")
                strength_df = pd.DataFrame([
                    {"Abbreviation":abbr, "Full Form":FULL_FORMS[abbr], "IT2TrFS":format_it2(it2)}
                    for abbr,it2 in LINGUISTIC_TERMS["strength"].items()
                ])
                st.dataframe(strength_df, hide_index=True, use_container_width=True)
            with col2:
                st.write("**Influence Terms**")
                infl_df = pd.DataFrame([
                    {"Abbreviation":abbr, "Full Form":FULL_FORMS[abbr], "IT2TrFS":format_it2(it2)}
                    for abbr,it2 in LINGUISTIC_TERMS["influence"].items()
                ])
                st.dataframe(infl_df, hide_index=True, use_container_width=True)

    with tab_analysis:
        with st.sidebar:
            st.header("⚙️ WINGS Configuration")
            n_components = st.number_input("Number of Components", min_value=2, max_value=25, value=3, key="w_ncomp")
            n_experts = st.number_input("Number of Experts", min_value=1, max_value=15, value=1, key="w_nexp")

            component_names = []
            for i in range(n_components):
                component_names.append(st.text_input(f"Name of Component {i+1}", value=f"C{i+1}", key=f"w_comp_{i}"))

            expert_weights = None
            if n_experts > 1:
                st.markdown("---")
                st.subheader("Expert Weights (sum=1)")
                weights = []
                for i in range(n_experts):
                    weights.append(st.number_input(f"Weight E{i+1}", min_value=0.0, max_value=1.0, value=round(1/n_experts, 4), step=0.01, key=f"w_w_{i}"))
                if not np.isclose(sum(weights), 1.0):
                    st.error(f"Weights must sum to 1.0 (now: {sum(weights):.4f})")
                    st.stop()
                expert_weights = weights

        if "experts_data" not in st.session_state:
            st.session_state.experts_data = {}

        for e in range(n_experts):
            if e not in st.session_state.experts_data:
                st.session_state.experts_data[e] = {
                    "strengths_linguistic": ["HR"]*n_components,
                    "influence_matrix_linguistic": [["ELI"]*n_components for _ in range(n_components)]
                }

        tabs = st.tabs([f"Expert {i+1}" for i in range(n_experts)]) if n_experts > 1 else [st.container()]

        strengths_list = []
        influence_list = []

        for e in range(n_experts):
            with tabs[e] if n_experts > 1 else tabs[0]:
                strengths = []
                st.write("**Strengths**")
                cols = st.columns(n_components)
                for i in range(n_components):
                    with cols[i]:
                        cur = st.session_state.experts_data[e]["strengths_linguistic"][i]
                        term = st.selectbox(component_names[i], options=list(LINGUISTIC_TERMS["strength"].keys()),
                                            index=list(LINGUISTIC_TERMS["strength"].keys()).index(cur),
                                            key=f"w_str_{e}_{i}")
                        st.session_state.experts_data[e]["strengths_linguistic"][i] = term
                        strengths.append(LINGUISTIC_TERMS["strength"][term])

                st.write("**Influence Matrix** (row influences column)")
                inf_mat = [[None]*n_components for _ in range(n_components)]
                for i in range(n_components):
                    row_cols = st.columns(n_components)
                    for j in range(n_components):
                        with row_cols[j]:
                            if i == j:
                                st.markdown("—")
                                inf_mat[i][j] = zero_it2()
                            else:
                                cur = st.session_state.experts_data[e]["influence_matrix_linguistic"][i][j]
                                term = st.selectbox(f"{component_names[i]}→{component_names[j]}",
                                                    options=list(LINGUISTIC_TERMS["influence"].keys()),
                                                    index=list(LINGUISTIC_TERMS["influence"].keys()).index(cur),
                                                    key=f"w_inf_{e}_{i}_{j}")
                                st.session_state.experts_data[e]["influence_matrix_linguistic"][i][j] = term
                                inf_mat[i][j] = LINGUISTIC_TERMS["influence"][term]

                strengths_list.append(strengths)
                influence_list.append(inf_mat)

        if st.button("🚀 Run IT2TrFS WINGS Analysis", type="primary", use_container_width=True, key="w_run"):
            with st.spinner("Calculating..."):
                results = wings_method_experts(strengths_list, influence_list, expert_weights)

            st.success("Done.")

            t1, t2, t3 = st.tabs(["Matrices", "Results", "Export"])
            with t1:
                st.subheader("Average SIDRM")
                st.dataframe(format_it2_df(results["average_sidrm"], component_names, component_names), use_container_width=True)
                st.subheader("Normalized Z")
                st.dataframe(format_it2_df(results["normalized_matrix"], component_names, component_names), use_container_width=True)
                st.subheader("Total T")
                st.dataframe(format_it2_df(results["total_matrix"], component_names, component_names), use_container_width=True)

            with t2:
                df_res = pd.DataFrame({
                    "Component": component_names,
                    "TI": results["total_impact_defuzz"],
                    "TR": results["total_receptivity_defuzz"],
                    "Engagement": results["engagement_defuzz"],
                    "Role": results["role_defuzz"],
                })
                df_res["Type"] = np.where(df_res["Role"] > 0, "Cause", "Effect")
                st.dataframe(df_res.style.format(precision=6), use_container_width=True, hide_index=True)

            with t3:
                doc = create_word_report(results, component_names, n_experts, expert_weights)
                st.markdown(get_word_download_link(doc), unsafe_allow_html=True)

# =========================================================
# MAIN NAVIGATION (TWO MODULES)
# =========================================================

def main():
    st.set_page_config(page_title="IT2TrFS Toolkit (WINGS + CoCoSo)", layout="wide")
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Choose a Module", ["IT2TrFS-WINGS", "IT2TrFS-CoCoSo"])

    if page == "IT2TrFS-WINGS":
        wings_app()
    else:
        cocoso_app()

if __name__ == "__main__":
    main()
