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
    # NOTE: kept consistent with your WINGS averaging style (scale a,b,c,d only; heights unchanged)
    Au, Al = A
    new_u = (k*Au[0], k*Au[1], k*Au[2], k*Au[3], Au[4], Au[5])
    new_l = (k*Al[0], k*Al[1], k*Al[2], k*Al[3], Al[4], Al[5])
    return (new_u, new_l)

def defuzz_it2(A):
    Au, Al = A
    # simple centroid-like average of (a,b,c,d) for UMF and LMF
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

    # protect against zeros
    eps = 1e-12

    u_a = np.prod([max(it2[0][0], eps)**w for it2, w in zip(it2_list, weights)])
    u_b = np.prod([max(it2[0][1], eps)**w for it2, w in zip(it2_list, weights)])
    u_c = np.prod([max(it2[0][2], eps)**w for it2, w in zip(it2_list, weights)])
    u_d = np.prod([max(it2[0][3], eps)**w for it2, w in zip(it2_list, weights)])

    l_a = np.prod([max(it2[1][0], eps)**w for it2, w in zip(it2_list, weights)])
    l_b = np.prod([max(it2[1][1], eps)**w for it2, w in zip(it2_list, weights)])
    l_c = np.prod([max(it2[1][2], eps)**w for it2, w in zip(it2_list, weights)])
    l_d = np.prod([max(it2[1][3], eps)**w for it2, w in zip(it2_list, weights)])

    uh1 = min([it2[0][4] for it2 in it2_list])
    uh2 = min([it2[0][5] for it2 in it2_list])
    lh1 = min([it2[1][4] for it2 in it2_list])
    lh2 = min([it2[1][5] for it2 in it2_list])

    return ((u_a, u_b, u_c, u_d, uh1, uh2),
            (l_a, l_b, l_c, l_d, lh1, lh2))

def it2_complement(A):
    """
    Simple trapezoid complement (for COST handling if you want IT2-level inversion):
    (a,b,c,d) -> (1-d, 1-c, 1-b, 1-a); heights unchanged.
    """
    Au, Al = A
    cu = (1-Au[3], 1-Au[2], 1-Au[1], 1-Au[0], Au[4], Au[5])
    cl = (1-Al[3], 1-Al[2], 1-Al[1], 1-Al[0], Al[4], Al[5])
    return (cu, cl)


# ============================================================
# 2) WINGS module (your original linguistic sets)
# ============================================================

LINGUISTIC_TERMS_WINGS = {
    "strength": {
        "VLR": ((0, 0.1, 0.1, 0.1, 1, 1), (0.0, 0.1, 0.1, 0.05, 0.9, 0.9)),
        "LR":  ((0.2, 0.3, 0.3, 0.4, 1, 1), (0.25, 0.3, 0.3, 0.35, 0.9, 0.9)),
        "MR":  ((0.4, 0.5, 0.5, 0.6, 1, 1), (0.45, 0.5, 0.5, 0.55, 0.9, 0.9)),
        "HR":  ((0.6, 0.7, 0.7, 0.8, 1, 1), (0.65, 0.7, 0.7, 0.75, 0.9, 0.9)),
        "VHR": ((0.8, 0.9, 0.9, 1,   1, 1), (0.85, 0.90,0.90,0.95, 0.9, 0.9))
    },
    "influence": {
        "ELI": ((0,   0.1, 0.1, 0.2, 1, 1), (0.05,0.1, 0.1, 0.15,0.9,0.9)),
        "VLI": ((0.1, 0.2, 0.2, 0.35,1, 1), (0.15,0.2, 0.2, 0.3, 0.9,0.9)),
        "LI":  ((0.2, 0.35,0.35,0.5, 1, 1), (0.25,0.35,0.35,0.45,0.9,0.9)),
        "MI":  ((0.35,0.5, 0.5, 0.65,1, 1), (0.40,0.5, 0.5, 0.6, 0.9,0.9)),
        "HI":  ((0.5, 0.65,0.65,0.8, 1, 1), (0.55,0.65,0.65,0.75,0.9,0.9)),
        "VHI": ((0.65,0.80,0.80,0.9, 1, 1), (0.7, 0.8, 0.8, 0.85,0.9,0.9)),
        "EHI": ((0.8, 0.9, 0.9, 1,   1, 1), (0.85,0.9, 0.9, 0.95,0.9,0.9))
    }
}

FULL_FORMS_WINGS = {
    "VLR": "Very Low Relevance",
    "LR":  "Low Relevance",
    "MR":  "Medium Relevance",
    "HR":  "High Relevance",
    "VHR": "Very High Relevance",
    "ELI": "Extremely Low Influence",
    "VLI": "Very Low Influence",
    "LI":  "Low Influence",
    "MI":  "Medium Influence",
    "HI":  "High Influence",
    "VHI": "Very High Influence",
    "EHI": "Extremely High Influence"
}

def identity_it2(n):
    I_mat = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for i in range(n):
        I_mat[i][i] = ((1,1,1,1,1,1), (1,1,1,1,1,1))
    return I_mat

def compute_total_relation_matrix(normalized_matrix):
    n = len(normalized_matrix)

    Z_4d = np.zeros((2, 2, n, n, 4))
    for i in range(n):
        for j in range(n):
            Au, Al = normalized_matrix[i][j]
            Z_4d[0, 0, i, j, :] = Au[:4]
            Z_4d[0, 1, i, j, :2] = Au[4:]
            Z_4d[1, 0, i, j, :] = Al[:4]
            Z_4d[1, 1, i, j, :2] = Al[4:]

    for i in range(2):
        for k in range(4):
            Z_component = Z_4d[i, 0, :, :, k]
            try:
                T_component = Z_component @ np.linalg.pinv(np.eye(n) - Z_component)
            except np.linalg.LinAlgError:
                T_component = np.zeros((n, n))
            Z_4d[i, 0, :, :, k] = T_component

    T = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for i in range(n):
        for j in range(n):
            T[i][j] = (
                (Z_4d[0,0,i,j,0], Z_4d[0,0,i,j,1], Z_4d[0,0,i,j,2], Z_4d[0,0,i,j,3],
                 Z_4d[0,1,i,j,0], Z_4d[0,1,i,j,1]),
                (Z_4d[1,0,i,j,0], Z_4d[1,0,i,j,1], Z_4d[1,0,i,j,2], Z_4d[1,0,i,j,3],
                 Z_4d[1,1,i,j,0], Z_4d[1,1,i,j,1])
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
            avg_sidrm[i][i] = add_it2(avg_sidrm[i][i], scalar_mul_it2(w, strengths_list[exp][i]))
            for j in range(n):
                if i != j:
                    avg_sidrm[i][j] = add_it2(avg_sidrm[i][j], scalar_mul_it2(w, influence_matrices_list[exp][i][j]))

    s = 0.0
    for i in range(n):
        for j in range(n):
            Au, Al = avg_sidrm[i][j]
            s += (Au[0]+Au[1]+Au[2]+Au[3]+Al[0]+Al[1]+Al[2]+Al[3])

    Z_mat = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for i in range(n):
        for j in range(n):
            Au, Al = avg_sidrm[i][j]
            Z_mat[i][j] = (
                (Au[0]/s if s else 0, Au[1]/s if s else 0, Au[2]/s if s else 0, Au[3]/s if s else 0, Au[4], Au[5]),
                (Al[0]/s if s else 0, Al[1]/s if s else 0, Al[2]/s if s else 0, Al[3]/s if s else 0, Al[4], Al[5])
            )

    T_mat = compute_total_relation_matrix(Z_mat)
    TI, TR = calculate_TI_TR(T_mat)
    engagement = [add_it2(TI[i], TR[i]) for i in range(n)]
    role = [sub_it2(TI[i], TR[i]) for i in range(n)]

    return {
        'average_sidrm': avg_sidrm,
        'normalized_matrix': Z_mat,
        'total_matrix': T_mat,
        'total_impact': TI,
        'total_receptivity': TR,
        'engagement': engagement,
        'role': role,
        'total_impact_defuzz': np.array([defuzz_it2(x) for x in TI]),
        'total_receptivity_defuzz': np.array([defuzz_it2(x) for x in TR]),
        'engagement_defuzz': np.array([defuzz_it2(x) for x in engagement]),
        'role_defuzz': np.array([defuzz_it2(x) for x in role]),
    }

def format_it2_df(mat, index, columns):
    df = pd.DataFrame(index=index, columns=columns)
    for i in range(len(index)):
        for j in range(len(columns)):
            df.iloc[i, j] = format_it2(mat[i][j])
    return df

def generate_flowchart_for_expert(expert_data, component_names, expert_idx=None):
    graph = graphviz.Digraph(comment=f'IT2TrFS WINGS Flowchart - Expert {expert_idx+1}' if expert_idx is not None else 'IT2TrFS WINGS Flowchart')
    graph.attr(rankdir='TD', size='8,8')

    for comp_idx, comp_name in enumerate(component_names):
        strength = expert_data['strengths_linguistic'][comp_idx]
        graph.node(comp_name, label=f"{comp_name} ({strength})", shape='box',
                   style='rounded,filled', fillcolor='lightblue', fontsize='12')

    for i, f in enumerate(component_names):
        for j, t in enumerate(component_names):
            if i == j:
                continue
            inf = expert_data['influence_matrix_linguistic'][i][j]
            if inf != "ELI":
                graph.edge(f, t, label=inf)
    return graph

def create_word_report(results, component_names, n_experts=1, expert_weights=None):
    doc = Document()
    title = doc.add_heading('IT2TrFS WINGS Analysis Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Number of experts: {n_experts}")
    if expert_weights and n_experts > 1:
        doc.add_paragraph("Expert weights: " + ", ".join([f"E{i+1}={w:.2f}" for i, w in enumerate(expert_weights)]))
    return doc

def get_word_download_link(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="it2trfs_wings_report.docx">Download Word Report</a>'


# ============================================================
# 3) IT2TrFS-CoCoSo (MANUAL INPUT like your IVIFN app)
# ============================================================

def default_cocoso_scale():
    """
    Default scale (editable). Replace these numbers with your paper’s scale if needed.
    Each row is an IT2TrFS number: UMF(a,b,c,d;uh1,uh2) and LMF(e,f,g,h;lh1,lh2)
    """
    rows = [
        ("Very Poor",  "VP", 0.00,0.10,0.10,0.20,1.00,1.00, 0.00,0.05,0.05,0.15,0.90,0.90),
        ("Poor",       "P",  0.10,0.20,0.20,0.35,1.00,1.00, 0.15,0.20,0.20,0.30,0.90,0.90),
        ("Medium Poor","MP", 0.20,0.35,0.35,0.50,1.00,1.00, 0.25,0.35,0.35,0.45,0.90,0.90),
        ("Fair",       "F",  0.35,0.50,0.50,0.65,1.00,1.00, 0.40,0.50,0.50,0.60,0.90,0.90),
        ("Medium Good","MG", 0.50,0.65,0.65,0.80,1.00,1.00, 0.55,0.65,0.65,0.75,0.90,0.90),
        ("Good",       "G",  0.65,0.80,0.80,0.90,1.00,1.00, 0.70,0.80,0.80,0.85,0.90,0.90),
        ("Very Good",  "VG", 0.80,0.90,0.90,1.00,1.00,1.00, 0.85,0.90,0.90,0.95,0.90,0.90),
    ]
    return pd.DataFrame(rows, columns=["Linguistic Attribute","Code","a","b","c","d","uh1","uh2","e","f","g","h","lh1","lh2"])

def scale_df_to_map(scale_df: pd.DataFrame):
    mapping = {}
    for _, r in scale_df.iterrows():
        code = str(r["Code"]).strip()
        if not code:
            continue
        it2 = (
            (float(r["a"]), float(r["b"]), float(r["c"]), float(r["d"]), float(r["uh1"]), float(r["uh2"])),
            (float(r["e"]), float(r["f"]), float(r["g"]), float(r["h"]), float(r["lh1"]), float(r["lh2"])),
        )
        mapping[code] = it2
    return mapping

def module_cocoso_manual():
    st.header("📊 IT2TrFS–CoCoSo (Manual input like IVIFN app)")

    st.markdown("""
    **Workflow**
    1) Define alternatives + criteria  
    2) Set criterion type (Benefit/Cost) + weights (sum=1)  
    3) Choose number of experts + expert weights (sum=1)  
    4) Fill expert evaluation matrices using linguistic codes (dropdown)  
    5) Run CoCoSo → see aggregated IT2TrFS matrix + crisp normalization + final K & rank
    """)

    # ---- Linguistic scale (editable) ----
    st.subheader("Step 0: Linguistic scale (editable)")
    if "cocoso_scale_df" not in st.session_state:
        st.session_state.cocoso_scale_df = default_cocoso_scale()

    st.session_state.cocoso_scale_df = st.data_editor(
        st.session_state.cocoso_scale_df,
        use_container_width=True,
        hide_index=True,
        key="cocoso_scale_editor"
    )

    scale_map = scale_df_to_map(st.session_state.cocoso_scale_df)
    code_options = list(scale_map.keys())
    if len(code_options) < 2:
        st.error("Please provide at least 2 linguistic codes in the scale table.")
        st.stop()

    # ---- Step 1: define alternatives + criteria ----
    st.subheader("Step 1: Define Alternatives and Criteria")
    c1, c2 = st.columns(2)
    alts_in = c1.text_input("Alternatives (comma-separated)", "T1, T2, T3", key="cocoso_alts_in")
    crits_in = c2.text_input("Criteria (comma-separated)", "C1, C2, C3", key="cocoso_crits_in")

    alternatives = [a.strip() for a in alts_in.split(",") if a.strip()]
    criteria = [c.strip() for c in crits_in.split(",") if c.strip()]

    if len(alternatives) < 2 or len(criteria) < 1:
        st.warning("Please enter at least 2 alternatives and at least 1 criterion.")
        return

    # ---- Step 2: criterion weights & types ----
    st.subheader("Step 2: Criterion type & weights")
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
            "Weight": st.column_config.NumberColumn("Weight", min_value=0.0, max_value=1.0, format="%.6f")
        },
        key="cocoso_crit_editor"
    )

    crit_types = st.session_state.cocoso_crit_df["Type"].tolist()
    crit_w = st.session_state.cocoso_crit_df["Weight"].astype(float).tolist()
    if not np.isclose(sum(crit_w), 1.0):
        st.error(f"Criterion weights must sum to 1. Current sum: {sum(crit_w):.6f}")
        st.stop()

    # ---- Step 3: experts & weights ----
    st.subheader("Step 3: Experts")
    n_exp = st.number_input("Number of experts", min_value=1, max_value=20, value=2, step=1, key="cocoso_nexp")

    st.markdown("**Expert weights** (must sum to 1.0)")
    exp_w = []
    cols = st.columns(n_exp) if n_exp > 1 else [st.container()]
    if n_exp > 1:
        for i in range(n_exp):
            with cols[i]:
                exp_w.append(st.number_input(f"E{i+1}", 0.0, 1.0, value=1.0/n_exp, step=0.05, format="%.2f", key=f"cocoso_expw_{i}"))
        if not np.isclose(sum(exp_w), 1.0):
            st.error(f"Expert weights must sum to 1. Current sum: {sum(exp_w):.2f}")
            st.stop()
    else:
        exp_w = [1.0]

    # ---- Step 4: expert matrices (dropdown) ----
    st.subheader("Step 4: Expert evaluation matrices (linguistic codes)")
    if "cocoso_expert_mats" not in st.session_state:
        st.session_state.cocoso_expert_mats = {}

    # init/reset if dimension changed
    need_reset = (
        len(st.session_state.cocoso_expert_mats) != n_exp or
        (n_exp > 0 and (
            list(st.session_state.cocoso_expert_mats.get(0, pd.DataFrame())).count != 0 and
            (set(st.session_state.cocoso_expert_mats[0].index) != set(alternatives) or
             set(st.session_state.cocoso_expert_mats[0].columns) != set(criteria))
        ))
    )
    if need_reset or len(st.session_state.cocoso_expert_mats) != n_exp:
        st.session_state.cocoso_expert_mats = {i: pd.DataFrame(code_options[0], index=alternatives, columns=criteria) for i in range(n_exp)}

    tabs = st.tabs([f"Expert {i+1}" for i in range(n_exp)])
    for i, tab in enumerate(tabs):
        with tab:
            st.caption(f"Fill codes using dropdowns (options come from your scale table).")
            st.session_state.cocoso_expert_mats[i] = st.data_editor(
                st.session_state.cocoso_expert_mats[i],
                use_container_width=True,
                column_config={c: st.column_config.SelectboxColumn(c, options=code_options) for c in criteria},
                key=f"cocoso_mat_{i}"
            )

    # ---- Step 5: Run CoCoSo ----
    st.subheader("Step 5: Run IT2TrFS–CoCoSo")
    if st.button("✅ Calculate CoCoSo Ranking", type="primary", use_container_width=True, key="cocoso_run"):
        # 5.1 aggregated IT2 decision matrix
        agg_it2 = pd.DataFrame(index=alternatives, columns=criteria, dtype=object)
        for a in alternatives:
            for c in criteria:
                it2s = []
                for ei in range(n_exp):
                    code = str(st.session_state.cocoso_expert_mats[ei].loc[a, c]).strip()
                    if code not in scale_map:
                        st.error(f"Unknown code '{code}' at (Alt={a}, Crit={c}) for Expert {ei+1}.")
                        st.stop()
                    it2s.append(scale_map[code])
                agg_it2.loc[a, c] = it2_weighted_avg(it2s, exp_w)

        st.markdown("### 5.1 Aggregated IT2TrFS decision matrix")
        show_agg = agg_it2.applymap(lambda x: format_it2(x) if x is not None else "N/A")
        st.dataframe(show_agg, use_container_width=True)

        # 5.2 crisp matrix by defuzzification (used for normalization + CoCoSo math)
        crisp = agg_it2.applymap(lambda x: defuzz_it2(x) if x is not None else np.nan)

        st.markdown("### 5.2 Crisp matrix (defuzzified)")
        st.dataframe(crisp.style.format(precision=6), use_container_width=True)

        # 5.3 normalization (standard CoCoSo crisp normalization)
        norm = crisp.copy()
        for j, c in enumerate(criteria):
            col = crisp[c].astype(float)
            if crit_types[j] == "Benefit":
                mx = np.nanmax(col.values)
                norm[c] = col / mx if mx != 0 else 0.0
            else:
                mn = np.nanmin(col.values)
                norm[c] = mn / col if np.all(col.values != 0) else 0.0

        st.markdown("### 5.3 Normalized crisp matrix")
        st.dataframe(norm.style.format(precision=6), use_container_width=True)

        # 5.4 S and P
        w = np.array(crit_w, dtype=float)
        S = norm.values @ w
        P = np.prod(np.power(np.maximum(norm.values, 1e-12), w), axis=1)

        res = pd.DataFrame({
            "Alternative": alternatives,
            "S": S,
            "P": P
        })

        st.markdown("### 5.4 CoCoSo S and P")
        st.dataframe(res.style.format(precision=6), use_container_width=True, hide_index=True)

        # 5.5 Kia/Kib/Kic/K (same structure as your IVIFN CoCoSo)
        s = res["S"].values
        p = res["P"].values

        Kia = (s + p) / np.sum(s + p) if np.sum(s + p) != 0 else np.zeros_like(s)
        Kib = (s / np.min(s) if np.min(s) != 0 else 0) + (p / np.min(p) if np.min(p) != 0 else 0)

        tau = 0.5
        denom = tau*np.max(s) + (1-tau)*np.max(p)
        Kic = (tau*s + (1-tau)*p) / denom if denom != 0 else np.zeros_like(s)

        K = np.power(Kia*Kib*Kic, 1/3) + (Kia + Kib + Kic)/3

        final = pd.DataFrame({
            "Alternative": alternatives,
            "Kia": Kia,
            "Kib": Kib,
            "Kic": Kic,
            "K": K
        })
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
# 4) IT2TrFS–WINGS module UI (kept simple)
# ============================================================

def module_wings():
    st.header("📊 IT2TrFS–WINGS")

    with st.sidebar:
        st.subheader("⚙️ WINGS Configuration")
        n_components = st.number_input("Number of Components", min_value=2, max_value=25, value=3, step=1, key="w_ncomp")
        n_experts = st.number_input("Number of Experts", min_value=1, max_value=15, value=1, step=1, key="w_nexp")

        component_names = [st.text_input(f"Component {i+1}", value=f"C{i+1}", key=f"w_comp_{i}") for i in range(n_components)]

        expert_weights = None
        if n_experts > 1:
            st.markdown("---")
            st.caption("Expert weights (sum=1)")
            ws = []
            cols = st.columns(n_experts)
            for i in range(n_experts):
                with cols[i]:
                    ws.append(st.number_input(f"E{i+1}", 0.0, 1.0, value=1.0/n_experts, step=0.05, key=f"w_expw_{i}"))
            if not np.isclose(sum(ws), 1.0):
                st.error("Expert weights must sum to 1.")
                st.stop()
            expert_weights = ws

    # session storage
    if "w_experts" not in st.session_state:
        st.session_state.w_experts = {}

    for ei in range(n_experts):
        if ei not in st.session_state.w_experts:
            st.session_state.w_experts[ei] = {
                "strengths": ["HR"] * n_components,
                "infl": [["ELI"] * n_components for _ in range(n_components)]
            }
        else:
            if len(st.session_state.w_experts[ei]["strengths"]) != n_components:
                st.session_state.w_experts[ei]["strengths"] = ["HR"] * n_components
            if len(st.session_state.w_experts[ei]["infl"]) != n_components:
                st.session_state.w_experts[ei]["infl"] = [["ELI"] * n_components for _ in range(n_components)]

    tabs = st.tabs([f"Expert {i+1}" for i in range(n_experts)]) if n_experts > 1 else [st.container()]

    strengths_list = []
    infl_list = []

    for ei in range(n_experts):
        with tabs[ei] if n_experts > 1 else tabs[0]:
            st.markdown("**Strengths**")
            cols = st.columns(n_components)
            strengths = []
            for i in range(n_components):
                with cols[i]:
                    cur = st.session_state.w_experts[ei]["strengths"][i]
                    pick = st.selectbox(component_names[i], list(LINGUISTIC_TERMS_WINGS["strength"].keys()),
                                        index=list(LINGUISTIC_TERMS_WINGS["strength"].keys()).index(cur),
                                        key=f"w_strength_{ei}_{i}")
                    st.session_state.w_experts[ei]["strengths"][i] = pick
                    strengths.append(LINGUISTIC_TERMS_WINGS["strength"][pick])

            st.markdown("**Influence matrix (row → column)**")
            infl = [[zero_it2() for _ in range(n_components)] for _ in range(n_components)]
            for i in range(n_components):
                row_cols = st.columns(n_components)
                for j in range(n_components):
                    with row_cols[j]:
                        if i == j:
                            st.markdown("—")
                        else:
                            cur = st.session_state.w_experts[ei]["infl"][i][j]
                            pick = st.selectbox("",
                                list(LINGUISTIC_TERMS_WINGS["influence"].keys()),
                                index=list(LINGUISTIC_TERMS_WINGS["influence"].keys()).index(cur),
                                key=f"w_infl_{ei}_{i}_{j}",
                                label_visibility="collapsed"
                            )
                            st.session_state.w_experts[ei]["infl"][i][j] = pick
                            infl[i][j] = LINGUISTIC_TERMS_WINGS["influence"][pick]

        strengths_list.append(strengths)
        infl_list.append(infl)

    if st.button("🚀 Run IT2TrFS–WINGS", type="primary", use_container_width=True, key="w_run"):
        res = wings_method_experts(strengths_list, infl_list, expert_weights)
        st.success("Done!")

        t1, t2, t3 = st.tabs(["Flowchart", "Matrices", "Results"])
        with t1:
            if n_experts > 1:
                for ei in range(n_experts):
                    ex_data = {
                        "strengths_linguistic": st.session_state.w_experts[ei]["strengths"],
                        "influence_matrix_linguistic": st.session_state.w_experts[ei]["infl"],
                    }
                    st.graphviz_chart(generate_flowchart_for_expert(ex_data, component_names, ei), use_container_width=True)
            else:
                ex_data = {
                    "strengths_linguistic": st.session_state.w_experts[0]["strengths"],
                    "influence_matrix_linguistic": st.session_state.w_experts[0]["infl"],
                }
                st.graphviz_chart(generate_flowchart_for_expert(ex_data, component_names), use_container_width=True)

        with t2:
            st.subheader("Average SIDRM")
            st.dataframe(format_it2_df(res["average_sidrm"], component_names, component_names), use_container_width=True)
            st.subheader("Normalized Z")
            st.dataframe(format_it2_df(res["normalized_matrix"], component_names, component_names), use_container_width=True)
            st.subheader("Total T")
            st.dataframe(format_it2_df(res["total_matrix"], component_names, component_names), use_container_width=True)

        with t3:
            out = pd.DataFrame({
                "Component": component_names,
                "TI": res["total_impact_defuzz"],
                "TR": res["total_receptivity_defuzz"],
                "Engagement": res["engagement_defuzz"],
                "Role": res["role_defuzz"],
                "Type": ["Cause" if x > 0 else "Effect" for x in res["role_defuzz"]],
            }).sort_values("Engagement", ascending=False)
            st.dataframe(out.style.format(precision=6), use_container_width=True, hide_index=True)

            doc = create_word_report(res, component_names, n_experts, expert_weights)
            st.markdown(get_word_download_link(doc), unsafe_allow_html=True)


# ============================================================
# 5) MAIN APP (sidebar 2 options)
# ============================================================

def main():
    st.set_page_config(page_title="IT2TrFS Toolkit (WINGS + CoCoSo)", layout="wide", page_icon="📊")
    st.title("🧰 IT2TrFS Toolkit")

    st.sidebar.header("Navigation")
    page = st.sidebar.radio("Choose a Model", ["IT2TrFS–WINGS", "IT2TrFS–CoCoSo"], index=0)

    if page == "IT2TrFS–WINGS":
        module_wings()
    else:
        module_cocoso_manual()

if __name__ == "__main__":
    main()
