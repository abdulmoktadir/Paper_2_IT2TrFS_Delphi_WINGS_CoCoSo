import streamlit as st
import numpy as np
import pandas as pd
import graphviz
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import base64
import matplotlib.pyplot as plt

# Define linguistic terms for IT2TrFS
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

# Full forms
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

def format_it2(it2):
    u, l = it2
    return f"(({u[0]:.6f},{u[1]:.6f},{u[2]:.6f},{u[3]:.6f};{u[4]:.1f},{u[5]:.1f}), ({l[0]:.6f},{l[1]:.6f},{l[2]:.6f},{l[3]:.6f};{l[4]:.1f},{l[5]:.1f}))"

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

def zero_it2():
    return ((0,0,0,0,1,1), (0,0,0,0,0.9,0.9))

def matrix_mul_it2(A_mat, B_mat):
    n = len(A_mat)
    C_mat = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for i in range(n):
        for j in range(n):
            for k in range(n):
                prod = mul_it2(A_mat[i][k], B_mat[k][j])
                C_mat[i][j] = add_it2(C_mat[i][j], prod)
    return C_mat

def defuzz_it2(A):
    Au, Al = A
    return (Au[0] + Au[1] + Au[2] + Au[3] + Al[0] + Al[1] + Al[2] + Al[3]) / 8

def identity_it2(n):
    I_mat = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for i in range(n):
        I_mat[i][i] = ((1, 1, 1, 1, 1, 1), (1, 1, 1, 1, 1, 1))
    return I_mat

def compute_total_relation_matrix(normalized_matrix):
    """
    Compute the total relation matrix T from normalized IT2TrFS matrix
    """
    n = len(normalized_matrix)
    T = [[zero_it2() for _ in range(n)] for _ in range(n)]
    I = identity_it2(n)
    
    # Convert normalized_matrix to 4D array for parameter-wise computation
    Z_4d = np.zeros((2, 2, n, n, 4))
    for i in range(n):
        for j in range(n):
            Au, Al = normalized_matrix[i][j]
            Z_4d[0, 0, i, j, :] = Au[:4]  # UMF parameters
            Z_4d[0, 1, i, j, :2] = Au[4:]  # UMF heights
            Z_4d[1, 0, i, j, :] = Al[:4]  # LMF parameters
            Z_4d[1, 1, i, j, :2] = Al[4:]  # LMF heights
    
    # Compute T for each parameter
    for i in range(2):  # UMF/LMF
        for j in range(2):  # Parameters/Heights
            if j == 0:  # Parameters
                for k in range(4):  # For each parameter
                    Z_component = Z_4d[i, j, :, :, k]
                    try:
                        T_component = Z_component @ np.linalg.pinv(np.eye(n) - Z_component)
                    except np.linalg.LinAlgError:
                        T_component = np.zeros((n, n))
                    Z_4d[i, j, :, :, k] = T_component
            else:  # Heights
                pass  # Heights remain unchanged
    
    # Reconstruct IT2TrFS matrix
    for i in range(n):
        for j in range(n):
            T[i][j] = (
                (Z_4d[0, 0, i, j, 0], Z_4d[0, 0, i, j, 1], Z_4d[0, 0, i, j, 2], Z_4d[0, 0, i, j, 3], Z_4d[0, 1, i, j, 0], Z_4d[0, 1, i, j, 1]),
                (Z_4d[1, 0, i, j, 0], Z_4d[1, 0, i, j, 1], Z_4d[1, 0, i, j, 2], Z_4d[1, 0, i, j, 3], Z_4d[1, 1, i, j, 0], Z_4d[1, 1, i, j, 1])
            )
    
    return T

def calculate_TI_TR(T):
    """
    Calculate total impact (TI) and total receptivity (TR) from total relation matrix
    """
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
    
    # Computing weighted average SIDRM from expert inputs
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
    
    # Compute sums for normalization
    s1U = 0.0; s2U = 0.0; s3U = 0.0; s4U = 0.0
    s1L = 0.0; s2L = 0.0; s3L = 0.0; s4L = 0.0
    for i in range(n):
        for j in range(n):
            Au, Al = avg_sidrm[i][j]
            s1U += Au[0]; s2U += Au[1]; s3U += Au[2]; s4U += Au[3]
            s1L += Al[0]; s2L += Al[1]; s3L += Al[2]; s4L += Al[3]
    s = s1U + s2U + s3U + s4U + s1L + s2L + s3L + s4L
    
    # Normalize SIDRM
    Z_mat = [[zero_it2() for _ in range(n)] for _ in range(n)]
    for i in range(n):
        for j in range(n):
            Au, Al = avg_sidrm[i][j]
            new_u = (Au[0]/s if s != 0 else 0, Au[1]/s if s != 0 else 0, Au[2]/s if s != 0 else 0, Au[3]/s if s != 0 else 0, Au[4], Au[5])
            new_l = (Al[0]/s if s != 0 else 0, Al[1]/s if s != 0 else 0, Al[2]/s if s != 0 else 0, Al[3]/s if s != 0 else 0, Al[4], Al[5])
            Z_mat[i][j] = (new_u, new_l)
    
    # Compute total relation matrix
    T_mat = compute_total_relation_matrix(Z_mat)
    
    # Calculate TI and TR
    TI, TR = calculate_TI_TR(T_mat)
    
    # Compute total engagement and role
    engagement = [zero_it2() for _ in range(n)]
    role = [zero_it2() for _ in range(n)]
    for i in range(n):
        engagement[i] = add_it2(TI[i], TR[i])
        role[i] = sub_it2(TI[i], TR[i])
    
    # Defuzzify only the final values
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
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Component'
    hdr_cells[1].text = 'Total Impact (TI)'
    hdr_cells[2].text = 'Total Receptivity (TR)'
    hdr_cells[3].text = 'Engagement (TI+TR)'
    hdr_cells[4].text = 'Role (TI-TR)'
    
    for i, name in enumerate(component_names):
        row_cells = table.add_row().cells
        row_cells[0].text = name
        row_cells[1].text = f"{results['total_impact_defuzz'][i]:.6f}"
        row_cells[2].text = f"{results['total_receptivity_defuzz'][i]:.6f}"
        row_cells[3].text = f"{results['engagement_defuzz'][i]:.6f}"
        row_cells[4].text = f"{results['role_defuzz'][i]:.6f}"
    
    doc.add_heading('Component Classification', level=1)
    
    class_table = doc.add_table(rows=1, cols=3)
    class_table.style = 'Table Grid'
    hdr_cells = class_table.rows[0].cells
    hdr_cells[0].text = 'Component'
    hdr_cells[1].text = 'Type'
    hdr_cells[2].text = 'Role (TI-TR)'
    
    for i, name in enumerate(component_names):
        status = "Cause" if results['role_defuzz'][i] > 0 else "Effect"
        row_cells = class_table.add_row().cells
        row_cells[0].text = name
        row_cells[1].text = status
        row_cells[2].text = f"{results['role_defuzz'][i]:.6f}"
    
    doc.add_heading('Matrices', level=1)
    
    doc.add_heading('Average SIDRM', level=2)
    df_avg = format_it2_df(results['average_sidrm'], component_names, component_names)
    add_dataframe_to_doc(doc, df_avg)
    
    doc.add_heading('Normalized Matrix Z', level=2)
    df_z = format_it2_df(results['normalized_matrix'], component_names, component_names)
    add_dataframe_to_doc(doc, df_z)
    
    doc.add_heading('Total Matrix T (IT2TrFS)', level=2)
    df_t = format_it2_df(results['total_matrix'], component_names, component_names)
    add_dataframe_to_doc(doc, df_t)
    
    doc.add_paragraph()
    doc.add_heading('Interpretation of Results', level=1)
    interpretation = doc.add_paragraph()
    interpretation.add_run("Total Impact (TI) ").bold = True
    interpretation.add_run("represents the outgoing influence of a component.")
    
    interpretation = doc.add_paragraph()
    interpretation.add_run("Total Receptivity (TR) ").bold = True
    interpretation.add_run("represents the incoming influence on a component.")
    
    interpretation = doc.add_paragraph()
    interpretation.add_run("Engagement (TI+TR) ").bold = True
    interpretation.add_run("indicates the overall involvement of a component in the system.")
    
    interpretation = doc.add_paragraph()
    interpretation.add_run("Role (TI-TR) ").bold = True
    interpretation.add_run("indicates whether a component is a cause or effect: ")
    interpretation.add_run("Positive values ").bold = True
    interpretation.add_run("indicate a component is a ")
    interpretation.add_run("Cause ").bold = True
    interpretation.add_run("(influences others more than it is influenced). ")
    interpretation.add_run("Negative values ").bold = True
    interpretation.add_run("indicate a component is an ")
    interpretation.add_run("Effect ").bold = True
    interpretation.add_run("(is influenced more than it influences).")
    
    return doc

def add_dataframe_to_doc(doc, df, precision=6):
    table = doc.add_table(rows=1, cols=len(df.columns)+1)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = ''
    
    for i, col in enumerate(df.columns):
        hdr_cells[i+1].text = str(col)
    
    for i, index in enumerate(df.index):
        row_cells = table.add_row().cells
        row_cells[0].text = str(index)
        
        for j, col in enumerate(df.columns):
            value = df.iloc[i, j]
            row_cells[j+1].text = str(value)
    
    doc.add_paragraph()

def get_word_download_link(doc):
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    b64 = base64.b64encode(file_stream.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="it2trfs_wings_analysis_report.docx">Download Word Report</a>'
    return href

def main():
    st.set_page_config(page_title="IT2TrFS WINGS Method Analysis", layout="wide", page_icon="📊")
    st.title("📊 IT2TrFS WINGS Method Analysis Platform")
    st.write("""
    This tool implements the Interval Type-2 Trapezoidal Fuzzy Sets Weighted Influence Non-linear Gauge System (IT2TrFS WINGS) method 
    for analyzing systems with interrelated components under uncertainty, incorporating input from multiple experts.
    """)
    
    tab_howto, tab_analysis = st.tabs(["📘 How to Use", "📊 Analysis"])
    
    with tab_howto:
        st.header("How to Use the IT2TrFS WINGS Analysis Platform")
        
        st.markdown("""
        ### Overview
        The IT2TrFS WINGS method is a decision-making tool that helps analyze complex systems with interrelated components, 
        handling uncertainty using Interval Type-2 Trapezoidal Fuzzy Sets (IT2TrFSs). This platform allows you to perform IT2TrFS WINGS analysis using linguistic terms mapped to IT2TrFS intervals.
        
        ### Step-by-Step Guide
        
        1. **Configuration (Sidebar)**
           - Specify the number of components in your system
           - Specify the number of experts
           - Name each component for easy reference
        
        2. **Input Data**
           - **Component Strengths**: For each component, specify its internal strength/relevance using linguistic terms
           - **Influence Matrix**: Define how each component influences others using linguistic terms
           - Use the expandable Linguistic Terms Reference if needed
        
        3. **Run Analysis**
           - Click the "Run IT2TrFS WINGS Analysis" button to process your inputs
           - The system will calculate total impact, receptivity, engagement, and role using IT2TrFS and defuzzification
        
        4. **Interpret Results**
           - **Flowchart**: Visual representation of components and their interactions
           - **Matrices**: View the IT2TrFS and crisp matrices
           - **Results**: See impact, receptivity, engagement, and role values for each component
           - **Classification**: Components are classified as Causes or Effects
           - **Visualization**: Graphical representations of the analysis
        
        ### Input Model
        - Use predefined linguistic terms for strength and influence assessments
        - Terms are mapped to IT2TrFS intervals for uncertainty handling
        - Multiple experts can provide assessments
        - Expert weights can be assigned for weighted averages
        
        ### Understanding the Results
        - **Total Impact (TI)**: Represents the outgoing influence of a component
        - **Total Receptivity (TR)**: Represents the incoming influence on a component
        - **Engagement (TI+TR)**: Indicates the overall involvement of a component
        - **Role (TI-TR)**: 
          - Positive values indicate a component is a **Cause** (influences others more than it's influenced)
          - Negative values indicate a component is an **Effect** (is influenced more than it influences others)
        
        ### Tips for Effective Use
        - Start with a small number of components to understand the method
        - Use descriptive names for components to make interpretation easier
        - Ensure all experts understand the linguistic term definitions
        - Review the flowchart to verify your inputs match your mental model of the system
        """)
        
        with st.expander("Linguistic Terms Reference"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Strength/Relevance Terms**")
                strength_data = []
                for abbr, it2 in LINGUISTIC_TERMS["strength"].items():
                    strength_data.append({
                        "Abbreviation": abbr,
                        "Full Form": FULL_FORMS[abbr],
                        "IT2TrFS Interval": format_it2(it2)
                    })
                strength_df = pd.DataFrame(strength_data)
                st.dataframe(strength_df, hide_index=True, use_container_width=True)
            
            with col2:
                st.write("**Influence Terms**")
                influence_data = []
                for abbr, it2 in LINGUISTIC_TERMS["influence"].items():
                    influence_data.append({
                        "Abbreviation": abbr,
                        "Full Form": FULL_FORMS[abbr],
                        "IT2TrFS Interval": format_it2(it2)
                    })
                influence_df = pd.DataFrame(influence_data)
                st.dataframe(influence_df, hide_index=True, use_container_width=True)
    
    with tab_analysis:
        with st.sidebar:
            st.header("⚙️ Configuration")
            n_components = st.number_input("Number of Components", min_value=2, max_value=25, value=3, help="How many components are in your system?")
            n_experts = st.number_input("Number of Experts", min_value=1, max_value=15, value=1, help="How many experts will provide assessments?")
            
            component_names = []
            for i in range(n_components):
                name = st.text_input(f"Name of Component {i+1}", value=f"C{i+1}", key=f"comp_name_{i}")
                component_names.append(name)
            
            expert_weights = None
            if n_experts > 1:
                st.markdown("---")
                st.subheader("Expert Weights")
                st.write("Assign weights to each expert (must sum to 1.0):")
                
                weights = []
                total_weight = 0
                
                for i in range(n_experts):
                    max_val = min(1.0, 1.0 - total_weight + (1.0/n_experts))
                    weight = st.number_input(
                        f"Weight for Expert {i+1}", 
                        min_value=0.0, 
                        max_value=max_val,
                        value=1.0/n_experts,
                        step=0.01,
                        format="%.2f",
                        key=f"weight_{i}",
                        help=f"Maximum allowed: {max_val:.2f}"
                    )
                    weights.append(weight)
                    total_weight += weight
                
                st.write(f"**Current total:** {total_weight:.2f}/1.0")
                
                if abs(total_weight - 1.0) > 0.001:
                    st.error(f"Weights must sum to 1.0. Current sum: {total_weight:.2f}")
                    st.stop()
                
                expert_weights = weights
            
            st.markdown("---")
            st.info("💡 **Tip**: Use the abbreviations for strength and influence assessments.")
        
        with st.expander("View Linguistic Terms Mapping", expanded=False):
            st.subheader("Linguistic Terms Mapping")
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Strength/Relevance Terms**")
                strength_data = []
                for abbr, it2 in LINGUISTIC_TERMS["strength"].items():
                    strength_data.append({
                        "Abbreviation": abbr,
                        "Full Form": FULL_FORMS[abbr],
                        "IT2TrFS Interval": format_it2(it2)
                    })
                strength_df = pd.DataFrame(strength_data)
                st.dataframe(strength_df, hide_index=True, use_container_width=True)
            
            with col2:
                st.write("**Influence Terms**")
                influence_data = []
                for abbr, it2 in LINGUISTIC_TERMS["influence"].items():
                    influence_data.append({
                        "Abbreviation": abbr,
                        "Full Form": FULL_FORMS[abbr],
                        "IT2TrFS Interval": format_it2(it2)
                    })
                influence_df = pd.DataFrame(influence_data)
                st.dataframe(influence_df, hide_index=True, use_container_width=True)
        
        if 'experts_data' not in st.session_state:
            st.session_state.experts_data = {}
        
        for expert_idx in range(n_experts):
            if expert_idx not in st.session_state.experts_data:
                st.session_state.experts_data[expert_idx] = {
                    'strengths_linguistic': ["HR" for _ in range(n_components)],
                    'influence_matrix_linguistic': [["ELI" for _ in range(n_components)] for _ in range(n_components)]
                }
            else:
                if len(st.session_state.experts_data[expert_idx]['strengths_linguistic']) != n_components:
                    st.session_state.experts_data[expert_idx]['strengths_linguistic'] = ["HR" for _ in range(n_components)]
                if len(st.session_state.experts_data[expert_idx]['influence_matrix_linguistic']) != n_components:
                    st.session_state.experts_data[expert_idx]['influence_matrix_linguistic'] = [["ELI" for _ in range(n_components)] for _ in range(n_components)]
        
        st.header(f"👨‍💼 Expert Input ({n_experts} Experts)" if n_experts > 1 else "👨‍💼 Data Input")
        
        expert_tabs = st.tabs([f"Expert {i+1}" for i in range(n_experts)]) if n_experts > 1 else [st.container()]
        
        strengths_list = []
        influence_matrices_list = []
        
        for expert_idx in range(n_experts):
            tab = expert_tabs[expert_idx] if n_experts > 1 else expert_tabs[0]
            
            with tab:
                if n_experts > 1:
                    st.subheader(f"Expert {expert_idx+1} Input")
                    if expert_weights:
                        st.write(f"**Weight:** {expert_weights[expert_idx]:.2f}")
                
                st.write("**Component Strengths/Relevance**")
                strengths = []
                
                strength_cols = st.columns(n_components + 1)
                with strength_cols[0]:
                    st.markdown("**Component**")
                for i in range(n_components):
                    with strength_cols[i + 1]:
                        st.markdown(f"**{component_names[i]}**")
                
                strength_input_cols = st.columns(n_components + 1)
                with strength_input_cols[0]:
                    st.markdown("**Strength Value**")
                
                for i in range(n_components):
                    with strength_input_cols[i + 1]:
                        current_strength = st.session_state.experts_data[expert_idx]['strengths_linguistic'][i]
                        
                        strength_term = st.selectbox(
                            f"Strength of {component_names[i]}", 
                            options=list(LINGUISTIC_TERMS["strength"].keys()),
                            index=list(LINGUISTIC_TERMS["strength"].keys()).index(current_strength),
                            key=f"strength_{expert_idx}_{i}",
                            help=FULL_FORMS[current_strength],
                            label_visibility="collapsed"
                        )
                        
                        interval = LINGUISTIC_TERMS["strength"][strength_term]
                        st.markdown(f"**UMF: {format_it2(interval)[1:-1]}**, LMF: {format_it2(interval)[1:-1]}")
                        
                        st.session_state.experts_data[expert_idx]['strengths_linguistic'][i] = strength_term
                        strengths.append(interval)
                
                st.write("**Influence Matrix**")
                st.write("Enter the influence between components (row influences column):")
                
                influence_matrix = [[None] * n_components for _ in range(n_components)]
                
                for i in range(n_components):
                    st.markdown(f"**Influences from {component_names[i]}**")
                    
                    header_cols = st.columns(n_components + 1)
                    with header_cols[0]:
                        st.markdown("**To →**")
                    for j in range(n_components):
                        with header_cols[j + 1]:
                            st.markdown(f"**{component_names[j]}**")
                    
                    input_cols = st.columns(n_components + 1)
                    with input_cols[0]:
                        st.markdown(f"**From {component_names[i]}**")
                    for j in range(n_components):
                        with input_cols[j + 1]:
                            if i == j:
                                st.markdown("—", help="Diagonal elements represent self-strength")
                            else:
                                current_influence = st.session_state.experts_data[expert_idx]['influence_matrix_linguistic'][i][j]
                                
                                influence_term = st.selectbox(
                                    f"{component_names[i]} → {component_names[j]}", 
                                    options=list(LINGUISTIC_TERMS["influence"].keys()),
                                    index=list(LINGUISTIC_TERMS["influence"].keys()).index(current_influence),
                                    key=f"inf_{expert_idx}_{i}_{j}",
                                    label_visibility="collapsed",
                                    help=FULL_FORMS[current_influence]
                                )
                                
                                interval = LINGUISTIC_TERMS["influence"][influence_term]
                                st.markdown(f"**UMF: {format_it2(interval)[1:-1]}**, LMF: {format_it2(interval)[1:-1]}")
                                
                                st.session_state.experts_data[expert_idx]['influence_matrix_linguistic'][i][j] = influence_term
                                influence_matrix[i][j] = interval
            
            strengths_list.append(strengths)
            influence_matrices_list.append(influence_matrix)
        
        if st.button("🚀 Run IT2TrFS WINGS Analysis", type="primary", use_container_width=True):
            with st.spinner("Calculating..."):
                results = wings_method_experts(strengths_list, influence_matrices_list, expert_weights)
            
            if results is None:
                return
            
            st.success("Analysis Complete!")
            
            tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
                "🔗 Flowchart", "📋 Expert Matrices", "🧮 IT2TrFS Matrices", 
                "📊 Results", "🏷️ Component Classification", "📈 Visualization", "📤 Export"
            ])
            
            with tab1:
                st.subheader("Component Interaction Flowchart")
                
                if n_experts > 1:
                    for expert_idx in range(n_experts):
                        st.subheader(f"Flowchart for Expert {expert_idx+1}")
                        flowchart = generate_flowchart_for_expert(
                            st.session_state.experts_data[expert_idx], 
                            component_names,
                            expert_idx
                        )
                        st.graphviz_chart(flowchart, use_container_width=True)
                else:
                    flowchart = generate_flowchart_for_expert(
                        st.session_state.experts_data[0], 
                        component_names
                    )
                    st.graphviz_chart(flowchart, use_container_width=True)
                
                st.markdown("""
                **Flowchart Explanation:**
                - **Nodes**: Represent components with their strength level (abbreviation) in parentheses
                - **Edges**: Show influences between components with their influence level (abbreviation)
                """)
            
            with tab2:
                st.subheader("Individual Expert SIDRMs")
                st.write("IT2TrFS WINGS averages SIDRMs directly across experts. Individual matrices are not stored separately.")
            
            with tab3:
                st.subheader("Average SIDRM")
                avg_df = format_it2_df(results['average_sidrm'], component_names, component_names)
                st.dataframe(avg_df, use_container_width=True)
                
                st.subheader("Normalized Matrix Z")
                z_df = format_it2_df(results['normalized_matrix'], component_names, component_names)
                st.dataframe(z_df, use_container_width=True)
                
                st.subheader("Total Matrix T (IT2TrFS)")
                t_df = format_it2_df(results['total_matrix'], component_names, component_names)
                st.dataframe(t_df, use_container_width=True)
            
            with tab4:
                st.subheader("Impact, Receptivity, Engagement, and Role Values")
                results_df = pd.DataFrame({
                    'Component': component_names,
                    'Total Impact (TI)': [format_it2(results['total_impact'][i]) for i in range(len(component_names))],
                    'Total Receptivity (TR)': [format_it2(results['total_receptivity'][i]) for i in range(len(component_names))],
                    'Engagement (TI+TR)': [format_it2(results['engagement'][i]) for i in range(len(component_names))],
                    'Role (TI-TR)': [format_it2(results['role'][i]) for i in range(len(component_names))]
                }).sort_values(by='Engagement (TI+TR)', key=lambda x: [defuzz_it2(results['engagement'][i]) for i in range(len(component_names))], ascending=False)
                
                styled_df = results_df.style.apply(
                    lambda x: ['background-color: #e6f3ff' if x.name == 'Engagement (TI+TR)' else '' for i in x],
                    axis=1
                ).format({
                    'Total Impact (TI)': '{}',
                    'Total Receptivity (TR)': '{}',
                    'Engagement (TI+TR)': '{}',
                    'Role (TI-TR)': '{}'
                })
                
                st.dataframe(styled_df, use_container_width=True, hide_index=True)
            
            with tab5:
                st.subheader("Component Classification")
                
                classification_data = []
                for i, (name, rel) in enumerate(zip(component_names, results['role'])):
                    status = "Cause" if defuzz_it2(rel) > 0 else "Effect"
                    classification_data.append({
                        "Component": name,
                        "Type": status,
                        "Role (TI-TR)": format_it2(rel),
                        "Engagement (TI+TR)": format_it2(results['engagement'][i])
                    })
                
                classification_df = pd.DataFrame(classification_data)
                
                cols = st.columns(3)
                for i, row in classification_df.iterrows():
                    with cols[i % 3]:
                        emoji = "➡️" if row['Type'] == 'Cause' else "⬅️"
                        st.metric(
                            label=f"{emoji} {row['Component']}",
                            value=row['Type'],
                            delta=f"Role: {row['Role (TI-TR)']}"
                        )
                
                fig, ax = plt.subplots(figsize=(10, 6))
                colors = ['#2ecc71' if defuzz_it2(results['role'][i]) > 0 else '#e74c3c' for i in range(len(component_names))]
                bars = ax.bar(classification_df['Component'], [defuzz_it2(results['role'][i]) for i in range(len(component_names))], color=colors)
                ax.set_title('Component Role Values')
                ax.set_ylabel('Role (TI-TR) (Defuzzified)')
                plt.xticks(rotation=45)
                st.pyplot(fig)
            
            with tab6:
                st.subheader("Visualization")
                
                fig, ax = plt.subplots(figsize=(10, 8))
                
                for i, name in enumerate(component_names):
                    color = 'green' if defuzz_it2(results['role'][i]) > 0 else 'red'
                    ax.scatter(defuzz_it2(results['engagement'][i]), defuzz_it2(results['role'][i]), s=150, color=color, alpha=0.7)
                    ax.annotate(name, (defuzz_it2(results['engagement'][i]), defuzz_it2(results['role'][i])), 
                                xytext=(5, 5), textcoords='offset points', fontsize=12)
                
                ax.axhline(y=0, color='gray', linestyle='--', alpha=0.7)
                ax.axvline(x=np.mean([defuzz_it2(results['engagement'][i]) for i in range(len(component_names))]), color='gray', linestyle='--', alpha=0.7)
                
                ax.set_xlabel('Engagement (TI+TR) (Defuzzified)')
                ax.set_ylabel('Role (TI-TR) (Defuzzified)')
                ax.set_title('Component Analysis: Engagement vs Role')
                ax.grid(True, alpha=0.3)
                
                st.pyplot(fig)
                
                fig2, ax2 = plt.subplots(figsize=(10, 6))
                y_pos = np.arange(len(component_names))
                ax2.barh(y_pos, [defuzz_it2(results['engagement'][i]) for i in range(len(component_names))], alpha=0.7)
                ax2.set_yticks(y_pos)
                ax2.set_yticklabels(component_names)
                ax2.set_xlabel('Engagement (TI+TR) (Defuzzified)')
                ax2.set_title('Component Engagement Ranking')
                ax2.grid(True, alpha=0.3, axis='x')
                
                st.pyplot(fig2)
                
                st.subheader("Cause-Effect Diagram")
                avg_engagement = np.mean([defuzz_it2(results['engagement'][i]) for i in range(len(component_names))])
                
                fig3, ax3 = plt.subplots(figsize=(10, 8))
                
                ax3.axhline(y=0, color='gray', linestyle='--', alpha=0.7)
                ax3.axvline(x=avg_engagement, color='gray', linestyle='--', alpha=0.7)
                
                for i, name in enumerate(component_names):
                    x = defuzz_it2(results['engagement'][i])
                    y = defuzz_it2(results['role'][i])
                    color = 'green' if y > 0 else 'red'
                    ax3.scatter(x, y, s=150, color=color, alpha=0.7)
                    ax3.annotate(name, (x, y), xytext=(5, 5), textcoords='offset points', fontsize=12)
                
                ax3.text(0.02, 0.98, "Cause Components", transform=ax3.transAxes, fontsize=14, 
                         verticalalignment='top', bbox=dict(boxstyle='round', facecolor='green', alpha=0.2))
                ax3.text(0.02, 0.02, "Effect Components", transform=ax3.transAxes, fontsize=14, 
                         verticalalignment='bottom', bbox=dict(boxstyle='round', facecolor='red', alpha=0.2))
                
                ax3.set_xlabel('Engagement (TI+TR) (Defuzzified)')
                ax3.set_ylabel('Role (TI-TR) (Defuzzified)')
                ax3.set_title('Cause-Effect Diagram')
                ax3.grid(True, alpha=0.3)
                
                st.pyplot(fig3)
            
            with tab7:
                st.subheader("Export Results")
                st.write("Download a comprehensive report of your IT2TrFS WINGS analysis in Word format.")
                
                doc = create_word_report(results, component_names, n_experts, expert_weights)
                
                html_link = get_word_download_link(doc)
                st.markdown(html_link, unsafe_allow_html=True)
                
                st.info("The Word report includes:")
                st.markdown("""
                - Analysis parameters and configuration
                - Impact, receptivity, engagement, and role results table
                - Component classification table
                - IT2TrFS and crisp matrices
                - Interpretation of results
                """)

if __name__ == "__main__":
    main()
