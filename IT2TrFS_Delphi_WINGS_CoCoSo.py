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
            div = max(umf_params) if len_
