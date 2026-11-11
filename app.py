import io
import math
import time
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Geometri Jalan Interaktif - Tulip Edition",
    layout="wide",
    page_icon="🌷"
)

# Custom CSS: Memastikan SEMUA Teks Berwarna Gelap & Kontras (Tidak Putih/Samar)
st.markdown("""
<style>
    /* Background Utama Gradasi Pastel Soft */
    .stApp {
        background: linear-gradient(135deg, #fff9e6 0%, #ffe6f0 50%, #e8f4f8 100%);
        color: #2D3142;
        font-family: 'Poppins', sans-serif;
    }
    
    /* Sidebar Kiri - Pink Soft dengan Teks Ungu Tua Tegas */
    section[data-testid="stSidebar"] {
        background-color: rgba(255, 235, 243, 0.95) !important;
        border-right: 2px solid #F8BBD0;
    }
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3, 
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span {
        color: #3A0056 !important;
        font-weight: 700 !important;
    }

    /* Tab Navigasi - Teks Ungu Tua Pekat (Sangat Jelas & Tidak Putih) */
    .stTabs [data-baseweb="tab-list"] {
        background-color: rgba(255, 255, 255, 0.95);
        border-radius: 15px;
        padding: 6px;
        border: 2px solid #F48FB1;
    }
    .stTabs [data-baseweb="tab"] {
        color: #3A0056 !important; /* Warna teks tab pasif sangat gelap & jelas */
        font-weight: 800 !important;
        font-size: 15px !important;
        border-radius: 10px;
        padding: 10px 18px;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(45deg, #D81B60, #AD1457) !important;
        color: #FFFFFF !important; /* Teks putih tebal saat tab di-klik */
        box-shadow: 0px 4px 12px rgba(216, 27, 96, 0.4);
    }
    
    /* Input Box & Number Input */
    div[data-baseweb="input"] {
        background-color: #FFFFFF !important;
        border-radius: 8px;
        color: #111111 !important;
        border: 1px solid #F48FB1 !important;
    }

    /* Tombol Utama */
    .stButton>button {
        background: linear-gradient(45deg, #C2185B, #E91E63);
        color: white !important;
        border: none;
        border-radius: 12px;
        font-weight: bold;
        box-shadow: 0px 4px 10px rgba(194, 24, 91, 0.3);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(45deg, #880E4F, #C2185B);
        transform: translateY(-2px);
    }
    
    /* Card Keterangan Status Warna Mobil */
    .status-card {
        padding: 15px;
        border-radius: 15px;
        margin-bottom: 15px;
        font-weight: bold;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }
    .status-oke {
        background-color: rgba(200, 230, 201, 0.95);
        color: #1B5E20;
        border: 2px solid #81C784;
    }
    .status-sedang {
        background-color: rgba(255, 245, 157, 0.95);
        color: #E65100;
        border: 2px solid #FFF176;
    }
    .status-bahaya {
        background-color: rgba(255, 205, 210, 0.95);
        color: #B71C1C;
        border: 2px solid #E57373;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 1. KEAMANAN & AUTHENTICATION
# ----------------------------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    col_top1, col_top2 = st.columns([4, 1])
    with col_top2:
        st.markdown(
            "<p style='text-align: right; font-size: 13px; color: #880E4F; margin: 0;'>pw: "
            "<span style='font-family: monospace; font-weight: bold; color: #C2185B;'>cantikaaa27</span></p>",
            unsafe_allow_html=True
        )

    st.markdown("<h1 style='text-align: center; color: #3A0056;'>🌷 Perencanaan Geometrik Jalan Interaktif</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #880E4F;'>Standar TPGJAK 1997 / RAE 2021</h4>", unsafe_allow_html=True)
    
    st.image(
        "https://images.unsplash.com/photo-1544620347-c4fd4a3d5957?q=80&w=1200&auto=format&fit=crop",
        caption="Sistem Geometri Jalan Interaktif (Tulip Edition)",
        use_column_width=True
    )
    st.markdown("---")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("💡 **Petunjuk:** Password akses adalah **cantikaaa27**")
        pwd_input = st.text_input("Masukkan Password Akses:", type="password", placeholder="Ketik password di sini...")

        if st.button("Masuk Aplikasi 🚀", use_container_width=True):
            if pwd_input.strip() == "cantikaaa27":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Password salah! Silakan periksa kembali tulisan di kanan atas.")

    st.stop()

# ----------------------------------------------------------------------------
# 2. LOGIKA MATEMATIS & GEOMETRIK (BM 1997 / 2021)
# ----------------------------------------------------------------------------
V_TAB = [20, 30, 40, 50, 60, 80, 100, 120]
R_FC_TAB = [60, 130, 250, 350, 500, 900, 1500, 2500]

def sta_str(s: float, d: int = 2) -> str:
    return f"{int(s // 1000)}+{s % 1000:0{4 + d}.{d}f}"

def f_max(V: float) -> float:
    return -0.00065 * V + 0.192 if V < 80 else -0.00125 * V + 0.24

def r_min(V: float, emax: float) -> float:
    return V**2 / (127 * (emax + f_max(V)))

def hitung_tikungan(V: float, delta: float, emax: float, en: float, Rc: float, sta_pi: float):
    fm = f_max(V)
    Rmin = r_min(V, emax)
    Dmax = 181913.53 * (emax + fm) / (V**2)
    D = 1432.39 / Rc
    e = -emax * (D**2) / (Dmax**2) + 2 * emax * D / Dmax
    e = max(e, en)

    re = 0.035 if V <= 70 else 0.025
    Ls_waktu = V * 3 / 3.6
    Ls_shortt = 0.022 * (V**3) / (Rc * 0.4) - 2.727 * V * e / 0.4
    Ls_re = (e + en) * V / (3.6 * re)
    Ls_min = max(Ls_waktu, Ls_shortt, Ls_re)
    Ls_pakai = math.ceil(Ls_min / 5) * 5

    R_fc = float(np.interp(V, V_TAB, R_FC_TAB))
    d2 = math.radians(delta / 2)
    
    res = dict(
        V=V, delta=delta, emax=emax, en=en, Rc=Rc, fm=fm, Rmin=Rmin,
        Dmax=Dmax, D=D, e=e, Ls_waktu=Ls_waktu, Ls_shortt=Ls_shortt,
        Ls_re=Ls_re, Ls_min=Ls_min, R_fc=R_fc, sta_pi=sta_pi, Ls_pakai=Ls_pakai
    )

    if Rc >= R_fc:
        Tc = Rc * math.tan(d2)
        Ec = Tc * math.tan(d2 / 2)
        Lc = math.radians(delta) * Rc
        res.update(
            jenis="Full Circle (FC)", Ls=0.0, Ls_sup=Ls_pakai, ths=0.0,
            dc=delta, Lc=Lc, T=Tc, E=Ec, Xs=0.0, Ys=0.0, p=0.0, k=0.0
        )
    else:
        ths = math.degrees(Ls_pakai / (2 * Rc))
        dc = delta - 2 * ths
        Lc = math.radians(dc) * Rc
        Ls = float(Ls_pakai)
        jenis = "Spiral-Circle-Spiral (S-C-S)"
        
        if Lc < 25:
            jenis = "Spiral-Spiral (S-S)"
            ths = delta / 2
            Ls = math.radians(ths) * 2 * Rc
            Lc, dc = 0.0, 0.0

        Xs = Ls * (1 - Ls**2 / (40 * Rc**2))
        Ys = (Ls**2) / (6 * Rc)
        p = Ys - Rc * (1 - math.cos(math.radians(ths)))
        k = Xs - Rc * math.sin(math.radians(ths))
        Ts = (Rc + p) * math.tan(d2) + k
        Es = (Rc + p) / math.cos(d2) - Rc
        res.update(
            jenis=jenis, Ls=Ls, Ls_sup=Ls, ths=ths, dc=dc, Lc=Lc, T=Ts,
            E=Es, Xs=Xs, Ys=Ys, p=p, k=k
        )

    res["Ltot"] = res["Lc"] + 2 * res["Ls"]
    res["sta_awal"] = sta_pi - res["T"]
    res["sta_akhir"] = res["sta_awal"] + res["Ltot"]
    res["v_safe"] = math.sqrt(127 * Rc * (res["e"] + fm))
    return res

def heading(s, Ls, Lc, Rc, dl):
    Ltot = 2 * Ls + Lc
    th = np.zeros_like(s)
    if Ls > 0:
        ths = Ls / (2 * Rc)
        m1 = (s > 0) & (s <= Ls)
        th[m1] = (s[m1]**2) / (2 * Rc * Ls)
        m2 = (s > Ls) & (s < Ls + Lc)
        th[m2] = ths + (s[m2] - Ls) / Rc
        m3 = (s >= Ls + Lc) & (s <= Ltot)
        th[m3] = dl - ((Ltot - s[m3])**2) / (2 * Rc * Ls)
    else:
        m = (s > 0) & (s <= Ltot)
        th[m] = s[m] / Rc
    th[s > Ltot] = dl
    return th

def hitung_profil(pvi_df, V):
    d = pvi_df.dropna().sort_values(pvi_df.columns[0]).reset_index(drop=True)
    if len(d) < 2:
        return None
    sta = d.iloc[:, 0].to_numpy(float)
    el = d.iloc[:, 1].to_numpy(float)
    lvm = d.iloc[:, 2].to_numpy(float)
    g = np.diff(el) / np.diff(sta)
    Jh = 0.278 * V * 2.5 + (V**2) / (254 * 0.40)

    kurva = []
    for i in range(1, len(sta) - 1):
        A = (g[i] - g[i - 1]) * 100
        if abs(A) < 1e-9:
            continue
        a = abs(A)
        if A < 0:
            L = a * (Jh**2) / 399
            if L < Jh:
                L = 2 * Jh - 399 / a
        else:
            L = a * (Jh**2) / (120 + 3.5 * Jh
