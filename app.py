import io
import math
import time
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

# ----------------------------------------------------------------------------
# KONFIGURASI HALAMAN STREAMLIT & TEMA
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Geometri Jalan Interaktif - Tulip Edition",
    layout="wide",
    page_icon="🌷"
)

# Custom CSS: Memastikan SEMUA Teks Berwarna Gelap & Kontras (Termasuk Tab Pink Tua)
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
        background-color: rgba(255, 235, 243, 0.98) !important;
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

    /* Tab Navigasi - Teks Pink Tua Pekat (#880E4F) agar Kontras & Tidak Nabrak Background */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #FFFFFF !important;
        border-radius: 15px;
        padding: 8px;
        border: 2px solid #F48FB1;
        box-shadow: 0 4px 12px rgba(136, 14, 79, 0.08);
    }
    .stTabs [data-baseweb="tab"],
    .stTabs [data-baseweb="tab"] p,
    .stTabs [data-baseweb="tab"] span,
    .stTabs button[role="tab"],
    .stTabs button[role="tab"] p {
        color: #880E4F !important; /* Warna Pink Tua Pekat */
        font-weight: 800 !important;
        font-size: 15px !important;
        border-radius: 10px;
        padding: 8px 14px;
        background-color: transparent !important;
    }
    .stTabs [data-baseweb="tab"]:hover,
    .stTabs button[role="tab"]:hover {
        background-color: #FCE4EC !important;
        color: #AD1457 !important;
    }
    .stTabs [aria-selected="true"],
    .stTabs button[aria-selected="true"] {
        background: linear-gradient(45deg, #AD1457, #D81B60) !important;
        box-shadow: 0px 4px 12px rgba(173, 20, 87, 0.35);
    }
    .stTabs [aria-selected="true"] p,
    .stTabs [aria-selected="true"] span,
    .stTabs button[aria-selected="true"] p {
        color: #FFFFFF !important; /* Teks putih bersih saat tab aktif dipilih */
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
        background-color: rgba(200, 230, 201, 0.98);
        color: #1B5E20;
        border: 2px solid #81C784;
    }
    .status-sedang {
        background-color: rgba(255, 245, 157, 0.98);
        color: #E65100;
        border: 2px solid #FFF176;
    }
    .status-bahaya {
        background-color: rgba(255, 205, 210, 0.98);
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
        use_container_width=True
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
    sign = "-" if s < 0 else ""
    s_abs = abs(s)
    return f"{sign}{int(s_abs // 1000)}+{s_abs % 1000:0{4 + d}.{d}f}"

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
            L = a * (Jh**2) / (120 + 3.5 * Jh)
            if L < Jh:
                L = 2 * Jh - (120 + 3.5 * Jh) / a
        
        L = math.ceil(max(L, 0.6 * V, V * 3 / 3.6) / 10) * 10
        if lvm[i] > 0:
            L = lvm[i]
            
        kurva.append(dict(
            i=i, sta=sta[i], el=el[i], A=A, Lv=L, Ev=a * L / 800,
            bvc=sta[i] - L / 2, evc=sta[i] + L / 2,
            jenis="Cembung" if A < 0 else "Cekung",
            g1=g[i - 1], g2=g[i]
        ))

    x = np.arange(sta[0], sta[-1] + 0.01, 1.0)
    y = np.interp(x, sta, el)
    for k in kurva:
        m = (x >= k["bvc"]) & (x <= k["evc"])
        dx = x[m] - k["bvc"]
        yb = k["el"] - k["g1"] * k["Lv"] / 2
        y[m] = yb + k["g1"] * dx + (k["g2"] - k["g1"]) / (2 * k["Lv"]) * (dx**2)
        
    return dict(sta=sta, el=el, g=g * 100, kurva=kurva, x=x, y=y, Jh=Jh)

# ----------------------------------------------------------------------------
# 3. SIDEBAR INPUT & CALCULATIONS
# ----------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Parameter Rencana")
    if st.button("🚪 Keluar / Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
        
    st.subheader("1. Kriteria Desain")
    V = st.number_input("Kecepatan Rencana Vr (km/jam)", 20, 120, 60, 10)
    v_sim = st.number_input("Kecepatan Kendaraan Simulasi (km/jam)", 10, 150, 65, 5)
    delta = st.number_input("Sudut Tikungan Δ (°)", 1.0, 179.0, 45.0, 0.5)
    emax = st.number_input("e maks (%)", 4.0, 12.0, 10.0, 0.5) / 100
    en = st.number_input("Kemiringan Normal en (%)", 1.0, 4.0, 2.0, 0.5) / 100
    
    Rmin = r_min(V, emax)
    st.caption(f"ℹ️ *R minimum acuan TPGJAK = {Rmin:.2f} m*")

    mode = st.radio("Jari-jari Tikungan Rc", ["Otomatis (Dibulatkan)", "Manual Input"])
    if mode.startswith("Oto"):
        Rc = float(math.ceil(Rmin / 10) * 10)
    else:
        Rc = st.number_input("Rc (m)", 10.0, 5000.0, 150.0, 5.0)
        
    sta_pi = st.number_input("STA PI (m)", 0.0, 100000.0, 300.0, 10.0)
    arah = st.radio("Arah Tikungan", ["Kiri", "Kanan"], horizontal=True)
    B = st.number_input("Lebar Perkerasan Jalan B (m)", 3.0, 30.0, 7.0, 0.5)

r = hitung_tikungan(V, delta, emax, en, Rc, sta_pi)
sg = 1 if arah == "Kiri" else -1
Ls, Lc = r["Ls"], r["Lc"]
dl = math.radians(r["delta"])
Ltot = r["Ltot"]
ext = math.ceil(max(120, 1.4 * r["T"]) / 10) * 10

s_vals = np.unique(np.concatenate([np.arange(-ext, Ltot + ext, 0.5), [0, Ls, Ls + Lc, Ltot]]))
th_vals = heading(s_vals, Ls, Lc, Rc, dl)
dS = np.diff(s_vals)
x_horiz = np.concatenate([[0], np.cumsum(dS * (np.cos(th_vals[:-1]) + np.cos(th_vals[1:])) / 2)])
y_horiz = np.concatenate([[0], np.cumsum(dS * (np.sin(th_vals[:-1]) + np.sin(th_vals[1:])) / 2)])
i0 = np.argmin(np.abs(s_vals))
x_horiz, y_horiz = x_horiz - x_horiz[i0], (y_horiz - y_horiz[i0]) * sg
ths_ = th_vals * sg
nx, ny = -np.sin(ths_), np.cos(ths_)

# State DataFrame Profil Vertikal & Tanah Asli
if "pvi_df" not in st.session_state:
    st.session_state.pvi_df = pd.DataFrame({
        "STA (m)": [0.0, 250.0, 500.0, 750.0],
        "Elevasi Rencana (m)": [100.0, 105.0, 101.0, 107.0],
        "Lv Manual (m)": [0.0, 0.0, 0.0, 0.0]
    })

if "tanah_df" not in st.session_state:
    st.session_state.tanah_df = pd.DataFrame({
        "STA (m)": [0.0, 125.0, 250.0, 375.0, 500.0, 625.0, 750.0],
        "Elevasi Tanah (m)": [101.0, 103.5, 104.0, 102.5, 101.8, 105.0, 106.5]
    })

if "anim_active" not in st.session_state:
    st.session_state.anim_active = False
if "anim_sta_idx" not in st.session_state:
    st.session_state.anim_sta_idx = 0

P_global = hitung_profil(st.session_state.pvi_df, V)
if P_global is not None:
    z_elev = np.interp(s_vals, P_global["x"], P_global["y"])
else:
    z_elev = np.zeros_like(s_vals)

# ----------------------------------------------------------------------------
# 4. EXPORT ENGINE (CAD DXF & EXCEL BM 2021)
# ----------------------------------------------------------------------------
def generate_excel_bm2021(r, pvi_data):
    wb = openpyxl.Workbook()
    font_title = Font(name="Calibri", size=14, bold=True, color="3A0056")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="C2185B", end_color="C2185B", fill_type="solid")
    fill_input = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
    border_thin = Border(left=Side(style='thin', color='F8BBD0'), right=Side(style='thin', color='F8BBD0'),
                         top=Side(style='thin', color='F8BBD0'), bottom=Side(style='thin', color='F8BBD0'))

    ws1 = wb.active
    ws1.title = "Tikungan Horisontal"
    ws1["A1"] = "REKAPITULASI KALKULASI GEOMETRIK JALAN (BINA MARGA 2021)"
    ws1["A1"].font = font_title

    headers = ["Parameter Geometrik", "Nilai", "Satuan", "Formula / Acuan BM 2021"]
    for col_i, h in enumerate(headers, start=1):
        cell = ws1.cell(row=3, column=col_i, value=h)
        cell.font = font_header
        cell.fill = fill_header

    fc = r["jenis"].startswith("Full")
    t_start = "TC" if fc else "TS"
    t_end = "CT" if fc else "ST"

    params = [
        ("Jenis Tikungan", r["jenis"], "-", "Kriteria R_fc (BM 2021)"),
        ("Kecepatan Rencana (Vr)", r["V"], "km/jam", "Input Desain"),
        ("Sudut Tikungan (Delta)", r["delta"], "°", "Input Desain"),
        ("Jari-Jari Tikungan (Rc)", r["Rc"], "m", "Input Desain"),
        ("Superelevasi Maksimum (e_max)", r["emax"] * 100, "%", "Standar Fisiologis"),
        ("Superelevasi Desain (e)", round(r["e"] * 100, 2), "%", "=-e_max*D^2/Dmax^2 + 2*e_max*D/Dmax"),
        ("Panjang Transisi (Ls)", r["Ls_sup"], "m", "=MAX(Ls_waktu, Ls_shortt, Ls_re)"),
        ("Panjang Tangen (T)", round(r["T"], 3), "m", "=(Rc+p)*TAN(RAD(Delta/2))+k"),
        ("Jarak Eksternal (E)", round(r["E"], 3), "m", "=(Rc+p)/COS(RAD(Delta/2))-Rc"),
        ("Panjang Total Tikungan (Ltot)", round(r["Ltot"], 3), "m", "=Lc + 2*Ls"),
        ("Kecepatan Aman Maksimum (V_safe)", round(r["v_safe"], 2), "km/jam", "=SQRT(127*Rc*(e + f_max))"),
        ("STA PI", sta_str(r["sta_pi"]), "m", "Input Desain"),
        (f"STA {t_start} (Awal Tikungan)", sta_str(r["sta_awal"]), "m", "=STA_PI - Tangen(T)"),
        (f"STA {t_end} (Akhir Tikungan)", sta_str(r["sta_akhir"]), "m", f"=STA_{t_start} + Ltot")
    ]

    for row_i, (p, v, u, f) in enumerate(params, start=4):
        ws1.cell(row=row_i, column=1, value=p).border = border_thin
        c_v = ws1.cell(row=row_i, column=2, value=v)
        c_v.border = border_thin
        c_v.fill = fill_input
        ws1.cell(row=row_i, column=3, value=u).border = border_thin
        ws1.cell(row=row_i, column=4, value=f).border = border_thin

    if pvi_data is not None:
        ws2 = wb.create_sheet(title="Alinyemen Vertikal")
        ws2["A1"] = "REKAPITULASI ALINYEMEN VERTIKAL"
        ws2["A1"].font = font_title
        
        v_headers = ["PVI", "STA (m)", "Elevasi PVI (m)", "Kelandaian g1 (%)", "Kelandaian g2 (%)", "Beda Kelandaian A (%)", "Panjang Kurva Lv (m)", "Jenis Lengkung"]
        for col_i, h in enumerate(v_headers, start=1):
            cell = ws2.cell(row=3, column=col_i, value=h)
            cell.font = font_header
            cell.fill = fill_header

        for idx, k in enumerate(pvi_data["kurva"], start=1):
            row_vals = [f"PVI-{idx}", k["sta"], k["el"], round(k["g1"]*100, 2), round(k["g2"]*100, 2), round(k["A"], 2), k["Lv"], k["jenis"]]
            for col_i, val in enumerate(row_vals, start=1):
                c = ws2.cell(row=idx+3, column=col_i, value=val)
                c.border = border_thin

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def export_dxf_full_autocad(x_h, y_h, z_e, s_v, r_dict, B_w, nx_v, ny_v):
    lines = [
        "0", "SECTION", "2", "HEADER", "0", "ENDSEC",
        "0", "SECTION", "2", "TABLES",
        "0", "TABLE", "2", "LAYER",
        "0", "LAYER", "2", "AS_JALAN_3D", "70", "0", "62", "7", "6", "CONTINUOUS",
        "0", "LAYER", "2", "TEPI_KIRI_3D", "70", "0", "62", "1", "6", "CONTINUOUS",
        "0", "LAYER", "2", "TEPI_KANAN_3D", "70", "0", "62", "5", "6", "CONTINUOUS",
        "0", "LAYER", "2", "DENAH_ANNOTATION", "70", "0", "62", "2", "6", "CONTINUOUS",
        "0", "LAYER", "2", "PROFIL_VERTIKAL", "70", "0", "62", "3", "6", "CONTINUOUS",
        "0", "ENDTAB", "0", "ENDSEC",
        "0", "SECTION", "2", "ENTITIES"
    ]

    lines.extend(["0", "POLYLINE", "8", "AS_JALAN_3D", "66", "1", "70", "8"])
    for x, y, z in zip(x_h, y_h, z_e):
        lines.extend(["0", "VERTEX", "8", "AS_JALAN_3D", "10", f"{float(x):.3f}", "20", f"{float(y):.3f}", "30", f"{float(z):.3f}"])
    lines.extend(["0", "SEQEND"])

    xl, yl = x_h + (B_w/2) * nx_v, y_h + (B_w/2) * ny_v
    xr, yr = x_h - (B_w/2) * nx_v, y_h - (B_w/2) * ny_v

    lines.extend(["0", "POLYLINE", "8", "TEPI_KIRI_3D", "66", "1", "70", "8"])
    for x, y, z in zip(xl, yl, z_e):
        lines.extend(["0", "VERTEX", "8", "TEPI_KIRI_3D", "10", f"{float(x):.3f}", "20", f"{float(y):.3f}", "30", f"{float(z):.3f}"])
    lines.extend(["0", "SEQEND"])

    lines.extend(["0", "POLYLINE", "8", "TEPI_KANAN_3D", "66", "1", "70", "8"])
    for idx_pt, (x, z) in enumerate(zip(xr, z_e)):
        lines.extend(["0", "VERTEX", "8", "TEPI_KANAN_3D", "10", f"{float(x):.3f}", "20", f"{float(yr[idx_pt]):.3f}", "30", f"{float(z):.3f}"])
    lines.extend(["0", "SEQEND"])

    y_off = -150.0
    lines.extend(["0", "POLYLINE", "8", "PROFIL_VERTIKAL", "66", "1", "70", "0"])
    for s, z in zip(s_v, z_e):
        lines.extend(["0", "VERTEX", "8", "PROFIL_VERTIKAL", "10", f"{float(s):.3f}", "20", f"{y_off + float(z):.3f}", "30", "0.0"])
    lines.extend(["0", "SEQEND"])

    lines.extend(["0", "ENDSEC", "0", "EOF"])
    return "\n".join(lines)

# ----------------------------------------------------------------------------
# 5. TAMPILAN APLIKASI UTAMA
# ----------------------------------------------------------------------------
st.markdown("<h1 style='color: #3A0056; text-align: center; font-weight: 800;'>🌸 Geometri Jalan Interaktif & Simulasi 3D</h1>", unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📍 Denah Tikungan (2D)", 
    "📈 Diagram Superelevasi", 
    "🚘 Simulasi 3D Mobil Jalan", 
    "📊 Profil Memanjang & Ekspor CAD",
    "📋 Tabel Hitungan & Excel"
])

# ----------------------------------------------------------------------------
# TAB 1: DENAH TIKUNGAN 2D (DETAIL SEPERTI REFERENSI)
# ----------------------------------------------------------------------------
with tab1:
    st.markdown("<h3 style='color: #3A0056;'>📍 Visualisasi Denah Trase Tikungan 2D</h3>", unsafe_allow_html=True)
    
    fig_2d = go.Figure()
    
    # 1. Badan Jalan (Arsiran Abu-abu Lembut) + Tepi Kiri & Kanan
    xl, yl = x_horiz + (B/2) * nx, y_horiz + (B/2) * ny
    xr, yr = x_horiz - (B/2) * nx, y_horiz - (B/2) * ny
    
    fig_2d.add_trace(go.Scatter(
        x=np.concatenate([xl, xr[::-1]]),
        y=np.concatenate([yl, yr[::-1]]),
        fill='toself',
        fillcolor='rgba(200, 200, 210, 0.25)',
        line=dict(color='rgba(0,0,0,0)'),
        name='Badan Perkerasan',
        hoverinfo='skip'
    ))
    fig_2d.add_trace(go.Scatter(x=xl, y=yl, mode='lines', name='Tepi Perkerasan', line=dict(color='#616161', width=1.5)))
    fig_2d.add_trace(go.Scatter(x=xr, y=yr, mode='lines', showlegend=False, line=dict(color='#616161', width=1.5)))
    
    # 2. As Jalan (Hitam Pekat & Jelas)
    fig_2d.add_trace(go.Scatter(x=x_horiz, y=y_horiz, mode='lines', name='As jalan', line=dict(color='#111111', width=3)))
    
    # 3. Garis Singgung Tangen (Biru Putus-putus) & Titik PI (Kotak Hitam)
    idx_start = np.argmin(np.abs(s_vals - 0))
    idx_end = np.argmin(np.abs(s_vals - r["Ltot"]))
    
    x_ts, y_ts = x_horiz[idx_start], y_horiz[idx_start]
    x_st, y_st = x_horiz[idx_end], y_horiz[idx_end]
    
    # Koordinat PI: berada pada perpanjangan tangen awal sejauh T
    x_pi = x_ts + r["T"] * np.cos(ths_[idx_start])
    y_pi = y_ts + r["T"] * np.sin(ths_[idx_start])
    
    fig_2d.add_trace(go.Scatter(
        x=[x_ts, x_pi, x_st],
        y=[y_ts, y_pi, y_st],
        mode='lines',
        name='Garis singgung',
        line=dict(color='#1565C0', width=1.8, dash='dash')
    ))
    
    fig_2d.add_trace(go.Scatter(
        x=[x_pi], y=[y_pi],
        mode='markers',
        name='Titik PI',
        marker=dict(size=10, color='#111111', symbol='square')
    ))
    
    # Anotasi PI dengan kotak background putih agar tidak nabrak garis
    fig_2d.add_annotation(
        x=x_pi, y=y_pi,
        text=f"<b>PI</b><br>Δ = {r['delta']:.1f}°<br>STA {sta_str(r['sta_pi'])}",
        showarrow=False,
        xshift=12, yshift=32 if arah == "Kanan" else -32,
        xanchor="left",
        font=dict(color="#111111", size=12),
        bgcolor="rgba(255,255,255,0.92)",
        bordercolor="#BDBDBD", borderwidth=1, borderpad=4
    )
    
    # 4. Titik Pusat Lingkaran (O) & Garis Jari-jari Hijau Putus-putus
    if r["jenis"].startswith("Full"):
        mid_idx = idx_start
    else:
        mid_idx = np.argmin(np.abs(s_vals - (r["Ltot"] / 2)))
        
    # Arah normal menuju pusat kelengkungan (kiri = +normal, kanan = -normal)
    x_center = x_horiz[mid_idx] + sg * r["Rc"] * nx[mid_idx]
    y_center = y_horiz[mid_idx] + sg * r["Rc"] * ny[mid_idx]
    
    fig_2d.add_trace(go.Scatter(
        x=[x_horiz[mid_idx], x_center],
        y=[y_horiz[mid_idx], y_center],
        mode='lines',
        name=f'Jari-jari Rc = {r["Rc"]:.0f} m',
        line=dict(color='#2E7D32', width=1.5, dash='dot')
    ))
    fig_2d.add_trace(go.Scatter(
        x=[x_center], y=[y_center],
        mode='markers',
        name='Pusat Lengkungan (O)',
        marker=dict(size=13, color='#2E7D32', symbol='cross-thin', line=dict(width=2, color='#2E7D32'))
    ))
    fig_2d.add_annotation(
        x=x_center, y=y_center,
        text=f"<b>R = {r['Rc']:.0f} m</b>",
        showarrow=False,
        xshift=14, xanchor="left",
        font=dict(color="#1B5E20", size=12),
        bgcolor="rgba(255,255,255,0.9)",
        borderpad=2
    )
    
    # 5. Garis Patok Stasioning Minor setiap 50 m (0+050, 0+100, dst.)
    sta_min_plot = math.ceil((r["sta_awal"] + s_vals[0]) / 50.0) * 50.0
    sta_max_plot = math.floor((r["sta_awal"] + s_vals[-1]) / 50.0) * 50.0
    tick_stas = np.arange(sta_min_plot, sta_max_plot + 1, 50.0)
    
    tick_half_len = max(B * 0.85, 5.5)
    for st_val in tick_stas:
        s_rel = st_val - r["sta_awal"]
        if s_vals[0] <= s_rel <= s_vals[-1]:
            idx_t = np.argmin(np.abs(s_vals - s_rel))
            xt0, yt0 = x_horiz[idx_t], y_horiz[idx_t]
            nxt, nyt = nx[idx_t], ny[idx_t]
            
            # Garis patok melintang kecil
            fig_2d.add_trace(go.Scatter(
                x=[xt0 - tick_half_len * nxt, xt0 + tick_half_len * nxt],
                y=[yt0 - tick_half_len * nyt, yt0 + tick_half_len * nyt],
                mode='lines',
                showlegend=False,
                hoverinfo='skip',
                line=dict(color='#757575', width=1.2)
            ))
            # Label STA di sisi luar tikungan agar rapi
            lbl_side = -sg
            fig_2d.add_annotation(
                x=xt0 + lbl_side * (tick_half_len + 2.5) * nxt,
                y=yt0 + lbl_side * (tick_half_len + 2.5) * nyt,
                text=sta_str(st_val, d=0),
                showarrow=False,
                font=dict(color="#424242", size=10),
                bgcolor="rgba(255,255,255,0.75)",
                borderpad=1
            )

    # 6. Titik-Titik Utama Tikungan (Merah Bulat + Label Merah Tua Berlatar Putih)
    fc = r["jenis"].startswith("Full")
    is_ss = r["jenis"].startswith("Spiral-Spiral")
    
    pts_x = [x_ts, x_st]
    pts_y = [y_ts, y_st]
    if fc:
        pts_lbls = [f"<b>TC</b><br>{sta_str(r['sta_awal'])}", f"<b>CT</b><br>{sta_str(r['sta_akhir'])}"]
    else:
        pts_lbls = [f"<b>TS</b><br>{sta_str(r['sta_awal'])}", f"<b>ST</b><br>{sta_str(r['sta_akhir'])}"]
        if is_ss:
            idx_mid = np.argmin(np.abs(s_vals - r["Ls"]))
            pts_x.append(x_horiz[idx_mid])
            pts_y.append(y_horiz[idx_mid])
            pts_lbls.append(f"<b>SC=CS</b><br>{sta_str(r['sta_awal'] + r['Ls'])}")
        else:
            idx_sc = np.argmin(np.abs(s_vals - r["Ls"]))
            idx_cs = np.argmin(np.abs(s_vals - (r["Ls"] + r["Lc"])))
            pts_x.extend([x_horiz[idx_sc], x_horiz[idx_cs]])
            pts_y.extend([y_horiz[idx_sc], y_horiz[idx_cs]])
            pts_lbls.extend([
                f"<b>SC</b><br>{sta_str(r['sta_awal'] + r['Ls'])}",
                f"<b>CS</b><br>{sta_str(r['sta_awal'] + r['Ls'] + r['Lc'])}"
            ])

    fig_2d.add_trace(go.Scatter(
        x=pts_x, y=pts_y,
        mode='markers',
        name='Titik Utama (TS/SC/CS/ST)',
        marker=dict(size=10, color='#D50000', symbol='circle', line=dict(width=1.5, color='#FFFFFF'))
    ))
    
    for px, py, plbl in zip(pts_x, pts_y, pts_lbls):
        fig_2d.add_annotation(
            x=px, y=py,
            text=plbl,
            showarrow=False,
            yshift=-26 if arah == "Kanan" else 26,
            font=dict(color="#8E0000", size=11),
            bgcolor="rgba(255,255,255,0.92)",
            bordercolor="#FFCDD2", borderwidth=1, borderpad=3
        )

    arah_txt = "tikungan ke kiri" if arah == "Kiri" else "tikungan ke kanan"
    fig_2d.update_layout(
        title=dict(
            text=f"Denah Trase Tikungan - {r['jenis']} ({arah_txt})",
            font=dict(color='#111111', size=17),
            x=0.5, xanchor='center'
        ),
        xaxis=dict(
            title="x (m)", title_font=dict(color='#111111', size=13),
            tickfont=dict(color='#222222'), gridcolor='#E0E0E0',
            showline=True, linecolor='#333333', mirror=True
        ),
        yaxis=dict(
            title="y (m)", title_font=dict(color='#111111', size=13),
            tickfont=dict(color='#222222'), gridcolor='#E0E0E0',
            showline=True, linecolor='#333333', mirror=True,
            scaleanchor="x", scaleratio=1
        ),
        legend=dict(
            bgcolor="rgba(255, 255, 255, 0.96)",
            bordercolor="#9E9E9E",
            borderwidth=1,
            font=dict(color="#111111", size=12),
            x=0.99, y=0.99, xanchor="right", yanchor="top"
        ),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        margin=dict(l=60, r=40, t=60, b=50),
        height=600
    )
    st.plotly_chart(fig_2d, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 2: DIAGRAM SUPERELEVASI (DETAIL ROTASI PADA SUMBU JALAN)
# ----------------------------------------------------------------------------
with tab2:
    st.markdown("<h3 style='color: #3A0056;'>📈 Diagram Superelevasi (Rotasi pada Sumbu Jalan)</h3>", unsafe_allow_html=True)
    
    e_desain = r["e"] * 100
    en_val = r["en"] * 100
    ls_val = r["Ls_sup"]
    ltot_val = r["Ltot"]
    sta_0 = r["sta_awal"]
    
    # Tentukan posisi titik kritis relatif terhadap STA awal (s = 0 di TS/TC)
    if r["jenis"].startswith("Full"):
        # Pada Full Circle, transisi biasanya 2/3 Ls di tangen dan 1/3 Ls di dalam lengkung
        s_start_trans = - (2.0 / 3.0) * ls_val
        s_full_start = (1.0 / 3.0) * ls_val
        s_full_end = ltot_val - (1.0 / 3.0) * ls_val
        s_end_trans = ltot_val + (2.0 / 3.0) * ls_val
        
        krit_s = [s_start_trans, 0.0, s_full_start, s_full_end, ltot_val, s_end_trans]
        krit_lbl = ["Awal Transisi", "TC", "e penuh", "e penuh", "CT", "Akhir Transisi"]
    else:
        s_start_trans = 0.0
        s_full_start = ls_val
        s_full_end = ltot_val - ls_val
        s_end_trans = ltot_val
        
        if r["jenis"].startswith("Spiral-Spiral"):
            krit_s = [0.0, ls_val, ltot_val]
            krit_lbl = ["TS", "SC=CS", "ST"]
        else:
            krit_s = [0.0, ls_val, ltot_val - ls_val, ltot_val]
            krit_lbl = ["TS", "SC", "CS", "ST"]

    # Hitung titik nol (kemiringan tepi luar = 0%) pada zona runout
    # Tepi luar naik linear dari -en ke +e sepanjang (s_full_start - s_start_trans)
    trans_len = s_full_start - s_start_trans
    s_zero_1 = s_start_trans + trans_len * (en_val / (e_desain + en_val))
    s_zero_2 = s_end_trans - trans_len * (en_val / (e_desain + en_val))
    
    # Titik belok tepi dalam saat tepi luar mencapai +en (crown removal selesai)
    s_crown_1 = s_start_trans + trans_len * ((2 * en_val) / (e_desain + en_val))
    s_crown_2 = s_end_trans - trans_len * ((2 * en_val) / (e_desain + en_val))

    margin_s = max(40.0, 0.25 * ltot_val)
    s_min_sup = min(0.0, s_start_trans) - margin_s
    s_max_sup = max(ltot_val, s_end_trans) + margin_s
    
    # Evaluasi piecewise tepat di titik-titik simpul agar garis patah tajam & akurat
    key_nodes = np.unique(np.sort([
        s_min_sup, s_start_trans, s_zero_1, s_crown_1,
        s_full_start, s_full_end, s_crown_2, s_zero_2, s_end_trans, s_max_sup
    ]))
    
    e_luar = np.zeros_like(key_nodes)
    e_dalam = np.zeros_like(key_nodes)
    
    for i, sv in enumerate(key_nodes):
        if sv <= s_start_trans or sv >= s_end_trans:
            e_luar[i] = -en_val
            e_dalam[i] = -en_val
        elif s_start_trans < sv < s_full_start:
            frac = (sv - s_start_trans) / trans_len
            e_luar[i] = -en_val + frac * (e_desain + en_val)
            e_dalam[i] = min(-en_val, -e_luar[i])
        elif s_full_start <= sv <= s_full_end:
            e_luar[i] = e_desain
            e_dalam[i] = -e_desain
        elif s_full_end < sv < s_end_trans:
            frac = (s_end_trans - sv) / trans_len
            e_luar[i] = -en_val + frac * (e_desain + en_val)
            e_dalam[i] = min(-en_val, -e_luar[i])
            
    sta_nodes = sta_0 + key_nodes
    
    fig_sup = go.Figure()
    
    # Arsiran transisi lembut di bawah kurva luar & dalam
    fig_sup.add_trace(go.Scatter(
        x=sta_nodes, y=e_luar,
        mode='lines',
        name='Tepi perkerasan luar',
        line=dict(color='#C62828', width=3),
        fill='tozeroy',
        fillcolor='rgba(198, 40, 40, 0.08)'
    ))
    
    fig_sup.add_trace(go.Scatter(
        x=sta_nodes, y=e_dalam,
        mode='lines',
        name='Tepi perkerasan dalam',
        line=dict(color='#0D47A1', width=3),
        fill='tozeroy',
        fillcolor='rgba(13, 71, 161, 0.05)'
    ))
    
    # Garis Sumbu Jalan (0%)
    fig_sup.add_hline(y=0, line_width=1.2, line_color="#424242")
    
    # Titik Nol Kemiringan (0%) pada Tepi Luar
    fig_sup.add_trace(go.Scatter(
        x=[sta_0 + s_zero_1, sta_0 + s_zero_2],
        y=[0.0, 0.0],
        mode='markers',
        name='Kemiringan Datar (0%)',
        marker=dict(size=8, color='#111111', symbol='circle')
    ))
    
    # Garis Vertikal Putus-putus & Label STA untuk TS, SC, CS, ST
    y_top_limit = max(e_desain + 3.5, 13.0)
    y_bot_limit = min(-e_desain - 3.5, -13.0)
    
    for ks, klbl in zip(krit_s, krit_lbl):
        sta_k = sta_0 + ks
        fig_sup.add_vline(x=sta_k, line_width=1, line_dash="dot", line_color="#757575")
        fig_sup.add_annotation(
            x=sta_k, y=y_top_limit - 1.2,
            text=f"<b>{klbl}</b><br>{sta_str(sta_k)}",
            showarrow=False,
            font=dict(color="#111111", size=11),
            bgcolor="rgba(255,255,255,0.92)",
            bordercolor="#E0E0E0", borderwidth=1, borderpad=3
        )
        
    # Anotasi nilai +e, -e, dan en
    mid_full_sta = sta_0 + (s_full_start + s_full_end) / 2
    fig_sup.add_annotation(
        x=mid_full_sta, y=e_desain + 0.9,
        text=f"<b>+e = {e_desain:.2f}%</b>",
        showarrow=False,
        font=dict(color="#B71C1C", size=12),
        bgcolor="rgba(255,255,255,0.92)", borderpad=2
    )
    fig_sup.add_annotation(
        x=mid_full_sta, y=-e_desain - 1.1,
        text=f"<b>-e = -{e_desain:.2f}%</b>",
        showarrow=False,
        font=dict(color="#0D47A1", size=12),
        bgcolor="rgba(255,255,255,0.92)", borderpad=2
    )
    fig_sup.add_annotation(
        x=sta_nodes[0] + margin_s * 0.45, y=-en_val - 1.1,
        text=f"<b>en = -{en_val:.1f}%</b>",
        showarrow=False,
        font=dict(color="#222222", size=11),
        bgcolor="rgba(255,255,255,0.92)", borderpad=2
    )
    
    # Format tick sumbu X menjadi stasioning (0+150, 0+200, dst.)
    tick_step = 50.0
    x_tick_vals = np.arange(
        math.ceil(sta_nodes[0] / tick_step) * tick_step,
        math.floor(sta_nodes[-1] / tick_step) * tick_step + 1,
        tick_step
    )
    x_tick_texts = [sta_str(v, d=0) for v in x_tick_vals]

    fig_sup.update_layout(
        title=dict(
            text="Diagram Superelevasi (rotasi pada sumbu jalan)",
            font=dict(color='#111111', size=17),
            x=0.5, xanchor='center'
        ),
        xaxis=dict(
            title="Stasioning",
            title_font=dict(color='#111111', size=13),
            tickfont=dict(color='#222222'),
            tickmode='array',
            tickvals=x_tick_vals,
            ticktext=x_tick_texts,
            gridcolor='#E0E0E0',
            showline=True, linecolor='#333333', mirror=True
        ),
        yaxis=dict(
            title="Kemiringan melintang (%)",
            title_font=dict(color='#111111', size=13),
            tickfont=dict(color='#222222'),
            range=[y_bot_limit, y_top_limit],
            gridcolor='#E0E0E0',
            showline=True, linecolor='#333333', mirror=True
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom", y=0.02,
            xanchor="center", x=0.5,
            bgcolor="rgba(255, 255, 255, 0.96)",
            bordercolor="#9E9E9E", borderwidth=1,
            font=dict(color="#111111", size=12)
        ),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        margin=dict(l=60, r=40, t=60, b=50),
        height=550
    )
    st.plotly_chart(fig_sup, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 3: PROFIL MEMANJANG & SIMULASI 3D
# ----------------------------------------------------------------------------
with tab3:
    st.markdown("<h3 style='color: #3A0056;'>🚘 Visualisasi 3D Jalan & Simulator Mobil Jalan Otomatis</h3>", unsafe_allow_html=True)
    
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.markdown("<b style='color:#3A0056;'>Edit Data PVI Alinyemen Vertikal:</b>", unsafe_allow_html=True)
        st.session_state.pvi_df = st.data_editor(st.session_state.pvi_df, num_rows="dynamic", key="pvi_ed", hide_index=True)
    with col_e2:
        st.markdown("<b style='color:#3A0056;'>Edit Data Elevasi Tanah Asli:</b>", unsafe_allow_html=True)
        st.session_state.tanah_df = st.data_editor(st.session_state.tanah_df, num_rows="dynamic", key="tnh_ed", hide_index=True)
        
    st.markdown("---")
    
    v_safe = r["v_safe"]
    if v_sim <= v_safe:
        car_color = "#1B5E20"
        status_box = f"""
        <div class="status-card status-oke">
            🟢 <b>KONDISI HIJAU (OKE):</b> Mobil Melaju Sangat Aman!<br>
            Kecepatan simulasi ({v_sim} km/jam) berada di bawah batas kecepatan aman tikungan ({v_safe:.1f} km/jam).
        </div>
        """
    elif v_sim <= v_safe + 10:
        car_color = "#E65100"
        status_box = f"""
        <div class="status-card status-sedang">
            🟡 <b>KONDISI KUNING (SEDANG / WASPADA):</b> Mobil Mendekati Batas Aman!<br>
            Kecepatan simulasi ({v_sim} km/jam) sedikit melampaui V_aman ({v_safe:.1f} km/jam). Harap kurangi kecepatan.
        </div>
        """
    else:
        car_color = "#B71C1C"
        status_box = f"""
        <div class="status-card status-bahaya">
            🔴 <b>KONDISI MERAH (TIDAK OKE / BAHAYA):</b> Over-Speed!<br>
            Kecepatan simulasi ({v_sim} km/jam) melampaui batas aman ({v_safe:.1f} km/jam). Berisiko selip atau kecelakaan.
        </div>
        """
        
    st.markdown(status_box, unsafe_allow_html=True)

    c_btn1, c_btn2, c_btn3 = st.columns([1, 1, 2])
    with c_btn1:
        if st.button("▶️ Jalan / Pause Mobil", use_container_width=True):
            st.session_state.anim_active = not st.session_state.anim_active
    with c_btn2:
        if st.button("🔄 Reset Ke Awal", use_container_width=True):
            st.session_state.anim_sta_idx = 0
            st.session_state.anim_active = False
            st.rerun()
    with c_btn3:
        speed_delay = st.slider("⚡ Kecepatan Animasi:", 0.01, 0.5, 0.05, 0.01)

    current_idx = st.session_state.anim_sta_idx
    selected_sta = st.slider(
        "🎮 Kontrol Posisi Mobil / STA (m):", 
        float(s_vals[0]), float(s_vals[-1]), 
        float(s_vals[current_idx]), 5.0
    )
    
    if not st.session_state.anim_active:
        st.session_state.anim_sta_idx = np.argmin(np.abs(s_vals - selected_sta))
        current_idx = st.session_state.anim_sta_idx

    xc, yc, zc = x_horiz[current_idx], y_horiz[current_idx], z_elev[current_idx]

    fig_3d = go.Figure()
    
    fig_3d.add_trace(go.Scatter3d(
        x=x_horiz, y=y_horiz, z=z_elev, mode='lines', name='As Jalan 3D',
        line=dict(color='#3A0056', width=6)
    ))
    fig_3d.add_trace(go.Scatter3d(x=xl, y=yl, z=z_elev, mode='lines', name='Tepi Kiri 3D', line=dict(color='#C2185B', width=3)))
    fig_3d.add_trace(go.Scatter3d(x=xr, y=yr, z=z_elev, mode='lines', name='Tepi Kanan 3D', line=dict(color='#C2185B', width=3)))
    
    fig_3d.add_trace(go.Scatter3d(
        x=[xc], y=[yc], z=[zc + 1.2], mode='markers+text', name='Mobil Simulasi',
        marker=dict(size=12, color=car_color, symbol='diamond'),
        text=[f"🚘 Mobil (STA {s_vals[current_idx]:.0f}m)"], textposition="top center",
        textfont=dict(color="#111111", size=12)
    ))

    dx_cam = np.cos(ths_[current_idx]) * 20
    dy_cam = np.sin(ths_[current_idx]) * 20
    
    camera_eye = dict(
        x=(xc - dx_cam) / 100,
        y=(yc - dy_cam) / 100,
        z=(zc + 3.0) / 100
    )

    fig_3d.update_layout(
        title=dict(text=f"Visualisasi 3D Trase Jalan & Simulator Mobil (STA {s_vals[current_idx]:.1f} m)", font=dict(color='#111111', size=16)),
        template="plotly_white",
        scene=dict(
            xaxis_title="X (m)", yaxis_title="Y (m)", zaxis_title="Elevasi Z (m)",
            aspectmode='data', camera_eye=camera_eye
        ),
        legend=dict(bgcolor="rgba(255,255,255,0.95)", font=dict(color="#111111")),
        margin=dict(l=0, r=0, b=0, t=40),
        paper_bgcolor="#FFFFFF",
        font=dict(color="#111111")
    )
    st.plotly_chart(fig_3d, use_container_width=True)

    if st.session_state.anim_active:
        time.sleep(speed_delay)
        st.session_state.anim_sta_idx += 4
        if st.session_state.anim_sta_idx >= len(s_vals):
            st.session_state.anim_sta_idx = 0
        st.rerun()

# ----------------------------------------------------------------------------
# TAB 4: PROFIL MEMANJANG & EKSPOR CAD
# ----------------------------------------------------------------------------
with tab4:
    st.markdown("<h3 style='color: #3A0056;'>📊 Profil Memanjang Vertikal Jalan & Ekspor CAD (.DXF)</h3>", unsafe_allow_html=True)
    
    if P_global is not None:
        t_df = st.session_state.tanah_df.dropna().sort_values(st.session_state.tanah_df.columns[0])
        fig_prof = go.Figure()
        
        fig_prof.add_trace(go.Scatter(
            x=t_df.iloc[:, 0], y=t_df.iloc[:, 1], mode='lines+markers',
            name='Elevasi Tanah Asli', line=dict(color='#1B5E20', dash='dash')
        ))
        fig_prof.add_trace(go.Scatter(
            x=s_vals, y=z_elev, mode='lines', name='Elevasi Rencana Jalan',
            line=dict(color='#C2185B', width=3)
        ))
        
        pvi_x = P_global["sta"]
        pvi_y = P_global["el"]
        fig_prof.add_trace(go.Scatter(
            x=pvi_x, y=pvi_y, mode='markers+text', name='Titik PVI',
            marker=dict(size=8, color='#3A0056', symbol='x'),
            text=[f"PVI-{i+1}" for i in range(len(pvi_x))], textposition="top center",
            textfont=dict(color="#111111")
        ))

        fig_prof.update_layout(
            title=dict(text="Grafik Profil Memanjang Vertikal (Tanah Asli vs Rencana)", font=dict(color='#111111', size=16)),
            xaxis=dict(title="Stasioning STA (m)", title_font=dict(color='#111111'), tickfont=dict(color='#222222'), gridcolor='#E0E0E0'),
            yaxis=dict(title="Elevasi Z (m)", title_font=dict(color='#111111'), tickfont=dict(color='#222222'), gridcolor='#E0E0E0'),
            legend=dict(bgcolor="rgba(255,255,255,0.95)", bordercolor="#9E9E9E", borderwidth=1, font=dict(color="#111111")),
            template="plotly_white", showlegend=True,
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
            font=dict(color="#111111")
        )
        st.plotly_chart(fig_prof, use_container_width=True)

    dxf_out = export_dxf_full_autocad(x_horiz, y_horiz, z_elev, s_vals, r, B, nx, ny)
    st.download_button(
        "📐 Unduh Gambar AutoCAD / Civil 3D (.DXF Multi-Layer)",
        dxf_out, file_name="Geometri_Jalan_BinaMarga.dxf", mime="application/dxf",
        use_container_width=True
    )

# ----------------------------------------------------------------------------
# TAB 5: TABEL HITUNGAN & EXCEL REPORT (BM 2021)
# ----------------------------------------------------------------------------
with tab5:
    st.markdown("<h3 style='color: #3A0056;'>📋 Tabel Hasil Perhitungan Tikungan & Laporan Excel</h3>", unsafe_allow_html=True)
    
    fc = r["jenis"].startswith("Full")
    a, b = ("TC", "CT") if fc else ("TS", "ST")
    rows = [
        ("Jenis Tikungan", r["jenis"], "Acuan Kriteria R_fc"),
        ("Kecepatan Rencana (Vr)", f"{r['V']:.0f}", "km/jam"),
        ("Sudut Tikungan (Δ)", f"{r['delta']:.2f}", "°"),
        ("Jari-Jari Rencana (Rc)", f"{r['Rc']:.2f}", "m"),
        ("Superelevasi Maksimum (e_max)", f"{r['emax']*100:.1f}", "%"),
        ("Superelevasi Desain (e)", f"{r['e']*100:.2f}", "%"),
        ("Panjang Transisi (Ls dipakai)", f"{r['Ls_sup']:.2f}", "m"),
        ("Panjang Tangen (T)", f"{r['T']:.3f}", "m"),
        ("Jarak Eksternal (E)", f"{r['E']:.3f}", "m"),
        ("Panjang Total Tikungan (Ltot)", f"{r['Ltot']:.3f}", "m"),
        ("Kecepatan Aman Maksimum (V_safe)", f"{r['v_safe']:.2f}", "km/jam"),
        (f"STA {a} (Awal Tikungan)", sta_str(r["sta_awal"]), "m"),
        ("STA PI", sta_str(r["sta_pi"]), "m"),
        (f"STA {b} (Akhir Tikungan)", sta_str(r["sta_akhir"]), "m"),
    ]
    
    st.dataframe(pd.DataFrame(rows, columns=["Parameter Geometrik", "Nilai Calculated", "Satuan / Keterangan"]), hide_index=True, use_container_width=True)

    excel_bytes = generate_excel_bm2021(r, P_global)
    st.download_button(
        "📊 Unduh Laporan Excel Ber-Formula Otomatis (.XLSX)",
        excel_bytes, file_name="Perhitungan_Geometri_Jalan_BM2021.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
