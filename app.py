import io
import math
import time
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

# Konfigurasi Halaman & Tema Cerah
st.set_page_config(
    page_title="Geometri Jalan Interaktif - Bina Marga",
    layout="wide",
    page_icon="🛣️"
)

# Custom CSS untuk Tampilan Cerah, Bersih, dan Modern
st.markdown("""
<style>
    .stApp {
        background-color: #F8F9FA;
        color: #212529;
    }
    div[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E9ECEF;
    }
    .status-card {
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 15px;
        font-weight: bold;
    }
    .status-oke {
        background-color: #D4EDDA;
        color: #155724;
        border: 1px solid #C3E6CB;
    }
    .status-sedang {
        background-color: #FFF3CD;
        color: #856404;
        border: 1px solid #FFEEBA;
    }
    .status-bahaya {
        background-color: #F8D7DA;
        color: #721C24;
        border: 1px solid #F5C6CB;
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
            "<p style='text-align: right; font-size: 12px; color: #6c757d; margin: 0;'>pw: "
            "<span style='font-family: monospace; font-weight: bold; color: #0d6efd;'>cantikaaa27</span></p>",
            unsafe_allow_html=True
        )

    st.markdown("<h1 style='text-align: center; color: #0f5132;'>🛣️ Perencanaan Geometrik Jalan Interaktif</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #6c757d;'>Standar TPGJAK 1997 / RAE 2021</h4>", unsafe_allow_html=True)
    
    st.image(
        "https://images.unsplash.com/photo-1544620347-c4fd4a3d5957?q=80&w=1200&auto=format&fit=crop",
        caption="Aplikasi Perencanaan Geometrik Jalan Interaktif & Visualisasi 3D",
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
# 3. SIDEBAR INPUT & SYSTEM CALCULATIONS
# ----------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Parameter Rencana Jalan")
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
ext = math.ceil(max(100, 1.3 * r["T"]) / 10) * 10

s_vals = np.unique(np.concatenate([np.arange(-ext, Ltot + ext, 1.0), [0, Ls, Ls + Lc, Ltot]]))
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

# Session State untuk Animasi Mobil Jalan
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
    font_title = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    ws1 = wb.active
    ws1.title = "Tikungan Horisontal"
    ws1["A1"] = "REKAPITULASI CALCULASI GEOMETRIK JALAN (BINA MARGA 2021)"
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
    for x, y, z in zip(xr, yr, z_e):
        lines.extend(["0", "VERTEX", "8", "TEPI_KANAN_3D", "10", f"{float(x):.3f}", "20", f"{float(y):.3f}", "30", f"{float(z):.3f}"])
    lines.extend(["0", "SEQEND"])

    y_off = -150.0
    lines.extend(["0", "POLYLINE", "8", "PROFIL_VERTIKAL", "66", "1", "70", "0"])
    for s, z in zip(s_v, z_e):
        lines.extend(["0", "VERTEX", "8", "PROFIL_VERTIKAL", "10", f"{float(s):.3f}", "20", f"{y_off + float(z):.3f}", "30", "0.0"])
    lines.extend(["0", "SEQEND"])

    lines.extend(["0", "ENDSEC", "0", "EOF"])
    return "\n".join(lines)

# ----------------------------------------------------------------------------
# 5. TAMPILAN APLIKASI UTAMA (TABS TEMA CERAH)
# ----------------------------------------------------------------------------
st.title("🛣️ Geometri Jalan Interaktif & Visualisasi 3D")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📍 Denah Tikungan (2D)", 
    "📈 Diagram Superelevasi", 
    "🚘 Simulasi 3D Mobil Jalan", 
    "📊 Profil Memanjang & Ekspor CAD",
    "📋 Tabel Hitungan & Excel"
])

# ----------------------------------------------------------------------------
# TAB 1: DENAH TIKUNGAN 2D
# ----------------------------------------------------------------------------
with tab1:
    st.subheader("📍 Visualisasi Denah Trase Tikungan 2D")
    
    fig_2d = go.Figure()
    fig_2d.add_trace(go.Scatter(x=x_horiz, y=y_horiz, mode='lines', name='As Jalan', line=dict(color='#0d6efd', width=3)))
    
    xl, yl = x_horiz + (B/2) * nx, y_horiz + (B/2) * ny
    xr, yr = x_horiz - (B/2) * nx, y_horiz - (B/2) * ny
    fig_2d.add_trace(go.Scatter(x=xl, y=yl, mode='lines', name='Tepi Kiri', line=dict(color='#6c757d', dash='dash')))
    fig_2d.add_trace(go.Scatter(x=xr, y=yr, mode='lines', name='Tepi Kanan', line=dict(color='#6c757d', dash='dash')))
    
    fc = r["jenis"].startswith("Full")
    t1_lbl, t2_lbl = ("TC", "CT") if fc else ("TS", "ST")
    
    idx_start = np.argmin(np.abs(s_vals - 0))
    idx_end = np.argmin(np.abs(s_vals - r["Ltot"]))
    
    pts_x = [x_horiz[idx_start], x_horiz[idx_end]]
    pts_y = [y_horiz[idx_start], y_horiz[idx_end]]
    pts_txt = [f"<b>STA {t1_lbl}</b> ({sta_str(r['sta_awal'])})", f"<b>STA {t2_lbl}</b> ({sta_str(r['sta_akhir'])})"]
    
    if not fc:
        idx_sc = np.argmin(np.abs(s_vals - r["Ls"]))
        idx_cs = np.argmin(np.abs(s_vals - (r["Ls"] + r["Lc"])))
        pts_x.extend([x_horiz[idx_sc], x_horiz[idx_cs]])
        pts_y.extend([y_horiz[idx_sc], y_horiz[idx_cs]])
        pts_txt.extend([f"<b>STA SC</b> ({sta_str(r['sta_awal']+r['Ls'])})", f"<b>STA CS</b> ({sta_str(r['sta_awal']+r['Ls']+r['Lc'])})"])

    fig_2d.add_trace(go.Scatter(
        x=pts_x, y=pts_y, mode='markers+text', name='Titik Utama Tikungan',
        marker=dict(size=10, color='#dc3545', symbol='diamond'),
        text=pts_txt, textposition="top center"
    ))

    fig_2d.update_layout(
        title=f"Denah Trase Tikungan ({r['jenis']}) - R_c = {r['Rc']} m, Vr = {r['V']} km/jam",
        xaxis_title="Koordinat X (m)", yaxis_title="Koordinat Y (m)",
        template="plotly_white", showlegend=True,
        yaxis=dict(scaleanchor="x", scaleratio=1),
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF"
    )
    st.plotly_chart(fig_2d, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 2: DIAGRAM SUPERELEVASI
# ----------------------------------------------------------------------------
with tab2:
    st.subheader("📈 Diagram Superelevasi Kontinyu Jalan")
    
    e_desain = r["e"] * 100
    en_val = r["en"] * 100
    ls_val = r["Ls_sup"]
    
    s_sup = np.linspace(-50, r["Ltot"] + 50, 200)
    e_left = np.zeros_like(s_sup)
    e_right = np.zeros_like(s_sup)
    
    for i, s_i in enumerate(s_sup):
        if s_i < 0:
            e_left[i] = -en_val
            e_right[i] = -en_val
        elif 0 <= s_i <= ls_val:
            frac = s_i / ls_val
            e_left[i] = -en_val + frac * (e_desain + en_val)
            e_right[i] = -en_val - frac * (e_desain - en_val)
        elif ls_val < s_i < r["Ltot"] - ls_val:
            e_left[i] = e_desain
            e_right[i] = -e_desain
        elif r["Ltot"] - ls_val <= s_i <= r["Ltot"]:
            frac = (r["Ltot"] - s_i) / ls_val
            e_left[i] = -en_val + frac * (e_desain + en_val)
            e_right[i] = -en_val - frac * (e_desain - en_val)
        else:
            e_left[i] = -en_val
            e_right[i] = -en_val
            
    fig_sup = go.Figure()
    fig_sup.add_trace(go.Scatter(x=s_sup, y=e_left, mode='lines', name='Tepi Kiri Jalan', line=dict(color='#0d6efd', width=2.5)))
    fig_sup.add_trace(go.Scatter(x=s_sup, y=e_right, mode='lines', name='Tepi Kanan Jalan', line=dict(color='#dc3545', width=2.5)))
    fig_sup.add_trace(go.Scatter(x=s_sup, y=np.zeros_like(s_sup), mode='lines', name='As Jalan (0%)', line=dict(color='#212529', dash='dash')))
    
    fig_sup.update_layout(
        title="Diagram Transisi Superelevasi Perkerasan Jalan",
        xaxis_title="Stasioning STA (m)", yaxis_title="Kemiringan Elevasi / Superelevasi (%)",
        template="plotly_white", showlegend=True,
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF"
    )
    st.plotly_chart(fig_sup, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 3: PROFIL MEMANJANG & SIMULASI 3D DRIVER'S EYE VIEW (ANIMASI)
# ----------------------------------------------------------------------------
with tab3:
    st.subheader("🚘 Visualisasi 3D Jalan & Simulator Mobil Jalan Otomatis")
    
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.markdown("**Edit Data PVI Alinyemen Vertikal:**")
        st.session_state.pvi_df = st.data_editor(st.session_state.pvi_df, num_rows="dynamic", key="pvi_ed", hide_index=True)
    with col_e2:
        st.markdown("**Edit Data Elevasi Tanah Asli:**")
        st.session_state.tanah_df = st.data_editor(st.session_state.tanah_df, num_rows="dynamic", key="tnh_ed", hide_index=True)
        
    st.markdown("---")
    
    # Indikator Keterangan Warna Mobil
    v_safe = r["v_safe"]
    if v_sim <= v_safe:
        car_color = "#198754"  # Hijau Oke
        status_box = f"""
        <div class="status-card status-oke">
            🟢 <b>KONDISI HIJAU (OKE):</b> Mobil Melaju Aman!<br>
            Kecepatan simulasi ({v_sim} km/jam) masih di bawah batas aman tikungan ({v_safe:.1f} km/jam).
        </div>
        """
    elif v_sim <= v_safe + 10:
        car_color = "#ffc107"  # Kuning Sedang
        status_box = f"""
        <div class="status-card status-sedang">
            🟡 <b>KONDISI KUNING (SEDANG / WASPADA):</b> Mobil Mendekati Batas Aman!<br>
            Kecepatan simulasi ({v_sim} km/jam) sedikit melampaui V_aman ({v_safe:.1f} km/jam). Pengemudi harus berhati-hati.
        </div>
        """
    else:
        car_color = "#dc3545"  # Merah Tidak Oke
        status_box = f"""
        <div class="status-card status-bahaya">
            🔴 <b>KONDISI MERAH (TIDAK OKE / BAHAYA):</b> Over-Speed!<br>
            Kecepatan simulasi ({v_sim} km/jam) jauh melebihi batas aman tikungan ({v_safe:.1f} km/jam). Berisiko terjadi kecelakaan/selip.
        </div>
        """
        
    st.markdown(status_box, unsafe_allow_html=True)

    # Kontrol Tombol Animasi Mobil Jalan
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

    # Slider Manual Posisi STA Mobil
    current_idx = st.session_state.anim_sta_idx
    selected_sta = st.slider(
        "🎮 Kontrol Posisi Mobil / STA (m):", 
        float(s_vals[0]), float(s_vals[-1]), 
        float(s_vals[current_idx]), 5.0
    )
    
    # Update index jika slider digeser manual
    if not st.session_state.anim_active:
        st.session_state.anim_sta_idx = np.argmin(np.abs(s_vals - selected_sta))
        current_idx = st.session_state.anim_sta_idx

    # Plotting Mesh 3D Jalan & Mobil
    xc, yc, zc = x_horiz[current_idx], y_horiz[current_idx], z_elev[current_idx]

    fig_3d = go.Figure()
    
    # As Jalan 3D
    fig_3d.add_trace(go.Scatter3d(
        x=x_horiz, y=y_horiz, z=z_elev, mode='lines', name='As Jalan 3D',
        line=dict(color='#212529', width=6)
    ))
    
    # Tepi Kiri & Kanan 3D
    fig_3d.add_trace(go.Scatter3d(x=xl, y=yl, z=z_elev, mode='lines', name='Tepi Kiri 3D', line=dict(color='#6c757d', width=3)))
    fig_3d.add_trace(go.Scatter3d(x=xr, y=yr, z=z_elev, mode='lines', name='Tepi Kanan 3D', line=dict(color='#6c757d', width=3)))
    
    # Marker Posisi Mobil Berjalan
    fig_3d.add_trace(go.Scatter3d(
        x=[xc], y=[yc], z=[zc + 1.2], mode='markers+text', name='Mobil Simulasi',
        marker=dict(size=12, color=car_color, symbol='diamond'),
        text=[f"🚘 Mobil (STA {s_vals[current_idx]:.0f}m)"], textposition="top center"
    ))

    # Setting Pandangan Kamera Pengemudi
    dx_cam = np.cos(ths_[current_idx]) * 20
    dy_cam = np.sin(ths_[current_idx]) * 20
    
    camera_eye = dict(
        x=(xc - dx_cam) / 100,
        y=(yc - dy_cam) / 100,
        z=(zc + 3.0) / 100
    )

    fig_3d.update_layout(
        title=f"Visualisasi 3D Trase Jalan & Simulator Mobil (STA {s_vals[current_idx]:.1f} m)",
        template="plotly_white",
        scene=dict(
            xaxis_title="X (m)", yaxis_title="Y (m)", zaxis_title="Elevasi Z (m)",
            aspectmode='data', camera_eye=camera_eye
        ),
        margin=dict(l=0, r=0, b=0, t=40),
        paper_bgcolor="#FFFFFF"
    )
    st.plotly_chart(fig_3d, use_container_width=True)

    # Loop Animasi Otomatis
    if st.session_state.anim_active:
        time.sleep(speed_delay)
        st.session_state.anim_sta_idx += 2
        if st.session_state.anim_sta_idx >= len(s_vals):
            st.session_state.anim_sta_idx = 0  # Loop kembali ke awal
        st.rerun()

# ----------------------------------------------------------------------------
# TAB 4: PROFIL MEMANJANG & EKSPOR CAD
# ----------------------------------------------------------------------------
with tab4:
    st.subheader("📊 Profil Memanjang Vertikal Jalan & Ekspor CAD (.DXF)")
    
    if P_global is not None:
        t_df = st.session_state.tanah_df.dropna().sort_values(st.session_state.tanah_df.columns[0])
        fig_prof = go.Figure()
        
        fig_prof.add_trace(go.Scatter(
            x=t_df.iloc[:, 0], y=t_df.iloc[:, 1], mode='lines+markers',
            name='Elevasi Tanah Asli', line=dict(color='#198754', dash='dash')
        ))
        fig_prof.add_trace(go.Scatter(
            x=s_vals, y=z_elev, mode='lines', name='Elevasi Rencana Jalan',
            line=dict(color='#dc3545', width=3)
        ))
        
        pvi_x = P_global["sta"]
        pvi_y = P_global["el"]
        fig_prof.add_trace(go.Scatter(
            x=pvi_x, y=pvi_y, mode='markers+text', name='Titik PVI',
            marker=dict(size=8, color='#0d6efd', symbol='x'),
            text=[f"PVI-{i+1}" for i in range(len(pvi_x))], textposition="top center"
        ))

        fig_prof.update_layout(
            title="Grafik Profil Memanjang Vertikal (Tanah Asli vs Rencana)",
            xaxis_title="Stasioning STA (m)", yaxis_title="Elevasi Z (m)",
            template="plotly_white", showlegend=True,
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF"
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
    st.subheader("📋 Tabel Hasil Perhitungan Tikungan & Laporan Excel")
    
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
