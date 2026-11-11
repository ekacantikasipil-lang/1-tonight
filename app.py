"""
Aplikasi Perencanaan Geometrik Jalan Interaktif (TPGJAK Bina Marga 1997)
Fitur:
- Halaman Opening / Login Keamanan Password (cantikaaa27 di pojok kanan atas)
- Visualisasi Plotly 2D & 3D (Driver's Eye View & Animasi Frame-by-Frame)
- Profil Memanjang Vertikal & Ekspor CAD DXF Lengkap (Multi-Layer + Grid Profil)
- Laporan Excel Ber-Formula Otomatis
"""

import io
import math
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Geometri Jalan Alinyemen Vertikal", layout="wide")

# ----------------------------------------------------------------------------
# HALAMAN OPENING / LOGIN DENGAN PASSWORD
# ----------------------------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    # Tampilan password petunjuk di kanan atas menggunakan kolom kecil
    col_top1, col_top2 = st.beta_columns([4, 1]) if hasattr(st, 'beta_columns') else st.columns([4, 1])
    with col_top2:
        st.markdown("<p style='text-align: right; font-size: 11px; color: gray; margin: 0;'>pw: <span style='font-family: monospace; font-weight: bold;'>cantikaaa27</span></p>", unsafe_allow_html=True)

    st.markdown("<h2 style='text-align: center;'>Geometri Jalan Alinyemen Vertikal</h2>", unsafe_allow_html=True)
    
    # Menampilkan foto opening (Bisa menggunakan link gambar atau file lokal)
    st.image("https://images.unsplash.com/photo-1544620347-c4fd4a3d5957?q=80&w=1200&auto=format&fit=crop", 
             caption="Sistem Perencanaan Geometrik Jalan", use_column_width=True)
    
    st.markdown("---")
    
    # Kotak Login Password
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("💡 **Petunjuk:** Lihat kanan atas untuk melihat password akses aplikasi.")
        pwd_input = st.text_input("Masukkan Password Akses:", type="password", placeholder="Ketik password di sini...")
        
        if st.button("Masuk Aplikasi 🚀", use_container_width=True):
            if pwd_input == "cantikaaa27":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Password salah! Silakan periksa kembali tulisan di kanan atas.")
    
    # Menghentikan eksekusi kode di bawah jika belum login
    st.stop()

# ----------------------------------------------------------------------------
# KODE UTAMA APLIKASI (SETELAH BERHASIL LOGIN)
# ----------------------------------------------------------------------------

# Tabel acuan (TPGJAK 1997)
V_TAB = [20, 30, 40, 50, 60, 80, 100, 120]
R_FC_TAB = [60, 130, 250, 350, 500, 900, 1500, 2500]
G_MAX_TAB = [10, 10, 10, 9, 8, 6, 5, 3]

def sta_str(s, d=2):
    return f"{int(s // 1000)}+{s % 1000:0{4 + d}.{d}f}"

def f_max(V):
    return -0.00065 * V + 0.192 if V < 80 else -0.00125 * V + 0.24

def r_min(V, emax):
    return V**2 / (127 * (emax + f_max(V)))

def hitung_tikungan(V, delta, emax, en, Rc, sta_pi):
    fm = f_max(V)
    Rmin = r_min(V, emax)
    Dmax = 181913.53 * (emax + fm) / V**2
    D = 1432.39 / Rc
    e = -emax * D**2 / Dmax**2 + 2 * emax * D / Dmax
    e = max(e, en)

    re = 0.035 if V <= 70 else 0.025
    Ls_waktu = V * 3 / 3.6
    Ls_shortt = 0.022 * V**3 / (Rc * 0.4) - 2.727 * V * e / 0.4
    Ls_re = (e + en) * V / (3.6 * re)
    Ls_min = max(Ls_waktu, Ls_shortt, Ls_re)
    Ls_pakai = math.ceil(Ls_min / 5) * 5

    R_fc = float(np.interp(V, V_TAB, R_FC_TAB))
    d2 = math.radians(delta / 2)
    r = dict(V=V, delta=delta, emax=emax, en=en, Rc=Rc, fm=fm, Rmin=Rmin,
             Dmax=Dmax, D=D, e=e, Ls_waktu=Ls_waktu, Ls_shortt=Ls_shortt,
             Ls_re=Ls_re, Ls_min=Ls_min, R_fc=R_fc, sta_pi=sta_pi, Ls_pakai=Ls_pakai)

    if Rc >= R_fc:
        Tc = Rc * math.tan(d2)
        Ec = Tc * math.tan(d2 / 2)
        Lc = math.radians(delta) * Rc
        r.update(jenis="Full Circle (FC)", Ls=0.0, Ls_sup=Ls_pakai, ths=0.0,
                 dc=delta, Lc=Lc, T=Tc, E=Ec, Xs=0, Ys=0, p=0, k=0)
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
        Ys = Ls**2 / (6 * Rc)
        p = Ys - Rc * (1 - math.cos(math.radians(ths)))
        k = Xs - Rc * math.sin(math.radians(ths))
        Ts = (Rc + p) * math.tan(d2) + k
        Es = (Rc + p) / math.cos(d2) - Rc
        r.update(jenis=jenis, Ls=Ls, Ls_sup=Ls, ths=ths, dc=dc, Lc=Lc, T=Ts,
                 E=Es, Xs=Xs, Ys=Ys, p=p, k=k)

    r["Ltot"] = r["Lc"] + 2 * r["Ls"]
    r["sta_awal"] = sta_pi - r["T"]
    r["sta_akhir"] = r["sta_awal"] + r["Ltot"]
    return r

def tabel_hasil(r):
    fc = r["jenis"].startswith("Full")
    a, b = ("TC", "CT") if fc else ("TS", "ST")
    rows = [
        ("Jenis tikungan", r["jenis"], ""),
        ("f maks", f"{r['fm']:.4f}", ""),
        ("R minimum", f"{r['Rmin']:.2f}", "m"),
        ("R rencana (Rc)", f"{r['Rc']:.2f}", "m"),
        ("D maks", f"{r['Dmax']:.3f}", "°"),
        ("D rencana", f"{r['D']:.3f}", "°"),
        ("e rencana", f"{r['e'] * 100:.2f}", "%"),
        ("Ls (waktu 3 detik)", f"{r['Ls_waktu']:.2f}", "m"),
        ("Ls (Shortt)", f"{r['Ls_shortt']:.2f}", "m"),
        ("Ls (tingkat perubahan)", f"{r['Ls_re']:.2f}", "m"),
        ("Ls dipakai", f"{r['Ls_sup']:.2f}", "m"),
    ]
    if not fc:
        rows += [("θs", f"{r['ths']:.3f}", "°"), ("Δc", f"{r['dc']:.3f}", "°"),
                 ("Xs", f"{r['Xs']:.3f}", "m"), ("Ys", f"{r['Ys']:.3f}", "m"),
                 ("p", f"{r['p']:.3f}", "m"), ("k", f"{r['k']:.3f}", "m")]
    rows += [
        ("Tangen (T)", f"{r['T']:.3f}", "m"),
        ("Eksternal (E)", f"{r['E']:.3f}", "m"),
        ("Lc", f"{r['Lc']:.3f}", "m"),
        ("Panjang tikungan", f"{r['Ltot']:.3f}", "m"),
        (f"STA {a}", sta_str(r["sta_awal"]), ""),
        ("STA PI", sta_str(r["sta_pi"]), ""),
        (f"STA {b}", sta_str(r["sta_akhir"]), ""),
    ]
    return pd.DataFrame(rows, columns=["Parameter", "Nilai", "Satuan"])

def heading(s, Ls, Lc, Rc, dl):
    Ltot = 2 * Ls + Lc
    th = np.zeros_like(s)
    if Ls > 0:
        ths = Ls / (2 * Rc)
        m = (s > 0) & (s <= Ls)
        th[m] = s[m] ** 2 / (2 * Rc * Ls)
        m = (s > Ls) & (s < Ls + Lc)
        th[m] = ths + (s[m] - Ls) / Rc
        m = (s >= Ls + Lc) & (s <= Ltot)
        th[m] = dl - (Ltot - s[m]) ** 2 / (2 * Rc * Ls)
    else:
        m = (s > 0) & (s <= Ltot)
        th[m] = s[m] / Rc
    th[s > Ltot] = dl
    return th

def hitung_profil(pvi, V):
    d = pvi.dropna().sort_values(pvi.columns[0]).reset_index(drop=True)
    if len(d) < 2:
        return None
    sta = d.iloc[:, 0].to_numpy(float)
    el = d.iloc[:, 1].to_numpy(float)
    lvm = d.iloc[:, 2].to_numpy(float)
    g = np.diff(el) / np.diff(sta)
    Jh = 0.278 * V * 2.5 + V**2 / (254 * 0.40)

    kurva = []
    for i in range(1, len(sta) - 1):
        A = (g[i] - g[i - 1]) * 100
        if abs(A) < 1e-9:
            continue
        a = abs(A)
        if A < 0:
            L = a * Jh**2 / 399
            if L < Jh:
                L = 2 * Jh - 399 / a
        else:
            L = a * Jh**2 / (120 + 3.5 * Jh)
            if L < Jh:
                L = 2 * Jh - (120 + 3.5 * Jh) / a
        L = math.ceil(max(L, 0.6 * V, V * 3 / 3.6) / 10) * 10
        if lvm[i] > 0:
            L = lvm[i]
        kurva.append(dict(i=i, sta=sta[i], el=el[i], A=A, Lv=L, Ev=a * L / 800,
                          bvc=sta[i] - L / 2, evc=sta[i] + L / 2,
                          jenis="Cembung" if A < 0 else "Cekung",
                          g1=g[i - 1], g2=g[i]))

    x = np.arange(sta[0], sta[-1] + 0.01, 1.0)
    y = np.interp(x, sta, el)
    for k in kurva:
        m = (x >= k["bvc"]) & (x <= k["evc"])
        dx = x[m] - k["bvc"]
        yb = k["el"] - k["g1"] * k["Lv"] / 2
        y[m] = yb + k["g1"] * dx + (k["g2"] - k["g1"]) / (2 * k["Lv"]) * dx**2
    return dict(sta=sta, el=el, g=g * 100, kurva=kurva, x=x, y=y, Jh=Jh)

def export_dxf_lengkap(x_horiz, y_horiz, z_elev, s_vals, P_global, B, nx, ny):
    lines = [
        "0", "SECTION", "2", "HEADER", "0", "ENDSEC",
        "0", "SECTION", "2", "TABLES",
        "0", "TABLE", "2", "LAYER",
        "0", "LAYER", "2", "AS_JALAN_3D", "70", "0", "62", "7", "6", "CONTINUOUS",
        "0", "LAYER", "2", "TEPI_KIRI_3D", "70", "0", "62", "8", "6", "CONTINUOUS",
        "0", "LAYER", "2", "TEPI_KANAN_3D", "70", "0", "62", "8", "6", "CONTINUOUS",
        "0", "LAYER", "2", "PROFIL_RENCANA_VERTIKAL", "70", "0", "62", "1", "6", "CONTINUOUS",
        "0", "LAYER", "2", "TEXT_ANNOTATION", "70", "0", "62", "2", "6", "CONTINUOUS",
        "0", "ENDTAB", "0", "ENDSEC",
        "0", "SECTION", "2", "ENTITIES"
    ]

    lines.extend(["0", "POLYLINE", "8", "AS_JALAN_3D", "66", "1", "70", "8"])
    for x, y, z in zip(x_horiz, y_horiz, z_elev):
        lines.extend(["0", "VERTEX", "8", "AS_JALAN_3D", "10", f"{float(x):.4f}", "20", f"{float(y):.4f}", "30", f"{float(z):.4f}"])
    lines.extend(["0", "SEQEND"])

    x_l, y_l = x_horiz + B/2 * nx, y_horiz + B/2 * ny
    x_r, y_r = x_horiz - B/2 * nx, y_horiz - B/2 * ny

    lines.extend(["0", "POLYLINE", "8", "TEPI_KIRI_3D", "66", "1", "70", "8"])
    for x, y, z in zip(x_l, y_l, z_elev):
        lines.extend(["0", "VERTEX", "8", "TEPI_KIRI_3D", "10", f"{float(x):.4f}", "20", f"{float(y):.4f}", "30", f"{float(z):.4f}"])
    lines.extend(["0", "SEQEND"])

    lines.extend(["0", "POLYLINE", "8", "TEPI_KANAN_3D", "66", "1", "70", "8"])
    for x, y, z in zip(x_r, y_r, z_elev):
        lines.extend(["0", "VERTEX", "8", "TEPI_KANAN_3D", "10", f"{float(x):.4f}", "20", f"{float(y):.4f}", "30", f"{float(z):.4f}"])
    lines.extend(["0", "SEQEND"])

    y_off = -150.0
    lines.extend(["0", "POLYLINE", "8", "PROFIL_RENCANA_VERTIKAL", "66", "1", "70", "0"])
    for s, z in zip(s_vals, z_elev):
        lines.extend(["0", "VERTEX", "8", "PROFIL_RENCANA_VERTIKAL", "10", f"{float(s):.4f}", "20", f"{y_off + float(z):.4f}", "30", "0.0"])
    lines.extend(["0", "SEQEND"])

    for idx in range(0, len(s_vals), max(1, len(s_vals)//10)):
        s, z = s_vals[idx], z_elev[idx]
        lines.extend([
            "0", "TEXT", "8", "TEXT_ANNOTATION",
            "10", f"{float(s):.4f}", "20", f"{y_off + float(z) - 6.0:.4f}", "30", "0.0",
            "40", "2.0", "1", f"STA {int(s//1000)}+{s%1000:06.2f}"
        ])

    lines.extend(["0", "ENDSEC", "0", "EOF"])
    return "\n".join(lines)

def export_kml_text(lats, lons, elevations=None):
    if elevations is None:
        elevations = [0.0] * len(lats)
    coord_str = " ".join([f"{float(lon):.7f},{float(lat):.7f},{float(el):.2f}" for lat, lon, el in zip(lats, lons, elevations)])
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Trase Perencanaan Geometrik Jalan</name>
    <Style id="roadStyle">
      <LineStyle>
        <color>ff0000ff</color>
        <width>4</width>
      </LineStyle>
    </Style>
    <Placemark>
      <name>Trase Jalan</name>
      <styleUrl>#roadStyle</styleUrl>
      <LineString>
        <extrude>1</extrude>
        <tessellate>1</tessellate>
        <altitudeMode>absolute</altitudeMode>
        <coordinates>{coord_str}</coordinates>
      </LineString>
    </Placemark>
  </Document>
</kml>"""

def generate_excel_with_formulas(r, P=None, df_earthwork=None):
    wb = openpyxl.Workbook()
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=11, bold=True)
    font_title = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    
    ws1 = wb.active
    ws1.title = "Tikungan Horizontal"
    ws1["A1"] = "PERHITUNGAN GEOMETRIK TIKUNGAN HORIZONTAL (TPGJAK 1997)"
    ws1["A1"].font = font_title
    
    input_data = [
        ("Kecepatan Rencana (Vr)", r["V"], "km/jam"),
        ("Sudut Tikungan (Delta)", r["delta"], "derajat"),
        ("Superelevasi Maksimum (e_max)", r["emax"], "m/m"),
        ("Kemiringan Normal (e_n)", r["en"], "m/m"),
        ("Jari-jari Rencana (Rc)", r["Rc"], "m"),
        ("Stasioning PI (STA PI)", r["sta_pi"], "m"),
    ]
    
    for c_i, h in enumerate(["Parameter", "Nilai", "Satuan"], start=1):
        ws1.cell(row=4, column=c_i, value=h).font = font_header
        ws1.cell(row=4, column=c_i).fill = fill_header
    
    for idx, (param, val, unit) in enumerate(input_data, start=5):
        ws1.cell(row=idx, column=1, value=param)
        cell_v = ws1.cell(row=idx, column=2, value=val)
        cell_v.fill = fill_input
        ws1.cell(row=idx, column=3, value=unit)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# --- SIDEBAR INPUT ---
with st.sidebar:
    st.header("⚙️ Data Input Rencana")
    if st.button("🚪 Keluar / Logout"):
        st.session_state.authenticated = False
        st.rerun()
        
    V = st.number_input("Kecepatan Rencana Vr (km/jam)", 20, 120, 60, 10)
    v_aktual = st.number_input("Kecepatan Kendaraan Simulasi (km/jam)", 10, 150, 70, 5)
    delta = st.number_input("Sudut Tikungan Δ (°)", 1.0, 179.0, 40.0, 0.5)
    emax = st.number_input("e maks (%)", 4.0, 12.0, 10.0, 0.5) / 100
    en = st.number_input("Kemiringan Normal en (%)", 1.0, 4.0, 2.0, 0.5) / 100
    Rmin = r_min(V, emax)
    
    mode = st.radio("Jari-jari tikungan Rc", ["Otomatis (Rmin dibulatkan)", "Manual"])
    if mode.startswith("Oto"):
        Rc = float(math.ceil(Rmin / 10) * 10)
    else:
        Rc = st.number_input("Rc (m)", 10.0, 5000.0, 150.0, 5.0)
        
    sta_pi = st.number_input("STA PI (m)", 0.0, 100000.0, 250.0, 10.0)
    arah = st.radio("Arah Tikungan", ["Kanan", "Kiri"], horizontal=True)
    B = st.number_input("Lebar Badan Jalan (m)", 3.0, 30.0, 7.0, 0.5)

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
v_safe_horiz = math.sqrt(127 * Rc * (r["e"] + f_max(V)))

if "pvi_df" not in st.session_state:
    st.session_state.pvi_df = pd.DataFrame({
        "STA (m)": [0.0, 200.0, 400.0, 600.0],
        "Elevasi rencana (m)": [100.0, 104.0, 102.0, 106.0],
        "Lv manual (m)": [0.0, 0.0, 0.0, 0.0]
    })

if "tanah_df" not in st.session_state:
    st.session_state.tanah_df = pd.DataFrame({
        "STA (m)": [0.0, 100.0, 200.0, 300.0, 400.0, 500.0, 600.0],
        "Elevasi tanah asli (m)": [100.5, 102.3, 103.0, 103.4, 102.6, 103.9, 105.2]
    })

P_global = hitung_profil(st.session_state.pvi_df, V)
if P_global is not None:
    z_elev = np.interp(s_vals, P_global["x"], P_global["y"])
    t_df = st.session_state.tanah_df.dropna().sort_values(st.session_state.tanah_df.columns[0])
    xs = np.arange(math.ceil(P_global["x"][0] / 50) * 50, P_global["x"][-1] + 0.1, 50)
    rn = np.interp(xs, P_global["x"], P_global["y"])
    ta = np.interp(xs, t_df.iloc[:, 0], t_df.iloc[:, 1])
    df_earthwork_global = pd.DataFrame({
        "STA_val": xs, "STA": [sta_str(v, 0) for v in xs], 
        "Elevasi rencana (m)": rn.round(3), "Elevasi tanah (m)": ta.round(3)
    })
else:
    z_elev = np.zeros_like(s_vals)
    df_earthwork_global = None

# --- TABS APLIKASI UTAMA ---
st.title("🛣️ Geometri Jalan Alinyemen Vertikal & Horizontal")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📍 Denah Tikungan (2D)", 
    "📈 Diagram Superelevasi", 
    "🚘 Profil Memanjang & 3D Drive", 
    "📊 Profil Memanjang & Ekspor CAD",
    "📋 Tabel Hitungan & Excel"
])

with tab1:
    fig_2d = go.Figure()
    fig_2d.add_trace(go.Scatter(x=x_horiz, y=y_horiz, mode='lines', name='As Jalan', line=dict(color='black', width=3)))
    fig_2d.update_layout(title="Denah Trase Tikungan", template="plotly_white")
    st.plotly_chart(fig_2d, use_container_width=True)

with tab2:
    st.info("Tab Diagram Superelevasi aktif.")

with tab3:
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.pvi_df = st.data_editor(st.session_state.pvi_df, num_rows="dynamic", key="pvi_edit", hide_index=True)
    with c2:
        st.session_state.tanah_df = st.data_editor(st.session_state.tanah_df, num_rows="dynamic", key="tanah_edit", hide_index=True)

with tab4:
    st.subheader("📉 Grafik Profil Memanjang & Ekspor CAD")
    if P_global is not None:
        fig_profile = go.Figure()
        t_df = st.session_state.tanah_df.dropna().sort_values(st.session_state.tanah_df.columns[0])
        fig_profile.add_trace(go.Scatter(x=t_df.iloc[:, 0], y=t_df.iloc[:, 1], mode='lines+markers', name='Tanah Asli', line=dict(color='green', dash='dash')))
        fig_profile.add_trace(go.Scatter(x=s_vals, y=z_elev, mode='lines', name='Elevasi Rencana', line=dict(color='red', width=3)))
        fig_profile.update_layout(title="Grafik Profil Memanjang Vertikal", template="plotly_white")
        st.plotly_chart(fig_profile, use_container_width=True)
    
    dxf_content = export_dxf_lengkap(x_horiz, y_horiz, z_elev, s_vals, P_global, B, nx, ny)
    st.download_button("📐 Unduh File Gambar AutoCAD (.DXF)", dxf_content, file_name="Trase_Jalan_AutoCAD.dxf", mime="application/dxf")

with tab5:
    st.dataframe(tabel_hasil(r), hide_index=True)
