"""
Aplikasi Perencanaan Geometrik Jalan Interaktif (TPGJAK Bina Marga 1997)
Fitur:
- Visualisasi Plotly 2D & 3D (Driver's Eye View & Animasi Frame-by-Frame)
- Pewarnaan Dinamis Mobil (Hijau/Kuning/Merah berdasarkan Batas Kecepatan Safe Speed & Kelandaian)
- Import Data Survei (CSV / Excel Total Station & GPS)
- Ekspor DXF (CAD R12 Polyline 3D) & KML (Google Earth)
- Overlay Trase di Peta Satelit / OpenStreetMap
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

st.set_page_config(page_title="Perencanaan Geometrik Jalan 3D & GIS", layout="wide")

# ----------------------------------------------------------------------------
# Tabel acuan (TPGJAK 1997)
# ----------------------------------------------------------------------------
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

# ----------------------------------------------------------------------------
# EXPORT CAD & KML
# ----------------------------------------------------------------------------
def export_dxf_text(x_coords, y_coords, z_coords=None):
    lines = ["0", "SECTION", "2", "HEADER", "0", "ENDSEC", "0", "SECTION", "2", "ENTITIES"]
    if z_coords is None:
        z_coords = [0.0] * len(x_coords)
    
    lines.extend(["0", "POLYLINE", "8", "TRASE_JALAN", "66", "1", "70", "8"])
    for x, y, z in zip(x_coords, y_coords, z_coords):
        lines.extend(["0", "VERTEX", "8", "TRASE_JALAN", "10", f"{float(x):.4f}", "20", f"{float(y):.4f}", "30", f"{float(z):.4f}"])
    lines.extend(["0", "SEQEND", "0", "ENDSEC", "0", "EOF"])
    return "\n".join(lines)

def export_kml_text(lats, lons, elevations=None):
    if elevations is None:
        elevations = [0.0] * len(lats)
    coord_str = " ".join([f"{float(lon):.7f},{float(lat):.7f},{float(el):.2f}" for lat, lon, el in zip(lats, lons, elevations)])
    kml = f"""<?xml version="1.0" encoding="UTF-8"?>
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
        <coordinates>
          {coord_str}
        </coordinates>
      </LineString>
    </Placemark>
  </Document>
</kml>"""
    return kml

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
    ws1.views.sheetView[0].showGridLines = True
    ws1["A1"] = "PERHITUNGAN GEOMETRIK TIKUNGAN HORIZONTAL (TPGJAK 1997)"
    ws1["A1"].font = font_title
    
    ws1.cell(row=3, column=1, value="1. DATA INPUT RENCANA").font = font_sub
    input_data = [
        ("Kecepatan Rencana (Vr)", r["V"], "km/jam", "B4"),
        ("Sudut Tikungan (Delta)", r["delta"], "derajat", "B5"),
        ("Superelevasi Maksimum (e_max)", r["emax"], "m/m", "B6"),
        ("Kemiringan Normal (e_n)", r["en"], "m/m", "B7"),
        ("Jari-jari Rencana (Rc)", r["Rc"], "m", "B8"),
        ("Stasioning PI (STA PI)", r["sta_pi"], "m", "B9"),
    ]
    
    for c_i, h in enumerate(["Parameter", "Nilai", "Satuan"], start=1):
        ws1.cell(row=4, column=c_i, value=h).font = font_header
        ws1.cell(row=4, column=c_i).fill = fill_header
    
    for idx, (param, val, unit, cell_ref) in enumerate(input_data, start=5):
        ws1.cell(row=idx, column=1, value=param)
        cell_v = ws1.cell(row=idx, column=2, value=val)
        cell_v.fill = fill_input
        ws1.cell(row=idx, column=3, value=unit)
        if "m/m" in unit:
            cell_v.number_format = '0.00%'
            
    ws1.cell(row=12, column=1, value="2. HASIL PERHITUNGAN BER-FORMULA").font = font_sub
    headers_calc = ["Deskripsi Parameter", "Formula Excel / Nilai", "Satuan", "Keterangan / Persamaan"]
    for col_num, h in enumerate(headers_calc, start=1):
        c = ws1.cell(row=13, column=col_num, value=h)
        c.font = font_header
        c.fill = fill_header

    calc_rows = [
        ("Faktor Gesekan Maksimum (f_max)", "=IF(B5<80, -0.00065*B5+0.192, -0.00125*B5+0.24)", "-", "TPGJAK 1997"),
        ("Jari-Jari Minimum (R_min)", "=B5^2 / (127 * (B7 + B14))", "m", "Vr^2 / (127*(emax+fm))"),
        ("Derajat Lengkung Maksimum (D_max)", "=181913.53 * (B7 + B14) / (B5^2)", "deg", "181913.53*(emax+fm)/Vr^2"),
        ("Derajat Lengkung Rencana (D)", "=1432.39 / B9", "deg", "1432.39 / Rc"),
        ("Superelevasi Rencana (e)", "=MAX(B8, -B7*(B17^2)/(B16^2) + 2*B7*B17/B16)", "%", "Metode Bina Marga"),
        ("Ls (Waktu 3 detik)", "=B5 * 3 / 3.6", "m", "Vr * 3 / 3.6"),
        ("Ls (Shortt)", "=0.022 * (B5^3) / (B9 * 0.4) - 2.727 * B5 * B18 / 0.4", "m", "Formula Shortt"),
        ("Ls (Kelandaian Relatif)", "=(B18 + B8) * B5 / (3.6 * IF(B5<=70, 0.035, 0.025))", "m", "Modulus kelandaian relatif"),
        ("Ls Minimum", "=MAX(B19, B20, B21)", "m", "Max dari 3 kriteria Ls"),
        ("Ls Dipakai (Pembulatan 5m)", "=CEILING(B22, 5)", "m", "Pembulatan ke atas kelipatan 5"),
        ("Sudut Lengkung Peralihan (Theta_s)", "=DEGREES(B23 / (2 * B9))", "deg", "Ls / (2 * Rc) dalam derajat"),
        ("Sudut Lengkung Lingkaran (Delta_c)", "=B6 - (2 * B24)", "deg", "Delta - 2*Theta_s"),
        ("Panjang Busur Lingkaran (Lc)", "=RADIANS(B25) * B9", "m", "Delta_c * Rc (rad)"),
        ("Jenis Tikungan", '=IF(B9>=1500, "Full Circle (FC)", IF(B26>=25, "Spiral-Circle-Spiral (S-C-S)", "Spiral-Spiral (S-S)"))', "-", "Evaluasi Kriteria Trase"),
        ("Pergeseran Tangen (Xs)", "=B23 * (1 - (B23^2) / (40 * (B9^2)))", "m", "Ls*(1 - Ls^2/(40*Rc^2))"),
        ("Pergeseran Absis (Ys)", "=(B23^2) / (6 * B9)", "m", "Ls^2 / (6*Rc)"),
        ("Penyimpangan Ordinat (p)", "=B29 - B9 * (1 - COS(RADIANS(B24)))", "m", "Ys - Rc*(1 - cos(Theta_s))"),
        ("Penyimpangan Absis (k)", "=B28 - B9 * SIN(RADIANS(B24))", "m", "Xs - Rc*sin(Theta_s)"),
        ("Panjang Tangen (T)", '=IF(ISNUMBER(SEARCH("Full", B27)), B9*TAN(RADIANS(B6/2)), (B9+B30)*TAN(RADIANS(B6/2)) + B31)', "m", "Jarak Tangen Utama"),
        ("Eksternal (E)", '=IF(ISNUMBER(SEARCH("Full", B27)), B32*TAN(RADIANS(B6/4)), (B9+B30)/COS(RADIANS(B6/2)) - B9)', "m", "Jarak Eksternal Utama"),
        ("Panjang Total Lengkung (L_tot)", "=B26 + (2 * B23)", "m", "Lc + 2*Ls"),
        ("STA Awal Tikungan (TC / TS)", "=B10 - B32", "m", "STA PI - T"),
        ("STA Akhir Tikungan (CT / ST)", "=B35 + B34", "m", "STA Awal + Ltot"),
    ]

    for idx, (desc, formula, unit, ket) in enumerate(calc_rows, start=14):
        ws1.cell(row=idx, column=1, value=desc)
        c_val = ws1.cell(row=idx, column=2, value=formula)
        ws1.cell(row=idx, column=3, value=unit)
        ws1.cell(row=idx, column=4, value=ket)
        if unit == "%":
            c_val.number_format = '0.00%'
        elif unit in ["m", "deg"]:
            c_val.number_format = '#,##0.00'

    if P and "kurva" in P and len(P["kurva"]) > 0:
        ws2 = wb.create_sheet(title="Lengkung Vertikal")
        ws2.views.sheetView[0].showGridLines = True
        ws2["A1"] = "PERHITUNGAN LENGKUNG VERTIKAL"
        ws2["A1"].font = font_title

        headers_v = ["STA PPV", "Elevasi PPV (m)", "g1 (%)", "g2 (%)", "A (%)", "Jh (m)", "Lv Rencana (m)", "Ev (m)", "STA PLV", "STA PTV", "Jenis Lengkung"]
        for col_num, h in enumerate(headers_v, start=1):
            c = ws2.cell(row=3, column=col_num, value=h)
            c.font = font_header
            c.fill = fill_header

        for r_idx, k in enumerate(P["kurva"], start=4):
            ws2.cell(row=r_idx, column=1, value=k["sta"]).number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=2, value=k["el"]).number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=3, value=k["g1"]).number_format = '0.00%'
            ws2.cell(row=r_idx, column=4, value=k["g2"]).number_format = '0.00%'
            ws2.cell(row=r_idx, column=5, value=f"=ABS(D{r_idx}-C{r_idx})").number_format = '0.00%'
            ws2.cell(row=r_idx, column=6, value=P["Jh"]).number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=7, value=k["Lv"]).number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=8, value=f"=E{r_idx}*100*G{r_idx}/800").number_format = '#,##0.000'
            ws2.cell(row=r_idx, column=9, value=f"=A{r_idx}-(G{r_idx}/2)").number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=10, value=f"=A{r_idx}+(G{r_idx}/2)").number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=11, value=f'=IF(D{r_idx}<C{r_idx},"Cembung","Cekung")')

    if df_earthwork is not None and not df_earthwork.empty:
        ws3 = wb.create_sheet(title="Galian Timbunan")
        ws3.views.sheetView[0].showGridLines = True
        ws3["A1"] = "EVALUASI GALIAN DAN TIMBUNAN TIAP 50 METER"
        ws3["A1"].font = font_title

        headers_ew = ["No", "Stasioning (STA)", "Elevasi Rencana (m)", "Elevasi Tanah Asli (m)", "Selisih Elevasi (m)", "Status Pekerjaan"]
        for col_num, h in enumerate(headers_ew, start=1):
            c = ws3.cell(row=3, column=col_num, value=h)
            c.font = font_header
            c.fill = fill_header

        for idx, row in df_earthwork.reset_index(drop=True).iterrows():
            r_i = idx + 4
            ws3.cell(row=r_i, column=1, value=idx + 1)
            ws3.cell(row=r_i, column=2, value=row["STA_val"]).number_format = '#,##0.00'
            ws3.cell(row=r_i, column=3, value=row["Elevasi rencana (m)"]).number_format = '#,##0.000'
            ws3.cell(row=r_i, column=4, value=row["Elevasi tanah (m)"]).number_format = '#,##0.000'
            ws3.cell(row=r_i, column=5, value=f"=C{r_i}-D{r_i}").number_format = '#,##0.000'
            ws3.cell(row=r_i, column=6, value=f'=IF(E{r_i}>0, "Timbunan", "Galian")')

    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                if cell.value is not None:
                    cell.border = thin_border
        for col in sheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ----------------------------------------------------------------------------
# STREAMLIT APPLICATION INITIALIZATION
# ----------------------------------------------------------------------------
st.title("🛣️ Perencanaan Geometrik Jalan (Plotly 3D & GIS)")
st.caption("Aplikasi Interaktif: Visualisasi Plotly 2D/3D, Animasi Frame, Evaluasi Kecepatan Safe Speed, Import Survey Data & Ekspor DXF/KML")

# --- SIDEBAR INPUT ---
with st.sidebar:
    st.header("⚙️ Data Input Rencana")
    V = st.number_input("Kecepatan Rencana Vr (km/jam)", 20, 120, 60, 10)
    v_aktual = st.number_input("Kecepatan Kendaraan Simulasi (km/jam)", 10, 150, 70, 5)
    delta = st.number_input("Sudut Tikungan Δ (°)", 1.0, 179.0, 40.0, 0.5)
    emax = st.number_input("e maks (%)", 4.0, 12.0, 10.0, 0.5) / 100
    en = st.number_input("Kemiringan Normal en (%)", 1.0, 4.0, 2.0, 0.5) / 100
    Rmin = r_min(V, emax)
    st.caption(f"R minimum = {Rmin:.1f} m")
    
    mode = st.radio("Jari-jari tikungan Rc", ["Otomatis (Rmin dibulatkan)", "Manual"])
    if mode.startswith("Oto"):
        Rc = float(math.ceil(Rmin / 10) * 10)
        st.caption(f"Rc = {Rc:.0f} m")
    else:
        Rc = st.number_input("Rc (m)", 10.0, 5000.0, 150.0, 5.0)
        
    sta_pi = st.number_input("STA PI (m)", 0.0, 100000.0, 250.0, 10.0)
    arah = st.radio("Arah Tikungan", ["Kanan", "Kiri"], horizontal=True)
    B = st.number_input("Lebar Badan Jalan (m)", 3.0, 30.0, 7.0, 0.5)

# --- PERHITUNGAN GEOMETRI HORISONTAL GLOBAL ---
r = hitung_tikungan(V, delta, emax, en, Rc, sta_pi)
if Rc < Rmin:
    st.error(f"⚠️ Rc = {Rc:.1f} m lebih kecil dari R minimum ({Rmin:.1f} m). Perbesar Rc.")

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

# --- INITIALISATION DEFAULT SESSION STATE (Mencegah NameError/AttributeError) ---
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

# --- PERHITUNGAN PROFIL VERTIKAL GLOBAL ---
P_global = hitung_profil(st.session_state.pvi_df, V)
if P_global is not None:
    z_elev = np.interp(s_vals, P_global["x"], P_global["y"])
    t_df = st.session_state.tanah_df.dropna().sort_values(st.session_state.tanah_df.columns[0])
    xs = np.arange(math.ceil(P_global["x"][0] / 50) * 50, P_global["x"][-1] + 0.1, 50)
    rn = np.interp(xs, P_global["x"], P_global["y"])
    ta = np.interp(xs, t_df.iloc[:, 0], t_df.iloc[:, 1])
    df_earthwork_global = pd.DataFrame({
        "STA_val": xs,
        "STA": [sta_str(v, 0) for v in xs], "Elevasi rencana (m)": rn.round(3),
        "Elevasi tanah (m)": ta.round(3), "Selisih (m)": (rn - ta).round(3),
        "Keterangan": np.where(rn - ta > 0, "Timbunan", "Galian")
    })
else:
    z_elev = np.zeros_like(s_vals)
    df_earthwork_global = None

# --- TABS APP ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📍 Denah Tikungan (Plotly 2D)", 
    "📈 Diagram Superelevasi", 
    "🚘 Profil Memanjang & 3D Drive", 
    "🗺️ GIS & Import/Export Data",
    "📊 Tabel Hitungan & Excel"
])

# ----------------------------------------------------------------------------
# TAB 1: DENAH TIKUNGAN PLOTLY
# ----------------------------------------------------------------------------
with tab1:
    fig_2d = go.Figure()
    fig_2d.add_trace(go.Scatter(x=x_horiz, y=y_horiz, mode='lines', name='As Jalan', line=dict(color='black', width=3)))
    fig_2d.add_trace(go.Scatter(x=x_horiz + B/2 * nx, y=y_horiz + B/2 * ny, mode='lines', name='Tepi Kiri', line=dict(color='gray', width=1)))
    fig_2d.add_trace(go.Scatter(x=x_horiz - B/2 * nx, y=y_horiz - B/2 * ny, mode='lines', name='Tepi Kanan', line=dict(color='gray', width=1)))
    
    fc = r["jenis"].startswith("Full")
    keys = [("TC", 0), ("CT", Lc)] if fc else ([("TS", 0), ("SC=CS", Ls), ("ST", Ltot)] if Lc == 0 else [("TS", 0), ("SC", Ls), ("CS", Ls + Lc), ("ST", Ltot)])
    
    for nm, sk in keys:
        px_val = float(np.interp(sk, s_vals, x_horiz))
        py_val = float(np.interp(sk, s_vals, y_horiz))
        fig_2d.add_trace(go.Scatter(
            x=[px_val], y=[py_val], mode='markers+text', name=nm,
            marker=dict(size=10, color='red'),
            text=[f"{nm}<br>{sta_str(r['sta_awal'] + sk)}"],
            textposition="top center"
        ))

    fig_2d.update_layout(
        title=f"Denah Trase Tikungan Plotly - {r['jenis']} (Tikungan ke {arah.lower()})",
        xaxis_title="X (m)", yaxis_title="Y (m)",
        yaxis=dict(scaleanchor="x", scaleratio=1),
        hovermode="closest", template="plotly_white"
    )
    st.plotly_chart(fig_2d, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 2: DIAGRAM SUPERELEVASI
# ----------------------------------------------------------------------------
with tab2:
    e_pct, n_pct = r["e"] * 100, r["en"] * 100
    Ls_sup = r["Ls_sup"]
    if fc:
        a_sta, b_sta = r["sta_awal"] - 0.75 * Ls_sup, r["sta_awal"] + 0.25 * Ls_sup
        c_sta, d_sta = r["sta_akhir"] - 0.25 * Ls_sup, r["sta_akhir"] + 0.75 * Ls_sup
    else:
        a_sta, b_sta = r["sta_awal"], r["sta_awal"] + r["Ls"]
        c_sta, d_sta = b_sta + r["Lc"], b_sta + r["Lc"] + r["Ls"]
    
    pad = 0.3 * Ls_sup
    xr = (b_sta - a_sta) * 2 * n_pct / (e_pct + n_pct) if (e_pct + n_pct) != 0 else 0
    
    fig_sup = go.Figure()
    fig_sup.add_trace(go.Scatter(
        x=[a_sta - pad, a_sta, b_sta, c_sta, d_sta, d_sta + pad],
        y=[-n_pct, -n_pct, e_pct, e_pct, -n_pct, -n_pct],
        mode='lines+markers', name='Tepi Perkerasan Luar', line=dict(color='red', width=3)
    ))
    fig_sup.add_trace(go.Scatter(
        x=[a_sta - pad, a_sta + xr, b_sta, c_sta, d_sta - xr, d_sta + pad],
        y=[-n_pct, -n_pct, -e_pct, -e_pct, -n_pct, -n_pct],
        mode='lines+markers', name='Tepi Perkerasan Dalam', line=dict(color='blue', width=3)
    ))
    fig_sup.add_hline(y=0, line_dash="dash", line_color="black")
    fig_sup.update_layout(
        title="Diagram Superelevasi (Plotly Interactive)",
        xaxis_title="Stasioning (m)", yaxis_title="Kemiringan Melintang (%)",
        template="plotly_white"
    )
    st.plotly_chart(fig_sup, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 3: PROFIL MEMANJANG, PLOTLY 3D & DRIVER'S EYE VIEW
# ----------------------------------------------------------------------------
with tab3:
    st.markdown("### 🚘 Profil Memanjang & Tampilan 3D Driver's Eye View")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.pvi_df = st.data_editor(
            st.session_state.pvi_df,
            num_rows="dynamic", key="pvi_editor", hide_index=True)
    with c2:
        st.session_state.tanah_df = st.data_editor(
            st.session_state.tanah_df,
            num_rows="dynamic", key="tanah_editor", hide_index=True)

    if P_global is not None:
        gmax = float(np.interp(V, V_TAB, G_MAX_TAB))
        dz = np.gradient(z_elev, s_vals) * 100
        
        car_colors = []
        status_texts = []
        
        for i, s_i in enumerate(s_vals):
            is_in_curve = (s_i >= r["sta_awal"]) and (s_i <= r["sta_akhir"])
            slope_i = abs(dz[i])
            
            if (is_in_curve and v_aktual > v_safe_horiz) or (slope_i > gmax):
                color = "red"
                status = "🔴 BAHAYA: Kecepatan / Kelandaian melebihi batas aman!"
            elif (is_in_curve and v_aktual > 0.85 * v_safe_horiz) or (slope_i > 0.85 * gmax):
                color = "yellow"
                status = "🟡 PERINGATAN: Mendekati batas aman tikungan/kelandaian."
            else:
                color = "green"
                status = "🟢 AMAN: Kecepatan & Kelandaian sesuai standar."
                
            car_colors.append(color)
            status_texts.append(status)

        fig_3d = go.Figure()
        fig_3d.add_trace(go.Scatter3d(
            x=x_horiz, y=y_horiz, z=z_elev,
            mode='lines', line=dict(color='black', width=6), name='As Jalan 3D'
        ))
        fig_3d.add_trace(go.Scatter3d(
            x=x_horiz + B/2 * nx, y=y_horiz + B/2 * ny, z=z_elev,
            mode='lines', line=dict(color='gray', width=3), name='Tepi Kiri 3D'
        ))
        fig_3d.add_trace(go.Scatter3d(
            x=x_horiz - B/2 * nx, y=y_horiz - B/2 * ny, z=z_elev,
            mode='lines', line=dict(color='gray', width=3), name='Tepi Kanan 3D'
        ))
        fig_3d.add_trace(go.Scatter3d(
            x=[x_horiz[0]], y=[y_horiz[0]], z=[z_elev[0] + 1.2],
            mode='markers+text',
            marker=dict(size=12, color=car_colors[0], symbol='square'),
            name='Mobil Simulasi', text=["🚘 Mobil"], textposition="top center"
        ))

        frames_list = []
        for i in range(len(s_vals)):
            frames_list.append(go.Frame(
                data=[
                    go.Scatter3d(
                        x=[x_horiz[i]], y=[y_horiz[i]], z=[z_elev[i] + 1.2],
                        mode='markers+text',
                        marker=dict(size=14, color=car_colors[i]),
                        text=[f"🚘 Mobil<br>{status_texts[i]}"]
                    )
                ],
                name=f"frame_{i}"
            ))
        fig_3d.frames = frames_list

        fig_3d.update_layout(
            title="🎮 Simulasi 3D Pandangan Pengemudi (Driver's Eye View Engine)",
            scene=dict(
                xaxis_title="X (m)", yaxis_title="Y (m)", zaxis_title="Elevasi Z (m)",
                aspectmode='data'
            ),
            updatemenus=[{
                "buttons": [
                    {
                        "args": [None, {"frame": {"duration": 50, "redraw": True}, "fromcurrent": True}],
                        "label": "▶️ Jalankan Animasi 3D",
                        "method": "animate"
                    },
                    {
                        "args": [[None], {"frame": {"duration": 0, "redraw": False}, "mode": "immediate"}],
                        "label": "⏸️ Pause",
                        "method": "animate"
                    }
                ],
                "type": "buttons", "showactive": False, "x": 0.1, "y": 0
            }],
            template="plotly_white"
        )
        st.plotly_chart(fig_3d, use_container_width=True)
        st.info(f"💡 Kecepatan Aman Tikungan (Safe Speed Horizontal): **{v_safe_horiz:.1f} km/jam** | Kecepatan Simulasi: **{v_aktual} km/jam**")
    else:
        st.warning("Silakan lengkapi data PPV minimal 2 baris untuk menghasilkan Profil Memanjang.")

# ----------------------------------------------------------------------------
# TAB 4: GIS, IMPORT & EXPORT (CSV, EXCEL, DXF, KML, MAPS)
# ----------------------------------------------------------------------------
with tab4:
    st.subheader("📥 Import Data Survei & 📤 Ekspor CAD/KML/GIS")
    st.markdown("#### 1. Import Titik Survei (Total Station / GPS)")
    file_upload = st.file_uploader("Unggah File Data Survei (CSV atau Excel)", type=["csv", "xlsx"])
    if file_upload is not None:
        try:
            if file_upload.name.endswith(".csv"):
                df_survey = pd.read_csv(file_upload)
            else:
                df_survey = pd.read_excel(file_upload)
            st.success("Data survei berhasil di-import!")
            st.dataframe(df_survey.head(), use_container_width=True)
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    st.markdown("---")
    st.markdown("#### 2. Spatial Map Overlay & Ekspor Format Cad / Google Earth")
    col_gis1, col_gis2 = st.columns(2)
    
    with col_gis1:
        st.markdown("**Format Koordinat Awal (UTM / Lat-Long):**")
        lat_ref = st.number_input("Latitude Reference Pier/Start Point", -11.0, 6.0, -7.5000, format="%.6f")
        lon_ref = st.number_input("Longitude Reference Pier/Start Point", 95.0, 141.0, 110.5000, format="%.6f")
        
        lats_projected = lat_ref + (y_horiz / 111000.0)
        lons_projected = lon_ref + (x_horiz / (111000.0 * math.cos(math.radians(lat_ref))))

        fig_map = px.scatter(
            pd.DataFrame({'Latitude': lats_projected, 'Longitude': lons_projected, 'STA': [sta_str(s) for s in s_vals]}),
            x='Longitude', y='Latitude', hover_name='STA', title="Proyeksi Geografis Trase Jalan (Lat/Long)"
        )
        fig_map.update_traces(marker=dict(size=6, color='blue'), mode='lines+markers')
        fig_map.update_layout(template="plotly_white")
        st.plotly_chart(fig_map, use_container_width=True)

    with col_gis2:
        st.markdown("**Unduh File Desain Trase Jalan:**")
        dxf_content = export_dxf_text(x_horiz, y_horiz, z_elev)
        st.download_button(
            label="📐 Unduh File CAD (.DXF R12)",
            data=dxf_content,
            file_name="Trase_Geometrik_Jalan.dxf",
            mime="application/dxf"
        )
        
        kml_content = export_kml_text(lats_projected, lons_projected, z_elev)
        st.download_button(
            label="🌍 Unduh File Google Earth (.KML)",
            data=kml_content,
            file_name="Trase_Geometrik_Jalan.kml",
            mime="application/vnd.google-earth.kml+xml"
        )

# ----------------------------------------------------------------------------
# TAB 5: TABEL HITUNGAN & EXCEL FORMULA
# ----------------------------------------------------------------------------
with tab5:
    st.subheader("📋 Tabel Perhitungan & Laporan Ber-Formula")
    st.dataframe(tabel_hasil(r), hide_index=True, height=500)
    
    excel_bytes = generate_excel_with_formulas(r, P_global, df_earthwork_global)
    st.download_button(
        label="📥 Unduh Laporan Excel Ber-Formula (.xlsx)",
        data=excel_bytes,
        file_name="Perencanaan_Geometrik_Jalan_Formulas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
