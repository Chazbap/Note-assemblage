import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import tempfile
import hashlib
import json

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ====== Supabase + live dashboard deps ======
from supabase import create_client
import plotly.graph_objects as go
from streamlit_autorefresh import st_autorefresh

# ==========================================================
# CONFIG / UI
# ==========================================================
st.set_page_config(page_title="🧪 Tableau Assemblage + Stocks", page_icon="🧪", layout="wide")

st.markdown(
    """
    <style>
      .block-container {padding-top: 1.2rem;}
      div[data-testid="stExpander"] details summary {font-size: 1.03rem;}
      .pill {display:inline-block; padding:.2rem .55rem; border-radius:999px; background:#f1f3f5; margin-right:.35rem;}
      .small {opacity:.75; font-size:.9rem;}
      .card {padding: .8rem 1rem; border-radius: 14px; background:#f8f9fa; border: 1px solid #e9ecef;}
      code {font-size: .9rem;}
    </style>
    """,
    unsafe_allow_html=True
)

st.title("🧪 Tableau Assemblage — avec gestion de stock + dégustation live")
st.markdown(
    """
    <span class="pill">C/N/M conservés</span>
    <span class="pill">Essais multiples</span>
    <span class="pill">Récap % fiable</span>
    <span class="pill">Stocks décrémentés par Code Produit</span>
    <span class="pill">Anti double-application</span>
    <span class="pill">Dégustation live (Supabase)</span>
    <span class="pill">PIN par essai</span>
    """,
    unsafe_allow_html=True
)

CEPAGE_LABEL = {"C": "Chardonnay", "N": "Pinot Noir", "M": "Meunier"}

# ==========================================================
# SUPABASE + PIN
# ==========================================================
@st.cache_resource
def sb():
    url = st.secrets.get("SUPABASE_URL", "")
    key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")
    if not url or not key:
        return None
    return create_client(url, key)

def supabase_ready():
    return sb() is not None

def _pin_pepper() -> str:
    return str(st.secrets.get("PIN_PEPPER", "") or "")

def hash_pin(pin: str) -> str:
    """
    Hash SHA-256 du PIN + pepper.
    Le pepper est un secret côté serveur (Streamlit Secrets).
    """
    pin = (pin or "").strip()
    h = hashlib.sha256()
    h.update((pin + _pin_pepper()).encode("utf-8"))
    return h.hexdigest()

def _clean_pin(pin: str) -> str:
    p = (pin or "").strip()
    # tolère espaces, mais on veut chiffres
    p = "".join([c for c in p if c.isdigit()])
    return p

def sup_create_essai(nom: str, cuves: list[str], pin: str) -> str:
    p = _clean_pin(pin)
    pin_hash = hash_pin(p) if p else None
    payload = {"nom": nom, "cuves": cuves, "pin_hash": pin_hash}
    res = sb().table("essais").insert(payload).execute()
    return res.data[0]["id"]

def sup_list_essais(limit: int = 50) -> pd.DataFrame:
    # inclut pin_hash pour pouvoir savoir si essai protégé
    res = sb().table("essais").select("id,created_at,nom,cuves,pin_hash").order("created_at", desc=True).limit(limit).execute()
    return pd.DataFrame(res.data)

def sup_get_essai(essai_id: str) -> dict:
    res = sb().table("essais").select("id,created_at,nom,cuves,pin_hash").eq("id", essai_id).single().execute()
    return res.data

def sup_upsert_note(essai_id: str, cuve: str, degustateur: str, notes: dict, commentaire: str):
    row = {
        "essai_id": essai_id,
        "cuve": cuve,
        "degustateur": degustateur,
        **notes,
        "commentaire": commentaire or "",
    }
    try:
        sb().table("notes").insert(row).execute()
    except Exception:
        sb().table("notes").update({**notes, "commentaire": commentaire or ""}) \
            .eq("essai_id", essai_id).eq("cuve", cuve).eq("degustateur", degustateur).execute()

@st.cache_data(ttl=2, show_spinner=False)
def sup_fetch_notes(essai_id: str) -> pd.DataFrame:
    res = sb().table("notes") \
        .select("created_at,essai_id,cuve,degustateur,acidite,amertume,mineralite,volume,sucrosite,defaut,commentaire") \
        .eq("essai_id", essai_id) \
        .order("created_at", desc=False) \
        .limit(5000) \
        .execute()
    df = pd.DataFrame(res.data)
    if df.empty:
        return df
    for c in ["acidite", "amertume", "mineralite", "volume", "sucrosite", "defaut"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

RADAR_AXES = ["Acidité", "Amertume", "Minéralité", "Volume", "Sucrosité", "Pureté"]  # Pureté = 6 - défaut

def radar_fig(df_cuve: pd.DataFrame, by_taster: bool = False):
    d = df_cuve.copy()
    d["purete"] = 6 - d["defaut"]

    mean_vals = [
        d["acidite"].mean(),
        d["amertume"].mean(),
        d["mineralite"].mean(),
        d["volume"].mean(),
        d["sucrosite"].mean(),
        d["purete"].mean(),
    ]

    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=mean_vals + [mean_vals[0]],
        theta=RADAR_AXES + [RADAR_AXES[0]],
        fill="toself",
        name="Moyenne",
    ))

    if by_taster:
        for degust, g in d.groupby("degustateur"):
            vals = [
                g["acidite"].mean(),
                g["amertume"].mean(),
                g["mineralite"].mean(),
                g["volume"].mean(),
                g["sucrosite"].mean(),
                g["purete"].mean(),
            ]
            fig.add_trace(go.Scatterpolar(
                r=vals + [vals[0]],
                theta=RADAR_AXES + [RADAR_AXES[0]],
                name=str(degust),
            ))

    fig.update_layout(
        margin=dict(l=10, r=10, t=25, b=10),
        polar=dict(radialaxis=dict(visible=True, range=[1, 5], dtick=1)),
        height=300,
        showlegend=by_taster,
    )
    return fig

# ==========================================================
# HELPERS COMMUNS
# ==========================================================
def norm_str_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip()

def normalize_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def safe_int_str(x):
    """Convertit vers int si possible, sinon renvoie str(x). Évite ValueError sur NaN/vides."""
    if pd.isna(x) or x is None or str(x).strip() == "":
        return ""
    try:
        return str(int(float(str(x).replace(",", ".").strip()))))
    except Exception:
        return str(x).strip()

def normalize_cuve_number(x):
    """Corrige des cuves typées en décimal (0,0351) -> 351."""
    if pd.isna(x):
        return x
    if isinstance(x, str):
        xs = x.strip().replace(" ", "").replace(",", ".")
        try:
            x = float(xs)
        except Exception:
            return x
    try:
        v = float(x)
    except Exception:
        return x
    if 0 < v < 1:
        v = v * 10000
    return int(round(v))

def to_cepage_code(v):
    if pd.isna(v):
        return ""
    s = str(v).strip()
    su = s.upper()
    if su in ("C", "N", "M"):
        return su
    if su.startswith("CHARD"):
        return "C"
    if "PINOT" in su and "NOIR" in su:
        return "N"
    if "MEUNIER" in su:
        return "M"
    return su

def essai_cols(e: int):
    return {
        "vol": f"Volume (L) E{e}",
        "solde": f"Solde E{e}",
        "qty": f"Quantité utilisée E{e}",
        "pct": f"% E{e}",
        "c250": f"250 E{e}",
        "c500": f"500 E{e}",
    }

def excel_cols_for_essai(e: int, start_base_col: int = 7):
    start = start_base_col + (e - 1) * 6
    return {"start": start, "vol": start, "solde": start + 1, "qty": start + 2, "pct": start + 3, "c250": start + 4, "c500": start + 5, "end": start + 5}

def force_excel_years_to_int(ws, year_col_letter: str, cuve_col_letter: str, start_row: int, end_row: int):
    for r in range(start_row, end_row + 1):
        cuv = ws[f"{cuve_col_letter}{r}"].value
        if cuv in (None, "", 0):
            continue
        v = ws[f"{year_col_letter}{r}"].value
        if isinstance(v, str):
            vv = v.strip()
            if vv.isdigit():
                ws[f"{year_col_letter}{r}"].value = int(vv)

def coerce_float(series):
    return pd.to_numeric(series, errors="coerce")

def find_header_row(excel_path, needle="Clé Produit en Cuve", sheet_name=0, max_scan=80):
    raw = pd.read_excel(excel_path, header=None, sheet_name=sheet_name)
    max_r = min(max_scan, len(raw))
    for i in range(max_r):
        row = raw.iloc[i].astype(str)
        if row.str.contains(needle, na=False).any():
            return i
    return None

# ==========================================================
# STOCK: anti double-application + delta
# ==========================================================
def make_fingerprint(file_bytes: bytes, essai: str, ref: str, date_conso) -> str:
    h = hashlib.sha256()
    h.update(file_bytes)
    h.update(str(essai).encode("utf-8"))
    h.update(str(ref).encode("utf-8"))
    h.update(str(date_conso).encode("utf-8"))
    return h.hexdigest()[:16]

def journal_has_fingerprint(journal_df: pd.DataFrame, fingerprint: str) -> bool:
    if journal_df is None or journal_df.empty:
        return False
    if "Fingerprint" not in journal_df.columns:
        return False
    return (journal_df["Fingerprint"].astype(str) == str(fingerprint)).any()

def build_delta_table(snapshot_df: pd.DataFrame, ledger_df: pd.DataFrame) -> pd.DataFrame:
    s = snapshot_df.copy()
    l = ledger_df.copy()

    s["Code Produit en Cuve"] = s["Code Produit en Cuve"].astype(str).str.strip()
    l["Code Produit en Cuve"] = l["Code Produit en Cuve"].astype(str).str.strip()

    if "Stock_Etat_L" not in s.columns:
        raise ValueError("snapshot_df doit contenir Stock_Etat_L")
    if "Stock restant (L)" not in l.columns:
        raise ValueError("ledger_df doit contenir Stock restant (L)")

    delta = s.merge(
        l[["Code Produit en Cuve", "Stock restant (L)"]],
        on="Code Produit en Cuve",
        how="outer"
    )

    delta["Stock_Etat_L"] = pd.to_numeric(delta["Stock_Etat_L"], errors="coerce").fillna(0.0)
    delta["Stock restant (L)"] = pd.to_numeric(delta["Stock restant (L)"], errors="coerce").fillna(0.0)
    delta["Écart (Etat - Ledger)"] = (delta["Stock_Etat_L"] - delta["Stock restant (L)"]).round(2)

    delta["__abs"] = delta["Écart (Etat - Ledger)"].abs()
    delta = delta.sort_values("__abs", ascending=False).drop(columns="__abs")
    return delta

# ==========================================================
# ONGLET 2 : STOCK UPDATE
# ==========================================================
def build_stock_snapshot(df_stock):
    required = {"Produit", "En Stock"}
    if not required.issubset(set(df_stock.columns)):
        missing = sorted(list(required - set(df_stock.columns)))
        raise ValueError(f"Colonnes manquantes dans l'état de stock: {missing}")

    df = df_stock.copy()
    df["Produit"] = df["Produit"].apply(normalize_str)
    df["En Stock"] = coerce_float(df["En Stock"]).fillna(0)

    agg = (
        df.groupby("Produit", as_index=False)
        .agg(Stock_Etat_L=("En Stock", "sum"))
        .sort_values("Produit")
    )
    agg.rename(columns={"Produit": "Code Produit en Cuve"}, inplace=True)
    return agg

def init_ledger_from_snapshot(snapshot_df):
    led = snapshot_df.copy()
    led["Stock initial (L)"] = led["Stock_Etat_L"]
    led["Consommé cumul (L)"] = 0.0
    led["Stock restant (L)"] = led["Stock initial (L)"] - led["Consommé cumul (L)"]
    led = led[["Code Produit en Cuve", "Stock initial (L)", "Consommé cumul (L)", "Stock restant (L)"]]
    return led

def read_ledger(ledger_xlsx):
    df = pd.read_excel(ledger_xlsx, sheet_name="STOCK_MAJ" if "STOCK_MAJ" in pd.ExcelFile(ledger_xlsx).sheet_names else 0)
    required = {"Code Produit en Cuve", "Stock initial (L)", "Consommé cumul (L)", "Stock restant (L)"}
    if not required.issubset(set(df.columns)):
        raise ValueError(
            "Ledger invalide. Colonnes attendues : Code Produit en Cuve, Stock initial (L), Consommé cumul (L), Stock restant (L)"
        )
    df["Code Produit en Cuve"] = df["Code Produit en Cuve"].apply(normalize_str)
    for c in ["Stock initial (L)", "Consommé cumul (L)", "Stock restant (L)"]:
        df[c] = coerce_float(df[c]).fillna(0)
    return df

def read_existing_journal(ledger_xlsx):
    try:
        xls = pd.ExcelFile(ledger_xlsx)
        if "JOURNAL" in xls.sheet_names:
            j = pd.read_excel(ledger_xlsx, sheet_name="JOURNAL")
            return j
        return None
    except Exception:
        return None

def read_consumption_from_assemblage(assemblage_xlsx, essai="E1"):
    header_row = find_header_row(assemblage_xlsx, needle="Clé Produit en Cuve")
    if header_row is None:
        raise ValueError("Impossible de trouver l'en-tête du tableau dans l'assemblage (Clé Produit en Cuve).")

    df = pd.read_excel(assemblage_xlsx, header=header_row)
    qty_col = f"Quantité utilisée {essai}"

    if "Code Produit en Cuve" not in df.columns:
        raise ValueError("Colonne 'Code Produit en Cuve' introuvable dans l'assemblage.")
    if qty_col not in df.columns:
        raise ValueError(f"Colonne '{qty_col}' introuvable dans l'assemblage.")

    d = df.copy()
    d["Code Produit en Cuve"] = d["Code Produit en Cuve"].apply(normalize_str)
    d[qty_col] = coerce_float(d[qty_col])

    d = d[(d["Code Produit en Cuve"] != "") & (d["Code Produit en Cuve"].str.upper() != "SOUS-TOTAL")]
    d = d[d[qty_col].fillna(0) > 0].copy()

    cons = (
        d.groupby("Code Produit en Cuve", as_index=False)
        .agg(**{"Consommé (L)": (qty_col, "sum")})
        .sort_values("Code Produit en Cuve")
    )
    cons["Consommé (L)"] = cons["Consommé (L)"].round(2)
    return cons

def apply_consumption(ledger_df, cons_df):
    led = ledger_df.copy()
    cons = cons_df.copy()

    merged = led.merge(cons, on="Code Produit en Cuve", how="outer")
    merged["Stock initial (L)"] = merged["Stock initial (L)"].fillna(0)
    merged["Consommé cumul (L)"] = merged["Consommé cumul (L)"].fillna(0)
    merged["Stock restant (L)"] = merged["Stock restant (L)"].fillna(merged["Stock initial (L)"] - merged["Consommé cumul (L)"])
    merged["Consommé (L)"] = merged["Consommé (L)"].fillna(0)

    merged["Stock restant après (L)"] = merged["Stock restant (L)"] - merged["Consommé (L)"]
    merged["Surconsommation (L)"] = np.where(merged["Stock restant après (L)"] < -1e-9, -merged["Stock restant après (L)"], 0)

    merged["Consommé cumul (L)"] = merged["Consommé cumul (L)"] + merged["Consommé (L)"]
    merged["Stock restant (L)"] = merged["Stock restant après (L)"]

    updated = merged.drop(columns=["Stock restant après (L)"]).copy()
    updated = updated[["Code Produit en Cuve", "Stock initial (L)", "Consommé cumul (L)", "Stock restant (L)", "Consommé (L)", "Surconsommation (L)"]]
    updated = updated.sort_values("Code Produit en Cuve")
    return updated

def export_stock_with_highlight_and_journal(updated_df, cons_df, journal_df, ref_assemblage, date_conso):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        out_path = tmp.name

    stock_sheet = "STOCK_MAJ"
    recap_sheet = "RECAP_CONSO"
    journal_sheet = "JOURNAL"

    recap = cons_df.copy()
    total = float(recap["Consommé (L)"].sum()) if not recap.empty else 0.0
    recap = pd.concat(
        [
            recap,
            pd.DataFrame([{"Code Produit en Cuve": "TOTAL", "Consommé (L)": round(total, 2)}]),
        ],
        ignore_index=True
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        updated_df.drop(columns=["Surconsommation (L)"], errors="ignore").to_excel(writer, sheet_name=stock_sheet, index=False)
        recap.to_excel(writer, sheet_name=recap_sheet, index=False)
        if journal_df is not None:
            journal_df.to_excel(writer, sheet_name=journal_sheet, index=False)

    wb = load_workbook(out_path)
    ws = wb[stock_sheet]
    ws2 = wb[recap_sheet]
    ws3 = wb[journal_sheet] if journal_sheet in wb.sheetnames else None

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    changed_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.insert_rows(1)
    ws.merge_cells("A1:E1")
    ws["A1"] = f"Stock mis à jour — {ref_assemblage} — {date_conso.strftime('%d/%m/%Y')}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    for cell in ws[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 14

    header_map = {ws.cell(row=2, column=c).value: c for c in range(1, ws.max_column + 1)}
    cons_col = header_map.get("Consommé (L)")

    for r in range(3, ws.max_row + 1):
        cons_val = ws.cell(row=r, column=cons_col).value if cons_col else 0
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            if c != 1:
                cell.number_format = "0.00"
        if cons_col and cons_val and float(cons_val) > 0:
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).fill = changed_fill

    ws2.insert_rows(1)
    ws2.merge_cells("A1:B1")
    ws2["A1"] = "RÉCAP des quantités utilisées (assemblage)"
    ws2["A1"].font = Font(bold=True, size=13)
    ws2["A1"].alignment = Alignment(horizontal="center")

    for cell in ws2[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    ws2.column_dimensions["A"].width = 24
    ws2.column_dimensions["B"].width = 16

    for r in range(3, ws2.max_row + 1):
        ws2[f"A{r}"].border = border
        ws2[f"B{r}"].border = border
        ws2[f"B{r}"].number_format = "0.00"
        if ws2[f"A{r}"].value == "TOTAL":
            ws2[f"A{r}"].font = Font(bold=True)
            ws2[f"B{r}"].font = Font(bold=True)
            ws2[f"A{r}"].fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
            ws2[f"B{r}"].fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")

    if ws3 is not None:
        for cell in ws3[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        for r in range(2, ws3.max_row + 1):
            for c in range(1, ws3.max_column + 1):
                ws3.cell(row=r, column=c).border = border

    wb.save(out_path)
    return out_path

# ==========================================================
# ONGLET 1 : ASSEMBLAGE (COMPLET)
# ==========================================================
def tab_assemblage():
    with st.sidebar:
        st.header("🧪 Assemblage — imports")
        uploaded_file_cuves = st.file_uploader("État cuverie (Excel)", type=["xlsx"], key="cuves_ass")
        uploaded_file_codes = st.file_uploader("Codes produits (Excel)", type=["xlsx"], key="codes_ass")
        uploaded_file_codes_assemblage = st.file_uploader("Liste produits ASSEMBLAGE (Excel)", type=["xlsx"], key="ass_list_ass")

        st.divider()
        st.header("🧪 Assemblage — paramètres")
        ESSAIS = st.number_input("Nombre d'essais", min_value=1, max_value=10, value=5, step=1, key="essais_ass")
        titre_excel = st.text_input("Titre du fichier", value="Assemblage Avril 2025", key="titre_ass")

    if not (uploaded_file_cuves and uploaded_file_codes and uploaded_file_codes_assemblage):
        st.info("👉 Importer les 3 fichiers (cuverie + codes + liste assemblage) dans la sidebar.")
        return

    df_cuves = pd.read_excel(uploaded_file_cuves)
    df_codes = pd.read_excel(uploaded_file_codes)
    df_codes_ass = pd.read_excel(uploaded_file_codes_assemblage)

    # Normalisation
    if "Produit" in df_cuves.columns:
        df_cuves["Produit"] = norm_str_series(df_cuves["Produit"])
    if "Cépage" in df_cuves.columns:
        df_cuves["Cépage"] = norm_str_series(df_cuves["Cépage"])
    if "N° Cuve" in df_cuves.columns:
        df_cuves["N° Cuve"] = df_cuves["N° Cuve"].apply(normalize_cuve_number)

    for d in (df_codes, df_codes_ass):
        for col in ["Code Produit en Cuve", "Clé Produit en Cuve", "Libéllé Produit en Cuve"]:
            if col in d.columns:
                d[col] = norm_str_series(d[col])

    # Check colonnes
    for col in ["En Stock", "Année", "Produit", "N° Cuve", "Cépage"]:
        if col not in df_cuves.columns:
            st.error(f"Le fichier cuverie doit contenir la colonne '{col}'.")
            return

    # Filtre stock
    df_cuves = df_cuves[df_cuves["En Stock"] > 0].copy()

    # Assemblages
    if "Code Produit en Cuve" not in df_codes_ass.columns:
        st.error("Le fichier ASSEMBLAGE doit contenir une colonne 'Code Produit en Cuve'.")
        return
    set_assemblages = set(df_codes_ass["Code Produit en Cuve"].dropna().astype(str).str.strip().tolist())

    # Type année/réserve
    df_cuves["Type"] = df_cuves["Année"].apply(lambda x: "Vin de l'année" if x >= 2025 else "Vin de réserve")

    df_cuves_ass = df_cuves[df_cuves["Produit"].isin(set_assemblages)].copy()
    df_cuves_std = df_cuves[~df_cuves["Produit"].isin(set_assemblages)].copy()
    df_cuves_std["CépageCode"] = df_cuves_std["Cépage"].apply(to_cepage_code)

    # UI sélection
    c1, c2, c3 = st.columns(3)
    c1.metric("Cuves en stock", int(len(df_cuves)))
    c2.metric("Cuves standard", int(df_cuves_std["N° Cuve"].nunique()))
    c3.metric("Cuves assemblage", int(df_cuves_ass["N° Cuve"].nunique()))

    st.subheader("🧩 Sélection des cuves")
    st.caption("Sélectionne des cuves par catégorie. Tu peux cumuler Assemblage + cépages classiques.")

    cuves_selectionnees = []

    if not df_cuves_ass.empty:
        with st.expander("🧩 ASSEMBLAGE", expanded=True):
            cuves_ass = st.multiselect(
                "Sélectionner les cuves ASSEMBLAGE",
                options=df_cuves_ass["N° Cuve"].tolist(),
                format_func=lambda x: (
                    f"{x} - {df_cuves_ass.loc[df_cuves_ass['N° Cuve'] == x, 'Produit'].values[0]} "
                    f"({df_cuves_ass.loc[df_cuves_ass['N° Cuve'] == x, 'En Stock'].values[0]} L)"
                ),
                key="assemblages_select"
            )
            cuves_selectionnees.extend(cuves_ass)

    cepage_codes = [c for c in df_cuves_std["CépageCode"].dropna().unique().tolist() if str(c).strip() != ""]
    if df_cuves_std.empty:
        st.warning("Aucune cuve standard (hors assemblage) trouvée en stock.")
        return

    def sort_key(x):
        order = {"C": 1, "N": 2, "M": 3}
        return (order.get(str(x).upper(), 99), str(x))

    for code in sorted(cepage_codes, key=sort_key):
        label = CEPAGE_LABEL.get(str(code).upper(), str(code))
        with st.expander(f"🍇 {label}", expanded=True):
            df_cepage = df_cuves_std[df_cuves_std["CépageCode"] == code].copy()

            df_annee = df_cepage[df_cepage["Type"] == "Vin de l'année"]
            if not df_annee.empty:
                st.markdown("**🟢 Vin de l'année (>= 2025)**")
                cuves_annee = st.multiselect(
                    f"{label} - Vin de l'année",
                    options=df_annee["N° Cuve"].tolist(),
                    format_func=lambda x, d=df_annee: (
                        f"{x} - {d.loc[d['N° Cuve'] == x, 'Produit'].values[0]} "
                        f"({d.loc[d['N° Cuve'] == x, 'En Stock'].values[0]} L)"
                    ),
                    key=f"{code}_annee_select"
                )
                cuves_selectionnees.extend(cuves_annee)

            df_reserve = df_cepage[df_cepage["Type"] == "Vin de réserve"]
            if not df_reserve.empty:
                st.markdown("**🟡 Vins de réserve (< 2025)**")
                cuves_reserve = st.multiselect(
                    f"{label} - Réserve",
                    options=df_reserve["N° Cuve"].tolist(),
                    format_func=lambda x, d=df_reserve: (
                        f"{x} - {d.loc[d['N° Cuve'] == x, 'Produit'].values[0]} "
                        f"({d.loc[d['N° Cuve'] == x, 'En Stock'].values[0]} L - {d.loc[d['N° Cuve'] == x, 'Année'].values[0]})"
                    ),
                    key=f"{code}_reserve_select"
                )
                cuves_selectionnees.extend(cuves_reserve)

    st.divider()
    st.write(f"✅ **Cuves sélectionnées : {len(set(cuves_selectionnees))}**")

    if not cuves_selectionnees:
        st.info("👉 Sélectionne au moins une cuve (standard ou assemblage) pour générer le fichier.")
        return

    df_selection = df_cuves[df_cuves["N° Cuve"].isin(cuves_selectionnees)].copy()

    # Fusion codes
    df_codes_all = pd.concat([df_codes, df_codes_ass], ignore_index=True)
    needed_cols = {"Code Produit en Cuve", "Clé Produit en Cuve", "Libéllé Produit en Cuve"}
    missing = needed_cols - set(df_codes_all.columns)
    if missing:
        st.error(
            f"Le fichier codes doit contenir : {', '.join(sorted(needed_cols))}. "
            f"Manquantes : {', '.join(sorted(missing))}"
        )
        return

    df_codes_all = df_codes_all.drop_duplicates(subset=["Code Produit en Cuve"], keep="first")

    df_selection = df_selection.merge(
        df_codes_all[["Code Produit en Cuve", "Clé Produit en Cuve", "Libéllé Produit en Cuve"]],
        how="left",
        left_on="Produit",
        right_on="Code Produit en Cuve"
    )

    # Préparer la liste de cuves pour dégustation
    cuves_for_tasting = (
        df_selection[["N° Cuve", "Produit"]]
        .drop_duplicates()
        .sort_values(["N° Cuve", "Produit"])
        .apply(lambda r: f"{safe_int_str(r['N° Cuve'])} - {str(r['Produit']).strip()}", axis=1)
        .tolist()
    )
    st.session_state["last_cuves_for_tasting"] = cuves_for_tasting
    st.session_state["last_titre_excel"] = st.session_state.get("titre_ass", "Assemblage")

    # Catégorie couleur (C/N/M/ASSEMBLAGE)
    df_selection["Catégorie couleur"] = df_selection.apply(
        lambda r: "ASSEMBLAGE" if str(r["Produit"]).strip() in set_assemblages else to_cepage_code(r["Cépage"]),
        axis=1
    )

    # Cépage affiché (texte)
    df_selection["Cépage_aff"] = df_selection["Catégorie couleur"].apply(
        lambda c: "Assemblage" if str(c).strip().upper() == "ASSEMBLAGE" else CEPAGE_LABEL.get(str(c).strip().upper(), str(c))
    )

    def _annee_export(row):
        if str(row["Produit"]).strip() in set_assemblages:
            return ""
        try:
            return int(row["Année"])
        except Exception:
            return row["Année"]

    df_selection["Année_export"] = df_selection.apply(_annee_export, axis=1)

    df_selection["Is_reserve"] = df_selection.apply(
        lambda r: 0 if str(r["Produit"]).strip() in set_assemblages else (1 if r["Type"] == "Vin de réserve" else 0),
        axis=1
    )

    df_export = df_selection[[
        "Clé Produit en Cuve",
        "N° Cuve",
        "Produit",
        "Libéllé Produit en Cuve",
        "Cépage_aff",
        "Année_export",
        "En Stock",
        "Catégorie couleur",
        "Is_reserve",
    ]].copy()

    df_export.columns = [
        "Clé Produit en Cuve",
        "N° Cuve",
        "Code Produit en Cuve",
        "Libellé Produit en Cuve",
        "Cépage",
        "Année",
        "Volume_base",
        "Catégorie couleur",
        "Is_reserve",
    ]

    df_export["__cat_order"] = df_export["Catégorie couleur"].apply(lambda x: 1 if str(x).strip().upper() != "ASSEMBLAGE" else 3)
    df_export["__reserve_order"] = df_export["Is_reserve"].apply(lambda x: 2 if int(x) == 1 else 1)

    def annee_sort(v):
        v = str(v).strip()
        if v == "":
            return -999999
        try:
            return int(v)
        except Exception:
            return -999999

    df_export["__annee"] = df_export["Année"].apply(annee_sort)

    df_export = (
        df_export.sort_values(
            by=["__cat_order", "__reserve_order", "Cépage", "__annee", "Code Produit en Cuve", "N° Cuve"],
            ascending=[True, True, True, False, True, True]
        )
        .drop(columns=["__cat_order", "__reserve_order", "__annee", "Is_reserve"])
    )

    df_sommaire = (
        df_export.groupby(["Cépage", "Année"], dropna=False)
        .agg(Nb_Cuves=("N° Cuve", "nunique"), Volume_L=("Volume_base", "sum"))
        .reset_index()
    )

    def libelle_bloc(cepage, annee):
        c = str(cepage).strip()
        a = str(annee).strip()
        return "ASSEMBLAGE" if c.lower() == "assemblage" else f"{c} {a}"

    df_sommaire["Libellé"] = df_sommaire.apply(lambda r: libelle_bloc(r["Cépage"], r["Année"]), axis=1)
    df_sommaire = df_sommaire[["Libellé", "Nb_Cuves", "Volume_L"]]
    df_sommaire = pd.concat([df_sommaire, pd.DataFrame([{
        "Libellé": "TOTAL",
        "Nb_Cuves": int(df_sommaire["Nb_Cuves"].sum()),
        "Volume_L": float(df_sommaire["Volume_L"].sum())
    }])], ignore_index=True)

    base_cols = ["Clé Produit en Cuve", "N° Cuve", "Code Produit en Cuve", "Libellé Produit en Cuve", "Cépage", "Année"]
    essais_cols = []
    ESSAIS = int(st.session_state.get("essais_ass", 5))
    for e in range(1, ESSAIS + 1):
        c = essai_cols(e)
        essais_cols.extend([c["vol"], c["solde"], c["qty"], c["pct"], c["c250"], c["c500"]])

    df_final = pd.DataFrame(columns=base_cols + essais_cols + ["Catégorie couleur"])

    for (cepage_aff, annee, cat_color), group in df_export.groupby(["Cépage", "Année", "Catégorie couleur"], sort=False):
        titre = "ASSEMBLAGE" if str(cepage_aff).strip().lower() == "assemblage" else f"{cepage_aff} {annee}"
        titre_row = {col: "" for col in df_final.columns}
        titre_row["Clé Produit en Cuve"] = titre
        titre_row["Catégorie couleur"] = cat_color
        df_final = pd.concat([df_final, pd.DataFrame([titre_row])], ignore_index=True)

        rows = []
        for _, r in group.iterrows():
            row = {col: "" for col in df_final.columns}
            row["Clé Produit en Cuve"] = r["Clé Produit en Cuve"]
            row["N° Cuve"] = r["N° Cuve"]
            row["Code Produit en Cuve"] = r["Code Produit en Cuve"]
            row["Libellé Produit en Cuve"] = r["Libellé Produit en Cuve"]
            row["Cépage"] = r["Cépage"]
            row["Année"] = r["Année"]
            row["Catégorie couleur"] = r["Catégorie couleur"]
            for e in range(1, ESSAIS + 1):
                cc = essai_cols(e)
                row[cc["vol"]] = r["Volume_base"]
            rows.append(row)

        df_final = pd.concat([df_final, pd.DataFrame(rows)], ignore_index=True)

        st_row = {col: "" for col in df_final.columns}
        st_row["Clé Produit en Cuve"] = "Sous-total"
        st_row["Année"] = annee
        st_row["Catégorie couleur"] = cat_color
        df_final = pd.concat([df_final, pd.DataFrame([st_row])], ignore_index=True)

    # =========================
    # EXCEL OUTPUT
    # =========================
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        fichier_excel = tmp.name
        df_final.to_excel(fichier_excel, index=False)

        wb = load_workbook(fichier_excel)
        ws = wb.active

        start_base_col = 7
        last_visible_col = 6 + ESSAIS * 6
        last_visible_letter = get_column_letter(last_visible_col)

        tech_col_idx = last_visible_col + 1
        tech_col_letter = get_column_letter(tech_col_idx)

        cuv_col = "B"
        ann_col = "F"
        titre_excel = st.session_state.get("titre_ass", "Assemblage")

        ws.insert_rows(1)
        ws.merge_cells(f"A1:{last_visible_letter}1")
        ws["A1"] = titre_excel
        ws["A1"].font = Font(bold=True, size=16)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26

        ws.column_dimensions[tech_col_letter].hidden = True

        fill_vert = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_rouge = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        fill_gris = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        fill_sous_total = PatternFill(start_color="595959", end_color="595959", fill_type="solid")
        fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        font_header = Font(bold=True, color="FFFFFF")

        border_epaisse = Border(
            top=Side(border_style="thick", color="000000"),
            left=Side(border_style="thick", color="000000"),
            right=Side(border_style="thick", color="000000"),
            bottom=Side(border_style="thick", color="000000")
        )

        n = len(df_sommaire)
        ws.insert_rows(2, amount=(1 + 1 + n + 1))
        ws["A2"] = "SOMMAIRE"
        ws["A2"].font = Font(bold=True, size=12)

        ws["A3"], ws["B3"], ws["C3"] = "Catégorie", "Nb cuves", "Volume (L)"
        for cell in (ws["A3"], ws["B3"], ws["C3"]):
            cell.fill = fill_header
            cell.font = font_header

        start = 4
        for ridx, row in enumerate(df_sommaire.itertuples(index=False), start=start):
            ws[f"A{ridx}"] = row.Libellé
            ws[f"B{ridx}"] = row.Nb_Cuves
            ws[f"C{ridx}"] = row.Volume_L
            ws[f"C{ridx}"].number_format = "0.00"

        after_summary_row = start + n
        header_row = after_summary_row + 1
        data_start_row = header_row + 1
        ws.freeze_panes = f"A{data_start_row}"

        for c in range(1, last_visible_col + 1):
            cell = ws[f"{get_column_letter(c)}{header_row}"]
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[header_row].height = 20

        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 10
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 34
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 8

        for e in range(1, ESSAIS + 1):
            cols = excel_cols_for_essai(e, start_base_col=start_base_col)
            ws.column_dimensions[get_column_letter(cols["vol"])].width = 12
            ws.column_dimensions[get_column_letter(cols["solde"])].width = 11
            ws.column_dimensions[get_column_letter(cols["qty"])].width = 18
            ws.column_dimensions[get_column_letter(cols["pct"])].width = 9
            ws.column_dimensions[get_column_letter(cols["c250"])].width = 9
            ws.column_dimensions[get_column_letter(cols["c500"])].width = 9

        for r in range(data_start_row, ws.max_row + 1):
            ws[f"B{r}"].number_format = "0"

        force_excel_years_to_int(ws, year_col_letter="F", cuve_col_letter="B", start_row=data_start_row, end_row=ws.max_row)

        current_start = data_start_row
        for r in range(data_start_row, ws.max_row + 1):
            if ws[f"A{r}"].value == "Sous-total":
                for e in range(1, ESSAIS + 1):
                    cols = excel_cols_for_essai(e, start_base_col=start_base_col)
                    for col_idx in range(cols["vol"], cols["end"] + 1):
                        col_letter = get_column_letter(col_idx)
                        ws[f"{col_letter}{r}"].value = f"=SUM({col_letter}{current_start}:{col_letter}{r-1})"
                        if col_idx == cols["pct"]:
                            ws[f"{col_letter}{r}"].number_format = "0.00%"
                        elif col_idx in (cols["c250"], cols["c500"]):
                            ws[f"{col_letter}{r}"].number_format = "0"
                        else:
                            ws[f"{col_letter}{r}"].number_format = "0.00"
                current_start = r + 1

        # ✅ RÉCAP % (via colonne tech)
        recap_rows = [
            ("RÉCAP % - 2025 (Vin de l'année)", None),
            ("Chardonnay", "C"),
            ("Pinot Noir", "N"),
            ("Meunier", "M"),
            ("Assemblage", "ASSEMBLAGE"),
            ("TOTAL 2025", "__TOTAL_2025__"),
            ("RÉCAP % - Total (avec réserves)", None),
            ("Chardonnay", "C"),
            ("Pinot Noir", "N"),
            ("Meunier", "M"),
            ("Assemblage", "ASSEMBLAGE"),
            ("% RÉSERVE", "__RESERVE__"),
            ("TOTAL GLOBAL", "__TOTAL_ALL__"),
        ]
        recap_start_row = ws.max_row + 1

        for k, (label, key) in enumerate(recap_rows):
            rr = recap_start_row + k
            ws[f"A{rr}"] = label
            ws[f"A{rr}"].font = Font(bold=True) if (key is None or str(key).startswith("__")) else Font(bold=False)

            for e in range(1, ESSAIS + 1):
                cols = excel_cols_for_essai(e, start_base_col=start_base_col)
                qty_col_letter = get_column_letter(cols["qty"])
                pct_col_letter = get_column_letter(cols["pct"])

                cuve_criteria = f'{cuv_col}:{cuv_col},">0"'
                denom_2025 = f"SUMIFS({qty_col_letter}:{qty_col_letter},{cuve_criteria},{ann_col}:{ann_col},2025)"
                denom_all = f"SUMIFS({qty_col_letter}:{qty_col_letter},{cuve_criteria})"

                if key is None:
                    ws[f"{pct_col_letter}{rr}"] = ""
                    continue

                if key == "__TOTAL_2025__":
                    ws[f"{pct_col_letter}{rr}"] = f"=IF({denom_2025}=0,0,1)"
                    ws[f"{pct_col_letter}{rr}"].number_format = "0.00%"
                    continue

                if key == "__TOTAL_ALL__":
                    ws[f"{pct_col_letter}{rr}"] = f"=IF({denom_all}=0,0,1)"
                    ws[f"{pct_col_letter}{rr}"].number_format = "0.00%"
                    continue

                if key == "__RESERVE__":
                    num_reserve = (
                        f"SUMIFS({qty_col_letter}:{qty_col_letter},"
                        f'{cuv_col}:{cuv_col},">0",'
                        f'{ann_col}:{ann_col},">0",'
                        f'{ann_col}:{ann_col},"<2025")'
                    )
                    ws[f"{pct_col_letter}{rr}"] = f"=IF({denom_all}=0,0,{num_reserve}/{denom_all})"
                    ws[f"{pct_col_letter}{rr}"].number_format = "0.00%"
                    continue

                is_block_2025 = rr < recap_start_row + 6
                if is_block_2025:
                    num = (
                        f"SUMIFS({qty_col_letter}:{qty_col_letter},"
                        f'{cuv_col}:{cuv_col},">0",'
                        f'{tech_col_letter}:{tech_col_letter},"{key}",'
                        f'{ann_col}:{ann_col},2025)'
                    )
                    ws[f"{pct_col_letter}{rr}"] = f"=IF({denom_2025}=0,0,{num}/{denom_2025})"
                else:
                    num = (
                        f"SUMIFS({qty_col_letter}:{qty_col_letter},"
                        f'{cuv_col}:{cuv_col},">0",'
                        f'{tech_col_letter}:{tech_col_letter},"{key}")'
                    )
                    ws[f"{pct_col_letter}{rr}"] = f"=IF({denom_all}=0,0,{num}/{denom_all})"
                ws[f"{pct_col_letter}{rr}"].number_format = "0.00%"

        dernier_row = ws.max_row + 1
        ws[f"A{dernier_row}"] = "TOTAL"
        ws[f"A{dernier_row}"].font = Font(bold=True, size=12)
        ws[f"A{dernier_row}"].fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
        ws.merge_cells(start_row=dernier_row, start_column=1, end_row=dernier_row, end_column=6)

        for e in range(1, ESSAIS + 1):
            cols = excel_cols_for_essai(e, start_base_col=start_base_col)
            for col_idx in range(cols["vol"], cols["end"] + 1):
                col_letter = get_column_letter(col_idx)
                refs = []
                for rr in range(data_start_row, dernier_row):
                    if ws[f"A{rr}"].value == "Sous-total":
                        refs.append(f"{col_letter}{rr}")
                if refs:
                    ws[f"{col_letter}{dernier_row}"] = f"=SUM({','.join(refs)})"
                    ws[f"{col_letter}{dernier_row}"].font = Font(bold=True)
                    if col_idx == cols["pct"]:
                        ws[f"{col_letter}{dernier_row}"].number_format = "0.00%"
                    elif col_idx in (cols["c250"], cols["c500"]):
                        ws[f"{col_letter}{dernier_row}"].number_format = "0"
                    else:
                        ws[f"{col_letter}{dernier_row}"].number_format = "0.00"

        # Couleurs C/N/M + sous-total
        for r in range(data_start_row, ws.max_row + 1):
            cat = str(ws[f"{tech_col_letter}{r}"].value).strip().upper()
            if cat == "C":
                for cell in ws[r]: cell.fill = fill_vert
            elif cat == "N":
                for cell in ws[r]: cell.fill = fill_rouge
            elif cat == "M" or "ASSEMBLAGE" in cat:
                for cell in ws[r]: cell.fill = fill_gris

            if ws[f"A{r}"].value == "Sous-total":
                for cell in ws[r]:
                    cell.fill = fill_sous_total
                    cell.font = Font(bold=True, color="FFA500")

        for r in range(data_start_row, ws.max_row + 1):
            val = ws[f"A{r}"].value
            cat_val = ws[f"{tech_col_letter}{r}"].value
            if val and val not in ("Sous-total", "TOTAL") and ws[f"B{r}"].value in (None, "") and cat_val not in (None, ""):
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_visible_col)
                ws[f"A{r}"].font = Font(bold=True, size=13)
                ws[f"A{r}"].alignment = Alignment(horizontal="left", vertical="center")

        for e in range(1, ESSAIS + 1):
            cols = excel_cols_for_essai(e, start_base_col=start_base_col)
            left_letter = get_column_letter(cols["start"])
            right_letter = get_column_letter(cols["end"])
            for r in range(header_row, ws.max_row + 1):
                cellL = ws[f"{left_letter}{r}"]
                cellL.border = Border(left=Side(border_style="thick", color="000000"),
                                     right=cellL.border.right, top=cellL.border.top, bottom=cellL.border.bottom)
                cellR = ws[f"{right_letter}{r}"]
                cellR.border = Border(right=Side(border_style="thick", color="000000"),
                                     left=cellR.border.left, top=cellR.border.top, bottom=cellR.border.bottom)

        groupe_debut = None
        for r in range(data_start_row, ws.max_row + 1):
            if ws[f"A{r}"].value not in ("Sous-total", "TOTAL", None, "") and ws[f"B{r}"].value not in (None, ""):
                if groupe_debut is None:
                    groupe_debut = r
            if ws[f"A{r}"].value == "Sous-total" and groupe_debut:
                for row in ws.iter_rows(min_row=groupe_debut, max_row=r, min_col=1, max_col=7):
                    for cell in row:
                        cell.border = border_epaisse
                groupe_debut = None

        total_row = ws.max_row
        for r in range(data_start_row, total_row):
            if ws[f"A{r}"].value == "Sous-total":
                continue
            if ws[f"B{r}"].value in (None, ""):
                continue

            for e in range(1, ESSAIS + 1):
                cols = excel_cols_for_essai(e, start_base_col=start_base_col)
                vol = get_column_letter(cols["vol"])
                solde = get_column_letter(cols["solde"])
                qty = get_column_letter(cols["qty"])
                pct = get_column_letter(cols["pct"])
                c250 = get_column_letter(cols["c250"])
                c500 = get_column_letter(cols["c500"])

                ws[f"{solde}{r}"].value = f"={vol}{r}-{qty}{r}"
                ws[f"{pct}{r}"].value = f"=IF({qty}{total_row}=0,0,{qty}{r}/{qty}{total_row})"
                ws[f"{pct}{r}"].number_format = "0.00%"

                ws[f"{c250}{r}"].value = f"={pct}{r}*250"
                ws[f"{c500}{r}"].value = f"={pct}{r}*500"
                ws[f"{c250}{r}"].number_format = "0"
                ws[f"{c500}{r}"].number_format = "0"

        wb.save(fichier_excel)

    st.success("✅ Fichier assemblage prêt !")
    with open(fichier_excel, "rb") as f:
        st.download_button("📥 Télécharger l'assemblage (Excel)", f, file_name="assemblage.xlsx", use_container_width=True)

    # =========================
    # Dégustation (création essai + PIN + lien)
    # =========================
    st.divider()
    st.subheader("🍷 Dégustation (optionnel)")

    if not supabase_ready():
        st.info("Supabase non configuré (ajoute SUPABASE_URL + SUPABASE_SERVICE_ROLE_KEY + APP_BASE_URL dans st.secrets).")
        return

    cuves_for_tasting = st.session_state.get("last_cuves_for_tasting", [])
    if not cuves_for_tasting:
        st.info("Aucune cuve en session pour créer un essai.")
        return

    default_essai_name = f"{titre_excel} — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    essai_name = st.text_input("Nom de l'essai dégustation", value=default_essai_name, key="essai_name_auto")

    st.caption("🔒 Définis un PIN (ex: 4 chiffres) : tu partageras le lien + le PIN aux dégustateurs.")
    pin = st.text_input("PIN (4 chiffres)", value="", max_chars=8, type="password", key="essai_pin")

    if st.button("✅ Créer l'essai dégustation (Supabase)", type="primary"):
        p = _clean_pin(pin)
        if len(p) < 4:
            st.error("PIN trop court. Mets au moins 4 chiffres.")
            st.stop()

        essai_id = sup_create_essai(essai_name, cuves_for_tasting, p)
        st.session_state["deg_essai_id"] = essai_id
        sup_fetch_notes.clear()

        base = (st.secrets.get("APP_BASE_URL", "") or "").rstrip("/")
        share_url = f"{base}/?essai={essai_id}" if base else ""

        st.success("Essai créé ✅")
        if share_url:
            st.write("➡️ **Lien à envoyer aux dégustateurs :**")
            st.code(share_url)
        st.write("🔑 **PIN à transmettre (séparément) :**")
        st.code(p)
        st.caption("Conseil : envoie le lien sur Teams + le PIN dans un second message.")

# ==========================================================
# ONGLET : DEGUSTATION LIVE (SUPABASE) + PIN
# ==========================================================
def tab_degustation_live():
    st.subheader("🍷 Dégustation Live (multi-dégustateurs, consolidation auto)")

    if not supabase_ready():
        st.error("Supabase non configuré. Ajoute SUPABASE_URL et SUPABASE_SERVICE_ROLE_KEY dans les Secrets Streamlit.")
        st.stop()

    # --- Auto-select essai via URL ?essai=...
    try:
        q = st.query_params  # Streamlit récent
        essai_from_url = q.get("essai", "")
    except Exception:
        q = st.experimental_get_query_params()
        essai_from_url = (q.get("essai", [""]) or [""])[0]

    if essai_from_url:
        st.session_state["deg_essai_id"] = essai_from_url

    st.markdown(
        """
        <div class="card">
          <div><b>Principe :</b> chaque note est enregistrée en base (Supabase). Le dashboard se met à jour automatiquement.</div>
          <div class="small">Pureté = 6 - Défaut (pour que “plus grand = mieux” sur l’araignée).</div>
          <div class="small"><b>Sécurité :</b> un PIN est requis par essai (si défini).</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Charger un essai
    df_ess = sup_list_essais()
    current = st.session_state.get("deg_essai_id", "")

    c1, c2 = st.columns([1.6, 1])
    with c1:
        if df_ess.empty:
            st.info("Aucun essai. Crée-en un depuis l’onglet Assemblage (bouton dégustation) ou ci-dessous.")
        else:
            df_ess["label"] = df_ess.apply(lambda r: f"{r['nom']}  —  {r['created_at']}", axis=1)
            options = dict(zip(df_ess["label"], df_ess["id"]))
            default_idx = 0
            if current and current in df_ess["id"].tolist():
                default_idx = df_ess["id"].tolist().index(current)

            chosen_label = st.selectbox("Essai à utiliser", list(options.keys()), index=default_idx)
            chosen_id = options[chosen_label]
            st.session_state["deg_essai_id"] = chosen_id

    with c2:
        st.caption("Créer un essai manuellement (si besoin)")
        default_name = f"Essai {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        nom = st.text_input("Nom", value=default_name, key="manual_essai_name")
        cuves_txt = st.text_area("Cuves (1 par ligne)", height=100, key="manual_cuves")
        pin_m = st.text_input("PIN (4 chiffres)", value="", max_chars=8, type="password", key="manual_pin")
        cuves = [x.strip() for x in cuves_txt.splitlines() if x.strip()]
        if st.button("➕ Créer essai (manuel)"):
            p = _clean_pin(pin_m)
            if not cuves:
                st.error("Ajoute au moins une cuve.")
            elif len(p) < 4:
                st.error("PIN trop court (min 4 chiffres).")
            else:
                new_id = sup_create_essai(nom, cuves, p)
                st.session_state["deg_essai_id"] = new_id
                st.success("Essai créé ✅")
                sup_fetch_notes.clear()

    essai_id = st.session_state.get("deg_essai_id", "")
    if not essai_id:
        st.stop()

    # Charger l'essai (inclut pin_hash)
    essai = sup_get_essai(essai_id)
    cuves = essai.get("cuves", []) or []
    pin_hash = essai.get("pin_hash", None)

    st.caption(f"Essai : **{essai.get('nom','')}** — {len(cuves)} cuve(s)")

    # --- Gate PIN (par essai) ---
    # On mémorise l'autorisation dans session_state par essai_id
    gate_key = f"pin_ok_{essai_id}"
    if gate_key not in st.session_state:
        st.session_state[gate_key] = False

    if pin_hash and not st.session_state[gate_key]:
        st.warning("🔒 Cet essai est protégé par un PIN.")
        pin_try = st.text_input
