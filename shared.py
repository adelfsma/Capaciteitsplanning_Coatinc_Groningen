import os
import json
from datetime import date, timedelta
from pathlib import Path
import numpy as np
import pandas as pd

APP_VERSION = "v2.2"
STATUS_MAP = {
    "uitgeleverd": "Verzinkt","PC Afgehaald": "Verzinkt","UB V Gereed": "UB","Afgehaald": "Verzinkt",
    "Gereed": "Verzinkt","UB": "UB","coat gereed": "Verzinkt","Opgehangen": "Niet verzinkt",
    "Voorbewerking uitvoeren": "Niet verzinkt","Productie gereed": "Niet verzinkt",
    "Nabewerking nog uitvoeren": "Verzinkt","Ontzinkt": "Niet verzinkt",
}
NL_DAY_ABBR = {0:"ma",1:"di",2:"wo",3:"do",4:"vr",5:"za",6:"zo"}
REQUIRED_FILES = ["OrderExport2G.xlsx","Export-1.xlsx","Export.xlsx","Export+1.xlsx","Export+2.xlsx","Export+3.xlsx","Export+4.xlsx","feestdagen.xlsx"]
PUBLISHED_DIR = Path("published_data")
METADATA_FILE = PUBLISHED_DIR / "metadata.json"

# Tijdvenster voor reserveringen (conform Power BI Merge1-logica)
RESERVERING_WINDOW_VOOR = 2   # dagen vóór vandaag
RESERVERING_WINDOW_NA   = 40  # dagen na vandaag
RESERVERING_LOCATIE     = "Coatinc Groningen"

def ensure_dirs():
    PUBLISHED_DIR.mkdir(parents=True, exist_ok=True)

def previous_workday(d: date) -> date:
    d = d - timedelta(days=1)
    while d.weekday() >= 5:
        d = d - timedelta(days=1)
    return d

def add_workdays(d: pd.Timestamp, days: int, holiday_dates: set) -> pd.Timestamp:
    out = pd.Timestamp(d).normalize()
    remaining = int(days)
    while remaining > 0:
        out = out + timedelta(days=1)
        if out.weekday() < 5 and out.date() not in holiday_dates:
            remaining -= 1
    return out

def subtract_workdays_existing_orders(d: pd.Timestamp, days: int) -> pd.Timestamp:
    out = pd.Timestamp(d).normalize()
    remaining = int(days)
    while remaining > 0:
        out = out - timedelta(days=1)
        if out.weekday() < 5:
            remaining -= 1
    return out

def coerce_numeric(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace("\u00a0", "", regex=False)
    s = s.str.replace("kg", "", case=False, regex=False)
    s = s.str.replace("ton", "", case=False, regex=False)
    s = s.str.replace(r"(?<=\d)[.](?=\d{3}\b)", "", regex=True)
    s = s.str.replace(r"(?<=\d),(?=\d{3}\b)", "", regex=True)
    s = s.str.replace(",", ".", regex=False)
    s = s.replace({"": np.nan, "nan": np.nan, "None": np.nan, "<NA>": np.nan})
    return pd.to_numeric(s, errors="coerce")

def extract_week_label(filename: str) -> str:
    name = filename.lower().replace(".xlsx", "")
    if "+" in name:
        return "+" + name.split("+")[-1]
    if "-" in name:
        return "-" + name.split("-")[-1]
    if name == "export":
        return "0"
    return "onbekend"

def format_nl_axis_label(ts: pd.Timestamp) -> str:
    ts = pd.Timestamp(ts)
    return f"{NL_DAY_ABBR[ts.weekday()]} {ts.day:02d}-{ts.month:02d}"

def format_pct(x):
    if pd.isna(x):
        return ""
    return f"{x:.1f}%"

def format_int(x):
    if pd.isna(x):
        return ""
    return f"{int(round(x, 0)):,}".replace(",", ".")

def stoplight(pct: float, is_holiday: bool) -> str:
    if is_holiday:
        return "🔴"
    if pd.isna(pct):
        return ""
    if pct < 80:
        return "🟢"
    if pct <= 100:
        return "🟠"
    return "🔴"

def build_horizon_and_include_holidays(start_dt: pd.Timestamp, production_workdays: int, holiday_dates: set):
    rows = []
    current = pd.Timestamp(start_dt).normalize()
    if current.weekday() >= 5:
        while current.weekday() >= 5:
            current += timedelta(days=1)
    productive_count = 0
    while productive_count < production_workdays:
        if current.weekday() < 5:
            is_holiday = current.date() in holiday_dates
            rows.append({"Verzinkdatum": current, "Is_feestdag_of_sluiting": is_holiday})
            if not is_holiday:
                productive_count += 1
        current += timedelta(days=1)
    return pd.DataFrame(rows)

def load_metadata():
    ensure_dirs()
    if METADATA_FILE.exists():
        with open(METADATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return None

def save_metadata(info: dict):
    ensure_dirs()
    with open(METADATA_FILE, "w", encoding="utf-8") as f:
        json.dump(info, f, indent=2, ensure_ascii=False)

def validate_feestdagen_xlsx(file_path: Path):
    df = pd.read_excel(file_path)
    required = {"Datum","Omschrijving","Type"}
    if not required.issubset(df.columns):
        raise ValueError("feestdagen.xlsx moet de kolommen Datum, Omschrijving en Type bevatten.")
    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
    if df["Datum"].isna().any():
        raise ValueError("feestdagen.xlsx bevat ongeldige datums.")
    return True

def validate_required_files_in_folder(folder: Path):
    missing = [f for f in REQUIRED_FILES if not (folder / f).exists()]
    if missing:
        raise FileNotFoundError("Ontbrekende bestanden: " + ", ".join(missing))
    validate_feestdagen_xlsx(folder / "feestdagen.xlsx")
    return True


def _build_reserveringen(order: pd.DataFrame, cgs_ordernummers: set) -> pd.DataFrame:
    """
    Identificeer reserveringen: orders in OrderExport2G die nog NIET in de
    CGS-export staan, binnen het planningsvenster en voor de juiste locatie.

    Conform de Power BI Merge1-logica:
      - Ordernummer niet in CGS-export (geen match op Hoofdorder)
      - Datum verzending tussen vandaag-2 en vandaag+40
      - Locatie V = "Coatinc Groningen"
    """
    today = pd.Timestamp(date.today()).normalize()
    window_start = today - timedelta(days=RESERVERING_WINDOW_VOOR)
    window_end   = today + timedelta(days=RESERVERING_WINDOW_NA)

    # Datum verzending parsen (flexibel: integer YYYYMMDD of echte datum)
    datum_col = order["Datum verzending"].copy()
    if pd.api.types.is_integer_dtype(datum_col) or datum_col.dropna().astype(str).str.match(r"^\d{8}$").all():
        datum_col = pd.to_datetime(datum_col.astype(str), format="%Y%m%d", errors="coerce")
    else:
        datum_col = pd.to_datetime(datum_col, dayfirst=True, errors="coerce")

    order = order.copy()
    order["Datum_verzending"] = datum_col

    # Locatiekolom bepalen (kan "Locatie V" of "LocatieV" heten)
    locatie_col = None
    for candidate in ["Locatie V", "LocatieV", "Locatie_V"]:
        if candidate in order.columns:
            locatie_col = candidate
            break

    mask = (
        (~order["Ordernummer"].isin(cgs_ordernummers)) &
        (order["Datum_verzending"] >= window_start) &
        (order["Datum_verzending"] <= window_end)
    )
    if locatie_col:
        mask &= order[locatie_col].astype(str).str.strip() == RESERVERING_LOCATIE

    reserveringen = order[mask].copy()

    if reserveringen.empty:
        return pd.DataFrame()

    # Gewicht: voor reserveringen is er geen CGS-deelorder splitsing —
    # het volledige 2G-gewicht per order wordt als één blok meegenomen.
    reserveringen["Nummer"]                 = reserveringen["Ordernummer"].astype(str)
    reserveringen["Leverdatum"]             = reserveringen["Datum_verzending"]
    reserveringen["Gewicht_export_kg"]      = np.nan
    reserveringen["Gewicht_2g_verdeeld_kg"] = reserveringen["Gewicht_order_kg"]
    reserveringen["Gewicht_effectief_kg"]   = reserveringen["Gewicht_order_kg"]
    reserveringen["Gewicht_bron"]           = "Reservering"
    reserveringen["Verzinkstatus"]          = "Niet verzinkt"
    reserveringen["Ordernummer_base"]       = reserveringen["Ordernummer"]
    reserveringen["Regels_per_order"]       = 1
    reserveringen["Bronbestand"]            = "OrderExport2G.xlsx"
    reserveringen["Bron_week"]              = "reservering"
    reserveringen["Status"]                 = "Reservering"
    # Aanleverdatum ontbreekt voor reserveringen; gebruik verzendatum als proxy
    if "Aanleverdatum" not in reserveringen.columns:
        reserveringen["Aanleverdatum"] = reserveringen["Datum_verzending"]

    return reserveringen


def load_published_data():
    ensure_dirs()
    validate_required_files_in_folder(PUBLISHED_DIR)

    # ── 1. CGS-exportbestanden inladen en samenvoegen ──────────────────────
    export_files = ["Export-1.xlsx","Export.xlsx","Export+1.xlsx","Export+2.xlsx","Export+3.xlsx","Export+4.xlsx"]
    export_frames = []
    export_file_summary = []
    for fname in export_files:
        fp = PUBLISHED_DIR / fname
        tmp = pd.read_excel(fp)
        tmp["Bronbestand"] = fname
        tmp["Bron_week"] = extract_week_label(fname)
        export_frames.append(tmp)
        export_file_summary.append({
            "Bronbestand": fname,
            "Bron_week": extract_week_label(fname),
            "Aantal_regels_ingelezen": len(tmp),
        })
    export = pd.concat(export_frames, ignore_index=True)

    # ── 2. OrderExport2G inladen ───────────────────────────────────────────
    order = pd.read_excel(PUBLISHED_DIR / "OrderExport2G.xlsx")
    holiday_df = pd.read_excel(PUBLISHED_DIR / "feestdagen.xlsx")

    # ── 3. Basiskolommen aanmaken ──────────────────────────────────────────
    export["Gewicht_export_kg"] = coerce_numeric(export["Gewicht"])
    export["Leverdatum"]        = pd.to_datetime(export["Datum"], dayfirst=True, errors="coerce")
    export["Verzinkstatus"]     = export["Status"].map(STATUS_MAP)
    export["Ordernummer_base"]  = coerce_numeric(
        export["Nummer"].astype(str).str.extract(r"(\d+)")[0]
    )

    order["Ordernummer"]      = coerce_numeric(order["Ordernummer"])
    order["Gewicht_order_kg"] = coerce_numeric(order["Gewicht(ton)"]) * 1000

    # ── 4. Gewicht verdelen over deelorders (definitieve CGS-orders) ───────
    row_counts = (
        export.groupby("Ordernummer_base", dropna=False)
              .size()
              .rename("Regels_per_order")
              .reset_index()
    )
    merged = export.merge(row_counts, on="Ordernummer_base", how="left")
    merged = merged.merge(
        order[["Ordernummer","Gewicht_order_kg"]],
        left_on="Ordernummer_base", right_on="Ordernummer", how="left"
    )
    merged["Gewicht_2g_verdeeld_kg"] = merged["Gewicht_order_kg"] / merged["Regels_per_order"]
    merged["Gewicht_bron"] = np.where(
        merged["Gewicht_export_kg"].fillna(0) > 0, "Export+", "OrderExport2G verdeeld"
    )
    merged["Gewicht_effectief_kg"] = np.where(
        merged["Gewicht_export_kg"].fillna(0) > 0,
        merged["Gewicht_export_kg"],
        merged["Gewicht_2g_verdeeld_kg"],
    )

    # ── 5. Reserveringen toevoegen (Power BI Merge1-logica) ────────────────
    cgs_ordernummers = set(merged["Ordernummer_base"].dropna().unique())
    reserveringen = _build_reserveringen(order, cgs_ordernummers)

    if not reserveringen.empty:
        merged = pd.concat([merged, reserveringen], ignore_index=True)
        export_file_summary.append({
            "Bronbestand": "OrderExport2G.xlsx (reserveringen)",
            "Bron_week": "reservering",
            "Aantal_regels_ingelezen": len(reserveringen),
        })

    # ── 6. Feestdagen opschonen ────────────────────────────────────────────
    holiday_df["Datum"] = pd.to_datetime(holiday_df["Datum"], errors="coerce").dt.date
    holiday_df = (
        holiday_df.dropna(subset=["Datum"])
                  .drop_duplicates(subset=["Datum","Omschrijving","Type"])
                  .sort_values("Datum")
                  .reset_index(drop=True)
    )

    return merged, pd.DataFrame(export_file_summary), "OrderExport2G.xlsx", holiday_df


def build_dashboard_data(
    df_raw: pd.DataFrame,
    holiday_df: pd.DataFrame,
    startdatum,
    capaciteit_kg: int,
    offset: int,
    kg_per_traverse: int,
):
    holiday_dates = set(holiday_df["Datum"].tolist())
    df = df_raw.copy()
    df["Verzinkdatum"] = df["Leverdatum"].apply(
        lambda x: subtract_workdays_existing_orders(x, offset) if pd.notna(x) else pd.NaT
    )
    start_ts  = pd.Timestamp(startdatum).normalize()
    today_ts  = pd.Timestamp(date.today()).normalize()

    df["Meegeteld_in_planning"] = "Nee"
    df["Reden_uitsluiting"]     = ""

    mask_status_niet_verzinkt = df["Verzinkstatus"] == "Niet verzinkt"
    mask_status_ub            = df["Verzinkstatus"] == "UB"
    mask_status_verzinkt      = df["Verzinkstatus"] == "Verzinkt"
    mask_binnen_horizon       = df["Verzinkdatum"] >= start_ts

    df.loc[mask_status_verzinkt,                          "Reden_uitsluiting"] = "Status = Verzinkt"
    df.loc[mask_status_ub,                                "Reden_uitsluiting"] = "Status = UB"
    df.loc[mask_status_niet_verzinkt & ~mask_binnen_horizon, "Reden_uitsluiting"] = "Voor startdatum rapport"
    df.loc[mask_status_niet_verzinkt &  mask_binnen_horizon, "Meegeteld_in_planning"] = "Ja"
    df.loc[mask_status_niet_verzinkt &  mask_binnen_horizon, "Reden_uitsluiting"]     = ""

    df_plan = df[df["Meegeteld_in_planning"] == "Ja"].copy()

    horizon = build_horizon_and_include_holidays(start_ts, 10, holiday_dates)

    # Aggregeer per dag — reserveringen en definitieve orders samen
    dag_orders = df_plan.groupby("Verzinkdatum", as_index=False).agg(
        Gewicht_kg=("Gewicht_effectief_kg", "sum"),
        Aantal_orders_te_verzinken=("Nummer", "count"),
        # Extra: uitsplitsing definitief vs. reservering voor transparantie
        Gewicht_definitief_kg=(
            "Gewicht_effectief_kg",
            lambda s: s[df_plan.loc[s.index, "Gewicht_bron"] != "Reservering"].sum(),
        ),
        Gewicht_reservering_kg=(
            "Gewicht_effectief_kg",
            lambda s: s[df_plan.loc[s.index, "Gewicht_bron"] == "Reservering"].sum(),
        ),
    )

    order_holiday_dates = df_plan[
        df_plan["Verzinkdatum"].dt.date.isin(holiday_dates)
    ][["Verzinkdatum"]].drop_duplicates()
    if not order_holiday_dates.empty:
        extra_holidays = order_holiday_dates.assign(Is_feestdag_of_sluiting=True)
        horizon = (
            pd.concat([horizon, extra_holidays], ignore_index=True)
              .drop_duplicates(subset=["Verzinkdatum"])
              .sort_values("Verzinkdatum")
              .reset_index(drop=True)
        )

    dag = horizon.merge(dag_orders, on="Verzinkdatum", how="left")
    dag["Gewicht_kg"]                   = dag["Gewicht_kg"].fillna(0.0)
    dag["Gewicht_definitief_kg"]        = dag["Gewicht_definitief_kg"].fillna(0.0)
    dag["Gewicht_reservering_kg"]       = dag["Gewicht_reservering_kg"].fillna(0.0)
    dag["Aantal_orders_te_verzinken"]   = dag["Aantal_orders_te_verzinken"].fillna(0).astype(int)
    dag["Capaciteit_kg"]   = np.where(dag["Is_feestdag_of_sluiting"], 0, capaciteit_kg)
    dag["Benutting_pct"]   = np.where(
        dag["Capaciteit_kg"] > 0,
        (dag["Gewicht_kg"] / dag["Capaciteit_kg"]) * 100,
        np.nan,
    )
    dag["Traverses_berekend"] = np.ceil(dag["Gewicht_kg"] / kg_per_traverse)
    dag["Status"]   = dag.apply(lambda r: stoplight(r["Benutting_pct"], r["Is_feestdag_of_sluiting"]), axis=1)
    dag["Jaar"]     = dag["Verzinkdatum"].dt.isocalendar().year.astype(int)
    dag["Week"]     = dag["Verzinkdatum"].dt.isocalendar().week.astype(int)
    dag["Label_nl"] = dag["Verzinkdatum"].apply(format_nl_axis_label)
    dag["Dagtype"]  = np.where(dag["Is_feestdag_of_sluiting"], "Feestdag / sluiting", "Werkdag")

    week = dag.groupby(["Jaar","Week"], as_index=False).agg(
        Gewicht_kg=("Gewicht_kg","sum"),
        Gewicht_definitief_kg=("Gewicht_definitief_kg","sum"),
        Gewicht_reservering_kg=("Gewicht_reservering_kg","sum"),
        Aantal_orders_te_verzinken=("Aantal_orders_te_verzinken","sum"),
        Traverses_berekend=("Traverses_berekend","sum"),
        Capaciteit_kg=("Capaciteit_kg","sum"),
    )
    week["Benutting_pct"] = np.where(
        week["Capaciteit_kg"] > 0,
        (week["Gewicht_kg"] / week["Capaciteit_kg"]) * 100,
        np.nan,
    )
    week["Status"] = week["Benutting_pct"].apply(lambda x: stoplight(x, False))

    def calculate_advice_date(day_df: pd.DataFrame, today_date: pd.Timestamp, holiday_dates: set):
        base_date  = add_workdays(today_date, 5, holiday_dates)
        productive = day_df[day_df["Is_feestdag_of_sluiting"] == False].copy()
        row = productive[productive["Verzinkdatum"] == base_date]
        if not row.empty and float(row.iloc[0]["Benutting_pct"]) <= 95:
            return base_date
        later = productive[
            (productive["Verzinkdatum"] > base_date) &
            (productive["Benutting_pct"] < 95)
        ]
        if not later.empty:
            return pd.Timestamp(later.iloc[0]["Verzinkdatum"])
        return base_date

    advies_datum = calculate_advice_date(dag, today_ts, holiday_dates)
    return df, df_plan, dag, week, advies_datum
