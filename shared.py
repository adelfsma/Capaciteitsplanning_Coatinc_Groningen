import json
import tempfile
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

APP_VERSION = "v2.2"

STATUS_MAP = {
    "uitgeleverd": "Verzinkt",
    "PC Afgehaald": "Verzinkt",
    "UB V Gereed": "UB",
    "Afgehaald": "Verzinkt",
    "Gereed": "Verzinkt",
    "UB": "UB",
    "coat gereed": "Verzinkt",
    "Opgehangen": "Niet verzinkt",
    "Voorbewerking uitvoeren": "Niet verzinkt",
    "Productie gereed": "Niet verzinkt",
    "Nabewerking nog uitvoeren": "Verzinkt",
    "Ontzinkt": "Niet verzinkt",
}

NL_DAY_ABBR = {0: "ma", 1: "di", 2: "wo", 3: "do", 4: "vr", 5: "za", 6: "zo"}

REQUIRED_FILES = [
    "OrderExport2G.xlsx",
    "Export-1.xlsx",
    "Export.xlsx",
    "Export+1.xlsx",
    "Export+2.xlsx",
    "Export+3.xlsx",
    "Export+4.xlsx",
    "feestdagen.xlsx",
]

_TEMP_DIR = Path(tempfile.gettempdir()) / "cap_planning_cache"

# ── Supabase client ────────────────────────────────────────────────────────────

@st.cache_resource
def _supabase_client():
    """Cached Supabase client; credentials come from .streamlit/secrets.toml."""
    from supabase import create_client
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)


def _get_bucket_name() -> str:
    """
    Bucket name from secrets.toml → [supabase] bucket.
    Falls back to 'capaciteitsplanning' if not set.
    Allows each location/deployment to use its own bucket.
    """
    try:
        return st.secrets["supabase"]["bucket"]
    except (KeyError, FileNotFoundError):
        return "capaciteitsplanning"


def _bucket():
    return _supabase_client().storage.from_(_get_bucket_name())


# ── Low-level cloud I/O ────────────────────────────────────────────────────────

def upload_file(filename: str, data: bytes) -> None:
    """Upload (or overwrite) a single file in the Supabase bucket."""
    _bucket().upload(filename, data, file_options={"upsert": "true"})


def download_file(filename: str) -> bytes:
    """Download a file from the Supabase bucket and return its bytes."""
    return _bucket().download(filename)


def list_cloud_files() -> list[str]:
    """Return a list of filenames currently in the bucket."""
    entries = _bucket().list()
    return [e["name"] for e in entries]


# ── Metadata ───────────────────────────────────────────────────────────────────

def save_metadata(info: dict) -> None:
    data = json.dumps(info, indent=2, ensure_ascii=False).encode("utf-8")
    upload_file("metadata.json", data)


def load_metadata() -> dict | None:
    try:
        data = download_file("metadata.json")
        return json.loads(data.decode("utf-8"))
    except Exception:
        return None


# ── Validation helpers (work on any Path folder) ───────────────────────────────

def validate_feestdagen_xlsx(file_path: Path) -> bool:
    df = pd.read_excel(file_path)
    required = {"Datum", "Omschrijving", "Type"}
    if not required.issubset(df.columns):
        raise ValueError(
            "feestdagen.xlsx moet de kolommen Datum, Omschrijving en Type bevatten."
        )
    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
    if df["Datum"].isna().any():
        raise ValueError("feestdagen.xlsx bevat ongeldige datums.")
    return True


def validate_required_files_in_folder(folder: Path) -> bool:
    missing = [f for f in REQUIRED_FILES if not (folder / f).exists()]
    if missing:
        raise FileNotFoundError("Ontbrekende bestanden: " + ", ".join(missing))
    validate_feestdagen_xlsx(folder / "feestdagen.xlsx")
    return True


# ── Publishing (manager → cloud) ───────────────────────────────────────────────

def publish_files(file_map: dict[str, bytes]) -> None:
    """Upload a dict of {filename: bytes} to the cloud bucket."""
    for fname, data in file_map.items():
        upload_file(fname, data)


# ── Loading published data (cloud → temp → DataFrame) ─────────────────────────

def _ensure_temp() -> Path:
    _TEMP_DIR.mkdir(parents=True, exist_ok=True)
    return _TEMP_DIR


def load_published_data():
    """
    Downloads all required files from the cloud bucket into a local temp
    directory, then processes them into DataFrames.

    Returns: (merged, export_file_summary, order_file_name, holiday_df)
    """
    tmp = _ensure_temp()

    # Check which files exist in the cloud
    try:
        available = set(list_cloud_files())
    except Exception as e:
        raise RuntimeError(f"Kan cloud-opslag niet bereiken: {e}")

    missing = [f for f in REQUIRED_FILES if f not in available]
    if missing:
        raise FileNotFoundError(
            "De volgende bestanden zijn nog niet gepubliceerd: "
            + ", ".join(missing)
        )

    # Download each required file to the temp dir
    for fname in REQUIRED_FILES:
        raw = download_file(fname)
        (tmp / fname).write_bytes(raw)

    # --- From here: same processing logic as before ---

    validate_required_files_in_folder(tmp)

    export_files = [
        "Export-1.xlsx",
        "Export.xlsx",
        "Export+1.xlsx",
        "Export+2.xlsx",
        "Export+3.xlsx",
        "Export+4.xlsx",
    ]
    export_frames = []
    export_file_summary = []

    for fname in export_files:
        fp = tmp / fname
        tmp_df = pd.read_excel(fp)
        tmp_df["Bronbestand"] = fname
        tmp_df["Bron_week"] = extract_week_label(fname)
        export_frames.append(tmp_df)
        export_file_summary.append(
            {
                "Bronbestand": fname,
                "Bron_week": extract_week_label(fname),
                "Aantal_regels_ingelezen": len(tmp_df),
            }
        )

    export = pd.concat(export_frames, ignore_index=True)
    order = pd.read_excel(tmp / "OrderExport2G.xlsx")
    holiday_df = pd.read_excel(tmp / "feestdagen.xlsx")

    export["Gewicht_export_kg"] = coerce_numeric(export["Gewicht"])
    export["Leverdatum"] = pd.to_datetime(export["Datum"], dayfirst=True, errors="coerce")
    export["Verzinkstatus"] = export["Status"].map(STATUS_MAP)
    export["Ordernummer_base"] = coerce_numeric(
        export["Nummer"].astype(str).str.extract(r"(\d+)")[0]
    )

    order["Ordernummer"] = coerce_numeric(order["Ordernummer"])
    order["Gewicht_order_kg"] = coerce_numeric(order["Gewicht(ton)"]) * 1000

    row_counts = (
        export.groupby("Ordernummer_base", dropna=False)
        .size()
        .rename("Regels_per_order")
        .reset_index()
    )
    merged = export.merge(row_counts, on="Ordernummer_base", how="left")
    merged = merged.merge(
        order[["Ordernummer", "Gewicht_order_kg"]],
        left_on="Ordernummer_base",
        right_on="Ordernummer",
        how="left",
    )
    merged["Gewicht_2g_verdeeld_kg"] = (
        merged["Gewicht_order_kg"] / merged["Regels_per_order"]
    )
    merged["Gewicht_bron"] = np.where(
        merged["Gewicht_export_kg"].fillna(0) > 0,
        "Export+",
        "OrderExport2G verdeeld",
    )
    merged["Gewicht_effectief_kg"] = np.where(
        merged["Gewicht_export_kg"].fillna(0) > 0,
        merged["Gewicht_export_kg"],
        merged["Gewicht_2g_verdeeld_kg"],
    )

    holiday_df["Datum"] = pd.to_datetime(holiday_df["Datum"], errors="coerce").dt.date
    holiday_df = (
        holiday_df.dropna(subset=["Datum"])
        .drop_duplicates(subset=["Datum", "Omschrijving", "Type"])
        .sort_values("Datum")
        .reset_index(drop=True)
    )

    return merged, pd.DataFrame(export_file_summary), "OrderExport2G.xlsx", holiday_df


# ── Date / calendar helpers ────────────────────────────────────────────────────

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


# ── Formatting helpers ─────────────────────────────────────────────────────────

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


# ── Dashboard builder ──────────────────────────────────────────────────────────

def build_horizon_and_include_holidays(
    start_dt: pd.Timestamp, production_workdays: int, holiday_dates: set
):
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

    start_ts = pd.Timestamp(startdatum).normalize()
    today_ts = pd.Timestamp(date.today()).normalize()

    df["Meegeteld_in_planning"] = "Nee"
    df["Reden_uitsluiting"] = ""

    mask_niet_verzinkt = df["Verzinkstatus"] == "Niet verzinkt"
    mask_ub = df["Verzinkstatus"] == "UB"
    mask_verzinkt = df["Verzinkstatus"] == "Verzinkt"
    mask_binnen_horizon = df["Verzinkdatum"] >= start_ts

    df.loc[mask_verzinkt, "Reden_uitsluiting"] = "Status = Verzinkt"
    df.loc[mask_ub, "Reden_uitsluiting"] = "Status = UB"
    df.loc[mask_niet_verzinkt & ~mask_binnen_horizon, "Reden_uitsluiting"] = (
        "Voor startdatum rapport"
    )
    df.loc[mask_niet_verzinkt & mask_binnen_horizon, "Meegeteld_in_planning"] = "Ja"
    df.loc[mask_niet_verzinkt & mask_binnen_horizon, "Reden_uitsluiting"] = ""

    df_plan = df[df["Meegeteld_in_planning"] == "Ja"].copy()

    horizon = build_horizon_and_include_holidays(start_ts, 10, holiday_dates)

    dag_orders = df_plan.groupby("Verzinkdatum", as_index=False).agg(
        Gewicht_kg=("Gewicht_effectief_kg", "sum"),
        Aantal_orders_te_verzinken=("Nummer", "count"),
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
    dag["Gewicht_kg"] = dag["Gewicht_kg"].fillna(0.0)
    dag["Aantal_orders_te_verzinken"] = dag["Aantal_orders_te_verzinken"].fillna(0).astype(int)
    dag["Capaciteit_kg"] = np.where(dag["Is_feestdag_of_sluiting"], 0, capaciteit_kg)
    dag["Benutting_pct"] = np.where(
        dag["Capaciteit_kg"] > 0,
        (dag["Gewicht_kg"] / dag["Capaciteit_kg"]) * 100,
        np.nan,
    )
    dag["Traverses_berekend"] = np.ceil(dag["Gewicht_kg"] / kg_per_traverse)
    dag["Status"] = dag.apply(
        lambda r: stoplight(r["Benutting_pct"], r["Is_feestdag_of_sluiting"]), axis=1
    )
    dag["Jaar"] = dag["Verzinkdatum"].dt.isocalendar().year.astype(int)
    dag["Week"] = dag["Verzinkdatum"].dt.isocalendar().week.astype(int)
    dag["Label_nl"] = dag["Verzinkdatum"].apply(format_nl_axis_label)
    dag["Dagtype"] = np.where(
        dag["Is_feestdag_of_sluiting"], "Feestdag / sluiting", "Werkdag"
    )

    week = dag.groupby(["Jaar", "Week"], as_index=False).agg(
        Gewicht_kg=("Gewicht_kg", "sum"),
        Aantal_orders_te_verzinken=("Aantal_orders_te_verzinken", "sum"),
        Traverses_berekend=("Traverses_berekend", "sum"),
        Capaciteit_kg=("Capaciteit_kg", "sum"),
    )
    week["Benutting_pct"] = np.where(
        week["Capaciteit_kg"] > 0,
        (week["Gewicht_kg"] / week["Capaciteit_kg"]) * 100,
        np.nan,
    )
    week["Status"] = week["Benutting_pct"].apply(lambda x: stoplight(x, False))

    def calculate_advice_date(day_df, today_date, holiday_dates):
        base_date = add_workdays(today_date, 5, holiday_dates)
        productive = day_df[day_df["Is_feestdag_of_sluiting"] == False].copy()
        row = productive[productive["Verzinkdatum"] == base_date]
        if not row.empty and float(row.iloc[0]["Benutting_pct"]) <= 95:
            return base_date
        later = productive[
            (productive["Verzinkdatum"] > base_date)
            & (productive["Benutting_pct"] < 95)
        ]
        if not later.empty:
            return pd.Timestamp(later.iloc[0]["Verzinkdatum"])
        return base_date

    advies_datum = calculate_advice_date(dag, today_ts, holiday_dates)
    return df, df_plan, dag, week, advies_datum
