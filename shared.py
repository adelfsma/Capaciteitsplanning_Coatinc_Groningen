import json
import re
import tempfile
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

APP_VERSION = "v2.5.8"


def get_app_environment() -> str:
    """Lees de app-omgeving uit Streamlit Secrets.

    Verwachte configuratie in Streamlit Cloud:

    [app]
    environment = "test"        # toont TEST-banner

    of:

    [app]
    environment = "production"  # geen TEST-banner

    Ontbreekt deze instelling, dan gedraagt de app zich als productie en wordt
    er geen TEST-banner getoond.
    """
    try:
        app_cfg = st.secrets.get("app", {})
        value = app_cfg.get("environment", "") if hasattr(app_cfg, "get") else ""
    except Exception:
        value = ""
    return str(value).strip().lower()


def is_test_environment() -> bool:
    """True als de app expliciet als testomgeving is ingesteld."""
    return get_app_environment() in {"test", "testing", "acceptatie", "staging"}


def get_page_title(base_title: str) -> str:
    """Voeg [TEST] toe aan de browser-tab als environment=test."""
    return f"[TEST] {base_title}" if is_test_environment() else base_title


def render_environment_banner(page_label: str = ""):
    """Toon alleen een duidelijke TEST-markering als [app] environment='test'."""
    if not is_test_environment():
        return

    suffix = f" – {page_label}" if page_label else ""
    st.sidebar.error("🔴 TESTOMGEVING")
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(90deg, #7f1d1d 0%, #dc2626 50%, #7f1d1d 100%);
            color: white;
            border: 4px solid #450a0a;
            border-radius: 14px;
            padding: 18px 24px;
            margin: 0 0 18px 0;
            text-align: center;
            box-shadow: 0 4px 14px rgba(127, 29, 29, 0.25);
        ">
            <div style="font-size: 52px; line-height: 1; font-weight: 900; letter-spacing: 0.16em;">
                TEST
            </div>
            <div style="font-size: 20px; font-weight: 700; margin-top: 8px; letter-spacing: 0.02em;">
                TESTOMGEVING{suffix} — niet gebruiken als productieversie
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

STATUS_MAP = {
    # Afgerond / opgehaald
    "uitgeleverd": "Verzinkt",
    "Afgehaald": "Verzinkt",
    "PC Afgehaald": "Verzinkt",
    "coat gereed": "Verzinkt",
    "Gereed": "Verzinkt",
    "Nabewerking nog uitvoeren": "Verzinkt",
    "Ontzinkt": "Verzinkt",                 # ontzinkt telt niet als 'nog te verzinken'
    # UB (uitbesteed)
    "UB": "UB",
    "UB V Gereed": "UB",
    # Nog te verzinken
    "Opgehangen": "Niet verzinkt",
    "PC Opgehangen": "Niet verzinkt",       # zelfde als Opgehangen
    "Voorbewerking uitvoeren": "Niet verzinkt",
    "Productie gereed": "Niet verzinkt",
    "Gereserveerd*": "Niet verzinkt",       # reserveringen meenemen in planning
    "Geblokkeerd": "Niet verzinkt",         # geblokkeerd maar nog niet verzinkt
    # Overig
    "meetrapport": "Verzinkt",              # administratief, geen capaciteitsimpact
    "Niet gezien": "Niet verzinkt",         # aanwezig maar nog niet ingepland
    "Binnengemeld": "Niet verzinkt",         # materiaal aangemeld, nog niet verzinkt
    "Gelost": "Niet verzinkt",               # materiaal gelost/ontvangen, nog niet verzinkt
}

# ── Zwarte voorraad (KPI) ───────────────────────────────────────────────────
# Mapping van de ruwe Status-waarde naar de subcategorie die meetelt in de
# KPI "Totale zwarte voorraad". Deze KPI is de som van de twee subcategorieën
# hieronder. Statussen die hier niet in voorkomen tellen niet mee (o.a. alles
# wat al verzinkt/UB is, en Reserveringen).
ZWARTE_VOORRAAD_MAP = {
    "Productie gereed":        "Productie gereed",
    "Voorbewerking uitvoeren": "Binnengemeld/Voorbewerking/Geblokkeerd",
    "Binnengemeld":            "Binnengemeld/Voorbewerking/Geblokkeerd",
    "Geblokkeerd":             "Binnengemeld/Voorbewerking/Geblokkeerd",
    "Gelost":                  "Binnengemeld/Voorbewerking/Geblokkeerd",
}

# Witte voorraad (KPI): materiaal dat al verzinkt/gereed is, maar administratief
# of logistiek nog in de voorraadstroom zit. De matching in de viewer gebeurt
# case-insensitive, zodat varianten als "coat gereed" en "Coat gereed" meetellen.
WITTE_VOORRAAD_STATUSSEN = {
    "Afgehaald",
    "Nabewerking nog uitvoeren",
    "PC Afgehaald",
    "Coat gereed",
}

# ── OTIF (On Time In Full) ────────────────────────────────────────────────────
# Mapping van de ruwe Status-waarde naar de OTIF-uitkomst per order.
#   - 'Ja'      → order is gereed op tijd
#   - 'Nee'     → order is (nog) niet gereed en telt dus als 'te laat'
#   - 'Nvt'     → order valt buiten de OTIF-noemer (bijv. al opgehaald PC)
# Statussen die niet in deze map staan krijgen 'Onbekend'; die tellen mee in
# het totaal (conform gebruikersformule) maar hebben geen impact op 'te laat'.
# Matching gebeurt case-insensitief.
OTIF_STATUS_MAP = {
    "uitgeleverd":               "Ja",
    "afgehaald":                 "Ja",
    "coat gereed":               "Ja",
    "ub v gereed":               "Ja",
    "pc afgehaald":              "Nvt",
    "productie gereed":          "Nee",
    "geblokkeerd":               "Nee",
    "opgehangen":                "Nee",
    "pc opgehangen":             "Nee",
    "ub":                        "Nee",
    "nabewerking nog uitvoeren": "Nee",
    "meetrapport":               "Nee",
}

# Drempels voor de OTIF-snelheidsmeter in het dashboard.
OTIF_GAUGE_GREEN_THRESHOLD  = 96.0   # ≥ dit percentage: groen
OTIF_GAUGE_ORANGE_THRESHOLD = 80.0   # ≥ dit percentage: oranje (anders rood)

# Kandidaat-kolomnamen voor de 'originele leverdatum' in de weekexports.
# De echte kolomnaam wisselt per bronsysteem; in de testomgeving kan tegen
# deze datum vergeleken worden in plaats van tegen 'Datum'.
ORIGINELE_LEVERDATUM_CANDIDATES = [
    "Originele datum",
    "Oorspronkelijke datum",
    "Originele leverdatum",
    "Oorspronkelijke leverdatum",
    "OrigineleDatum",
    "OrigineleLeverdatum",
    "Datum origineel",
    "Origineel Datum",
    "Origineel leverdatum",
]

NL_DAY_ABBR = {0: "ma", 1: "di", 2: "wo", 3: "do", 4: "vr", 5: "za", 6: "zo"}

# Standaard (Groningen) vereiste bestanden. Per vestiging kan een kleinere
# vereiste set gelden (zie _compute_file_lists); OrderExport2G.xlsx is altijd
# vereist. De uiteindelijke REQUIRED_FILES/OPTIONAL_FILES worden lager in dit
# bestand berekend, zodra de locatie-config beschikbaar is.
_DEFAULT_REQUIRED_FILES = [
    "OrderExport2G.xlsx",
    "Export-1.xlsx",
    "Export.xlsx",
    "Export+1.xlsx",
    "Export+2.xlsx",
    "Export+3.xlsx",
    "Export+4.xlsx",
    "feestdagen.xlsx",
]

# Optioneel: CGS-planningsexport (verzinken productielijn) en debiteurenexport.
# Als aanwezig, worden CGS-orders die NIET al in de weekexports zitten toegevoegd.
# Als debtor-export.xlsx aanwezig is, wordt het klantsegment/materialtype gekoppeld.
DEBTOR_EXPORT_FILE = "debtor-export.xlsx"
# Klanten met 24-uursservice (scheepsleidingen). Als dit bestand aanwezig is,
# worden deze klanten (op klantnummer) altijd als 'Aanleveren depot = 1'
# behandeld. Optioneel en per vestiging: ontbreekt het, dan gebeurt er niets.
SCHEEPSLEIDINGEN_FILE = "scheepsleidingen.xlsx"
_BASE_OPTIONAL_FILES = ["Export_CGS.xlsx", DEBTOR_EXPORT_FILE, SCHEEPSLEIDINGEN_FILE]

OPTIONAL_FILE_LABELS = {
    "Export_CGS.xlsx": "Upload Export_CGS.xlsx (optioneel – CGS verzinkplanning)",
    DEBTOR_EXPORT_FILE: "Upload debtor-export.xlsx (optioneel – segment/type materiaal per klant)",
    SCHEEPSLEIDINGEN_FILE: "Upload scheepsleidingen.xlsx (optioneel – klanten met 24-uursservice; worden als depot behandeld)",
    "feestdagen.xlsx": "Upload feestdagen.xlsx (optioneel voor deze vestiging – feestdagenkalender)",
    "Export+1.xlsx": "Upload Export+1.xlsx (optioneel voor deze vestiging)",
    "Export+2.xlsx": "Upload Export+2.xlsx (optioneel voor deze vestiging)",
    "Export+3.xlsx": "Upload Export+3.xlsx (optioneel voor deze vestiging)",
    "Export+4.xlsx": "Upload Export+4.xlsx (optioneel voor deze vestiging)",
}

MATERIAALTYPE_ORDER = ["Constructie", "Maatwerk", "Seriewerk", "Overig / onbekend"]
MATERIAALTYPE_DAG_COLS = {
    "Constructie": "KG_Constructie",
    "Maatwerk": "KG_Maatwerk",
    "Seriewerk": "KG_Seriewerk",
    "Overig / onbekend": "KG_Overig_onbekend",
}

# ── Reserveringen (Power BI Merge1-logica) ─────────────────────────────────────
# Orders in OrderExport2G die nog NIET in de CGS-exports staan,
# binnen het planningsvenster en voor de juiste locatie.
RESERVERING_WINDOW_VOOR = 2    # dagen vóór vandaag
RESERVERING_WINDOW_NA   = 40   # dagen na vandaag


# ── Vestiging-specifieke configuratie ──────────────────────────────────────────
# Alles wat per vestiging verschilt komt uit secrets.toml → [locatie].
# Zo bedient één codebase meerdere vestigingen. Ontbreekt de sectie of een veld,
# dan valt de waarde terug op de Groningse standaard (backwards compatible).
#
# [locatie]
# naam                = "Coatinc Groningen"           # titels
# logo                = "logo_coatinc_groningen.png"  # bestand in repo-root
# reservering_locatie = "Coatinc Groningen"           # exacte match op 'Locatie V'
# beheer_wachtwoord   = "coatinc2026"                 # wachtwoord beheeromgeving
# poetsen_otif_uitzondering = false                   # true = 'Afgehaald' telt niet OK voor poetsen-orders
# max_capaciteit_min = 50                             # dagcapaciteit-slider (ton): minimum
# max_capaciteit_max = 90                             # dagcapaciteit-slider (ton): maximum
# max_capaciteit_default = 60                         # dagcapaciteit-slider (ton): startwaarde
# kg_traverse_constructie = 1300                      # default KG per traverse Constructie
# kg_traverse_maatwerk = 840                          # default KG per traverse Maatwerk
# kg_traverse_seriewerk = 930                         # default KG per traverse Seriewerk
# otif_uitsluiten_statussen = ""                      # komma-gescheiden statussen buiten OTIF, bijv. "UB"
# verplichte_bestanden = ""                           # kleinere vereiste set, bijv. "OrderExport2G.xlsx, Export-1.xlsx, Export.xlsx"

_LOCATIE_DEFAULTS = {
    "naam": "Coatinc Groningen",
    "logo": "logo_coatinc_groningen.png",
    "beheer_wachtwoord": "coatinc2026",
}


def _get_locatie_setting(key: str, default: str) -> str:
    try:
        cfg = st.secrets["locatie"]
        value = cfg.get(key, "") if hasattr(cfg, "get") else ""
    except (KeyError, FileNotFoundError):
        value = ""
    value = str(value).strip()
    return value if value else default


def get_locatie_naam() -> str:
    """Weergavenaam van de vestiging (gebruikt in de paginatitels)."""
    return _get_locatie_setting("naam", _LOCATIE_DEFAULTS["naam"])


def get_locatie_logo() -> str:
    """Bestandsnaam van het logo in de repo-root."""
    return _get_locatie_setting("logo", _LOCATIE_DEFAULTS["logo"])


def get_beheer_wachtwoord() -> str:
    """Wachtwoord voor de beheeromgeving."""
    return _get_locatie_setting("beheer_wachtwoord", _LOCATIE_DEFAULTS["beheer_wachtwoord"])


def get_reservering_locatie() -> str:
    """Waarde waarop 'Locatie V' exact moet matchen voor reserveringen.
    Valt terug op de vestigingsnaam als niet apart gezet."""
    return _get_locatie_setting("reservering_locatie", get_locatie_naam())


def get_poetsen_otif_actief() -> bool:
    """True als de poetsen-uitzondering op OTIF actief is voor deze vestiging.
    Zet in secrets: [locatie] poetsen_otif_uitzondering = true.
    Effect: voor orders waarvan PrijsCategorie 'poetsen' bevat, telt de status
    'Afgehaald' op de peildatum als niet OK (i.p.v. OK). Overige statussen
    volgen de normale OTIF-regels. Standaard uit (Groningen ongewijzigd)."""
    val = _get_locatie_setting("poetsen_otif_uitzondering", "false")
    return str(val).strip().lower() in ("true", "1", "ja", "yes", "aan")


def _get_locatie_number(key: str, default):
    """Lees een numerieke [locatie]-instelling; val terug op default bij ontbreken
    of ongeldige waarde. Geeft int terug als de waarde geheel is, anders float."""
    try:
        cfg = st.secrets["locatie"]
        raw = cfg.get(key, None) if hasattr(cfg, "get") else None
    except (KeyError, FileNotFoundError):
        raw = None
    if raw is None or (isinstance(raw, str) and not raw.strip()):
        return default
    try:
        f = float(raw)
    except (TypeError, ValueError):
        return default
    return int(f) if f.is_integer() else f


def get_max_capaciteit_config() -> tuple[int, int, int, int]:
    """(min, max, default, step) voor de dagcapaciteit-slider (ton).
    Groningen-defaults: 50–90, start 60, stap 5. Per vestiging instelbaar via
    [locatie]: max_capaciteit_min / _max / _default / _step."""
    mn = int(_get_locatie_number("max_capaciteit_min", 50))
    mx = int(_get_locatie_number("max_capaciteit_max", 90))
    step = int(_get_locatie_number("max_capaciteit_step", 5))
    default = int(_get_locatie_number("max_capaciteit_default", 60))
    if mx < mn:
        mn, mx = mx, mn
    default = min(max(default, mn), mx)   # binnen [min, max] houden
    return mn, mx, default, max(step, 1)


def get_kg_traverse_defaults() -> tuple[int, int, int]:
    """Default KG-per-traverse voor (Constructie, Maatwerk, Seriewerk).
    Groningen-defaults 1300 / 840 / 930. Per vestiging instelbaar via [locatie]:
    kg_traverse_constructie / _maatwerk / _seriewerk."""
    return (
        int(_get_locatie_number("kg_traverse_constructie", 1300)),
        int(_get_locatie_number("kg_traverse_maatwerk", 840)),
        int(_get_locatie_number("kg_traverse_seriewerk", 930)),
    )


def get_otif_uitsluiten_statussen() -> set:
    """Statussen die volledig buiten de OTIF-berekening blijven (niet in totaal
    en niet als te laat). Per vestiging instelbaar via [locatie]:
    otif_uitsluiten_statussen = "UB"  (komma-gescheiden voor meerdere).
    Standaard leeg (Groningen ongewijzigd)."""
    raw = _get_locatie_setting("otif_uitsluiten_statussen", "")
    if not raw:
        return set()
    return {s.strip() for s in str(raw).split(",") if s.strip()}


def _split_required_optional(raw: str) -> tuple[list, list]:
    """Pure helper: bepaal (required, optional) uit de config-string
    'verplichte_bestanden'. Leeg → Groningen-defaults."""
    if not raw or not str(raw).strip():
        return list(_DEFAULT_REQUIRED_FILES), list(_BASE_OPTIONAL_FILES)

    wanted = {s.strip() for s in str(raw).split(",") if s.strip()}
    required = [f for f in _DEFAULT_REQUIRED_FILES if f in wanted]
    if "OrderExport2G.xlsx" not in required:
        required = ["OrderExport2G.xlsx"] + required
    demoted = [f for f in _DEFAULT_REQUIRED_FILES if f not in required]
    optional = demoted + list(_BASE_OPTIONAL_FILES)
    return required, optional


def _compute_file_lists() -> tuple[list, list]:
    """Bepaal (REQUIRED_FILES, OPTIONAL_FILES) voor deze vestiging.

    Standaard (Groningen): alle weekexports + feestdagen vereist. Per vestiging
    kan een kleinere vereiste set gelden via [locatie]:
        verplichte_bestanden = "OrderExport2G.xlsx, Export-1.xlsx, Export.xlsx"
    De overige standaard-vereiste bestanden verschuiven dan naar optioneel.
    OrderExport2G.xlsx blijft altijd vereist (harde afhankelijkheid)."""
    return _split_required_optional(_get_locatie_setting("verplichte_bestanden", ""))


# Definitieve, per-vestiging bepaalde bestandslijsten (door manager- en viewer-app
# geïmporteerd). Berekend bij import; per app/vestiging vast via de secrets.
REQUIRED_FILES, OPTIONAL_FILES = _compute_file_lists()


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
    # feestdagen kan per vestiging optioneel zijn; alleen valideren indien aanwezig.
    if (folder / "feestdagen.xlsx").exists():
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


# ── Debtor segment / materiaaltype helpers ─────────────────────────────────────

def normalize_materiaaltype(segment) -> str:
    """
    Normaliseer het klantsegment uit debtor-export naar de materiaaltype-indeling
    die in het dashboard wordt getoond.
    """
    if pd.isna(segment):
        return "Overig / onbekend"

    s = str(segment).strip().lower()
    if not s:
        return "Overig / onbekend"
    if "constructie" in s:
        return "Constructie"
    if "maatwerk" in s:
        return "Maatwerk"
    if "serie" in s:
        return "Seriewerk"
    return "Overig / onbekend"


def _normalize_name_key(series: pd.Series) -> pd.Series:
    """Maak een robuuste sleutel voor klantnamen."""
    s = series.fillna("").astype(str).str.lower().str.strip()
    s = s.str.replace(r"[^a-z0-9]+", " ", regex=True)
    s = s.str.replace(r"\b(bv|b v|b\.v|b\.v\.|nv|n v|n\.v|holding|groep|group)\b", "", regex=True)
    s = s.str.replace(r"\s+", " ", regex=True).str.strip()
    return s


def _normalize_number_key(series: pd.Series) -> pd.Series:
    """Maak een robuuste sleutel voor debiteurnummers."""
    numeric = coerce_numeric(series)
    out = numeric.round(0).astype("Int64").astype(str)
    return out.replace({"<NA>": ""})


def _find_first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    exact = {c.lower().strip(): c for c in df.columns}
    for candidate in candidates:
        if candidate.lower().strip() in exact:
            return exact[candidate.lower().strip()]
    return None


def validate_debtor_export_xlsx(file_path: Path) -> bool:
    """Valideer de optionele debtor-export voor segmentkoppeling."""
    try:
        debtors = pd.read_excel(file_path, sheet_name="debtors", nrows=5)
    except ValueError as e:
        raise ValueError("debtor-export.xlsx moet een tabblad 'debtors' bevatten.") from e

    required = {"number", "name", "segment"}
    missing = required.difference(debtors.columns)
    if missing:
        raise ValueError(
            "debtor-export.xlsx moet in tabblad 'debtors' de kolommen "
            + ", ".join(sorted(required))
            + " bevatten. Ontbrekend: "
            + ", ".join(sorted(missing))
        )
    return True


def _load_debtor_segment_lookup(tmp: Path) -> pd.DataFrame:
    """
    Lees debtor-export.xlsx en bouw een lookup op debiteurnummer en klantnaam.
    Het bestand is optioneel: bij ontbreken blijven segmenten 'Overig / onbekend'.
    """
    debtor_path = tmp / DEBTOR_EXPORT_FILE
    if not debtor_path.exists():
        return pd.DataFrame()

    validate_debtor_export_xlsx(debtor_path)
    debtors = pd.read_excel(debtor_path, sheet_name="debtors")

    lookup = debtors[["number", "name", "segment"]].copy()
    lookup["Debiteurnummer_key"] = _normalize_number_key(lookup["number"])
    lookup["Klantnaam_key"] = _normalize_name_key(lookup["name"])
    lookup["Segment_debtor_export"] = lookup["segment"].fillna("").astype(str).str.strip()
    lookup["Materiaaltype"] = lookup["segment"].apply(normalize_materiaaltype)

    if "is_active" in debtors.columns:
        lookup["_active_rank"] = debtors["is_active"].astype(str).str.lower().eq("active").astype(int)
    else:
        lookup["_active_rank"] = 0

    lookup = lookup.sort_values("_active_rank", ascending=False)
    lookup = lookup.drop_duplicates(subset=["Debiteurnummer_key"], keep="first")
    lookup_name = lookup.drop_duplicates(subset=["Klantnaam_key"], keep="first")
    lookup = pd.concat([lookup, lookup_name], ignore_index=True).drop_duplicates()
    return lookup[["Debiteurnummer_key", "Klantnaam_key", "Segment_debtor_export", "Materiaaltype"]]


def _load_scheepsleidingen_ids(tmp: Path) -> set:
    """
    Lees scheepsleidingen.xlsx en geef de set genormaliseerde klantnummers van
    klanten met 24-uursservice. Deze klanten worden verderop altijd als
    'Aanleveren depot = 1' behandeld.

    Het bestand is optioneel: ontbreekt het, dan een lege set (geen effect).
    Als de kolom 'is_active' aanwezig is, tellen alleen actieve klanten mee.
    Het klantnummer wordt flexibel gezocht (kolom 'number' of varianten).
    """
    path = tmp / SCHEEPSLEIDINGEN_FILE
    if not path.exists():
        return set()

    df = pd.read_excel(path)
    num_col = _find_first_existing_column(
        df,
        ["number", "Debiteurnummer", "ID Debiteur", "Klantnummer", "customer_number", "debtor_number"],
    )
    if not num_col:
        return set()

    if "is_active" in df.columns:
        df = df[df["is_active"].astype(str).str.strip().str.lower() == "active"]

    keys = _normalize_number_key(df[num_col])
    return {k for k in keys if k}


def _force_scheepsleidingen_depot(order: pd.DataFrame, scheeps_ids: set) -> pd.DataFrame:
    """
    Zet 'Aanleveren depot' op 1 voor orders van scheepsleidingen-klanten
    (24-uursservice), gematcht op klantnummer. Retourneert het (aangepaste)
    order-DataFrame. Lege set of ontbrekende kolommen → ongewijzigd.
    """
    if not scheeps_ids or order is None or order.empty:
        return order
    if "Aanleveren depot" not in order.columns:
        return order
    num_col = _find_first_existing_column(
        order,
        ["ID Debiteur", "Debiteurnummer", "Klantnummer", "customer_number", "number"],
    )
    if not num_col:
        return order
    is_scheeps = _normalize_number_key(order[num_col]).isin(scheeps_ids)
    order.loc[is_scheeps, "Aanleveren depot"] = 1
    return order


def _attach_debtor_segments(df: pd.DataFrame, debtor_lookup: pd.DataFrame) -> pd.DataFrame:
    """Voeg Segment_debtor_export en Materiaaltype toe aan orderregels."""
    out = df.copy()
    out["Segment_debtor_export"] = ""
    out["Materiaaltype"] = "Overig / onbekend"

    if debtor_lookup.empty:
        return out

    number_candidates = [
        "Debiteurnummer", "Debiteur nummer", "Debiteur_nummer", "DebiteurNr", "Debiteur nr",
        "Klantnummer", "Klant nummer", "customer_number", "debtor_number", "number",
    ]
    name_candidates = ["Klantnaam", "Debiteur", "Debiteurnaam", "name"]

    matched = pd.Series(False, index=out.index)

    debtor_number_col = _find_first_existing_column(out, number_candidates)
    if debtor_number_col:
        number_lookup = (
            debtor_lookup[debtor_lookup["Debiteurnummer_key"].astype(str).str.strip() != ""]
            .drop_duplicates(subset=["Debiteurnummer_key"])
            .set_index("Debiteurnummer_key")
        )
        keys = _normalize_number_key(out[debtor_number_col])
        seg_raw = keys.map(number_lookup["Segment_debtor_export"])
        mat = keys.map(number_lookup["Materiaaltype"])
        mask = mat.notna()
        out.loc[mask, "Segment_debtor_export"] = seg_raw[mask].fillna("").values
        out.loc[mask, "Materiaaltype"] = mat[mask].values
        matched = matched | mask

    debtor_name_col = _find_first_existing_column(out, name_candidates)
    if debtor_name_col:
        name_lookup = (
            debtor_lookup[debtor_lookup["Klantnaam_key"].astype(str).str.strip() != ""]
            .drop_duplicates(subset=["Klantnaam_key"])
            .set_index("Klantnaam_key")
        )
        keys = _normalize_name_key(out[debtor_name_col])
        seg_raw = keys.map(name_lookup["Segment_debtor_export"])
        mat = keys.map(name_lookup["Materiaaltype"])
        mask = mat.notna() & ~matched
        out.loc[mask, "Segment_debtor_export"] = seg_raw[mask].fillna("").values
        out.loc[mask, "Materiaaltype"] = mat[mask].values

    out["Materiaaltype"] = out["Materiaaltype"].fillna("Overig / onbekend").apply(normalize_materiaaltype)
    return out


def _load_cgs_as_export_rows(tmp: Path, existing_cgs_ids: set) -> tuple[pd.DataFrame, int]:
    """
    Lees Export_CGS.xlsx en zet de rijen om naar hetzelfde formaat als de weekexports.
    Alleen orders met ProductieLijn=VZ die NIET al in de weekexports zitten worden toegevoegd.
    Geeft (dataframe, aantal_toegevoegd) terug.
    """
    cgs_path = tmp / "Export_CGS.xlsx"
    if not cgs_path.exists():
        return pd.DataFrame(), 0

    cgs = pd.read_excel(cgs_path)

    # Alleen verzinken orders die nog niet in exports zitten
    cgs = cgs[cgs["ProduktieLijn"] == "VZ"].copy()
    cgs = cgs[~cgs["OrderID"].isin(existing_cgs_ids)].copy()

    if cgs.empty:
        return pd.DataFrame(), 0

    # Status bepalen op basis van StatusDiensten-string
    def _cgs_status(s: str) -> str:
        s = str(s)
        if "R" in s[13:16]:
            return "Gereserveerd*"
        if "U" in s[13:16]:
            return "Productie gereed"
        return "Productie gereed"

    rows = []
    for _, r in cgs.iterrows():
        rows.append({
            "Status": _cgs_status(r["StatusDiensten"]),
            "Nummer": str(r.get("ReferentieKlant", r["OrderID"])),
            "Datum": r["LeverDatum"],
            "Gewicht": r["Gewicht"],
            "Bronbestand": "Export_CGS.xlsx",
            "Bron_week": "CGS",
            "CgsNummer": r["OrderID"],
        })

    df = pd.DataFrame(rows)
    return df, len(df)


def _parse_order_date(series: pd.Series) -> pd.Series:
    """
    Parseer datumkolommen uit OrderExport2G.
    Ondersteunt Excel-datums, tekst-datums en integer YYYYMMDD.
    """
    try:
        as_str = series.dropna().astype(str).str.strip()
        if len(as_str) > 0 and as_str.str.match(r"^\d{8}$").all():
            return pd.to_datetime(series.astype(str), format="%Y%m%d", errors="coerce")
        return pd.to_datetime(series, dayfirst=True, errors="coerce")
    except Exception:
        return pd.to_datetime(series, errors="coerce")


def _build_reserveringen(order: pd.DataFrame, cgs_ordernummers: set) -> pd.DataFrame:
    """
    Identificeer reserveringen: orders in OrderExport2G die nog NIET in de
    CGS-exports staan, binnen het planningsvenster en uitsluitend voor
    Locatie V = get_reservering_locatie() (per vestiging via secrets).

    Aangepaste reserveringslogica:
      - Alleen meenemen als Locatie V exact gelijk is aan de ingestelde
        vestigingswaarde (get_reservering_locatie()) na trimmen.
        Lege Locatie V of ontbrekende Locatie V wordt niet meegenomen.
      - Verzinkdatum reservering:
          1. Als Leverdatum V is gevuld: Leverdatum V - 2 werkdagen.
          2. Als Leverdatum V leeg is: Datum verzending + 2 werkdagen.
        Dus nooit Datum verzending - 2 werkdagen.
    """
    today = pd.Timestamp(date.today()).normalize()
    window_start = today - timedelta(days=RESERVERING_WINDOW_VOOR)
    window_end = today + timedelta(days=RESERVERING_WINDOW_NA)

    order = order.copy()

    if "Datum verzending" not in order.columns:
        return pd.DataFrame()

    order["Datum_verzending"] = _parse_order_date(order["Datum verzending"])

    # Zoek de kolom "Leverdatum V" flexibel.
    leverdatum_v_col = None
    for candidate in ["Leverdatum V", "LeverdatumV", "Leverdatum_V"]:
        if candidate in order.columns:
            leverdatum_v_col = candidate
            break

    if leverdatum_v_col:
        order["Leverdatum_V"] = _parse_order_date(order[leverdatum_v_col])
    else:
        order["Leverdatum_V"] = pd.NaT

    # Zoek de kolom "Locatie V" flexibel.
    locatie_col = None
    for candidate in ["Locatie V", "LocatieV", "Locatie_V"]:
        if candidate in order.columns:
            locatie_col = candidate
            break

    # Locatie V is verplicht voor reserveringen. Als deze kolom ontbreekt,
    # nemen we geen reserveringen mee om foutieve locatiebelasting te voorkomen.
    if not locatie_col:
        return pd.DataFrame()

    locatie_match = order[locatie_col].astype(str).str.strip().eq(get_reservering_locatie())

    # Basisfilter: niet in CGS-exports + binnen tijdvenster + juiste locatie.
    # Het tijdvenster blijft gebaseerd op Datum verzending, conform bronexport.
    mask = (
        (~order["Ordernummer"].isin(cgs_ordernummers))
        & (order["Datum_verzending"] >= window_start)
        & (order["Datum_verzending"] <= window_end)
        & locatie_match
    )

    reserveringen = order[mask].copy()
    if reserveringen.empty:
        return pd.DataFrame()

    # Bereken verzinkdatum voor reserveringen volgens aangepaste businessregel.
    # Depot Amsterdam (Aanleveren depot=1): Leverdatum V - 1 / Datum verzending + 1
    # Overig         (Aanleveren depot=0): Leverdatum V - 2 / Datum verzending + 3
    def _calc_reservering_verzinkdatum(r):
        is_depot = int(r.get("Aanleveren depot", 0)) == 1
        if pd.notna(r["Leverdatum_V"]):
            if is_depot:
                return subtract_workdays_existing_orders(r["Leverdatum_V"], 1)
            else:
                return pd.NaT  # geen override → build_dashboard_data gebruikt sidebar offset
        if pd.notna(r["Datum_verzending"]):
            if is_depot:
                return add_workdays_existing_orders(r["Datum_verzending"], 1)
            else:
                return pd.NaT  # geen override → Leverdatum=DV+5, dan sidebar offset
        return pd.NaT

    reserveringen["Verzinkdatum_reservering"] = reserveringen.apply(
        _calc_reservering_verzinkdatum, axis=1
    )

    # Voor display en aansluiting op bestaande structuur:
    # Leverdatum = Leverdatum V indien gevuld, anders Datum verzending.
    def _calc_leverdatum_basis(r):
        if pd.notna(r["Leverdatum_V"]):
            return r["Leverdatum_V"]
        is_depot = int(r.get("Aanleveren depot", 0)) == 1
        if pd.notna(r["Datum_verzending"]):
            # Depot=1: Leverdatum = DV + 2 werkdagen
            # Depot=0: Leverdatum = DV + 5 werkdagen
            return add_workdays_existing_orders(
                r["Datum_verzending"], 2 if is_depot else 5
            )
        return pd.NaT

    reserveringen["Leverdatum_reservering_basis"] = reserveringen.apply(
        _calc_leverdatum_basis, axis=1
    )

    # Maak kolommen aan die aansluiten op de merged-structuur.
    reserveringen["Klantnaam"] = reserveringen["Debiteurnaam"] if "Debiteurnaam" in reserveringen.columns else ""
    debtor_number_col = _find_first_existing_column(
        reserveringen,
        [
            "Debiteurnummer", "Debiteur nummer", "Debiteur_nummer", "DebiteurNr",
            "Debiteur nr", "Klantnummer", "Klant nummer", "customer_number",
            "debtor_number",
        ],
    )
    if debtor_number_col and debtor_number_col != "Debiteurnummer":
        reserveringen["Debiteurnummer"] = reserveringen[debtor_number_col]
    reserveringen["Nummer"] = reserveringen["Ordernummer"].astype(str)
    reserveringen["Leverdatum"] = pd.to_datetime(reserveringen["Leverdatum_reservering_basis"], errors="coerce")
    reserveringen["Datum"] = reserveringen["Datum_verzending"]
    reserveringen["Status"] = "Reservering"
    reserveringen["Verzinkstatus"] = "Niet verzinkt"
    reserveringen["Gewicht"] = reserveringen["Gewicht_order_kg"]
    reserveringen["Gewicht_export_kg"] = np.nan
    reserveringen["Gewicht_2g_verdeeld_kg"] = reserveringen["Gewicht_order_kg"]
    reserveringen["Gewicht_effectief_kg"] = reserveringen["Gewicht_order_kg"]
    reserveringen["Gewicht_bron"] = "Reservering"
    reserveringen["Ordernummer_base"] = reserveringen["Ordernummer"]
    reserveringen["Regels_per_order"] = 1
    if "Debiteurnaam" in order.columns:
        reserveringen["Debiteurnaam"] = reserveringen["Debiteurnaam"] if "Debiteurnaam" in reserveringen.columns else ""
    reserveringen["Bronbestand"] = "OrderExport2G.xlsx"
    reserveringen["Bron_week"] = "reservering"

    return reserveringen


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

    # Download required files
    for fname in REQUIRED_FILES:
        raw = download_file(fname)
        (tmp / fname).write_bytes(raw)

    # Download optional files if available
    for fname in OPTIONAL_FILES:
        if fname in available:
            raw = download_file(fname)
            (tmp / fname).write_bytes(raw)

    validate_required_files_in_folder(tmp)

    # ── 1. CGS-weekexports inladen en samenvoegen ─────────────────────────
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
        if not fp.exists():
            continue   # per vestiging optioneel; overslaan indien niet gepubliceerd
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

    if not export_frames:
        raise FileNotFoundError(
            "Geen enkel Export-weekbestand gevonden. Minimaal één (bijv. "
            "Export-1.xlsx of Export.xlsx) is nodig."
        )

    export = pd.concat(export_frames, ignore_index=True)

    # Verzamel alle CgsNummers die al in de weekexports zitten
    existing_cgs_ids: set = set()
    if "CgsNummer" in export.columns:
        existing_cgs_ids = set(export["CgsNummer"].dropna().astype(int))

    # ── 2. OrderExport2G en feestdagen inladen ────────────────────────────
    order = pd.read_excel(tmp / "OrderExport2G.xlsx")
    feestdagen_fp = tmp / "feestdagen.xlsx"
    if feestdagen_fp.exists():
        holiday_df = pd.read_excel(feestdagen_fp)
    else:
        # feestdagen kan per vestiging optioneel zijn → lege kalender.
        holiday_df = pd.DataFrame(columns=["Datum", "Omschrijving", "Type"])
    debtor_lookup = _load_debtor_segment_lookup(tmp)

    # Scheepsleidingen-klanten (24-uursservice) altijd als depot behandelen,
    # gelijk aan 'Aanleveren depot = 1'. We forceren dit op de order-data zodat
    # het meeloopt in zowel de OTIF-depotshift/verzinkdatum (via de merge) als
    # de reserveringen. Optioneel bestand: ontbreekt het, dan geen effect.
    order = _force_scheepsleidingen_depot(order, _load_scheepsleidingen_ids(tmp))

    # ── 3. Basiskolommen aanmaken ─────────────────────────────────────────
    export["Gewicht_export_kg"] = coerce_numeric(export["Gewicht"])
    export["Leverdatum"] = pd.to_datetime(export["Datum"], dayfirst=True, errors="coerce")

    # Originele leverdatum (optioneel; wordt in de testomgeving als alternatief
    # peildatum-referentie voor OTIF gebruikt). Kolomnaam wisselt per bron; we
    # zoeken hem case-insensitief op basis van bekende kandidaten.
    originele_col = find_originele_leverdatum_column(export)
    if originele_col:
        export["Originele_leverdatum"] = pd.to_datetime(
            export[originele_col], dayfirst=True, errors="coerce"
        )
    else:
        export["Originele_leverdatum"] = pd.NaT

    export["Verzinkstatus"] = export["Status"].map(STATUS_MAP)
    export["Ordernummer_base"] = coerce_numeric(
        export["Nummer"].astype(str).str.extract(r"(\d+)")[0]
    )

    order["Ordernummer"]      = coerce_numeric(order["Ordernummer"])
    order["Gewicht_order_kg"] = coerce_numeric(order["Gewicht(ton)"]) * 1000

    # ── 4. Gewicht verdelen over deelorders (definitieve CGS-orders) ──────
    row_counts = (
        export.groupby("Ordernummer_base", dropna=False)
        .size()
        .rename("Regels_per_order")
        .reset_index()
    )
    merged = export.merge(row_counts, on="Ordernummer_base", how="left")
    # Debiteurnaam en debiteurnummer meenemen als kolommen beschikbaar zijn in OrderExport2G
    order_merge_cols = ["Ordernummer", "Gewicht_order_kg"]
    if "Debiteurnaam" in order.columns:
        order_merge_cols.append("Debiteurnaam")

    debtor_number_col_order = _find_first_existing_column(
        order,
        [
            "Debiteurnummer", "Debiteur nummer", "Debiteur_nummer", "DebiteurNr",
            "Debiteur nr", "Klantnummer", "Klant nummer", "customer_number",
            "debtor_number",
        ],
    )
    if debtor_number_col_order and debtor_number_col_order not in order_merge_cols:
        if debtor_number_col_order != "Debiteurnummer":
            order["Debiteurnummer"] = order[debtor_number_col_order]
        order_merge_cols.append("Debiteurnummer")

    if "Aanleveren depot" in order.columns:
        order_merge_cols.append("Aanleveren depot")
    # Leverdatum V meenemen zodat depot-orders de juiste leverdatum krijgen
    if "Leverdatum V" in order.columns:
        order["Leverdatum_V_2G"] = pd.to_datetime(order["Leverdatum V"], dayfirst=True, errors="coerce")
        order_merge_cols.append("Leverdatum_V_2G")
    merged = merged.merge(
        order[order_merge_cols],
        left_on="Ordernummer_base",
        right_on="Ordernummer",
        how="left",
    )
    merged["Gewicht_2g_verdeeld_kg"] = (
        merged["Gewicht_order_kg"] / merged["Regels_per_order"]
    )
    # Klantnaam: gebruik Debiteur uit CGS-export, vul aan met Debiteurnaam uit OrderExport2G
    if "Debiteur" in merged.columns and "Debiteurnaam" in merged.columns:
        merged["Klantnaam"] = merged["Debiteur"].where(
            merged["Debiteur"].notna() & (merged["Debiteur"].astype(str).str.strip() != ""),
            merged["Debiteurnaam"]
        )
    elif "Debiteur" in merged.columns:
        merged["Klantnaam"] = merged["Debiteur"]
    elif "Debiteurnaam" in merged.columns:
        merged["Klantnaam"] = merged["Debiteurnaam"]
    else:
        merged["Klantnaam"] = ""

    # ── Coatinc 24 Amsterdam altijd als depot behandelen ─────────────────
    if "Klantnaam" in merged.columns and "Aanleveren depot" in merged.columns:
        is_c24 = merged["Klantnaam"].astype(str).str.contains(
            "Coatinc 24 Amsterdam", case=False, na=False
        )
        merged.loc[is_c24, "Aanleveren depot"] = 1

    # Leverdatum voor export-orders = Datum uit Export-bestanden (ongewijzigd).
    # Leverdatum voor reserveringen = Leverdatum V uit OrderExport2G
    # (wordt afgehandeld in _build_reserveringen).

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

    # ── 5. Optionele CGS-planningsexport toevoegen ────────────────────────
    cgs_df, cgs_count = _load_cgs_as_export_rows(tmp, existing_cgs_ids)
    if cgs_count > 0:
        cgs_df["Gewicht_export_kg"]      = coerce_numeric(cgs_df["Gewicht"])
        cgs_df["Leverdatum"]             = pd.to_datetime(cgs_df["Datum"], errors="coerce")
        cgs_df["Verzinkstatus"]          = cgs_df["Status"].map(STATUS_MAP)
        cgs_df["Ordernummer_base"]       = np.nan
        cgs_df["Gewicht_order_kg"]       = np.nan
        cgs_df["Regels_per_order"]       = 1
        cgs_df["Gewicht_2g_verdeeld_kg"] = np.nan
        cgs_df["Gewicht_bron"]           = "CGS"
        cgs_df["Gewicht_effectief_kg"]   = cgs_df["Gewicht_export_kg"]
        merged = pd.concat([merged, cgs_df], ignore_index=True)
        export_file_summary.append({
            "Bronbestand": "Export_CGS.xlsx",
            "Bron_week": "CGS",
            "Aantal_regels_ingelezen": cgs_count,
        })

    # ── 6. Reserveringen toevoegen (Power BI Merge1-logica) ───────────────
    # Orders in OrderExport2G die NIET in de weekexports of CGS-export zitten
    # worden als reservering meegenomen in de capaciteitsplanning.
    cgs_ordernummers = set(
        merged["Ordernummer_base"].dropna().astype(float).astype(int).unique()
    )
    reserveringen = _build_reserveringen(order, cgs_ordernummers)

    if not reserveringen.empty:
        merged = pd.concat([merged, reserveringen], ignore_index=True)
        export_file_summary.append({
            "Bronbestand": "OrderExport2G.xlsx (reserveringen)",
            "Bron_week": "reservering",
            "Aantal_regels_ingelezen": len(reserveringen),
        })

    # ── 7. Segment / materiaaltype koppelen vanuit debtor-export ─────────────
    merged = _attach_debtor_segments(merged, debtor_lookup)
    if not debtor_lookup.empty:
        export_file_summary.append({
            "Bronbestand": DEBTOR_EXPORT_FILE,
            "Bron_week": "segmenten",
            "Aantal_regels_ingelezen": len(debtor_lookup),
        })

    # ── 8. Feestdagen opschonen ───────────────────────────────────────────
    holiday_df["Datum"] = pd.to_datetime(holiday_df["Datum"], errors="coerce").dt.date
    holiday_df = (
        holiday_df.dropna(subset=["Datum"])
        .drop_duplicates(subset=["Datum", "Omschrijving", "Type"])
        .sort_values("Datum")
        .reset_index(drop=True)
    )

    return merged, pd.DataFrame(export_file_summary), "OrderExport2G.xlsx", holiday_df


# ── Date / calendar helpers ────────────────────────────────────────────────────

def previous_workday(d: date, holiday_dates: set | None = None) -> date:
    """
    Retourneer de eerstvolgende werkdag vóór `d`. Slaat weekenden altijd over.
    Als `holiday_dates` wordt meegegeven, worden feestdagen ook overgeslagen.
    Backwards compatible: zonder feestdagen-set werkt de functie precies zoals
    voorheen.
    """
    d = d - timedelta(days=1)
    while d.weekday() >= 5 or (holiday_dates is not None and d in holiday_dates):
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



def add_workdays_existing_orders(d: pd.Timestamp, days: int) -> pd.Timestamp:
    """
    Tel werkdagen vooruit voor bestaande/reserveringsorders.
    Alleen weekenden worden uitgesloten; feestdagen worden in de capaciteitstabel
    zichtbaar gemaakt met capaciteit 0.
    """
    out = pd.Timestamp(d).normalize()
    remaining = int(days)
    while remaining > 0:
        out = out + timedelta(days=1)
        if out.weekday() < 5:
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

    if "Materiaaltype" not in df.columns:
        df["Materiaaltype"] = "Overig / onbekend"
    df["Materiaaltype"] = df["Materiaaltype"].fillna("Overig / onbekend").apply(normalize_materiaaltype)

    # Zwarte voorraad: subcategorie o.b.v. Status, ongeacht planningshorizon.
    df["Zwarte_voorraad_categorie"] = df["Status"].map(ZWARTE_VOORRAAD_MAP)

    # Verzinkdatum: depot Amsterdam (Aanleveren depot=1) → -1 werkdag, overig → sidebar offset
    def _row_offset(row):
        if pd.notna(row.get("Aanleveren depot")) and int(row["Aanleveren depot"]) == 1:
            return 1
        return offset

    df["Verzinkdatum"] = df.apply(
        lambda row: subtract_workdays_existing_orders(row["Leverdatum"], _row_offset(row))
        if pd.notna(row["Leverdatum"]) else pd.NaT,
        axis=1,
    )

    # Reserveringen krijgen een expliciete verzinkdatum volgens de businessregel:
    # Leverdatum V - 2 werkdagen; als Leverdatum V leeg is: Datum verzending + 2 werkdagen.
    # Die override voorkomt dat reserveringen per ongeluk via Leverdatum/Dag-offset opnieuw
    # worden berekend.
    if "Verzinkdatum_reservering" in df.columns:
        override = pd.to_datetime(df["Verzinkdatum_reservering"], errors="coerce")
        df.loc[override.notna(), "Verzinkdatum"] = override[override.notna()]

    start_ts = pd.Timestamp(startdatum).normalize()
    today_ts = pd.Timestamp(date.today()).normalize()

    df["Meegeteld_in_planning"] = "Nee"
    df["Reden_uitsluiting"]     = ""

    mask_niet_verzinkt   = df["Verzinkstatus"] == "Niet verzinkt"
    mask_ub              = df["Verzinkstatus"] == "UB"
    mask_verzinkt        = df["Verzinkstatus"] == "Verzinkt"
    mask_binnen_horizon  = df["Verzinkdatum"] >= start_ts

    df.loc[mask_verzinkt,                             "Reden_uitsluiting"] = "Status = Verzinkt"
    df.loc[mask_ub,                                   "Reden_uitsluiting"] = "Status = UB"
    df.loc[mask_niet_verzinkt & ~mask_binnen_horizon, "Reden_uitsluiting"] = "Voor startdatum rapport"
    df.loc[mask_niet_verzinkt &  mask_binnen_horizon, "Meegeteld_in_planning"] = "Ja"
    df.loc[mask_niet_verzinkt &  mask_binnen_horizon, "Reden_uitsluiting"]     = ""

    # Coat-orders horen niet in de verzinkstraat-capaciteitsplanning en worden
    # daarom altijd uitgesloten, ongeacht status/horizon. Twee patronen
    # (case-insensitief), zelfde classificatie als bij de OTIF-KPI:
    #   1) Coat-alleen: '<cijfers>C<cijfers>' zonder V ervoor (202616873C1)
    #   2) Coat-deel van combi-order: eindigt op '-C' (202615352VC1-C)
    # Het verzink-deel van combi-orders (202615352VC1-V) blijft wél meetellen.
    # Geldt voor alle vestigingen (CGR en CAL); geen locatie-specifieke flag.
    if "Nummer" in df.columns:
        mask_coat_order = df["Nummer"].apply(is_otif_excluded_coat_order)
        df.loc[mask_coat_order, "Meegeteld_in_planning"] = "Nee"
        df.loc[mask_coat_order, "Reden_uitsluiting"]     = "Coat-order"

    df_plan = df[df["Meegeteld_in_planning"] == "Ja"].copy()

    # Uitsplitsing definitief vs. reservering voor transparantie in dashboard
    df_plan["Gewicht_definitief_kg"]  = np.where(
        df_plan["Gewicht_bron"] != "Reservering",
        df_plan["Gewicht_effectief_kg"],
        0.0,
    )
    df_plan["Gewicht_reservering_kg"] = np.where(
        df_plan["Gewicht_bron"] == "Reservering",
        df_plan["Gewicht_effectief_kg"],
        0.0,
    )

    horizon = build_horizon_and_include_holidays(start_ts, 10, holiday_dates)

    dag_orders = df_plan.groupby("Verzinkdatum", as_index=False).agg(
        Gewicht_kg=("Gewicht_effectief_kg", "sum"),
        Gewicht_definitief_kg=("Gewicht_definitief_kg", "sum"),
        Gewicht_reservering_kg=("Gewicht_reservering_kg", "sum"),
        Aantal_orders_te_verzinken=("Nummer", "count"),
    )

    # Uitsplitsing tonnage per materiaaltype/segment. Dit is bewust los van
    # definitief/reservering, zodat de bestaande grafiek intact blijft en de
    # materiaal-mix als tweede beeld kan worden getoond.
    if not df_plan.empty:
        dag_materiaal = (
            df_plan.pivot_table(
                index="Verzinkdatum",
                columns="Materiaaltype",
                values="Gewicht_effectief_kg",
                aggfunc="sum",
                fill_value=0.0,
            )
            .reset_index()
        )
    else:
        dag_materiaal = pd.DataFrame(columns=["Verzinkdatum"])

    for materiaaltype in MATERIAALTYPE_ORDER:
        if materiaaltype not in dag_materiaal.columns:
            dag_materiaal[materiaaltype] = 0.0

    dag_materiaal = dag_materiaal[["Verzinkdatum"] + MATERIAALTYPE_ORDER].rename(
        columns=MATERIAALTYPE_DAG_COLS
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
    dag = dag.merge(dag_materiaal, on="Verzinkdatum", how="left")
    dag["Gewicht_kg"]               = dag["Gewicht_kg"].fillna(0.0)
    dag["Gewicht_definitief_kg"]    = dag["Gewicht_definitief_kg"].fillna(0.0)
    dag["Gewicht_reservering_kg"]   = dag["Gewicht_reservering_kg"].fillna(0.0)
    for col in MATERIAALTYPE_DAG_COLS.values():
        dag[col] = dag[col].fillna(0.0)
    dag["Aantal_orders_te_verzinken"] = dag["Aantal_orders_te_verzinken"].fillna(0).astype(int)
    dag["Capaciteit_kg"] = np.where(dag["Is_feestdag_of_sluiting"], 0, capaciteit_kg)
    dag["Benutting_pct"] = np.where(
        dag["Capaciteit_kg"] > 0,
        (dag["Gewicht_kg"] / dag["Capaciteit_kg"]) * 100,
        np.nan,
    )
    dag["Traverses_berekend"] = np.ceil(dag["Gewicht_kg"] / kg_per_traverse)
    dag["Status"]   = dag.apply(
        lambda r: stoplight(r["Benutting_pct"], r["Is_feestdag_of_sluiting"]), axis=1
    )
    dag["Jaar"]     = dag["Verzinkdatum"].dt.isocalendar().year.astype(int)
    dag["Week"]     = dag["Verzinkdatum"].dt.isocalendar().week.astype(int)
    dag["Label_nl"] = dag["Verzinkdatum"].apply(format_nl_axis_label)
    dag["Dagtype"]  = np.where(
        dag["Is_feestdag_of_sluiting"], "Feestdag / sluiting", "Werkdag"
    )

    week_agg = {
        "Gewicht_kg": ("Gewicht_kg", "sum"),
        "Gewicht_definitief_kg": ("Gewicht_definitief_kg", "sum"),
        "Gewicht_reservering_kg": ("Gewicht_reservering_kg", "sum"),
        "Aantal_orders_te_verzinken": ("Aantal_orders_te_verzinken", "sum"),
        "Traverses_berekend": ("Traverses_berekend", "sum"),
        "Capaciteit_kg": ("Capaciteit_kg", "sum"),
    }
    for col in MATERIAALTYPE_DAG_COLS.values():
        week_agg[col] = (col, "sum")

    week = dag.groupby(["Jaar", "Week"], as_index=False).agg(**week_agg)
    week["Benutting_pct"] = np.where(
        week["Capaciteit_kg"] > 0,
        (week["Gewicht_kg"] / week["Capaciteit_kg"]) * 100,
        np.nan,
    )
    week["Status"] = week["Benutting_pct"].apply(lambda x: stoplight(x, False))

    def calculate_advice_date(day_df, today_date, holiday_dates):
        base_date  = add_workdays(today_date, 5, holiday_dates)
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


# ── OTIF helpers (KPI: percentage orders op tijd gereed) ──────────────────────

def map_otif_status(status) -> str:
    """
    Vertaal een ruwe Status-waarde naar de OTIF-uitkomst 'Ja', 'Nee', 'Nvt'
    of 'Onbekend'. Matching is case-insensitief.
    """
    if pd.isna(status):
        return "Onbekend"
    key = str(status).strip().casefold()
    return OTIF_STATUS_MAP.get(key, "Onbekend")


# ── OTIF: uitsluiten van (poeder)coat-orders ──────────────────────────────────
# Poeder-coat orders horen niet in de OTIF van de verzinkstraat.
# Twee patronen om uit te sluiten:
#   1) Coat-alleen orders — typecode is enkel 'C' (geen V ervoor).
#      Voorbeeld: 202616661C2
#   2) Coat-deel van gecombineerde verzink+coat-orders — eindigt op '-C'.
#      Voorbeeld: 202614382VC1-C
# Wél meegenomen (verzink):
#   - Verzink-alleen orders (bevatten helemaal geen C): 202614382V1
#   - Verzink-deel van gecombineerde orders (VC zonder '-C' suffix):
#     202614382VC1
_OTIF_COAT_ONLY_PATTERN = re.compile(r"^\d+C\d*$", re.IGNORECASE)


def is_otif_excluded_coat_order(ordernummer) -> bool:
    """
    Retourneer True voor coat-orders die NIET in de OTIF-berekening horen.

    Uitgesloten:
    - Coat-alleen: ordernummer matcht `<cijfers>C<optionele cijfers>` zonder
      V ervoor (bijv. 202616661C2).
    - Coat-deel van een gecombineerde verzink+coat-order: ordernummer eindigt
      op '-C' (bijv. 202614382VC1-C).

    Niet uitgesloten (blijft dus meetellen in OTIF):
    - Verzink-alleen orders zoals 202614382V1
    - Verzink-deel van gecombineerde orders zoals 202614382VC1
    - Ordernummers die niet aan bovenstaande patronen voldoen

    Case-insensitief. NaN en lege waarden worden nooit uitgesloten.
    """
    if pd.isna(ordernummer):
        return False
    s = str(ordernummer).strip()
    if not s:
        return False
    if s.upper().endswith("-C"):
        return True
    if _OTIF_COAT_ONLY_PATTERN.match(s):
        return True
    return False


def find_originele_leverdatum_column(df: pd.DataFrame) -> str | None:
    """
    Zoek in `df` een kolom die de originele leverdatum bevat. Retourneert de
    (originele) kolomnaam of None als geen kandidaat gevonden is.
    """
    if df is None or df.empty:
        return None
    return _find_first_existing_column(df, ORIGINELE_LEVERDATUM_CANDIDATES)


def _empty_otif_result(peildatum, date_column: str, date_column_label: str) -> dict:
    return {
        "totaal": 0,
        "gereed": 0,
        "niet_gereed": 0,
        "nvt": 0,
        "onbekend": 0,
        "otif_pct": None,
        "peildatum": peildatum,
        "peildatum_vorige_werkdag": None,
        "date_column": date_column,
        "date_column_label": date_column_label,
        "depot_shift_actief": False,
        "aantal_depot": 0,
        "aantal_coat_uitgesloten": 0,
        "orders": pd.DataFrame(),
    }


def compute_otif(
    merged: pd.DataFrame,
    peildatum,
    date_column: str = "Leverdatum",
    date_column_label: str | None = None,
    holiday_dates: set | None = None,
    depot_column: str = "Aanleveren depot",
    apply_depot_shift: bool = True,
    poetsen_afgehaald_niet_ok: bool = False,
    prijscategorie_column: str = "PrijsCategorie",
    uitsluiten_statussen: set | None = None,
) -> dict:
    """
    Bereken de OTIF-KPI over orders waarvan de te toetsen datum gelijk is aan
    de peildatum.

    Uitgangspunten:
    - Reserveringen (Bron_week == 'reservering') worden niet meegeteld; dit
      zijn planningsobjecten, geen orders in productie.
    - De 'peildatum' is standaard de dag van de laatste publicatie (de dag
      waarop de data is geëxporteerd/geüpload).
    - `date_column` bepaalt welk datumveld tegen de peildatum wordt gehouden:
        * 'Leverdatum'           → conform de 'Datum' uit de weekexports
        * 'Originele_leverdatum' → variant t.o.v. de originele afspraakdatum

    Depot-shift (op verzoek):
    - Orders met `Aanleveren depot = 1` hebben een deadline van 17:00 op hun
      leverdatum. Omdat de brondata rond 08:30 wordt geëxporteerd, worden
      deze orders pas de eerstvolgende werkdag zinvol op OTIF getoetst.
    - Op peildatum P worden dus meegenomen:
        (a) orders met `date_column == P` en `Aanleveren depot != 1`, én
        (b) orders met `date_column == vorige_werkdag(P)` en
            `Aanleveren depot == 1`.
    - `holiday_dates` zorgt ervoor dat 'vorige werkdag' feestdagen overslaat.
    - Zonder depot-kolom of met `apply_depot_shift=False` valt de logica
      terug op de oorspronkelijke datum-vergelijking.

    Poetsen-uitzondering (per vestiging, standaard uit):
    - Met `poetsen_afgehaald_niet_ok=True` telt voor orders waarvan de kolom
      `prijscategorie_column` (default 'PrijsCategorie') de tekst 'poetsen'
      bevat, de status 'Afgehaald' als niet OK ('Nee') i.p.v. OK. Alle andere
      statussen volgen de normale OTIF-mapping.

    Coat-uitsluiting (altijd actief):
    - Poeder-coat orders horen niet in de OTIF van de verzinkstraat en worden
      weggefilterd. Twee patronen (case-insensitief):
        1) Coat-alleen — ordernummer past bij `<cijfers>C<optionele cijfers>`
           zonder V ervoor (bijv. 202616661C2).
        2) Coat-deel van gecombineerde verzink+coat-order — ordernummer
           eindigt op '-C' (bijv. 202614382VC1-C).
    - Het verzink-deel van gecombineerde orders (bijv. 202614382VC1 zonder
      '-C') blijft wél meetellen — dat is het gedeelte dat op de verzinkstraat
      geproduceerd wordt.

    Formule conform gebruikersspecificatie:
        OTIF = (1 - aantal te laat / totaal aantal orders) * 100 %

    Waar 'totaal' alle meetellende orders betreft (inclusief Nvt en
    Onbekend) en 'te laat' het aantal orders met OTIF-status = 'Nee'.
    """
    label = date_column_label or date_column
    if merged is None or merged.empty or date_column not in merged.columns:
        return _empty_otif_result(peildatum, date_column, label)

    df = merged.copy()

    # Reserveringen zijn nog geen 'in productie' orders en tellen niet mee.
    if "Bron_week" in df.columns:
        df = df[df["Bron_week"].astype(str).str.strip().str.lower() != "reservering"].copy()

    peil_ts = pd.Timestamp(peildatum).normalize()
    date_series = pd.to_datetime(df[date_column], errors="coerce").dt.normalize()

    # Depot-vlag: NaN/leeg = geen depot. Coatinc 24 Amsterdam is elders in
    # load_published_data al op 1 gezet. Waarde-vergelijking op '== 1' zodat
    # tekstuele varianten ('1'/'0') na coercion netjes werken.
    depot_shift_actief = bool(apply_depot_shift) and (depot_column in df.columns)
    if depot_shift_actief:
        depot_num = pd.to_numeric(df[depot_column], errors="coerce").fillna(0)
        is_depot = depot_num == 1
        vorige_werkdag = previous_workday(peildatum, holiday_dates)
        vorige_ts = pd.Timestamp(vorige_werkdag).normalize()
        # Niet-depot orders op peildatum + depot-orders op vorige werkdag.
        mask = (
            ((date_series == peil_ts)  & (~is_depot))
            |
            ((date_series == vorige_ts) & ( is_depot))
        )
    else:
        is_depot = pd.Series(False, index=df.index)
        vorige_werkdag = None
        mask = date_series == peil_ts

    subset = df.loc[mask].copy()

    # Statussen die volledig buiten de OTIF-telling blijven (per vestiging),
    # bijv. 'UB'. Exacte match op status (getrimd, hoofdletterongevoelig), zodat
    # bijv. 'UB V Gereed' NIET wordt uitgesloten.
    if uitsluiten_statussen and "Status" in subset.columns:
        excl = {str(s).strip().casefold() for s in uitsluiten_statussen}
        subset = subset[
            ~subset["Status"].astype(str).str.strip().str.casefold().isin(excl)
        ].copy()

    if subset.empty or "Status" not in subset.columns:
        result = _empty_otif_result(peildatum, date_column, label)
        result["peildatum_vorige_werkdag"] = vorige_werkdag
        result["depot_shift_actief"] = depot_shift_actief
        return result

    subset["OTIF_status"] = subset["Status"].apply(map_otif_status)

    # Poetsen-uitzondering (per vestiging): voor orders waarvan PrijsCategorie
    # 'poetsen' bevat, telt de status 'Afgehaald' op de peildatum als niet OK
    # ('Nee') in plaats van OK. Overige statussen volgen de normale regels.
    if poetsen_afgehaald_niet_ok and prijscategorie_column in subset.columns:
        is_poetsen = (
            subset[prijscategorie_column].astype(str).str.contains("poetsen", case=False, na=False)
        )
        is_afgehaald = subset["Status"].astype(str).str.strip().str.casefold() == "afgehaald"
        subset.loc[is_poetsen & is_afgehaald, "OTIF_status"] = "Nee"

    # Sluit poeder-coat orders uit; deze horen niet in de OTIF van de
    # verzinkstraat. Filteren gebeurt NA de datum/depot-mask en NA de
    # status-mapping (inclusief poetsen-uitzondering) zodat de telling
    # meteen zinvol is voor de gekozen peildatum.
    aantal_coat_uitgesloten = 0
    if "Nummer" in subset.columns:
        is_coat_excl = subset["Nummer"].apply(is_otif_excluded_coat_order)
        aantal_coat_uitgesloten = int(is_coat_excl.sum())
        if aantal_coat_uitgesloten:
            subset = subset.loc[~is_coat_excl].copy()

    # Kan leeg worden als álle relevante orders coat waren.
    if subset.empty:
        result = _empty_otif_result(peildatum, date_column, label)
        result["peildatum_vorige_werkdag"] = vorige_werkdag
        result["depot_shift_actief"] = depot_shift_actief
        result["aantal_coat_uitgesloten"] = aantal_coat_uitgesloten
        return result

    # Boolean voor duidelijkheid in de detailtabel én in de samenvatting.
    if depot_shift_actief:
        subset["Is_depot"] = is_depot.loc[subset.index]
    else:
        subset["Is_depot"] = False

    gereed      = int((subset["OTIF_status"] == "Ja").sum())
    niet_gereed = int((subset["OTIF_status"] == "Nee").sum())
    nvt         = int((subset["OTIF_status"] == "Nvt").sum())
    onbekend    = int((subset["OTIF_status"] == "Onbekend").sum())
    totaal      = len(subset)
    aantal_depot = int(subset["Is_depot"].sum())

    # OTIF-percentage conform de gespecificeerde formule.
    otif_pct = (1 - niet_gereed / totaal) * 100 if totaal > 0 else None

    return {
        "totaal": totaal,
        "gereed": gereed,
        "niet_gereed": niet_gereed,
        "nvt": nvt,
        "onbekend": onbekend,
        "otif_pct": otif_pct,
        "peildatum": peildatum,
        "peildatum_vorige_werkdag": vorige_werkdag,
        "date_column": date_column,
        "date_column_label": label,
        "depot_shift_actief": depot_shift_actief,
        "aantal_depot": aantal_depot,
        "aantal_coat_uitgesloten": aantal_coat_uitgesloten,
        "orders": subset,
    }


def get_peildatum_from_metadata(meta: dict | None) -> date:
    """
    Bepaal de OTIF-peildatum uit de publicatiedatum in metadata. Valt terug
    op vandaag als de metadata ontbreekt of niet parseerbaar is.
    """
    if not meta:
        return date.today()
    raw = str(meta.get("published_at", "")).strip()
    if not raw:
        return date.today()

    for fmt in (
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ):
        try:
            return pd.to_datetime(raw, format=fmt).date()
        except Exception:
            continue

    # Laatste redmiddel: laat pandas de datum inschatten.
    try:
        return pd.to_datetime(raw, dayfirst=True, errors="raise").date()
    except Exception:
        return date.today()
