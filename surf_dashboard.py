import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import requests
import plotly.graph_objects as go
from datetime import datetime, timezone
import math
import os

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="East Coast Surf Dashboard",
    page_icon="🏄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f0f4f8; }
    .stApp { background-color: #f0f4f8; }
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        text-align: center;
        border-left: 4px solid #1B3A5C;
    }
    .metric-label { font-size: 12px; color: #888; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-value { font-size: 28px; font-weight: 700; color: #1B3A5C; margin: 4px 0; }
    .metric-unit  { font-size: 12px; color: #aaa; }
    .swell-badge {
        display: inline-block;
        background: #1F7A8C;
        color: white;
        font-size: 32px;
        font-weight: 800;
        padding: 10px 28px;
        border-radius: 10px;
        letter-spacing: 2px;
    }
    .spot-match   { background: #C6EFCE; border-radius: 8px; }
    .spot-nomatch { background: #f5f5f5; color: #aaa; border-radius: 8px; }
    div[data-testid="stMetricValue"] { font-size: 22px !important; }
    .stSelectbox label { font-weight: 600; }
    .section-title {
        font-size: 18px;
        font-weight: 700;
        color: #1B3A5C;
        margin: 18px 0 8px 0;
        border-bottom: 2px solid #BDD7EE;
        padding-bottom: 6px;
    }
    .live-badge {
        background: #27ae60;
        color: white;
        font-size: 11px;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 20px;
        margin-left: 8px;
    }
    .stale-badge {
        background: #e67e22;
        color: white;
        font-size: 11px;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 20px;
        margin-left: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "East_Coast_Surf_Spots.xlsx")

BUOYS = {
    "44007 – Portland, ME":   "44007",
    "44013 – Boston, MA":     "44013",
    "44017 – Montauk, NY":    "44017",
    "44025 – New York Bight": "44025",
    "44027 – Jonesport, ME":  "44027",
}

# Approximate lat/lon for each spot (for map)
SPOT_COORDS = {
    "Higgins Beach":       (43.5589, -70.3294),
    "Old Orchard Beach":   (43.5168, -70.3740),
    "Scarborough Beach":   (43.5467, -70.3629),
    "Fortunes Rocks Beach":(43.4598, -70.4097),
    "Kennebunk Beach":     (43.3490, -70.4768),
    "Kennebunkport":       (43.3620, -70.4780),
    "Lox Point":           (43.3551, -70.4815),
    "Wells Jetties":       (43.3049, -70.5618),
    "Wells Beach":         (43.3180, -70.5622),
    "Moody Beach":         (43.2958, -70.5665),
    "Ogunquit Beach":      (43.2468, -70.5847),
    "Ogunquit Rivermouth": (43.2401, -70.5967),
    "Long Sands Beach":    (43.1704, -70.6427),
    "Short Sands Beach":   (43.1738, -70.6483),
}

SWELL_DIRECTIONS = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"]
DB_SWELL_COLS    = ["N", "NE", "E", "SE", "S", "SW"]  # columns in Excel

QUALITY_MAP = {1: "⭐", 2: "⭐⭐", 3: "⭐⭐⭐", 4: "⭐⭐⭐⭐", 5: "⭐⭐⭐⭐⭐"}

# NOAA tide stations near Southern Maine surf spots
TIDE_STATIONS = {
    "Wells, ME (8419317)":          "8419317",
    "Portland, ME (8418150)":       "8418150",
    "Portsmouth, NH (8423898)":     "8423898",
}
TIDE_DEFAULT = "Wells, ME (8419317)"

# Surfline spot IDs — add more as you find them from your Surfline URLs
SURFLINE_SPOTS = {
    "Long Sands Beach": "5842041f4e65fad6a77089e3",
    # e.g. "Ogunquit Beach": "<spot_id_from_surfline_url>",
}
SURFLINE_DEFAULT_SPOT = "Long Sands Beach"

# ── Helpers ───────────────────────────────────────────────────────────────────
def degrees_to_cardinal(deg):
    """Convert wave direction degrees to 8-point cardinal."""
    try:
        deg = float(deg)
        if deg >= 999:
            return None
        dirs = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"]
        idx = round(deg / 45) % 8
        return dirs[idx]
    except (TypeError, ValueError):
        return None

def cardinal_to_degrees(card):
    mapping = {"N": 0, "NE": 45, "E": 90, "SE": 135,
               "S": 180, "SW": 225, "W": 270, "NW": 315}
    return mapping.get(card, 0)

def meters_to_feet(m):
    try:
        v = float(m)
        return round(v * 3.281, 1) if v < 999 else None
    except (TypeError, ValueError):
        return None

def ms_to_mph(ms):
    try:
        v = float(ms)
        return round(v * 2.237, 1) if v < 999 else None
    except (TypeError, ValueError):
        return None

def star_rating(stars_str):
    """Convert '★★★★☆' style string to int count of ★."""
    if not stars_str:
        return 0
    return str(stars_str).count("★")

def degrees_to_arrow(deg):
    """Map swell degrees to a directional arrow (points toward swell source)."""
    try:
        deg = float(deg) % 360
    except (TypeError, ValueError):
        return "·"
    arrows = ["↑", "↗", "→", "↘", "↓", "↙", "←", "↖"]
    idx = round(deg / 45) % 8
    return arrows[idx]

def degrees_to_cardinal_16(deg):
    """16-point cardinal for finer LOTUS direction labels (e.g. ESE, SSE)."""
    try:
        deg = float(deg) % 360
    except (TypeError, ValueError):
        return "—"
    dirs = ["N","NNE","NE","ENE","E","ESE","SE","SSE",
            "S","SSW","SW","WSW","W","WNW","NW","NNW"]
    idx = round(deg / 22.5) % 16
    return dirs[idx]

# ── Data loaders ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def load_spots():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Spots Database", header=2)
    df.columns = df.columns.str.strip()
    # Rename unnamed cols if needed
    df = df.dropna(subset=["Spot Name"])
    df["Spot Name"] = df["Spot Name"].astype(str).str.strip()
    # Quality: count ★ characters
    df["Quality_Int"] = df["Quality ★"].apply(star_rating)
    return df

@st.cache_data(ttl=1800)  # refresh every 30 min
def fetch_buoy(buoy_id):
    """Fetch the latest observation from NDBC realtime2."""
    url = f"https://www.ndbc.noaa.gov/data/realtime2/{buoy_id}.txt"
    try:
        resp = requests.get(url, timeout=12)
        resp.raise_for_status()
        lines = [l for l in resp.text.strip().split("\n") if l]
        # Line 0 = col headers (starts with #YY)
        # Line 1 = units (starts with #yr)
        # Line 2+ = data rows
        col_line = lines[0].lstrip("#").split()
        data_rows = [l for l in lines[2:] if not l.startswith("#")]
        if not data_rows:
            return None, "No data rows found"
        latest = data_rows[0].split()
        if len(latest) < len(col_line):
            return None, "Row length mismatch"
        record = dict(zip(col_line, latest))
        return record, None
    except requests.exceptions.Timeout:
        return None, "Request timed out — NOAA may be slow"
    except requests.exceptions.ConnectionError:
        return None, "Could not reach NOAA NDBC. Check internet connection."
    except Exception as e:
        return None, str(e)

def parse_buoy(record):
    """Extract clean values from raw NDBC record dict."""
    if not record:
        return {}
    wvht_ft = meters_to_feet(record.get("WVHT"))
    wspd_mph = ms_to_mph(record.get("WSPD"))
    gst_mph  = ms_to_mph(record.get("GST"))
    dpd      = record.get("DPD", "MM")
    mwd_raw  = record.get("MWD", "MM")
    wdir_raw = record.get("WDIR", "MM")
    atmp     = record.get("ATMP", "MM")
    wtmp     = record.get("WTMP", "MM")

    mwd_cardinal  = degrees_to_cardinal(mwd_raw)  if mwd_raw  != "MM" else None
    wdir_cardinal = degrees_to_cardinal(wdir_raw) if wdir_raw != "MM" else None

    # Timestamp
    try:
        yy = record.get("YY", record.get("#YY", "00"))
        mm = record.get("MM", "01")
        dd = record.get("DD", "01")
        hh = record.get("hh", "00")
        mn = record.get("mm", "00")
        obs_dt = datetime(int(yy), int(mm), int(dd), int(hh), int(mn),
                          tzinfo=timezone.utc)
    except Exception:
        obs_dt = None

    def safe_float(v, sentinel=99):
        try:
            f = float(v)
            return None if f >= sentinel * 10 else f
        except (TypeError, ValueError):
            return None

    return {
        "wvht_ft":      wvht_ft,
        "dpd_sec":      safe_float(dpd),
        "mwd_deg":      safe_float(mwd_raw),
        "mwd_card":     mwd_cardinal,
        "wdir_card":    wdir_cardinal,
        "wspd_mph":     wspd_mph,
        "gst_mph":      gst_mph,
        "atmp_c":       safe_float(atmp),
        "wtmp_c":       safe_float(wtmp),
        "obs_dt":       obs_dt,
    }

@st.cache_data(ttl=1800)
def fetch_surfline_forecast(spot_id):
    """Fetch LOTUS wave forecast from Surfline's undocumented API."""
    url = "https://services.surfline.com/kbyg/spots/forecasts/wave"
    params = {"spotId": spot_id, "days": 1, "intervalHours": 1}
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=12)
        resp.raise_for_status()
        return resp.json(), None
    except requests.exceptions.Timeout:
        return None, "Surfline request timed out"
    except requests.exceptions.ConnectionError:
        return None, "Could not reach Surfline API"
    except Exception as e:
        return None, str(e)

def parse_lotus(data):
    """Extract the current forecast entry and swell components from LOTUS response."""
    if not data:
        return None
    try:
        wave_entries = data["data"]["wave"]
        if not wave_entries:
            return None
        # Find the entry closest to now
        now_ts = datetime.now(timezone.utc).timestamp()
        closest = min(wave_entries, key=lambda e: abs(e["timestamp"] - now_ts))
        swells = closest.get("swells", [])
        surf   = closest.get("surf", {})
        # Filter out flat/zero swells
        active_swells = [s for s in swells if s.get("height", 0) > 0]
        # Sort by optimalScore desc, then height desc
        active_swells.sort(key=lambda s: (s.get("optimalScore", 0),
                                           s.get("height", 0)), reverse=True)
        # Dominant = first after sort
        dominant = active_swells[0] if active_swells else None
        dominant_card = degrees_to_cardinal(dominant["direction"]) if dominant else None
        return {
            "swells":        active_swells,
            "dominant":      dominant,
            "dominant_card": dominant_card,
            "surf_min_ft":   round(surf.get("min", 0) * 3.281, 1),
            "surf_max_ft":   round(surf.get("max", 0) * 3.281, 1),
            "optimal_score": surf.get("optimalScore", 0),
            "timestamp":     closest["timestamp"],
        }
    except (KeyError, IndexError, TypeError):
        return None

@st.cache_data(ttl=3600)
def fetch_tide(station_id):
    """Fetch today's hourly tide predictions + hi/lo from NOAA Tides & Currents."""
    today = datetime.now().strftime("%Y%m%d")
    base  = "https://api.tidesandcurrents.noaa.gov/api/prod/datagetter"
    common = dict(station=station_id, datum="MLLW", time_zone="lst_ldt",
                  units="english", application="surf_dashboard", format="json")
    try:
        # Hourly curve
        r_hourly = requests.get(base, params={**common,
            "product": "predictions", "interval": "h",
            "begin_date": today, "end_date": today}, timeout=10)
        r_hourly.raise_for_status()
        hourly = r_hourly.json().get("predictions", [])

        # Hi/Lo markers
        r_hilo = requests.get(base, params={**common,
            "product": "predictions", "interval": "hilo",
            "begin_date": today, "end_date": today}, timeout=10)
        r_hilo.raise_for_status()
        hilo = r_hilo.json().get("predictions", [])

        return {"hourly": hourly, "hilo": hilo}, None
    except requests.exceptions.ConnectionError:
        return None, "Could not reach NOAA Tides API"
    except Exception as e:
        return None, str(e)

def parse_tide(data):
    """Return times, heights, hi/lo events, and current tide state."""
    if not data or not data.get("hourly"):
        return None
    try:
        times, heights = [], []
        for pt in data["hourly"]:
            times.append(datetime.strptime(pt["t"], "%Y-%m-%d %H:%M"))
            heights.append(float(pt["v"]))

        now_local = datetime.now()

        # Interpolate current height between two nearest hourly readings
        current_ht = None
        state       = "—"
        for i in range(len(times) - 1):
            if times[i] <= now_local <= times[i + 1]:
                frac = (now_local - times[i]).seconds / 3600
                current_ht = round(heights[i] + frac * (heights[i+1] - heights[i]), 2)
                state = "📈 Rising" if heights[i+1] > heights[i] else "📉 Falling"
                break

        # Hi/Lo events
        hilo_events = []
        for pt in data.get("hilo", []):
            t = datetime.strptime(pt["t"], "%Y-%m-%d %H:%M")
            hilo_events.append({
                "time":   t,
                "height": round(float(pt["v"]), 2),
                "type":   pt["type"],   # "H" or "L"
                "label":  "High" if pt["type"] == "H" else "Low",
                "past":   t < now_local,
            })

        # Next tide event
        upcoming = [e for e in hilo_events if not e["past"]]
        next_event = upcoming[0] if upcoming else None

        return {
            "times":      times,
            "heights":    heights,
            "current_ht": current_ht,
            "state":      state,
            "hilo":       hilo_events,
            "next_event": next_event,
            "now":        now_local,
        }
    except Exception:
        return None

def match_spots(df, swell_dir):
    """Return df with 'Match' bool column for the given swell direction."""
    df = df.copy()
    if swell_dir in DB_SWELL_COLS:
        df["Match"] = df[swell_dir].astype(str).str.strip().str.upper() == "Y"
    else:
        # W or NW not in DB — mark as offshore
        df["Match"] = False
    return df

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/4/47/PNG_transparency_demonstration_1.png/1px-PNG_transparency_demonstration_1.png",
             width=1)   # spacer trick
    st.markdown("## 🏄 Surf Dashboard")
    st.markdown("---")

    st.markdown("### 🌊 Tide Station")
    selected_tide_name = st.selectbox("Tide station", list(TIDE_STATIONS.keys()),
                                       index=0, label_visibility="collapsed")
    tide_station_id = TIDE_STATIONS[selected_tide_name]

    st.markdown("### 📡 NOAA Buoy")
    selected_buoy_name = st.selectbox("Select nearest buoy", list(BUOYS.keys()),
                                       index=0, label_visibility="collapsed")
    buoy_id = BUOYS[selected_buoy_name]

    if st.button("🔄 Refresh Live Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.markdown("### 🎯 Filter Spots")

    skill_options = ["All Levels", "Beginner", "Beginner–Int",
                     "Intermediate", "Int–Adv", "Advanced"]
    skill_filter = st.selectbox("Skill Level", skill_options)

    region_filter = st.multiselect(
        "Region",
        ["Maine – North", "Maine – South"],
        default=["Maine – North", "Maine – South"]
    )

    st.markdown("---")
    st.markdown("### ✏️ Manual Override")
    manual_on = st.toggle("Override swell direction")
    if manual_on:
        manual_dir = st.selectbox("Select swell direction",
                                   SWELL_DIRECTIONS, index=1)

    st.markdown("---")
    st.markdown(
        "<small style='color:#aaa'>"
        "🔮 Swell: Surfline LOTUS<br>"
        "📡 Observed: NOAA NDBC<br>"
        "🌊 Tides: NOAA Tides & Currents<br>"
        "📖 Spots: Stormrider Guide NA<br>"
        "Refreshes every 30–60 min</small>",
        unsafe_allow_html=True
    )

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(
    "<h1 style='color:#1B3A5C; margin-bottom:2px;'>🌊 East Coast Surf Spot Dashboard</h1>"
    "<p style='color:#888; margin-top:0;'>Maine Coast · "
    "LOTUS spot forecast + NOAA NDBC observed conditions</p>",
    unsafe_allow_html=True
)

# ── Fetch buoy + Surfline data in parallel ────────────────────────────────────
with st.spinner("Fetching live conditions…"):
    raw_record, fetch_error = fetch_buoy(buoy_id)
    buoy_data = parse_buoy(raw_record)
    spot_id   = SURFLINE_SPOTS[SURFLINE_DEFAULT_SPOT]
    lotus_raw, lotus_error = fetch_surfline_forecast(spot_id)
    lotus     = parse_lotus(lotus_raw)
    tide_raw, tide_error   = fetch_tide(tide_station_id)
    tide      = parse_tide(tide_raw)

# ── Conditions banner ─────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📊 Current Ocean Conditions</div>',
            unsafe_allow_html=True)

if fetch_error or not buoy_data:
    st.warning(f"⚠️ Could not load buoy data: {fetch_error or 'Unknown error'}. "
               f"Using manual swell direction.")
    live_available = False
else:
    live_available = True
    # Freshness badge
    if buoy_data.get("obs_dt"):
        now_utc = datetime.now(timezone.utc)
        age_min = (now_utc - buoy_data["obs_dt"]).total_seconds() / 60
        age_str = f"{int(age_min)}m ago" if age_min < 60 else f"{age_min/60:.1f}h ago"
        badge = "live-badge" if age_min < 90 else "stale-badge"
        st.markdown(
            f"<span style='color:#888; font-size:13px;'>Buoy <strong>{buoy_id}</strong> "
            f"— observation as of {buoy_data['obs_dt'].strftime('%H:%M UTC')}</span>"
            f"<span class='{badge}'>{age_str}</span>",
            unsafe_allow_html=True
        )

col1, col2, col3, col4, col5 = st.columns(5)

def metric_card(label, value, unit=""):
    disp = str(value) if value is not None else "—"
    return f"""<div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{disp}</div>
        <div class="metric-unit">{unit}</div>
    </div>"""

with col1:
    wvht = buoy_data.get("wvht_ft") if live_available else None
    st.markdown(metric_card("Wave Height", wvht, "ft"), unsafe_allow_html=True)

with col2:
    dpd = buoy_data.get("dpd_sec") if live_available else None
    st.markdown(metric_card("Wave Period", dpd, "sec"), unsafe_allow_html=True)

with col3:
    mwd_card = buoy_data.get("mwd_card") if live_available else None
    mwd_deg  = buoy_data.get("mwd_deg")  if live_available else None
    deg_str  = f"({int(mwd_deg)}°)" if mwd_deg is not None else ""
    st.markdown(metric_card("Swell Direction", mwd_card or "—", deg_str),
                unsafe_allow_html=True)

with col4:
    wspd = buoy_data.get("wspd_mph") if live_available else None
    wdir = buoy_data.get("wdir_card") if live_available else None
    st.markdown(metric_card("Wind Speed", wspd, f"mph from {wdir or '—'}"),
                unsafe_allow_html=True)

with col5:
    wtmp = buoy_data.get("wtmp_c") if live_available else None
    wtmp_f = round(wtmp * 9/5 + 32, 1) if wtmp is not None else None
    st.markdown(metric_card("Water Temp", wtmp_f, "°F"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── LOTUS Swell Components ────────────────────────────────────────────────────
st.markdown('<div class="section-title">🔮 LOTUS Forecast — Swell Components</div>',
            unsafe_allow_html=True)

if lotus_error or not lotus:
    st.warning(f"⚠️ Could not load Surfline LOTUS data: {lotus_error or 'No data'}. "
               "Spot matching will use NOAA buoy direction instead.")
else:
    forecast_dt = datetime.fromtimestamp(lotus["timestamp"], tz=timezone.utc)
    surf_range  = f"{lotus['surf_min_ft']}–{lotus['surf_max_ft']} ft"
    score_stars = ["", "●", "●●", "●●●"][min(lotus["optimal_score"], 3)]
    score_color = ["#aaa", "#e67e22", "#27ae60", "#1F7A8C"][min(lotus["optimal_score"], 3)]

    lotus_col, surf_col = st.columns([3, 1])

    with lotus_col:
        rows_html = ""
        for i, swell in enumerate(lotus["swells"][:3]):
            h    = round(swell.get("height", 0), 1)  # already in feet from Surfline API
            p    = int(swell.get("period", 0))
            deg  = swell.get("direction", 0)
            card = degrees_to_cardinal_16(deg)
            arr  = degrees_to_arrow(deg)
            opt  = swell.get("optimalScore", 0)
            # Dominant swell gets highlighted row
            row_bg    = "#e8f7f0" if i == 0 else ("white" if i == 1 else "#fafafa")
            row_weight= "700" if i == 0 else "400"
            opt_dot   = f"<span style='color:#1F7A8C; font-size:10px; margin-left:6px;'>{'●' * opt}</span>" if opt else ""
            rows_html += f"""
            <tr style="background:{row_bg}; border-bottom:1px solid #eee;">
                <td style="padding:10px 14px; font-size:22px; color:#1F7A8C;">{arr}</td>
                <td style="padding:10px 8px; font-size:15px; font-weight:{row_weight}; color:#1B3A5C;">{h} ft</td>
                <td style="padding:10px 8px; font-size:15px; font-weight:{row_weight}; color:#1B3A5C;">{p}s</td>
                <td style="padding:10px 8px; font-size:15px; font-weight:{row_weight}; color:#444;">
                    {card}&nbsp;<span style="color:#999; font-size:12px;">{int(deg)}°</span>{opt_dot}
                </td>
            </tr>"""

        st.markdown(f"""
        <div style="background:white; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.08);
                    overflow:hidden; border-left:4px solid #1F7A8C;">
            <div style="background:#1F7A8C; padding:8px 14px;">
                <span style="color:white; font-size:12px; font-weight:700;">
                    LOTUS &nbsp;·&nbsp; {SURFLINE_DEFAULT_SPOT} &nbsp;·&nbsp;
                    {forecast_dt.strftime('%H:%M UTC')}
                </span>
            </div>
            <table style="width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="background:#f0f4f8;">
                        <th style="padding:6px 14px; font-size:10px; color:#888; text-align:left;
                                   font-weight:600; text-transform:uppercase;">Dir</th>
                        <th style="padding:6px 8px; font-size:10px; color:#888; text-align:left;
                                   font-weight:600; text-transform:uppercase;">Height</th>
                        <th style="padding:6px 8px; font-size:10px; color:#888; text-align:left;
                                   font-weight:600; text-transform:uppercase;">Period</th>
                        <th style="padding:6px 8px; font-size:10px; color:#888; text-align:left;
                                   font-weight:600; text-transform:uppercase;">Swell</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>""", unsafe_allow_html=True)

    with surf_col:
        st.markdown(f"""
        <div style="background:white; border-radius:12px; padding:20px;
                    box-shadow:0 2px 8px rgba(0,0,0,0.08); text-align:center;
                    border-top:4px solid {score_color}; height:100%;">
            <div style="font-size:11px; color:#888; font-weight:600;
                        text-transform:uppercase; margin-bottom:8px;">Surf Height</div>
            <div style="font-size:26px; font-weight:800; color:#1B3A5C;">{surf_range}</div>
            <div style="font-size:20px; color:{score_color}; margin-top:8px;">{score_stars}</div>
            <div style="font-size:10px; color:#aaa; margin-top:6px;">Optimal Score</div>
            <hr style="margin:12px 0; border-color:#eee;">
            <div style="font-size:11px; color:#888; font-weight:600;
                        text-transform:uppercase; margin-bottom:4px;">Dominant Swell</div>
            <div style="font-size:20px; font-weight:800; color:#1F7A8C;">
                {lotus.get('dominant_card') or '—'}
            </div>
            <div style="font-size:11px; color:#aaa;">driving spot match</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Determine active swell direction ──────────────────────────────────────────
# Priority: Manual > LOTUS dominant > NOAA buoy > default
if manual_on:
    active_swell = manual_dir
    swell_source = "Manual override"
elif lotus and lotus.get("dominant_card"):
    active_swell = lotus["dominant_card"]
    swell_source = f"LOTUS forecast · {SURFLINE_DEFAULT_SPOT}"
elif live_available and buoy_data.get("mwd_card"):
    active_swell = buoy_data["mwd_card"]
    swell_source = f"NOAA buoy {buoy_id}"
else:
    active_swell = "NE"
    swell_source = "Default (no live data)"

# Swell direction display
sw_col, info_col = st.columns([1, 3])
with sw_col:
    st.markdown(
        f"<div style='text-align:center; padding:10px 0;'>"
        f"<div style='font-size:12px; color:#888; font-weight:600; text-transform:uppercase;'>"
        f"Active Swell Direction</div>"
        f"<div class='swell-badge'>{active_swell}</div>"
        f"<div style='font-size:11px; color:#aaa; margin-top:6px;'>{swell_source}</div>"
        f"</div>",
        unsafe_allow_html=True
    )

# Compass rose (mini plotly polar)
with info_col:
    deg = cardinal_to_degrees(active_swell)
    arrow_x = math.sin(math.radians(deg))
    arrow_y = math.cos(math.radians(deg))
    fig_compass = go.Figure()
    fig_compass.add_trace(go.Scatterpolar(
        r=[0, 0.85], theta=[0, deg],
        mode="lines",
        line=dict(color="#1F7A8C", width=4),
        showlegend=False
    ))
    fig_compass.add_trace(go.Scatterpolar(
        r=[0.85], theta=[deg],
        mode="markers",
        marker=dict(symbol="arrow", size=16, color="#1F7A8C",
                    angleref="previous", angle=180),
        showlegend=False
    ))
    fig_compass.update_layout(
        polar=dict(
            radialaxis=dict(visible=False, range=[0, 1]),
            angularaxis=dict(
                tickmode="array",
                tickvals=[0, 45, 90, 135, 180, 225, 270, 315],
                ticktext=["N", "NE", "E", "SE", "S", "SW", "W", "NW"],
                direction="clockwise", rotation=90,
                tickfont=dict(size=11, color="#555")
            ),
            bgcolor="#f8fbff"
        ),
        margin=dict(l=20, r=20, t=20, b=20),
        height=180,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig_compass, use_container_width=False, config={"displayModeBar": False})

st.markdown("---")

# ── Load spots & apply filters ────────────────────────────────────────────────
try:
    df_spots = load_spots()
except Exception as e:
    st.error(f"Could not load Spots Database from Excel: {e}")
    st.stop()

# Match spots to swell
df_spots = match_spots(df_spots, active_swell)

# Region filter
if region_filter:
    df_spots = df_spots[df_spots["Region"].isin(region_filter)]

# Skill filter
if skill_filter != "All Levels":
    df_spots = df_spots[
        df_spots["Skill Level"].str.contains(
            skill_filter.split("–")[0], case=False, na=False
        )
    ]

# Sort: matches first, then by quality desc
df_spots["_sort"] = df_spots["Match"].astype(int) * 10 + df_spots["Quality_Int"]
df_sorted = df_spots.sort_values("_sort", ascending=False).reset_index(drop=True)

# ── Tide ──────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">🌊 Tide — Today\'s Curve</div>',
            unsafe_allow_html=True)

if tide_error or not tide:
    st.warning(f"⚠️ Could not load tide data: {tide_error or 'No data'}")
else:
    tide_left, tide_right = st.columns([3, 1])

    with tide_left:
        fig_tide = go.Figure()

        # Shaded area under curve
        fig_tide.add_trace(go.Scatter(
            x=tide["times"], y=tide["heights"],
            mode="lines", name="Tide Height",
            line=dict(color="#1F7A8C", width=2.5, shape="spline"),
            fill="tozeroy",
            fillcolor="rgba(31,122,140,0.12)",
        ))

        # Hi/Lo markers
        for ev in tide["hilo"]:
            color  = "#1B3A5C" if ev["type"] == "H" else "#E8873A"
            symbol = "triangle-up" if ev["type"] == "H" else "triangle-down"
            fig_tide.add_trace(go.Scatter(
                x=[ev["time"]], y=[ev["height"]],
                mode="markers+text",
                marker=dict(symbol=symbol, size=14, color=color),
                text=[f'{ev["label"]}<br>{ev["height"]}ft<br>{ev["time"].strftime("%I:%M %p")}'],
                textposition="top center" if ev["type"] == "H" else "bottom center",
                textfont=dict(size=10, color=color),
                showlegend=False,
            ))

        # Current time vertical line — use add_shape + add_annotation
        # (avoids a Plotly bug with add_vline on datetime axes in Python 3.14)
        if tide["current_ht"] is not None:
            now_str = tide["now"].strftime("%Y-%m-%d %H:%M")
            fig_tide.add_shape(
                type="line",
                x0=now_str, x1=now_str,
                y0=0, y1=1,
                xref="x", yref="paper",
                line=dict(color="#E8873A", width=2, dash="dash"),
            )
            fig_tide.add_annotation(
                x=now_str, y=1,
                xref="x", yref="paper",
                text=f"NOW  {tide['current_ht']}ft",
                showarrow=False,
                xanchor="left",
                font=dict(color="#E8873A", size=11),
                bgcolor="rgba(255,255,255,0.7)",
            )

        fig_tide.update_layout(
            margin=dict(l=10, r=10, t=10, b=10),
            height=220,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="white",
            showlegend=False,
            xaxis=dict(
                tickformat="%I %p", showgrid=True,
                gridcolor="#f0f0f0", tickfont=dict(size=10),
            ),
            yaxis=dict(
                title="Height (ft)", showgrid=True,
                gridcolor="#f0f0f0", tickfont=dict(size=10),
                rangemode="tozero",
            ),
        )
        st.plotly_chart(fig_tide, use_container_width=True,
                        config={"displayModeBar": False})

    with tide_right:
        ht   = tide["current_ht"]
        st8  = tide["state"]
        nxt  = tide["next_event"]
        nxt_label = ""
        if nxt:
            mins_away = int((nxt["time"] - tide["now"]).total_seconds() / 60)
            hrs, mins = divmod(mins_away, 60)
            nxt_time  = nxt["time"].strftime("%I:%M %p")
            nxt_label = (f"{'⬆️ High' if nxt['type']=='H' else '⬇️ Low'} "
                         f"@ {nxt_time}<br>"
                         f"<span style='color:#aaa'>{hrs}h {mins}m away · "
                         f"{nxt['height']}ft</span>")

        st.markdown(f"""
        <div style="background:white; border-radius:12px; padding:18px;
                    box-shadow:0 2px 8px rgba(0,0,0,0.08); text-align:center;
                    border-top:4px solid #1F7A8C;">
            <div style="font-size:11px; color:#888; font-weight:600;
                        text-transform:uppercase; margin-bottom:6px;">Current Tide</div>
            <div style="font-size:30px; font-weight:800; color:#1B3A5C;">
                {f"{ht} ft" if ht is not None else "—"}</div>
            <div style="font-size:14px; color:#1F7A8C; margin:4px 0 14px 0;">
                {st8}</div>
            <hr style="border-color:#eee; margin:10px 0;">
            <div style="font-size:11px; color:#888; font-weight:600;
                        text-transform:uppercase; margin-bottom:6px;">Next</div>
            <div style="font-size:12px; color:#444; line-height:1.6;">
                {nxt_label if nxt_label else "—"}</div>
            <hr style="border-color:#eee; margin:10px 0;">
            <div style="font-size:10px; color:#bbb;">{selected_tide_name}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("---")

# ── Map ───────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">🗺️ Spot Map</div>', unsafe_allow_html=True)

map_data = []
for _, row in df_sorted.iterrows():
    name = row["Spot Name"]
    coords = SPOT_COORDS.get(name)
    if coords:
        map_data.append({
            "name":   name,
            "lat":    coords[0],
            "lon":    coords[1],
            "match":  row["Match"],
            "quality":row["Quality_Int"],
            "tide":   row.get("Best Tide", ""),
            "skill":  row.get("Skill Level", ""),
            "wave":   row.get("Wave Type", ""),
            "swell":  row.get("Primary Swell", ""),
        })

if map_data:
    fig_map = go.Figure()

    # Non-matching spots (grey)
    grey = [d for d in map_data if not d["match"]]
    if grey:
        fig_map.add_trace(go.Scattermapbox(
            lat=[d["lat"] for d in grey],
            lon=[d["lon"] for d in grey],
            mode="markers+text",
            marker=dict(size=12, color="#cccccc", opacity=0.7),
            text=[d["name"] for d in grey],
            textposition="top right",
            textfont=dict(size=10, color="#aaa"),
            customdata=[[d["tide"], d["skill"], d["wave"], d["swell"], d["quality"]]
                        for d in grey],
            hovertemplate=(
                "<b>%{text}</b><br>"
                "Swell: %{customdata[3]}<br>"
                "Tide: %{customdata[0]}<br>"
                "Wave: %{customdata[2]}<br>"
                "Skill: %{customdata[1]}<br>"
                "Quality: %{customdata[4]}★<br>"
                "<i>Not ideal for current swell</i><extra></extra>"
            ),
            name="Not Ideal",
            showlegend=True,
        ))

    # Matching spots (teal/green, sized by quality)
    green = [d for d in map_data if d["match"]]
    if green:
        sizes = [10 + d["quality"] * 4 for d in green]
        fig_map.add_trace(go.Scattermapbox(
            lat=[d["lat"] for d in green],
            lon=[d["lon"] for d in green],
            mode="markers+text",
            marker=dict(size=sizes, color="#1F7A8C", opacity=0.92),
            text=[d["name"] for d in green],
            textposition="top right",
            textfont=dict(size=11, color="#1B3A5C"),
            customdata=[[d["tide"], d["skill"], d["wave"], d["swell"], d["quality"]]
                        for d in green],
            hovertemplate=(
                "<b>%{text}</b> ✅<br>"
                "Swell: %{customdata[3]}<br>"
                "Tide: %{customdata[0]}<br>"
                "Wave: %{customdata[2]}<br>"
                "Skill: %{customdata[1]}<br>"
                "Quality: %{customdata[4]}★<extra></extra>"
            ),
            name="Good for Current Swell",
            showlegend=True,
        ))

    # Buoy marker
    buoy_coords = {
        "44007": (43.525, -70.141),
        "44013": (42.346, -70.651),
        "44017": (40.693, -72.048),
        "44025": (40.251, -73.166),
        "44027": (44.270, -67.316),
    }
    if buoy_id in buoy_coords:
        blat, blon = buoy_coords[buoy_id]
        fig_map.add_trace(go.Scattermapbox(
            lat=[blat], lon=[blon],
            mode="markers+text",
            marker=dict(size=14, color="#E8873A", symbol="circle"),
            text=[f"BUOY {buoy_id}"],
            textposition="bottom right",
            textfont=dict(size=10, color="#E8873A"),
            hovertemplate=f"<b>NOAA Buoy {buoy_id}</b><br>{selected_buoy_name}<extra></extra>",
            name=f"Buoy {buoy_id}",
            showlegend=True,
        ))

    center_lat = sum(d["lat"] for d in map_data) / len(map_data)
    center_lon = sum(d["lon"] for d in map_data) / len(map_data)

    fig_map.update_layout(
        mapbox=dict(
            style="carto-positron",
            center=dict(lat=center_lat, lon=center_lon),
            zoom=9.5,
        ),
        margin=dict(l=0, r=0, t=0, b=0),
        height=430,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.01,
            xanchor="left", x=0,
            bgcolor="rgba(255,255,255,0.85)",
            bordercolor="#ddd", borderwidth=1,
            font=dict(size=11)
        ),
        paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig_map, use_container_width=True, config={"displayModeBar": False})

st.markdown("---")

# ── Ranked spot table ─────────────────────────────────────────────────────────
match_count = df_sorted["Match"].sum()
total_count  = len(df_sorted)

st.markdown(
    f'<div class="section-title">🏆 Ranked Spots — '
    f'<span style="color:#1F7A8C;">{int(match_count)} of {total_count} spots work '
    f'for {active_swell} swell</span></div>',
    unsafe_allow_html=True
)

DISPLAY_COLS = {
    "Spot Name":     "Spot",
    "Region":        "Region",
    "Primary Swell": "Best Swell",
    "Best Tide":     "Tide",
    "Wave Type":     "Wave Type",
    "Skill Level":   "Skill",
    "Crowd":         "Crowd",
    "Quality ★":     "Quality",
}

for idx, row in df_sorted.iterrows():
    match = row["Match"]
    name  = row["Spot Name"]
    qual  = row["Quality_Int"]
    stars = "★" * qual + "☆" * (5 - qual)
    bg    = "#e8f7f0" if match else "#f9f9f9"
    border = "#1F7A8C" if match else "#ddd"
    icon  = "✅" if match else "—"
    opacity = "1" if match else "0.55"

    with st.container():
        st.markdown(
            f"""<div style="background:{bg}; border-left:4px solid {border};
                border-radius:8px; padding:12px 16px; margin-bottom:8px;
                opacity:{opacity};">
                <span style="font-size:18px;">{icon}</span>
                <strong style="font-size:15px; color:#1B3A5C; margin-left:8px;">{name}</strong>
                <span style="color:#E8873A; margin-left:10px; font-size:15px;">{stars}</span>
                <span style="float:right; font-size:12px; color:#888;">
                    {row.get('Wave Type','—')} &nbsp;|&nbsp;
                    Tide: {row.get('Best Tide','—')} &nbsp;|&nbsp;
                    {row.get('Skill Level','—')} &nbsp;|&nbsp;
                    Swell: {row.get('Primary Swell','—')}
                </span>
            </div>""",
            unsafe_allow_html=True
        )

st.markdown("<br>", unsafe_allow_html=True)

# ── Expandable notes ──────────────────────────────────────────────────────────
with st.expander("📋 Full Spot Details & Notes"):
    for idx, row in df_sorted.iterrows():
        match = row["Match"]
        icon  = "✅" if match else "—"
        st.markdown(f"**{icon} {row['Spot Name']}** — {row.get('Region','')}, {row.get('Town','')}")
        notes = row.get("Notes", "")
        if notes and str(notes) != "nan":
            st.markdown(f"<small style='color:#555;'>{notes}</small>", unsafe_allow_html=True)
        st.markdown("---")

# ── Live Cams ─────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📷 Live Cams — Southern Maine Coast</div>',
            unsafe_allow_html=True)

st.markdown(
    "<p style='font-size:13px; color:#555; margin-bottom:12px;'>"
    "Click any camera to open a live view in a new tab. "
    "<span style='color:#E8873A; font-weight:600;'>Surfline cams</span> require your "
    "Surfline login — full embedded streams coming once the Mac Mini is set up. 🤙"
    "</p>",
    unsafe_allow_html=True
)

CAMS = [
    # (Spot Name, Location Label, URL, Source, Surfline?)
    ("Long Sands Beach",    "York, ME",       "https://www.surfline.com/surf-report/long-sands-beach/5842041f4e65fad6a77089e3?camId=59b04bce375c71ff6213f6d0", "Surfline", True),
    ("Short Sands Beach",   "York, ME",       "https://worldcam.eu/webcams/north-america/maine-usa/9943-york-short-sands-beach",   "WorldCam", False),
    ("Ogunquit Beach",      "Ogunquit, ME",   "https://worldcam.eu/webcams/north-america/maine-usa/30020-ogunquit-beach",           "WorldCam", False),
    ("Ogunquit Beachmere",  "Ogunquit, ME",   "https://worldcam.eu/webcams/north-america/maine-usa/24711-ogunquit-beachmere-inn",  "WorldCam", False),
    ("Ogunquit Sea Chambers","Ogunquit, ME",  "https://worldcam.eu/webcams/north-america/maine-usa/30149-ogunquit-sea-chambers-motel", "WorldCam", False),
    ("Wells Beach",         "Wells, ME",      "https://worldcam.eu/webcams/north-america/maine-usa/9944-wells-beach-lafayettes-oceanfront-resort", "WorldCam", False),
    ("Ogunquit Norseman",   "Ogunquit, ME",   "https://worldcam.eu/webcams/north-america/maine-usa/24694-ogunquit-norseman-resort", "WorldCam", False),
]

cam_cols = st.columns(4)
for i, (name, location, url, source, is_surfline) in enumerate(CAMS):
    col = cam_cols[i % 4]
    with col:
        badge_color = "#E8873A" if is_surfline else "#1F7A8C"
        badge_text  = "🔒 Surfline" if is_surfline else "🌐 Free"
        st.markdown(
            f"""<a href="{url}" target="_blank" style="text-decoration:none;">
            <div style="background:white; border-radius:10px; padding:12px 14px;
                        margin-bottom:10px; box-shadow:0 2px 6px rgba(0,0,0,0.08);
                        border-top:3px solid {badge_color};
                        transition:box-shadow 0.2s;">
                <div style="font-size:13px; font-weight:700; color:#1B3A5C;">{name}</div>
                <div style="font-size:11px; color:#888; margin:2px 0 6px 0;">{location}</div>
                <span style="background:{badge_color}; color:white; font-size:10px;
                             font-weight:600; padding:2px 7px; border-radius:20px;">
                    {badge_text}</span>
                <span style="float:right; font-size:18px;">📷</span>
            </div></a>""",
            unsafe_allow_html=True
        )

st.markdown("---")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown(
    "<div style='text-align:center; color:#aaa; font-size:11px; margin-top:20px;'>"
    "Data sources: NOAA NDBC (live buoy) · The Stormrider Surf Guide North America · "
    "Built with Streamlit + Plotly · "
    f"Page loaded: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    "</div>",
    unsafe_allow_html=True
)
