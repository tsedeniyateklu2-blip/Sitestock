"""
SiteStock v2 - Construction Site Material Management System
Run with:  py -m streamlit run app.py

Features
  * Login with roles (Admin / Storekeeper / Manager)
  * Materials, deliveries (with receipt photo) and daily usage logging
  * Edit / delete records
  * Dashboard: KPIs, alerts, planned-vs-actual, usage trends, cost & budget,
    stock-out / waste forecast, supplier comparison
  * Excel and PDF report export
"""
import hashlib
import html
import io
import os
import secrets
import sqlite3
from contextlib import closing
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = os.environ.get("SITESTOCK_DB", str(BASE_DIR / "sitestock.db"))
RECEIPT_DIR = BASE_DIR / "receipts"
CURRENCY = "ETB"
OVER_USAGE_THRESHOLD = 10.0      # % above expected usage that triggers a warning
DEFAULT_PROGRESS = 75
PALETTE = ["#1e5a8a", "#f28c28", "#2e9e5b", "#d64545", "#7c5cbf", "#16a3b8"]

STATUS_LABEL = {
    "ok": "🟢 OK",
    "low": "🔴 Low stock",
    "over": "🟡 Over-usage",
    "error": "⚫ Data error",
}
STATUS_PLAIN = {"ok": "OK", "low": "LOW STOCK", "over": "OVER-USAGE", "error": "DATA ERROR"}
STATUS_COLOR = {
    "🟢 OK": "#2e9e5b",
    "🔴 Low stock": "#d64545",
    "🟡 Over-usage": "#e0a100",
    "⚫ Data error": "#555555",
}

THEME_TOML = """[theme]
base = "light"
primaryColor = "#f28c28"
backgroundColor = "#f4f7fb"
secondaryBackgroundColor = "#ffffff"
textColor = "#1b2b3a"

[browser]
gatherUsageStats = false
"""


def ensure_theme_file():
    """Create .streamlit/config.toml once so the light theme applies on next start."""
    cfg = Path.cwd() / ".streamlit" / "config.toml"
    if not cfg.exists():
        try:
            cfg.parent.mkdir(exist_ok=True)
            cfg.write_text(THEME_TOML, encoding="utf-8")
        except OSError:
            pass


ensure_theme_file()
st.set_page_config(page_title="SiteStock", page_icon="🏗️", layout="wide")


# --------------------------------------------------------------------------
# Styling
# --------------------------------------------------------------------------
CSS = """
<style>
.block-container{padding-top:1.4rem;padding-bottom:3rem;max-width:1300px}
footer{visibility:hidden}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#0f2a43 0%,#16406a 100%)}
[data-testid="stSidebar"] *{color:#eaf2fa}
.brand{font-size:1.55rem;font-weight:800;letter-spacing:.02em;margin-bottom:2px}
.brand span{color:#f28c28}
.brand-sub{font-size:.78rem;opacity:.75;margin-bottom:14px}
.usercard{background:rgba(255,255,255,.10);border-radius:12px;padding:10px 12px;margin-bottom:14px;font-size:.88rem}
.usercard b{font-size:.95rem}
.hero{background:linear-gradient(120deg,#0f2a43 0%,#1e5a8a 62%,#f28c28 150%);color:#fff;padding:20px 26px;border-radius:16px;display:flex;gap:18px;align-items:center;margin-bottom:18px;box-shadow:0 6px 18px rgba(15,42,67,.18)}
.hero h1{color:#fff !important;margin:0;padding:0;font-size:1.7rem;line-height:1.2}
.hero p{margin:3px 0 0;opacity:.9;color:#fff}
.hero-icon{font-size:2.3rem;background:rgba(255,255,255,.16);border-radius:14px;padding:6px 14px}
.kpi{background:#fff;border-radius:14px;padding:14px 16px;display:flex;gap:12px;align-items:center;box-shadow:0 2px 10px rgba(15,42,67,.08);margin-bottom:10px;min-height:88px}
.kpi-icon{font-size:1.7rem}
.kpi-label{font-size:.72rem;color:#5b6b7b;text-transform:uppercase;letter-spacing:.05em}
.kpi-value{font-size:1.4rem;font-weight:700;color:#0f2a43;line-height:1.25}
.kpi-sub{font-size:.74rem;color:#7a8896}
.note{border-radius:12px;padding:11px 16px;margin:7px 0;font-size:.92rem;border-left:5px solid}
.note.red{background:#fdecec;border-color:#d64545;color:#7a1f1f}
.note.amber{background:#fff6e0;border-color:#e0a100;color:#6b4e00}
.note.green{background:#e8f6ee;border-color:#2e9e5b;color:#1d5b37}
.note.blue{background:#e7f0fa;border-color:#1e5a8a;color:#133a5c}
.section-title{font-size:1.1rem;font-weight:700;color:#0f2a43;margin:18px 0 6px}
.stTabs [data-baseweb="tab-list"]{gap:6px}
.stTabs [data-baseweb="tab"]{background:#fff;border-radius:10px 10px 0 0;padding:8px 16px;height:auto}
.stTabs [aria-selected="true"]{background:#f28c28}
.stTabs [aria-selected="true"] p{color:#fff !important}
div.stButton>button, div.stDownloadButton>button, .stFormSubmitButton>button{background:#f28c28;color:#fff;border:0;border-radius:10px;font-weight:600;padding:.5rem 1.2rem}
div.stButton>button:hover, div.stDownloadButton>button:hover, .stFormSubmitButton>button:hover{background:#d9761a;color:#fff}
[data-testid="stForm"]{background:#fff;border-radius:14px;border:1px solid #e1e8f0;box-shadow:0 2px 10px rgba(15,42,67,.05)}
</style>
"""


CSS_PRO = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, .stApp, [class*="st-"], button, input, textarea, select{font-family:'Inter','Segoe UI',system-ui,-apple-system,sans-serif}
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:translateY(0)}}
@keyframes slideIn{from{opacity:0;transform:translateX(-14px)}to{opacity:1;transform:translateX(0)}}
@keyframes gradientShift{0%{background-position:0% 50%}50%{background-position:100% 50%}100%{background-position:0% 50%}}
@keyframes pulseDot{0%{box-shadow:0 0 0 0 rgba(214,69,69,.55)}70%{box-shadow:0 0 0 9px rgba(214,69,69,0)}100%{box-shadow:0 0 0 0 rgba(214,69,69,0)}}
@keyframes growBar{from{width:0}}

/* blueprint-grid canvas */
.stApp{background-color:#f4f7fb;background-image:radial-gradient(circle at 1px 1px,rgba(30,90,138,.09) 1px,transparent 0);background-size:24px 24px}
::-webkit-scrollbar{width:10px;height:10px}
::-webkit-scrollbar-thumb{background:#b9c8d8;border-radius:8px}
::-webkit-scrollbar-thumb:hover{background:#f28c28}
h1,h2,h3{letter-spacing:-.01em}

/* animated hero */
.hero{position:relative;overflow:hidden;background:linear-gradient(120deg,#0f2a43,#1e5a8a,#f28c28,#1e5a8a,#0f2a43);background-size:300% 300%;animation:gradientShift 14s ease infinite,fadeUp .5s ease both;border:1px solid rgba(255,255,255,.12)}
.hero:before{content:"";position:absolute;right:-60px;top:-70px;width:220px;height:220px;border-radius:50%;background:rgba(255,255,255,.08)}
.hero:after{content:"";position:absolute;right:90px;bottom:-90px;width:170px;height:170px;border-radius:50%;background:rgba(255,255,255,.06)}
.hero h1{font-weight:800;text-shadow:0 2px 8px rgba(0,0,0,.25)}
.hero-icon{backdrop-filter:blur(4px);border:1px solid rgba(255,255,255,.25)}

/* cards */
.kpi{animation:fadeUp .55s ease both;transition:transform .2s ease,box-shadow .2s ease;border:1px solid #e6edf5}
.kpi:hover{transform:translateY(-4px);box-shadow:0 12px 26px rgba(15,42,67,.16)}
.kpi-icon{width:46px;height:46px;display:flex;align-items:center;justify-content:center;border-radius:12px;background:#f1f5fa}
.kpi-value{font-weight:800}
.note{animation:slideIn .45s ease both;box-shadow:0 2px 8px rgba(15,42,67,.05)}
.section-title{border-left:5px solid #f28c28;padding-left:10px;margin:22px 0 10px}

/* material cards */
.mgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:16px;margin:6px 0 10px}
.mcard{background:#fff;border-radius:16px;padding:16px 18px;border:1px solid #e6edf5;box-shadow:0 2px 10px rgba(15,42,67,.06);animation:fadeUp .6s ease both;transition:transform .2s ease,box-shadow .2s ease;position:relative;overflow:hidden}
.mcard:before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:var(--accent,#1e5a8a)}
.mcard:hover{transform:translateY(-5px);box-shadow:0 14px 28px rgba(15,42,67,.16)}
.mcard-top{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:10px}
.mname{font-weight:700;font-size:1.05rem;color:#0f2a43}
.munit{font-size:.74rem;color:#7a8896;font-weight:500}
.mnums{display:flex;justify-content:space-between;gap:6px;margin:6px 0 12px;text-align:center}
.mnums div{flex:1;background:#f5f8fc;border-radius:10px;padding:7px 4px}
.mnums b{display:block;font-size:1.12rem;color:#0f2a43;font-weight:800}
.mnums small{font-size:.68rem;color:#6b7b8b;text-transform:uppercase;letter-spacing:.04em}
.bar-label{display:flex;justify-content:space-between;font-size:.72rem;color:#5b6b7b;margin:7px 0 3px;font-weight:600}
.bar{height:8px;border-radius:8px;background:#e8eef5;overflow:hidden}
.bar span{display:block;height:100%;border-radius:8px;background:linear-gradient(90deg,#1e5a8a,#3d8bc7);animation:growBar 1.1s ease both}
.bar.usage span{background:linear-gradient(90deg,#f28c28,#f7b267)}
.bar.danger span{background:linear-gradient(90deg,#d64545,#ef7f7f)}
.mfoot{display:flex;justify-content:space-between;margin-top:12px;font-size:.76rem;color:#6b7b8b}

/* status pills */
.pill{display:inline-flex;align-items:center;gap:6px;padding:3px 11px;border-radius:999px;font-size:.72rem;font-weight:700;letter-spacing:.02em}
.pill:before{content:"";width:7px;height:7px;border-radius:50%;background:currentColor}
.pill.green{background:#e3f5ea;color:#1f8a4c}
.pill.red{background:#fde8e8;color:#c13030}
.pill.red:before{animation:pulseDot 1.6s infinite}
.pill.amber{background:#fff3d6;color:#a77800}
.pill.grey{background:#ececec;color:#555}

/* sidebar navigation */
[data-testid="stSidebar"]{border-right:1px solid rgba(255,255,255,.08)}
[data-testid="stSidebar"] [role="radiogroup"]{gap:4px}
[data-testid="stSidebar"] [role="radiogroup"] label{padding:9px 12px;border-radius:10px;transition:background .15s ease,transform .15s ease;cursor:pointer;width:100%}
[data-testid="stSidebar"] [role="radiogroup"] label:hover{background:rgba(255,255,255,.10);transform:translateX(3px)}
[data-testid="stSidebar"] [role="radiogroup"] label > div:first-child{display:none}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked){background:linear-gradient(90deg,#f28c28,#f5a54f);box-shadow:0 4px 12px rgba(242,140,40,.35)}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) *{color:#fff !important;font-weight:700}
.brand{display:flex;align-items:center;gap:10px}
.brand svg{flex:none}
.usercard{border:1px solid rgba(255,255,255,.12);display:flex;align-items:center;gap:10px}
.avatar{width:36px;height:36px;border-radius:50%;background:linear-gradient(135deg,#f28c28,#f7b267);display:flex;align-items:center;justify-content:center;font-weight:800;color:#fff;flex:none}

/* inputs, tables, tabs, buttons */
.stTextInput input,.stNumberInput input,.stDateInput input,[data-baseweb="select"]>div,[data-baseweb="textarea"]{border-radius:10px !important}
.stTextInput input:focus,.stNumberInput input:focus{border-color:#f28c28 !important;box-shadow:0 0 0 3px rgba(242,140,40,.2) !important}
[data-testid="stDataFrame"],[data-testid="stDataEditor"]{border-radius:14px;overflow:hidden;box-shadow:0 2px 12px rgba(15,42,67,.08);border:1px solid #e1e8f0}
[data-testid="stPlotlyChart"]{background:#fff;border-radius:16px;padding:8px;box-shadow:0 2px 12px rgba(15,42,67,.07);border:1px solid #e6edf5;animation:fadeUp .6s ease both}
.stTabs [data-baseweb="tab"]{border:1px solid #e1e8f0;border-bottom:none;font-weight:600;transition:background .15s ease}
.stTabs [data-baseweb="tab"]:hover{background:#fff1e0}
.stTabs [data-baseweb="tab-highlight"]{display:none}
div.stButton>button,div.stDownloadButton>button,.stFormSubmitButton>button{box-shadow:0 4px 12px rgba(242,140,40,.3);transition:transform .15s ease,box-shadow .15s ease}
div.stButton>button:hover,div.stDownloadButton>button:hover,.stFormSubmitButton>button:hover{transform:translateY(-2px);box-shadow:0 8px 18px rgba(242,140,40,.4)}
[data-testid="stForm"]{padding:18px 22px}
[data-testid="stExpander"]{border-radius:12px;border:1px solid #e1e8f0;background:#fff}
[data-testid="stSidebar"] [data-testid="stSlider"]{padding:4px 6px}

/* login */
.login-panel{background:linear-gradient(150deg,#0f2a43,#1e5a8a);color:#fff;border-radius:18px;padding:30px 30px 26px;box-shadow:0 14px 34px rgba(15,42,67,.25);animation:fadeUp .6s ease both;min-height:360px}
.login-panel h2{color:#fff !important;font-weight:800;margin:0 0 6px}
.login-panel p{color:#dbe8f5;margin:0 0 14px}
.login-panel ul{list-style:none;padding:0;margin:0}
.login-panel li{padding:7px 0;border-bottom:1px solid rgba(255,255,255,.12);font-size:.93rem;color:#fff}

/* footer */
.app-footer{margin-top:40px;padding:16px 0 4px;text-align:center;font-size:.78rem;color:#7a8896;border-top:1px solid #dfe7ef}
.app-footer b{color:#0f2a43}
</style>
"""

BRAND_SVG = (
    '<svg width="36" height="36" viewBox="0 0 36 36" xmlns="http://www.w3.org/2000/svg">'
    '<rect width="36" height="36" rx="9" fill="#f28c28"/>'
    '<path d="M7 28h22M10 28V14l8-6 8 6v14M15 28v-7h6v7" fill="none" stroke="#fff" stroke-width="2.2" '
    'stroke-linecap="round" stroke-linejoin="round"/></svg>'
)

PILL_CLASS = {"ok": "green", "low": "red", "over": "amber", "error": "grey"}
PILL_TEXT = {"ok": "In good shape", "low": "Low stock", "over": "Over-usage", "error": "Data error"}
ACCENT = {"ok": "#2e9e5b", "low": "#d64545", "over": "#e0a100", "error": "#555555"}


def inject_css():
    st.markdown(CSS + CSS_PRO, unsafe_allow_html=True)


def material_cards(df):
    """Visual card grid: one card per material with progress bars and a status pill."""
    cards = []
    for _, r in df.iterrows():
        key = r["status_key"]
        used_pct = min(100.0, (r["used"] / r["planned_qty"] * 100) if r["planned_qty"] > 0 else 0.0)
        days = f"{r['days_left']:.0f}" if pd.notna(r["days_left"]) else "-"
        bar_cls = "bar usage danger" if key == "over" else "bar usage"
        cards.append(
            f'<div class="mcard" style="--accent:{ACCENT[key]}">'
            f'<div class="mcard-top"><div><div class="mname">{html.escape(str(r["name"]))}</div>'
            f'<div class="munit">{html.escape(str(r["unit"]))}</div></div>'
            f'<span class="pill {PILL_CLASS[key]}">{PILL_TEXT[key]}</span></div>'
            f'<div class="mnums"><div><b>{fmt(r["stock"])}</b><small>in stock</small></div>'
            f'<div><b>{fmt(r["used"])}</b><small>used</small></div>'
            f'<div><b>{days}</b><small>days left</small></div></div>'
            f'<div class="bar-label"><span>Delivered vs plan</span><span>{r["delivered_pct"]:.0f}%</span></div>'
            f'<div class="bar"><span style="width:{r["delivered_pct"]:.0f}%"></span></div>'
            f'<div class="bar-label"><span>Used vs plan</span><span>{used_pct:.0f}%</span></div>'
            f'<div class="{bar_cls}"><span style="width:{used_pct:.0f}%"></span></div>'
            f'<div class="mfoot"><span>Spent {fmt(r["cost"])} {CURRENCY}</span>'
            f'<span>{r["over_usage_pct"]:+.1f}% vs expected</span></div></div>'
        )
    st.markdown('<div class="mgrid">' + "".join(cards) + "</div>", unsafe_allow_html=True)


def footer():
    st.markdown(
        f'<div class="app-footer"><b>SiteStock</b> · Construction Material Management System · '
        f'Built with Python, Streamlit and SQLite · {date.today():%Y}</div>',
        unsafe_allow_html=True,
    )


def hero(title, subtitle, icon="🏗️"):
    st.markdown(
        f'<div class="hero"><div class="hero-icon">{icon}</div>'
        f"<div><h1>{html.escape(title)}</h1><p>{html.escape(subtitle)}</p></div></div>",
        unsafe_allow_html=True,
    )


def kpi(icon, label, value, sub="", color="#1e5a8a"):
    st.markdown(
        f'<div class="kpi" style="border-left:6px solid {color}"><div class="kpi-icon">{icon}</div>'
        f'<div><div class="kpi-label">{html.escape(label)}</div>'
        f'<div class="kpi-value">{html.escape(str(value))}</div>'
        f'<div class="kpi-sub">{html.escape(sub)}</div></div></div>',
        unsafe_allow_html=True,
    )


def note(kind, text_html):
    st.markdown(f'<div class="note {kind}">{text_html}</div>', unsafe_allow_html=True)


def section(title):
    st.markdown(f'<div class="section-title">{html.escape(title)}</div>', unsafe_allow_html=True)


def style_fig(fig, height=380):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=50, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.22),
        font=dict(color="#1b2b3a"),
        title_font_size=16,
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#e3eaf2")
    return fig


def fmt(n, digits=0):
    try:
        return f"{float(n):,.{digits}f}"
    except (TypeError, ValueError):
        return "-"


# --------------------------------------------------------------------------
# Database
# --------------------------------------------------------------------------
def conn():
    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA foreign_keys = ON")
    return con


def run(sql, params=()):
    with closing(conn()) as con:
        con.execute(sql, params)
        con.commit()


def run_many(statements):
    with closing(conn()) as con:
        for sql, params in statements:
            con.execute(sql, params)
        con.commit()


def query(sql, params=()):
    with closing(conn()) as con:
        return pd.read_sql_query(sql, con, params=params)


def _has_column(con, table, column):
    return any(row[1] == column for row in con.execute(f"PRAGMA table_info({table})"))


def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(8)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000).hex()
    return f"{salt}${digest}"


def verify_password(password, stored):
    try:
        salt, _ = stored.split("$", 1)
    except ValueError:
        return False
    return secrets.compare_digest(hash_password(password, salt), stored)


DEMO_USERS = [
    ("admin", "System Admin", "Admin", "admin123"),
    ("store", "Site Storekeeper", "Storekeeper", "store123"),
    ("manager", "Project Manager", "Manager", "manager123"),
]


def init_db():
    with closing(conn()) as con:
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS materials (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                unit TEXT NOT NULL,
                planned_qty REAL NOT NULL,
                reorder_level REAL NOT NULL,
                planned_price REAL NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS deliveries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                material_id INTEGER NOT NULL,
                date TEXT NOT NULL,
                supplier TEXT,
                qty REAL NOT NULL,
                unit_price REAL NOT NULL,
                receipt TEXT,
                FOREIGN KEY (material_id) REFERENCES materials(id)
            );
            CREATE TABLE IF NOT EXISTS usage (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                material_id INTEGER NOT NULL,
                date TEXT NOT NULL,
                area TEXT,
                qty REAL NOT NULL,
                FOREIGN KEY (material_id) REFERENCES materials(id)
            );
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                full_name TEXT NOT NULL,
                role TEXT NOT NULL,
                password_hash TEXT NOT NULL
            );
            """
        )
        # upgrade databases created by the first version of the app
        if not _has_column(con, "materials", "planned_price"):
            con.execute("ALTER TABLE materials ADD COLUMN planned_price REAL NOT NULL DEFAULT 0")
        if not _has_column(con, "deliveries", "receipt"):
            con.execute("ALTER TABLE deliveries ADD COLUMN receipt TEXT")
        if con.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
            for username, full_name, role, pw in DEMO_USERS:
                con.execute(
                    "INSERT INTO users (username, full_name, role, password_hash) VALUES (?,?,?,?)",
                    (username, full_name, role, hash_password(pw)),
                )
        con.commit()


def get_materials():
    return query("SELECT * FROM materials ORDER BY name")


def save_receipt(uploaded):
    if uploaded is None:
        return None
    RECEIPT_DIR.mkdir(exist_ok=True)
    safe = "".join(c for c in uploaded.name if c.isalnum() or c in "._-") or "receipt"
    fname = f"{datetime.now():%Y%m%d_%H%M%S}_{safe}"
    (RECEIPT_DIR / fname).write_bytes(bytes(uploaded.getbuffer()))
    return fname


def load_sample_data():
    today = date.today()

    def d(days_ago):
        return (today - timedelta(days=days_ago)).isoformat()

    materials = [
        ("Cement", "bags", 800, 100, 950),
        ("Rebar (12mm)", "pcs", 1200, 150, 410),
        ("Sand", "m3", 60, 8, 1200),
        ("Hollow blocks", "pcs", 5000, 500, 28),
    ]
    for name, unit, planned, reorder, price in materials:
        run(
            "INSERT OR IGNORE INTO materials (name, unit, planned_qty, reorder_level, planned_price) "
            "VALUES (?,?,?,?,?)",
            (name, unit, planned, reorder, price),
        )
    ids = {r["name"]: int(r["id"]) for _, r in get_materials().iterrows()}

    deliveries = [
        ("Cement", 35, "Habesha Cement", 300, 950), ("Cement", 22, "Habesha Cement", 250, 965),
        ("Cement", 8, "Dire Cement", 150, 980),
        ("Rebar (12mm)", 34, "Steel Plus", 600, 410), ("Rebar (12mm)", 15, "Steel Plus", 400, 420),
        ("Rebar (12mm)", 5, "Addis Steel", 200, 405),
        ("Sand", 34, "Local Quarry", 25, 1200), ("Sand", 14, "Local Quarry", 20, 1250),
        ("Sand", 4, "River Sand Co.", 10, 1180),
        ("Hollow blocks", 25, "BlockWorks", 2000, 28), ("Hollow blocks", 12, "BlockWorks", 1500, 29),
        ("Hollow blocks", 4, "Merkato Blocks", 1000, 27),
    ]
    usages = [
        ("Cement", 33, "Foundation", 60), ("Cement", 30, "Foundation", 70), ("Cement", 26, "Foundation", 55),
        ("Cement", 21, "Columns", 50), ("Cement", 17, "Columns", 60), ("Cement", 12, "Columns", 65),
        ("Cement", 9, "Slab", 70), ("Cement", 6, "Slab", 80), ("Cement", 3, "Slab", 75), ("Cement", 1, "Slab", 60),
        ("Rebar (12mm)", 31, "Foundation", 120), ("Rebar (12mm)", 24, "Foundation", 140),
        ("Rebar (12mm)", 19, "Columns", 150), ("Rebar (12mm)", 14, "Columns", 160),
        ("Rebar (12mm)", 10, "Columns", 130), ("Rebar (12mm)", 5, "Slab", 150), ("Rebar (12mm)", 2, "Slab", 180),
        ("Sand", 28, "Foundation", 6), ("Sand", 20, "Masonry", 8), ("Sand", 15, "Masonry", 7),
        ("Sand", 11, "Masonry", 9), ("Sand", 7, "Masonry", 8), ("Sand", 3, "Masonry", 9), ("Sand", 1, "Plaster", 6),
        ("Hollow blocks", 20, "Ground floor walls", 600), ("Hollow blocks", 16, "Ground floor walls", 700),
        ("Hollow blocks", 12, "Ground floor walls", 650), ("Hollow blocks", 8, "First floor walls", 600),
        ("Hollow blocks", 5, "First floor walls", 500), ("Hollow blocks", 2, "First floor walls", 450),
    ]
    statements = []
    for name, days, supplier, qty, price in deliveries:
        statements.append((
            "INSERT INTO deliveries (material_id, date, supplier, qty, unit_price) VALUES (?,?,?,?,?)",
            (ids[name], d(days), supplier, qty, price),
        ))
    for name, days, area, qty in usages:
        statements.append((
            "INSERT INTO usage (material_id, date, area, qty) VALUES (?,?,?,?)",
            (ids[name], d(days), area, qty),
        ))
    run_many(statements)


def reset_all_data():
    run_many([("DELETE FROM usage", ()), ("DELETE FROM deliveries", ()), ("DELETE FROM materials", ())])


def delete_material_cascade(material_id):
    run_many([
        ("DELETE FROM usage WHERE material_id = ?", (material_id,)),
        ("DELETE FROM deliveries WHERE material_id = ?", (material_id,)),
        ("DELETE FROM materials WHERE id = ?", (material_id,)),
    ])


# --------------------------------------------------------------------------
# Calculations
# --------------------------------------------------------------------------
def _status_of(row):
    if row["stock"] < 0:
        return "error"
    if row["stock"] <= row["reorder_level"]:
        return "low"
    if row["over_usage_pct"] > OVER_USAGE_THRESHOLD:
        return "over"
    return "ok"


def get_summary(progress_pct):
    """One row per material with stock, waste, cost, budget and forecast figures."""
    df = query(
        """
        SELECT m.id, m.name, m.unit, m.planned_qty, m.reorder_level, m.planned_price,
               COALESCE((SELECT SUM(qty) FROM deliveries d WHERE d.material_id = m.id), 0) AS delivered,
               COALESCE((SELECT SUM(qty) FROM usage u WHERE u.material_id = m.id), 0) AS used,
               COALESCE((SELECT SUM(qty * unit_price) FROM deliveries d WHERE d.material_id = m.id), 0) AS cost
        FROM materials m
        ORDER BY m.name
        """
    )
    if df.empty:
        return df

    progress = float(progress_pct)
    df["stock"] = df["delivered"] - df["used"]
    df["expected_used"] = df["planned_qty"] * progress / 100.0
    df["over_usage_pct"] = (
        (df["used"] - df["expected_used"]) / df["expected_used"].where(df["expected_used"] > 0) * 100
    ).fillna(0.0)
    df["avg_price"] = (df["cost"] / df["delivered"].where(df["delivered"] > 0)).fillna(df["planned_price"])
    df["waste_qty"] = (df["used"] - df["expected_used"]).clip(lower=0) if progress > 0 else 0.0
    df["waste_cost"] = df["waste_qty"] * df["avg_price"]
    df["budget"] = df["planned_qty"] * df["planned_price"]
    df["delivered_pct"] = (df["delivered"] / df["planned_qty"].where(df["planned_qty"] > 0) * 100).fillna(0).clip(0, 100)

    # forecast
    df["projected_total"] = df["used"] / (progress / 100.0) if progress > 0 else df["planned_qty"]
    df["to_order"] = (df["projected_total"] - df["delivered"]).clip(lower=0)
    df["projected_spend"] = df["projected_total"] * df["avg_price"]

    usage = query("SELECT material_id, date, qty FROM usage")
    usage["date"] = pd.to_datetime(usage["date"], errors="coerce")
    today = pd.Timestamp(date.today())
    rates = []
    for _, row in df.iterrows():
        u = usage[usage["material_id"] == row["id"]].dropna(subset=["date"])
        rate = 0.0
        if not u.empty:
            recent = u[u["date"] >= today - pd.Timedelta(days=13)]["qty"].sum()
            if recent > 0:
                rate = recent / 14.0
            else:
                span = max(1, (u["date"].max() - u["date"].min()).days + 1)
                rate = u["qty"].sum() / span
        rates.append(rate)
    df["daily_rate"] = rates
    df["days_left"] = [
        (max(s, 0) / r) if r > 0 else float("nan") for s, r in zip(df["stock"], df["daily_rate"])
    ]
    df["runout_date"] = [
        (date.today() + timedelta(days=float(x))).strftime("%d %b %Y") if pd.notna(x) else "-"
        for x in df["days_left"]
    ]

    df["status_key"] = df.apply(_status_of, axis=1)
    df["status_label"] = df["status_key"].map(STATUS_LABEL)
    return df


def usage_with_names():
    df = query(
        "SELECT u.id, m.name AS material, u.date, u.area, u.qty "
        "FROM usage u JOIN materials m ON m.id = u.material_id ORDER BY u.date"
    )
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["area"] = df["area"].fillna("").replace("", "Unspecified")
    return df


def deliveries_with_names():
    df = query(
        "SELECT d.id, m.name AS material, d.date, d.supplier, d.qty, d.unit_price, d.receipt "
        "FROM deliveries d JOIN materials m ON m.id = d.material_id ORDER BY d.date"
    )
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["supplier"] = df["supplier"].fillna("").replace("", "Unknown")
    return df


# --------------------------------------------------------------------------
# Report export
# --------------------------------------------------------------------------
def build_excel(progress):
    from openpyxl.styles import Alignment, Font, PatternFill

    df = get_summary(progress)
    summary = df[[
        "name", "unit", "planned_qty", "delivered", "used", "stock", "reorder_level", "expected_used",
        "over_usage_pct", "waste_cost", "avg_price", "cost", "budget", "status_label",
    ]].copy()
    summary["status_label"] = df["status_key"].map(STATUS_PLAIN)
    summary.columns = [
        "Material", "Unit", "Planned qty", "Delivered", "Used", "Stock on hand", "Reorder level",
        "Expected used", "Over-usage %", f"Waste cost ({CURRENCY})", f"Avg price ({CURRENCY})",
        f"Spent ({CURRENCY})", f"Budget ({CURRENCY})", "Status",
    ]
    forecast = df[[
        "name", "unit", "daily_rate", "days_left", "runout_date", "projected_total", "planned_qty", "to_order",
        "projected_spend", "budget",
    ]].copy()
    forecast.columns = [
        "Material", "Unit", "Daily usage rate", "Days of stock left", "Estimated run-out", "Projected total need",
        "Planned qty", "Still to order", f"Projected spend ({CURRENCY})", f"Budget ({CURRENCY})",
    ]
    deliveries = deliveries_with_names().drop(columns=["id", "receipt"])
    deliveries["date"] = deliveries["date"].dt.strftime("%Y-%m-%d")
    deliveries.columns = ["Material", "Date", "Supplier", "Quantity", f"Unit price ({CURRENCY})"]
    usage = usage_with_names().drop(columns=["id"])
    usage["date"] = usage["date"].dt.strftime("%Y-%m-%d")
    usage.columns = ["Material", "Date", "Work area", "Quantity"]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, frame in [("Summary", summary), ("Forecast", forecast), ("Deliveries", deliveries), ("Usage", usage)]:
            frame.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="0F2A43")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col) + 3
                ws.column_dimensions[col[0].column_letter].width = min(max(width, 12), 34)
            ws.freeze_panes = "A2"
    return buf.getvalue()


def _pdf_text(value):
    return str(value).encode("latin-1", "replace").decode("latin-1")


def build_pdf(progress):
    from fpdf import FPDF

    df = get_summary(progress)
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_fill_color(15, 42, 67)
    pdf.rect(0, 0, 297, 28, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_xy(10, 7)
    pdf.cell(0, 9, "SiteStock - Material Status Report")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_xy(10, 18)
    pdf.cell(0, 6, f"Generated {date.today():%d %b %Y}   |   Project progress: {progress}%")
    pdf.set_text_color(27, 43, 58)
    pdf.set_y(34)

    if df.empty:
        pdf.set_font("Helvetica", "", 12)
        pdf.cell(0, 10, "No materials recorded yet.")
        return bytes(pdf.output())

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(
        0, 7,
        _pdf_text(
            f"Total spent: {df['cost'].sum():,.0f} {CURRENCY}   |   Budget: {df['budget'].sum():,.0f} {CURRENCY}   |   "
            f"Estimated waste cost: {df['waste_cost'].sum():,.0f} {CURRENCY}"
        ),
        new_x="LMARGIN", new_y="NEXT",
    )
    pdf.ln(3)

    cols = [("Material", 55), ("Unit", 18), ("Planned", 24), ("Delivered", 24), ("Used", 24),
            ("Stock", 24), ("Over %", 22), (f"Cost ({CURRENCY})", 32), ("Status", 54)]
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(15, 42, 67)
    pdf.set_text_color(255, 255, 255)
    for i, (title, w) in enumerate(cols):
        last = i == len(cols) - 1
        pdf.cell(w, 8, title, border=1, fill=True, align="C",
                 new_x="LMARGIN" if last else "RIGHT", new_y="NEXT" if last else "TOP")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(27, 43, 58)
    for n, (_, r) in enumerate(df.iterrows()):
        pdf.set_fill_color(240, 245, 250) if n % 2 == 0 else pdf.set_fill_color(255, 255, 255)
        values = [
            _pdf_text(r["name"]), _pdf_text(r["unit"]), fmt(r["planned_qty"]), fmt(r["delivered"]), fmt(r["used"]),
            fmt(r["stock"]), f"{r['over_usage_pct']:+.1f}", fmt(r["cost"]), STATUS_PLAIN[r["status_key"]],
        ]
        for i, ((_, w), v) in enumerate(zip(cols, values)):
            last = i == len(cols) - 1
            pdf.cell(w, 8, v, border=1, fill=True, align="L" if i in (0, 8) else "R",
                     new_x="LMARGIN" if last else "RIGHT", new_y="NEXT" if last else "TOP")

    pdf.ln(6)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Alerts and forecast", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    lines = []
    for _, r in df.iterrows():
        if r["status_key"] == "low":
            lines.append(f"- {r['name']}: LOW STOCK - {fmt(r['stock'])} {r['unit']} left (reorder level {fmt(r['reorder_level'])}).")
        if r["status_key"] == "over":
            lines.append(f"- {r['name']}: used {r['over_usage_pct']:.1f}% more than expected at {progress}% progress.")
        if pd.notna(r["days_left"]) and r["days_left"] <= 14:
            lines.append(f"- {r['name']}: about {r['days_left']:.0f} days of stock left (around {r['runout_date']}).")
        if r["to_order"] > 0:
            lines.append(f"- {r['name']}: about {fmt(r['to_order'])} {r['unit']} still to order to finish at the current usage rate.")
    if not lines:
        lines = ["- All materials are within limits."]
    for line in lines:
        pdf.multi_cell(0, 6, _pdf_text(line), new_x="LMARGIN", new_y="NEXT")

    pdf.set_y(-12)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(120, 130, 140)
    pdf.cell(0, 6, "Generated by SiteStock", align="C")
    return bytes(pdf.output())


# --------------------------------------------------------------------------
# Pages
# --------------------------------------------------------------------------
def current_progress():
    return st.session_state.get("progress", DEFAULT_PROGRESS)


def empty_state():
    note("blue", "No materials yet. Add materials in <b>Materials</b>, or load an example in <b>Sample data / reset</b>.")


def page_dashboard():
    hero("Dashboard", "Live overview of stock, usage, cost and forecasts", "📊")
    progress = current_progress()
    df = get_summary(progress)
    if df.empty:
        empty_state()
        return

    low = df[df["status_key"] == "low"]
    over = df[df["status_key"] == "over"]
    total_spent = df["cost"].sum()
    total_budget = df["budget"].sum()
    used_pct = (total_spent / total_budget * 100) if total_budget > 0 else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        kpi("📦", "Materials tracked", len(df), f"progress set to {progress}%", "#1e5a8a")
    with c2:
        kpi("💰", f"Total spent ({CURRENCY})", fmt(total_spent), f"{used_pct:.0f}% of budget" if total_budget else "no budget set", "#f28c28")
    with c3:
        kpi("🗑️", f"Est. waste ({CURRENCY})", fmt(df["waste_cost"].sum()), "usage above plan", "#e0a100")
    with c4:
        kpi("🔴", "Low-stock items", len(low), "at or below reorder level", "#d64545")
    with c5:
        kpi("🟡", "Over-usage items", len(over), f"> {OVER_USAGE_THRESHOLD:.0f}% above plan", "#7c5cbf")

    tabs = st.tabs(["Overview", "Usage trends", "Cost & budget", "Forecast", "Suppliers"])

    # ---- Overview
    with tabs[0]:
        if low.empty and over.empty:
            note("green", "✅ All materials are within limits.")
        for _, r in low.iterrows():
            tail = f" At the current rate it lasts about <b>{r['days_left']:.0f} days</b>." if pd.notna(r["days_left"]) else ""
            note("red", f"🔴 <b>{html.escape(r['name'])}</b> is low: {fmt(r['stock'])} {html.escape(r['unit'])} left "
                        f"(reorder level {fmt(r['reorder_level'])}).{tail}")
        for _, r in over.iterrows():
            note("amber", f"🟡 <b>{html.escape(r['name'])}</b> used <b>{r['over_usage_pct']:.1f}%</b> more than expected at {progress}% progress.")

        section("Materials at a glance")
        material_cards(df)

        section("Material summary")
        table = df[["name", "unit", "planned_qty", "delivered", "delivered_pct", "used", "stock",
                    "reorder_level", "over_usage_pct", "cost", "status_label"]]
        st.dataframe(
            table,
            hide_index=True,
            column_config={
                "name": "Material",
                "unit": "Unit",
                "planned_qty": st.column_config.NumberColumn("Planned", format="%.0f"),
                "delivered": st.column_config.NumberColumn("Delivered", format="%.0f"),
                "delivered_pct": st.column_config.ProgressColumn("Delivered vs plan (%)", min_value=0, max_value=100, format="%.0f"),
                "used": st.column_config.NumberColumn("Used", format="%.0f"),
                "stock": st.column_config.NumberColumn("Stock on hand", format="%.0f"),
                "reorder_level": st.column_config.NumberColumn("Reorder level", format="%.0f"),
                "over_usage_pct": st.column_config.NumberColumn("Over-usage (%)", format="%.1f"),
                "cost": st.column_config.NumberColumn(f"Cost ({CURRENCY})", format="%.0f"),
                "status_label": "Status",
            },
        )

        left, right = st.columns(2)
        with left:
            long = df.melt(id_vars=["name"], value_vars=["expected_used", "used"], var_name="Series", value_name="Quantity")
            long["Series"] = long["Series"].map({"expected_used": "Expected", "used": "Actual"})
            fig = px.bar(long, x="name", y="Quantity", color="Series", barmode="group",
                         color_discrete_map={"Expected": "#9bb7d4", "Actual": "#f28c28"},
                         title=f"Expected vs actual usage at {progress}% progress", labels={"name": ""})
            st.plotly_chart(style_fig(fig))
        with right:
            fig = px.bar(df, x="name", y="stock", color="status_label", color_discrete_map=STATUS_COLOR,
                         title="Stock on hand vs reorder level", labels={"name": "", "stock": "Stock", "status_label": "Status"})
            fig.add_trace(go.Scatter(x=df["name"], y=df["reorder_level"], mode="markers", name="Reorder level",
                                     marker=dict(symbol="line-ew", size=38, line=dict(width=3, color="#1b2b3a"))))
            st.plotly_chart(style_fig(fig))

    # ---- Usage trends
    with tabs[1]:
        usage = usage_with_names()
        if usage.empty:
            note("blue", "No usage logged yet.")
        else:
            cum = usage.sort_values("date").copy()
            cum["Cumulative usage"] = cum.groupby("material")["qty"].cumsum()
            fig = px.line(cum, x="date", y="Cumulative usage", color="material", markers=True,
                          color_discrete_sequence=PALETTE, title="Cumulative usage over time", labels={"date": "", "material": "Material"})
            st.plotly_chart(style_fig(fig))

            choice = st.selectbox("Show daily usage by work area for", list(df["name"]), key="trend_material")
            one = usage[usage["material"] == choice]
            if one.empty:
                note("blue", "No usage logged for this material.")
            else:
                fig = px.bar(one, x="date", y="qty", color="area", color_discrete_sequence=PALETTE,
                             title=f"{choice}: usage per entry by work area", labels={"date": "", "qty": "Quantity", "area": "Work area"})
                st.plotly_chart(style_fig(fig))

    # ---- Cost & budget
    with tabs[2]:
        long = df.melt(id_vars=["name"], value_vars=["budget", "cost", "projected_spend"], var_name="Series", value_name="Amount")
        long["Series"] = long["Series"].map({"budget": "Budget", "cost": "Spent so far", "projected_spend": "Projected final spend"})
        fig = px.bar(long, x="name", y="Amount", color="Series", barmode="group",
                     color_discrete_map={"Budget": "#9bb7d4", "Spent so far": "#f28c28", "Projected final spend": "#d64545"},
                     title=f"Budget vs actual and projected spend ({CURRENCY})", labels={"name": ""})
        st.plotly_chart(style_fig(fig))
        if (df["budget"] <= 0).all():
            note("blue", "Tip: set a <b>planned unit price</b> on each material to see budget comparisons.")
        else:
            over_budget = df[(df["budget"] > 0) & (df["projected_spend"] > df["budget"])]
            for _, r in over_budget.iterrows():
                note("amber", f"💸 <b>{html.escape(r['name'])}</b> is projected to cost {fmt(r['projected_spend'])} {CURRENCY}, "
                              f"over its budget of {fmt(r['budget'])} {CURRENCY}.")

        left, right = st.columns(2)
        usage = usage_with_names()
        with left:
            if usage.empty:
                note("blue", "No usage logged yet.")
            else:
                priced = usage.merge(df[["name", "avg_price"]], left_on="material", right_on="name")
                priced["cost"] = priced["qty"] * priced["avg_price"]
                by_area = priced.groupby("area", as_index=False)["cost"].sum()
                fig = px.pie(by_area, names="area", values="cost", hole=0.5, color_discrete_sequence=PALETTE,
                             title=f"Material cost by work area ({CURRENCY})")
                st.plotly_chart(style_fig(fig))
        with right:
            deliveries = deliveries_with_names()
            if deliveries.empty:
                note("blue", "No deliveries logged yet.")
            else:
                fig = px.line(deliveries, x="date", y="unit_price", color="material", markers=True,
                              color_discrete_sequence=PALETTE, title=f"Unit price per delivery ({CURRENCY})",
                              labels={"date": "", "unit_price": "Unit price", "material": "Material"})
                st.plotly_chart(style_fig(fig))

    # ---- Forecast
    with tabs[3]:
        note("blue", "Forecast uses the <b>last 14 days</b> of usage to estimate daily consumption, and the progress slider to project the total need.")
        fc = df.copy()
        fc["urgency"] = fc["days_left"].apply(
            lambda x: "No usage data" if pd.isna(x) else ("Critical (< 7 days)" if x < 7 else ("Soon (7-14 days)" if x < 14 else "Comfortable"))
        )
        for _, r in fc.iterrows():
            if pd.notna(r["days_left"]) and r["days_left"] < 14:
                kind = "red" if r["days_left"] < 7 else "amber"
                note(kind, f"⏳ <b>{html.escape(r['name'])}</b> will run out in about <b>{r['days_left']:.0f} days</b> (around {r['runout_date']}).")
            if r["to_order"] > 0:
                note("blue", f"🛒 <b>{html.escape(r['name'])}</b>: about <b>{fmt(r['to_order'])} {html.escape(r['unit'])}</b> still to order to finish the project at the current rate.")
        shown = fc[fc["days_left"].notna()]
        if not shown.empty:
            fig = px.bar(shown, y="name", x="days_left", orientation="h", color="urgency",
                         color_discrete_map={"Critical (< 7 days)": "#d64545", "Soon (7-14 days)": "#e0a100", "Comfortable": "#2e9e5b"},
                         title="Days of stock left at the current usage rate", labels={"name": "", "days_left": "Days"})
            st.plotly_chart(style_fig(fig, 340))

        section("Forecast table")
        ftable = fc[["name", "unit", "daily_rate", "days_left", "runout_date", "projected_total", "planned_qty", "to_order"]]
        st.dataframe(
            ftable,
            hide_index=True,
            column_config={
                "name": "Material",
                "unit": "Unit",
                "daily_rate": st.column_config.NumberColumn("Daily usage", format="%.1f"),
                "days_left": st.column_config.NumberColumn("Days left", format="%.0f"),
                "runout_date": "Estimated run-out",
                "projected_total": st.column_config.NumberColumn("Projected total need", format="%.0f"),
                "planned_qty": st.column_config.NumberColumn("Planned", format="%.0f"),
                "to_order": st.column_config.NumberColumn("Still to order", format="%.0f"),
            },
        )

    # ---- Suppliers
    with tabs[4]:
        deliveries = deliveries_with_names()
        if deliveries.empty:
            note("blue", "No deliveries logged yet.")
        else:
            deliveries["spend"] = deliveries["qty"] * deliveries["unit_price"]
            sup = deliveries.groupby(["material", "supplier"], as_index=False).agg(
                deliveries=("id", "count"), quantity=("qty", "sum"), spend=("spend", "sum"))
            sup["avg_price"] = sup["spend"] / sup["quantity"]
            choice = st.selectbox("Compare suppliers for", list(df["name"]), key="supplier_material")
            one = sup[sup["material"] == choice].sort_values("avg_price")
            if one.empty:
                note("blue", "No deliveries for this material yet.")
            else:
                best = one.iloc[0]
                if len(one) > 1:
                    worst = one.iloc[-1]
                    saving = (worst["avg_price"] - best["avg_price"]) * one["quantity"].sum()
                    note("green", f"🏆 Cheapest for <b>{html.escape(choice)}</b>: <b>{html.escape(best['supplier'])}</b> at "
                                  f"{best['avg_price']:,.0f} {CURRENCY} per unit. Buying everything from them instead of the dearest "
                                  f"supplier would have saved about {saving:,.0f} {CURRENCY}.")
                else:
                    note("blue", f"Only one supplier so far for {html.escape(choice)}: <b>{html.escape(best['supplier'])}</b>.")
                fig = px.bar(one, x="supplier", y="avg_price", color="supplier", color_discrete_sequence=PALETTE,
                             title=f"Average unit price by supplier: {choice} ({CURRENCY})", labels={"supplier": "", "avg_price": "Avg unit price"})
                fig.update_layout(showlegend=False)
                st.plotly_chart(style_fig(fig, 340))
            section("All suppliers")
            st.dataframe(
                sup.sort_values(["material", "avg_price"]),
                hide_index=True,
                column_config={
                    "material": "Material", "supplier": "Supplier",
                    "deliveries": st.column_config.NumberColumn("Deliveries", format="%d"),
                    "quantity": st.column_config.NumberColumn("Total quantity", format="%.0f"),
                    "spend": st.column_config.NumberColumn(f"Total spend ({CURRENCY})", format="%.0f"),
                    "avg_price": st.column_config.NumberColumn(f"Avg unit price ({CURRENCY})", format="%.0f"),
                },
            )


def page_materials():
    hero("Materials & plan", "Define each material with its planned quantity, price and reorder level", "🧱")
    note("blue", "Enter the planned total quantity from the BOQ. The <b>planned unit price</b> is used for budget comparison.")

    with st.form("add_material", clear_on_submit=True):
        c1, c2 = st.columns(2)
        name = c1.text_input("Material name (e.g. Cement)")
        unit = c2.text_input("Unit (e.g. bags, m3, pcs)")
        c3, c4, c5 = st.columns(3)
        planned = c3.number_input("Planned total quantity", min_value=0.0, step=1.0)
        price = c4.number_input(f"Planned unit price ({CURRENCY})", min_value=0.0, step=1.0)
        reorder = c5.number_input("Reorder level (alert at or below)", min_value=0.0, step=1.0)
        if st.form_submit_button("➕ Add material"):
            if not name.strip() or not unit.strip() or planned <= 0:
                st.error("Please enter a name, a unit and a planned quantity above 0.")
            else:
                try:
                    run("INSERT INTO materials (name, unit, planned_qty, reorder_level, planned_price) VALUES (?,?,?,?,?)",
                        (name.strip(), unit.strip(), planned, reorder, price))
                    st.success(f"Added {name.strip()}.")
                except sqlite3.IntegrityError:
                    st.error("That material already exists.")

    mats = get_materials()
    if mats.empty:
        return
    section("Edit materials")
    mats["delete"] = False
    edited = st.data_editor(
        mats, key="materials_editor", hide_index=True, disabled=["id"],
        column_config={
            "id": None,
            "name": "Material", "unit": "Unit",
            "planned_qty": st.column_config.NumberColumn("Planned qty", min_value=0.0),
            "reorder_level": st.column_config.NumberColumn("Reorder level", min_value=0.0),
            "planned_price": st.column_config.NumberColumn(f"Planned price ({CURRENCY})", min_value=0.0),
            "delete": st.column_config.CheckboxColumn("Delete?"),
        },
    )
    confirm = st.checkbox("I understand deleting a material also deletes all its deliveries and usage", key="confirm_del_mat")
    if st.button("💾 Save material changes"):
        try:
            for _, r in edited.iterrows():
                if r["delete"]:
                    if confirm:
                        delete_material_cascade(int(r["id"]))
                else:
                    run("UPDATE materials SET name=?, unit=?, planned_qty=?, reorder_level=?, planned_price=? WHERE id=?",
                        (str(r["name"]).strip(), str(r["unit"]).strip(), float(r["planned_qty"]),
                         float(r["reorder_level"]), float(r["planned_price"]), int(r["id"])))
            if edited["delete"].any() and not confirm:
                st.warning("Tick the confirmation box to delete materials. Other changes were saved.")
            else:
                st.success("Changes saved.")
        except sqlite3.IntegrityError:
            st.error("Two materials cannot share the same name.")


def page_delivery():
    hero("Log a delivery", "Record what arrived on site, from whom and at what price", "🚚")
    mats = get_materials()
    if mats.empty:
        empty_state()
        return
    with st.form("delivery", clear_on_submit=True):
        name = st.selectbox("Material", list(mats["name"]))
        c1, c2 = st.columns(2)
        d = c1.date_input("Delivery date", value=date.today())
        supplier = c2.text_input("Supplier")
        c3, c4 = st.columns(2)
        qty = c3.number_input("Quantity delivered", min_value=0.0, step=1.0)
        price = c4.number_input(f"Unit price ({CURRENCY})", min_value=0.0, step=1.0)
        receipt = st.file_uploader("Receipt photo (optional)", type=["png", "jpg", "jpeg", "pdf"])
        if st.form_submit_button("💾 Save delivery"):
            if qty <= 0:
                st.error("Quantity must be greater than 0.")
            else:
                mid = int(mats.loc[mats["name"] == name, "id"].iloc[0])
                fname = save_receipt(receipt)
                run("INSERT INTO deliveries (material_id, date, supplier, qty, unit_price, receipt) VALUES (?,?,?,?,?,?)",
                    (mid, d.isoformat(), supplier.strip(), qty, price, fname))
                st.success("Delivery saved." + (" Receipt attached." if fname else ""))


def page_usage():
    hero("Log daily usage", "Record the material used and where it went", "🏗️")
    mats = get_materials()
    if mats.empty:
        empty_state()
        return
    with st.form("usage", clear_on_submit=True):
        name = st.selectbox("Material", list(mats["name"]))
        c1, c2 = st.columns(2)
        d = c1.date_input("Date used", value=date.today())
        area = c2.text_input("Work area / task (e.g. Columns, Slab)")
        qty = st.number_input("Quantity used", min_value=0.0, step=1.0)
        if st.form_submit_button("💾 Save usage"):
            if qty <= 0:
                st.error("Quantity must be greater than 0.")
            else:
                mid = int(mats.loc[mats["name"] == name, "id"].iloc[0])
                stock = float(query(
                    "SELECT COALESCE((SELECT SUM(qty) FROM deliveries WHERE material_id=?),0) - "
                    "COALESCE((SELECT SUM(qty) FROM usage WHERE material_id=?),0) AS stock",
                    (mid, mid))["stock"].iloc[0])
                run("INSERT INTO usage (material_id, date, area, qty) VALUES (?,?,?,?)",
                    (mid, d.isoformat(), area.strip(), qty))
                st.success("Usage saved.")
                if qty > stock:
                    st.warning(f"You used more than the recorded stock ({stock:,.0f}). Check for a missing delivery entry.")


def _to_date(value):
    ts = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(ts) else ts.date().isoformat()


def page_records():
    hero("Edit records", "Fix mistakes: change values or tick Delete and save", "✏️")
    tab_d, tab_u = st.tabs(["Deliveries", "Usage"])

    with tab_d:
        df = deliveries_with_names().sort_values("date", ascending=False)
        if df.empty:
            note("blue", "No deliveries yet.")
        else:
            df["date"] = df["date"].dt.date
            df["delete"] = False
            edited = st.data_editor(
                df, key="deliveries_editor", hide_index=True, disabled=["id", "material", "receipt"],
                column_config={
                    "id": None, "material": "Material",
                    "date": st.column_config.DateColumn("Date"),
                    "supplier": "Supplier",
                    "qty": st.column_config.NumberColumn("Quantity", min_value=0.0),
                    "unit_price": st.column_config.NumberColumn(f"Unit price ({CURRENCY})", min_value=0.0),
                    "receipt": "Receipt file",
                    "delete": st.column_config.CheckboxColumn("Delete?"),
                },
            )
            if st.button("💾 Save delivery changes"):
                for _, r in edited.iterrows():
                    if r["delete"]:
                        run("DELETE FROM deliveries WHERE id = ?", (int(r["id"]),))
                    elif float(r["qty"]) > 0 and _to_date(r["date"]):
                        run("UPDATE deliveries SET date=?, supplier=?, qty=?, unit_price=? WHERE id=?",
                            (_to_date(r["date"]), str(r["supplier"] or "").strip(), float(r["qty"]),
                             float(r["unit_price"]), int(r["id"])))
                st.success("Deliveries updated.")
            with_receipt = df[df["receipt"].notna()]
            if not with_receipt.empty:
                section("View a receipt")
                labels = {int(r["id"]): f"#{int(r['id'])}  {r['material']}  {r['date']}  ({r['receipt']})" for _, r in with_receipt.iterrows()}
                pick = st.selectbox("Delivery with receipt", list(labels), format_func=lambda k: labels[k])
                fname = with_receipt[with_receipt["id"] == pick]["receipt"].iloc[0]
                path = RECEIPT_DIR / fname
                if path.exists():
                    if path.suffix.lower() in (".png", ".jpg", ".jpeg"):
                        st.image(str(path), width=420)
                    else:
                        st.download_button("Download receipt (PDF)", data=path.read_bytes(), file_name=fname)
                else:
                    st.warning("Receipt file not found in the receipts folder.")

    with tab_u:
        df = usage_with_names().sort_values("date", ascending=False)
        if df.empty:
            note("blue", "No usage yet.")
        else:
            df["date"] = df["date"].dt.date
            df["delete"] = False
            edited = st.data_editor(
                df, key="usage_editor", hide_index=True, disabled=["id", "material"],
                column_config={
                    "id": None, "material": "Material",
                    "date": st.column_config.DateColumn("Date"),
                    "area": "Work area",
                    "qty": st.column_config.NumberColumn("Quantity", min_value=0.0),
                    "delete": st.column_config.CheckboxColumn("Delete?"),
                },
            )
            if st.button("💾 Save usage changes"):
                for _, r in edited.iterrows():
                    if r["delete"]:
                        run("DELETE FROM usage WHERE id = ?", (int(r["id"]),))
                    elif float(r["qty"]) > 0 and _to_date(r["date"]):
                        area = "" if r["area"] == "Unspecified" else str(r["area"] or "").strip()
                        run("UPDATE usage SET date=?, area=?, qty=? WHERE id=?",
                            (_to_date(r["date"]), area, float(r["qty"]), int(r["id"])))
                st.success("Usage updated.")


def page_reports():
    hero("Reports", "Download the weekly status report as Excel or PDF", "📄")
    progress = current_progress()
    df = get_summary(progress)
    if df.empty:
        empty_state()
        return
    c1, c2, c3 = st.columns(3)
    with c1:
        kpi("💰", f"Total spent ({CURRENCY})", fmt(df["cost"].sum()), "all deliveries", "#f28c28")
    with c2:
        kpi("🗑️", f"Estimated waste ({CURRENCY})", fmt(df["waste_cost"].sum()), f"at {progress}% progress", "#e0a100")
    with c3:
        kpi("🔴", "Items needing attention", int((df["status_key"] != "ok").sum()), "low stock or over-usage", "#d64545")

    section("Preview")
    st.dataframe(
        df[["name", "unit", "planned_qty", "delivered", "used", "stock", "over_usage_pct", "cost", "status_label"]],
        hide_index=True,
        column_config={
            "name": "Material", "unit": "Unit",
            "planned_qty": st.column_config.NumberColumn("Planned", format="%.0f"),
            "delivered": st.column_config.NumberColumn("Delivered", format="%.0f"),
            "used": st.column_config.NumberColumn("Used", format="%.0f"),
            "stock": st.column_config.NumberColumn("Stock", format="%.0f"),
            "over_usage_pct": st.column_config.NumberColumn("Over-usage (%)", format="%.1f"),
            "cost": st.column_config.NumberColumn(f"Cost ({CURRENCY})", format="%.0f"),
            "status_label": "Status",
        },
    )
    section("Download")
    d1, d2 = st.columns(2)
    stamp = date.today().strftime("%Y%m%d")
    with d1:
        try:
            st.download_button("⬇️ Excel report (.xlsx)", data=build_excel(progress),
                               file_name=f"sitestock_report_{stamp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except ImportError:
            st.error("Excel export needs the openpyxl package: run  py -m pip install openpyxl")
    with d2:
        try:
            st.download_button("⬇️ PDF report (.pdf)", data=build_pdf(progress),
                               file_name=f"sitestock_report_{stamp}.pdf", mime="application/pdf")
        except ImportError:
            st.error("PDF export needs the fpdf2 package: run  py -m pip install fpdf2")


def page_users():
    hero("Users", "Create accounts and assign roles", "👥")
    note("blue", "<b>Admin</b>: everything. <b>Storekeeper</b>: log deliveries/usage, edit records, view dashboard. <b>Manager</b>: dashboard and reports only.")
    users = query("SELECT id, username, full_name, role FROM users ORDER BY username")
    st.dataframe(users, hide_index=True, column_config={"id": None, "username": "Username", "full_name": "Full name", "role": "Role"})

    with st.form("add_user", clear_on_submit=True):
        c1, c2 = st.columns(2)
        username = c1.text_input("Username")
        full_name = c2.text_input("Full name")
        c3, c4 = st.columns(2)
        role = c3.selectbox("Role", ["Storekeeper", "Manager", "Admin"])
        password = c4.text_input("Password", type="password")
        if st.form_submit_button("➕ Create user"):
            if not username.strip() or not full_name.strip() or len(password) < 4:
                st.error("Enter a username, a full name and a password of at least 4 characters.")
            else:
                try:
                    run("INSERT INTO users (username, full_name, role, password_hash) VALUES (?,?,?,?)",
                        (username.strip().lower(), full_name.strip(), role, hash_password(password)))
                    st.success(f"User {username.strip().lower()} created.")
                except sqlite3.IntegrityError:
                    st.error("That username already exists.")

    me = st.session_state.user["username"]
    others = [u for u in users["username"] if u != me]
    if others:
        section("Remove a user")
        target = st.selectbox("User", others)
        if st.button("🗑️ Delete user"):
            run("DELETE FROM users WHERE username = ?", (target,))
            st.success(f"Deleted {target}.")


def page_data():
    hero("Sample data & reset", "Try the app with example data, or wipe it for a new project", "🧪")
    st.write("Load a ready-made example: 4 materials with 5 weeks of deliveries and usage, relative to today.")
    if st.button("Load sample data"):
        if get_materials().empty:
            load_sample_data()
            st.success("Sample data loaded. Open the Dashboard.")
        else:
            st.warning("Materials already exist. Reset first to avoid duplicates.")
    st.divider()
    st.write("Delete all materials, deliveries and usage (user accounts are kept).")
    confirm = st.checkbox("I understand this deletes all project data")
    if st.button("Reset all project data", disabled=not confirm):
        reset_all_data()
        st.success("All project data deleted.")


# --------------------------------------------------------------------------
# Login and navigation
# --------------------------------------------------------------------------
def login_page():
    st.markdown("<div style='height:6vh'></div>", unsafe_allow_html=True)
    left, right = st.columns([1.15, 1])
    with left:
        st.markdown(
            '<div class="login-panel"><div style="display:flex;align-items:center;gap:12px;margin-bottom:14px">'
            + BRAND_SVG + '<h2>SiteStock</h2></div>'
            '<p>Know exactly what is on site, what has been used and what you will need next.</p>'
            '<ul><li>📦 Live stock for cement, rebar, sand, blocks and more</li>'
            '<li>🚨 Instant low-stock and over-usage alerts</li>'
            '<li>⏳ Stock-out and waste forecasts</li>'
            '<li>💰 Budget, cost and supplier comparison</li>'
            '<li>📄 One-click Excel and PDF reports</li></ul></div>',
            unsafe_allow_html=True,
        )
    with right:
        with st.form("login"):
            st.markdown("#### Sign in")
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign in")
        if submitted:
            rows = query("SELECT username, full_name, role, password_hash FROM users WHERE username = ?",
                         (username.strip().lower(),))
            if not rows.empty and verify_password(password, rows.iloc[0]["password_hash"]):
                st.session_state.user = {
                    "username": rows.iloc[0]["username"],
                    "full_name": rows.iloc[0]["full_name"],
                    "role": rows.iloc[0]["role"],
                }
                st.rerun()
            else:
                st.error("Wrong username or password.")
        with st.expander("Demo accounts"):
            st.markdown("- **admin** / admin123 (everything)\n- **store** / store123 (storekeeper)\n- **manager** / manager123 (reports and dashboard)")


PAGE_FUNCS = {
    "📊 Dashboard": page_dashboard,
    "🧱 Materials": page_materials,
    "🚚 Log delivery": page_delivery,
    "🏗️ Log usage": page_usage,
    "✏️ Edit records": page_records,
    "📄 Reports": page_reports,
    "👥 Users": page_users,
    "🧪 Sample data / reset": page_data,
}
PAGES_BY_ROLE = {
    "Admin": list(PAGE_FUNCS),
    "Storekeeper": ["📊 Dashboard", "🚚 Log delivery", "🏗️ Log usage", "✏️ Edit records"],
    "Manager": ["📊 Dashboard", "📄 Reports"],
}


def main():
    init_db()
    inject_css()
    if st.session_state.get("user") is None:
        login_page()
        return

    user = st.session_state.user
    with st.sidebar:
        st.markdown(f'<div class="brand">{BRAND_SVG}<div>Site<span>Stock</span></div></div>'
                    '<div class="brand-sub">Material management for construction sites</div>',
                    unsafe_allow_html=True)
        initials = "".join(w[0] for w in user["full_name"].split()[:2]).upper() or "U"
        st.markdown(f'<div class="usercard"><div class="avatar">{html.escape(initials)}</div>'
                    f'<div><b>{html.escape(user["full_name"])}</b><br>{html.escape(user["role"])}</div></div>',
                    unsafe_allow_html=True)
        choice = st.radio("Menu", PAGES_BY_ROLE.get(user["role"], ["📊 Dashboard"]), label_visibility="collapsed")
        st.markdown("---")
        st.slider("Project progress (%)", 0, 100, DEFAULT_PROGRESS, key="progress",
                  help="How much of the project is complete. Drives expected usage, waste and forecasts.")
        if st.button("Sign out"):
            st.session_state.user = None
            st.rerun()
    PAGE_FUNCS[choice]()
    footer()


main()
