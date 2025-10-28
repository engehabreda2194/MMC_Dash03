"""
Arsal Facility Management – MMC Dashboard (single-file Streamlit)

Run locally:
	streamlit run "Maximo_DEV/MMC_Dash03.py"

Requirements (install in your venv):
	pip install streamlit pandas plotly requests openpyxl

Deployment notes (Streamlit Community Cloud):
- Set your secrets in the app settings (MMC_DATA_URL, MMC_EXCEL_SHEET, MMC_ARSAL_LOGO, MMC_CLIENT_LOGO).
- A fallback to environment variables is supported for local runs.
- Add runtime.txt with Python version (e.g., 3.11) for best compatibility.
"""

from __future__ import annotations

import io
import os
from datetime import datetime, timedelta
import base64
from typing import Optional, Tuple, Dict, Any, List
from textwrap import dedent
import warnings

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
# Note: components import removed as it's unused in this display-only version


# -------------------------
# Page config and theme
# -------------------------
st.set_page_config(
	page_title="Arsal FM – MMC Dashboard",
	page_icon="🟧",
	layout="wide",
	initial_sidebar_state="collapsed",
)

PRIMARY = "#FF6B00"      # Bright Orange
SECONDARY = "#F5F5F5"    # Light Gray
ACCENT = "#333333"       # Charcoal Gray
TEXT = "#000000"         # Black

# Brand assets (direct view links)
arsal_logo_url = "https://drive.google.com/uc?export=view&id=1lAmXKLuyFH0IS840z7PoecsU_eYSsloh"
client_logo_url = "https://drive.google.com/uc?export=view&id=1XHP8Str5Iwt9YinIgfisG6xTnOM222Py"
def get_secret_env(key: str, default: Optional[str] = None) -> Optional[str]:
	"""Read a config value from st.secrets first, then environment variables.

	Returns default if the key is not present or empty in both places.
	"""
	try:
		# st.secrets behaves like a dict; guard for when secrets not configured
		if hasattr(st, "secrets") and key in st.secrets and str(st.secrets.get(key)).strip():
			return str(st.secrets.get(key))
	except Exception:
		pass
	val = os.environ.get(key)
	return val if (val is not None and str(val).strip()) else default



def _gdrive_to_direct(u: str) -> str:
	"""Convert Google Drive file link to direct view URL if needed."""
	if not u:
		return u
	if "drive.google.com/file/d/" in u:
		try:
			file_id = u.split("/file/d/")[1].split("/")[0]
			return f"https://drive.google.com/uc?export=view&id={file_id}"
		except Exception:
			return u
	return u


def _file_to_data_uri(path: str) -> Optional[str]:
	try:
		with open(path, "rb") as f:
			b = f.read()
		mime = "image/png"
		ext = os.path.splitext(path)[1].lower()
		if ext in [".jpg", ".jpeg"]:
			mime = "image/jpeg"
		elif ext == ".svg":
			mime = "image/svg+xml"
		elif ext == ".webp":
			mime = "image/webp"
		return f"data:{mime};base64,{base64.b64encode(b).decode()}"
	except Exception:
		return None


def _resolve_logo_src(default_url: str, env_key: str, local_names: list[str], keywords: Optional[list[str]] = None) -> str:
	"""Prefer local asset (as data URI), then env URL, then default URL.

	Additionally supports fuzzy search by keywords in the assets folder (case-insensitive),
	so names like "MMC Logo.jpg" or "ARSAL Logo.png" are detected automatically.
	"""
	# 1) Local assets under Maximo_DEV/assets (exact names first)
	base_dir = os.path.dirname(__file__)
	assets_dir = os.path.join(base_dir, "assets")
	for name in local_names:
		p = os.path.join(assets_dir, name)
		if os.path.exists(p):
			data_uri = _file_to_data_uri(p)
			if data_uri:
				return data_uri
	# 1b) Fuzzy search by keywords
	if os.path.isdir(assets_dir) and keywords:
		try:
			for f in os.listdir(assets_dir):
				fl = f.lower()
				if any(kw.lower() in fl for kw in keywords) and os.path.splitext(f)[1].lower() in [".png", ".jpg", ".jpeg", ".svg", ".webp"]:
					p = os.path.join(assets_dir, f)
					data_uri = _file_to_data_uri(p)
					if data_uri:
						return data_uri
		except Exception:
			pass
	# 2) Secrets/Environment override
	env_url = get_secret_env(env_key)
	if env_url:
		return _gdrive_to_direct(env_url)
	# 3) Default
	return _gdrive_to_direct(default_url)


def _parse_dates_safely(series: pd.Series) -> pd.Series:
	"""Parse a date-like Series robustly without noisy warnings.

	Tries a set of common formats first (vectorized and fast), picks the one with
	the most non-NaT matches. Falls back to generic parser with warnings suppressed.
	"""
	if series is None:
		return series
	# Ensure string input (avoid mixed types) and strip whitespace
	s = series.astype(str).str.strip()
	# If already datetime-like, return as-is
	if getattr(series, 'dtype', None) is not None and str(series.dtype).startswith('datetime'):
		return series

	formats = [
		"%Y-%m-%d",
		"%Y/%m/%d",
		"%d/%m/%Y",
		"%m/%d/%Y",
		"%d-%m-%Y",
		"%m-%d-%Y",
		"%Y-%m-%d %H:%M:%S",
		"%d/%m/%Y %H:%M",
		"%m/%d/%Y %H:%M",
	]
	bests: Tuple[Optional[pd.Series], int] = (None, -1)
	for fmt in formats:
		parsed = pd.to_datetime(s, format=fmt, errors="coerce")
		non_na = int(parsed.notna().sum())
		if non_na > bests[1]:
			bests = (parsed, non_na)
			if non_na == len(s):
				break

	if bests[0] is not None and bests[1] > 0:
		return bests[0]

	# Fallback: generic parsing with warnings suppressed; prefer dayfirst to handle DD/MM/YYYY
	with warnings.catch_warnings():
		warnings.simplefilter("ignore", category=UserWarning)
		return pd.to_datetime(s, errors="coerce", dayfirst=True)


def _svg_placeholder(text: str, bg: str = "#FF6B00", fg: str = "#FFFFFF") -> str:
	"""Create a simple SVG badge as data URI used as a fallback logo."""
	svg = f"""
	<svg xmlns='http://www.w3.org/2000/svg' width='128' height='64'>
		<rect width='100%' height='100%' fill='{bg}' rx='8' ry='8'/>
		<text x='50%' y='50%' dominant-baseline='middle' text-anchor='middle'
				font-family='Segoe UI, Poppins, Arial' font-size='28' font-weight='700' fill='{fg}'>
			{text}
		</text>
	</svg>
	""".strip()
	uri = "data:image/svg+xml;base64," + base64.b64encode(svg.encode("utf-8")).decode("ascii")
	return uri

# Qualitative, diverse color palette for charts (not only orange)
COLOR_SEQ = (
	px.colors.qualitative.Vivid
	+ px.colors.qualitative.Safe
	+ px.colors.qualitative.Set2
	+ px.colors.qualitative.Pastel
)
ORANGE_SEQ = ["#FF6B00", "#FF8C33", "#FFB366", "#FFD1A3"]

# Fixed OneDrive/SharePoint CSV direct link (update this to your public CSV download URL)
# Example patterns:
# - SharePoint: https://<tenant>.sharepoint.com/.../download.aspx?share=...
# - OneDrive: https://api.onedrive.com/v1.0/shares/.../root/content
# - Google Sheets CSV export: https://docs.google.com/spreadsheets/d/<ID>/export?format=csv
FIXED_DATA_URL = get_secret_env(
	"MMC_DATA_URL",
	# Default to your Google Sheets link; loader will convert it to CSV export automatically
	"https://docs.google.com/spreadsheets/d/1T6dndJHd33ZW3i4e9LIOGl1BPkxDsYme/edit?usp=sharing&ouid=115201289744778991707&rtpof=true&sd=true",
)


def inject_brand_css() -> None:
	st.markdown(
		f"""
		<style>
		:root {{
			--accent: {PRIMARY};
			--light-bg: #fffaf5;
			--shadow: 0 8px 18px rgba(0,0,0,0.08);
			--card-radius: 12px;
		}}
		body {{
			background: linear-gradient(135deg, var(--light-bg), #ffffff);
			font-family: 'Poppins', 'Segoe UI', sans-serif;
			zoom: 1.05; /* Slight scale-up for TV readability */
		}}
		.brand-header {{
			display: grid;
			/* Fix left/right columns to keep logo + title perfectly aligned across tabs */
			grid-template-columns: 160px 1fr 160px;
			align-items: center;
			/* Fix header block height so it doesn't jump between tabs */
			min-height: 110px; /* Taller to fit subtitle line */
			padding: 0.6rem 1rem; background: rgba(255,255,255,0.9);
			border-bottom: 2px solid var(--accent);
			box-shadow: 0 2px 10px rgba(0,0,0,0.05);
			border-radius: 12px; backdrop-filter: blur(8px);
			margin-bottom: 8px;
		}}
		/* Fix logo box size so layout stays stable when switching tabs */
		.logo-img {{ height:64px; width:64px; border-radius:10px; object-fit: contain; }}
		.brand-left {{ text-align: left; }}
		.brand-center {{ text-align: center; }}
		.brand-right {{ text-align: right; }}
		.brand-subtitle {{ margin: 4px 0 0 0; color:#555; font-size:14px; font-weight:600; }}
		/* Keep title on one line and consistent size across tabs */
		.brand-center h3 {{
			font-size: 38px;
			line-height: 1.2;
			margin: 0;
			white-space: nowrap;
		}}
		.kpi-card {{
			background: rgba(255,255,255,0.9);
			border: 1px solid rgba(255,255,255,0.5);
			border-radius: 20px;
			box-shadow: 0 2px 10px rgba(0,0,0,0.08);
			padding: 18px;
		}}
		.kpi-title {{ color:#666; font-size:18px; margin:0; }}
		.kpi-value {{ color:#111; font-weight:900; font-size:44px; margin:0; }}
		.kpi-emoji {{ font-size:24px; margin-right:8px; }}
		h1, h2, h3 {{ color: var(--accent); }}
		.stButton>button {{
			background: linear-gradient(135deg, var(--accent), #ff933f);
			color: white; border: none; border-radius: 10px; padding: 10px 18px; font-weight:700;
			box-shadow: 0 6px 14px rgba(255,107,0,0.22);
		}}
		.stButton>button:hover {{ filter: brightness(1.05); }}
		.card {{ background:white; border:1px solid rgba(0,0,0,0.08); border-radius: var(--card-radius); box-shadow: var(--shadow); padding: 10px 12px; }}
		.muted {{ color:#666; font-size:14px; }}
		header {{ visibility: hidden; }}
		/* Remove extra padding for TV style */
		.block-container {{ padding-top: 0.8rem; padding-bottom: 0.8rem; }}
		</style>
		""",
		unsafe_allow_html=True,
	)


## Removed sidebar hints and inputs to make a pure display dashboard


def brand_header() -> None:
	"""Top header with Aarsal logo (left), centered title, and Client logo (right)."""
	# Current date & time (local)
	now_str = datetime.now().strftime("%a, %d %b %Y – %H:%M")
	left_src = _resolve_logo_src(
		default_url=arsal_logo_url,
		env_key="MMC_ARSAL_LOGO",
		local_names=["arsal_logo.png", "arsal_logo.jpg", "arsal_logo.jpeg", "arsal_logo.svg", "arsal_logo.webp"],
		keywords=["arsal", "aarsal", "arsal logo"],
	)
	right_src = _resolve_logo_src(
		default_url=client_logo_url,
		env_key="MMC_CLIENT_LOGO",
		local_names=["client_logo.png", "client_logo.jpg", "client_logo.jpeg", "client_logo.svg", "client_logo.webp"],
		keywords=["client", "mmc", "almajdouie", "mmc logo"],
	)
	left_fallback = _svg_placeholder("Arsal", bg=PRIMARY)
	right_fallback = _svg_placeholder("Client", bg=PRIMARY)
	st.markdown(
		f"""
		<div class='brand-header'>
			<div class='brand-left'>
				<img src="{left_src}" class="logo-img" onerror="this.onerror=null; this.src='{left_fallback}';" />
			</div>
			<div class='brand-center'>
				<h3 style='margin:0;'>MMC Project Dashboard</h3>
				<div class='brand-subtitle'>{now_str}</div>
			</div>
			<div class='brand-right'>
				<img src="{right_src}" class="logo-img" onerror="this.onerror=null; this.src='{right_fallback}';" />
			</div>
		</div>
		""",
		unsafe_allow_html=True,
	)


# -------------------------
# Data loading and caching
# -------------------------

@st.cache_data(ttl=60, show_spinner=False)
def load_data(url: str) -> Tuple[pd.DataFrame, datetime]:
	"""Download CSV from fixed link (SharePoint/OneDrive/Google Sheets CSV export)."""
	if not url:
		raise ValueError("No data URL configured.")
	headers = {"User-Agent": "Mozilla/5.0"}

	def ensure_csv(u: str) -> str:
		# Google Sheets share -> force CSV
		if "docs.google.com/spreadsheets" in u and "/export?" not in u:
			try:
				sheet_id = u.split("/d/")[1].split("/")[0]
				return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
			except Exception:
				return u
		# Encourage direct content for SharePoint/OneDrive links
		if ("download.aspx" in u) or ("download=1" in u) or ("/content" in u):
			return u
		return f"{u}&download=1" if ("?" in u) else f"{u}?download=1"

	eff_url = ensure_csv(url)
	try:
		# Use a shorter timeout to avoid long hangs on TVs
		resp = requests.get(eff_url, headers=headers, timeout=12, allow_redirects=True)
		resp.raise_for_status()
	except requests.RequestException as e:
		raise RuntimeError(f"HTTP error while fetching CSV. Ensure the link is public/direct. Details: {e}") from e

	# Try CSV first; if it fails, try Excel as fallback
	content = resp.content
	with io.BytesIO(content) as buf:
		try:
			df = pd.read_csv(buf, on_bad_lines="skip")
		except Exception:
			buf.seek(0)
			try:
				# Read Excel completely. If MMC_EXCEL_SHEET is set, select that sheet (case-insensitive),
				# otherwise read and concatenate all sheets to ensure we don't miss data.
				xls = pd.ExcelFile(buf, engine="openpyxl")
				prefer_sheet = get_secret_env("MMC_EXCEL_SHEET")
				if prefer_sheet:
					# case-insensitive match
					match = None
					for sn in xls.sheet_names:
						if sn.strip().lower() == prefer_sheet.strip().lower():
							match = sn
							break
					sheet_to_read = match or xls.sheet_names[0]
					df = xls.parse(sheet_name=sheet_to_read, engine="openpyxl")
				else:
					frames = []
					for sn in xls.sheet_names:
						frames.append(xls.parse(sheet_name=sn, engine="openpyxl"))
					df = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
			except Exception as e:
				raise RuntimeError("Unable to parse data as CSV or Excel.") from e

	# Normalize columns
	df.columns = [str(c).strip() for c in df.columns]
	# Parse likely date columns (best-effort)
	for col in [
		"Reported Date",
		"Create Date",
		"Created Date",
		"Target Date",
		"Completion Date",
		"Actual Finish",
		"Actual Start",
		"Status Date",
		"Last Modified",
		"Change Date",
	]:
		if col in df.columns:
			df[col] = _parse_dates_safely(df[col])

	return df, datetime.now()


def clear_cache_and_rerun():
	st.cache_data.clear()
	try:
		# Newer Streamlit versions
		st.rerun()
	except Exception:
		# Backward compatibility
		st.experimental_rerun()


def make_sample_data() -> pd.DataFrame:
	"""Fallback sample dataset for TV demo when data URL is unavailable."""
	now = datetime.now()
	data = {
		"Status": [
			"WPLAN", "PLANCOMP", "WQAPPRC", "QAPPRC", "WSCH", "SCHEDCOMP",
			"COMP", "COMP", "QREJECTC", "WPLAN", "PLANCOMP", "CAN"
		],
		"Work Type": [
			"PM", "CM", "ADW", "PM", "CM", "ADW",
			"PM", "CM", "ADW", "PM", "CM", "PM"
		],
		"Work Category": [
			"Routine", "Routine", "Emergency", "Routine", "Emergency", "Routine",
			"Routine", "Emergency", "Routine", "Contract", "Routine", "Routine"
		],
		"Reported Date": [now - timedelta(days=d) for d in [30,28,25,22,20,18,15,12,10,8,5,2]],
		"Target Date": [now - timedelta(days=d) for d in [25,26,22,20,18,16,12,10,8,6,2,1]],
		"Completion Date": [
			pd.NaT, pd.NaT, pd.NaT, pd.NaT, pd.NaT, pd.NaT,
			now - timedelta(days=13), now - timedelta(days=9), pd.NaT, pd.NaT, pd.NaT, pd.NaT
		],
		"Site/Location": [
			"Site A", "Site A", "Site B", "Site B", "Site C", "Site C",
			"Site A", "Site B", "Site C", "Site A", "Site B", "Site C"
		],
		"Planner Name": [
			"Planner 1", "Planner 2", "Planner 1", "Planner 3", "Planner 2", "Planner 3",
			"Planner 1", "Planner 2", "Planner 1", "Planner 3", "Planner 2", "Planner 3"
		],
	}
	df = pd.DataFrame(data)
	return df


def parse_excel_bytes(content: bytes) -> pd.DataFrame:
	with io.BytesIO(content) as buf:
		# Try openpyxl
		try:
			xls = pd.ExcelFile(buf, engine="openpyxl")
			prefer_sheet = get_secret_env("MMC_EXCEL_SHEET")
			if prefer_sheet:
				match = None
				for sn in xls.sheet_names:
					if sn.strip().lower() == prefer_sheet.strip().lower():
						match = sn
						break
				sheet_to_read = match or xls.sheet_names[0]
				df = xls.parse(sheet_name=sheet_to_read, engine="openpyxl")
			else:
				frames = [xls.parse(sheet_name=sn, engine="openpyxl") for sn in xls.sheet_names]
				df = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
		except Exception:
			# Fallback: reposition and try CSV
			buf.seek(0)
			try:
				df = pd.read_csv(buf)
			except Exception as e:
				raise RuntimeError("Uploaded file format not recognized as Excel/CSV. If it's .xls, install 'xlrd' or re-save as .xlsx.") from e
	# Normalize columns
	df.columns = [str(c).strip() for c in df.columns]
	# Parse likely date columns (best-effort)
	for col in [
		"Reported Date",
		"Create Date",
		"Created Date",
		"Target Date",
		"Completion Date",
		"Actual Finish",
		"Actual Start",
		"Status Date",
		"Last Modified",
		"Change Date",
	]:
		if col in df.columns:
			df[col] = _parse_dates_safely(df[col])
	return df


# -------------------------
# KPI helpers
# -------------------------

def compute_kpis(df: pd.DataFrame) -> Dict[str, Optional[float]]:
	"""Compute KPIs using the specified status mapping and rules.

	Status mapping:
	- WPLAN: Open – not yet acted on
	- PLANCOMP: With procurement
	- WQAPPRC: Waiting client quotation approval
	- QREJECTC: Rejected by client
	- QAPPRC: Client approved (waiting execution)
	- WSCH: Materials delivered / service executed (waiting completion confirmation)
	- SCHEDCOMP: Work executed, pending invoices/docs
	- COMP: Fully completed
	- CAN: Cancelled – exclude from all stats
	"""
	# Normalize status column
	status_col = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	if status_col is not None:
		status_series = df[status_col].astype(str).str.upper().str.strip()
		# Common synonyms/variants
		status_series = (
			status_series
			.replace({
				"CANCELLED": "CAN",
				"CANCELED": "CAN",
				"CANCEL": "CAN",
				"CLOSED": "CLOSE",
			})
		)
	else:
		status_series = pd.Series([None] * len(df))

	# Exclude cancelled and 'Contract' category from all KPIs, and exclude Location == 'MMC-ARSAL'
	not_cancelled = ~(status_series == "CAN")
	cat_col = next((c for c in ["Work Category", "Category"] if c in df.columns), None)
	if cat_col is not None:
		cat_series = df[cat_col].astype(str).str.upper().str.strip()
		not_contract = ~(cat_series == "CONTRACT")
	else:
		not_contract = pd.Series([True] * len(df))

	loc_col = next((c for c in ["Location", "Site/Location", "Site"] if c in df.columns), None)
	if loc_col is not None:
		loc_series = df[loc_col].astype(str).str.upper().str.strip()
		not_mmc_arsal = ~(loc_series == "MMC-ARSAL")
	else:
		not_mmc_arsal = pd.Series([True] * len(df))

	mask = not_cancelled & not_contract & not_mmc_arsal
	dfx = df[mask].copy()
	status_series = status_series[mask]

	total = len(dfx)

	OPEN_CODES = {"WPLAN", "PLANCOMP", "WQAPPRC", "QAPPRC", "WSCH", "SCHEDCOMP"}
	COMPLETED_CODE = "COMP"
	CLOSED_CODE = "CLOSE"
	REJECTED_CODE = "QREJECTC"

	open_mask = status_series.isin(OPEN_CODES)
	completed_mask = status_series.eq(COMPLETED_CODE)
	closed_mask = status_series.eq(CLOSED_CODE)
	rejected_mask = status_series.eq(REJECTED_CODE)

	# Work type distribution (PM, CM, ADW)
	wt_col = next((c for c in ["Work Type", "Type", "WORKTYPE", "worktype"] if c in dfx.columns), None)
	if wt_col:
		wt_upper = dfx[wt_col].astype(str).str.upper()
		pm_mask = wt_upper.str.contains("PM", na=False)
		cm_mask = wt_upper.str.contains("CM", na=False)
		adw_mask = wt_upper.str.contains("ADW", na=False)
	else:
		pm_mask = pd.Series([False] * total)
		cm_mask = pd.Series([False] * total)
		adw_mask = pd.Series([False] * total)
	# Work type percentages and CM/PM ratio
	pm_n = int(pm_mask.sum())
	cm_n = int(cm_mask.sum())
	adw_n = int(adw_mask.sum())
	wt_total = max(1, pm_n + cm_n + adw_n)
	pm_pct = 100 * pm_n / wt_total
	cm_pct = 100 * cm_n / wt_total
	adw_pct = 100 * adw_n / wt_total
	cm_pm_ratio = 100 * (cm_n / max(1, pm_n))
	# PM/CM only share (per request: PM% over PM+CM, CM% = 100 - PM%)
	pmcm_total = max(1, pm_n + cm_n)
	pm_pct_pc = 100 * pm_n / pmcm_total
	cm_pct_pc = 100 - pm_pct_pc

	# Avg completion time
	if "Reported Date" in dfx.columns:
		reported = dfx["Reported Date"]
	elif "Created Date" in dfx.columns:
		reported = dfx["Created Date"]
	else:
		reported = pd.Series([pd.NaT] * total)

	completion = dfx.get("Completion Date", dfx.get("Actual Finish", pd.Series([pd.NaT] * total)))
	valid = reported.notna() & completion.notna()
	avg_days = (completion[valid] - reported[valid]).dt.total_seconds().mean() / 86400 if valid.any() else None

	# % Closed on time (if Target Date exists)
	target = dfx.get("Target Date")
	if target is not None and completion is not None:
		on_time_mask = completion.notna() & target.notna() & (completion <= target)
		# Denominator: completed with valid target+completion
		valid_comp = completion.notna() & target.notna()
		pct_on_time = 100 * (on_time_mask.sum() / max(1, valid_comp.sum()))
	else:
		pct_on_time = None

	# Completion Rate (completed over total, excluding cancelled already)
	completion_rate = 100 * (completed_mask.sum() / max(1, total))
	# Backlog Rate (open over total)
	backlog_rate = 100 * (open_mask.sum() / max(1, total))
	# Rework Rate (QREJECTC over total)
	rework_rate = 100 * (rejected_mask.sum() / max(1, total))

	return {
		"total": total,
		"open": int(open_mask.sum()),
		"completed": int(completed_mask.sum()),
		"closed": int(closed_mask.sum()),
		"rejected": int(rejected_mask.sum()),
		"pm": pm_n,
		"cm": cm_n,
		"adw": adw_n,
		"pm_pct": pm_pct,
		"cm_pct": cm_pct,
		"adw_pct": adw_pct,
		"cm_pm_ratio": cm_pm_ratio,
		"pm_pct_pc": pm_pct_pc,
		"cm_pct_pc": cm_pct_pc,
		"avg_days": avg_days,
		"pct_on_time": pct_on_time,
		"completion_rate": completion_rate,
		"backlog_rate": backlog_rate,
		"rework_rate": rework_rate,
	}


def kpi_card(label: str, value: str, emoji: str) -> None:
	st.markdown(
		f"""
		<div class='kpi-card'>
		  <p class='kpi-title'><span class='kpi-emoji'>{emoji}</span>{label}</p>
		  <p class='kpi-value'>{value}</p>
		</div>
		""",
		unsafe_allow_html=True,
	)


# -------------------------
# Insights
# -------------------------

def compute_insights(df: pd.DataFrame) -> Dict[str, Optional[float]]:
	"""Aggregate additional analytics for service provider performance."""
	out: Dict[str, Optional[float]] = {}
	total = len(df)

	status_col = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	done_mask = pd.Series([False] * total)
	if status_col is not None:
		su = df[status_col].astype(str).str.upper()
		done_mask = su.isin(["COMP", "CLOSE"])  # completed or closed
	done_n = int(done_mask.sum())
	out["completion_rate"] = 100 * (done_n / max(1, total))

	# On-time among those with both dates
	target = pd.to_datetime(df.get("Target Date"), errors="coerce") if "Target Date" in df.columns else None
	completion = pd.to_datetime(df.get("Completion Date"), errors="coerce") if "Completion Date" in df.columns else None
	if target is not None and completion is not None:
		valid = completion.notna() & target.notna()
		on_time = (completion <= target) & valid
		out["on_time_rate"] = 100 * (on_time.sum() / max(1, valid.sum()))
	else:
		out["on_time_rate"] = None

	# Month-over-month change in completed
	if completion is not None:
		months = completion.dt.to_period("M").astype(str)
		ser = pd.DataFrame({"month": months[done_mask]})
		if not ser.empty:
			cnt = ser.value_counts().rename_axis(["month"]).reset_index(name="count").sort_values("month")
			if len(cnt) >= 2:
				last, prev = cnt.iloc[-1]["count"], cnt.iloc[-2]["count"]
				out["mom_change"] = 100 * ((last - prev) / max(1, prev))
				out["last_month"] = str(cnt.iloc[-1]["month"])
			elif len(cnt) == 1:
				out["mom_change"] = None
				out["last_month"] = str(cnt.iloc[-1]["month"])
			else:
				out["mom_change"] = None
				out["last_month"] = None
		else:
			out["mom_change"] = None
			out["last_month"] = None
	else:
		out["mom_change"] = None
		out["last_month"] = None

	# Top category and planner
	cat_col = next((c for c in ["Work Category", "Category"] if c in df.columns), None)
	plan_col = next((c for c in ["Planner Name", "Planner"] if c in df.columns), None)
	try:
		out["top_category"] = (
			df[cat_col].astype(str).value_counts().idxmax() if cat_col else None
		)
	except Exception:
		out["top_category"] = None
	try:
		out["top_planner"] = (
			df[plan_col].astype(str).value_counts().idxmax() if plan_col else None
		)
	except Exception:
		out["top_planner"] = None

	out["total_wos"] = total
	out["done_wos"] = done_n

	# Additional KPIs for provider performance
	# Rework proxy using QREJECTC presence
	status_col2 = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	rework_n = 0
	if status_col2 is not None:
		su2 = df[status_col2].astype(str).str.upper()
		rework_n = int((su2 == "QREJECTC").sum())
	out["rework_rate"] = 100 * (rework_n / max(1, total))

	# Backlog rate (open codes)
	OPEN_CODES = {"WPLAN", "PLANCOMP", "WQAPPRC", "QAPPRC", "WSCH", "SCHEDCOMP"}
	open_n = 0
	if status_col2 is not None:
		su3 = df[status_col2].astype(str).str.upper()
		open_n = int(su3.isin(OPEN_CODES).sum())
	out["backlog_rate"] = 100 * (open_n / max(1, total))

	# Overdue open rate (Target passed and not completed)
	target = pd.to_datetime(df.get("Target Date"), errors="coerce") if "Target Date" in df.columns else None
	if target is not None and status_col2 is not None:
		su4 = df[status_col2].astype(str).str.upper()
		is_open = su4.isin(OPEN_CODES)
		overdue_open = is_open & target.notna() & (target < pd.Timestamp.now())
		out["overdue_open_rate"] = 100 * (overdue_open.sum() / max(1, is_open.sum())) if is_open.any() else None
	else:
		out["overdue_open_rate"] = None

	# Average backlog age (days) for open items
	reported = pd.to_datetime(df.get("Reported Date"), errors="coerce") if "Reported Date" in df.columns else None
	if reported is not None and status_col2 is not None:
		su5 = df[status_col2].astype(str).str.upper()
		is_open = su5.isin(OPEN_CODES)
		ages = (pd.Timestamp.now() - reported[is_open]).dt.total_seconds() / 86400
		out["avg_backlog_age"] = float(ages.mean()) if is_open.any() else None
	else:
		out["avg_backlog_age"] = None

	# SLA breach rate among completed
	if target is not None and completion is not None:
		valid_comp = completion.notna() & target.notna()
		late = valid_comp & (completion > target)
		out["sla_breach_rate"] = 100 * (late.sum() / max(1, valid_comp.sum())) if valid_comp.any() else None
	else:
		out["sla_breach_rate"] = None

	# PM compliance (PM Completed on time / PM Completed with target)
	wt_col = next((c for c in ["Work Type", "Type", "WORKTYPE", "worktype"] if c in df.columns), None)
	if wt_col is not None and target is not None and completion is not None:
		wt_u = df[wt_col].astype(str).str.upper()
		is_pm = wt_u.str.contains("PM", na=False)
		valid = is_pm & completion.notna() & target.notna()
		ontime = valid & (completion <= target)
		out["pm_compliance"] = 100 * (ontime.sum() / max(1, valid.sum())) if valid.any() else None
	else:
		out["pm_compliance"] = None

	# First-Time Fix Rate (approx) = 100 - rework_rate
	out["ftfr"] = 100 - (out["rework_rate"] or 0)
	return out


def insight_card(label: str, value: str, emoji: str) -> None:
	kpi_card(label, value, emoji)

# -------------------------
# Sidebar and filters
# -------------------------

## Removed: build_sidebar (no inputs in display mode)


## Removed: apply_filters (no filters in display mode)


# -------------------------
# Charts
# -------------------------

def _apply_chart_theme(fig, title: str):
	fig.update_layout(
		title={"text": title, "x": 0.02, "xanchor": "left"},
		margin=dict(l=10, r=10, t=50, b=10),
		font=dict(size=18, color="#222"),
		legend=dict(font=dict(size=16)),
		paper_bgcolor="#f5f6fa",
		plot_bgcolor="#eef1f5",
	)
	fig.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.05)")
	fig.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.05)")
	return fig


def make_charts(df: pd.DataFrame) -> None:
	"""Render all charts full-width, stacked vertically for maximum readability."""
	# Prefer Location over Site when both are present
	c_location = next((c for c in ["Location", "Site/Location", "Site"] if c in df.columns), None)
	c_status = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	c_planner = next((c for c in ["Planner Name", "Planner"] if c in df.columns), None)
	c_date = next((c for c in ["Reported Date", "Created Date", "Create Date"] if c in df.columns), None)
	c_category = next((c for c in ["Work Category", "Category"] if c in df.columns), None)
	c_type = next((c for c in ["Work Type", "Type", "WORKTYPE", "worktype"] if c in df.columns), None)

	# 1) Full-width: Work Orders by Location (aggregated) with count labels and small tick text
	if c_location:
		loc_counts = (
			df.groupby(c_location).size().reset_index(name="count").sort_values("count", ascending=False)
		)
		fig_loc = px.bar(
			loc_counts,
			x=c_location,
			y="count",
			title="Work Orders by Location",
			text="count",
			color_discrete_sequence=ORANGE_SEQ,
			labels={c_location: "Location", "count": "WO count"},
		)
		fig_loc.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
		fig_loc.update_traces(hovertemplate="Location=%{x}<br>WO count=%{y}<extra></extra>")
		fig_loc.update_layout(bargap=0.15)
		fig_loc.update_xaxes(tickangle=-90, tickfont=dict(size=9), categoryorder="total descending")
		fig_loc = _apply_chart_theme(fig_loc, "Work Orders by Location")
		fig_loc.update_layout(height=600)
		st.plotly_chart(fig_loc, use_container_width=True)

		# 1b) Full-width: CM vs PM per Location – grouped bars (all records in sheet)
		if c_type:
			df_cmpm = df.copy()

			# Keep only CM and PM
			wt_upper = df_cmpm[c_type].astype(str).str.upper()
			# Include any values that equal or contain CM/PM to be tolerant of variations
			is_cm = wt_upper.str.contains("CM", na=False)
			is_pm = wt_upper.str.contains("PM", na=False)
			df_cmpm = df_cmpm[is_cm | is_pm].copy()
			df_cmpm["_wt"] = wt_upper.where(is_pm, "CM").where(is_cm, "PM")

			if not df_cmpm.empty:
				counts = df_cmpm.groupby([c_location, "_wt"]).size().reset_index(name="count")
				# Order locations by total count desc
				order = counts.groupby(c_location)["count"].sum().sort_values(ascending=False).index.tolist()
				fig_cmpm = px.bar(
					counts,
					x=c_location,
					y="count",
					color="_wt",
					barmode="group",
					title="CM vs PM per Location",
					labels={"count": "WO count", "_wt": "Work Type", c_location: "Location"},
					color_discrete_map={"CM": "#FF6B00", "PM": "#2E86AB"},
					text="count",
				)
				fig_cmpm.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
				fig_cmpm.update_traces(hovertemplate="Location=%{x}<br>Work Type=%{legendgroup}<br>WO count=%{y}<extra></extra>")
				fig_cmpm.update_xaxes(tickangle=-90, tickfont=dict(size=9), categoryorder="array", categoryarray=order)
				fig_cmpm.update_layout(legend_title_text="Work Type")
				fig_cmpm = _apply_chart_theme(fig_cmpm, "CM vs PM per Location")
				fig_cmpm.update_layout(height=650)
				st.plotly_chart(fig_cmpm, use_container_width=True)
			else:
				st.info("No CM/PM completed records available for Location breakdown.")
		else:
			st.info("Work Type column not found for CM/PM per Location chart.")
	else:
		st.info("Location column not found.")

	# 8) Full-width: Work Orders by Priority (with small-location labels on bars)
	c_priority = next((c for c in ["Priority", "PRIORITY", "priority", "PRIORITY CODE", "PRIORITYCODE"] if c in df.columns), None)
	if c_priority:
		prio_str = df[c_priority].astype(str).str.strip()
		dfp = df.copy()
		dfp["_prio"] = prio_str
		# Build location labels (small) per priority if Location exists
		if c_location:
			def _fmt_locs(series: pd.Series) -> str:
				vals = pd.unique(series.astype(str))
				vals = sorted([v for v in vals if v and v.lower() != "nan"])  # clean
				# Show ALL locations, compact with bullet separators
				return " • ".join(vals)
			loc_text = dfp.groupby("_prio")[c_location].apply(_fmt_locs).rename("locs_text")
			counts = dfp.groupby("_prio").size().reset_index(name="count").merge(loc_text, on="_prio", how="left")
		else:
			counts = dfp.groupby("_prio").size().reset_index(name="count")
			counts["locs_text"] = ""

		# Numeric sort if possible
		counts["prio_num"] = pd.to_numeric(counts["_prio"], errors="coerce")
		counts = counts.sort_values(["prio_num", "_prio"], ascending=[True, True])

		# Friendly labels next to each priority
		prio_name_map = {"1": "Urgent", "2": "High", "3": "Medium", "4": "Low"}
		counts["prio_name"] = counts["_prio"].map(prio_name_map).fillna("")
		counts["prio_label"] = counts.apply(lambda r: f"{r['_prio']} - {r['prio_name']}" if r["prio_name"] else str(r["_prio"]), axis=1)

		# Color map: Priority 1 in red
		unique_prios = counts["_prio"].astype(str).tolist()
		color_map = {p: px.colors.qualitative.Vivid[i % len(px.colors.qualitative.Vivid)] for i, p in enumerate(unique_prios)}
		color_map["1"] = "#d62728"

		fig_prio = px.bar(
			counts,
			x="_prio",
			y="count",
			color="_prio",
			text="locs_text",
			title="Work Orders by Priority",
			labels={"_prio": "Priority", "count": "WO count"},
			color_discrete_map=color_map,
			custom_data=["locs_text", "count"],
		)
		# Place full location list inside the bar, tiny font, vertical for space efficiency
		fig_prio.update_traces(textposition="inside", textangle=90, insidetextfont=dict(size=7, color="white"), cliponaxis=False)
		fig_prio.update_traces(hovertemplate="Priority=%{x}<br>WO count=%{y}<br>Locations=%{customdata[0]}<extra></extra>")
		fig_prio.update_xaxes(categoryorder="array", categoryarray=counts["_prio"].tolist())
		# Replace tick text with friendly labels (e.g., '1 - Urgent')
		fig_prio.update_xaxes(tickmode="array", tickvals=counts["_prio"].tolist(), ticktext=counts["prio_label"].tolist())
		fig_prio.update_layout(legend_title_text="Priority")
		fig_prio.update_layout(hovermode="closest")
		fig_prio = _apply_chart_theme(fig_prio, "Work Orders by Priority")
		fig_prio.update_layout(height=560)
		st.plotly_chart(fig_prio, use_container_width=True)
	else:
		st.info("Priority column not found.")

	# 2) Full-width: Work Orders by Status (bigger donut to show small slices)
	if c_status:
		# Exclude CAN from status pie if present
		s = df[c_status].astype(str).str.upper()
		s = s[s != "CAN"]
		fig = px.pie(s, names=s, title="Work Orders by Status", hole=0.35, color_discrete_sequence=COLOR_SEQ)
		# Show both count and percent for each slice
		fig.update_traces(
			textposition="outside",
			textinfo="label+value+percent",
			texttemplate="%{label}: %{value} (%{percent})",
			textfont_size=16,
			pull=0,
		)
		fig.update_layout(showlegend=True)
		fig = _apply_chart_theme(fig, "Work Orders by Status")
		# Extra top margin so outside labels aren't clipped at the top edge
		fig.update_layout(margin=dict(l=20, r=20, t=180, b=30))
		fig.update_layout(height=760)
		st.plotly_chart(fig, use_container_width=True)
	else:
		st.info("Status column not found.")

	# 3) Full-width: Work Orders Over Time (monthly with all months shown)
	if c_date:
		date_series = pd.to_datetime(df[c_date], errors="coerce")
		valid = date_series.notna()
		if valid.any():
			month_period = date_series.dt.to_period("M")
			cnt = month_period.value_counts().rename_axis("month").reset_index(name="count")
			min_p, max_p = month_period[valid].min(), month_period[valid].max()
			all_months = pd.period_range(min_p, max_p, freq="M")
			full_df = pd.DataFrame({"month": all_months}).merge(cnt, on="month", how="left").fillna({"count": 0})
			full_df["month_start"] = full_df["month"].dt.to_timestamp()
			fig = px.line(full_df, x="month_start", y="count", title="Work Orders Over Time", markers=True,
						  color_discrete_sequence=ORANGE_SEQ)
			fig.update_xaxes(dtick="M1", tickformat="%b %Y")
			fig = _apply_chart_theme(fig, "Work Orders Over Time")
			fig.update_layout(height=500)
			st.plotly_chart(fig, use_container_width=True)
		else:
			st.info("No valid dates found for time series chart.")
	else:
		st.info("Date column not found.")

	# 4) Full-width: Work Orders by Planner (aggregated for readability)
	if c_planner:
		plan_counts = df.groupby(c_planner).size().reset_index(name="count").sort_values("count", ascending=False)
		fig = px.bar(
			plan_counts,
			x=c_planner,
			y="count",
			title="Work Orders by Planner",
			text="count",
			color_discrete_sequence=COLOR_SEQ,
			labels={c_planner: "Planner", "count": "WO count"},
		)
		fig.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
		fig.update_traces(hovertemplate="Planner=%{x}<br>WO count=%{y}<extra></extra>")
		fig.update_xaxes(tickangle=-60, tickfont=dict(size=10), categoryorder="total descending")
		fig = _apply_chart_theme(fig, "Work Orders by Planner")
		fig.update_layout(height=520)
		st.plotly_chart(fig, use_container_width=True)
	else:
		st.info("Planner column not found.")

	# 5) Full-width: Work Orders by Category
	if c_category:
		cat_counts = df.groupby(c_category).size().reset_index(name="count").sort_values("count", ascending=False)
		fig = px.bar(
			cat_counts,
			x=c_category,
			y="count",
			title="Work Orders by Category",
			text="count",
			color_discrete_sequence=COLOR_SEQ,
			labels={c_category: "Category", "count": "WO count"},
		)
		fig.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
		fig.update_traces(hovertemplate="Category=%{x}<br>WO count=%{y}<extra></extra>")
		fig.update_xaxes(tickangle=-30, tickfont=dict(size=12), categoryorder="total descending")
		fig = _apply_chart_theme(fig, "Work Orders by Category")
		fig.update_layout(height=500)
		st.plotly_chart(fig, use_container_width=True)
	else:
		st.info("Work Category column not found.")

	# 6) Full-width: Work Orders by Type
	if c_type:
		type_counts = df.groupby(c_type).size().reset_index(name="count").sort_values("count", ascending=False)
		fig = px.bar(
			type_counts,
			x=c_type,
			y="count",
			title="Work Orders by Type",
			text="count",
			color_discrete_sequence=COLOR_SEQ,
			labels={c_type: "Type", "count": "WO count"},
		)
		fig.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
		fig.update_traces(hovertemplate="Type=%{x}<br>WO count=%{y}<extra></extra>")
		fig.update_xaxes(tickangle=-30, tickfont=dict(size=12), categoryorder="total descending")
		fig = _apply_chart_theme(fig, "Work Orders by Type")
		fig.update_layout(height=480)
		st.plotly_chart(fig, use_container_width=True)
	else:
		st.info("Work Type column not found.")

	# 7) Full-width: On-time vs Delayed per month
	# Build month column from completion date if available else reported
	date_for_month = None
	if "Completion Date" in df.columns:
		date_for_month = pd.to_datetime(df["Completion Date"], errors="coerce")
	elif c_date:
		date_for_month = pd.to_datetime(df[c_date], errors="coerce")

	if date_for_month is not None:
		month = date_for_month.dt.to_period("M").astype(str)
		# On-time/Delayed only when both Target and Completion available
		target = pd.to_datetime(df.get("Target Date"), errors="coerce") if "Target Date" in df.columns else None
		if target is not None and "Completion Date" in df.columns:
			completion = pd.to_datetime(df["Completion Date"], errors="coerce")
			on_time = (completion.notna() & target.notna() & (completion <= target))
			delayed = (completion.notna() & target.notna() & (completion > target))
			dt = pd.DataFrame({"month": month, "On Time": on_time.astype(int), "Delayed": delayed.astype(int)})
			grp = dt.groupby("month").sum().reset_index()
			long = grp.melt(id_vars="month", var_name="Status", value_name="count")
			fig = px.bar(
				long,
				x="month",
				y="count",
				color="Status",
				barmode="stack",
				color_discrete_sequence=["#2ca02c", "#d62728"],
				title="On-Time vs Delayed by Month",
				labels={"month": "Month", "count": "WO count", "Status": "Status"},
				custom_data=["Status"],
			)
			fig.update_traces(hovertemplate="Month=%{x}<br>Status=%{customdata[0]}<br>WO count=%{y}<extra></extra>")
			fig = _apply_chart_theme(fig, "On-Time vs Delayed by Month")
			fig.update_layout(height=520)
			st.plotly_chart(fig, use_container_width=True)
		else:
			st.info("Target/Completion dates not sufficient for on-time analysis.")

		# Completed MoM trend (COMP or CLOSE) with all months shown (including zero-count months)
		if c_status is not None:
			status_u = df[c_status].astype(str).str.upper()
			valid_dates = date_for_month.notna()
			if valid_dates.any():
				is_done = status_u.isin(["COMP", "CLOSE"]) & valid_dates
				# Work with Period[M] to build a full continuous monthly range
				month_period = date_for_month.dt.to_period("M")
				done_months = month_period[is_done]
				cnt = done_months.value_counts().rename_axis("month").reset_index(name="count") if is_done.any() else pd.DataFrame({"month": [], "count": []})

				# Build full month range from min to max available dates
				min_p = month_period[valid_dates].min()
				max_p = month_period[valid_dates].max()
				if pd.isna(min_p) or pd.isna(max_p):
					st.info("No valid dates for monthly completion trend.")
				else:
					all_months = pd.period_range(min_p, max_p, freq="M")
					full_df = pd.DataFrame({"month": all_months}).merge(cnt, on="month", how="left").fillna({"count": 0})
					# Use month start as datetime for clean date axis; force monthly ticks
					full_df["month_start"] = full_df["month"].dt.to_timestamp()
					fig2 = px.line(full_df, x="month_start", y="count", markers=True, color_discrete_sequence=ORANGE_SEQ, title="Completed Orders per Month")
					fig2.update_xaxes(dtick="M1", tickformat="%b %Y")
					fig2 = _apply_chart_theme(fig2, "Completed Orders per Month")
					fig2.update_layout(height=480)
					st.plotly_chart(fig2, use_container_width=True)
			else:
				st.info("No valid dates for monthly completion trend.")
		else:
			st.info("Status column not found for monthly completion trend.")
	else:
		st.info("No suitable date column for monthly trends.")


# -------------------------
# Pages
# -------------------------

# Modern KPI styling and helpers for the Performance KPIs tab
def inject_perf_kpi_css(brand_color: str = "#FFA500") -> None:
		css = f"""
				<style>
					:root {{
						--brand: {brand_color};
						--bg-card: #ffffff;
						--text: #1f2937;
						--muted: #6b7280;
						--good: #16a34a;
						--warn: #f59e0b;
						--bad: #ef4444;
						--shadow: 0 4px 16px rgba(0,0,0,0.08);
						--radius: 14px;
					}}

					.kpi-section {{
						margin: 8px 0 24px 0;
					}}
					.kpi-title {{
						display:flex; align-items:center; gap:12px;
						margin: 4px 0 12px 0;
						font-weight: 700; color: var(--text);
					}}
					.kpi-title .accent {{
						width: 18px; height: 18px; border-radius: 4px; background: var(--brand);
						box-shadow: 0 6px 16px rgba(255,165,0,0.35);
					}}
					.kpi-subtitle {{
						margin-top:-4px; color: var(--muted); font-size: 0.88rem;
					}}

					.kpi-grid {{
						display: grid;
						grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
						gap: 16px;
					}}

					.kpi-card.modern {{
						background: var(--bg-card);
						border-radius: var(--radius);
						box-shadow: var(--shadow);
						padding: 14px 16px;
						transition: transform .15s ease, box-shadow .15s ease;
						border: 1px solid rgba(0,0,0,0.04);
					}}
					.kpi-card.modern:hover {{ transform: translateY(-2px); box-shadow: 0 10px 24px rgba(0,0,0,0.10); }}

					.kpi-head {{
						display:flex; align-items:center; justify-content:space-between; gap:8px; margin-bottom: 6px;
					}}
					.kpi-left {{ display:flex; align-items:center; gap:10px; min-width:0; }}
					.kpi-icon {{
						width: 30px; height: 30px; display:grid; place-items:center;
						border-radius:10px; background: rgba(0,0,0,0.04); font-size: 18px;
					}}
					.kpi-label {{
						font-weight: 600; color: var(--muted); white-space: nowrap; overflow:hidden; text-overflow: ellipsis;
					}}

					.kpi-tip {{
						position: relative; cursor: help; color: var(--muted); font-size: 14px; line-height: 1;
						padding: 4px 8px; border-radius: 8px; background: rgba(0,0,0,0.04);
					}}
					.kpi-tip:hover::after {{
						content: attr(data-tip);
						position: absolute; z-index: 10; top: 120%; right: 0;
						max-width: 260px; background: #111827; color:#fff; padding:10px 12px; border-radius: 10px; font-size:12px;
						box-shadow: 0 8px 24px rgba(0,0,0,0.25);
						white-space: normal;
					}}

					.kpi-value.modern {{
						font-size: 1.8rem; font-weight: 800; color: var(--text);
						letter-spacing: -0.02em;
					}}
					.kpi-sub {{
						margin-top: 4px; color: var(--muted); font-size: 0.86rem;
					}}

					/* color states for values */
					.kpi-card.modern.good .kpi-value.modern {{ color: var(--good); }}
					.kpi-card.modern.warn .kpi-value.modern {{ color: var(--warn); }}
					.kpi-card.modern.bad  .kpi-value.modern {{ color: var(--bad);  }}

					/* tiny top accent using brand color */
					.kpi-card.modern::before {{
						content:""; display:block; height: 4px; border-radius: 12px 12px 0 0; 
						background: linear-gradient(90deg, var(--brand), rgba(255,165,0,0.25));
						margin: -14px -16px 10px -16px;
					}}

					/* Typography smoothing */
					.kpi-value.modern, .kpi-label, .kpi-title {{ -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }}
				</style>
				"""
		st.markdown(dedent(css), unsafe_allow_html=True)

# Ensure HTML starts at column 0 to avoid Markdown code blocks
def _flush_left(s: str) -> str:
	try:
		return "\n".join((line.lstrip() for line in s.splitlines()))
	except Exception:
		return s

def _fmt(v: Any, unit: str = "") -> str:
		if v is None or v == "-" or (isinstance(v, float) and (v != v)):
				return "—"
		try:
				if unit == "%":
						return f"{float(v):.0f}%"
				if unit == "pct1":
						return f"{float(v):.1f}%"
				if unit == "d":
						return f"{float(v):.1f} d"
				if unit == "int":
						return f"{int(v):,}"
				return str(v)
		except Exception:
				return str(v)

def _state(metric_key: str, val: Optional[float]) -> str:
		if val is None:
				return ""
		try:
				v = float(val)
		except Exception:
				return ""

		# Higher is better
		higher_better = {"completion_rate", "on_time_rate", "pm_pct_pc", "first_time_fix", "pm_compliance"}
		# Lower is better
		lower_better = {"backlog_rate", "rework_rate", "overdue_open", "avg_backlog_age"}

		if metric_key in higher_better:
				if v >= 90: return "good"
				if v >= 75: return "warn"
				return "bad"
		if metric_key in lower_better:
				if v <= 10: return "good"
				if v <= 20: return "warn"
				return "bad"
		return ""

def _kpi_html(icon: str, label: str, value: str, tip: Optional[str], state_cls: str) -> str:
	tip_html = f'<span class="kpi-tip" data-tip="{tip}">?</span>' if tip else ""
	# Build without any leading indentation
	return (
		f'<div class="kpi-card modern {state_cls}">'  
		f'<div class="kpi-head">'  
		f'<div class="kpi-left">'  
		f'<div class="kpi-icon">{icon}</div>'  
		f'<div class="kpi-label">{label}</div>'  
		f'</div>'  
		f'{tip_html}'  
		f'</div>'  
		f'<div class="kpi-value modern">{value}</div>'  
		f'</div>'
	)

def render_kpi_section(title: str, items: List[Dict[str, Any]], subtitle: Optional[str] = None) -> None:
    items_html = ''.join([
        _kpi_html(
            icon=i.get('icon','📊'),
            label=i['label'],
            value=_fmt(i.get('value'), i.get('unit','')),
            tip=i.get('tip'),
            state_cls=_state(i.get('key',''), i.get('value'))
        ) for i in items
    ])
    subtitle_html = f'<div class="kpi-subtitle">{subtitle}</div>' if subtitle else ''
    html = (
        '<div class="kpi-section">'
        '<div class="kpi-title"><span class="accent"></span><span>' + str(title) + '</span></div>'
        + subtitle_html +
        '<div class="kpi-grid">' + items_html + '</div>'
        '</div>'
    )
    st.markdown(_flush_left(html), unsafe_allow_html=True)

def page_kpis(df: pd.DataFrame, updated_at: Optional[datetime]) -> None:
	# Compute core metrics
	kpis = compute_kpis(df)
	ins = compute_insights(df)

	# Brand-aware KPI UI
	inject_perf_kpi_css("#FFA500")

	# Map metrics to modern sections
	overall = [
		{"key":"", "label":"Total WOs",   "value": kpis.get("total"),     "unit":"int", "icon":"🧰", "tip":"Total work orders in scope after exclusions."},
		{"key":"", "label":"Open Orders", "value": kpis.get("open"),      "unit":"int", "icon":"📂", "tip":"Currently open work orders."},
		{"key":"", "label":"Completed",   "value": kpis.get("completed"), "unit":"int", "icon":"✅", "tip":"Completed (may include closed)."},
		{"key":"", "label":"Closed",      "value": kpis.get("closed"),    "unit":"int", "icon":"📦", "tip":"Fully closed WOs."},
		{"key":"", "label":"Rejected",    "value": kpis.get("rejected"),  "unit":"int", "icon":"❌", "tip":"Rejected/voided after QC."},
	]
	render_kpi_section(
		"Overall Performance",
		overall,
		subtitle=f"Last Updated: {updated_at.strftime('%Y-%m-%d %H:%M:%S') if updated_at else datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
	)

	completion = [
		{"key":"completion_rate", "label":"Completion Rate", "value": kpis.get("completion_rate"), "unit":"%", "icon":"📈",
		 "tip":"Completed / (Completed + Open). Target ≥ 90%"},
		{"key":"on_time_rate", "label":"On-Time Rate", "value": kpis.get("pct_on_time"), "unit":"%", "icon":"⏱️",
		 "tip":"Completed on or before target date. Target ≥ 90%"},
		{"key":"avg_completion", "label":"Avg Completion", "value": kpis.get("avg_days"), "unit":"d", "icon":"🕰️",
		 "tip":"Average turnaround time from creation to completion."},
		{"key":"rework_rate", "label":"Rework Rate", "value": kpis.get("rework_rate"), "unit":"%", "icon":"🧭",
		 "tip":"Percent of WOs that required rework. Lower is better."},
		{"key":"backlog_rate", "label":"Backlog Rate", "value": kpis.get("backlog_rate"), "unit":"%", "icon":"📥",
		 "tip":"Open / Total. Lower is better; target ≤ 10%"},
	]
	render_kpi_section("Completion Metrics", completion)

	split = [
		{"key":"pm_pct_pc", "label":"PM % (PM+CM)", "value": kpis.get("pm_pct_pc"), "unit":"%", "icon":"🛠️",
		 "tip":"Share of Preventive in PM+CM only."},
		{"key":"cm_pct_pc", "label":"CM % (PM+CM)", "value": kpis.get("cm_pct_pc"), "unit":"%", "icon":"🔧",
		 "tip":"Share of Corrective in PM+CM only."},
		{"key":"", "label":"CM/PM Ratio", "value": kpis.get("cm_pm_ratio"), "unit":"pct1", "icon":"📊",
		 "tip":"Corrective to Preventive ratio (CM / PM)."},
	]
	render_kpi_section("Work Type Split", split)

	# Service Provider sections
	svc_top_line = [
		{"key":"", "label":"Total WOs (Svc)", "value": ins.get("total_wos"), "unit":"int", "icon":"🧰",
		 "tip":"Total service-provider work orders in scope."},
		{"key":"completion_rate", "label":"Completion Rate", "value": ins.get("completion_rate"), "unit":"%", "icon":"📈",
		 "tip":"Provider completion rate; target ≥ 90%."},
		{"key":"on_time_rate", "label":"On-Time (valid)", "value": ins.get("on_time_rate"), "unit":"%", "icon":"⏱️",
		 "tip":"On-time only where target dates exist."},
		{"key":"", "label":"MoM Completed", "value": ins.get("mom_change"), "unit":"pct1", "icon":"📅",
		 "tip":"Month-over-month change in completed count."},
		{"key":"", "label":"Top Category", "value": ins.get("top_category"), "unit":"", "icon":"🏷️",
		 "tip":"Category with most WOs for the period."},
	]
	render_kpi_section("Service Provider Performance", svc_top_line)

	if ins.get("top_planner"):
		st.info(f"Top Planner: {ins.get('top_planner')}", icon="🧑‍💼")

	svc_ops = [
		{"key":"first_time_fix", "label":"First-Time Fix", "value": ins.get("ftfr"), "unit":"%", "icon":"🧩",
		 "tip":"Fixed on first visit; higher is better."},
		{"key":"", "label":"Done (Svc)", "value": ins.get("done_wos"), "unit":"int", "icon":"🧾",
		 "tip":"Completed by service provider."},
		{"key":"overdue_open", "label":"Overdue Open", "value": ins.get("overdue_open_rate"), "unit":"%", "icon":"⏰",
		 "tip":"Open WOs past target date; lower is better."},
		{"key":"avg_backlog_age", "label":"Avg Backlog Age", "value": ins.get("avg_backlog_age"), "unit":"d", "icon":"📦",
		 "tip":"Average age of open backlog in days."},
		{"key":"pm_compliance", "label":"PM Compliance", "value": ins.get("pm_compliance"), "unit":"%", "icon":"✅",
		 "tip":"On-time PMs / scheduled PMs; target ≥ 95%."},
	]
	render_kpi_section("Operational Quality (Svc)", svc_ops)

	# On-time completion rate per Location (Service Level per Location)
	st.markdown("---")
	st.subheader("Service Level per Location (On-Time %)")
	c_location = next((c for c in ["Location", "Site/Location", "Site"] if c in df.columns), None)
	if c_location and ("Target Date" in df.columns) and ("Completion Date" in df.columns):
		tgt = pd.to_datetime(df["Target Date"], errors="coerce")
		comp = pd.to_datetime(df["Completion Date"], errors="coerce")
		valid = comp.notna() & tgt.notna()
		if valid.any():
			dloc = df.loc[valid, [c_location]].copy()
			dloc["on_time"] = (comp[valid] <= tgt[valid]).astype(int)
			grp = dloc.groupby(c_location)["on_time"].agg(["count", "sum"]).reset_index()
			grp["on_time_rate"] = 100 * grp["sum"] / grp["count"].replace(0, 1)
			grp = grp.sort_values("on_time_rate", ascending=False)
			fig_sl = px.bar(
				grp,
				x=c_location,
				y="on_time_rate",
				title="On-Time Completion Rate by Location",
				labels={c_location: "Location", "on_time_rate": "On-Time %"},
				text=grp["on_time_rate"].round(0).astype(int).astype(str) + "%",
				color_discrete_sequence=ORANGE_SEQ,
			)
			fig_sl.update_traces(textposition="outside", cliponaxis=False)
			fig_sl.update_xaxes(tickangle=-45, tickfont=dict(size=10))
			fig_sl = _apply_chart_theme(fig_sl, "On-Time Completion Rate by Location")
			fig_sl.update_layout(height=520)
			st.plotly_chart(fig_sl, use_container_width=True)
		else:
			st.info("No valid Target/Completion dates to compute on-time rate per Location.")
	else:
		st.info("Columns for Location/Target/Completion are not sufficient to compute per-location service level.")


def page_analytics(df: pd.DataFrame) -> None:
	# Display-only analytics
	make_charts(df)


# -------------------------
# Main
# -------------------------

def main() -> None:
	inject_brand_css()
	brand_header()
	# Pure display: no sidebar, no inputs

	# Optional daily auto-refresh (24h) if streamlit_autorefresh is installed
	try:
		from streamlit_autorefresh import st_autorefresh  # type: ignore
		# Refresh every 60 seconds
		st_autorefresh(interval=60_000, key="refresh")
	except Exception:
		pass

	# Top tabs for switching pages (Analytics first)
	tabs = st.tabs(["Analytics & Charts", "Performance KPIs"])

	# Load data
	df: Optional[pd.DataFrame] = None
	updated_at: Optional[datetime] = None
	error = None
	data_url = FIXED_DATA_URL
	if (not data_url) or ("example.com" in str(data_url)):
		error = "No fixed data URL configured."
	else:
		try:
			with st.spinner("👉 Importing data from MAXIMO..."):
				df, updated_at = load_data(data_url)
		except KeyboardInterrupt:
			error = "Data loading interrupted by user."
		except Exception as exc:
			error = str(exc)

	if error:
		st.warning("Using sample data (failed to load fixed data URL). Configure MMC_DATA_URL env var or update FIXED_DATA_URL in code.")
		df = make_sample_data()
		updated_at = datetime.now()

	if df is None:
		return

	# Exclude cancelled and Work Category == 'Contract' from all stats/charts globally
	status_col = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	if status_col:
		st_u = df[status_col].astype(str).str.upper().str.strip()
		st_u = st_u.replace({"CANCELLED": "CAN", "CANCELED": "CAN", "CANCEL": "CAN"})
		df = df[st_u != "CAN"].copy()
	cat_col = next((c for c in ["Work Category", "Category"] if c in df.columns), None)
	if cat_col:
		df = df[df[cat_col].astype(str).str.upper().str.strip() != "CONTRACT"].copy()

	# Exclude specific location globally from all stats/charts: 'MMC-ARSAL'
	loc_col = next((c for c in ["Location", "Site/Location", "Site"] if c in df.columns), None)
	if loc_col:
		loc_u = df[loc_col].astype(str).str.upper().str.strip()
		df = df[loc_u != "MMC-ARSAL"].copy()

	# Clean & de-duplicate to avoid double counting and keep the most informative/latest rows
	def _deduplicate_rows(dfin: pd.DataFrame) -> pd.DataFrame:
		# Trim strings
		obj_cols = dfin.select_dtypes(include=["object"]).columns
		for c in obj_cols:
			dfin[c] = dfin[c].astype(str).str.strip()
		# Candidate unique ID columns
		id_candidates = [
			"WO", "WO Number", "WONUM", "Work Order", "WORKORDER", "TICKET", "Ticket ID", "SR", "REQUESTID"
		]
		id_col = next((c for c in id_candidates if c in dfin.columns), None)
		if not id_col:
			return dfin.drop_duplicates().reset_index(drop=True)
		# Build a composite sort key: more non-nulls first, then latest timestamp across known date columns
		date_cols = [
			"Completion Date", "Actual Finish", "Target Date", "Reported Date", "Created Date", "Create Date", "Status Date", "Last Modified", "Change Date"
		]
		present = [c for c in date_cols if c in dfin.columns]
		if present:
			ts = pd.DataFrame({c: pd.to_datetime(dfin[c], errors="coerce") for c in present}).max(axis=1)
		else:
			ts = pd.Series([pd.NaT] * len(dfin))
		df2 = dfin.copy()
		df2["__nn"] = dfin.notna().sum(axis=1)
		df2["__ts"] = ts
		df2 = df2.sort_values(["__nn", "__ts"], ascending=[False, False])
		df2 = df2.drop_duplicates(subset=id_col, keep="first").drop(columns=["__nn", "__ts"], errors="ignore")
		return df2.reset_index(drop=True)

	df = _deduplicate_rows(df)

	with tabs[0]:
		page_analytics(df)
	with tabs[1]:
		page_kpis(df, updated_at)


if __name__ == "__main__":
	main()

=======
"""
Arsal Facility Management – MMC Dashboard (single-file Streamlit)

Run locally:
	streamlit run "Maximo_DEV/MMC_Dash03.py"

Requirements (install in your venv):
	pip install streamlit pandas plotly requests openpyxl

Deployment notes (Streamlit Community Cloud):
- Set your secrets in the app settings (MMC_DATA_URL, MMC_EXCEL_SHEET, MMC_ARSAL_LOGO, MMC_CLIENT_LOGO).
- A fallback to environment variables is supported for local runs.
- Add runtime.txt with Python version (e.g., 3.11) for best compatibility.
"""

from __future__ import annotations

import io
import os
from datetime import datetime, timedelta
import base64
from typing import Optional, Tuple, Dict, Any, List
from textwrap import dedent
import warnings

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
# Note: components import removed as it's unused in this display-only version


# -------------------------
# Page config and theme
# -------------------------
st.set_page_config(
	page_title="Arsal FM – MMC Dashboard",
	page_icon="🟧",
	layout="wide",
	initial_sidebar_state="collapsed",
)

PRIMARY = "#FF6B00"      # Bright Orange
SECONDARY = "#F5F5F5"    # Light Gray
ACCENT = "#333333"       # Charcoal Gray
TEXT = "#000000"         # Black

# Brand assets (direct view links)
arsal_logo_url = "https://drive.google.com/uc?export=view&id=1lAmXKLuyFH0IS840z7PoecsU_eYSsloh"
client_logo_url = "https://drive.google.com/uc?export=view&id=1XHP8Str5Iwt9YinIgfisG6xTnOM222Py"
def get_secret_env(key: str, default: Optional[str] = None) -> Optional[str]:
	"""Read a config value from st.secrets first, then environment variables.

	Returns default if the key is not present or empty in both places.
	"""
	try:
		# st.secrets behaves like a dict; guard for when secrets not configured
		if hasattr(st, "secrets") and key in st.secrets and str(st.secrets.get(key)).strip():
			return str(st.secrets.get(key))
	except Exception:
		pass
	val = os.environ.get(key)
	return val if (val is not None and str(val).strip()) else default



def _gdrive_to_direct(u: str) -> str:
	"""Convert Google Drive file link to direct view URL if needed."""
	if not u:
		return u
	if "drive.google.com/file/d/" in u:
		try:
			file_id = u.split("/file/d/")[1].split("/")[0]
			return f"https://drive.google.com/uc?export=view&id={file_id}"
		except Exception:
			return u
	return u


def _file_to_data_uri(path: str) -> Optional[str]:
	try:
		with open(path, "rb") as f:
			b = f.read()
		mime = "image/png"
		ext = os.path.splitext(path)[1].lower()
		if ext in [".jpg", ".jpeg"]:
			mime = "image/jpeg"
		elif ext == ".svg":
			mime = "image/svg+xml"
		elif ext == ".webp":
			mime = "image/webp"
		return f"data:{mime};base64,{base64.b64encode(b).decode()}"
	except Exception:
		return None


def _resolve_logo_src(default_url: str, env_key: str, local_names: list[str], keywords: Optional[list[str]] = None) -> str:
	"""Prefer local asset (as data URI), then env URL, then default URL.

	Additionally supports fuzzy search by keywords in the assets folder (case-insensitive),
	so names like "MMC Logo.jpg" or "ARSAL Logo.png" are detected automatically.
	"""
	# 1) Local assets under Maximo_DEV/assets (exact names first)
	base_dir = os.path.dirname(__file__)
	assets_dir = os.path.join(base_dir, "assets")
	for name in local_names:
		p = os.path.join(assets_dir, name)
		if os.path.exists(p):
			data_uri = _file_to_data_uri(p)
			if data_uri:
				return data_uri
	# 1b) Fuzzy search by keywords
	if os.path.isdir(assets_dir) and keywords:
		try:
			for f in os.listdir(assets_dir):
				fl = f.lower()
				if any(kw.lower() in fl for kw in keywords) and os.path.splitext(f)[1].lower() in [".png", ".jpg", ".jpeg", ".svg", ".webp"]:
					p = os.path.join(assets_dir, f)
					data_uri = _file_to_data_uri(p)
					if data_uri:
						return data_uri
		except Exception:
			pass
	# 2) Secrets/Environment override
	env_url = get_secret_env(env_key)
	if env_url:
		return _gdrive_to_direct(env_url)
	# 3) Default
	return _gdrive_to_direct(default_url)


def _parse_dates_safely(series: pd.Series) -> pd.Series:
	"""Parse a date-like Series robustly without noisy warnings.

	Tries a set of common formats first (vectorized and fast), picks the one with
	the most non-NaT matches. Falls back to generic parser with warnings suppressed.
	"""
	if series is None:
		return series
	# Ensure string input (avoid mixed types) and strip whitespace
	s = series.astype(str).str.strip()
	# If already datetime-like, return as-is
	if getattr(series, 'dtype', None) is not None and str(series.dtype).startswith('datetime'):
		return series

	formats = [
		"%Y-%m-%d",
		"%Y/%m/%d",
		"%d/%m/%Y",
		"%m/%d/%Y",
		"%d-%m-%Y",
		"%m-%d-%Y",
		"%Y-%m-%d %H:%M:%S",
		"%d/%m/%Y %H:%M",
		"%m/%d/%Y %H:%M",
	]
	bests: Tuple[Optional[pd.Series], int] = (None, -1)
	for fmt in formats:
		parsed = pd.to_datetime(s, format=fmt, errors="coerce")
		non_na = int(parsed.notna().sum())
		if non_na > bests[1]:
			bests = (parsed, non_na)
			if non_na == len(s):
				break

	if bests[0] is not None and bests[1] > 0:
		return bests[0]

	# Fallback: generic parsing with warnings suppressed; prefer dayfirst to handle DD/MM/YYYY
	with warnings.catch_warnings():
		warnings.simplefilter("ignore", category=UserWarning)
		return pd.to_datetime(s, errors="coerce", dayfirst=True)


def _svg_placeholder(text: str, bg: str = "#FF6B00", fg: str = "#FFFFFF") -> str:
	"""Create a simple SVG badge as data URI used as a fallback logo."""
	svg = f"""
	<svg xmlns='http://www.w3.org/2000/svg' width='128' height='64'>
		<rect width='100%' height='100%' fill='{bg}' rx='8' ry='8'/>
		<text x='50%' y='50%' dominant-baseline='middle' text-anchor='middle'
				font-family='Segoe UI, Poppins, Arial' font-size='28' font-weight='700' fill='{fg}'>
			{text}
		</text>
	</svg>
	""".strip()
	uri = "data:image/svg+xml;base64," + base64.b64encode(svg.encode("utf-8")).decode("ascii")
	return uri

# Qualitative, diverse color palette for charts (not only orange)
COLOR_SEQ = (
	px.colors.qualitative.Vivid
	+ px.colors.qualitative.Safe
	+ px.colors.qualitative.Set2
	+ px.colors.qualitative.Pastel
)
ORANGE_SEQ = ["#FF6B00", "#FF8C33", "#FFB366", "#FFD1A3"]

# Fixed OneDrive/SharePoint CSV direct link (update this to your public CSV download URL)
# Example patterns:
# - SharePoint: https://<tenant>.sharepoint.com/.../download.aspx?share=...
# - OneDrive: https://api.onedrive.com/v1.0/shares/.../root/content
# - Google Sheets CSV export: https://docs.google.com/spreadsheets/d/<ID>/export?format=csv
FIXED_DATA_URL = get_secret_env(
	"MMC_DATA_URL",
	# Default to your Google Sheets link; loader will convert it to CSV export automatically
	"https://docs.google.com/spreadsheets/d/1T6dndJHd33ZW3i4e9LIOGl1BPkxDsYme/edit?usp=sharing&ouid=115201289744778991707&rtpof=true&sd=true",
)


def inject_brand_css() -> None:
	st.markdown(
		f"""
		<style>
		:root {{
			--accent: {PRIMARY};
			--light-bg: #fffaf5;
			--shadow: 0 8px 18px rgba(0,0,0,0.08);
			--card-radius: 12px;
		}}
		body {{
			background: linear-gradient(135deg, var(--light-bg), #ffffff);
			font-family: 'Poppins', 'Segoe UI', sans-serif;
			zoom: 1.05; /* Slight scale-up for TV readability */
		}}
		.brand-header {{
			display: grid;
			/* Fix left/right columns to keep logo + title perfectly aligned across tabs */
			grid-template-columns: 160px 1fr 160px;
			align-items: center;
			/* Fix header block height so it doesn't jump between tabs */
			min-height: 110px; /* Taller to fit subtitle line */
			padding: 0.6rem 1rem; background: rgba(255,255,255,0.9);
			border-bottom: 2px solid var(--accent);
			box-shadow: 0 2px 10px rgba(0,0,0,0.05);
			border-radius: 12px; backdrop-filter: blur(8px);
			margin-bottom: 8px;
		}}
		/* Fix logo box size so layout stays stable when switching tabs */
		.logo-img {{ height:64px; width:64px; border-radius:10px; object-fit: contain; }}
		.brand-left {{ text-align: left; }}
		.brand-center {{ text-align: center; }}
		.brand-right {{ text-align: right; }}
		.brand-subtitle {{ margin: 4px 0 0 0; color:#555; font-size:14px; font-weight:600; }}
		/* Keep title on one line and consistent size across tabs */
		.brand-center h3 {{
			font-size: 38px;
			line-height: 1.2;
			margin: 0;
			white-space: nowrap;
		}}
		.kpi-card {{
			background: rgba(255,255,255,0.9);
			border: 1px solid rgba(255,255,255,0.5);
			border-radius: 20px;
			box-shadow: 0 2px 10px rgba(0,0,0,0.08);
			padding: 18px;
		}}
		.kpi-title {{ color:#666; font-size:18px; margin:0; }}
		.kpi-value {{ color:#111; font-weight:900; font-size:44px; margin:0; }}
		.kpi-emoji {{ font-size:24px; margin-right:8px; }}
		h1, h2, h3 {{ color: var(--accent); }}
		.stButton>button {{
			background: linear-gradient(135deg, var(--accent), #ff933f);
			color: white; border: none; border-radius: 10px; padding: 10px 18px; font-weight:700;
			box-shadow: 0 6px 14px rgba(255,107,0,0.22);
		}}
		.stButton>button:hover {{ filter: brightness(1.05); }}
		.card {{ background:white; border:1px solid rgba(0,0,0,0.08); border-radius: var(--card-radius); box-shadow: var(--shadow); padding: 10px 12px; }}
		.muted {{ color:#666; font-size:14px; }}
		header {{ visibility: hidden; }}
		/* Remove extra padding for TV style */
		.block-container {{ padding-top: 0.8rem; padding-bottom: 0.8rem; }}
		</style>
		""",
		unsafe_allow_html=True,
	)


## Removed sidebar hints and inputs to make a pure display dashboard


def brand_header() -> None:
	"""Top header with Aarsal logo (left), centered title, and Client logo (right)."""
	# Current date & time (local)
	now_str = datetime.now().strftime("%a, %d %b %Y – %H:%M")
	left_src = _resolve_logo_src(
		default_url=arsal_logo_url,
		env_key="MMC_ARSAL_LOGO",
		local_names=["arsal_logo.png", "arsal_logo.jpg", "arsal_logo.jpeg", "arsal_logo.svg", "arsal_logo.webp"],
		keywords=["arsal", "aarsal", "arsal logo"],
	)
	right_src = _resolve_logo_src(
		default_url=client_logo_url,
		env_key="MMC_CLIENT_LOGO",
		local_names=["client_logo.png", "client_logo.jpg", "client_logo.jpeg", "client_logo.svg", "client_logo.webp"],
		keywords=["client", "mmc", "almajdouie", "mmc logo"],
	)
	left_fallback = _svg_placeholder("Arsal", bg=PRIMARY)
	right_fallback = _svg_placeholder("Client", bg=PRIMARY)
	st.markdown(
		f"""
		<div class='brand-header'>
			<div class='brand-left'>
				<img src="{left_src}" class="logo-img" onerror="this.onerror=null; this.src='{left_fallback}';" />
			</div>
			<div class='brand-center'>
				<h3 style='margin:0;'>MMC Project Dashboard</h3>
				<div class='brand-subtitle'>{now_str}</div>
			</div>
			<div class='brand-right'>
				<img src="{right_src}" class="logo-img" onerror="this.onerror=null; this.src='{right_fallback}';" />
			</div>
		</div>
		""",
		unsafe_allow_html=True,
	)


# -------------------------
# Data loading and caching
# -------------------------

@st.cache_data(ttl=60, show_spinner=False)
def load_data(url: str) -> Tuple[pd.DataFrame, datetime]:
	"""Download CSV from fixed link (SharePoint/OneDrive/Google Sheets CSV export)."""
	if not url:
		raise ValueError("No data URL configured.")
	headers = {"User-Agent": "Mozilla/5.0"}

	def ensure_csv(u: str) -> str:
		# Google Sheets share -> force CSV
		if "docs.google.com/spreadsheets" in u and "/export?" not in u:
			try:
				sheet_id = u.split("/d/")[1].split("/")[0]
				return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
			except Exception:
				return u
		# Encourage direct content for SharePoint/OneDrive links
		if ("download.aspx" in u) or ("download=1" in u) or ("/content" in u):
			return u
		return f"{u}&download=1" if ("?" in u) else f"{u}?download=1"

	eff_url = ensure_csv(url)
	try:
		# Use a shorter timeout to avoid long hangs on TVs
		resp = requests.get(eff_url, headers=headers, timeout=12, allow_redirects=True)
		resp.raise_for_status()
	except requests.RequestException as e:
		raise RuntimeError(f"HTTP error while fetching CSV. Ensure the link is public/direct. Details: {e}") from e

	# Try CSV first; if it fails, try Excel as fallback
	content = resp.content
	with io.BytesIO(content) as buf:
		try:
			df = pd.read_csv(buf, on_bad_lines="skip")
		except Exception:
			buf.seek(0)
			try:
				# Read Excel completely. If MMC_EXCEL_SHEET is set, select that sheet (case-insensitive),
				# otherwise read and concatenate all sheets to ensure we don't miss data.
				xls = pd.ExcelFile(buf, engine="openpyxl")
				prefer_sheet = get_secret_env("MMC_EXCEL_SHEET")
				if prefer_sheet:
					# case-insensitive match
					match = None
					for sn in xls.sheet_names:
						if sn.strip().lower() == prefer_sheet.strip().lower():
							match = sn
							break
					sheet_to_read = match or xls.sheet_names[0]
					df = xls.parse(sheet_name=sheet_to_read, engine="openpyxl")
				else:
					frames = []
					for sn in xls.sheet_names:
						frames.append(xls.parse(sheet_name=sn, engine="openpyxl"))
					df = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
			except Exception as e:
				raise RuntimeError("Unable to parse data as CSV or Excel.") from e

	# Normalize columns
	df.columns = [str(c).strip() for c in df.columns]
	# Parse likely date columns (best-effort)
	for col in [
		"Reported Date",
		"Create Date",
		"Created Date",
		"Target Date",
		"Completion Date",
		"Actual Finish",
		"Actual Start",
		"Status Date",
		"Last Modified",
		"Change Date",
	]:
		if col in df.columns:
			df[col] = _parse_dates_safely(df[col])

	return df, datetime.now()


def clear_cache_and_rerun():
	st.cache_data.clear()
	try:
		# Newer Streamlit versions
		st.rerun()
	except Exception:
		# Backward compatibility
		st.experimental_rerun()


def make_sample_data() -> pd.DataFrame:
	"""Fallback sample dataset for TV demo when data URL is unavailable."""
	now = datetime.now()
	data = {
		"Status": [
			"WPLAN", "PLANCOMP", "WQAPPRC", "QAPPRC", "WSCH", "SCHEDCOMP",
			"COMP", "COMP", "QREJECTC", "WPLAN", "PLANCOMP", "CAN"
		],
		"Work Type": [
			"PM", "CM", "ADW", "PM", "CM", "ADW",
			"PM", "CM", "ADW", "PM", "CM", "PM"
		],
		"Work Category": [
			"Routine", "Routine", "Emergency", "Routine", "Emergency", "Routine",
			"Routine", "Emergency", "Routine", "Contract", "Routine", "Routine"
		],
		"Reported Date": [now - timedelta(days=d) for d in [30,28,25,22,20,18,15,12,10,8,5,2]],
		"Target Date": [now - timedelta(days=d) for d in [25,26,22,20,18,16,12,10,8,6,2,1]],
		"Completion Date": [
			pd.NaT, pd.NaT, pd.NaT, pd.NaT, pd.NaT, pd.NaT,
			now - timedelta(days=13), now - timedelta(days=9), pd.NaT, pd.NaT, pd.NaT, pd.NaT
		],
		"Site/Location": [
			"Site A", "Site A", "Site B", "Site B", "Site C", "Site C",
			"Site A", "Site B", "Site C", "Site A", "Site B", "Site C"
		],
		"Planner Name": [
			"Planner 1", "Planner 2", "Planner 1", "Planner 3", "Planner 2", "Planner 3",
			"Planner 1", "Planner 2", "Planner 1", "Planner 3", "Planner 2", "Planner 3"
		],
	}
	df = pd.DataFrame(data)
	return df


def parse_excel_bytes(content: bytes) -> pd.DataFrame:
	with io.BytesIO(content) as buf:
		# Try openpyxl
		try:
			xls = pd.ExcelFile(buf, engine="openpyxl")
			prefer_sheet = get_secret_env("MMC_EXCEL_SHEET")
			if prefer_sheet:
				match = None
				for sn in xls.sheet_names:
					if sn.strip().lower() == prefer_sheet.strip().lower():
						match = sn
						break
				sheet_to_read = match or xls.sheet_names[0]
				df = xls.parse(sheet_name=sheet_to_read, engine="openpyxl")
			else:
				frames = [xls.parse(sheet_name=sn, engine="openpyxl") for sn in xls.sheet_names]
				df = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
		except Exception:
			# Fallback: reposition and try CSV
			buf.seek(0)
			try:
				df = pd.read_csv(buf)
			except Exception as e:
				raise RuntimeError("Uploaded file format not recognized as Excel/CSV. If it's .xls, install 'xlrd' or re-save as .xlsx.") from e
	# Normalize columns
	df.columns = [str(c).strip() for c in df.columns]
	# Parse likely date columns (best-effort)
	for col in [
		"Reported Date",
		"Create Date",
		"Created Date",
		"Target Date",
		"Completion Date",
		"Actual Finish",
		"Actual Start",
		"Status Date",
		"Last Modified",
		"Change Date",
	]:
		if col in df.columns:
			df[col] = _parse_dates_safely(df[col])
	return df


# -------------------------
# KPI helpers
# -------------------------

def compute_kpis(df: pd.DataFrame) -> Dict[str, Optional[float]]:
	"""Compute KPIs using the specified status mapping and rules.

	Status mapping:
	- WPLAN: Open – not yet acted on
	- PLANCOMP: With procurement
	- WQAPPRC: Waiting client quotation approval
	- QREJECTC: Rejected by client
	- QAPPRC: Client approved (waiting execution)
	- WSCH: Materials delivered / service executed (waiting completion confirmation)
	- SCHEDCOMP: Work executed, pending invoices/docs
	- COMP: Fully completed
	- CAN: Cancelled – exclude from all stats
	"""
	# Normalize status column
	status_col = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	if status_col is not None:
		status_series = df[status_col].astype(str).str.upper().str.strip()
		# Common synonyms/variants
		status_series = (
			status_series
			.replace({
				"CANCELLED": "CAN",
				"CANCELED": "CAN",
				"CANCEL": "CAN",
				"CLOSED": "CLOSE",
			})
		)
	else:
		status_series = pd.Series([None] * len(df))

	# Exclude cancelled and 'Contract' category from all KPIs, and exclude Location == 'MMC-ARSAL'
	not_cancelled = ~(status_series == "CAN")
	cat_col = next((c for c in ["Work Category", "Category"] if c in df.columns), None)
	if cat_col is not None:
		cat_series = df[cat_col].astype(str).str.upper().str.strip()
		not_contract = ~(cat_series == "CONTRACT")
	else:
		not_contract = pd.Series([True] * len(df))

	loc_col = next((c for c in ["Location", "Site/Location", "Site"] if c in df.columns), None)
	if loc_col is not None:
		loc_series = df[loc_col].astype(str).str.upper().str.strip()
		not_mmc_arsal = ~(loc_series == "MMC-ARSAL")
	else:
		not_mmc_arsal = pd.Series([True] * len(df))

	mask = not_cancelled & not_contract & not_mmc_arsal
	dfx = df[mask].copy()
	status_series = status_series[mask]

	total = len(dfx)

	OPEN_CODES = {"WPLAN", "PLANCOMP", "WQAPPRC", "QAPPRC", "WSCH", "SCHEDCOMP"}
	COMPLETED_CODE = "COMP"
	CLOSED_CODE = "CLOSE"
	REJECTED_CODE = "QREJECTC"

	open_mask = status_series.isin(OPEN_CODES)
	completed_mask = status_series.eq(COMPLETED_CODE)
	closed_mask = status_series.eq(CLOSED_CODE)
	rejected_mask = status_series.eq(REJECTED_CODE)

	# Work type distribution (PM, CM, ADW)
	wt_col = next((c for c in ["Work Type", "Type", "WORKTYPE", "worktype"] if c in dfx.columns), None)
	if wt_col:
		wt_upper = dfx[wt_col].astype(str).str.upper()
		pm_mask = wt_upper.str.contains("PM", na=False)
		cm_mask = wt_upper.str.contains("CM", na=False)
		adw_mask = wt_upper.str.contains("ADW", na=False)
	else:
		pm_mask = pd.Series([False] * total)
		cm_mask = pd.Series([False] * total)
		adw_mask = pd.Series([False] * total)
	# Work type percentages and CM/PM ratio
	pm_n = int(pm_mask.sum())
	cm_n = int(cm_mask.sum())
	adw_n = int(adw_mask.sum())
	wt_total = max(1, pm_n + cm_n + adw_n)
	pm_pct = 100 * pm_n / wt_total
	cm_pct = 100 * cm_n / wt_total
	adw_pct = 100 * adw_n / wt_total
	cm_pm_ratio = 100 * (cm_n / max(1, pm_n))
	# PM/CM only share (per request: PM% over PM+CM, CM% = 100 - PM%)
	pmcm_total = max(1, pm_n + cm_n)
	pm_pct_pc = 100 * pm_n / pmcm_total
	cm_pct_pc = 100 - pm_pct_pc

	# Avg completion time
	if "Reported Date" in dfx.columns:
		reported = dfx["Reported Date"]
	elif "Created Date" in dfx.columns:
		reported = dfx["Created Date"]
	else:
		reported = pd.Series([pd.NaT] * total)

	completion = dfx.get("Completion Date", dfx.get("Actual Finish", pd.Series([pd.NaT] * total)))
	valid = reported.notna() & completion.notna()
	avg_days = (completion[valid] - reported[valid]).dt.total_seconds().mean() / 86400 if valid.any() else None

	# % Closed on time (if Target Date exists)
	target = dfx.get("Target Date")
	if target is not None and completion is not None:
		on_time_mask = completion.notna() & target.notna() & (completion <= target)
		# Denominator: completed with valid target+completion
		valid_comp = completion.notna() & target.notna()
		pct_on_time = 100 * (on_time_mask.sum() / max(1, valid_comp.sum()))
	else:
		pct_on_time = None

	# Completion Rate (completed over total, excluding cancelled already)
	completion_rate = 100 * (completed_mask.sum() / max(1, total))
	# Backlog Rate (open over total)
	backlog_rate = 100 * (open_mask.sum() / max(1, total))
	# Rework Rate (QREJECTC over total)
	rework_rate = 100 * (rejected_mask.sum() / max(1, total))

	return {
		"total": total,
		"open": int(open_mask.sum()),
		"completed": int(completed_mask.sum()),
		"closed": int(closed_mask.sum()),
		"rejected": int(rejected_mask.sum()),
		"pm": pm_n,
		"cm": cm_n,
		"adw": adw_n,
		"pm_pct": pm_pct,
		"cm_pct": cm_pct,
		"adw_pct": adw_pct,
		"cm_pm_ratio": cm_pm_ratio,
		"pm_pct_pc": pm_pct_pc,
		"cm_pct_pc": cm_pct_pc,
		"avg_days": avg_days,
		"pct_on_time": pct_on_time,
		"completion_rate": completion_rate,
		"backlog_rate": backlog_rate,
		"rework_rate": rework_rate,
	}


def kpi_card(label: str, value: str, emoji: str) -> None:
	st.markdown(
		f"""
		<div class='kpi-card'>
		  <p class='kpi-title'><span class='kpi-emoji'>{emoji}</span>{label}</p>
		  <p class='kpi-value'>{value}</p>
		</div>
		""",
		unsafe_allow_html=True,
	)


# -------------------------
# Insights
# -------------------------

def compute_insights(df: pd.DataFrame) -> Dict[str, Optional[float]]:
	"""Aggregate additional analytics for service provider performance."""
	out: Dict[str, Optional[float]] = {}
	total = len(df)

	status_col = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	done_mask = pd.Series([False] * total)
	if status_col is not None:
		su = df[status_col].astype(str).str.upper()
		done_mask = su.isin(["COMP", "CLOSE"])  # completed or closed
	done_n = int(done_mask.sum())
	out["completion_rate"] = 100 * (done_n / max(1, total))

	# On-time among those with both dates
	target = pd.to_datetime(df.get("Target Date"), errors="coerce") if "Target Date" in df.columns else None
	completion = pd.to_datetime(df.get("Completion Date"), errors="coerce") if "Completion Date" in df.columns else None
	if target is not None and completion is not None:
		valid = completion.notna() & target.notna()
		on_time = (completion <= target) & valid
		out["on_time_rate"] = 100 * (on_time.sum() / max(1, valid.sum()))
	else:
		out["on_time_rate"] = None

	# Month-over-month change in completed
	if completion is not None:
		months = completion.dt.to_period("M").astype(str)
		ser = pd.DataFrame({"month": months[done_mask]})
		if not ser.empty:
			cnt = ser.value_counts().rename_axis(["month"]).reset_index(name="count").sort_values("month")
			if len(cnt) >= 2:
				last, prev = cnt.iloc[-1]["count"], cnt.iloc[-2]["count"]
				out["mom_change"] = 100 * ((last - prev) / max(1, prev))
				out["last_month"] = str(cnt.iloc[-1]["month"])
			elif len(cnt) == 1:
				out["mom_change"] = None
				out["last_month"] = str(cnt.iloc[-1]["month"])
			else:
				out["mom_change"] = None
				out["last_month"] = None
		else:
			out["mom_change"] = None
			out["last_month"] = None
	else:
		out["mom_change"] = None
		out["last_month"] = None

	# Top category and planner
	cat_col = next((c for c in ["Work Category", "Category"] if c in df.columns), None)
	plan_col = next((c for c in ["Planner Name", "Planner"] if c in df.columns), None)
	try:
		out["top_category"] = (
			df[cat_col].astype(str).value_counts().idxmax() if cat_col else None
		)
	except Exception:
		out["top_category"] = None
	try:
		out["top_planner"] = (
			df[plan_col].astype(str).value_counts().idxmax() if plan_col else None
		)
	except Exception:
		out["top_planner"] = None

	out["total_wos"] = total
	out["done_wos"] = done_n

	# Additional KPIs for provider performance
	# Rework proxy using QREJECTC presence
	status_col2 = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	rework_n = 0
	if status_col2 is not None:
		su2 = df[status_col2].astype(str).str.upper()
		rework_n = int((su2 == "QREJECTC").sum())
	out["rework_rate"] = 100 * (rework_n / max(1, total))

	# Backlog rate (open codes)
	OPEN_CODES = {"WPLAN", "PLANCOMP", "WQAPPRC", "QAPPRC", "WSCH", "SCHEDCOMP"}
	open_n = 0
	if status_col2 is not None:
		su3 = df[status_col2].astype(str).str.upper()
		open_n = int(su3.isin(OPEN_CODES).sum())
	out["backlog_rate"] = 100 * (open_n / max(1, total))

	# Overdue open rate (Target passed and not completed)
	target = pd.to_datetime(df.get("Target Date"), errors="coerce") if "Target Date" in df.columns else None
	if target is not None and status_col2 is not None:
		su4 = df[status_col2].astype(str).str.upper()
		is_open = su4.isin(OPEN_CODES)
		overdue_open = is_open & target.notna() & (target < pd.Timestamp.now())
		out["overdue_open_rate"] = 100 * (overdue_open.sum() / max(1, is_open.sum())) if is_open.any() else None
	else:
		out["overdue_open_rate"] = None

	# Average backlog age (days) for open items
	reported = pd.to_datetime(df.get("Reported Date"), errors="coerce") if "Reported Date" in df.columns else None
	if reported is not None and status_col2 is not None:
		su5 = df[status_col2].astype(str).str.upper()
		is_open = su5.isin(OPEN_CODES)
		ages = (pd.Timestamp.now() - reported[is_open]).dt.total_seconds() / 86400
		out["avg_backlog_age"] = float(ages.mean()) if is_open.any() else None
	else:
		out["avg_backlog_age"] = None

	# SLA breach rate among completed
	if target is not None and completion is not None:
		valid_comp = completion.notna() & target.notna()
		late = valid_comp & (completion > target)
		out["sla_breach_rate"] = 100 * (late.sum() / max(1, valid_comp.sum())) if valid_comp.any() else None
	else:
		out["sla_breach_rate"] = None

	# PM compliance (PM Completed on time / PM Completed with target)
	wt_col = next((c for c in ["Work Type", "Type", "WORKTYPE", "worktype"] if c in df.columns), None)
	if wt_col is not None and target is not None and completion is not None:
		wt_u = df[wt_col].astype(str).str.upper()
		is_pm = wt_u.str.contains("PM", na=False)
		valid = is_pm & completion.notna() & target.notna()
		ontime = valid & (completion <= target)
		out["pm_compliance"] = 100 * (ontime.sum() / max(1, valid.sum())) if valid.any() else None
	else:
		out["pm_compliance"] = None

	# First-Time Fix Rate (approx) = 100 - rework_rate
	out["ftfr"] = 100 - (out["rework_rate"] or 0)
	return out


def insight_card(label: str, value: str, emoji: str) -> None:
	kpi_card(label, value, emoji)

# -------------------------
# Sidebar and filters
# -------------------------

## Removed: build_sidebar (no inputs in display mode)


## Removed: apply_filters (no filters in display mode)


# -------------------------
# Charts
# -------------------------

def _apply_chart_theme(fig, title: str):
	fig.update_layout(
		title={"text": title, "x": 0.02, "xanchor": "left"},
		margin=dict(l=10, r=10, t=50, b=10),
		font=dict(size=18, color="#222"),
		legend=dict(font=dict(size=16)),
		paper_bgcolor="#f5f6fa",
		plot_bgcolor="#eef1f5",
	)
	fig.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.05)")
	fig.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.05)")
	return fig


def make_charts(df: pd.DataFrame) -> None:
	"""Render all charts full-width, stacked vertically for maximum readability."""
	# Prefer Location over Site when both are present
	c_location = next((c for c in ["Location", "Site/Location", "Site"] if c in df.columns), None)
	c_status = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	c_planner = next((c for c in ["Planner Name", "Planner"] if c in df.columns), None)
	c_date = next((c for c in ["Reported Date", "Created Date", "Create Date"] if c in df.columns), None)
	c_category = next((c for c in ["Work Category", "Category"] if c in df.columns), None)
	c_type = next((c for c in ["Work Type", "Type", "WORKTYPE", "worktype"] if c in df.columns), None)

	# 1) Full-width: Work Orders by Location (aggregated) with count labels and small tick text
	if c_location:
		loc_counts = (
			df.groupby(c_location).size().reset_index(name="count").sort_values("count", ascending=False)
		)
		fig_loc = px.bar(
			loc_counts,
			x=c_location,
			y="count",
			title="Work Orders by Location",
			text="count",
			color_discrete_sequence=ORANGE_SEQ,
			labels={c_location: "Location", "count": "WO count"},
		)
		fig_loc.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
		fig_loc.update_traces(hovertemplate="Location=%{x}<br>WO count=%{y}<extra></extra>")
		fig_loc.update_layout(bargap=0.15)
		fig_loc.update_xaxes(tickangle=-90, tickfont=dict(size=9), categoryorder="total descending")
		fig_loc = _apply_chart_theme(fig_loc, "Work Orders by Location")
		fig_loc.update_layout(height=600)
		st.plotly_chart(fig_loc, use_container_width=True)

		# 1b) Full-width: CM vs PM per Location – grouped bars (all records in sheet)
		if c_type:
			df_cmpm = df.copy()

			# Keep only CM and PM
			wt_upper = df_cmpm[c_type].astype(str).str.upper()
			# Include any values that equal or contain CM/PM to be tolerant of variations
			is_cm = wt_upper.str.contains("CM", na=False)
			is_pm = wt_upper.str.contains("PM", na=False)
			df_cmpm = df_cmpm[is_cm | is_pm].copy()
			df_cmpm["_wt"] = wt_upper.where(is_pm, "CM").where(is_cm, "PM")

			if not df_cmpm.empty:
				counts = df_cmpm.groupby([c_location, "_wt"]).size().reset_index(name="count")
				# Order locations by total count desc
				order = counts.groupby(c_location)["count"].sum().sort_values(ascending=False).index.tolist()
				fig_cmpm = px.bar(
					counts,
					x=c_location,
					y="count",
					color="_wt",
					barmode="group",
					title="CM vs PM per Location",
					labels={"count": "WO count", "_wt": "Work Type", c_location: "Location"},
					color_discrete_map={"CM": "#FF6B00", "PM": "#2E86AB"},
					text="count",
				)
				fig_cmpm.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
				fig_cmpm.update_traces(hovertemplate="Location=%{x}<br>Work Type=%{legendgroup}<br>WO count=%{y}<extra></extra>")
				fig_cmpm.update_xaxes(tickangle=-90, tickfont=dict(size=9), categoryorder="array", categoryarray=order)
				fig_cmpm.update_layout(legend_title_text="Work Type")
				fig_cmpm = _apply_chart_theme(fig_cmpm, "CM vs PM per Location")
				fig_cmpm.update_layout(height=650)
				st.plotly_chart(fig_cmpm, use_container_width=True)
			else:
				st.info("No CM/PM completed records available for Location breakdown.")
		else:
			st.info("Work Type column not found for CM/PM per Location chart.")
	else:
		st.info("Location column not found.")

	# 8) Full-width: Work Orders by Priority (with small-location labels on bars)
	c_priority = next((c for c in ["Priority", "PRIORITY", "priority", "PRIORITY CODE", "PRIORITYCODE"] if c in df.columns), None)
	if c_priority:
		prio_str = df[c_priority].astype(str).str.strip()
		dfp = df.copy()
		dfp["_prio"] = prio_str
		# Build location labels (small) per priority if Location exists
		if c_location:
			def _fmt_locs(series: pd.Series) -> str:
				vals = pd.unique(series.astype(str))
				vals = sorted([v for v in vals if v and v.lower() != "nan"])  # clean
				# Show ALL locations, compact with bullet separators
				return " • ".join(vals)
			loc_text = dfp.groupby("_prio")[c_location].apply(_fmt_locs).rename("locs_text")
			counts = dfp.groupby("_prio").size().reset_index(name="count").merge(loc_text, on="_prio", how="left")
		else:
			counts = dfp.groupby("_prio").size().reset_index(name="count")
			counts["locs_text"] = ""

		# Numeric sort if possible
		counts["prio_num"] = pd.to_numeric(counts["_prio"], errors="coerce")
		counts = counts.sort_values(["prio_num", "_prio"], ascending=[True, True])

		# Friendly labels next to each priority
		prio_name_map = {"1": "Urgent", "2": "High", "3": "Medium", "4": "Low"}
		counts["prio_name"] = counts["_prio"].map(prio_name_map).fillna("")
		counts["prio_label"] = counts.apply(lambda r: f"{r['_prio']} - {r['prio_name']}" if r["prio_name"] else str(r["_prio"]), axis=1)

		# Color map: Priority 1 in red
		unique_prios = counts["_prio"].astype(str).tolist()
		color_map = {p: px.colors.qualitative.Vivid[i % len(px.colors.qualitative.Vivid)] for i, p in enumerate(unique_prios)}
		color_map["1"] = "#d62728"

		fig_prio = px.bar(
			counts,
			x="_prio",
			y="count",
			color="_prio",
			text="locs_text",
			title="Work Orders by Priority",
			labels={"_prio": "Priority", "count": "WO count"},
			color_discrete_map=color_map,
			custom_data=["locs_text", "count"],
		)
		# Place full location list inside the bar, tiny font, vertical for space efficiency
		fig_prio.update_traces(textposition="inside", textangle=90, insidetextfont=dict(size=7, color="white"), cliponaxis=False)
		fig_prio.update_traces(hovertemplate="Priority=%{x}<br>WO count=%{y}<br>Locations=%{customdata[0]}<extra></extra>")
		fig_prio.update_xaxes(categoryorder="array", categoryarray=counts["_prio"].tolist())
		# Replace tick text with friendly labels (e.g., '1 - Urgent')
		fig_prio.update_xaxes(tickmode="array", tickvals=counts["_prio"].tolist(), ticktext=counts["prio_label"].tolist())
		fig_prio.update_layout(legend_title_text="Priority")
		fig_prio.update_layout(hovermode="closest")
		fig_prio = _apply_chart_theme(fig_prio, "Work Orders by Priority")
		fig_prio.update_layout(height=560)
		st.plotly_chart(fig_prio, use_container_width=True)
	else:
		st.info("Priority column not found.")

	# 2) Full-width: Work Orders by Status (bigger donut to show small slices)
	if c_status:
		# Exclude CAN from status pie if present
		s = df[c_status].astype(str).str.upper()
		s = s[s != "CAN"]
		fig = px.pie(s, names=s, title="Work Orders by Status", hole=0.35, color_discrete_sequence=COLOR_SEQ)
		# Show both count and percent for each slice
		fig.update_traces(
			textposition="outside",
			textinfo="label+value+percent",
			texttemplate="%{label}: %{value} (%{percent})",
			textfont_size=16,
			pull=0,
		)
		fig.update_layout(showlegend=True)
		fig = _apply_chart_theme(fig, "Work Orders by Status")
		# Extra top margin so outside labels aren't clipped at the top edge
		fig.update_layout(margin=dict(l=20, r=20, t=180, b=30))
		fig.update_layout(height=760)
		st.plotly_chart(fig, use_container_width=True)
	else:
		st.info("Status column not found.")

	# 3) Full-width: Work Orders Over Time (monthly with all months shown)
	if c_date:
		date_series = pd.to_datetime(df[c_date], errors="coerce")
		valid = date_series.notna()
		if valid.any():
			month_period = date_series.dt.to_period("M")
			cnt = month_period.value_counts().rename_axis("month").reset_index(name="count")
			min_p, max_p = month_period[valid].min(), month_period[valid].max()
			all_months = pd.period_range(min_p, max_p, freq="M")
			full_df = pd.DataFrame({"month": all_months}).merge(cnt, on="month", how="left").fillna({"count": 0})
			full_df["month_start"] = full_df["month"].dt.to_timestamp()
			fig = px.line(full_df, x="month_start", y="count", title="Work Orders Over Time", markers=True,
						  color_discrete_sequence=ORANGE_SEQ)
			fig.update_xaxes(dtick="M1", tickformat="%b %Y")
			fig = _apply_chart_theme(fig, "Work Orders Over Time")
			fig.update_layout(height=500)
			st.plotly_chart(fig, use_container_width=True)
		else:
			st.info("No valid dates found for time series chart.")
	else:
		st.info("Date column not found.")

	# 4) Full-width: Work Orders by Planner (aggregated for readability)
	if c_planner:
		plan_counts = df.groupby(c_planner).size().reset_index(name="count").sort_values("count", ascending=False)
		fig = px.bar(
			plan_counts,
			x=c_planner,
			y="count",
			title="Work Orders by Planner",
			text="count",
			color_discrete_sequence=COLOR_SEQ,
			labels={c_planner: "Planner", "count": "WO count"},
		)
		fig.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
		fig.update_traces(hovertemplate="Planner=%{x}<br>WO count=%{y}<extra></extra>")
		fig.update_xaxes(tickangle=-60, tickfont=dict(size=10), categoryorder="total descending")
		fig = _apply_chart_theme(fig, "Work Orders by Planner")
		fig.update_layout(height=520)
		st.plotly_chart(fig, use_container_width=True)
	else:
		st.info("Planner column not found.")

	# 5) Full-width: Work Orders by Category
	if c_category:
		cat_counts = df.groupby(c_category).size().reset_index(name="count").sort_values("count", ascending=False)
		fig = px.bar(
			cat_counts,
			x=c_category,
			y="count",
			title="Work Orders by Category",
			text="count",
			color_discrete_sequence=COLOR_SEQ,
			labels={c_category: "Category", "count": "WO count"},
		)
		fig.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
		fig.update_traces(hovertemplate="Category=%{x}<br>WO count=%{y}<extra></extra>")
		fig.update_xaxes(tickangle=-30, tickfont=dict(size=12), categoryorder="total descending")
		fig = _apply_chart_theme(fig, "Work Orders by Category")
		fig.update_layout(height=500)
		st.plotly_chart(fig, use_container_width=True)
	else:
		st.info("Work Category column not found.")

	# 6) Full-width: Work Orders by Type
	if c_type:
		type_counts = df.groupby(c_type).size().reset_index(name="count").sort_values("count", ascending=False)
		fig = px.bar(
			type_counts,
			x=c_type,
			y="count",
			title="Work Orders by Type",
			text="count",
			color_discrete_sequence=COLOR_SEQ,
			labels={c_type: "Type", "count": "WO count"},
		)
		fig.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
		fig.update_traces(hovertemplate="Type=%{x}<br>WO count=%{y}<extra></extra>")
		fig.update_xaxes(tickangle=-30, tickfont=dict(size=12), categoryorder="total descending")
		fig = _apply_chart_theme(fig, "Work Orders by Type")
		fig.update_layout(height=480)
		st.plotly_chart(fig, use_container_width=True)
	else:
		st.info("Work Type column not found.")

	# 7) Full-width: On-time vs Delayed per month
	# Build month column from completion date if available else reported
	date_for_month = None
	if "Completion Date" in df.columns:
		date_for_month = pd.to_datetime(df["Completion Date"], errors="coerce")
	elif c_date:
		date_for_month = pd.to_datetime(df[c_date], errors="coerce")

	if date_for_month is not None:
		month = date_for_month.dt.to_period("M").astype(str)
		# On-time/Delayed only when both Target and Completion available
		target = pd.to_datetime(df.get("Target Date"), errors="coerce") if "Target Date" in df.columns else None
		if target is not None and "Completion Date" in df.columns:
			completion = pd.to_datetime(df["Completion Date"], errors="coerce")
			on_time = (completion.notna() & target.notna() & (completion <= target))
			delayed = (completion.notna() & target.notna() & (completion > target))
			dt = pd.DataFrame({"month": month, "On Time": on_time.astype(int), "Delayed": delayed.astype(int)})
			grp = dt.groupby("month").sum().reset_index()
			long = grp.melt(id_vars="month", var_name="Status", value_name="count")
			fig = px.bar(
				long,
				x="month",
				y="count",
				color="Status",
				barmode="stack",
				color_discrete_sequence=["#2ca02c", "#d62728"],
				title="On-Time vs Delayed by Month",
				labels={"month": "Month", "count": "WO count", "Status": "Status"},
				custom_data=["Status"],
			)
			fig.update_traces(hovertemplate="Month=%{x}<br>Status=%{customdata[0]}<br>WO count=%{y}<extra></extra>")
			fig = _apply_chart_theme(fig, "On-Time vs Delayed by Month")
			fig.update_layout(height=520)
			st.plotly_chart(fig, use_container_width=True)
		else:
			st.info("Target/Completion dates not sufficient for on-time analysis.")

		# Completed MoM trend (COMP or CLOSE) with all months shown (including zero-count months)
		if c_status is not None:
			status_u = df[c_status].astype(str).str.upper()
			valid_dates = date_for_month.notna()
			if valid_dates.any():
				is_done = status_u.isin(["COMP", "CLOSE"]) & valid_dates
				# Work with Period[M] to build a full continuous monthly range
				month_period = date_for_month.dt.to_period("M")
				done_months = month_period[is_done]
				cnt = done_months.value_counts().rename_axis("month").reset_index(name="count") if is_done.any() else pd.DataFrame({"month": [], "count": []})

				# Build full month range from min to max available dates
				min_p = month_period[valid_dates].min()
				max_p = month_period[valid_dates].max()
				if pd.isna(min_p) or pd.isna(max_p):
					st.info("No valid dates for monthly completion trend.")
				else:
					all_months = pd.period_range(min_p, max_p, freq="M")
					full_df = pd.DataFrame({"month": all_months}).merge(cnt, on="month", how="left").fillna({"count": 0})
					# Use month start as datetime for clean date axis; force monthly ticks
					full_df["month_start"] = full_df["month"].dt.to_timestamp()
					fig2 = px.line(full_df, x="month_start", y="count", markers=True, color_discrete_sequence=ORANGE_SEQ, title="Completed Orders per Month")
					fig2.update_xaxes(dtick="M1", tickformat="%b %Y")
					fig2 = _apply_chart_theme(fig2, "Completed Orders per Month")
					fig2.update_layout(height=480)
					st.plotly_chart(fig2, use_container_width=True)
			else:
				st.info("No valid dates for monthly completion trend.")
		else:
			st.info("Status column not found for monthly completion trend.")
	else:
		st.info("No suitable date column for monthly trends.")


# -------------------------
# Pages
# -------------------------

# Modern KPI styling and helpers for the Performance KPIs tab
def inject_perf_kpi_css(brand_color: str = "#FFA500") -> None:
		css = f"""
				<style>
					:root {{
						--brand: {brand_color};
						--bg-card: #ffffff;
						--text: #1f2937;
						--muted: #6b7280;
						--good: #16a34a;
						--warn: #f59e0b;
						--bad: #ef4444;
						--shadow: 0 4px 16px rgba(0,0,0,0.08);
						--radius: 14px;
					}}

					.kpi-section {{
						margin: 8px 0 24px 0;
					}}
					.kpi-title {{
						display:flex; align-items:center; gap:12px;
						margin: 4px 0 12px 0;
						font-weight: 700; color: var(--text);
					}}
					.kpi-title .accent {{
						width: 18px; height: 18px; border-radius: 4px; background: var(--brand);
						box-shadow: 0 6px 16px rgba(255,165,0,0.35);
					}}
					.kpi-subtitle {{
						margin-top:-4px; color: var(--muted); font-size: 0.88rem;
					}}

					.kpi-grid {{
						display: grid;
						grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
						gap: 16px;
					}}

					.kpi-card.modern {{
						background: var(--bg-card);
						border-radius: var(--radius);
						box-shadow: var(--shadow);
						padding: 14px 16px;
						transition: transform .15s ease, box-shadow .15s ease;
						border: 1px solid rgba(0,0,0,0.04);
					}}
					.kpi-card.modern:hover {{ transform: translateY(-2px); box-shadow: 0 10px 24px rgba(0,0,0,0.10); }}

					.kpi-head {{
						display:flex; align-items:center; justify-content:space-between; gap:8px; margin-bottom: 6px;
					}}
					.kpi-left {{ display:flex; align-items:center; gap:10px; min-width:0; }}
					.kpi-icon {{
						width: 30px; height: 30px; display:grid; place-items:center;
						border-radius:10px; background: rgba(0,0,0,0.04); font-size: 18px;
					}}
					.kpi-label {{
						font-weight: 600; color: var(--muted); white-space: nowrap; overflow:hidden; text-overflow: ellipsis;
					}}

					.kpi-tip {{
						position: relative; cursor: help; color: var(--muted); font-size: 14px; line-height: 1;
						padding: 4px 8px; border-radius: 8px; background: rgba(0,0,0,0.04);
					}}
					.kpi-tip:hover::after {{
						content: attr(data-tip);
						position: absolute; z-index: 10; top: 120%; right: 0;
						max-width: 260px; background: #111827; color:#fff; padding:10px 12px; border-radius: 10px; font-size:12px;
						box-shadow: 0 8px 24px rgba(0,0,0,0.25);
						white-space: normal;
					}}

					.kpi-value.modern {{
						font-size: 1.8rem; font-weight: 800; color: var(--text);
						letter-spacing: -0.02em;
					}}
					.kpi-sub {{
						margin-top: 4px; color: var(--muted); font-size: 0.86rem;
					}}

					/* color states for values */
					.kpi-card.modern.good .kpi-value.modern {{ color: var(--good); }}
					.kpi-card.modern.warn .kpi-value.modern {{ color: var(--warn); }}
					.kpi-card.modern.bad  .kpi-value.modern {{ color: var(--bad);  }}

					/* tiny top accent using brand color */
					.kpi-card.modern::before {{
						content:""; display:block; height: 4px; border-radius: 12px 12px 0 0; 
						background: linear-gradient(90deg, var(--brand), rgba(255,165,0,0.25));
						margin: -14px -16px 10px -16px;
					}}

					/* Typography smoothing */
					.kpi-value.modern, .kpi-label, .kpi-title {{ -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }}
				</style>
				"""
		st.markdown(dedent(css), unsafe_allow_html=True)

# Ensure HTML starts at column 0 to avoid Markdown code blocks
def _flush_left(s: str) -> str:
	try:
		return "\n".join((line.lstrip() for line in s.splitlines()))
	except Exception:
		return s

def _fmt(v: Any, unit: str = "") -> str:
		if v is None or v == "-" or (isinstance(v, float) and (v != v)):
				return "—"
		try:
				if unit == "%":
						return f"{float(v):.0f}%"
				if unit == "pct1":
						return f"{float(v):.1f}%"
				if unit == "d":
						return f"{float(v):.1f} d"
				if unit == "int":
						return f"{int(v):,}"
				return str(v)
		except Exception:
				return str(v)

def _state(metric_key: str, val: Optional[float]) -> str:
		if val is None:
				return ""
		try:
				v = float(val)
		except Exception:
				return ""

		# Higher is better
		higher_better = {"completion_rate", "on_time_rate", "pm_pct_pc", "first_time_fix", "pm_compliance"}
		# Lower is better
		lower_better = {"backlog_rate", "rework_rate", "overdue_open", "avg_backlog_age"}

		if metric_key in higher_better:
				if v >= 90: return "good"
				if v >= 75: return "warn"
				return "bad"
		if metric_key in lower_better:
				if v <= 10: return "good"
				if v <= 20: return "warn"
				return "bad"
		return ""

def _kpi_html(icon: str, label: str, value: str, tip: Optional[str], state_cls: str) -> str:
	tip_html = f'<span class="kpi-tip" data-tip="{tip}">?</span>' if tip else ""
	# Build without any leading indentation
	return (
		f'<div class="kpi-card modern {state_cls}">'  
		f'<div class="kpi-head">'  
		f'<div class="kpi-left">'  
		f'<div class="kpi-icon">{icon}</div>'  
		f'<div class="kpi-label">{label}</div>'  
		f'</div>'  
		f'{tip_html}'  
		f'</div>'  
		f'<div class="kpi-value modern">{value}</div>'  
		f'</div>'
	)

def render_kpi_section(title: str, items: List[Dict[str, Any]], subtitle: Optional[str] = None) -> None:
    items_html = ''.join([
        _kpi_html(
            icon=i.get('icon','📊'),
            label=i['label'],
            value=_fmt(i.get('value'), i.get('unit','')),
            tip=i.get('tip'),
            state_cls=_state(i.get('key',''), i.get('value'))
        ) for i in items
    ])
    subtitle_html = f'<div class="kpi-subtitle">{subtitle}</div>' if subtitle else ''
    html = (
        '<div class="kpi-section">'
        '<div class="kpi-title"><span class="accent"></span><span>' + str(title) + '</span></div>'
        + subtitle_html +
        '<div class="kpi-grid">' + items_html + '</div>'
        '</div>'
    )
    st.markdown(_flush_left(html), unsafe_allow_html=True)

def page_kpis(df: pd.DataFrame, updated_at: Optional[datetime]) -> None:
	# Compute core metrics
	kpis = compute_kpis(df)
	ins = compute_insights(df)

	# Brand-aware KPI UI
	inject_perf_kpi_css("#FFA500")

	# Map metrics to modern sections
	overall = [
		{"key":"", "label":"Total WOs",   "value": kpis.get("total"),     "unit":"int", "icon":"🧰", "tip":"Total work orders in scope after exclusions."},
		{"key":"", "label":"Open Orders", "value": kpis.get("open"),      "unit":"int", "icon":"📂", "tip":"Currently open work orders."},
		{"key":"", "label":"Completed",   "value": kpis.get("completed"), "unit":"int", "icon":"✅", "tip":"Completed (may include closed)."},
		{"key":"", "label":"Closed",      "value": kpis.get("closed"),    "unit":"int", "icon":"📦", "tip":"Fully closed WOs."},
		{"key":"", "label":"Rejected",    "value": kpis.get("rejected"),  "unit":"int", "icon":"❌", "tip":"Rejected/voided after QC."},
	]
	render_kpi_section(
		"Overall Performance",
		overall,
		subtitle=f"Last Updated: {updated_at.strftime('%Y-%m-%d %H:%M:%S') if updated_at else datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
	)

	completion = [
		{"key":"completion_rate", "label":"Completion Rate", "value": kpis.get("completion_rate"), "unit":"%", "icon":"📈",
		 "tip":"Completed / (Completed + Open). Target ≥ 90%"},
		{"key":"on_time_rate", "label":"On-Time Rate", "value": kpis.get("pct_on_time"), "unit":"%", "icon":"⏱️",
		 "tip":"Completed on or before target date. Target ≥ 90%"},
		{"key":"avg_completion", "label":"Avg Completion", "value": kpis.get("avg_days"), "unit":"d", "icon":"🕰️",
		 "tip":"Average turnaround time from creation to completion."},
		{"key":"rework_rate", "label":"Rework Rate", "value": kpis.get("rework_rate"), "unit":"%", "icon":"🧭",
		 "tip":"Percent of WOs that required rework. Lower is better."},
		{"key":"backlog_rate", "label":"Backlog Rate", "value": kpis.get("backlog_rate"), "unit":"%", "icon":"📥",
		 "tip":"Open / Total. Lower is better; target ≤ 10%"},
	]
	render_kpi_section("Completion Metrics", completion)

	split = [
		{"key":"pm_pct_pc", "label":"PM % (PM+CM)", "value": kpis.get("pm_pct_pc"), "unit":"%", "icon":"🛠️",
		 "tip":"Share of Preventive in PM+CM only."},
		{"key":"cm_pct_pc", "label":"CM % (PM+CM)", "value": kpis.get("cm_pct_pc"), "unit":"%", "icon":"🔧",
		 "tip":"Share of Corrective in PM+CM only."},
		{"key":"", "label":"CM/PM Ratio", "value": kpis.get("cm_pm_ratio"), "unit":"pct1", "icon":"📊",
		 "tip":"Corrective to Preventive ratio (CM / PM)."},
	]
	render_kpi_section("Work Type Split", split)

	# Service Provider sections
	svc_top_line = [
		{"key":"", "label":"Total WOs (Svc)", "value": ins.get("total_wos"), "unit":"int", "icon":"🧰",
		 "tip":"Total service-provider work orders in scope."},
		{"key":"completion_rate", "label":"Completion Rate", "value": ins.get("completion_rate"), "unit":"%", "icon":"📈",
		 "tip":"Provider completion rate; target ≥ 90%."},
		{"key":"on_time_rate", "label":"On-Time (valid)", "value": ins.get("on_time_rate"), "unit":"%", "icon":"⏱️",
		 "tip":"On-time only where target dates exist."},
		{"key":"", "label":"MoM Completed", "value": ins.get("mom_change"), "unit":"pct1", "icon":"📅",
		 "tip":"Month-over-month change in completed count."},
		{"key":"", "label":"Top Category", "value": ins.get("top_category"), "unit":"", "icon":"🏷️",
		 "tip":"Category with most WOs for the period."},
	]
	render_kpi_section("Service Provider Performance", svc_top_line)

	if ins.get("top_planner"):
		st.info(f"Top Planner: {ins.get('top_planner')}", icon="🧑‍💼")

	svc_ops = [
		{"key":"first_time_fix", "label":"First-Time Fix", "value": ins.get("ftfr"), "unit":"%", "icon":"🧩",
		 "tip":"Fixed on first visit; higher is better."},
		{"key":"", "label":"Done (Svc)", "value": ins.get("done_wos"), "unit":"int", "icon":"🧾",
		 "tip":"Completed by service provider."},
		{"key":"overdue_open", "label":"Overdue Open", "value": ins.get("overdue_open_rate"), "unit":"%", "icon":"⏰",
		 "tip":"Open WOs past target date; lower is better."},
		{"key":"avg_backlog_age", "label":"Avg Backlog Age", "value": ins.get("avg_backlog_age"), "unit":"d", "icon":"📦",
		 "tip":"Average age of open backlog in days."},
		{"key":"pm_compliance", "label":"PM Compliance", "value": ins.get("pm_compliance"), "unit":"%", "icon":"✅",
		 "tip":"On-time PMs / scheduled PMs; target ≥ 95%."},
	]
	render_kpi_section("Operational Quality (Svc)", svc_ops)

	# On-time completion rate per Location (Service Level per Location)
	st.markdown("---")
	st.subheader("Service Level per Location (On-Time %)")
	c_location = next((c for c in ["Location", "Site/Location", "Site"] if c in df.columns), None)
	if c_location and ("Target Date" in df.columns) and ("Completion Date" in df.columns):
		tgt = pd.to_datetime(df["Target Date"], errors="coerce")
		comp = pd.to_datetime(df["Completion Date"], errors="coerce")
		valid = comp.notna() & tgt.notna()
		if valid.any():
			dloc = df.loc[valid, [c_location]].copy()
			dloc["on_time"] = (comp[valid] <= tgt[valid]).astype(int)
			grp = dloc.groupby(c_location)["on_time"].agg(["count", "sum"]).reset_index()
			grp["on_time_rate"] = 100 * grp["sum"] / grp["count"].replace(0, 1)
			grp = grp.sort_values("on_time_rate", ascending=False)
			fig_sl = px.bar(
				grp,
				x=c_location,
				y="on_time_rate",
				title="On-Time Completion Rate by Location",
				labels={c_location: "Location", "on_time_rate": "On-Time %"},
				text=grp["on_time_rate"].round(0).astype(int).astype(str) + "%",
				color_discrete_sequence=ORANGE_SEQ,
			)
			fig_sl.update_traces(textposition="outside", cliponaxis=False)
			fig_sl.update_xaxes(tickangle=-45, tickfont=dict(size=10))
			fig_sl = _apply_chart_theme(fig_sl, "On-Time Completion Rate by Location")
			fig_sl.update_layout(height=520)
			st.plotly_chart(fig_sl, use_container_width=True)
		else:
			st.info("No valid Target/Completion dates to compute on-time rate per Location.")
	else:
		st.info("Columns for Location/Target/Completion are not sufficient to compute per-location service level.")


def page_analytics(df: pd.DataFrame) -> None:
	# Display-only analytics
	make_charts(df)


# -------------------------
# Main
# -------------------------

def main() -> None:
	inject_brand_css()
	brand_header()
	# Pure display: no sidebar, no inputs

	# Optional daily auto-refresh (24h) if streamlit_autorefresh is installed
	try:
		from streamlit_autorefresh import st_autorefresh  # type: ignore
		# Refresh every 60 seconds
		st_autorefresh(interval=60_000, key="refresh")
	except Exception:
		pass

	# Top tabs for switching pages (Analytics first)
	tabs = st.tabs(["Analytics & Charts", "Performance KPIs"])

	# Load data
	df: Optional[pd.DataFrame] = None
	updated_at: Optional[datetime] = None
	error = None
	data_url = FIXED_DATA_URL
	if (not data_url) or ("example.com" in str(data_url)):
		error = "No fixed data URL configured."
	else:
		try:
			with st.spinner("👉 Importing data from MAXIMO..."):
				df, updated_at = load_data(data_url)
		except KeyboardInterrupt:
			error = "Data loading interrupted by user."
		except Exception as exc:
			error = str(exc)

	if error:
		st.warning("Using sample data (failed to load fixed data URL). Configure MMC_DATA_URL env var or update FIXED_DATA_URL in code.")
		df = make_sample_data()
		updated_at = datetime.now()

	if df is None:
		return

	# Exclude cancelled and Work Category == 'Contract' from all stats/charts globally
	status_col = next((c for c in ["Status", "STATUS", "status"] if c in df.columns), None)
	if status_col:
		st_u = df[status_col].astype(str).str.upper().str.strip()
		st_u = st_u.replace({"CANCELLED": "CAN", "CANCELED": "CAN", "CANCEL": "CAN"})
		df = df[st_u != "CAN"].copy()
	cat_col = next((c for c in ["Work Category", "Category"] if c in df.columns), None)
	if cat_col:
		df = df[df[cat_col].astype(str).str.upper().str.strip() != "CONTRACT"].copy()

	# Exclude specific location globally from all stats/charts: 'MMC-ARSAL'
	loc_col = next((c for c in ["Location", "Site/Location", "Site"] if c in df.columns), None)
	if loc_col:
		loc_u = df[loc_col].astype(str).str.upper().str.strip()
		df = df[loc_u != "MMC-ARSAL"].copy()

	# Clean & de-duplicate to avoid double counting and keep the most informative/latest rows
	def _deduplicate_rows(dfin: pd.DataFrame) -> pd.DataFrame:
		# Trim strings
		obj_cols = dfin.select_dtypes(include=["object"]).columns
		for c in obj_cols:
			dfin[c] = dfin[c].astype(str).str.strip()
		# Candidate unique ID columns
		id_candidates = [
			"WO", "WO Number", "WONUM", "Work Order", "WORKORDER", "TICKET", "Ticket ID", "SR", "REQUESTID"
		]
		id_col = next((c for c in id_candidates if c in dfin.columns), None)
		if not id_col:
			return dfin.drop_duplicates().reset_index(drop=True)
		# Build a composite sort key: more non-nulls first, then latest timestamp across known date columns
		date_cols = [
			"Completion Date", "Actual Finish", "Target Date", "Reported Date", "Created Date", "Create Date", "Status Date", "Last Modified", "Change Date"
		]
		present = [c for c in date_cols if c in dfin.columns]
		if present:
			ts = pd.DataFrame({c: pd.to_datetime(dfin[c], errors="coerce") for c in present}).max(axis=1)
		else:
			ts = pd.Series([pd.NaT] * len(dfin))
		df2 = dfin.copy()
		df2["__nn"] = dfin.notna().sum(axis=1)
		df2["__ts"] = ts
		df2 = df2.sort_values(["__nn", "__ts"], ascending=[False, False])
		df2 = df2.drop_duplicates(subset=id_col, keep="first").drop(columns=["__nn", "__ts"], errors="ignore")
		return df2.reset_index(drop=True)

	df = _deduplicate_rows(df)

	with tabs[0]:
		page_analytics(df)
	with tabs[1]:
		page_kpis(df, updated_at)


if __name__ == "__main__":
	main()
