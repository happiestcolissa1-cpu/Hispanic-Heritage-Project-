"""
Hispanic Heritage Historical Data Explorer
Streamlit app for World Bank indicator exploration and function analysis.
"""
from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
import math
import re
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import requests
import streamlit as st
import sympy as sp
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle
from xlsxwriter import __version__ as xlsxwriter_version

st.set_page_config(page_title="Hispanic Heritage Historical Data Explorer", page_icon="📈", layout="wide")
st.title("Hispanic Heritage Historical Data Explorer")
st.caption("Hispanic Heritage Data Lab • authentic World Bank indicators • student-friendly modeling")
st.info("Brazil is included as a Latin American comparison country; it is not a Spanish-speaking country.")

COUNTRIES = {
    "Mexico": "MEX", "Brazil": "BRA", "Argentina": "ARG", "Colombia": "COL",
    "Chile": "CHL", "Peru": "PER", "Venezuela": "VEN", "Ecuador": "ECU", "Uruguay": "URY"
}
INDICATORS = {
    "Life expectancy at birth (years)": ("SP.DYN.LE00.IN", "Years", "Average number of years a newborn is expected to live if current age-specific mortality rates persist."),
    "Unemployment, national estimate (% of labor force)": ("SL.UEM.TOTL.NE.ZS", "% of labor force", "Unemployment as a share of the labor force, using national estimates."),
    "GDP per capita (current US$)": ("NY.GDP.PCAP.CD", "Current US dollars per person", "Gross domestic product divided by midyear population. This is not average wealth or household income."),
    "Crude birth rate (per 1,000 people)": ("SP.DYN.CBRT.IN", "Births per 1,000 people per year", "Annual live births per 1,000 people; this is a rate, not a birth count."),
    "Net migration": ("SM.POP.NETM", "People (net estimate)", "Net number of migrants during the period; it is not an emigration count."),
    "Intentional homicide rate (per 100,000 people)": ("VC.IHR.PSRC.P5", "Deaths per 100,000 people", "Intentional homicide victims per 100,000 people."),
    "Consumer price inflation (annual %)": ("FP.CPI.TOTL.ZG", "% per year", "Annual percentage change in the consumer price index."),
    "Central government debt (% of GDP)": ("GC.DOD.TOTL.GD.ZS", "% of GDP", "Central government debt as a percentage of GDP."),
    "Total reserves including gold (current US$)": ("FI.RES.TOTL.CD", "Current US dollars", "Total reserves including gold, valued in current US dollars."),
    "External balance on goods and services (% of GDP)": ("NE.RSB.GNFS.ZS", "% of GDP", "External balance on goods and services as a share of GDP.")
}
WB_API = "https://api.worldbank.org/v2/country/{country}/indicator/{indicator}"
WB_SOURCE = "https://data.worldbank.org/indicator/"
TODAY_YEAR = date.today().year
START_YEAR, END_YEAR = TODAY_YEAR - 70, TODAY_YEAR - 1

CONTEXTS = {
    "Business": ("Revenue", "thousand US dollars", "Time", "years", "A revenue-like measure; interpret derivatives as change per year."),
    "Engineering": ("Displacement", "meters", "Time", "seconds", "For position, the first derivative is velocity and the second is acceleration."),
    "Medical": ("Concentration", "mg/L", "Time", "hours", "A concentration-like measure; derivatives describe how concentration changes."),
    "Sports science": ("Performance", "score units", "Time", "seconds", "A performance measure changing over time."),
    "Biology": ("Population", "organisms", "Time", "days", "A population-like measure; derivatives describe population growth."),
    "Bioengineering": ("Signal", "signal units", "Time", "seconds", "A measured signal and its rates of change."),
    "Economic forecasting": ("Indicator", "indicator units", "Time", "years", "A trend model; extrapolation is uncertain and not a causal explanation."),
    "Weather": ("Temperature", "°C", "Time", "hours", "Temperature change and its rate of change."),
    "Custom": ("Output", "output units", "Input", "input units", "Define what the variables mean in your own scenario.")
}
SUGGESTED = {
    "Business": "a*x**2 + b*x + c", "Engineering": "a*x**2 + b*x + c",
    "Medical": "a*exp(-k*x)", "Sports science": "a*sin(k*x) + b",
    "Biology": "a*exp(k*x)", "Bioengineering": "a*sin(k*x) + b",
    "Economic forecasting": "a*x**2 + b*x + c", "Weather": "a*sin(k*x) + b",
    "Custom": "a*x + b"
}

def wb_fetch(country_code: str, indicator_code: str, start: int, end: int) -> tuple[pd.DataFrame, str | None]:
    url = WB_API.format(country=country_code, indicator=indicator_code)
    try:
        response = requests.get(url, params={"format": "json", "date": f"{start}:{end}", "per_page": 200}, timeout=18)
        response.raise_for_status()
        payload = response.json()
        if not isinstance(payload, list) or len(payload) < 2 or not payload[1]:
            return pd.DataFrame(columns=["year", "value"]), "The World Bank API returned no observations for this selection."
        rows = [{"year": int(item["date"]), "value": item.get("value")} for item in payload[1] if item.get("date")]
        frame = pd.DataFrame(rows)
        frame["value"] = pd.to_numeric(frame["value"], errors="coerce")
        frame = frame.dropna(subset=["value"]).sort_values("year").reset_index(drop=True)
        if frame.empty:
            return frame, "The API response contained no numeric observations for this period."
        return frame, None
    except requests.Timeout:
        return pd.DataFrame(columns=["year", "value"]), "Request timed out. Try again later or check your connection."
    except (requests.RequestException, ValueError, TypeError, KeyError) as exc:
        return pd.DataFrame(columns=["year", "value"]), f"Could not retrieve data: {exc.__class__.__name__}. Check connection/API status."

def validate_data(frame: pd.DataFrame, indicator_name: str) -> tuple[pd.DataFrame, list[str]]:
    errors = []
    result = frame.copy()
    if not {"year", "value"}.issubset(result.columns):
        return pd.DataFrame(columns=["year", "value"]), ["Table must have year and value columns."]
    result["year"] = pd.to_numeric(result["year"], errors="coerce")
    result["value"] = pd.to_numeric(result["value"], errors="coerce")
    if result["year"].isna().any() or not np.all(np.isclose(result["year"].dropna() % 1, 0)):
        errors.append("Years must be whole-number integers.")
    if result["value"].isna().any():
        errors.append("Values must be numeric and non-empty.")
    result = result.dropna(subset=["year", "value"]).copy()
    result["year"] = result["year"].astype(int)
    if result["year"].duplicated().any():
        errors.append("Each year must be unique. Remove duplicate years before fitting.")
    if ((result["year"] < 1800) | (result["year"] > 2200)).any():
        errors.append("Years must be between 1800 and 2200.")
    if indicator_name.startswith("Life expectancy") and ((result.value < 0) | (result.value > 120)).any():
        errors.append("Life expectancy values must be between 0 and 120 years.")
    if indicator_name.startswith("Unemployment") and ((result.value < 0) | (result.value > 100)).any():
        errors.append("Unemployment percentages must be between 0 and 100.")
    if indicator_name.startswith("Crude birth rate") and (result.value < 0).any():
        errors.append("Crude birth rate cannot be negative.")
    return result.sort_values("year").reset_index(drop=True), errors

def fit_polynomial(frame: pd.DataFrame, degree: int):
    if frame["year"].nunique() < degree + 1:
        raise ValueError(f"Degree {degree} needs at least {degree + 1} distinct years; current data has {frame['year'].nunique()}.")
    base = int(frame["year"].min())
    t = frame["year"].to_numpy(dtype=float) - base
    y = frame["value"].to_numpy(dtype=float)
    coeff = np.polyfit(t, y, degree)
    pred = np.polyval(coeff, t)
    ss_res = float(np.sum((y - pred) ** 2))
    ss_tot = float(np.sum((y - np.mean(y)) ** 2))
    r2 = 1.0 - ss_res / ss_tot if ss_tot else float("nan")
    mae = float(np.mean(np.abs(y - pred)))
    residuals = y - pred
    return {"base": base, "degree": degree, "coeff": coeff, "pred": pred, "r2": r2, "mae": mae,
            "residuals": residuals, "mean": float(np.mean(y)), "n": len(y),
            "min_year": int(frame.year.min()), "max_year": int(frame.year.max())}

def poly_equation(model, unit=""):
    coeff = model["coeff"]
    pieces = []
    d = model["degree"]
    for i, c in enumerate(coeff):
        power = d - i
        if abs(c) < 1e-12: continue
        term = f"{c:.8g}" if power == 0 else (f"{c:.8g}·t" if power == 1 else f"{c:.8g}·t^{power}")
        pieces.append(term)
    return "f(t) = " + (" + ".join(pieces).replace("+ -", "- ") if pieces else "0") + f", where t = year − {model['base']}"

def predict(model, years):
    return np.polyval(model["coeff"], np.asarray(years, dtype=float) - model["base"])

def safe_sympy_function(expression: str):
    if len(expression) > 250 or not re.fullmatch(r"[0-9a-zA-Z_+\-*/().,\s^]*", expression):
        raise ValueError("Use ordinary math characters, numbers, x, and supported functions only.")
    expr = sp.sympify(expression.replace("^", "**"), locals={"x": sp.Symbol("x"), "sin": sp.sin, "cos": sp.cos,
        "tan": sp.tan, "exp": sp.exp, "log": sp.log, "sqrt": sp.sqrt, "pi": sp.pi, "E": sp.E})
    x = sp.Symbol("x")
    if expr.free_symbols - {x}:
        raise ValueError("Use x as the independent variable; replace other symbols with numeric constants.")
    if expr.has(sp.zoo, sp.nan, sp.oo, -sp.oo):
        raise ValueError("Function contains an undefined or infinite expression.")
    f = sp.lambdify(x, expr, modules=["numpy"])
    d1, d2 = sp.diff(expr, x), sp.diff(expr, x, 2)
    return x, expr, d1, d2, f, sp.lambdify(x, d1, modules=["numpy"]), sp.lambdify(x, d2, modules=["numpy"])

def scalar(fn, x):
    try:
        v = np.asarray(fn(float(x))).item()
        return float(v) if np.isfinite(v) else float("nan")
    except Exception:
        return float("nan")

def find_roots(fn, lo, hi, samples=1800):
    grid = np.linspace(lo, hi, samples)
    vals = np.array([scalar(fn, z) for z in grid])
    roots = []
    for i in range(len(grid)-1):
        if not (np.isfinite(vals[i]) and np.isfinite(vals[i+1])): continue
        if abs(vals[i]) < 1e-5: roots.append(grid[i])
        elif vals[i] * vals[i+1] < 0:
            a, b = grid[i], grid[i+1]
            for _ in range(35):
                m = (a+b)/2
                fm = scalar(fn,m)
                if not np.isfinite(fm): break
                if scalar(fn,a)*fm <= 0: b=m
                else: a=m
            roots.append((a+b)/2)
    out=[]
    for r in roots:
        if not out or abs(r-out[-1]) > (hi-lo)/samples*3: out.append(float(r))
    return out

def make_chart(years, values, fitted_years=None, fitted_values=None, estimates=None, boundary=None, x_label="Year", y_label="Value", title="Model"):
    fig, ax = plt.subplots(figsize=(10, 4.8))
    ax.scatter(years, values, label="Observed data", s=26, zorder=3)
    if fitted_years is not None:
        hist = np.asarray(fitted_years) <= (boundary if boundary is not None else max(years))
        ax.plot(np.asarray(fitted_years)[hist], np.asarray(fitted_values)[hist], label="Fitted curve (historical range)", linewidth=2)
        if (~hist).any():
            ax.plot(np.asarray(fitted_years)[~hist], np.asarray(fitted_values)[~hist], "--", label="Extrapolated segment", linewidth=2)
        if boundary is not None: ax.axvline(boundary, linestyle=":", label="Historical boundary")
    if estimates:
        ax.scatter([e[0] for e in estimates], [e[1] for e in estimates], marker="x", s=80, label="Calculated estimate")
    ax.set(xlabel=x_label, ylabel=y_label, title=title)
    ax.grid(True, alpha=.25); ax.legend(loc="best"); fig.tight_layout()
    return fig

def excel_bytes(summary: dict, data: pd.DataFrame, calculated: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame([summary]).to_excel(writer, index=False, sheet_name="Summary")
        data.to_excel(writer, index=False, sheet_name="Working data")
        calculated.to_excel(writer, index=False, sheet_name="Calculated years")
        for ws in writer.sheets.values():
            ws.freeze_panes(1, 0); ws.set_column(0, 12, 22)
    return output.getvalue()

def pdf_bytes(summary_text: str, chart_png: bytes, data: pd.DataFrame, calculated: pd.DataFrame):
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(letter), rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    story = [Paragraph("Hispanic Heritage Historical Data Explorer", styles["Title"]),
             Paragraph(summary_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>"), styles["BodyText"]), Spacer(1, 10)]
    story.append(RLImage(BytesIO(chart_png), width=650, height=300))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Working data preview", styles["Heading2"]))
    preview = data.head(15).astype(str).values.tolist()
    if preview:
        tbl = Table([list(data.columns)] + preview, repeatRows=1)
        tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#e8eef7")),("GRID",(0,0),(-1,-1),.4,colors.grey),("FONTSIZE",(0,0),(-1,-1),8)]))
        story.append(tbl)
    if not calculated.empty:
        story.append(Spacer(1,8)); story.append(Paragraph("Calculated-year results", styles["Heading2"]))
        tbl2 = Table([list(calculated.columns)] + calculated.astype(str).values.tolist()[:20], repeatRows=1)
        tbl2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#e8eef7")),("GRID",(0,0),(-1,-1),.4,colors.grey),("FONTSIZE",(0,0),(-1,-1),8)]))
        story.append(tbl2)
    doc.build(story)
    return output.getvalue()

with st.sidebar:
    st.header("Historical data")
    country_name = st.selectbox("Country", list(COUNTRIES))
    indicator_name = st.selectbox("Indicator", list(INDICATORS))
    degree = st.slider("Polynomial degree", 3, 8, 3)
    sample_increment = st.slider("Model marker increment (years)", 1, 10, 3)
    horizon = st.slider("Forward display horizon (years)", 0, 50, 10)
    compare_country = st.selectbox("Comparison country", ["None"] + [c for c in COUNTRIES if c != country_name])
    printer_mode = st.checkbox("Printer-friendly report view", False)
    course = st.selectbox("Mathematics pathway", ["Honors Algebra 2", "Precalculus", "Calculus"])
    st.divider()
    st.caption(f"Requested window: {START_YEAR}–{END_YEAR} (current year − 70 through current year − 1).")

indicator_code, units, definition = INDICATORS[indicator_name]
country_code = COUNTRIES[country_name]
cache_key = f"source_{country_code}_{indicator_code}_{START_YEAR}_{END_YEAR}"
if cache_key not in st.session_state:
    with st.spinner("Retrieving official World Bank observations…"):
        fetched, fetch_error = wb_fetch(country_code, indicator_code, START_YEAR, END_YEAR)
    st.session_state[cache_key] = fetched
    st.session_state[cache_key+"_error"] = fetch_error
source_data = st.session_state[cache_key].copy()
fetch_error = st.session_state.get(cache_key+"_error")
if fetch_error:
    st.error(fetch_error)
st.markdown(f"**Definition:** {definition}  \n**Units:** {units}  \n**Source:** [World Bank indicator page]({WB_SOURCE}{indicator_code}) · [API endpoint]({WB_API.format(country=country_code, indicator=indicator_code)})  \n**Retrieved:** {date.today().isoformat()}")
st.caption(f"Available observations: {len(source_data)}. Actual coverage: {source_data.year.min() if not source_data.empty else 'none'}–{source_data.year.max() if not source_data.empty else 'none'}. Missing years are not filled or fabricated.")

if "working_data" not in st.session_state or st.session_state.get("working_key") != cache_key:
    st.session_state.working_data = source_data.copy()
    st.session_state.working_key = cache_key
    st.session_state.edited = False
st.subheader("Validated, editable working data")
st.write("Edit years and values below. Source data remains unchanged. Use **Reset to source** to discard edits.")
col1, col2 = st.columns([1, 5])
with col1:
    if st.button("Reset to source"):
        st.session_state.working_data = source_data.copy()
        st.session_state.edited = False
        st.rerun()
with col2:
    st.write("Status: **USER-MODIFIED SCENARIO**" if st.session_state.edited else "**Source-derived data**")
edited_input = st.data_editor(st.session_state.working_data, num_rows="dynamic", use_container_width=True, key=f"editor_{cache_key}")
validated, validation_errors = validate_data(edited_input, indicator_name)
if not edited_input.equals(st.session_state.working_data):
    st.session_state.edited = not edited_input.reset_index(drop=True).equals(source_data.reset_index(drop=True))
st.session_state.working_data = edited_input.copy()
for err in validation_errors: st.error(err)

if len(validated) < degree + 1:
    st.warning(f"Cannot fit degree {degree}: need at least {degree+1} valid distinct years; currently {validated.year.nunique()}. Lower the degree or repair the table.")
    model = None
else:
    try: model = fit_polynomial(validated, degree)
    except (ValueError, np.linalg.LinAlgError) as exc:
        st.error(str(exc)); model = None

if model:
    st.subheader("Polynomial model")
    st.latex(poly_equation(model))
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Observations", model["n"]); m2.metric("In-sample R²", f"{model['r2']:.5f}")
    m3.metric("In-sample MAE", f"{model['mae']:.5g} {units}"); m4.metric("Observed mean", f"{model['mean']:.5g} {units}")
    st.caption(f"Centered time t = year − {model['base']}. Historical fitted range: {model['min_year']}–{model['max_year']}. Metrics are in-sample, not independent validation.")
    residual_table = validated[["year"]].copy()
    residual_table["observed"] = validated.value.to_numpy()
    residual_table["fitted"] = model["pred"]
    residual_table["residual"] = model["residuals"]
    with st.expander("Residuals and chronological holdout"):
        st.dataframe(residual_table, use_container_width=True)
        if len(validated) >= max(2*(degree+1), 12):
            split = max(degree+1, int(len(validated)*.8))
            train = validated.iloc[:split]
            test = validated.iloc[split:]
            if len(test):
                try:
                    hm = fit_polynomial(train, min(degree, train.year.nunique()-1))
                    hp = predict(hm, test.year.to_numpy())
                    st.write(f"Chronological holdout MAE (last {len(test)} observations): {np.mean(np.abs(test.value.to_numpy()-hp)):.5g} {units}. Separate from in-sample metrics.")
                except Exception as exc: st.write(f"Holdout score unavailable: {exc}")
        else: st.write("Not enough observations for the app's chronological holdout threshold.")

    st.subheader("Calculate fitted value by year")
    single_year = st.number_input("Single year (1950–2100)", min_value=1950, max_value=2100, value=min(max(END_YEAR,1950),2100), step=1)
    if st.button("Calculate Value"):
        v = float(predict(model, [single_year])[0])
        st.session_state.single_result = {"year": int(single_year), "value": v, "key": (country_code,indicator_code,degree,tuple(validated.year),tuple(np.round(validated.value,8)))}
    current_key = (country_code,indicator_code,degree,tuple(validated.year),tuple(np.round(validated.value,8)))
    if st.session_state.get("single_result",{}).get("key") == current_key:
        r = st.session_state.single_result
        status = "within observed fitted range" if model["min_year"] <= r["year"] <= model["max_year"] else "OUTSIDE observed range (extrapolation)"
        st.success(f"{r['year']}: {r['value']:.8g} {units} — {status}. This is a fitted model estimate, not necessarily an actual observation.")
        if abs(r["value"]) > max(1e6, abs(model["mean"])*1000):
            st.warning("This estimate is unusually large relative to the observed scale. Polynomial extrapolation can diverge rapidly; it has not been clipped.")
    st.markdown("**Batch years**")
    batch_mode = st.radio("Batch input type", ["Year range", "Comma-separated custom years"], horizontal=True)
    batch_years = []
    invalid_batch = []
    if batch_mode == "Year range":
        a,b = st.columns(2)
        with a: first_year = st.number_input("Start year", min_value=1950, max_value=2100, value=max(1950,model["min_year"]), step=1)
        with b: last_year = st.number_input("End year", min_value=1950, max_value=2100, value=min(2100,model["max_year"]), step=1)
        step_year = st.slider("Range step (years)",1,10,1)
        if first_year <= last_year: batch_years = list(range(int(first_year),int(last_year)+1,int(step_year)))
        else: invalid_batch.append("Start year must be less than or equal to end year.")
    else:
        custom = st.text_input("Years, separated by commas", f"{model['min_year']},{model['max_year']}")
        for item in custom.split(","):
            try:
                n=float(item.strip())
                if not n.is_integer() or n<1950 or n>2100: invalid_batch.append(f"Invalid year: {item.strip()} (use integer years 1950–2100).")
                else: batch_years.append(int(n))
            except ValueError: invalid_batch.append(f"Invalid year: {item.strip()} (not a number).")
        if len(batch_years) != len(set(batch_years)): invalid_batch.append("Duplicate years in custom input.")
    for err in invalid_batch: st.warning(err)
    batch_years = sorted(set(batch_years))[:500]
    calc = pd.DataFrame(columns=["year","estimate","units","range_status"])
    if batch_years:
        vals = predict(model,batch_years)
        calc = pd.DataFrame({"year":batch_years,"estimate":vals,"units":units,
            "range_status":["within observed range" if model["min_year"]<=y<=model["max_year"] else "outside observed range (extrapolation)" for y in batch_years]})
        st.dataframe(calc, use_container_width=True)
    st.subheader("Historical fit and estimates")
    plot_start = min(model["min_year"], min(batch_years) if batch_years else model["min_year"])
    plot_end = max(model["max_year"], max(batch_years) if batch_years else model["max_year"]) + int(horizon)
    plot_end = min(plot_end, 2200)
    curve_years = np.linspace(plot_start, plot_end, max(250, (plot_end-plot_start+1)*3))
    curve_values = predict(model, curve_years)
    sampled = validated.iloc[::sample_increment]
    fig = make_chart(validated.year, validated.value, curve_years, curve_values,
        [(int(y),float(v)) for y,v in zip(batch_years,vals)] if batch_years else None,
        boundary=model["max_year"], y_label=units, title=f"{country_name}: {indicator_name}")
    ax = fig.axes[0]
    ax.scatter(sampled.year, sampled.value, marker="s", s=35, facecolors="none", label=f"Sample markers every {sample_increment} row(s)")
    ax.legend(); st.pyplot(fig, use_container_width=True)
    chart_buffer=BytesIO(); fig.savefig(chart_buffer,format="png",dpi=140,bbox_inches="tight"); chart_png=chart_buffer.getvalue()
    plt.close(fig)
    if compare_country != "None":
        with st.spinner(f"Loading comparison data for {compare_country}…"):
            comp, comp_err = wb_fetch(COUNTRIES[compare_country],indicator_code,START_YEAR,END_YEAR)
        if comp_err: st.warning(f"Comparison data unavailable: {comp_err}")
        else:
            st.write(f"Comparison coverage: {comp.year.min()}–{comp.year.max()} ({len(comp)} observations). Different coverage and missing years can make comparisons incomplete.")
            fig2, ax2 = plt.subplots(figsize=(10,4))
            ax2.plot(validated.year,validated.value,marker="o",label=country_name)
            ax2.plot(comp.year,comp.value,marker="o",label=compare_country)
            ax2.set(xlabel="Year",ylabel=units,title=f"{indicator_name}: country comparison"); ax2.grid(True,alpha=.25); ax2.legend(); fig2.tight_layout()
            st.pyplot(fig2,use_container_width=True); plt.close(fig2)
    st.subheader("Mathematical interpretation")
    tlo, thi = model["min_year"]-model["base"], model["max_year"]-model["base"]
    d1coeff=np.polyder(model["coeff"]); d2coeff=np.polyder(d1coeff)
    if course == "Honors Algebra 2":
        st.write(f"**Domain used for historical fit:** {model['min_year']}–{model['max_year']}; model can be evaluated beyond this, but that is extrapolation. **Range on observed years:** {validated.value.min():.5g} to {validated.value.max():.5g} {units} (observed values, not the full polynomial range).")
        st.write("**Graphical trend:** inspect whether the fitted curve follows the observed points. **End behavior:** a polynomial's ends are governed by its leading term; high-degree models can grow rapidly. **Average rate:** change in fitted value divided by elapsed years.")
        if len(validated)>1: st.write(f"Average observed rate over the data window: {(validated.value.iloc[-1]-validated.value.iloc[0])/(validated.year.iloc[-1]-validated.year.iloc[0]):.5g} {units} per year.")
    elif course == "Precalculus":
        st.write("Degree controls possible bends/turning behavior; coefficient sizes and signs affect scale, direction, and end behavior. Concavity describes whether average rates tend to increase or decrease. Polynomial degree alone does not guarantee a turning point.")
        st.write(f"Numerically estimated turning-point candidates in observed window: {[round(model['base']+x,3) for x in find_roots(lambda x: np.polyval(d1coeff,x),tlo,thi)]}")
        st.write(f"Numerically estimated concavity-change candidates: {[round(model['base']+x,3) for x in find_roots(lambda x: np.polyval(d2coeff,x),tlo,thi)]}. Candidates require checking the behavior on both sides.")
    else:
        st.write("For t = year − base year, the first derivative is change in indicator units per year; the second derivative is change in that rate per year. Numerical candidates below are approximations.")
        d1 = np.polyder(model["coeff"]); d2 = np.polyder(d1)
        roots1=find_roots(lambda x:np.polyval(d1,x),tlo,thi)
        roots2=find_roots(lambda x:np.polyval(d2,x),tlo,thi)
        st.write(f"**First derivative:** `{np.array2string(d1, precision=6)}` (coefficients in descending powers of centered t).")
        st.write(f"**Second derivative:** `{np.array2string(d2, precision=6)}`.")
        st.write(f"**Critical candidates:** {[round(model['base']+x,3) for x in roots1]}; classify by checking derivative sign changes.")
        st.write(f"**Inflection candidates:** {[round(model['base']+x,3) for x in roots2]}; retain only where concavity changes.")
        st.write(f"**Absolute extrema on the closed observed interval:** compare fitted values at endpoints and critical candidates: {model['min_year']}–{model['max_year']}.")
    st.warning("Uncertainty: residuals and MAE describe in-sample fit error, not a validated 95% prediction interval. Polynomial fits are descriptive, sensitive to degree, and do not establish causes. Predictions outside observed years are especially uncertain.")
else:
    calc = pd.DataFrame(columns=["year","estimate","units","range_status"])
    chart_png = b""

st.divider()
st.subheader("Context and function lab")
context = st.selectbox("Context", list(CONTEXTS))
default_y, default_yunit, default_x, default_xunit, context_desc = CONTEXTS[context]
context_description = st.text_input("Context description", context_desc)
x_name = st.text_input("Independent variable name", default_x)
x_unit = st.text_input("Independent variable units", default_xunit)
y_name = st.text_input("Dependent variable name", default_y)
y_unit = st.text_input("Dependent variable units", default_yunit)
function_mode = st.radio("Function source", ["World Bank curve of best fit", "Suggested function", "Custom mathematical function"], horizontal=True)
if function_mode == "World Bank curve of best fit":
    expr_text = None
elif function_mode == "Suggested function":
    expr_text = st.text_input("Suggested function template", SUGGESTED[context])
    p1, p2, p3, p4 = st.columns(4)
    with p1: param_a = st.number_input("Parameter a", value=1.0, key=f"pa_{context}")
    with p2: param_b = st.number_input("Parameter b", value=0.0, key=f"pb_{context}")
    with p3: param_c = st.number_input("Parameter c", value=0.0, key=f"pc_{context}")
    with p4: param_k = st.number_input("Parameter k", value=0.2, key=f"pk_{context}")
    for symbol, value in [("a", param_a), ("b", param_b), ("c", param_c), ("k", param_k)]:
        expr_text = re.sub(rf"\b{symbol}\b", f"({value})", expr_text)
else:
    expr_text = st.text_input("Enter function using x, e.g. 2*x**2 - 3*x + 1", "2*x**2 - 3*x + 1")
if function_mode == "World Bank curve of best fit" and model:
    expr_text = "polynomial fit (centered on year)"
    # for derivative analysis use a symbolic centered polynomial translated into x
    x_symbol = sp.Symbol("x")
    expr = sum(sp.Float(float(c))*(x_symbol-model["base"])**(model["degree"]-i) for i,c in enumerate(model["coeff"]))
    d1expr, d2expr = sp.diff(expr,x_symbol), sp.diff(expr,x_symbol,2)
    f = sp.lambdify(x_symbol,expr,"numpy"); d1f=sp.lambdify(x_symbol,d1expr,"numpy"); d2f=sp.lambdify(x_symbol,d2expr,"numpy")
elif function_mode != "World Bank curve of best fit":
    try:
        x_symbol, expr, d1expr, d2expr, f, d1f, d2f = safe_sympy_function(expr_text)
    except Exception as exc:
        st.error(f"Function input error: {exc}")
        expr = d1expr = d2expr = f = d1f = d2f = None
else:
    expr = d1expr = d2expr = f = d1f = d2f = None

if f is not None:
    st.latex(sp.latex(expr))
    if course == "Calculus":
        st.markdown(f"**First derivative:** ${sp.latex(d1expr)}$  \n**Second derivative:** ${sp.latex(d2expr)}$")
    default_lo, default_hi = (float(model["min_year"]), float(model["max_year"])) if function_mode == "World Bank curve of best fit" and model else (0.0,10.0)
    left,right = st.columns(2)
    with left:
        xmin = st.number_input("Analysis interval start", value=float(default_lo), key=f"xmin_{context}_{function_mode}")
    with right:
        xmax = st.number_input("Analysis interval end", value=float(default_hi), key=f"xmax_{context}_{function_mode}")
    if xmax <= xmin: st.error("Analysis interval end must be greater than its start.")
    else:
        point = st.slider("Current input x", min_value=float(xmin), max_value=float(xmax), value=float((xmin+xmax)/2), key=f"point_{context}_{function_mode}")
        c1,c2 = st.columns(2)
        with c1: aroc_start = st.slider("Average-rate interval start",float(xmin),float(xmax),float(xmin),key=f"aroc_a_{context}_{function_mode}")
        with c2: aroc_end = st.slider("Average-rate interval end",float(xmin),float(xmax),float(xmax),key=f"aroc_b_{context}_{function_mode}")
        show_f=st.checkbox("Show function",True); show_d1=st.checkbox("Show first derivative",course=="Calculus"); show_d2=st.checkbox("Show second derivative",False)
        v=scalar(f,point); slope=scalar(d1f,point); accel=scalar(d2f,point)
        st.write(f"At x = {point:.5g}: f(x) = {v:.7g} {y_unit}; f′(x) = {slope:.7g} {y_unit}/{x_unit}; f″(x) = {accel:.7g} {y_unit}/{x_unit}².")
        if abs(aroc_end-aroc_start)>1e-12:
            va,vb=scalar(f,aroc_start),scalar(f,aroc_end)
            ar=(vb-va)/(aroc_end-aroc_start) if np.isfinite(va) and np.isfinite(vb) else float("nan")
            st.write(f"Average rate of change on [{aroc_start:.5g}, {aroc_end:.5g}]: {ar:.7g} {y_unit}/{x_unit}.")
        else: st.warning("Choose different interval endpoints to calculate average rate of change.")
        st.write(f"**Context interpretation:** {context_description} At this point, the output is {v:.5g} {y_unit}; the instantaneous rate is {slope:.5g} {y_unit} per {x_unit}; the second derivative describes how that rate is changing.")
        if course == "Calculus":
            st.write("**Tangent line:** y = f(x₀) + f′(x₀)(x − x₀). **Secant line:** connects the two interval endpoints when both function values are defined.")
            crit=find_roots(d1f,xmin,xmax); infl=find_roots(d2f,xmin,xmax)
            classified=[]
            for r in crit:
                eps=max((xmax-xmin)*1e-4,1e-5); before=scalar(d1f,r-eps); after=scalar(d1f,r+eps)
                kind="local maximum" if before>0 and after<0 else "local minimum" if before<0 and after>0 else "stationary candidate (not classified as extremum)"
                classified.append((round(r,5),kind))
            st.write(f"Critical points (numerical candidates): {classified}")
            inflections=[]
            for r in infl:
                eps=max((xmax-xmin)*1e-4,1e-5); before=scalar(d2f,r-eps); after=scalar(d2f,r+eps)
                if np.isfinite(before) and np.isfinite(after) and before*after<0: inflections.append(round(r,5))
            st.write(f"Inflection points where sampled concavity changes (numerical): {inflections}")
            st.write(f"Closed-interval absolute extrema candidates include endpoints and critical points: x = {xmin}, {xmax}, { [round(z,5) for z,_ in classified] }.")
        grid=np.linspace(xmin,xmax,900)
        fig3,ax3=plt.subplots(figsize=(10,4.5))
        for fn,label,enabled in [(f,f"{y_name} f(x)",show_f),(d1f,"First derivative",show_d1),(d2f,"Second derivative",show_d2)]:
            if enabled:
                ys=np.array([scalar(fn,z) for z in grid])
                # Break line across undefined regions; NaNs naturally split matplotlib lines.
                ys[~np.isfinite(ys)] = np.nan
                ax3.plot(grid,ys,label=label)
        if show_f and np.isfinite(v):
            ax3.scatter([point],[v],s=65,label="Current point")
            if np.isfinite(slope):
                ax3.plot(grid, v+slope*(grid-point), "--", label="Tangent line")
        if abs(aroc_end-aroc_start)>1e-12 and np.isfinite(va) and np.isfinite(vb):
            sec_slope=(vb-va)/(aroc_end-aroc_start)
            ax3.plot(grid,va+sec_slope*(grid-aroc_start),":",label="Secant line")
        ax3.axvline(aroc_start,alpha=.25); ax3.axvline(aroc_end,alpha=.25)
        ax3.set(xlabel=f"{x_name} ({x_unit})",ylabel=f"{y_name} / derivative units",title=f"{context}: function analysis")
        ax3.grid(True,alpha=.25); ax3.legend(); fig3.tight_layout(); st.pyplot(fig3,use_container_width=True); plt.close(fig3)
        with st.expander("Playback controls (manual stepping)"):
            st.caption("Use the input slider to move the point; playback is represented by step controls so it remains stable on hosted Streamlit.")
            speed=st.select_slider("Step size",options=[0.1,0.25,0.5,1,2,5],value=1.0)
            pcol1,pcol2=st.columns(2)
            with pcol1:
                if st.button("Step forward"): st.session_state[f"point_{context}_{function_mode}"]=min(xmax,point+speed); st.rerun()
            with pcol2:
                if st.button("Reset point"): st.session_state[f"point_{context}_{function_mode}"]=(xmin+xmax)/2; st.rerun()
        st.caption("Domain/discontinuity note: numerical plotting skips non-finite values, but automatic domain detection is not complete for every arbitrary expression. Check rational denominators and logarithm/square-root restrictions before interpreting results.")

summary = {
    "app": "Hispanic Heritage Historical Data Explorer", "country": country_name, "country_code": country_code,
    "indicator": indicator_name, "indicator_code": indicator_code, "units": units, "retrieved": date.today().isoformat(),
    "source_status": "user-modified scenario" if st.session_state.get("edited") else "source-derived",
    "observation_count": len(validated) if "validated" in locals() else 0,
    "degree": degree, "base_year": model["base"] if model else None,
    "equation": poly_equation(model) if model else "Model unavailable",
    "r_squared_in_sample": model["r2"] if model else None, "mae_in_sample": model["mae"] if model else None
}
st.subheader("Export current selection")
data_export = validated if "validated" in locals() else pd.DataFrame(columns=["year","value"])
csv_data = data_export.to_csv(index=False).encode("utf-8")
st.download_button("Download working data CSV",csv_data,file_name="hispanic_heritage_working_data.csv",mime="text/csv")
calc_export = calc if "calc" in locals() else pd.DataFrame(columns=["year","estimate","units","range_status"])
st.download_button("Download Excel workbook (.xlsx)",excel_bytes(summary,data_export,calc_export),file_name="hispanic_heritage_explorer.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
if model and chart_png:
    summary_text = "\n".join([f"{k}: {v}" for k,v in summary.items()]) + "\nLimitations: in-sample fit; not causal; extrapolation uncertain; no validated 95% prediction interval."
    st.download_button("Download PDF report",pdf_bytes(summary_text,chart_png,data_export,calc_export),file_name="hispanic_heritage_report.pdf",mime="application/pdf")
if printer_mode:
    st.markdown("---\n## Printer-friendly summary")
    st.write(f"{country_name} • {indicator_name} • {units} • Retrieved {date.today().isoformat()}")
    st.write(f"Source coverage: {source_data.year.min() if not source_data.empty else 'none'}–{source_data.year.max() if not source_data.empty else 'none'}; {len(source_data)} numeric observations.")
    if model: st.write(poly_equation(model), f"In-sample R² = {model['r2']:.5g}; in-sample MAE = {model['mae']:.5g} {units}.")
    st.write("Predictions are model estimates, not official forecasts. Extrapolations are uncertain; no causal claim is made.")
