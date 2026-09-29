"""
Stream-lite — Baseline Table Builder
=====================================
A Streamlit app that turns an uploaded master chart (Excel/CSV) into a
publication-ready "Table 1": descriptive statistics per variable, with
automatic parametric/non-parametric test selection.

Run with:
    pip install streamlit pandas numpy scipy openpyxl python-docx
    streamlit run streamlite_app.py

How test selection works
-------------------------
Numeric variables, 2 groups   -> Welch's t-test (parametric) or
                                  Mann-Whitney U / Wilcoxon rank-sum (non-parametric)
Numeric variables, 3+ groups  -> One-way ANOVA or Kruskal-Wallis H
Categorical variables         -> Chi-square test of independence, automatically
                                  switched to Fisher's exact test for 2x2 tables
                                  when any expected cell count is below 5

Two table layouts
-----------------
1. Standard Table 1: a categorical grouping variable goes in the columns and
   the selected variables go down the rows.
2. Continuous-variable layout: pick one continuous variable (e.g. Age) for the
   columns and the selected categorical factors (e.g. Sex, Smoking) go down the
   rows. For each factor level you get n and mean \u00B1 SD / median (IQR) of the
   continuous variable, plus a test comparing the levels of that factor
   (Welch t / Mann-Whitney for 2 levels, ANOVA / Kruskal-Wallis for 3+).

3. Custom table: several continuous variables side by side, up to two nested
   grouping variables (order 1 = outer, order 2 = inner) splitting every
   column, factors as rows, and either the column groups or the factor levels
   compared with a test.

Parametric vs non-parametric is decided automatically via the
D'Agostino-Pearson omnibus normality test (scipy.stats.normaltest),
unless the user forces one or the other.
"""

import io
import itertools
import html as _html
import numpy as np
import pandas as pd
import streamlit as st
from scipy import stats
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import Workbook
from openpyxl.styles import Font as OpenpyxlFont, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Stream-lite · Baseline Table Builder", layout="wide")

# --------------------------------------------------------------------------
# Statistical helpers
# --------------------------------------------------------------------------

def detect_type(series: pd.Series):
    """Guess whether a column is numerical or categorical."""
    non_missing = series.dropna()
    non_missing = non_missing[non_missing.astype(str).str.strip() != ""]
    n = len(non_missing)
    if n == 0:
        return "categorical", 0, 0
    numeric_coerced = pd.to_numeric(non_missing, errors="coerce")
    numeric_ratio = numeric_coerced.notna().mean()
    unique_n = non_missing.astype(str).str.strip().nunique()
    if numeric_ratio >= 0.9 and unique_n > 10:
        return "numerical", n, unique_n
    return "categorical", n, unique_n


def effective_type(meta):
    """Resolve a variable's working type: if the user left it on 'auto',
    use the auto-detected type; otherwise use their manual override."""
    return meta["detected"] if meta["type"] == "auto" else meta["type"]


def is_normal(arr, alpha):
    """D'Agostino-Pearson omnibus normality test. Needs n>=8; smaller
    samples are treated as non-normal (safer default)."""
    arr = np.asarray(arr, dtype=float)
    arr = arr[~np.isnan(arr)]
    if len(arr) < 8:
        return False
    if np.all(arr == arr[0]):
        return True
    try:
        _, p = stats.normaltest(arr)
        return p > alpha
    except Exception:
        return False


def welch_t_test(a, b):
    a, b = np.asarray(a, float), np.asarray(b, float)
    res = stats.ttest_ind(a, b, equal_var=False)
    v1, v2, n1, n2 = a.var(ddof=1), b.var(ddof=1), len(a), len(b)
    df = (v1 / n1 + v2 / n2) ** 2 / ((v1 / n1) ** 2 / (n1 - 1) + (v2 / n2) ** 2 / (n2 - 1))
    return {"name": "Independent t-test (Welch)",
            "stat": f"t={res.statistic:.2f}, df={df:.1f}",
            "p": float(res.pvalue)}


def mann_whitney_test(a, b):
    res = stats.mannwhitneyu(a, b, alternative="two-sided")
    return {"name": "Mann-Whitney U (Wilcoxon rank-sum)",
            "stat": f"U={res.statistic:.1f}",
            "p": float(res.pvalue)}


def anova_test(groups):
    res = stats.f_oneway(*groups)
    k = len(groups)
    N = sum(len(g) for g in groups)
    return {"name": "One-way ANOVA",
            "stat": f"F={res.statistic:.2f}, df={k-1},{N-k}",
            "p": float(res.pvalue)}


def kruskal_test(groups):
    res = stats.kruskal(*groups)
    return {"name": "Kruskal-Wallis H",
            "stat": f"H={res.statistic:.2f}, df={len(groups)-1}",
            "p": float(res.pvalue)}


def chi_or_fisher_test(table, yates_correction=False):
    """RxC contingency table -> chi-square, auto-falling back to Fisher's
    exact test for 2x2 tables with low expected counts.
    yates_correction controls scipy's Yates' continuity correction, which
    only has an effect on 2x2 tables (scipy silently ignores it otherwise).
    Off by default, since the correction is conservative and not universally
    recommended; many modern guidelines prefer the uncorrected chi-square or
    Fisher's exact test for small 2x2 tables instead."""
    table = np.array(table)
    chi2, p, dof, expected = stats.chi2_contingency(table, correction=yates_correction)
    min_e = float(expected.min())
    if table.shape == (2, 2) and min_e < 5:
        _, p_fisher = stats.fisher_exact(table)
        return {"name": "Fisher's exact test", "stat": "—", "p": float(p_fisher), "min_expected": min_e}
    name = "Chi-square test"
    if table.shape == (2, 2) and yates_correction:
        name = "Chi-square test (Yates-corrected)"
    return {"name": name, "stat": f"\u03C7\u00B2={chi2:.2f}, df={dof}",
            "p": float(p), "min_expected": min_e}


def fmt_p(p):
    if p is None or (isinstance(p, float) and np.isnan(p)):
        return "—"
    return "<0.001" if p < 0.001 else f"{p:.3f}"


def fmt_num(x, d=3):
    return f"{x:.{d}f}"


# --------------------------------------------------------------------------
# Table 1 builder
# --------------------------------------------------------------------------

def format_numeric_cell(values, display_mode, use_param, digits=3):
    """Format one group's numeric summary according to the chosen display mode.
    display_mode: 'auto' | 'mean_sd' | 'median_iqr' | 'both'
    digits: decimal places for every number shown."""
    if len(values) == 0:
        return "\u2014"
    show_mean = display_mode == "mean_sd" or display_mode == "both" or (display_mode == "auto" and use_param)
    show_median = display_mode == "median_iqr" or display_mode == "both" or (display_mode == "auto" and not use_param)
    parts = []
    if show_mean:
        sd_txt = fmt_num(np.std(values, ddof=1), digits) if len(values) > 1 else "\u2014"
        parts.append(f"{fmt_num(np.mean(values), digits)} \u00B1 {sd_txt}")
    if show_median:
        q1, med, q3 = np.percentile(values, [25, 50, 75])
        parts.append(f"{fmt_num(med, digits)} ({fmt_num(q1, digits)}\u2013{fmt_num(q3, digits)})")
    return "; ".join(parts)


def numeric_label(col, display_mode, use_param):
    if display_mode == "mean_sd":
        return f"{col}, mean \u00B1 SD"
    if display_mode == "median_iqr":
        return f"{col}, median (IQR)"
    if display_mode == "both":
        return f"{col}, mean \u00B1 SD; median (IQR)"
    return f"{col}, mean \u00B1 SD" if use_param else f"{col}, median (IQR)"


def build_table1(df, var_meta, group_col, test_mode, alpha, display_mode="auto",
                  pct_mode="column", pct_digits=2, yates_correction=False):
    """Returns (display_rows, csv_rows, footnote_flags).
    display_rows: list of dicts describing each printed row for st rendering.
    csv_rows: list of lists for CSV / clipboard export.
    display_mode controls how numeric summaries are shown: 'auto' (mean±SD if
    normal else median (IQR)), 'mean_sd', 'median_iqr', or 'both'. This is
    independent of test_mode, which controls whether t-test/ANOVA or
    Wilcoxon/Kruskal-Wallis is used for the comparison.
    pct_mode controls the denominator used for categorical n (%) cells when a
    grouping variable is set: 'column' (default) expresses each cell as a
    percentage of its own group's total (columns sum to ~100%); 'row'
    expresses each cell as a percentage of that category's total across all
    groups (rows sum to ~100%). Ignored when there is no grouping variable
    (percentages are always of the overall n in that case).
    pct_digits controls the number of decimal places shown for all
    categorical percentages.
    yates_correction controls whether Yates' continuity correction is applied
    to the chi-square test for 2x2 tables. Off by default. Has no effect on
    Fisher's exact test or on chi-square tables larger than 2x2.
    """
    group_levels = []
    if group_col:
        group_levels = sorted(df[group_col].dropna().astype(str).str.strip().unique().tolist())

    header = ["Variable"]
    if group_col:
        for lv in group_levels:
            n = (df[group_col].astype(str).str.strip() == lv).sum()
            header.append(f"{lv} (n={n})")
        header += ["Test", "Statistic", "p"]
    else:
        header.append(f"Overall (n={len(df)})")

    csv_rows = [header]
    display_rows = []
    flags = set()

    variables = [c for c in df.columns if var_meta[c]["use"] and c != group_col]

    for col in variables:
        vtype = effective_type(var_meta[col])

        if vtype == "numerical":
            numeric_series = pd.to_numeric(df[col], errors="coerce")

            if group_col:
                group_arrays = []
                for lv in group_levels:
                    mask = (df[group_col].astype(str).str.strip() == lv) & numeric_series.notna()
                    group_arrays.append(numeric_series[mask].values)
            else:
                group_arrays = None

            all_vals = numeric_series.dropna().values

            # use_param governs which TEST is used (t-test/ANOVA vs Wilcoxon/KW);
            # it is decided by test_mode regardless of the display_mode setting.
            if test_mode == "parametric":
                use_param = True
            elif test_mode == "nonparametric":
                use_param = False
            else:
                if group_col:
                    use_param = all(is_normal(g, alpha) for g in group_arrays if len(g) > 0)
                else:
                    use_param = is_normal(all_vals, alpha)

            label = numeric_label(col, display_mode, use_param)

            cells = []
            if group_col:
                for g in group_arrays:
                    cells.append(format_numeric_cell(g, display_mode, use_param))
            else:
                cells.append(format_numeric_cell(all_vals, display_mode, use_param))

            test_result = None
            if group_col and len(group_levels) >= 2:
                nonempty = [g for g in group_arrays if len(g) > 1]  # need >=2 points per group for variance
                if len(nonempty) >= 2:
                    try:
                        if len(group_levels) == 2:
                            test_result = welch_t_test(*nonempty) if use_param else mann_whitney_test(*nonempty)
                        else:
                            test_result = anova_test(nonempty) if use_param else kruskal_test(nonempty)
                    except Exception:
                        test_result = None
                        flags.add("skipped")

            display_rows.append({"kind": "var", "label": label, "cells": cells, "test": test_result})
            csv_rows.append([label, *cells,
                              test_result["name"] if test_result else "",
                              test_result["stat"] if test_result else "",
                              fmt_p(test_result["p"]) if test_result else ""])

        else:  # categorical
            series = df[col].astype(str).str.strip()
            series = series.where(df[col].notna() & (series != ""), other=np.nan)
            levels = sorted(series.dropna().unique().tolist())

            contingency = None
            test_result = None
            if group_col and len(group_levels) >= 2 and len(levels) >= 2:
                group_series = df[group_col].astype(str).str.strip()
                contingency = [[int(((series == lv) & (group_series == glv)).sum()) for glv in group_levels]
                               for lv in levels]
                # Skip the test if any row/column is entirely zero — chi2_contingency
                # (and Fisher's exact) require every row and column to have at least
                # one observation, otherwise the table is degenerate.
                arr = np.array(contingency)
                if arr.size > 0 and arr.shape[0] >= 2 and arr.shape[1] >= 2 \
                        and (arr.sum(axis=0) > 0).all() and (arr.sum(axis=1) > 0).all():
                    try:
                        res = chi_or_fisher_test(contingency, yates_correction=yates_correction)
                        test_result = res
                        if res["name"].startswith("Fisher"):
                            flags.add("fisher")
                        else:
                            flags.add("chi2")
                            if res["min_expected"] < 5:
                                flags.add("lowE")
                    except Exception:
                        test_result = None
                        flags.add("skipped")
            elif group_col:
                # Build a zero contingency table skeleton for the n(%) display below,
                # even though no test is run (fewer than 2 non-empty levels/groups).
                group_series = df[group_col].astype(str).str.strip()
                contingency = [[int(((series == lv) & (group_series == glv)).sum()) for glv in group_levels]
                               for lv in levels]

            display_rows.append({"kind": "varheader", "label": f"{col}, n (%)"})
            csv_rows.append([f"{col}, n (%)"])
            flags.add("catpct")

            # Non-missing total for this variable, overall and per group. Used
            # as the % denominator so missing values are excluded (rather than
            # counting them against the group/overall total), giving true
            # "percent of observed" values.
            total_nonmissing = int(series.notna().sum())
            group_nonmissing_totals = None
            if group_col and levels:
                group_nonmissing_totals = [sum(contingency[i][gi] for i in range(len(levels)))
                                            for gi in range(len(group_levels))]

            for i, lv in enumerate(levels):
                cells = []
                if group_col:
                    row_total = sum(contingency[i]) if pct_mode == "row" else None
                    for gi, glv in enumerate(group_levels):
                        n = contingency[i][gi]
                        if pct_mode == "row":
                            denom = row_total
                        else:
                            denom = group_nonmissing_totals[gi]
                        pct = 100 * n / denom if denom else 0.0
                        cells.append(f"{n} ({pct:.{pct_digits}f}%)")
                else:
                    n = int((series == lv).sum())
                    pct = 100 * n / total_nonmissing if total_nonmissing else 0.0
                    cells.append(f"{n} ({pct:.{pct_digits}f}%)")

                is_last = i == len(levels) - 1
                display_rows.append({"kind": "level", "label": lv, "cells": cells,
                                      "test": test_result if is_last else None})
                csv_rows.append([f"  {lv}", *cells,
                                  test_result["name"] if is_last and test_result else "",
                                  test_result["stat"] if is_last and test_result else "",
                                  fmt_p(test_result["p"]) if is_last and test_result else ""])

    return header, display_rows, csv_rows, flags


def build_table_by_factors(df, var_meta, outcome_col, test_mode, alpha, display_mode="auto"):
    """Alternate layout: ONE continuous variable (outcome_col) sits in the
    columns and the selected categorical factors run down the rows.

    For every level of every factor it reports n (non-missing values of the
    continuous variable) and its mean +/- SD and/or median (IQR), then tests
    whether the continuous variable differs across the levels of that factor:
        2 levels  -> Welch t-test or Mann-Whitney U
        3+ levels -> one-way ANOVA or Kruskal-Wallis H
    Parametric vs non-parametric follows test_mode ('auto' = D'Agostino-Pearson
    on every level). display_mode only controls how the summary is printed.

    Factors = variables ticked 'Use' whose effective type is categorical
    (the continuous variable itself is never used as a factor).
    Returns (header, display_rows, csv_rows, flags), same shape as build_table1.
    """
    outcome = pd.to_numeric(df[outcome_col], errors="coerce")
    factors = [c for c in df.columns
               if var_meta[c]["use"] and c != outcome_col
               and effective_type(var_meta[c]) == "categorical"]

    if display_mode == "mean_sd":
        stat_headers = [f"{outcome_col}, mean \u00B1 SD"]
    elif display_mode == "median_iqr":
        stat_headers = [f"{outcome_col}, median (IQR)"]
    elif display_mode == "both":
        stat_headers = [f"{outcome_col}, mean \u00B1 SD", f"{outcome_col}, median (IQR)"]
    else:
        stat_headers = [f"{outcome_col}, mean \u00B1 SD or median (IQR)"]

    header = ["Factor", "n", *stat_headers, "Test", "Statistic", "p"]
    csv_rows = [header]
    display_rows = []
    flags = set()

    def make_cells(values, use_param):
        n = len(values)
        if display_mode == "both":
            return [str(n), format_numeric_cell(values, "mean_sd", True),
                    format_numeric_cell(values, "median_iqr", False)]
        return [str(n), format_numeric_cell(values, display_mode, use_param)]

    def pick_param(arrays):
        if test_mode == "parametric":
            return True
        if test_mode == "nonparametric":
            return False
        return all(is_normal(a, alpha) for a in arrays if len(a) > 0)

    # Overall row
    all_vals = outcome.dropna().values
    overall_param = pick_param([all_vals])
    overall_cells = make_cells(all_vals, overall_param)
    display_rows.append({"kind": "var", "label": f"Overall", "cells": overall_cells, "test": None})
    csv_rows.append(["Overall", *overall_cells, "", "", ""])

    for col in factors:
        series = df[col].astype(str).str.strip()
        series = series.where(df[col].notna() & (series != ""), other=np.nan)
        levels = sorted(series.dropna().unique().tolist())
        if not levels:
            continue

        arrays = [outcome[(series == lv) & outcome.notna()].values for lv in levels]
        use_param = pick_param(arrays)

        test_result = None
        if len(levels) >= 2:
            usable = [a for a in arrays if len(a) > 1]   # need >=2 points per level
            if any(len(a) < 2 for a in arrays):
                flags.add("smalln")
            if len(usable) >= 2:
                try:
                    if len(usable) == 2:
                        test_result = welch_t_test(*usable) if use_param else mann_whitney_test(*usable)
                    else:
                        test_result = anova_test(usable) if use_param else kruskal_test(usable)
                except Exception:
                    test_result = None
                    flags.add("skipped")
            else:
                flags.add("skipped")

        label = f"{col}"
        if display_mode == "auto":
            label += ", mean \u00B1 SD" if use_param else ", median (IQR)"
        display_rows.append({"kind": "varheader", "label": label})
        csv_rows.append([label])

        for i, (lv, arr) in enumerate(zip(levels, arrays)):
            cells = make_cells(arr, use_param)
            is_last = i == len(levels) - 1
            display_rows.append({"kind": "level", "label": lv, "cells": cells,
                                  "test": test_result if is_last else None})
            csv_rows.append([f"  {lv}", *cells,
                              test_result["name"] if is_last and test_result else "",
                              test_result["stat"] if is_last and test_result else "",
                              fmt_p(test_result["p"]) if is_last and test_result else ""])

    return header, display_rows, csv_rows, flags


def factor_footnotes(display_mode, alpha, outcome_col, flags):
    """Footnotes for the continuous-variable-in-columns layout."""
    if display_mode == "mean_sd":
        desc = "mean \u00B1 SD"
    elif display_mode == "median_iqr":
        desc = "median (IQR)"
    elif display_mode == "both":
        desc = "mean \u00B1 SD and median (IQR)"
    else:
        desc = (f"mean \u00B1 SD (all levels assessed as normal via D'Agostino-Pearson test, "
                f"\u03B1={alpha}) or median (IQR) otherwise")
    notes = [
        f"{outcome_col} is reported as {desc} within each level of the factor. n is the number of "
        f"non-missing {outcome_col} values; observations with a missing factor value are excluded "
        f"from that factor's rows.",
        "Two-level factors were compared with Welch's t-test (parametric) or the Mann-Whitney U test "
        "(non-parametric); factors with three or more levels with one-way ANOVA or the Kruskal-Wallis "
        "H test.",
    ]
    if "smalln" in flags:
        notes.append("Levels with fewer than 2 non-missing observations were excluded from the "
                     "significance test for that factor.")
    if "skipped" in flags:
        notes.append("Note: a statistical test could not be computed for one or more factors "
                     "(e.g. insufficient data) and was left blank.")
    notes.append(f"Bold p-values indicate statistical significance at \u03B1={alpha}.")
    return notes


def render_table_markdown(header, display_rows, group_col, alpha):
    """Render the Table 1 as an HTML table with a three-line (journal-style) look."""
    css = """
    <style>
    table.pub { width:100%; border-collapse:collapse; font-family: Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif; font-size: 9pt; }
    table.pub thead th { border-top:2px solid #1E2A32; border-bottom:1px solid #1E2A32;
                          padding:8px 10px; text-align:left; font-family: inherit; }
    table.pub tbody td { padding:5px 10px; }
    table.pub tbody tr.var td { font-weight:700; padding-top:10px; }
    table.pub tbody tr.level td.name { padding-left:20px; color:#555; font-weight:400; }
    table.pub td.stat { font-family: inherit; text-align:center; font-size:9pt;}
    table.pub td.sig { font-weight:700; }
    table.pub tbody tr.lastrow td { border-bottom:2px solid #1E2A32; padding-bottom:10px; }
    </style>
    """
    html = css + '<table class="pub"><thead><tr>'
    for h in header:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"

    for idx, row in enumerate(display_rows):
        is_last_of_block = (idx == len(display_rows) - 1) or \
                            (display_rows[idx + 1]["kind"] in ("var", "varheader"))
        cls = "var" if row["kind"] in ("var", "varheader") else "level"
        cls += " lastrow" if is_last_of_block else ""
        html += f'<tr class="{cls}">'

        if row["kind"] == "varheader":
            html += f'<td>{row["label"]}</td>'
            colspan = len(header) - 1
            html += f'<td colspan="{colspan}"></td>'
        else:
            name_cls = "name" if row["kind"] == "level" else ""
            html += f'<td class="{name_cls}">{row["label"]}</td>'
            for c in row["cells"]:
                html += f'<td class="stat">{c}</td>'
            if group_col:
                t = row.get("test")
                if t:
                    sig = "sig" if t["p"] < alpha else ""
                    html += f'<td>{t["name"]}</td><td class="stat">{t["stat"]}</td><td class="stat {sig}">{fmt_p(t["p"])}</td>'
                else:
                    html += "<td>—</td><td>—</td><td>—</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html


def _set_cell_border(cell, **kwargs):
    """Add borders to a single table cell. kwargs like
    top={'sz':12,'val':'single','color':'000000'}."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right'):
        if edge in kwargs:
            spec = kwargs[edge]
            tag = f'w:{edge}'
            el = tcBorders.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                tcBorders.append(el)
            el.set(qn('w:val'), spec.get('val', 'single'))
            el.set(qn('w:sz'), str(spec.get('sz', 8)))
            el.set(qn('w:color'), spec.get('color', '000000'))


def _set_run_font(run, name="Calibri", size=9, bold=None, italic=None):
    run.font.name = name
    run.font.size = Pt(size)
    # Ensure the font also applies to complex-script/east-asian text runs in Word
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:cs'), name)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic


def descriptive_footnote(display_mode, alpha):
    if display_mode == "mean_sd":
        return "Continuous variables reported as mean \u00B1 SD; categorical variables reported as n (%)."
    if display_mode == "median_iqr":
        return "Continuous variables reported as median (IQR); categorical variables reported as n (%)."
    if display_mode == "both":
        return "Continuous variables reported as mean \u00B1 SD and median (IQR); categorical variables reported as n (%)."
    return (f"Continuous variables reported as mean \u00B1 SD (assessed as normal via D'Agostino-Pearson "
            f"test, \u03B1={alpha}) or median (IQR) otherwise; categorical variables reported as n (%).")


def pct_basis_footnote(group_col, pct_mode):
    """Explain the denominator used for categorical n (%) cells."""
    if not group_col:
        return None
    if pct_mode == "row":
        return ("Percentages for categorical variables are row-wise: each n (%) is a share of that "
                "category's total across all groups (rows sum to ~100%).")
    return ("Percentages for categorical variables are column-wise: each n (%) is a share of its own "
            "group's total (columns sum to ~100%).")


def build_excel(csv_rows, sheet_name="Table 1"):
    """Build an .xlsx file directly with openpyxl (bypassing pandas'
    ExcelWriter/openpyxl sheet-swap, which on some pandas/openpyxl/Python
    version combinations raises 'IndexError: At least one sheet must be
    visible'). csv_rows is the same list-of-lists used for the CSV export,
    with csv_rows[0] as the header row."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_state = "visible"

    for row in csv_rows:
        ws.append(row)

    for xl_row in ws.iter_rows():
        for cell in xl_row:
            bold = cell.row == 1
            cell.font = OpenpyxlFont(name="Calibri", size=9, bold=bold)

    for column_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = min(max(length + 2, 10), 45)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Custom table: continuous variable(s) x grouping (order 1 / order 2) x factor rows
# --------------------------------------------------------------------------

def _layout_header(hdr):
    """Give every header cell a (row, col) grid position, honouring rowspans.
    hdr is a list of header rows; each cell is {'text', 'colspan', 'rowspan'}.
    Returns a list of (row, col, cell)."""
    occupied, placed = set(), []
    for r, row in enumerate(hdr):
        c = 0
        for cell in row:
            while (r, c) in occupied:
                c += 1
            placed.append((r, c, cell))
            for dr in range(cell["rowspan"]):
                for dc in range(cell["colspan"]):
                    occupied.add((r + dr, c + dc))
            c += cell["colspan"]
    return placed


def run_comparison(arrays, use_param, flags):
    """Compare a list of numeric arrays: Welch t / Mann-Whitney for 2 groups,
    ANOVA / Kruskal-Wallis for 3+. Groups with <2 values are dropped (flagged).
    Returns a result dict or None."""
    if len(arrays) < 2:
        return None
    usable = [a for a in arrays if len(a) > 1]
    if len(usable) < len(arrays):
        flags.add("smalln")
    if len(usable) < 2:
        flags.add("skipped")
        return None
    try:
        if len(usable) == 2:
            return welch_t_test(*usable) if use_param else mann_whitney_test(*usable)
        return anova_test(usable) if use_param else kruskal_test(usable)
    except Exception:
        flags.add("skipped")
        return None


def build_custom_table(df, cont_vars, group_cols, group_levels, factor_cols, test_mode,
                       display_mode, digits, alpha, compare="groups", show_n=True):
    """Custom stratified table.

    cont_vars    : one or more continuous variables -> top-level column blocks
    group_cols   : 0-2 grouping variables (order 1 = outer split, order 2 = inner split);
                   every block is split into one column per level (or level combination)
    group_levels : {group_col: [levels to show]}
    factor_cols  : categorical variables -> row blocks (one row per level)
    test_mode    : 'auto' | 'parametric' | 'nonparametric'
    display_mode : 'auto' | 'mean_sd' | 'median_iqr' | 'both'   (auto = by normality)
    compare      : 'groups' -> for each row level, compare the column groups
                              (test columns at the far right)
                   'levels' -> for each column group, compare the levels of the factor
                              (p-value row under each factor)
                   'none'   -> no tests
    Returns (hdr, rows, csv_rows, flags).
    """
    flags = set()
    outs = {v: pd.to_numeric(df[v], errors="coerce") for v in cont_vars}

    def clean(col):
        s = df[col].astype(str).str.strip()
        return s.where(df[col].notna() & (s != ""), other=np.nan)

    gseries = [clean(gc) for gc in group_cols]
    glevels = [list(group_levels.get(gc) or sorted(s.dropna().unique().tolist()))
               for gc, s in zip(group_cols, gseries)]
    combos = list(itertools.product(*glevels)) if group_cols else [()]
    masks = []
    for combo in combos:
        m = pd.Series(True, index=df.index)
        for s, lv in zip(gseries, combo):
            m = m & (s == lv)
        masks.append(m)
    ngc = len(combos)

    if not group_cols and compare == "groups":
        compare = "levels"          # nothing to compare across columns
    n_test_cols = 3 * len(cont_vars) if compare == "groups" else 0

    # ---- multi-level header ------------------------------------------------
    n_hdr = 1 + len(group_cols)
    hdr = [[] for _ in range(n_hdr)]
    hdr[0].append({"text": "Variable", "colspan": 1, "rowspan": n_hdr})
    for v in cont_vars:
        hdr[0].append({"text": v, "colspan": ngc, "rowspan": 1})
        if len(group_cols) >= 1:
            span = len(glevels[1]) if len(group_cols) == 2 else 1
            for l1 in glevels[0]:
                hdr[1].append({"text": l1, "colspan": span, "rowspan": 1})
        if len(group_cols) == 2:
            for _l1 in glevels[0]:
                for l2 in glevels[1]:
                    hdr[2].append({"text": l2, "colspan": 1, "rowspan": 1})
    if compare == "groups":
        for v in cont_vars:
            hdr[0].append({"text": f"{v}: comparison" if len(cont_vars) > 1 else "Comparison",
                           "colspan": 3, "rowspan": 1})
        for _v in cont_vars:
            for t in ("Test", "Statistic", "p"):
                hdr[1].append({"text": t, "colspan": 1, "rowspan": n_hdr - 1})

    # ---- helpers -------------------------------------------------------------
    def pick_param(norm_ok):
        if test_mode == "parametric":
            return True
        if test_mode == "nonparametric":
            return False
        return norm_ok

    def fmt_cell(a, norm_ok):
        if len(a) == 0:
            return "\u2014"
        txt = format_numeric_cell(a, display_mode, norm_ok, digits)
        return f"{txt} [n={len(a)}]" if show_n else txt

    def test_block(res):
        if res is None:
            return ["\u2014"] * 3, [False] * 3
        return [res["name"], res["stat"], fmt_p(res["p"])], [False, False, bool(res["p"] < alpha)]

    rows = []

    # ---- overall row -----------------------------------------------------------
    data, tests, tsig = [], [], []
    for v in cont_vars:
        arrs = [outs[v][m & outs[v].notna()].values for m in masks]
        norm_ok = all(is_normal(a, alpha) for a in arrs if len(a) > 0)
        data += [fmt_cell(a, norm_ok) for a in arrs]
        if compare == "groups":
            cells_t, sig_t = test_block(run_comparison(arrs, pick_param(norm_ok), flags))
            tests += cells_t
            tsig += sig_t
    rows.append({"kind": "var", "label": "Overall", "cells": data + tests,
                 "sig": [False] * len(data) + tsig})

    # ---- factor blocks -------------------------------------------------------
    for fc in factor_cols:
        fs = clean(fc)
        levels = sorted(fs.dropna().unique().tolist())
        if not levels:
            continue
        # arrays[v][level_index][column_index]
        arrays = {v: [[outs[v][masks[ci] & (fs == lv) & outs[v].notna()].values for ci in range(ngc)]
                      for lv in levels] for v in cont_vars}
        norm_ok = {v: all(is_normal(a, alpha) for lvrow in arrays[v] for a in lvrow if len(a) > 0)
                   for v in cont_vars}

        hint = []
        for v in cont_vars:
            h = ("mean \u00B1 SD" if norm_ok[v] else "median (IQR)") if display_mode == "auto" else ""
            hint += [h] * ngc
        rows.append({"kind": "varheader", "label": fc, "cells": hint + [""] * n_test_cols,
                     "sig": [False] * (len(hint) + n_test_cols)})

        for li, lv in enumerate(levels):
            data, tests, tsig = [], [], []
            for v in cont_vars:
                data += [fmt_cell(a, norm_ok[v]) for a in arrays[v][li]]
                if compare == "groups":
                    cells_t, sig_t = test_block(
                        run_comparison(arrays[v][li], pick_param(norm_ok[v]), flags))
                    tests += cells_t
                    tsig += sig_t
            rows.append({"kind": "level", "label": lv, "cells": data + tests,
                         "sig": [False] * len(data) + tsig})

        if compare == "levels" and len(levels) >= 2:
            pdata, psig, any_res = [], [], False
            for v in cont_vars:
                for ci in range(ngc):
                    res = run_comparison([arrays[v][li][ci] for li in range(len(levels))],
                                         pick_param(norm_ok[v]), flags)
                    ptxt = fmt_p(res["p"]) if res is not None else "\u2014"
                    if res is None or ptxt == "\u2014":
                        pdata.append("\u2014")
                        psig.append(False)
                        continue
                    any_res = True
                    ptxt = f"p{ptxt}" if ptxt.startswith("<") else f"p={ptxt}"
                    pdata.append(f"{res['stat']}; {ptxt}" if res["stat"] != "\u2014" else ptxt)
                    psig.append(bool(res["p"] < alpha))
            if any_res:
                rows.append({"kind": "prow", "label": "p-value", "cells": pdata + [""] * n_test_cols,
                             "sig": psig + [False] * n_test_cols})

    # ---- flat CSV (header rows flattened) -------------------------------------
    ncols = 1 + len(rows[0]["cells"])
    grid = [[""] * ncols for _ in hdr]
    for r, c, cell in _layout_header(hdr):
        grid[r][c] = cell["text"]
    csv_rows = grid[:]
    for row in rows:
        label = ("  " + row["label"]) if row["kind"] in ("level", "prow") else row["label"]
        csv_rows.append([label, *row["cells"]])
    return hdr, rows, csv_rows, flags


def custom_footnotes(cont_vars, group_cols, display_mode, test_mode, alpha, compare, show_n, flags):
    """Footnotes for the custom table."""
    vars_txt = ", ".join(cont_vars)
    if display_mode == "mean_sd":
        desc = "mean \u00B1 SD"
    elif display_mode == "median_iqr":
        desc = "median (IQR)"
    elif display_mode == "both":
        desc = "mean \u00B1 SD; median (IQR)"
    else:
        desc = (f"mean \u00B1 SD when all groups in that factor block were normally distributed "
                f"(D'Agostino-Pearson, \u03B1={alpha}) and median (IQR) otherwise (the italic label "
                f"under each factor name shows which)")
    strat = ""
    if group_cols:
        strat = (" Columns are stratified by " + group_cols[0]
                 + (f" (order 1) and then {group_cols[1]} (order 2)." if len(group_cols) == 2 else "."))
    n_txt = " [n=\u2026] is the number of non-missing values in that cell." if show_n else ""
    notes = [f"Values of {vars_txt} within each column group and row level are reported as {desc}."
             f"{strat}{n_txt} Observations with a missing group or factor value are excluded from the "
             f"relevant cells."]

    if test_mode == "parametric":
        basis = "Parametric tests were forced."
    elif test_mode == "nonparametric":
        basis = "Non-parametric tests were forced."
    else:
        basis = "Parametric or non-parametric tests were chosen by normality (D'Agostino-Pearson)."
    if compare == "groups" and group_cols:
        notes.append("p-values compare the column groups within each row (Welch's t-test or Mann-Whitney U "
                     "for two groups; one-way ANOVA or Kruskal-Wallis H for three or more). " + basis)
    elif compare in ("groups", "levels"):
        notes.append("p-values compare the levels of each factor within each column group (Welch's t-test "
                     "or Mann-Whitney U for two levels; one-way ANOVA or Kruskal-Wallis H for three or "
                     "more) and are shown in the p-value row under each factor. " + basis)
    if "smalln" in flags:
        notes.append("Groups with fewer than 2 non-missing observations were excluded from the "
                     "significance tests.")
    if "skipped" in flags:
        notes.append("Note: a statistical test could not be computed for one or more comparisons "
                     "(e.g. insufficient data) and was left blank.")
    if compare != "none":
        notes.append(f"Bold p-values indicate statistical significance at \u03B1={alpha}.")
    return notes


def render_custom_html(hdr, rows, alpha):
    """HTML version of the custom table (multi-level header, journal-style rules)."""
    n_hdr = len(hdr)
    css = """
    <style>
    table.cust { width:100%; border-collapse:collapse; font-family: Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif; font-size:9pt; }
    table.cust th { padding:6px 10px; text-align:center; font-weight:700; }
    table.cust th.lbl { text-align:left; }
    table.cust th.ht { border-top:2px solid #1E2A32; }
    table.cust th.hb { border-bottom:1px solid #1E2A32; }
    table.cust th.hsub { border-bottom:1px solid #888; }
    table.cust td { padding:5px 10px; }
    table.cust td.stat { text-align:center; white-space:nowrap; }
    table.cust td.hint { text-align:center; font-weight:400; font-style:italic; color:#666; font-size:8pt; }
    table.cust td.name { padding-left:20px; color:#555; }
    table.cust td.sig { font-weight:700; }
    table.cust tr.var td, table.cust tr.varheader td { font-weight:700; padding-top:10px; }
    table.cust tr.varheader td.hint { font-weight:400; }
    table.cust tr.prow td { font-style:italic; }
    table.cust tr.blockend td { border-bottom:1px solid #999; padding-bottom:8px; }
    table.cust tr.final td { border-bottom:2px solid #1E2A32; padding-bottom:8px; }
    </style>
    """
    out = [css, '<div style="overflow-x:auto"><table class="cust"><thead>']
    for r, row in enumerate(hdr):
        out.append("<tr>")
        for ci, cell in enumerate(row):
            cs, rs = cell["colspan"], cell["rowspan"]
            cls = []
            if r == 0:
                cls.append("ht")
                if ci == 0:
                    cls.append("lbl")
            if r + rs == n_hdr:
                cls.append("hb")
            elif cs > 1:
                cls.append("hsub")
            attrs = (f' colspan="{cs}"' if cs > 1 else "") + (f' rowspan="{rs}"' if rs > 1 else "")
            out.append(f'<th class="{" ".join(cls)}"{attrs}>{_html.escape(str(cell["text"]))}</th>')
        out.append("</tr>")
    out.append("</thead><tbody>")

    for idx, row in enumerate(rows):
        nxt = rows[idx + 1]["kind"] if idx + 1 < len(rows) else None
        cls = [row["kind"]]
        if nxt is None:
            cls.append("final")
        elif nxt in ("var", "varheader"):
            cls.append("blockend")
        out.append(f'<tr class="{" ".join(cls)}">')
        name_cls = "" if row["kind"] in ("var", "varheader") else "name"
        out.append(f'<td class="{name_cls}">{_html.escape(str(row["label"]))}</td>')
        for txt, sig in zip(row["cells"], row["sig"]):
            c = "hint" if row["kind"] == "varheader" else "stat"
            if sig:
                c += " sig"
            out.append(f'<td class="{c}">{_html.escape(str(txt))}</td>')
        out.append("</tr>")
    out.append("</tbody></table></div>")
    return "".join(out)


def build_excel_custom(hdr, rows, sheet_name="Table 1"):
    """Excel export of the custom table, with merged multi-level header cells."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_state = "visible"
    n_hdr = len(hdr)
    ncols = 1 + len(rows[0]["cells"])
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    widths = [10] * ncols

    for r, c, cell in _layout_header(hdr):
        cs, rs = cell["colspan"], cell["rowspan"]
        ws.cell(row=r + 1, column=c + 1, value=cell["text"])
        if cs > 1 or rs > 1:
            ws.merge_cells(start_row=r + 1, start_column=c + 1, end_row=r + rs, end_column=c + cs)
        xc = ws.cell(row=r + 1, column=c + 1)
        xc.font = OpenpyxlFont(name="Calibri", size=9, bold=True)
        xc.alignment = center
        if cs == 1:
            widths[c] = max(widths[c], len(str(cell["text"])) + 2)

    for i, row in enumerate(rows):
        rr = n_hdr + 1 + i
        label = ("    " + row["label"]) if row["kind"] in ("level", "prow") else row["label"]
        lc = ws.cell(row=rr, column=1, value=label)
        lc.font = OpenpyxlFont(name="Calibri", size=9, bold=row["kind"] in ("var", "varheader"),
                               italic=row["kind"] == "prow")
        widths[0] = max(widths[0], len(label) + 2)
        for j, (txt, sig) in enumerate(zip(row["cells"], row["sig"])):
            xc = ws.cell(row=rr, column=j + 2, value=txt)
            xc.font = OpenpyxlFont(name="Calibri", size=8 if row["kind"] == "varheader" else 9,
                                   bold=bool(sig), italic=row["kind"] in ("varheader", "prow"))
            xc.alignment = Alignment(horizontal="center")
            widths[j + 1] = max(widths[j + 1], len(str(txt)) + 2)

    for ci, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(ci + 1)].width = min(w, 45)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def build_docx_custom(hdr, rows, footnotes, title):
    """Word export of the custom table: landscape page, merged multi-level header,
    three-line look, Calibri 9pt (8pt when the table is wide)."""
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(0.6))
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(9)

    h = doc.add_heading(title, level=2)
    for r in h.runs:
        _set_run_font(r, size=11, bold=True)

    n_hdr = len(hdr)
    ncols = 1 + len(rows[0]["cells"])
    fs = 9 if ncols <= 10 else 8
    table = doc.add_table(rows=n_hdr + len(rows), cols=ncols)
    table.autofit = True

    for r, c, cell in _layout_header(hdr):
        cs, rs = cell["colspan"], cell["rowspan"]
        tc = table.cell(r, c)
        if cs > 1 or rs > 1:
            tc = tc.merge(table.cell(r + rs - 1, c + cs - 1))
        tc.text = str(cell["text"])
        for p in tc.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c == 0 else WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                _set_run_font(run, size=fs, bold=True)
        if r == 0:
            _set_cell_border(tc, top={'sz': 12, 'val': 'single'})
        if r + rs == n_hdr:
            _set_cell_border(tc, bottom={'sz': 8, 'val': 'single'})
        elif cs > 1:
            _set_cell_border(tc, bottom={'sz': 4, 'val': 'single'})

    nr = len(rows)
    for i, row in enumerate(rows):
        r = n_hdr + i
        cells = [table.cell(r, j) for j in range(ncols)]
        label = ("    " + row["label"]) if row["kind"] in ("level", "prow") else row["label"]
        cells[0].text = label
        for run in cells[0].paragraphs[0].runs:
            _set_run_font(run, size=fs, bold=row["kind"] in ("var", "varheader"),
                          italic=row["kind"] == "prow")
        for j, txt in enumerate(row["cells"]):
            cell = cells[j + 1]
            cell.text = str(txt)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    _set_run_font(run, size=fs - 1 if row["kind"] == "varheader" else fs,
                                  bold=bool(row["sig"][j]),
                                  italic=row["kind"] in ("varheader", "prow"))
        nxt = rows[i + 1]["kind"] if i + 1 < nr else None
        if nxt is None:
            for c in cells:
                _set_cell_border(c, bottom={'sz': 12, 'val': 'single'})
        elif nxt in ("var", "varheader"):
            for c in cells:
                _set_cell_border(c, bottom={'sz': 4, 'val': 'single'})

    doc.add_paragraph()
    for f in footnotes:
        p = doc.add_paragraph(f)
        for run in p.runs:
            _set_run_font(run, size=9, italic=True)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def build_docx(header, display_rows, group_col, alpha, flags, display_mode="auto",
                pct_mode="column", yates_correction=False, title="Table 1. Baseline characteristics",
                footnotes_override=None):
    """Build a Word document with a three-line (journal-style) table.
    All table and footnote text uses Calibri 9pt."""
    doc = Document()

    # Set the document default (Normal style) to Calibri 9pt so anything
    # not explicitly styled below still matches.
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(9)

    h = doc.add_heading(title, level=2)
    for r in h.runs:
        _set_run_font(r, size=11, bold=True)

    ncols = len(header)
    table = doc.add_table(rows=1, cols=ncols)
    table.autofit = True

    hdr_cells = table.rows[0].cells
    for i, htext in enumerate(header):
        hdr_cells[i].text = str(htext)
        for p in hdr_cells[i].paragraphs:
            for run in p.runs:
                _set_run_font(run, size=9, bold=True)
        _set_cell_border(hdr_cells[i], top={'sz': 12, 'val': 'single'},
                          bottom={'sz': 8, 'val': 'single'})

    n_rows = len(display_rows)
    for idx, row in enumerate(display_rows):
        cells = table.add_row().cells
        is_last = idx == n_rows - 1
        next_is_new_block = is_last or (display_rows[idx + 1]["kind"] in ("var", "varheader"))

        if row["kind"] == "varheader":
            cells[0].text = row["label"]
            for run in cells[0].paragraphs[0].runs:
                _set_run_font(run, size=9, bold=True)
            for c in cells[1:]:
                c.text = ""
        else:
            label = ("    " + row["label"]) if row["kind"] == "level" else row["label"]
            cells[0].text = label
            for run in cells[0].paragraphs[0].runs:
                _set_run_font(run, size=9, bold=(row["kind"] == "var"))
            ci = 1
            for val in row["cells"]:
                cells[ci].text = str(val)
                for p in cells[ci].paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        _set_run_font(run, size=9)
                ci += 1
            if group_col:
                t = row.get("test")
                if t:
                    cells[ci].text = t["name"]
                    for run in cells[ci].paragraphs[0].runs:
                        _set_run_font(run, size=9)
                    ci += 1
                    cells[ci].text = t["stat"]
                    for run in cells[ci].paragraphs[0].runs:
                        _set_run_font(run, size=9)
                    ci += 1
                    cells[ci].text = fmt_p(t["p"])
                    for p in cells[ci].paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in p.runs:
                            _set_run_font(run, size=9, bold=(t["p"] < alpha))
                else:
                    cells[ci].text = ""; ci += 1
                    cells[ci].text = ""; ci += 1
                    cells[ci].text = ""

        if next_is_new_block:
            for c in cells:
                _set_cell_border(c, bottom={'sz': 8, 'val': 'single'})

    for c in table.rows[-1].cells:
        _set_cell_border(c, bottom={'sz': 12, 'val': 'single'})

    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    if run.font.size is None:
                        _set_run_font(run, size=9)

    footnotes = [descriptive_footnote(display_mode, alpha)]
    pct_note = pct_basis_footnote(group_col, pct_mode)
    if pct_note:
        footnotes.append(pct_note)
    if "catpct" in flags:
        footnotes.append("Percentages for categorical variables are calculated among non-missing "
                          "responses for that variable (missing values excluded from the denominator).")
    if "chi2" in flags:
        chi2_note = "Chi-square test of independence used for categorical comparisons with adequate expected cell counts"
        chi2_note += " (Yates' continuity correction applied to 2\u00D72 tables)." if yates_correction \
            else " (no continuity correction applied)."
        footnotes.append(chi2_note)
    if "fisher" in flags:
        footnotes.append("Fisher's exact test used in place of chi-square when a 2\u00D72 table had an expected cell count below 5.")
    if "lowE" in flags:
        footnotes.append("Caution: one or more categorical comparisons above have expected cell counts below 5; chi-square approximation may be unreliable.")
    if "skipped" in flags:
        footnotes.append("A statistical test could not be computed for one or more variables and was left blank.")
    footnotes.append(f"Bold p-values indicate statistical significance at \u03B1={alpha}.")

    if footnotes_override is not None:
        footnotes = list(footnotes_override)

    doc.add_paragraph()  # spacer
    for f in footnotes:
        p = doc.add_paragraph(f)
        for run in p.runs:
            _set_run_font(run, size=9, italic=True)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.title("Stream-lite · Baseline Table Builder")
st.caption(
    "Upload a master chart. Assign variable types, choose a grouping variable, and "
    "Stream-lite auto-selects the correct test — t-test or Wilcoxon for numeric variables, "
    "chi-square or Fisher's exact for categorical — and lays out a three-line publication table."
)

st.markdown("### 1. Upload master chart")
uploaded = st.file_uploader("Excel (.xlsx/.xls) or CSV. First row must be column headers.",
                             type=["xlsx", "xls", "csv"])

if uploaded is not None:
    is_excel = not uploaded.name.lower().endswith(".csv")
    sheet_name = None

    try:
        if is_excel:
            xls = pd.ExcelFile(uploaded)
            sheet_names = xls.sheet_names
            if len(sheet_names) > 1:
                sheet_name = st.selectbox("Select sheet", options=sheet_names)
            else:
                sheet_name = sheet_names[0]
            df = pd.read_excel(xls, sheet_name=sheet_name)
        else:
            df = pd.read_csv(uploaded)
    except Exception as e:
        st.error(f"Could not read that file: {e}")
        st.stop()

    st.success(f"Loaded **{uploaded.name}**"
                + (f" · sheet **{sheet_name}**" if sheet_name else "")
                + f" — {len(df)} rows, {len(df.columns)} columns")

    st.markdown("### 2. Variable types")
    st.caption(
        "Stream-lite guesses numerical vs. categorical from the data, but no variable is "
        "included automatically — check **Use** for each variable you want in the table, "
        "and change **Type** if 'Auto' picked the wrong one."
    )

    dataset_key = f"{uploaded.name}::{sheet_name}"
    if "var_meta" not in st.session_state or st.session_state.get("_last_file") != dataset_key:
        var_meta = {}
        total_rows = len(df)
        for col in df.columns:
            vtype, n, unique_n = detect_type(df[col])
            var_meta[col] = {
                "type": "auto",
                "detected": vtype,
                "use": False,
                "n": n,
                "missing": total_rows - n,
                "unique": unique_n,
            }
        st.session_state["var_meta"] = var_meta
        st.session_state["_last_file"] = dataset_key

    var_meta = st.session_state["var_meta"]

    editor_df = pd.DataFrame([
        {
            "Variable": col,
            "Use": var_meta[col]["use"],
            "Type": var_meta[col]["type"],
            "Auto-detected": var_meta[col]["detected"].capitalize(),
            "n (non-missing)": var_meta[col]["n"],
            "n (missing)": var_meta[col]["missing"],
            "Unique values": var_meta[col]["unique"],
        }
        for col in df.columns
    ])

    edited = st.data_editor(
        editor_df,
        column_config={
            "Use": st.column_config.CheckboxColumn(required=True),
            "Type": st.column_config.SelectboxColumn(options=["auto", "numerical", "categorical"], required=True),
            "Auto-detected": st.column_config.TextColumn(disabled=True),
            "Variable": st.column_config.TextColumn(disabled=True),
            "n (non-missing)": st.column_config.NumberColumn(disabled=True),
            "n (missing)": st.column_config.NumberColumn(disabled=True),
            "Unique values": st.column_config.NumberColumn(disabled=True),
        },
        hide_index=True,
        use_container_width=True,
        key="var_editor",
    )

    for _, row in edited.iterrows():
        var_meta[row["Variable"]]["use"] = bool(row["Use"])
        var_meta[row["Variable"]]["type"] = row["Type"]

    st.markdown("### 3. Table layout, grouping & test selection")

    layout = st.radio(
        "Table layout",
        options=["group", "factors", "custom"],
        format_func=lambda x: {
            "group": "Standard: categorical grouping variable in columns",
            "factors": "Continuous variable in columns, factors in rows",
            "custom": "Custom table: continuous variable(s) \u00D7 groups (stratified) \u00D7 factor rows",
        }[x],
        horizontal=True,
        help="Standard: pick a categorical grouping variable (e.g. Treatment) and every ticked "
             "variable is summarised per group. Continuous variable in columns: pick one continuous "
             "variable (e.g. Age) and every ticked categorical variable becomes a factor whose levels "
             "are shown as rows. Custom table: several continuous variables, up to two nested grouping "
             "variables (e.g. Sex, then Treatment) split every column, and you choose the factor rows.",
    )
    factor_mode = layout == "factors"
    custom_mode = layout == "custom"

    categorical_cols = [c for c in df.columns if effective_type(var_meta[c]) == "categorical"]
    numerical_cols = [c for c in df.columns if effective_type(var_meta[c]) == "numerical"]

    group_col = None
    outcome_col = None
    pct_mode, pct_digits, yates_correction = "column", 2, False
    can_generate = True
    factor_cols = []
    custom_cfg = {}

    def _levels_of(col):
        s = df[col].dropna().astype(str).str.strip()
        return sorted(s[s != ""].unique().tolist())

    if custom_mode:
        NONE_LBL = "\u2014 None \u2014"
        cc1, cc2 = st.columns(2)
        with cc1:
            cont_vars = st.multiselect(
                "1) Continuous variable(s) \u2014 columns",
                options=numerical_cols,
                help="One or more. Only variables whose Type is numerical are listed.",
            )
            g1 = st.selectbox("2) Grouping variable \u2014 order 1 (outer column split)",
                              [NONE_LBL] + categorical_cols)
            g1 = None if g1 == NONE_LBL else g1
            g2 = st.selectbox("Grouping variable \u2014 order 2 (inner split, optional)",
                              [NONE_LBL] + [c for c in categorical_cols if c != g1],
                              disabled=g1 is None)
            g2 = None if (g1 is None or g2 == NONE_LBL) else g2
            custom_groups = [g for g in (g1, g2) if g]

            custom_levels = {}
            for gc in custom_groups:
                lv_all = _levels_of(gc)
                custom_levels[gc] = st.multiselect(
                    f"Levels of {gc} to show as columns", options=lv_all, default=lv_all,
                    key=f"lv_{gc}_{dataset_key}",
                    help="Deselect a level to hide it from the table.",
                )

            factor_options = [c for c in categorical_cols if c not in custom_groups]
            factor_default = [c for c in factor_options if var_meta[c]["use"]]
            custom_factors = st.multiselect(
                "3) Variables / factors \u2014 rows",
                options=factor_options, default=factor_default,
                help="Each factor becomes a block of rows, one row per level. Pre-filled with the "
                     "categorical variables ticked under Use.",
            )
        with cc2:
            test_mode = st.radio(
                "4) Test selection",
                options=["auto", "parametric", "nonparametric"],
                format_func=lambda x: {"auto": "Based on normality",
                                        "parametric": "Force parametric",
                                        "nonparametric": "Force non-parametric"}[x],
                horizontal=True,
            )
            digits = st.number_input("5) Decimal places (descriptive statistics)",
                                      min_value=0, max_value=6, value=2, step=1)
            display_mode = st.radio(
                "6) Descriptive statistics display",
                options=["auto", "mean_sd", "median_iqr", "both"],
                format_func=lambda x: {"auto": "Based on normality",
                                        "mean_sd": "Mean \u00B1 SD",
                                        "median_iqr": "Median (IQR)",
                                        "both": "Both"}[x],
                horizontal=True,
            )
            compare = st.radio(
                "p-value compares",
                options=["groups", "levels", "none"],
                format_func=lambda x: {
                    "groups": "Column groups within each row (test columns on the right)",
                    "levels": "Factor levels within each column group (p-value row)",
                    "none": "No tests",
                }[x],
                help="'Column groups': e.g. Male vs Female for the row Smoking = Yes. "
                     "'Factor levels': e.g. Smoking Yes vs No within Males. If no grouping variable "
                     "is chosen, factor levels are compared.",
            )
            show_n = st.checkbox("Show n in each cell", value=True)
            alpha = st.number_input("Significance level (\u03B1)", min_value=0.001, max_value=0.5,
                                     value=0.05, step=0.01)

        if not cont_vars:
            st.warning("Select at least one continuous variable for the columns.")
            can_generate = False
        if any(len(custom_levels[g]) == 0 for g in custom_groups):
            st.warning("Keep at least one level of each grouping variable.")
            can_generate = False
        if not custom_factors:
            st.info("No factor rows selected \u2014 the table will only contain the Overall row.")
        if compare == "groups" and not custom_groups:
            st.info("No grouping variable chosen, so factor levels will be compared instead.")
        if can_generate:
            n_cols = max(1, int(np.prod([len(custom_levels[g]) for g in custom_groups]))) * len(cont_vars)
            if n_cols > 12:
                st.warning(f"This table will have {n_cols} data columns \u2014 it may be wide in Word/Excel.")
        custom_cfg = {"cont_vars": cont_vars, "groups": custom_groups, "levels": custom_levels,
                      "factors": custom_factors, "digits": int(digits), "compare": compare,
                      "show_n": show_n}
    else:
        col1, col2, col3, col4 = st.columns([2, 2, 1, 1.3])
        with col1:
            if factor_mode:
                outcome_col = st.selectbox(
                    "Continuous variable (columns)",
                    options=numerical_cols,
                    help="Only variables whose Type is numerical (auto-detected or set manually) are listed.",
                )
            else:
                group_col = st.selectbox(
                    "Grouping variable",
                    options=["\u2014 None (descriptive only) \u2014"] + categorical_cols,
                )
                group_col = None if group_col == "\u2014 None (descriptive only) \u2014" else group_col

            display_mode = st.radio(
                "Descriptive statistics display",
                options=["auto", "mean_sd", "median_iqr", "both"],
                format_func=lambda x: {"auto": "Auto (normality-based)",
                                        "mean_sd": "Mean \u00B1 SD",
                                        "median_iqr": "Median (IQR)",
                                        "both": "Both"}[x],
                horizontal=True,
            )

        with col2:
            test_mode = st.radio(
                "Numeric test selection",
                options=["auto", "parametric", "nonparametric"],
                format_func=lambda x: {"auto": "Auto (normality-based)",
                                        "parametric": "Force parametric",
                                        "nonparametric": "Force non-parametric"}[x],
                horizontal=True,
            )
            st.caption("Display and test selection are independent \u2014 e.g. you can show both mean\u00B1SD "
                        "and median (IQR) while still testing with Wilcoxon based on normality.")

            if not factor_mode:
                yates_correction = st.checkbox(
                    "Apply Yates' continuity correction (2\u00D72 chi-square)",
                    value=False,
                    help="Only affects chi-square tests on 2\u00D72 tables (ignored for larger tables and for "
                         "Fisher's exact test). Off by default \u2014 the uncorrected chi-square is generally "
                         "preferred today; Yates' correction is conservative and can reduce power.",
                )

        with col3:
            alpha = st.number_input("Significance level (\u03B1)", min_value=0.001, max_value=0.5,
                                     value=0.05, step=0.01)

        with col4:
            if not factor_mode:
                pct_mode = st.radio(
                    "Categorical % basis",
                    options=["column", "row"],
                    format_func=lambda x: "Column-wise (\u00F7 group n)" if x == "column" else "Row-wise (\u00F7 category n)",
                    disabled=not group_col,
                    help="Column-wise: each n(%) is a share of its own group's total (columns sum to ~100%). "
                         "Row-wise: each n(%) is a share of that category's total across all groups (rows sum "
                         "to ~100%). Only applies when a grouping variable is selected.",
                )
                pct_digits = st.number_input("% decimal places", min_value=0, max_value=4, value=2, step=1)

        if factor_mode:
            if not numerical_cols:
                st.warning("No numerical variables found. Set **Type** to *numerical* for the continuous "
                           "variable in the table above.")
                can_generate = False
            else:
                factor_cols = [c for c in categorical_cols if var_meta[c]["use"] and c != outcome_col]
                if not factor_cols:
                    st.warning("No factors selected \u2014 check **Use** for at least one categorical variable "
                               "in the table above. Each one becomes a block of rows.")
                    can_generate = False
                else:
                    st.info(f"Columns: **{outcome_col}** \u00B7 Factors (rows): " + ", ".join(f"**{c}**" for c in factor_cols)
                            + ". Ticked numerical variables are ignored in this layout.")
        else:
            n_selected = sum(1 for c in df.columns if var_meta[c]["use"] and c != group_col)
            if n_selected == 0:
                st.warning("No variables selected \u2014 check **Use** for at least one variable in the table above.")
                can_generate = False
            if group_col:
                levels = df[group_col].dropna().astype(str).str.strip().unique().tolist()
                counts = ", ".join(
                    f"{lv} (n={(df[group_col].astype(str).str.strip() == lv).sum()})" for lv in sorted(levels)
                )
                if len(levels) < 2:
                    st.warning(f"Groups in **{group_col}**: {counts} \u2014 need at least 2 groups to run comparisons.")
                    can_generate = False
                else:
                    st.info(f"Groups in **{group_col}**: {counts}")

    if st.button("Generate Table 1", type="primary", disabled=not can_generate):
        if custom_mode:
            cfg = custom_cfg
            hdr, crows, csv_rows, flags = build_custom_table(
                df, cfg["cont_vars"], cfg["groups"], cfg["levels"], cfg["factors"],
                test_mode, display_mode, cfg["digits"], alpha,
                compare=cfg["compare"], show_n=cfg["show_n"],
            )
            title = "Table 1. " + ", ".join(cfg["cont_vars"])
            if cfg["groups"]:
                title += " by " + " and ".join(cfg["groups"])
            eff_compare = "levels" if (cfg["compare"] == "groups" and not cfg["groups"]) else cfg["compare"]
            st.session_state["result"] = {
                "layout": "custom", "custom": {"hdr": hdr, "rows": crows},
                "header": None, "display_rows": None, "csv_rows": csv_rows, "flags": flags,
                "group_col": None, "outcome_col": None, "alpha": alpha,
                "display_mode": display_mode, "pct_mode": "column", "pct_digits": 2, "yates": False,
                "title": title,
                "footnotes": custom_footnotes(cfg["cont_vars"], cfg["groups"], display_mode, test_mode,
                                               alpha, eff_compare, cfg["show_n"], flags),
            }
        elif factor_mode:
            header, display_rows, csv_rows, flags = build_table_by_factors(
                df, var_meta, outcome_col, test_mode, alpha, display_mode,
            )
            st.session_state["result"] = {
                "layout": "factors", "header": header, "display_rows": display_rows,
                "csv_rows": csv_rows, "flags": flags, "group_col": outcome_col,
                "outcome_col": outcome_col, "alpha": alpha, "display_mode": display_mode,
                "pct_mode": "column", "pct_digits": 2, "yates": False,
                "title": f"Table 1. {outcome_col} by factor levels",
                "footnotes": factor_footnotes(display_mode, alpha, outcome_col, flags),
            }
        else:
            header, display_rows, csv_rows, flags = build_table1(
                df, var_meta, group_col, test_mode, alpha, display_mode,
                pct_mode=pct_mode, pct_digits=pct_digits, yates_correction=yates_correction,
            )
            st.session_state["result"] = {
                "layout": "group", "header": header, "display_rows": display_rows,
                "csv_rows": csv_rows, "flags": flags, "group_col": group_col,
                "outcome_col": None, "alpha": alpha, "display_mode": display_mode,
                "pct_mode": pct_mode, "pct_digits": pct_digits, "yates": yates_correction,
                "title": "Table 1. Baseline characteristics", "footnotes": None,
            }

    if "result" in st.session_state:
        res = st.session_state["result"]
        header, display_rows, csv_rows, flags = (res["header"], res["display_rows"],
                                                  res["csv_rows"], res["flags"])
        result_group_col, result_alpha = res["group_col"], res["alpha"]
        result_display_mode, result_pct_mode = res["display_mode"], res["pct_mode"]
        result_yates_correction = res["yates"]

        st.markdown(f"### {res['title']}")
        if res["layout"] == "custom":
            st.markdown(render_custom_html(res["custom"]["hdr"], res["custom"]["rows"], result_alpha),
                         unsafe_allow_html=True)
        else:
            st.markdown(render_table_markdown(header, display_rows, result_group_col, result_alpha),
                         unsafe_allow_html=True)

        if res["layout"] in ("factors", "custom"):
            footnotes = res["footnotes"]
        else:
            footnotes = [descriptive_footnote(result_display_mode, result_alpha)]
            pct_note = pct_basis_footnote(result_group_col, result_pct_mode)
            if pct_note:
                footnotes.append(pct_note)
            if "catpct" in flags:
                footnotes.append("Percentages for categorical variables are calculated among non-missing "
                                  "responses for that variable (missing values excluded from the denominator).")
            if "chi2" in flags:
                chi2_note = "Chi-square test of independence used for categorical comparisons with adequate expected cell counts"
                chi2_note += " (Yates' continuity correction applied to 2\u00D72 tables)." if result_yates_correction \
                    else " (no continuity correction applied)."
                footnotes.append(chi2_note)
            if "fisher" in flags:
                footnotes.append("Fisher's exact test used in place of chi-square when a 2\u00D72 table had an expected cell count below 5.")
            if "lowE" in flags:
                footnotes.append("Caution: one or more categorical comparisons above have expected cell counts below 5; chi-square approximation may be unreliable.")
            if "skipped" in flags:
                footnotes.append("Note: a statistical test could not be computed for one or more variables (e.g. insufficient data in a group) and was left blank.")
            footnotes.append(f"Bold p-values indicate statistical significance at \u03B1={result_alpha}.")
        st.caption("  \n".join(footnotes))

        dl_col1, dl_col2, dl_col3 = st.columns(3)

        with dl_col1:
            csv_buf = io.StringIO()
            pd.DataFrame(csv_rows).to_csv(csv_buf, index=False, header=False)
            st.download_button("Download CSV", csv_buf.getvalue(), file_name="table1.csv", mime="text/csv")

        with dl_col2:
            if res["layout"] == "custom":
                excel_bytes = build_excel_custom(res["custom"]["hdr"], res["custom"]["rows"])
            else:
                excel_bytes = build_excel(csv_rows, sheet_name="Table 1")
            st.download_button("Download Excel", excel_bytes, file_name="table1.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with dl_col3:
            if res["layout"] == "custom":
                docx_bytes = build_docx_custom(res["custom"]["hdr"], res["custom"]["rows"],
                                                footnotes, res["title"])
            else:
                docx_bytes = build_docx(header, display_rows, result_group_col, result_alpha, flags,
                                         result_display_mode, pct_mode=result_pct_mode,
                                         yates_correction=result_yates_correction, title=res["title"],
                                         footnotes_override=footnotes if res["layout"] == "factors" else None)
            st.download_button("Download Word", docx_bytes, file_name="table1.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

else:
    st.info("Upload a file to get started. Nothing leaves your machine \u2014 the app runs locally.")
