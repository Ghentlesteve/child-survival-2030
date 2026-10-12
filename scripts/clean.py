"""Clean UN IGME / MMEIG / WPP estimates and World Bank groups into tidy tables for Power BI.

Run from the project root:  python scripts/clean.py
Inputs  (data/raw/):  unicef_mortality.csv, wb_country_groups.xlsx
Outputs (data/clean/): areas.csv, indicators.csv, estimates.csv, progress.csv, years.csv
"""
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW, CLEAN = ROOT / "data" / "raw", ROOT / "data" / "clean"
CLEAN.mkdir(parents=True, exist_ok=True)

BASE_YEAR, TARGET_YEAR = 2015, 2030

# indicator_code: (short name, unit, per, SDG target, target label, secondary threshold, order)
INDICATORS = {
    "CME_MRY0T4": ("Under-5 mortality rate", "deaths per 1,000 live births", 1000, 25, "SDG 3.2: 25 or less by 2030", None, 1),
    "CME_MRM0":   ("Neonatal mortality rate", "deaths per 1,000 live births", 1000, 12, "SDG 3.2: 12 or less by 2030", None, 2),
    "MNCH_MMR":   ("Maternal mortality ratio", "deaths per 100,000 live births", 100000, 70, "SDG 3.1: global ratio below 70 by 2030", 140, 3),
    "CME_MRY0":   ("Infant mortality rate", "deaths per 1,000 live births", 1000, None, "No SDG target", None, 4),
    "CME_SBR":    ("Stillbirth rate", "stillbirths per 1,000 total births", 1000, 12, "Every Newborn Action Plan: 12 or less by 2030", None, 5),
    "CME_TMY0T4": ("Under-5 deaths", "deaths", 1, None, "", None, 6),
    "CME_TMM0":   ("Neonatal deaths", "deaths", 1, None, "", None, 7),
    "DM_BRTS":    ("Births", "live births", 1, None, "", None, 8),
    "DM_POP_TOT": ("Population", "people", 1, None, "", None, 9),
}
TARGETED = ["CME_MRY0T4", "CME_MRM0", "MNCH_MMR"]

# Official aggregates kept from the UNICEF file (as computed by UN IGME / MMEIG). The World Bank regional rows
# in the download carry only births and population, so regions are UNICEF reporting regions.
AGGREGATES = {
    "WORLD": ("World", "World"),
    "UNICEF_SSA": ("Sub-Saharan Africa", "Region"), "UNICEF_SA": ("South Asia", "Region"),
    "UNICEF_MENA": ("Middle East and North Africa", "Region"), "UNICEF_EAP": ("East Asia and Pacific", "Region"),
    "UNICEF_ECA": ("Europe and Central Asia", "Region"), "UNICEF_LAC": ("Latin America and the Caribbean", "Region"),
    "UNICEF_NA": ("North America", "Region"),
    "WB_LI": ("Low income", "Income group"), "WB_LMI": ("Lower middle income", "Income group"),
    "WB_UMI": ("Upper middle income", "Income group"), "WB_HI": ("High income", "Income group"),
}
# The global maternal mortality ratio is published under the UNICEF global code, not WORLD.
WORLD_ALIASES = {"UNICEF_REP_REG_GLOBAL": "WORLD"}


def main():
    raw = pd.read_csv(RAW / "unicef_mortality.csv", low_memory=False)
    n_raw = len(raw)
    d = raw[(raw.SEX == "_T") & raw.INDICATOR.isin(INDICATORS)].copy()
    d = d.rename(columns={"REF_AREA": "area_code", "Geographic area": "area_name", "INDICATOR": "indicator_code",
                          "TIME_PERIOD": "year", "OBS_VALUE": "value", "LOWER_BOUND": "lower", "UPPER_BOUND": "upper"})
    # Births and population are published in thousands.
    mult = 10.0 ** d.UNIT_MULTIPLIER.fillna(0)
    for c in ["value", "lower", "upper"]:
        d[c] = d[c] * mult

    for src, dst in WORLD_ALIASES.items():
        extra = d[(d.area_code == src) & ~d.set_index(["indicator_code", "year"]).index.isin(
            d[d.area_code == dst].set_index(["indicator_code", "year"]).index)].copy()
        extra["area_code"] = dst
        d = pd.concat([d, extra], ignore_index=True)

    # ------------------------------------------------------------- areas
    groups = pd.read_excel(RAW / "wb_country_groups.xlsx", sheet_name="List of economies")
    groups = groups.rename(columns={"Code": "area_code", "Region": "wb_region", "Income group": "income_group",
                                    "Lending category": "lending_category"})
    countries = set(d.loc[(d.indicator_code == "CME_MRY0T4") & d.area_code.str.fullmatch("[A-Z]{3}"), "area_code"])
    names = d.drop_duplicates("area_code").set_index("area_code").area_name
    areas = pd.DataFrame({"area_code": sorted(countries)})
    areas["area_name"] = areas.area_code.map(names)
    areas["area_type"] = "Country"
    areas = areas.merge(groups[["area_code", "wb_region", "income_group", "lending_category"]], on="area_code", how="left")
    areas["wb_region"] = areas.wb_region.fillna("Not classified by the World Bank")
    areas["income_group"] = areas.income_group.fillna("Not classified")
    agg = pd.DataFrame([{"area_code": k, "area_name": v[0], "area_type": v[1]} for k, v in AGGREGATES.items()])
    areas = pd.concat([areas, agg], ignore_index=True)
    areas["is_country"] = areas.area_type == "Country"
    areas["sort_order"] = np.where(areas.is_country, 100, areas.area_type.map({"World": 0, "Region": 10, "Income group": 20}))
    areas = areas.sort_values(["sort_order", "area_name"]).reset_index(drop=True)
    areas.to_csv(CLEAN / "areas.csv", index=False)

    # ------------------------------------------------------------- estimates
    est = d[d.area_code.isin(areas.area_code)][["area_code", "indicator_code", "year", "value", "lower", "upper"]]
    est = est.drop_duplicates(["area_code", "indicator_code", "year"]).sort_values(["area_code", "indicator_code", "year"])
    assert not est.duplicated(["area_code", "indicator_code", "year"]).any()
    est.to_csv(CLEAN / "estimates.csv", index=False, float_format="%.4f")

    # ------------------------------------------------------------- indicators
    latest_year = est.groupby("indicator_code").year.max()
    ind = pd.DataFrame([{"indicator_code": k, "indicator": v[0], "unit": v[1], "per": v[2], "sdg_target": v[3],
                         "target_label": v[4], "secondary_threshold": v[5], "sort_order": v[6],
                         "has_target": v[3] is not None, "latest_year": int(latest_year.get(k, 0)),
                         "has_uncertainty": bool(est.loc[est.indicator_code == k, "lower"].notna().any())}
                        for k, v in INDICATORS.items()])
    ind.to_csv(CLEAN / "indicators.csv", index=False)

    # ------------------------------------------------------------- progress toward 2030
    births = est[est.indicator_code == "DM_BRTS"].set_index(["area_code", "year"]).value
    rows = []
    for code in TARGETED:
        target = INDICATORS[code][3]
        sub = est[est.indicator_code == code]
        for area, g in sub.groupby("area_code"):
            g = g.set_index("year")
            ly = int(g.index.max())
            latest = g.loc[ly]
            if BASE_YEAR not in g.index:
                continue
            base = g.loc[BASE_YEAR, "value"]
            v = latest.value
            arr = np.log(base / v) / (ly - BASE_YEAR) if base > 0 and v > 0 else np.nan
            if v <= target:
                req, status = 0.0, "Target met"
            else:
                req = np.log(v / target) / (TARGET_YEAR - ly)
                if arr >= req:
                    status = "On track"
                elif arr > 0:
                    status = "Acceleration needed"
                else:
                    status = "No progress"
            proj = v * np.exp(-arr * (TARGET_YEAR - ly)) if pd.notna(arr) else np.nan
            if v <= target:
                year_reached = np.nan
            elif arr > 0:
                year_reached = ly + np.log(v / target) / arr
            else:
                year_reached = np.nan
            b = births.get((area, min(ly, int(latest_year["DM_BRTS"]))), np.nan)
            # Crisis flag (child mortality only): UN IGME includes deaths from conflicts and disasters, so a
            # year-on-year jump of more than 20% between 2014 and the latest year distorts progress since 2015.
            crisis = False
            if code.startswith("CME_"):
                window = g.loc[BASE_YEAR - 1:ly, "value"]
                crisis = bool((window.pct_change() > 0.2).any())
            if v <= target:
                when = "Already met"
            elif pd.isna(year_reached):
                when = "Not at current pace"
            elif year_reached <= TARGET_YEAR:
                when = "By 2030"
            elif year_reached <= 2050:
                when = "2031-2050"
            elif year_reached <= 2100:
                when = "2051-2100"
            else:
                when = "After 2100"
            per = INDICATORS[code][2]
            rows.append({
                "area_code": area, "indicator_code": code, "base_year": BASE_YEAR, "base_value": base,
                "latest_year": ly, "latest_value": v, "latest_lower": latest.lower, "latest_upper": latest.upper,
                "target": target, "gap_to_target": max(v - target, 0),
                "arr_achieved": arr, "arr_required": req, "status": status,
                "status_order": {"Target met": 1, "On track": 2, "Acceleration needed": 3, "No progress": 4}[status],
                "projected_2030": proj,
                "year_target_reached": year_reached if pd.notna(year_reached) and year_reached <= 2100 else np.nan,
                "when_target_reached": when,
                "crisis_in_period": crisis,
                "births_latest": b,
                # deaths each year that would not occur if the rate were at the target (an illustration)
                "excess_deaths_vs_target": max(v - target, 0) * b / per if pd.notna(b) else np.nan,
                "above_secondary": bool(INDICATORS[code][5] and v > INDICATORS[code][5]),
            })
    prog = pd.DataFrame(rows)
    prog.to_csv(CLEAN / "progress.csv", index=False, float_format="%.6f")

    # ------------------------------------------------------------- years
    years = pd.DataFrame({"year": range(1990, TARGET_YEAR + 1)})
    years["period"] = np.select([years.year < 2000, years.year < BASE_YEAR, years.year <= 2024],
                                ["1990s", "MDG era (2000-2014)", "SDG era (2015+)"], default="Projection")
    years.to_csv(CLEAN / "years.csv", index=False)

    # ------------------------------------------------------------- report
    print(f"raw rows {n_raw:,}; estimates {len(est):,}; areas {len(areas)} "
          f"({int(areas.is_country.sum())} countries); progress rows {len(prog)}")
    c = prog[prog.area_code.isin(areas.loc[areas.is_country, "area_code"])]
    print(c.pivot_table(index="status", columns="indicator_code", values="area_code", aggfunc="count").to_string())


if __name__ == "__main__":
    main()
