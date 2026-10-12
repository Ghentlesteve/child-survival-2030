"""Generate the Power BI semantic model (PBIP / TMDL) for Child Survival 2030.

Run from the project root with Power BI Desktop closed:  python scripts/make_model.py
Writes powerbi/child-survival-2030.SemanticModel/ and powerbi/child-survival-2030.pbip.
Open the .pbip, then Refresh to load data/clean/.
"""
import json
import uuid
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
CLEAN = ROOT / "data" / "clean"
PBI = ROOT / "powerbi"
NAME = "child-survival-2030"
SM = PBI / f"{NAME}.SemanticModel"
DEF = SM / "definition"
WIN_FOLDER = r"C:\Users\Kwenev Steve\Documents\Web\Advanced Dashboards\child-survival-2030\data\clean" + "\\"
T = "\t"

TABLES = ["areas", "indicators", "years", "estimates", "progress"]
HIDDEN_COLS = {"estimates": {"area_code", "indicator_code", "year"}, "progress": {"area_code", "indicator_code"}}
SORT_BY = {("progress", "status"): "status_order", ("areas", "area_type"): "sort_order",
           ("indicators", "indicator"): "sort_order"}
FORMATS = {"value": "#,0.0", "lower": "#,0.0", "upper": "#,0.0", "arr_achieved": "0.0%", "arr_required": "0.0%",
           "year": "0", "latest_year": "0", "base_year": "0", "year_target_reached": "0"}


def tag():
    return str(uuid.uuid4())


def col_types(df):
    out = {}
    for c, dt in df.dtypes.items():
        if dt == bool:
            out[c] = ("boolean", "type logical")
        elif pd.api.types.is_integer_dtype(dt):
            out[c] = ("int64", "Int64.Type")
        elif pd.api.types.is_float_dtype(dt):
            out[c] = ("double", "type number")
        else:
            out[c] = ("string", "type text")
    return out


# ------------------------------------------------------------------ measures
# (table, name, expression, format)
M = []


def m(table, name, expr, fmt=None, folder=None):
    M.append((table, name, " ".join(expr.split()), fmt, folder))


# core values (one area, one indicator)
m("estimates", "Estimate", "IF ( HASONEVALUE ( indicators[indicator_code] ) && HASONEVALUE ( areas[area_code] ), SUM ( estimates[value] ) )", "#,0.0", "Values")
m("estimates", "Lower Bound", "IF ( HASONEVALUE ( indicators[indicator_code] ) && HASONEVALUE ( areas[area_code] ), SUM ( estimates[lower] ) )", "#,0.0", "Values")
m("estimates", "Upper Bound", "IF ( HASONEVALUE ( indicators[indicator_code] ) && HASONEVALUE ( areas[area_code] ), SUM ( estimates[upper] ) )", "#,0.0", "Values")
m("estimates", "Latest Year", "MAX ( indicators[latest_year] )", "0", "Values")
m("estimates", "Latest Value", "VAR ly = MAX ( indicators[latest_year] ) RETURN CALCULATE ( [Estimate], REMOVEFILTERS ( years ), years[year] = ly )", "#,0.0", "Values")
m("estimates", "SDG Target", "IF ( HASONEVALUE ( indicators[indicator_code] ), MAX ( indicators[sdg_target] ) )", "#,0", "Values")
m("estimates", "Profile Title", "SELECTEDVALUE ( areas[area_name], \"Right-click a country on any page, then Drill through\" )", None, "Labels")
m("estimates", "Profile Subtitle", "IF ( HASONEVALUE ( areas[area_code] ), SELECTEDVALUE ( areas[wb_region] ) & \"  ·  \" & SELECTEDVALUE ( areas[income_group] ) & \"  ·  \" & FORMAT ( [Selected Area Deaths Share], \"0.0%\" ) & \" of the world's under-5 deaths in \" & FORMAT ( CALCULATE ( MAX ( indicators[latest_year] ), REMOVEFILTERS ( indicators ), indicators[indicator_code] = \"CME_TMY0T4\" ), \"0\" ) )", None, "Labels")
m("estimates", "Indicator Title", "SELECTEDVALUE ( indicators[indicator], \"Select an indicator\" ) & \", \" & SELECTEDVALUE ( indicators[unit] )", None, "Labels")
m("estimates", "Target Label", "SELECTEDVALUE ( indicators[target_label] )", None, "Labels")

# world headline numbers
W = "REMOVEFILTERS ( areas ), REMOVEFILTERS ( indicators ), areas[area_code] = \"WORLD\""
m("estimates", "World Under-5 Deaths", f"CALCULATE ( [Latest Value], {W}, indicators[indicator_code] = \"CME_TMY0T4\" )", "#,0", "Headlines")
m("estimates", "World Neonatal Deaths", f"CALCULATE ( [Latest Value], {W}, indicators[indicator_code] = \"CME_TMM0\" )", "#,0", "Headlines")
m("estimates", "World Under-5 Rate", f"CALCULATE ( [Latest Value], {W}, indicators[indicator_code] = \"CME_MRY0T4\" )", "#,0.0", "Headlines")
m("estimates", "World Under-5 Rate 1990", f"CALCULATE ( [Estimate], {W}, indicators[indicator_code] = \"CME_MRY0T4\", REMOVEFILTERS ( years ), years[year] = 1990 )", "#,0.0", "Headlines")
m("estimates", "World Under-5 Rate Change Since 1990", "DIVIDE ( [World Under-5 Rate] - [World Under-5 Rate 1990], [World Under-5 Rate 1990] )", "+0%;-0%;0%", "Headlines")
m("estimates", "Nigeria Under-5 Deaths", "CALCULATE ( [Latest Value], REMOVEFILTERS ( areas ), REMOVEFILTERS ( indicators ), areas[area_code] = \"NGA\", indicators[indicator_code] = \"CME_TMY0T4\" )", "#,0", "Headlines")
m("estimates", "Nigeria Share of World Under-5 Deaths", "DIVIDE ( [Nigeria Under-5 Deaths], [World Under-5 Deaths] )", "0%", "Headlines")
m("estimates", "Selected Area Deaths Share", "DIVIDE ( CALCULATE ( [Latest Value], REMOVEFILTERS ( indicators ), indicators[indicator_code] = \"CME_TMY0T4\" ), [World Under-5 Deaths] )", "0.0%", "Headlines")

# progress toward 2030 (from progress table)
m("progress", "Countries", "CALCULATE ( COUNTROWS ( progress ), areas[is_country] = TRUE () )", "#,0", "Progress")
m("progress", "Countries Off Track", "CALCULATE ( [Countries], progress[status] IN { \"Acceleration needed\", \"No progress\" } )", "#,0", "Progress")
m("progress", "Countries Met or On Track", "CALCULATE ( [Countries], progress[status] IN { \"Target met\", \"On track\" } )", "#,0", "Progress")
m("progress", "Share Off Track", "DIVIDE ( [Countries Off Track], [Countries] )", "0%", "Progress")
m("progress", "Progress Status", "SELECTEDVALUE ( progress[status] )", None, "Progress")
m("progress", "When Target Reached", "SELECTEDVALUE ( progress[when_target_reached] )", None, "Progress")
m("progress", "Base Value 2015", "IF ( HASONEVALUE ( areas[area_code] ), MAX ( progress[base_value] ) )", "#,0.0", "Progress")
m("progress", "Latest Rate", "IF ( HASONEVALUE ( areas[area_code] ), MAX ( progress[latest_value] ) )", "#,0.0", "Progress")
m("progress", "Change Since 2015", "DIVIDE ( [Latest Rate] - [Base Value 2015], [Base Value 2015] )", "+0%;-0%;0%", "Progress")
m("progress", "ARR Achieved", "IF ( HASONEVALUE ( areas[area_code] ), MAX ( progress[arr_achieved] ) )", "0.0%", "Progress")
m("progress", "ARR Required", "IF ( HASONEVALUE ( areas[area_code] ), MAX ( progress[arr_required] ) )", "0.0%", "Progress")
m("progress", "Speed-up Needed", "VAR a = [ARR Achieved] VAR r = [ARR Required] RETURN IF ( r > 0 && a > 0, DIVIDE ( r, a ) )", "0.0\"x\"", "Progress")
m("progress", "Gap to Target", "IF ( HASONEVALUE ( areas[area_code] ), MAX ( progress[gap_to_target] ) )", "#,0.0", "Progress")
m("progress", "Projected 2030", "IF ( HASONEVALUE ( areas[area_code] ), MAX ( progress[projected_2030] ) )", "#,0.0", "Progress")
m("progress", "Deaths Above Target", "CALCULATE ( SUM ( progress[excess_deaths_vs_target] ), areas[is_country] = TRUE () )", "#,0", "Progress")
m("progress", "Crisis Note", "IF ( HASONEVALUE ( areas[area_code] ) && CALCULATE ( COUNTROWS ( progress ), progress[crisis_in_period] = TRUE () ) > 0, \"Crisis years since 2015: progress figure less reliable\" )", None, "Progress")
m("progress", "Status Colour",
  "SWITCH ( [Progress Status], \"Target met\", \"#0CA30C\", \"On track\", \"#6DA7EC\", \"Acceleration needed\", \"#EC835A\", \"No progress\", \"#D03B3B\" )",
  None, "Progress")

# scenarios (one area, one indicator)
m("progress", "Scenario ARR Value", "SELECTEDVALUE ( 'Scenario ARR'[Scenario ARR], 5 ) / 100", "0.0%", "Scenario")
m("progress", "Peer ARR",
  "VAR p = SELECTEDVALUE ( 'Peer Country'[peer_code] ) RETURN IF ( NOT ISBLANK ( p ), CALCULATE ( MAX ( progress[arr_achieved] ), REMOVEFILTERS ( areas ), progress[area_code] = p ) )",
  "0.0%", "Scenario")
PATH = ("VAR y = MAX ( years[year] ) VAR ly = CALCULATE ( MAX ( progress[latest_year] ), REMOVEFILTERS ( years ) ) "
        "VAR lv = CALCULATE ( MAX ( progress[latest_value] ), REMOVEFILTERS ( years ) ) VAR a = {a} "
        "RETURN IF ( HASONEVALUE ( areas[area_code] ) && y >= ly && NOT ISBLANK ( a ), lv * EXP ( - a * ( y - ly ) ) )")
m("progress", "Path at Current Pace", PATH.format(a="CALCULATE ( MAX ( progress[arr_achieved] ), REMOVEFILTERS ( years ) )"), "#,0.0", "Scenario")
m("progress", "Path at Required Pace", PATH.format(a="CALCULATE ( MAX ( progress[arr_required] ), REMOVEFILTERS ( years ) )"), "#,0.0", "Scenario")
m("progress", "Path at Chosen Pace", PATH.format(a="[Scenario ARR Value]"), "#,0.0", "Scenario")
m("progress", "Path at Peer Pace", PATH.format(a="[Peer ARR]"), "#,0.0", "Scenario")
m("progress", "Target Line", "VAR y = MAX ( years[year] ) RETURN IF ( y >= 2000, [SDG Target] )", "#,0", "Scenario")
REACH = ("VAR ly = MAX ( progress[latest_year] ) VAR lv = MAX ( progress[latest_value] ) VAR t = MAX ( progress[target] ) VAR a = {a} "
         "RETURN IF ( HASONEVALUE ( areas[area_code] ), IF ( lv <= t, \"Already met\", IF ( a <= 0, \"Not at this pace\", "
         "VAR yr = ly + LN ( lv / t ) / a RETURN IF ( yr > 2100, \"After 2100\", FORMAT ( ROUNDUP ( yr, 0 ), \"0\" ) ) ) ) )")
m("progress", "Year Reached at Chosen Pace", REACH.format(a="[Scenario ARR Value]"), None, "Scenario")
m("progress", "Year Reached at Current Pace", REACH.format(a="MAX ( progress[arr_achieved] )"), None, "Scenario")
m("progress", "Year Reached at Peer Pace", REACH.format(a="[Peer ARR]"), None, "Scenario")
m("progress", "Lives Saved by 2030 at Chosen Pace",
  "VAR ly = MAX ( progress[latest_year] ) VAR lv = MAX ( progress[latest_value] ) VAR ac = MAX ( progress[arr_achieved] ) "
  "VAR s = [Scenario ARR Value] VAR b = MAX ( progress[births_latest] ) VAR per = MAX ( indicators[per] ) "
  "RETURN IF ( HASONEVALUE ( areas[area_code] ) && HASONEVALUE ( indicators[indicator_code] ), "
  "SUMX ( FILTER ( ALL ( years[year] ), years[year] > ly && years[year] <= 2030 ), VAR k = years[year] - ly RETURN ( lv * EXP ( - ac * k ) - lv * EXP ( - s * k ) ) * b / per ) )",
  "#,0", "Scenario")
m("progress", "Scenario Title",
  "SELECTEDVALUE ( areas[area_name] ) & \": \" & SELECTEDVALUE ( indicators[indicator] ) & \" to 2030 at different rates of reduction\"",
  None, "Labels")

# ------------------------------------------------------------------ writers


def table_tmdl(name, df):
    types = col_types(df)
    lines = [f"table {name}", f"{T}lineageTag: {tag()}", ""]
    for (t, mname, expr, fmt, folder) in M:
        if t != name:
            continue
        lines.append(f"{T}measure '{mname}' = {expr}")
        if fmt:
            lines.append(f"{T}{T}formatString: {fmt}")
        if folder:
            lines.append(f"{T}{T}displayFolder: {folder}")
        lines.append(f"{T}{T}lineageTag: {tag()}")
        lines.append("")
    for c, (dt, _) in types.items():
        lines.append(f"{T}column {c}")
        lines.append(f"{T}{T}dataType: {dt}")
        if c in FORMATS and dt != "string":
            lines.append(f"{T}{T}formatString: {FORMATS[c]}")
        if c in HIDDEN_COLS.get(name, set()) or c.endswith("_order"):
            lines.append(f"{T}{T}isHidden")
        lines.append(f"{T}{T}lineageTag: {tag()}")
        lines.append(f"{T}{T}summarizeBy: none")
        lines.append(f"{T}{T}sourceColumn: {c}")
        if (name, c) in SORT_BY:
            lines.append(f"{T}{T}sortByColumn: {SORT_BY[(name, c)]}")
        lines.append("")
    typed = ", ".join(f'{{"{c}", {mt}}}' for c, (_, mt) in types.items())
    lines += [f"{T}partition {name} = m", f"{T}{T}mode: import", f"{T}{T}source =",
              f"{T}{T}{T}{T}let",
              f'{T}{T}{T}{T}{T}Source = Csv.Document(File.Contents(DataFolder & "{name}.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),',
              f"{T}{T}{T}{T}{T}Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),",
              f'{T}{T}{T}{T}{T}Typed = Table.TransformColumnTypes(Promoted, {{{typed}}}, "en-US")',
              f"{T}{T}{T}{T}in", f"{T}{T}{T}{T}{T}Typed", ""]
    return "\n".join(lines)


def calc_group():
    ly = "VAR ly = MAX ( indicators[latest_year] )"
    items = [
        ("1990", "CALCULATE ( SELECTEDMEASURE (), REMOVEFILTERS ( years ), years[year] = 1990 )", None),
        ("2015", "CALCULATE ( SELECTEDMEASURE (), REMOVEFILTERS ( years ), years[year] = 2015 )", None),
        ("Latest", f"{ly} RETURN CALCULATE ( SELECTEDMEASURE (), REMOVEFILTERS ( years ), years[year] = ly )", None),
        ("Change since 1990", f"{ly} VAR a = CALCULATE ( SELECTEDMEASURE (), REMOVEFILTERS ( years ), years[year] = 1990 ) "
                              "VAR b = CALCULATE ( SELECTEDMEASURE (), REMOVEFILTERS ( years ), years[year] = ly ) RETURN DIVIDE ( b - a, a )", "\"+0%;-0%;0%\""),
        ("Change since 2015", f"{ly} VAR a = CALCULATE ( SELECTEDMEASURE (), REMOVEFILTERS ( years ), years[year] = 2015 ) "
                              "VAR b = CALCULATE ( SELECTEDMEASURE (), REMOVEFILTERS ( years ), years[year] = ly ) RETURN DIVIDE ( b - a, a )", "\"+0%;-0%;0%\""),
        ("Yearly reduction since 2015", f"{ly} VAR a = CALCULATE ( SELECTEDMEASURE (), REMOVEFILTERS ( years ), years[year] = 2015 ) "
                                        "VAR b = CALCULATE ( SELECTEDMEASURE (), REMOVEFILTERS ( years ), years[year] = ly ) "
                                        "RETURN IF ( a > 0 && b > 0, LN ( a / b ) / ( ly - 2015 ) )", "\"0.0%\""),
    ]
    lines = ["table 'Time View'", f"{T}lineageTag: {tag()}", "", f"{T}calculationGroup", f"{T}{T}precedence: 1", ""]
    for i, (n, e, f) in enumerate(items):
        lines.append(f"{T}{T}calculationItem '{n}' = {e}")
        lines.append(f"{T}{T}{T}ordinal: {i}")
        if f:
            lines.append(f"{T}{T}{T}formatStringDefinition = {f}")
        lines.append("")
    lines += [f"{T}column 'Time View'", f"{T}{T}dataType: string", f"{T}{T}lineageTag: {tag()}",
              f"{T}{T}summarizeBy: none", f"{T}{T}sourceColumn: Name", f"{T}{T}sortByColumn: Ordinal", "",
              f"{T}column Ordinal", f"{T}{T}dataType: int64", f"{T}{T}formatString: 0", f"{T}{T}isHidden",
              f"{T}{T}lineageTag: {tag()}", f"{T}{T}summarizeBy: sum", f"{T}{T}sourceColumn: Ordinal", "",
              f"{T}partition 'Time View' = calculationGroup", ""]
    return "\n".join(lines)


def field_parameter():
    rows = [("Latest rate", "progress", "Latest Rate"), ("Change since 2015", "progress", "Change Since 2015"),
            ("Yearly reduction achieved", "progress", "ARR Achieved"), ("Yearly reduction needed", "progress", "ARR Required"),
            ("Speed-up needed", "progress", "Speed-up Needed"), ("Gap to target", "progress", "Gap to Target"),
            ("Deaths above target", "progress", "Deaths Above Target")]
    src = ", ".join(f'("{lbl}", NAMEOF(\'{t}\'[{mn}]), {i})' for i, (lbl, t, mn) in enumerate(rows))
    return "\n".join([
        "table 'Ranking Metric'", f"{T}lineageTag: {tag()}", "",
        f"{T}column 'Ranking Metric'", f"{T}{T}dataType: string", f"{T}{T}lineageTag: {tag()}", f"{T}{T}summarizeBy: none",
        f"{T}{T}sourceColumn: [Value1]", f"{T}{T}sortByColumn: 'Ranking Metric Order'", "",
        f"{T}{T}relatedColumnDetails", f"{T}{T}{T}groupByColumn: 'Ranking Metric Fields'", "",
        f"{T}column 'Ranking Metric Fields'", f"{T}{T}dataType: string", f"{T}{T}isHidden", f"{T}{T}lineageTag: {tag()}",
        f"{T}{T}summarizeBy: none", f"{T}{T}sourceColumn: [Value2]", "",
        f"{T}{T}extendedProperty ParameterMetadata =", f"{T}{T}{T}{T}{{", f'{T}{T}{T}{T}  "version": 3,', f'{T}{T}{T}{T}  "kind": 2',
        f"{T}{T}{T}{T}}}", "",
        f"{T}column 'Ranking Metric Order'", f"{T}{T}dataType: int64", f"{T}{T}isHidden", f"{T}{T}formatString: 0",
        f"{T}{T}lineageTag: {tag()}", f"{T}{T}summarizeBy: sum", f"{T}{T}sourceColumn: [Value3]", "",
        f"{T}partition 'Ranking Metric' = calculated", f"{T}{T}mode: import", f"{T}{T}source = {{{src}}}", "",
        f"{T}annotation PBI_Id = {uuid.uuid4().hex}", ""])


def calc_table(name, cols, source):
    lines = [f"table '{name}'", f"{T}lineageTag: {tag()}", ""]
    for c, dt, fmt, hidden in cols:
        lines += [f"{T}column '{c}'", f"{T}{T}dataType: {dt}"]
        if fmt:
            lines.append(f"{T}{T}formatString: {fmt}")
        if hidden:
            lines.append(f"{T}{T}isHidden")
        lines += [f"{T}{T}lineageTag: {tag()}", f"{T}{T}summarizeBy: none", f"{T}{T}sourceColumn: [{c}]", ""]
    lines += [f"{T}partition '{name}' = calculated", f"{T}{T}mode: import", f"{T}{T}source = {source}", ""]
    return "\n".join(lines)


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def main():
    dfs = {t: pd.read_csv(CLEAN / f"{t}.csv", keep_default_na=True) for t in TABLES}
    for t in (DEF / "tables").glob("*.tmdl") if (DEF / "tables").exists() else []:
        t.write_text("", encoding="utf-8")  # cleared; overwritten below (files cannot be deleted here)
    for t, df in dfs.items():
        write(DEF / "tables" / f"{t}.tmdl", table_tmdl(t, df))
    write(DEF / "tables" / "Time View.tmdl", calc_group())
    write(DEF / "tables" / "Ranking Metric.tmdl", field_parameter())
    write(DEF / "tables" / "Scenario ARR.tmdl", calc_table(
        "Scenario ARR", [("Scenario ARR", "double", "0.0\\%", False)],
        "SELECTCOLUMNS ( GENERATESERIES ( 0, 30, 0.5 ), \"Scenario ARR\", [Value] )"))
    write(DEF / "tables" / "Peer Country.tmdl", calc_table(
        "Peer Country", [("Peer Country", "string", None, False), ("peer_code", "string", None, True)],
        "SELECTCOLUMNS ( FILTER ( areas, areas[is_country] = TRUE () ), \"Peer Country\", areas[area_name], \"peer_code\", areas[area_code] )"))

    rels = [("estimates", "area_code", "areas", "area_code"), ("estimates", "indicator_code", "indicators", "indicator_code"),
            ("estimates", "year", "years", "year"), ("progress", "area_code", "areas", "area_code"),
            ("progress", "indicator_code", "indicators", "indicator_code")]
    write(DEF / "relationships.tmdl", "\n".join(
        f"relationship {uuid.uuid4()}\n{T}fromColumn: {a}.{b}\n{T}toColumn: {c}.{d}\n" for a, b, c, d in rels))
    write(DEF / "expressions.tmdl",
          f'expression DataFolder = "{WIN_FOLDER}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
          f"{T}lineageTag: {tag()}\n\n{T}annotation PBI_ResultType = Text\n")
    write(DEF / "database.tmdl", "database\n\tcompatibilityLevel: 1601\n")
    all_tables = TABLES + ["Time View", "Ranking Metric", "Scenario ARR", "Peer Country"]
    refs = "\n".join(f"ref table '{t}'" if " " in t else f"ref table {t}" for t in all_tables)
    write(DEF / "model.tmdl", "\n".join([
        "model Model", f"{T}culture: en-US", f"{T}defaultPowerBIDataSourceVersion: powerBI_V3",
        f"{T}discourageImplicitMeasures", f"{T}sourceQueryCulture: en-US", f"{T}dataAccessOptions",
        f"{T}{T}legacyRedirects", f"{T}{T}returnErrorValuesAsNull", "",
        f"annotation PBI_QueryOrder = {json.dumps(TABLES + ['DataFolder'])}", "",
        "annotation __PBI_TimeIntelligenceEnabled = 0", "",
        'annotation PBI_ProTooling = ["TMDLView_Desktop","DaxQueryView_Desktop","DevMode"]', "",
        refs, ""]))
    write(SM / "definition.pbism", json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.1", "settings": {}}, indent=2))
    plat = SM / ".platform"
    if not plat.exists():
        write(plat, json.dumps({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": "SemanticModel", "displayName": NAME},
            "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}}, indent=2))
    write(PBI / f"{NAME}.pbip", json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True}}, indent=2))
    write(PBI / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
    print(f"model written: {len(all_tables)} tables, {len(M)} measures")


if __name__ == "__main__":
    main()
