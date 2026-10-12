"""Generate the Power BI report (legacy report.json), theme and report folder for Child Survival 2030.

Run from the project root with Power BI Desktop closed:  python scripts/make_report.py
Then open powerbi/child-survival-2030.pbip.
"""
import json
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
NAME = "child-survival-2030"
REPORT = ROOT / "powerbi" / f"{NAME}.Report"
THEME_NAME = "ChildSurvival2030.json"

# ---------------------------------------------------------------- palette
BLUE, ORANGE, AQUA, YELLOW, MAGENTA, GREEN, VIOLET, RED = (
    "#2A78D6", "#EB6834", "#1BAF7A", "#EDA100", "#E87BA4", "#008300", "#4A3AA7", "#E34948")
PAGE_BG, CARD_BG, BORDER = "#F4F3EF", "#FFFFFF", "#E4E2DC"
INK, INK2, INK3 = "#0B0B0B", "#52514E", "#8A8984"

THEME = {
    "name": "Child Survival 2030",
    "dataColors": [BLUE, ORANGE, AQUA, YELLOW, MAGENTA, GREEN, VIOLET, RED],
    "foreground": INK, "foregroundNeutralSecondary": INK2, "foregroundNeutralTertiary": INK3,
    "background": CARD_BG, "backgroundLight": "#F0EFEC", "backgroundNeutral": BORDER,
    "tableAccent": BLUE, "good": "#0CA30C", "neutral": "#FAB219", "bad": "#D03B3B",
    "maximum": "#1C5CAB", "center": "#86B6EF", "minimum": "#F0EFEC",
    "textClasses": {
        "callout": {"fontSize": 26, "fontFace": "Segoe UI Semibold", "color": INK},
        "title": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": INK},
        "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": INK},
        "label": {"fontSize": 10, "fontFace": "Segoe UI", "color": INK2},
    },
    "visualStyles": {
        "*": {"*": {
            "background": [{"show": True, "color": {"solid": {"color": CARD_BG}}, "transparency": 0}],
            "border": [{"show": True, "color": {"solid": {"color": BORDER}}, "radius": 8}],
            "dropShadow": [{"show": False}],
            "padding": [{"top": 10, "bottom": 10, "left": 14, "right": 14}],
            "title": [{"show": True, "fontSize": 12, "fontColor": {"solid": {"color": INK}},
                       "fontFamily": "Segoe UI Semibold"}],
            "categoryAxis": [{"gridlineShow": False, "labelColor": {"solid": {"color": INK2}},
                              "showAxisTitle": False}],
            "valueAxis": [{"gridlineShow": True, "gridlineColor": {"solid": {"color": "#ECEAE4"}},
                           "gridlineStyle": "solid", "labelColor": {"solid": {"color": INK2}},
                           "showAxisTitle": False}],
            "legend": [{"show": True, "position": "Top", "showTitle": False,
                        "labelColor": {"solid": {"color": INK2}}, "fontSize": 9}],
        }},
        "page": {"*": {
            "background": [{"color": {"solid": {"color": PAGE_BG}}, "transparency": 0}],
            "outspace": [{"color": {"solid": {"color": PAGE_BG}}, "transparency": 0}],
        }},
        "lineChart": {"*": {"lineStyles": [{"strokeWidth": 2, "showMarker": False}]}},
        "textbox": {"*": {"background": [{"show": False}], "border": [{"show": False}],
                          "padding": [{"top": 0, "bottom": 0, "left": 0, "right": 0}]}},
        "slicer": {"*": {"header": [{"show": True, "fontColor": {"solid": {"color": INK2}},
                                     "textSize": 9}],
                         "items": [{"fontColor": {"solid": {"color": INK}}, "textSize": 10}]}},
        "card": {"*": {"labels": [{"color": {"solid": {"color": INK}}, "fontSize": 26}],
                       "categoryLabels": [{"show": False}]}},
        "tableEx": {"*": {"grid": [{"gridVertical": False, "rowPadding": 6}],
                          "columnHeaders": [{"fontColor": {"solid": {"color": INK2}}}]}},
        "pivotTable": {"*": {"grid": [{"gridVertical": False, "rowPadding": 4}],
                             "columnHeaders": [{"fontColor": {"solid": {"color": INK2}}}]}},
    },
}

# ---------------------------------------------------------------- query helpers
ALIASES = {"estimates": "e", "progress": "p", "areas": "a", "indicators": "i", "years": "y",
           "Time View": "t", "Ranking Metric": "r", "Scenario ARR": "s", "Peer Country": "c"}


def lit(v):
    if isinstance(v, bool):
        return {"expr": {"Literal": {"Value": "true" if v else "false"}}}
    if isinstance(v, (int, float)):
        return {"expr": {"Literal": {"Value": f"{v}D"}}}
    return {"expr": {"Literal": {"Value": "'" + str(v).replace("'", "''") + "'"}}}


def measure_expr(table, name):
    return {"expr": {"Measure": {"Expression": {"SourceRef": {"Entity": table}}, "Property": name}}}


def color(hex_):
    return {"solid": {"color": lit(hex_)}}


class Field:
    def __init__(self, table, prop, kind="col", agg=None):
        self.table, self.prop, self.kind, self.agg = table, prop, kind, agg

    @property
    def ref(self):
        if self.agg is not None:
            return f"{['Sum','Avg','Count','Min','Max'][self.agg]}({self.table}.{self.prop})"
        return f"{self.table}.{self.prop}"

    def select(self):
        src = {"Expression": {"SourceRef": {"Source": ALIASES[self.table]}}, "Property": self.prop}
        if self.kind == "measure":
            body = {"Measure": src}
        else:
            body = {"Column": src}
            if self.agg is not None:
                body = {"Aggregation": {"Expression": {"Column": src}, "Function": self.agg}}
        return {**body, "Name": self.ref}

    def expr(self):
        src = {"Expression": {"SourceRef": {"Source": ALIASES[self.table]}}, "Property": self.prop}
        return {"Measure": src} if self.kind == "measure" else {"Column": src}


def C(t, p, agg=None):
    return Field(t, p, "col", agg)


def M(t, p):
    return Field(t, p, "measure")


def literal_value(v):
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, int):
        return f"{v}L"
    if isinstance(v, float):
        return f"{v}D"
    return "'" + str(v).replace("'", "''") + "'"


def in_filter(field, values, name=None):
    """Categorical 'In' filter definition (used by filters and slicer defaults)."""
    a = ALIASES[field.table]
    return {
        "Version": 2,
        "From": [{"Name": a, "Entity": field.table, "Type": 0}],
        "Where": [{"Condition": {"In": {
            "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": a}}, "Property": field.prop}}],
            "Values": [[{"Literal": {"Value": literal_value(v)}}] for v in values]}}}],
    }


def filter_entry(field, values):
    h = hashlib.md5(f"{field.ref}{values}".encode()).hexdigest()[:20]
    return {
        "name": "Filter" + h,
        "expression": {"Column": {"Expression": {"SourceRef": {"Entity": field.table}}, "Property": field.prop}},
        "filter": in_filter(field, values),
        "type": "Categorical",
        "howCreated": 1,
        "isHiddenInViewMode": False,
    }


_counter = [0]


def vid(prefix):
    _counter[0] += 1
    return hashlib.md5(f"{prefix}{_counter[0]}".encode()).hexdigest()[:20]


def visual(vtype, x, y, w, h, projections=None, objects=None, vc=None, title=None, sort=None,
           filters=None, name=None, extra=None, display_names=None, z=None):
    projections = projections or {}
    fields, seen = [], set()
    for role, items in projections.items():
        for f in items:
            if f.ref not in seen:
                seen.add(f.ref)
                fields.append(f)
    tables = []
    for f in fields:
        if f.table not in tables:
            tables.append(f.table)
    if sort:
        for f, _ in sort:
            if f.table not in tables:
                tables.append(f.table)
    if vtype == "lineChart":
        objects = {"legend": [{"properties": {"showTitle": lit(False)}}], **(objects or {})}
    sv = {"visualType": vtype}
    if fields:
        sv["projections"] = {role: [{"queryRef": f.ref, **({"active": True} if vtype == "slicer" else {})}
                                    for f in items] for role, items in projections.items()}
        pq = {"Version": 2,
              "From": [{"Name": ALIASES[t], "Entity": t, "Type": 0} for t in tables],
              "Select": [f.select() for f in fields]}
        if sort:
            pq["OrderBy"] = [{"Direction": d, "Expression": f.expr()} for f, d in sort]
        sv["prototypeQuery"] = pq
    if display_names:
        sv["columnProperties"] = {f.ref: {"displayName": dn} for f, dn in display_names}
    sv["drillFilterOtherVisuals"] = True
    if objects:
        sv["objects"] = objects
    vco = dict(vc or {})
    if title is not None:
        if isinstance(title, Field):
            vco["title"] = [{"properties": {"show": lit(True), "text": measure_expr(title.table, title.prop)}}]
        elif title is False:
            vco["title"] = [{"properties": {"show": lit(False)}}]
        else:
            vco["title"] = [{"properties": {"show": lit(True), "text": lit(title)}}]
    if vco:
        sv["vcObjects"] = vco
    if extra:
        sv.update(extra)
    name = name or vid(vtype)
    zz = z if z is not None else _counter[0] * 100
    cfg = {"name": name,
           "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": zz, "width": w, "height": h,
                                               "tabOrder": zz}}],
           "singleVisual": sv}
    return {"x": x, "y": y, "z": zz, "width": w, "height": h,
            "config": json.dumps(cfg, ensure_ascii=False),
            "filters": json.dumps(filters or [], ensure_ascii=False),
            "_name": name}


def textbox(x, y, w, h, paragraphs):
    """paragraphs: list of list of (text, style dict)."""
    paras = [{"textRuns": [{"value": t, "textStyle": s} for t, s in runs]} for runs in paragraphs]
    return visual("textbox", x, y, w, h, objects={"general": [{"properties": {"paragraphs": paras}}]})


def header(title, subtitle):
    return textbox(24, 14, 740, 70, [
        [(title, {"fontFamily": "Segoe UI Semibold", "fontSize": "20pt", "color": INK})],
        [(subtitle, {"fontFamily": "Segoe UI", "fontSize": "10pt", "color": INK2})],
    ])


def slicer(field, x, y, w, h, title, mode="Dropdown", single=True, default=None, sync=None):
    objects = {"data": [{"properties": {"mode": lit(mode)}}],
               "header": [{"properties": {"show": lit(True), "text": lit(title)}}]}
    if mode != "Between":
        objects["selection"] = [{"properties": {"singleSelect": lit(single),
                                                "strictSingleSelect": lit(single)}}]
    if default is not None:
        objects["general"] = [{"properties": {"filter": {"filter": in_filter(field, default)}}}]
    extra = {"syncGroup": {"groupName": sync, "fieldChanges": True, "filterChanges": True}} if sync else None
    return visual("slicer", x, y, w, h, projections={"Values": [field]}, objects=objects, title=False,
                  extra=extra, vc={"padding": [{"properties": {"top": lit(4), "bottom": lit(4)}}]})


def card(measure, x, y, w, h, title, fmt_units=True, precision=None, size=26):
    objects = {"labels": [{"properties": {"fontSize": lit(size), **({"labelDisplayUnits": lit(1)} if fmt_units else {}),
                                          **({"labelPrecision": lit(precision)} if precision is not None else {})}}],
               "categoryLabels": [{"properties": {"show": lit(False)}}]}
    return visual("card", x, y, w, h, projections={"Values": [measure]}, objects=objects, title=title)


def series_colors(field, mapping):
    return [{"properties": {"fill": color(c)},
             "selector": {"data": [{"scopeId": {"Comparison": {
                 "ComparisonKind": 0,
                 "Left": {"Column": {"Expression": {"SourceRef": {"Entity": field.table}}, "Property": field.prop}},
                 "Right": {"Literal": {"Value": literal_value(v)}}}}}]}} for v, c in mapping.items()]


TOOLTIP_PAGE = "c5000000000000000tip1"


def tooltip_vc():
    return {"visualTooltip": [{"properties": {"type": lit("ReportPage"), "section": lit(TOOLTIP_PAGE)}}]}


def section(display, name, visuals, filters=None, no_filter_pairs=(), hidden=False, tooltip=False):
    cfg = {}
    if no_filter_pairs:
        cfg["relationships"] = [{"source": s, "target": t, "type": 3} for s, t in no_filter_pairs]
    if hidden:
        cfg["visibility"] = 1
    if tooltip:
        cfg["objects"] = {
            "pageInformation": [{"properties": {"pageInformationType": lit("Tooltip")}}],
            "background": [{"properties": {"color": {"solid": {"color": lit(CARD_BG)}}, "transparency": lit(0)}}],
        }
        cfg["type"] = 1  # 1 = tooltip page
    return {
        "config": json.dumps(cfg),
        "displayName": display,
        "displayOption": 3 if tooltip else 1,
        "filters": json.dumps(filters or [], ensure_ascii=False),
        "height": 240.0 if tooltip else 720.0,
        "name": name,
        "visualContainers": [{k: v for k, v in vc.items() if k != "_name"} for vc in visuals],
        "width": 320.0 if tooltip else 1280.0,
    }



# ================================================================= Child Survival 2030 pages
STATUS_COLOURS = {"Target met": "#0CA30C", "On track": "#6DA7EC", "Acceleration needed": "#EC835A", "No progress": "#D03B3B"}
TARGETED = ["Under-5 mortality rate", "Neonatal mortality rate", "Maternal mortality ratio"]
TARGETED_CODES = ["CME_MRY0T4", "CME_MRM0", "MNCH_MMR"]
note = {"fontFamily": "Segoe UI", "fontSize": "9pt", "color": INK3}
body = {"fontFamily": "Segoe UI", "fontSize": "10pt", "color": INK2}
bold = {"fontFamily": "Segoe UI Semibold", "fontSize": "10pt", "color": INK}
h3 = {"fontFamily": "Segoe UI Semibold", "fontSize": "12pt", "color": INK}

ind_name, ind_code = C("indicators", "indicator"), C("indicators", "indicator_code")
yr = C("years", "year")
area, area_code, area_type = C("areas", "area_name"), C("areas", "area_code"), C("areas", "area_type")
region, income, is_country = C("areas", "wb_region"), C("areas", "income_group"), C("areas", "is_country")
status = C("progress", "status")
time_view = C("Time View", "Time View")
E = lambda n: M("estimates", n)  # noqa: E731
P = lambda n: M("progress", n)  # noqa: E731


def bool_filter(field, value=True):
    f = filter_entry(field, ["x"])
    a = ALIASES[field.table]
    f["filter"]["Where"][0]["Condition"]["In"]["Values"] = [[{"Literal": {"Value": "true" if value else "false"}}]]
    return f


COUNTRIES_ONLY = [bool_filter(is_country)]


def ind_slicer(x, y, w=250, default="Under-5 mortality rate", sync="indicator"):
    s = slicer(ind_name, x, y, w, 62, "Indicator", default=[default], sync=sync)
    s["filters"] = json.dumps([filter_entry(ind_code, TARGETED_CODES)])
    return s


def footnote(text):
    return textbox(24, 686, 1232, 24, [[(text, note)]])


def status_points():
    return series_colors(status, STATUS_COLOURS)


SOURCE = ("Source: UN IGME 2025 and UN MMEIG 2025 estimates via the UNICEF Data Warehouse; World Bank country groups (FY27). "
          "All figures are modelled estimates.")

# ---------------------------------------------------------------- page 1: global picture
p1 = [header("Where the world stands on child survival",
             "Deaths before age 5, in the first 28 days and in pregnancy or childbirth, 1990 to the latest UN estimates.")]
s1 = ind_slicer(1006, 18)
p1.append(s1)
cw, cg, cy = 296, 16, 96
p1.append(card(E("World Under-5 Deaths"), 24, cy, cw, 96, "Children who died before age 5, 2024", fmt_units=False, precision=2))
p1.append(card(E("World Under-5 Rate"), 24 + (cw + cg), cy, cw, 96, "World under-5 deaths per 1,000 live births", fmt_units=False))
p1.append(card(E("World Under-5 Rate Change Since 1990"), 24 + 2 * (cw + cg), cy, cw, 96, "Change in the world rate since 1990", fmt_units=False))
p1.append(card(E("Nigeria Share of World Under-5 Deaths"), 24 + 3 * (cw + cg), cy, cw, 96, "Share of these deaths in Nigeria", fmt_units=False))
trend = visual("lineChart", 24, 208, 740, 470,
               projections={"Category": [yr], "Y": [E("Estimate")], "Series": [area]},
               filters=[filter_entry(area_type, ["World", "Region"])],
               objects={"dataPoint": series_colors(area, {"World": INK})},
               title=E("Indicator Title"))
p1.append(trend)
p1.append(visual("barChart", 780, 208, 476, 228,
                 projections={"Category": [area], "Y": [E("Latest Value")]},
                 filters=[filter_entry(area_type, ["Income group"])],
                 sort=[(E("Latest Value"), 2)],
                 objects={"labels": [{"properties": {"show": lit(True), "color": color(INK2)}}],
                          "dataPoint": [{"properties": {"fill": color(BLUE)}}],
                          "valueAxis": [{"properties": {"show": lit(False)}}]},
                 title="Latest rate by World Bank income group"))
p1.append(visual("barChart", 780, 450, 476, 228,
                 projections={"Category": [status], "Y": [P("Countries")]},
                 objects={"labels": [{"properties": {"show": lit(True), "color": color(INK2)}}],
                          "dataPoint": status_points(),
                          "valueAxis": [{"properties": {"show": lit(False)}}]},
                 vc=tooltip_vc(), title="Countries by progress toward the 2030 target"))
p1.append(footnote(SOURCE + " Regions are UNICEF reporting regions. Income-group figures are not published for maternal mortality."))
page1 = section("Global picture", "c5000000000000000p01", p1)

# ---------------------------------------------------------------- page 2: on track?
p2 = [header("Who is on track for 2030?",
             "Yearly rate of reduction each country has achieved since 2015, against the rate it needs to reach the SDG target by 2030.")]
p2.append(ind_slicer(1006, 18))
p2.append(slicer(region, 744, 18, 250, 62, "Region", single=False, sync="region"))
p2.append(card(P("Countries Met or On Track"), 24, cy, cw, 96, "Countries that met the target or are on track"))
p2.append(card(P("Countries Off Track"), 24 + (cw + cg), cy, cw, 96, "Countries off track"))
p2.append(card(P("Share Off Track"), 24 + 2 * (cw + cg), cy, cw, 96, "Share of countries off track", fmt_units=False))
p2.append(card(P("Deaths Above Target"), 24 + 3 * (cw + cg), cy, cw, 96, "Deaths a year above target levels", fmt_units=False, precision=2))
scatter = visual("scatterChart", 24, 208, 616, 470,
                 projections={"Category": [area], "Series": [status], "X": [P("ARR Required")], "Y": [P("ARR Achieved")]},
                 filters=COUNTRIES_ONLY + [filter_entry(status, ["On track", "Acceleration needed", "No progress"])],
                 objects={"dataPoint": status_points()},
                 vc=tooltip_vc(),
                 title="Each dot is a country: yearly reduction achieved since 2015 (up) vs needed by 2030 (right)")
p2.append(scatter)
rank_slicer = slicer(C("Ranking Metric", "Ranking Metric"), 656, 208, 600, 62, "Rank countries by")
_cfg = json.loads(rank_slicer["config"])
_cfg["singleVisual"]["objects"]["general"] = [{"properties": {"filter": {"filter": in_filter(
    C("Ranking Metric", "Ranking Metric Fields"), ["'progress'[Speed-up Needed]"])}}}]
rank_slicer["config"] = json.dumps(_cfg, ensure_ascii=False)
p2.append(rank_slicer)
ranking = visual("barChart", 656, 278, 600, 400,
                 projections={"Category": [area], "Y": [P("Speed-up Needed")]},
                 filters=COUNTRIES_ONLY, sort=[(P("Speed-up Needed"), 2)],
                 extra={"queryFieldParametersByRole": {"Y": [{"index": 0, "length": 1, "expr": {"Column": {
                     "Expression": {"SourceRef": {"Entity": "Ranking Metric"}}, "Property": "Ranking Metric"}}}]},
                        "hasDefaultSort": True},
                 objects={"labels": [{"properties": {"show": lit(True), "color": color(INK2)}}],
                          "dataPoint": [{"properties": {"fill": color(BLUE)}}],
                          "valueAxis": [{"properties": {"show": lit(False)}}]},
                 vc=tooltip_vc(), title="Countries ranked")
p2.append(ranking)
p2.append(footnote("Yearly rate of reduction = average yearly % fall in the rate since 2015. Countries that already meet the target "
                   "are not shown on the chart. Right-click a country and choose Drill through to see its profile."))
page2 = section("On track for 2030?", "c5000000000000000p02", p2)

# ---------------------------------------------------------------- page 3: compare countries
p3 = [header("Compare countries",
             "Rates in 1990, 2015 and the latest year, with change and yearly reduction since 2015 (one calculation applied to every column).")]
p3.append(ind_slicer(1006, 18))
p3.append(slicer(region, 744, 18, 250, 62, "Region", single=False, sync="region"))
p3.append(visual("pivotTable", 24, 96, 760, 582,
                 projections={"Rows": [area], "Columns": [time_view], "Values": [E("Estimate")]},
                 filters=COUNTRIES_ONLY, display_names=[(area, "Country")],
                 objects={"subTotals": [{"properties": {"rowSubtotals": lit(False), "columnSubtotals": lit(False)}}]},
                 title=E("Indicator Title")))
p3.append(visual("tableEx", 800, 96, 456, 582,
                 projections={"Values": [area, P("Progress Status"), P("When Target Reached"), P("Speed-up Needed")]},
                 filters=COUNTRIES_ONLY,
                 display_names=[(area, "Country"), (P("When Target Reached"), "Target reached at current pace"),
                                (P("Speed-up Needed"), "Speed-up needed"), (P("Progress Status"), "Status")],
                 objects={"values": [{"properties": {"fontColor": {"solid": {"color": measure_expr("progress", "Status Colour")}}},
                                      "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}],
                                                   "metadata": "progress.Progress Status"}}],
                          "total": [{"properties": {"totals": lit(False)}}]},
                 vc=tooltip_vc(), title="Status and outlook"))
p3.append(footnote("Speed-up needed = yearly reduction required from now to 2030, divided by the yearly reduction achieved since 2015."))
page3 = section("Compare countries", "c5000000000000000p03", p3)

# ---------------------------------------------------------------- page 4: country profile (drill-through)
p4 = []
p4.append(visual("card", 24, 14, 900, 70, projections={"Values": [E("Profile Title")]},
                 objects={"labels": [{"properties": {"fontSize": lit(22)}}],
                          "categoryLabels": [{"properties": {"show": lit(False)}}]},
                 vc={"background": [{"properties": {"show": lit(False)}}], "border": [{"properties": {"show": lit(False)}}]},
                 title=False))
p4.append(visual("card", 24, 62, 900, 34, projections={"Values": [E("Profile Subtitle")]},
                 objects={"labels": [{"properties": {"fontSize": lit(10), "color": color(INK2)}}],
                          "categoryLabels": [{"properties": {"show": lit(False)}}]},
                 vc={"background": [{"properties": {"show": lit(False)}}], "border": [{"properties": {"show": lit(False)}}]},
                 title=False))
lw, lg = 405, 16
for i, (code, title) in enumerate([("CME_MRY0T4", "Under-5 deaths per 1,000 live births"),
                                   ("CME_MRM0", "Newborn deaths per 1,000 live births"),
                                   ("MNCH_MMR", "Maternal deaths per 100,000 live births")]):
    p4.append(visual("lineChart", 24 + i * (lw + lg), 104, lw, 330,
                     projections={"Category": [yr], "Y": [E("Estimate"), E("Lower Bound"), E("Upper Bound"), P("Target Line")]},
                     filters=[filter_entry(ind_code, [code]), filter_entry(yr, list(range(1990, 2025)))],
                     display_names=[(E("Estimate"), "Estimate"), (E("Lower Bound"), "Lower bound"),
                                    (E("Upper Bound"), "Upper bound"), (P("Target Line"), "SDG target")],
                     objects={"dataPoint": [
                         {"properties": {"fill": color(BLUE)}, "selector": {"metadata": "estimates.Estimate"}},
                         {"properties": {"fill": color("#B7D3F6")}, "selector": {"metadata": "estimates.Lower Bound"}},
                         {"properties": {"fill": color("#B7D3F6")}, "selector": {"metadata": "estimates.Upper Bound"}},
                         {"properties": {"fill": color(INK3)}, "selector": {"metadata": "progress.Target Line"}}],
                         "lineStyles": [
                         {"properties": {"strokeWidth": lit(1)}, "selector": {"metadata": "estimates.Lower Bound"}},
                         {"properties": {"strokeWidth": lit(1)}, "selector": {"metadata": "estimates.Upper Bound"}},
                         {"properties": {"strokeWidth": lit(1), "lineStyle": lit("dashed")}, "selector": {"metadata": "progress.Target Line"}}]},
                     title=title))
p4.append(visual("tableEx", 24, 450, 1232, 228,
                 projections={"Values": [ind_name, P("Base Value 2015"), P("Latest Rate"), E("SDG Target"), P("ARR Achieved"),
                                         P("ARR Required"), P("Progress Status"), P("When Target Reached"), P("Crisis Note")]},
                 display_names=[(ind_name, "Indicator"), (P("Base Value 2015"), "2015"), (P("Latest Rate"), "Latest"),
                                (E("SDG Target"), "Target"), (P("ARR Achieved"), "Yearly reduction since 2015"),
                                (P("ARR Required"), "Yearly reduction needed"), (P("When Target Reached"), "Target reached at current pace"),
                                (P("Crisis Note"), "Note"), (P("Progress Status"), "Status")],
                 filters=[filter_entry(ind_code, TARGETED_CODES)],
                 objects={"values": [{"properties": {"fontColor": {"solid": {"color": measure_expr("progress", "Status Colour")}}},
                                      "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}],
                                                   "metadata": "progress.Progress Status"}}],
                          "total": [{"properties": {"totals": lit(False)}}]},
                 title="Progress toward the 2030 targets"))
p4.append(footnote("Shaded lines show the UN IGME uncertainty range (90% interval); maternal mortality has no range in this data. "
                   "Right-click any country on another page and choose Drill through > Country profile."))
drill = {"name": hashlib.md5(b"drillthrough-area_name").hexdigest()[:20],
         "expression": {"Column": {"Expression": {"SourceRef": {"Entity": "areas"}}, "Property": "area_name"}},
         "type": "Categorical", "howCreated": 5}  # 5 = drill-through field
back = visual("actionButton", 1176, 24, 80, 36,
              objects={"icon": [{"properties": {"shapeType": lit("back")}, "selector": {"id": "default"}}]},
              vc={"visualLink": [{"properties": {"show": lit(True), "type": lit("Back")}}]})
p4.append(back)
page4 = section("Country profile", "c5000000000000000p04", p4, filters=[drill])

# ---------------------------------------------------------------- page 5: Nigeria in focus
p5 = [header("Nigeria: what would it take?",
             "18% of the world's under-5 deaths and the highest under-5 death rate of any country; progress since 2015 has stalled.")]
p5.append(ind_slicer(1006, 18, default="Under-5 mortality rate", sync=None))
cw5 = 196
for i, (msr, t) in enumerate([(P("Latest Rate"), "Latest rate"), (P("ARR Achieved"), "Yearly reduction since 2015"),
                              (P("ARR Required"), "Yearly reduction needed for 2030"),
                              (P("Year Reached at Current Pace"), "Target reached at current pace"),
                              (P("Year Reached at Chosen Pace"), "Target reached at chosen pace"),
                              (P("Lives Saved by 2030 at Chosen Pace"), "Lives saved 2025-2030 at chosen pace")]):
    p5.append(card(msr, 24 + i * (cw5 + 11.2), cy, cw5, 96, t, fmt_units=False, size=20,
                   precision=2 if "Lives" in t else None))
p5.append(slicer(C("Scenario ARR", "Scenario ARR"), 24, 208, 300, 62, "Chosen yearly reduction (% a year)", default=[10.0]))
p5.append(slicer(C("Peer Country", "Peer Country"), 340, 208, 300, 62, "Compare with the pace of", default=["Ghana"]))
p5.append(visual("lineChart", 24, 280, 900, 398,
                 projections={"Category": [yr], "Y": [E("Estimate"), P("Path at Current Pace"), P("Path at Chosen Pace"),
                                                      P("Path at Peer Pace"), P("Path at Required Pace"), P("Target Line")]},
                 filters=[filter_entry(yr, list(range(2000, 2031)))],
                 display_names=[(E("Estimate"), "UN estimate"), (P("Path at Current Pace"), "At current pace"),
                                (P("Path at Chosen Pace"), "At chosen pace"), (P("Path at Peer Pace"), "At peer's pace"),
                                (P("Path at Required Pace"), "Pace needed for 2030"), (P("Target Line"), "SDG target")],
                 objects={"dataPoint": [
                     {"properties": {"fill": color(INK)}, "selector": {"metadata": "estimates.Estimate"}},
                     {"properties": {"fill": color(INK3)}, "selector": {"metadata": "progress.Path at Current Pace"}},
                     {"properties": {"fill": color(BLUE)}, "selector": {"metadata": "progress.Path at Chosen Pace"}},
                     {"properties": {"fill": color(ORANGE)}, "selector": {"metadata": "progress.Path at Peer Pace"}},
                     {"properties": {"fill": color(GREEN)}, "selector": {"metadata": "progress.Path at Required Pace"}},
                     {"properties": {"fill": color(INK3)}, "selector": {"metadata": "progress.Target Line"}}],
                     "lineStyles": [
                     {"properties": {"lineStyle": lit("dashed")}, "selector": {"metadata": "progress.Path at Current Pace"}},
                     {"properties": {"lineStyle": lit("dotted"), "strokeWidth": lit(1)}, "selector": {"metadata": "progress.Target Line"}}]},
                 title=P("Scenario Title")))
p5.append(textbox(940, 208, 316, 470, [
    [("How to read this", h3)], [("", body)],
    [("The black line is the UN estimate. From the latest year, the other lines project the rate forward at a constant "
      "yearly rate of reduction:", body)], [("", body)],
    [("Current pace: ", bold), ("Nigeria's own average since 2015.", body)],
    [("Chosen pace: ", bold), ("set it with the slider.", body)],
    [("Peer's pace: ", bold), ("what the selected country achieved since 2015.", body)],
    [("Pace needed: ", bold), ("the rate that reaches the target exactly in 2030.", body)], [("", body)],
    [("Lives saved compares the chosen pace with the current pace from 2025 to 2030, using the latest number of births.", body)],
    [("", body)],
    [("These are simple illustrations, not forecasts. Real progress depends on conflict, funding, health workers and "
      "coverage of care, which are not in this data.", note)]]))
page5 = section("Nigeria in focus", "c5000000000000000p05", p5, filters=[filter_entry(area_code, ["NGA"])])

# ---------------------------------------------------------------- page 6: about
p6 = [header("About the data", "Sources, definitions, methods and limitations.")]
cols = [
    [[("Sources", h3)], [("", body)],
     [("UN IGME 2025: ", bold), ("under-5, infant and neonatal mortality, deaths and stillbirths, 1990-2024, with 90% uncertainty ranges.", body)],
     [("UN MMEIG 2025: ", bold), ("maternal mortality ratio, 2000-2023.", body)],
     [("UN World Population Prospects: ", bold), ("births and population.", body)],
     [("World Bank FY27 country groups: ", bold), ("region and income group.", body)], [("", body)],
     [("All downloaded from the UNICEF Data Warehouse and the World Bank in October 2026.", body)], [("", body)],
     [("Targets", h3)], [("", body)],
     [("SDG 3.2: under-5 mortality 25 or less and neonatal mortality 12 or less per 1,000 live births by 2030.", body)],
     [("SDG 3.1: global maternal mortality below 70 per 100,000 live births; used here as a benchmark for each country.", body)]],
    [[("Methods", h3)], [("", body)],
     [("Yearly rate of reduction (ARR) = ln(rate 2015 / rate latest) / years between them.", body)],
     [("Needed ARR = ln(latest rate / target) / (2030 - latest year).", body)], [("", body)],
     [("Status: ", bold), ("target met; on track (achieved ARR at least the needed ARR); acceleration needed (falling, but too slowly); "
                           "no progress (no fall since 2015).", body)], [("", body)],
     [("Projections apply a constant ARR from the latest year. Regional (UNICEF reporting regions), income group and world "
       "figures are the official birth-weighted UN aggregates, not averages of countries.", body)]],
    [[("Limitations", h3)], [("", body)],
     [("All figures are modelled estimates. Where death registration is weak, including Nigeria, they rest on surveys and models "
       "and carry wide uncertainty.", body)], [("", body)],
     [("UN IGME includes deaths from conflicts and disasters. Countries with a sharp jump since 2014 (such as South Sudan, "
       "Central African Republic and Somalia) are flagged, because their progress since 2015 is distorted.", body)], [("", body)],
     [("Each country's region and income group come from the World Bank's FY27 groups; regional trend lines use "
       "UNICEF reporting regions, so a few countries sit in a differently named region in each.", body)]],
]
for i, c in enumerate(cols):
    p6.append(textbox(24 + i * 416, 100, 400, 560, c))
page6 = section("About the data", "c5000000000000000p06", p6)

# ---------------------------------------------------------------- tooltip page
tip = [
    visual("card", 0, 0, 320, 64, projections={"Values": [P("Progress Status")]},
           objects={"labels": [{"properties": {"fontSize": lit(16)}}], "categoryLabels": [{"properties": {"show": lit(False)}}]},
           vc={"border": [{"properties": {"show": lit(False)}}]}, title=E("Profile Title")),
    visual("lineChart", 0, 64, 320, 176,
           projections={"Category": [yr], "Y": [E("Estimate"), P("Target Line")]},
           objects={"dataPoint": [{"properties": {"fill": color(BLUE)}, "selector": {"metadata": "estimates.Estimate"}},
                                  {"properties": {"fill": color(INK3)}, "selector": {"metadata": "progress.Target Line"}}],
                    "legend": [{"properties": {"show": lit(False)}}]},
           vc={"border": [{"properties": {"show": lit(False)}}]}, title=E("Indicator Title")),
]
page_tip = section("Country tooltip", TOOLTIP_PAGE, tip, hidden=True, tooltip=True)


# ---------------------------------------------------------------- write report folder
def main():
    REPORT.mkdir(parents=True, exist_ok=True)
    cfg = {"version": "5.59",
           "themeCollection": {"baseTheme": {"name": "CY24SU10", "type": 2, "version": "5.59"},
                               "customTheme": {"name": THEME_NAME, "type": 1, "version": "5.65"}},
           "activeSectionIndex": 0, "defaultDrillFilterOtherVisuals": True,
           "settings": {"useNewFilterPaneExperience": True, "allowChangeFilterTypes": True,
                        "useStylableVisualContainerHeader": True, "queryLimitOption": 6, "useEnhancedTooltips": True,
                        "exportDataMode": 1, "useDefaultAggregateDisplayName": True},
           "objects": {"section": [{"properties": {"verticalAlignment": {"expr": {"Literal": {"Value": "'Top'"}}}}}]}}
    # Bookmarks are authored in Power BI Desktop (View > Bookmarks); keep any that already exist.
    old_path = REPORT / "report.json"
    if old_path.exists():
        try:
            old_cfg = json.loads(json.loads(old_path.read_text(encoding="utf-8"))["config"])
            if old_cfg.get("bookmarks"):
                cfg["bookmarks"] = old_cfg["bookmarks"]
        except (ValueError, KeyError):
            pass
    rj = {"config": json.dumps(cfg), "layoutOptimization": 0,
          "resourcePackages": [
              {"resourcePackage": {"disabled": False, "items": [{"name": "CY24SU10", "path": "BaseThemes/CY24SU10.json", "type": 202}],
                                   "name": "SharedResources", "type": 2}},
              {"resourcePackage": {"disabled": False, "items": [{"name": THEME_NAME, "path": THEME_NAME, "type": 201}],
                                   "name": "RegisteredResources", "type": 1}}],
          "sections": [page1, page2, page3, page4, page5, page6, page_tip]}
    for i, sec in enumerate(rj["sections"]):
        sec["ordinal"] = i
    # drill-through pages are registered as "pods" bound to their section
    # acceptsFilterContext 1 = "Keep all filters" off: only the country is passed, not the indicator slicer
    rj["pods"] = [{"boundSection": page4["name"], "config": json.dumps({"acceptsFilterContext": 1}),
                   "name": hashlib.md5(b"pod-country-profile").hexdigest()[:20],
                   "parameters": json.dumps([{"name": hashlib.md5(b"pod-param-area").hexdigest()[:20],
                                              "boundFilter": drill["name"],
                                              "fieldExpr": drill["expression"]}]),
                   "type": 1}]
    (REPORT / "report.json").write_text(json.dumps(rj, indent=2, ensure_ascii=False), encoding="utf-8")
    theme_dir = REPORT / "StaticResources" / "RegisteredResources"
    theme_dir.mkdir(parents=True, exist_ok=True)
    (theme_dir / THEME_NAME).write_text(json.dumps(THEME, indent=2), encoding="utf-8")
    (REPORT / "definition.pbir").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}}, indent=2), encoding="utf-8")
    plat = REPORT / ".platform"
    if not plat.exists():
        import uuid
        plat.write_text(json.dumps({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": "Report", "displayName": NAME},
            "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}}, indent=2), encoding="utf-8")
    base = REPORT / "StaticResources" / "SharedResources" / "BaseThemes" / "CY24SU10.json"
    if not base.exists():
        print("NOTE: copy CY24SU10.json into", base.parent)
    print("report written:", sum(len(s["visualContainers"]) for s in rj["sections"]), "visuals on", len(rj["sections"]), "pages")


if __name__ == "__main__":
    main()
