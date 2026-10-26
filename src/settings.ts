/*
 * Defaults come from the Lean Excel brand tokens (CLAUDE.md, .lean-excel/standards/BRAND.md).
 * Format pane model. Object and property names here must match capabilities.json
 * and must NEVER be renamed after the first public release (CLAUDE.md rule 5):
 * renaming silently resets every user's saved formatting.
 */
import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { tokens as T } from "./core";

import Card = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;
import Slice = formattingSettings.Slice;

export const BIN_RULES = [
    { value: "scott", displayName: "Automatic (Excel, Scott's rule)" },
    { value: "freedman-diaconis", displayName: "Automatic, robust to outliers (Freedman–Diaconis)" },
    { value: "sturges", displayName: "Automatic, small samples (Sturges)" },
    { value: "sqrt", displayName: "Square root of count" },
    { value: "count", displayName: "Number of bins" },
    { value: "width", displayName: "Bin width" }
];

export const Y_MEASURES = [
    { value: "count", displayName: "Count" },
    { value: "percent", displayName: "% of total" },
    { value: "density", displayName: "Density" }
];

const color = (hex: string): powerbi.ThemeColorData => ({ value: hex }) as powerbi.ThemeColorData;

export class BinsCard extends Card {
    name = "bins";
    displayName = "Bins";
    description = "How values are grouped. Automatic matches Excel's histogram chart.";

    rule = new formattingSettings.ItemDropdown({ name: "rule", displayName: "Bin method", items: BIN_RULES, value: BIN_RULES[0] });
    binCount = new formattingSettings.NumUpDown({
        name: "binCount", displayName: "Number of bins", value: 10,
        options: { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 1 }, maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 500 } }
    });
    binWidth = new formattingSettings.NumUpDown({
        name: "binWidth", displayName: "Bin width", value: 1,
        options: { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } }
    });
    roundEdges = new formattingSettings.ToggleSwitch({ name: "roundEdges", displayName: "Round bin edges (e.g. 42, 44, 46)", value: false });
    start = new formattingSettings.TextInput({ name: "start", displayName: "First bin starts at", value: "", placeholder: "Auto (data minimum)" });
    underflow = new formattingSettings.TextInput({ name: "underflow", displayName: "Underflow bin (≤ value)", value: "", placeholder: "Off" });
    overflow = new formattingSettings.TextInput({ name: "overflow", displayName: "Overflow bin (> value)", value: "", placeholder: "Off" });

    slices: Slice[] = [this.rule, this.binCount, this.binWidth, this.roundEdges, this.start, this.underflow, this.overflow];
}

export class BarsCard extends Card {
    name = "bars";
    displayName = "Bars";

    fill = new formattingSettings.ColorPicker({ name: "fill", displayName: "Color", value: color("") });
    gap = new formattingSettings.Slider({
        name: "gap", displayName: "Gap between bars (%)", value: 2,
        options: { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 }, maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 50 } }
    });
    showLabels = new formattingSettings.ToggleSwitch({ name: "showLabels", displayName: "Data labels", value: false });
    labelColor = new formattingSettings.ColorPicker({ name: "labelColor", displayName: "Label color", value: color(T["color.neutral.700"]) });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Label size", value: 9 });

    slices: Slice[] = [this.fill, this.gap, this.showLabels, this.labelColor, this.fontSize];
}

export class YAxisCard extends Card {
    name = "yAxis";
    displayName = "Y axis";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: true });
    topLevelSlice = this.show;
    measure = new formattingSettings.ItemDropdown({ name: "measure", displayName: "Show as", items: Y_MEASURES, value: Y_MEASURES[0] });
    showTitle = new formattingSettings.ToggleSwitch({ name: "showTitle", displayName: "Title", value: true });
    gridlines = new formattingSettings.ToggleSwitch({ name: "gridlines", displayName: "Gridlines", value: true });
    fontColor = new formattingSettings.ColorPicker({ name: "fontColor", displayName: "Color", value: color(T["color.neutral.muted"]) });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Text size", value: 9 });

    slices: Slice[] = [this.measure, this.showTitle, this.gridlines, this.fontColor, this.fontSize];
}

export class XAxisCard extends Card {
    name = "xAxis";
    displayName = "X axis";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: true });
    topLevelSlice = this.show;
    showTitle = new formattingSettings.ToggleSwitch({ name: "showTitle", displayName: "Title", value: true });
    decimals = new formattingSettings.NumUpDown({
        name: "decimals", displayName: "Decimal places (−1 = auto)", value: -1,
        options: { minValue: { type: powerbi.visuals.ValidatorType.Min, value: -1 }, maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 9 } }
    });
    fontColor = new formattingSettings.ColorPicker({ name: "fontColor", displayName: "Color", value: color(T["color.neutral.muted"]) });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Text size", value: 9 });

    slices: Slice[] = [this.showTitle, this.decimals, this.fontColor, this.fontSize];
}

export class OverlaysCard extends Card {
    name = "overlays";
    displayName = "Statistics overlays";

    normalCurve = new formattingSettings.ToggleSwitch({ name: "normalCurve", displayName: "Normal curve", value: false });
    curveColor = new formattingSettings.ColorPicker({ name: "curveColor", displayName: "Curve color", value: color(T["color.neutral.ink"]) });
    meanLine = new formattingSettings.ToggleSwitch({ name: "meanLine", displayName: "Mean line", value: false });
    medianLine = new formattingSettings.ToggleSwitch({ name: "medianLine", displayName: "Median line", value: false });
    statColor = new formattingSettings.ColorPicker({ name: "statColor", displayName: "Mean / median color", value: color(T["color.neutral.ink"]) });
    cumulative = new formattingSettings.ToggleSwitch({ name: "cumulative", displayName: "Cumulative % line", value: false });
    cumulativeColor = new formattingSettings.ColorPicker({ name: "cumulativeColor", displayName: "Cumulative color", value: color(T["color.chart.7"]) });

    slices: Slice[] = [this.normalCurve, this.curveColor, this.meanLine, this.medianLine, this.statColor, this.cumulative, this.cumulativeColor];
}

export class SpecLimitsCard extends Card {
    name = "specLimits";
    displayName = "Specification limits";
    description = "Fixed values here, or bind measures to the LSL / USL / Target fields (measures win).";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: true });
    topLevelSlice = this.show;
    lsl = new formattingSettings.TextInput({ name: "lsl", displayName: "Lower spec limit (LSL)", value: "", placeholder: "None" });
    usl = new formattingSettings.TextInput({ name: "usl", displayName: "Upper spec limit (USL)", value: "", placeholder: "None" });
    target = new formattingSettings.TextInput({ name: "target", displayName: "Target", value: "", placeholder: "None" });
    lineColor = new formattingSettings.ColorPicker({ name: "lineColor", displayName: "Limit line color", value: color(T["color.status.notMet"]) });
    targetColor = new formattingSettings.ColorPicker({ name: "targetColor", displayName: "Target line color", value: color(T["color.petrol.600"]) });
    outOfSpecFill = new formattingSettings.ColorPicker({ name: "outOfSpecFill", displayName: "Out-of-spec bar color", value: color(T["color.status.notMet"]) });
    showCapability = new formattingSettings.ToggleSwitch({ name: "showCapability", displayName: "Show Pp, Ppk and expected out of spec", value: true });

    slices: Slice[] = [this.lsl, this.usl, this.target, this.lineColor, this.targetColor, this.outOfSpecFill, this.showCapability];
}

export class SummaryCard extends Card {
    name = "summary";
    displayName = "Summary statistics";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: true });
    topLevelSlice = this.show;
    fontColor = new formattingSettings.ColorPicker({ name: "fontColor", displayName: "Color", value: color(T["color.neutral.muted"]) });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Text size", value: 9 });

    slices: Slice[] = [this.fontColor, this.fontSize];
}

export class SmallMultiplesCard extends Card {
    name = "smallMultiples";
    displayName = "Small multiples (Histogram Pro)";
    description = "Applies when a field is in Split by. Every panel uses the same bins so shapes compare fairly.";

    sameYAxis = new formattingSettings.ToggleSwitch({ name: "sameYAxis", displayName: "Same Y axis on every panel", value: true });
    columns = new formattingSettings.NumUpDown({ name: "columns", displayName: "Panels per row (0 = automatic)", value: 0 });
    maxPanels = new formattingSettings.NumUpDown({ name: "maxPanels", displayName: "Most panels to show", value: 12 });
    titleSize = new formattingSettings.NumUpDown({ name: "titleSize", displayName: "Panel title size", value: 10 });

    slices: Slice[] = [this.sameYAxis, this.columns, this.maxPanels, this.titleSize];
}

export class HistogramSettings extends Model {
    bins = new BinsCard();
    bars = new BarsCard();
    yAxis = new YAxisCard();
    xAxis = new XAxisCard();
    overlays = new OverlaysCard();
    specLimits = new SpecLimitsCard();
    summary = new SummaryCard();
    smallMultiples = new SmallMultiplesCard();

    cards = [this.bins, this.bars, this.yAxis, this.xAxis, this.overlays, this.specLimits, this.summary, this.smallMultiples];

    /** Show only the slices relevant to the chosen options. Call before building the pane. */
    applyVisibility(): void {
        const rule = String(this.bins.rule.value?.value ?? "scott");
        this.bins.binCount.visible = rule === "count";
        this.bins.binWidth.visible = rule === "width";
        this.bins.roundEdges.visible = rule !== "width";
        this.bars.labelColor.visible = this.bars.fontSize.visible = this.bars.showLabels.value;
        this.overlays.curveColor.visible = this.overlays.normalCurve.value;
        this.overlays.statColor.visible = this.overlays.meanLine.value || this.overlays.medianLine.value;
        this.overlays.cumulativeColor.visible = this.overlays.cumulative.value;
    }
}

/** Parse a format-pane text box as a number; blank or invalid means "not set". Accepts "1,5" in comma-decimal locales. */
export function parseOptionalNumber(text: string | undefined): number | undefined {
    if (text === undefined || text === null) return undefined;
    let t = String(text).trim().replace(/\s/g, "");
    if (t === "") return undefined;
    if (t.includes(",") && !t.includes(".")) t = t.replace(",", ".");
    else t = t.replace(/,/g, "");
    const n = Number(t);
    return Number.isFinite(n) ? n : undefined;
}
