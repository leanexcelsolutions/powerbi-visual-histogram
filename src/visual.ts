/*
 * Histogram by Lean Excel — certified-ready Power BI visual. Everything is free except
 * Split by (small multiples), which needs a Histogram Pro licence (src/license.ts).
 * Spec: docs/03-products/histogram/SPEC.md. Maths: @leanexcel/powerbi-visuals-core.
 *
 * Certification rules (CLAUDE.md): no network, no raw HTML injection, d3 submodules only,
 * rendering events on every path. All text is set with textContent.
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { valueFormatter, textMeasurementService } from "powerbi-visuals-utils-formattingutils";
import { select, Selection } from "d3-selection";
import { scaleLinear, ScaleLinear } from "d3-scale";
import { axisLeft, axisRight } from "d3-axis";
import { binValues, weightedSummary, normalPdf, capability, Capability, Bin, BinRule, Summary, tokens as T } from "./core";

import { HistogramSettings, parseOptionalNumber } from "./settings";
import { readData, HistogramData, ROW_LIMIT } from "./data";
import { ProLicence } from "./license";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ITooltipService = powerbi.extensibility.ITooltipService;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import ISandboxExtendedColorPalette = powerbi.extensibility.ISandboxExtendedColorPalette;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;

type D3Sel<T extends Element> = Selection<T, unknown, null, undefined>;

const FONT = "'Segoe UI', wf_segoe-ui_normal, helvetica, arial, sans-serif";
const SVG_NS = "http://www.w3.org/2000/svg";
/** From this magnitude, numbers are shown with display units (bn, T) instead of in full. */
const LARGE_VALUE = 1e9;
/** Power BI's formatter shows numbers beyond this in scientific notation. */
const SCIENTIFIC_VALUE = 1e15;
/** Format-pane font sizes are points (like native visuals); SVG/CSS need pixels. */
const px = (pt: number): number => Math.round(pt * 4 / 3 * 10) / 10;

/** Everything the renderer needs for one frame. */
interface Model {
    data: HistogramData;
    bins: Bin[];
    binOf: number[];
    rowsByBin: number[][];
    highlightByBin?: number[];
    summary: Summary;
    total: number;
    regularWidth: number;
    lsl?: number;
    usl?: number;
    target?: number;
    outOfSpecShare?: number;
    capability?: Capability | null;
    /** Weighted blanks in this model's rows. */
    blanks: number;
    format: (v: number) => string;
    formatPrecise: (v: number) => string;
}

/** One chart: the whole visual, or one small-multiple panel. */
interface Panel {
    el: HTMLElement;
    titleEl?: HTMLElement;
    summaryEl: HTMLElement;
    noteEl: HTMLElement;
    svg: D3Sel<SVGSVGElement>;
    model: Model;
}

/** Natural order for group labels: numbers numerically, text alphabetically, "(Blank)" last. */
function compareLabels(a: string, b: string): number {
    if (a === b) return 0;
    if (a === "(Blank)") return 1;
    if (b === "(Blank)") return -1;
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}

interface Colors {
    bar: string; text: string; axis: string; grid: string; label: string;
    curve: string; stat: string; cumulative: string; limit: string; target: string; outOfSpec: string;
    selectedStroke: string; highContrast: boolean; background: string;
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private events: IVisualEventService;
    private selectionManager: ISelectionManager;
    private tooltipService: ITooltipService;
    private localization: ILocalizationManager;
    private formattingService = new FormattingSettingsService();
    private settings = new HistogramSettings();

    private root: HTMLElement;
    private summaryEl: HTMLElement;
    private noteEl: HTMLElement;
    private landingEl: HTMLElement;
    private svg: D3Sel<SVGSVGElement>;

    private model: Model | null = null;
    private gridEl: HTMLElement;
    private panels: Panel[] = [];
    private multiples = false;
    private licence: ProLicence;
    private lastOptions: VisualUpdateOptions | null = null;
    /** Heights of the summary strips and notes the current layout was built with. */
    private usedHeights = new WeakMap<Element, number>();
    private sizeWatch: ResizeObserver | undefined;
    private relayoutQueued = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipService = options.host.tooltipService;
        this.localization = options.host.createLocalizationManager();

        this.root = document.createElement("div");
        this.root.className = "lx-hist";
        options.element.appendChild(this.root);

        this.summaryEl = this.el("div", "lx-summary");
        this.noteEl = this.el("div", "lx-note");
        this.landingEl = this.el("div", "lx-landing");
        this.svg = this.createSvg(this.root);
        this.gridEl = this.el("div", "lx-grid-panels");
        this.gridEl.style.display = "none";
        this.licence = new ProLicence(this.host);
        // Power BI can draw a visual while its container is hidden or not laid out yet (dashboard
        // tiles, pages rendered in the background), then show it without another update. The
        // summary strip then measured 0 px and the chart overflowed its box. Watch the strip and
        // redraw once when its real height differs from the one the layout used.
        if (typeof ResizeObserver !== "undefined") {
            this.sizeWatch = new ResizeObserver(entries => {
                const changed = entries.some(e => {
                    const used = this.usedHeights.get(e.target);
                    return used !== undefined && Math.abs((e.target as HTMLElement).offsetHeight - used) > 1;
                });
                if (changed) this.queueRelayout();
            });
        }
        this.root.addEventListener("contextmenu", (e: MouseEvent) => {
            if (!this.interactive || (e.target as Element).closest(".lx-bar")) return;
            this.selectionManager.showContextMenu({} as ISelectionId, { x: e.clientX, y: e.clientY });
            e.preventDefault();
        });
        // Keep highlight state in sync with bookmarks and other visuals.
        this.selectionManager.registerOnSelectCallback(() => this.applySelectionState());
    }

    public update(options: VisualUpdateOptions): void {
        this.events.renderingStarted(options);
        this.lastOptions = options;
        try {
            const dataView = options.dataViews?.[0];
            this.settings = this.formattingService.populateFormattingSettingsModel(HistogramSettings, dataView);
            const state = readData(dataView, this.host);
            const wantsSplit = state.kind === "ok" && !!state.data.split;
            if (wantsSplit && this.licence.state === "pending") {
                // Wait for the licence answer before the first frame, so exports and the first
                // paint never show the wrong view. ProLicence times out on its own.
                this.licence.ready.then(() => {
                    if (this.lastOptions !== options) return; // a newer update already drew
                    this.draw(options, state);
                    this.events.renderingFinished(options);
                }).catch(err => this.events.renderingFailed(options, String(err)));
                return;
            }
            this.draw(options, state);
            this.events.renderingFinished(options);
        } catch (err) {
            this.events.renderingFailed(options, err instanceof Error ? err.message : String(err));
        }
    }

    /** Redraw once with the last update, after the browser has finished its current layout. */
    private queueRelayout(): void {
        if (this.relayoutQueued || !this.lastOptions) return;
        this.relayoutQueued = true;
        setTimeout(() => {
            this.relayoutQueued = false;
            if (this.lastOptions) this.update(this.lastOptions);
        }, 0);
    }

    private draw(options: VisualUpdateOptions, state: ReturnType<typeof readData>): void {
        const { width, height } = options.viewport;
        if (state.kind === "empty") {
            this.model = null;
            this.licence.reset();
            this.showLanding();
            return;
        }
        if (state.kind === "noNumbers") {
            this.model = null;
            this.showMessage(this.t("NoNumbers_Title"), this.t("NoNumbers_Body"));
            return;
        }
        const data = state.data;
        this.model = this.buildModel(data);
        const extraNotes: string[] = [];
        if (data.split) {
            const st = this.licence.state;
            if (st === "pro") {
                this.licence.reset();
                this.renderMultiples(width, height);
                return;
            }
            // Free view: all data combined, never blocked. Say why once, in plain words.
            if (st === "free") {
                this.licence.notifyBlocked(this.t("Pro_Blocked_Tooltip", data.splitName));
                extraNotes.push(this.t("Note_ProSplit", data.splitName));
            } else if (st === "unsupported") {
                extraNotes.push(this.t("Note_ProUnsupported", data.splitName));
            } else {
                extraNotes.push(this.t("Note_ProUnavailable", data.splitName));
            }
        } else {
            this.licence.reset();
        }
        this.renderSingle(width, height, extraNotes);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        this.settings.applyVisibility();
        return this.formattingService.buildFormattingModel(this.settings);
    }

    // ------------------------------------------------------------------ model

    /**
     * Bins always come from ALL rows, so small-multiple panels share identical edges.
     * With a mask, counts, statistics and selection are limited to the masked rows.
     */
    private buildModel(data: HistogramData, mask?: boolean[]): Model {
        const s = this.settings;
        const rule = String(s.bins.rule.value?.value ?? "scott");
        const result = binValues(data.values, {
            rule: (["scott", "freedman-diaconis", "sturges", "sqrt"].includes(rule) ? rule : "scott") as BinRule,
            binCount: rule === "count" ? s.bins.binCount.value : undefined,
            binWidth: rule === "width" && s.bins.binWidth.value > 0 ? s.bins.binWidth.value : undefined,
            start: parseOptionalNumber(s.bins.start.value),
            underflow: parseOptionalNumber(s.bins.underflow.value),
            overflow: parseOptionalNumber(s.bins.overflow.value),
            nice: s.bins.roundEdges.value
        }, data.weights);

        let bins = result.bins;
        let binOf = result.binOf;
        const weightOf = (i: number) => (data.weights ? data.weights[i] : 1);
        if (mask) {
            binOf = binOf.map((b, i) => (mask[i] ? b : -1));
            const counts = bins.map(() => 0);
            binOf.forEach((b, i) => { if (b >= 0) counts[b] += weightOf(i); });
            const tot = counts.reduce((a, b) => a + b, 0);
            let run = 0;
            bins = bins.map((b, k) => {
                run += counts[k];
                return { ...b, count: counts[k], share: tot > 0 ? counts[k] / tot : 0, cumulativeShare: tot > 0 ? run / tot : 0 };
            });
        }
        const values = mask ? data.values.map((v, i) => (mask[i] ? v : null)) : data.values;
        let blanks = data.blanks;
        if (mask) {
            blanks = 0;
            data.values.forEach((v, i) => { if (mask[i] && v === null) blanks += data.hasFrequency ? weightOf(i) : 1; });
        }
        const rowsByBin: number[][] = bins.map((): number[] => []);
        binOf.forEach((b, i) => { if (b >= 0) rowsByBin[b].push(i); });

        let highlightByBin: number[] | undefined;
        if (data.highlights) {
            highlightByBin = bins.map(() => 0);
            binOf.forEach((b, i) => { if (b >= 0) highlightByBin![b] += data.highlights![i]; });
        }

        const summary = weightedSummary(values, data.weights);
        const total = bins.reduce((t, b) => t + b.count, 0);
        const regular = bins.filter(b => b.kind === "regular");
        const regularWidth = regular.length ? (regular[0].x1 - regular[0].x0) : 0;

        // Spec limits: bound measures win over typed values.
        const show = s.specLimits.show.value;
        const lsl = show ? (data.lsl ?? parseOptionalNumber(s.specLimits.lsl.value)) : undefined;
        const usl = show ? (data.usl ?? parseOptionalNumber(s.specLimits.usl.value)) : undefined;
        const target = show ? (data.target ?? parseOptionalNumber(s.specLimits.target.value)) : undefined;
        let outOfSpecShare: number | undefined;
        if (lsl !== undefined || usl !== undefined) {
            let out = 0;
            values.forEach((v, i) => {
                if (v === null) return;
                const w = data.weights ? data.weights[i] : 1;
                if ((lsl !== undefined && v < lsl) || (usl !== undefined && v > usl)) out += w;
            });
            outOfSpecShare = total > 0 ? out / total : 0;
        }
        const cap = s.specLimits.showCapability.value ? capability(summary.mean, summary.stdDev, lsl, usl) : null;

        // Number formatting: honour the column's format string and the report locale; pick
        // enough decimals to tell bin edges apart unless the user fixed the decimals.
        // Very large values (a billion and up, e.g. 16-digit numbers) use display units such as
        // "1,234.57T", as native Power BI axes do; otherwise the formatter falls back to
        // scientific notation and every label reads "1E+15". Smaller data keeps plain numbers.
        const decimals = s.xAxis.decimals.value;
        const maxAbs = Math.max(Math.abs(summary.min), Math.abs(summary.max));
        const unitValue = isFinite(maxAbs) && maxAbs >= LARGE_VALUE ? maxAbs : 0;
        const unitBase = valueFormatter.create({ format: data.valueFormat, value: unitValue, cultureSelector: this.host.locale }).displayUnit?.value || 1;
        const scaledWidth = regularWidth / unitBase;
        // Above 1e15 Power BI always switches to scientific notation (1.2345E+15), so the
        // decimals apply to the mantissa: enough of them to tell neighbouring bin edges apart.
        const autoPrecision = maxAbs > SCIENTIFIC_VALUE && regularWidth > 0
            ? Math.min(9, Math.max(0, Math.floor(Math.log10(maxAbs)) - Math.floor(Math.log10(regularWidth))))
            : scaledWidth > 0 ? Math.min(6, Math.max(0, 1 - Math.floor(Math.log10(scaledWidth)))) : 2;
        const precision = decimals >= 0 ? decimals : autoPrecision;
        const fmt = valueFormatter.create({ format: data.valueFormat, value: unitValue, precision, cultureSelector: this.host.locale });
        const fmtPrecise = valueFormatter.create({ format: data.valueFormat, value: unitValue, precision: Math.min(9, precision + 1), cultureSelector: this.host.locale });

        return {
            data, bins, binOf, rowsByBin, highlightByBin, summary, total, regularWidth,
            lsl, usl, target, outOfSpecShare, capability: cap, blanks,
            format: (v: number) => fmt.format(v),
            formatPrecise: (v: number) => fmtPrecise.format(v)
        };
    }

    // ------------------------------------------------------------------ render

    private colors(): Colors {
        const s = this.settings;
        const palette = this.host.colorPalette as ISandboxExtendedColorPalette;
        if (palette.isHighContrast) {
            const fg = palette.foreground.value, bg = palette.background.value, sel = palette.foregroundSelected.value;
            return {
                bar: bg, text: fg, axis: fg, grid: fg, label: fg, curve: fg, stat: fg, cumulative: fg,
                limit: fg, target: fg, outOfSpec: fg, selectedStroke: sel, highContrast: true, background: bg
            };
        }
        const pick = (c: powerbi.ThemeColorData | undefined, fallback: string) => (c && c.value) ? c.value : fallback;
        return {
            bar: pick(s.bars.fill.value, palette.getColor("histogram-bars").value),
            text: pick(s.summary.fontColor.value, T["color.neutral.muted"]),
            axis: pick(s.yAxis.fontColor.value, T["color.neutral.muted"]),
            grid: T["color.neutral.line"],
            label: pick(s.bars.labelColor.value, T["color.neutral.700"]),
            curve: pick(s.overlays.curveColor.value, T["color.neutral.ink"]),
            stat: pick(s.overlays.statColor.value, T["color.neutral.ink"]),
            cumulative: pick(s.overlays.cumulativeColor.value, T["color.chart.7"]),
            limit: pick(s.specLimits.lineColor.value, T["color.status.notMet"]),
            target: pick(s.specLimits.targetColor.value, T["color.petrol.600"]),
            outOfSpec: pick(s.specLimits.outOfSpecFill.value, T["color.status.notMet"]),
            selectedStroke: T["color.neutral.ink"],
            highContrast: false,
            background: "transparent"
        };
    }

    /** The whole visual is one histogram (no Split by, or no Pro licence). */
    private renderSingle(width: number, height: number, extraNotes: string[]): void {
        this.multiples = false;
        this.landingEl.style.display = "none";
        this.gridEl.style.display = "none";
        this.gridEl.replaceChildren();
        this.panels = [];
        const p: Panel = { el: this.root, summaryEl: this.summaryEl, noteEl: this.noteEl, svg: this.svg, model: this.model! };
        this.renderPanel(p, width, height, { notes: this.dataNotes(p.model.data).concat(extraNotes) });
        this.applySelectionState();
    }

    /** Histogram Pro: one panel per Split by group, same bins, optional shared Y axis. */
    private renderMultiples(width: number, height: number): void {
        const data = this.model!.data;
        const sm = this.settings.smallMultiples;
        this.multiples = true;
        this.landingEl.style.display = "none";
        this.summaryEl.style.display = "none";
        this.svg.style("display", "none");

        // Groups in natural order.
        const labels = Array.from(new Set(data.split!)).sort(compareLabels);
        const maxPanels = Math.max(1, Math.min(50, Math.round(sm.maxPanels.value || 12)));
        const shown = labels.slice(0, maxPanels);

        // Notes for the whole visual go once, above the grid.
        const notes = this.dataNotes(data);
        if (labels.length > shown.length) notes.push(this.t("Note_PanelsCapped", String(shown.length), String(labels.length), data.splitName));
        this.noteEl.textContent = notes.join(" ");
        this.noteEl.style.display = notes.length ? "block" : "none";
        const top = notes.length ? this.noteEl.offsetHeight : 0;

        // Grid shape: fixed columns, or the best automatic column count.
        const n = shown.length;
        let cols = Math.round(sm.columns.value || 0);
        if (!(cols >= 1)) {
            let best = 1, bestScore = Infinity;
            for (let k = 1; k <= n; k++) {
                const r = Math.ceil(n / k);
                const aspect = (width / k) / Math.max(1, (height - top) / r);
                // Panels near 4:3, and few empty cells in the last row.
                const score = Math.abs(Math.log(aspect / (4 / 3))) + (r * k - n) / (r * k);
                if (score < bestScore) { best = k; bestScore = score; }
            }
            cols = best;
        }
        cols = Math.min(cols, n);
        const rows = Math.ceil(n / cols);
        const gap = 8;
        const pw = Math.floor((width - gap * (cols - 1)) / cols);
        const ph = Math.floor((height - top - gap * (rows - 1)) / rows);

        const grid = this.gridEl;
        grid.replaceChildren();
        grid.style.display = "grid";
        grid.style.gridTemplateColumns = `repeat(${cols}, ${pw}px)`;
        grid.style.gap = `${gap}px`;

        const c = this.colors();
        const models = shown.map(label => this.buildModel(data, data.split!.map(g => g === label)));
        this.panels = shown.map((label, k) => {
            const el = this.el("div", "lx-panel", grid);
            el.style.width = `${pw}px`;
            el.style.height = `${ph}px`;
            const titleEl = this.el("div", "lx-panel-title", el);
            titleEl.textContent = label;
            titleEl.title = `${data.splitName}: ${label}`;
            titleEl.style.fontSize = `${px(sm.titleSize.value)}px`;
            titleEl.style.color = c.highContrast ? c.text : T["color.neutral.ink"];
            const summaryEl = this.el("div", "lx-summary", el);
            const noteEl = this.el("div", "lx-note", el);
            const svg = this.createSvg(el);
            return { el, titleEl, summaryEl, noteEl, svg, model: models[k] };
        });

        // Shared Y axis: the largest bar or curve peak across panels.
        let yMax: number | undefined;
        if (sm.sameYAxis.value) yMax = Math.max(...this.panels.map(p => this.yExtent(p.model)));
        for (const p of this.panels) {
            const titleH = p.titleEl ? p.titleEl.offsetHeight : 0;
            this.renderPanel(p, pw, ph - titleH, { yMax, notes: [], panel: true });
        }
        this.applySelectionState();
    }

    /** Notes about the data set up (Frequency, truncation) for the whole visual. */
    private dataNotes(data: HistogramData): string[] {
        const notes: string[] = [];
        if (!data.hasFrequency) notes.push(this.t("Note_Distinct", data.valueName));
        if (data.frequencyIssue === "sumOfValue") {
            notes.push(this.t("Note_FreqSum", data.valueName));
            if (data.unrecoverableZeros) notes.push(this.t("Note_FreqZeros", this.formatCount(data.unrecoverableZeros)));
        }
        if (data.frequencyIssue === "notWhole") notes.push(this.t("Note_FreqDecimals", data.valueName));
        if (data.truncated) notes.push(this.t("Note_Truncated", this.formatCount(ROW_LIMIT)));
        return notes;
    }

    private yMeasure(): string {
        return String(this.settings.yAxis.measure.value?.value ?? "count");
    }

    /** Top of a model's chart in Y units: tallest bar, or the normal-curve peak if higher. */
    private yExtent(m: Model): number {
        const measure = this.yMeasure();
        let top = 0;
        for (const b of m.bins) {
            let v = b.count;
            if (measure === "percent") v = m.total > 0 ? b.count / m.total : 0;
            else if (measure === "density") {
                const w = b.kind === "regular" ? (b.x1 - b.x0) : m.regularWidth;
                v = m.total > 0 && w > 0 ? b.count / (m.total * w) : 0;
            }
            top = Math.max(top, v);
        }
        if (this.settings.overlays.normalCurve.value && m.summary.stdDev > 0 && m.regularWidth > 0) {
            const scale = measure === "percent" ? m.regularWidth : measure === "density" ? 1 : m.total * m.regularWidth;
            top = Math.max(top, normalPdf(m.summary.mean, m.summary.mean, m.summary.stdDev) * scale);
        }
        return top;
    }

    private renderPanel(p: Panel, width: number, height: number, opts: { yMax?: number; notes: string[]; panel?: boolean }): void {
        const m = p.model;
        const s = this.settings;
        const c = this.colors();
        p.svg.style("display", null);

        const compact = height < 140 || width < 180;

        // --- summary strip and note (HTML, wraps naturally)
        this.renderSummary(p.summaryEl, m, c, s.summary.show.value && !compact, height, !!opts.panel);
        p.noteEl.textContent = opts.notes.join(" ");
        p.noteEl.style.display = opts.notes.length && !compact ? "block" : "none";

        const top = (p.summaryEl.style.display === "none" ? 0 : p.summaryEl.offsetHeight)
            + (p.noteEl.style.display === "none" ? 0 : p.noteEl.offsetHeight);
        for (const el of [p.summaryEl, p.noteEl]) {
            this.usedHeights.set(el, el.offsetHeight);
            this.sizeWatch?.observe(el);
        }
        const svgH = Math.max(0, height - top);
        p.svg.attr("width", width).attr("height", svgH);
        p.svg.selectAll("*").remove();
        p.svg.attr("role", "group").attr("aria-label", this.t("Aria_Chart", m.data.valueName) + (p.titleEl ? `, ${p.titleEl.textContent}` : ""));

        // --- measures
        const measure = String(s.yAxis.measure.value?.value ?? "count");
        const yOf = (b: Bin, count = b.count): number => {
            if (measure === "percent") return m.total > 0 ? count / m.total : 0;
            if (measure === "density") {
                const w = b.kind === "regular" ? (b.x1 - b.x0) : m.regularWidth;
                return m.total > 0 && w > 0 ? count / (m.total * w) : 0;
            }
            return count;
        };
        const curveScale = measure === "percent" ? m.regularWidth : measure === "density" ? 1 : m.total * m.regularWidth;
        const showCurve = s.overlays.normalCurve.value && m.summary.stdDev > 0 && m.regularWidth > 0;
        let yMax = Math.max(...m.bins.map(b => yOf(b)), 0);
        if (showCurve) yMax = Math.max(yMax, normalPdf(m.summary.mean, m.summary.mean, m.summary.stdDev) * curveScale);
        if (opts.yMax !== undefined && opts.yMax > 0) yMax = opts.yMax;
        if (!(yMax > 0)) yMax = 1;

        const yDigits = measure === "count"
            ? (Number.isInteger(m.total) ? 0 : 1)
            : Math.min(6, Math.max(0, (measure === "percent" ? 1 : 2) - Math.floor(Math.log10(yMax * (measure === "percent" ? 100 : 1)))));
        const yIntl = new Intl.NumberFormat(this.locale(), measure === "percent"
            ? { style: "percent", maximumFractionDigits: Math.max(0, yDigits - 1) }
            : { maximumFractionDigits: yDigits });
        const yFormat = { format: (v: number) => yIntl.format(v) };

        // --- margins from measured text
        const yFont = px(s.yAxis.fontSize.value), xFont = px(s.xAxis.fontSize.value);
        const showY = s.yAxis.show.value && !compact || (s.yAxis.show.value && width >= 120);
        const showX = s.xAxis.show.value;
        const yTicks = svgH < 220 ? 2 : 5;
        const yScaleTmp = scaleLinear().domain([0, yMax]).nice(yTicks);
        const yTickLabels = yScaleTmp.ticks(yTicks).map(v => yFormat.format(v));
        const textW = (t: string, size: number) => textMeasurementService.measureSvgTextWidth({ text: t, fontFamily: FONT, fontSize: `${size}px` });
        const showCumulative = s.overlays.cumulative.value;
        const margin = {
            top: 20,
            left: showY ? Math.ceil(Math.max(...yTickLabels.map(t => textW(t, yFont)), 0)) + 10 + (s.yAxis.showTitle.value && !compact ? yFont + 6 : 0) : 6,
            // Room for the last x label (centred on the right edge) or the cumulative % axis.
            right: Math.max(showCumulative ? textW("100%", yFont) + 12 : 10,
                showX && m.bins.length ? textW(m.format(m.bins[m.bins.length - 1].kind === "overflow" ? m.bins[m.bins.length - 1].x0 : m.bins[m.bins.length - 1].x1), xFont) / 2 + 4 : 0),
            bottom: showX ? xFont + 10 + (s.xAxis.showTitle.value && !compact ? xFont + 8 : 0) : 6
        };
        const plotW = Math.max(10, width - margin.left - margin.right);
        // Slots: every bin gets equal width (as in Excel); overlays map values onto the regular span.
        const slotW = plotW / m.bins.length;
        const firstReg = m.bins.findIndex(b => b.kind === "regular");
        const lastReg = firstReg < 0 ? -1 : firstReg + m.bins.filter(b => b.kind === "regular").length - 1;
        const regX0 = firstReg >= 0 ? firstReg * slotW : 0;
        const regX1 = lastReg >= 0 ? (lastReg + 1) * slotW : plotW;
        const regV0 = firstReg >= 0 ? m.bins[firstReg].x0 : 0;
        const regV1 = lastReg >= 0 ? m.bins[lastReg].x1 : 1;
        const xOfValue = (v: number): number => {
            if (regV1 === regV0) return (regX0 + regX1) / 2;
            const px = regX0 + (v - regV0) / (regV1 - regV0) * (regX1 - regX0);
            return Math.max(0, Math.min(plotW, px));
        };

        const markers: { v: number; label: string; color: string; dash: string }[] = [];
        if (m.lsl !== undefined) markers.push({ v: m.lsl, label: `${this.t("Marker_LSL")} ${m.format(m.lsl)}`, color: c.limit, dash: "5,3" });
        if (m.usl !== undefined) markers.push({ v: m.usl, label: `${this.t("Marker_USL")} ${m.format(m.usl)}`, color: c.limit, dash: "5,3" });
        const nLimits = markers.length;
        if (m.target !== undefined) markers.push({ v: m.target, label: `${this.t("Marker_Target")} ${m.format(m.target)}`, color: c.target, dash: "2,2" });
        if (s.overlays.meanLine.value) markers.push({ v: m.summary.mean, label: `${this.t("Marker_Mean")} ${m.formatPrecise(m.summary.mean)}`, color: c.stat, dash: "" });
        if (s.overlays.medianLine.value) markers.push({ v: m.summary.median, label: `${this.t("Marker_Median")} ${m.formatPrecise(m.summary.median)}`, color: c.stat, dash: "6,2,1,2" });
        // Marker labels sit above the plot. Clamp each inside the plot width (so the LSL label
        // never runs into the y-axis labels) and stack clashing labels upward, growing the top margin.
        // Short charts keep only the spec-limit labels; target, mean and median keep their lines.
        const shortChart = svgH < 220;
        const markerLayout = this.layoutMarkers(markers.map((mk, k) => ({
            ...mk, x: xOfValue(mk.v), label: shortChart && k >= nLimits ? "" : mk.label
        })), plotW, xFont);
        const markerRows = markerLayout.reduce((r, mk) => Math.max(r, mk.row + 1), 0);
        if (markerRows > 1) margin.top += (markerRows - 1) * (xFont + 3);
        const plotH = Math.max(10, svgH - margin.top - margin.bottom);
        if (svgH < 40 || width < 60) return; // too small to draw anything meaningful

        const g = p.svg.append("g").attr("transform", `translate(${margin.left},${margin.top})`);
        const y = scaleLinear().domain([0, yMax]).nice(yTicks).range([plotH, 0]);

        // --- gridlines and y axis
        if (s.yAxis.gridlines.value && !c.highContrast) {
            g.append("g").attr("class", "lx-grid").selectAll("line").data(y.ticks(yTicks)).enter().append("line")
                .attr("x1", 0).attr("x2", plotW).attr("y1", d => y(d)).attr("y2", d => y(d))
                .attr("stroke", c.grid).attr("stroke-width", 1);
        }
        if (showY) {
            const ax = g.append("g").attr("class", "lx-axis").call(axisLeft(y).ticks(yTicks).tickSize(0).tickPadding(6).tickFormat(v => yFormat.format(v as number)));
            this.styleAxis(ax, c.highContrast ? c.axis : (s.yAxis.fontColor.value?.value || T["color.neutral.muted"]), yFont);
            if (s.yAxis.showTitle.value && !compact) {
                const titles: Record<string, string> = { count: this.t(m.data.hasFrequency ? "Axis_Count" : "Axis_CountDistinct"), percent: this.t("Axis_Percent"), density: this.t("Axis_Density") };
                g.append("text").attr("class", "lx-title")
                    .attr("transform", `translate(${-margin.left + yFont},${plotH / 2}) rotate(-90)`)
                    .attr("text-anchor", "middle").attr("font-size", yFont).attr("fill", c.highContrast ? c.axis : (s.yAxis.fontColor.value?.value || T["color.neutral.muted"]))
                    .text(titles[measure] ?? "Count");
            }
        }

        // --- bars
        const gapFrac = Math.min(0.5, Math.max(0, s.bars.gap.value / 100));
        const barInset = Math.min(slotW * gapFrac / 2, slotW / 2 - 0.5);
        const outOfSpec = (b: Bin): boolean => {
            const mid = b.kind === "underflow" ? b.x1 : b.kind === "overflow" ? b.x0 + Number.EPSILON : (b.x0 + b.x1) / 2;
            return (m.lsl !== undefined && mid < m.lsl) || (m.usl !== undefined && mid > m.usl);
        };
        const barsG = g.append("g").attr("class", "lx-bars").attr("role", "listbox").attr("aria-multiselectable", "true");
        m.bins.forEach((b, i) => {
            const x = i * slotW + barInset;
            const w = Math.max(0.5, slotW - 2 * barInset);
            const v = yOf(b);
            const bar = barsG.append("rect").attr("class", "lx-bar")
                .attr("x", x).attr("width", w)
                .attr("y", y(v)).attr("height", Math.max(0, plotH - y(v)))
                .attr("fill", outOfSpec(b) && !c.highContrast ? c.outOfSpec : c.bar)
                .attr("data-bin", i)
                .attr("tabindex", 0)
                .attr("role", "option")
                .attr("aria-label", this.binLabel(m, b) + `: ${this.formatCount(b.count)} (${(b.share * 100).toFixed(1)}%)`);
            if (c.highContrast) bar.attr("stroke", c.axis).attr("stroke-width", 1);

            if (m.highlightByBin) {
                const hv = yOf(b, m.highlightByBin[i]);
                bar.attr("fill-opacity", 0.35);
                barsG.append("rect").attr("class", "lx-hl").attr("pointer-events", "none")
                    .attr("x", x).attr("width", w).attr("y", y(hv)).attr("height", Math.max(0, plotH - y(hv)))
                    .attr("fill", outOfSpec(b) && !c.highContrast ? c.outOfSpec : (c.highContrast ? c.axis : c.bar));
            }

            if (s.bars.showLabels.value && b.count > 0 && w >= 14) {
                const label = measure === "percent" ? `${(b.share * 100).toFixed(1)}%` : this.formatCount(b.count);
                barsG.append("text").attr("class", "lx-label").attr("pointer-events", "none")
                    .attr("x", x + w / 2).attr("y", y(v) - 3).attr("text-anchor", "middle")
                    .attr("font-size", px(s.bars.fontSize.value)).attr("fill", c.label).text(label);
            }
        });
        this.bindBarEvents(barsG, p);

        // --- x axis: bin edges (thinned to fit), with under/overflow slot labels
        if (showX) this.renderXAxis(g, m, slotW, plotH, firstReg, lastReg, xFont, c, compact);

        // --- overlays
        if (showCurve) {
            const pts: string[] = [];
            const steps = Math.max(40, Math.min(200, Math.round(regX1 - regX0)));
            for (let k = 0; k <= steps; k++) {
                const v = regV0 + (regV1 - regV0) * k / steps;
                const yy = y(normalPdf(v, m.summary.mean, m.summary.stdDev) * curveScale);
                pts.push(`${(k === 0 ? "M" : "L")}${xOfValue(v).toFixed(2)},${Math.max(0, yy).toFixed(2)}`);
            }
            g.append("path").attr("class", "lx-curve").attr("d", pts.join("")).attr("fill", "none")
                .attr("stroke", c.curve).attr("stroke-width", 1.5).attr("pointer-events", "none");
        }

        this.renderMarkers(g, markerLayout, plotH, xFont);

        if (showCumulative) {
            const yc = scaleLinear().domain([0, 1]).range([plotH, 0]);
            const pts = [`M0,${plotH}`].concat(m.bins.map((b, i) => `L${((i + 1) * slotW).toFixed(2)},${yc(b.cumulativeShare).toFixed(2)}`));
            g.append("path").attr("class", "lx-cumulative").attr("d", pts.join("")).attr("fill", "none")
                .attr("stroke", c.cumulative).attr("stroke-width", 1.5).attr("pointer-events", "none");
            const ax = g.append("g").attr("class", "lx-axis").attr("transform", `translate(${plotW},0)`)
                .call(axisRight(yc).tickValues([0, 0.25, 0.5, 0.75, 1]).tickSize(0).tickPadding(6).tickFormat(v => `${Math.round((v as number) * 100)}%`));
            this.styleAxis(ax, c.highContrast ? c.axis : c.cumulative, yFont);
        }

        g.append("line").attr("class", "lx-baseline").attr("x1", 0).attr("x2", plotW).attr("y1", plotH).attr("y2", plotH)
            .attr("stroke", c.highContrast ? c.axis : T["color.neutral.400"]).attr("stroke-width", 1);
    }

    private renderXAxis(g: D3Sel<SVGGElement>, m: Model, slotW: number, plotH: number, firstReg: number, lastReg: number,
        size: number, c: Colors, compact: boolean): void {
        const s = this.settings;
        const color = c.highContrast ? c.axis : (s.xAxis.fontColor.value?.value || T["color.neutral.muted"]);
        const ax = g.append("g").attr("class", "lx-axis lx-xaxis").attr("transform", `translate(0,${plotH})`);
        const widthOf = (t: string) => textMeasurementService.measureSvgTextWidth({ text: t, fontFamily: FONT, fontSize: `${size}px` });
        const hasUnder = m.bins.some(b => b.kind === "underflow");
        const hasOver = m.bins.some(b => b.kind === "overflow");
        const items: { x: number; text: string; w: number }[] = [];
        if (firstReg >= 0) {
            // The edge next to an underflow/overflow bin repeats that bin's "≤ a" / "> b" label, so skip it.
            for (let i = hasUnder ? firstReg + 1 : firstReg; i <= (hasOver ? lastReg : lastReg + 1); i++) {
                const v = i <= lastReg ? m.bins[i].x0 : m.bins[lastReg].x1;
                const text = m.format(v);
                items.push({ x: i * slotW, text, w: widthOf(text) });
            }
        }
        // Degenerate single bin (all values equal): one centred label.
        if (firstReg >= 0 && firstReg === lastReg && m.bins[firstReg].x0 === m.bins[firstReg].x1) {
            const text = m.format(m.bins[firstReg].x0);
            items.splice(0, items.length, { x: (firstReg + 0.5) * slotW, text, w: widthOf(text) });
        }
        // Underflow/overflow labels are always shown; edge labels must not collide with them.
        const reserved = m.bins.map((b, i) => ({ b, i })).filter(o => o.b.kind !== "regular").map(o => {
            const text = o.b.kind === "underflow" ? `≤ ${m.format(o.b.x1)}` : `> ${m.format(o.b.x0)}`;
            return { x: (o.i + 0.5) * slotW, text, w: widthOf(text) };
        });
        const gap = 8;
        const clear = (a: { x: number; w: number }, b: { x: number; w: number }) => Math.abs(a.x - b.x) - (a.w + b.w) / 2 >= gap;
        // Thin edge labels at a regular stride (every 1st, 2nd, 5th… edge) so the spacing stays even.
        const free = items.filter(it => reserved.every(r => clear(r, it)));
        let kept: typeof items = [];
        for (const stride of [1, 2, 3, 4, 5, 10, 20, 25, 50, 100, 200, 500]) {
            kept = free.filter((_, k) => k % stride === 0);
            if (kept.every((it, k) => k === 0 || clear(kept[k - 1], it))) break;
        }
        const last = free[free.length - 1];
        if (last && kept[kept.length - 1] !== last && (!kept.length || clear(kept[kept.length - 1], last))) kept.push(last);
        kept.concat(reserved).forEach(it => {
            if (!reserved.includes(it)) ax.append("line").attr("x1", it.x).attr("x2", it.x).attr("y1", 0).attr("y2", 4).attr("stroke", color);
            ax.append("text").attr("x", it.x).attr("y", size + 5).attr("text-anchor", "middle")
                .attr("font-size", size).attr("fill", color).text(it.text);
        });
        if (s.xAxis.showTitle.value && !compact) {
            ax.append("text").attr("class", "lx-title").attr("x", (m.bins.length * slotW) / 2).attr("y", 2 * size + 12)
                .attr("text-anchor", "middle").attr("font-size", size).attr("fill", color).text(m.data.valueName);
        }
    }

    private layoutMarkers(markers: { x: number; label: string; color: string; dash: string }[], plotW: number, size: number):
        { x: number; labelX: number; row: number; label: string; color: string; dash: string }[] {
        const sorted = markers.slice().sort((a, b) => a.x - b.x);
        const placed: { labelX: number; w: number; row: number }[] = [];
        return sorted.map(mk => {
            if (!mk.label) return { ...mk, labelX: mk.x, row: 0 }; // line only
            const w = textMeasurementService.measureSvgTextWidth({ text: mk.label, fontFamily: FONT, fontSize: `${size}px` });
            const labelX = w >= plotW ? plotW / 2 : Math.max(w / 2, Math.min(plotW - w / 2, mk.x));
            let row = 0;
            while (placed.some(p => p.row === row && Math.abs(p.labelX - labelX) < (p.w + w) / 2 + 6)) row++;
            placed.push({ labelX, w, row });
            return { ...mk, labelX, row };
        });
    }

    private renderMarkers(g: D3Sel<SVGGElement>, markers: { x: number; labelX: number; row: number; label: string; color: string; dash: string }[], plotH: number, size: number): void {
        for (const mk of markers) {
            g.append("line").attr("class", "lx-marker").attr("pointer-events", "none")
                .attr("x1", mk.x).attr("x2", mk.x).attr("y1", 0).attr("y2", plotH)
                .attr("stroke", mk.color).attr("stroke-width", 1.5).attr("stroke-dasharray", mk.dash || null);
            if (!mk.label) continue;
            g.append("text").attr("class", "lx-marker-label").attr("pointer-events", "none")
                .attr("x", mk.labelX).attr("y", -4 - mk.row * (size + 3)).attr("text-anchor", "middle")
                .attr("font-size", size).attr("fill", mk.color).text(mk.label);
        }
    }

    private renderSummary(el: HTMLElement, m: Model, c: Colors, show: boolean, availHeight: number, panel = false): void {
        el.replaceChildren();
        el.style.display = show ? "flex" : "none";
        if (!show) return;
        el.style.color = c.text;
        const fontSizePx = px(this.settings.summary.fontSize.value);
        el.style.fontSize = `${fontSizePx}px`;

        // Dynamic max height based on available visual height:
        // Visuals >= 260px height comfortably accommodate up to 3 badge rows so narrowing width never drops stats.
        // Moderate visuals (>= 180px) allow 2 rows. Short visuals get 1 row.
        const rowHeight = Math.round(fontSizePx * 1.3 + 12);
        const maxRows = availHeight >= 260 ? 3 : (availHeight >= 180 ? 2 : 1);
        const maxH = maxRows * rowHeight + 6;
        el.style.maxHeight = `${maxH}px`;

        const items: { el: HTMLElement; priority: number }[] = [];
        // priority: 1 = keep longest; higher numbers are dropped first when space runs out.
        const add = (label: string, value: string, status?: "met" | "at-risk" | "not-met", tip?: string, priority = 3) => {
            const item = this.el("span", "lx-stat", el);
            items.push({ el: item, priority });
            if (tip) item.title = tip;
            this.el("span", "lx-stat-label", item).textContent = label;
            const v = this.el("span", "lx-stat-value", item);
            let displayVal = value;
            if (status === "met") displayVal += " ✔";
            else if (status === "at-risk") displayVal += " ▲";
            else if (status === "not-met") displayVal += " ✖";
            v.textContent = displayVal;

            if (status && !c.highContrast) {
                item.classList.add(`lx-stat-${status}`);
            }
        };
        const sm = m.summary;
        add(this.t("Stat_N"), this.formatCount(sm.count), undefined, undefined, 1);
        add(this.t("Stat_Mean"), m.formatPrecise(sm.mean), undefined, undefined, 1);
        add(this.t("Stat_StdDev"), m.formatPrecise(sm.stdDev), undefined, undefined, 2);
        // Panels show the same short set everywhere, so they compare line by line.
        if (!panel) {
            add(this.t("Stat_Median"), m.formatPrecise(sm.median), undefined, undefined, 4);
            add(this.t("Stat_Min"), m.format(sm.min), undefined, undefined, 5);
            add(this.t("Stat_Max"), m.format(sm.max), undefined, undefined, 5);
        }
        if (m.outOfSpecShare !== undefined) {
            const status = m.outOfSpecShare > 0 ? "not-met" : "met";
            add(this.t("Stat_OutOfSpec"), `${(m.outOfSpecShare * 100).toFixed(1)}%`, status, undefined, 1);
        }
        const cap = m.capability;
        if (cap) {
            const idx = new Intl.NumberFormat(this.locale(), { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            if (cap.pp !== undefined && !panel) {
                const ppStatus = cap.pp >= 1.33 ? "met" : cap.pp >= 1.0 ? "at-risk" : "not-met";
                add(this.t("Stat_Pp"), idx.format(cap.pp), ppStatus, this.t("Stat_Pp_Tip"), 3);
            }
            if (cap.ppk !== undefined) {
                const ppkStatus = cap.ppk >= 1.33 ? "met" : cap.ppk >= 1.0 ? "at-risk" : "not-met";
                add(this.t("Stat_Ppk"), idx.format(cap.ppk), ppkStatus, this.t("Stat_Ppk_Tip"), 1);
            }
            if (!panel) add(this.t("Stat_ExpectedOut"), this.formatShare(cap.expectedOutShare), undefined, this.t("Stat_ExpectedOut_Tip"), 3);
        }
        if (m.blanks > 0 && !panel) add(this.t("Stat_Blanks"), this.formatCount(m.blanks), undefined, undefined, 6);

        // Never show a half-cut line: drop the least important stats (a whole priority level at a
        // time, so Min and Max go together) until everything fits.
        const levels = Array.from(new Set(items.map(i => i.priority))).sort((a, b) => b - a);
        for (const level of levels) {
            if (el.scrollHeight <= el.clientHeight + 2 || level === 1) break;
            items.filter(i => i.priority === level).forEach(i => i.el.remove());
        }
    }

    private showLanding(): void {
        this.svg.style("display", "none");
        this.gridEl.style.display = "none";
        this.gridEl.replaceChildren();
        this.panels = [];
        this.summaryEl.style.display = "none";
        this.noteEl.style.display = "none";
        const el = this.landingEl;
        el.replaceChildren();
        el.style.display = "flex";
        const card = this.el("div", "lx-landing-card", el);
        this.el("div", "lx-landing-title", card).textContent = this.t("Landing_Title");
        this.el("div", "lx-landing-sub", card).textContent = this.t("Landing_Sub");
        const ol = this.el("ol", "lx-landing-steps", card);
        ["Landing_Step1", "Landing_Step2", "Landing_Step3"].map(k => this.t(k)).forEach(t => { this.el("li", "", ol).textContent = t; });
    }

    private showMessage(title: string, body: string): void {
        this.showLanding();
        const card = this.landingEl.firstElementChild as HTMLElement;
        card.replaceChildren();
        this.el("div", "lx-landing-title", card).textContent = title;
        this.el("div", "lx-landing-sub", card).textContent = body;
    }

    // ------------------------------------------------------------------ interaction

    private bindBarEvents(barsG: D3Sel<SVGGElement>, p: Panel): void {
        const node = barsG.node()!;
        const m = p.model;
        const binAt = (e: Event): number | null => {
            const t = (e.target as Element).closest(".lx-bar");
            return t ? Number(t.getAttribute("data-bin")) : null;
        };
        const choose = (i: number, multi: boolean) => {
            if (!this.interactive) return;
            const ids = m.rowsByBin[i].map(r => m.data.ids[r]);
            if (!ids.length) return;
            const already = this.isBinSelected(m, i);
            const p = (already && !multi) ? this.selectionManager.clear() : this.selectionManager.select(ids, multi);
            p.then(() => this.applySelectionState());
        };
        node.addEventListener("click", (e: MouseEvent) => {
            const i = binAt(e);
            if (i === null) return;
            choose(i, e.ctrlKey || e.metaKey || e.shiftKey);
            e.stopPropagation();
        });
        node.addEventListener("keydown", (e: KeyboardEvent) => {
            const i = binAt(e);
            if (i === null) return;
            if (e.key === "Enter" || e.key === " ") {
                choose(i, e.ctrlKey || e.metaKey || e.shiftKey);
                e.preventDefault();
            }
        });
        node.addEventListener("contextmenu", (e: MouseEvent) => {
            const i = binAt(e);
            if (i === null || !this.interactive) return;
            const first = m.rowsByBin[i][0];
            this.selectionManager.showContextMenu(first !== undefined ? m.data.ids[first] : ({} as ISelectionId), { x: e.clientX, y: e.clientY });
            e.preventDefault();
            e.stopPropagation();
        });
        const tip = (e: MouseEvent, kind: "show" | "move") => {
            const i = binAt(e);
            if (i === null) { this.tooltipService.hide({ immediately: true, isTouchEvent: false }); return; }
            const rect = this.root.getBoundingClientRect();
            const args = {
                coordinates: [e.clientX - rect.left, e.clientY - rect.top],
                isTouchEvent: false,
                dataItems: this.tooltipItems(p, i),
                identities: m.rowsByBin[i].map(r => m.data.ids[r])
            };
            if (kind === "show") this.tooltipService.show(args); else this.tooltipService.move(args);
        };
        node.addEventListener("mouseover", (e) => tip(e, "show"));
        node.addEventListener("mousemove", (e) => tip(e, "move"));
        node.addEventListener("mouseleave", () => this.tooltipService.hide({ immediately: false, isTouchEvent: false }));
    }

    private tooltipItems(p: Panel, i: number): VisualTooltipDataItem[] {
        const m = p.model;
        const b = m.bins[i];
        const items: VisualTooltipDataItem[] = [];
        if (p.titleEl && m.data.splitName) items.push({ displayName: m.data.splitName, value: p.titleEl.textContent || "" });
        items.push(
            { displayName: m.data.valueName, value: this.binLabel(m, b) },
            { displayName: this.t(m.data.hasFrequency ? "Tip_Count" : "Tip_Distinct"), value: this.formatCount(b.count) },
            { displayName: this.t("Tip_Percent"), value: `${(b.share * 100).toFixed(1)}%` },
            { displayName: this.t("Tip_Cumulative"), value: `${(b.cumulativeShare * 100).toFixed(1)}%` }
        );
        if (m.highlightByBin) items.push({ displayName: this.t("Tip_Highlighted"), value: this.formatCount(m.highlightByBin[i]) });
        return items;
    }

    private isBinSelected(m: Model, i: number): boolean {
        const selected = this.selectionManager.getSelectionIds() as ISelectionId[];
        if (!selected.length) return false;
        const keys = new Set(selected.map(id => id.getKey()));
        return m.rowsByBin[i].some(r => keys.has(m.data.ids[r].getKey()));
    }

    /** Dim unselected bins; mark selected ones for screen readers. */
    private applySelectionState(): void {
        if (!this.model) return;
        const selected = this.selectionManager.getSelectionIds() as ISelectionId[];
        const keys = new Set(selected.map(id => id.getKey()));
        const any = keys.size > 0;
        const c = this.colors();
        const targets = this.multiples ? this.panels : [{ svg: this.svg, model: this.model }];
        for (const { svg, model: m } of targets) svg.selectAll<SVGRectElement, unknown>(".lx-bar").each(function () {
            const i = Number(this.getAttribute("data-bin"));
            const on = !any || m.rowsByBin[i].some(r => keys.has(m.data.ids[r].getKey()));
            this.setAttribute("aria-selected", String(any && on));
            this.style.opacity = on ? "1" : "0.4";
            if (c.highContrast) this.setAttribute("stroke", any && on ? c.selectedStroke : c.axis);
        });
    }

    // ------------------------------------------------------------------ helpers

    /** Localized string; {0}, {1}… are replaced with the args. Falls back to the key. */
    private t(key: string, ...args: (string | undefined)[]): string {
        let text = this.localization.getDisplayName(key) || key;
        args.forEach((a, k) => { text = text.split(`{${k}}`).join(a ?? ""); });
        return text;
    }

    /** False on dashboard tiles and other read-only hosts: no selection or context menu then. */
    private get interactive(): boolean {
        return this.host.hostCapabilities?.allowInteractions !== false;
    }

    private binLabel(m: Model, b: Bin): string {
        if (b.kind === "underflow") return `≤ ${m.format(b.x1)}`;
        if (b.kind === "overflow") return `> ${m.format(b.x0)}`;
        const first = m.bins.findIndex(x => x.kind === "regular") === m.bins.indexOf(b);
        return `${first ? "[" : "("}${m.format(b.x0)}, ${m.format(b.x1)}]`;
    }

    /** Exact counts with thousands separators (no "2K" display units). */
    /** Share as a percentage, plus parts per million when small: "0.48% (4,817 ppm)". */
    private formatShare(share: number): string {
        const pct = share * 100;
        const digits = pct >= 10 ? 1 : pct >= 0.1 ? 2 : 3;
        const p = new Intl.NumberFormat(this.locale(), { maximumFractionDigits: digits, minimumFractionDigits: Math.min(digits, 1) }).format(pct);
        if (share >= 0.01) return `${p}%`;
        const ppm = new Intl.NumberFormat(this.locale(), { maximumFractionDigits: share * 1e6 < 10 ? 1 : 0 }).format(share * 1e6);
        return `${p}% (${ppm} ppm)`;
    }

    private formatCount(v: number): string {
        return new Intl.NumberFormat(this.locale(), { maximumFractionDigits: Number.isInteger(v) ? 0 : 1 }).format(v);
    }

    private locale(): string {
        try { return Intl.NumberFormat.supportedLocalesOf([this.host.locale || "en-US"])[0] || "en-US"; } catch { return "en-US"; }
    }

    private styleAxis(ax: D3Sel<SVGGElement>, color: string, size: number): void {
        ax.select(".domain").remove();
        ax.selectAll("text").attr("fill", color).attr("font-size", size).attr("font-family", FONT);
    }

    /** A chart surface. Clicking its empty space clears the selection. */
    private createSvg(parent: HTMLElement): D3Sel<SVGSVGElement> {
        const node = document.createElementNS(SVG_NS, "svg") as SVGSVGElement;
        node.setAttribute("class", "lx-chart");
        parent.appendChild(node);
        node.addEventListener("click", (e) => {
            if (!this.interactive || (e.target as Element).closest(".lx-bar")) return;
            this.selectionManager.clear().then(() => this.applySelectionState());
        });
        return select(node);
    }

    private el(tag: string, cls: string, parent: HTMLElement = this.root): HTMLElement {
        const e = document.createElement(tag);
        if (cls) e.className = cls;
        parent.appendChild(e);
        return e;
    }
}
