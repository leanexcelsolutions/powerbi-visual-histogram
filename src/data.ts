/*
 * Converts the Power BI data view into a plain model. No DOM here.
 *
 * Data contract (capabilities.json):
 *   value      Grouping  numeric column. Power BI merges identical values into one row,
 *                        so without a Frequency measure each DISTINCT value counts once.
 *   frequency  Measure   usually "Count of <value>": gives exact counts per distinct value.
 *   lsl/usl/target Measure (optional) spec limits; first finite value wins.
 * Report-page tooltips receive the selection ids of every row in the hovered bin,
 * so Power BI aggregates the tooltip page correctly for the whole bin.
 */
import powerbi from "powerbi-visuals-api";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";

import DataView = powerbi.DataView;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import ISelectionId = powerbi.visuals.ISelectionId;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

export const ROW_LIMIT = 30000;

export interface HistogramData {
    values: (number | null)[];
    /** Present when a Frequency measure is bound. */
    weights?: number[];
    /** Per row: highlighted weight when another visual cross-highlights; undefined when not highlighting. */
    highlights?: number[];
    ids: ISelectionId[];
    hasFrequency: boolean;
    /** Weighted count of blank / non-numeric values (excluded from the chart). */
    blanks: number;
    /** True when Power BI returned the maximum number of rows, so data may be cut off. */
    truncated: boolean;
    valueName: string;
    /** Present when a Split by field is bound: the group label of every row. */
    split?: string[];
    splitName?: string;
    valueFormat?: string;
    lsl?: number;
    usl?: number;
    target?: number;
    /**
     * Frequency set up wrongly. "sumOfValue": Frequency is Sum of the Value column, so each weight
     * is value × count; counts were recovered as weight ÷ value. "notWhole": weights have decimals,
     * so Frequency is probably not a count.
     */
    frequencyIssue?: "sumOfValue" | "notWhole";
    /** With "sumOfValue": rows whose value is 0, so their count can't be recovered (left out). */
    unrecoverableZeros?: number;
}

export type DataState =
    | { kind: "empty" }                 // no Value field bound
    | { kind: "noNumbers"; data: HistogramData } // bound but nothing numeric
    | { kind: "ok"; data: HistogramData };

function roleOf(column: DataViewValueColumn, role: string): boolean {
    return !!column.source.roles && !!column.source.roles[role];
}

/** True when the Frequency column is Sum of the same column bound to Value (a common setup slip). */
export function isSumOfValue(freq: powerbi.DataViewMetadataColumn, value: powerbi.DataViewMetadataColumn): boolean {
    const fq = (freq.queryName || "").replace(/\s+/g, "").toLowerCase();
    const vq = (value.queryName || "").replace(/\s+/g, "").toLowerCase();
    if (fq && vq && fq === `sum(${vq})`) return true;
    // Fallback for hosts that don't fill queryName: Power BI's default display name.
    return !!freq.displayName && !!value.displayName && freq.displayName === `Sum of ${value.displayName}`;
}

/** Recover per-row counts from weights that are value × count. Returns null if they don't fit. */
function recoverCounts(values: (number | null)[], weights: number[]): { counts: number[]; zeros: number } | null {
    const counts = new Array<number>(weights.length);
    let zeros = 0;
    for (let i = 0; i < weights.length; i++) {
        const v = values[i];
        const w = weights[i];
        if (v === null) { counts[i] = 0; continue; }
        if (v === 0) { counts[i] = 0; zeros++; continue; }
        const k = w / v;
        const r = Math.round(k);
        if (r < 1 || Math.abs(k - r) > 1e-6 * Math.max(1, r)) return null;
        counts[i] = r;
    }
    return { counts, zeros };
}

/** Readable group label; blanks become "(Blank)" like native visuals. */
function splitLabel(v: powerbi.PrimitiveValue, format?: string): string {
    if (v === null || v === undefined || v === "") return "(Blank)";
    return valueFormatter.format(v, format);
}

function firstFinite(values: powerbi.PrimitiveValue[] | undefined): number | undefined {
    if (!values) return undefined;
    for (const v of values) {
        const n = typeof v === "number" ? v : NaN;
        if (Number.isFinite(n)) return n;
    }
    return undefined;
}

export function readData(dataView: DataView | undefined, host: IVisualHost): DataState {
    const categorical = dataView?.categorical;
    const cats = categorical?.categories ?? [];
    const category = cats.find(c => !!c.source.roles?.["value"]);
    if (!category) return { kind: "empty" };
    const splitCat = cats.find(c => !!c.source.roles?.["split"]);

    const columns = categorical.values ? Array.from(categorical.values) : [];
    const freqCol = columns.find(c => roleOf(c, "frequency"));
    const n = category.values.length;

    const values: (number | null)[] = new Array(n);
    const ids: ISelectionId[] = new Array(n);
    let weights: number[] | undefined;
    if (freqCol) weights = new Array(n);
    let blanks = 0;
    let numericRows = 0;

    for (let i = 0; i < n; i++) {
        const raw = category.values[i];
        const v = typeof raw === "number" && Number.isFinite(raw) ? raw : null;
        const fw = freqCol ? Number(freqCol.values[i]) : 1;
        const w = Number.isFinite(fw) && fw > 0 ? fw : 0;
        if (weights) weights[i] = w;
        values[i] = v;
        if (v === null) blanks += freqCol ? w : 1;
        else if (w > 0) numericRows++;
        const b = host.createSelectionIdBuilder().withCategory(category, i);
        if (splitCat) b.withCategory(splitCat, i);
        ids[i] = b.createSelectionId();
    }

    // Frequency set to Sum of the Value column instead of Count: recover the true counts.
    let frequencyIssue: HistogramData["frequencyIssue"];
    let unrecoverableZeros: number | undefined;
    let highlightScale: number[] | undefined;
    if (freqCol && weights && isSumOfValue(freqCol.source, category.source)) {
        const rec = recoverCounts(values, weights);
        if (rec) {
            highlightScale = weights.map((w, i) => (w > 0 ? rec.counts[i] / w : 0));
            weights = rec.counts;
            frequencyIssue = "sumOfValue";
            unrecoverableZeros = rec.zeros || undefined;
            blanks = 0; // a summed blank is blank; its count is unknown
        }
    }
    if (!frequencyIssue && weights && weights.some(w => w > 0 && Math.abs(w - Math.round(w)) > 1e-9)) {
        frequencyIssue = "notWhole";
    }
    if (frequencyIssue === "sumOfValue") {
        numericRows = 0;
        for (let i = 0; i < n; i++) if (values[i] !== null && weights![i] > 0) numericRows++;
    }

    // Highlights: use the frequency column's highlights if present, else any measure's.
    let highlights: number[] | undefined;
    const hlCol = (freqCol && freqCol.highlights) ? freqCol : columns.find(c => !!c.highlights);
    if (hlCol && hlCol.highlights) {
        highlights = new Array(n);
        for (let i = 0; i < n; i++) {
            const h = hlCol.highlights[i];
            if (hlCol === freqCol) {
                const hn = Number(h);
                const h0 = Number.isFinite(hn) && hn > 0 ? hn : 0;
                highlights[i] = highlightScale ? Math.round(h0 * highlightScale[i]) : h0;
            } else {
                highlights[i] = h !== null && h !== undefined ? (weights ? weights[i] : 1) : 0;
            }
        }
    }

    const data: HistogramData = {
        values,
        weights,
        highlights,
        ids,
        hasFrequency: !!freqCol,
        blanks,
        truncated: n >= ROW_LIMIT,
        valueName: category.source.displayName,
        split: splitCat ? splitCat.values.map(v => splitLabel(v, splitCat.source.format)) : undefined,
        splitName: splitCat?.source.displayName,
        valueFormat: category.source.format,
        lsl: firstFinite(columns.find(c => roleOf(c, "lsl"))?.values),
        usl: firstFinite(columns.find(c => roleOf(c, "usl"))?.values),
        target: firstFinite(columns.find(c => roleOf(c, "target"))?.values),
        frequencyIssue,
        unrecoverableZeros
    };
    return numericRows > 0 ? { kind: "ok", data } : { kind: "noNumbers", data };
}
