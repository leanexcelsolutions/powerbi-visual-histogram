/**
 * Statistics engine shared by the Statistical Pack (histogram, box plot, Pareto)
 * and, later, SPC Pro. Pure functions, no DOM, no Power BI API, so they can be
 * unit-tested in Node and reused in any visual.
 *
 * Conventions
 * - Inputs may contain null / undefined / NaN / Infinity: they are dropped, never thrown on.
 * - Quantiles default to the "inclusive" method (Excel QUARTILE.INC / R type 7) so
 *   numbers match what Excel users already see. "exclusive" matches QUARTILE.EXC.
 * - Standard deviation is the sample (n - 1) version, matching Excel STDEV.S.
 */

export type QuantileMethod = 'inclusive' | 'exclusive';

/** Keep finite numbers only and sort ascending. */
export function cleanSorted(values: ReadonlyArray<number | null | undefined>): number[] {
  const out: number[] = [];
  for (const v of values) {
    if (typeof v === 'number' && Number.isFinite(v)) out.push(v);
  }
  return out.sort((a, b) => a - b);
}

export function mean(sorted: ReadonlyArray<number>): number {
  if (sorted.length === 0) return NaN;
  let sum = 0;
  for (const v of sorted) sum += v;
  return sum / sorted.length;
}

/** Sample standard deviation (n - 1). Returns 0 for a single value, NaN for none. */
export function stdDev(sorted: ReadonlyArray<number>): number {
  const n = sorted.length;
  if (n === 0) return NaN;
  if (n === 1) return 0;
  const m = mean(sorted);
  let ss = 0;
  for (const v of sorted) ss += (v - m) * (v - m);
  return Math.sqrt(ss / (n - 1));
}

/**
 * Quantile of an ascending-sorted array.
 * inclusive: h = (n - 1) p      (Excel QUARTILE.INC, PERCENTILE.INC)
 * exclusive: h = (n + 1) p - 1  (Excel QUARTILE.EXC), clamped to the data range
 */
export function quantile(sorted: ReadonlyArray<number>, p: number, method: QuantileMethod = 'inclusive'): number {
  const n = sorted.length;
  if (n === 0 || !Number.isFinite(p)) return NaN;
  if (n === 1) return sorted[0];
  const clampedP = Math.min(1, Math.max(0, p));
  let h = method === 'exclusive' ? (n + 1) * clampedP - 1 : (n - 1) * clampedP;
  h = Math.min(n - 1, Math.max(0, h));
  const lo = Math.floor(h);
  const hi = Math.ceil(h);
  return sorted[lo] + (h - lo) * (sorted[hi] - sorted[lo]);
}

export interface Summary {
  count: number;
  min: number;
  max: number;
  mean: number;
  stdDev: number;
  q1: number;
  median: number;
  q3: number;
  iqr: number;
  /** Q1 - k * IQR */
  lowerFence: number;
  /** Q3 + k * IQR */
  upperFence: number;
  /** Lowest value inside the fences (box plot whisker end). */
  whiskerLow: number;
  /** Highest value inside the fences (box plot whisker end). */
  whiskerHigh: number;
  outliers: number[];
}

/**
 * Five-number summary plus Tukey fences. k = 1.5 is the standard box plot rule;
 * k = 3 flags only "far out" values.
 */
export function summarize(
  values: ReadonlyArray<number | null | undefined>,
  options: { method?: QuantileMethod; fenceK?: number } = {}
): Summary {
  const sorted = cleanSorted(values);
  const method = options.method ?? 'inclusive';
  const k = options.fenceK ?? 1.5;
  const n = sorted.length;
  if (n === 0) {
    return {
      count: 0, min: NaN, max: NaN, mean: NaN, stdDev: NaN, q1: NaN, median: NaN, q3: NaN,
      iqr: NaN, lowerFence: NaN, upperFence: NaN, whiskerLow: NaN, whiskerHigh: NaN, outliers: []
    };
  }
  const q1 = quantile(sorted, 0.25, method);
  const median = quantile(sorted, 0.5, method);
  const q3 = quantile(sorted, 0.75, method);
  const iqr = q3 - q1;
  const lowerFence = q1 - k * iqr;
  const upperFence = q3 + k * iqr;
  const inside = sorted.filter(v => v >= lowerFence && v <= upperFence);
  return {
    count: n,
    min: sorted[0],
    max: sorted[n - 1],
    mean: mean(sorted),
    stdDev: stdDev(sorted),
    q1,
    median,
    q3,
    iqr,
    lowerFence,
    upperFence,
    whiskerLow: inside.length ? inside[0] : sorted[0],
    whiskerHigh: inside.length ? inside[inside.length - 1] : sorted[n - 1],
    outliers: sorted.filter(v => v < lowerFence || v > upperFence)
  };
}

// ---------------------------------------------------------------------------
// Histogram binning
// ---------------------------------------------------------------------------

export type BinRule = 'scott' | 'freedman-diaconis' | 'sturges' | 'sqrt';

export type BinKind = 'underflow' | 'regular' | 'overflow';

export interface Bin {
  kind: BinKind;
  /** Lower edge. Exclusive, except the first regular bin which includes the minimum (Excel convention). */
  x0: number;
  /** Upper edge, inclusive. Underflow bin: values <= x1. Overflow bin: values > x0. */
  x1: number;
  count: number;
  /** count / total, 0..1 */
  share: number;
  /** running share including this bin, 0..1 */
  cumulativeShare: number;
}

export interface BinOptions {
  /** Automatic rule. Default 'scott' (what Excel's histogram chart uses). */
  rule?: BinRule;
  /** Fixed number of bins; overrides rule. */
  binCount?: number;
  /** Fixed bin width; overrides binCount and rule. */
  binWidth?: number;
  /** Where the first regular bin starts; defaults to the data minimum (or the underflow value). */
  start?: number;
  /** Values <= underflow go into one leading "underflow" bin (Excel "Underflow bin"). */
  underflow?: number;
  /** Values > overflow go into one trailing "overflow" bin (Excel "Overflow bin"). */
  overflow?: number;
  /**
   * Round the automatic bin width up to 1, 2, 2.5 or 5 x 10^k and align edges to
   * multiples of it (e.g. 42, 44, 46 instead of 42.80, 43.94, 45.09). Ignored when
   * binWidth is set. Not Excel's default, so off unless asked for.
   */
  nice?: boolean;
  /** Upper cap on regular bins so a bad setting cannot freeze the visual. Default 500. */
  maxBins?: number;
}

/** Smallest "nice" step (1, 2, 2.5, 5 x 10^k) that is >= raw. */
export function niceStep(raw: number): number {
  if (!(raw > 0) || !Number.isFinite(raw)) return 1;
  const exp = Math.floor(Math.log10(raw));
  const base = Math.pow(10, exp);
  for (const m of [1, 2, 2.5, 5, 10]) {
    const step = m * base;
    if (step >= raw * (1 - 1e-12)) return step;
  }
  return 10 * base;
}

/** Suggested number of bins for a rule; always >= 1. */
export function suggestBinCount(sorted: ReadonlyArray<number>, rule: BinRule = 'scott'): number {
  const n = sorted.length;
  if (n <= 1) return 1;
  return suggestFromStats(n, sorted[n - 1] - sorted[0], stdDev(sorted),
    quantile(sorted, 0.75) - quantile(sorted, 0.25), rule);
}

function suggestFromStats(n: number, range: number, sd: number, iqr: number, rule: BinRule): number {
  if (n <= 1 || !(range > 0)) return 1;
  const sturges = Math.max(1, Math.ceil(Math.log2(n) + 1));
  let width: number;
  switch (rule) {
    case 'sturges': return sturges;
    case 'sqrt': return Math.max(1, Math.ceil(Math.sqrt(n)));
    case 'freedman-diaconis': width = 2 * iqr / Math.cbrt(n); break;
    case 'scott':
    default: width = 3.49 * sd / Math.cbrt(n); break;
  }
  // Degenerate spread (e.g. IQR = 0): fall back to Sturges.
  if (!(width > 0)) return sturges;
  return Math.max(1, Math.ceil(range / width - 1e-9));
}

/**
 * Excel-compatible histogram bins. Bins are right-closed, (x0, x1], and the first
 * regular bin also includes its lower edge, so the minimum is always counted.
 * Optional weights (same length as values) support pre-aggregated frequencies;
 * a value with weight w counts w times. Non-finite or negative weights count as 0.
 */
export function createBins(
  values: ReadonlyArray<number | null | undefined>,
  options: BinOptions = {},
  weights?: ReadonlyArray<number | null | undefined>
): Bin[] {
  return binValues(values, options, weights).bins;
}

export interface BinResult {
  bins: Bin[];
  /** For each input index, the index into bins, or -1 if the value was dropped (blank or zero weight). */
  binOf: number[];
}

/** Same as createBins, plus which bin every input row landed in (for selection and highlighting). */
export function binValues(
  values: ReadonlyArray<number | null | undefined>,
  options: BinOptions = {},
  weights?: ReadonlyArray<number | null | undefined>
): BinResult {
  const binOf = new Array<number>(values.length).fill(-1);
  // Triples of [value, weight, input index], dropping invalid values.
  const pairs: Array<[number, number, number]> = [];
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (typeof v !== 'number' || !Number.isFinite(v)) continue;
    let w = 1;
    if (weights) {
      const raw = weights[i];
      w = typeof raw === 'number' && Number.isFinite(raw) && raw > 0 ? raw : 0;
    }
    if (w > 0) pairs.push([v, w, i]);
  }
  if (pairs.length === 0) return { bins: [], binOf };
  pairs.sort((a, b) => a[0] - b[0]);
  const total = pairs.reduce((t, p) => t + p[1], 0);

  const hasUnder = options.underflow !== undefined && Number.isFinite(options.underflow);
  const hasOver = options.overflow !== undefined && Number.isFinite(options.overflow);
  const under = hasUnder ? (options.underflow as number) : -Infinity;
  const over = hasOver ? (options.overflow as number) : Infinity;
  const maxBins = options.maxBins ?? 500;

  let underCount = 0;
  let overCount = 0;
  const middle: Array<[number, number, number]> = [];
  const UNDER = -2, OVER = -3; // provisional markers, resolved to final indices below
  for (const p of pairs) {
    if (hasUnder && p[0] <= under) { underCount += p[1]; binOf[p[2]] = UNDER; }
    else if (hasOver && p[0] > over) { overCount += p[1]; binOf[p[2]] = OVER; }
    else middle.push(p);
  }

  const regular: Array<{ x0: number; x1: number; count: number }> = [];
  if (middle.length > 0) {
    const min = middle[0][0];
    const max = middle[middle.length - 1][0];
    let start = hasUnder ? under : min;
    if (options.start !== undefined && Number.isFinite(options.start) && options.start <= min) start = options.start;
    const end = hasOver ? Math.max(over, max) : max;

    let end2 = end;
    if (end === start) {
      regular.push({ x0: start, x1: end, count: middle.reduce((t, p) => t + p[1], 0) });
      for (const p of middle) binOf[p[2]] = 0;
    } else {
      let width: number;
      const fixedWidth = !!(options.binWidth && options.binWidth > 0);
      if (fixedWidth) {
        width = options.binWidth as number;
      } else {
        const count = options.binCount && options.binCount > 0
          ? Math.floor(options.binCount)
          : ruleBinCount(middle, start, end, options.rule ?? 'scott');
        width = (end - start) / count;
        if (options.nice) {
          width = niceStep(width);
          // Align the start to a multiple of the step unless the user fixed it or an underflow edge sets it.
          const userStart = options.start !== undefined && Number.isFinite(options.start) && options.start <= min;
          if (!userStart && !hasUnder) start = Math.floor(start / width + 1e-9) * width;
          if (!hasOver) end2 = start + Math.ceil((end - start) / width - 1e-9) * width;
        }
      }
      const snapEnd = !fixedWidth && !options.nice;
      let binCount = Math.max(1, Math.ceil((end2 - start) / width - 1e-9));
      if (binCount > maxBins) {
        binCount = maxBins;
        width = (end2 - start) / binCount;
      }
      // Edges computed as start + i * width, then rounded to the step's precision to avoid 0.30000000000000004.
      const dec = Math.min(12, Math.max(0, 10 - Math.floor(Math.log10(Math.abs(width) || 1))));
      const edge = (i: number) => Number((start + i * width).toFixed(dec));
      for (let i = 0; i < binCount; i++) regular.push({ x0: edge(i), x1: edge(i + 1), count: 0 });
      // Snap the last edge to the true end to avoid floating-point drift.
      if (snapEnd) regular[binCount - 1].x1 = end;
      for (const [v, w, idx] of middle) {
        let i = Math.ceil((v - start) / width - 1e-9) - 1; // right-closed
        if (i < 0) i = 0;
        if (i >= binCount) i = binCount - 1;
        regular[i].count += w;
        binOf[idx] = i;
      }
    }
  }

  const out: Array<{ kind: BinKind; x0: number; x1: number; count: number }> = [];
  if (hasUnder) out.push({ kind: 'underflow', x0: middle.length ? Math.min(pairs[0][0], under) : pairs[0][0], x1: under, count: underCount });
  for (const r of regular) out.push({ kind: 'regular', ...r });
  if (hasOver) out.push({ kind: 'overflow', x0: over, x1: Math.max(pairs[pairs.length - 1][0], over), count: overCount });

  // Resolve indices: regular bins shift by one when an underflow bin leads.
  const offset = hasUnder ? 1 : 0;
  for (let i = 0; i < binOf.length; i++) {
    const b = binOf[i];
    if (b === UNDER) binOf[i] = 0;
    else if (b === OVER) binOf[i] = out.length - 1;
    else if (b >= 0) binOf[i] = b + offset;
  }

  let running = 0;
  const bins = out.map(b => {
    running += b.count;
    return { ...b, share: b.count / total, cumulativeShare: running / total };
  });
  return { bins, binOf };
}

function ruleBinCount(sortedPairs: ReadonlyArray<readonly [number, number, ...number[]]>, start: number, end: number, rule: BinRule): number {
  const st = weightedStats(sortedPairs);
  return suggestFromStats(st.n, end - start, st.sd, st.iqr, rule);
}

/** Exact weighted statistics; weights act as integer frequencies (counts). */
function weightedStats(sortedPairs: ReadonlyArray<readonly [number, number, ...number[]]>) {
  let n = 0, sum = 0;
  for (const [v, w] of sortedPairs) { n += w; sum += v * w; }
  const m = n > 0 ? sum / n : NaN;
  let ss = 0;
  for (const [v, w] of sortedPairs) ss += w * (v - m) * (v - m);
  const sd = n > 1 ? Math.sqrt(ss / (n - 1)) : (n === 1 ? 0 : NaN);
  // Value at 0-based rank k in the expanded (repeated) sorted list.
  const at = (k: number): number => {
    let cum = 0;
    for (const [v, w] of sortedPairs) { cum += w; if (k < cum) return v; }
    return sortedPairs[sortedPairs.length - 1][0];
  };
  const q = (p: number): number => {
    if (n <= 0) return NaN;
    const h = (n - 1) * p, lo = Math.floor(h);
    const a = at(lo), b = at(Math.min(Math.ceil(h), n - 1));
    return a + (h - lo) * (b - a);
  };
  const q1 = q(0.25), q3 = q(0.75);
  return { n, mean: m, sd, q1, median: q(0.5), q3, iqr: q3 - q1 };
}

/**
 * Summary for pre-aggregated data: weights[i] is how many times values[i] occurs
 * (the Frequency measure). Same conventions as summarize(); outliers are listed
 * once per distinct value.
 */
export function weightedSummary(
  values: ReadonlyArray<number | null | undefined>,
  weights?: ReadonlyArray<number | null | undefined>,
  options: { fenceK?: number } = {}
): Summary {
  if (!weights) return summarize(values, options);
  const pairs: Array<[number, number]> = [];
  values.forEach((v, i) => {
    const w = weights[i];
    if (typeof v === 'number' && Number.isFinite(v) && typeof w === 'number' && Number.isFinite(w) && w > 0) pairs.push([v, w]);
  });
  if (pairs.length === 0) return summarize([]);
  pairs.sort((a, b) => a[0] - b[0]);
  const st = weightedStats(pairs);
  const k = options.fenceK ?? 1.5;
  const lowerFence = st.q1 - k * st.iqr, upperFence = st.q3 + k * st.iqr;
  const inside = pairs.filter(([v]) => v >= lowerFence && v <= upperFence);
  return {
    count: st.n, min: pairs[0][0], max: pairs[pairs.length - 1][0], mean: st.mean, stdDev: st.sd,
    q1: st.q1, median: st.median, q3: st.q3, iqr: st.iqr, lowerFence, upperFence,
    whiskerLow: inside.length ? inside[0][0] : pairs[0][0],
    whiskerHigh: inside.length ? inside[inside.length - 1][0] : pairs[pairs.length - 1][0],
    outliers: pairs.filter(([v]) => v < lowerFence || v > upperFence).map(([v]) => v)
  };
}

/** Normal probability density, for the optional bell-curve overlay. */
export function normalPdf(x: number, mu: number, sigma: number): number {
  if (!(sigma > 0)) return NaN;
  const z = (x - mu) / sigma;
  return Math.exp(-0.5 * z * z) / (sigma * Math.sqrt(2 * Math.PI));
}

/**
 * Standard normal cumulative distribution, same as Excel NORM.S.DIST(z, TRUE).
 * Uses erfc via W. J. Cody's rational approximations (relative error below 1e-14),
 * so tail probabilities stay accurate in parts per million and beyond.
 */
export function normalCdf(x: number, mu = 0, sigma = 1): number {
  if (!(sigma > 0)) return NaN;
  const z = (x - mu) / sigma;
  return 0.5 * erfc(-z / Math.SQRT2);
}

function erfc(x: number): number {
  const ax = Math.abs(x);
  let r: number;
  if (ax < 0.5) {
    // erf series region
    const t = x * x;
    const top = (((0.185777706184603153 * t + 3.16112374387056560) * t + 113.864154151050156) * t + 377.485237685302021) * t + 3209.37758913846947;
    const bot = (((t + 23.6012909523441209) * t + 244.024637934444173) * t + 1282.61652607737228) * t + 2844.23683343917062;
    return 1 - x * top / bot;
  } else if (ax < 4) {
    const top = (((((((2.15311535474403846e-8 * ax + 0.564188496988670089) * ax + 8.88314979438837594) * ax + 66.1191906371416295) * ax
      + 298.635138197400131) * ax + 881.952221241769090) * ax + 1712.04761263407058) * ax + 2051.07837782607147) * ax + 1230.33935479799725;
    const bot = (((((((ax + 15.7449261107098347) * ax + 117.693950891312499) * ax + 537.181101862009858) * ax + 1621.38957456669019) * ax
      + 3290.79923573345963) * ax + 4362.61909014324716) * ax + 3439.36767414372164) * ax + 1230.33935480374942;
    r = Math.exp(-ax * ax) * top / bot;
  } else {
    const z = 1 / (ax * ax);
    const top = ((((0.0163153871373020978 * z + 0.305326634961232344) * z + 0.360344899949804439) * z + 0.125781726111229246) * z + 0.0160837851487422766) * z + 6.58749161529837803e-4;
    const bot = ((((z + 2.56852019228982242) * z + 1.87295284992346725) * z + 0.527905102951428412) * z + 0.0605183413124413191) * z + 2.33520497626869185e-3;
    r = Math.exp(-ax * ax) / ax * (0.564189583547756287 - z * top / bot);
  }
  return x < 0 ? 2 - r : r;
}

export interface Capability {
  /** (USL − LSL) / 6σ. Needs both limits. */
  pp?: number;
  /** Min of the one-sided indices below. */
  ppk?: number;
  /** (USL − mean) / 3σ */
  ppu?: number;
  /** (mean − LSL) / 3σ */
  ppl?: number;
  /** Share of a normal distribution with this mean and σ that falls outside the limits. */
  expectedOutShare: number;
}

/**
 * Overall (long-term) process performance from the mean and overall standard deviation
 * (STDEV.S). Same as the Excel formulas (USL−LSL)/(6*STDEV.S) and
 * MIN((USL−AVERAGE)/(3*STDEV.S), (AVERAGE−LSL)/(3*STDEV.S)).
 * Cp/Cpk need a within-subgroup sigma from time-ordered data, which a histogram doesn't have,
 * so they are deliberately not computed here.
 */
export function capability(mean: number, sigma: number, lsl?: number, usl?: number): Capability | null {
  if (!(sigma > 0) || !Number.isFinite(mean)) return null;
  const hasL = lsl !== undefined && Number.isFinite(lsl);
  const hasU = usl !== undefined && Number.isFinite(usl);
  if (!hasL && !hasU) return null;
  if (hasL && hasU && (usl as number) <= (lsl as number)) return null;
  const out: Capability = { expectedOutShare: 0 };
  if (hasU) {
    out.ppu = ((usl as number) - mean) / (3 * sigma);
    out.expectedOutShare += 1 - normalCdf(usl as number, mean, sigma);
  }
  if (hasL) {
    out.ppl = (mean - (lsl as number)) / (3 * sigma);
    out.expectedOutShare += normalCdf(lsl as number, mean, sigma);
  }
  if (hasL && hasU) out.pp = ((usl as number) - (lsl as number)) / (6 * sigma);
  out.ppk = Math.min(out.ppu ?? Infinity, out.ppl ?? Infinity);
  return out;
}

// ---------------------------------------------------------------------------
// Pareto
// ---------------------------------------------------------------------------

export interface ParetoItem<T = unknown> {
  label: string;
  value: number;
  /** Anything the visual needs to carry through (e.g. a selection id). */
  payload?: T;
}

export interface ParetoRow<T = unknown> extends ParetoItem<T> {
  rank: number;
  share: number;
  cumulativeShare: number;
  /** True for the "vital few": items up to and including the one that crosses the threshold. */
  isVitalFew: boolean;
}

/**
 * Sort descending, compute share and cumulative share, flag the vital few.
 * Negative and non-finite values are dropped (Pareto is only meaningful for
 * non-negative contributions); ties keep their input order.
 */
export function pareto<T>(items: ReadonlyArray<ParetoItem<T>>, threshold = 0.8): ParetoRow<T>[] {
  const valid = items
    .map((item, index) => ({ item, index }))
    .filter(({ item }) => Number.isFinite(item.value) && item.value >= 0)
    .sort((a, b) => b.item.value - a.item.value || a.index - b.index)
    .map(({ item }) => item);
  const total = valid.reduce((s, i) => s + i.value, 0);
  let running = 0;
  let crossed = false;
  return valid.map((item, i) => {
    running += item.value;
    const share = total > 0 ? item.value / total : 0;
    const cumulativeShare = total > 0 ? running / total : 0;
    const isVitalFew = !crossed;
    if (cumulativeShare >= threshold) crossed = true;
    return { ...item, rank: i + 1, share, cumulativeShare, isVitalFew };
  });
}
