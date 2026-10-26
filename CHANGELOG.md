# Changelog: Histogram by Lean Excel

Every published version has an entry. Format: version, date, then Added / Changed / Fixed.

## 1.1.2 (to submit for AppSource and certification)

### Fixed
- When Power BI drew the visual while its container was hidden or not laid out yet (seen on a dashboard in the Power BI service), the chart could overflow its box and cut off the x-axis. The visual now notices when the summary strip's real height differs from the one it used and redraws once.

## 1.1.1 (not published; superseded by 1.1.2)

### Fixed
- Very large values (16-digit numbers, Microsoft's validation test) showed every axis label as "1E+15". Values from a billion upwards now use Power BI display units (bn, T), and values beyond 10^15 show enough decimals in scientific notation to tell bins apart. Smaller values are unchanged.

### Changed
- Package description no longer uses "Excel-style"; it says the bins match Microsoft Excel (Microsoft trademark guidelines).
- Package GitHub link points at the certification repository.

## 1.1.0 (tested by the owner; not published)

### Added
- **Histogram Pro: Split by** (small multiples). One panel per group with identical bins, same Y axis (toggle), automatic or fixed grid, up to 12 panels by default. Each panel shows n, mean, std dev, out of spec and Ppk; cross-filter, highlight and tooltips work per panel. Needs a Histogram Pro licence (Microsoft-managed, one-month free trial).
- Licence check through Microsoft's Licensing API. Without a licence, or where Power BI can't check licences, Split by shows all data combined with a short note. Nothing else is ever blocked.
- **Metric Badges for Summary Statistics**: individual styled pills/chips with subtle backgrounds, borders, and Lean Excel status indicators (✔ / ▲ / ✖) for capability ($Pp$, $Ppk$) and out-of-spec % with high-contrast support.

### Changed
- Support link in the package now goes to the Histogram's own page, www.leanexcel.com/support/histogram-by-lean-excel (company standard: one privacy, support and terms page per product).

### Fixed
- **Visual resize layout robustness**: Fixed summary strip from being crushed into a sliver by flexbox when visual height decreases. Adaptive multi-row wrapping (up to 3 rows on normal charts) prevents capability statistics ($Pp$, $Ppk$, Expected out) from being dropped when width is narrowed.
- **Marker label collision**: Increased top margin clearance so spec limit labels (LSL, USL, Target) and mean lines never overlap the summary bar.

### Changed
- Short charts and small panels keep only the spec-limit labels (target, mean and median keep their lines) and use fewer y-axis ticks. Compact mode threshold lowered to 140px height so compact cards retain summary statistics.

## 1.0.1 (superseded by 1.1.0)

From the owner's first Power BI Desktop test and a review against the main competitor.

### Added
- Pp, Ppk and expected out of spec (normal curve, in % and ppm) in the summary strip when spec limits are set. Hover a figure for a plain explanation. Toggle: Specification limits > "Show Pp, Ppk and expected out of spec". Matches Excel's formulas.

### Fixed
- Frequency left on **Sum** of the Value column (instead of Count) weighted each value by its size: wrong n, mean, shape, and far too many bins. The visual now detects this, recovers the true counts and shows a note on how to fix the field. Frequency with decimals also shows a note.
- Spec-limit and mean/median labels no longer overlap the y-axis labels or each other: labels stay inside the plot and stack upwards.
- The summary strip wraps to a second line instead of hiding statistics on narrower visuals. If two lines aren't enough, the least important figures (blanks, then min and max, then median) drop out whole; no half-cut text.
- X-axis labels: no more overlap between the "≤ a" / "> b" labels and their neighbouring edge; edge labels are thinned at an even step (every 2nd, 5th… edge) instead of an irregular one.

### Changed
- Version 1.0.1 so Power BI Desktop loads the new build instead of a cached 1.0.0.

## 1.0.0 (release candidate, superseded by 1.0.1)

### Brand
- Lean Excel brand standard applied: default colours from brand tokens, store logo and screenshots in Lean Green, support at www.leanexcel.com, contact hello@leanexcel.com.

### Added
- Excel-compatible histogram: automatic bins (Scott's rule, Excel's default), Freedman–Diaconis, Sturges, square root, fixed bin count or bin width.
- Right-closed bins like Excel, with optional underflow (≤) and overflow (>) bins.
- Optional round bin edges (e.g. 42, 44, 46).
- Exact counts using a Frequency field (Count of Value) or pre-counted data.
- Y axis as count, % of total or density.
- Overlays: normal curve, mean line, median line, cumulative % line with its own axis.
- Specification limits (LSL, USL, target) from fixed values or measures, out-of-spec bars and "% out of spec".
- Summary strip: n, mean, std dev, median, min, max, out of spec, blanks excluded.
- Cross-filtering by bin (Ctrl+click to add bins), highlighting from other visuals, context menu, report-page tooltips for the whole bin, keyboard navigation, high-contrast mode, read-only tile support.
- Theme-aware colours, report locale number formats, up to 30,000 distinct values.
