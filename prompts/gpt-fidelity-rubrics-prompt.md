# ChatGPT 5.2 Pro Handoff: Excel Library Fidelity Scoring Rubrics

## Context

We're building **ExcelBench**, a comprehensive benchmark suite that compares Excel libraries' feature parity across Python and Rust ecosystems. Unlike typical benchmarks that focus on performance, we're measuring **feature fidelity** - how accurately each library can read and write Excel features compared to native Excel.

This is a research-grade project intended to serve as a definitive reference for developers choosing between libraries.

## Your Task

Design detailed **fidelity scoring rubrics** for each Excel feature we're testing. Each feature needs a clear 0-3 scoring definition that:
- Is objective and reproducible (two people scoring the same library should get the same result)
- Captures meaningful gradations of support (not just "works/doesn't work")
- Accounts for the read vs. write distinction where relevant
- Identifies specific edge cases that differentiate score levels

## Scoring Scale

- **0 - Unsupported**: Feature cannot be read or written at all. Library throws an error, silently ignores, or corrupts data.
- **1 - Minimal**: Basic recognition of the feature exists but with significant limitations or data loss. Usable only in trivial cases.
- **2 - Functional**: Feature works for common use cases. Some edge cases may fail or have minor fidelity loss. Suitable for most real-world usage.
- **3 - Complete**: Full fidelity with native Excel. All edge cases handled. Round-trip preserves all data. Indistinguishable from Excel's own output.

## Features Requiring Rubrics

### Tier 1 - Essential

1. **Cell Values**
   - String, number, date, boolean, error values
   - Edge cases: very long strings, Unicode/emoji, scientific notation, dates before 1900, blank vs. empty string

2. **Formulas**
   - Basic functions: SUM, AVERAGE, COUNT, IF, VLOOKUP, INDEX/MATCH
   - Key distinction: Can the library read the formula string? Can it read the cached/calculated value? Can it evaluate formulas itself?
   - Edge cases: array formulas (legacy and dynamic), circular references, external references, volatile functions (NOW, RAND), structured references

3. **Basic Text Formatting**
   - Bold, italic, underline, strikethrough
   - Font family, font size, font color
   - Edge cases: rich text (multiple formats in one cell), theme colors vs. explicit RGB, font that doesn't exist on system

4. **Cell Background Color**
   - Solid fill colors
   - Edge cases: pattern fills, gradient fills, theme colors

5. **Number Formats**
   - Built-in formats: currency, percentage, date/time, accounting
   - Custom format strings
   - Edge cases: locale-specific formats, format strings with conditions, four-part format strings (positive;negative;zero;text)

6. **Cell Alignment**
   - Horizontal: left, center, right, justify, distributed
   - Vertical: top, middle, bottom
   - Text wrap, shrink to fit
   - Text rotation (0-180 degrees, vertical text)
   - Indent level
   - Edge cases: rotation + wrap interaction, distributed alignment with indent

7. **Borders**
   - Border positions: top, bottom, left, right, diagonal
   - Border styles: thin, medium, thick, dashed, dotted, double, hair, etc.
   - Border colors
   - Edge cases: diagonal borders (up, down, both), border style "none" vs. no border, medium-dashed vs. dashed distinction

8. **Column Widths / Row Heights**
   - Explicit dimensions
   - Auto-fit / best-fit flags
   - Hidden rows/columns
   - Edge cases: very narrow columns, default width vs. explicit width, width unit conversion accuracy

9. **Multiple Sheets**
   - Reading/writing multiple worksheets
   - Sheet names (including special characters, Unicode)
   - Sheet order, active sheet, sheet visibility (visible, hidden, very hidden)
   - Edge cases: maximum sheet name length, duplicate handling, sheet color tabs

### Tier 2 - Standard

10. **Merged Cells**
    - Reading merged ranges, writing merged ranges
    - Edge cases: value in top-left vs. other cells of merge, formatting of merged cells, partial overlap detection

11. **Conditional Formatting**
    - Rule types: cell value rules, top/bottom rules, above/below average, duplicate/unique, text contains
    - Data bars, color scales, icon sets
    - Formula-based rules
    - Edge cases: multiple rules on same range, rule priority/stop-if-true, rules referencing other sheets

12. **Data Validation**
    - Validation types: list (dropdown), whole number, decimal, date, time, text length, custom formula
    - Input message, error alert (style, title, message)
    - Edge cases: list from range vs. explicit list, cross-sheet validation, allowing blank

13. **Hyperlinks**
    - URL links, email links, file links
    - Internal links (to cell, to sheet)
    - Link display text vs. target
    - Edge cases: very long URLs, special characters in URLs, tooltip text

14. **Images/Embedded Objects**
    - Inserting images (PNG, JPEG, GIF, etc.)
    - Image positioning (cell-anchored vs. free-floating)
    - Image sizing, aspect ratio
    - Edge cases: image compression, embedded vs. linked images, alt text

15. **Pivot Tables**
    - Reading pivot table structure and data
    - Creating basic pivot tables
    - Row/column/value/filter fields
    - Aggregation functions (sum, count, average, etc.)
    - Edge cases: calculated fields, grouping (date grouping, numeric ranges), multiple value fields, pivot table styles

16. **Comments and Notes**
    - Cell comments (threaded comments in modern Excel)
    - Legacy notes
    - Author, timestamp, reply threads
    - Edge cases: formatted text in comments, comment positioning/sizing, resolved status

17. **Freeze Panes / Split Views**
    - Freeze rows, freeze columns, freeze both
    - Split position
    - Edge cases: freeze with hidden rows/columns, pane selection state

### Tier 3 - Advanced

18. **Charts**
    - Chart types: column, bar, line, pie, scatter, area, combo
    - Chart elements: title, legend, axis labels, data labels, gridlines
    - Data source references
    - Edge cases: secondary axis, trendlines, error bars, chart in chart sheet vs. embedded

19. **Named Ranges**
    - Workbook-scoped names, sheet-scoped names
    - Names referencing ranges, formulas, constants
    - Edge cases: names with special characters, dynamic named ranges (using OFFSET/INDEX), name conflicts

20. **Complex Conditional Formatting**
    - Nested rules, formula-based with complex logic
    - Color scales with custom midpoints
    - Icon sets with custom thresholds
    - Edge cases: rules using relative references, rules across non-contiguous ranges

21. **Tables (Structured References)**
    - Creating/reading Excel Tables (ListObjects)
    - Table styles, header row, total row
    - Structured reference syntax in formulas
    - Edge cases: table resize, calculated columns, table intersections

22. **Print Settings**
    - Page orientation, margins, paper size
    - Print area, print titles (repeat rows/columns)
    - Headers and footers (including dynamic fields like page number, date)
    - Page breaks (manual and automatic)
    - Edge cases: scaling (fit to page), print quality, black & white setting

23. **Protection**
    - Sheet protection (with password, allowed actions)
    - Workbook protection (structure, windows)
    - Cell-level protection (locked, hidden formula)
    - Edge cases: protection without password, specific user permissions (Excel 2007+)

## Deliverable Format

For each feature, provide:

```markdown
## [Feature Name]

### Scoring Rubric

| Score | Read Capability | Write Capability |
|-------|-----------------|------------------|
| 0 | [specific criteria] | [specific criteria] |
| 1 | [specific criteria] | [specific criteria] |
| 2 | [specific criteria] | [specific criteria] |
| 3 | [specific criteria] | [specific criteria] |

### Key Edge Cases to Test
- [Edge case 1]: Tests [what aspect]
- [Edge case 2]: Tests [what aspect]
- ...

### Scoring Notes
[Any additional guidance for consistent scoring, common ambiguities, etc.]
```

## Important Considerations

1. **Read vs. Write may differ**: A library might perfectly read borders but write them incorrectly, or vice versa. The rubric should allow independent scoring.

2. **"Partial" must be specific**: Don't just say "partial support" - specify exactly what subset counts as level 1 vs. level 2.

3. **Corruption vs. Omission**: Distinguish between a library that silently drops a feature (might be acceptable) vs. one that corrupts the file or causes Excel to show repair dialogs (never acceptable, should be 0).

4. **Theme vs. Explicit**: Many features have "theme-aware" versions (theme colors, theme fonts). Decide whether theme support is required for level 3 or is a separate consideration.

5. **Version differences**: Excel has evolved (conditional formatting expanded significantly in 2007/2010, threaded comments in Office 365). Note which Excel version features appeared in if relevant to scoring.

## Example Output (for reference)

Here's a partial example of what we're looking for:

---

## Borders

### Scoring Rubric

| Score | Read Capability | Write Capability |
|-------|-----------------|------------------|
| 0 | Cannot detect borders exist; throws error or ignores | Cannot create any borders; file corrupted or borders missing |
| 1 | Detects border presence but loses style/color detail (e.g., all borders become "thin black") | Can create borders but limited to basic styles (thin/medium/thick, black only) |
| 2 | Correctly reads border position, style, and color for common styles; may lose exotic styles (hair, slantDashDot) or diagonal borders | Creates borders with full style and color support; diagonal borders may be unsupported |
| 3 | Perfect fidelity: all 13 border styles, all positions including diagonals, RGB and theme colors preserved | Creates any border style Excel can, including diagonals; theme colors mapped correctly |

### Key Edge Cases to Test
- **Diagonal borders**: Both up-diagonal and down-diagonal simultaneously
- **All 13 styles**: Particularly hair, mediumDashed, dashDot, mediumDashDot, dashDotDot, mediumDashDotDot, slantDashDot
- **Theme colors**: Border using theme color (e.g., "Accent1") vs. explicit RGB
- **Mixed borders**: Cell with different styles on each edge
- **Border vs. no border**: Distinguishing "none" style from absence of border element

### Scoring Notes
- If a library converts theme colors to their RGB equivalent on read, this is acceptable for score 2 (minor fidelity loss).
- If diagonal borders are silently dropped but all other borders preserved, this is score 2.
- Border "none" explicitly set should be distinguishable from never having had a border; if this distinction is lost, cap at score 2.

---

Please generate complete rubrics for all 23 features listed above with this level of detail.
