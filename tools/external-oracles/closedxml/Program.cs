using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO.Compression;
using ClosedXML.Excel;

try
{
    using var stdin = Console.OpenStandardInput();
    var request = JsonSerializer.Deserialize<OracleRequest>(stdin, JsonOptions.Value)
        ?? throw new InvalidOperationException("Empty oracle request.");

    var payload = request.Operation switch
    {
        "write_fixture" => WriteFixture(request),
        "read_metadata" => ReadMetadata(request),
        _ => throw new InvalidOperationException($"Unsupported operation '{request.Operation}'.")
    };

    Console.WriteLine(JsonSerializer.Serialize(payload, JsonOptions.Value));
    return 0;
}
catch (Exception ex)
{
    Console.WriteLine(JsonSerializer.Serialize(new
    {
        error = "closedxml_oracle_failed",
        message = ex.Message
    }, JsonOptions.Value));
    return 1;
}

static object WriteFixture(OracleRequest request)
{
    if (string.IsNullOrWhiteSpace(request.OutputPath))
    {
        throw new InvalidOperationException("write_fixture requires output_path.");
    }

    var payload = request.Payload.Deserialize<WritePayload>(JsonOptions.Value)
        ?? new WritePayload();

    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(request.OutputPath))!);
    using var workbook = new XLWorkbook();
    ConfigureSheets(workbook, payload.Sheets);
    ApplyCells(workbook, payload.Cells);
    ApplyRichText(workbook, payload.RichText);
    ApplyComments(workbook, payload.Comments);
    ApplyTables(workbook, payload.Tables);
    ApplyConditionalFormats(workbook, payload.ConditionalFormats);
    ApplyPivots(workbook, payload.Pivots);
    ApplyProtection(workbook, payload.Protection);
    workbook.SaveAs(request.OutputPath);

    return new
    {
        fixture_id = request.FixtureId,
        operation = request.Operation,
        output_path = request.OutputPath,
        tool = "closedxml",
        counts = new
        {
            sheets = workbook.Worksheets.Count,
            cells = payload.Cells.Count,
            rich_text = payload.RichText.Count,
            comments = payload.Comments.Count,
            tables = payload.Tables.Count,
            conditional_formats = payload.ConditionalFormats.Count,
            pivots = payload.Pivots.Count,
            protected_sheets = payload.Protection.Count
        }
    };
}

static object ReadMetadata(OracleRequest request)
{
    var inputPath = string.IsNullOrWhiteSpace(request.InputPath)
        ? request.OutputPath
        : request.InputPath;
    if (string.IsNullOrWhiteSpace(inputPath))
    {
        throw new InvalidOperationException("read_metadata requires input_path or output_path.");
    }

    using var workbookPackage = ZipFile.OpenRead(inputPath);
    var partNames = workbookPackage.Entries.Select(entry => entry.FullName).ToArray();
    return new
    {
        fixture_id = request.FixtureId,
        operation = request.Operation,
        input_path = inputPath,
        tool = "closedxml",
        counts = new
        {
            tables = CountParts(partNames, "tables/table"),
            pivot_tables = CountParts(partNames, "pivotTables/pivotTable"),
            pivot_caches = CountParts(partNames, "pivotCache/pivotCacheDefinition"),
            comments = CountParts(partNames, "comments"),
            vml_drawings = CountParts(partNames, "drawings/vmlDrawing"),
            worksheet_parts = CountParts(partNames, "worksheets/sheet")
        }
    };
}

static void ConfigureSheets(XLWorkbook workbook, List<SheetSpec> sheets)
{
    if (sheets.Count == 0)
    {
        workbook.Worksheets.Add("Sheet1");
        return;
    }

    foreach (var sheet in sheets)
    {
        workbook.Worksheets.Add(sheet.Name);
    }
}

static void ApplyCells(XLWorkbook workbook, List<CellSpec> cells)
{
    foreach (var cell in cells)
    {
        var target = workbook.Worksheet(cell.Sheet).Cell(cell.Cell);
        if (cell.Type == "formula")
        {
            target.FormulaA1 = cell.Formula ?? "";
            continue;
        }

        switch (cell.Value.ValueKind)
        {
            case JsonValueKind.String:
                target.Value = cell.Value.GetString() ?? "";
                break;
            case JsonValueKind.Number:
                target.Value = cell.Value.TryGetInt64(out var integer)
                    ? integer
                    : cell.Value.GetDouble();
                break;
            case JsonValueKind.True:
            case JsonValueKind.False:
                target.Value = cell.Value.GetBoolean();
                break;
            case JsonValueKind.Null:
            case JsonValueKind.Undefined:
                break;
            default:
                target.Value = cell.Value.ToString();
                break;
        }
    }
}

static void ApplyRichText(XLWorkbook workbook, List<RichTextSpec> items)
{
    foreach (var item in items)
    {
        var richText = workbook.Worksheet(item.Sheet).Cell(item.Cell).CreateRichText();
        foreach (var run in item.Runs)
        {
            var richString = richText.AddText(run.Text);
            if (run.Bold)
            {
                richString.SetBold();
            }
            if (run.Italic)
            {
                richString.SetItalic();
            }
            if (!string.IsNullOrWhiteSpace(run.FontColor))
            {
                richString.SetFontColor(XLColor.FromHtml(run.FontColor));
            }
        }
    }
}

static void ApplyComments(XLWorkbook workbook, List<CommentSpec> comments)
{
    foreach (var item in comments)
    {
        var comment = workbook.Worksheet(item.Sheet).Cell(item.Cell).CreateComment();
        if (!string.IsNullOrWhiteSpace(item.Author))
        {
            comment.Author = item.Author;
        }
        comment.AddText(item.Text);
    }
}

static void ApplyTables(XLWorkbook workbook, List<TableSpec> tables)
{
    foreach (var table in tables)
    {
        var range = workbook.Worksheet(table.Sheet).Range(table.Range);
        var created = range.CreateTable(table.Name);
        created.Theme = XLTableTheme.TableStyleMedium2;
    }
}

static void ApplyConditionalFormats(XLWorkbook workbook, List<ConditionalFormatSpec> formats)
{
    foreach (var format in formats)
    {
        var range = workbook.Worksheet(format.Sheet).Range(format.Range);
        switch (format.Type)
        {
            case "3_color_scale":
                range.AddConditionalFormat().ColorScale()
                    .LowestValue(XLColor.Red)
                    .Midpoint(XLCFContentType.Percent, 50, XLColor.Yellow)
                    .HighestValue(XLColor.Green);
                break;
            case "data_bar":
                range.AddConditionalFormat().DataBar(XLColor.Blue)
                    .LowestValue()
                    .HighestValue();
                break;
            case "cell":
                ApplyCellRule(range, format);
                break;
        }
    }
}

static void ApplyCellRule(IXLRange range, ConditionalFormatSpec format)
{
    if (!double.TryParse(format.Value, out var value))
    {
        value = 0;
    }

    var rule = format.Criteria switch
    {
        ">" => range.AddConditionalFormat().WhenGreaterThan(value),
        "<" => range.AddConditionalFormat().WhenLessThan(value),
        ">=" => range.AddConditionalFormat().WhenEqualOrGreaterThan(value),
        "<=" => range.AddConditionalFormat().WhenEqualOrLessThan(value),
        _ => range.AddConditionalFormat().WhenGreaterThan(value)
    };
    rule.Fill.SetBackgroundColor(XLColor.LightGreen);
}

static void ApplyPivots(XLWorkbook workbook, List<PivotSpec> pivots)
{
    foreach (var pivot in pivots)
    {
        var sourceRange = ResolveRange(workbook, pivot.DataRange);
        var (targetSheet, targetCell) = ResolveCell(workbook, pivot.Cell);
        var created = targetSheet.PivotTables.Add(pivot.Name, targetCell, sourceRange);
        foreach (var row in pivot.Rows)
        {
            created.RowLabels.Add(row.Name);
        }
        foreach (var column in pivot.Columns)
        {
            created.ColumnLabels.Add(column.Name);
        }
        foreach (var value in pivot.Data)
        {
            created.Values.Add(value.Name).SetSummaryFormula(XLPivotSummary.Sum);
        }
    }
}

static void ApplyProtection(XLWorkbook workbook, List<ProtectionSpec> protections)
{
    foreach (var protection in protections)
    {
        var worksheet = workbook.Worksheet(protection.Sheet);
        if (string.IsNullOrWhiteSpace(protection.Password))
        {
            worksheet.Protect();
        }
        else
        {
            worksheet.Protect(protection.Password);
        }
    }
}

static IXLRange ResolveRange(XLWorkbook workbook, string reference)
{
    var parts = reference.Split('!', 2);
    if (parts.Length != 2)
    {
        throw new InvalidOperationException($"Expected sheet-qualified range: {reference}");
    }
    return workbook.Worksheet(parts[0]).Range(parts[1]);
}

static (IXLWorksheet Sheet, IXLCell Cell) ResolveCell(XLWorkbook workbook, string reference)
{
    var parts = reference.Split('!', 2);
    if (parts.Length != 2)
    {
        throw new InvalidOperationException($"Expected sheet-qualified cell: {reference}");
    }
    var sheet = workbook.Worksheet(parts[0]);
    return (sheet, sheet.Cell(parts[1]));
}

static int CountParts(IEnumerable<string> partNames, string prefix)
{
    return partNames.Count(part => part.Contains(prefix, StringComparison.Ordinal));
}

sealed record OracleRequest(
    [property: JsonPropertyName("fixture_id")] string FixtureId,
    [property: JsonPropertyName("operation")] string Operation,
    [property: JsonPropertyName("payload")] JsonElement Payload,
    [property: JsonPropertyName("input_path")] string? InputPath,
    [property: JsonPropertyName("output_path")] string? OutputPath);

sealed class WritePayload
{
    [JsonPropertyName("sheets")]
    public List<SheetSpec> Sheets { get; set; } = [];

    [JsonPropertyName("cells")]
    public List<CellSpec> Cells { get; set; } = [];

    [JsonPropertyName("rich_text")]
    public List<RichTextSpec> RichText { get; set; } = [];

    [JsonPropertyName("comments")]
    public List<CommentSpec> Comments { get; set; } = [];

    [JsonPropertyName("tables")]
    public List<TableSpec> Tables { get; set; } = [];

    [JsonPropertyName("conditional_formats")]
    public List<ConditionalFormatSpec> ConditionalFormats { get; set; } = [];

    [JsonPropertyName("pivots")]
    public List<PivotSpec> Pivots { get; set; } = [];

    [JsonPropertyName("protection")]
    public List<ProtectionSpec> Protection { get; set; } = [];
}

sealed record SheetSpec([property: JsonPropertyName("name")] string Name);

sealed record CellSpec(
    [property: JsonPropertyName("sheet")] string Sheet,
    [property: JsonPropertyName("cell")] string Cell,
    [property: JsonPropertyName("type")] string? Type,
    [property: JsonPropertyName("value")] JsonElement Value,
    [property: JsonPropertyName("formula")] string? Formula);

sealed record RichTextSpec(
    [property: JsonPropertyName("sheet")] string Sheet,
    [property: JsonPropertyName("cell")] string Cell,
    [property: JsonPropertyName("runs")] List<RichRunSpec> Runs);

sealed record RichRunSpec(
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("bold")] bool Bold,
    [property: JsonPropertyName("italic")] bool Italic,
    [property: JsonPropertyName("font_color")] string? FontColor);

sealed record CommentSpec(
    [property: JsonPropertyName("sheet")] string Sheet,
    [property: JsonPropertyName("cell")] string Cell,
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("author")] string? Author);

sealed record TableSpec(
    [property: JsonPropertyName("sheet")] string Sheet,
    [property: JsonPropertyName("range")] string Range,
    [property: JsonPropertyName("name")] string Name);

sealed record ConditionalFormatSpec(
    [property: JsonPropertyName("sheet")] string Sheet,
    [property: JsonPropertyName("range")] string Range,
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("criteria")] string? Criteria,
    [property: JsonPropertyName("value")] string? Value);

sealed record PivotSpec(
    [property: JsonPropertyName("data_range")] string DataRange,
    [property: JsonPropertyName("cell")] string Cell,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("rows")] List<PivotField> Rows,
    [property: JsonPropertyName("columns")] List<PivotField> Columns,
    [property: JsonPropertyName("data")] List<PivotField> Data);

sealed record PivotField([property: JsonPropertyName("name")] string Name);

sealed record ProtectionSpec(
    [property: JsonPropertyName("sheet")] string Sheet,
    [property: JsonPropertyName("password")] string? Password);

static class JsonOptions
{
    public static readonly JsonSerializerOptions Value = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        WriteIndented = false
    };
}
