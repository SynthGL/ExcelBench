using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Serialization;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

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
        error = "npoi_oracle_failed",
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
    using var workbook = new XSSFWorkbook();
    ConfigureSheets(workbook, payload.Sheets);
    ApplyCells(workbook, payload.Cells);
    ApplyRichText(workbook, payload.RichText);
    ApplyComments(workbook, payload.Comments);
    ApplyMergedRanges(workbook, payload.MergedRanges);
    ApplyProtection(workbook, payload.Protection);

    using (var stream = File.Create(request.OutputPath))
    {
        workbook.Write(stream, leaveOpen: false);
    }

    return new
    {
        fixture_id = request.FixtureId,
        operation = request.Operation,
        output_path = request.OutputPath,
        tool = "npoi",
        counts = new
        {
            sheets = workbook.NumberOfSheets,
            cells = payload.Cells.Count,
            formulas = payload.Cells.Count(cell => cell.Type == "formula"),
            rich_text = payload.RichText.Count,
            comments = payload.Comments.Count,
            merged_ranges = payload.MergedRanges.Count,
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
        tool = "npoi",
        counts = new
        {
            worksheets = CountParts(partNames, "worksheets/sheet"),
            shared_strings = CountParts(partNames, "sharedStrings"),
            comments = CountParts(partNames, "comments"),
            vml_drawings = CountParts(partNames, "drawings/vmlDrawing"),
            calc_chain = CountParts(partNames, "calcChain")
        }
    };
}

static void ConfigureSheets(XSSFWorkbook workbook, List<SheetSpec> sheets)
{
    if (sheets.Count == 0)
    {
        workbook.CreateSheet("Sheet1");
        return;
    }

    foreach (var sheet in sheets)
    {
        workbook.CreateSheet(sheet.Name);
    }
}

static void ApplyCells(XSSFWorkbook workbook, List<CellSpec> cells)
{
    foreach (var item in cells)
    {
        var cell = GetOrCreateCell(workbook, item.Sheet, item.Cell);
        if (item.Type == "formula")
        {
            cell.SetCellFormula(item.Formula ?? "");
            continue;
        }

        switch (item.Value.ValueKind)
        {
            case JsonValueKind.String:
                cell.SetCellValue(item.Value.GetString() ?? "");
                break;
            case JsonValueKind.Number:
                cell.SetCellValue(item.Value.GetDouble());
                break;
            case JsonValueKind.True:
            case JsonValueKind.False:
                cell.SetCellValue(item.Value.GetBoolean());
                break;
            case JsonValueKind.Null:
            case JsonValueKind.Undefined:
                break;
            default:
                cell.SetCellValue(item.Value.ToString());
                break;
        }
    }
}

static void ApplyRichText(XSSFWorkbook workbook, List<RichTextSpec> items)
{
    foreach (var item in items)
    {
        var cell = GetOrCreateCell(workbook, item.Sheet, item.Cell);
        var richText = new XSSFRichTextString();
        foreach (var run in item.Runs)
        {
            var font = (XSSFFont)workbook.CreateFont();
            font.IsBold = run.Bold;
            font.IsItalic = run.Italic;
            richText.Append(run.Text, font);
        }
        cell.SetCellValue(richText);
    }
}

static void ApplyComments(XSSFWorkbook workbook, List<CommentSpec> comments)
{
    foreach (var item in comments)
    {
        var sheet = workbook.GetSheet(item.Sheet)
            ?? throw new InvalidOperationException($"Sheet not found: {item.Sheet}");
        var cell = GetOrCreateCell(workbook, item.Sheet, item.Cell);
        var drawing = sheet.CreateDrawingPatriarch();
        var anchor = workbook.GetCreationHelper().CreateClientAnchor();
        anchor.Col1 = cell.ColumnIndex + 1;
        anchor.Col2 = cell.ColumnIndex + 4;
        anchor.Row1 = cell.RowIndex;
        anchor.Row2 = cell.RowIndex + 3;
        var comment = drawing.CreateCellComment(anchor);
        comment.String = workbook.GetCreationHelper().CreateRichTextString(item.Text);
        comment.Author = item.Author ?? "NPOI Oracle";
        cell.CellComment = comment;
    }
}

static void ApplyMergedRanges(XSSFWorkbook workbook, List<MergedRangeSpec> mergedRanges)
{
    foreach (var item in mergedRanges)
    {
        var sheet = workbook.GetSheet(item.Sheet)
            ?? throw new InvalidOperationException($"Sheet not found: {item.Sheet}");
        sheet.AddMergedRegion(CellRangeAddress.ValueOf(item.Range));
    }
}

static void ApplyProtection(XSSFWorkbook workbook, List<ProtectionSpec> protections)
{
    foreach (var item in protections)
    {
        var sheet = workbook.GetSheet(item.Sheet)
            ?? throw new InvalidOperationException($"Sheet not found: {item.Sheet}");
        sheet.ProtectSheet(item.Password ?? "");
    }
}

static ICell GetOrCreateCell(XSSFWorkbook workbook, string sheetName, string address)
{
    var sheet = workbook.GetSheet(sheetName)
        ?? throw new InvalidOperationException($"Sheet not found: {sheetName}");
    var cellRef = new CellReference(address);
    var row = sheet.GetRow(cellRef.Row) ?? sheet.CreateRow(cellRef.Row);
    return row.GetCell(cellRef.Col) ?? row.CreateCell(cellRef.Col);
}

static int CountParts(IEnumerable<string> partNames, string fragment)
{
    return partNames.Count(part => part.Contains(fragment, StringComparison.Ordinal));
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

    [JsonPropertyName("merged_ranges")]
    public List<MergedRangeSpec> MergedRanges { get; set; } = [];

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
    [property: JsonPropertyName("italic")] bool Italic);

sealed record CommentSpec(
    [property: JsonPropertyName("sheet")] string Sheet,
    [property: JsonPropertyName("cell")] string Cell,
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("author")] string? Author);

sealed record MergedRangeSpec(
    [property: JsonPropertyName("sheet")] string Sheet,
    [property: JsonPropertyName("range")] string Range);

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
