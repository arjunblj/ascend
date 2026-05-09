using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;

internal static class Program
{
record RunnerArgs(string Operation, string? File, int Rows, int Cols, string Workload, int Repeat, int Warmup, bool Json);
record Coord(int Row, int Col);

static void Main(string[] argv)
{
    var args = ParseArgs(argv);
    for (var index = 0; index < args.Warmup; index++)
    {
        if (args.Operation == "read") ReadAssertions(args.File!);
        else WriteWorkbookBytes(args.Workload, args.Rows, args.Cols);
    }
    var samples = new List<Dictionary<string, object>>();
    Dictionary<string, object?>? assertions = null;
    for (var index = 0; index < args.Repeat; index++)
    {
        var stopwatch = Stopwatch.StartNew();
        if (args.Operation == "read")
        {
            assertions = ReadAssertions(args.File!);
        }
        else
        {
            var bytes = WriteWorkbookBytes(args.Workload, args.Rows, args.Cols);
            stopwatch.Stop();
            assertions = WriteAssertions(bytes, args.Workload, args.Rows, args.Cols);
            samples.Add(MemorySample(stopwatch.Elapsed.TotalMilliseconds));
            continue;
        }
        stopwatch.Stop();
        samples.Add(MemorySample(stopwatch.Elapsed.TotalMilliseconds));
    }
    var payload = new Dictionary<string, object?>
    {
        ["assertions"] = assertions ?? new Dictionary<string, object?>(),
        ["samples"] = samples
    };
    var options = new JsonSerializerOptions { WriteIndented = !args.Json };
    Console.WriteLine(JsonSerializer.Serialize(payload, options));
}

static RunnerArgs ParseArgs(string[] argv)
{
    string? operation = null;
    string? file = null;
    var rows = 0;
    var cols = 0;
    var workload = "dense-values";
    var repeat = 1;
    var warmup = 0;
    var json = false;
    for (var index = 0; index < argv.Length; index++)
    {
        switch (argv[index])
        {
            case "--operation":
                operation = argv[++index];
                break;
            case "--file":
                file = argv[++index];
                break;
            case "--rows":
                rows = Math.Max(1, int.Parse(argv[++index]));
                break;
            case "--cols":
                cols = Math.Max(1, int.Parse(argv[++index]));
                break;
            case "--workload":
                workload = argv[++index];
                break;
            case "--repeat":
                repeat = Math.Max(1, int.Parse(argv[++index]));
                break;
            case "--warmup":
                warmup = Math.Max(0, int.Parse(argv[++index]));
                break;
            case "--json":
                json = true;
                break;
            default:
                throw new ArgumentException($"Unsupported argument {argv[index]}");
        }
    }
    if (operation != "read" && operation != "write") throw new ArgumentException("--operation must be read or write");
    if (operation == "read" && file is null) throw new ArgumentException("--file is required");
    if (operation == "write" && (rows <= 0 || cols <= 0))
        throw new ArgumentException("--rows and --cols are required for write");
    return new RunnerArgs(operation, file, rows, cols, workload, repeat, warmup, json);
}

static Dictionary<string, object?> ReadAssertions(string path)
{
    using var workbook = new XLWorkbook(path);
    var sheetNames = new List<string>();
    var usedRanges = new List<string>();
    var physicalUsedRanges = new List<string>();
    var semanticRefs = new List<string>();
    var semanticValues = new List<string>();
    var formulaTexts = new List<string>();
    var cellCount = 0;
    var physicalCellCount = 0;
    var formulaCount = 0;
    var commentCount = 0;
    var hyperlinkCount = 0;
    var dataValidationCount = 0;
    var conditionalFormatCount = 0;

    foreach (var worksheet in workbook.Worksheets)
    {
        sheetNames.Add(worksheet.Name);
        var semanticCoords = new List<Coord>();
        var physicalCoords = new List<Coord>();
        foreach (var cell in worksheet.CellsUsed(XLCellsUsedOptions.AllContents))
        {
            physicalCellCount++;
            var coord = new Coord(cell.Address.RowNumber, cell.Address.ColumnNumber);
            physicalCoords.Add(coord);
            if (cell.HasHyperlink) hyperlinkCount++;
            if (cell.HasComment) commentCount++;
            var payload = ScalarPayload(cell);
            if (payload is null) continue;
            cellCount++;
            semanticCoords.Add(coord);
            var reference = CellRef(worksheet.Name, coord.Row, coord.Col);
            semanticRefs.Add(reference);
            semanticValues.Add($"{reference}\t{payload}");
            if (cell.HasFormula)
            {
                formulaCount++;
                formulaTexts.Add($"{reference}={cell.FormulaA1}");
            }
        }
        dataValidationCount += worksheet.DataValidations.Count();
        conditionalFormatCount += worksheet.ConditionalFormats.Count();
        usedRanges.Add(UsedRange(worksheet.Name, semanticCoords));
        physicalUsedRanges.Add(UsedRange(worksheet.Name, physicalCoords));
    }

    return new Dictionary<string, object?>
    {
        ["runnerVersion"] = typeof(XLWorkbook).Assembly.GetName().Version?.ToString() ?? "unknown",
        ["sheetCount"] = workbook.Worksheets.Count,
        ["sheetNamesHash"] = HashLines(sheetNames.Select((name, index) => $"{index}:{name}")),
        ["cellCount"] = cellCount,
        ["physicalCellCount"] = physicalCellCount,
        ["formulaCount"] = formulaCount,
        ["usedRangeCount"] = usedRanges.Count,
        ["firstUsedRange"] = usedRanges.FirstOrDefault(),
        ["firstPhysicalUsedRange"] = physicalUsedRanges.FirstOrDefault(),
        ["usedRangesHash"] = HashLines(usedRanges),
        ["physicalUsedRangesHash"] = HashLines(physicalUsedRanges),
        ["semanticCellRefsHash"] = HashLines(semanticRefs),
        ["semanticCellValuesHash"] = HashLines(semanticValues),
        ["formulaTextHash"] = HashLines(formulaTexts),
        ["readCommentCount"] = commentCount,
        ["readHyperlinkCount"] = hyperlinkCount,
        ["readDataValidationCount"] = dataValidationCount,
        ["readConditionalFormatCount"] = conditionalFormatCount,
        ["readDefinedNameCount"] = workbook.DefinedNames.Count()
    };
}

static string? ScalarPayload(IXLCell cell)
{
    var value = cell.HasFormula ? cell.CachedValue : cell.Value;
    if (value.IsBlank) return null;
    if (value.IsBoolean) return $"b:{(value.GetBoolean() ? "true" : "false")}";
    if (value.IsNumber) return $"n:{CanonicalNumber(value.GetNumber())}";
    if (value.IsText) return $"s:{value.GetText()}";
    if (value.IsError) return $"e:{value.GetError()}";
    if (value.IsDateTime) return $"s:{value.GetDateTime():O}";
    if (value.IsTimeSpan) return $"s:{value.GetTimeSpan()}";
    return $"s:{value}";
}

static byte[] WriteWorkbookBytes(string workload, int rows, int cols)
{
    using var workbook = new XLWorkbook();
    var worksheet = workbook.Worksheets.Add("Data");
    for (var row = 0; row < rows; row++)
    {
        for (var col = 0; col < cols; col++)
        {
            var value = WorkloadValue(workload, row, col, cols);
            if (value is null) continue;
            var cell = worksheet.Cell(row + 1, col + 1);
            if (value is int intValue) cell.SetValue(intValue);
            else if (value is double doubleValue) cell.SetValue(doubleValue);
            else if (value is bool boolValue) cell.SetValue(boolValue);
            else cell.SetValue(value.ToString());
            if (workload == "styles-heavy")
            {
                switch ((row + col) % 4)
                {
                    case 1:
                        cell.Style.Fill.BackgroundColor = XLColor.LightYellow;
                        break;
                    case 2:
                        cell.Style.Fill.BackgroundColor = XLColor.LightGreen;
                        break;
                    case 3:
                        cell.Style.NumberFormat.Format = "#,##0.00";
                        break;
                }
            }
        }
    }
    using var output = new MemoryStream();
    workbook.SaveAs(output);
    return output.ToArray();
}

static Dictionary<string, object?> WriteAssertions(byte[] bytes, string workload, int rows, int cols)
{
    using var workbook = new XLWorkbook(new MemoryStream(bytes));
    var worksheet = workbook.Worksheet("Data");
    var semanticValues = new List<string>();
    var cellCount = 0;
    var formulaCount = 0;
    foreach (var cell in worksheet.CellsUsed(XLCellsUsedOptions.AllContents))
    {
        if (cell.HasFormula) formulaCount++;
        var payload = ScalarPayload(cell);
        if (payload is null) continue;
        cellCount++;
        semanticValues.Add($"{CellRef(worksheet.Name, cell.Address.RowNumber, cell.Address.ColumnNumber)}\t{payload}");
    }
    var observedHash = HashLines(semanticValues);
    var expectedHash = ExpectedValuesHash(workload, rows, cols);
    var expectedCells = ExpectedCellCount(workload, rows, cols);
    return new Dictionary<string, object?>
    {
        ["runnerVersion"] = typeof(XLWorkbook).Assembly.GetName().Version?.ToString() ?? "unknown",
        ["workload"] = workload,
        ["bytes"] = bytes.Length,
        ["reopenOk"] = true,
        ["tablePartCount"] = 0,
        ["commentPartCount"] = 0,
        ["vmlDrawingPartCount"] = 0,
        ["worksheetHyperlinkCount"] = 0,
        ["worksheetDataValidationCount"] = 0,
        ["worksheetConditionalFormattingCount"] = 0,
        ["definedNameCount"] = 0,
        ["sheetCount"] = workbook.Worksheets.Count,
        ["cellCount"] = cellCount,
        ["formulaCount"] = formulaCount,
        ["expectedCellCount"] = expectedCells,
        ["cellCountMatches"] = cellCount == expectedCells,
        ["semanticCellValuesHash"] = observedHash,
        ["expectedSemanticCellValuesHash"] = expectedHash,
        ["semanticCellValuesHashMatches"] = observedHash == expectedHash
    };
}

static object? WorkloadValue(string workload, int row, int col, int cols)
{
    if (workload == "dense-values") return row * cols + col;
    if (workload == "mixed-10pct-text")
    {
        var key = row * cols + col;
        return key % 10 == 0 ? $"text-{key:00000000}" : key;
    }
    if (workload == "mixed-50pct-text")
    {
        var key = row * cols + col;
        return key % 2 == 0 ? $"text-{key:00000000}" : key;
    }
    if (workload == "mixed-closedxml-10text-5number") return col < 10 ? "Hello world" : col - 10;
    if (workload == "plain-text") return $"text-{row * cols + col:00000000}";
    if (workload is "selected-sheet" or "metadata-only" or "warm-workflow") return row * cols + col;
    if (workload == "feature-rich") return row == 0 && col == 0 ? "Ascend" : row * cols + col;
    if (workload == "styles-heavy") return (row + 1) * (col + 1);
    if (workload == "formula-heavy")
    {
        var baseValue = row + 1;
        if (col == 0) return baseValue;
        if (col == 1) return baseValue * 2;
        return baseValue * 3 + col;
    }
    if (workload == "table-heavy")
    {
        if (row == 0) return $"Column {col + 1}";
        if (col % 3 == 0) return row;
        if (col % 3 == 1) return $"item-{row}-{col}";
        return row * cols + col;
    }
    if (workload == "sparse-wide")
    {
        if (col == 0) return row;
        if (col == cols - 1) return $"edge-{row}-{cols}";
        if ((row * 31 + col * 17) % 97 == 0) return row * cols + col;
        return null;
    }
    var fallbackKey = row * cols + col;
    return (col % 5) switch
    {
        0 => $"sku-{fallbackKey:00000000}",
        1 => $"region-{row % 17 + 1}",
        2 => $"customer-{row % 997}-segment-{col % 13}",
        3 => $"note row {row} col {col} token {fallbackKey % 104729}",
        _ => fallbackKey % 2 == 0 ? $"status-open-{fallbackKey % 31}" : $"status-closed-{fallbackKey % 29}"
    };
}

static string ExpectedValuesHash(string workload, int rows, int cols)
{
    var values = new List<string>();
    for (var row = 0; row < rows; row++)
    {
        for (var col = 0; col < cols; col++)
        {
            var value = WorkloadValue(workload, row, col, cols);
            if (value is null) continue;
            values.Add($"Data!{ColumnName(col + 1)}{row + 1}\t{ScalarPayload(value)}");
        }
    }
    return HashLines(values);
}

static int ExpectedCellCount(string workload, int rows, int cols)
{
    var count = 0;
    for (var row = 0; row < rows; row++)
    {
        for (var col = 0; col < cols; col++)
        {
            if (WorkloadValue(workload, row, col, cols) is not null) count++;
        }
    }
    return count;
}

static string ScalarPayload(object value)
{
    if (value is bool boolValue) return $"b:{(boolValue ? "true" : "false")}";
    if (value is int intValue) return $"n:{intValue}";
    if (value is long longValue) return $"n:{longValue}";
    if (value is double doubleValue) return $"n:{CanonicalNumber(doubleValue)}";
    if (value is decimal decimalValue) return $"n:{CanonicalNumber((double)decimalValue)}";
    return $"s:{value}";
}

static Dictionary<string, object> MemorySample(double durationMs)
{
    using var process = Process.GetCurrentProcess();
    var workingSet = process.WorkingSet64;
    return new Dictionary<string, object>
    {
        ["durationMs"] = durationMs,
        ["peakRssBytes"] = workingSet,
        ["rssAfterBytes"] = workingSet,
        ["heapUsedBytes"] = GC.GetTotalMemory(false)
    };
}

static string CanonicalNumber(double value)
{
    if (value == 0) return "0";
    return Math.Truncate(value) == value
        ? ((long)value).ToString(System.Globalization.CultureInfo.InvariantCulture)
        : value.ToString("G15", System.Globalization.CultureInfo.InvariantCulture);
}

static string CellRef(string sheetName, int row, int col) => $"{sheetName}!{ColumnName(col)}{row}";

static string UsedRange(string sheetName, List<Coord> coords)
{
    if (coords.Count == 0) return $"{sheetName}!empty";
    var minRow = coords.Min(coord => coord.Row);
    var minCol = coords.Min(coord => coord.Col);
    var maxRow = coords.Max(coord => coord.Row);
    var maxCol = coords.Max(coord => coord.Col);
    return $"{sheetName}!{ColumnName(minCol)}{minRow}:{ColumnName(maxCol)}{maxRow}";
}

static string ColumnName(int oneBasedColumn)
{
    var col = oneBasedColumn;
    var result = "";
    while (col > 0)
    {
        col--;
        result = (char)('A' + col % 26) + result;
        col /= 26;
    }
    return result;
}

static string HashLines(IEnumerable<string> lines)
{
    using var sha = SHA256.Create();
    foreach (var line in lines.Order(StringComparer.Ordinal))
    {
        var bytes = Encoding.UTF8.GetBytes(line);
        sha.TransformBlock(Encoding.UTF8.GetBytes($"{line.Length}:"), 0, $"{line.Length}:".Length, null, 0);
        sha.TransformBlock(bytes, 0, bytes.Length, null, 0);
        sha.TransformBlock(new byte[] { (byte)'\n' }, 0, 1, null, 0);
    }
    sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
    return Convert.ToHexString(sha.Hash!).ToLowerInvariant();
}
}
