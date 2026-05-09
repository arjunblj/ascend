using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

internal static class Program
{
record RunnerArgs(string Operation, int Rows, int Cols, string Workload, int Repeat, int Warmup, bool Json);

static void Main(string[] argv)
{
    var args = ParseArgs(argv);
    if (args.Operation != "write") throw new ArgumentException("NPOI runner currently supports --operation write");
    for (var index = 0; index < args.Warmup; index++) WriteWorkbookBytes(args.Workload, args.Rows, args.Cols);
    var samples = new List<Dictionary<string, object>>();
    Dictionary<string, object?>? assertions = null;
    for (var index = 0; index < args.Repeat; index++)
    {
        var stopwatch = Stopwatch.StartNew();
        var bytes = WriteWorkbookBytes(args.Workload, args.Rows, args.Cols);
        stopwatch.Stop();
        assertions = WriteAssertions(bytes, args.Workload, args.Rows, args.Cols);
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
    if (operation != "write") throw new ArgumentException("--operation must be write");
    if (rows <= 0 || cols <= 0) throw new ArgumentException("--rows and --cols are required for write");
    return new RunnerArgs(operation, rows, cols, workload, repeat, warmup, json);
}

static byte[] WriteWorkbookBytes(string workload, int rows, int cols)
{
    using var workbook = new XSSFWorkbook();
    var worksheet = workbook.CreateSheet("Data");
    var styles = CreateStyles(workbook);
    for (var rowIndex = 0; rowIndex < rows; rowIndex++)
    {
        var row = worksheet.CreateRow(rowIndex);
        for (var colIndex = 0; colIndex < cols; colIndex++)
        {
            var value = WorkloadValue(workload, rowIndex, colIndex, cols);
            if (value is null) continue;
            var cell = row.CreateCell(colIndex);
            if (value is int intValue) cell.SetCellValue(intValue);
            else if (value is double doubleValue) cell.SetCellValue(doubleValue);
            else if (value is bool boolValue) cell.SetCellValue(boolValue);
            else cell.SetCellValue(value.ToString());
            if (workload == "styles-heavy") cell.CellStyle = styles[(rowIndex + colIndex) % styles.Length];
        }
    }
    using var output = new MemoryStream();
    workbook.Write(output, true);
    return output.ToArray();
}

static ICellStyle[] CreateStyles(IWorkbook workbook)
{
    var normal = workbook.CreateCellStyle();
    var yellow = workbook.CreateCellStyle();
    yellow.FillForegroundColor = IndexedColors.LightYellow.Index;
    yellow.FillPattern = FillPattern.SolidForeground;
    var green = workbook.CreateCellStyle();
    green.FillForegroundColor = IndexedColors.LightGreen.Index;
    green.FillPattern = FillPattern.SolidForeground;
    var number = workbook.CreateCellStyle();
    number.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.00");
    return new[] { normal, yellow, green, number };
}

static Dictionary<string, object?> WriteAssertions(byte[] bytes, string workload, int rows, int cols)
{
    using var workbook = WorkbookFactory.Create(new MemoryStream(bytes));
    var worksheet = workbook.GetSheet("Data");
    var semanticValues = new List<string>();
    var cellCount = 0;
    foreach (IRow row in worksheet)
    {
        foreach (ICell cell in row)
        {
            var payload = ScalarPayload(cell);
            if (payload is null) continue;
            cellCount++;
            semanticValues.Add($"{CellRef(worksheet.SheetName, cell.RowIndex + 1, cell.ColumnIndex + 1)}\t{payload}");
        }
    }
    var observedHash = HashLines(semanticValues);
    var expectedHash = ExpectedValuesHash(workload, rows, cols);
    var expectedCells = ExpectedCellCount(workload, rows, cols);
    return new Dictionary<string, object?>
    {
        ["runnerVersion"] = typeof(XSSFWorkbook).Assembly.GetName().Version?.ToString() ?? "unknown",
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
        ["sheetCount"] = workbook.NumberOfSheets,
        ["cellCount"] = cellCount,
        ["expectedCellCount"] = expectedCells,
        ["cellCountMatches"] = cellCount == expectedCells,
        ["semanticCellValuesHash"] = observedHash,
        ["expectedSemanticCellValuesHash"] = expectedHash,
        ["semanticCellValuesHashMatches"] = observedHash == expectedHash
    };
}

static string? ScalarPayload(ICell cell)
{
    var type = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
    return type switch
    {
        CellType.Blank => null,
        CellType.Numeric => $"n:{CanonicalNumber(cell.NumericCellValue)}",
        CellType.String => $"s:{cell.StringCellValue}",
        CellType.Boolean => $"b:{(cell.BooleanCellValue ? "true" : "false")}",
        CellType.Error => $"e:{cell.ErrorCellValue}",
        _ => null
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
        var prefix = Encoding.UTF8.GetBytes($"{line.Length}:");
        var bytes = Encoding.UTF8.GetBytes(line);
        sha.TransformBlock(prefix, 0, prefix.Length, null, 0);
        sha.TransformBlock(bytes, 0, bytes.Length, null, 0);
        sha.TransformBlock(new byte[] { (byte)'\n' }, 0, 1, null, 0);
    }
    sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
    return Convert.ToHexString(sha.Hash!).ToLowerInvariant();
}
}
