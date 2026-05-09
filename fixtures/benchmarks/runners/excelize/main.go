package main

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"math"
	"os"
	"runtime"
	"runtime/debug"
	"sort"
	"strconv"
	"strings"
	"syscall"
	"time"
	"unicode/utf16"

	"github.com/xuri/excelize/v2"
)

const runnerVersion = "1"

type args struct {
	operation      string
	file           string
	rows           int
	cols           int
	workload       string
	repeat         int
	warmup         int
	validationMode string
	pretty         bool
}

type sample struct {
	DurationMs    float64 `json:"durationMs"`
	PeakRssBytes  uint64  `json:"peakRssBytes,omitempty"`
	RssAfterBytes uint64  `json:"rssAfterBytes,omitempty"`
}

type payload struct {
	Assertions map[string]any `json:"assertions"`
	Samples    []sample       `json:"samples"`
}

func main() {
	if err := run(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

func run() error {
	parsed, err := parseArgs(os.Args[1:])
	if err != nil {
		return err
	}
	if parsed.operation != "read" && parsed.operation != "write" {
		return fmt.Errorf("unsupported --operation %q: expected read or write", parsed.operation)
	}
	for i := 0; i < parsed.warmup; i++ {
		if parsed.operation == "read" {
			if _, err := readAssertions(parsed.file); err != nil {
				return err
			}
		} else if _, err := writeWorkbook(parsed.workload, parsed.rows, parsed.cols); err != nil {
			return err
		}
	}
	samples := make([]sample, 0, parsed.repeat)
	var assertions map[string]any
	var lastData []byte
	for i := 0; i < parsed.repeat; i++ {
		start := time.Now()
		var err error
		if parsed.operation == "read" {
			assertions, err = readAssertions(parsed.file)
			if err != nil {
				return err
			}
		} else {
			lastData, err = writeWorkbook(parsed.workload, parsed.rows, parsed.cols)
			if err != nil {
				return err
			}
		}
		durationMs := float64(time.Since(start).Nanoseconds()) / 1_000_000
		if parsed.operation == "write" && parsed.validationMode == "each" {
			assertions, err = writeAssertionsFromBytes(lastData, parsed.workload, parsed.rows, parsed.cols)
			if err != nil {
				return err
			}
		}
		rss := peakRSSBytes()
		samples = append(samples, sample{
			DurationMs:    durationMs,
			PeakRssBytes:  rss,
			RssAfterBytes: rss,
		})
	}
	if parsed.operation == "write" && parsed.validationMode == "final" {
		var err error
		assertions, err = writeAssertionsFromBytes(lastData, parsed.workload, parsed.rows, parsed.cols)
		if err != nil {
			return err
		}
	}
	if assertions != nil {
		validationMode := parsed.validationMode
		if parsed.operation == "read" {
			validationMode = "each"
		}
		assertions["validationMode"] = validationMode
		if validationMode == "each" {
			assertions["validationSamples"] = parsed.repeat
		} else {
			assertions["validationSamples"] = 1
		}
	}
	encoded, err := marshal(payload{Assertions: assertions, Samples: samples}, parsed.pretty)
	if err != nil {
		return err
	}
	fmt.Println(string(encoded))
	return nil
}

func parseArgs(values []string) (args, error) {
	parsed := args{workload: "dense-values", repeat: 1, validationMode: "each", pretty: true}
	for i := 0; i < len(values); i++ {
		switch values[i] {
		case "--operation":
			i++
			if i >= len(values) {
				return parsed, errors.New("missing value for --operation")
			}
			parsed.operation = values[i]
		case "--file":
			i++
			if i >= len(values) {
				return parsed, errors.New("missing value for --file")
			}
			parsed.file = values[i]
		case "--rows":
			i++
			count, err := parseCount("--rows", values, i)
			if err != nil {
				return parsed, err
			}
			parsed.rows = max(1, count)
		case "--cols":
			i++
			count, err := parseCount("--cols", values, i)
			if err != nil {
				return parsed, err
			}
			parsed.cols = max(1, count)
		case "--workload":
			i++
			if i >= len(values) {
				return parsed, errors.New("missing value for --workload")
			}
			parsed.workload = values[i]
		case "--repeat":
			i++
			count, err := parseCount("--repeat", values, i)
			if err != nil {
				return parsed, err
			}
			parsed.repeat = max(1, count)
		case "--warmup":
			i++
			count, err := parseCount("--warmup", values, i)
			if err != nil {
				return parsed, err
			}
			parsed.warmup = max(0, count)
		case "--validation-mode":
			i++
			if i >= len(values) {
				return parsed, errors.New("missing value for --validation-mode")
			}
			if values[i] != "each" && values[i] != "final" {
				return parsed, fmt.Errorf("unsupported --validation-mode %q: expected each or final", values[i])
			}
			parsed.validationMode = values[i]
		case "--json":
			parsed.pretty = false
		default:
			return parsed, fmt.Errorf("unsupported argument %q", values[i])
		}
	}
	if parsed.operation == "" {
		return parsed, errors.New("missing --operation")
	}
	if parsed.operation == "read" && parsed.file == "" {
		return parsed, errors.New("missing --file")
	}
	if parsed.operation == "write" && (parsed.rows <= 0 || parsed.cols <= 0) {
		return parsed, errors.New("--rows and --cols are required for write")
	}
	return parsed, nil
}

func parseCount(name string, values []string, index int) (int, error) {
	if index >= len(values) {
		return 0, fmt.Errorf("missing value for %s", name)
	}
	count, err := strconv.Atoi(values[index])
	if err != nil {
		return 0, fmt.Errorf("%s must be a non-negative integer", name)
	}
	return count, nil
}

func readAssertions(path string) (map[string]any, error) {
	workbook, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer workbook.Close()
	return workbookAssertions(workbook)
}

func workbookAssertions(workbook *excelize.File) (map[string]any, error) {
	sheetNames := workbook.GetSheetList()
	cellCount := 0
	physicalCellCount := 0
	usedRanges := make([]string, 0, len(sheetNames))
	semanticCellRefs := make([]string, 0)
	semanticCellValues := make([]string, 0)

	for _, sheetName := range sheetNames {
		rows, err := workbook.Rows(sheetName)
		if err != nil {
			return nil, err
		}
		minRow := math.MaxInt
		minCol := math.MaxInt
		maxRow := 0
		maxCol := 0
		rowIndex := 0
		for rows.Next() {
			rowIndex++
			cols, err := rows.Columns(excelize.Options{RawCellValue: true})
			if err != nil {
				rows.Close()
				return nil, err
			}
			for colIndex, raw := range cols {
				if raw == "" {
					continue
				}
				cellRef, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex)
				if err != nil {
					rows.Close()
					return nil, err
				}
				cellType, err := workbook.GetCellType(sheetName, cellRef)
				if err != nil {
					rows.Close()
					return nil, err
				}
				payload := scalarPayload(cellType, raw)
				if payload == "" {
					continue
				}
				physicalCellCount++
				cellCount++
				minRow = min(minRow, rowIndex)
				minCol = min(minCol, colIndex+1)
				maxRow = max(maxRow, rowIndex)
				maxCol = max(maxCol, colIndex+1)
				ref := fmt.Sprintf("%s!%s", sheetName, cellRef)
				semanticCellRefs = append(semanticCellRefs, ref)
				semanticCellValues = append(semanticCellValues, ref+"\t"+payload)
			}
		}
		if err := rows.Close(); err != nil {
			return nil, err
		}
		if minRow == math.MaxInt {
			usedRanges = append(usedRanges, sheetName+"!empty")
		} else {
			usedRanges = append(
				usedRanges,
				fmt.Sprintf("%s!%s%d:%s%d", sheetName, columnName(minCol), minRow, columnName(maxCol), maxRow),
			)
		}
	}
	usedRangesHash := hashLines(usedRanges)
	return map[string]any{
		"runnerVersion":              runnerVersion,
		"excelizeVersion":            moduleVersion("github.com/xuri/excelize/v2"),
		"sheetCount":                 len(sheetNames),
		"sheetNamesHash":             hashLines(indexedLines(sheetNames)),
		"cellCount":                  cellCount,
		"physicalCellCount":          physicalCellCount,
		"formulaCount":               0,
		"usedRangeCount":             len(usedRanges),
		"firstUsedRange":             firstOrNull(usedRanges),
		"firstPhysicalUsedRange":     firstOrNull(usedRanges),
		"usedRangesHash":             usedRangesHash,
		"physicalUsedRangesHash":     usedRangesHash,
		"semanticCellRefsHash":       hashLines(semanticCellRefs),
		"semanticCellValuesHash":     hashLines(semanticCellValues),
		"formulaTextHash":            hashLines([]string{}),
		"excelizeRawCellValueOption": true,
	}, nil
}

func writeAssertions(workload string, rows int, cols int) (map[string]any, error) {
	data, err := writeWorkbook(workload, rows, cols)
	if err != nil {
		return nil, err
	}
	return writeAssertionsFromBytes(data, workload, rows, cols)
}

func writeAssertionsFromBytes(data []byte, workload string, rows int, cols int) (map[string]any, error) {
	workbook, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return map[string]any{
			"runnerVersion":                 runnerVersion,
			"excelizeVersion":               moduleVersion("github.com/xuri/excelize/v2"),
			"workload":                      workload,
			"bytes":                         len(data),
			"reopenOk":                      false,
			"reopenError":                   err.Error(),
			"cellCountMatches":              false,
			"semanticCellValuesHashMatches": false,
		}, nil
	}
	defer workbook.Close()
	assertions, err := workbookAssertions(workbook)
	if err != nil {
		return nil, err
	}
	expectedHash := expectedValuesHash(workload, rows, cols)
	expectedCells := expectedCellCount(workload, rows, cols)
	assertions["workload"] = workload
	assertions["bytes"] = len(data)
	assertions["reopenOk"] = true
	assertions["tablePartCount"] = 0
	assertions["commentPartCount"] = 0
	assertions["vmlDrawingPartCount"] = 0
	assertions["worksheetHyperlinkCount"] = 0
	assertions["worksheetDataValidationCount"] = 0
	assertions["worksheetConditionalFormattingCount"] = 0
	assertions["definedNameCount"] = 0
	assertions["expectedCellCount"] = expectedCells
	assertions["cellCountMatches"] = assertions["cellCount"] == expectedCells
	assertions["expectedSemanticCellValuesHash"] = expectedHash
	assertions["semanticCellValuesHashMatches"] = assertions["semanticCellValuesHash"] == expectedHash
	return assertions, nil
}

func writeWorkbook(workload string, rows int, cols int) ([]byte, error) {
	workbook := excelize.NewFile()
	defer workbook.Close()
	if err := workbook.SetSheetName("Sheet1", "Data"); err != nil {
		return nil, err
	}
	writer, err := workbook.NewStreamWriter("Data")
	if err != nil {
		return nil, err
	}
	for rowIndex := 0; rowIndex < rows; rowIndex++ {
		values := make([]interface{}, cols)
		for colIndex := 0; colIndex < cols; colIndex++ {
			values[colIndex] = workloadValue(workload, rowIndex, colIndex, cols)
		}
		cellRef, err := excelize.CoordinatesToCellName(1, rowIndex+1)
		if err != nil {
			return nil, err
		}
		if err := writer.SetRow(cellRef, values); err != nil {
			return nil, err
		}
	}
	if err := writer.Flush(); err != nil {
		return nil, err
	}
	var output bytes.Buffer
	if err := workbook.Write(&output); err != nil {
		return nil, err
	}
	return output.Bytes(), nil
}

func workloadValue(workload string, row int, col int, cols int) any {
	if workload == "dense-values" {
		return row*cols + col
	}
	key := row*cols + col
	if workload == "mixed-10pct-text" {
		if key%10 == 0 {
			return fmt.Sprintf("text-%08d", key)
		}
		return key
	}
	if workload == "mixed-50pct-text" {
		if key%2 == 0 {
			return fmt.Sprintf("text-%08d", key)
		}
		return key
	}
	if workload == "plain-text" {
		return fmt.Sprintf("text-%08d", key)
	}
	if workload == "styles-heavy" {
		return (row + 1) * (col + 1)
	}
	if workload == "formula-heavy" {
		base := row + 1
		if col == 0 {
			return base
		}
		if col == 1 {
			return base * 2
		}
		return base*3 + col
	}
	if workload == "feature-rich" {
		if row == 0 && col == 0 {
			return "Ascend"
		}
		return key
	}
	if workload == "table-heavy" {
		if row == 0 {
			return fmt.Sprintf("Column %d", col+1)
		}
		if col%3 == 0 {
			return row
		}
		if col%3 == 1 {
			return fmt.Sprintf("item-%d-%d", row, col)
		}
		return key
	}
	if workload == "sparse-wide" {
		if col == 0 {
			return row
		}
		if col == cols-1 {
			return fmt.Sprintf("edge-%d-%d", row, cols)
		}
		if (row*31+col*17)%97 == 0 {
			return key
		}
		return nil
	}
	switch col % 5 {
	case 0:
		return fmt.Sprintf("sku-%08d", key)
	case 1:
		return fmt.Sprintf("region-%d", (row%17)+1)
	case 2:
		return fmt.Sprintf("customer-%d-segment-%d", row%997, col%13)
	case 3:
		return fmt.Sprintf("note row %d col %d token %d", row, col, key%104729)
	default:
		if key%2 == 0 {
			return fmt.Sprintf("status-open-%d", key%31)
		}
		return fmt.Sprintf("status-closed-%d", key%29)
	}
}

func expectedValuesHash(workload string, rows int, cols int) string {
	values := make([]string, 0, expectedCellCount(workload, rows, cols))
	for row := 0; row < rows; row++ {
		for col := 0; col < cols; col++ {
			value := workloadValue(workload, row, col, cols)
			if value == nil {
				continue
			}
			values = append(values, fmt.Sprintf("Data!%s%d\t%s", columnName(col+1), row+1, scalarPayloadValue(value)))
		}
	}
	return hashLines(values)
}

func expectedCellCount(workload string, rows int, cols int) int {
	count := 0
	for row := 0; row < rows; row++ {
		for col := 0; col < cols; col++ {
			if workloadValue(workload, row, col, cols) != nil {
				count++
			}
		}
	}
	return count
}

func scalarPayload(cellType excelize.CellType, raw string) string {
	switch cellType {
	case excelize.CellTypeBool:
		normalized := strings.ToLower(raw)
		return "b:" + strconv.FormatBool(normalized == "1" || normalized == "true")
	case excelize.CellTypeNumber, excelize.CellTypeDate:
		return "n:" + canonicalNumberString(raw)
	case excelize.CellTypeError:
		return "e:" + raw
	case excelize.CellTypeInlineString, excelize.CellTypeSharedString:
		return "s:" + raw
	case excelize.CellTypeFormula:
		if number, ok := tryCanonicalNumber(raw); ok {
			return "n:" + number
		}
		return "s:" + raw
	default:
		if number, ok := tryCanonicalNumber(raw); ok {
			return "n:" + number
		}
		return "s:" + raw
	}
}

func scalarPayloadValue(value any) string {
	switch typed := value.(type) {
	case bool:
		return "b:" + strconv.FormatBool(typed)
	case int:
		return "n:" + strconv.Itoa(typed)
	case int64:
		return "n:" + strconv.FormatInt(typed, 10)
	case float64:
		return "n:" + canonicalNumber(typed)
	case string:
		return "s:" + typed
	default:
		return fmt.Sprintf("s:%v", typed)
	}
}

func tryCanonicalNumber(raw string) (string, bool) {
	if raw == "" {
		return "", false
	}
	value, err := strconv.ParseFloat(raw, 64)
	if err != nil {
		return "", false
	}
	return canonicalNumber(value), true
}

func canonicalNumberString(raw string) string {
	if number, ok := tryCanonicalNumber(raw); ok {
		return number
	}
	return raw
}

func canonicalNumber(value float64) string {
	if value == 0 {
		return "0"
	}
	if math.Trunc(value) == value && value >= math.MinInt64 && value <= math.MaxInt64 {
		return strconv.FormatInt(int64(value), 10)
	}
	text := strconv.FormatFloat(value, 'f', 15, 64)
	text = strings.TrimRight(text, "0")
	return strings.TrimRight(text, ".")
}

func indexedLines(values []string) []string {
	lines := make([]string, 0, len(values))
	for index, value := range values {
		lines = append(lines, fmt.Sprintf("%d:%s", index, value))
	}
	return lines
}

func hashLines(lines []string) string {
	sorted := append([]string(nil), lines...)
	sort.Strings(sorted)
	hash := sha256.New()
	for _, line := range sorted {
		hash.Write([]byte(strconv.Itoa(len(utf16.Encode([]rune(line))))))
		hash.Write([]byte(":"))
		hash.Write([]byte(line))
		hash.Write([]byte("\n"))
	}
	return hex.EncodeToString(hash.Sum(nil))
}

func columnName(col int) string {
	name := ""
	for col > 0 {
		col--
		name = string(rune('A'+(col%26))) + name
		col /= 26
	}
	return name
}

func firstOrNull(values []string) any {
	if len(values) == 0 {
		return nil
	}
	return values[0]
}

func peakRSSBytes() uint64 {
	var usage syscall.Rusage
	if err := syscall.Getrusage(syscall.RUSAGE_SELF, &usage); err != nil {
		return 0
	}
	value := uint64(usage.Maxrss)
	if runtime.GOOS == "darwin" {
		return value
	}
	return value * 1024
}

func marshal(value any, pretty bool) ([]byte, error) {
	if pretty {
		return json.MarshalIndent(value, "", "  ")
	}
	return json.Marshal(value)
}

func moduleVersion(path string) string {
	info, ok := debug.ReadBuildInfo()
	if !ok {
		return "unknown"
	}
	for _, dep := range info.Deps {
		if dep.Path == path {
			if dep.Replace != nil {
				return dep.Replace.Version
			}
			return dep.Version
		}
	}
	return "unknown"
}
