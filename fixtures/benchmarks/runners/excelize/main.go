package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"math"
	"os"
	pathpkg "path"
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
	operation       string
	file            string
	rows            int
	cols            int
	workload        string
	repeat          int
	warmup          int
	validationMode  string
	editSheet       string
	editRef         string
	editValueType   string
	editValue       any
	editValueSet    bool
	editOldValue    any
	editOldValueSet bool
	pretty          bool
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

type Coord struct {
	Row int
	Col int
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
	if parsed.operation != "read" && parsed.operation != "write" && parsed.operation != "edit-roundtrip" {
		return fmt.Errorf("unsupported --operation %q: expected read, write, or edit-roundtrip", parsed.operation)
	}
	for i := 0; i < parsed.warmup; i++ {
		if parsed.operation == "read" {
			if _, err := readAssertions(parsed.file); err != nil {
				return err
			}
		} else if parsed.operation == "edit-roundtrip" {
			data, err := editRoundtripBytes(parsed.file, parsed.editSheet, parsed.editRef, parsed.editValue)
			if err != nil {
				return err
			}
			if _, err := editRoundtripAssertions(parsed.file, data, parsed.editSheet, parsed.editRef, parsed.editValue, parsed.editOldValue, parsed.editOldValueSet, parsed.editValueType); err != nil {
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
		} else if parsed.operation == "edit-roundtrip" {
			lastData, err = editRoundtripBytes(parsed.file, parsed.editSheet, parsed.editRef, parsed.editValue)
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
		} else if parsed.operation == "edit-roundtrip" {
			assertions, err = editRoundtripAssertions(parsed.file, lastData, parsed.editSheet, parsed.editRef, parsed.editValue, parsed.editOldValue, parsed.editOldValueSet, parsed.editValueType)
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
		if parsed.operation == "read" || parsed.operation == "edit-roundtrip" {
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
	parsed := args{workload: "dense-values", repeat: 1, validationMode: "each", editValueType: "number", pretty: true}
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
		case "--edit-sheet":
			i++
			if i >= len(values) {
				return parsed, errors.New("missing value for --edit-sheet")
			}
			parsed.editSheet = values[i]
		case "--edit-ref":
			i++
			if i >= len(values) {
				return parsed, errors.New("missing value for --edit-ref")
			}
			parsed.editRef = values[i]
		case "--edit-value-type":
			i++
			if i >= len(values) {
				return parsed, errors.New("missing value for --edit-value-type")
			}
			if values[i] != "number" && values[i] != "string" && values[i] != "boolean" {
				return parsed, fmt.Errorf("unsupported --edit-value-type %q: expected number, string, or boolean", values[i])
			}
			parsed.editValueType = values[i]
		case "--edit-value":
			i++
			if i >= len(values) {
				return parsed, errors.New("missing value for --edit-value")
			}
			value, err := parseEditValue(values[i], parsed.editValueType)
			if err != nil {
				return parsed, fmt.Errorf("--edit-value %w", err)
			}
			parsed.editValue = value
			parsed.editValueSet = true
		case "--edit-old-value":
			i++
			if i >= len(values) {
				return parsed, errors.New("missing value for --edit-old-value")
			}
			value, err := parseEditValue(values[i], parsed.editValueType)
			if err != nil {
				return parsed, fmt.Errorf("--edit-old-value %w", err)
			}
			parsed.editOldValue = value
			parsed.editOldValueSet = true
		case "--json":
			parsed.pretty = false
		default:
			return parsed, fmt.Errorf("unsupported argument %q", values[i])
		}
	}
	if parsed.operation == "" {
		return parsed, errors.New("missing --operation")
	}
	if (parsed.operation == "read" || parsed.operation == "edit-roundtrip") && parsed.file == "" {
		return parsed, errors.New("missing --file")
	}
	if parsed.operation == "write" && (parsed.rows <= 0 || parsed.cols <= 0) {
		return parsed, errors.New("--rows and --cols are required for write")
	}
	if parsed.operation == "edit-roundtrip" && (parsed.editSheet == "" || parsed.editRef == "" || !parsed.editValueSet) {
		return parsed, errors.New("edit-roundtrip requires --edit-sheet, --edit-ref, and --edit-value")
	}
	return parsed, nil
}

func parseEditValue(raw string, valueType string) (any, error) {
	switch valueType {
	case "number":
		value, err := strconv.ParseFloat(raw, 64)
		if err != nil {
			return nil, errors.New("must be numeric")
		}
		return value, nil
	case "boolean":
		normalized := strings.ToLower(raw)
		if normalized == "1" || normalized == "true" {
			return true, nil
		}
		if normalized == "0" || normalized == "false" {
			return false, nil
		}
		return nil, errors.New("must be boolean")
	case "string":
		return raw, nil
	default:
		return nil, fmt.Errorf("has unsupported type %q", valueType)
	}
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

func editRoundtripBytes(path string, sheetName string, ref string, value any) ([]byte, error) {
	workbook, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer workbook.Close()
	if err := workbook.SetCellValue(sheetName, ref, value); err != nil {
		return nil, err
	}
	output, err := workbook.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return output.Bytes(), nil
}

func editRoundtripAssertions(path string, data []byte, sheetName string, ref string, expectedValue any, oldValue any, oldValueSet bool, valueType string) (map[string]any, error) {
	original, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	assertions, err := roundtripZipAssertions(data)
	if err != nil {
		return nil, err
	}
	observed := readZipCellValue(data, sheetName, ref)
	assertions["runnerVersion"] = runnerVersion
	assertions["excelizeVersion"] = moduleVersion("github.com/xuri/excelize/v2")
	assertions["bytes"] = len(data)
	assertions["byteIdentical"] = sha256.Sum256(data) == sha256.Sum256(original)
	assertions["editSheetName"] = sheetName
	assertions["editRef"] = ref
	assertions["editValueType"] = valueType
	if !oldValueSet {
		assertions["editOldValue"] = nil
	} else {
		assertions["editOldValue"] = oldValue
	}
	assertions["editExpectedValue"] = expectedValue
	assertions["editObservedValue"] = observed
	assertions["editCellValueMatches"] = editValuesEqual(observed, expectedValue)
	return assertions, nil
}

func roundtripZipAssertions(data []byte) (map[string]any, error) {
	assertions, err := workbookShapeAssertionsFromZip(data, "roundtrip")
	if err != nil {
		return nil, err
	}
	for key, value := range packageFingerprintAssertions(data) {
		assertions[key] = value
	}
	for key, value := range featureSummaryAssertions(data) {
		assertions[key] = value
	}
	return assertions, nil
}

func workbookAssertions(workbook *excelize.File) (map[string]any, error) {
	sheetNames := workbook.GetSheetList()
	cellCount := 0
	physicalCellCount := 0
	readCommentCount := 0
	readHyperlinkCount := 0
	readDataValidationCount := 0
	readConditionalFormatCount := 0
	usedRanges := make([]string, 0, len(sheetNames))
	semanticCellRefsHash := newOrderedLineHasher()
	semanticCellValuesHash := newOrderedLineHasher()

	for _, sheetName := range sheetNames {
		comments, err := workbook.GetComments(sheetName)
		if err != nil {
			return nil, err
		}
		readCommentCount += len(comments)
		dataValidations, err := workbook.GetDataValidations(sheetName)
		if err != nil {
			return nil, err
		}
		readDataValidationCount += len(dataValidations)
		conditionalFormats, err := workbook.GetConditionalFormats(sheetName)
		if err != nil {
			return nil, err
		}
		readConditionalFormatCount += len(conditionalFormats)
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
				hasHyperlink, _, err := workbook.GetCellHyperLink(sheetName, cellRef)
				if err != nil {
					rows.Close()
					return nil, err
				}
				if hasHyperlink {
					readHyperlinkCount++
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
				semanticCellRefsHash.Add(ref)
				semanticCellValuesHash.Add(ref + "\t" + payload)
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
		"runnerVersion":                 runnerVersion,
		"excelizeVersion":               moduleVersion("github.com/xuri/excelize/v2"),
		"sheetCount":                    len(sheetNames),
		"sheetNamesHash":                hashLines(indexedLines(sheetNames)),
		"cellCount":                     cellCount,
		"physicalCellCount":             physicalCellCount,
		"formulaCount":                  0,
		"usedRangeCount":                len(usedRanges),
		"firstUsedRange":                firstOrNull(usedRanges),
		"firstPhysicalUsedRange":        firstOrNull(usedRanges),
		"usedRangesHash":                usedRangesHash,
		"physicalUsedRangesHash":        usedRangesHash,
		"semanticCellRefsHash":          semanticCellRefsHash.Digest(),
		"orderedSemanticCellRefsHash":   semanticCellRefsHash.Digest(),
		"semanticCellValuesHash":        semanticCellValuesHash.Digest(),
		"orderedSemanticCellValuesHash": semanticCellValuesHash.Digest(),
		"formulaTextHash":               hashLines([]string{}),
		"excelizeRawCellValueOption":    true,
		"readCommentCount":              readCommentCount,
		"readHyperlinkCount":            readHyperlinkCount,
		"readDataValidationCount":       readDataValidationCount,
		"readConditionalFormatCount":    readConditionalFormatCount,
		"readDefinedNameCount":          len(workbook.GetDefinedName()),
	}, nil
}

func generatedWorkbookAssertions(workbook *excelize.File, workload string, expectedRows int, expectedCols int) (map[string]any, error) {
	sheetNames := workbook.GetSheetList()
	cellCount := 0
	physicalCellCount := 0
	usedRanges := make([]string, 0, len(sheetNames))
	semanticCellRefsHash := newOrderedLineHasher()
	semanticCellValuesHash := newOrderedLineHasher()
	columnNames := make([]string, expectedCols)
	for col := 0; col < expectedCols; col++ {
		columnNames[col] = columnName(col + 1)
	}

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
				colNumber := colIndex + 1
				colName := columnName(colNumber)
				if colIndex < len(columnNames) {
					colName = columnNames[colIndex]
				}
				expected := workloadValue(workload, rowIndex-1, colIndex, expectedCols)
				payload := scalarPayloadForExpectedValue(expected, raw)
				if payload == "" {
					continue
				}
				physicalCellCount++
				cellCount++
				minRow = min(minRow, rowIndex)
				minCol = min(minCol, colNumber)
				maxRow = max(maxRow, rowIndex)
				maxCol = max(maxCol, colNumber)
				ref := fmt.Sprintf("%s!%s%d", sheetName, colName, rowIndex)
				semanticCellRefsHash.Add(ref)
				semanticCellValuesHash.Add(ref + "\t" + payload)
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
		"runnerVersion":                  runnerVersion,
		"excelizeVersion":                moduleVersion("github.com/xuri/excelize/v2"),
		"sheetCount":                     len(sheetNames),
		"sheetNamesHash":                 hashLines(indexedLines(sheetNames)),
		"cellCount":                      cellCount,
		"physicalCellCount":              physicalCellCount,
		"formulaCount":                   0,
		"usedRangeCount":                 len(usedRanges),
		"firstUsedRange":                 firstOrNull(usedRanges),
		"firstPhysicalUsedRange":         firstOrNull(usedRanges),
		"usedRangesHash":                 usedRangesHash,
		"physicalUsedRangesHash":         usedRangesHash,
		"semanticCellRefsHash":           semanticCellRefsHash.Digest(),
		"orderedSemanticCellRefsHash":    semanticCellRefsHash.Digest(),
		"semanticCellValuesHash":         semanticCellValuesHash.Digest(),
		"orderedSemanticCellValuesHash":  semanticCellValuesHash.Digest(),
		"formulaTextHash":                hashLines([]string{}),
		"excelizeRawCellValueOption":     true,
		"generatedWriteValueValidation":  true,
		"generatedWriteExpectedRowCount": expectedRows,
		"generatedWriteExpectedColCount": expectedCols,
		"generatedWriteTypeValidation":   "workload-shape",
	}, nil
}

type generatedSheetSummary struct {
	cellCount                     int
	physicalCellCount             int
	formulaCount                  int
	firstUsedRange                any
	usedRangesHash                string
	semanticCellRefsHash          string
	orderedSemanticCellValuesHash string
}

func generatedZipAssertions(data []byte, workload string, expectedRows int, expectedCols int) (map[string]any, error) {
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
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
	sheetNames, err := workbookSheetNames(reader)
	if err != nil {
		return nil, err
	}
	sharedStrings, err := workbookSharedStrings(reader)
	if err != nil {
		return nil, err
	}
	summary, err := generatedSheetSummaryFromZip(reader, sharedStrings, workload, expectedRows, expectedCols)
	if err != nil {
		return nil, err
	}
	return map[string]any{
		"runnerVersion":                  runnerVersion,
		"excelizeVersion":                moduleVersion("github.com/xuri/excelize/v2"),
		"sheetCount":                     len(sheetNames),
		"sheetNamesHash":                 hashLines(indexedLines(sheetNames)),
		"cellCount":                      summary.cellCount,
		"physicalCellCount":              summary.physicalCellCount,
		"formulaCount":                   summary.formulaCount,
		"usedRangeCount":                 1,
		"firstUsedRange":                 summary.firstUsedRange,
		"firstPhysicalUsedRange":         summary.firstUsedRange,
		"usedRangesHash":                 summary.usedRangesHash,
		"physicalUsedRangesHash":         summary.usedRangesHash,
		"semanticCellRefsHash":           summary.semanticCellRefsHash,
		"orderedSemanticCellRefsHash":    summary.semanticCellRefsHash,
		"semanticCellValuesHash":         summary.orderedSemanticCellValuesHash,
		"orderedSemanticCellValuesHash":  summary.orderedSemanticCellValuesHash,
		"formulaTextHash":                hashLines([]string{}),
		"excelizeRawCellValueOption":     true,
		"generatedWriteZipValidation":    true,
		"generatedWriteExpectedRowCount": expectedRows,
		"generatedWriteExpectedColCount": expectedCols,
		"generatedWriteTypeValidation":   "workload-shape",
		"generatedWriteValueValidation":  true,
	}, nil
}

func workbookSheetNames(reader *zip.Reader) ([]string, error) {
	body, err := zipFileBytes(reader, "xl/workbook.xml")
	if err != nil {
		return nil, err
	}
	decoder := xml.NewDecoder(bytes.NewReader(body))
	names := []string{}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return names, nil
		}
		if err != nil {
			return nil, err
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "sheet" {
			continue
		}
		if name := xmlAttr(start.Attr, "name"); name != "" {
			names = append(names, name)
		}
	}
}

func workbookSharedStrings(reader *zip.Reader) ([]string, error) {
	file := zipFile(reader, "xl/sharedStrings.xml")
	if file == nil {
		return nil, nil
	}
	stream, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer stream.Close()
	decoder := xml.NewDecoder(stream)
	values := []string{}
	var current strings.Builder
	inString := false
	inText := false
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return values, nil
		}
		if err != nil {
			return nil, err
		}
		switch typed := token.(type) {
		case xml.StartElement:
			if typed.Name.Local == "si" {
				current.Reset()
				inString = true
			} else if inString && typed.Name.Local == "t" {
				inText = true
			}
		case xml.CharData:
			if inText {
				current.Write([]byte(typed))
			}
		case xml.EndElement:
			if typed.Name.Local == "t" {
				inText = false
			} else if typed.Name.Local == "si" {
				values = append(values, current.String())
				inString = false
			}
		}
	}
}

func generatedSheetSummaryFromZip(reader *zip.Reader, sharedStrings []string, workload string, expectedRows int, expectedCols int) (generatedSheetSummary, error) {
	file := zipFile(reader, "xl/worksheets/sheet1.xml")
	if file == nil {
		return generatedSheetSummary{}, errors.New("missing xl/worksheets/sheet1.xml")
	}
	stream, err := file.Open()
	if err != nil {
		return generatedSheetSummary{}, err
	}
	defer stream.Close()
	decoder := xml.NewDecoder(stream)
	columnNames := columnNameCache(expectedCols)
	semanticCellRefsHash := newOrderedLineHasher()
	semanticCellValuesHash := newOrderedLineHasher()
	cellCount := 0
	formulaCount := 0
	minRow := math.MaxInt
	minCol := math.MaxInt
	maxRow := 0
	maxCol := 0
	rowIndex := 0
	implicitRow := 0
	implicitCol := 0
	inCell := false
	inValue := false
	inInlineText := false
	cellType := ""
	cellCol := 0
	var cellValue strings.Builder
	var inlineValue strings.Builder
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return generatedSheetSummary{}, err
		}
		switch typed := token.(type) {
		case xml.StartElement:
			switch typed.Name.Local {
			case "row":
				rowIndex = positiveIntAttr(typed.Attr, "r")
				if rowIndex == 0 {
					rowIndex = implicitRow + 1
				}
				implicitRow = rowIndex
				implicitCol = 0
			case "c":
				inCell = true
				cellType = xmlAttr(typed.Attr, "t")
				cellValue.Reset()
				inlineValue.Reset()
				cellCol = columnIndexFromCellRef(xmlAttr(typed.Attr, "r"))
				if cellCol == 0 {
					cellCol = implicitCol + 1
				}
				implicitCol = cellCol
			case "f":
				if inCell {
					formulaCount++
				}
			case "v":
				if inCell {
					inValue = true
				}
			case "t":
				if inCell {
					inInlineText = true
				}
			}
		case xml.CharData:
			if inValue {
				cellValue.Write([]byte(typed))
			}
			if inInlineText {
				inlineValue.Write([]byte(typed))
			}
		case xml.EndElement:
			switch typed.Name.Local {
			case "v":
				inValue = false
			case "t":
				inInlineText = false
			case "c":
				payload := generatedCellPayload(cellType, cellValue.String(), inlineValue.String(), sharedStrings)
				if payload != "" {
					cellCount++
					minRow = min(minRow, rowIndex)
					minCol = min(minCol, cellCol)
					maxRow = max(maxRow, rowIndex)
					maxCol = max(maxCol, cellCol)
					colName := columnName(cellCol)
					if cellCol > 0 && cellCol <= len(columnNames) {
						colName = columnNames[cellCol-1]
					}
					ref := fmt.Sprintf("Data!%s%d", colName, rowIndex)
					semanticCellRefsHash.Add(ref)
					semanticCellValuesHash.Add(ref + "\t" + payload)
				}
				inCell = false
			}
		}
	}
	usedRanges := []string{"Data!empty"}
	if minRow != math.MaxInt {
		usedRanges = []string{
			fmt.Sprintf("Data!%s%d:%s%d", columnName(minCol), minRow, columnName(maxCol), maxRow),
		}
	}
	return generatedSheetSummary{
		cellCount:                     cellCount,
		physicalCellCount:             cellCount,
		formulaCount:                  formulaCount,
		firstUsedRange:                firstOrNull(usedRanges),
		usedRangesHash:                hashLines(usedRanges),
		semanticCellRefsHash:          semanticCellRefsHash.Digest(),
		orderedSemanticCellValuesHash: semanticCellValuesHash.Digest(),
	}, nil
}

type workbookSheet struct {
	name string
	path string
}

type relationship struct {
	id         string
	relType    string
	target     string
	targetMode string
}

type sheetZipSummary struct {
	cellCount         int
	physicalCellCount int
	formulaCount      int
	usedRange         string
	physicalUsedRange string
	semanticRefs      []string
	semanticValues    []string
	formulaTexts      []string
}

func workbookShapeAssertionsFromZip(data []byte, prefix string) (map[string]any, error) {
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	sheets, err := workbookSheets(reader)
	if err != nil {
		return nil, err
	}
	sharedStrings, err := workbookSharedStrings(reader)
	if err != nil {
		return nil, err
	}
	sheetNames := make([]string, 0, len(sheets))
	usedRanges := make([]string, 0, len(sheets))
	physicalUsedRanges := make([]string, 0, len(sheets))
	semanticRefs := []string{}
	semanticValues := []string{}
	formulaTexts := []string{}
	cellCount := 0
	physicalCellCount := 0
	formulaCount := 0
	for _, sheet := range sheets {
		sheetNames = append(sheetNames, sheet.name)
		summary, err := sheetSummaryFromZip(reader, sharedStrings, sheet.name, sheet.path)
		if err != nil {
			return nil, err
		}
		cellCount += summary.cellCount
		physicalCellCount += summary.physicalCellCount
		formulaCount += summary.formulaCount
		usedRanges = append(usedRanges, summary.usedRange)
		physicalUsedRanges = append(physicalUsedRanges, summary.physicalUsedRange)
		semanticRefs = append(semanticRefs, summary.semanticRefs...)
		semanticValues = append(semanticValues, summary.semanticValues...)
		formulaTexts = append(formulaTexts, summary.formulaTexts...)
	}
	return map[string]any{
		prefixKey(prefix, "sheetCount"):             len(sheetNames),
		prefixKey(prefix, "sheetNamesHash"):         hashLines(indexedLines(sheetNames)),
		prefixKey(prefix, "cellCount"):              cellCount,
		prefixKey(prefix, "physicalCellCount"):      physicalCellCount,
		prefixKey(prefix, "formulaCount"):           formulaCount,
		prefixKey(prefix, "usedRangeCount"):         len(usedRanges),
		prefixKey(prefix, "firstUsedRange"):         firstOrNull(usedRanges),
		prefixKey(prefix, "firstPhysicalUsedRange"): firstOrNull(physicalUsedRanges),
		prefixKey(prefix, "usedRangesHash"):         hashLines(usedRanges),
		prefixKey(prefix, "physicalUsedRangesHash"): hashLines(physicalUsedRanges),
		prefixKey(prefix, "semanticCellRefsHash"):   hashLines(semanticRefs),
		prefixKey(prefix, "semanticCellValuesHash"): hashLines(semanticValues),
		prefixKey(prefix, "formulaTextHash"):        hashLines(formulaTexts),
	}, nil
}

func sheetSummaryFromZip(reader *zip.Reader, sharedStrings []string, sheetName string, sheetPath string) (sheetZipSummary, error) {
	file := zipFile(reader, sheetPath)
	if file == nil {
		return sheetZipSummary{}, fmt.Errorf("missing %s", sheetPath)
	}
	stream, err := file.Open()
	if err != nil {
		return sheetZipSummary{}, err
	}
	defer stream.Close()
	decoder := xml.NewDecoder(stream)
	semanticRefs := []string{}
	semanticValues := []string{}
	formulaTexts := []string{}
	semanticCoords := []Coord{}
	physicalCoords := []Coord{}
	rowIndex := 0
	implicitRow := 0
	implicitCol := 0
	inCell := false
	inValue := false
	inInlineText := false
	inFormula := false
	cellType := ""
	cellRow := 0
	cellCol := 0
	cellRef := ""
	hasFormula := false
	var cellValue strings.Builder
	var inlineValue strings.Builder
	var formulaValue strings.Builder
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return sheetZipSummary{}, err
		}
		switch typed := token.(type) {
		case xml.StartElement:
			switch typed.Name.Local {
			case "row":
				rowIndex = positiveIntAttr(typed.Attr, "r")
				if rowIndex == 0 {
					rowIndex = implicitRow + 1
				}
				implicitRow = rowIndex
				implicitCol = 0
			case "c":
				inCell = true
				cellType = xmlAttr(typed.Attr, "t")
				cellValue.Reset()
				inlineValue.Reset()
				formulaValue.Reset()
				hasFormula = false
				cellRef = xmlAttr(typed.Attr, "r")
				cellCol = columnIndexFromCellRef(cellRef)
				cellRow = rowIndexFromCellRef(cellRef)
				if cellCol == 0 {
					cellCol = implicitCol + 1
				}
				if cellRow == 0 {
					cellRow = rowIndex
				}
				implicitCol = cellCol
				if cellRow > 0 && cellCol > 0 {
					physicalCoords = append(physicalCoords, Coord{Row: cellRow, Col: cellCol})
				}
			case "f":
				if inCell {
					hasFormula = true
					inFormula = true
				}
			case "v":
				if inCell {
					inValue = true
				}
			case "t":
				if inCell {
					inInlineText = true
				}
			}
		case xml.CharData:
			if inValue {
				cellValue.Write([]byte(typed))
			}
			if inInlineText {
				inlineValue.Write([]byte(typed))
			}
			if inFormula {
				formulaValue.Write([]byte(typed))
			}
		case xml.EndElement:
			switch typed.Name.Local {
			case "f":
				inFormula = false
			case "v":
				inValue = false
			case "t":
				inInlineText = false
			case "c":
				ref := cellRef
				if ref == "" && cellRow > 0 && cellCol > 0 {
					ref = fmt.Sprintf("%s%d", columnName(cellCol), cellRow)
				}
				if hasFormula {
					formulaTexts = append(formulaTexts, fmt.Sprintf("%s!%s=%s", sheetName, ref, formulaValue.String()))
				}
				payload := zipCellPayload(cellType, cellValue.String(), inlineValue.String(), sharedStrings, hasFormula)
				if payload != "" && cellRow > 0 && cellCol > 0 {
					semanticCoords = append(semanticCoords, Coord{Row: cellRow, Col: cellCol})
					fullRef := fmt.Sprintf("%s!%s", sheetName, ref)
					semanticRefs = append(semanticRefs, fullRef)
					semanticValues = append(semanticValues, fullRef+"\t"+payload)
				}
				inCell = false
			}
		}
	}
	return sheetZipSummary{
		cellCount:         len(semanticCoords),
		physicalCellCount: len(physicalCoords),
		formulaCount:      len(formulaTexts),
		usedRange:         UsedRange(sheetName, semanticCoords),
		physicalUsedRange: UsedRange(sheetName, physicalCoords),
		semanticRefs:      semanticRefs,
		semanticValues:    semanticValues,
		formulaTexts:      formulaTexts,
	}, nil
}

func zipCellPayload(cellType string, raw string, inline string, sharedStrings []string, hasFormula bool) string {
	if hasFormula {
		return "empty"
	}
	return generatedCellPayload(cellType, raw, inline, sharedStrings)
}

func readZipCellValue(data []byte, sheetName string, ref string) any {
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil
	}
	sheets, err := workbookSheets(reader)
	if err != nil {
		return nil
	}
	sharedStrings, err := workbookSharedStrings(reader)
	if err != nil {
		return nil
	}
	for _, sheet := range sheets {
		if sheet.name != sheetName {
			continue
		}
		file := zipFile(reader, sheet.path)
		if file == nil {
			return nil
		}
		stream, err := file.Open()
		if err != nil {
			return nil
		}
		defer stream.Close()
		decoder := xml.NewDecoder(stream)
		inTarget := false
		inValue := false
		inInlineText := false
		cellType := ""
		var value strings.Builder
		var inline strings.Builder
		for {
			token, err := decoder.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				return nil
			}
			switch typed := token.(type) {
			case xml.StartElement:
				if typed.Name.Local == "c" {
					inTarget = xmlAttr(typed.Attr, "r") == ref
					cellType = xmlAttr(typed.Attr, "t")
					value.Reset()
					inline.Reset()
				} else if inTarget && typed.Name.Local == "v" {
					inValue = true
				} else if inTarget && typed.Name.Local == "t" {
					inInlineText = true
				}
			case xml.CharData:
				if inValue {
					value.Write([]byte(typed))
				}
				if inInlineText {
					inline.Write([]byte(typed))
				}
			case xml.EndElement:
				if typed.Name.Local == "v" {
					inValue = false
				} else if typed.Name.Local == "t" {
					inInlineText = false
				} else if typed.Name.Local == "c" && inTarget {
					return zipCellScalarValue(cellType, value.String(), inline.String(), sharedStrings)
				}
			}
		}
	}
	return nil
}

func zipCellScalarValue(cellType string, raw string, inline string, sharedStrings []string) any {
	switch cellType {
	case "b":
		normalized := strings.ToLower(raw)
		return normalized == "1" || normalized == "true"
	case "inlineStr":
		return inline
	case "s":
		index, err := strconv.Atoi(raw)
		if err != nil || index < 0 || index >= len(sharedStrings) {
			return ""
		}
		return sharedStrings[index]
	case "str":
		return raw
	default:
		if raw == "" {
			return nil
		}
		value, err := strconv.ParseFloat(raw, 64)
		if err != nil {
			return raw
		}
		return value
	}
}

func editValuesEqual(observed any, expected any) bool {
	observedNumber, observedIsNumber := observed.(float64)
	expectedNumber, expectedIsNumber := expected.(float64)
	if observedIsNumber && expectedIsNumber {
		return observedNumber == expectedNumber
	}
	return observed == expected
}

func writeAssertions(workload string, rows int, cols int) (map[string]any, error) {
	data, err := writeWorkbook(workload, rows, cols)
	if err != nil {
		return nil, err
	}
	return writeAssertionsFromBytes(data, workload, rows, cols)
}

func writeAssertionsFromBytes(data []byte, workload string, rows int, cols int) (map[string]any, error) {
	if canUseGeneratedWorkbookAssertions(workload) {
		assertions, err := generatedZipAssertions(data, workload, rows, cols)
		if err != nil {
			return nil, err
		}
		return finalizeWriteAssertions(assertions, data, workload, rows, cols), nil
	}
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
	var assertions map[string]any
	if canUseGeneratedWorkbookAssertions(workload) {
		assertions, err = generatedWorkbookAssertions(workbook, workload, rows, cols)
	} else {
		assertions, err = workbookAssertions(workbook)
	}
	if err != nil {
		return nil, err
	}
	return finalizeWriteAssertions(assertions, data, workload, rows, cols), nil
}

func finalizeWriteAssertions(assertions map[string]any, data []byte, workload string, rows int, cols int) map[string]any {
	shouldComputeExpectedHash := rows*cols <= 500_000 || !canUseGeneratedWorkbookAssertions(workload)
	expectedCells := generatedCellCount(workload, rows, cols)
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
	if shouldComputeExpectedHash {
		expectedOrderedHash := expectedOrderedValuesHash(workload, rows, cols)
		assertions["expectedOrderedSemanticCellValuesHash"] = expectedOrderedHash
		assertions["semanticCellValuesHashMatches"] = assertions["orderedSemanticCellValuesHash"] == expectedOrderedHash
	}
	return assertions
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
	output, err := os.CreateTemp("", "ascend-excelize-*.xlsx")
	if err != nil {
		return nil, err
	}
	outputPath := output.Name()
	if err := output.Close(); err != nil {
		_ = os.Remove(outputPath)
		return nil, err
	}
	defer os.Remove(outputPath)
	if err := workbook.SaveAs(outputPath); err != nil {
		return nil, err
	}
	return os.ReadFile(outputPath)
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
	if workload == "mixed-closedxml-10text-5number" {
		if col < 10 {
			return "Hello world"
		}
		return col - 10
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

func expectedOrderedValuesHash(workload string, rows int, cols int) string {
	hash := newOrderedLineHasher()
	for row := 0; row < rows; row++ {
		for col := 0; col < cols; col++ {
			value := workloadValue(workload, row, col, cols)
			if value == nil {
				continue
			}
			hash.Add(fmt.Sprintf("Data!%s%d\t%s", columnName(col+1), row+1, scalarPayloadValue(value)))
		}
	}
	return hash.Digest()
}

func canUseGeneratedWorkbookAssertions(workload string) bool {
	switch workload {
	case "dense-values",
		"mixed-10pct-text",
		"mixed-50pct-text",
		"mixed-closedxml-10text-5number",
		"plain-text",
		"string-heavy",
		"sparse-wide":
		return true
	default:
		return false
	}
}

func generatedCellCount(workload string, rows int, cols int) int {
	if workload != "sparse-wide" {
		return rows * cols
	}
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

func scalarPayloadForExpectedValue(expected any, raw string) string {
	switch expected.(type) {
	case bool:
		normalized := strings.ToLower(raw)
		return "b:" + strconv.FormatBool(normalized == "1" || normalized == "true")
	case int, int64, float64:
		return "n:" + canonicalNumberString(raw)
	case string:
		return "s:" + raw
	case nil:
		if number, ok := tryCanonicalNumber(raw); ok {
			return "n:" + number
		}
		return "s:" + raw
	default:
		return "s:" + raw
	}
}

func generatedCellPayload(cellType string, raw string, inline string, sharedStrings []string) string {
	switch cellType {
	case "b":
		normalized := strings.ToLower(raw)
		return "b:" + strconv.FormatBool(normalized == "1" || normalized == "true")
	case "inlineStr":
		return "s:" + inline
	case "s":
		index, err := strconv.Atoi(raw)
		if err != nil || index < 0 || index >= len(sharedStrings) {
			return ""
		}
		return "s:" + sharedStrings[index]
	case "str":
		return "s:" + raw
	default:
		if raw == "" {
			return ""
		}
		return "n:" + canonicalNumberString(raw)
	}
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

func zipFile(reader *zip.Reader, name string) *zip.File {
	for _, file := range reader.File {
		if file.Name == name {
			return file
		}
	}
	return nil
}

func zipFileBytes(reader *zip.Reader, name string) ([]byte, error) {
	file := zipFile(reader, name)
	if file == nil {
		return nil, fmt.Errorf("missing %s", name)
	}
	stream, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer stream.Close()
	return io.ReadAll(stream)
}

func workbookSheets(reader *zip.Reader) ([]workbookSheet, error) {
	workbookBody, err := zipFileBytes(reader, "xl/workbook.xml")
	if err != nil {
		return nil, err
	}
	relsBody, err := zipFileBytes(reader, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, err
	}
	rels := parseRelationshipXML(relsBody)
	relsByID := map[string]relationship{}
	for _, rel := range rels {
		relsByID[rel.id] = rel
	}
	decoder := xml.NewDecoder(bytes.NewReader(workbookBody))
	sheets := []workbookSheet{}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return sheets, nil
		}
		if err != nil {
			return nil, err
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "sheet" {
			continue
		}
		name := xmlAttr(start.Attr, "name")
		rid := xmlAttr(start.Attr, "id")
		rel, ok := relsByID[rid]
		if name == "" || !ok || !strings.HasSuffix(rel.relType, "/worksheet") {
			continue
		}
		sheets = append(sheets, workbookSheet{
			name: name,
			path: resolveOOXMLPath("xl/workbook.xml", rel.target),
		})
	}
}

func packageFingerprintAssertions(data []byte) map[string]any {
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return map[string]any{}
	}
	partPaths := zipPartPaths(reader)
	contentTypeLines := contentTypeFingerprintLines(zipText(reader, "[Content_Types].xml"))
	relationshipLines := []string{}
	for _, relsPath := range partPaths {
		if !isRelationshipPart(relsPath) {
			continue
		}
		relsXML := zipText(reader, relsPath)
		sourcePart := sourcePartForRelationships(relsPath)
		for _, rel := range parseRelationshipXML([]byte(relsXML)) {
			targetMode := rel.targetMode
			if targetMode == "" {
				targetMode = "Internal"
			}
			resolvedTarget := rel.target
			if targetMode != "External" {
				resolvedTarget = resolveOOXMLPath(sourcePart, rel.target)
			}
			relationshipLines = append(
				relationshipLines,
				strings.Join([]string{rootSourcePart(sourcePart), rel.relType, targetMode, rel.target, resolvedTarget}, "\t"),
			)
		}
	}
	preservedPaths := []string{}
	preservedContentLines := []string{}
	for _, partPath := range partPaths {
		if !isPreservedNonCellPart(partPath) {
			continue
		}
		preservedPaths = append(preservedPaths, partPath)
		preservedContentLines = append(preservedContentLines, partPath+"\t"+sha256Hex(zipBytes(reader, partPath)))
	}
	return map[string]any{
		"roundtripPackagePartCount":             len(partPaths),
		"roundtripPackagePartNamesHash":         hashLines(partPaths),
		"roundtripPackageContentTypeCount":      len(contentTypeLines),
		"roundtripPackageContentTypesHash":      hashLines(contentTypeLines),
		"roundtripPackageRelationshipCount":     len(relationshipLines),
		"roundtripPackageRelationshipGraphHash": hashLines(relationshipLines),
		"roundtripPreservedPartCount":           len(preservedPaths),
		"roundtripPreservedPartNamesHash":       hashLines(preservedPaths),
		"roundtripPreservedPartContentHash":     hashLines(preservedContentLines),
	}
}

func featureSummaryAssertions(data []byte) map[string]any {
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return map[string]any{}
	}
	partPaths := zipPartPaths(reader)
	sheets, _ := workbookSheets(reader)
	featurePartLines := []string{}
	for _, partPath := range partPaths {
		if kind := classifyFeaturePart(partPath); kind != "" {
			featurePartLines = append(featurePartLines, kind+"\t"+partPath)
		}
	}
	worksheetFeatureLines := []string{}
	for _, sheet := range sheets {
		worksheetFeatureLines = append(
			worksheetFeatureLines,
			worksheetFeatureEntries(sheet.path, zipText(reader, sheet.path))...,
		)
	}
	definedNameLines := definedNameEntries(zipText(reader, "xl/workbook.xml"))
	sharedStringFeatureLines := sharedStringFeatureEntries(zipText(reader, "xl/sharedStrings.xml"))
	calcChainFeatureLines := calcChainFeatureEntries(zipText(reader, "xl/calcChain.xml"))
	inventoryLines := append([]string{}, featurePartLines...)
	inventoryLines = append(inventoryLines, worksheetFeatureLines...)
	inventoryLines = append(inventoryLines, definedNameLines...)
	inventoryLines = append(inventoryLines, sharedStringFeatureLines...)
	inventoryLines = append(inventoryLines, calcChainFeatureLines...)
	return map[string]any{
		"roundtripTablePartCount": countPartPaths(partPaths, func(path string) bool {
			return strings.HasPrefix(path, "xl/tables/") && strings.HasSuffix(path, ".xml")
		}),
		"roundtripChartPartCount": countPartPaths(partPaths, func(path string) bool {
			return strings.HasPrefix(path, "xl/charts/") && strings.HasSuffix(path, ".xml")
		}),
		"roundtripChartExPartCount": countPartPaths(partPaths, func(path string) bool {
			return strings.HasPrefix(path, "xl/chartEx/") && strings.HasSuffix(path, ".xml")
		}),
		"roundtripDrawingPartCount": countPartPaths(partPaths, func(path string) bool {
			return strings.HasPrefix(path, "xl/drawings/") && strings.HasSuffix(path, ".xml")
		}),
		"roundtripVmlDrawingPartCount": countPartPaths(partPaths, func(path string) bool {
			return strings.HasPrefix(path, "xl/drawings/") && strings.HasSuffix(path, ".vml")
		}),
		"roundtripPivotTablePartCount": countPartPaths(partPaths, func(path string) bool {
			return strings.HasPrefix(path, "xl/pivotTables/") && strings.HasSuffix(path, ".xml")
		}),
		"roundtripPivotCachePartCount": countPartPaths(partPaths, func(path string) bool {
			return strings.HasPrefix(path, "xl/pivotCache/") && strings.HasSuffix(path, ".xml")
		}),
		"roundtripSlicerPartCount":  countPartPaths(partPaths, isSlicerPart),
		"roundtripCommentPartCount": countPartPaths(partPaths, isCommentPart),
		"roundtripThreadedCommentPartCount": countPartPaths(partPaths, func(path string) bool {
			return strings.HasPrefix(path, "xl/threadedComments/") && strings.HasSuffix(path, ".xml")
		}),
		"roundtripMediaPartCount": countPartPaths(partPaths, func(path string) bool { return strings.HasPrefix(path, "xl/media/") }),
		"roundtripExternalLinkPartCount": countPartPaths(partPaths, func(path string) bool {
			return strings.HasPrefix(path, "xl/externalLinks/") && strings.HasSuffix(path, ".xml")
		}),
		"roundtripConnectionPartCount":                 countPartPaths(partPaths, func(path string) bool { return path == "xl/connections.xml" }),
		"roundtripCustomXmlPartCount":                  countPartPaths(partPaths, func(path string) bool { return strings.HasPrefix(path, "customXml/") }),
		"roundtripWorksheetHyperlinkCount":             countLinesWithPrefix(worksheetFeatureLines, "worksheet-hyperlink\t"),
		"roundtripWorksheetDataValidationCount":        countLinesWithPrefix(worksheetFeatureLines, "worksheet-data-validation\t"),
		"roundtripWorksheetConditionalFormattingCount": countLinesWithPrefix(worksheetFeatureLines, "worksheet-conditional-formatting\t"),
		"roundtripDefinedNameCount":                    len(definedNameLines),
		"roundtripFeaturePartNamesHash":                hashLines(featurePartLines),
		"roundtripFeatureInventoryHash":                hashLines(inventoryLines),
	}
}

func xmlAttr(attrs []xml.Attr, name string) string {
	for _, attr := range attrs {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

func parseRelationshipXML(body []byte) []relationship {
	decoder := xml.NewDecoder(bytes.NewReader(body))
	rels := []relationship{}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return rels
		}
		if err != nil {
			return rels
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}
		targetMode := xmlAttr(start.Attr, "TargetMode")
		if targetMode == "" {
			targetMode = "Internal"
		}
		rels = append(rels, relationship{
			id:         xmlAttr(start.Attr, "Id"),
			relType:    xmlAttr(start.Attr, "Type"),
			target:     xmlAttr(start.Attr, "Target"),
			targetMode: targetMode,
		})
	}
}

func resolveOOXMLPath(sourcePart string, target string) string {
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(pathpkg.Clean(target), "/")
	}
	base := pathpkg.Dir(sourcePart)
	if sourcePart == "" || sourcePart == "." {
		base = "."
	}
	return strings.TrimPrefix(pathpkg.Clean(pathpkg.Join(base, target)), "./")
}

func rootSourcePart(sourcePart string) string {
	if sourcePart == "" {
		return "/"
	}
	return sourcePart
}

func zipPartPaths(reader *zip.Reader) []string {
	paths := []string{}
	for _, file := range reader.File {
		if !strings.HasSuffix(file.Name, "/") {
			paths = append(paths, file.Name)
		}
	}
	sort.Strings(paths)
	return paths
}

func zipBytes(reader *zip.Reader, name string) []byte {
	file := zipFile(reader, name)
	if file == nil {
		return nil
	}
	stream, err := file.Open()
	if err != nil {
		return nil
	}
	defer stream.Close()
	body, err := io.ReadAll(stream)
	if err != nil {
		return nil
	}
	return body
}

func zipText(reader *zip.Reader, name string) string {
	return string(zipBytes(reader, name))
}

func sha256Hex(data []byte) string {
	digest := sha256.Sum256(data)
	return hex.EncodeToString(digest[:])
}

func contentTypeFingerprintLines(xmlText string) []string {
	decoder := xml.NewDecoder(strings.NewReader(xmlText))
	lines := []string{}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return lines
		}
		if err != nil {
			return lines
		}
		start, ok := token.(xml.StartElement)
		if !ok {
			continue
		}
		switch start.Name.Local {
		case "Default":
			extension := xmlAttr(start.Attr, "Extension")
			contentType := xmlAttr(start.Attr, "ContentType")
			if extension != "" && contentType != "" {
				lines = append(lines, "Default\t"+extension+"\t"+contentType)
			}
		case "Override":
			partName := strings.TrimPrefix(xmlAttr(start.Attr, "PartName"), "/")
			contentType := xmlAttr(start.Attr, "ContentType")
			if partName != "" && contentType != "" {
				lines = append(lines, "Override\t"+partName+"\t"+contentType)
			}
		}
	}
}

func isRelationshipPart(path string) bool {
	return path == "_rels/.rels" || strings.HasSuffix(path, ".rels")
}

func sourcePartForRelationships(relsPath string) string {
	if relsPath == "_rels/.rels" {
		return ""
	}
	marker := "/_rels/"
	index := strings.LastIndex(relsPath, marker)
	if index < 0 || !strings.HasSuffix(relsPath, ".rels") {
		return ""
	}
	return relsPath[:index] + "/" + strings.TrimSuffix(relsPath[index+len(marker):], ".rels")
}

func isPreservedNonCellPart(path string) bool {
	if path == "[Content_Types].xml" ||
		path == "_rels/.rels" ||
		strings.HasSuffix(path, ".rels") ||
		path == "xl/workbook.xml" ||
		path == "xl/sharedStrings.xml" ||
		path == "xl/calcChain.xml" ||
		path == "docProps/app.xml" ||
		path == "docProps/core.xml" {
		return false
	}
	if strings.HasPrefix(path, "xl/worksheets/sheet") && strings.HasSuffix(path, ".xml") {
		return false
	}
	return true
}

func classifyFeaturePart(path string) string {
	switch {
	case strings.HasPrefix(path, "xl/tables/") && strings.HasSuffix(path, ".xml"):
		return "table-part"
	case strings.HasPrefix(path, "xl/charts/") && strings.HasSuffix(path, ".xml"):
		return "chart-part"
	case strings.HasPrefix(path, "xl/chartEx/") && strings.HasSuffix(path, ".xml"):
		return "chart-ex-part"
	case strings.HasPrefix(path, "xl/drawings/") && strings.HasSuffix(path, ".xml"):
		return "drawing-part"
	case strings.HasPrefix(path, "xl/drawings/") && strings.HasSuffix(path, ".vml"):
		return "vml-drawing-part"
	case strings.HasPrefix(path, "xl/pivotTables/") && strings.HasSuffix(path, ".xml"):
		return "pivot-table-part"
	case strings.HasPrefix(path, "xl/pivotCache/") && strings.HasSuffix(path, ".xml"):
		return "pivot-cache-part"
	case isSlicerPart(path):
		return "slicer-part"
	case isCommentPart(path):
		return "comment-part"
	case strings.HasPrefix(path, "xl/threadedComments/") && strings.HasSuffix(path, ".xml"):
		return "threaded-comment-part"
	case strings.HasPrefix(path, "xl/media/"):
		return "media-part"
	case strings.HasPrefix(path, "xl/externalLinks/") && strings.HasSuffix(path, ".xml"):
		return "external-link-part"
	case path == "xl/connections.xml":
		return "connection-part"
	case path == "xl/calcChain.xml":
		return "calc-chain-part"
	case strings.HasPrefix(path, "customXml/"):
		return "custom-xml-part"
	default:
		return ""
	}
}

func worksheetFeatureEntries(path string, xmlText string) []string {
	decoder := xml.NewDecoder(strings.NewReader(xmlText))
	lines := []string{}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return lines
		}
		if err != nil {
			return lines
		}
		start, ok := token.(xml.StartElement)
		if !ok {
			continue
		}
		switch start.Name.Local {
		case "mergeCell":
			lines = append(lines, "worksheet-merge-cell\t"+path+"\t"+xmlAttr(start.Attr, "ref"))
		case "autoFilter":
			lines = append(lines, "worksheet-auto-filter\t"+path+"\t"+xmlAttr(start.Attr, "ref"))
		case "sortState":
			lines = append(lines, "worksheet-sort-state\t"+path+"\t"+canonicalXMLAttrs(start.Attr))
		case "sheetProtection":
			lines = append(lines, "worksheet-sheet-protection\t"+path+"\t"+canonicalXMLAttrs(start.Attr))
		case "sheetView":
			lines = append(lines, "worksheet-sheet-view\t"+path+"\t"+canonicalXMLAttrs(start.Attr))
		case "hyperlink":
			lines = append(lines, "worksheet-hyperlink\t"+path+"\t"+xmlAttr(start.Attr, "ref"))
		case "dataValidation":
			lines = append(lines, "worksheet-data-validation\t"+path+"\t"+xmlAttr(start.Attr, "sqref"))
		case "conditionalFormatting":
			lines = append(lines, "worksheet-conditional-formatting\t"+path+"\t"+xmlAttr(start.Attr, "sqref"))
		}
	}
}

func sharedStringFeatureEntries(xmlText string) []string {
	decoder := xml.NewDecoder(strings.NewReader(xmlText))
	lines := []string{}
	index := -1
	inString := false
	hasRichText := false
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return lines
		}
		if err != nil {
			return lines
		}
		switch typed := token.(type) {
		case xml.StartElement:
			if typed.Name.Local == "si" {
				index++
				inString = true
				hasRichText = false
			} else if inString && typed.Name.Local == "r" {
				hasRichText = true
			}
		case xml.EndElement:
			if typed.Name.Local == "si" {
				if hasRichText {
					lines = append(lines, fmt.Sprintf("shared-string-rich-text\t%d\tpresent", index))
				}
				inString = false
			}
		}
	}
}

func calcChainFeatureEntries(xmlText string) []string {
	decoder := xml.NewDecoder(strings.NewReader(xmlText))
	lines := []string{}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return lines
		}
		if err != nil {
			return lines
		}
		start, ok := token.(xml.StartElement)
		if ok && start.Name.Local == "c" {
			lines = append(lines, "calc-chain-cell\t"+canonicalXMLAttrs(start.Attr))
		}
	}
}

func definedNameEntries(xmlText string) []string {
	decoder := xml.NewDecoder(strings.NewReader(xmlText))
	lines := []string{}
	inName := false
	name := ""
	localSheetID := ""
	var value strings.Builder
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return lines
		}
		if err != nil {
			return lines
		}
		switch typed := token.(type) {
		case xml.StartElement:
			if typed.Name.Local == "definedName" {
				inName = true
				name = xmlAttr(typed.Attr, "name")
				localSheetID = xmlAttr(typed.Attr, "localSheetId")
				value.Reset()
			}
		case xml.CharData:
			if inName {
				value.Write([]byte(typed))
			}
		case xml.EndElement:
			if typed.Name.Local == "definedName" {
				lines = append(lines, "defined-name\t"+name+"\t"+localSheetID+"\t"+value.String())
				inName = false
			}
		}
	}
}

func canonicalXMLAttrs(attrs []xml.Attr) string {
	values := make([]string, 0, len(attrs))
	for _, attr := range attrs {
		values = append(values, attr.Name.Local+"="+attr.Value)
	}
	sort.Strings(values)
	return strings.Join(values, ";")
}

func isSlicerPart(path string) bool {
	return (strings.HasPrefix(path, "xl/slicers/") || strings.HasPrefix(path, "xl/slicerCaches/")) && strings.HasSuffix(path, ".xml")
}

func isCommentPart(path string) bool {
	return (strings.HasPrefix(path, "xl/comments") || strings.HasPrefix(path, "xl/comments/")) && strings.HasSuffix(path, ".xml")
}

func countPartPaths(paths []string, predicate func(string) bool) int {
	count := 0
	for _, path := range paths {
		if predicate(path) {
			count++
		}
	}
	return count
}

func countLinesWithPrefix(lines []string, prefix string) int {
	count := 0
	for _, line := range lines {
		if strings.HasPrefix(line, prefix) {
			count++
		}
	}
	return count
}

func prefixKey(prefix string, name string) string {
	if prefix == "" {
		return name
	}
	return prefix + strings.ToUpper(name[:1]) + name[1:]
}

func positiveIntAttr(attrs []xml.Attr, name string) int {
	raw := xmlAttr(attrs, name)
	if raw == "" {
		return 0
	}
	value, err := strconv.Atoi(raw)
	if err != nil || value <= 0 {
		return 0
	}
	return value
}

func columnIndexFromCellRef(ref string) int {
	value := 0
	for _, ch := range ref {
		if ch < 'A' || ch > 'Z' {
			break
		}
		value = value*26 + int(ch-'A'+1)
	}
	return value
}

func rowIndexFromCellRef(ref string) int {
	value := 0
	inRow := false
	for _, ch := range ref {
		if ch >= '0' && ch <= '9' {
			inRow = true
			value = value*10 + int(ch-'0')
		} else if inRow {
			break
		}
	}
	return value
}

func columnNameCache(cols int) []string {
	names := make([]string, max(0, cols))
	for col := 1; col <= cols; col++ {
		names[col-1] = columnName(col)
	}
	return names
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
	hasher := newOrderedLineHasher()
	for _, line := range sorted {
		hasher.Add(line)
	}
	return hasher.Digest()
}

type orderedLineHasher struct {
	hash hashHash
}

type hashHash interface {
	Write([]byte) (int, error)
	Sum([]byte) []byte
}

func newOrderedLineHasher() *orderedLineHasher {
	return &orderedLineHasher{hash: sha256.New()}
}

func (hasher *orderedLineHasher) Add(line string) {
	hasher.hash.Write([]byte(strconv.Itoa(len(utf16.Encode([]rune(line))))))
	hasher.hash.Write([]byte(":"))
	hasher.hash.Write([]byte(line))
	hasher.hash.Write([]byte("\n"))
}

func (hasher *orderedLineHasher) Digest() string {
	return hex.EncodeToString(hasher.hash.Sum(nil))
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

func UsedRange(sheetName string, coords []Coord) string {
	if len(coords) == 0 {
		return sheetName + "!empty"
	}
	minRow := math.MaxInt
	minCol := math.MaxInt
	maxRow := 0
	maxCol := 0
	for _, coord := range coords {
		minRow = min(minRow, coord.Row)
		minCol = min(minCol, coord.Col)
		maxRow = max(maxRow, coord.Row)
		maxCol = max(maxCol, coord.Col)
	}
	return fmt.Sprintf("%s!%s%d:%s%d", sheetName, columnName(minCol), minRow, columnName(maxCol), maxRow)
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
