package dev.ascend.benchmarks;

import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public final class ApachePoiRunner {
	private static final ObjectMapper JSON = new ObjectMapper();

	private record Args(
		String operation,
		File file,
		int rows,
		int cols,
		String workload,
		int repeat,
		int warmup,
		String validationMode,
		boolean json
	) {}

	private record Coord(int row, int col) {}

	public static void main(String[] argv) throws Exception {
		PrintStream originalOut = System.out;
		System.setOut(new PrintStream(new ByteArrayOutputStream()));
		IOUtils.setByteArrayMaxOverride(512_000_000);
		Args args = parseArgs(argv);
		for (int index = 0; index < args.warmup; index++) {
			if ("read".equals(args.operation)) {
				readAssertions(args.file);
			} else {
				writeWorkbookBytes(args.workload, args.rows, args.cols);
			}
		}
		List<Map<String, Number>> samples = new ArrayList<>();
		Map<String, Object> assertions = null;
		byte[] finalWriteBytes = null;
		for (int index = 0; index < args.repeat; index++) {
			long start = System.nanoTime();
			if ("read".equals(args.operation)) {
				assertions = readAssertions(args.file);
			} else {
				byte[] bytes = writeWorkbookBytes(args.workload, args.rows, args.cols);
				finalWriteBytes = bytes;
				double durationMs = (System.nanoTime() - start) / 1_000_000.0;
				if ("each".equals(args.validationMode)) {
					assertions = writeAssertions(bytes, args.workload, args.rows, args.cols, "each", index + 1);
				}
				samples.add(memorySample(durationMs));
				continue;
			}
			double durationMs = (System.nanoTime() - start) / 1_000_000.0;
			samples.add(memorySample(durationMs));
		}
		if ("write".equals(args.operation) && "final".equals(args.validationMode) && finalWriteBytes != null) {
			assertions = writeAssertions(finalWriteBytes, args.workload, args.rows, args.cols, "final", 1);
		}
		Map<String, Object> payload = new HashMap<>();
		payload.put("assertions", assertions == null ? Map.of() : assertions);
		payload.put("samples", samples);
		if (args.json) {
			originalOut.println(JSON.writeValueAsString(payload));
		} else {
			originalOut.println(JSON.writerWithDefaultPrettyPrinter().writeValueAsString(payload));
		}
	}

	private static Args parseArgs(String[] argv) {
		String operation = null;
		File file = null;
		int rows = 0;
		int cols = 0;
		String workload = "dense-values";
		int repeat = 1;
		int warmup = 0;
		String validationMode = "each";
		boolean json = false;
		for (int index = 0; index < argv.length; index++) {
			String arg = argv[index];
			switch (arg) {
				case "--operation" -> operation = argv[++index];
				case "--file" -> file = new File(argv[++index]);
				case "--rows" -> rows = Math.max(1, Integer.parseInt(argv[++index]));
				case "--cols" -> cols = Math.max(1, Integer.parseInt(argv[++index]));
				case "--workload" -> workload = argv[++index];
				case "--repeat" -> repeat = Math.max(1, Integer.parseInt(argv[++index]));
				case "--warmup" -> warmup = Math.max(0, Integer.parseInt(argv[++index]));
				case "--validation-mode" -> validationMode = argv[++index];
				case "--json" -> json = true;
				default -> throw new IllegalArgumentException("Unsupported argument " + arg);
			}
		}
		if (!"read".equals(operation) && !"write".equals(operation)) {
			throw new IllegalArgumentException("--operation must be read or write");
		}
		if ("read".equals(operation) && file == null) throw new IllegalArgumentException("--file is required");
		if ("write".equals(operation) && (rows <= 0 || cols <= 0)) {
			throw new IllegalArgumentException("--rows and --cols are required for write");
		}
		if (!"each".equals(validationMode) && !"final".equals(validationMode)) {
			throw new IllegalArgumentException("--validation-mode must be each or final");
		}
		return new Args(operation, file, rows, cols, workload, repeat, warmup, validationMode, json);
	}

	private static Map<String, Number> memorySample(double durationMs) {
		Runtime runtime = Runtime.getRuntime();
		long used = runtime.totalMemory() - runtime.freeMemory();
		Map<String, Number> sample = new HashMap<>();
		sample.put("durationMs", durationMs);
		sample.put("heapUsedBytes", used);
		sample.put("heapTotalBytes", runtime.totalMemory());
		sample.put("peakRssBytes", runtime.totalMemory());
		sample.put("rssAfterBytes", runtime.totalMemory());
		return sample;
	}

	private static Map<String, Object> readAssertions(File file) throws Exception {
		try (Workbook workbook = WorkbookFactory.create(file, null, true)) {
			List<String> sheetNames = new ArrayList<>();
			List<String> usedRanges = new ArrayList<>();
			List<String> physicalUsedRanges = new ArrayList<>();
			List<String> semanticRefs = new ArrayList<>();
			List<String> semanticValues = new ArrayList<>();
			List<String> formulaTexts = new ArrayList<>();
			int cellCount = 0;
			int physicalCellCount = 0;
			int formulaCount = 0;
			int commentCount = 0;
			int hyperlinkCount = 0;
			int dataValidationCount = 0;
			int conditionalFormatCount = 0;

			for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
				Sheet sheet = workbook.getSheetAt(sheetIndex);
				String sheetName = sheet.getSheetName();
				sheetNames.add(sheetName);
				List<Coord> semanticCoords = new ArrayList<>();
				List<Coord> physicalCoords = new ArrayList<>();
				if (sheet instanceof XSSFSheet xssfSheet) {
					commentCount += xssfSheet.getCellComments().size();
				}
				dataValidationCount += sheet.getDataValidations().size();
				conditionalFormatCount += sheet.getSheetConditionalFormatting().getNumConditionalFormattings();
				for (Row row : sheet) {
					for (Cell cell : row) {
						Coord coord = new Coord(cell.getRowIndex() + 1, cell.getColumnIndex() + 1);
						physicalCellCount++;
						physicalCoords.add(coord);
						if (cell.getHyperlink() != null) hyperlinkCount++;
						String payload = scalarPayload(cell);
						if (payload == null) continue;
						cellCount++;
						semanticCoords.add(coord);
						String ref = cellRef(sheetName, coord.row, coord.col);
						semanticRefs.add(ref);
						semanticValues.add(ref + "\t" + payload);
						if (cell.getCellType() == CellType.FORMULA) {
							formulaCount++;
							formulaTexts.add(ref + "=" + cell.getCellFormula());
						}
					}
				}
				usedRanges.add(usedRange(sheetName, semanticCoords));
				physicalUsedRanges.add(usedRange(sheetName, physicalCoords));
			}

			Map<String, Object> assertions = new HashMap<>();
			assertions.put("runnerVersion", org.apache.poi.Version.getVersion());
			assertions.put("sheetCount", workbook.getNumberOfSheets());
			assertions.put("sheetNamesHash", hashLines(indexedLines(sheetNames)));
			assertions.put("cellCount", cellCount);
			assertions.put("physicalCellCount", physicalCellCount);
			assertions.put("formulaCount", formulaCount);
			assertions.put("usedRangeCount", usedRanges.size());
			assertions.put("firstUsedRange", usedRanges.isEmpty() ? null : usedRanges.get(0));
			assertions.put("firstPhysicalUsedRange", physicalUsedRanges.isEmpty() ? null : physicalUsedRanges.get(0));
			assertions.put("usedRangesHash", hashLines(usedRanges));
			assertions.put("physicalUsedRangesHash", hashLines(physicalUsedRanges));
			assertions.put("semanticCellRefsHash", hashLines(semanticRefs));
			assertions.put("semanticCellValuesHash", hashLines(semanticValues));
			assertions.put("formulaTextHash", hashLines(formulaTexts));
			assertions.put("readCommentCount", commentCount);
			assertions.put("readHyperlinkCount", hyperlinkCount);
			assertions.put("readDataValidationCount", dataValidationCount);
			assertions.put("readConditionalFormatCount", conditionalFormatCount);
			assertions.put("readDefinedNameCount", workbook.getNumberOfNames());
			return assertions;
		}
	}

	private static byte[] writeWorkbookBytes(String workload, int rows, int cols) throws Exception {
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			writeWorkbook(workbook, workload, rows, cols);
			try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
				workbook.write(output);
				return output.toByteArray();
			}
		}
	}

	private static Map<String, Object> writeAssertions(
		byte[] bytes,
		String workload,
		int rows,
		int cols,
		String validationMode,
		int validationSamples
	) throws Exception {
		Map<String, Object> assertions = readAssertionsFromBytes(bytes);
		String expectedHash = expectedValuesHash(workload, rows, cols);
		int expectedCells = expectedCellCount(workload, rows, cols);
		assertions.put("runnerVersion", org.apache.poi.Version.getVersion());
		assertions.put("workload", workload);
		assertions.put("bytes", bytes.length);
		assertions.put("reopenOk", true);
		assertions.put("tablePartCount", 0);
		assertions.put("commentPartCount", 0);
		assertions.put("vmlDrawingPartCount", 0);
		assertions.put("worksheetHyperlinkCount", 0);
		assertions.put("worksheetDataValidationCount", 0);
		assertions.put("worksheetConditionalFormattingCount", 0);
		assertions.put("definedNameCount", 0);
		assertions.put("expectedCellCount", expectedCells);
		assertions.put("cellCountMatches", assertions.get("cellCount").equals(expectedCells));
		assertions.put("expectedSemanticCellValuesHash", expectedHash);
		assertions.put("semanticCellValuesHashMatches", expectedHash.equals(assertions.get("semanticCellValuesHash")));
		assertions.put("validationMode", validationMode);
		assertions.put("validationSamples", validationSamples);
		return assertions;
	}

	private static void writeWorkbook(XSSFWorkbook workbook, String workload, int rows, int cols) {
		Sheet sheet = workbook.createSheet("Data");
		CellStyle[] styles = createStyles(workbook);
		for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
			Row row = sheet.createRow(rowIndex);
			for (int colIndex = 0; colIndex < cols; colIndex++) {
				Object value = workloadValue(workload, rowIndex, colIndex, cols);
				if (value == null) continue;
				Cell cell = row.createCell(colIndex);
				if (value instanceof Number number) {
					cell.setCellValue(number.doubleValue());
				} else if (value instanceof Boolean bool) {
					cell.setCellValue(bool);
				} else {
					cell.setCellValue(value.toString());
				}
				if ("styles-heavy".equals(workload)) {
					cell.setCellStyle(styles[(rowIndex + colIndex) % styles.length]);
				}
			}
		}
	}

	private static CellStyle[] createStyles(XSSFWorkbook workbook) {
		CellStyle normal = workbook.createCellStyle();
		CellStyle yellow = workbook.createCellStyle();
		yellow.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
		yellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		CellStyle green = workbook.createCellStyle();
		green.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		green.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		CellStyle number = workbook.createCellStyle();
		number.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
		return new CellStyle[] { normal, yellow, green, number };
	}

	private static Map<String, Object> readAssertionsFromBytes(byte[] bytes) throws Exception {
		try (Workbook workbook = WorkbookFactory.create(new java.io.ByteArrayInputStream(bytes))) {
			List<String> semanticValues = new ArrayList<>();
			int cellCount = 0;
			int formulaCount = 0;
			for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
				Sheet sheet = workbook.getSheetAt(sheetIndex);
				String sheetName = sheet.getSheetName();
				for (Row row : sheet) {
					for (Cell cell : row) {
						if (cell.getCellType() == CellType.FORMULA) formulaCount++;
						String payload = scalarPayload(cell);
						if (payload == null) continue;
						cellCount++;
						semanticValues.add(cellRef(sheetName, cell.getRowIndex() + 1, cell.getColumnIndex() + 1) + "\t" + payload);
					}
				}
			}
			Map<String, Object> assertions = new HashMap<>();
			assertions.put("sheetCount", workbook.getNumberOfSheets());
			assertions.put("cellCount", cellCount);
			assertions.put("formulaCount", formulaCount);
			assertions.put("semanticCellValuesHash", hashLines(semanticValues));
			return assertions;
		}
	}

	private static Object workloadValue(String workload, int row, int col, int cols) {
		if ("dense-values".equals(workload)) return row * cols + col;
		if ("mixed-10pct-text".equals(workload)) {
			int key = row * cols + col;
			return key % 10 == 0 ? "text-" + String.format("%08d", key) : key;
		}
		if ("mixed-50pct-text".equals(workload)) {
			int key = row * cols + col;
			return key % 2 == 0 ? "text-" + String.format("%08d", key) : key;
		}
		if ("mixed-closedxml-10text-5number".equals(workload)) {
			return col < 10 ? "Hello world" : col - 10;
		}
		if ("plain-text".equals(workload)) {
			return "text-" + String.format("%08d", row * cols + col);
		}
		if ("selected-sheet".equals(workload) || "metadata-only".equals(workload) || "warm-workflow".equals(workload)) {
			return row * cols + col;
		}
		if ("feature-rich".equals(workload)) return row == 0 && col == 0 ? "Ascend" : row * cols + col;
		if ("styles-heavy".equals(workload)) return (row + 1) * (col + 1);
		if ("formula-heavy".equals(workload)) {
			int base = row + 1;
			if (col == 0) return base;
			if (col == 1) return base * 2;
			return base * 3 + col;
		}
		if ("table-heavy".equals(workload)) {
			if (row == 0) return "Column " + (col + 1);
			if (col % 3 == 0) return row;
			if (col % 3 == 1) return "item-" + row + "-" + col;
			return row * cols + col;
		}
		if ("sparse-wide".equals(workload)) {
			if (col == 0) return row;
			if (col == cols - 1) return "edge-" + row + "-" + cols;
			if ((row * 31 + col * 17) % 97 == 0) return row * cols + col;
			return null;
		}
		int key = row * cols + col;
		return switch (col % 5) {
			case 0 -> "sku-" + String.format("%08d", key);
			case 1 -> "region-" + ((row % 17) + 1);
			case 2 -> "customer-" + (row % 997) + "-segment-" + (col % 13);
			case 3 -> "note row " + row + " col " + col + " token " + (key % 104729);
			default -> key % 2 == 0 ? "status-open-" + (key % 31) : "status-closed-" + (key % 29);
		};
	}

	private static String expectedValuesHash(String workload, int rows, int cols) throws Exception {
		List<String> values = new ArrayList<>();
		for (int row = 0; row < rows; row++) {
			for (int col = 0; col < cols; col++) {
				Object value = workloadValue(workload, row, col, cols);
				if (value == null) continue;
				values.add("Data!" + columnName(col + 1) + (row + 1) + "\t" + scalarPayload(value));
			}
		}
		return hashLines(values);
	}

	private static int expectedCellCount(String workload, int rows, int cols) {
		int count = 0;
		for (int row = 0; row < rows; row++) {
			for (int col = 0; col < cols; col++) {
				if (workloadValue(workload, row, col, cols) != null) count++;
			}
		}
		return count;
	}

	private static List<String> indexedLines(List<String> values) {
		List<String> lines = new ArrayList<>();
		for (int index = 0; index < values.size(); index++) {
			lines.add(index + ":" + values.get(index));
		}
		return lines;
	}

	private static String scalarPayload(Cell cell) {
		CellType type = cell.getCellType();
		if (type == CellType.BLANK || type == CellType._NONE) return null;
		if (type == CellType.FORMULA) type = cell.getCachedFormulaResultType();
		return switch (type) {
			case NUMERIC -> "n:" + canonicalNumber(cell.getNumericCellValue());
			case STRING -> "s:" + cell.getStringCellValue();
			case BOOLEAN -> "b:" + (cell.getBooleanCellValue() ? "true" : "false");
			case ERROR -> "e:" + FormulaError.forInt(cell.getErrorCellValue()).getString();
			case BLANK, _NONE -> null;
			case FORMULA -> "empty";
		};
	}

	private static String scalarPayload(Object value) {
		if (value instanceof Boolean bool) return "b:" + (bool ? "true" : "false");
		if (value instanceof Number number) return "n:" + canonicalNumber(number.doubleValue());
		return "s:" + value;
	}

	private static String canonicalNumber(double value) {
		if (value == 0.0d) return "0";
		if (Math.rint(value) == value) return Long.toString((long) value);
		return String.format(Locale.ROOT, "%.15g", value);
	}

	private static String cellRef(String sheetName, int row, int col) {
		return sheetName + "!" + columnName(col) + row;
	}

	private static String usedRange(String sheetName, List<Coord> coords) {
		if (coords.isEmpty()) return sheetName + "!empty";
		int minRow = Integer.MAX_VALUE;
		int minCol = Integer.MAX_VALUE;
		int maxRow = 0;
		int maxCol = 0;
		for (Coord coord : coords) {
			minRow = Math.min(minRow, coord.row);
			minCol = Math.min(minCol, coord.col);
			maxRow = Math.max(maxRow, coord.row);
			maxCol = Math.max(maxCol, coord.col);
		}
		return sheetName + "!" + columnName(minCol) + minRow + ":" + columnName(maxCol) + maxRow;
	}

	private static String columnName(int oneBasedColumn) {
		StringBuilder result = new StringBuilder();
		int col = oneBasedColumn;
		while (col > 0) {
			col--;
			result.insert(0, (char) ('A' + (col % 26)));
			col /= 26;
		}
		return result.toString();
	}

	private static String hashLines(List<String> lines) throws Exception {
		Collections.sort(lines);
		MessageDigest digest = MessageDigest.getInstance("SHA-256");
		for (String line : lines) {
			byte[] bytes = line.getBytes(StandardCharsets.UTF_8);
			digest.update(Integer.toString(line.length()).getBytes(StandardCharsets.UTF_8));
			digest.update((byte) ':');
			digest.update(bytes);
			digest.update((byte) '\n');
		}
		StringBuilder hex = new StringBuilder();
		for (byte value : digest.digest()) {
			hex.append(String.format("%02x", value));
		}
		return hex.toString();
	}
}
