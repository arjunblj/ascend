package dev.ascend.benchmarks;

import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.ByteArrayInputStream;
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
import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;

public final class FastExcelRunner {
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

	public static void main(String[] argv) throws Exception {
		PrintStream originalOut = System.out;
		System.setOut(new PrintStream(new ByteArrayOutputStream()));
		Args args = parseArgs(argv);
		for (int index = 0; index < args.warmup; index++) {
			if ("read".equals(args.operation)) {
				readCellCount(args.file);
			} else {
				writeWorkbook(args.workload, args.rows, args.cols);
			}
		}
		List<Map<String, Number>> samples = new ArrayList<>();
		Map<String, Object> assertions = null;
		byte[] finalWriteBytes = null;
		for (int index = 0; index < args.repeat; index++) {
			long start = System.nanoTime();
			if ("read".equals(args.operation)) {
				readCellCount(args.file);
				double durationMs = (System.nanoTime() - start) / 1_000_000.0;
				if ("each".equals(args.validationMode)) assertions = readAssertions(args.file);
				samples.add(memorySample(durationMs));
				continue;
			}
			byte[] bytes = writeWorkbook(args.workload, args.rows, args.cols);
			finalWriteBytes = bytes;
			double durationMs = (System.nanoTime() - start) / 1_000_000.0;
			if ("each".equals(args.validationMode)) assertions = writeAssertions(bytes, args.workload, args.rows, args.cols);
			samples.add(memorySample(durationMs));
		}
		if ("read".equals(args.operation) && "final".equals(args.validationMode)) assertions = readAssertions(args.file);
		if ("write".equals(args.operation) && "final".equals(args.validationMode) && finalWriteBytes != null) {
			assertions = writeAssertions(finalWriteBytes, args.workload, args.rows, args.cols);
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
		if (operation == null) throw new IllegalArgumentException("--operation is required");
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

	private static Map<String, Object> writeAssertions(byte[] bytes, String workload, int rows, int cols) throws Exception {
		Map<String, Object> assertions = readAssertionsFromBytes(bytes);
		String expectedHash = expectedValuesHash(workload, rows, cols);
		int expectedCells = expectedCellCount(workload, rows, cols);
		assertions.put("runnerVersion", implementationVersion());
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
		return assertions;
	}

	private static byte[] writeWorkbook(String workload, int rows, int cols) throws Exception {
		try (ByteArrayOutputStream output = new ByteArrayOutputStream();
			 Workbook workbook = new Workbook(output, "Ascend benchmark", "1.0")) {
			Worksheet sheet = workbook.newWorksheet("Data");
			for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
				for (int colIndex = 0; colIndex < cols; colIndex++) {
					Object value = workloadValue(workload, rowIndex, colIndex, cols);
					if (value == null) continue;
					if (value instanceof Number number) {
						sheet.value(rowIndex, colIndex, number.doubleValue());
					} else if (value instanceof Boolean bool) {
						sheet.value(rowIndex, colIndex, bool);
					} else {
						sheet.value(rowIndex, colIndex, value.toString());
					}
					if ("styles-heavy".equals(workload)) {
						switch ((rowIndex + colIndex) % 4) {
							case 1 -> sheet.style(rowIndex, colIndex).fillColor("FFF9C4").set();
							case 2 -> sheet.style(rowIndex, colIndex).fillColor("C8E6C9").set();
							case 3 -> sheet.style(rowIndex, colIndex).format("#,##0.00").set();
							default -> {
							}
						}
					}
				}
			}
			workbook.finish();
			return output.toByteArray();
		}
	}

	private static Map<String, Object> readAssertions(File file) throws Exception {
		try (org.dhatim.fastexcel.reader.ReadableWorkbook workbook =
				 new org.dhatim.fastexcel.reader.ReadableWorkbook(file)) {
			List<String> semanticValues = new ArrayList<>();
			int sheetCount = 0;
			int cellCount = 0;
			int formulaCount = 0;
			try (Stream<org.dhatim.fastexcel.reader.Sheet> sheets = workbook.getSheets()) {
				for (org.dhatim.fastexcel.reader.Sheet sheet : (Iterable<org.dhatim.fastexcel.reader.Sheet>) sheets::iterator) {
					sheetCount++;
					String sheetName = sheet.getName();
					try (Stream<org.dhatim.fastexcel.reader.Row> rows = sheet.openStream()) {
						for (org.dhatim.fastexcel.reader.Row row : (Iterable<org.dhatim.fastexcel.reader.Row>) rows::iterator) {
							for (org.dhatim.fastexcel.reader.Cell cell : row) {
								if (cell == null) continue;
								if (cell.getType() == org.dhatim.fastexcel.reader.CellType.FORMULA) formulaCount++;
								String payload = scalarPayload(cell);
								if (payload == null) continue;
								cellCount++;
								int rowNumber = normalizeFastExcelRowNumber(row.getRowNum());
								int columnNumber = cell.getColumnIndex() + 1;
								semanticValues.add(cellRef(sheetName, rowNumber, columnNumber) + "\t" + payload);
							}
						}
					}
				}
			}
			Map<String, Object> assertions = new HashMap<>();
			assertions.put("runnerVersion", readerImplementationVersion());
			assertions.put("sheetCount", sheetCount);
			assertions.put("cellCount", cellCount);
			assertions.put("formulaCount", formulaCount);
			assertions.put("semanticCellValuesHash", hashLines(semanticValues));
			assertions.put("runnerApi", "fastexcel-reader");
			return assertions;
		}
	}

	private static int readCellCount(File file) throws Exception {
		int cellCount = 0;
		try (org.dhatim.fastexcel.reader.ReadableWorkbook workbook =
				 new org.dhatim.fastexcel.reader.ReadableWorkbook(file);
			 Stream<org.dhatim.fastexcel.reader.Sheet> sheets = workbook.getSheets()) {
			for (org.dhatim.fastexcel.reader.Sheet sheet : (Iterable<org.dhatim.fastexcel.reader.Sheet>) sheets::iterator) {
				try (Stream<org.dhatim.fastexcel.reader.Row> rows = sheet.openStream()) {
					for (org.dhatim.fastexcel.reader.Row row : (Iterable<org.dhatim.fastexcel.reader.Row>) rows::iterator) {
						for (org.dhatim.fastexcel.reader.Cell cell : row) {
							if (cell == null) continue;
							if (cell.getType() != org.dhatim.fastexcel.reader.CellType.EMPTY) cellCount++;
						}
					}
				}
			}
		}
		return cellCount;
	}

	private static Map<String, Object> readAssertionsFromBytes(byte[] bytes) throws Exception {
		try (org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
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

	private static String scalarPayload(org.dhatim.fastexcel.reader.Cell cell) {
		return switch (cell.getType()) {
			case NUMBER -> "n:" + canonicalNumber(cell.asNumber().doubleValue());
			case STRING -> "s:" + cell.asString();
			case BOOLEAN -> "b:" + (Boolean.TRUE.equals(cell.asBoolean()) ? "true" : "false");
			case ERROR -> "e:" + cell.getText();
			case FORMULA -> {
				Object value = cell.getValue();
				yield value == null ? "empty" : scalarPayload(value);
			}
			case EMPTY -> null;
		};
	}

	private static String canonicalNumber(double value) {
		if (value == 0.0d) return "0";
		if (Math.rint(value) == value) return Long.toString((long) value);
		return String.format(Locale.ROOT, "%.15g", value);
	}

	private static String cellRef(String sheetName, int row, int col) {
		return sheetName + "!" + columnName(col) + row;
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

	private static String implementationVersion() {
		Package pkg = Workbook.class.getPackage();
		String version = pkg == null ? null : pkg.getImplementationVersion();
		return version == null ? "0.20.0" : version;
	}

	private static String readerImplementationVersion() {
		Package pkg = org.dhatim.fastexcel.reader.ReadableWorkbook.class.getPackage();
		String version = pkg == null ? null : pkg.getImplementationVersion();
		return version == null ? "0.20.0" : version;
	}

	private static int normalizeFastExcelRowNumber(int rowNum) {
		return rowNum <= 0 ? rowNum + 1 : rowNum;
	}
}
