use calamine::{open_workbook_auto, Data, Reader};
use rust_xlsxwriter::{Format, FormatAlign, Formula, Table, TableColumn, Workbook};
use serde_json::json;
use sha2::{Digest, Sha256};
use std::env;
use std::fs::{read, remove_file};
use std::path::PathBuf;
use std::time::{Instant, SystemTime, UNIX_EPOCH};

const RUST_XLSXWRITER_VERSION: &str = "0.94.0";

#[derive(Clone, Copy, PartialEq, Eq)]
enum Workload {
    DenseValues,
    Mixed10PctText,
    Mixed50PctText,
    PlainText,
    StringHeavy,
    SparseWide,
    StylesHeavy,
    FormulaHeavy,
    TableHeavy,
    FeatureRich,
}

struct Args {
    operation: String,
    rows: usize,
    cols: usize,
    workload: Workload,
    repeat: usize,
    warmup: usize,
    pretty_json: bool,
}

fn main() {
    if let Err(error) = run() {
        eprintln!("{error}");
        std::process::exit(1);
    }
}

fn run() -> Result<(), String> {
    let args = parse_args()?;
    if args.operation != "write" {
        return Err(format!(
            "Unsupported --operation {:?}. Expected write.",
            args.operation
        ));
    }

    for _ in 0..args.warmup {
        let path = write_workbook(args.workload, args.rows, args.cols)?;
        let _ = remove_file(path);
    }

    let mut samples = Vec::with_capacity(args.repeat);
    let mut assertions = None;
    for _ in 0..args.repeat {
        let start = Instant::now();
        let path = write_workbook(args.workload, args.rows, args.cols)?;
        let duration_ms = start.elapsed().as_secs_f64() * 1000.0;
        let bytes = read(&path).map_err(|error| error.to_string())?;
        assertions = Some(write_assertions(
            &path,
            bytes.len(),
            args.workload,
            args.rows,
            args.cols,
        )?);
        let _ = remove_file(path);
        let peak_rss_bytes = peak_rss_bytes();
        samples.push(json!({
            "durationMs": duration_ms,
            "peakRssBytes": peak_rss_bytes,
            "rssAfterBytes": peak_rss_bytes
        }));
    }

    let payload = json!({
        "assertions": assertions.ok_or("No samples were produced")?,
        "samples": samples,
    });
    if args.pretty_json {
        println!(
            "{}",
            serde_json::to_string_pretty(&payload).map_err(|error| error.to_string())?
        );
    } else {
        println!(
            "{}",
            serde_json::to_string(&payload).map_err(|error| error.to_string())?
        );
    }
    Ok(())
}

fn parse_args() -> Result<Args, String> {
    let mut operation = None;
    let mut rows = None;
    let mut cols = None;
    let mut workload = Workload::DenseValues;
    let mut repeat = 1usize;
    let mut warmup = 0usize;
    let mut pretty_json = true;
    let mut args = env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--operation" => operation = args.next(),
            "--rows" => rows = Some(parse_count("--rows", args.next())?),
            "--cols" => cols = Some(parse_count("--cols", args.next())?),
            "--workload" => workload = parse_workload(args.next())?,
            "--repeat" => repeat = parse_count("--repeat", args.next())?,
            "--warmup" => warmup = parse_count("--warmup", args.next())?,
            "--json" => pretty_json = false,
            other => return Err(format!("Unsupported argument {other:?}")),
        }
    }
    Ok(Args {
        operation: operation.ok_or("Missing --operation")?,
        rows: rows.ok_or("Missing --rows")?,
        cols: cols.ok_or("Missing --cols")?,
        workload,
        repeat: repeat.max(1),
        warmup,
        pretty_json,
    })
}

fn parse_count(name: &str, value: Option<String>) -> Result<usize, String> {
    value
        .ok_or_else(|| format!("Missing value for {name}"))?
        .parse::<usize>()
        .map_err(|_| format!("{name} must be a non-negative integer"))
}

fn parse_workload(value: Option<String>) -> Result<Workload, String> {
    match value.as_deref().unwrap_or("dense-values") {
        "dense-values" => Ok(Workload::DenseValues),
        "mixed-10pct-text" => Ok(Workload::Mixed10PctText),
        "mixed-50pct-text" => Ok(Workload::Mixed50PctText),
        "plain-text" => Ok(Workload::PlainText),
        "string-heavy" => Ok(Workload::StringHeavy),
        "sparse-wide" => Ok(Workload::SparseWide),
        "styles-heavy" => Ok(Workload::StylesHeavy),
        "formula-heavy" => Ok(Workload::FormulaHeavy),
        "table-heavy" => Ok(Workload::TableHeavy),
        "feature-rich" => Ok(Workload::FeatureRich),
        other => Err(format!("Unsupported --workload {other:?}")),
    }
}

fn workload_name(workload: Workload) -> &'static str {
    match workload {
        Workload::DenseValues => "dense-values",
        Workload::Mixed10PctText => "mixed-10pct-text",
        Workload::Mixed50PctText => "mixed-50pct-text",
        Workload::PlainText => "plain-text",
        Workload::StringHeavy => "string-heavy",
        Workload::SparseWide => "sparse-wide",
        Workload::StylesHeavy => "styles-heavy",
        Workload::FormulaHeavy => "formula-heavy",
        Workload::TableHeavy => "table-heavy",
        Workload::FeatureRich => "feature-rich",
    }
}

fn write_workbook(workload: Workload, rows: usize, cols: usize) -> Result<PathBuf, String> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name("Data")
        .map_err(|error| error.to_string())?;
    let style_format = Format::new().set_align(FormatAlign::Center);

    for row in 0..rows {
        for col in 0..cols {
            let Some(value) = workload_value(workload, row, col, cols) else {
                continue;
            };
            let row_num = u32::try_from(row).map_err(|_| "row exceeds u32".to_string())?;
            let col_num = u16::try_from(col).map_err(|_| "column exceeds u16".to_string())?;
            match value {
                WorkloadValue::Number(value) => {
                    if workload == Workload::StylesHeavy {
                        worksheet
                            .write_with_format(row_num, col_num, value as f64, &style_format)
                            .map_err(|error| error.to_string())?;
                    } else {
                        worksheet
                            .write_number(row_num, col_num, value as f64)
                            .map_err(|error| error.to_string())?;
                    }
                }
                WorkloadValue::Text(value) => {
                    worksheet
                        .write_string(row_num, col_num, &value)
                        .map_err(|error| error.to_string())?;
                }
                WorkloadValue::Formula { formula, result } => {
                    worksheet
                        .write_formula(
                            row_num,
                            col_num,
                            Formula::new(&formula).set_result(result.to_string()),
                        )
                        .map_err(|error| error.to_string())?;
                }
            }
        }
    }

    if workload == Workload::TableHeavy && rows > 0 && cols > 0 {
        let columns: Vec<TableColumn> = (0..cols)
            .map(|index| TableColumn::new().set_header(format!("Column {}", index + 1)))
            .collect();
        let table = Table::new().set_columns(&columns);
        worksheet
            .add_table(
                0,
                0,
                u32::try_from(rows - 1).map_err(|_| "row exceeds u32".to_string())?,
                u16::try_from(cols - 1).map_err(|_| "column exceeds u16".to_string())?,
                &table,
            )
            .map_err(|error| error.to_string())?;
    }

    let path = temp_path();
    workbook.save(&path).map_err(|error| error.to_string())?;
    Ok(path)
}

enum WorkloadValue {
    Number(i64),
    Text(String),
    Formula { formula: String, result: i64 },
}

fn workload_value(
    workload: Workload,
    row: usize,
    col: usize,
    cols: usize,
) -> Option<WorkloadValue> {
    match workload {
        Workload::DenseValues => Some(WorkloadValue::Number((row * cols + col) as i64)),
        Workload::Mixed10PctText => {
            let key = row * cols + col;
            if key % 10 == 0 {
                Some(WorkloadValue::Text(format!("text-{key:08}")))
            } else {
                Some(WorkloadValue::Number(key as i64))
            }
        }
        Workload::Mixed50PctText => {
            let key = row * cols + col;
            if key % 2 == 0 {
                Some(WorkloadValue::Text(format!("text-{key:08}")))
            } else {
                Some(WorkloadValue::Number(key as i64))
            }
        }
        Workload::PlainText => Some(WorkloadValue::Text(format!("text-{:08}", row * cols + col))),
        Workload::StylesHeavy => Some(WorkloadValue::Number(((row + 1) * (col + 1)) as i64)),
        Workload::FormulaHeavy => {
            let base = (row + 1) as i64;
            if col == 0 {
                Some(WorkloadValue::Number(base))
            } else if col == 1 {
                Some(WorkloadValue::Number(base * 2))
            } else {
                let result = base * 3 + col as i64;
                Some(WorkloadValue::Formula {
                    formula: format!("=A{}+B{}+{}", row + 1, row + 1, col),
                    result,
                })
            }
        }
        Workload::FeatureRich => {
            if row == 0 && col == 0 {
                Some(WorkloadValue::Text("Ascend".to_string()))
            } else {
                Some(WorkloadValue::Number((row * cols + col) as i64))
            }
        }
        Workload::TableHeavy => {
            if row == 0 {
                Some(WorkloadValue::Text(format!("Column {}", col + 1)))
            } else if col % 3 == 0 {
                Some(WorkloadValue::Number(row as i64))
            } else if col % 3 == 1 {
                Some(WorkloadValue::Text(format!("item-{row}-{col}")))
            } else {
                Some(WorkloadValue::Number((row * cols + col) as i64))
            }
        }
        Workload::SparseWide => {
            if col == 0 {
                Some(WorkloadValue::Number(row as i64))
            } else if col == cols.saturating_sub(1) {
                Some(WorkloadValue::Text(format!("edge-{row}-{cols}")))
            } else if (row * 31 + col * 17) % 97 == 0 {
                Some(WorkloadValue::Number((row * cols + col) as i64))
            } else {
                None
            }
        }
        Workload::StringHeavy => {
            let key = row * cols + col;
            match col % 5 {
                0 => Some(WorkloadValue::Text(format!("sku-{key:08}"))),
                1 => Some(WorkloadValue::Text(format!("region-{}", (row % 17) + 1))),
                2 => Some(WorkloadValue::Text(format!(
                    "customer-{}-segment-{}",
                    row % 997,
                    col % 13
                ))),
                3 => Some(WorkloadValue::Text(format!(
                    "note row {row} col {col} token {}",
                    key % 104_729
                ))),
                _ if key % 2 == 0 => Some(WorkloadValue::Text(format!("status-open-{}", key % 31))),
                _ => Some(WorkloadValue::Text(format!("status-closed-{}", key % 29))),
            }
        }
    }
}

fn write_assertions(
    path: &PathBuf,
    bytes: usize,
    workload: Workload,
    rows: usize,
    cols: usize,
) -> Result<serde_json::Value, String> {
    let mut workbook = open_workbook_auto(path).map_err(|error| error.to_string())?;
    let range = workbook
        .worksheet_range("Data")
        .map_err(|error| error.to_string())?;
    let mut cell_count = 0usize;
    let mut semantic_cell_values = Vec::new();
    for (row, col, value) in range.used_cells() {
        let Some(payload) = scalar_payload(value) else {
            continue;
        };
        cell_count += 1;
        let ref_name = format!("Data!{}{}", column_name(col), row + 1);
        semantic_cell_values.push(format!("{ref_name}\t{payload}"));
    }
    let observed_hash = hash_lines(semantic_cell_values);
    let expected_hash = expected_values_hash(workload, rows, cols);
    let expected_cells = expected_cell_count(workload, rows, cols);
    Ok(json!({
        "runnerVersion": env!("CARGO_PKG_VERSION"),
        "rustXlsxWriterVersion": RUST_XLSXWRITER_VERSION,
        "workload": workload_name(workload),
        "bytes": bytes,
        "formulaCount": expected_formula_count(workload, rows, cols),
        "tablePartCount": if workload == Workload::TableHeavy { 1 } else { 0 },
        "commentPartCount": 0,
        "vmlDrawingPartCount": 0,
        "worksheetHyperlinkCount": 0,
        "worksheetDataValidationCount": 0,
        "worksheetConditionalFormattingCount": 0,
        "definedNameCount": 0,
        "sheetCount": workbook.sheet_names().len(),
        "cellCount": cell_count,
        "expectedCellCount": expected_cells,
        "cellCountMatches": cell_count == expected_cells,
        "semanticCellValuesHash": observed_hash,
        "expectedSemanticCellValuesHash": expected_hash,
        "semanticCellValuesHashMatches": observed_hash == expected_hash,
    }))
}

fn expected_formula_count(workload: Workload, rows: usize, cols: usize) -> usize {
    if workload == Workload::FormulaHeavy {
        rows * cols.saturating_sub(2)
    } else {
        0
    }
}

fn expected_cell_count(workload: Workload, rows: usize, cols: usize) -> usize {
    let mut count = 0usize;
    for row in 0..rows {
        for col in 0..cols {
            if workload_value(workload, row, col, cols).is_some() {
                count += 1;
            }
        }
    }
    count
}

fn expected_values_hash(workload: Workload, rows: usize, cols: usize) -> String {
    let mut lines = Vec::new();
    for row in 0..rows {
        for col in 0..cols {
            let Some(value) = workload_value(workload, row, col, cols) else {
                continue;
            };
            let payload = match value {
                WorkloadValue::Number(value) => format!("n:{value}"),
                WorkloadValue::Text(value) => format!("s:{value}"),
                WorkloadValue::Formula { result, .. } => format!("n:{result}"),
            };
            lines.push(format!("Data!{}{}\t{payload}", column_name(col), row + 1));
        }
    }
    hash_lines(lines)
}

fn scalar_payload(value: &Data) -> Option<String> {
    match value {
        Data::Empty => None,
        Data::Bool(value) => Some(format!("b:{}", if *value { "true" } else { "false" })),
        Data::Int(value) => Some(format!("n:{value}")),
        Data::Float(value) => Some(format!("n:{}", canonical_number(*value))),
        Data::String(value) => Some(format!("s:{value}")),
        Data::DateTime(value) => Some(format!("n:{}", canonical_number(value.as_f64()))),
        Data::DateTimeIso(value) | Data::DurationIso(value) => Some(format!("s:{value}")),
        Data::Error(value) => Some(format!("e:{value:?}")),
    }
}

fn canonical_number(value: f64) -> String {
    if value == 0.0 {
        "0".to_string()
    } else if value.fract() == 0.0 {
        format!("{}", value as i64)
    } else {
        value.to_string()
    }
}

fn hash_lines<I>(lines: I) -> String
where
    I: IntoIterator<Item = String>,
{
    let mut lines: Vec<String> = lines.into_iter().collect();
    lines.sort();
    let mut hash = Sha256::new();
    for line in lines {
        hash.update(format!("{}:", line.chars().count()).as_bytes());
        hash.update(line.as_bytes());
        hash.update(b"\n");
    }
    hex_lower(hash.finalize().as_slice())
}

fn column_name(mut col: usize) -> String {
    let mut name = String::new();
    col += 1;
    while col > 0 {
        col -= 1;
        name.insert(0, char::from(b'A' + (col % 26) as u8));
        col /= 26;
    }
    name
}

fn temp_path() -> PathBuf {
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or(0);
    env::temp_dir().join(format!(
        "ascend-rust-xlsxwriter-{}-{nanos}.xlsx",
        std::process::id()
    ))
}

fn peak_rss_bytes() -> u64 {
    let mut usage = std::mem::MaybeUninit::<libc::rusage>::uninit();
    let status = unsafe { libc::getrusage(libc::RUSAGE_SELF, usage.as_mut_ptr()) };
    if status != 0 {
        return 0;
    }
    let max_rss = unsafe { usage.assume_init().ru_maxrss };
    if cfg!(target_os = "macos") {
        max_rss as u64
    } else {
        (max_rss as u64) * 1024
    }
}

fn hex_lower(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut text = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        text.push(HEX[(byte >> 4) as usize] as char);
        text.push(HEX[(byte & 0x0f) as usize] as char);
    }
    text
}
