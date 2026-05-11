use calamine::{open_workbook_auto, Data, Reader};
use serde_json::{json, Map, Value};
use sha2::{Digest, Sha256};
use std::env;
use std::path::PathBuf;
use std::time::Instant;

const CALAMINE_VERSION: &str = "0.34.0";

struct Args {
    operation: String,
    file: PathBuf,
    ordered_hashes: bool,
    repeat: usize,
    warmup: usize,
    pretty_json: bool,
}

struct WorkbookAssertions {
    sheet_count: usize,
    sheet_names_hash: String,
    cell_count: usize,
    physical_cell_count: usize,
    used_ranges: Vec<String>,
    semantic_cell_refs_hash: String,
    semantic_cell_values_hash: String,
    ordered_semantic_cell_refs_hash: String,
    ordered_semantic_cell_values_hash: String,
}

fn main() {
    if let Err(error) = run() {
        eprintln!("{error}");
        std::process::exit(1);
    }
}

fn run() -> Result<(), String> {
    let args = parse_args()?;
    if args.operation != "read" {
        return Err(format!(
            "Unsupported --operation {:?}. Expected read.",
            args.operation
        ));
    }

    for _ in 0..args.warmup {
        read_assertions(&args.file, args.ordered_hashes)?;
    }

    let mut samples = Vec::with_capacity(args.repeat);
    let mut assertions = None;
    for _ in 0..args.repeat {
        let start = Instant::now();
        let read = read_assertions(&args.file, args.ordered_hashes)?;
        let duration_ms = start.elapsed().as_secs_f64() * 1000.0;
        assertions = Some(read);
        let peak_rss_bytes = peak_rss_bytes();
        samples.push(json!({
            "durationMs": duration_ms,
            "peakRssBytes": peak_rss_bytes,
            "rssAfterBytes": peak_rss_bytes
        }));
    }

    let payload = json!({
        "assertions": assertions_to_json(assertions.ok_or("No samples were produced")?),
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
    let mut file = None;
    let mut repeat = 1usize;
    let mut warmup = 0usize;
    let mut ordered_hashes = false;
    let mut pretty_json = true;
    let mut args = env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--operation" => operation = args.next(),
            "--file" => file = args.next().map(PathBuf::from),
            "--repeat" => repeat = parse_count("--repeat", args.next())?,
            "--warmup" => warmup = parse_count("--warmup", args.next())?,
            "--ordered-hashes" => ordered_hashes = true,
            "--json" => pretty_json = false,
            other => return Err(format!("Unsupported argument {other:?}")),
        }
    }
    Ok(Args {
        operation: operation.ok_or("Missing --operation")?,
        file: file.ok_or("Missing --file")?,
        ordered_hashes,
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

fn read_assertions(path: &PathBuf, ordered_hashes: bool) -> Result<WorkbookAssertions, String> {
    let mut workbook = open_workbook_auto(path).map_err(|error| error.to_string())?;
    let sheet_names = workbook.sheet_names().to_vec();
    let mut cell_count = 0usize;
    let mut physical_cell_count = 0usize;
    let mut used_ranges = Vec::with_capacity(sheet_names.len());
    let mut semantic_cell_refs = Vec::new();
    let mut semantic_cell_values = Vec::new();
    let mut ordered_semantic_cell_refs = OrderedLineHasher::new();
    let mut ordered_semantic_cell_values = OrderedLineHasher::new();

    for sheet_name in &sheet_names {
        let range = workbook
            .worksheet_range(sheet_name)
            .map_err(|error| error.to_string())?;
        let (start_row, start_col) = range.start().unwrap_or((0, 0));
        let start_row = start_row as usize;
        let start_col = start_col as usize;
        let mut min_row = usize::MAX;
        let mut min_col = usize::MAX;
        let mut max_row = 0usize;
        let mut max_col = 0usize;

        for (row, col, value) in range.used_cells() {
            physical_cell_count += 1;
            let Some(payload) = scalar_payload(value) else {
                continue;
            };
            let absolute_row = start_row + row;
            let absolute_col = start_col + col;
            min_row = min_row.min(absolute_row);
            min_col = min_col.min(absolute_col);
            max_row = max_row.max(absolute_row);
            max_col = max_col.max(absolute_col);
            cell_count += 1;
            let ref_name = format!(
                "{}!{}{}",
                sheet_name,
                column_name(absolute_col),
                absolute_row + 1
            );
            ordered_semantic_cell_refs.update(&ref_name);
            ordered_semantic_cell_values.update(&format!("{ref_name}\t{payload}"));
            if !ordered_hashes {
                semantic_cell_refs.push(ref_name.clone());
                semantic_cell_values.push(format!("{ref_name}\t{payload}"));
            }
        }

        used_ranges.push(if min_row == usize::MAX {
            format!("{sheet_name}!empty")
        } else {
            format!(
                "{}!{}{}:{}{}",
                sheet_name,
                column_name(min_col),
                min_row + 1,
                column_name(max_col),
                max_row + 1
            )
        });
    }

    Ok(WorkbookAssertions {
        sheet_count: sheet_names.len(),
        sheet_names_hash: hash_lines(
            sheet_names
                .iter()
                .enumerate()
                .map(|(index, name)| format!("{index}:{name}")),
        ),
        cell_count,
        physical_cell_count,
        used_ranges: used_ranges.clone(),
        semantic_cell_refs_hash: hash_lines(semantic_cell_refs),
        semantic_cell_values_hash: hash_lines(semantic_cell_values),
        ordered_semantic_cell_refs_hash: ordered_semantic_cell_refs.finish(),
        ordered_semantic_cell_values_hash: ordered_semantic_cell_values.finish(),
    })
}

fn assertions_to_json(assertions: WorkbookAssertions) -> Value {
    let first_used_range = assertions
        .used_ranges
        .first()
        .map(|value| Value::String(value.clone()))
        .unwrap_or(Value::Null);
    let used_range_count = assertions.used_ranges.len();
    let used_ranges_hash = hash_lines(assertions.used_ranges);
    let mut object = Map::new();
    object.insert(
        "runnerVersion".to_string(),
        json!(env!("CARGO_PKG_VERSION")),
    );
    object.insert("calamineVersion".to_string(), json!(CALAMINE_VERSION));
    object.insert("sheetCount".to_string(), json!(assertions.sheet_count));
    object.insert(
        "sheetNamesHash".to_string(),
        json!(assertions.sheet_names_hash),
    );
    object.insert("cellCount".to_string(), json!(assertions.cell_count));
    object.insert(
        "physicalCellCount".to_string(),
        json!(assertions.physical_cell_count),
    );
    object.insert("formulaCount".to_string(), json!(0));
    object.insert("usedRangeCount".to_string(), json!(used_range_count));
    object.insert("firstUsedRange".to_string(), first_used_range.clone());
    object.insert("firstPhysicalUsedRange".to_string(), first_used_range);
    object.insert(
        "usedRangesHash".to_string(),
        json!(used_ranges_hash.clone()),
    );
    object.insert(
        "physicalUsedRangesHash".to_string(),
        json!(used_ranges_hash),
    );
    object.insert(
        "semanticCellRefsHash".to_string(),
        json!(assertions.semantic_cell_refs_hash),
    );
    object.insert(
        "semanticCellValuesHash".to_string(),
        json!(assertions.semantic_cell_values_hash),
    );
    object.insert(
        "orderedSemanticCellRefsHash".to_string(),
        json!(assertions.ordered_semantic_cell_refs_hash),
    );
    object.insert(
        "orderedSemanticCellValuesHash".to_string(),
        json!(assertions.ordered_semantic_cell_values_hash),
    );
    object.insert(
        "formulaTextHash".to_string(),
        json!(hash_lines(Vec::<String>::new())),
    );
    object.insert(
        "orderedFormulaTextHash".to_string(),
        json!(OrderedLineHasher::new().finish()),
    );
    object.insert("readCommentCount".to_string(), json!(0));
    object.insert("readHyperlinkCount".to_string(), json!(0));
    object.insert("readDataValidationCount".to_string(), json!(0));
    object.insert("readConditionalFormatCount".to_string(), json!(0));
    object.insert("readDefinedNameCount".to_string(), json!(0));
    Value::Object(object)
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
    } else {
        value.to_string()
    }
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

fn hash_lines<I>(lines: I) -> String
where
    I: IntoIterator<Item = String>,
{
    let mut lines: Vec<String> = lines.into_iter().collect();
    lines.sort();
    let mut hash = Sha256::new();
    for line in lines {
        hash.update(format!("{}:", line.encode_utf16().count()).as_bytes());
        hash.update(line.as_bytes());
        hash.update(b"\n");
    }
    hex_lower(hash.finalize().as_slice())
}

struct OrderedLineHasher {
    hash: Sha256,
}

impl OrderedLineHasher {
    fn new() -> Self {
        Self {
            hash: Sha256::new(),
        }
    }

    fn update(&mut self, line: &str) {
        self.hash
            .update(format!("{}:", line.encode_utf16().count()).as_bytes());
        self.hash.update(line.as_bytes());
        self.hash.update(b"\n");
    }

    fn finish(self) -> String {
        hex_lower(self.hash.finalize().as_slice())
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
