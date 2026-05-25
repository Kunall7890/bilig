export function ironCalcRustSidecarCargoToml(): string {
  return `[package]
name = "bilig-ironcalc-rust-bench"
version = "0.1.0"
edition = "2021"
publish = false

[dependencies]
ironcalc_base = "=0.7.1"
serde = { version = "1", features = ["derive"] }
serde_json = "1"
`
}

export function ironCalcRustSidecarMainRs(): string {
  return String.raw`use ironcalc_base::{cell::CellValue, Model, UserModel};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::BTreeMap;
use std::env;
use std::fs;
use std::time::Instant;

#[derive(Deserialize)]
struct RunnerInput {
    benchmark: BenchmarkInput,
    workloads: Vec<WorkloadInput>,
}

#[derive(Deserialize)]
struct BenchmarkInput {
    #[serde(rename = "sampleCount")]
    sample_count: usize,
    #[serde(rename = "warmupCount")]
    warmup_count: usize,
}

#[derive(Deserialize)]
struct WorkloadInput {
    workload: String,
    sheets: Vec<SheetInput>,
    operation: OperationInput,
    observations: Vec<ObservationInput>,
}

#[derive(Deserialize)]
#[serde(tag = "kind")]
enum OperationInput {
    #[serde(rename = "build")]
    Build,
    #[serde(rename = "cell-edits")]
    CellEdits {
        edits: Vec<CellEditInput>,
        evaluation: EvaluationMode,
    },
    #[serde(rename = "single-cell-edit")]
    SingleCellEdit { edit: CellEditInput },
    #[serde(rename = "insert-rows")]
    InsertRows {
        #[serde(rename = "sheetName")]
        sheet_name: String,
        index: usize,
    },
    #[serde(rename = "delete-rows")]
    DeleteRows {
        #[serde(rename = "sheetName")]
        sheet_name: String,
        index: usize,
    },
    #[serde(rename = "move-rows")]
    MoveRows {
        #[serde(rename = "sheetName")]
        sheet_name: String,
        index: usize,
        delta: i32,
    },
    #[serde(rename = "insert-columns")]
    InsertColumns {
        #[serde(rename = "sheetName")]
        sheet_name: String,
        index: usize,
    },
    #[serde(rename = "delete-columns")]
    DeleteColumns {
        #[serde(rename = "sheetName")]
        sheet_name: String,
        index: usize,
    },
    #[serde(rename = "move-columns")]
    MoveColumns {
        #[serde(rename = "sheetName")]
        sheet_name: String,
        index: usize,
        delta: i32,
    },
    #[serde(rename = "rename-sheet")]
    RenameSheet {
        #[serde(rename = "oldName")]
        old_name: String,
        #[serde(rename = "newName")]
        new_name: String,
    },
    #[serde(rename = "range-read")]
    RangeRead {
        #[serde(rename = "sheetName")]
        sheet_name: String,
        #[serde(rename = "startRow")]
        start_row: usize,
        #[serde(rename = "startCol")]
        start_col: usize,
        #[serde(rename = "endRow")]
        end_row: usize,
        #[serde(rename = "endCol")]
        end_col: usize,
        summary: RangeReadSummary,
        #[serde(rename = "middleCol")]
        middle_col: Option<usize>,
    },
}

#[derive(Deserialize)]
enum RangeReadSummary {
    #[serde(rename = "dense")]
    Dense,
    #[serde(rename = "sparse-wide")]
    SparseWide,
    #[serde(rename = "formula-grid")]
    FormulaGrid,
}

#[derive(Deserialize)]
enum EvaluationMode {
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "paused")]
    Paused,
}

#[derive(Deserialize)]
struct SheetInput {
    name: String,
    cells: Vec<Vec<Value>>,
}

#[derive(Deserialize)]
struct CellEditInput {
    #[serde(rename = "sheetName")]
    sheet_name: String,
    row: usize,
    col: usize,
    value: Value,
}

#[derive(Deserialize)]
struct ObservationInput {
    key: String,
    #[serde(rename = "sheetName")]
    sheet_name: String,
    row: usize,
    col: usize,
}

#[derive(Serialize)]
struct RunnerOutput {
    engine: EngineOutput,
    results: Vec<WorkloadOutput>,
}

#[derive(Serialize)]
struct EngineOutput {
    #[serde(rename = "crate")]
    crate_name: &'static str,
    version: &'static str,
}

#[derive(Serialize)]
struct WorkloadOutput {
    workload: String,
    #[serde(rename = "apiPath")]
    api_path: &'static str,
    #[serde(rename = "elapsedMs")]
    elapsed_ms: Vec<f64>,
    verification: BTreeMap<String, Value>,
}

fn main() -> Result<(), String> {
    let args: Vec<String> = env::args().collect();
    if args.len() != 3 {
        return Err("usage: bilig-ironcalc-rust-bench <input.json> <output.json>".to_string());
    }
    let input_text = fs::read_to_string(&args[1]).map_err(|error| format!("failed to read input: {error}"))?;
    let input: RunnerInput = serde_json::from_str(&input_text).map_err(|error| format!("failed to parse input: {error}"))?;
    let mut results = Vec::with_capacity(input.workloads.len());
    for workload in &input.workloads {
        results.push(run_workload(workload, input.benchmark.warmup_count, input.benchmark.sample_count)?);
    }
    let output = RunnerOutput {
        engine: EngineOutput {
            crate_name: "ironcalc_base",
            version: "0.7.1",
        },
        results,
    };
    let output_text = serde_json::to_string_pretty(&output).map_err(|error| format!("failed to serialize output: {error}"))?;
    fs::write(&args[2], format!("{output_text}\n")).map_err(|error| format!("failed to write output: {error}"))?;
    Ok(())
}

fn run_workload(workload: &WorkloadInput, warmup_count: usize, sample_count: usize) -> Result<WorkloadOutput, String> {
    for _ in 0..warmup_count {
        let _ = run_sample(workload).map_err(|error| format!("{}: {error}", workload.workload))?;
    }
    let mut elapsed_ms = Vec::with_capacity(sample_count);
    let mut verification: Option<BTreeMap<String, Value>> = None;
    let mut api_path: Option<&'static str> = None;
    for _ in 0..sample_count {
        let sample = run_sample(workload).map_err(|error| format!("{}: {error}", workload.workload))?;
        if let Some(expected) = &verification {
            if expected != &sample.verification {
                return Err(format!(
                    "verification drifted for {}: expected {}, got {}",
                    workload.workload,
                    serde_json::to_string(expected).unwrap_or_default(),
                    serde_json::to_string(&sample.verification).unwrap_or_default()
                ));
            }
        } else {
            verification = Some(sample.verification.clone());
            api_path = Some(sample.api_path);
        }
        elapsed_ms.push(sample.elapsed_ms);
    }
    Ok(WorkloadOutput {
        workload: workload.workload.clone(),
        api_path: api_path.unwrap_or("Model"),
        elapsed_ms,
        verification: verification.unwrap_or_default(),
    })
}

#[derive(Clone)]
struct SampleOutput {
    api_path: &'static str,
    elapsed_ms: f64,
    verification: BTreeMap<String, Value>,
}

fn run_sample(workload: &WorkloadInput) -> Result<SampleOutput, String> {
    match &workload.operation {
        OperationInput::Build => {
            let started = Instant::now();
            let model = build_model(&workload.sheets)?;
            let elapsed_ms = started.elapsed().as_secs_f64() * 1000.0;
            Ok(SampleOutput {
                api_path: "Model",
                elapsed_ms,
                verification: observe_model(&model, &workload.sheets, &workload.observations)?,
            })
        }
        OperationInput::SingleCellEdit { edit } => {
            let model = build_model(&workload.sheets)?;
            let mut user_model = UserModel::from_model(model);
            let edit_sheet_index = sheet_index(&workload.sheets, &edit.sheet_name)?;
            let input = cell_input(&edit.value)?.unwrap_or_default();
            let started = Instant::now();
            user_model.set_user_input(edit_sheet_index, to_ironcalc_index(edit.row), to_ironcalc_index(edit.col), &input)?;
            let elapsed_ms = started.elapsed().as_secs_f64() * 1000.0;
            Ok(SampleOutput {
                api_path: "UserModel",
                elapsed_ms,
                verification: observe_model(user_model.get_model(), &workload.sheets, &workload.observations)?,
            })
        }
        OperationInput::CellEdits { edits, evaluation } => {
            let model = build_model(&workload.sheets)?;
            let mut user_model = UserModel::from_model(model);
            let started = Instant::now();
            if matches!(evaluation, EvaluationMode::Paused) {
                user_model.pause_evaluation();
            }
            for edit in edits {
                let sheet = sheet_index(&workload.sheets, &edit.sheet_name)?;
                let input = cell_input(&edit.value)?.unwrap_or_default();
                user_model.set_user_input(sheet, to_ironcalc_index(edit.row), to_ironcalc_index(edit.col), &input)?;
            }
            if matches!(evaluation, EvaluationMode::Paused) {
                user_model.resume_evaluation();
                user_model.evaluate();
            }
            let elapsed_ms = started.elapsed().as_secs_f64() * 1000.0;
            Ok(SampleOutput {
                api_path: "UserModel",
                elapsed_ms,
                verification: observe_model(user_model.get_model(), &workload.sheets, &workload.observations)?,
            })
        }
        OperationInput::InsertRows { sheet_name, index } => {
            run_user_model_operation(workload, |user_model| {
                let sheet = sheet_index(&workload.sheets, sheet_name)?;
                user_model.insert_rows(sheet, to_ironcalc_index(*index), 1)
            })
        }
        OperationInput::DeleteRows { sheet_name, index } => {
            run_user_model_operation(workload, |user_model| {
                let sheet = sheet_index(&workload.sheets, sheet_name)?;
                user_model.delete_rows(sheet, to_ironcalc_index(*index), 1)
            })
        }
        OperationInput::MoveRows { sheet_name, index, delta } => {
            run_user_model_operation(workload, |user_model| {
                let sheet = sheet_index(&workload.sheets, sheet_name)?;
                user_model.move_row_action(sheet, to_ironcalc_index(*index), *delta)
            })
        }
        OperationInput::InsertColumns { sheet_name, index } => {
            run_user_model_operation(workload, |user_model| {
                let sheet = sheet_index(&workload.sheets, sheet_name)?;
                user_model.insert_columns(sheet, to_ironcalc_index(*index), 1)
            })
        }
        OperationInput::DeleteColumns { sheet_name, index } => {
            run_user_model_operation(workload, |user_model| {
                let sheet = sheet_index(&workload.sheets, sheet_name)?;
                user_model.delete_columns(sheet, to_ironcalc_index(*index), 1)
            })
        }
        OperationInput::MoveColumns { sheet_name, index, delta } => {
            run_user_model_operation(workload, |user_model| {
                let sheet = sheet_index(&workload.sheets, sheet_name)?;
                user_model.move_column_action(sheet, to_ironcalc_index(*index), *delta)
            })
        }
        OperationInput::RenameSheet { old_name, new_name } => {
            run_user_model_operation(workload, |user_model| {
                let sheet = sheet_index(&workload.sheets, old_name)?;
                user_model.rename_sheet(sheet, new_name)
            })
        }
        OperationInput::RangeRead {
            sheet_name,
            start_row,
            start_col,
            end_row,
            end_col,
            summary,
            middle_col,
        } => {
            let model = build_model(&workload.sheets)?;
            let sheet = sheet_index(&workload.sheets, sheet_name)?;
            let started = Instant::now();
            let values = read_range_values(&model, sheet, *start_row, *start_col, *end_row, *end_col)?;
            let verification = summarize_range_read(&values, summary, *middle_col)?;
            let elapsed_ms = started.elapsed().as_secs_f64() * 1000.0;
            Ok(SampleOutput {
                api_path: "Model",
                elapsed_ms,
                verification,
            })
        }
    }
}

fn run_user_model_operation<F>(workload: &WorkloadInput, execute: F) -> Result<SampleOutput, String>
where
    F: FnOnce(&mut UserModel<'_>) -> Result<(), String>,
{
    let model = build_model(&workload.sheets)?;
    let mut user_model = UserModel::from_model(model);
    let started = Instant::now();
    execute(&mut user_model)?;
    let elapsed_ms = started.elapsed().as_secs_f64() * 1000.0;
    Ok(SampleOutput {
        api_path: "UserModel",
        elapsed_ms,
        verification: observe_model(user_model.get_model(), &workload.sheets, &workload.observations)?,
    })
}

fn build_model(sheets: &[SheetInput]) -> Result<Model<'static>, String> {
    let first_sheet = sheets.first().ok_or_else(|| "benchmark workload has no sheets".to_string())?;
    let mut model = Model::new_empty("bilig-ironcalc-rust-bench", "en", "UTC", "en")?;
    model.rename_sheet_by_index(0, &first_sheet.name)?;
    for (sheet_index, sheet) in sheets.iter().enumerate().skip(1) {
        model.insert_sheet(&sheet.name, sheet_index as u32, None)?;
    }
    for (sheet_index, sheet) in sheets.iter().enumerate() {
        for (row_index, row) in sheet.cells.iter().enumerate() {
            for (col_index, cell) in row.iter().enumerate() {
                if let Some(input) = cell_input(cell)? {
                    model.set_user_input(sheet_index as u32, to_ironcalc_index(row_index), to_ironcalc_index(col_index), input)?;
                }
            }
        }
    }
    model.evaluate();
    Ok(model)
}

fn observe_model(
    model: &Model<'_>,
    sheets: &[SheetInput],
    observations: &[ObservationInput],
) -> Result<BTreeMap<String, Value>, String> {
    let mut verification = BTreeMap::new();
    for observation in observations {
        let sheet = sheet_index(sheets, &observation.sheet_name)?;
        let value = model.get_cell_value_by_index(
            sheet,
            to_ironcalc_index(observation.row),
            to_ironcalc_index(observation.col),
        )?;
        verification.insert(observation.key.clone(), normalize_cell_value(value)?);
    }
    Ok(verification)
}

fn read_range_values(
    model: &Model<'_>,
    sheet: u32,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
) -> Result<Vec<Vec<Value>>, String> {
    let mut values = Vec::with_capacity(end_row.saturating_sub(start_row) + 1);
    for row in start_row..=end_row {
        let mut row_values = Vec::with_capacity(end_col.saturating_sub(start_col) + 1);
        for col in start_col..=end_col {
            row_values.push(normalize_cell_value(model.get_cell_value_by_index(
                sheet,
                to_ironcalc_index(row),
                to_ironcalc_index(col),
            )?)?);
        }
        values.push(row_values);
    }
    Ok(values)
}

fn summarize_range_read(
    values: &[Vec<Value>],
    summary: &RangeReadSummary,
    middle_col: Option<usize>,
) -> Result<BTreeMap<String, Value>, String> {
    match summary {
        RangeReadSummary::Dense => summarize_dense_range_read(values),
        RangeReadSummary::SparseWide => summarize_sparse_wide_range_read(values, middle_col),
        RangeReadSummary::FormulaGrid => summarize_formula_grid_range_read(values),
    }
}

fn summarize_dense_range_read(values: &[Vec<Value>]) -> Result<BTreeMap<String, Value>, String> {
    let mut summary = BTreeMap::new();
    summary.insert("readCols".to_string(), json_number(values.first().map_or(0, Vec::len) as f64)?);
    summary.insert("readRows".to_string(), json_number(values.len() as f64)?);
    summary.insert("terminalValue".to_string(), range_value(values, values.len().saturating_sub(1), values.first().map_or(0, Vec::len).saturating_sub(1)));
    summary.insert("topLeftValue".to_string(), range_value(values, 0, 0));
    Ok(summary)
}

fn summarize_sparse_wide_range_read(values: &[Vec<Value>], middle_col: Option<usize>) -> Result<BTreeMap<String, Value>, String> {
    let middle_col = middle_col.ok_or_else(|| "sparse-wide range read requires middleCol".to_string())?;
    let mut summary = BTreeMap::new();
    let empty_value = range_value(values, 0, 1);
    summary.insert(
        "emptyValue".to_string(),
        if empty_value.is_null() { Value::String(String::new()) } else { empty_value },
    );
    summary.insert("middleValue".to_string(), range_value(values, values.len().saturating_sub(1), middle_col));
    summary.insert("readCols".to_string(), json_number(values.first().map_or(0, Vec::len) as f64)?);
    summary.insert("readRows".to_string(), json_number(values.len() as f64)?);
    summary.insert("terminalValue".to_string(), range_value(values, values.len().saturating_sub(1), values.first().map_or(0, Vec::len).saturating_sub(1)));
    summary.insert("topLeftValue".to_string(), range_value(values, 0, 0));
    Ok(summary)
}

fn summarize_formula_grid_range_read(values: &[Vec<Value>]) -> Result<BTreeMap<String, Value>, String> {
    let mut summary = BTreeMap::new();
    summary.insert("leadingFormulaValue".to_string(), range_value(values, 0, 0));
    summary.insert("readCols".to_string(), json_number(values.first().map_or(0, Vec::len) as f64)?);
    summary.insert("readRows".to_string(), json_number(values.len() as f64)?);
    summary.insert("terminalFormulaValue".to_string(), range_value(values, values.len().saturating_sub(1), values.first().map_or(0, Vec::len).saturating_sub(1)));
    Ok(summary)
}

fn range_value(values: &[Vec<Value>], row: usize, col: usize) -> Value {
    values.get(row).and_then(|row_values| row_values.get(col)).cloned().unwrap_or(Value::Null)
}

fn json_number(value: f64) -> Result<Value, String> {
    serde_json::Number::from_f64(value)
        .map(Value::Number)
        .ok_or_else(|| format!("non-finite range summary number: {value}"))
}

fn cell_input(value: &Value) -> Result<Option<String>, String> {
    match value {
        Value::Null => Ok(None),
        Value::Bool(value) => Ok(Some(if *value { "TRUE".to_string() } else { "FALSE".to_string() })),
        Value::Number(value) => Ok(Some(value.to_string())),
        Value::String(value) => Ok(Some(value.clone())),
        Value::Array(_) | Value::Object(_) => Err(format!("unsupported cell input value: {value}")),
    }
}

fn normalize_cell_value(value: CellValue) -> Result<Value, String> {
    match value {
        CellValue::None => Ok(Value::Null),
        CellValue::Boolean(value) => Ok(Value::Bool(value)),
        CellValue::Number(value) => serde_json::Number::from_f64(value)
            .map(Value::Number)
            .ok_or_else(|| format!("non-finite IronCalc number: {value}")),
        CellValue::String(value) => Ok(Value::String(value)),
    }
}

fn sheet_index(sheets: &[SheetInput], sheet_name: &str) -> Result<u32, String> {
    sheets
        .iter()
        .position(|sheet| sheet.name == sheet_name)
        .map(|index| index as u32)
        .ok_or_else(|| format!("sheet not found: {sheet_name}"))
}

fn to_ironcalc_index(zero_based_index: usize) -> i32 {
    (zero_based_index + 1) as i32
}
`
}
