use rootcellar_core::model::CellRef;
use rootcellar_core::{load_workbook_model, CellValue, NoopEventSink, TraceContext};
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};
use std::{env, io};

fn workspace_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .expect("crate should live under <repo>/crates/rootcellar-cli")
        .to_path_buf()
}

fn temp_root(test_name: &str) -> PathBuf {
    let suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|clock| clock.as_nanos())
        .unwrap_or(0);
    let path = env::temp_dir().join(format!("rootcellar-macro-{test_name}-{suffix}"));
    fs::create_dir_all(&path).expect("create temp test dir");
    path
}

fn read_macro_result(path: &Path) -> io::Result<BTreeMap<(u32, u32), (CellValue, Option<String>)>> {
    let mut sink = NoopEventSink;
    let workbook =
        load_workbook_model(path, &mut sink, &TraceContext::root()).map_err(io::Error::other)?;
    let sheet = workbook
        .sheets
        .get("Sheet1")
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "Sheet1 not found"))?;

    let mut cells = BTreeMap::new();
    for (CellRef { row, col }, cell) in &sheet.cells {
        cells.insert((*row, *col), (cell.value.clone(), cell.formula.clone()));
    }
    Ok(cells)
}

fn run_macro_command(args: &[&str], extra_env: &[(&str, &str)]) -> std::process::Output {
    let mut command = Command::new("cargo");
    let mut command_args = vec!["run", "-q", "-p", "rootcellar-cli", "--", "run-macro"];
    command_args.extend_from_slice(args);
    command.args(&command_args);
    command.current_dir(workspace_root());
    command.env("ROOTCELLAR_PYTHON", "python");
    for (key, value) in extra_env {
        command.env(key, value);
    }
    command.output().expect("run-macro command should spawn")
}

#[test]
fn run_macro_cli_applies_mutations_and_recalculates_cells() {
    if Command::new("python").arg("-V").output().is_err() {
        eprintln!("python missing; skipping run-macro integration test");
        return;
    }

    let root = workspace_root();
    let temp = temp_root("success");
    let input = temp.join("input.xlsx");
    let output = temp.join("output.xlsx");
    fs::copy(root.join("normalized.xlsx"), &input).expect("copy template workbook");

    let macro_script = temp.join("macro_user.py");
    fs::write(
        &macro_script,
        "def run_macro(ctx, args):\n    base = int(args.get('base', '11'))\n    ctx.set_value('Sheet1', 'A1', base)\n    ctx.set_formula('Sheet1', 'B1', '=A1+1')\n    ctx.set_range_values('Sheet1', 'C1:D2', base + 1)\n    return {'base': base}\n",
    )
    .expect("write macro script");

    let result = run_macro_command(
        &[
            "--macro-script",
            macro_script.to_str().expect("macro script path"),
            "--macro-name",
            "run_macro",
            "--allow",
            "fs.write",
            "--arg",
            "base=9",
            input.to_str().expect("input path"),
            output.to_str().expect("output path"),
        ],
        &[],
    );

    assert!(
        result.status.success(),
        "macro run should succeed: {}",
        String::from_utf8_lossy(&result.stderr)
    );

    let cells = read_macro_result(&output).expect("read output workbook");
    assert_eq!(
        cells.get(&(1, 1)).map(|(value, _)| value.clone()),
        Some(CellValue::Number(9.0))
    );
    assert_eq!(
        cells.get(&(1, 2)).map(|(_, formula)| formula.as_deref()),
        Some(Some("=A1+1"))
    );
    assert_eq!(
        cells.get(&(1, 3)).map(|(value, _)| value.clone()),
        Some(CellValue::Number(10.0))
    );
    assert_eq!(
        cells.get(&(2, 4)).map(|(value, _)| value.clone()),
        Some(CellValue::Number(10.0))
    );

    fs::remove_dir_all(&temp).expect("cleanup macro integration temp dir");
}

#[test]
fn run_macro_cli_denies_mutation_without_permission() {
    if Command::new("python").arg("-V").output().is_err() {
        eprintln!("python missing; skipping run-macro denial integration test");
        return;
    }

    let root = workspace_root();
    let temp = temp_root("denied");
    let input = temp.join("input.xlsx");
    let output = temp.join("output.xlsx");
    fs::copy(root.join("normalized.xlsx"), &input).expect("copy template workbook");

    let macro_script = temp.join("macro_user.py");
    fs::write(
        &macro_script,
        "def run_macro(ctx, args):\n    ctx.set_value('Sheet1', 'A1', 123)\n",
    )
    .expect("write macro script");

    let result = run_macro_command(
        &[
            "--macro-script",
            macro_script.to_str().expect("macro script path"),
            "--macro-name",
            "run_macro",
            "--allow",
            "fs.read",
            input.to_str().expect("input path"),
            output.to_str().expect("output path"),
        ],
        &[],
    );

    assert!(!result.status.success());
    let stderr = String::from_utf8_lossy(&result.stderr);
    assert!(stderr.contains("macro runtime error"));

    fs::remove_dir_all(&temp).expect("cleanup macro integration temp dir");
}

#[test]
fn run_macro_cli_fails_with_invalid_worker_response() {
    if Command::new("python").arg("-V").output().is_err() {
        eprintln!("python missing; skipping malformed-worker-response integration test");
        return;
    }

    let temp = temp_root("invalid-worker");
    let root = workspace_root();
    let input = temp.join("input.xlsx");
    fs::copy(root.join("normalized.xlsx"), &input).expect("copy template workbook");

    let output = temp.join("output.xlsx");
    let macro_script = temp.join("macro_user.py");
    fs::write(
        &macro_script,
        "def run_macro(ctx, args):\n    return {'ok': True}\n",
    )
    .expect("write macro script");

    let bad_worker = temp.join("bad_worker.py");
    fs::write(&bad_worker, "print(\"not-json\")\n").expect("write invalid worker script");

    let result = run_macro_command(
        &[
            "--macro-script",
            macro_script.to_str().expect("macro script path"),
            "--macro-name",
            "run_macro",
            "--allow",
            "fs.write",
            input.to_str().expect("input path"),
            output.to_str().expect("output path"),
        ],
        &[(
            "ROOTCELLAR_SCRIPT_WORKER",
            bad_worker.to_str().expect("bad worker path"),
        )],
    );

    assert!(!result.status.success());
    let stderr = String::from_utf8_lossy(&result.stderr);
    assert!(stderr.contains("invalid worker response"));

    fs::remove_dir_all(&temp).expect("cleanup macro integration temp dir");
}
