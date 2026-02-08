use std::fs;
use std::path::PathBuf;

fn emit_version(env_key: &str, version: &str) {
    // Expose resolved dependency versions to the Rust code.
    println!("cargo:rustc-env={env_key}={version}");
}

fn parse_lock_version(lock: &str, name: &str) -> Option<String> {
    // Very small Cargo.lock parser: find a [[package]] entry with matching name.
    // This is stable enough for our "what version did we build against?" reporting.
    let needle = format!("name = \"{}\"", name);
    let mut in_pkg = false;
    let mut saw_name = false;
    for line in lock.lines() {
        let line = line.trim();
        if line == "[[package]]" {
            in_pkg = true;
            saw_name = false;
            continue;
        }
        if !in_pkg {
            continue;
        }
        if line.starts_with("name = ") {
            saw_name = line == needle;
        }
        if saw_name && line.starts_with("version = ") {
            // version = "0.25.0"
            if let Some(v) = line.split('"').nth(1) {
                return Some(v.to_string());
            }
        }
    }
    None
}

fn main() {
    println!("cargo:rerun-if-changed=Cargo.toml");
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=Cargo.lock");

    let manifest_dir = std::env::var("CARGO_MANIFEST_DIR").ok();
    let Some(manifest_dir) = manifest_dir else {
        return;
    };
    let lock_path = PathBuf::from(manifest_dir).join("Cargo.lock");
    let Ok(lock) = fs::read_to_string(&lock_path) else {
        return;
    };

    if let Some(v) = parse_lock_version(&lock, "calamine") {
        emit_version("EXCELBENCH_DEP_CALAMINE_VERSION", &v);
    }
    if let Some(v) = parse_lock_version(&lock, "rust_xlsxwriter") {
        emit_version("EXCELBENCH_DEP_RUST_XLSXWRITER_VERSION", &v);
    }
    if let Some(v) = parse_lock_version(&lock, "umya-spreadsheet") {
        emit_version("EXCELBENCH_DEP_UMYA_VERSION", &v);
    }
}
