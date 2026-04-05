mod dicom;
mod patient;
mod progress;
mod types;
mod zip_handler;

use clap::Parser;
use rayon::prelude::*;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;
use types::{OutputPayload, Stats};

#[derive(Parser)]
#[command(name = "study-agg-engine", about = "DICOM study aggregation engine")]
struct Args {
    /// Path to directory or ZIP file to process
    path: PathBuf,

    /// Base64-encoded password for encrypted ZIPs
    #[arg(long)]
    password: Option<String>,

    /// Maximum number of worker threads
    #[arg(long, default_value_t = 0)]
    max_workers: usize,

    /// Maximum nested ZIP depth
    #[arg(long, default_value_t = 5)]
    max_nested_zips: u32,
}

fn main() {
    let args = Args::parse();
    let start = Instant::now();

    // Configure rayon thread pool
    let num_workers = if args.max_workers > 0 {
        args.max_workers
    } else {
        std::thread::available_parallelism()
            .map(|n| n.get())
            .unwrap_or(4)
    };
    rayon::ThreadPoolBuilder::new()
        .num_threads(num_workers)
        .build_global()
        .ok();

    // Decode password
    let password: Option<Vec<u8>> = args.password.as_ref().and_then(|b64| {
        use base64::Engine;
        base64::engine::general_purpose::STANDARD
            .decode(b64)
            .ok()
    });

    let input_path = &args.path;
    if !input_path.exists() {
        progress::error(format!("Path does not exist: {}", input_path.display()));
        std::process::exit(1);
    }

    let seven_zip = zip_handler::find_7zip();
    let files_scanned = AtomicU64::new(0);
    let dicom_valid = AtomicU64::new(0);

    let all_studies = if input_path.is_file() && has_zip_extension(input_path) {
        process_single_zip(
            input_path,
            password.as_deref(),
            seven_zip.as_deref(),
            args.max_nested_zips,
            &files_scanned,
            &dicom_valid,
        )
    } else if input_path.is_dir() {
        process_directory(
            input_path,
            password.as_deref(),
            seven_zip.as_deref(),
            args.max_nested_zips,
            &files_scanned,
            &dicom_valid,
        )
    } else {
        progress::error(format!(
            "Invalid input: {} is not a directory or ZIP file",
            input_path.display()
        ));
        std::process::exit(1);
    };

    if all_studies.is_empty() {
        progress::error("No DICOM studies found");
        std::process::exit(1);
    }

    progress::progress(95, "Merging patient records...");
    let patients = patient::merge_patients(all_studies);

    if patients.is_empty() {
        progress::error("No patients with identifiable information found");
        std::process::exit(1);
    }

    let elapsed = start.elapsed();
    let payload = OutputPayload {
        stats: Stats {
            files_scanned: files_scanned.load(Ordering::Relaxed),
            dicom_valid: dicom_valid.load(Ordering::Relaxed),
            patients_found: patients.len() as u64,
            elapsed_ms: elapsed.as_millis() as u64,
        },
        patients,
    };

    progress::progress(100, "Done");

    // Write JSON to stdout
    if let Err(e) = serde_json::to_writer(std::io::stdout().lock(), &payload) {
        progress::error(format!("Failed to write JSON output: {}", e));
        std::process::exit(1);
    }
}

fn has_zip_extension(path: &Path) -> bool {
    path.extension()
        .and_then(|e| e.to_str())
        .map(|e| e.eq_ignore_ascii_case("zip"))
        .unwrap_or(false)
}

fn process_single_zip(
    zip_path: &Path,
    password: Option<&[u8]>,
    seven_zip: Option<&Path>,
    max_nested: u32,
    files_scanned: &AtomicU64,
    dicom_valid: &AtomicU64,
) -> Vec<types::StudyInfo> {
    // Check if password is needed
    let enc = zip_handler::detect_encryption(zip_path);
    if enc != zip_handler::EncryptionType::None
        && enc != zip_handler::EncryptionType::Corrupted
        && password.is_none()
    {
        progress::password_needed();
        std::process::exit(2);
    }

    progress::progress(0, format!("Processing ZIP: {}", zip_path.display()));
    let studies = zip_handler::process_zip(zip_path, password, seven_zip, 0, max_nested);
    let valid = studies.len() as u64;
    // files_scanned is approximate for ZIPs — we count valid as scanned
    files_scanned.fetch_add(valid, Ordering::Relaxed);
    dicom_valid.fetch_add(valid, Ordering::Relaxed);
    studies
}

fn process_directory(
    dir_path: &Path,
    password: Option<&[u8]>,
    seven_zip: Option<&Path>,
    max_nested: u32,
    files_scanned: &AtomicU64,
    dicom_valid: &AtomicU64,
) -> Vec<types::StudyInfo> {
    progress::progress(0, "Scanning directory...");

    // Single-pass directory walk: separate ZIPs and DICOM candidates
    let mut zip_files: Vec<PathBuf> = Vec::new();
    let mut dicom_candidates: Vec<PathBuf> = Vec::new();
    let mut scan_count: u64 = 0;

    for entry in jwalk::WalkDir::new(dir_path)
        .into_iter()
        .filter_map(|e| e.ok())
    {
        let path = entry.path();
        if !path.is_file() {
            continue;
        }

        // Skip hidden directories/files
        if path
            .components()
            .any(|c| c.as_os_str().to_string_lossy().starts_with('.'))
        {
            continue;
        }

        scan_count += 1;
        if scan_count % 500 == 0 {
            progress::progress(
                2,
                format!("Scanning directory... ({} files found)", scan_count),
            );
        }

        let name_lower = path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_lowercase();

        if name_lower.ends_with(".zip") {
            zip_files.push(path.to_path_buf());
        } else {
            // Check if it could be a DICOM file
            let ext = path
                .extension()
                .and_then(|e| e.to_str())
                .unwrap_or("")
                .to_lowercase();
            let is_candidate = matches!(ext.as_str(), "dcm" | "ima" | "dicom" | "")
                || ext.chars().all(|c| c.is_ascii_digit());
            if is_candidate {
                dicom_candidates.push(path.to_path_buf());
            }
        }
    }

    files_scanned.store(scan_count, Ordering::Relaxed);
    progress::progress(
        5,
        format!(
            "Found {} ZIPs, {} potential DICOM files",
            zip_files.len(),
            dicom_candidates.len()
        ),
    );

    let mut all_studies: Vec<types::StudyInfo> = Vec::new();

    // Check if any ZIPs need a password
    if !zip_files.is_empty() {
        let any_encrypted = zip_files.iter().any(|z| {
            let enc = zip_handler::detect_encryption(z);
            enc != zip_handler::EncryptionType::None
                && enc != zip_handler::EncryptionType::Corrupted
        });
        if any_encrypted && password.is_none() {
            progress::password_needed();
            std::process::exit(2);
        }

        // Process ZIPs sequentially
        for (i, zip_path) in zip_files.iter().enumerate() {
            let pct = 10 + ((i as i32 * 45) / zip_files.len().max(1) as i32);
            progress::progress(
                pct,
                format!(
                    "Processing ZIP {}/{}: {}",
                    i + 1,
                    zip_files.len(),
                    zip_path.file_name().unwrap_or_default().to_string_lossy()
                ),
            );
            let studies =
                zip_handler::process_zip(zip_path, password, seven_zip, 0, max_nested);
            let count = studies.len() as u64;
            dicom_valid.fetch_add(count, Ordering::Relaxed);
            all_studies.extend(studies);
        }
    }

    // Process loose DICOM files in parallel
    let dicom_start_pct = if zip_files.is_empty() { 5 } else { 55 };
    progress::progress(
        dicom_start_pct,
        format!("Processing {} DICOM candidates...", dicom_candidates.len()),
    );

    let processed_count = AtomicU64::new(0);
    let total_candidates = dicom_candidates.len() as u64;

    let dicom_studies: Vec<types::StudyInfo> = dicom_candidates
        .par_iter()
        .filter_map(|path| {
            let count = processed_count.fetch_add(1, Ordering::Relaxed);
            if count % 100 == 0 && total_candidates > 0 {
                let pct = dicom_start_pct
                    + ((count as i32).min(total_candidates as i32 - 1)
                        * (90 - dicom_start_pct))
                        / (total_candidates as i32).max(1);
                progress::progress(
                    pct.min(90),
                    format!("Processing DICOMs: {}/{}", count, total_candidates),
                );
            }

            if !dicom::is_valid_dicom(path) {
                return None;
            }
            dicom::extract_tags(path).map(|mut info| {
                info.source_path = Some(path.to_string_lossy().into_owned());
                info
            })
        })
        .collect();

    dicom_valid.fetch_add(dicom_studies.len() as u64, Ordering::Relaxed);
    all_studies.extend(dicom_studies);

    progress::progress(
        92,
        format!("Found {} valid DICOM studies", all_studies.len()),
    );
    all_studies
}
