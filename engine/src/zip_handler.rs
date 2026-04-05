use crate::dicom;
use crate::progress;
use crate::types::StudyInfo;
use rayon::prelude::*;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::Command;

#[derive(Debug, PartialEq)]
pub enum EncryptionType {
    None,
    Traditional,
    Aes,
    Corrupted,
    Unknown,
}

/// Detect ZIP encryption type by inspecting entry flags.
pub fn detect_encryption(path: &Path) -> EncryptionType {
    let file = match File::open(path) {
        Ok(f) => f,
        Err(_) => return EncryptionType::Unknown,
    };
    let reader = std::io::BufReader::new(file);
    let mut archive = match zip::ZipArchive::new(reader) {
        Ok(a) => a,
        Err(_) => return EncryptionType::Corrupted,
    };

    for i in 0..archive.len() {
        let entry = match archive.by_index_raw(i) {
            Ok(e) => e,
            Err(_) => continue,
        };
        if entry.encrypted() {
            let method_str = format!("{:?}", entry.compression());
            if method_str.contains("Aes") || method_str.contains("99") {
                return EncryptionType::Aes;
            }
            return EncryptionType::Traditional;
        }
    }
    EncryptionType::None
}

/// Find 7-Zip executable on the system.
pub fn find_7zip() -> Option<PathBuf> {
    let candidates = [
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
    ];
    for path in &candidates {
        let p = PathBuf::from(path);
        if p.is_file() {
            return Some(p);
        }
    }
    which_7z()
}

fn which_7z() -> Option<PathBuf> {
    for name in &["7z", "7za"] {
        if let Ok(output) = Command::new("where").arg(name).output() {
            if output.status.success() {
                let stdout = String::from_utf8_lossy(&output.stdout);
                if let Some(line) = stdout.lines().next() {
                    let p = PathBuf::from(line.trim());
                    if p.is_file() {
                        return Some(p);
                    }
                }
            }
        }
    }
    None
}

/// Extract a ZIP file using 7-Zip subprocess to a temp directory.
fn extract_with_7zip(
    seven_zip: &Path,
    zip_path: &Path,
    dest: &Path,
    password: Option<&[u8]>,
) -> Result<(), String> {
    let mut cmd = Command::new(seven_zip);
    cmd.arg("x")
        .arg(zip_path)
        .arg(format!("-o{}", dest.display()))
        .arg("-y");

    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        cmd.creation_flags(0x08000000); // CREATE_NO_WINDOW
    }

    if let Some(pw) = password {
        let pw_str = String::from_utf8_lossy(pw);
        cmd.arg(format!("-p{}", pw_str));
    }

    let output = cmd
        .output()
        .map_err(|e| format!("Failed to run 7-Zip: {}", e))?;

    if output.status.success() {
        Ok(())
    } else {
        let stderr = String::from_utf8_lossy(&output.stderr).to_lowercase();
        let stdout = String::from_utf8_lossy(&output.stdout).to_lowercase();
        if stderr.contains("wrong password") || stdout.contains("wrong password")
            || stderr.contains("cannot open encrypted") || stdout.contains("cannot open encrypted")
        {
            Err("WRONG_PASSWORD".into())
        } else {
            Err(format!("7-Zip failed: {}", String::from_utf8_lossy(&output.stderr)))
        }
    }
}

// ─── In-memory streaming (fast path) ────────────────────────────────────────

/// An entry read into memory from a ZIP archive.
struct MemEntry {
    name: String,
    data: Vec<u8>,
}

/// Stream a ZIP archive, parsing DICOM entries on-the-fly without buffering
/// all file contents in memory. Only nested ZIPs are buffered (they're typically small).
/// Returns parsed StudyInfo results or Err for password/corruption issues.
fn stream_and_parse_zip(
    zip_path: &Path,
    password: Option<&[u8]>,
    source_label: &str,
    seven_zip: Option<&Path>,
    nested_level: u32,
    max_nested: u32,
) -> Result<Vec<StudyInfo>, String> {
    let file = File::open(zip_path)
        .map_err(|e| format!("Cannot open ZIP {}: {}", zip_path.display(), e))?;
    let reader = std::io::BufReader::new(file);
    let mut archive = zip::ZipArchive::new(reader)
        .map_err(|e| format!("Invalid ZIP: {}", e))?;

    let total = archive.len();
    let zip_name = zip_path
        .file_name()
        .unwrap_or_default()
        .to_string_lossy();

    let mut all_studies: Vec<StudyInfo> = Vec::new();
    let mut nested_zips: Vec<MemEntry> = Vec::new();
    let mut dicom_count: u64 = 0;

    for i in 0..total {
        // Progress: 5-90% across all entries
        if total > 0 && (i % 200 == 0 || i + 1 == total) {
            let pct = 5 + ((i as i32 * 85) / total.max(1) as i32);
            progress::progress(pct, format!("Processing {}: {}/{} entries", zip_name, i, total));
        }

        let mut entry = if let Some(pw) = password {
            match archive.by_index_decrypt(i, pw) {
                Ok(e) => e,
                Err(e) => {
                    let msg = format!("{}", e);
                    if msg.to_lowercase().contains("password")
                        || msg.to_lowercase().contains("invalid")
                    {
                        return Err("WRONG_PASSWORD".into());
                    }
                    continue;
                }
            }
        } else {
            match archive.by_index(i) {
                Ok(e) => e,
                Err(_) => continue,
            }
        };

        if entry.is_dir() {
            continue;
        }

        let name = match entry.enclosed_name() {
            Some(n) => n.to_string_lossy().to_string(),
            None => continue,
        };

        let name_lower = name.to_lowercase();

        if name_lower.ends_with(".zip") {
            // Buffer nested ZIPs for later processing (typically small)
            if nested_level < max_nested {
                let mut buf = Vec::with_capacity(entry.size() as usize);
                if entry.read_to_end(&mut buf).is_ok() {
                    nested_zips.push(MemEntry { name, data: buf });
                }
            }
            continue;
        }

        // Check if it could be DICOM by extension
        let ext = Path::new(&name)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();
        let is_candidate = matches!(ext.as_str(), "dcm" | "ima" | "dicom" | "")
            || ext.chars().all(|c| c.is_ascii_digit());

        if !is_candidate {
            continue;
        }

        // Read entry into a temporary buffer, parse immediately, then drop
        let mut buf = Vec::with_capacity(entry.size() as usize);
        if entry.read_to_end(&mut buf).is_err() {
            continue;
        }

        if dicom::is_valid_dicom_bytes(&buf) {
            if let Some(mut info) = dicom::extract_tags_from_bytes(&buf) {
                info.source_path = Some(source_label.to_string());
                all_studies.push(info);
                dicom_count += 1;
            }
        }
        // buf is dropped here — memory freed immediately
    }

    // Process nested ZIPs
    for nested in nested_zips {
        let studies = process_zip_from_bytes(
            &nested.data,
            &nested.name,
            password,
            seven_zip,
            nested_level + 1,
            max_nested,
        );
        all_studies.extend(studies);
    }

    progress::progress(92, format!("Found {} DICOM studies in {}", dicom_count, zip_name));
    Ok(all_studies)
}

/// Process a nested ZIP that exists only in memory (from a parent archive).
fn process_zip_from_bytes(
    data: &[u8],
    name: &str,
    password: Option<&[u8]>,
    seven_zip: Option<&Path>,
    nested_level: u32,
    max_nested: u32,
) -> Vec<StudyInfo> {
    if nested_level > max_nested {
        return Vec::new();
    }

    let indent = "  ".repeat(nested_level as usize);
    progress::progress(-1, format!("{}Processing nested {}...", indent, name));

    // Try streaming parse from memory via zip crate
    let cursor = std::io::Cursor::new(data);
    match zip::ZipArchive::new(cursor) {
        Ok(mut archive) => {
            let total = archive.len();
            let mut all_studies: Vec<StudyInfo> = Vec::new();
            let mut nested_zips: Vec<MemEntry> = Vec::new();

            for i in 0..total {
                let mut entry = if let Some(pw) = password {
                    match archive.by_index_decrypt(i, pw) {
                        Ok(e) => e,
                        Err(_) => continue,
                    }
                } else {
                    match archive.by_index(i) {
                        Ok(e) => e,
                        Err(_) => continue,
                    }
                };
                if entry.is_dir() {
                    continue;
                }
                let entry_name = match entry.enclosed_name() {
                    Some(n) => n.to_string_lossy().to_string(),
                    None => continue,
                };

                let entry_lower = entry_name.to_lowercase();
                if entry_lower.ends_with(".zip") {
                    if nested_level < max_nested {
                        let mut buf = Vec::with_capacity(entry.size() as usize);
                        if entry.read_to_end(&mut buf).is_ok() {
                            nested_zips.push(MemEntry { name: entry_name, data: buf });
                        }
                    }
                    continue;
                }

                let ext = Path::new(&entry_name)
                    .extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("")
                    .to_lowercase();
                let is_candidate = matches!(ext.as_str(), "dcm" | "ima" | "dicom" | "")
                    || ext.chars().all(|c| c.is_ascii_digit());
                if !is_candidate {
                    continue;
                }

                let mut buf = Vec::with_capacity(entry.size() as usize);
                if entry.read_to_end(&mut buf).is_err() {
                    continue;
                }
                if dicom::is_valid_dicom_bytes(&buf) {
                    if let Some(mut info) = dicom::extract_tags_from_bytes(&buf) {
                        info.source_path = Some(name.to_string());
                        all_studies.push(info);
                    }
                }
            }

            for nested in nested_zips {
                let studies = process_zip_from_bytes(
                    &nested.data, &nested.name, password, seven_zip,
                    nested_level + 1, max_nested,
                );
                all_studies.extend(studies);
            }

            all_studies
        }
        Err(_) => {
            // Can't read from memory — write to temp file and use 7-Zip
            if let Some(sz) = seven_zip {
                let temp_dir = match tempfile::tempdir() {
                    Ok(d) => d,
                    Err(_) => return Vec::new(),
                };
                let temp_zip = temp_dir.path().join(name);
                if std::fs::write(&temp_zip, data).is_err() {
                    return Vec::new();
                }
                process_zip_disk_fallback(&temp_zip, password, Some(sz), nested_level, max_nested)
            } else {
                Vec::new()
            }
        }
    }
}

// ─── Disk-based extraction (fallback for 7-Zip) ─────────────────────────────

/// Fallback: extract to disk via 7-Zip, then walk and parse.
fn process_zip_disk_fallback(
    zip_path: &Path,
    password: Option<&[u8]>,
    seven_zip: Option<&Path>,
    nested_level: u32,
    max_nested: u32,
) -> Vec<StudyInfo> {
    let temp_dir = match tempfile::tempdir() {
        Ok(d) => d,
        Err(e) => {
            progress::error(format!("Cannot create temp dir: {}", e));
            return Vec::new();
        }
    };

    let indent = "  ".repeat(nested_level as usize);
    progress::progress(
        -1,
        format!(
            "{}Extracting {} (7-Zip)...",
            indent,
            zip_path.file_name().unwrap_or_default().to_string_lossy()
        ),
    );

    if let Some(sz) = seven_zip {
        if let Err(e) = extract_with_7zip(sz, zip_path, temp_dir.path(), password) {
            if e == "WRONG_PASSWORD" {
                progress::error(format!(
                    "Wrong password for {}",
                    zip_path.file_name().unwrap_or_default().to_string_lossy()
                ));
            }
            return Vec::new();
        }
    } else {
        return Vec::new();
    }

    // Walk extracted files
    let mut nested_zips = Vec::new();
    let mut dicom_candidates = Vec::new();

    for entry in jwalk::WalkDir::new(temp_dir.path())
        .into_iter()
        .filter_map(|e| e.ok())
    {
        let path = entry.path();
        if !path.is_file() {
            continue;
        }
        let name_lower = path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_lowercase();
        if name_lower.ends_with(".zip") {
            nested_zips.push(path.to_path_buf());
        } else {
            dicom_candidates.push(path.to_path_buf());
        }
    }

    let mut all_studies = Vec::new();
    let source = zip_path.to_string_lossy().to_string();

    // Nested ZIPs from disk
    for nested in &nested_zips {
        let studies = process_zip(nested, password, seven_zip, nested_level + 1, max_nested);
        all_studies.extend(studies);
    }

    // DICOM candidates from disk (parallel)
    let dicom_studies: Vec<StudyInfo> = dicom_candidates
        .par_iter()
        .filter_map(|path| {
            if dicom::is_valid_dicom(path) {
                dicom::extract_tags(path).map(|mut info| {
                    info.source_path = Some(source.clone());
                    info
                })
            } else {
                None
            }
        })
        .collect();

    all_studies.extend(dicom_studies);
    all_studies
}

// ─── Public entry point ──────────────────────────────────────────────────────

/// Process a ZIP file on disk. Uses in-memory streaming when possible,
/// falls back to 7-Zip disk extraction for encrypted archives the zip crate can't handle.
pub fn process_zip(
    zip_path: &Path,
    password: Option<&[u8]>,
    seven_zip: Option<&Path>,
    nested_level: u32,
    max_nested: u32,
) -> Vec<StudyInfo> {
    if nested_level > max_nested {
        return Vec::new();
    }

    let enc = detect_encryption(zip_path);
    if enc == EncryptionType::Corrupted {
        progress::error(format!(
            "Corrupted ZIP: {}",
            zip_path.file_name().unwrap_or_default().to_string_lossy()
        ));
        return Vec::new();
    }

    if enc != EncryptionType::None && password.is_none() {
        // Encrypted but no password — caller handles this
        return Vec::new();
    }

    let indent = "  ".repeat(nested_level as usize);
    let zip_name = zip_path.file_name().unwrap_or_default().to_string_lossy();
    let source_label = zip_path.to_string_lossy().to_string();

    // Try streaming parse (fast path — no temp dir, process-as-you-go)
    progress::progress(2, format!("{}Processing {}...", indent, zip_name));

    match stream_and_parse_zip(zip_path, password, &source_label, seven_zip, nested_level, max_nested) {
        Ok(studies) => {
            return studies;
        }
        Err(e) if e == "WRONG_PASSWORD" => {
            progress::error(format!("Wrong password for {}", zip_name));
            return Vec::new();
        }
        Err(_) => {
            // zip crate couldn't handle it — fall through to 7-Zip
        }
    }

    // Disk fallback via 7-Zip
    if seven_zip.is_some() {
        process_zip_disk_fallback(zip_path, password, seven_zip, nested_level, max_nested)
    } else {
        progress::error(format!("Cannot extract {}: unsupported format and 7-Zip not available", zip_name));
        Vec::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    #[test]
    fn test_find_7zip_returns_option() {
        let _ = find_7zip();
    }

    #[test]
    fn test_detect_encryption_nonexistent() {
        assert_eq!(
            detect_encryption(Path::new("nonexistent.zip")),
            EncryptionType::Unknown
        );
    }

    #[test]
    fn test_detect_encryption_not_a_zip() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("notazip.zip");
        std::fs::write(&path, b"this is not a zip file").unwrap();
        assert_eq!(detect_encryption(&path), EncryptionType::Corrupted);
    }

    #[test]
    fn test_detect_encryption_unencrypted() {
        let dir = tempfile::tempdir().unwrap();
        let zip_path = dir.path().join("test.zip");
        {
            let file = File::create(&zip_path).unwrap();
            let mut writer = zip::ZipWriter::new(file);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            writer.start_file("test.txt", options).unwrap();
            writer.write_all(b"hello").unwrap();
            writer.finish().unwrap();
        }
        assert_eq!(detect_encryption(&zip_path), EncryptionType::None);
    }

    #[test]
    fn test_process_zip_unencrypted() {
        let dir = tempfile::tempdir().unwrap();
        let zip_path = dir.path().join("test.zip");
        {
            let file = File::create(&zip_path).unwrap();
            let mut writer = zip::ZipWriter::new(file);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            writer.start_file("hello.txt", options).unwrap();
            writer.write_all(b"world").unwrap();
            writer.finish().unwrap();
        }
        // Should succeed without crashing (no DICOM files, so empty result)
        let results = process_zip(&zip_path, None, None, 0, 5);
        assert!(results.is_empty());
    }
}
