use crate::types::StudyInfo;
use memmap2::Mmap;
use std::collections::HashSet;
use std::fs::File;
use std::path::Path;

// DICOM tag constants (group, element) as u32: (group << 16) | element
const PATIENT_ID: u32 = 0x0010_0020;
const PATIENT_NAME: u32 = 0x0010_0010;
const PATIENT_BIRTH_DATE: u32 = 0x0010_0030;
const STUDY_DATE: u32 = 0x0008_0020;
const STUDY_DESCRIPTION: u32 = 0x0008_1030;
const SERIES_DESCRIPTION: u32 = 0x0008_103E;
const MODALITY: u32 = 0x0008_0060;
const STUDY_INSTANCE_UID: u32 = 0x0020_000D;
const SERIES_INSTANCE_UID: u32 = 0x0020_000E;
const SERIES_NUMBER: u32 = 0x0020_0011;
const PROTOCOL_NAME: u32 = 0x0018_1030;
const REQUESTED_PROCEDURE_DESC: u32 = 0x0032_1060;
const STUDY_COMMENTS: u32 = 0x0032_4000;
const INSTITUTION_NAME: u32 = 0x0008_0080;
const INSTITUTION_ADDRESS: u32 = 0x0008_0081;
const INSTITUTIONAL_DEPT_NAME: u32 = 0x0008_1040;

// Pixel data tag — stop parsing when we hit this
const PIXEL_DATA: u32 = 0x7FE0_0010;

// Sequence/item delimiter tags
const ITEM_TAG: u32 = 0xFFFE_E000;
const ITEM_DELIM: u32 = 0xFFFE_E00D;
const SEQ_DELIM: u32 = 0xFFFE_E0DD;

/// VRs that use 4-byte length in explicit VR (the "long" VRs)
const LONG_VRS: &[&[u8; 2]] = &[
    b"OB", b"OD", b"OF", b"OL", b"OW", b"SQ", b"UC", b"UN", b"UR", b"UT",
];

const SKIP_EXTENSIONS: &[&str] = &[
    ".pdf", ".txt", ".exe", ".bat", ".inf", ".chm", ".log", ".xml", ".html",
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".doc", ".docx",
    ".xls", ".xlsx", ".zip", ".rar", ".7z",
];

/// Tags we care about — build a set for O(1) lookup.
fn wanted_tags() -> HashSet<u32> {
    [
        PATIENT_ID, PATIENT_NAME, PATIENT_BIRTH_DATE, STUDY_DATE,
        STUDY_DESCRIPTION, SERIES_DESCRIPTION, MODALITY,
        STUDY_INSTANCE_UID, SERIES_INSTANCE_UID, SERIES_NUMBER,
        PROTOCOL_NAME, REQUESTED_PROCEDURE_DESC, STUDY_COMMENTS,
        INSTITUTION_NAME, INSTITUTION_ADDRESS, INSTITUTIONAL_DEPT_NAME,
    ]
    .into_iter()
    .collect()
}

/// Quick validation: is this likely a DICOM file?
pub fn is_valid_dicom(path: &Path) -> bool {
    // Size check
    let meta = match std::fs::metadata(path) {
        Ok(m) => m,
        Err(_) => return false,
    };
    if meta.len() < 132 || meta.is_dir() {
        return false;
    }

    // Extension skip
    if let Some(ext) = path.extension().and_then(|e| e.to_str()) {
        let ext_lower = format!(".{}", ext.to_lowercase());
        if SKIP_EXTENSIONS.contains(&ext_lower.as_str()) {
            return false;
        }
    }

    // mmap and check DICM magic at offset 128
    let file = match File::open(path) {
        Ok(f) => f,
        Err(_) => return false,
    };
    let mmap = match unsafe { Mmap::map(&file) } {
        Ok(m) => m,
        Err(_) => return false,
    };
    if mmap.len() < 132 {
        return false;
    }

    // Standard DICOM: 128-byte preamble + "DICM"
    if &mmap[128..132] == b"DICM" {
        return true;
    }

    // Non-standard: some files skip the preamble entirely.
    // Check if the first bytes look like a DICOM tag (group 0008 is common).
    if mmap.len() >= 4 {
        let group = u16::from_le_bytes([mmap[0], mmap[1]]);
        // Group 0008 (study/series info) or 0002 (file meta) are valid starts
        if group == 0x0008 || group == 0x0002 {
            return true;
        }
    }

    false
}

/// Extract the 15 needed tags from a DICOM file using zero-copy mmap parsing.
pub fn extract_tags(path: &Path) -> Option<StudyInfo> {
    let file = File::open(path).ok()?;
    let mmap = unsafe { Mmap::map(&file) }.ok()?;
    let mut info = extract_tags_from_bytes(&mmap[..])?;
    info.source_path = Some(path.to_string_lossy().into_owned());
    Some(info)
}

/// Extract tags from an in-memory byte buffer (for streaming from ZIP archives).
pub fn extract_tags_from_bytes(data: &[u8]) -> Option<StudyInfo> {
    if data.len() < 4 {
        return None;
    }

    // Determine start offset: skip 132-byte preamble if "DICM" present
    let mut offset = if data.len() >= 132 && &data[128..132] == b"DICM" {
        132
    } else {
        // No preamble — check for group 0002 or 0008 at start
        let group = read_u16_le(data, 0)?;
        if group == 0x0002 || group == 0x0008 {
            0
        } else {
            return None;
        }
    };

    let wanted = wanted_tags();
    let mut found: std::collections::HashMap<u32, String> = std::collections::HashMap::new();
    let mut is_explicit_vr = None; // auto-detect
    let mut found_count = 0;
    let total_wanted = wanted.len();

    // Parse file meta information (group 0002) — always explicit VR little-endian
    // Then parse dataset — VR determined by transfer syntax or auto-detected
    while offset + 4 <= data.len() {
        // Read tag
        let group = read_u16_le(data, offset)?;
        let element = read_u16_le(data, offset + 2)?;
        let tag = ((group as u32) << 16) | (element as u32);

        // Stop at pixel data
        if tag == PIXEL_DATA {
            break;
        }

        // Handle sequence/item delimiters
        if tag == ITEM_TAG || tag == ITEM_DELIM || tag == SEQ_DELIM {
            offset += 4;
            if offset + 4 <= data.len() {
                let len = read_u32_le(data, offset)?;
                offset += 4;
                if len != 0xFFFFFFFF && len != 0 {
                    offset += len as usize;
                }
            }
            continue;
        }

        offset += 4; // past tag

        // Determine VR and length
        let (vr_bytes, value_len, header_size) = if group == 0x0002 {
            // File meta is always explicit VR LE
            parse_explicit_vr_element(data, offset)?
        } else {
            // Auto-detect VR on first non-meta element
            if is_explicit_vr.is_none() {
                is_explicit_vr = Some(detect_explicit_vr(data, offset));
            }
            if is_explicit_vr.unwrap_or(true) {
                parse_explicit_vr_element(data, offset)?
            } else {
                parse_implicit_vr_element(data, offset)?
            }
        };

        offset += header_size;

        // Undefined length — skip sequences we can't handle simply
        if value_len == 0xFFFFFFFF {
            // Skip to the sequence delimiter
            offset = skip_undefined_length(data, offset);
            continue;
        }

        let value_end = offset + value_len as usize;
        if value_end > data.len() {
            break; // truncated file
        }

        // Extract value if this is a wanted tag
        if wanted.contains(&tag) && !found.contains_key(&tag) {
            let value_bytes = &data[offset..value_end];
            let value = decode_string(value_bytes, &vr_bytes);
            if !value.is_empty() {
                found.insert(tag, value);
                found_count += 1;
                if found_count >= total_wanted {
                    break; // got everything
                }
            }
        }

        offset = value_end;
    }

    if found.is_empty() {
        return None;
    }

    // Build StudyInfo from extracted tags
    let patient_name_raw = found.get(&PATIENT_NAME).cloned().unwrap_or_default();
    let patient_name = normalize_patient_name(&patient_name_raw);

    let patient_dob = found
        .get(&PATIENT_BIRTH_DATE)
        .map(|d| format_dicom_date(d))
        .unwrap_or_else(|| "Unknown".into());

    let study_date = found
        .get(&STUDY_DATE)
        .map(|d| format_dicom_date(d))
        .unwrap_or_else(|| "Unknown".into());

    let study_desc = found.get(&STUDY_DESCRIPTION).cloned().unwrap_or_default();
    let requested_proc = found.get(&REQUESTED_PROCEDURE_DESC).cloned().unwrap_or_default();
    let protocol = found.get(&PROTOCOL_NAME).cloned().unwrap_or_default();
    let series_desc_raw = found.get(&SERIES_DESCRIPTION).cloned().unwrap_or_default();
    let modality = found.get(&MODALITY).cloned().unwrap_or_default();

    let study_description = resolve_study_description(
        &study_desc,
        &requested_proc,
        &protocol,
        &series_desc_raw,
        &modality,
    );

    let series_description = if !series_desc_raw.is_empty() {
        series_desc_raw
    } else {
        let sn = found.get(&SERIES_NUMBER).cloned().unwrap_or_default();
        if !sn.is_empty() {
            format!("Series {}", sn)
        } else {
            "Unknown Series".into()
        }
    };

    Some(StudyInfo {
        patient_id: found.get(&PATIENT_ID).cloned().filter(|s| !s.is_empty()),
        patient_name: if patient_name.is_empty() { "Unknown".into() } else { patient_name },
        patient_dob,
        study_date,
        study_description,
        series_description,
        modality,
        study_instance_uid: found.get(&STUDY_INSTANCE_UID).cloned().filter(|s| !s.is_empty()),
        series_instance_uid: found.get(&SERIES_INSTANCE_UID).cloned().filter(|s| !s.is_empty()),
        series_number: found.get(&SERIES_NUMBER).cloned().filter(|s| !s.is_empty()),
        institution_name: found.get(&INSTITUTION_NAME).cloned().filter(|s| !s.is_empty()),
        institution_address: found.get(&INSTITUTION_ADDRESS).cloned().filter(|s| !s.is_empty()),
        department_name: found.get(&INSTITUTIONAL_DEPT_NAME).cloned().filter(|s| !s.is_empty()),
        source_path: None,
    })
}

/// Quick validation on in-memory bytes: does this look like DICOM?
pub fn is_valid_dicom_bytes(data: &[u8]) -> bool {
    if data.len() < 132 {
        return false;
    }
    // Standard: DICM at offset 128
    if &data[128..132] == b"DICM" {
        return true;
    }
    // No preamble: group 0002 or 0008 at start
    if data.len() >= 4 {
        let group = u16::from_le_bytes([data[0], data[1]]);
        if group == 0x0002 || group == 0x0008 {
            return true;
        }
    }
    false
}

// --- Internal helpers ---

#[inline]
fn read_u16_le(data: &[u8], offset: usize) -> Option<u16> {
    if offset + 2 > data.len() {
        return None;
    }
    Some(u16::from_le_bytes([data[offset], data[offset + 1]]))
}

#[inline]
fn read_u32_le(data: &[u8], offset: usize) -> Option<u32> {
    if offset + 4 > data.len() {
        return None;
    }
    Some(u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
}

/// Try to detect if the dataset uses explicit VR by checking if bytes at
/// the current offset look like a valid 2-char VR.
fn detect_explicit_vr(data: &[u8], offset: usize) -> bool {
    if offset + 2 > data.len() {
        return false;
    }
    let a = data[offset];
    let b = data[offset + 1];
    // Valid VR characters are uppercase ASCII letters
    a.is_ascii_uppercase() && b.is_ascii_uppercase()
}

/// Parse an explicit VR element header. Returns (vr_bytes, value_length, header_bytes_consumed).
fn parse_explicit_vr_element(data: &[u8], offset: usize) -> Option<([u8; 2], u32, usize)> {
    if offset + 4 > data.len() {
        return None;
    }
    let vr = [data[offset], data[offset + 1]];

    let is_long = LONG_VRS.iter().any(|v| **v == vr);
    if is_long {
        // VR (2) + reserved (2) + length (4) = 8 bytes
        if offset + 8 > data.len() {
            return None;
        }
        let len = read_u32_le(data, offset + 4)?;
        Some((vr, len, 8))
    } else {
        // VR (2) + length (2) = 4 bytes
        let len = read_u16_le(data, offset + 2)? as u32;
        Some((vr, len, 4))
    }
}

/// Parse an implicit VR element header. Returns (vr_bytes [0,0], value_length, header_bytes_consumed).
fn parse_implicit_vr_element(data: &[u8], offset: usize) -> Option<([u8; 2], u32, usize)> {
    if offset + 4 > data.len() {
        return None;
    }
    let len = read_u32_le(data, offset)?;
    Some(([0, 0], len, 4))
}

/// Skip an undefined-length sequence by scanning for the sequence delimiter.
fn skip_undefined_length(data: &[u8], mut offset: usize) -> usize {
    while offset + 8 <= data.len() {
        let group = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let element = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
        let tag = ((group as u32) << 16) | (element as u32);

        if tag == SEQ_DELIM {
            offset += 8; // skip delimiter + its 4-byte zero length
            return offset;
        }

        // Item tag — read item length
        if tag == ITEM_TAG {
            let len = read_u32_le(data, offset + 4).unwrap_or(0);
            offset += 8;
            if len != 0xFFFFFFFF {
                offset += len as usize;
            }
            // if item has undefined length, just keep scanning
            continue;
        }

        // Unknown — advance byte by byte as last resort
        offset += 1;
    }
    data.len() // ran off the end
}

/// Decode DICOM value bytes to a trimmed string.
/// Handles both ASCII and basic multi-byte (strip null padding, trim whitespace).
fn decode_string(bytes: &[u8], _vr: &[u8; 2]) -> String {
    // Strip trailing nulls and spaces (DICOM padding)
    let end = bytes
        .iter()
        .rposition(|&b| b != 0 && b != b' ')
        .map(|i| i + 1)
        .unwrap_or(0);
    let trimmed = &bytes[..end];

    // Try UTF-8 first, fall back to latin-1
    match std::str::from_utf8(trimmed) {
        Ok(s) => s.trim().to_string(),
        Err(_) => trimmed.iter().map(|&b| b as char).collect::<String>().trim().to_string(),
    }
}

/// Format a DICOM date (YYYYMMDD) to MM-DD-YYYY.
fn format_dicom_date(raw: &str) -> String {
    let clean = raw.trim();
    if clean.len() == 8 && clean.chars().all(|c| c.is_ascii_digit()) {
        format!("{}-{}-{}", &clean[4..6], &clean[6..8], &clean[0..4])
    } else {
        "Unknown".into()
    }
}

/// Normalize patient name: replace separators, collapse whitespace, title case.
fn normalize_patient_name(raw: &str) -> String {
    let cleaned = raw
        .replace('^', " ")
        .replace(',', " ")
        .replace('_', " ");
    let parts: Vec<&str> = cleaned.split_whitespace().collect();
    if parts.is_empty() {
        return String::new();
    }
    parts
        .iter()
        .map(|p| title_case(p))
        .collect::<Vec<_>>()
        .join(" ")
}

/// Title-case a single word.
fn title_case(s: &str) -> String {
    let mut chars = s.chars();
    match chars.next() {
        None => String::new(),
        Some(c) => {
            let upper: String = c.to_uppercase().collect();
            let lower: String = chars.flat_map(|c| c.to_lowercase()).collect();
            format!("{}{}", upper, lower)
        }
    }
}

/// Resolve study description using the priority chain.
fn resolve_study_description(
    study_desc: &str,
    requested_proc: &str,
    protocol: &str,
    series_desc: &str,
    modality: &str,
) -> String {
    let meaningful = |s: &str| {
        let lower = s.trim().to_lowercase();
        !lower.is_empty() && lower != "study" && lower != "unknown"
    };

    let desc = if meaningful(study_desc) {
        study_desc.to_string()
    } else if meaningful(requested_proc) {
        // Strip CPT code prefixes (common in dental/medical)
        let trimmed = requested_proc.trim();
        if trimmed.len() > 5
            && trimmed[..5].chars().all(|c| c.is_ascii_digit())
            && trimmed.as_bytes().get(5).copied() == Some(b' ')
        {
            trimmed[6..].to_string()
        } else {
            trimmed.to_string()
        }
    } else if meaningful(protocol) {
        protocol.to_string()
    } else if meaningful(series_desc) {
        format!("{} Study", series_desc)
    } else if !modality.trim().is_empty() {
        format!("{} Study", modality.trim())
    } else {
        "Study".into()
    };

    // Title case and clean up
    desc.split_whitespace()
        .map(|w| title_case(w))
        .collect::<Vec<_>>()
        .join(" ")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_dicom_date() {
        assert_eq!(format_dicom_date("20260315"), "03-15-2026");
        assert_eq!(format_dicom_date("bad"), "Unknown");
        assert_eq!(format_dicom_date(""), "Unknown");
        assert_eq!(format_dicom_date("2026031"), "Unknown");
    }

    #[test]
    fn test_normalize_patient_name() {
        assert_eq!(normalize_patient_name("DOE^JOHN"), "Doe John");
        assert_eq!(normalize_patient_name("DOE,JOHN M"), "Doe John M");
        assert_eq!(normalize_patient_name("doe_john"), "Doe John");
        assert_eq!(normalize_patient_name("  SMITH  JANE  "), "Smith Jane");
        assert_eq!(normalize_patient_name(""), "");
    }

    #[test]
    fn test_resolve_study_description() {
        assert_eq!(
            resolve_study_description("CHEST XRAY", "", "", "", "CR"),
            "Chest Xray"
        );
        assert_eq!(
            resolve_study_description("", "73110 WRIST 2+ VIEWS", "", "", "CR"),
            "Wrist 2+ Views"
        );
        assert_eq!(
            resolve_study_description("", "", "AP Pelvis", "", ""),
            "Ap Pelvis"
        );
        assert_eq!(
            resolve_study_description("", "", "", "Sagittal T1", "MR"),
            "Sagittal T1 Study"
        );
        assert_eq!(
            resolve_study_description("", "", "", "", "CT"),
            "Ct Study"
        );
        assert_eq!(
            resolve_study_description("", "", "", "", ""),
            "Study"
        );
        // "Study" and "unknown" are not meaningful
        assert_eq!(
            resolve_study_description("Study", "", "", "", "CR"),
            "Cr Study"
        );
    }

    #[test]
    fn test_title_case() {
        assert_eq!(title_case("HELLO"), "Hello");
        assert_eq!(title_case("world"), "World");
        assert_eq!(title_case(""), "");
    }

    #[test]
    fn test_detect_explicit_vr() {
        // "CS" is a valid VR
        assert!(detect_explicit_vr(&[b'C', b'S'], 0));
        // Lowercase is not
        assert!(!detect_explicit_vr(&[b'c', b's'], 0));
        // Numbers are not
        assert!(!detect_explicit_vr(&[b'1', b'2'], 0));
    }

    #[test]
    fn test_decode_string() {
        let vr = [0, 0];
        assert_eq!(decode_string(b"HELLO\0\0 ", &vr), "HELLO");
        assert_eq!(decode_string(b"  test  \0", &vr), "test");
        assert_eq!(decode_string(b"", &vr), "");
    }

    #[test]
    fn test_is_valid_dicom_rejects_small_files() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("small.dcm");
        std::fs::write(&path, b"too small").unwrap();
        assert!(!is_valid_dicom(&path));
    }

    #[test]
    fn test_is_valid_dicom_accepts_magic() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.dcm");
        let mut data = vec![0u8; 256];
        // Write DICM at offset 128
        data[128] = b'D';
        data[129] = b'I';
        data[130] = b'C';
        data[131] = b'M';
        std::fs::write(&path, &data).unwrap();
        assert!(is_valid_dicom(&path));
    }

    #[test]
    fn test_is_valid_dicom_accepts_no_preamble() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("nopreamble.dcm");
        // Group 0008 at start (no preamble)
        let mut data = vec![0u8; 256];
        data[0] = 0x08;
        data[1] = 0x00;
        std::fs::write(&path, &data).unwrap();
        assert!(is_valid_dicom(&path));
    }

    #[test]
    fn test_is_valid_dicom_skips_extensions() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.pdf");
        let mut data = vec![0u8; 256];
        data[128..132].copy_from_slice(b"DICM");
        std::fs::write(&path, &data).unwrap();
        assert!(!is_valid_dicom(&path));
    }
}
