use serde::Serialize;
use std::collections::{BTreeMap, BTreeSet};

/// Raw per-file extraction result from DICOM parsing.
#[derive(Debug, Clone)]
pub struct StudyInfo {
    pub patient_id: Option<String>,
    pub patient_name: String,
    pub patient_dob: String,
    pub study_date: String,
    pub study_description: String,
    pub series_description: String,
    pub modality: String,
    pub study_instance_uid: Option<String>,
    pub series_instance_uid: Option<String>,
    pub series_number: Option<String>,
    pub institution_name: Option<String>,
    pub institution_address: Option<String>,
    pub department_name: Option<String>,
    pub source_path: Option<String>,
}

/// A single study within a patient record.
#[derive(Debug, Clone, Serialize)]
pub struct Study {
    pub study_date: String,
    pub study_description: String,
    pub all_series: BTreeSet<String>,
}

/// A merged patient record.
#[derive(Debug, Clone, Serialize)]
pub struct Patient {
    pub patient_id: Option<String>,
    pub patient_name: String,
    pub patient_dob: String,
    pub institution_name: Option<String>,
    pub institution_address: Option<String>,
    pub department_name: Option<String>,
    pub studies: BTreeMap<String, Study>,
}

/// Stats about the processing run.
#[derive(Debug, Clone, Serialize)]
pub struct Stats {
    pub files_scanned: u64,
    pub dicom_valid: u64,
    pub patients_found: u64,
    pub elapsed_ms: u64,
}

/// Top-level output payload written to stdout as JSON.
#[derive(Debug, Clone, Serialize)]
pub struct OutputPayload {
    pub patients: BTreeMap<String, Patient>,
    pub stats: Stats,
}
