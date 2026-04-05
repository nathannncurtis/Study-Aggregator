use crate::types::{Patient, Study, StudyInfo};
use std::collections::BTreeMap;

/// Normalize a patient name for matching: uppercase, sort parts alphabetically.
/// Returns None if the name is empty or "Unknown".
pub fn normalize_name(name: &str) -> Option<Vec<String>> {
    let cleaned = name
        .replace('^', " ")
        .replace(',', " ")
        .replace('_', " ");
    let mut parts: Vec<String> = cleaned
        .split_whitespace()
        .map(|s| s.to_uppercase())
        .filter(|s| !s.is_empty())
        .collect();
    if parts.is_empty() || (parts.len() == 1 && parts[0] == "UNKNOWN") {
        return None;
    }
    parts.sort();
    Some(parts)
}

/// Check if two names match after normalization.
pub fn names_match(a: &str, b: &str) -> bool {
    match (normalize_name(a), normalize_name(b)) {
        (Some(na), Some(nb)) => na == nb,
        _ => false,
    }
}

/// Check if a patient has all unknown/empty identifying information.
fn is_all_unknown(patient: &Patient) -> bool {
    let id_unknown = patient.patient_id.as_ref().map_or(true, |id| id.trim().is_empty());
    let name_unknown = patient.patient_name.trim().is_empty()
        || patient.patient_name.trim() == "Unknown";
    let dob_unknown = patient.patient_dob.trim().is_empty()
        || patient.patient_dob.trim() == "Unknown";
    id_unknown && name_unknown && dob_unknown
}

/// Merge a list of study records into grouped patient records.
/// Matches by patient ID first, then by normalized name with DOB conflict detection.
pub fn merge_patients(studies: Vec<StudyInfo>) -> BTreeMap<String, Patient> {
    let mut patients: Vec<(String, Patient)> = Vec::new();

    for study in &studies {
        let mut matched_idx: Option<usize> = None;

        // Phase 1: match by patient ID
        if let Some(ref pid) = study.patient_id {
            if !pid.is_empty() {
                for (i, (_, p)) in patients.iter().enumerate() {
                    if p.patient_id.as_deref() == Some(pid.as_str()) {
                        matched_idx = Some(i);
                        break;
                    }
                }
            }
        }

        // Phase 2: match by normalized name (skip if DOB conflict)
        if matched_idx.is_none() {
            for (i, (_, p)) in patients.iter().enumerate() {
                if names_match(&study.patient_name, &p.patient_name) {
                    // Check DOB conflict
                    if study.patient_dob != "Unknown"
                        && p.patient_dob != "Unknown"
                        && study.patient_dob != p.patient_dob
                    {
                        continue; // DOB conflict — separate patients
                    }
                    matched_idx = Some(i);
                    break;
                }
            }
        }

        if let Some(idx) = matched_idx {
            let (_, ref mut existing) = patients[idx];

            // Update DOB if existing is Unknown
            if existing.patient_dob == "Unknown" && study.patient_dob != "Unknown" {
                existing.patient_dob = study.patient_dob.clone();
            }

            // Update facility info if not set
            if existing.institution_name.is_none() {
                existing.institution_name = study.institution_name.clone();
            }
            if existing.institution_address.is_none() {
                existing.institution_address = study.institution_address.clone();
            }
            if existing.department_name.is_none() {
                existing.department_name = study.department_name.clone();
            }

            // Add study/series
            add_study_to_patient(existing, study);
        } else {
            // Create new patient
            let key = format!(
                "{}_{}_{}",
                study.patient_name,
                study.patient_dob,
                study.patient_id.as_deref().unwrap_or("NO_ID")
            );
            let mut patient = Patient {
                patient_id: study.patient_id.clone(),
                patient_name: study.patient_name.clone(),
                patient_dob: study.patient_dob.clone(),
                institution_name: study.institution_name.clone(),
                institution_address: study.institution_address.clone(),
                department_name: study.department_name.clone(),
                studies: BTreeMap::new(),
            };
            add_study_to_patient(&mut patient, study);
            patients.push((key, patient));
        }
    }

    // Filter out all-unknown patients
    patients
        .into_iter()
        .filter(|(_, p)| !is_all_unknown(p))
        .collect()
}

fn add_study_to_patient(patient: &mut Patient, study: &StudyInfo) {
    let study_uid = study
        .study_instance_uid
        .clone()
        .unwrap_or_else(|| format!("{}_{}", study.study_date, study.study_description));

    let entry = patient.studies.entry(study_uid).or_insert_with(|| Study {
        study_date: study.study_date.clone(),
        study_description: study.study_description.clone(),
        all_series: std::collections::BTreeSet::new(),
    });

    let series_uid = study.series_instance_uid.clone().unwrap_or_else(|| {
        format!(
            "{}_{}",
            study.series_number.as_deref().unwrap_or("Unknown"),
            study.series_description
        )
    });
    entry.all_series.insert(series_uid);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_normalize_name() {
        assert_eq!(normalize_name("DOE^JOHN"), Some(vec!["DOE".into(), "JOHN".into()]));
        assert_eq!(normalize_name("John Smith"), Some(vec!["JOHN".into(), "SMITH".into()]));
        assert_eq!(normalize_name("Smith, John"), Some(vec!["JOHN".into(), "SMITH".into()]));
        assert_eq!(normalize_name("Unknown"), None);
        assert_eq!(normalize_name(""), None);
    }

    #[test]
    fn test_names_match() {
        assert!(names_match("DOE^JOHN", "John Doe"));
        assert!(names_match("SMITH, JANE", "jane_smith"));
        assert!(!names_match("DOE^JOHN", "SMITH^JANE"));
        assert!(!names_match("Unknown", "DOE^JOHN"));
    }

    #[test]
    fn test_merge_by_patient_id() {
        let studies = vec![
            make_study("12345", "DOE^JOHN", "01-15-1990", "Study A"),
            make_study("12345", "DOE^JOHN", "01-15-1990", "Study B"),
        ];
        let result = merge_patients(studies);
        assert_eq!(result.len(), 1);
        let patient = result.values().next().unwrap();
        assert_eq!(patient.studies.len(), 2);
    }

    #[test]
    fn test_merge_by_name() {
        let studies = vec![
            make_study("", "DOE^JOHN", "01-15-1990", "Study A"),
            make_study("", "John Doe", "01-15-1990", "Study B"),
        ];
        let result = merge_patients(studies);
        assert_eq!(result.len(), 1);
    }

    #[test]
    fn test_dob_conflict_creates_separate_patients() {
        let studies = vec![
            make_study("", "DOE^JOHN", "01-15-1990", "Study A"),
            make_study("", "John Doe", "02-20-1985", "Study B"),
        ];
        let result = merge_patients(studies);
        assert_eq!(result.len(), 2);
    }

    #[test]
    fn test_all_unknown_filtered() {
        let studies = vec![make_study("", "Unknown", "Unknown", "Study A")];
        let result = merge_patients(studies);
        assert_eq!(result.len(), 0);
    }

    #[test]
    fn test_dob_update_on_merge() {
        let studies = vec![
            make_study("12345", "DOE^JOHN", "Unknown", "Study A"),
            make_study("12345", "DOE^JOHN", "01-15-1990", "Study B"),
        ];
        let result = merge_patients(studies);
        assert_eq!(result.len(), 1);
        let patient = result.values().next().unwrap();
        assert_eq!(patient.patient_dob, "01-15-1990");
    }

    fn make_study(id: &str, name: &str, dob: &str, desc: &str) -> StudyInfo {
        StudyInfo {
            patient_id: if id.is_empty() { None } else { Some(id.into()) },
            patient_name: name.into(),
            patient_dob: dob.into(),
            study_date: "01-01-2026".into(),
            study_description: desc.into(),
            series_description: "Series 1".into(),
            modality: "CR".into(),
            study_instance_uid: Some(format!("1.2.3.{}", desc)),
            series_instance_uid: Some(format!("1.2.3.{}.1", desc)),
            series_number: Some("1".into()),
            institution_name: None,
            institution_address: None,
            department_name: None,
            source_path: None,
        }
    }
}
