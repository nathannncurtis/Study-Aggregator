use serde::Serialize;
use std::io::Write;
use std::sync::Mutex;

static STDERR_LOCK: Mutex<()> = Mutex::new(());

#[derive(Debug, Serialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum ProgressMessage {
    Progress {
        percent: i32,
        message: String,
    },
    PasswordNeeded,
    Error {
        message: String,
    },
}

/// Emit a progress message as a JSON line to stderr.
/// Thread-safe via mutex to prevent interleaved output from rayon.
pub fn emit(msg: &ProgressMessage) {
    let _lock = STDERR_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    if let Ok(json) = serde_json::to_string(msg) {
        let _ = writeln!(std::io::stderr(), "{}", json);
    }
}

pub fn progress(percent: i32, message: impl Into<String>) {
    emit(&ProgressMessage::Progress {
        percent,
        message: message.into(),
    });
}

pub fn error(message: impl Into<String>) {
    emit(&ProgressMessage::Error {
        message: message.into(),
    });
}

pub fn password_needed() {
    emit(&ProgressMessage::PasswordNeeded);
}
