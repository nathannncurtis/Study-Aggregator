# Study Aggregator

A fast Windows desktop tool for extracting and summarizing DICOM study metadata from folders, ZIP archives, and optical drives. Results are copied to your clipboard as a formatted report.

As of **v4.0**, the DICOM parsing hot path is implemented in Rust and runs as a native subprocess, delivering roughly a **100-400x** speedup over the previous pydicom-based implementation on typical office hardware.

## Features

- **Native Rust DICOM engine** - zero-copy memory-mapped tag extraction, no Python per-file overhead
- **Parallel processing** - directory walking (jwalk) and DICOM parsing (rayon) use every CPU core
- **ZIP archive support** - unencrypted and encrypted (AES + traditional) ZIPs, nested archives, streamed parsing so multi-gigabyte archives don't balloon RAM
- **Patient merging** - groups studies by patient via ID and normalized name matching with DOB conflict detection
- **Handles non-conformant DICOMs** - implicit/explicit VR, missing preamble, truncated files, undefined-length sequences, private tags
- **Clipboard output** - formatted patient/study report ready to paste into any application

## Architecture

```
PyQt5 GUI  ->  subprocess  ->  study-agg-engine.exe  (Rust)
           <-  stderr: progress JSON lines
           <-  stdout: final result JSON
```

The Python wrapper handles UI, input/password prompts, and clipboard. Everything performance-sensitive lives in the Rust engine under `engine/`.

## Installation

Download the latest `StudyAggregatorSetup.exe` from the [Releases](https://github.com/nathannncurtis/Study-Aggregator/releases) page and run it. The installer places the app in `%APPDATA%\Study Aggregator` and optionally creates a desktop shortcut.

## Usage

Drag and drop a folder, ZIP file, or drive onto `Study Aggregator.exe`, or run from the command line:

```
"Study Aggregator.exe" <path_to_directory_or_zip>
```

When processing finishes, a formatted report is copied to your clipboard and a confirmation dialog appears.

## Building from Source

**Requirements:**
- Python 3.12
- Rust (stable toolchain) - install via [rustup](https://rustup.rs/)
- Windows
- [Inno Setup 6](https://jrsoftware.org/isinfo.php) (for installer compilation)

**Steps:**

```cmd
pip install -r req.txt
pip install coil-compiler
build.bat
```

`build.bat` compiles the Rust engine (`cargo build --release`), bundles the Python app with Coil, and copies the engine binary into `dist\Study Aggregator\`. After it finishes, open `StudyAggSetup.iss` in Inno Setup to produce `Output\StudyAggregatorSetup.exe`.

**Run from source (without bundling):**

```cmd
cd engine && cargo build --release && cd ..
python "Study Aggregator.py" <path>
```

## Project Structure

| Path | Description |
|---|---|
| `Study Aggregator.py` | PyQt5 GUI wrapper |
| `engine/` | Rust DICOM engine (Cargo crate) |
| `engine/src/dicom.rs` | Zero-copy mmap DICOM parser |
| `engine/src/zip_handler.rs` | Streaming ZIP extraction |
| `engine/src/patient.rs` | Patient merging and name normalization |
| `coil.toml` | Coil bundling configuration |
| `build.bat` | Full build script |
| `StudyAggSetup.iss` | Inno Setup installer script |
| `req.txt` | Python dependencies |
| `version.txt` | Current version |

## Dependencies

**Python:**
- [PyQt5](https://pypi.org/project/PyQt5/) - GUI
- [clipboard](https://pypi.org/project/clipboard/) - clipboard access

**Rust (see `engine/Cargo.toml`):**
- `memmap2` - memory-mapped file I/O
- `zip` - archive handling
- `jwalk` - parallel directory walking
- `rayon` - data parallelism
- `clap`, `serde`, `serde_json`, `tempfile`, `base64`

## License

MIT License. See [LICENSE](LICENSE).
