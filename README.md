# Study Aggregator

A Windows desktop application for processing, organizing, and aggregating DICOM medical imaging studies. Study Aggregator reads DICOM files from directories, ZIP archives, and optical drives, then generates organized patient reports via clipboard or PDF.

## Features

- **DICOM Processing** - Reads and extracts metadata from DICOM files (.dcm, .ima, .dicom, extensionless)
- **ZIP Archive Support** - Handles encrypted (AES & traditional) and unencrypted ZIP files with nested archive support up to 10 levels deep
- **Patient Merging** - Intelligently groups studies by patient using ID, name, and DOB matching with conflict detection
- **Multiple Output Formats** - Export as formatted clipboard text or PDF reports using a fillable template
- **Context Menu Integration** - Right-click "Study Aggregator" option on files, folders, drives, and ZIP files
- **CD/Drive Support** - Reads from optical drives, network drives, and removable media
- **Auto Updates** - Daily scheduled update checks against GitHub releases with automatic download
- **Multi-threaded** - Parallel DICOM processing (4x CPU cores, max 32 threads) for fast file scanning

## Installation

Download the latest `StudyAggregatorSetup.exe` from [GitHub Releases](https://github.com/nathannncurtis/Study-Aggregator/releases) and run the installer.

The installer will:
- Install to `%APPDATA%\Study Aggregator`
- Register Windows context menu entries
- Create a daily scheduled task for update checks

## Usage

**Right-click** any folder, ZIP file, or drive in Windows Explorer and select **Study Aggregator** from the context menu.

Or run from the command line:

```
"Study Aggregator.exe" <path_to_directory_or_zip>
```

On launch you'll be prompted to choose an output mode:
- **Clipboard** - Copies a formatted text report to clipboard
- **PDF** - Generates a PDF report saved to your chosen directory
- **Both** - Produces both outputs

## Development Setup

**Requirements:**
- Python 3.13
- Windows

**Install dependencies:**

```
pip install -r req.txt
```

**Run from source:**

```
python "Study Aggregator.py" <path>
```

**Build executable:**

```
python setup.py build
```

The frozen executable is output to `build/exe.win-amd64-3.13/`. Use [Inno Setup](https://jrsoftware.org/isinfo.php) with `StudyAggSetup.iss` to compile the installer.

## Project Structure

| File | Description |
|---|---|
| `Study Aggregator.py` | Main application |
| `setup.py` | cx_Freeze build configuration |
| `build.bat` | Build and code signing script |
| `reg.py` | Register context menu and update scheduler |
| `unreg.py` | Unregister context menu entries |
| `update_checker.py` | Standalone update checker utility |
| `bd.pdf` | PDF report template |
| `req.txt` | Python dependencies |
| `StudyAggSetup.iss` | Inno Setup installer script |
| `version.txt` | Current version number |

## Dependencies

- [PyQt5](https://pypi.org/project/PyQt5/) - GUI framework
- [pydicom](https://pypi.org/project/pydicom/) - DICOM file parsing
- [pypdf](https://pypi.org/project/pypdf/) - PDF template filling
- [reportlab](https://pypi.org/project/reportlab/) - PDF generation
- [pyzipper](https://pypi.org/project/pyzipper/) - Encrypted ZIP handling
- [numpy](https://pypi.org/project/numpy/) - Required by pydicom
- [clipboard](https://pypi.org/project/clipboard/) - Clipboard access

**Optional:** [7-Zip](https://www.7-zip.org/) - Preferred for ZIP extraction (falls back to pyzipper)

## License

Proprietary software owned by Ronsin Litigation Support Services. See [LICENSE](LICENSE) for details.
