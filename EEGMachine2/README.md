
# EEG Machine PowerShell Scripts

**ALL scripts in this folder are READ-ONLY operations.**

- ✅ Original hospital data is **NEVER modified**
- ✅ Scripts only **READ** data from Arc API
- ✅ All output goes to **EXPORTED COPIES** on external storage
- ✅ Patient data is **ANONYMIZED** before export (HIPAA compliant)

The term "Removed Segments" in outputs refers to **filtering labels in the exported CSV files**, NOT deletion of original data.

## Overview

This folder contains PowerShell scripts designed to run on the **Windows EEG machine** with the **Arc API** (version 3.2+) installed. These scripts extract, process, and export EEG data for integration with the EEGChat conversational intelligence system.

## Prerequisites

1. **Windows OS** with Arc EEG software installed
2. **Arc API COM Component** registered (`Arc.Api.ArcApi`)
3. **PowerShell 5.1+** (Windows built-in) or **PowerShell 7+**
4. **Valid Arc credentials** with API access
5. **External drive** for data export (recommended)

## Scripts Overview

| Script | Purpose | Output | Modifies Original? |
|--------|---------|--------|-------------------|
| `0-config.ps1` | Configuration and helper functions | N/A | ❌ No |
| `1-dataExploration.ps1` | System overview, record counts | `system_overview_*.txt` | ❌ No |
| `2-recordExploration.ps1` | List all records with patient info | `records_export_*.csv` | ❌ No |
| `3-singleRecordExploration.ps1` | Deep dive into one record | `record_*_channels.csv` | ❌ No |
| `4-segmentFiltering.ps1` | Filter time segments (export only) | `segments_*.csv` | ❌ No |
| `5-eventanalysis.ps1` | Analyze events (seizures, spikes) | `events_*.csv` | ❌ No |
| `6-waveformDataExtraction.ps1` | Export EEG waveforms (EDF + CSV) | `*.edf`, `waveform_*.csv` | ❌ No |
| `7-documentManagement.ps1` | Export documents and notes | `documents_*.csv` | ❌ No |
| `8-Anonymization.ps1` | Export anonymized patient data | `patient_anonymized_*.json` | ❌ No |
| `9-qualitycontrol.ps1` | Assess data quality | `quality_*.csv` | ❌ No |
| `10-masterExport.ps1` | **Full export to external drive** | Complete data package | ❌ No |

## Quick Start

### Option 1: Full Export to External Drive (RECOMMENDED)

This is the easiest way to export all data at once:

```powershell
cd C:\path\to\EEGchat\EEGMachine

# Connect your external drive (e.g., E:)
# Run master export
.\10-masterExport.ps1 -ExportDrive "E:" -ExportFolder "EEGChat_Export"

# Export specific records only
.\10-masterExport.ps1 -ExportDrive "E:" -StartRecordIndex 0 -EndRecordIndex 10

# Export full EDF files (entire recording, not just first hour)
.\10-masterExport.ps1 -ExportDrive "E:" -ExportFullEDF
```

This creates a complete package on your external drive:
```
E:\EEGChat_Export\Export_20260127_120000\
├── EDF/                    ← EEG waveforms (standard format)
├── Metadata/               ← Record and channel info
├── Events/                 ← Clinical events (seizures, spikes)
├── Segments/               ← Time segment data
├── PatientData_Anonymized/ ← De-identified patient info
├── QualityControl/         ← Data quality reports
├── EXPORT_MANIFEST.json    ← Export summary
└── README.txt              ← Usage instructions
```

### Option 2: Run Individual Scripts

```powershell
cd C:\path\to\EEGchat\EEGMachine

# 1. First, configure settings
notepad .\0-config.ps1

# 2. Verify connectivity
.\1-dataExploration.ps1

# 3. List all records
.\2-recordExploration.ps1

# 4. Export EDF + CSV for a specific record
.\6-waveformDataExtraction.ps1 -RecordIndex 0 -ExportFullRecord

# 5. Export anonymized patient data
.\8-Anonymization.ps1 -RecordIndex 0

# 6. Run quality control
.\9-qualitycontrol.ps1
```

## Configuration Options

### Date Range
```powershell
$script:START_DATE = Get-Date "2024-01-01"
$script:END_DATE = Get-Date "2025-12-31"
```

### Quality Control Thresholds
```powershell
$script:QC_MIN_DURATION_MINUTES = 1    # Minimum recording duration
$script:QC_MIN_CHANNELS = 8            # Minimum channels required
$script:QC_MAX_GAPS = 10               # Maximum data gaps allowed
$script:QC_MIN_CONNECTION_RATE = 90    # Minimum electrode connection %
```

### Segment Filtering
```powershell
$script:MIN_SEGMENT_DURATION = 60      # Minimum segment (seconds)
$script:MAX_GAP_TO_MERGE = 10          # Merge gaps shorter than this
```

### Waveform Extraction
```powershell
$script:WAVEFORM_DURATION = 60         # Duration to extract (seconds)
$script:MAX_CSV_SAMPLES = 10000        # Limit CSV file size
```

## Output Files

All exports are saved to the configured output directory with timestamps.

### EDF Files (European Data Format)
- **Standard format** for EEG data
- Can be opened with:
  - [EDFbrowser](https://www.teuniz.net/edfbrowser/) (free, cross-platform)
  - [EEGLAB](https://sccn.ucsd.edu/eeglab/) (MATLAB)
  - [MNE-Python](https://mne.tools/): `mne.io.read_raw_edf('file.edf')`
  - Any EDF-compatible EEG software

### CSV Files
- Record metadata, events, segments
- Open with Excel, Python pandas, R

### JSON Files
- Structured data for programmatic access
- Anonymized patient information

## Troubleshooting

### "Cannot create Arc API COM object"
- Ensure Arc software is installed
- Run as Administrator
- Check COM registration: `Get-ChildItem HKLM:\SOFTWARE\Classes -Name | Select-String "Arc.Api"`

### "Login failed"
- Verify credentials in `0-config.ps1`
- Check Arc server connectivity
- Confirm user has API access permissions

### "Record open failed"
- Record may be in use by another application
- Try a different record index
- Check record status in Arc software

### Script execution disabled
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

## Data Flow to EEGChat

After running the export scripts:

1. **Disconnect external drive** from EEG machine
2. **Connect to EEGChat processing machine**
3. **Run Python data ingestion**:

```python
from eegchat.data_ingestion import DataIngestionModule

# Point to your export folder
ingestion = DataIngestionModule(data_dir="E:/EEGChat_Export/Export_20260127_120000")

# Process all records
ingestion.process_all_records()
```

## Security & Compliance

### HIPAA Compliance
The `8-Anonymization.ps1` and `10-masterExport.ps1` scripts implement HIPAA Safe Harbor de-identification:
- Names → REDACTED
- Dates → Age ranges only
- MRN → Hashed (preserves linkage)
- SSN → REDACTED
- Addresses → State only, ZIP truncated to 3 digits
- Contact info → REDACTED

## API Reference

These scripts use the Arc API v3.2. Key interfaces:
- `IArcApi` - Main API entry point
- `IRecord` - EEG record access
- `IRecordData` - Waveform and event data
- `IPatient` - Patient demographics
- `IEvent` - Clinical events

See `API_ARC.txt` for full documentation.

---
