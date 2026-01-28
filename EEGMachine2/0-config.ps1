# ============================================
# EEGChat - PowerShell Scripts Configuration
# ============================================
# IMPORTANT: Update these settings before running on EEG machine

# API Credentials (use environment variables in production)
$script:EEG_USERNAME = "username"    # Change to actual username
$script:EEG_PASSWORD = "password"    # Change to actual password

# Output Directory - ensure this exists on the EEG machine
$script:OUTPUT_DIR = "C:\EEGChat_Output"

# Create output directory if it doesn't exist
if (-not (Test-Path $script:OUTPUT_DIR)) {
    New-Item -ItemType Directory -Path $script:OUTPUT_DIR -Force | Out-Null
}

# Date range for record queries (modify as needed)
$script:START_DATE = Get-Date "2024-01-01"
$script:END_DATE = Get-Date "2025-12-31"

# Quality Control Thresholds
$script:QC_MIN_DURATION_MINUTES = 1
$script:QC_MIN_CHANNELS = 8
$script:QC_MAX_GAPS = 10
$script:QC_MIN_CONNECTION_RATE = 90

# Segment Filtering Configuration
$script:MIN_SEGMENT_DURATION = 60      # seconds
$script:MAX_GAP_TO_MERGE = 10          # seconds
$script:MIN_END_SEGMENT_DURATION = 120 # seconds

# Waveform Extraction Defaults
$script:WAVEFORM_START_OFFSET = 0      # seconds
$script:WAVEFORM_DURATION = 60         # seconds
$script:MAX_CSV_SAMPLES = 10000        # limit CSV export size

# Anonymization Settings
$script:ANONYMIZE_OUTPUT = $true       # Always anonymize patient data

# Function to initialize API with error handling
function Initialize-ArcApi {
    try {
        $api = New-Object -ComObject Arc.Api.ArcApi
        $result = $api.Login($script:EEG_USERNAME, $script:EEG_PASSWORD)
        
        if (-not $result.IsSuccess) {
            Write-Error "Login failed: $($result.ErrorMessage)"
            return $null
        }
        
        Write-Host "API initialized successfully. Version: $($api.Version)" -ForegroundColor Green
        return $api
    }
    catch {
        Write-Error "Failed to initialize Arc API: $($_.Exception.Message)"
        Write-Host "Ensure Arc API is installed and COM object is registered." -ForegroundColor Yellow
        return $null
    }
}

# Function to safely dispose API
function Close-ArcApi {
    param($api)
    
    if ($null -ne $api) {
        try {
            $api.Dispose()
            Write-Host "API session closed." -ForegroundColor Green
        }
        catch {
            Write-Warning "Error disposing API: $($_.Exception.Message)"
        }
    }
}

# Function to get timestamp for filenames
function Get-Timestamp {
    return Get-Date -Format "yyyyMMdd_HHmmss"
}

# Function to export to CSV with error handling
function Export-SafeCsv {
    param(
        [string]$Content,
        [string]$FileName
    )
    
    $filePath = Join-Path $script:OUTPUT_DIR $FileName
    try {
        $Content | Out-File $filePath -Encoding UTF8
        Write-Host "Exported: $filePath" -ForegroundColor Cyan
        return $filePath
    }
    catch {
        Write-Error "Failed to export $FileName : $($_.Exception.Message)"
        return $null
    }
}

Write-Host "Configuration loaded. Output directory: $($script:OUTPUT_DIR)" -ForegroundColor Green
