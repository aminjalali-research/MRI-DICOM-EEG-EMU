# ========================================
# SAFETY CHECKLIST SCRIPT
# Run BEFORE starting anonymization
# ========================================

Write-Host "=== PRE-ANONYMIZATION SAFETY CHECKLIST ==="
Write-Host ""

$checks = @(
    @{
        Name = "Database Backup"
        Question = "Do you have a recent backup of the CadLink database?"
    },
    @{
        Name = "Export Directory"
        Question = "Is the export directory on a drive with sufficient space?"
    },
    @{
        Name = "Read-Only Approach"
        Question = "Do you understand that only EXPORTED files will be modified?"
    },
    @{
        Name = "Mapping File Security"
        Question = "Do you have a secure location to store the anonymization mapping?"
    },
    @{
        Name = "Verification Plan"
        Question = "Will you run the verification script after anonymization?"
    }
)

$allPassed = $true

foreach ($check in $checks) {
    Write-Host "`n$($check.Name):"
    Write-Host "  $($check.Question)"
    $response = Read-Host "  [Y/N]"
    
    if ($response -ne "Y" -and $response -ne "y") {
        Write-Host "  ✗ NOT READY" -ForegroundColor Red
        $allPassed = $false
    } else {
        Write-Host "  ✓ READY" -ForegroundColor Green
    }
}

Write-Host "`n========================================="
if ($allPassed) {
    Write-Host "✓ ALL CHECKS PASSED - Ready to proceed" -ForegroundColor Green
} else {
    Write-Host "✗ CHECKLIST INCOMPLETE - Address issues before proceeding" -ForegroundColor Red
}
Write-Host "========================================="
