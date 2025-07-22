# Archive Google Apps Script Deployments
# Archives all deployments of version X or lower for a Google Apps Script project
# Uses clasp undeploy to move deployments to archived state (not permanently deleted)
# Allows choosing between dev and prod configurations

param(
    [Parameter(Mandatory=$true)]
    [string]$MaxVersion,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("dev", "prod")]
    [string]$Environment = "dev"
)

# Function to compare version numbers
function Compare-Version {
    param([string]$Version1, [string]$Version2)
    
    $v1 = $Version1 -replace '^[^\d]*', '' -split '\.' | ForEach-Object { [int]$_ }
    $v2 = $Version2 -replace '^[^\d]*', '' -split '\.' | ForEach-Object { [int]$_ }
    
    $maxLength = [Math]::Max($v1.Length, $v2.Length)
    
    for ($i = 0; $i -lt $maxLength; $i++) {
        $num1 = if ($i -lt $v1.Length) { $v1[$i] } else { 0 }
        $num2 = if ($i -lt $v2.Length) { $v2[$i] } else { 0 }
        
        if ($num1 -gt $num2) { return 1 }
        if ($num1 -lt $num2) { return -1 }
    }
    return 0
}

# Load configuration based on environment
$configFile = if ($Environment -eq "prod") { ".clasp.prod.json" } else { ".clasp.dev.json" }

if (!(Test-Path $configFile)) {
    Write-Host "Error: Configuration file $configFile not found!" -ForegroundColor Red
    exit 1
}

$config = Get-Content $configFile | ConvertFrom-Json
$scriptId = $config.scriptId
$projectId = $config.projectId

Write-Host "Using $Environment configuration:" -ForegroundColor Cyan
Write-Host "  Script ID: $scriptId" -ForegroundColor Gray
Write-Host "  Project ID: $projectId" -ForegroundColor Gray
Write-Host "  Max Version: $MaxVersion" -ForegroundColor Gray
Write-Host ""

# Check if clasp is available
try {
    $claspVersion = clasp --version
    Write-Host "Using clasp: $claspVersion" -ForegroundColor Green
} catch {
    Write-Host "Error: clasp not found. Please install clasp first." -ForegroundColor Red
    exit 1
}

# Get deployments list
Write-Host "Fetching deployments..." -ForegroundColor Yellow

try {
    # Temporarily copy the config to .clasp.json for clasp commands
    Copy-Item $configFile ".clasp.json" -Force
    
    # Get deployments
    $deploymentsOutput = clasp deployments 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error getting deployments: $deploymentsOutput" -ForegroundColor Red
        exit 1
    }
    
    # Parse deployments
    $deployments = @()
    $lines = $deploymentsOutput -split "`n"
    
    foreach ($line in $lines) {
        if ($line -match "^\s*-\s*(\S+)\s+@(\S+)\s+(.*)$") {
            # Skip @head version
            if ($matches[2] -ne "head") {
                $deployments += @{
                    Id = $matches[1]
                    Version = $matches[2]
                    Description = $matches[3].Trim()
                }
            }
        }
    }
    
    if ($deployments.Count -eq 0) {
        Write-Host "No deployments found." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "Found $($deployments.Count) deployment(s)" -ForegroundColor Green
    
    # Filter deployments by version
    $deploymentsToArchive = @()
    
    foreach ($deployment in $deployments) {
        $compareResult = Compare-Version $deployment.Version $MaxVersion
        if ($compareResult -le 0) {  # Version is equal or lower
            $deploymentsToArchive += $deployment
        }
    }
    
    if ($deploymentsToArchive.Count -eq 0) {
        Write-Host "No deployments found with version $MaxVersion or lower." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "`nDeployments to archive (version $MaxVersion or lower):" -ForegroundColor Cyan
    foreach ($deployment in $deploymentsToArchive) {
        Write-Host "  - $($deployment.Id) @$($deployment.Version): $($deployment.Description)" -ForegroundColor Gray
    }
    
    # Confirm archiving
    Write-Host ""
    $confirm = Read-Host "Archive $($deploymentsToArchive.Count) deployment(s)? (y/N)"
    
    if ($confirm -ne "y" -and $confirm -ne "Y") {
        Write-Host "Archiving cancelled." -ForegroundColor Yellow
        exit 0
    }
    
    # Archive each deployment
    $successCount = 0
    
    foreach ($deployment in $deploymentsToArchive) {
        try {
            Write-Host "Archiving deployment $($deployment.Id) @$($deployment.Version)..." -ForegroundColor Yellow
            
            # Archive the deployment using clasp undeploy (moves to archived state)
            $archiveResult = clasp undeploy $deployment.Id 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  Successfully archived deployment $($deployment.Id)" -ForegroundColor Green
                $successCount++
            } else {
                Write-Host "  Error archiving deployment $($deployment.Id): $archiveResult" -ForegroundColor Red
            }
            
        } catch {
            Write-Host "  Error archiving deployment $($deployment.Id): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "`nArchiving complete!" -ForegroundColor Green
    Write-Host "Successfully archived: $successCount/$($deploymentsToArchive.Count) deployments" -ForegroundColor Green
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
} finally {
    # Remove the temporary .clasp.json file
    if (Test-Path ".clasp.json") {
        Remove-Item ".clasp.json" -Force
    }
}