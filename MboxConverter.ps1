#!/usr/bin/env pwsh
<#
.SYNOPSIS
    PowerShell wrapper for MBOX Converter - works on Windows, Linux, macOS

.DESCRIPTION
    This script provides PowerShell-friendly functions for converting MBOX files.
    It wraps the Python mbox_converter.py tool with PowerShell pipeline support
    and structured output.

.EXAMPLE
    # Convert single file
    Convert-MBox -Path inbox.mbox -Format csv

.EXAMPLE
    # Batch convert with filtering
    Get-ChildItem *.mbox | Convert-MBox -Format eml -DateAfter "2023-01-01"

.EXAMPLE
    # Get MBOX information
    Get-MBoxInfo -Path inbox.mbox

.NOTES
    Requires Python 3.8+ with mbox_converter.py
#>

# Script configuration
$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConverterScript = Join-Path $ScriptDir 'mbox_converter.py'

function Test-PythonAvailable {
    <#
    .SYNOPSIS
        Check if Python is available
    #>
    try {
        $null = & python --version 2>&1
        return $true
    } catch {
        return $false
    }
}

function Get-PythonCommand {
    <#
    .SYNOPSIS
        Get the Python command (python or python3)
    #>
    if (Get-Command python3 -ErrorAction SilentlyContinue) {
        return 'python3'
    }
    return 'python'
}

function Convert-MBox {
    <#
    .SYNOPSIS
        Convert MBOX files to CSV, EML, TXT, or PST format

    .PARAMETER Path
        Path to MBOX file(s). Accepts pipeline input.

    .PARAMETER Format
        Output format: csv, eml, txt, pst

    .PARAMETER OutputDirectory
        Directory for output files

    .PARAMETER DateAfter
        Only include emails after this date (YYYY-MM-DD)

    .PARAMETER DateBefore
        Only include emails before this date (YYYY-MM-DD)

    .PARAMETER FromPattern
        Filter by sender (regex pattern)

    .PARAMETER SubjectPattern
        Filter by subject (regex pattern)

    .PARAMETER BodyContains
        Filter by body content (substring)

    .PARAMETER DryRun
        Preview conversion without writing files

    .PARAMETER ShowProgress
        Show progress bar during conversion

    .EXAMPLE
        Convert-MBox -Path inbox.mbox -Format csv -OutputDirectory ./output

    .EXAMPLE
        Get-ChildItem *.mbox | Convert-MBox -Format eml -DateAfter "2023-01-01"
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias('FullName')]
        [string[]]$Path,

        [Parameter(Mandatory=$false)]
        [ValidateSet('csv', 'eml', 'txt', 'pst')]
        [string]$Format = 'csv',

        [Parameter(Mandatory=$false)]
        [Alias('OutputDir', 'Out')]
        [string]$OutputDirectory,

        [Parameter(Mandatory=$false)]
        [string]$DateAfter,

        [Parameter(Mandatory=$false)]
        [string]$DateBefore,

        [Parameter(Mandatory=$false)]
        [string]$FromPattern,

        [Parameter(Mandatory=$false)]
        [string]$ToPattern,

        [Parameter(Mandatory=$false)]
        [string]$SubjectPattern,

        [Parameter(Mandatory=$false)]
        [string]$BodyContains,

        [Parameter(Mandatory=$false)]
        [string]$Encoding = 'utf-8',

        [Parameter(Mandatory=$false)]
        [switch]$DryRun,

        [Parameter(Mandatory=$false)]
        [switch]$ShowProgress,

        [Parameter(Mandatory=$false)]
        [switch]$Quiet
    )

    begin {
        if (-not (Test-PythonAvailable)) {
            throw "Python is not available. Please install Python 3.8+"
        }

        if (-not (Test-Path $ConverterScript)) {
            throw "Converter script not found: $ConverterScript"
        }

        $python = Get-PythonCommand
        $allPaths = @()
    }

    process {
        foreach ($p in $Path) {
            $allPaths += $p
        }
    }

    end {
        if ($allPaths.Count -eq 0) {
            Write-Warning "No MBOX files specified"
            return
        }

        # Build arguments
        $args = @('convert')
        $args += $allPaths
        $args += '--format', $Format

        if ($OutputDirectory) { $args += '--output-dir', $OutputDirectory }
        if ($DateAfter) { $args += '--date-after', $DateAfter }
        if ($DateBefore) { $args += '--date-before', $DateBefore }
        if ($FromPattern) { $args += '--from-pattern', $FromPattern }
        if ($ToPattern) { $args += '--to-pattern', $ToPattern }
        if ($SubjectPattern) { $args += '--subject-pattern', $SubjectPattern }
        if ($BodyContains) { $args += '--body-contains', $BodyContains }
        if ($Encoding -ne 'utf-8') { $args += '--encoding', $Encoding }
        if ($DryRun) { $args += '--dry-run' }
        if ($ShowProgress) { $args += '--progress' }
        if ($Quiet) { $args += '--quiet' }

        Write-Verbose "Running: $python $ConverterScript $($args -join ' ')"

        if ($PSCmdlet.ShouldProcess($allPaths -join ', ', 'Convert MBOX')) {
            try {
                & $python $ConverterScript @args
                $exitCode = $LASTEXITCODE

                # Return structured result
                [PSCustomObject]@{
                    Files = $allPaths
                    Format = $Format
                    OutputDirectory = $OutputDirectory
                    ExitCode = $exitCode
                    Success = ($exitCode -eq 0)
                }
            } catch {
                Write-Error "Conversion failed: $_"
            }
        }
    }
}

function Get-MBoxInfo {
    <#
    .SYNOPSIS
        Get information about MBOX file(s)

    .PARAMETER Path
        Path to MBOX file(s). Accepts pipeline input.

    .EXAMPLE
        Get-MBoxInfo -Path inbox.mbox

    .EXAMPLE
        Get-ChildItem *.mbox | Get-MBoxInfo
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias('FullName')]
        [string[]]$Path
    )

    begin {
        if (-not (Test-PythonAvailable)) {
            throw "Python is not available. Please install Python 3.8+"
        }

        $python = Get-PythonCommand
        $allPaths = @()
    }

    process {
        foreach ($p in $Path) {
            $allPaths += $p
        }
    }

    end {
        $args = @('info', '--json') + $allPaths

        try {
            $output = & $python $ConverterScript @args | Out-String
            $info = $output | ConvertFrom-Json

            foreach ($item in $info) {
                [PSCustomObject]@{
                    Path = $item.path
                    TotalEmails = $item.total_emails
                    UniqueSenders = $item.unique_senders
                    WithAttachments = $item.emails_with_attachments
                    FileSizeMB = $item.file_size_mb
                    DateRange = if ($item.date_range) { "$($item.date_range.earliest) to $($item.date_range.latest)" } else { 'N/A' }
                }
            }
        } catch {
            Write-Error "Failed to get MBOX info: $_"
        }
    }
}

function Get-MBoxEmails {
    <#
    .SYNOPSIS
        List emails in MBOX file(s) with optional filtering

    .PARAMETER Path
        Path to MBOX file(s)

    .PARAMETER Limit
        Maximum number of emails to return

    .PARAMETER DateAfter
        Filter by date

    .PARAMETER FromPattern
        Filter by sender (regex)

    .PARAMETER SubjectPattern
        Filter by subject (regex)

    .EXAMPLE
        Get-MBoxEmails -Path inbox.mbox -Limit 50 -SubjectPattern "invoice"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias('FullName')]
        [string[]]$Path,

        [Parameter(Mandatory=$false)]
        [int]$Limit = 100,

        [Parameter(Mandatory=$false)]
        [string]$DateAfter,

        [Parameter(Mandatory=$false)]
        [string]$DateBefore,

        [Parameter(Mandatory=$false)]
        [string]$FromPattern,

        [Parameter(Mandatory=$false)]
        [string]$SubjectPattern,

        [Parameter(Mandatory=$false)]
        [string]$BodyContains
    )

    begin {
        if (-not (Test-PythonAvailable)) {
            throw "Python is not available. Please install Python 3.8+"
        }

        $python = Get-PythonCommand
        $allPaths = @()
    }

    process {
        foreach ($p in $Path) {
            $allPaths += $p
        }
    }

    end {
        $args = @('list', '--json', '--limit', $Limit) + $allPaths

        if ($DateAfter) { $args += '--date-after', $DateAfter }
        if ($DateBefore) { $args += '--date-before', $DateBefore }
        if ($FromPattern) { $args += '--from-pattern', $FromPattern }
        if ($SubjectPattern) { $args += '--subject-pattern', $SubjectPattern }
        if ($BodyContains) { $args += '--body-contains', $BodyContains }

        try {
            $output = & $python $ConverterScript @args | Out-String
            $emails = $output | ConvertFrom-Json

            foreach ($email in $emails) {
                [PSCustomObject]@{
                    File = $email.file
                    Index = $email.index
                    From = $email.from
                    To = $email.to
                    Subject = $email.subject
                    Date = $email.date
                }
            }
        } catch {
            Write-Error "Failed to list emails: $_"
        }
    }
}

# Export functions
Export-ModuleMember -Function Convert-MBox, Get-MBoxInfo, Get-MBoxEmails

# Display help if run directly
if ($MyInvocation.InvocationName -ne '.') {
    Write-Host @"

MBOX Converter PowerShell Functions
====================================

Available Commands:
  Convert-MBox    - Convert MBOX files to CSV/EML/TXT/PST
  Get-MBoxInfo    - Get information about MBOX files
  Get-MBoxEmails  - List emails with filtering

Examples:
  # Convert to CSV
  Convert-MBox -Path inbox.mbox -Format csv

  # Batch convert with pipeline
  Get-ChildItem *.mbox | Convert-MBox -Format eml -OutputDirectory ./output

  # Filter by date and sender
  Convert-MBox -Path inbox.mbox -Format csv -DateAfter "2023-01-01" -FromPattern "@company.com"

  # Get file info
  Get-MBoxInfo -Path inbox.mbox

  # Search emails
  Get-MBoxEmails -Path inbox.mbox -SubjectPattern "invoice" -Limit 50

To import as module:
  Import-Module ./MboxConverter.ps1

"@
}
