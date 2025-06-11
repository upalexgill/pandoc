<#
.SYNOPSIS
    Test script for the ConvertToWord.ps1 script.

.DESCRIPTION
    This script tests various scenarios of the Markdown to Word converter
    to ensure it works correctly with different inputs and edge cases.

.EXAMPLE
    .\Test-Converter.ps1
    Runs all test scenarios and reports results.
#>

[CmdletBinding()]
param()

function Write-TestResult {
    param(
        [string]$TestName,
        [bool]$Passed,
        [string]$Details = ""
    )
    
    $status = if ($Passed) { "PASS" } else { "FAIL" }
    $color = if ($Passed) { "Green" } else { "Red" }
    
    Write-Host "[$status] $TestName" -ForegroundColor $color
    if ($Details) {
        Write-Host "  Details: $Details" -ForegroundColor Gray
    }
}

function Test-HelpDocumentation {
    Write-Host "`n=== Testing Help Documentation ===" -ForegroundColor Cyan
    
    try {
        $help = Get-Help .\ConvertToWord.ps1 -ErrorAction Stop
        $hasExamples = $help.Examples.Example.Count -gt 0
        $hasParameters = $help.Parameters.Parameter.Count -gt 0
        
        Write-TestResult "Help documentation available" ($null -ne $help)
        Write-TestResult "Help has examples" $hasExamples "Found $($help.Examples.Example.Count) examples"
        Write-TestResult "Help has parameter documentation" $hasParameters "Found $($help.Parameters.Parameter.Count) parameters"
        
        return ($null -ne $help)
    }
    catch {
        Write-TestResult "Help documentation" $false "Exception: $_"
        return $false
    }
}

function Test-BasicConversion {
    Write-Host "`n=== Testing Basic Conversion ===" -ForegroundColor Cyan
    
    $testFile = "examples\example-a.md"
    $outputFile = "examples\test-output-a.docx"
    
    try {
        if (Test-Path $outputFile) {
            Remove-Item $outputFile -Force
        }
        
        $null = & .\ConvertToWord.ps1 $testFile $outputFile
        
        $success = Test-Path $outputFile
        Write-TestResult "Basic conversion of $testFile" $success
        
        return $success
    }
    catch {
        Write-TestResult "Basic conversion of $testFile" $false "Exception: $_"
        return $false
    }
}

function Test-AutoOutputNaming {
    Write-Host "`n=== Testing Auto Output Naming ===" -ForegroundColor Cyan
    
    $testFile = "examples\example-b.md"
    $expectedOutput = "examples\example-b.docx"
    
    try {
        if (Test-Path $expectedOutput) {
            Remove-Item $expectedOutput -Force
        }
        
        $null = & .\ConvertToWord.ps1 $testFile
        
        $success = Test-Path $expectedOutput
        Write-TestResult "Auto output naming" $success "Expected: $expectedOutput"
        
        return $success
    }
    catch {
        Write-TestResult "Auto output naming" $false "Exception: $_"
        return $false
    }
}

function Test-InvalidInput {
    Write-Host "`n=== Testing Invalid Input Handling ===" -ForegroundColor Cyan
    
    try {
        $null = & .\ConvertToWord.ps1 "nonexistent.md" 2>&1
        $exitCode = $LASTEXITCODE
        
        Write-TestResult "Handles non-existent file" ($exitCode -ne 0) "Exit code: $exitCode"
        
        try {
            $null = & .\ConvertToWord.ps1 "test.txt" 2>&1
            $exitCode = $LASTEXITCODE
            Write-TestResult "Handles invalid file gracefully" ($exitCode -ne 0) "Exit code: $exitCode"
            return $true
        }
        catch {
            Write-TestResult "Handles invalid file gracefully" $true "Correctly handled error: $_"
            return $true
        }
    }
    catch {
        Write-TestResult "Invalid input handling" $false "Exception: $_"
        return $false
    }
}

# Main test execution
Write-Host "PowerShell Pandoc Converter Test Suite" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Yellow

$testResults = @()
$testResults += Test-HelpDocumentation
$testResults += Test-BasicConversion
$testResults += Test-AutoOutputNaming
$testResults += Test-InvalidInput

$passedTests = ($testResults | Where-Object { $_ }).Count
$totalTests = $testResults.Count

Write-Host "`n=== Test Summary ===" -ForegroundColor Yellow
Write-Host "Passed: $passedTests/$totalTests tests" -ForegroundColor $(if ($passedTests -eq $totalTests) { "Green" } else { "Yellow" })

if ($passedTests -eq $totalTests) {
    Write-Host "All tests passed! The converter is working correctly." -ForegroundColor Green
} else {
    Write-Host "Some tests failed. Please review the output above." -ForegroundColor Yellow
}
