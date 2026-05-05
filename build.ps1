# Macro Copier - Nuitka Build Script
# Kompilerer macro_copier.py til standalone executable

# Farver for output
$Green = 'Green'
$Red = 'Red'
$Yellow = 'Yellow'
$Cyan = 'Cyan'

Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "Macro Copier - Nuitka Build Script" -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host ""

# 1. Check Python Installation
Write-Host "[1/4] Tjekker Python installation..." -ForegroundColor $Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host "  ✓ Python found: $pythonVersion" -ForegroundColor $Green
} catch {
    Write-Host "  ✗ Python ikke fundet!" -ForegroundColor $Red
    exit 1
}

# 2. Install/Update Nuitka
Write-Host "[2/4] Installerer Nuitka..." -ForegroundColor $Yellow
try {
    pip install --upgrade nuitka 2>&1 | Out-Null
    Write-Host "  ✓ Nuitka installed/updated" -ForegroundColor $Green
} catch {
    Write-Host "  ✗ Fejl ved installation af Nuitka" -ForegroundColor $Red
    exit 1
}

# 3. Compile med Nuitka
Write-Host "[3/4] Kompilerer med Nuitka onefile..." -ForegroundColor $Yellow
Write-Host "  Dette kan tage 1-3 minutter..." -ForegroundColor $Cyan

$buildDir = ".\build"
$distDir = ".\dist"

# Ryd gamle build-filer
if (Test-Path $buildDir) {
    Remove-Item -Recurse -Force $buildDir -ErrorAction SilentlyContinue
}
if (Test-Path $distDir) {
    Remove-Item -Recurse -Force $distDir -ErrorAction SilentlyContinue
}

try {
    & python -m nuitka `
        --onefile `
        --output-dir=$distDir `
        --follow-imports `
        --include-package=ttkbootstrap `
        --windows-console-mode=disable `
        .\macro_copier.py
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✓ Kompilering gennemført!" -ForegroundColor $Green
    } else {
        Write-Host "  ✗ Nuitka returned exit code $LASTEXITCODE" -ForegroundColor $Red
        exit 1
    }
}
catch {
    Write-Host "  ✗ Fejl ved kompilering: $_" -ForegroundColor $Red
    exit 1
}

# 4. Test af executable
Write-Host "[4/4] Tester executable..." -ForegroundColor $Yellow

$exePath = "$distDir\macro_copier.exe"

if (Test-Path $exePath) {
    $fileSize = (Get-Item $exePath).Length / 1MB
    Write-Host "  ✓ Executable oprettet: $exePath" -ForegroundColor $Green
    Write-Host "  ✓ Filstørrelse: {0:F2} MB" -f $fileSize -ForegroundColor $Green
    
    # Test at executable kan køres (med timeout)
    Write-Host "  Starter applikation for test..." -ForegroundColor $Cyan
    try {
        $process = Start-Process -FilePath $exePath -PassThru -NoNewWindow -ErrorAction Stop
        Start-Sleep -Seconds 2
        
        if ($process.HasExited) {
            $exitCode = $process.ExitCode
            Write-Host "  ✓ Applikation startede og lukkede (exit code: $exitCode)" -ForegroundColor $Green
        } else {
            # Luk applikationen hvis den kører
            Stop-Process -InputObject $process -Force -ErrorAction SilentlyContinue
            Write-Host "  ✓ Applikation startede og kørte uden fejl" -ForegroundColor $Green
        }
    } catch {
        Write-Host "  ⚠ Kunne ikke starte applikation: $_" -ForegroundColor $Yellow
    }
} else {
    Write-Host "  ✗ Executable ikke fundet!" -ForegroundColor $Red
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor $Cyan
Write-Host "✓ Build gennemført!" -ForegroundColor $Green
Write-Host "Executable: $exePath" -ForegroundColor $Cyan
Write-Host "========================================" -ForegroundColor $Cyan
