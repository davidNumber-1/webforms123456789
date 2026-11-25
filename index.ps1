param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$TextFile,

    [ValidateSet('Notepad', 'Word')]
    [string]$App = 'Notepad',

    [double]$DelaySeconds = 2,

     [switch]$NewTab
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

# Escape characters that SendKeys treats as special: + ^ % ~ ( ) { } [ ]
function ConvertTo-SendKeysText([string]$s) {
    return [regex]::Replace($s, '[\+\^%~\(\)\{\}\[\]]', '{$0}')
}

# --- read the txt file (line by line) ---
if (-not (Test-Path -LiteralPath $TextFile)) {
    Write-Host "Text file not found: $TextFile" -ForegroundColor Red
    exit 1
}
$lines = @(Get-Content -LiteralPath $TextFile)
if ($lines.Count -eq 0) {
    Write-Host "Text file is empty." -ForegroundColor Red
    exit 1
}
$totalWords = ($lines | ForEach-Object { @($_ -split '\s+' | Where-Object { $_ -ne '' }).Count } | Measure-Object -Sum).Sum
Write-Host "$($lines.Count) lines, $totalWords words, $DelaySeconds sec per word."

# --- open the app ---
if ($App -eq 'Word') {
    $proc = Start-Process winword -ArgumentList '/w' -PassThru   # /w = new blank document
    $title = 'Word'
    Start-Sleep -Seconds 6
} else {
    $proc = Start-Process notepad -PassThru
    $title = 'Notepad'
    Start-Sleep -Seconds 2
}

# --- give it focus ---
$shell = New-Object -ComObject WScript.Shell
if (-not $shell.AppActivate($proc.Id)) { $shell.AppActivate($title) | Out-Null }
Start-Sleep -Seconds 1

if ($App -eq 'Notepad' -and $NewTab) {
    [System.Windows.Forms.SendKeys]::SendWait('^t')
    Start-Sleep -Seconds 1
}



$n = 0
for ($l = 0; $l -lt $lines.Count; $l++) {
    $words = @($lines[$l] -split '\s+' | Where-Object { $_ -ne '' })

    for ($i = 0; $i -lt $words.Count; $i++) {
        $n++
        $text = ConvertTo-SendKeysText $words[$i]
        if ($i -lt $words.Count - 1) { $text += ' ' }
        [System.Windows.Forms.SendKeys]::SendWait($text)
        Write-Host ("line {0}/{1}  word {2}/{3}: {4}" -f ($l + 1), $lines.Count, $n, $totalWords, $words[$i])
        Start-Sleep -Milliseconds ([int]($DelaySeconds * 1000))
    }

    # end of line -> Enter (blank lines also get an Enter, no delay)
    if ($l -lt $lines.Count - 1) {
        [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    }
}

Write-Host "Done." -ForegroundColor Green