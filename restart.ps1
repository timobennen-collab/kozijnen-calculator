# Sluit een draaiende Kozijnen calculator (main.py) en start opnieuw.
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $dir

Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -and $_.CommandLine -match 'kozijnen-calculator' -and $_.CommandLine -match 'main\.py' } |
    ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }

Start-Sleep -Milliseconds 400
Start-Process -FilePath "python" -ArgumentList "main.py" -WorkingDirectory $dir
