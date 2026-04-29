$ErrorActionPreference = "Stop"

Write-Host "== ЛР7: запуск MongoDB + сервера + smoke-test ==" -ForegroundColor Cyan

$root = $PSScriptRoot
$port = 5007
$baseUrl = "http://127.0.0.1:$port"

function Wait-HttpReady {
    param(
        [string]$Url,
        [int]$TimeoutSec = 30
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            $resp = Invoke-WebRequest -UseBasicParsing "$Url/connect/temperature" -TimeoutSec 2
            if ($resp.StatusCode -eq 200) {
                return $true
            }
        } catch {
            Start-Sleep -Milliseconds 500
        }
    }
    return $false
}

try {
    $svc = Get-Service MongoDB -ErrorAction Stop
    if ($svc.Status -ne "Running") {
        Start-Service MongoDB
        Start-Sleep -Seconds 1
    }
    Write-Host "[OK] MongoDB service: Running" -ForegroundColor Green
} catch {
    Write-Host "[ERR] Сервис MongoDB не найден. Установи MongoDB Server." -ForegroundColor Red
    exit 1
}

$existing = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if ($existing) {
    try {
        Stop-Process -Id $existing.OwningProcess -Force -ErrorAction Stop
        Write-Host "[WARN] Освобожден порт $port (остановлен PID $($existing.OwningProcess))" -ForegroundColor Yellow
        Start-Sleep -Milliseconds 500
    } catch {
        Write-Host "[ERR] Порт $port занят и процесс не удалось остановить." -ForegroundColor Red
        exit 1
    }
}

$server = Start-Process -FilePath "npm.cmd" -ArgumentList "start" -WorkingDirectory $root -PassThru
Write-Host "[OK] Сервер запускается, PID=$($server.Id)" -ForegroundColor Green

if (-not (Wait-HttpReady -Url $baseUrl -TimeoutSec 35)) {
    Write-Host "[ERR] Сервер не поднялся на $baseUrl" -ForegroundColor Red
    exit 1
}
Write-Host "[OK] Сервер доступен: $baseUrl" -ForegroundColor Green

$null = Invoke-WebRequest -UseBasicParsing "$baseUrl/command/temperature?value=21.4" | Out-Null
$null = Invoke-WebRequest -UseBasicParsing "$baseUrl/command/temperature?value=24.9" | Out-Null
$null = Invoke-WebRequest -UseBasicParsing "$baseUrl/command/temperature?value=26.2" | Out-Null
Write-Host "[OK] Тестовые команды отправлены" -ForegroundColor Green

$check = @"
const { MongoClient } = require("mongodb");
(async () => {
  const client = new MongoClient("mongodb://127.0.0.1:27017");
  await client.connect();
  const db = client.db("iot_logger_db");
  const temp = await db.collection("Temperature").countDocuments();
  const heater = await db.collection("Heater").countDocuments();
  const lastTemp = await db.collection("Temperature").find().sort({ _id: -1 }).limit(1).toArray();
  const lastHeater = await db.collection("Heater").find().sort({ _id: -1 }).limit(1).toArray();
  console.log(JSON.stringify({
    tempCount: temp,
    heaterCount: heater,
    lastTemp: lastTemp[0] || null,
    lastHeater: lastHeater[0] || null
  }, null, 2));
  await client.close();
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
"@

$tmpCheck = Join-Path $root ".tmp-lab7-mongo-check.cjs"
Set-Content -Path $tmpCheck -Value $check -Encoding UTF8
node $tmpCheck
if ($LASTEXITCODE -ne 0) {
    Remove-Item $tmpCheck -Force -ErrorAction SilentlyContinue
    throw "MongoDB check failed"
}
Remove-Item $tmpCheck -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Done. Lab7 server is running at $baseUrl" -ForegroundColor Cyan
Write-Host ("To stop server run: Stop-Process -Id {0}" -f $server.Id) -ForegroundColor DarkGray
