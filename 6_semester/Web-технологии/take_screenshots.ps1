$edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
$baseDir = $PWD.Path
$photoDir = Join-Path $baseDir 'photo'
if (-not (Test-Path $photoDir)) { New-Item -ItemType Directory -Path $photoDir | Out-Null }

$p3 = (Get-ChildItem -Path $baseDir -Recurse -Filter 'page3.html')[0].FullName
$p4 = (Get-ChildItem -Path $baseDir -Recurse -Filter 'index.html' | Where-Object { $_.DirectoryName -match '4' })[0].FullName
$p5 = (Get-ChildItem -Path $baseDir -Recurse -Filter 'index.html' | Where-Object { $_.DirectoryName -match '5' })[0].FullName
$p6 = (Get-ChildItem -Path $baseDir -Recurse -Filter 'index.html' | Where-Object { $_.DirectoryName -match '6' })[0].FullName

$targets = @(
    @($p3, "$photoDir\screenshot_lr3.png"),
    @($p4, "$photoDir\screenshot_lr4.png"),
    @($p5, "$photoDir\screenshot_lr5.png"),
    @($p6, "$photoDir\screenshot_lr6.png")
)

foreach ($t in $targets) {
    if (-not $t[0]) { Write-Host "Skipping unknown path"; continue }
    if (Test-Path $t[1]) { Remove-Item $t[1] -Force -ErrorAction SilentlyContinue }
    $url = $t[0]
    $out = $t[1]
    Write-Host "Capturing to $out"
    
    $job = Start-Job -ArgumentList $edge, $url, $out -ScriptBlock {
        param($e, $u, $o)
        $p = Start-Process -FilePath $e -ArgumentList "--headless --disable-gpu --window-size=1280,1080 --screenshot=`"$o`" `"$u`"" -PassThru -WindowStyle Hidden
        $p.WaitForExit()
        return $p.ExitCode
    }
    
    Wait-Job $job -Timeout 15 | Out-Null
    if ($job.State -ne 'Completed') {
        Write-Host "Timeout! Killing job and Edge."
        Stop-Job $job
        Stop-Process -Name "msedge" -Force -ErrorAction SilentlyContinue
    } else {
        $res = Receive-Job $job
        Write-Host "Done capturing with exit code: $res"
    }
    Remove-Job $job
}
Write-Host "All screenshots processed."
