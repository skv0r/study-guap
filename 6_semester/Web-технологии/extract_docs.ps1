param($OutDir)
$baseDir = $PWD.Path
$files = Get-ChildItem -Path $baseDir -Recurse -Include *.doc, *.docx

foreach ($f in $files) {
    # Skip LR1 and LR2 based on path name if we can, or just extract them all!
    # Wait, LR1 and LR2 are fine if we extract them again. 
    $file = $f.FullName
    $outName = $f.BaseName + ".txt"
    $outPath = Join-Path $OutDir $outName
    Write-Host "Extracting $($f.Name)..."

    $job = Start-Job -ArgumentList $file, $outPath -ScriptBlock {
        param($inFile, $outFile)
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            $doc = $word.Documents.Open($inFile, $false, $true)
            $text = $doc.Content.Text
            $doc.Close($false)
            $word.Quit()
            $text | Out-File -FilePath $outFile -Encoding utf8
            return "SUCCESS"
        } catch {
            return "ERROR: $_"
        }
    }
    
    Wait-Job $job -Timeout 10 | Out-Null
    if ($job.State -ne 'Completed') {
        Write-Host "Timeout reading $($f.Name)! Killing WINWORD."
        Stop-Job $job
        Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
    } else {
        $result = Receive-Job $job
        Write-Host "Result for $($f.Name) : $result"
    }
    Remove-Job $job
}
Write-Host "Extraction finished."
