@echo off
:: =====================================================================
::   FLATTEN PDFs — Drag-and-Drop Tool
:: =====================================================================
::   Double-click to open a browser-based drag-and-drop tool
::   that flattens PDFs using qpdf (free, reliable, handles
::   forms, annotations, signatures, and complex documents).
::
::   - No admin rights needed
::   - First run downloads qpdf (~23 MB, one time only)
::   - Opens in your default browser
::   - Files never leave your computer
:: =====================================================================

title Flatten PDFs - Starting...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$lines = @(); $reading = $false; foreach ($line in (Get-Content -LiteralPath '%~f0' -Encoding UTF8)) { if ($line -eq '# >>>PS_START<<<') { $reading = $true; continue }; if ($line -eq '# >>>PS_END<<<') { break }; if ($reading -and $line.StartsWith('## ')) { $lines += $line.Substring(3) } elseif ($reading -and $line -eq '##') { $lines += '' } }; $sp = Join-Path $env:TEMP 'Flatten-PDFs-Server.ps1'; $lines -join [Environment]::NewLine | Set-Content -LiteralPath $sp -Encoding UTF8; try { & $sp } finally { Remove-Item $sp -Force -ErrorAction SilentlyContinue }"

exit /b

# >>>PS_START<<<
## # ═══════════════════════════════════════════════════════════════════
## #  qpdf setup
## # ═══════════════════════════════════════════════════════════════════
## $qpdfVersion = "12.3.2"
## $qpdfDir     = Join-Path $env:LOCALAPPDATA "qpdf"
## $qpdfBinDir  = Join-Path $qpdfDir "qpdf-$qpdfVersion-msvc64"
## $qpdfExe     = Join-Path $qpdfBinDir "bin\qpdf.exe"
## $qpdfZipUrl  = "https://github.com/qpdf/qpdf/releases/download/v$qpdfVersion/qpdf-$qpdfVersion-msvc64.zip"
## $qpdfZip     = Join-Path $qpdfDir "qpdf-$qpdfVersion-msvc64.zip"
##
## if (-not (Test-Path $qpdfExe)) {
##     Write-Host "  qpdf not found. Downloading (~23 MB, one-time setup)..." -ForegroundColor Yellow
##     if (-not (Test-Path $qpdfDir)) { New-Item -ItemType Directory -Path $qpdfDir -Force | Out-Null }
##     try {
##         [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
##         $ProgressPreference = 'SilentlyContinue'
##         Invoke-WebRequest -Uri $qpdfZipUrl -OutFile $qpdfZip -UseBasicParsing
##         Write-Host "  Extracting..." -ForegroundColor Yellow
##         Expand-Archive -Path $qpdfZip -DestinationPath $qpdfDir -Force
##         Remove-Item $qpdfZip -Force -ErrorAction SilentlyContinue
##         Write-Host "  qpdf ready." -ForegroundColor Green
##     } catch {
##         Write-Host "  ERROR: Download failed: $($_.Exception.Message)" -ForegroundColor Red
##         Write-Host "  Download manually from: $qpdfZipUrl" -ForegroundColor Yellow
##         Read-Host "Press Enter to exit"
##         exit 1
##     }
## }
##
## if (-not (Test-Path $qpdfExe)) {
##     Write-Host "  ERROR: qpdf.exe not found at $qpdfExe" -ForegroundColor Red
##     Read-Host "Press Enter to exit"
##     exit 1
## }
##
## # ═══════════════════════════════════════════════════════════════════
## #  Ghostscript setup (for page resizing — downloaded only if needed)
## # ═══════════════════════════════════════════════════════════════════
## $gsVersion = "10.07.0"
## $gsTag     = "gs$($gsVersion.Replace('.',''))"
## $gsDir     = Join-Path $env:LOCALAPPDATA "ghostscript\gs$($gsVersion.Replace('.',''))"
## $gsExe     = Join-Path $gsDir "bin\gswin64c.exe"
## $gsInstallerUrl = "https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/$gsTag/${gsTag}w64.exe"
## $gsInstaller    = Join-Path $env:TEMP "${gsTag}w64.exe"
##
## function Ensure-Ghostscript {
##     if (Test-Path $gsExe) { return $true }
##     Write-Host "  Ghostscript not found. Downloading for page resizing (~62 MB, one-time)..." -ForegroundColor Yellow
##     try {
##         [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
##         $ProgressPreference = 'SilentlyContinue'
##         Invoke-WebRequest -Uri $gsInstallerUrl -OutFile $gsInstaller -UseBasicParsing
##         Write-Host "  Installing silently (no admin needed)..." -ForegroundColor Yellow
##         # NSIS silent install to user-writable location
##         $installProc = Start-Process -FilePath $gsInstaller -ArgumentList "/S /D=$gsDir" -PassThru -Wait
##         Remove-Item $gsInstaller -Force -ErrorAction SilentlyContinue
##         if (Test-Path $gsExe) {
##             Write-Host "  Ghostscript ready." -ForegroundColor Green
##         }
##     } catch {
##         Write-Host "  WARNING: Ghostscript download failed. Page resizing skipped." -ForegroundColor Yellow
##         Remove-Item $gsInstaller -Force -ErrorAction SilentlyContinue
##         return $false
##     }
##     return (Test-Path $gsExe)
## }
##
## # US Letter in points: 612 x 792 (tolerance of 2 points)
## function Test-NeedsResize {
##     param([string]$pdfPath)
##     try {
##         $psi = New-Object System.Diagnostics.ProcessStartInfo
##         $psi.FileName = $qpdfExe
##         $psi.Arguments = "--json --json-key=pages `"$pdfPath`""
##         $psi.UseShellExecute = $false
##         $psi.RedirectStandardOutput = $true
##         $psi.RedirectStandardError = $true
##         $psi.CreateNoWindow = $true
##         $proc = [System.Diagnostics.Process]::Start($psi)
##         $jsonOut = $proc.StandardOutput.ReadToEnd()
##         $proc.WaitForExit()
##         $json = $jsonOut | ConvertFrom-Json
##         foreach ($page in $json.pages) {
##             $mb = $page.mediabox
##             if ($null -eq $mb) { continue }
##             $w = [math]::Abs($mb[2] - $mb[0])
##             $h = [math]::Abs($mb[3] - $mb[1])
##             # Check both orientations (portrait and landscape)
##             $isLetterPortrait  = ([math]::Abs($w - 612) -le 2) -and ([math]::Abs($h - 792) -le 2)
##             $isLetterLandscape = ([math]::Abs($w - 792) -le 2) -and ([math]::Abs($h - 612) -le 2)
##             if (-not $isLetterPortrait -and -not $isLetterLandscape) {
##                 return $true
##             }
##         }
##         return $false
##     } catch {
##         return $false
##     }
## }
##
## function Resize-ToLetter {
##     param([string]$inputPath, [string]$outputPath)
##     $psi = New-Object System.Diagnostics.ProcessStartInfo
##     $psi.FileName = $gsExe
##     $psi.Arguments = "-sDEVICE=pdfwrite -sPAPERSIZE=letter -dFIXEDMEDIA -dPDFFitPage -dCompatibilityLevel=1.7 -dNOPAUSE -dQUIET -dBATCH `"-sOutputFile=$outputPath`" `"$inputPath`""
##     $psi.UseShellExecute = $false
##     $psi.RedirectStandardError = $true
##     $psi.CreateNoWindow = $true
##     $proc = [System.Diagnostics.Process]::Start($psi)
##     $stderr = $proc.StandardError.ReadToEnd()
##     $proc.WaitForExit()
##     return @{ ExitCode = $proc.ExitCode; Error = $stderr }
## }
##
## # ═══════════════════════════════════════════════════════════════════
## #  Temp directory for processing
## # ═══════════════════════════════════════════════════════════════════
## $tempDir = Join-Path $env:TEMP "flatten-pdfs-work"
## if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }
##
## # ═══════════════════════════════════════════════════════════════════
## #  HTML page (served from memory)
## # ═══════════════════════════════════════════════════════════════════
## $htmlPage = @'
## <!DOCTYPE html>
## <html lang="en">
## <head>
## <meta charset="UTF-8">
## <meta name="viewport" content="width=device-width, initial-scale=1.0">
## <title>Flatten PDFs</title>
## <style>
##   @import url('https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600&display=swap');
##   * { margin: 0; padding: 0; box-sizing: border-box; }
##   :root {
##     --bg: #1a1a2e; --surface: #16213e; --surface-raised: #1c2a4a;
##     --border: #2a3a5c; --border-active: #4a6fa5;
##     --text: #e0e6f0; --text-dim: #7a8ba8;
##     --accent: #4fc3f7; --accent-glow: rgba(79,195,247,0.15);
##     --success: #66bb6a; --success-bg: rgba(102,187,106,0.1);
##     --error: #ef5350; --error-bg: rgba(239,83,80,0.1);
##     --processing: #ffa726;
##   }
##   body {
##     font-family: 'DM Sans', sans-serif; background: var(--bg); color: var(--text);
##     min-height: 100vh; display: flex; flex-direction: column; align-items: center;
##     padding: 40px 20px;
##     background-image: radial-gradient(ellipse at 20% 50%, rgba(79,195,247,0.04) 0%, transparent 50%),
##                        radial-gradient(ellipse at 80% 20%, rgba(102,187,106,0.03) 0%, transparent 50%);
##   }
##   h1 { font-size: 28px; font-weight: 600; letter-spacing: -0.5px; margin-bottom: 6px; }
##   .subtitle { color: var(--text-dim); font-size: 14px; margin-bottom: 32px; }
##   .app-container { width: 100%; max-width: 640px; }
##   .drop-zone {
##     border: 2px dashed var(--border); border-radius: 16px; padding: 48px 24px;
##     text-align: center; cursor: pointer; transition: all 0.25s ease;
##     background: var(--surface); position: relative; overflow: hidden;
##   }
##   .drop-zone::before {
##     content: ''; position: absolute; inset: 0; background: var(--accent-glow);
##     opacity: 0; transition: opacity 0.25s ease; pointer-events: none;
##   }
##   .drop-zone.drag-over { border-color: var(--accent); transform: scale(1.01); }
##   .drop-zone.drag-over::before { opacity: 1; }
##   .drop-zone:hover { border-color: var(--border-active); }
##   .drop-icon { font-size: 48px; margin-bottom: 12px; display: block; opacity: 0.7; }
##   .drop-text { font-size: 16px; font-weight: 500; margin-bottom: 4px; }
##   .drop-hint { font-size: 13px; color: var(--text-dim); }
##   .file-input { display: none; }
##   .file-list { margin-top: 20px; display: flex; flex-direction: column; gap: 8px; }
##   .file-item {
##     display: flex; align-items: center; gap: 12px; padding: 14px 16px;
##     background: var(--surface); border: 1px solid var(--border); border-radius: 10px;
##     animation: slideIn 0.25s ease;
##   }
##   @keyframes slideIn { from { opacity:0; transform:translateY(-8px); } to { opacity:1; transform:translateY(0); } }
##   .file-item.success { border-color: var(--success); background: var(--success-bg); }
##   .file-item.error { border-color: var(--error); background: var(--error-bg); }
##   .file-icon { font-size: 20px; flex-shrink: 0; width: 32px; text-align: center; }
##   .file-info { flex: 1; min-width: 0; }
##   .file-name { font-size: 14px; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
##   .file-meta { font-size: 12px; color: var(--text-dim); margin-top: 2px; }
##   .file-meta.error-msg { color: var(--error); }
##   .file-actions { flex-shrink: 0; }
##   .download-btn {
##     background: var(--success); color: #1a1a2e; border: none; padding: 7px 14px;
##     border-radius: 6px; font-size: 13px; font-weight: 600; font-family: inherit;
##     cursor: pointer; transition: all 0.15s ease; text-decoration: none; display: inline-block;
##   }
##   .download-btn:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(102,187,106,0.3); }
##   .batch-bar {
##     margin-top: 16px; display: flex; align-items: center; justify-content: space-between;
##     padding: 14px 16px; background: var(--surface-raised); border: 1px solid var(--border);
##     border-radius: 10px; animation: slideIn 0.25s ease; flex-wrap: wrap; gap: 10px;
##   }
##   .batch-bar .summary { font-size: 14px; color: var(--text-dim); }
##   .batch-bar .summary strong { color: var(--text); }
##   .batch-actions { display: flex; gap: 8px; align-items: center; }
##   .open-folder-btn {
##     background: var(--accent); color: #1a1a2e; border: none; padding: 9px 20px;
##     border-radius: 8px; font-size: 14px; font-weight: 600; font-family: inherit;
##     cursor: pointer; transition: all 0.15s ease;
##   }
##   .open-folder-btn:hover { transform: translateY(-1px); box-shadow: 0 4px 16px rgba(79,195,247,0.3); }
##   .clear-btn {
##     background: none; border: 1px solid var(--border); color: var(--text-dim);
##     padding: 9px 16px; border-radius: 8px; font-size: 13px; font-family: inherit;
##     cursor: pointer; transition: all 0.15s ease;
##   }
##   .clear-btn:hover { border-color: var(--text-dim); color: var(--text); }
##   .spinner {
##     width: 20px; height: 20px; border: 2.5px solid var(--border);
##     border-top-color: var(--processing); border-radius: 50%; animation: spin 0.7s linear infinite;
##   }
##   @keyframes spin { to { transform: rotate(360deg); } }
##   .footer { margin-top: 40px; text-align: center; font-size: 12px; color: var(--text-dim); line-height: 1.6; max-width: 440px; }
## </style>
## </head>
## <body>
## <h1>Flatten PDFs</h1>
## <p class="subtitle">Drop PDFs here to flatten form fields, annotations &amp; signatures</p>
## <div class="app-container">
##   <div class="drop-zone" id="dropZone">
##     <span class="drop-icon">&#128196;</span>
##     <div class="drop-text">Drag &amp; drop PDF files here</div>
##     <div class="drop-hint">or click to browse</div>
##     <input type="file" class="file-input" id="fileInput" multiple accept=".pdf">
##   </div>
##   <div class="file-list" id="fileList"></div>
##   <div id="batchBar"></div>
## </div>
## <div class="footer">
##   Powered by qpdf &mdash; files are processed locally and never leave your computer.<br>
##   Forms, annotations, stamps, and signatures are all flattened reliably.
## </div>
## <script>
## const dropZone = document.getElementById('dropZone');
## const fileInput = document.getElementById('fileInput');
## const fileList = document.getElementById('fileList');
## const batchBar = document.getElementById('batchBar');
## let results = [];
## let lastOutputFolder = '';
##
## dropZone.addEventListener('click', () => fileInput.click());
## dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
## dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
## dropZone.addEventListener('drop', e => {
##   e.preventDefault(); dropZone.classList.remove('drag-over');
##   const files = [...e.dataTransfer.files].filter(f => f.name.toLowerCase().endsWith('.pdf'));
##   if (files.length) processFiles(files);
## });
## fileInput.addEventListener('change', () => {
##   const files = [...fileInput.files];
##   if (files.length) processFiles(files);
##   fileInput.value = '';
## });
##
## async function processFiles(files) {
##   for (const file of files) {
##     if (results.some(r => r.name === file.name)) continue;
##     const itemEl = addFileItem(file.name, file.size);
##     try {
##       const formData = new FormData();
##       formData.append('pdf', file);
##       const resp = await fetch('/flatten', { method: 'POST', body: formData });
##       const data = await resp.json();
##       if (data.success) {
##         lastOutputFolder = data.outputFolder || '';
##         results.push({ name: file.name, success: true, origSize: file.size, newSize: data.newSize, downloadUrl: data.downloadUrl });
##         updateFileItem(itemEl, { success: true, name: file.name, origSize: file.size, newSize: data.newSize, downloadUrl: data.downloadUrl });
##       } else {
##         results.push({ name: file.name, success: false, error: data.error });
##         updateFileItem(itemEl, { success: false, name: file.name, error: data.error });
##       }
##     } catch (err) {
##       results.push({ name: file.name, success: false, error: err.message });
##       updateFileItem(itemEl, { success: false, name: file.name, error: err.message });
##     }
##   }
##   updateBatchBar();
## }
##
## function addFileItem(name, size) {
##   const el = document.createElement('div');
##   el.className = 'file-item';
##   el.innerHTML = `<div class="file-icon"><div class="spinner"></div></div>
##     <div class="file-info"><div class="file-name">${esc(name)}</div>
##     <div class="file-meta">Flattening\u2026 (${fmt(size)})</div></div>`;
##   fileList.appendChild(el);
##   return el;
## }
##
## function updateFileItem(el, d) {
##   if (d.success) {
##     el.className = 'file-item success';
##     const pct = d.origSize > 0 ? Math.round((1 - d.newSize / d.origSize) * 100) : 0;
##     const s = pct > 0 ? ` \u00b7 ${pct}% smaller` : pct < 0 ? ` \u00b7 ${Math.abs(pct)}% larger` : '';
##     el.innerHTML = `<div class="file-icon">\u2705</div>
##       <div class="file-info"><div class="file-name">${esc(d.name)}</div>
##       <div class="file-meta">${fmt(d.origSize)} \u2192 ${fmt(d.newSize)}${s}</div></div>
##       <div class="file-actions"><a class="download-btn" href="${d.downloadUrl}" download="${esc(d.name)}">Download</a></div>`;
##   } else {
##     el.className = 'file-item error';
##     el.innerHTML = `<div class="file-icon">\u274c</div>
##       <div class="file-info"><div class="file-name">${esc(d.name)}</div>
##       <div class="file-meta error-msg">${esc(d.error || 'Unknown error')}</div></div>`;
##   }
## }
##
## function updateBatchBar() {
##   const ok = results.filter(r => r.success), fail = results.filter(r => !r.success);
##   if (!ok.length) { batchBar.innerHTML = ''; return; }
##   const tO = ok.reduce((s,r) => s + r.origSize, 0), tN = ok.reduce((s,r) => s + r.newSize, 0);
##   batchBar.innerHTML = `<div class="batch-bar">
##     <div class="summary"><strong>${ok.length}</strong> flattened${fail.length ? `, <strong>${fail.length}</strong> failed` : ''}
##     \u00b7 ${fmt(tO)} \u2192 ${fmt(tN)}</div>
##     <div class="batch-actions">
##     ${lastOutputFolder ? '<button class="open-folder-btn" onclick="openFolder()">Open Output Folder</button>' : ''}
##     <button class="clear-btn" onclick="clearAll()">Clear</button></div></div>`;
## }
##
## async function openFolder() {
##   if (lastOutputFolder) await fetch('/open-folder?path=' + encodeURIComponent(lastOutputFolder));
## }
##
## function clearAll() { results = []; lastOutputFolder = ''; fileList.innerHTML = ''; batchBar.innerHTML = ''; }
## function fmt(b) { return b < 1024 ? b+' B' : b < 1048576 ? (b/1024).toFixed(1)+' KB' : (b/1048576).toFixed(1)+' MB'; }
## function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }
## </script>
## </body>
## </html>
## '@
##
## # ═══════════════════════════════════════════════════════════════════
## #  HTTP Server
## # ═══════════════════════════════════════════════════════════════════
## $port = 8739
## $prefix = "http://localhost:$port/"
##
## # Find an available port
## for ($p = $port; $p -lt $port + 20; $p++) {
##     try {
##         $testListener = New-Object System.Net.HttpListener
##         $testListener.Prefixes.Add("http://localhost:$p/")
##         $testListener.Start()
##         $testListener.Stop()
##         $testListener.Close()
##         $port = $p
##         $prefix = "http://localhost:$port/"
##         break
##     } catch { continue }
## }
##
## $listener = New-Object System.Net.HttpListener
## $listener.Prefixes.Add($prefix)
##
## try {
##     $listener.Start()
## } catch {
##     Write-Host "ERROR: Could not start local server on port $port" -ForegroundColor Red
##     Write-Host $_.Exception.Message -ForegroundColor Red
##     Read-Host "Press Enter to exit"
##     exit 1
## }
##
## Write-Host ""
## Write-Host "  Flatten PDFs is running at $prefix" -ForegroundColor Cyan
## Write-Host "  (Keep this window open while using the tool)" -ForegroundColor Gray
## Write-Host "  Press Ctrl+C to stop." -ForegroundColor Gray
## Write-Host ""
##
## # Open browser
## Start-Process $prefix
##
## # ── Request loop ──────────────────────────────────────────────────
## try {
##     while ($listener.IsListening) {
##         $context  = $listener.GetContext()
##         $request  = $context.Request
##         $response = $context.Response
##
##         $path = $request.Url.AbsolutePath
##
##         if ($path -eq "/" -and $request.HttpMethod -eq "GET") {
##             # Serve the HTML page
##             $buffer = [System.Text.Encoding]::UTF8.GetBytes($htmlPage)
##             $response.ContentType = "text/html; charset=utf-8"
##             $response.ContentLength64 = $buffer.Length
##             $response.OutputStream.Write($buffer, 0, $buffer.Length)
##
##         } elseif ($path -eq "/flatten" -and $request.HttpMethod -eq "POST") {
##             # Parse multipart form data to get the uploaded PDF
##             try {
##                 # Read entire request body as bytes in one pass (stream is not seekable)
##                 $ms = New-Object System.IO.MemoryStream
##                 $request.InputStream.CopyTo($ms)
##                 $allBytes = $ms.ToArray()
##                 $ms.Dispose()
##
##                 $boundary = $request.ContentType.Split('=')[1].Trim()
##
##                 # Parse headers from the UTF-8 representation (only first 2000 bytes for headers)
##                 $headerStr = [System.Text.Encoding]::UTF8.GetString($allBytes, 0, [math]::Min(2000, $allBytes.Length))
##
##                 # Extract filename
##                 $filenameMatch = [regex]::Match($headerStr, 'filename="([^"]+)"')
##                 $filename = if ($filenameMatch.Success) { $filenameMatch.Groups[1].Value } else { "input.pdf" }
##                 # Decode any URL-encoded characters (e.g. %20 for spaces)
##                 $filename = [System.Uri]::UnescapeDataString($filename)
##                 # Strip any path components and sanitize
##                 $filename = [System.IO.Path]::GetFileName($filename)
##                 # Remove any characters illegal in Windows filenames
##                 $illegal = [System.IO.Path]::GetInvalidFileNameChars()
##                 foreach ($ch in $illegal) { $filename = $filename.Replace([string]$ch, '_') }
##                 if ([string]::IsNullOrWhiteSpace($filename)) { $filename = "input.pdf" }
##
##                 # Find end of MIME headers (blank line = 0D 0A 0D 0A)
##                 $headerEndPattern = [byte[]]@(13, 10, 13, 10)
##                 $startIdx = -1
##                 for ($i = 0; $i -lt [math]::Min($allBytes.Length - 4, 2000); $i++) {
##                     if ($allBytes[$i] -eq 13 -and $allBytes[$i+1] -eq 10 -and $allBytes[$i+2] -eq 13 -and $allBytes[$i+3] -eq 10) {
##                         $startIdx = $i + 4
##                         break
##                     }
##                 }
##                 if ($startIdx -lt 0) { throw "Could not find end of MIME headers" }
##
##                 # Find closing boundary from the end (search backwards for CR LF -- boundary)
##                 $closingBytes = [System.Text.Encoding]::ASCII.GetBytes("`r`n--$boundary")
##                 $endIdx = $allBytes.Length
##                 for ($i = $allBytes.Length - $closingBytes.Length; $i -ge $startIdx; $i--) {
##                     $match = $true
##                     for ($j = 0; $j -lt $closingBytes.Length; $j++) {
##                         if ($allBytes[$i + $j] -ne $closingBytes[$j]) { $match = $false; break }
##                     }
##                     if ($match) { $endIdx = $i; break }
##                 }
##
##                 $pdfLength = $endIdx - $startIdx
##                 $pdfBytes = New-Object byte[] $pdfLength
##                 [Array]::Copy($allBytes, $startIdx, $pdfBytes, 0, $pdfLength)
##
##                 # Write input PDF to temp
##                 $inputPath  = Join-Path $tempDir "input_$([System.IO.Path]::GetRandomFileName()).pdf"
##                 $outputPath = Join-Path $tempDir "output_$([System.IO.Path]::GetRandomFileName()).pdf"
##                 [System.IO.File]::WriteAllBytes($inputPath, $pdfBytes)
##
##                 # Also save to a persistent output folder
##                 $outputFolder = Join-Path $tempDir "Flattened_PDFs"
##                 if (-not (Test-Path $outputFolder)) { New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null }
##                 $persistPath = Join-Path $outputFolder $filename
##
##                 # Run qpdf
##                 $psi = New-Object System.Diagnostics.ProcessStartInfo
##                 $psi.FileName = $qpdfExe
##                 $psi.Arguments = "--decrypt --flatten-annotations=all --generate-appearances `"$inputPath`" `"$outputPath`""
##                 $psi.UseShellExecute = $false
##                 $psi.RedirectStandardError = $true
##                 $psi.CreateNoWindow = $true
##                 $proc = [System.Diagnostics.Process]::Start($psi)
##                 $stderr = $proc.StandardError.ReadToEnd()
##                 $proc.WaitForExit()
##
##                 # Clean up input
##                 Remove-Item $inputPath -Force -ErrorAction SilentlyContinue
##
##                 if ((Test-Path $outputPath) -and $proc.ExitCode -le 3) {
##                     # Check if any pages need resizing to Letter
##                     Start-Sleep -Milliseconds 500
##                     $needsResize = $false
##                     try { $needsResize = Test-NeedsResize $outputPath } catch { }
##                     $finalPath = $outputPath
##                     if ($needsResize) {
##                         if (Ensure-Ghostscript) {
##                             Start-Sleep -Milliseconds 500
##                             $resizedPath = Join-Path $tempDir "resized_$([System.IO.Path]::GetRandomFileName()).pdf"
##                             $gsResult = Resize-ToLetter -inputPath $outputPath -outputPath $resizedPath
##                             if ((Test-Path $resizedPath) -and $gsResult.ExitCode -eq 0) {
##                                 # Use the resized file as final output instead of replacing
##                                 $finalPath = $resizedPath
##                             } else {
##                                 Remove-Item $resizedPath -Force -ErrorAction SilentlyContinue
##                             }
##                         }
##                     }
##
##                     # Copy to persistent location (may fail if file is open in a viewer)
##                     try { Copy-Item $finalPath $persistPath -Force } catch { }
##                     $newSize = (Get-Item $finalPath).Length
##
##                     # Serve the download URL as a unique path
##                     $dlId = [System.IO.Path]::GetFileNameWithoutExtension($finalPath)
##                     $jsonResp = @{ success = $true; newSize = $newSize; downloadUrl = "/download/$dlId"; outputFolder = $outputFolder } | ConvertTo-Json
##                 } else {
##                     $errMsg = if ($stderr) { $stderr.Trim().Substring(0, [math]::Min(200, $stderr.Trim().Length)) } else { "qpdf failed with exit code $($proc.ExitCode)" }
##                     $jsonResp = @{ success = $false; error = $errMsg } | ConvertTo-Json
##                     Remove-Item $outputPath -Force -ErrorAction SilentlyContinue
##                 }
##
##             } catch {
##                 $jsonResp = @{ success = $false; error = $_.Exception.Message } | ConvertTo-Json
##             }
##
##             $buffer = [System.Text.Encoding]::UTF8.GetBytes($jsonResp)
##             $response.ContentType = "application/json"
##             $response.ContentLength64 = $buffer.Length
##             $response.OutputStream.Write($buffer, 0, $buffer.Length)
##
##         } elseif ($path.StartsWith("/download/") -and $request.HttpMethod -eq "GET") {
##             # Serve a flattened PDF for download
##             $dlId = $path.Substring("/download/".Length)
##             $dlPath = Join-Path $tempDir "$dlId.pdf"
##             if (Test-Path $dlPath) {
##                 $fileBytes = [System.IO.File]::ReadAllBytes($dlPath)
##                 $response.ContentType = "application/pdf"
##                 $response.ContentLength64 = $fileBytes.Length
##                 $response.OutputStream.Write($fileBytes, 0, $fileBytes.Length)
##             } else {
##                 $response.StatusCode = 404
##                 $buffer = [System.Text.Encoding]::UTF8.GetBytes("File not found")
##                 $response.ContentLength64 = $buffer.Length
##                 $response.OutputStream.Write($buffer, 0, $buffer.Length)
##             }
##
##         } elseif ($path -eq "/open-folder" -and $request.HttpMethod -eq "GET") {
##             # Open a folder in Explorer
##             $folderPath = $request.QueryString["path"]
##             if ($folderPath -and (Test-Path $folderPath)) {
##                 Start-Process explorer.exe $folderPath
##             }
##             $buffer = [System.Text.Encoding]::UTF8.GetBytes('{"ok":true}')
##             $response.ContentType = "application/json"
##             $response.ContentLength64 = $buffer.Length
##             $response.OutputStream.Write($buffer, 0, $buffer.Length)
##
##         } else {
##             $response.StatusCode = 404
##             $buffer = [System.Text.Encoding]::UTF8.GetBytes("Not found")
##             $response.ContentLength64 = $buffer.Length
##             $response.OutputStream.Write($buffer, 0, $buffer.Length)
##         }
##
##         $response.Close()
##     }
## } finally {
##     $listener.Stop()
##     $listener.Close()
##     # Clean up temp files
##     Get-ChildItem -Path $tempDir -Filter "output_*" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
##     Get-ChildItem -Path $tempDir -Filter "input_*" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
## }
# >>>PS_END<<<