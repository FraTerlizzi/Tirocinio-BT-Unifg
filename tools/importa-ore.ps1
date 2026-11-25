# Rigenera dati.js dal registro ore (foglio "I anno": COGNOME|NOME|MATRICOLA|ORE; foglio "Terzo Anno": N.|STUDENTE|ORE).
# I nomi diventano "Cognome N." e le matricole vengono oscurate (es. 6342**): i dati completi non entrano mai nel repository.
# Uso:  powershell -File tools\importa-ore.ps1 -Origine "C:\percorso\ore_tirocinio_studenti.xlsx"
param([Parameter(Mandatory)][string]$Origine)

$dest = Join-Path $PSScriptRoot "..\dati.js"
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
$xl.AutomationSecurity = 3   # macro disattivate
$ti = (Get-Culture).TextInfo
function Mask($m) { $s = [string]$m; if (-not $s) { return "" }; if ($s.Length -ge 6) { return $s.Substring(0,4) + "**" }; return $s.Substring(0,[Math]::Min(2,$s.Length)) + "**" }
function Ore($v) { $s = [string]$v
  if ($s -match '^\d+(\.\d+)?$') { return @{ ore = [int][double]$s } }
  if ($s -match 'erasmus') { return @{ ore = $null; erasmus = ($s -replace '(?i)erasmus\s*','').Trim() } }
  return @{ ore = $null } }

$out = @()
try {
  $wb = $xl.Workbooks.Open($Origine, 0, $true)
  $v = $wb.Worksheets.Item("I anno").UsedRange.Value2
  for ($r = 2; $r -le $v.GetLength(0); $r++) {
    $cog = [string]$v[$r,1]; $nom = [string]$v[$r,2]
    if (-not $nom -or $cog -like 'TOT*') { continue }
    if ($cog -eq 'True') { $cog = 'VERO' }   # Excel converte il cognome "Vero" in booleano
    $o = Ore $v[$r,4]
    $rec = [ordered]@{ n = $ti.ToTitleCase($cog.ToLower().Trim()) + " " + $nom.Trim().Substring(0,1).ToUpper() + "."; m = Mask $v[$r,3]; anno = 1; ore = $o.ore }
    if ($o.erasmus) { $rec.erasmus = $o.erasmus }
    $out += [pscustomobject]$rec
  }
  $v = $wb.Worksheets.Item("Terzo Anno").UsedRange.Value2
  for ($r = 2; $r -le $v.GetLength(0); $r++) {
    $n = [string]$v[$r,2]; if (-not $n) { continue }
    $o = Ore $v[$r,3]
    $out += [pscustomobject][ordered]@{ n = $n.Trim(); m = ""; anno = 3; ore = $o.ore }
  }
  $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }

$js = "// Studenti reali dal registro ore (I e III anno). Nomi ridotti a cognome e iniziale, matricole oscurate.`n// Generato da tools/importa-ore.ps1: non modificare a mano.`nwindow.STUDENTI_REALI = " + ($out | ConvertTo-Json -Compress) + ";`n"
[IO.File]::WriteAllText($dest, $js, (New-Object Text.UTF8Encoding $false))
"Scritti $($out.Count) studenti in $dest"
