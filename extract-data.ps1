param(
  [string]$XlsxPath = "C:\Users\CS407740.PROXIMIDAD\OneDrive - FEMSA Proximidad & Salud\Documentos\02. Pessoal\1. FOLARI\Financeiro\PN - Instituto Folari.xlsx"
)

$ErrorActionPreference = 'Stop'
$nsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$nsRel = 'http://schemas.openxmlformats.org/package/2006/relationships'

function Unzip-Xlsx {
  param([string]$Path)
  $tmp = Join-Path $env:TEMP ("folari_extract_" + [guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Path $tmp | Out-Null
  Copy-Item $Path (Join-Path $tmp 'book.zip')
  Expand-Archive (Join-Path $tmp 'book.zip') -DestinationPath (Join-Path $tmp 'x') -Force
  return (Join-Path $tmp 'x')
}

function Get-SharedStrings {
  param([string]$Base)
  $path = Join-Path $Base 'xl\sharedStrings.xml'
  if (-not (Test-Path $path)) { return @() }
  [xml]$xml = Get-Content $path -Encoding UTF8
  $list = New-Object System.Collections.Generic.List[string]
  foreach ($si in $xml.sst.si) {
    if ($si.t) {
      $list.Add([string]$si.t)
    } else {
      $text = ($si.r | ForEach-Object { $_.t }) -join ''
      $list.Add([string]$text)
    }
  }
  return $list
}

function Parse-CellRef {
  param([string]$Ref)
  if ($Ref -match '^([A-Z]+)(\d+)$') {
    $colLetters = $matches[1]
    $row = [int]$matches[2]
    $col = 0
    foreach ($ch in $colLetters.ToCharArray()) {
      $col = $col * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return @{ Row = $row; Col = $col }
  }
  return $null
}

function Get-CellValue {
  param($Cell, $SharedStrings)
  if ($null -eq $Cell) { return $null }
  $type = $Cell.t
  $raw = $Cell.v
  if ($null -eq $raw -or $raw -eq '') { return $null }
  if ($type -eq 's') {
    return $SharedStrings[[int]$raw]
  }
  if ($type -eq 'b') {
    return [bool][int]$raw
  }
  if ($type -eq 'str') {
    return [string]$raw
  }
  if ($raw -match '^-?\d+(\.\d+)?$') {
    return [double]$raw
  }
  return [string]$raw
}

function Read-SheetGrid {
  param([string]$Base, [string]$SheetFile, [string[]]$SharedStrings)
  $path = Join-Path $Base "xl\worksheets\$SheetFile"
  [xml]$xml = Get-Content $path -Encoding UTF8
  $grid = @{}
  foreach ($row in $xml.worksheet.sheetData.row) {
    foreach ($cell in $row.c) {
      $ref = Parse-CellRef $cell.r
      if ($null -eq $ref) { continue }
      $grid["$($ref.Row):$($ref.Col)"] = Get-CellValue $cell $SharedStrings
    }
  }
  return $grid
}

function Get-GridValue {
  param($Grid, [int]$Row, [int]$Col)
  $key = "$Row`:$Col"
  if ($Grid.ContainsKey($key)) { return $Grid[$key] }
  return $null
}

function Normalize-YesNo {
  param($Value)
  if ($null -eq $Value) { return $null }
  $s = ([string]$Value).Trim().ToLowerInvariant()
  if ($s -in @('sim','s','yes','true','1')) { return 'Sim' }
  if ($s -in @('nao','não','n','no','false','0')) { return 'Não' }
  return [string]$Value
}

function Parse-ExcelDate {
  param($Value)
  if ($null -eq $Value) { return $null }
  if ($Value -is [double]) {
    return [datetime]::FromOADate($Value).Date
  }
  $s = [string]$Value
  if ($s -match '^\d{4}-\d{2}-\d{2}') {
    return [datetime]::Parse($s).Date
  }
  $formats = @('d/M/yyyy','dd/MM/yyyy','M/d/yyyy','MM/dd/yyyy')
  foreach ($f in $formats) {
    try { return [datetime]::ParseExact($s, $f, [Globalization.CultureInfo]::InvariantCulture).Date } catch {}
  }
  try { return [datetime]::Parse($s, [Globalization.CultureInfo]::GetCultureInfo('pt-BR')).Date } catch { return $null }
}

function To-IsoDate { param([datetime]$Date) return $Date.ToString('yyyy-MM-dd') }
function To-IsoMonth { param([datetime]$Date) return $Date.ToString('yyyy-MM') }
function Monday-Of { param([datetime]$Date)
  $dow = ([int]$Date.DayOfWeek + 6) % 7
  return $Date.AddDays(-$dow).Date
}

$base = Unzip-Xlsx $XlsxPath
$shared = Get-SharedStrings $base
$funilGrid = Read-SheetGrid $base 'sheet5.xml' $shared
$fluxoGrid = Read-SheetGrid $base 'sheet9.xml' $shared

# Headers are on row 2 in Funil_Vendas
$headers = @{}
for ($c = 1; $c -le 40; $c++) {
  $h = Get-GridValue $funilGrid 2 $c
  if ($null -ne $h -and [string]$h -ne '' -and [string]$h -notmatch '^\d+(\.\d+)?$') {
    $headers[[string]$h] = $c
  }
}

"HEADERS:"
$headers.GetEnumerator() | Sort-Object Value | ForEach-Object { "$($_.Value): $($_.Key)" }

$colDataConsulta = $headers['Data Consulta']
$colCompareceu = $headers['Compareceu Consulta?']
$colProposta = $headers['Possui Proposta?']
$colPropostaPaga = $headers['Tratamento Pago?']
if (-not $colPropostaPaga) { $colPropostaPaga = $headers['Proposta Paga?'] }
if (-not $colPropostaPaga) { $colPropostaPaga = $headers['Proposta paga?'] }

$rows = @()
for ($r = 3; $r -le 200; $r++) {
  $nome = Get-GridValue $funilGrid $r 1
  $dataConsulta = Get-GridValue $funilGrid $r $colDataConsulta
  if ($null -eq $nome -and $null -eq $dataConsulta) { continue }
  if ([string]$nome -match '^(TOTAL|Total|Tabela|Dados )') { continue }
  $rows += [pscustomobject]@{
    Row = $r
    Nome = $nome
    DataConsulta = Parse-ExcelDate $dataConsulta
    Compareceu = Normalize-YesNo (Get-GridValue $funilGrid $r $colCompareceu)
    Proposta = Normalize-YesNo (Get-GridValue $funilGrid $r $colProposta)
    PropostaPaga = Normalize-YesNo (Get-GridValue $funilGrid $r $colPropostaPaga)
  }
}

$agendados = $rows.Count
$compareceram = @($rows | Where-Object { $_.Compareceu -eq 'Sim' }).Count
$proposta = @($rows | Where-Object { $_.Proposta -eq 'Sim' }).Count
$propostaPaga = @($rows | Where-Object { $_.PropostaPaga -eq 'Sim' }).Count

"`nFUNIL:"
"agendados=$agendados compareceram=$compareceram proposta=$proposta propostaPaga=$propostaPaga rows=$($rows.Count)"

$daily = @{}
foreach ($row in $rows) {
  if ($null -eq $row.DataConsulta) { continue }
  $key = To-IsoDate $row.DataConsulta
  if (-not $daily.ContainsKey($key)) { $daily[$key] = 0 }
  $daily[$key]++
}
$dailySeries = $daily.GetEnumerator() | Sort-Object Name | ForEach-Object { ,@($_.Name, $_.Value) }

$weekly = @{}
foreach ($row in $rows) {
  if ($null -eq $row.DataConsulta) { continue }
  $mon = Monday-Of $row.DataConsulta
  $key = To-IsoDate $mon
  if (-not $weekly.ContainsKey($key)) { $weekly[$key] = 0 }
  $weekly[$key]++
}
$weeklySeries = $weekly.GetEnumerator() | Sort-Object Name | ForEach-Object { ,@($_.Name, $_.Value) }

$monthly = @{}
foreach ($row in $rows) {
  if ($null -eq $row.DataConsulta) { continue }
  $key = To-IsoMonth $row.DataConsulta
  if (-not $monthly.ContainsKey($key)) { $monthly[$key] = 0 }
  $monthly[$key]++
}
$monthlySeries = $monthly.GetEnumerator() | Sort-Object Name | ForEach-Object { ,@($_.Name, $_.Value) }

"`nCONSULTAS daily=$($dailySeries.Count) weekly=$($weeklySeries.Count) monthly=$($monthlySeries.Count)"

# Fluxo_Caixa headers (linha 1)
$fluxoHeaders = @{}
for ($c = 1; $c -le 12; $c++) {
  $h = Get-GridValue $fluxoGrid 1 $c
  if ($null -ne $h -and [string]$h -ne '' -and [string]$h -notmatch '^\d+(\.\d+)?$') {
    $fluxoHeaders[[string]$h] = $c
  }
}

"`nFLUXO HEADERS:"
$fluxoHeaders.GetEnumerator() | Sort-Object Value | ForEach-Object { "$($_.Value): $($_.Key)" }

$colValor = $fluxoHeaders['Valor Pagamento']
if (-not $colValor) { $colValor = $fluxoHeaders['Valor'] }
$colReceb = $fluxoHeaders['Data Recebimento']
if (-not $colReceb) { $colReceb = $fluxoHeaders['Recebimento'] }
if (-not $colReceb) { $colReceb = $fluxoHeaders['Data de Recebimento'] }

$today = Get-Date
$totalFaturamento = 0.0
$totalCaixa = 0.0
$aReceber = 0.0
$totalDespesas = 0.0
$recebMensais = @{}
$despMensais = @{}
$despCats = @{}

for ($r = 2; $r -le 500; $r++) {
  $nomeFluxo = Get-GridValue $fluxoGrid $r 1
  $val = Get-GridValue $fluxoGrid $r $colValor
  if ($null -eq $val -or $val -eq '') { continue }
  if ($null -eq $nomeFluxo -or [string]$nomeFluxo -eq '') { continue }
  if ([string]$nomeFluxo -match '^(TOTAL|Total|Subtotal|Resumo)') { continue }
  $amount = [double]$val
  $recvDate = Parse-ExcelDate (Get-GridValue $fluxoGrid $r $colReceb)
  if ($null -eq $recvDate) { continue }
  $monthKey = To-IsoMonth $recvDate
  if ($amount -gt 0) {
    $totalFaturamento += $amount
    if (-not $recebMensais.ContainsKey($monthKey)) { $recebMensais[$monthKey] = 0.0 }
    $recebMensais[$monthKey] += $amount
    if ($recvDate.Date -le $today.Date) { $totalCaixa += $amount }
    else { $aReceber += $amount }
  } elseif ($amount -lt 0) {
    $abs = [math]::Abs($amount)
    $totalDespesas += $abs
    if (-not $despMensais.ContainsKey($monthKey)) { $despMensais[$monthKey] = 0.0 }
    $despMensais[$monthKey] += $abs
    $catName = [string]$nomeFluxo
    if (-not $despCats.ContainsKey($catName)) { $despCats[$catName] = 0.0 }
    $despCats[$catName] += $abs
  }
}

$recebSeries = $recebMensais.GetEnumerator() | Sort-Object Name | ForEach-Object { ,@($_.Name, [math]::Round($_.Value)) }
$despSeries = $despMensais.GetEnumerator() | Sort-Object Name | ForEach-Object { ,@($_.Name, [math]::Round($_.Value)) }
$despCatSeries = $despCats.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { ,@($_.Name, [math]::Round($_.Value)) }

"`nFINANCEIRO totalFaturamento=$([math]::Round($totalFaturamento)) totalCaixa=$([math]::Round($totalCaixa)) aReceber=$([math]::Round($aReceber)) totalDespesas=$([math]::Round($totalDespesas))"
"recebimentosMensais count=$($recebSeries.Count) despesasMensais count=$($despSeries.Count)"

$geradoEm = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

$dataObj = [ordered]@{
  geradoEm = $geradoEm
  funil = [ordered]@{
    agendados = $agendados
    compareceram = $compareceram
    proposta = $proposta
    propostaPaga = $propostaPaga
  }
  consultas = [ordered]@{
    daily = $dailySeries
    weekly = $weeklySeries
    monthly = $monthlySeries
  }
  financeiro = [ordered]@{
    totalFaturamento = [math]::Round($totalFaturamento)
    totalCaixa = [math]::Round($totalCaixa)
    aReceber = [math]::Round($aReceber)
    totalDespesas = [math]::Round($totalDespesas)
    despesasPorCategoria = $despCatSeries
    despesasMensais = $despSeries
  }
  recebimentosMensais = $recebSeries
}

$jsonPath = Join-Path $PSScriptRoot 'data-snapshot.json'
$dataObj | ConvertTo-Json -Depth 6 | Set-Content $jsonPath -Encoding UTF8
"JSON written to $jsonPath"
