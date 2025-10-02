param(
  [Parameter(Mandatory=$true)] [string] $ExtractRoot,
  [Parameter(Mandatory=$true)] [string] $OutFile
)

$ErrorActionPreference = 'Stop'

# Najdi model (model.bim nebo model.tmsl.json)
$modelJson = Get-ChildItem -Path $ExtractRoot -Recurse -Filter model.* |
  Where-Object { $_.Name -match 'model\.(bim|tmsl\.json)$' } |
  Select-Object -First 1
if (-not $modelJson) { throw "Model file not found under $ExtractRoot" }

$model = Get-Content $modelJson.FullName -Raw | ConvertFrom-Json

# Power Query (M) – volitelné
$powerQueryPath = Join-Path $ExtractRoot 'DataMashup/PowerQueryFormulas/Section1.m'
$hasM = Test-Path $powerQueryPath

function Esc($s){ if($null -eq $s){return ''} return ($s -replace '\|','\\|') }

New-Item -ItemType Directory -Force -Path (Split-Path $OutFile) | Out-Null

$md = @()
$md += "# Model Documentation"
$md += "`nGenerated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
$md += "`nSource: $($modelJson.FullName)`n"

$tables = @($model.model.tables)
$rels   = @($model.model.relationships)
$roles  = @($model.model.roles)
$partitions = @()
foreach($t in $tables){ if($t.partitions){ $partitions += $t.partitions } }

$md += "`n## Summary"
$md += "- Tables: $($tables.Count)"
$md += "- Relationships: $($rels.Count)"
$md += "- Roles: $($roles.Count)"
$md += "- Partitions: $($partitions.Count)`n"

$md += "`n## Tables"
foreach($t in $tables){
  $md += "`n### `$(Esc $t.name)`"
  if($t.description){ $md += "> $(($t.description -replace '\r?\n',' '))`n" }

  $cols = @($t.columns)
  if($cols.Count -gt 0){
    $md += "`n**Columns**"
    $md += "`n| Name | Data type | Format | Is Hidden | Description |"
    $md += "`n|---|---|---|---:|---|"
    foreach($c in $cols){
      $md += "`n| $(Esc $c.name) | $(Esc $c.dataType) | $(Esc $c.formatString) | $([bool]$c.isHidden) | $(Esc $c.description) |"
    }
  }

  $measures = @($t.measures)
  if($measures.Count -gt 0){
    $md += "`n`n**Measures (DAX)**`n"
    foreach($m in $measures){
      $expr = ($m.expression -replace '\r?\n',"`n    ")
      $md += "`n- **$(Esc $m.name)**`n`n    ```DAX`n    $expr`n    ````n"
      if($m.description){ $md += "`n    _$(Esc $m.description)_`n" }
    }
  }

  $hiers = @($t.hierarchies)
  if($hiers.Count -gt 0){
    $md += "`n`n**Hierarchies**"
    foreach($h in $hiers){
      $levels = ($h.levels | ForEach-Object { $_.name }) -join ' › '
      $md += "`n- $(Esc $h.name): $levels"
    }
  }

  $tparts = @($t.partitions)
  if($tparts.Count -gt 0){
    $md += "`n`n**Partitions**"
    $md += "`n| Name | Mode | Source |"
    $md += "`n|---|---|---|"
    foreach($p in $tparts){
      $src = $p.source.type
      if($p.source.expression){ $src = "$src (M)" }
      $md += "`n| $(Esc $p.name) | $(Esc $p.mode) | $(Esc $src) |"
    }
  }
}

if($rels.Count -gt 0){
  $md += "`n`n## Relationships"
  $md += "`n| From | To | Cardinality | CrossFilter | IsActive |"
  $md += "`n|---|---|---|---|---:|"
  foreach($r in $rels){
    $from = "$(Esc $r.fromTable).$(Esc $r.fromColumn)"
    $to   = "$(Esc $r.toTable).$(Esc $r.toColumn)"
    $md += "`n| $from | $to | $(Esc $r.cardinality) | $(Esc $r.crossFilteringBehavior) | $([bool]$r.isActive) |"
  }
}

if($roles.Count -gt 0){
  $md += "`n`n## Roles"
  foreach($role in $roles){
    $md += "`n### $(Esc $role.name)"
    if($role.description){ $md += "`n$(Esc $role.description)" }
    $md += "`nMembers: $(@($role.members).Count)"
    if($role.tablePermissions){
      $md += "`n**Table filters**"
      foreach($tp in $role.tablePermissions){
        $md += "`n- $(Esc $tp.table) → `$(Esc $tp.filterExpression)`"
      }
    }
  }
}

if($hasM){
  $md += "`n`n## Power Query (M)"
  $mText = Get-Content $powerQueryPath -Raw
  $lines = ($mText -split "\r?\n")
  $snippet = $lines[0..([Math]::Min(60, $lines.Count-1))] -join "`n"
  $md += "`n> Showing first ~60 lines of Section1.m; full file saved as `docs/Section1.m`"
  $md += "`n```m`n$snippet`n````n"
  New-Item -ItemType Directory -Force -Path (Split-Path $OutFile) | Out-Null
  Set-Content -Path (Join-Path (Split-Path $OutFile) 'Section1.m') -Value $mText -Encoding UTF8
}

$mdText = ($md -join "`n")
Set-Content -Path $OutFile -Value $mdText -Encoding UTF8
Write-Host "Documentation written to $OutFile"
