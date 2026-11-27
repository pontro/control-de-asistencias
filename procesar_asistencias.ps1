param(
    [string]$RutaEntrada = "",
    [string]$RutaSalidaDetalle = "",
    [string]$RutaSalidaResumen = "",
    [int]$TotalEntradasEsperadas = 18,
    [int]$TotalSalidasEsperadas = 17,
    [string]$HoraLimiteEntrada = "07:35",
    [switch]$ExcluirFinesDeSemana = $false
)

$directorio = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($RutaEntrada)) { $RutaEntrada = Join-Path $directorio "asistencias_crudo.csv" }
if ([string]::IsNullOrWhiteSpace($RutaSalidaDetalle)) { $RutaSalidaDetalle = Join-Path $directorio "asistencias_detalle.csv" }
if ([string]::IsNullOrWhiteSpace($RutaSalidaResumen)) { $RutaSalidaResumen = Join-Path $directorio "asistencias_resumen.csv" }

$hlParts = $HoraLimiteEntrada -split ":"
$hlH = [int]$hlParts[0]
$hlM = if ($hlParts.Count -ge 2) { [int]$hlParts[1] } else { 0 }
$limiteSegundos = ($hlH * 3600) + ($hlM * 60)

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   PROCESADOR DE ASISTENCIAS - ISAAC NEWTON               " -ForegroundColor Cyan
Write-Host "   Base Oficial: $TotalEntradasEsperadas Entradas / $TotalSalidasEsperadas Salidas (Hora Limite: $HoraLimiteEntrada)" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

if (-not (Test-Path $RutaEntrada)) {
    Write-Error "No se encontró el archivo de entrada: $RutaEntrada"
    exit 1
}

# Leer líneas del archivo
$lineas = Get-Content -Path $RutaEntrada -Encoding UTF8
$registrosCrudos = @()

foreach ($linea in $lineas) {
    if ([string]::IsNullOrWhiteSpace($linea)) { continue }
    
    $partes = $linea -split ","
    if ($partes.Count -lt 3) { continue }
    
    $nombre = $partes[0].Trim().Trim('"')
    if ($nombre -ieq "Nombre" -or $nombre -eq "") { continue }
    
    # Detección dinámica de columna de fecha
    $fechaIdx = -1
    for ($i = 1; $i -lt $partes.Count; $i++) {
        if ($partes[$i] -match "\b\d{1,2}/\d{1,2}/\d{4}\b") {
            $fechaIdx = $i
            break
        }
    }
    if ($fechaIdx -eq -1) { continue }
    
    # Detección dinámica de columna de hora
    $horaIdx = -1
    for ($i = 1; $i -lt $partes.Count; $i++) {
        if ($i -ne $fechaIdx -and $partes[$i] -match "\b\d{1,2}:\d{2}") {
            $horaIdx = $i
            break
        }
    }
    if ($horaIdx -eq -1) { continue }
    
    # Detección dinámica de AM/PM
    $ampmIdx = -1
    for ($i = 1; $i -lt $partes.Count; $i++) {
        if ($i -ne $fechaIdx -and $i -ne $horaIdx -and $partes[$i].Trim().Trim('"') -match "^(a\.?\s*m\.?|p\.?\s*m\.?|am|pm)$") {
            $ampmIdx = $i
            break
        }
    }
    
    $diaStr = $partes[$fechaIdx].Trim().Trim('"')
    $horaStr = $partes[$horaIdx].Trim().Trim('"')
    $periodo = if ($ampmIdx -ne -1) { $partes[$ampmIdx].Trim().Trim('"') } else { "" }
    
    $notasArr = @()
    for ($i = 1; $i -lt $partes.Count; $i++) {
        if ($i -ne $fechaIdx -and $i -ne $horaIdx -and $i -ne $ampmIdx) {
            $val = $partes[$i].Trim().Trim('"')
            if (-not [string]::IsNullOrWhiteSpace($val) -and $val -notmatch "^(?i)(new organization|general|default)$") {
                $notasArr += $val
            }
        }
    }
    $nota = $notasArr -join ", "
    
    $fechaSolo = ($diaStr -split " ")[0]
    
    $esPM = ($periodo -match "p\.?\s*m\.?" -or $horaStr -match "p\.?\s*m\.?" -or $horaStr -match "(?i)\bpm\b" -or $periodo -match "(?i)\bpm\b")
    $esAM = ($periodo -match "a\.?\s*m\.?" -or $horaStr -match "a\.?\s*m\.?" -or $horaStr -match "(?i)\bam\b" -or $periodo -match "(?i)\bam\b")

    
    $horaLimpia = ($horaStr -replace "[^\d:]", "").Trim()
    $horaParts = $horaLimpia -split ":"
    if ($horaParts.Count -ge 2) {
        $h = [int]$horaParts[0]
        $m = [int]$horaParts[1]
        $s = if ($horaParts.Count -ge 3) { [int]$horaParts[2] } else { 0 }
        
        if ($esPM -and $h -lt 12) { $h += 12 }
        if ($esAM -and $h -eq 12) { $h = 0 }
        
        $totalSegundos = ($h * 3600) + ($m * 60) + $s
        $hora24Str = "{0:D2}:{1:D2}:{2:D2}" -f $h, $m, $s
        
        # Validar si es fin de semana
        $fp = $fechaSolo -split "/"
        $esSabadoODomingo = $false
        if ($fp.Count -eq 3) {
            $dt = Get-Date -Year ([int]$fp[2]) -Month ([int]$fp[1]) -Day ([int]$fp[0])
            if ($dt.DayOfWeek -in "Saturday", "Sunday") {
                $esSabadoODomingo = $true
            }
        }

        if ($ExcluirFinesDeSemana -and $esSabadoODomingo) {
            continue
        }

        $registrosCrudos += [PSCustomObject]@{
            Nombre = $nombre
            Fecha = $fechaSolo
            Hora24 = $hora24Str
            Segundos = $totalSegundos
            Nota = $nota
            EsFinDeSemana = $esSabadoODomingo
        }
    }
}

Write-Host "Total de registros leídos: $($registrosCrudos.Count)" -ForegroundColor Green

# Agrupar por Nombre y Fecha
$grupos = $registrosCrudos | Group-Object -Property Nombre, Fecha
$detalles = @()

foreach ($g in $grupos) {
    $items = $g.Group | Sort-Object Segundos
    $nombre = $items[0].Nombre
    $fecha = $items[0].Fecha
    $esFinDeSemana = $items[0].EsFinDeSemana
    
    $notasUnidas = ($items | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Nota) } | ForEach-Object { $_.Nota } | Select-Object -Unique) -join "; "
    
    $esRetraso = $false
    if ($items.Count -ge 2) {
        $primero = $items[0]
        $ultimo = $items[-1]
        
        if ($ultimo.Segundos - $primero.Segundos -lt 60) {
            if ($primero.Segundos -lt (12 * 3600)) {
                $horaEntrada = $primero.Hora24
                $horaSalida = ""
                $esRetraso = $primero.Segundos -gt $limiteSegundos
                $estado = if ($esRetraso) { "Retraso" } else { "Solo Entrada (Marcaje Duplicado)" }
            } else {
                $horaEntrada = ""
                $horaSalida = $ultimo.Hora24
                $estado = "Solo Salida (Marcaje Duplicado)"
            }
        } else {
            $horaEntrada = $primero.Hora24
            $horaSalida = $ultimo.Hora24
            $esRetraso = $primero.Segundos -gt $limiteSegundos
            $estado = if ($esRetraso) { "Retraso" } else { "Completo" }
        }
    } else {
        $unica = $items[0]
        if ($unica.Segundos -lt (12 * 3600)) {
            $horaEntrada = $unica.Hora24
            $horaSalida = ""
            $esRetraso = $unica.Segundos -gt $limiteSegundos
            $estado = if ($esRetraso) { "Retraso" } else { "Solo Entrada (Falta Salida)" }
        } else {
            $horaEntrada = ""
            $horaSalida = $unica.Hora24
            $estado = "Solo Salida (Falta Entrada)"
        }
    }

    if ($esFinDeSemana) {
        $estado += " [Fin de Semana]"
    }
    
    $fParts = $fecha -split "/"
    $fechaSort = "{0}-{1:D2}-{2:D2}" -f $fParts[2], [int]$fParts[1], [int]$fParts[0]

    $detalles += [PSCustomObject]@{
        Nombre = $nombre
        Fecha = $fecha
        FechaSort = $fechaSort
        HoraEntrada = $horaEntrada
        HoraSalida = $horaSalida
        Estado = $estado
        EsRetraso = $esRetraso
        Observaciones = $notasUnidas
    }
}

$detalles = $detalles | Sort-Object Nombre, FechaSort


# Generar Resumen por Persona
$resumen = @()
$gruposPersona = $detalles | Group-Object -Property Nombre

foreach ($gp in $gruposPersona) {
    $pNombre = $gp.Name
    $pItems = $gp.Group
    
    $diasAsistidos = @($pItems).Count
    
    $entradasCount = @($pItems | Where-Object { -not [string]::IsNullOrWhiteSpace($_.HoraEntrada) }).Count
    $salidasCount = @($pItems | Where-Object { -not [string]::IsNullOrWhiteSpace($_.HoraSalida) }).Count
    $retrasosCount = @($pItems | Where-Object { $_.EsRetraso -eq $true }).Count
    
    $faltasEntrada = [Math]::Max(0, $TotalEntradasEsperadas - $entradasCount)
    $faltasSalida = [Math]::Max(0, $TotalSalidasEsperadas - $salidasCount)
    $totalFaltas = $faltasEntrada + $faltasSalida + $retrasosCount
    
    # Ratios contra la Base Oficial
    $entradasRatio = "$entradasCount / $TotalEntradasEsperadas"
    $salidasRatio = "$salidasCount / $TotalSalidasEsperadas"
    
    $resumen += [PSCustomObject]@{
        "Empleado" = $pNombre
        "Entradas" = $entradasRatio
        "Salidas" = $salidasRatio
        "Días Asistidos" = $diasAsistidos
        "Retrasos" = $retrasosCount
        "Faltas Totales" = $totalFaltas
    }
}




# Exportar
$detallesExport = $detalles | Select-Object Nombre, Fecha, HoraEntrada, HoraSalida, Estado, Observaciones
$detallesExport | Export-Csv -Path $RutaSalidaDetalle -NoTypeInformation -Encoding UTF8
$resumen | Export-Csv -Path $RutaSalidaResumen -NoTypeInformation -Encoding UTF8


Write-Host "`n✅ ¡Procesamiento completado con éxito!" -ForegroundColor Green
Write-Host "   -> Detalle generado en: $RutaSalidaDetalle" -ForegroundColor Yellow
Write-Host "   -> Resumen generado en: $RutaSalidaResumen" -ForegroundColor Yellow
