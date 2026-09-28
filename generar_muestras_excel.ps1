$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$baseDir = "c:\Users\Nando Pontro\Desktop\Projects\isaacnewton"

# ----------------------------------------------------
# 1. FORMATO 1: 5 COLUMNAS (Nombre, Fecha, Hora, Periodo, Observaciones)
# ----------------------------------------------------
$wb1 = $excel.Workbooks.Add()
$ws1 = $wb1.Worksheets.Item(1)
$ws1.Name = "Asistencias_5_Cols"

$data1 = @(
    @("Nombre", "Fecha", "Hora", "Periodo", "Observaciones"),
    @("Carlos Mendoza", "01/09/2026", "07:45:00", "a. m.", "Entrada puntual"),
    @("Carlos Mendoza", "01/09/2026", "06:10:00", "p. m.", "Salida regular"),
    @("Carlos Mendoza", "02/09/2026", "07:50:00", "a. m.", "Entrada regular"),
    @("Carlos Mendoza", "02/09/2026", "06:05:00", "p. m.", "Salida regular"),
    @("Carlos Mendoza", "03/09/2026", "07:48:00", "a. m.", "Solo vino en la mañana (falta salida)"),
    @("Lucia Fernandez", "01/09/2026", "07:40:00", "a. m.", "Entrada puntual"),
    @("Lucia Fernandez", "01/09/2026", "06:30:00", "p. m.", "Salida regular"),
    @("Lucia Fernandez", "02/09/2026", "06:15:00", "p. m.", "Solo marcó en la tarde (falta entrada)"),
    @("Lucia Fernandez", "03/09/2026", "07:35:00", "a. m.", "Entrada puntual"),
    @("Lucia Fernandez", "03/09/2026", "06:00:00", "p. m.", "Salida regular"),
    @("Roberto Gomez", "01/09/2026", "07:55:00", "a. m.", "Entrada normal"),
    @("Roberto Gomez", "01/09/2026", "07:55:30", "a. m.", "Doble marcaje entrada (deduplicar)"),
    @("Roberto Gomez", "01/09/2026", "05:45:00", "p. m.", "Salida regular"),
    @("Roberto Gomez", "02/09/2026", "07:50:00", "a. m.", "Entrada normal"),
    @("Roberto Gomez", "02/09/2026", "05:50:00", "p. m.", "Salida normal"),
    @("Mariana Rios", "01/09/2026", "08:05:00", "a. m.", "Entrada con retardo"),
    @("Mariana Rios", "01/09/2026", "06:00:00", "p. m.", "Salida regular"),
    @("Mariana Rios", "02/09/2026", "08:00:00", "a. m.", "Entrada normal"),
    @("Mariana Rios", "02/09/2026", "06:15:00", "p. m.", "Salida normal"),
    @("Mariana Rios", "03/09/2026", "07:55:00", "a. m.", "Entrada normal"),
    @("Mariana Rios", "03/09/2026", "06:10:00", "p. m.", "Salida normal")
)

for ($r = 0; $r -lt $data1.Length; $r++) {
    for ($c = 0; $c -lt $data1[$r].Length; $c++) {
        $ws1.Cells.Item($r + 1, $c + 1).Value2 = $data1[$r][$c]
    }
}
$ws1.UsedRange.EntireColumn.AutoFit() | Out-Null
$path1 = Join-Path $baseDir "ejemplo_formato_5_columnas.xlsx"
$wb1.SaveAs($path1, 51)
$wb1.Close()

# ----------------------------------------------------
# 2. FORMATO 2: 4 COLUMNAS CON HORA Y PERIODO UNIFICADOS
# ----------------------------------------------------
$wb2 = $excel.Workbooks.Add()
$ws2 = $wb2.Worksheets.Item(1)
$ws2.Name = "Asistencias_4_Cols_Periodo"

$data2 = @(
    @("Nombre", "Fecha", "Hora y Periodo", "Observaciones"),
    @("Carlos Mendoza", "01/09/2026", "07:45:00 a. m.", "Entrada puntual"),
    @("Carlos Mendoza", "01/09/2026", "06:10:00 p. m.", "Salida regular"),
    @("Carlos Mendoza", "02/09/2026", "07:50:00 a. m.", "Entrada regular"),
    @("Carlos Mendoza", "02/09/2026", "06:05:00 p. m.", "Salida regular"),
    @("Carlos Mendoza", "03/09/2026", "07:48:00 a. m.", "Solo marcó en la mañana"),
    @("Lucia Fernandez", "01/09/2026", "07:40:00 a. m.", "Entrada puntual"),
    @("Lucia Fernandez", "01/09/2026", "06:30:00 p. m.", "Salida regular"),
    @("Lucia Fernandez", "02/09/2026", "06:15:00 p. m.", "Solo marcó en la tarde"),
    @("Lucia Fernandez", "03/09/2026", "07:35:00 a. m.", "Entrada puntual"),
    @("Lucia Fernandez", "03/09/2026", "06:00:00 p. m.", "Salida regular"),
    @("Roberto Gomez", "01/09/2026", "07:55:00 AM", "Marcaje con formato AM/PM mayúsculas"),
    @("Roberto Gomez", "01/09/2026", "05:45:00 PM", "Salida normal"),
    @("Roberto Gomez", "02/09/2026", "07:50:00 AM", "Entrada normal"),
    @("Roberto Gomez", "02/09/2026", "05:50:00 PM", "Salida normal"),
    @("Mariana Rios", "01/09/2026", "08:05:00 a. m.", "Entrada con retardo"),
    @("Mariana Rios", "01/09/2026", "06:00:00 p. m.", "Salida regular"),
    @("Mariana Rios", "02/09/2026", "08:00:00 a. m.", "Entrada normal"),
    @("Mariana Rios", "02/09/2026", "06:15:00 p. m.", "Salida normal"),
    @("Mariana Rios", "03/09/2026", "07:55:00 a. m.", "Entrada normal"),
    @("Mariana Rios", "03/09/2026", "06:10:00 p. m.", "Salida normal")
)

for ($r = 0; $r -lt $data2.Length; $r++) {
    for ($c = 0; $c -lt $data2[$r].Length; $c++) {
        $ws2.Cells.Item($r + 1, $c + 1).Value2 = $data2[$r][$c]
    }
}
$ws2.UsedRange.EntireColumn.AutoFit() | Out-Null
$path2 = Join-Path $baseDir "ejemplo_formato_4_columnas_ampm.xlsx"
$wb2.SaveAs($path2, 51)
$wb2.Close()

# ----------------------------------------------------
# 3. FORMATO 3: 4 COLUMNAS CON HORA 24 HORAS DIRECTA
# ----------------------------------------------------
$wb3 = $excel.Workbooks.Add()
$ws3 = $wb3.Worksheets.Item(1)
$ws3.Name = "Asistencias_4_Cols_24h"

$data3 = @(
    @("Nombre", "Fecha", "Hora (24h)", "Observaciones"),
    @("Carlos Mendoza", "01/09/2026", "07:45:00", "Entrada en formato militar"),
    @("Carlos Mendoza", "01/09/2026", "18:10:00", "Salida en formato militar"),
    @("Carlos Mendoza", "02/09/2026", "07:50:00", "Entrada"),
    @("Carlos Mendoza", "02/09/2026", "18:05:00", "Salida"),
    @("Carlos Mendoza", "03/09/2026", "07:48:00", "Solo entrada"),
    @("Lucia Fernandez", "01/09/2026", "07:40:00", "Entrada"),
    @("Lucia Fernandez", "01/09/2026", "18:30:00", "Salida"),
    @("Lucia Fernandez", "02/09/2026", "18:15:00", "Solo salida tarde"),
    @("Lucia Fernandez", "03/09/2026", "07:35:00", "Entrada"),
    @("Lucia Fernandez", "03/09/2026", "18:00:00", "Salida"),
    @("Roberto Gomez", "01/09/2026", "07:55:00", "Entrada"),
    @("Roberto Gomez", "01/09/2026", "17:45:00", "Salida"),
    @("Roberto Gomez", "02/09/2026", "07:50:00", "Entrada"),
    @("Roberto Gomez", "02/09/2026", "17:50:00", "Salida"),
    @("Mariana Rios", "01/09/2026", "08:05:00", "Entrada"),
    @("Mariana Rios", "01/09/2026", "18:00:00", "Salida"),
    @("Mariana Rios", "02/09/2026", "08:00:00", "Entrada"),
    @("Mariana Rios", "02/09/2026", "18:15:00", "Salida"),
    @("Mariana Rios", "03/09/2026", "07:55:00", "Entrada"),
    @("Mariana Rios", "03/09/2026", "18:10:00", "Salida")
)

for ($r = 0; $r -lt $data3.Length; $r++) {
    for ($c = 0; $c -lt $data3[$r].Length; $c++) {
        $ws3.Cells.Item($r + 1, $c + 1).Value2 = $data3[$r][$c]
    }
}
$ws3.UsedRange.EntireColumn.AutoFit() | Out-Null
$path3 = Join-Path $baseDir "ejemplo_formato_4_columnas_24h.xlsx"
$wb3.SaveAs($path3, 51)
$wb3.Close()

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "3 Archivos creados exitosamente."
