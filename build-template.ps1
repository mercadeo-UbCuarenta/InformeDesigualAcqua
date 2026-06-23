$ErrorActionPreference = "Stop"
$root = Get-Location
$xlsx = Join-Path $root "plantilla-informe-evento.xlsx"
$tmp = Join-Path $root (".xlsx-build-" + [guid]::NewGuid().ToString())

New-Item -ItemType Directory -Path $tmp | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "_rels") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "xl\_rels") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "xl\worksheets") | Out-Null

function X([object]$value) {
  [System.Security.SecurityElement]::Escape([string]$value)
}

function ColName([int]$number) {
  $name = ""
  while ($number -gt 0) {
    $mod = ($number - 1) % 26
    $name = [char](65 + $mod) + $name
    $number = [math]::Floor(($number - 1) / 26)
  }
  $name
}

function Write-Sheet($path, [array]$rows) {
  $sb = [System.Text.StringBuilder]::new()
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
  for ($r = 0; $r -lt $rows.Count; $r++) {
    $rowNumber = $r + 1
    [void]$sb.Append("<row r=""$rowNumber"">")
    $row = $rows[$r]
    for ($c = 0; $c -lt $row.Count; $c++) {
      $cellRef = (ColName ($c + 1)) + $rowNumber
      $value = $row[$c]
      if ($null -eq $value -or [string]$value -eq "") { continue }
      if ($value -is [int] -or $value -is [double] -or $value -is [decimal]) {
        [void]$sb.Append("<c r=""$cellRef""><v>$value</v></c>")
      } else {
        [void]$sb.Append("<c r=""$cellRef"" t=""inlineStr""><is><t>$(X $value)</t></is></c>")
      }
    }
    [void]$sb.Append("</row>")
  }
  [void]$sb.Append("</sheetData></worksheet>")
  [IO.File]::WriteAllText($path, $sb.ToString(), [Text.UTF8Encoding]::new($false))
}

$sheets = @(
  @{
    Name = "Evento"
    Rows = @(
      @("Campo", "Valor"),
      @("Nombre evento", "Levi's Colina"),
      @("Tipo evento", "Reinauguracion de tienda"),
      @("Tienda", "Tienda Colina"),
      @("Ciudad", "Bogota"),
      @("Fecha evento", 46185),
      @("Periodo medicion", "Previo: 01 al 11 de junio | Evento: 12 de junio | Posterior: 13 al 15 de junio de 2026"),
      @("Venta resultado", 0),
      @("Venta cumplimiento", 0),
      @("Venta dia evento", 0),
      @("ATV dia evento", 0),
      @("Gasto total", 16700000),
      @("Resumen ejecutivo", "El evento genero una reapertura con alto impacto visual, buena asistencia de clientes invitados y una venta superior a la meta definida."),
      @("Analisis dia evento", "La jornada concentro el mayor impacto comercial: alta asistencia, buen cierre de clientes invitados y mejores indicadores de ATV y UPT frente a la meta del dia."),
      @("Hallazgo principal", "Los clientes invitados respondieron mejor a la experiencia personalizada que a la comunicacion masiva."),
      @("Recomendacion", "Replicar el formato de atencion VIP, reforzando agendamiento previo y seguimiento post-evento por CRM."),
      @("Conclusion", "El evento fue rentable para el objetivo comercial y fortalecio la percepcion de tienda renovada."),
      @("Proximos pasos", "Realizar seguimiento a clientes asistentes durante los proximos 15 dias y medir recompra.")
    )
  },
  @{
    Name = "Ventas"
    Rows = @(
      @("Fecha", "Factura", "Tienda", "Cliente", "Nombre", "Apellido", "Email", "Total", "Descuento", "Costo envio", "Impuestos"),
      @(46184, "FAC-001", "Tienda ejemplo", "C-001", "", "", "", 650000, 0, 0, 0),
      @(46185, "FAC-002", "Tienda ejemplo", "C-002", "", "", "", 980000, 0, 0, 0),
      @(46186, "FAC-003", "Tienda ejemplo", "C-003", "", "", "", 720000, 0, 0, 0)
    )
  },
  @{
    Name = "Trafico"
    Rows = @(
      @("Fecha", "Punto de venta", "Trafico exterior", "Trafico individual", "Tasa de conversion (Individual)"),
      @(46184, "Tienda ejemplo", 2500, 75, 0.03),
      @(46185, "Tienda ejemplo", 3200, 420, 0.1313),
      @(46186, "Tienda ejemplo", 2800, 110, 0.0393)
    )
  },
  @{
    Name = "Pauta Digital"
    Rows = @(
      @("PLATAFORMA", "IMPRESIONES", "CLICS", "CTR", "VIDEO STARTS", "25% COMPLETE", "50% COMPLETE", "75% COMPLETE", "100% COMPLETE", "COMPLETE RATE", "CONSUMO"),
      @("Youtube", 1000000, 1500, 0.0015, 800000, 680000, 590000, 540000, 500000, 0.625, 4000000),
      @(),
      @("PLATAFORMA", "IMPRESIONES", "CLICS", "CTR", "CONSUMO"),
      @("META", 650000, 1400, 0.0022, 2000000)
    )
  },
  @{
    Name = "Momentos"
    Rows = @(
      @("Momento", "Titulo", "Analisis", "Indicador"),
      @("Previo", "Comunicacion de apertura", "CRM, redes y contacto desde tienda prepararon la convocatoria antes de la reapertura.", "Base impactada: 520"),
      @("Dia del evento", "Conversion y experiencia", "La experiencia en tienda genero trafico calificado, contenido organico y cierre comercial.", 'Venta dia: $52,8M'),
      @("Posterior", "Sostenimiento de venta", "Seguimiento a asistentes, remarketing y comunicacion de tienda sostienen la venta posterior.", "Acciones activas: 3")
    )
  },
  @{
    Name = "Gastos"
    Rows = @(
      @("Categoria", "Proveedor", "Concepto", "Modalidad", "Gasto", "Observaciones"),
      @("Experiencia", "Proveedor A", "Catering y punto de hidratacion", "Efectivo", 2700000, "Pago directo al proveedor"),
      @("Experiencia", "Proveedor B", "Produccion del evento", "Efectivo", 2100000, "Pago directo al proveedor"),
      @("Comunicacion", "Influencer", "Contenido de apertura", "Canje de prendas", 1500000, "Entregado en producto")
    )
  },
  @{
    Name = "Acciones"
    Rows = @(
      @("Accion", "Tipo", "Responsable", "Fecha", "Estado", "Descripcion", "Indicador"),
      @("Invitacion clientes VIP", "CRM previo", "Marketing", "01 al 11 de junio", "Ejecutada", "Segmentacion de clientes activos y comunicacion personalizada para agendar asistencia.", "Base impactada: 520"),
      @("Experiencia de reapertura", "BTL dia evento", "Retail + Marketing", "12 de junio", "Ejecutada", "DJ, catering y foto cabina para generar permanencia, contenido organico y conversacion en tienda.", "86 asistentes"),
      @("Seguimiento asistentes", "CRM posterior", "Marketing + tienda", "13 al 15 de junio", "Activa", "Contacto posterior para agradecer asistencia, sostener interes y promover recompra.", "34 clientes VIP")
    )
  },
  @{
    Name = "Dinamicas"
    Rows = @(
      @("Dinamica", "Mecanica", "Resultado"),
      @("Regalo por compra", "Obsequio desde compra minima", "Impulso al ticket promedio"),
      @("Bono recompra", "Bono para una compra posterior", "Sostenimiento de venta")
    )
  },
  @{
    Name = "Resultados"
    Rows = @(
      @("Titulo", "Impacto", "Lectura", "Prioridad"),
      @("Venta incremental", "+18%", "La tienda supero el promedio de venta frente a fines de semana comparables.", "Alta"),
      @("Clientes VIP", "34", "Base priorizada con mejor conversion y alto potencial de recompra.", "Alta"),
      @("Conversion en tienda", "28%", "La dinamica comercial ayudo a cerrar ventas durante la jornada.", "Media")
    )
  },
  @{
    Name = "Fotos"
    Rows = @(
      @("Titulo", "Categoria", "Descripcion", "Imagen URL"),
      @("Reapertura con alto impacto visual", "Experiencia", "Ambientacion de tienda y experiencia inicial para clientes invitados.", "assets/evento-principal.jpg"),
      @("Interaccion con asistentes", "Cliente", "Atencion personalizada para impulsar conversion y recompra.", "assets/evento-experiencia.jpg"),
      @("Activacion comercial", "Dinamica", "Beneficio tactico para fortalecer ATV, UPT y cierre de venta.", "assets/evento-dinamica.jpg")
    )
  }
)

for ($i = 0; $i -lt $sheets.Count; $i++) {
  Write-Sheet (Join-Path $tmp "xl\worksheets\sheet$($i + 1).xml") $sheets[$i].Rows
}

$contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
for ($i = 0; $i -lt $sheets.Count; $i++) {
  $contentTypes += "<Override PartName=""/xl/worksheets/sheet$($i + 1).xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>"
}
$contentTypes += "</Types>"
[IO.File]::WriteAllText((Join-Path $tmp "[Content_Types].xml"), $contentTypes, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $tmp "_rels\.rels"), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', [Text.UTF8Encoding]::new($false))

$workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
$relationships = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
for ($i = 0; $i -lt $sheets.Count; $i++) {
  $id = $i + 1
  $workbook += "<sheet name=""$(X $sheets[$i].Name)"" sheetId=""$id"" r:id=""rId$id""/>"
  $relationships += "<Relationship Id=""rId$id"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$id.xml""/>"
}
$workbook += "</sheets></workbook>"
$relationships += "</Relationships>"
[IO.File]::WriteAllText((Join-Path $tmp "xl\workbook.xml"), $workbook, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $tmp "xl\_rels\workbook.xml.rels"), $relationships, [Text.UTF8Encoding]::new($false))

if ([IO.File]::Exists($xlsx)) {
  [IO.File]::Delete($xlsx)
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $xlsx)
Get-Item $xlsx | Select-Object Name, Length, LastWriteTime
