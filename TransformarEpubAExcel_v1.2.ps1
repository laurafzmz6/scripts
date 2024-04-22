# Verificar si el script se esta ejecutando como administrador
 if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Solicitar elevacion de permisos ejecutando el script nuevamente con privilegios elevados
    Start-Process -FilePath "powershell.exe" -ArgumentList "-File `"$PSCommandPath`"" -Verb RunAs
   exit
}

# Verifica si el modulo ImportExcel esta instalado
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    # Instala el modulo solo si no esta instalado
    Install-Module -Name ImportExcel -Force -SkipPublisherCheck
}

# Importa el modulo ImportExcel
Import-Module -Name ImportExcel -Force


# Ruta al archivo EPUB y carpeta de salida para imagenes
$epubPath = Get-ChildItem -Path $PSScriptRoot -Filter *.epub | Select-Object -First 1 -ExpandProperty FullName

# Verifica si se encontro un archivo EPUB
if ($epubPath -eq $null) {
    Write-Host "No se encontro ningun archivo EPUB en la carpeta actual."
    exit
}
Write-Host "Procesando el archivo EPUB: $($epubPath)"

# Nombre del archivo EPUB
$epubName = (Get-Item $epubPath).BaseName

# Crea una carpeta para los resultados si no existe
$resultFolder = Join-Path -Path $PSScriptRoot -ChildPath $epubName
if (-not (Test-Path $resultFolder)) {
    $null = New-Item -Path $resultFolder -ItemType Directory
}

# Crea una copia del archivo EPUB en formato ZIP dentro de la carpeta de resultados
$zipPath = Join-Path -Path $resultFolder -ChildPath ($epubName  + ".zip")
Copy-Item -Path $epubPath -Destination $zipPath -Force

# Crea una carpeta temporal para extraer el contenido del EPUB
$tempFolder = Join-Path -Path $resultFolder -ChildPath ($epubName  + "_temp")
Expand-Archive -Path $zipPath -DestinationPath $tempFolder -Force


# Busca archivos OPF en el directorio temporal
$opfFiles = Get-ChildItem -Path $tempFolder -Recurse -Filter *.opf -File | Select-Object -First 1

# Buscar archivos XHTML en el directorio temporal
$xhtmlFiles = Get-ChildItem -Path $tempFolder -Filter *.xhtml -File -Recurse

# Buscar archivos HTML en el directorio temporal
$htmlFiles = Get-ChildItem -Path $tempFolder -Filter *.html -File -Recurse

# Buscar archivos HTM en el directorio temporal
$htmFiles = Get-ChildItem -Path $tempFolder -Filter *.htm -File -Recurse

# Combinar los resultados en una sola lista
$xhtmlFiles += $htmlFiles
$xhtmlFiles += $htmFiles

# Verificar si se encontraron archivos XHTML/HTML
if ($xhtmlFiles.Count -eq 0) {
    Write-Host "No se encontraron archivos XHTML/HTML en el directorio temporal."
    exit
}

# Obtener el directorio que contiene los archivos XHTML/HTML
$xhtmlDirectoryName = $xhtmlFiles[0].Directory.Name

# Construir la ruta de acceso a la carpeta que contiene los archivos XHTML/HTML
$xhtmlFolderPath = Join-Path -Path $tempFolder -ChildPath $xhtmlDirectoryName

# Buscar archivos de hojas de estilo en el directorio temporal
$styleFiles = Get-ChildItem -Path $tempFolder -File -Recurse | Where-Object { $_.Extension -match '\.css$|\.scss$' }

# Verificar si se encontraron archivos de hojas de estilo
if ($styleFiles.Count -eq 0) {
    Write-Host "No se encontraron archivos de hojas de estilo en el directorio temporal."
}

# Obtener el directorio que contiene los archivos de hojas de estilo
$styleDirectoryName = $styleFiles[0].Directory.Name

# Construir la ruta de acceso a la carpeta que contiene las imagenes
$styleFolderPath = Join-Path -Path $tempFolder -ChildPath $styleDirectoryName

# Ruta de la carpeta de salida para imÃ¡genes
$outputStylePath = Join-Path $resultFolder "style"

# Verifica si la carpeta de salida existe
if (-not (Test-Path $outputStylePath)) {
    # Intenta crear la carpeta de salida
    try {
        $null = New-Item -ItemType Directory -Path $outputStylePath -ErrorAction Stop
        Write-Host "Carpeta de salida creada en: $outputStylePath"
    }
    catch {
        Write-Host "Error al crear la carpeta de salida: $_"
        exit
    }
}
else {
    Write-Host "La carpeta de salida ya existe en: $outputStylePath"
}

# Copiar las imagenes al directorio de salida
foreach ($styleFile in $styleFiles) {
    $destinationPath = $outputStylePath
    Copy-Item -Path $styleFile.FullName -Destination $destinationPath -Force
}


# Buscar archivos de imagenes en el directorio temporal
$imgFiles = Get-ChildItem -Path $tempFolder -File -Recurse | Where-Object { $_.Extension -match '\.jpg$|\.jpeg$|\.png$|\.gif$' }

# Verificar si se encontraron archivos de imagenes
if ($imgFiles.Count -eq 0) {
    Write-Host "No se encontraron archivos de imagenes en el directorio temporal."
}

# Obtener el directorio que contiene los archivos de imagenes
$imgDirectoryName = $imgFiles[0].Directory.Name

# Construir la ruta de acceso a la carpeta que contiene las imagenes
$imgFolderPath = Join-Path -Path $tempFolder -ChildPath $imgDirectoryName

# Ruta de la carpeta de salida para imagenes
$outputImagePath = Join-Path $resultFolder $imgDirectoryName

# Verifica si la carpeta de salida existe
if (-not (Test-Path $outputImagePath)) {
    # Intenta crear la carpeta de salida
    try {
        $null = New-Item -ItemType Directory -Path $outputImagePath -ErrorAction Stop
        Write-Host "Carpeta de salida creada en: $outputImagePath"
    }
    catch {
        Write-Host "Error al crear la carpeta de salida: $_"
        exit
    }
}
else {
    Write-Host "La carpeta de salida ya existe en: $outputImagePath"
}

# Copiar las imagenes al directorio de salida
foreach ($imgFile in $imgFiles) {
    $destinationPath = $outputImagePath
    Copy-Item -Path $imgFile.FullName -Destination $destinationPath -Force
}

# Crear un nuevo archivo Excel con el mismo nombre que el EPUB
$excelPath = Join-Path $resultFolder ($epubName  + "_output.xlsx")
$package = New-Object OfficeOpenXml.ExcelPackage
$worksheet = $package.Workbook.Worksheets.Add("$epubName")

# Añadir encabezados para identificar la etiqueta, la clase, la ruta de las imagenes y el contenido
$worksheet.Cells[1, 1].Value = "tag"
$worksheet.Cells[1, 2].Value = "class"
$worksheet.Cells[1, 3].Value = "src"
$worksheet.Cells[1, 4].Value = "original"

# Obtener el orden de los archivos XHTML segun el archivo OPF
$xhtmlOrder = @()
$opfContent = Get-Content $opfFiles.FullName -Raw
$opfXml = [xml]$opfContent



foreach ($itemref in $opfXml.package.spine.itemref) {
    $xhtmlId = $itemref.idref
    $xhtmlOrder += $xhtmlId
    
}

# Recorrer cada archivo XHTML o HTML en el orden especificado por el archivo OPF
$rowIndex = 2  # Comenzar desde la segunda fila (despues de los encabezados)

foreach ($xhtmlId in $xhtmlOrder) {
    # Buscar el archivo XHTML que coincide con el ID actual en la lista de archivos
    $xhtmlFile = $xhtmlFiles | Where-Object { $_.Name -match "^$xhtmlId(\..*)?$" } | Select-Object -First 1
   

# Si no se encuentra el archivo XHTML directamente, buscar referencias en el OPF
if ($xhtmlFile -eq $null) {
    foreach ($item in $opfXml.package.manifest.item) {
        if ($item.id -eq $xhtmlId) {
            # Obtener solo el nombre del archivo sin la ruta
            $xhtmlFileName = Split-Path -Leaf $item.href
            $xhtmlFile = $xhtmlFiles | Where-Object { $_.Name -eq $xhtmlFileName } | Select-Object -First 1
            break
        }
    }
}

    if ($xhtmlFile -ne $null) {
        # Leer el contenido del archivo XHTML/HTML
        $xhtmlContent = Get-Content $xhtmlFile.FullName -Raw -Encoding UTF8
        Write-Host "Procesando archivo XHTML/HTML: $($xhtmlFile.FullName)..."


        # Reemplazar la referencia a la entidad "&nbsp;" con un espacio en blanco
        $xhtmlContent = $xhtmlContent -replace '&nbsp;', ' '

        # Reemplazar la referencia a la entidad "&mdash;" con un guión largo
        $xhtmlContent = $xhtmlContent -replace '&mdash;', '—'

        # Reemplazar la referencia a la entidad "&crarr;" con un retorno de carro
        $xhtmlContent = $xhtmlContent -replace '&crarr;', [char]13


        # Intentar convertir el contenido a un objeto XML
        try {
            # Extraer el texto del XHTML/HTML y las rutas de las imagenes
            $xhtmlDoc = [xml]$xhtmlContent
        }
        catch {
            Write-Host "Error al procesar el archivo XHTML/HTML: $($xhtmlFile.FullName)"
            Write-Host "Detalles del error: $_"
            continue  # Pasar al siguiente archivo
        }

        $xhtmlNodes = $xhtmlDoc.SelectNodes("//*")  # Seleccionar todos los nodos

        foreach ($node in $xhtmlNodes) {
            $type = $node.LocalName  # Tipo de elemento (etiqueta)
            $content = $node.InnerText  # Contenido del elemento
            $class = $node.GetAttribute("class")  # Clase del elemento


            if ($type -eq "img") {
                # Si el elemento es una imagen, extraer la ruta de la imagen (atributo src)
                $src = $node.GetAttribute("src")
                $worksheet.Cells[$rowIndex, 1].Value = $type
                $worksheet.Cells[$rowIndex, 2].Value = $class
                $worksheet.Cells[$rowIndex, 3].Value = $src
                $rowIndex++
            }
            elseif ($type -eq "p" -or $type -match "^h[1-6]$") {
                # Si el tipo de elemento es un parrafo o un encabezado (h1, h2, ..., h6)
                # Agregar el tipo de contenido y el contenido a las celdas correspondientes en Excel
                $worksheet.Cells[$rowIndex, 1].Value = $type
                $worksheet.Cells[$rowIndex, 2].Value = $class
                $worksheet.Cells[$rowIndex, 4].Value = $content
                $rowIndex++
            }
        }
    }
}

# Guardar el archivo Excel en la carpeta de resultados
$package.SaveAs($excelPath)
$package.Dispose()  # Libera los recursos utilizados por el paquete

Write-Host "Proceso completado para el archivo EPUB: $epubName"
Write-Host "Archivo Excel generado en: $excelPath"

# Elimina la carpeta temporal despues de procesar el EPUB
Remove-Item -Path $tempFolder -Recurse -Force

# Elimina el archivo ZIP
Remove-Item -Path $zipPath -Force

# Abrir automaticamente el archivo Excel
Start-Process $excelPath

# Leer una entrada del usuario para mantener PowerShell abierto
# Read-Host "Presiona Enter para salir..."
