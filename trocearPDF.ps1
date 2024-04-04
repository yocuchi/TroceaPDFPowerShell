# Verificar si el módulo PSWritePDF está instalado
if (-not (Get-Module -ListAvailable -Name PSWritePDF)) {
    # Módulo no instalado, instalarlo automáticamente
    Write-Host "Instalando el módulo PSWritePDF..."
    Install-Module -Name PSWritePDF -Force -Scope CurrentUser -AllowClobber -Repository PSGallery -Verbose
}

# Importar el módulo PSWritePDF
Import-Module PSWritePDF


# Verificar si el módulo iText7Module está instalado
if (-not (Get-Module -ListAvailable -Name iText7Module)) {
    # Módulo no instalado, instalarlo automáticamente
    Write-Host "Instalando el módulo iText7Module..."
    Install-Module -Name iText7Module -Force -Scope CurrentUser -AllowClobber -Repository PSGallery -Verbose
}

# Importar el módulo iText7Module
Import-Module iText7Module

# Pregunta al usuario por el código del paciente
$codigoPaciente = Read-Host "Ingrese el código del paciente"

# Pregunta al usuario por la semana de la visita
$semanaVisita = Read-Host "Ingrese la semana de la visita"

# Obtener la fecha actual
$fechaActual = Get-Date -Format "yyyyMMdd"

# Obtener la ruta del archivo PDF desde el primer argumento
if ($args.Count -eq 0) {
    Write-Host "Arrastra un archivo PDF sobre este script para trocearlo."
    exit
}

$pdfPath = $args[0]

# Verificar si el archivo PDF existe
if (-not (Test-Path -Path $pdfPath -PathType Leaf)) {
    Write-Host "El archivo PDF especificado no existe: $pdfPath"
    exit
}

# Pregunta al usuario por las páginas de inicio y fin para cada tipo de fichero
$paginasIMP_VIAL = Read-Host "Ingrese las páginas de inicio y fin para IMP_VIAL (separadas por coma)"
$paginasLRF = Read-Host "Ingrese las páginas de inicio y fin para LRF (separadas por coma)"
$paginasAWB = Read-Host "Ingrese las páginas de inicio y fin para AWB (separadas por coma)"
$paginasVRF = Read-Host "Ingrese las páginas de inicio y fin para VRF (separadas por coma)"

# Convertir las respuestas del usuario en arreglos de páginas para cada tipo de fichero
$paginasIMP_VIAL = $paginasIMP_VIAL -split ","
$paginasLRF = $paginasLRF -split ","
$paginasAWB = $paginasAWB -split ","
$paginasVRF = $paginasVRF -split ","

# Trocear el PDF
function TrocearPDF($pdfPath, $paginas, $nombreArchivo) {
    $reader = [iText.Kernel.Pdf.PdfReader]::new($pdfPath)
    $pdfDoc = [iText.Kernel.Pdf.PdfDocument]::new($reader)
    $outputPath = Join-Path -Path (Split-Path -Path $pdfPath) -ChildPath $nombreArchivo

    $copy = [iText.Kernel.Pdf.PdfWriter]::new($outputPath)
    $copyPdfDoc = [iText.Kernel.Pdf.PdfDocument]::new($copy)

    foreach ($pagina in $paginas) {
        Write-Host "pagina $pagina"
        $page = $pdfDoc.GetPage([int]$pagina).CopyTo($copyPDfDoc)
        $copyPdfDoc.AddPage($page)
    }

    $copyPdfDoc.Close()
    $pdfDoc.Close()
}
# Trocear el PDF en los tipos de fichero IMP_VIAL, LRF, AWB y VRF
TrocearPDF $pdfPath $paginasIMP_VIAL "$codigoPaciente _IMP_VIAL_$semanaVisita_ $fechaActual.pdf"
TrocearPDF $pdfPath $paginasLRF "$codigoPaciente _LRF_$semanaVisita_ $fechaActual.pdf"
TrocearPDF $pdfPath $paginasAWB "$codigoPaciente _AWB_$semanaVisita_ $fechaActual.pdf"
TrocearPDF $pdfPath $paginasVRF "$codigoPaciente _VRF_$semanaVisita_ $fechaActual.pdf"

Write-Host "PDF troceado exitosamente."


$paginasVRF = Read-Host "Ingrese las páginas de inicio y fin para VRF (separadas por coma)"