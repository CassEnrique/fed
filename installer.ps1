# ============================================================
#  BLOQUE DE ELEVACIÓN AUTOMÁTICA (para doble clic)
# ============================================================

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

    Write-Host "Solicitando permisos de administrador..." -ForegroundColor Yellow

    # Reinicia el script como Administrador
    Start-Process PowerShell -Verb RunAs -ArgumentList @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`""
    )
    exit
}

# ============================================================
#  AQUÍ EMPIEZA TU SCRIPT CON PERMISOS DE ADMINISTRADOR
# ============================================================

Write-Host "✅ Ejecutando con privilegios de administrador" -ForegroundColor Green

# ==========================================
# PASO 0: CONFIGURACIÓN INICIAL
# ==========================================
$ErrorActionPreference = "Continue"

# URLs
$otherUrls = @(
    "https://github.com/oschwartz10612/poppler-windows/releases/download/v26.09.0-0/Release-26.09.0-0.zip",
    "https://pub-cb2496f16a314ed6a74ac9038851ddf9.r2.dev/FED.zip"
)

$architecture = [System.Environment]::Is64BitOperatingSystem
$baseUrl = if ($architecture) {
    "https://aka.ms/highdpimfc2013x64enu"
} else {
    "https://aka.ms/highdpimfc2013x86enu"
}

# Rutas destino
$popplerDest = "C:\poppler"
$popplerBin = "$popplerDest\Library\bin"
$fedDest = "C:\fed"

# Crear carpeta temporal
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "Setup_$(Get-Random)"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
Write-Output "Archivos temporales en: $tempDir"

$failedOperations = @()

# ==========================================
# RESOLVER NOMBRE DEL INSTALADOR MFC (.exe)
# ==========================================
try {
    $webRequest = [System.Net.WebRequest]::Create($baseUrl)
    $webRequest.Method = "HEAD"
    $webRequest.AllowAutoRedirect = $true
    $webResponse = $webRequest.GetResponse()
    $finalUrl = $webResponse.ResponseUri.AbsoluteUri
    $webResponse.Close()

    $mfcFileName = [System.IO.Path]::GetFileName($finalUrl)
    if ([string]::IsNullOrWhiteSpace($mfcFileName) -or $mfcFileName -notlike "*.exe") {
        $mfcFileName = if ($architecture) { "vcredist_x64.exe" } else { "vcredist_x86.exe" }
    }
}
catch {
    Write-Output "⚠️ No se resolvió el enlace MFC: $_"
    $finalUrl = $baseUrl
    $mfcFileName = if ($architecture) { "vcredist_x64.exe" } else { "vcredist_x86.exe" }
}

# ==========================================
# DESCARGA DE ARCHIVOS
# ==========================================
# 1. Descargar runtime MFC (Visual C++ Redistributable)
$mfcPath = Join-Path $tempDir $mfcFileName
try {
    Write-Output "⬇️ Descargando vcredist: $mfcFileName ..."
    Invoke-WebRequest -Uri $finalUrl -OutFile $mfcPath -UseBasicParsing
    Write-Output "✅ Descargado: $mfcPath"
} catch {
    Write-Output "❌ Error descargando MFC: $_"
    $failedOperations += "MFC Download"
}

# 2. Descargar archivos ZIP
foreach ($url in $otherUrls) {
    try {
        $zipName = [System.IO.Path]::GetFileName($url)
        if ([string]::IsNullOrWhiteSpace($zipName)) { $zipName = "download_$(Get-Random).zip" }
        $zipPath = Join-Path $tempDir $zipName

        Write-Output "⬇️ Descargando: $zipName ..."
        Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing
        Write-Output "✅ Descargado: $zipPath"
    } catch {
        Write-Output "❌ Error descargando '$url': $_"
        $failedOperations += $url
    }
}

# ==========================================
# INSTALAR MFC (vcredist silencioso)
# ==========================================
if (Test-Path $mfcPath) {
    Write-Output "Instalando paquete de compatibilidad (vc++ redistributable)..."
    try {
        $proc = Start-Process -FilePath $mfcPath -ArgumentList "/install", "/quiet", "/norestart" -Wait -PassThru -ErrorAction Stop
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            Write-Output "✅ Instalación completada (código: $($proc.ExitCode))."
        } else {
            Write-Output "❌ Instalación falló: Código de salida $($proc.ExitCode)"
            $failedOperations += "MFC Installation"
        }
    } catch {
        Write-Output "❌ No se pudo ejecutar el instalador: $_"
        $failedOperations += "MFC Execution"
    }
} else {
    Write-Output "❌ No se encontró el instalador MFC para ejecutar."
    $failedOperations += "MFC Missing"
}

# ==========================================
# DESCOMPRIMIR Y COPIAR POPPLER A C:\poppler
# ==========================================
$popplerZip = Get-ChildItem -Path $tempDir -Filter "Release-26.09.0-0.zip" | Select-Object -First 1
if ($popplerZip) {
    $extractPoppler = Join-Path $tempDir "poppler_extract"
    New-Item -ItemType Directory -Path $extractPoppler -Force | Out-Null

    try {
        Write-Output "Descomprimiendo: $($popplerZip.Name) ..."
        Expand-Archive -Path $popplerZip.FullName -DestinationPath $extractPoppler -Force

        # Buscar carpeta extraída: poppler-26.09.0
        $popplerFolder = Get-ChildItem -Path $extractPoppler -Directory | Where-Object { $_.Name -like "poppler-*" } | Select-Object -First 1
        if ($popplerFolder) {
            Write-Output "Carpeta encontrada: $($popplerFolder.FullName)"

            # Crear C:\poppler y copiar contenido
            New-Item -ItemType Directory -Path $popplerDest -Force | Out-Null
            Write-Output "Copiando contenido a $popplerDest ..."
            Copy-Item -Path "$($popplerFolder.FullName)\*" -Destination $popplerDest -Recurse -Force
            Write-Output "✅ Poppler instalado en: $popplerDest"

            # Verificar que existan: Library y share
            if (Test-Path "$popplerDest\Library") { Write-Output "✅ Directorio Library presente." }
            if (Test-Path "$popplerDest\share") { Write-Output "✅ Directorio share presente." }
        } else {
            Write-Output "❌ No se encontró la carpeta poppler-* dentro del zip."
            $failedOperations += "Poppler folder not found"
        }
    } catch {
        Write-Output "❌ Error descomprimiendo poppler: $_"
        $failedOperations += "Poppler Extraction"
    }
}

# ==========================================
# AGREGAR C:\poppler\Library\bin AL PATH DEL SISTEMA
# ==========================================
$pathToAdd = "$popplerBin"
try {
    # Obtener el PATH del sistema
    $regPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    $scope = "Machine"
    $currentPath = [Environment]::GetEnvironmentVariable("Path", $scope)

    if ($currentPath -notlike "*$pathToAdd*") {
        Write-Output "Agregando al PATH del sistema: $pathToAdd"
        $newPath = "$currentPath;$pathToAdd"
        # Asegurarse de que no haya doble punto y coma
        $newPath = (@($newPath -split ';').Where({$_}) | Select-Object -Unique) -join ';'

        # Set-ItemProperty -Path "HKLM:\$regPath" -Name "Path" -Value $newPath
        # Actualizar la variable PATH en el contexto actual también
        [Environment]::SetEnvironmentVariable("Path", $newPath, $scope)
        Write-Output "✅ PATH actualizado."
    } else {
        Write-Output "✅ La ruta $pathToAdd ya está en el PATH del sistema."
    }
} catch {
    Write-Output "❌ Error al actualizar el PATH del sistema: $_"
    $failedOperations += "PATH Update"
}

# ==========================================
# COPIAR FED.zip a C:\fed
# ==========================================
$fedZip = Get-ChildItem -Path $tempDir -Filter "FED.zip" | Select-Object -First 1
if ($fedZip) {
    try {
        New-Item -ItemType Directory -Path $fedDest -Force | Out-Null
        Write-Output "Descomprimiendo FED.zip en $fedDest ..."
        Expand-Archive -Path $fedZip.FullName -DestinationPath $fedDest -Force
        Write-Output "✅ Archivo FED instalado en: $fedDest"
    } catch {
        Write-Output "❌ Error copiando FED: $_"
        $failedOperations += "FED Copy"
    }
}

# ==========================================
# RESUMEN FINAL
# ==========================================
Write-Output "`n✅✅✅ PROCESO COMPLETADO"
Write-Output "Instalación en: C:\poppler"
Write-Output "PATH incluye: $popplerBin"
Write-Output "Archivos FED en: $fedDest"

if ($failedOperations.Count -gt 0) {
    Write-Output "`n❌ ALGUNAS OPERACIONES FALLARON:"
    $failedOperations | ForEach-Object { Write-Output "   - $_" }
} else {
    Write-Output "`nTodo completado sin errores."
}

# Abrir carpeta destino para revisión (opcional)
Explorer $fedDest

Write-Host "`nPresiona cualquier tecla para salir..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
