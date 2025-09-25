$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

# Configuración
$ProjectPath = "C:\Users\junior.marketing\Documents\rpa_downloads_extranet_gls"
$LogPath = "$ProjectPath\logs"
$VenvPath = "$ProjectPath\.venv"

# Crear directorio de logs si no existe
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force
}

# Función de logging
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    Write-Host $logMessage
    try {
        Add-Content -Path "$LogPath\scheduler.log" -Value $logMessage -Force
    } catch {
        Write-Host "ERROR writing to scheduler.log: $_"
    }
}

try {
    Write-Log "Iniciando RPA GLS"
    
    # Cambiar al directorio del proyecto
    Set-Location $ProjectPath
    
    # Verificar que el entorno virtual existe
    if (-not (Test-Path "$VenvPath\Scripts\python.exe")) {
        throw "No se ha encontrado el entorno virtual en $VenvPath"
    }
    
    # Activar entorno virtual y ejecutar
    $pythonExe = "$VenvPath\Scripts\python.exe"
    $arguments = @("main.py")
    
    Write-Log "Ejecutando: $pythonExe $arguments"
    
    # Ejecutar con captura de salida
    $processOutput = & $pythonExe $arguments 2>&1
    $exitCode = $LASTEXITCODE

    # Escribir output a archivos
    $processOutput | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } | 
        Out-File "$LogPath\scheduler_error.log" -Append -Encoding UTF8
    
    $processOutput | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] } | 
        Out-File "$LogPath\scheduler_output.log" -Append -Encoding UTF8
    
    if ($exitCode -eq 0) {
        Write-Log "Proceso finalizado con éxito (Codigo: $exitCode)"
    } else {
        Write-Log "Proceso fallido con código de salida: $($process.ExitCode)"
        exit $exitCode
    }
    
} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Log "SEGUIMIENTO: $($_.ScriptStackTrace)"
    exit 1
} finally {
    Write-Log "Proceso finalizado"
}