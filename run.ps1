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
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($7zipUrl, $installerPath)
        Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait
        Remove-Item $installerPath -Force
        return $true
    } catch {
        Write-Host "Erro ao instalar 7-Zip: $_" -ForegroundColor Red
        return $false
    }
}

# Função para extrair arquivo RAR
function Extract-RAR {
    param (
        [string]$rarPath,
        [string]$destination
    )
    
    $7zip = "C:\Program Files\7-Zip\7z.exe"
    if (Test-Path $7zip) {
        Write-Host "Extraindo com 7-Zip..." -ForegroundColor Yellow
        $process = Start-Process -FilePath $7zip -ArgumentList "x", "`"$rarPath`"", "-y", "-o`"$destination`"" -NoNewWindow -PassThru -Wait
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
        Write-Host "7-Zip não encontrado." -ForegroundColor Yellow
        if (-not (Install-7Zip)) {
            throw "Não foi possível instalar o 7-Zip. Por favor, instale manualmente de www.7-zip.org"
        }
        Write-Host "7-Zip instalado com sucesso!" -ForegroundColor Green
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
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($url, $outputPath)
            
            if (Test-Path $outputPath) {
                $fileSize = (Get-Item $outputPath).Length
                if ($fileSize -gt 0) {
                    $success = $true
                    Write-Host "Download concluído! Tamanho: $([math]::Round($fileSize/1MB, 2)) MB" -ForegroundColor Green
                } else {
                    Write-Host "Arquivo vazio, tentando novamente..." -ForegroundColor Yellow
                    Remove-Item $outputPath -Force
                }
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
    
    if (Extract-RAR -rarPath $outputPath -destination $extractPath) {
        Write-Host "Arquivo extraído com sucesso!" -ForegroundColor Green
        
        # Procura por executáveis
        $exeFiles = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse
        if ($exeFiles.Count -gt 0) {
            Write-Host "Executável encontrado: $($exeFiles[0].Name)" -ForegroundColor Green
            Write-Host "Executando programa..." -ForegroundColor Green
            Start-Process $exeFiles[0].FullName
        } else {
            Write-Host "Conteúdo da pasta extraída:" -ForegroundColor Yellow
            Get-ChildItem -Path $extractPath -Recurse | ForEach-Object {
                Write-Host " - $($_.FullName)"
            }
            throw "Nenhum executável encontrado na pasta extraída."
        }
    } else {
        Write-Host "Conteúdo da pasta temporária:" -ForegroundColor Yellow
        Get-ChildItem -Path $tempFolder -Recurse | ForEach-Object {
            Write-Host " - $($_.FullName)"
        }
        throw "Erro ao extrair o arquivo RAR. Verifique se o arquivo está correto."
    }
    
} catch {
    Write-Host "Erro: $_" -ForegroundColor Red
    Write-Host "Pressione qualquer tecla para continuar..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
} finally {
    # Aguarda antes de limpar
    Start-Sleep -Seconds 5
    # Limpa arquivos temporários
    Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
} 
