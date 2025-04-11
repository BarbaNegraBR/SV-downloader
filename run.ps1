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

# Cria pasta SVteste dentro de Downloads
$downloadsFolder = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
$tempFolder = Join-Path $downloadsFolder "SVteste"
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
    
    # Download com retry e verificações adicionais
    $maxRetries = 3
    $retryCount = 0
    $success = $false
    
    while (-not $success -and $retryCount -lt $maxRetries) {
        try {
            # Limpa arquivo anterior se existir
            if (Test-Path $outputPath) {
                Remove-Item $outputPath -Force
            }

            # Usa Invoke-WebRequest para mais controle
            $response = Invoke-WebRequest -Uri $url -OutFile $outputPath -PassThru
            
            if (Test-Path $outputPath) {
                $fileSize = (Get-Item $outputPath).Length
                if ($fileSize -gt 0) {
                    # Verifica os primeiros bytes do arquivo para confirmar que é RAR
                    $bytes = Get-Content $outputPath -Encoding Byte -TotalCount 4
                    if ($bytes[0] -eq 0x52 -and $bytes[1] -eq 0x61 -and $bytes[2] -eq 0x72) {
                        $success = $true
                        Write-Host "Download concluído! Tamanho: $([math]::Round($fileSize/1MB, 2)) MB" -ForegroundColor Green
                        Write-Host "Assinatura do arquivo RAR verificada com sucesso" -ForegroundColor Green
                        Write-Host "Arquivo salvo em: $outputPath" -ForegroundColor Green
                    } else {
                        Write-Host "Arquivo baixado não tem assinatura RAR válida, tentando novamente..." -ForegroundColor Yellow
                        Remove-Item $outputPath -Force
                    }
                } else {
                    Write-Host "Arquivo vazio, tentando novamente..." -ForegroundColor Yellow
                    Remove-Item $outputPath -Force
                }
            }
        } catch {
            $retryCount++
            Write-Host "Erro no download: $_" -ForegroundColor Red
            if ($retryCount -lt $maxRetries) {
                Write-Host "Tentativa $retryCount de $maxRetries falhou. Tentando novamente..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
        }
    }
    
    if (-not $success) {
        throw "Não foi possível baixar o arquivo RAR válido após $maxRetries tentativas"
    }

    # Verifica se o arquivo é realmente um RAR
    Write-Host "Verificando arquivo..." -ForegroundColor Yellow
    $7zip = "C:\Program Files\7-Zip\7z.exe"
    
    # Lista o conteúdo antes de tentar extrair
    Write-Host "Listando conteúdo do arquivo..." -ForegroundColor Yellow
    $listProcess = Start-Process -FilePath $7zip -ArgumentList "l", $outputPath -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$tempFolder\7z_list.log" -RedirectStandardError "$tempFolder\7z_list.error"
    
    if ($listProcess.ExitCode -ne 0) {
        if (Test-Path "$tempFolder\7z_list.error") {
            Write-Host "Erro ao listar conteúdo:" -ForegroundColor Red
            Get-Content "$tempFolder\7z_list.error"
        }
        if (Test-Path "$tempFolder\7z_list.log") {
            Write-Host "Log da listagem:" -ForegroundColor Yellow
            Get-Content "$tempFolder\7z_list.log"
        }
        throw "O arquivo baixado não é um arquivo RAR válido"
    }
    
    # Extrai o arquivo
    Write-Host "Extraindo arquivo..." -ForegroundColor Yellow
    $extractPath = Join-Path $tempFolder "extracted"
    New-Item -ItemType Directory -Force -Path $extractPath | Out-Null
    
    # Tenta extrair usando comando direto do 7z com tratamento de erros melhorado
    $extractProcess = Start-Process -FilePath $7zip -ArgumentList "x", "-y", "-o$extractPath", $outputPath -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$tempFolder\7z.log" -RedirectStandardError "$tempFolder\7z.error"
    
    # Verifica o resultado da extração
    if ($extractProcess.ExitCode -eq 0) {
        Write-Host "Arquivo extraído com sucesso!" -ForegroundColor Green
        Write-Host "Arquivos extraídos em: $extractPath" -ForegroundColor Green
        
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
        # Mostra logs de erro se disponíveis
        if (Test-Path "$tempFolder\7z.error") {
            Write-Host "Log de erro do 7-Zip:" -ForegroundColor Red
            Get-Content "$tempFolder\7z.error"
        }
        if (Test-Path "$tempFolder\7z.log") {
            Write-Host "Log do 7-Zip:" -ForegroundColor Yellow
            Get-Content "$tempFolder\7z.log"
        }
        throw "Erro ao extrair o arquivo RAR. Código de saída: $($extractProcess.ExitCode)"
    }
    
} catch {
    Write-Host "Erro: $_" -ForegroundColor Red
    Write-Host "Pressione qualquer tecla para continuar..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
} 
