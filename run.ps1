# Função para mostrar banner bonito
function Show-Banner {
    param (
        [string]$Title
    )
    
    $width = 70
    $border = "=" * $width
    $padding = " " * (($width - $Title.Length) / 2)
    
    Write-Host "`n$border" -ForegroundColor Cyan
    Write-Host "$padding$Title" -ForegroundColor White
    Write-Host "$border`n" -ForegroundColor Cyan
}

# Função para verificar se o 7-Zip está instalado
function Test-7Zip {
    $7zipPath = "C:\Program Files\7-Zip\7z.exe"
    
    Show-Banner "VERIFICANDO 7-ZIP"
    
    if (Test-Path $7zipPath) {
        Write-Host "✓ 7-Zip encontrado em: $7zipPath" -ForegroundColor Green
        return @{
            Name = "7-Zip"
            Path = $7zipPath
        }
    }
    
    Write-Host "❌ 7-Zip não encontrado." -ForegroundColor Red
    return $null
}

# Função para instalar o 7-Zip automaticamente
function Install-7Zip {
    Show-Banner "INSTALANDO 7-ZIP"
    Write-Host "Iniciando instalação automática..." -ForegroundColor Yellow
    
    try {
        # URL do instalador do 7-Zip
        $7zipUrl = "https://www.7-zip.org/a/7z2301-x64.exe"
        $installerPath = "$env:TEMP\7zip_installer.exe"
        
        # Baixa o instalador
        Write-Host "`nBaixando instalador do 7-Zip..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $7zipUrl -OutFile $installerPath
        
        # Instala silenciosamente
        Write-Host "Instalando 7-Zip..." -ForegroundColor Yellow
        Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait
        
        # Remove o instalador
        Remove-Item $installerPath -Force
        
        # Verifica se instalou corretamente
        if (Test-Path "C:\Program Files\7-Zip\7z.exe") {
            Write-Host "✓ 7-Zip instalado com sucesso!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "❌ Falha ao instalar 7-Zip automaticamente." -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "❌ Erro ao instalar 7-Zip: $_" -ForegroundColor Red
        return $false
    }
}

# Função para limpar arquivos temporários
function Clean-TempFiles {
    param (
        [string]$folder
    )
    
    # Lista de arquivos para limpar
    $filesToClean = @(
        "7z.error",
        "7z.log",
        "7z_list.error",
        "7z_list.log",
        "programa.zip",
        "programa.rar"
    )
    
    foreach ($file in $filesToClean) {
        $filePath = Join-Path $folder $file
        if (Test-Path $filePath) {
            Remove-Item $filePath -Force
            Write-Host "Arquivo temporário removido: $file" -ForegroundColor Gray
        }
    }
}

# Define o link padrão do Dropbox
$url = "https://www.dropbox.com/scl/fi/yje55jikt3can3g3unmru/servidordownload.rar?rlkey=t8okqd5jfamelgp3cttjglqbn&dl=1"

# Cria pasta SVteste dentro de Downloads
$downloadsFolder = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
$tempFolder = Join-Path $downloadsFolder "SVteste"
New-Item -ItemType Directory -Force -Path $tempFolder | Out-Null

try {
    Show-Banner "EXTRATOR DE ARQUIVOS"
    
    # Verifica se o 7-Zip está instalado
    $archiveTool = Test-7Zip
    if (-not $archiveTool) {
        if (Install-7Zip) {
            $archiveTool = @{
                Name = "7-Zip"
                Path = "C:\Program Files\7-Zip\7z.exe"
            }
        } else {
            throw "Não foi possível instalar o 7-Zip automaticamente.`nPor favor, instale manualmente de www.7-zip.org"
        }
    }
    
    Write-Host "`nUsando 7-Zip para extrair o arquivo" -ForegroundColor Green

    # Download do arquivo
    Show-Banner "DOWNLOAD DO ARQUIVO"
    Write-Host "📥 Baixando arquivo do servidor..." -ForegroundColor Yellow
    Write-Host "📁 Local de destino: Downloads\SVteste" -ForegroundColor Yellow
    $outputPath = Join-Path $tempFolder "programa.rar"
    
    # Download com retry e verificações adicionais
    $maxRetries = 3
    $retryCount = 0
    $success = $false
    
    while (-not $success -and $retryCount -lt $maxRetries) {
        try {
            # Limpa arquivo anterior se existir
            if (Test-Path $outputPath) {
                Write-Host "🗑️ Removendo download anterior..." -ForegroundColor Gray
                Remove-Item $outputPath -Force
            }

            # Inicia o download
            Write-Host "`n⏳ Baixando arquivo..." -ForegroundColor Yellow
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($url, $outputPath)
            
            if (Test-Path $outputPath) {
                $fileSize = (Get-Item $outputPath).Length
                if ($fileSize -gt 0) {
                    # Verifica os primeiros bytes do arquivo para confirmar que é RAR
                    $bytes = Get-Content $outputPath -Encoding Byte -TotalCount 4
                    if ($bytes[0] -eq 0x52 -and $bytes[1] -eq 0x61 -and $bytes[2] -eq 0x72) {
                        $success = $true
                        Show-Banner "DOWNLOAD CONCLUÍDO"
                        Write-Host "✓ Tamanho do arquivo: $([math]::Round($fileSize/1MB, 2)) MB" -ForegroundColor Green
                        Write-Host "✓ Arquivo RAR verificado com sucesso" -ForegroundColor Green
                        Write-Host "✓ Salvo em: Downloads\SVteste\programa.rar" -ForegroundColor Green
                    } else {
                        Write-Host "`n❌ Arquivo baixado não é um RAR válido" -ForegroundColor Red
                        Remove-Item $outputPath -Force
                    }
                } else {
                    Write-Host "`n❌ Arquivo vazio" -ForegroundColor Red
                    Remove-Item $outputPath -Force
                }
            }
        } catch {
            $retryCount++
            Write-Host "`n❌ Erro no download: $_" -ForegroundColor Red
            if ($retryCount -lt $maxRetries) {
                Write-Host "↻ Tentativa $retryCount de $maxRetries - Aguarde..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
        }
    }
    
    if (-not $success) {
        throw "❌ Não foi possível baixar o arquivo após $maxRetries tentativas"
    }

    # Verifica se o arquivo é realmente um RAR
    Show-Banner "VERIFICANDO ARQUIVO"
    Write-Host "🔍 Analisando arquivo RAR..." -ForegroundColor Yellow
    
    $listProcess = Start-Process -FilePath $archiveTool.Path -ArgumentList "l", $outputPath -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$tempFolder\7z_list.log" -RedirectStandardError "$tempFolder\7z_list.error"
    
    if ($listProcess.ExitCode -ne 0) {
        Show-Banner "ERRO NA VERIFICAÇÃO"
        Write-Host "❌ Arquivo inválido ou corrompido" -ForegroundColor Red
        if (Test-Path "$tempFolder\7z_list.error") {
            Write-Host "Detalhes do erro:" -ForegroundColor Red
            Get-Content "$tempFolder\7z_list.error"
        }
        throw "O arquivo não é um RAR válido"
    }
    
    # Extrai o arquivo
    Show-Banner "EXTRAINDO ARQUIVOS"
    Write-Host "📦 Descompactando arquivos..." -ForegroundColor Yellow
    Write-Host "📁 Local: Downloads\SVteste" -ForegroundColor Yellow
    
    # Inicia a extração
    $extractProcess = Start-Process -FilePath $archiveTool.Path -ArgumentList "x", "-y", "-o$tempFolder", $outputPath -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$tempFolder\7z.log" -RedirectStandardError "$tempFolder\7z.error"
    
    # Verifica o resultado da extração
    if ($extractProcess.ExitCode -eq 0) {
        Show-Banner "EXTRAÇÃO CONCLUÍDA"
        Write-Host "✓ Todos os arquivos foram extraídos" -ForegroundColor Green
        Write-Host "✓ Local: Downloads\SVteste" -ForegroundColor Green
        
        # Lista todos os arquivos extraídos
        Show-Banner "ARQUIVOS EXTRAÍDOS"
        Get-ChildItem -Path $tempFolder -Recurse | ForEach-Object {
            $relativePath = $_.FullName.Replace($tempFolder, "").TrimStart("\")
            if ($_.PSIsContainer) {
                Write-Host " [📁] $relativePath" -ForegroundColor Cyan
            } else {
                $extension = $_.Extension.ToLower()
                $size = [math]::Round($_.Length/1KB, 2)
                Write-Host " [📄] $relativePath ($size KB)" -ForegroundColor Green
            }
        }
        
        Show-Banner "LIMPEZA"
        Write-Host "🧹 Removendo arquivos temporários..." -ForegroundColor Yellow
        Clean-TempFiles -folder $tempFolder
        
        # Verifica se o arquivo RAR foi removido
        if (Test-Path $outputPath) {
            Write-Host "🗑️ Removendo arquivo RAR original..." -ForegroundColor Yellow
            Remove-Item $outputPath -Force
            Write-Host "✓ Arquivo RAR removido" -ForegroundColor Green
        }
        
        Show-Banner "PROCESSO CONCLUÍDO"
        Write-Host "✨ Todos os arquivos estão em: Downloads\SVteste" -ForegroundColor Green
        Write-Host "✨ Processo finalizado com sucesso!" -ForegroundColor Green
    } else {
        Show-Banner "ERRO NA EXTRAÇÃO"
        Write-Host "❌ Falha ao extrair os arquivos" -ForegroundColor Red
        if (Test-Path "$tempFolder\7z.error") {
            Write-Host "Detalhes do erro:" -ForegroundColor Red
            Get-Content "$tempFolder\7z.error"
        }
        throw "Não foi possível extrair o arquivo RAR"
    }
    
} catch {
    Show-Banner "ERRO"
    Write-Host "❌ $($_)" -ForegroundColor Red
    Write-Host "`nPressione qualquer tecla para continuar..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
} 
