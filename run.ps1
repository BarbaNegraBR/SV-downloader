# Função para verificar se o 7-Zip ou WinRAR está instalado
function Test-ArchiveTools {
    $tools = @{
        "7-Zip" = "C:\Program Files\7-Zip\7z.exe"
        "WinRAR" = "C:\Program Files\WinRAR\WinRAR.exe"
    }
    
    foreach ($tool in $tools.GetEnumerator()) {
        if (Test-Path $tool.Value) {
            return @{
                Name = $tool.Key
                Path = $tool.Value
            }
        }
    }
    
    return $null
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
    # Verifica qual ferramenta de arquivo está instalada
    $archiveTool = Test-ArchiveTools
    if (-not $archiveTool) {
        throw "Nenhuma ferramenta de arquivo encontrada. Por favor, instale o 7-Zip (www.7-zip.org) ou o WinRAR (www.rarlab.com)"
    }
    
    Write-Host "Usando $($archiveTool.Name) para extrair o arquivo" -ForegroundColor Yellow

    # Download do arquivo
    Write-Host "Baixando arquivo de: $url" -ForegroundColor Yellow
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
    
    # Lista o conteúdo antes de tentar extrair
    Write-Host "Listando conteúdo do arquivo..." -ForegroundColor Yellow
    
    if ($archiveTool.Name -eq "7-Zip") {
        $listProcess = Start-Process -FilePath $archiveTool.Path -ArgumentList "l", $outputPath -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$tempFolder\7z_list.log" -RedirectStandardError "$tempFolder\7z_list.error"
    } else {
        $listProcess = Start-Process -FilePath $archiveTool.Path -ArgumentList "l", $outputPath -NoNewWindow -PassThru -Wait
    }
    
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
    
    if ($archiveTool.Name -eq "7-Zip") {
        $extractProcess = Start-Process -FilePath $archiveTool.Path -ArgumentList "x", "-y", "-o$tempFolder", $outputPath -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$tempFolder\7z.log" -RedirectStandardError "$tempFolder\7z.error"
    } else {
        $extractProcess = Start-Process -FilePath $archiveTool.Path -ArgumentList "x", "-y", "-o$tempFolder", $outputPath -NoNewWindow -PassThru -Wait
    }
    
    # Verifica o resultado da extração
    if ($extractProcess.ExitCode -eq 0) {
        Write-Host "Arquivo extraído com sucesso!" -ForegroundColor Green
        Write-Host "Arquivos extraídos em: $tempFolder" -ForegroundColor Green
        
        # Lista todos os arquivos extraídos
        Write-Host "`nArquivos encontrados na pasta:" -ForegroundColor Yellow
        Get-ChildItem -Path $tempFolder -Recurse | ForEach-Object {
            $relativePath = $_.FullName.Replace($tempFolder, "").TrimStart("\")
            if ($_.PSIsContainer) {
                Write-Host " [Pasta] $relativePath" -ForegroundColor Cyan
            } else {
                $extension = $_.Extension.ToLower()
                $size = [math]::Round($_.Length/1KB, 2)
                Write-Host " [Arquivo] $relativePath ($size KB)" -ForegroundColor Green
            }
        }
        
        Write-Host "`nOs arquivos foram extraídos com sucesso para: $tempFolder" -ForegroundColor Green
        Write-Host "Você pode encontrar os arquivos na pasta Downloads\SVteste" -ForegroundColor Yellow
        
        # Limpa arquivos temporários após mostrar os arquivos extraídos
        Write-Host "`nLimpando arquivos temporários..." -ForegroundColor Yellow
        Clean-TempFiles -folder $tempFolder
        
        # Verifica se o arquivo RAR foi removido
        if (Test-Path $outputPath) {
            Write-Host "Removendo arquivo RAR..." -ForegroundColor Yellow
            Remove-Item $outputPath -Force
            Write-Host "Arquivo RAR removido com sucesso" -ForegroundColor Green
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
