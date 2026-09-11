# =========================================================
# Impressao automatica - Entrada -> Impresso / Erro
# =========================================================

function Obter-PastaBase {
    try {
        $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($exePath -notmatch 'powershell(_ise)?\.exe$') { return Split-Path $exePath -Parent }
    } catch { }
    if ($PSScriptRoot) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return Split-Path $MyInvocation.MyCommand.Path -Parent }
    return (Get-Location).Path
}
$PastaBase = Obter-PastaBase

# --- Pastas (relativas a onde o programa esta instalado) ---
$Entrada    = Join-Path $PastaBase "Entrada"
$Impresso   = Join-Path $PastaBase "Impresso"
$ErroPasta  = Join-Path $PastaBase "Erro"
$LogPasta   = Join-Path $PastaBase "Log"
$LogArquivo = Join-Path $LogPasta "impressao.log"
$FolhaBranco = Join-Path $PastaBase "Separador\Branco.pdf"

# --- Configuracoes de impressao (ajuste conforme sua necessidade) ---
$NomeImpressora        = ""     # vazio = usa a impressora padrao do Windows. Ou informe o nome exato, ex: "HP LaserJet M404"
$Copias                = 1      # quantas vezes repetir a impressao de cada arquivo
$UsarSeparador         = $true  # imprime a folha em branco entre um documento e outro
$ForcarDuplex          = $false # tenta configurar frente-e-verso na impressora antes de imprimir (pode exigir executar como administrador)
$MaxTentativas         = 3      # tentativas antes de desistir do arquivo e mover para a pasta Erro
$TimeoutSpoolerSegundos = 150   # tempo maximo esperando a fila de impressao liberar um arquivo grande (livro)
$EsperaSemDeteccaoSegundos = 20    # se o job nunca aparecer na fila (comum em algumas impressoras de rede), segue em frente apos esse tempo - Word/Excel/verbo generico
$EsperaSemDeteccaoPDFSegundos = 60 # mesma logica, mas para PDF via Acrobat, que demora mais pra abrir/renderizar antes de mandar pra fila, especialmente em documentos com varias paginas
$IntervaloVarreduraSegundos = 10 # intervalo entre cada verificacao da pasta Entrada

# --- Preparacao ---
foreach ($pasta in @($Entrada, $Impresso, $ErroPasta, $LogPasta)) {
    if (-not (Test-Path $pasta)) { New-Item -ItemType Directory -Path $pasta -Force | Out-Null }
}

function Write-Log {
    param([string]$Mensagem)
    $linha = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Mensagem"
    Write-Host $linha
    try { Add-Content -Path $LogArquivo -Value $linha -Encoding UTF8 } catch {}
}

function Aguardar-Spooler {
    param(
        [string]$NomeArquivo,
        [int]$TimeoutSegundos = 60,
        [int]$EsperaSemDeteccaoOverride = -1
    )
    $espera = if ($EsperaSemDeteccaoOverride -ge 0) { $EsperaSemDeteccaoOverride } else { $EsperaSemDeteccaoSegundos }
    $baseNome = [System.IO.Path]::GetFileNameWithoutExtension($NomeArquivo)
    $inicio = Get-Date
    $jobVisto = $false
    while ((New-TimeSpan -Start $inicio -End (Get-Date)).TotalSeconds -lt $TimeoutSegundos) {
        $decorridos = (New-TimeSpan -Start $inicio -End (Get-Date)).TotalSeconds
        if (-not $jobVisto -and $decorridos -ge $espera) {
            # Em algumas impressoras de rede o job nunca aparece no Win32_PrintJob mesmo imprimindo normalmente.
            # Depois do tempo padrao de processamento, segue em frente em vez de esperar o timeout inteiro.
            return
        }
        try {
            $jobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop |
                    Where-Object { $_.Document -like "*$baseNome*" }
        } catch {
            # WMI indisponivel neste sistema: usa espera fixa de seguranca
            Start-Sleep -Seconds 5
            return
        }
        if ($jobs -and $jobs.Count -gt 0) {
            $jobVisto = $true
        } elseif ($jobVisto) {
            return
        }
        Start-Sleep -Seconds 1
    }
    if ($jobVisto) {
        Write-Log "AVISO: tempo limite esperando o spooler liberar '$NomeArquivo'. Seguindo em frente."
    } else {
        Write-Log "AVISO: nenhum job de impressao foi detectado na fila para '$NomeArquivo' (confira a impressora fisicamente)."
    }
}

function Fechar-ProcessosNovos {
    param([int[]]$PidsAntes)
    $nomesAlvo = 'AcroRd32', 'Acrobat', 'WINWORD', 'EXCEL', 'POWERPNT'
    Start-Sleep -Seconds 1
    $processosAtuais = Get-Process -Name $nomesAlvo -ErrorAction SilentlyContinue
    foreach ($p in $processosAtuais) {
        if ($PidsAntes -notcontains $p.Id) {
            try {
                Stop-Process -Id $p.Id -Force -ErrorAction Stop
                Write-Log "Processo residual fechado: $($p.ProcessName) (PID $($p.Id))"
            } catch {
                Write-Log "AVISO: nao foi possivel fechar $($p.ProcessName) (PID $($p.Id)): $($_.Exception.Message)"
            }
        }
    }
}

function Resolver-ImpressoraAlvo {
    if ($NomeImpressora -ne "") { return $NomeImpressora }
    $p = Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue | Where-Object { $_.Default -eq $true }
    if ($p) { return $p.Name }
    return $null
}

function Obter-AcrobatExecutavel {
    $caminhos = @(
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe"
    )
    foreach ($c in $caminhos) {
        try {
            $v = (Get-ItemProperty -Path $c -ErrorAction Stop).'(default)'
            if ($v -and (Test-Path -LiteralPath $v)) { return $v }
        } catch { }
    }
    $fixos = @(
        "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    )
    foreach ($f in $fixos) { if (Test-Path $f) { return $f } }
    return $null
}
$AcrobatExe = Obter-AcrobatExecutavel
if ($AcrobatExe) { Write-Log "PDF: impressao silenciosa via Adobe Acrobat habilitada ($AcrobatExe)" }
else { Write-Log "AVISO: Adobe Acrobat nao encontrado; PDFs vao usar o verbo de impressao padrao do Windows (pode abrir dialogo)." }

function Imprimir-UmaCopia {
    param([string]$CaminhoArquivo, [string]$NomeParaSpooler)

    $ehPdf = [System.IO.Path]::GetExtension($CaminhoArquivo).ToLower() -eq ".pdf"

    if ($ehPdf -and $AcrobatExe) {
        $impressoraAlvo = Resolver-ImpressoraAlvo
        if ($impressoraAlvo) {
            $argList = @("/t", $CaminhoArquivo, $impressoraAlvo)
            Start-Process -FilePath $AcrobatExe -ArgumentList $argList -WindowStyle Minimized -ErrorAction Stop
            Aguardar-Spooler -NomeArquivo $NomeParaSpooler -TimeoutSegundos $TimeoutSpoolerSegundos -EsperaSemDeteccaoOverride $EsperaSemDeteccaoPDFSegundos
            return
        }
    }

    Start-Process -FilePath $CaminhoArquivo -Verb Print -WindowStyle Minimized -ErrorAction Stop
    Aguardar-Spooler -NomeArquivo $NomeParaSpooler -TimeoutSegundos $TimeoutSpoolerSegundos
}

# --- Impressora padrao / duplex (opcional) ---
if ($NomeImpressora -ne "") {
    try {
        $wshNetwork = New-Object -ComObject WScript.Network
        $wshNetwork.SetDefaultPrinter($NomeImpressora)
        Write-Log "Impressora padrao definida para: $NomeImpressora"
    } catch {
        Write-Log "AVISO: nao foi possivel definir a impressora '$NomeImpressora' como padrao: $($_.Exception.Message)"
    }
}

if ($ForcarDuplex) {
    try {
        Import-Module PrintManagement -ErrorAction Stop
        $printerAlvo = if ($NomeImpressora -ne "") { $NomeImpressora } else { (Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }).Name }
        Set-PrintConfiguration -PrinterName $printerAlvo -DuplexingMode TwoSidedLongEdge -ErrorAction Stop
        Write-Log "Frente e verso configurado na impressora: $printerAlvo"
    } catch {
        Write-Log "AVISO: nao foi possivel configurar frente e verso automaticamente ($($_.Exception.Message)). Configure manualmente nas propriedades da impressora, se necessario."
    }
}

Write-Log "=== Script de impressao iniciado ==="

$tentativas = @{}

while ($true) {

    $arquivos = Get-ChildItem -Path $Entrada -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notlike '~$*' } |
                Sort-Object Name

    foreach ($arquivo in $arquivos) {
        $nome   = $arquivo.Name
        $origem = $arquivo.FullName

        if (-not $tentativas.ContainsKey($nome)) { $tentativas[$nome] = 0 }

        try {
            Write-Log "Iniciando impressao: $nome (tentativa $($tentativas[$nome] + 1)/$MaxTentativas)"

            $pidsAntes = (Get-Process -Name AcroRd32, Acrobat, WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue).Id

            for ($c = 1; $c -le $Copias; $c++) {
                Imprimir-UmaCopia -CaminhoArquivo $origem -NomeParaSpooler $nome
            }

            if ($UsarSeparador -and (Test-Path $FolhaBranco)) {
                Imprimir-UmaCopia -CaminhoArquivo $FolhaBranco -NomeParaSpooler "Branco.pdf"
            }

            Fechar-ProcessosNovos -PidsAntes $pidsAntes

            $destino = Join-Path $Impresso $nome
            Move-Item -LiteralPath $origem -Destination $destino -Force -ErrorAction Stop

            Write-Log "OK: '$nome' impresso e movido para Impresso."
            $tentativas.Remove($nome)

        } catch {
            $tentativas[$nome]++
            Write-Log "ERRO ao imprimir '$nome' (tentativa $($tentativas[$nome])/$MaxTentativas): $($_.Exception.Message)"

            if ($tentativas[$nome] -ge $MaxTentativas) {
                try {
                    $destinoErro = Join-Path $ErroPasta $nome
                    Move-Item -LiteralPath $origem -Destination $destinoErro -Force -ErrorAction Stop
                    $motivoArquivo = Join-Path $ErroPasta "$nome.erro.txt"
                    "Falhou apos $MaxTentativas tentativas.`r`nUltimo erro: $($_.Exception.Message)`r`nData: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" |
                        Out-File -FilePath $motivoArquivo -Encoding UTF8
                    Write-Log "'$nome' movido para a pasta Erro apos $MaxTentativas tentativas."
                } catch {
                    Write-Log "FALHA CRITICA: nao foi possivel mover '$nome' para a pasta Erro: $($_.Exception.Message)"
                }
                $tentativas.Remove($nome)
            }
        }
    }

    Start-Sleep -Seconds $IntervaloVarreduraSegundos
}
