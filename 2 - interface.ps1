Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

# =========================================================
# Estado compartilhado entre a interface (thread principal)
# e o motor de impressao (runspace em segundo plano)
# =========================================================
$sync = [hashtable]::Synchronized(@{})
$sync.Entrada             = Join-Path $PastaBase "Entrada"
$sync.Impresso            = Join-Path $PastaBase "Impresso"
$sync.ErroPasta           = Join-Path $PastaBase "Erro"
$sync.LogPasta            = Join-Path $PastaBase "Log"
$sync.LogArquivo          = Join-Path $sync.LogPasta "impressao.log"
$sync.FolhaBranco         = Join-Path $PastaBase "Separador\Branco.pdf"
$sync.NomeImpressora      = ""
$sync.Copias              = 1
$sync.UsarSeparador       = $true
$sync.ForcarDuplex        = $false
$sync.MaxTentativas       = 3
$sync.TimeoutSpooler      = 150
$sync.EsperaSemDeteccao   = 20   # segundos: se o job nunca aparecer na fila (comum em algumas impressoras de rede), segue em frente apos esse tempo - usado para Word/Excel/verbo generico
$sync.EsperaSemDeteccaoPDF = 60  # segundos: mesma logica, mas para PDF via Acrobat, que demora mais pra abrir/renderizar antes de mandar pra fila, especialmente em documentos com varias paginas
$sync.IntervaloVarredura  = 10
$sync.AutomacaoAtiva      = $true
$sync.Encerrar            = $false
$sync.UltimaVarredura     = ""
$sync.LogQueue            = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$sync.ManualQueue         = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$sync.StatusPorArquivo    = [hashtable]::Synchronized(@{})

foreach ($pasta in @($sync.Entrada, $sync.Impresso, $sync.ErroPasta, $sync.LogPasta)) {
    if (-not (Test-Path $pasta)) { New-Item -ItemType Directory -Path $pasta -Force | Out-Null }
}

# =========================================================
# Motor de impressao (roda em runspace separado, nao trava a janela)
# =========================================================
$workerScript = {
    param($sync)

    function Write-LogSync {
        param([string]$Mensagem)
        $linha = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Mensagem"
        $sync.LogQueue.Enqueue($linha)
        try { Add-Content -Path $sync.LogArquivo -Value $linha -Encoding UTF8 } catch {}
    }

    function Aguardar-Spooler {
        param([string]$NomeArquivo, [int]$TimeoutSegundos = 60, [int]$EsperaSemDeteccaoOverride = -1)
        $espera = if ($EsperaSemDeteccaoOverride -ge 0) { $EsperaSemDeteccaoOverride } else { $sync.EsperaSemDeteccao }
        $baseNome = [System.IO.Path]::GetFileNameWithoutExtension($NomeArquivo)
        $inicio = Get-Date
        $jobVisto = $false
        while ((New-TimeSpan -Start $inicio -End (Get-Date)).TotalSeconds -lt $TimeoutSegundos) {
            $decorridos = (New-TimeSpan -Start $inicio -End (Get-Date)).TotalSeconds
            if (-not $jobVisto -and $decorridos -ge $espera) {
                # Nesta impressora o job pode nunca aparecer no Win32_PrintJob mesmo imprimindo normalmente.
                # Depois do tempo padrao de processamento, segue em frente em vez de esperar o timeout inteiro.
                return
            }
            try {
                $jobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop |
                        Where-Object { $_.Document -like "*$baseNome*" }
            } catch {
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
            Write-LogSync "AVISO: tempo limite esperando o spooler liberar '$NomeArquivo'."
        } else {
            Write-LogSync "AVISO: nenhum job de impressao foi detectado na fila para '$NomeArquivo' (confira a impressora fisicamente)."
        }
    }

    function Fechar-ProcessosNovos {
        param([int[]]$PidsAntes)
        $nomesAlvo = 'AcroRd32', 'Acrobat', 'WINWORD', 'EXCEL', 'POWERPNT'
        Start-Sleep -Seconds 1
        foreach ($p in (Get-Process -Name $nomesAlvo -ErrorAction SilentlyContinue)) {
            if ($PidsAntes -notcontains $p.Id) {
                try { Stop-Process -Id $p.Id -Force -ErrorAction Stop } catch {}
            }
        }
    }

    function Resolver-ImpressoraAlvo {
        if ($sync.NomeImpressora -and $sync.NomeImpressora -ne "") { return $sync.NomeImpressora }
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
    $script:acrobatExe = Obter-AcrobatExecutavel
    if ($script:acrobatExe) { Write-LogSync "PDF: impressao silenciosa via Adobe Acrobat habilitada ($($script:acrobatExe))" }
    else { Write-LogSync "AVISO: Adobe Acrobat nao encontrado; PDFs vao usar o verbo de impressao padrao do Windows (pode abrir dialogo)." }

    function Imprimir-UmaCopia {
        param([string]$CaminhoArquivo, [string]$NomeParaSpooler)

        $ehPdf = [System.IO.Path]::GetExtension($CaminhoArquivo).ToLower() -eq ".pdf"

        if ($ehPdf -and $script:acrobatExe) {
            $impressoraAlvo = Resolver-ImpressoraAlvo
            if ($impressoraAlvo) {
                $argList = @("/t", $CaminhoArquivo, $impressoraAlvo)
                Start-Process -FilePath $script:acrobatExe -ArgumentList $argList -WindowStyle Minimized -ErrorAction Stop
                Aguardar-Spooler -NomeArquivo $NomeParaSpooler -TimeoutSegundos $sync.TimeoutSpooler -EsperaSemDeteccaoOverride $sync.EsperaSemDeteccaoPDF
                return
            }
        }

        Start-Process -FilePath $CaminhoArquivo -Verb Print -WindowStyle Minimized -ErrorAction Stop
        Aguardar-Spooler -NomeArquivo $NomeParaSpooler -TimeoutSegundos $sync.TimeoutSpooler
    }

    function Imprimir-Documento {
        param([string]$CaminhoArquivo)
        $nome = Split-Path $CaminhoArquivo -Leaf
        $pidsAntes = (Get-Process -Name AcroRd32, Acrobat, WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue).Id
        for ($c = 1; $c -le [int]$sync.Copias; $c++) {
            Imprimir-UmaCopia -CaminhoArquivo $CaminhoArquivo -NomeParaSpooler $nome
        }
        if ($sync.UsarSeparador -and (Test-Path $sync.FolhaBranco)) {
            Imprimir-UmaCopia -CaminhoArquivo $sync.FolhaBranco -NomeParaSpooler "Branco.pdf"
        }
        Fechar-ProcessosNovos -PidsAntes $pidsAntes
    }

    function Drain-ManualQueue {
        $caminho = $null
        while ($sync.ManualQueue.TryDequeue([ref]$caminho)) {
            $nome = Split-Path $caminho -Leaf
            $sync.StatusPorArquivo[$caminho] = "Imprimindo..."
            try {
                Imprimir-Documento -CaminhoArquivo $caminho
                $sync.StatusPorArquivo[$caminho] = "Impresso"
                Write-LogSync "OK (manual): $nome"
            } catch {
                $sync.StatusPorArquivo[$caminho] = "Erro: $($_.Exception.Message)"
                Write-LogSync "ERRO (manual) '$nome': $($_.Exception.Message)"
            }
        }
    }

    if ($sync.NomeImpressora -and $sync.NomeImpressora -ne "") {
        try {
            Import-Module PrintManagement -ErrorAction Stop
            $printerAlvo = $sync.NomeImpressora
            Set-PrintConfiguration -PrinterName $printerAlvo -DuplexingMode TwoSidedLongEdge -ErrorAction Stop
        } catch { }
    }

    $tentativasAuto = @{}
    $ultimaImpressoraAplicada = ""
    Write-LogSync "=== Motor de impressao iniciado ==="

    while (-not $sync.Encerrar) {

        if ($sync.NomeImpressora -and $sync.NomeImpressora -ne "" -and $sync.NomeImpressora -ne $ultimaImpressoraAplicada) {
            try {
                (New-Object -ComObject WScript.Network).SetDefaultPrinter($sync.NomeImpressora)
                $ultimaImpressoraAplicada = $sync.NomeImpressora
                Write-LogSync "Impressora padrao definida para: $($sync.NomeImpressora)"
                if ($sync.ForcarDuplex) {
                    try {
                        Import-Module PrintManagement -ErrorAction Stop
                        Set-PrintConfiguration -PrinterName $sync.NomeImpressora -DuplexingMode TwoSidedLongEdge -ErrorAction Stop
                        Write-LogSync "Frente e verso configurado em: $($sync.NomeImpressora)"
                    } catch {
                        Write-LogSync "AVISO: nao foi possivel configurar frente e verso: $($_.Exception.Message)"
                    }
                }
            } catch {
                Write-LogSync "AVISO: nao foi possivel definir impressora '$($sync.NomeImpressora)': $($_.Exception.Message)"
            }
        }

        Drain-ManualQueue

        if ($sync.AutomacaoAtiva) {
            $arquivos = Get-ChildItem -Path $sync.Entrada -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name

            foreach ($arquivo in $arquivos) {
                $nome = $arquivo.Name
                $origem = $arquivo.FullName
                if (-not $tentativasAuto.ContainsKey($nome)) { $tentativasAuto[$nome] = 0 }

                try {
                    Write-LogSync "Iniciando impressao automatica: $nome (tentativa $($tentativasAuto[$nome] + 1)/$($sync.MaxTentativas))"
                    Imprimir-Documento -CaminhoArquivo $origem
                    $destino = Join-Path $sync.Impresso $nome
                    Move-Item -LiteralPath $origem -Destination $destino -Force -ErrorAction Stop
                    Write-LogSync "OK: '$nome' impresso e movido para Impresso."
                    $tentativasAuto.Remove($nome)
                } catch {
                    $tentativasAuto[$nome]++
                    Write-LogSync "ERRO ao imprimir '$nome' (tentativa $($tentativasAuto[$nome])/$($sync.MaxTentativas)): $($_.Exception.Message)"
                    if ($tentativasAuto[$nome] -ge $sync.MaxTentativas) {
                        try {
                            $destinoErro = Join-Path $sync.ErroPasta $nome
                            Move-Item -LiteralPath $origem -Destination $destinoErro -Force -ErrorAction Stop
                            "Falhou apos $($sync.MaxTentativas) tentativas.`r`nUltimo erro: $($_.Exception.Message)`r`nData: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" |
                                Out-File -FilePath (Join-Path $sync.ErroPasta "$nome.erro.txt") -Encoding UTF8
                            Write-LogSync "'$nome' movido para a pasta Erro apos $($sync.MaxTentativas) tentativas."
                        } catch {
                            Write-LogSync "FALHA CRITICA: nao foi possivel mover '$nome' para Erro: $($_.Exception.Message)"
                        }
                        $tentativasAuto.Remove($nome)
                    }
                }

                Drain-ManualQueue
            }
        }

        $sync.UltimaVarredura = Get-Date -Format 'HH:mm:ss'

        for ($i = 0; $i -lt ($sync.IntervaloVarredura * 2) -and -not $sync.Encerrar; $i++) {
            Start-Sleep -Milliseconds 500
            Drain-ManualQueue
        }
    }
    Write-LogSync "=== Motor de impressao encerrado ==="
}

$runspace = [runspacefactory]::CreateRunspace()
$runspace.ApartmentState = "MTA"
$runspace.ThreadOptions = "ReuseThread"
$runspace.Open()
$motorPS = [powershell]::Create()
$motorPS.Runspace = $runspace
[void]$motorPS.AddScript($workerScript).AddArgument($sync)
$motorHandle = $motorPS.BeginInvoke()

# =========================================================
# Interface grafica
# =========================================================
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "Impressao Automatica"
$form.ClientSize = New-Object System.Drawing.Size(800, 610)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(720, 500)

# --- Impressora ---
$grpImpressora = New-Object System.Windows.Forms.GroupBox
$grpImpressora.Text = "Selecionar Impressora"
$grpImpressora.Location = New-Object System.Drawing.Point(10, 10)
$grpImpressora.Size = New-Object System.Drawing.Size(230, 90)

$comboImpressora = New-Object System.Windows.Forms.ComboBox
$comboImpressora.Location = New-Object System.Drawing.Point(10, 25)
$comboImpressora.Size = New-Object System.Drawing.Size(210, 24)
$comboImpressora.DropDownStyle = 'DropDownList'
foreach ($imp in [System.Drawing.Printing.PrinterSettings]::InstalledPrinters) {
    [void]$comboImpressora.Items.Add($imp)
}
$padrao = (New-Object System.Drawing.Printing.PrinterSettings).PrinterName
if ($comboImpressora.Items.Contains($padrao)) { $comboImpressora.SelectedItem = $padrao }
$sync.NomeImpressora = $comboImpressora.SelectedItem
$comboImpressora.Add_SelectedIndexChanged({ $sync.NomeImpressora = $comboImpressora.SelectedItem })

$btnPropriedades = New-Object System.Windows.Forms.Button
$btnPropriedades.Text = "Propriedades da Impressora"
$btnPropriedades.Location = New-Object System.Drawing.Point(10, 55)
$btnPropriedades.Size = New-Object System.Drawing.Size(210, 24)
$btnPropriedades.Add_Click({
    if ($comboImpressora.SelectedItem) {
        Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /p /n `"$($comboImpressora.SelectedItem)`""
    }
})

$grpImpressora.Controls.AddRange(@($comboImpressora, $btnPropriedades))

# --- Configuracoes ---
$grpConfig = New-Object System.Windows.Forms.GroupBox
$grpConfig.Text = "Configuracoes"
$grpConfig.Location = New-Object System.Drawing.Point(10, 110)
$grpConfig.Size = New-Object System.Drawing.Size(230, 155)

$lblCopias = New-Object System.Windows.Forms.Label
$lblCopias.Text = "Copias:"
$lblCopias.Location = New-Object System.Drawing.Point(10, 25)
$lblCopias.Size = New-Object System.Drawing.Size(55, 20)

$numCopias = New-Object System.Windows.Forms.NumericUpDown
$numCopias.Location = New-Object System.Drawing.Point(75, 23)
$numCopias.Size = New-Object System.Drawing.Size(55, 20)
$numCopias.Minimum = 1
$numCopias.Maximum = 50
$numCopias.Value = 1
$numCopias.Add_ValueChanged({ $sync.Copias = [int]$numCopias.Value })

$chkDuplex = New-Object System.Windows.Forms.CheckBox
$chkDuplex.Text = "Forcar frente e verso"
$chkDuplex.Location = New-Object System.Drawing.Point(10, 55)
$chkDuplex.Size = New-Object System.Drawing.Size(210, 24)
$chkDuplex.Add_CheckedChanged({ $sync.ForcarDuplex = $chkDuplex.Checked })

$chkSeparador = New-Object System.Windows.Forms.CheckBox
$chkSeparador.Text = "Imprimir folha separadora"
$chkSeparador.Checked = $true
$chkSeparador.Location = New-Object System.Drawing.Point(10, 82)
$chkSeparador.Size = New-Object System.Drawing.Size(210, 24)
$chkSeparador.Add_CheckedChanged({ $sync.UsarSeparador = $chkSeparador.Checked })

$chkAutomacao = New-Object System.Windows.Forms.CheckBox
$chkAutomacao.Text = "Automacao da pasta Entrada ativa"
$chkAutomacao.Checked = $true
$chkAutomacao.Location = New-Object System.Drawing.Point(10, 109)
$chkAutomacao.Size = New-Object System.Drawing.Size(210, 24)
$chkAutomacao.Add_CheckedChanged({ $sync.AutomacaoAtiva = $chkAutomacao.Checked })

$grpConfig.Controls.AddRange(@($lblCopias, $numCopias, $chkDuplex, $chkSeparador, $chkAutomacao))

# --- Log ---
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = "Log"
$grpLog.Location = New-Object System.Drawing.Point(10, 275)
$grpLog.Size = New-Object System.Drawing.Size(230, 325)
$grpLog.Anchor = 'Top,Bottom,Left'

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = 'Vertical'
$logBox.ReadOnly = $true
$logBox.Location = New-Object System.Drawing.Point(10, 20)
$logBox.Size = New-Object System.Drawing.Size(210, 295)
$logBox.Anchor = 'Top,Bottom,Left,Right'
$logBox.Font = New-Object System.Drawing.Font("Consolas", 8)
$grpLog.Controls.Add($logBox)

# --- Lista de documentos (fila manual) ---
$grpLista = New-Object System.Windows.Forms.GroupBox
$grpLista.Text = "Lista de Documentos"
$grpLista.Location = New-Object System.Drawing.Point(250, 10)
$grpLista.Size = New-Object System.Drawing.Size(540, 390)
$grpLista.Anchor = 'Top,Bottom,Left,Right'

$btnAddArquivos = New-Object System.Windows.Forms.Button
$btnAddArquivos.Text = "Adicionar Arquivos"
$btnAddArquivos.Location = New-Object System.Drawing.Point(10, 20)
$btnAddArquivos.Size = New-Object System.Drawing.Size(95, 26)

$btnAddPasta = New-Object System.Windows.Forms.Button
$btnAddPasta.Text = "Adicionar Pasta"
$btnAddPasta.Location = New-Object System.Drawing.Point(110, 20)
$btnAddPasta.Size = New-Object System.Drawing.Size(85, 26)

$btnMoverCima = New-Object System.Windows.Forms.Button
$btnMoverCima.Text = [char]0x2191
$btnMoverCima.Location = New-Object System.Drawing.Point(200, 20)
$btnMoverCima.Size = New-Object System.Drawing.Size(40, 26)

$btnMoverBaixo = New-Object System.Windows.Forms.Button
$btnMoverBaixo.Text = [char]0x2193
$btnMoverBaixo.Location = New-Object System.Drawing.Point(245, 20)
$btnMoverBaixo.Size = New-Object System.Drawing.Size(40, 26)

$btnRemover = New-Object System.Windows.Forms.Button
$btnRemover.Text = "Remover"
$btnRemover.Location = New-Object System.Drawing.Point(295, 20)
$btnRemover.Size = New-Object System.Drawing.Size(65, 26)

$btnLimpar = New-Object System.Windows.Forms.Button
$btnLimpar.Text = "Limpar Lista"
$btnLimpar.Location = New-Object System.Drawing.Point(365, 20)
$btnLimpar.Size = New-Object System.Drawing.Size(75, 26)

$btnImprimir = New-Object System.Windows.Forms.Button
$btnImprimir.Text = "Imprimir"
$btnImprimir.Location = New-Object System.Drawing.Point(445, 20)
$btnImprimir.Size = New-Object System.Drawing.Size(85, 26)
$btnImprimir.BackColor = [System.Drawing.Color]::LightSteelBlue
$btnImprimir.Anchor = 'Top,Right'

$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 55)
$listView.Size = New-Object System.Drawing.Size(520, 325)
$listView.View = 'Details'
$listView.CheckBoxes = $true
$listView.FullRowSelect = $true
$listView.AllowDrop = $true
$listView.Anchor = 'Top,Bottom,Left,Right'
[void]$listView.Columns.Add("Nome", 230)
[void]$listView.Columns.Add("Pasta", 160)
[void]$listView.Columns.Add("Status", 110)

function Adicionar-ArquivoNaLista {
    param([string]$Caminho)
    if (-not (Test-Path -LiteralPath $Caminho -PathType Leaf)) { return }
    foreach ($item in $listView.Items) {
        if ($item.Tag -eq $Caminho) { return }
    }
    $info = Get-Item -LiteralPath $Caminho
    $item = New-Object System.Windows.Forms.ListViewItem($info.Name)
    [void]$item.SubItems.Add($info.DirectoryName)
    [void]$item.SubItems.Add("Aguardando")
    $item.Tag = $Caminho
    $item.Checked = $true
    [void]$listView.Items.Add($item)
    $sync.StatusPorArquivo[$Caminho] = "Aguardando"
}

$btnAddArquivos.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Multiselect = $true
    $ofd.Filter = "Todos os arquivos (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') {
        foreach ($f in $ofd.FileNames) { Adicionar-ArquivoNaLista -Caminho $f }
    }
})

$btnAddPasta.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog() -eq 'OK') {
        Get-ChildItem -Path $fbd.SelectedPath -File | ForEach-Object { Adicionar-ArquivoNaLista -Caminho $_.FullName }
    }
})

$btnRemover.Add_Click({
    $selecionados = @($listView.SelectedItems)
    foreach ($item in $selecionados) {
        $sync.StatusPorArquivo.Remove($item.Tag)
        $listView.Items.Remove($item)
    }
})

$btnMoverCima.Add_Click({
    $indices = @($listView.SelectedIndices) | Sort-Object
    foreach ($idx in $indices) {
        if ($idx -eq 0) { continue }
        $item = $listView.Items[$idx]
        $listView.Items.RemoveAt($idx)
        $listView.Items.Insert($idx - 1, $item)
        $item.Selected = $true
    }
})

$btnMoverBaixo.Add_Click({
    $indices = @($listView.SelectedIndices) | Sort-Object -Descending
    foreach ($idx in $indices) {
        if ($idx -ge $listView.Items.Count - 1) { continue }
        $item = $listView.Items[$idx]
        $listView.Items.RemoveAt($idx)
        $listView.Items.Insert($idx + 1, $item)
        $item.Selected = $true
    }
})

$btnLimpar.Add_Click({
    foreach ($item in $listView.Items) { $sync.StatusPorArquivo.Remove($item.Tag) }
    $listView.Items.Clear()
})

$btnImprimir.Add_Click({
    foreach ($item in $listView.Items) {
        if ($item.Checked) {
            $sync.StatusPorArquivo[$item.Tag] = "Na fila"
            $sync.ManualQueue.Enqueue($item.Tag)
        }
    }
})

$listView.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' } else { $e.Effect = 'None' }
})
$listView.Add_DragDrop({
    param($s, $e)
    $caminhos = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    foreach ($c in $caminhos) {
        if (Test-Path -LiteralPath $c -PathType Container) {
            Get-ChildItem -Path $c -File | ForEach-Object { Adicionar-ArquivoNaLista -Caminho $_.FullName }
        } else {
            Adicionar-ArquivoNaLista -Caminho $c
        }
    }
})

$grpLista.Controls.AddRange(@($btnAddArquivos, $btnAddPasta, $btnMoverCima, $btnMoverBaixo, $btnRemover, $btnLimpar, $btnImprimir, $listView))

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($btnMoverCima, "Mover selecionado para cima (ordem de impressao)")
$toolTip.SetToolTip($btnMoverBaixo, "Mover selecionado para baixo (ordem de impressao)")

# --- Rodape: acesso as pastas + status do motor ---
$grpRodape = New-Object System.Windows.Forms.GroupBox
$grpRodape.Text = "Pastas e status"
$grpRodape.Location = New-Object System.Drawing.Point(250, 410)
$grpRodape.Size = New-Object System.Drawing.Size(540, 190)
$grpRodape.Anchor = 'Top,Bottom,Left,Right'

$btnAbrirEntrada = New-Object System.Windows.Forms.Button
$btnAbrirEntrada.Text = "Abrir Entrada"
$btnAbrirEntrada.Location = New-Object System.Drawing.Point(10, 25)
$btnAbrirEntrada.Size = New-Object System.Drawing.Size(160, 26)
$btnAbrirEntrada.Add_Click({ Start-Process explorer.exe $sync.Entrada })

$btnAbrirImpresso = New-Object System.Windows.Forms.Button
$btnAbrirImpresso.Text = "Abrir Impresso"
$btnAbrirImpresso.Location = New-Object System.Drawing.Point(180, 25)
$btnAbrirImpresso.Size = New-Object System.Drawing.Size(160, 26)
$btnAbrirImpresso.Add_Click({ Start-Process explorer.exe $sync.Impresso })

$btnAbrirErro = New-Object System.Windows.Forms.Button
$btnAbrirErro.Text = "Abrir Erro"
$btnAbrirErro.Location = New-Object System.Drawing.Point(350, 25)
$btnAbrirErro.Size = New-Object System.Drawing.Size(160, 26)
$btnAbrirErro.Add_Click({ Start-Process explorer.exe $sync.ErroPasta })

$lblStatusMotor = New-Object System.Windows.Forms.Label
$lblStatusMotor.Text = "Motor de impressao: iniciando..."
$lblStatusMotor.Location = New-Object System.Drawing.Point(10, 65)
$lblStatusMotor.Size = New-Object System.Drawing.Size(510, 60)
$lblStatusMotor.Anchor = 'Top,Left,Right'

$grpRodape.Controls.AddRange(@($btnAbrirEntrada, $btnAbrirImpresso, $btnAbrirErro, $lblStatusMotor))

$form.Controls.AddRange(@($grpImpressora, $grpConfig, $grpLog, $grpLista, $grpRodape))

# --- Timer de atualizacao da interface (le o estado compartilhado, sem travar a janela) ---
$timerUI = New-Object System.Windows.Forms.Timer
$timerUI.Interval = 500
$timerUI.Add_Tick({
    $linha = $null
    while ($sync.LogQueue.TryDequeue([ref]$linha)) {
        $logBox.AppendText("$linha`r`n")
    }

    foreach ($item in $listView.Items) {
        $caminho = $item.Tag
        if ($sync.StatusPorArquivo.ContainsKey($caminho)) {
            $novoStatus = $sync.StatusPorArquivo[$caminho]
            if ($item.SubItems[2].Text -ne $novoStatus) {
                $item.SubItems[2].Text = $novoStatus
            }
        }
    }

    $pendentes = 0
    try { $pendentes = (Get-ChildItem -Path $sync.Entrada -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '~$*' }).Count } catch {}
    $estadoAuto = if ($sync.AutomacaoAtiva) { "ativa" } else { "pausada" }
    $lblStatusMotor.Text = "Motor de impressao: rodando`r`nAutomacao da pasta Entrada: $estadoAuto`r`nArquivos aguardando em Entrada: $pendentes`r`nUltima verificacao: $($sync.UltimaVarredura)"
})
$timerUI.Start()

$form.Add_FormClosing({
    $timerUI.Stop()
    $sync.Encerrar = $true
    Start-Sleep -Milliseconds 400
    try { $motorPS.Stop() } catch {}
    try { $motorPS.Dispose() } catch {}
    try { $runspace.Close() } catch {}
})

[void]$form.ShowDialog()
