# Impressao Automatica

Ferramenta para Windows que imprime arquivos automaticamente a partir de uma pasta monitorada, com opcao de interface grafica para fila manual (arrastar e soltar, escolher impressora, reordenar).

## Como funciona

- **Automacao de pasta**: qualquer arquivo colocado em `Entrada/` e impresso automaticamente e movido para `Impresso/` (ou para `Erro/` apos falhas repetidas, com um `.txt` explicando o motivo).
- **Fila manual** (na interface grafica): arraste arquivos ou pastas, escolha a impressora, reordene com os botoes de seta e clique em Imprimir.
- Entre cada documento, opcionalmente imprime uma folha separadora (`Separador/Branco.pdf`).
- Log completo em `Log/impressao.log`.

## Arquivos

| Arquivo | Descricao |
|---|---|
| `1 - impressao.ps1` | Motor de impressao sem interface (roda em segundo plano, so automacao de pasta). |
| `2 - interface.ps1` | Interface grafica (WinForms) com fila manual + automacao de pasta rodando junto. |
| `Impressao.exe` | Versao compilada de `2 - interface.ps1` (gerada com [ps2exe](https://github.com/MScholtes/PS2EXE), nao versionada no repositorio). |

## Requisitos

- Windows com PowerShell.
- **Adobe Acrobat** (Reader ou DC) instalado, para impressao silenciosa de PDF sem abrir dialogos. Sem ele, PDFs caem no verbo de impressao padrao do Windows, que pode nao funcionar silenciosamente dependendo do leitor de PDF padrao configurado.

## Uso

Rodar direto o script:

```powershell
powershell -ExecutionPolicy Bypass -File "2 - interface.ps1"
```

Ou compilar para `.exe` com [ps2exe](https://www.powershellgallery.com/packages/ps2exe):

```powershell
Install-Module ps2exe -Scope CurrentUser
Invoke-ps2exe -inputFile "2 - interface.ps1" -outputFile "Impressao.exe" -STA -noConsole -winFormsDPIAware -title "Impressao Automatica"
```

O programa detecta sozinho a pasta onde esta instalado (funciona em qualquer local/PC, desde que a estrutura de pastas `Entrada/Impresso/Erro/Log/Separador` esteja junto).

## Configuracao

No topo de cada script (ou nos controles da interface grafica): nome da impressora, numero de copias, frente e verso, folha separadora liga/desliga, e o intervalo de verificacao da pasta `Entrada`.
