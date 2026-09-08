# Impressao Automatica

Ferramenta para Windows que imprime arquivos automaticamente a partir de uma pasta monitorada, com interface grafica para fila manual (arrastar e soltar, escolher impressora, reordenar antes de imprimir).

Desenvolvido por [Jean Vieira](mailto:jean.vieira@hotmail.com), contador, pra eliminar a etapa manual de abrir cada arquivo e mandar imprimir um por um.

## Baixar (Windows, pronto pra usar)

Nao precisa instalar nada. Baixe o zip (contem o `.exe`, o `LEIA-ME.txt` com instrucoes e a `LICENSE.txt`) e extraia:

**[Impressao_Automatica.zip (GitHub Releases)](https://github.com/jeanvieir4/impressao-automatica/releases/download/v1.0.0/Impressao_Automatica.zip)**

O LEIA-ME.txt tambem esta neste repositorio em [`Para_Equipe/LEIA-ME.txt`](Para_Equipe/LEIA-ME.txt).

Na primeira execucao o Windows SmartScreen avisa "Editor desconhecido" - isso e normal pra um executavel sem certificado de assinatura paga, nao e virus. Clique em "Mais informacoes" -> "Executar assim mesmo". Se o antivirus colocar em quarentena, restaure e adicione uma excecao.

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
| `Impressao.exe` | Versao compilada de `2 - interface.ps1` (gerada com [ps2exe](https://github.com/MScholtes/PS2EXE)), disponivel na secao [Baixar](#baixar-windows-pronto-pra-usar), nao versionada no repositorio. |

## Requisitos

- Windows com PowerShell.
- **Adobe Acrobat** (Reader ou DC) instalado, para impressao silenciosa de PDF sem abrir dialogos. Sem ele, PDFs caem no verbo de impressao padrao do Windows, que pode nao funcionar silenciosamente dependendo do leitor de PDF padrao configurado.

## Rodar a partir do codigo-fonte

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

## Aviso de responsabilidade

Ferramenta oferecida como esta, gratuita, sem garantia. Confira sempre se os documentos importantes realmente saíram da impressora antes de descartar o original ou dar uma tarefa por concluida.

## Licenca

[MIT](LICENSE) - livre pra usar, copiar e repassar, inclusive em trabalho comercial.

## Apoie o projeto

Se este programa te ajudou, voce pode apoiar o desenvolvimento via Pix:

**Chave Pix:** jean.vieira@hotmail.com

![QR Code Pix](./assets/qrcode-pix.png)

## Contato

Duvidas, bug ou sugestao: Jean Vieira - (49) 99907-9884 - jean.vieira@hotmail.com
