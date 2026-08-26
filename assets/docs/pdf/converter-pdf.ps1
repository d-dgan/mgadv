# Converte para PDF os documentos gerados pelo site.
#
# Usa o proprio Word que ja esta instalado no computador, entao o PDF sai
# identico ao que aparece na tela — timbrado, logo, rodape e fontes. Nenhum
# arquivo sai do computador: nao ha envio para servidor nem para a internet.
#
# Pode receber os arquivos arrastados por cima do atalho. Sem nenhum arquivo,
# converte todos os .docx da pasta em que o script esta.

param([Parameter(ValueFromRemainingArguments = $true)] [string[]] $Arquivos)

$ErrorActionPreference = 'Stop'
$wdFormatPDF = 17

function Escrever($texto, $cor = 'Gray') { Write-Host $texto -ForegroundColor $cor }

# Sem argumentos: pega tudo o que for .docx na pasta do proprio script.
if (-not $Arquivos -or $Arquivos.Count -eq 0) {
    $pasta = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $Arquivos = Get-ChildItem -Path $pasta -Filter *.docx -File | ForEach-Object { $_.FullName }
}

# Arquivo temporario do Word (~$nome.docx) nao deve ser convertido.
$Arquivos = $Arquivos | Where-Object {
    $_ -and (Test-Path $_) -and ([IO.Path]::GetExtension($_) -eq '.docx') -and
    -not ([IO.Path]::GetFileName($_)).StartsWith('~$')
}

if ($Arquivos.Count -eq 0) {
    Escrever 'Nenhum documento do Word foi encontrado nesta pasta.' 'Yellow'
    Escrever ''
    # Causa mais comum: clicar duas vezes aqui dentro do proprio ZIP. O Windows
    # deixa abrir um arquivo sem extrair, mas copia so ele para uma pasta
    # temporaria — os documentos ficam para tras e nao ha o que converter.
    Escrever 'O motivo mais provavel e que o ZIP ainda nao foi extraido.' 'Yellow'
    Escrever ''
    Escrever 'Feche esta janela e faca assim:'
    Escrever '  1. Clique com o botao direito no arquivo .zip que voce baixou'
    Escrever '  2. Escolha "Extrair tudo" e confirme'
    Escrever '  3. Abra a pasta que apareceu e clique duas vezes aqui de novo'
    Escrever ''
    Read-Host 'Pressione Enter para fechar'
    exit 1
}

Escrever "Convertendo $($Arquivos.Count) documento(s) para PDF..." 'Cyan'
Escrever ''

$word = $null
$convertidos = 0
$falhas = @()

try {
    try {
        $word = New-Object -ComObject Word.Application
    } catch {
        Escrever 'Nao foi possivel abrir o Microsoft Word.' 'Red'
        Escrever 'Confira se o Word esta instalado neste computador.' 'Red'
        Read-Host 'Pressione Enter para fechar'
        exit 1
    }

    $word.Visible = $false
    $word.DisplayAlerts = 0          # wdAlertsNone: nao trava esperando caixa de dialogo

    foreach ($arquivo in $Arquivos) {
        $nome = [IO.Path]::GetFileName($arquivo)
        $pdf  = [IO.Path]::ChangeExtension($arquivo, '.pdf')
        $doc  = $null
        try {
            # ReadOnly, para nunca alterar o .docx original por acidente.
            $doc = $word.Documents.Open($arquivo, $false, $true)
            $doc.ExportAsFixedFormat($pdf, $wdFormatPDF)
            $convertidos++
            Escrever ("  OK   {0}  ->  {1}" -f $nome, [IO.Path]::GetFileName($pdf)) 'Green'
        } catch {
            $falhas += $nome
            Escrever ("  ERRO {0}: {1}" -f $nome, $_.Exception.Message) 'Red'
        } finally {
            if ($doc) { $doc.Close($false) }
        }
    }
} finally {
    # O Word fica rodando em segundo plano se nao for encerrado explicitamente.
    if ($word) { try { $word.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Escrever ''
if ($falhas.Count -eq 0) {
    Escrever "Pronto: $convertidos PDF(s) criados na mesma pasta dos documentos." 'Green'
} else {
    Escrever "$convertidos convertido(s), $($falhas.Count) com erro: $($falhas -join ', ')" 'Yellow'
}
Escrever ''
Read-Host 'Pressione Enter para fechar'
