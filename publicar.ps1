# Publica o programa na dist deste PC e na pasta publica, COM A VERSAO NO NOME:
#   DownloaderFotos_v<versao>.exe  (+ LEIA-ME.txt)
# (pedido do Rafael, 13/10/2026: so o arquivo com versao, sem DownloaderFotos.exe)
#
# - As versoes anteriores (DownloaderFotos_v*.exe e o antigo DownloaderFotos.exe)
#   sao apagadas. Se alguma estiver aberta (o Windows nao deixa apagar .exe em
#   uso), fica la ate fecharem.
# - Se a MESMA versao for gerada de novo com ela aberta, a nova espera em
#   _pendente\ e entra no lugar quando liberar.
# - O que ficar pendente uma tarefa agendada do Windows resolve de 5 em 5
#   minutos; quando nao sobra nada, ela se remove sozinha.
#
#   powershell -ExecutionPolicy Bypass -File publicar.ps1            (dist + pasta publica)
#   powershell -ExecutionPolicy Bypass -File publicar.ps1 -SoDist    (so a dist; usado pelo gerar_exe.bat)
#   ... -GitHub    tambem cria/atualiza a Release v<versao> no GitHub (repositorio deste projeto) com o
#                  .exe e o LEIA-ME pra baixar (precisa do gh logado; faca o push antes)
#   ... -SemTestes nao roda o teste completo (testes\regressao.py) antes de publicar
#
# Antes de ir pra pasta publica, o .exe novo passa pelo teste completo
# (testes\regressao.py --exe): propostas fixas de verdade comparadas com o
# gabarito. Falhou = nada e publicado (pedido de 15/10/2026).
#
# Log: %LOCALAPPDATA%\DownloaderFotos\logs\publicar.log
param(
    [string]$Novo,      # exe recem-gerado (padrao: build\exe\DownloaderFotos.exe)
    [switch]$SoDist,
    [switch]$GitHub,    # tambem cria/atualiza a Release no GitHub
    [switch]$SemTestes, # pula o teste completo antes da pasta publica
    [switch]$Agendado   # chamada pela tarefa agendada
)

$repo = $PSScriptRoot
$publica = "\\servidor\pasta_publica"
$tarefa = "DownloaderFotos - trocar versao nova"
$logDir = Join-Path $env:LOCALAPPDATA "DownloaderFotos\logs"
New-Item -ItemType Directory -Force $logDir | Out-Null
$logArq = Join-Path $logDir "publicar.log"
$versao = [regex]::Match((Get-Content (Join-Path $repo "comum.py") -Raw), 'VERSAO = "([^"]+)"').Groups[1].Value
$nome = "DownloaderFotos_v$versao.exe"

function Log($msg) {
    $linha = "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') $msg"
    Add-Content -Path $logArq -Value $linha -Encoding UTF8
    Write-Host $linha
}

function Hash($arq) {
    if (Test-Path $arq) { return (Get-FileHash $arq).Hash } else { return "" }
}

function Colocar($origem, $pasta) {
    # poe $origem como DownloaderFotos_v<versao>.exe; $true se ficou igual
    $alvo = Join-Path $pasta $nome
    if ((Hash $alvo) -eq (Hash $origem)) { return $true }
    try { Copy-Item $origem $alvo -Force -ErrorAction Stop } catch { return $false }
    return ((Hash $alvo) -eq (Hash $origem))
}

function ApagarAntigas($pasta) {
    # apaga DownloaderFotos*.exe que nao sao a versao atual; devolve quantos ficaram (abertos)
    $ficaram = 0
    foreach ($v in @(Get-ChildItem $pasta -Filter "DownloaderFotos*.exe" -ErrorAction SilentlyContinue)) {
        if ($v.Name -ieq $nome) { continue }
        try { Remove-Item $v.FullName -Force -ErrorAction Stop; Log "apagada a versao antiga $($v.FullName)" }
        catch { $ficaram++ }
    }
    return $ficaram
}

function Pendente($pasta) {
    $p = Join-Path (Join-Path $pasta "_pendente") $nome
    if (Test-Path $p) { return $p } else { return $null }
}

function Agendar {
    $script = $PSCommandPath
    # conhost --headless: roda sem piscar janela de console a cada 5 minutos
    $acao = New-ScheduledTaskAction -Execute "conhost.exe" `
        -Argument "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$script`" -Agendado"
    $gatilho = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
        -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days 30)
    $cfg = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $tarefa -Action $acao -Trigger $gatilho -Settings $cfg -Force `
        -Description "Termina a publicacao do DownloaderFotos quando ninguem estiver usando (publicar.ps1). Se remove sozinha." | Out-Null
    Log "tarefa agendada '$tarefa' (a cada 5 min)"
}

function Resolver($pasta) {
    # termina o que ficou pendente numa pasta; devolve quantas coisas ainda faltam
    $falta = 0
    $pend = Pendente $pasta
    if ($pend) {
        if (Colocar $pend $pasta) {
            Log "trocado: $pasta\$nome (estava em _pendente)"
            Remove-Item $pend -Force -ErrorAction SilentlyContinue
            Remove-Item (Join-Path $pasta "_pendente") -Force -ErrorAction SilentlyContinue  # so se vazia
        } else { $falta++ }
    }
    if (Test-Path (Join-Path $pasta $nome)) { $falta += (ApagarAntigas $pasta) }  # so apaga antigas com a atual no lugar
    $dirPend = Join-Path $pasta "_pendente"
    if ((Test-Path $dirPend) -and -not (Get-ChildItem $dirPend -Force)) {
        Remove-Item $dirPend -Force -ErrorAction SilentlyContinue  # sobra vazia
    }
    return $falta
}

# ------------------------------------------------------------ tarefa agendada
if ($Agendado) {
    $pendentes = 0
    foreach ($pasta in @("$repo\dist", $publica)) {
        try {
            if (-not (Test-Path $pasta)) { $pendentes++; continue }  # P: fora do ar: tenta depois
            $pendentes += (Resolver $pasta)
        } catch { $pendentes++ }
    }
    if ($pendentes -eq 0) {
        Unregister-ScheduledTask -TaskName $tarefa -Confirm:$false -ErrorAction SilentlyContinue
        Log "tudo publicado: tarefa '$tarefa' removida"
    }
    exit 0
}

# ------------------------------------------------------------ publicar
if (-not $Novo) { $Novo = Join-Path $repo "build\exe\DownloaderFotos.exe" }
if (-not (Test-Path $Novo)) { Log "ERRO: nao achei $Novo"; exit 1 }
# copia a origem pra um lugar temporario: ela pode ser um DownloaderFotos*.exe da
# propria dist, que ApagarAntigas apagaria antes de chegar na pasta publica
$tmpNovo = Join-Path $env:TEMP "publicar_novo"
New-Item -ItemType Directory -Force $tmpNovo | Out-Null
Copy-Item $Novo (Join-Path $tmpNovo "DownloaderFotos_novo.exe") -Force
$Novo = Join-Path $tmpNovo "DownloaderFotos_novo.exe"
if (-not $SoDist -and -not $SemTestes) {
    Log "teste completo do .exe novo (testes\regressao.py) antes de publicar..."
    python (Join-Path $repo "testes\regressao.py") --exe $Novo
    if ($LASTEXITCODE -ne 0) {
        Log "ERRO: o teste completo falhou - NADA foi publicado (veja acima; confira e, se a diferenca for esperada, rode testes\regressao.py --gravar)"
        exit 1
    }
    Log "teste completo OK"
}
$pastas = @("$repo\dist")
if (-not $SoDist) { $pastas += $publica }
$agendar = $false
foreach ($pasta in $pastas) {
    if (-not (Test-Path $pasta)) {
        if ($pasta -eq $publica) { Log "ERRO: pasta publica $publica fora do ar"; continue }
        New-Item -ItemType Directory -Force $pasta | Out-Null
    }
    if (Colocar $Novo $pasta) {
        Log "publicado: $pasta\$nome"
    } else {
        $dirPend = Join-Path $pasta "_pendente"
        New-Item -ItemType Directory -Force $dirPend | Out-Null
        Copy-Item $Novo (Join-Path $dirPend $nome) -Force
        Log "$pasta\$nome esta aberto: a versao nova espera em _pendente\ (troca agendada)"
        $agendar = $true
    }
    if ((Resolver $pasta) -gt 0) { $agendar = $true }
}
if (-not $SoDist -and (Test-Path $publica)) {
    Copy-Item (Join-Path $repo "dist\LEIA-ME.txt") (Join-Path $publica "LEIA-ME.txt") -Force
    Log "publicado: $publica\LEIA-ME.txt"
}
if ($agendar) { Agendar }

if ($GitHub) {
    # Release v<versao> com o .exe (nome com versao) e o LEIA-ME, no repositorio de onde o projeto foi clonado
    $repoGh = (gh repo view --json nameWithOwner -q .nameWithOwner 2>$null)
    if (-not $repoGh) { Log "ERRO: nao achei o repositorio no GitHub (gh logado?)"; exit 1 }
    $tag = "v$versao"
    $leia = Join-Path $repo "dist\LEIA-ME.txt"
    $tmpDir = Join-Path $env:TEMP "publicar_gh"
    New-Item -ItemType Directory -Force $tmpDir | Out-Null
    $exeGh = Join-Path $tmpDir $nome
    Copy-Item $Novo $exeGh -Force
    gh release view $tag --repo $repoGh 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) {
        gh release upload $tag $exeGh $leia --repo $repoGh --clobber 2>&1 | Out-Null
        $ok = ($LASTEXITCODE -eq 0)
        gh release delete-asset $tag "DownloaderFotos.exe" --repo $repoGh --yes 2>$null | Out-Null  # nome antigo, se houver
    } else {
        gh release create $tag $exeGh $leia --repo $repoGh --target main --title "Downloader de Fotos $versao" `
            --notes "Baixe o $nome e rode (Windows 10/11 + Excel instalado). Veja o LEIA-ME.txt." 2>&1 | Out-Null
        $ok = ($LASTEXITCODE -eq 0)
    }
    if ($ok) { Log "GitHub: Release $tag atualizada com $nome e o LEIA-ME" }
    else { Log "ERRO: nao consegui publicar a Release $tag no GitHub (gh logado? push feito?)" }
}
