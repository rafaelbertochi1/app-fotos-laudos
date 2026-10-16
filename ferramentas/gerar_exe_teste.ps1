# Gera os 2 .exe de TESTE do pre-preenchimento do formulario do banco
# (pedido do Rafael, 09/10/2026 - testar os dois e aprovar um):
#   DownloaderFotos_TESTE_Basico.exe   -> so o que a planilha do banco ja traz preenchido
#   DownloaderFotos_TESTE_Completo.exe -> tudo que a Central de Gestão tiver
# Cada um leva um preenchimento.txt embutido (comum.PREENCHIMENTO) e mostra
# "TESTE - preenchimento ..." na janela. Vao para dist\teste e para a pasta
# publica (\\servidor\pasta_publica\TESTE preenchimento Banco D).
#
#   powershell -ExecutionPolicy Bypass -File ferramentas\gerar_exe_teste.ps1 [-SemPublica]
# (usado so no teste das duas edicoes; a aprovada, "completo", ja e o padrao)
param([switch]$SemPublica)

$repo = Split-Path $PSScriptRoot -Parent  # este script fica em ferramentas\
Set-Location $repo
$publica = "\\servidor\pasta_publica\TESTE preenchimento Banco D"
$saida = Join-Path $repo "dist\teste"
New-Item -ItemType Directory -Force $saida | Out-Null

foreach ($edicao in @("Basico", "Completo")) {
    $tmp = Join-Path $repo "build\edicao_$edicao"
    New-Item -ItemType Directory -Force $tmp | Out-Null
    Set-Content -Path (Join-Path $tmp "preenchimento.txt") -Value $edicao.ToLower() -Encoding ASCII
    $nome = "DownloaderFotos_TESTE_$edicao"
    Write-Host "== gerando $nome.exe"
    # caminhos absolutos: com --specpath o PyInstaller resolve os relativos a partir da pasta do spec
    python -m PyInstaller --noconfirm --clean --onefile --windowed --name $nome `
        --icon "$repo\assets\icone.ico" `
        --add-data "$repo\assets\icone.ico;assets" `
        --add-data "$repo\assets\logo_empresa.png;assets" `
        --add-data "$repo\assets\assinatura_responsavel.png;assets" `
        --add-data "$repo\modelos;modelos" `
        --add-data "$tmp\preenchimento.txt;." `
        --collect-data sv_ttk --hidden-import win32timezone `
        --exclude-module playwright --exclude-module openpyxl `
        --distpath "$repo\dist_build" --workpath "$repo\build\$nome" --specpath "$repo\build\$nome" `
        "$repo\app.py" 2>&1 | Select-Object -Last 1
    if (-not (Test-Path "dist_build\$nome.exe")) { Write-Host "ERRO: nao gerou $nome.exe"; exit 1 }
    Copy-Item "dist_build\$nome.exe" (Join-Path $saida "$nome.exe") -Force
    if (-not $SemPublica) {
        New-Item -ItemType Directory -Force $publica | Out-Null
        try {
            Copy-Item "dist_build\$nome.exe" (Join-Path $publica "$nome.exe") -Force -ErrorAction Stop
            Write-Host "publicado: $publica\$nome.exe"
        } catch {
            Write-Host "AVISO: $publica\$nome.exe esta aberto - feche e rode de novo"
        }
    }
}
Write-Host "Pronto: $saida"
