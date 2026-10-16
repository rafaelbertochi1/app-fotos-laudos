@echo off
REM Gera o programa (um arquivo só, com a pasta "modelos" embutida:
REM modelos\<banco>\*.xlsm). Para trocar/adicionar modelos sem gerar de novo:
REM crie modelos\<banco>\ ao lado do .exe e coloque os .xlsm lá (têm prioridade).
REM O .exe é gerado em build\exe\ (pasta de trabalho, ninguém deve usar o de lá),
REM publicado na dist\ como DownloaderFotos_v<versão>.exe pelo publicar.ps1 (apaga
REM a versão anterior; se estiver aberta, troca sozinho depois) e apagado de build\exe\.
REM Para publicar também na pasta pública: publicar.ps1 sem -SoDist.
cd /d "%~dp0"
python -m pip install -r requirements.txt --quiet
REM testes rapidos (segundos): regra quebrada nao vira .exe
python -m unittest discover -s testes
if errorlevel 1 (
  echo.
  echo TESTES FALHARAM: o .exe NAO foi gerado. Veja acima.
  goto fim
)
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name DownloaderFotos ^
  --icon assets\icone.ico ^
  --add-data "assets\icone.ico;assets" ^
  --add-data "assets\logo_empresa.png;assets" ^
  --add-data "assets\assinatura_responsavel.png;assets" ^
  --add-data "modelos;modelos" ^
  --collect-data sv_ttk ^
  --hidden-import win32timezone ^
  --exclude-module playwright ^
  --exclude-module openpyxl ^
  --distpath build\exe ^
  app.py
if errorlevel 1 goto fim
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0publicar.ps1" -SoDist -Novo "%~dp0build\exe\DownloaderFotos.exe"
del /q "%~dp0build\exe\DownloaderFotos.exe" 2>nul
echo.
echo Pronto: dist\DownloaderFotos_v^<versao^>.exe
:fim
