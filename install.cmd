# Activation-Office-Project-and-Visio-2019
Script CMD
Para Windows 10 Pro >>> Key Hacking: Office, Project e Visio.
Recurso de Ativação definitiva do MS Office e complementos 
1 - Abra o menu executar - Windows + R ou acesse através da cortana
2 - Digite cmd (admin)
3 - Copie todo o código
4 - Cole

Nota: Conexão à Internet On

=================================
1. Microsoft Office Professional Plus 2019

if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
set "cmd=cscript //nologo ospp.vbs"
%cmd% /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul 2>&1
%cmd% /dstatus | findstr "Office19ProPlus2019VL"
if not %errorlevel% == 0 (for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do %cmd% /inslic:"..\root\Licenses16\%x")
%cmd% /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
%cmd% /sethst:kms.lotro.cc & %cmd% /act
cls & %cmd% /dstatus
echo

2. Microsoft Project Professional 2019

if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
set "cmd=cscript //nologo ospp.vbs"
%cmd% /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B >nul 2>&1
%cmd% /dstatus | findstr "Office19ProjectPro2019VL"
if not %errorlevel% == 0 (for /f %x in ('dir /b ..\root\Licenses16\ProjectPro2019VL*.xrm-ms') do %cmd% /inslic:"..\root\Licenses16\%x")
%cmd% /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
%cmd% /sethst:kms.lotro.cc & %cmd% /act
cls & %cmd% /dstatus
echo

3. MS Visio Professional 2019

if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
set "cmd=cscript //nologo ospp.vbs"
%cmd% /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB >nul 2>&1
%cmd% /dstatus | findstr "Office19VisioPro2019VL"
if not %errorlevel% == 0 (for /f %x in ('dir /b ..\root\Licenses16\VisioPro2019VL*.xrm-ms') do %cmd% /inslic:"..\root\Licenses16\%x")
%cmd% /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
%cmd% /sethst:kms.lotro.cc & %cmd% /act
cls & %cmd% /dstatus
echo
