just copy all commands individually and paste them into a cmd window that was started as administrator as soon as your office is ready.

64Bits: cd /d %ProgramFiles%\Microsoft Office\Office16       
32bits: cd /d %ProgramFiles(x86)%\Microsoft Office\Office16  

for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"

cscript ospp.vbs /setprt:1688

cscript ospp.vbs /unpkey:6F7TH >nul

cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH

cscript ospp.vbs /sethst:s8.uk.to

cscript ospp.vbs /act

presionar ENTER
----------------------------------------------------------
