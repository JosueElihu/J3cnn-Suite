@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio\VC98\Bin\cl.exe" /nologo /MD
set bin_dir=..\..

pushd %~dp0

%cl_exe% /LD root13.c /Feroot13.dll /link
copy root13.dll %bin_dir% > nul
:cleanup
del /q *.exp *.lib *.obj *.dll ~$*
popd
pause