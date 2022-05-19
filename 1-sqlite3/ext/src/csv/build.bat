@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio\VC98\Bin\cl.exe" /nologo /MD
set bin_dir=..\..

pushd %~dp0

%cl_exe% csv.c -link -dll -out:csv.dll
copy csv.dll %bin_dir% > nul

:cleanup
del /q *.exp *.lib *.obj *.dll ~$*
popd
pause