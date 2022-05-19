@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio\VC98\Bin\cl.exe" /nologo /MD
set bin_dir=..\..
set bin_name=sha1.dll

pushd %~dp0

%cl_exe% sha1.c -link -dll -out:%bin_name%
copy %bin_name% %bin_dir% > nul
del /q *.exp *.lib *.obj *.dll ~$*
pause