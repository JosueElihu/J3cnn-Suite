@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio\VC98\Bin\cl.exe" /nologo /MD
set bin_dir=..\..
set bin_name=cksumvfs.dll

pushd %~dp0

%cl_exe% cksumvfs.c -link -dll -out:%bin_name%
copy %bin_name% %bin_dir% > nul
del /q *.exp *.lib *.obj *.dll ~$*
pause