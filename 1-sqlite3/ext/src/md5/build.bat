@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio\VC98\Bin\cl.exe" /nologo /MD
set bin_dir=..\..

pushd %~dp0
%cl_exe% /LD md5.c /Femd5.dll /link /DEF:md5.def
copy md5.dll %bin_dir% > nul
:cleanup
del /q *.exp *.lib *.obj *.dll ~$*
popd
pause