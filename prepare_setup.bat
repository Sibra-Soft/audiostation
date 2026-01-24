@echo off

cd ./project

XCOPY ".\deps\*.ocx" ".\setup\deps" /s /i
XCOPY ".\deps\*.dll" ".\setup\deps" /s /i

XCOPY ".\build\*.exe" ".\setup\build" /s /i