@echo off

"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe" "C:\Source\Repos\outsystems_advancedexcel\Advanced_Excel\Source\NET\Advanced_Excel.sln" /nologo /verbosity:minimal /target:Rebuild /property:Configuration=Release