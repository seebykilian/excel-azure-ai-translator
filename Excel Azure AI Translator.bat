@echo off

REM Change the directory to the directory of the batch file
cd /d %~dp0

REM Set the title of the command prompt window to "Excel Azure AI Translator"
title Excel Azure AI Translator

REM Run the command to start the application
dotnet run
