@echo off
@REM Get Current folder path "%~dp0"
set strPath=%~dp0FolderSync.vbs
echo  App Path: %strPath%
@REM Create schedule :weekly monday 11：00
schtasks /create /tn "SyncFoldersTask" /ru system /tr %strPath% /sc weekly /d mon /st 11:00
echo Task have created successfully
pause 
