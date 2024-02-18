@echo off
cls

Title LayoutModificationTools
PowerShell -ExecutionPolicy ByPass -NoExit -Command "pushd '%~dp0'; Import-Module -Name ((Get-Location).Path); Write-Host \""[How to use]`nExport-TaskbarPinnedAppsLayout -ExportPath C:\Windows\OEM\TaskbarLayoutModification.xml -EditSortOrder`nExport-StartPinnedAppsLayout -ExportPath $env:UserProfile\Desktop\LayoutModification.json -EditSortOrder`n\"""
