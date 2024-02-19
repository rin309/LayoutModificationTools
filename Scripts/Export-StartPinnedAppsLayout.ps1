#Requires -Version 5.0

<#
 .Synopsis
   Outputs the startmenu layout to a file

 .Description
   Export startmenu pinned icons to a file

 .Parameter ExportPath
   Specifies an absolute path to a layout file

 .Parameter EditSortOrder
   Edit the sort order with the simple editor

 .Parameter SortOrder
   Edit default sort order list

   [default list]
   - Microsoft Edge
   - Microsoft Word
   - Microsoft Excel
   - Microsoft PowerPoint
   - Microsoft Access
   - Microsoft Publisher
   - Microsoft Outlook

 .Example
   Export-StartPinnedAppsLayout -ExportPath $env:UserProfile\Desktop\LayoutModification.json -EditSortOrder

 .Example
   Export-StartPinnedAppsLayout -ExportPath $env:UserProfile\Desktop\LayoutModification.json -SortOrder @("MSEdge", "Microsoft.Office.WINWORD.EXE.15", "Microsoft.Office.EXCEL.EXE.15", "Microsoft.Office.POWERPNT.EXE.15", "Microsoft.Office.OUTLOOK.EXE.15", "Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookforWindows", "windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel", "Microsoft.VisualStudioCode")

#>
Function Export-StartPinnedAppsLayout{
    Param(
        [Parameter(Mandatory)][String]$ExportPath="",
        [Switch]$EditSortOrder = $False,
        [String[]]$SortOrder = $Null
    )
    Class Item{
        [String]$Name
        [String]$AppLinkPath
        [String]$AppLinkType
        [String]$Path
        [Bool]$IsSelected
        [String] DisplayText([Int]$Length){
            $Text = " $($This.Name) [$($This.AppLinkPath)]"
            $DisplayTextLength = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetByteCount($Text.PadRight($Length))
            If ($Length -eq 0){
                $DisplayTextLength = $DisplayTextLength
            }
            ElseIf($DisplayTextLength -ge $Length){
                $DisplayTextLength = $Length - ($DisplayTextLength - $Length)
            }
            Else{
                $DisplayTextLength = $Length
            }
            Return $Text.PadRight($DisplayTextLength + 1)
        }
    }

    If (((Get-CimInstance -Query 'Select * from Win32_OperatingSystem').Caption) -notlike "Microsoft Windows 11*"){
        throw (New-Object System.NotSupportedException)
    }

    If (-not (Test-Path (Split-Path $ExportPath -Parent) -PathType Container)){
        Write-Error ([System.IO.FileNotFoundException]::new("宛先ディレクトリが見つかりませんでした: [$ExportPath]")) -ErrorAction Stop
    }

    Add-Type -AssemblyName System.Runtime.InteropServices | Out-Null

    $Definition = '[DllImport("shlwapi.dll", BestFitMapping = false, CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = false, ThrowOnUnmappableChar = true)]
    public static extern int SHLoadIndirectString(string pszSource, System.Text.StringBuilder pszOutBuf, int cchOutBuf, IntPtr ppvReserved);'
    $ShellLightweightUtilityFunctions = Add-Type -MemberDefinition $Definition -Name "Win32SHLoadIndirectString" -PassThru
    [System.Text.StringBuilder]$Buffer = 1024

    $UnpinFromStartText = "スタートからピン留めを解除(&N)"
    # -51201: スタート メニューにピン留めする
    # -51394: スタートからピン留めを外す(&P)
    # -51609: スタートからピン留めを解除(&N)
    # -51606: スタート にピン留めする(&P)
    $Source = "@shell32.dll,-51609"
    If ($ShellLightweightUtilityFunctions::SHLoadIndirectString($Source, $Buffer ,$Buffer.Capacity, [System.IntPtr]::Zero) -eq 0){
        $UnpinFromStartText = $Buffer.ToString()
    }

    $TemporaryItems = New-Object "System.Collections.Generic.List[Item]"
    Write-Progress -Activity "スタート メニューにピン留めされたプログラムを調べています" -Status "Export-StartLayout" -PercentComplete 10
    Export-StartLayout -Path $ExportPath
    $OriginalStartLayout = Get-Content -Path $ExportPath -Encoding UTF8 | ConvertFrom-Json
    Write-Progress -Activity "スタート メニューにピン留めされたプログラムを調べています" -Status "shell:AppsFolder" -PercentComplete 30
    $PinnedItems = New-Object "System.Collections.Generic.List[Item]"
    ((New-Object -Com Shell.Application).NameSpace("shell:AppsFolder").Items()) | Select-Object Name, Path, @{Name="IsPinnedOnTaskbar";Expression={$UnpinFromStartText -in (($_.Verbs() | Select-Object Name).Name)}} | Where-Object "IsPinnedOnTaskbar" | ForEach-Object {$PinnedItems.Add((New-Object Item -Property @{Name = $_.Name; AppLinkPath = $_.Path}))}
    Write-Progress -Activity "スタート メニューにピン留めされたプログラムを調べています" -PercentComplete 70
    $OriginalStartLayout.pinnedList | ForEach-Object {
        $Value = $_
        $_ | Get-Member -MemberType NoteProperty | ForEach-Object {
            If ($_.Name -eq "desktopAppLink"){
                $AppName = ([System.Io.Path]::GetFileNameWithoutExtension($Value.desktopAppLink)).Replace("File Explorer","エクスプローラー")
                $TemporaryItems.Add((New-Object Item -Property @{Name = ($PinnedItems | Where-Object Name -eq $AppName).Name; AppLinkType = "desktopAppId"; AppLinkPath = ($PinnedItems | Where-Object Name -eq $AppName).AppLinkPath}))
            }
            Else{
                $TemporaryItems.Add((New-Object Item -Property @{Name = ($PinnedItems | Where-Object AppLinkPath -eq $Value.($_.Name)).Name; AppLinkType = $_.Name; AppLinkPath = $Value.($_.Name)}))
            }
        }
    }

    Write-Progress -Activity "スタート メニューにピン留めされたプログラムを調べています" -Completed

    If ($TemporaryItems.Count -eq 0){
        Write-Warning "スタート メニューにピン留めされたアプリが見つかりませんでした"
        Break
    }

    If ($SortOrder -ne $Null){
        $TemporaryItems = $TemporaryItems | Sort-Object {$Index = $SortOrder.IndexOf($_.AppLinkPath); If ($Index -eq -1){$Index = 2147483647}; Return $Index}
    }

    If ($EditSortOrder){
        $TemporaryItems[0].IsSelected = $True
        Show-Menu -Items $TemporaryItems -Title "項目の並び替え"
    }

    $Items = [PSCustomObject]@{} | Select-Object "pinnedList"
    $Items.pinnedList = @()

    $TemporaryItems | ForEach-Object {
        $Items.pinnedList += [PSCustomObject]@{($_.AppLinkType) = $_.AppLinkPath}
    }

    $Items | ConvertTo-Json -Compress | Out-File -Encoding utf8 -FilePath $ExportPath
}
Export-ModuleMember -Function Export-StartPinnedAppsLayout