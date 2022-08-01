using namespace System.Xml
using namespace System.Xml.Linq

#Requires -Version 5.0

<#
 .Synopsis
   Outputs the taskbar layout to a file

 .Description
   Export taskbar pinned icons to a file

 .Parameter Path
   Specifies an absolute path to a layout file

 .Parameter EditSortOrder
   Edit the sort order with the simple editor

 .Example
   Export-TaskbarPinnedAppsLayout -Path C:\Windows\OEM\TaskbarLayoutModification.xml -EditSortOrder

#>
Function Export-TaskbarPinnedAppsLayout{
    Param(
        [Parameter(Mandatory)][String]$Path="",
        [Switch]$EditSortOrder = $False
    )
    Class Item{
        [String]$Name
        [String]$DesktopApplicationID
        [String]$Path
        [Bool]$IsSelected
        [String] DisplayText([Int]$Length){
            $Text = " $($This.Name) [$($This.DesktopApplicationID)]"
            Write-Verbose $Text
            $DisplayTextLength = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetByteCount($Text.PadRight($Length))
            If ($DisplayTextLength -ge $Length){
                $DisplayTextLength = $Length - ($DisplayTextLength - $Length)
            }
            Else{
                $DisplayTextLength = $Length
            }
            Return $Text.PadRight($DisplayTextLength + 1)
        }
    }

    If (-not (Test-Path (Split-Path $Path -Parent) -PathType Container)){
        Write-Error ([System.IO.FileNotFoundException]::new("宛先ディレクトリが見つかりませんでした: [$Path]")) -ErrorAction Stop
    }

    Add-Type -AssemblyName System.Runtime.InteropServices,System.Xml,System.Xml.Linq | Out-Null

    $Definition = '[DllImport("shlwapi.dll", BestFitMapping = false, CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = false, ThrowOnUnmappableChar = true)]
    public static extern int SHLoadIndirectString(string pszSource, System.Text.StringBuilder pszOutBuf, int cchOutBuf, IntPtr ppvReserved);'
    $ShellLightweightUtilityFunctions = Add-Type -MemberDefinition $Definition -Name "Win32SHLoadIndirectString" -PassThru
    [System.Text.StringBuilder]$Buffer = 1024

    $UnpinFromTaskbarText = "タスク バーからピン留めを外す(&K)"
    # -5386: タスク バーにピン留めする(&K)
    # -5387: タスク バーからピン留めを外す(&K)
    $Source = "@shell32.dll,-5387"
    If ($ShellLightweightUtilityFunctions::SHLoadIndirectString($Source, $Buffer ,$Buffer.Capacity, [System.IntPtr]::Zero) -eq 0){
        $UnpinFromTaskbarText = $Buffer.ToString()
    }

    $TemporaryItems = New-Object "System.Collections.Generic.List[Item]"
    Write-Progress -Activity "タスクバーにピン留されたプログラムを調べています" -Status "shell:User Pinned\Taskbar" -PercentComplete 10
    ((New-Object -Com Shell.Application).NameSpace("shell:User Pinned\Taskbar").Items() | Where-Object IsLink) | Select-Object Name, @{Name="DesktopApplicationID";Expression={@((New-Object -Com Shell.Application).NameSpace("shell:AppsFolder").Items() | Where-Object Name -eq $_.Name)[0].Path}}, Path, @{Name="IsPinnedOnTaskbar";Expression={$True}}, @{Name="DesktopApp";Expression={"DesktopApplicationLinkPath"}} | ForEach-Object {$TemporaryItems.Add((New-Object Item -Property @{Name = $_.Name; Path = $_.Path; DesktopApplicationID = $_.DesktopApplicationID}))}
    Write-Progress -Activity "タスクバーにピン留されたプログラムを調べています" -Status "shell:User Pinned\Taskbar" -PercentComplete 30
    ((New-Object -Com Shell.Application).NameSpace("shell:AppsFolder").Items()) | Select-Object Name, Path, @{Name="IsPinnedOnTaskbar";Expression={$UnpinFromTaskbarText -in (($_.Verbs() | Select-Object Name).Name)}}, @{Name="DesktopApp";Expression={"DesktopApplicationID"}} | Where-Object "IsPinnedOnTaskbar" | ForEach-Object {$TemporaryItems.Add((New-Object Item -Property @{Name = $_.Name; DesktopApplicationID = $_.Path}))}
    Write-Progress -Activity "タスクバーにピン留されたプログラムを調べています" -Completed

    $TemporaryItems = $TemporaryItems | Where-Object {-not ([String]::IsNullOrEmpty($_.DesktopApplicationID))} | Sort-Object -Unique DesktopApplicationID
    $Items = New-Object "System.Collections.Generic.List[Item]"
    $TemporaryItems | ForEach-Object {$Items.Add($_)}

    If ($EditSortOrder){
        Show-Menu -Items $Items -Title "項目の並び替え"
    }

    $Xml = [XDocument]::new(
        [XDeclaration]::new("1.0", "utf-8", "yes"),
        [XElement]::new("{http://schemas.microsoft.com/Start/2014/LayoutModification}LayoutModificationTemplate", @(
            #[XNamespace]"http://schemas.microsoft.com/Start/2014/LayoutModification",
            [XAttribute]::new("{http://www.w3.org/2000/xmlns/}defaultlayout", "http://schemas.microsoft.com/Start/2014/FullDefaultLayout"),
            [XAttribute]::new("{http://www.w3.org/2000/xmlns/}start", "http://schemas.microsoft.com/Start/2014/StartLayout"),
            [XAttribute]::new("{http://www.w3.org/2000/xmlns/}taskbar", "http://schemas.microsoft.com/Start/2014/TaskbarLayout"),
            [XAttribute]::new("version", "1"),
            [XElement]::new("{http://schemas.microsoft.com/Start/2014/LayoutModification}CustomTaskbarLayoutCollection", @(
                [XElement]::new("{http://schemas.microsoft.com/Start/2014/FullDefaultLayout}TaskbarLayout", @(
                    [XElement]::new("{http://schemas.microsoft.com/Start/2014/TaskbarLayout}TaskbarPinList", @(
                        #[XElement]::new("{http://schemas.microsoft.com/Start/2014/TaskbarLayout}DesktopApp", @(
                            #[XAttribute]::new("DesktopApplicationID", "Microsoft.Windows.Explorer")
                        #)),
                        #[XElement]::new("{http://schemas.microsoft.com/Start/2014/TaskbarLayout}DesktopApp", @(
                            #[XAttribute]::new("DesktopApplicationLinkPath", "%APPDATA%\Microsoft\Windows\Start Menu\Programs\System Tools\Command Prompt.lnk")
                        #))
                    ))
                ))
            ))
        ))
    )

    $Items | ForEach-Object {
        $Item = [XElement]::new("{http://schemas.microsoft.com/Start/2014/TaskbarLayout}DesktopApp", @(
                                [XAttribute]::new("DesktopApplicationID", $_.DesktopApplicationID)
                            ))
        $Xml.Root.Element("{http://schemas.microsoft.com/Start/2014/LayoutModification}CustomTaskbarLayoutCollection").Element("{http://schemas.microsoft.com/Start/2014/FullDefaultLayout}TaskbarLayout").Element("{http://schemas.microsoft.com/Start/2014/TaskbarLayout}TaskbarPinList").Add($Item)
    }

    $XmlWriterSettings = [XmlWriterSettings] @{
        Encoding = [System.Text.Encoding]::UTF8
        Indent = $True
        #NewLineOnAttributes = $True
    }
    $Writer = [XmlTextWriter]::Create($Path, $XmlWriterSettings)
    $Xml.Save($Writer)
    $Writer.Dispose()
}
Export-ModuleMember -Function Export-TaskbarPinnedAppsLayout