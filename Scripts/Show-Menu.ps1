Function Show-Menu($Items, $Title){
    Function Write-Menu($Items, $Length, $Title){
        [System.Console]::SetCursorPosition(0, 0)
        Write-Host $Title -NoNewline -ForegroundColor Gray
        Write-Host "`n`n" -NoNewline
        $Items | ForEach-Object {
            If ($_.IsSelected){
                Write-Host ($_.DisplayText($Length)) -ForegroundColor White -BackgroundColor DarkCyan -NoNewline
            }
            Else{
                Write-Host ($_.DisplayText($Length)) -ForegroundColor White -NoNewline
            }
            Write-Host "`n" -NoNewline
        }
        If ($Items.Count -lt $ItemsCount){
            (1 .. ($ItemsCount - $items.Count)) | ForEach-Object{
                Write-Host "$(''.PadRight($Length))`n" -NoNewline
            }
        }
        Write-Host "`n" -NoNewline
        Write-Host "Esc" -NoNewline -BackgroundColor Gray -ForegroundColor Black
        Write-Host " 終了  " -NoNewline -ForegroundColor Gray
        Write-Host "↑・↓" -NoNewline -BackgroundColor Gray -ForegroundColor Black
        Write-Host " 項目の選択  " -NoNewline -ForegroundColor Gray
        Write-Host "PgUp・PgDn" -NoNewline -BackgroundColor Gray -ForegroundColor Black
        Write-Host " 選択項目の移動  " -NoNewline -ForegroundColor Gray
        Write-Host "Del" -NoNewline -BackgroundColor Gray -ForegroundColor Black
        Write-Host " 選択項目の削除  " -NoNewline -ForegroundColor Gray
    }

    If ($psISE -ne $Null){
        Write-Error -Exception ([System.NotSupportedException]"PowerShell ISE では動作しません")
        Break
    }

    $ItemsCount = $Items.Count
    $Length = ($Items | Select-Object @{Name="ByteLength";Expression={[System.Text.Encoding]::GetEncoding("Shift_JIS").GetByteCount(" $($_.Name) [$($_.DesktopApplicationID)]")}} | Sort-Object ByteLength -Descending)[0].ByteLength
    Clear-Host
    [System.Console]::CursorVisible = $False

    While ($True){
        Write-Menu -Items $Items -Length $Length -Title $Title
        $PressedKey = [System.Console]::ReadKey($True)
        If ($PressedKey.Key -eq [System.ConsoleKey]::Escape -and $PressedKey.Modifiers -eq 0){
            [System.Console]::CursorVisible = $True
            Clear-Host
            Break
        }
        ElseIf ($PressedKey.Key -eq [System.ConsoleKey]::Home -and $PressedKey.Modifiers -eq 0){
            $Index = $Items.IndexOf(($Items | Where-Object IsSelected))
            $Items[$Index].IsSelected = $False
            $Items[0].IsSelected = $True
        }
        ElseIf ($PressedKey.Key -eq [System.ConsoleKey]::End -and $PressedKey.Modifiers -eq 0){
            $Index = $Items.IndexOf(($Items | Where-Object IsSelected))
            $Items[$Index].IsSelected = $False
            $Items[$Items.Count - 1].IsSelected = $True
        }
        ElseIf ($PressedKey.Key -eq [System.ConsoleKey]::UpArrow -and $PressedKey.Modifiers -eq 0){
            $Index = $Items.IndexOf(($Items | Where-Object IsSelected))
            $Items[$Index].IsSelected = $False
            If ($Index -eq 0){
                $Items[$Items.Count - 1].IsSelected = $True
            }
            Else{
                $Items[$Index - 1].IsSelected = $True
            }
        }
        ElseIf ($PressedKey.Key -eq [System.ConsoleKey]::DownArrow -and $PressedKey.Modifiers -eq 0){
            $Index = $Items.IndexOf(($Items | Where-Object IsSelected))
            $Items[$Index].IsSelected = $False
            If ($Index -eq ($Items.Count - 1)){
                $Items[0].IsSelected = $True
            }
            Else{
                $Items[$Index + 1].IsSelected = $True
            }
        }
        ElseIf ($PressedKey.Key -eq [System.ConsoleKey]::PageUp -and $PressedKey.Modifiers -eq 0){
            $Item = ($Items | Where-Object IsSelected)
            $Index = $Items.IndexOf($Item)
            If ($Index -ne 0){
                $Items.RemoveAt($Index)
                $Items.Insert($Index - 1, $Item)
            }
        }
        ElseIf ($PressedKey.Key -eq [System.ConsoleKey]::PageDown -and $PressedKey.Modifiers -eq 0){
            $Item = ($Items | Where-Object IsSelected)
            $Index = $Items.IndexOf($Item)
            If ($Index -ne ($Items.Count - 1)){
                $Items.RemoveAt($Index)
                $Items.Insert($Index + 1, $Item)
            }
        }
        ElseIf ($PressedKey.Key -eq [System.ConsoleKey]::Delete -and $PressedKey.Modifiers -eq 0){
            $Index = $Items.IndexOf(($Items | Where-Object IsSelected))
            $Items.RemoveAt($Index)
            If ($Items.Count -eq 0){
                [System.Console]::CursorVisible = $True
                Clear-Host
                Break
            }
            Else{
                $Items[0].IsSelected = $True
            }
        }
        ElseIf ($PressedKey.Key -eq [System.ConsoleKey]::F5 -and $PressedKey.Modifiers -eq 0){
            Clear-Host
            [System.Console]::CursorVisible = $False
        }
    }
}
