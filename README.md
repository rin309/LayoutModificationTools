# LayoutModificationTools
LayoutModification.xml の作成を支援します

# LayoutModificationTools のインストール
下記 URL を参照してください。
https://github.com/rin309/LayoutModificationTools/wiki/%E3%82%A4%E3%83%B3%E3%82%B9%E3%83%88%E3%83%BC%E3%83%AB%E3%80%81%E3%81%82%E3%82%8B%E3%81%84%E3%81%AF%E3%83%AD%E3%83%BC%E3%82%AB%E3%83%AB%E3%81%A7%E3%81%AE%E4%BD%BF%E7%94%A8

# TaskbarLayoutModification.xml を出力: 使い方1
タスク バーにピン留めされたアプリの一覧を出力します。
```
Export-TaskbarPinnedAppsLayout -Path C:\Windows\OEM\TaskbarLayoutModification.xml -EditSortOrder
```
- クリーンインストール直後など、通常 C:\Windows\OEM\ は存在しません。あらかじめフォルダーを作成してから実行してください。
- C:\Windows フォルダーへは通常のアクセス権では書き込みできません。上記コマンドのように直接保存される場合は、PowerShell を管理者として実行してから実行してください。

順番が再現されないのは仕様です。
https://user-images.githubusercontent.com/760251/182179704-b8a88716-6411-44ad-b0d8-2c9b06b26f54.mp4
代わりに、動画のように順番がバラバラでも並び替えができるよう、編集機能を有しています。

# TaskbarLayoutModification.xml を出力: 使い方2
テキストエディタで編集したほうが早いかもしれませんが、並び替えの順番は自分で定義することも可能です。
```
Export-TaskbarPinnedAppsLayout -Path C:\Windows\OEM\TaskbarLayoutModification.xml -SortOrder @("MSEdge", "Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge", "Microsoft.WindowsStore_8wekyb3d8bbwe!App", "Microsoft.Windows.Explorer", "Microsoft.Office.WINWORD.EXE.15", "Microsoft.Office.EXCEL.EXE.15", "Microsoft.Office.POWERPNT.EXE.15", "Microsoft.Office.MSACCESS.EXE.15", "Microsoft.Office.MSPUB.EXE.15", "Microsoft.Office.OUTLOOK.EXE.15", "Microsoft.VisualStudioCode")
```
