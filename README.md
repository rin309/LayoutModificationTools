# LayoutModificationTools
LayoutModification.xml の作成を支援します

# LayoutModificationTools のインストール
下記 URL を参照してください。
https://github.com/rin309/LayoutModificationTools/wiki/%E3%82%A4%E3%83%B3%E3%82%B9%E3%83%88%E3%83%BC%E3%83%AB%E3%80%81%E3%81%82%E3%82%8B%E3%81%84%E3%81%AF%E3%83%AD%E3%83%BC%E3%82%AB%E3%83%AB%E3%81%A7%E3%81%AE%E4%BD%BF%E7%94%A8

# 使い方1: TaskbarLayoutModification.xml を出力
```
Export-TaskbarPinnedAppsLayout -Path C:\Windows\OEM\TaskbarLayoutModification.xml -EditSortOrder
```
- クリーンインストール直後など、通常 C:\Windows\OEM\ は存在しません。あらかじめフォルダーを作成してから実行してください。
- C:\Windows フォルダーへは通常のアクセス権では書き込みできません。上記コマンドのように直接保存される場合は、PowerShell を管理者として実行してから実行してください。

