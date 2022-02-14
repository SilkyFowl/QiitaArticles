<!--
title:   Object Browser In PowerShellを見つけた話
tags:    PowerShell
id:      a39888b53bcdb73edab9
private: false
-->
## 初めに

Poswershellで少し凝ったことをしたくなった場合、**CUIに表示された文字列の実態はオブジェクトである**と認識することはとても重要です。

>PowerShell は、**.NET 上に構築された**タスクベースのコマンド ライン シェルおよびスクリプト言語です。 PowerShell は、システム管理者およびパワー ユーザーが、オペレーティング システム (Linux、macOS、および Windows) とプロセスを管理するタスクを迅速に自動化するのに役立ちます。
[PowerShell スクリプト - PowerShell | Microsoft Docs](https://docs.microsoft.com/ja-jp/powershell/scripting/overview?view=powershell-7)

最近、その助けになるであろうスクリプトが公開されていることを知りました。
Powershell版ObjectBrowserです。ここからダウンロードできます。
[PowerShell Object Browser](https://gallery.technet.microsoft.com/PowerShell-Object-Browser-847d62c1)

解説はこちら
[How to Create and Use an Object Browser In PowerShell](https://social.technet.microsoft.com/wiki/contents/articles/28651.how-to-create-and-use-an-object-browser-in-powershell.aspx)

## 使い方

#### スクリプトを開いたpowershellプロセスで使えるクラスを表示


![操作a](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/027c902e-0f3a-84fe-388f-31642390994f.jpeg)

![操作A.gif](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/5b481c77-a8c5-a14e-531d-ba39d9bf2ac4.gif)

色んなことがわかりますが、特にValuesタブ最終行の`UnderlyingSystemType`がとても助かります……

#### 任意のObjectを表示

![操作1.gif](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/81428a19-b351-15a3-a1da-43aeaa19421a.gif)
![操作2.gif](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/c04b144a-807e-4107-127c-7497f5bc5f83.gif)

調べたいObjectがある時はPowershellタブから追加できます。
`Add-Node [Object]$Object [String]$Name`
`Remove-Node [Object]$Node`

##### 応用:コマンド結果の比較
![操作i.gif](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/ec4fbb5a-0311-5b93-3b1b-4adbf796cf82.gif)
![操作ii.gif](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/82d70238-950a-6aa2-26dd-7d1e38f3dfe9.gif)
パイプライン処理の結果とコレクションメソッド処理の結果を比較しています。

## 何に使うの?

VScodeのデバッグ機能や、`gm -InputObject $Foo -Force | Out-GridView`だと、深い階層の情報を調べるのはなかなかメンドクサイです……
そんな時にお世話になっています。
![応用.jpg](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/9cb6585b-da8a-e525-646f-6ff4b3b6b558.jpeg)

たとえば、Nugetからダウンロードしてきたライブラリの使いかたを調べる助けになったりしてます。

## Powershell 7で使うためには

Powershell7で動かすには L5260以降の` Generate-PSObjectExplorerForm`の修正が必要です。
参考 : [Windows フォームに関する破壊的変更 - .NET Core | Microsoft Docs](https://docs.microsoft.com/ja-jp/dotnet/core/compatibility/winforms)
- `System.Windows.Forms.MainMenu` → ` System.Windows.Forms.MenuStrip`
- `System.Windows.Forms.MenuItem` → ` System.Windows.Forms.ToolStripMenuItem`
- `$MainMenu.MenuItems.Add` → `$MainMenu.Items.Add`
- `$mnuView.MenuItems.Add` → `$mnuView.DropDownItems.Add`
- L5821 `$Form.ShowDialog()| Out-Null` の前行に `$Form.Controls.Add($MainMenu)`を追加
  > // Add the MenuStrip last.
  > // This is important for correct placement in the z-order.
     [MenuStrip クラス (System.Windows.Forms) | Microsoft Docs](https://docs.microsoft.com/ja-jp/dotnet/api/system.windows.forms.menustrip?view=netcore-3.1) コード例より

## 余談

- 起動する際、`Start-ThreadJob`を使うといい感じでした。
- 二次配布、GitHubの使い方とかそこらへんが疎いので勉強中です。