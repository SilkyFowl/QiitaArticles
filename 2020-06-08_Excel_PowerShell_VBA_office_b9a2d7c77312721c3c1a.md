<!--
title:   OfficeをCOM Object経由でPowershellから扱うときの面倒を少しマシにする
tags:    Excel,PowerShell,VBA,office
id:      b9a2d7c77312721c3c1a
private: false
-->
# ~~※最適解はCOMを使わないことです~~正しく扱えば問題ありませんでした

2020/10/03追記

**適切に解放すれば面倒くさくありませんでした。**
[OfficeをCOM Object経由でPowershellから扱うのはそれほど面倒じゃありませんでした](2020-10-04_Excel_PowerShell_VBA_office_b4b6271619bd6d3824f7.md)

~~Excelなら[ImportExcel](https://github.com/dfinke/ImportExcel)を使いましょう。
Wordなら[PSWriteWord](https://evotec.xyz/hub/scripts/pswriteword-powershell-module)がすごそうです。
幸い、PowershellはCOMを使わずとも大体何とかなります。
モジュール機能はFirefoxやChromeの拡張機能のようにPowershellに強力な機能を付与します。
ただし、それでもCOMを使わないといけない時は往々にしてあります。
**そもそも拡張機能はソフトウェアのインストールと判断されて導入が難しい場合も少なくないと思います。**
そんなわけで、COMを扱う作法を守りつつも面倒くさみを軽減したいという記事です。~~

## ~~結論（コードはこちら）~~

~~考え方は単純です。
コード内で生成されたCOMオブジェクトの情報をあらかじめ専用のスタックに格納して、最後にまとめて後始末するというものです。~~

```powershell:lazyPracticeExcelComObject.ps1

using namespace System.Management.Automation
using namespace System.Collections.Generic
using namespace System.Runtime.InteropServices
using namespace Microsoft.Office.Interop.Excel

param(
    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Path to one locations.")]
    [ValidateScript( { Test-Path $_  -Include "*.xlsx" })]
    [string]
    $ExcelPath
)

# VBAのEnumを使うためにアセンブリをロード
# Powershell 7.1ではアセンブリ名で見つけることが出来なかったのでPathで指定
@(
    "C:\Windows\assembly\GAC_MSIL\Microsoft.VisualBasic\*\Microsoft.VisualBasic.dll"
    "C:\Windows\assembly\GAC_MSIL\office\*\OFFICE.DLL"
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Vbe.Interop\*\Microsoft.Vbe.Interop.dll"
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\*\Microsoft.Office.Interop.Excel.dll"
).ForEach{
    Add-Type -path $_
}

# COMObjectに拡張メソッドを追加
Update-TypeData -TypeName System.__ComObject -MemberType ScriptMethod -MemberName Tee -Value {
    param(
        [Stack[WeakReference]]
        $stack
    )
    $stack.Push([WeakReference]::new($this))

    Write-Output $this
}

# Comを解放するために使う弱参照のスタック
$refs = [Stack[WeakReference]]::new()

# Excel操作本体
$app = (New-Object -ComObject Excel.Application).Tee($refs)
$book = $app.Workbooks.Tee($refs).Open($ExcelPath).Tee($refs)
$book.Worksheets.Tee($refs).Item('Sheet1').Tee($refs).Range('A1:B2').Tee($refs)._NewEnum.Tee($refs).ForEach{
    $_.Text
    [void][Marshal]::ReleaseComObject($_)
}

# (Get-CimInstance -Class Win32_Process -Filter "ProcessId = $((ps excel).id)" -Property "CommandLine").CommandLine

# ファイルを閉じる
$book.Close()

# COMの開放
While ($refs.Count) {
    # スタックから弱参照を取得
    $comRef = $refs.Pop()

    # 解放するCOMを参照してる変数を全て取得
    $comVar = (Get-Variable).where{ [object]::ReferenceEquals($comRef.Target, $_.Value) }

    # Applicationオブジェクトであるかの判定
    $isApp = $comRef.Target -is [Microsoft.Office.Interop.Excel.Application]

    # アプリケーションの終了前にガベージ コレクトを強制
    if ($isApp) {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        $comRef.Target.Quit()
    }

    # COMObjectの解放
    # ※正しく動作していればApplicationオブジェクトが解放された瞬間にExcelが終了する
    while ([Marshal]::ReleaseComObject($comRef.Target)) { }
    $comRef.Target = $null

    # 変数を削除
    $comVar | Remove-Variable
    Remove-Variable comRef

    # Application オブジェクトのガベージ コレクトを強制
    if ($isApp) {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }
}

```

## ~~そもそも、なんで面倒なの？~~

~~COMオブジェクトは他のプロセスにあるオブジェクトを参照したものです。
異なるプロセス間の通信を簡易化するための仕組みです。
イメージとしてはExcelプロセスの頭脳に**COMオブジェクトの参照カウンター数だけ**アンテナを突き刺してリモコンで操作しているようなものでしょうか？~~

>pwsh-chan「このファイルを開いて？」
>EXCEL「あっ あっ（`$workbook.Open("ファイルパス")`）」
>pwsh-chan「Sheets1のA1の値は？」
>EXCEL「あっ」
>EXCEL「`[Workbook]`インスタンスの あっ」
>EXCEL「`[Worksheets]`の子要素である`[Worksheet]`の`[Cells]`を取得して」
>EXCEL「あぅ」
>EXCEL「`Sheets1 という名前のシートが見つかりませんでした` あっ」
>pwsh-chan「......Sheet1のA1の値は？」
>EXCEL「あっ あっ あっ」

~~この仕組みのおかげでプロセス間通信の面倒を軽減できます。ただし、同時に様々な面倒の発生源になります。~~

```powershell:通常のExcel
PS D:\> Start-Process excel
PS D:\> (Get-CimInstance -Class Win32_Process -Filter "ProcessId = $((ps excel).id)" -Property "CommandLine").CommandLine
"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
PS D:\> ps excel | kill
```

~~通常の場合、プロセスのコマンドラインはExcelの実行ファイルのパスだけです。それに対し、COMオブジェクトでExcelを操作する場合は` /automation -Embedding`というオプションが付いています。（すでに実行しているプロセスからCOMオブジェクトのインスタンスを生成する場合については割愛します）~~

```powershell:操られてるExcel
# 冒頭のスクリプトのコメントアウトを解除する
(Get-CimInstance -Class Win32_Process -Filter "ProcessId = $((ps excel).id)" -Property "CommandLine").CommandLine
# "C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE" /automation -Embedding
```

~~COMオブジェクトの適切な後始末とは、操作対象のプロセスとのつながりを全て切断する（解放する）ことだとイメージするとわかりやすいかもしれません。
適切に後始末すると操作対象のプロセスは自動で終了します。
この辺の.netでCOMを扱うベストプラクティスについてはこちらの記事が大変参考になります。~~

[.NETを使ったOfficeの自動化が面倒なはずがない―そう考えていた時期が俺にもありました。](https://qiita.com/mima_ita/items/aa811423d8c4410eca71)

ベストプラクティスの要点は以下の2点です。

- 生成されたCOMオブジェクトは全て明示的に開放すること。
- `Application`クラスを解剖する前後でガベージコレクトを強制すること。

~~これを守るコードを普通に書くと、処理中に扱うCOMオブジェクトの数だけコードが膨れてしまいます。
とても面倒なのでなんとかしましょう。~~

### 余談:PowershellでCOMObjectのForeachを行う方法

今回、PowershellにおけるCOMの扱われ方を調査した際、[COM.Basic.Tests.ps1](https://github.com/daxian-dbw/PowerShell/blob/097a161f491f81f54c728bd871f999e3e8261a27/test/powershell/engine/COM/COM.Basic.Tests.ps1#L28-L37)で興味深いコードを見つけました。

```powershell:COM.Basic.Tests.ps1
        It "Should enumerate IEnumVariant interface object without exception" {
            $shell = New-Object -ComObject "Shell.Application"
            $folder = $shell.Namespace("$TESTDRIVE")
            $items = $folder.Items()

            ## $enumVariant is an IEnumVariant interface of all items belong to the folder, and it should be enumerated.
            $enumVariant = $items._NewEnum()
            $items.Count | Should -Be 3
            $enumVariant | Measure-Object | ForEach-Object Count | Should -Be $items.Count
        }
```

PowershellでCOMオブジェクトのForeachを行いたい場合は`_NewEnum`を取得すると良いみたいです。

## ~~メンドクサイに立ち向かう~~

~~ベストプラクティスの「メンドクサイ」を分解してみましょう。~~

- 明示的に開放するためにCOMオブジェクトを変数で保持しなければならないためメソッドチェインが使えない。
- 生成されるCOMオブジェクトの数だけ冗長な開放処理コードを書く必要がある。

~~つまり、こう出来ればいいわけです。~~

- メソッドチェインの過程で生成されるCOMオブジェクトを適切に解放出来る仕組みを作る
- 解放順序は階層が深い順で`Application`クラスを解放する前後でガベージコレクトを強制する

~~これを実現するためにアレコレ行った結果、冒頭のスクリプトが出来上がりました。コードが増えましたが生成されるCOMが増えてもこれ以上に解放処理コードを増やす必要がなくなりました。
面倒は多少軽減出来た気がします。~~

## 課題

- ~~`Tee()`メソッドの後はインテリセンスが効かない（`OutputTypeAttribute`を使ったけど効かなかった）~~

## ~~この先書きたいこと~~

- 今回のコードの詳細な解説
- スクリプトをモジュール化する

~~Powershellはなかなかに癖が強いとは思いますが、慣れると楽しいです~~