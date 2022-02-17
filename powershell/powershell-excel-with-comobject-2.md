<!--
title:   OfficeをCOM Object経由でPowershellから扱うのはそれほど面倒じゃありませんでした
tags:    Excel,PowerShell,VBA,office
id:      b4b6271619bd6d3824f7
private: false
-->
# これは何？

こんな記事を書きましたが、実はPowershellでCOMオブジェクト(RCW)を扱うのは面倒じゃありませんでしたという記事です。
[OfficeをCOM Object経由でPowershellから扱うときの面倒を少しマシにする](2020-06-08_Excel_PowerShell_VBA_office_b9a2d7c77312721c3c1a.md)

# COMオブジェクト(RCW)を扱うコツ

参考記事:[Answer: Clean up Excel Interop Objects with IDisposable](https://stackoverflow.com/a/25135685?stw=2 )

## 要約（DeepL）

- 「2ドットルール」を守ってオブジェクトの参照を保存して`Marshal.ReleaseComObject()`を呼び出す必要はない
- `GC.Collect()`、`GC.WaitForPendingFinalizers()`を呼び出して明示的にガベージコレクタを実行すれば解放される
- だたし、**どこにコードを書くのかが非常に重要**。よくある記事の場合、Excel関連の操作を小さなヘルパーメソッドに移動するとうまくいく

# やってみた

EnumRcw.exe…COMの解放漏れをチェックするツールをお借りしました。
> [.NETを使った別プロセスのOfficeの自動化が面倒なはずがない―そう考えていた時期が俺にもありました。](https://qiita.com/mima_ita/items/aa811423d8c4410eca71)
>
> [RCWのオブジェクトの状況を監視する。](https://qiita.com/mima_ita/items/aa811423d8c4410eca71#rcw%E3%81%AE%E3%82%AA%E3%83%96%E3%82%B8%E3%82%A7%E3%82%AF%E3%83%88%E3%81%AE%E7%8A%B6%E6%B3%81%E3%82%92%E7%9B%A3%E8%A6%96%E3%81%99%E3%82%8B)

## pwsh 7.1 rc-1の場合

検証をして気付いたのですが、正しく参照が解放されているとExcelプロセスは`Quit()`をしなくても終了するようです。
なお、Excelを可視化して通常通りの操作を行うと参照が全て解放されていてもプロセスは終了せず、手動でExcelを閉じると終了しました。
これはこちらの記事のようにインスタンスが再利用されたと解釈しています。

>[Office オートメーションで割り当てたオブジェクトを解放する - Part2](https://social.msdn.microsoft.com/Forums/ja-JP/0d9c6273-bade-4f6a-a0de-5adb748d15eb/office-part2)
>前回の投稿 Office のプロセス インスタンス制御について でも一部ご紹介しましたが、Office アプリケーションは、オブジェクト生成等を実施した際に、ワークステーション単位で管理されているランニング オブジェクト テーブル (ROT) 等に登録していて、様々な状況下で OLE や DDE 等のテクノロジーを基盤としてオブジェクトを再利用する動作となります。

もう1点気になったExcel操作以降の表示順序です。(後述するWindows Powershell5.1から変化している)
最新のPowershellはCOMオブジェクトの扱い方が改良されてます。その辺が関連しているのかもしれません。

```powershell:pwsh7.1rc-1
$EnumRcw = '~\Repos\MemoryCheck\bin\x64\Release\EnumRcw.exe'

'実行前'
& $EnumRcw pwsh

'Excel操作'
& {
  $app = New-Object -ComObject Excel.Application
  $app.Workbooks.Open((Convert-Path Temp:\Events.xlsx)).Sheets["Sheet1"].Range("A1:C7").foreach{
    $_ | select NumberFormat,Value2
  }

  $app.Workbooks.Close()
}

'操作おわり'

'この時点ではExcelプロセスは残ったまま'
& $EnumRcw pwsh

'GCを強制'
[gc]::Collect() #
[gc]::WaitForPendingFinalizers() #

& $EnumRcw pwsh
```

<details><summary>結果</summary><div>

```shell-session:結果
実行前
Excel操作

pwsh
pwsh 1852 =======================================================
Amd64
     24EE7923360           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE7923878           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE7923920           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE7928F80           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE7929030           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE7929188           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE7929230           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE7929370           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE7929420           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE7929560           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE7929610           24 System.__ComObject 2 False System.Management.Automation.ComInterop.IDispatch,System.Management.Automation.IDispatch,System.Runtime.InteropServices.IDispatch,
     24EE7929888           24 System.__ComObject 44 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE794ECB8           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.IEnumVARIANT,
     24EE794EE50           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE796D448           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE798B680           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE79A98C8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE79C7B00           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE79E5DB0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7A04000           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7A22238           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7A40490           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7A5E778           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7A7C9B0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7A9ABF8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7AB8E48           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7AD7098           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7AF52F0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7B13540           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7B31778           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7B4FAD8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7B6DD28           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7B8BF60           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7BAA1B8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7BE78C8           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
NumberFormat Value2
------------ ------
G/標準         MemberType
G/標準         Name
G/標準         DeclaringType
G/標準         Event
G/標準         UnhandledException
G/標準         System.AppContext
G/標準         Event
G/標準         FirstChanceException
G/標準         System.AppContext
G/標準         Event
G/標準         ProcessExit
G/標準         System.AppContext
G/標準         Event
G/標準         UnhandledException
G/標準         System.AppDomain
G/標準         Event
G/標準         DomainUnload
G/標準         System.AppDomain
G/標準         Event
G/標準         FirstChanceException
G/標準         System.AppDomain
操作おわり
この時点ではExcelプロセスは残ったまま
GCを強制
pwsh
pwsh 1852 =======================================================
Amd64
```

</div></details>

## Windows Powershell5.1の場合

Windows10標準機能なPowershellはこちらです。
最後に残る2つのオブジェクトはベストプラクティスを守っても残るようです。
このコードを実行したあとも、Excelプロセスは`Quit()`をしていないのに終了するので正しく参照が解放されていると考えて良いでしょう。

>[実験９：PowerShellで実行した場合](https://qiita.com/mima_ita/items/aa811423d8c4410eca71#%E5%AE%9F%E9%A8%93%EF%BC%99powershell%E3%81%A7%E5%AE%9F%E8%A1%8C%E3%81%97%E3%81%9F%E5%A0%B4%E5%90%88)
>いわゆるベストプラクティスを守ったスクリプトを実行しても、操作終了後に「System.Runtime.InteropServices.ComTypes.ITypeInfo」が残ります。このオブジェクトはSystem.Collections.Concurrent.ConcurrentDictionaryから参照されています。

```powershell
$excelPath = gci D:\temp\Events.xlsx

$EnumRcw = '~\Repos\MemoryCheck\bin\x64\Release\EnumRcw.exe'

'実行前'
& $EnumRcw powershell

'Excel操作'
& {
  $app = New-Object -ComObject Excel.Application
  $app.Workbooks.Open($excelPath).Sheets["Sheet1"].Range("A1:C7") | select NumberFormat,Value2

  $app.Workbooks.Close()
}
'操作おわり'

'この時点では参照が残ったまま'
& $EnumRcw powershell

'GCを強制'
[gc]::Collect() #
[gc]::WaitForPendingFinalizers() #

& $EnumRcw powershell
```

<details><summary>結果</summary><div>

```shell-session:結果
実行前
powershell
powershell 14216 =======================================================
Amd64
     1C9018E7970           32 System.__ComObject 1 False Windows.Foundation.Diagnostics.IAsyncCausalityTracerStatics,
     1C9018E79D0           32 Windows.Foundation.Diagnostics.TracingStatusChangedEventArgs 1 False Windows.Foundation.Diagnostics.ITracingStatusChangedEventArgs,
     1C901B99008           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1C901B99028           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1C901B994C8           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1C901B994E8           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1C901B99508           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1C901E3E158           32 System.__ComObject (GetRCWDataに失敗)
     1C901E96FA0           32 System.__ComObject (GetRCWDataに失敗)
Excel操作

NumberFormat Value2       
------------ ------
General      MemberType
General      Name
General      DeclaringType
General      Event        
General      UnhandledE...
General      System.App...
General      Event        
General      FirstChanc...
General      System.App...
General      Event        
General      ProcessExit  
General      System.App...
General      Event        
General      UnhandledE...
General      System.App...
General      Event        
General      DomainUnload 
General      System.App...
General      Event        
General      FirstChanc...
General      System.App...
操作おわり
この時点では参照が残ったまま
powershell
powershell 14216 =======================================================
Amd64
     1C9018E7970           32 System.__ComObject 1 False Windows.Foundation.Diagnostics.IAsyncCausalityTracerStatics,
     1C9018E79D0           32 Windows.Foundation.Diagnostics.TracingStatusChangedEventArgs 1 False Windows.Foundation.Diagnostics.ITracingStatusChangedEventArgs,
     1C901B99008           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1C901B99028           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1C901B994C8           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1C901B994E8           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1C901B99508           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1C901E3E158           32 System.__ComObject (GetRCWDataに失敗)
     1C901E96FA0           32 System.__ComObject (GetRCWDataに失敗)
     1C9028B7EC8           32 Microsoft.Office.Interop.Excel.ApplicationClass 1 False System.Management.Automation.IDispatch,Microsoft.Office.Interop.Excel._Application,
     1C9028C5050           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1C902987960           32 System.__ComObject 1 False Microsoft.Office.Interop.Excel.Workbooks,System.Management.Automation.ComInterop.IDispatch,
     1C90299FB50           32 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1C9029F23F0           32 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     1C902A046E0           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1C902A511E8           32 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     1C902A625F0           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1C902AC6D30           32 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     1C902AD5F48           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1C902B25848           32 System.__ComObject 2 False
     1C902B43938           32 System.__ComObject 1 False
     1C902B43B28           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902B46E70           32 System.__ComObject 21 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1C902B7AF08           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902BA0990           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902BC8388           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902BEE238           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902C141C8           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902C3A248           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902C60190           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902C86220           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902CAC118           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902CD1FB0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902CF7E68           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902D1DD60           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902D43D28           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902D69C20           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902D8FC28           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902DB5DF0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902DDBCB8           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902E01EC0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902E27D70           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902E4DC78           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1C902E7D7A0           32 System.__ComObject 1 False Microsoft.Office.Interop.Excel.Workbooks,System.Management.Automation.ComInterop.IDispatch,
GCを強制
powershell
powershell 14216 =======================================================
Amd64
     1C9018E64A8           32 System.__ComObject 1 False Windows.Foundation.Diagnostics.IAsyncCausalityTracerStatics,
     1C901CE14A8           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
```

</div></details>

# 更に踏み込む

この結果を見るに、「ベストプラクティス」とCOMオブジェクト操作コードの分離はCOMオブジェクトの参照の解放という点で同じ事をしていることになります。
両者の共通点は全てのCOMオブジェクトを参照不可にしていることです。割愛しますが、ガベージコレクトの仕組み上、参照可能なポイントが全て消滅することは重要な意味を持ちます。

参考:[ガベージ コレクションの基礎 | Microsoft Docs](https://docs.microsoft.com/ja-jp/dotnet/standard/garbage-collection/fundamentals)

何が言いたいかというと、理屈上はCOMオブジェクトを割り当てた全ての変数に`$null`を代入してからガベージコレクトを実行しても正しい解放が可能です。
何より重要なのは、**この条件ならPowershellで行うのはとても簡単**ということです。

## pwsh 7.1 rc-1の場合

```powershell
$EnumRcw = '~\Repos\MemoryCheck\bin\x64\Release\EnumRcw.exe'

'実行前'
& $EnumRcw pwsh

'Excel操作'

$app = New-Object -ComObject Excel.Application
$app.Workbooks.Open((Convert-Path Temp:\Events.xlsx)).Sheets["Sheet1"].Range("A1:C7").foreach{
    $_ | select NumberFormat, Value2
}

$app.Workbooks.Close()

'操作おわり'

'この時点ではExcelプロセスは残ったまま'
& $EnumRcw pwsh

'__ComObjectを参照している変数をクリアする($val=$nullと同等の処理)
※VBAのオブジェクトを参照している変数にNothingを設定するのと同じパターン'
gv | ? Value -is [__ComObject] | clv


& $EnumRcw pwsh
'GCを強制'
[gc]::Collect() #
[gc]::WaitForPendingFinalizers() #

& $EnumRcw pwsh
```

<details><summary>結果</summary><div>

```shell-session:結果
実行前
pwsh
pwsh 1852 =======================================================
Amd64
Excel操作

pwsh
pwsh 1852 =======================================================
Amd64
     24EE709F2D8           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE709F7F0           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE709F898           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE70A4DD0           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70A4E80           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE70A4FC0           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70A5068           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE70A51A8           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70A5258           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE70A5398           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70A5448           24 System.__ComObject 2 False System.Management.Automation.ComInterop.IDispatch,System.Management.Automation.IDispatch,System.Runtime.InteropServices.IDispatch,
     24EE70A56C0           24 System.__ComObject 44 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70CAAD8           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.IEnumVARIANT,
     24EE70CAC70           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE70E9250           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7107488           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE71256D0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7143908           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7161BB8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE717FE20           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE719E058           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE71BC2B0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE71DA598           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE71F87D0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7216A18           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7234C68           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7252EA0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE72710F8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE728F348           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE72AD580           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE72CB8E0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE72E9B30           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7307D80           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7325FD8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE73636D0           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
NumberFormat Value2
------------ ------
G/標準         MemberType
G/標準         Name
G/標準         DeclaringType
G/標準         Event
G/標準         UnhandledException
G/標準         System.AppContext
G/標準         Event
G/標準         FirstChanceException
G/標準         System.AppContext
G/標準         Event
G/標準         ProcessExit
G/標準         System.AppContext
G/標準         Event
G/標準         UnhandledException
G/標準         System.AppDomain
G/標準         Event
G/標準         DomainUnload
G/標準         System.AppDomain
G/標準         Event
G/標準         FirstChanceException
G/標準         System.AppDomain
操作おわり
この時点ではExcelプロセスは残ったまま
__ComObjectを参照している変数をクリアする($val=$nullと同等の処理)
※VBAのオブジェクトを参照している変数にNothingを設定するのと同じパターン
pwsh
pwsh 1852 =======================================================
Amd64
     24EE709F2D8           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE709F7F0           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE709F898           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE70A4DD0           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70A4E80           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE70A4FC0           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70A5068           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE70A51A8           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70A5258           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE70A5398           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70A5448           24 System.__ComObject 2 False System.Management.Automation.ComInterop.IDispatch,System.Management.Automation.IDispatch,System.Runtime.InteropServices.IDispatch,
     24EE70A56C0           24 System.__ComObject 44 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE70CAAD8           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.IEnumVARIANT,
     24EE70CAC70           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE70E9250           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7107488           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE71256D0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7143908           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7161BB8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE717FE20           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE719E058           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE71BC2B0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE71DA598           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE71F87D0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7216A18           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7234C68           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7252EA0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE72710F8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE728F348           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE72AD580           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE72CB8E0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE72E9B30           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7307D80           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE7325FD8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE73636D0           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
GCを強制
pwsh
pwsh 1852 =======================================================
Amd64
```

</div></details>

## pwsh 7.1 rc-1の場合（何処かに暗黙の参照が残るような操作を行った場合）

```powershell
$EnumRcw = '~\Repos\MemoryCheck\bin\x64\Release\EnumRcw.exe'

'実行前'
& $EnumRcw pwsh

'Excel操作'

$app = New-Object -ComObject Excel.Application
$app.Workbooks.Open((Convert-Path Temp:\Events.xlsx)).Sheets["Sheet1"].Range("A1:C7") | select NumberFormat, Value2

$app.Workbooks.Close()


'操作おわり'

'この時点ではExcelプロセスは残ったまま'
& $EnumRcw pwsh

'__ComObjectを参照している変数をクリアする($val=$nullと同等の処理)
※VBAのオブジェクトを参照している変数にNothingを設定するのと同じパターン'
gv | ? Value -is [__ComObject] | clv


& $EnumRcw pwsh
'GCを強制'
[gc]::Collect() #
[gc]::WaitForPendingFinalizers() #

& $EnumRcw pwsh

'自動変数をクリーンアップ'
1|%{$_} > $null
[gc]::Collect() #

& $EnumRcw pwsh
```

<details><summary>結果</summary><div>

```shell-session:結果
実行前
pwsh
pwsh 1852 =======================================================
Amd64
Excel操作

NumberFormat Value2
------------ ------
G/標準         MemberType
G/標準         Name
G/標準         DeclaringType
G/標準         Event
G/標準         UnhandledExc…
G/標準         System.AppCo…
G/標準         Event
G/標準         FirstChanceE…
G/標準         System.AppCo…
G/標準         Event
G/標準         ProcessExit
G/標準         System.AppCo…
G/標準         Event
G/標準         UnhandledExc…
G/標準         System.AppDo…
G/標準         Event
G/標準         DomainUnload
G/標準         System.AppDo…
G/標準         Event
G/標準         FirstChanceE…
G/標準         System.AppDo…
操作おわり
この時点ではExcelプロセスは残ったまま
pwsh
pwsh 1852 =======================================================
Amd64
     24EE80DE490           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80DE9A8           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80DEA50           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80E3FA8           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80E4058           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80E4198           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80E4240           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80E4380           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80E4430           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80E4570           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80E4620           24 System.__ComObject 2 False System.Runtime.InteropServices.IDispatch,
     24EE80E77A0           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.IEnumVARIANT,
     24EE80E7938           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE80E8850           24 System.__ComObject 42 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE8112358           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81314E8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81506A0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8170810           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE818FCA8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81AF0D0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81CE478           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81ED8A8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE820CCE8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE822C090           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE824B448           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE826A888           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8289C30           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE82A9060           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE82C84A0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE82E7830           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8306C00           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8326040           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE83453D0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8364818           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8383F40           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
__ComObjectを参照している変数をクリアする($val=$nullと同等の処理)
※VBAのオブジェクトを参照している変数にNothingを設定するのと同じパターン
pwsh
pwsh 1852 =======================================================
Amd64
     24EE80DE490           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80DE9A8           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80DEA50           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80E3FA8           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80E4058           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80E4198           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80E4240           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80E4380           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80E4430           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     24EE80E4570           24 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE80E4620           24 System.__ComObject 2 False System.Runtime.InteropServices.IDispatch,
     24EE80E77A0           24 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.IEnumVARIANT,
     24EE80E7938           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE80E8850           24 System.__ComObject 42 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     24EE8112358           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81314E8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81506A0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8170810           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE818FCA8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81AF0D0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81CE478           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE81ED8A8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE820CCE8           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE822C090           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE824B448           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE826A888           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8289C30           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE82A9060           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE82C84A0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE82E7830           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8306C00           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8326040           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE83453D0           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8364818           24 System.__ComObject 1 False System.Management.Automation.IDispatch,System.Management.Automation.ComInterop.IDispatch,
     24EE8383F40           24 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
GCを強制
pwsh
pwsh 1852 =======================================================
Amd64
     24EE4BD3F68           24 System.__ComObject 2 False System.Runtime.InteropServices.IDispatch,
自動変数をクリーンアップ
pwsh
pwsh 1852 =======================================================
Amd64
```

</div></details>

## Windows Powershell5.1の場合

```powershell
$excelPath = gci D:\temp\Events.xlsx

$EnumRcw = '~\Repos\MemoryCheck\bin\x64\Release\EnumRcw.exe'

'実行前'
& $EnumRcw powershell

'Excel操作'

$app = New-Object -ComObject Excel.Application
$app.Workbooks.Open($excelPath).Sheets["Sheet1"].Range("A1:C7") | select NumberFormat, Value2

$app.Workbooks.Close()

'操作おわり'

'この時点では参照が残ったまま'
& $EnumRcw powershell

'__ComObjectを参照している変数をクリアする($val=$nullと同等の処理)
※VBAのオブジェクトを参照している変数にNothingを設定するのと同じパターン'
gv | ? Value -is [__ComObject] | clv


& $EnumRcw powershell
'GCを強制'
[gc]::Collect() #
[gc]::WaitForPendingFinalizers() #

& $EnumRcw powershell

'自動変数をクリーンアップ'
1|%{$_} > $null
[gc]::Collect() #

& $EnumRcw powershell
```

<details><summary>結果</summary><div>

```shell-session:結果
実行前
powershell
powershell 11648 =======================================================
Amd64
     1F5BFF47970           32 System.__ComObject 1 False Windows.Foundation.Diagnostics.IAsyncCausalityTracerStatics,
     1F5BFF479D0           32 Windows.Foundation.Diagnostics.TracingStatusChangedEventArgs 1 False Windows.Foundation.Diagnostics.ITracingStatusChangedEventArgs,
     1F5C024D740           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1F5C024D760           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1F5C024DC00           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1F5C024DC20           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1F5C024DC40           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1F5C049C338           32 System.__ComObject (GetRCWDataに失敗)
     1F5C04F7190           32 System.__ComObject (GetRCWDataに失敗)
Excel操作

NumberFormat Value2       
------------ ------
General      MemberType
General      Name
General      DeclaringType
General      Event
General      UnhandledE...
General      System.App...
General      Event        
General      FirstChanc...
General      System.App...
General      Event        
General      ProcessExit  
General      System.App...
General      Event        
General      UnhandledE...
General      System.App...
General      Event        
General      DomainUnload 
General      System.App...
General      Event        
General      FirstChanc...
General      System.App...
操作おわり
この時点では参照が残ったまま
powershell
powershell 11648 =======================================================
Amd64
     1F5BFF47970           32 System.__ComObject 1 False Windows.Foundation.Diagnostics.IAsyncCausalityTracerStatics,
     1F5BFF479D0           32 Windows.Foundation.Diagnostics.TracingStatusChangedEventArgs 1 False Windows.Foundation.Diagnostics.ITracingStatusChangedEventArgs,
     1F5C024D740           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1F5C024D760           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1F5C024DC00           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1F5C024DC20           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1F5C024DC40           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1F5C049C338           32 System.__ComObject (GetRCWDataに失敗)
     1F5C04F7190           32 System.__ComObject (GetRCWDataに失敗)
     1F5C0EEC298           32 Microsoft.Office.Interop.Excel.ApplicationClass 1 False System.Management.Automation.IDispatch,Microsoft.Office.Interop.Excel._Application,
     1F5C0EF9648           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C0FBBF58           32 System.__ComObject 1 False Microsoft.Office.Interop.Excel.Workbooks,System.Management.Automation.ComInterop.IDispatch,
     1F5C0FD45E0           32 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C1028500           32 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     1F5C103B7B8           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C10882D8           32 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     1F5C1099140           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C10FD880           32 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     1F5C110CAB0           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C115C3B0           32 System.__ComObject 2 False
     1F5C1179D98           32 System.__ComObject 1 False 
     1F5C1179F88           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C117E108           32 System.__ComObject 21 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C11B2198           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C11D7C00           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C11FD710           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1225540           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C124B4C8           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1271540           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1297480           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C12BD508           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C12E33F8           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C13092A0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C132F150           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1355040           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C137B000           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C13A0EF0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C13C6EF0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C13ED0B0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1412F70           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1439170           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C145F018           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1484F18           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C14B4A90           32 System.__ComObject 1 False Microsoft.Office.Interop.Excel.Workbooks,System.Management.Automation.ComInterop.IDispatch,
__ComObjectを参照している変数をクリアする($val=$nullと同等の処理)
※VBAのオブジェクトを参照している変数にNothingを設定するのと同じパターン
powershell
powershell 11648 =======================================================
Amd64
     1F5BFF47970           32 System.__ComObject 1 False Windows.Foundation.Diagnostics.IAsyncCausalityTracerStatics,
     1F5BFF479D0           32 Windows.Foundation.Diagnostics.TracingStatusChangedEventArgs 1 False Windows.Foundation.Diagnostics.ITracingStatusChangedEventArgs,
     1F5C024D740           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1F5C024D760           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1F5C024DC00           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1F5C024DC20           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyEnum,
     1F5C024DC40           32 System.__ComObject 1 False Microsoft.Win32.IAssemblyName,
     1F5C049C338           32 System.__ComObject (GetRCWDataに失敗)
     1F5C04F7190           32 System.__ComObject (GetRCWDataに失敗)
     1F5C0EEC298           32 Microsoft.Office.Interop.Excel.ApplicationClass 1 False System.Management.Automation.IDispatch,Microsoft.Office.Interop.Excel._Application,
     1F5C0EF9648           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C0FBBF58           32 System.__ComObject 1 False Microsoft.Office.Interop.Excel.Workbooks,System.Management.Automation.ComInterop.IDispatch,
     1F5C0FD45E0           32 System.__ComObject 2 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C1028500           32 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     1F5C103B7B8           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C10882D8           32 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     1F5C1099140           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C10FD880           32 System.__ComObject 1 False System.Management.Automation.ComInterop.IDispatch,
     1F5C110CAB0           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C115C3B0           32 System.__ComObject 2 False
     1F5C1179D98           32 System.__ComObject 1 False 
     1F5C1179F88           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C117E108           32 System.__ComObject 21 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C11B2198           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C11D7C00           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C11FD710           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1225540           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C124B4C8           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1271540           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1297480           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C12BD508           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C12E33F8           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C13092A0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C132F150           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1355040           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C137B000           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C13A0EF0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C13C6EF0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C13ED0B0           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1412F70           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1439170           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C145F018           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C1484F18           32 System.__ComObject 1 False System.Management.Automation.IDispatch,
     1F5C14B4A90           32 System.__ComObject 1 False Microsoft.Office.Interop.Excel.Workbooks,System.Management.Automation.ComInterop.IDispatch,
GCを強制
powershell
powershell 11648 =======================================================
Amd64
     1F5BFF464A8           32 System.__ComObject 1 False Windows.Foundation.Diagnostics.IAsyncCausalityTracerStatics,
     1F5C0343638           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
     1F5C03C1F60           32 System.__ComObject 2 False 
自動変数をクリーンアップ
powershell
powershell 11648 =======================================================
Amd64
     1F5BFF464A8           32 System.__ComObject 1 False Windows.Foundation.Diagnostics.IAsyncCausalityTracerStatics,
     1F5C0343638           32 System.__ComObject 1 False System.Runtime.InteropServices.ComTypes.ITypeInfo,
```

</div></details>

# 逆に考えるんだ 「解放処理するべき変数が分からなくてもいいさ」と考えるんだ

重要なのは`gv | ? Value -is [__ComObject] | clv`です。エイリアスを使わないと以下のようになります。

```powershell
Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable
```

`Get-Variable`はPowershellに設定されている変数を取得するコマンドレットです。
引数を設定しない場合、現在のスコープで取得可能な全ての変数を取得出来ます。
`Where-Object Value -is [__ComObject]`でフィルタリングして型が`__ComObject`であるものだけを処理の対象にします。
そして`Clear-Variable`で該当する全ての変数の内容をクリアします。これは変数に`$null`を代入するのと同等な操作です。
この方法の嬉しいとこは解放が必要な変数の数や名前を覚える必要が無いことです。
`[System.Runtime.InteropServices.Marshal]::ReleaseComO　bject()`を使う場合と違って解放する順序を気にする必要もあり ません。

# それでも残る参照への対処

`Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable`でも解放されない参照はPowershell内部の自動処理の影響で通常ではアクセス出来ない所に参照が残っていると考えると良いでしょう。
COMオブジェクトをパイプラインで`Select-Object`する場合などが該当します。
この場合は、普通のオブジェクトで同じ操作を行えば参照が上書きされるようです。
`1|%{$_} > $null`を行っているのはこのためです。　　
新しい参照でCOMへの参照を上書きすることで解放を実現します。
※もしかしたら`Get-Variable`のパラメータを適切に設定すれば不要かも知れませんが今は分からないのでこの方法を使っています。

# おわりに

今回のアプローチとは逆に、[Microsoft.Diagnostics.Runtime](https://github.com/microsoft/clrmd)のアセンブリをPowershellでロード、自らのプロセスの未解放COMオブジェクトのアドレスを収集させるバックグラウンドジョブを生成、そのアドレスを元に一つ残らずReleaseCOMObjectする方法も出来そうですが手間なので必要になったら検討します。
