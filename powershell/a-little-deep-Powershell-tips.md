<!--
title:   少しディープなPowershell備忘録
tags:    PowerShell
id:      9a4cc6f760e6abe96048
private: false
-->
# これは何

個人的なPowershell備忘録です。
思い出したり記事を探すのが手間だったりするものを中心に。
ディープな部分も含めてとりあえず記録しておこうと思ったものを追記予定。
**特に明記してない場合は`Powershell 7.x`の内容です。**

## 未翻訳ドキュメント要約と私的補足

Powershellの未翻訳ドキュメントからあれこれ
知らなかったそんな機能……って沢山感じました。

### [About calculated properties](https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_calculated_properties?view=powershell-7)

一部コマンドレットの`properties`パラメーターや`Groupby`パラメータでは連想配列で手の込んだ処理が出来ます。
その解説。

#### 連想配列のキー構成

- `name`/`label` - 作成するプロパティの名前を指定する。どちらを使っても良い。
- `expression` - プロパティの値を計算するためのスクリプトブロックを指定する。
- `alignment` - 表形式の出力を生成するコマンドレットで、値が列にどのように表示されるかを定義するために使用する。値は `'left'`、`'center'`、または `'right'` のいずれかでなければならない。
- `formatstring` - フォーマット文字列を指定する。文法は`-f`演算子、`[string]::Format`と同じ。
- `width` - カラムの幅を指定する。0より大きい数値を指定しなければならない。
- `depth` - `Format-Custom`の`-Depth`パラメータは全プロパティの展開深度を指定するが、`depth`キーを使うと個別に深度を設定出来る。
- `ascending` / `descending` - ソート順を指定する。`boolean`値。

`expression`のみが必須。連想配列ではなく`Scriptblock`を指定した場合は`expression`のみが指定される。
連想配列キーは他のPowershellのパラメータ指定と同様に被りが起きないところまで省略可能。
手打ちでは1文字だけで指定していく場合がほとんど。

```powershell
# 省略なしの場合
ps code | ft id,
            name,
            @{name='Handles';expression={$_.Handles};alignment='left'},
            @{name='CPU';expression={$_.CPU};formatstring="####0.00"}

# 省略した場合
ps code | ft id,name,@{n='Handles';e={$_.Handles};a='left'},@{n='CPU';e={$_.CPU};f="####0.00"}
```

結果

```powershell
   Id Name Handles    CPU
   -- ---- -------    ---
 6924 Code 1059    756.02
 7344 Code 213       0.95
 7552 Code 464       1.92
15900 Code 398      49.06
15948 Code 279       7.13
16116 Code 403      33.30
21404 Code 214       1.19
22532 Code 215       0.48
36224 Code 501     274.56
38748 Code 259       1.86
40000 Code 305       0.11
40684 Code 535     210.70
41032 Code 1836    225.59
41332 Code 223       1.45
41564 Code 220     204.81
41724 Code 218       0.53
```

#### `Compare-Object`

> - `expression`

`Compare-Object`はプロパティで比較基準を設定出来る。

`[Process]`型を様々な方法で比較してみる。

```powershell:下準備
$ps= get-process code
$ps[0] | fl
$ps[1] | fl
```

結果

```console
Id      : 6924
Handles : 1065
CPU     : 536.8125
SI      : 1
Name    : Code


Id      : 7344
Handles : 213
CPU     : 0.953125
SI      : 1
Name    : Code

```

通常はNameが比較対象になるので「同じ」だと判定される。

```Powershell
 Compare-Object  $ps[0] $ps[1]
```

 プロパティ名を指定すると比較対象が変わる。

```Powershell
 Compare-Object  $ps[0] $ps[1] -Property Id

  Id SideIndicator
  -- -------------
7344 =>
6924 <=

```

`Select-Object`と同様に複数プロパティの指定も可能。

```Powershell
 Compare-Object  $ps[0] $ps[1] -Property Name,Id

Name   Id SideIndicator
----   -- -------------
Code 7344 =>
Code 6924 <=

```

プロパティで演算を行った例。
`Id`は偶数なので`{$_.Id % 2}`は両方とも0で同じなのでこの結果になる。

```Powershell
Compare-Object  $ps[0] $ps[1] -Property Name,{$_.Id % 2}
```

演算なしとの比較

```Powershell
Compare-Object  $ps[0] $ps[1] -Property Name,{$_.Id % 2},Id

Name $_.Id % 2   Id SideIndicator
---- ---------   -- -------------
Code         0 7344 =>
Code         0 6924 <=
```

##### `Compare-Object`応用

###### 型を比較する

メンバーの型定義をすることによりダックタイピングを実現出来るかも？
`[ValidateScript]`で使えそう。
※Idは可視化のためにつけている。

```powershell
Compare-Object  $ps[0] $ps[1] -Property {$_.Name.Gettype()},{$_.Id.Gettype()},{$_.Parent.Gettype()},Id

$_.Name.Gettype()   : System.String
$_.Id.Gettype()     : System.Int32
$_.Parent.Gettype() : System.Diagnostics.Process
Id                  : 7344
SideIndicator       : =>

$_.Name.Gettype()   : System.String
$_.Id.Gettype()     : System.Int32
$_.Parent.Gettype() : System.Diagnostics.Process
Id                  : 6924
SideIndicator       : <=

```

###### 発想を自由にする

Get-Processをスクリプトブロック内で行うことで、Compare-Object実行時に実行されているプロセス数を比較している。

```Powershell
Compare-Object pwsh code -Property {(get-process $_).count}

(get-process $_).count SideIndicator
---------------------- -------------
                    16 =>
                     4 <=
```

##### `Compare-Object`実践

###### ファイル比較

準備

```powershell
PS >"---test.txt---"
>> cat .\test.txt
>>
>> "`n---test.log---"
>> cat .\test.log
---test.txt---
Line 1
Line 2
Line 3
Line 4
Line 5

---test.log---

Name                           Value
----                           -----
args                           {}
e                              System.IO.FileSystemEventArgs
Error                          {}
event                          System.Management.Automation.PSEventArgs
eventArgs                      System.IO.FileSystemEventArgs
eventSubscriber                System.Management.Automation.PSEventSubscriber
false                          False
MyInvocation                   System.Management.Automation.InvocationInfo
null
PSBoundParameters              {[source, System.IO.FileSystemWatcher], [e, System.IO.FileSystemEventArgs]}
PSCommandPath
PSDefaultParameterValues       {}
PSScriptRoot
script:Error                   {}
sender                         System.IO.FileSystemWatcher
source                         System.IO.FileSystemWatcher
true                           True

```

パス名は文字列比較になってしまう

```powershell
PS >Compare-Object .\test.txt .\test.log

InputObject SideIndicator
----------- -------------
.\test.log  =>
.\test.txt  <=

```

公式のドキュメントではこのようにして比較してる

```powershell

PS >Compare-Object (Get-Content .\test.txt) (Get-Content .\test.log)

InputObject                                                                                                SideIndicator
-----------                                                                                                -------------
                                                                                                           =>
Name                           Value                                                                       =>
----                           -----                                                                       =>
args                           {}                                                                          =>
e                              System.IO.FileSystemEventArgs                                               =>
Error                          {}                                                                          =>
event                          System.Management.Automation.PSEventArgs                                    =>
eventArgs                      System.IO.FileSystemEventArgs                                               =>
eventSubscriber                System.Management.Automation.PSEventSubscriber                              =>
false                          False                                                                       =>
MyInvocation                   System.Management.Automation.InvocationInfo                                 =>
null                                                                                                       =>
PSBoundParameters              {[source, System.IO.FileSystemWatcher], [e, System.IO.FileSystemEventArgs]} =>
PSCommandPath                                                                                              =>
PSDefaultParameterValues       {}                                                                          =>
PSScriptRoot                                                                                               =>
script:Error                   {}                                                                          =>
sender                         System.IO.FileSystemWatcher                                                 =>
source                         System.IO.FileSystemWatcher                                                 =>
true                           True                                                                        =>
                                                                                                           =>
Line 1                                                                                                     <=
Line 2                                                                                                     <=
Line 3                                                                                                     <=
Line 4                                                                                                     <=
Line 5                                                                                                     <=

```

`Property`パラメータで演算を行うとこうなる
結果が違うのは公式の例はファイル内容の1行ごとの文字列動詞を比較しているのに対し、この例ではファイル内容全体で比較しているため。

```powershell
PS >Compare-Object .\test.txt .\test.log -Property {cat $_}

cat $_                                                                                                               SideIndicato
                                                                                                                     r
------                                                                                                               ------------
{, Name                           Value, ----                           -----, args                           {}...} =>
{Line 1, Line 2, Line 3, Line 4...}                                                                                  <=

```

ファイルハッシュの比較も行える

```powershell
PS >Compare-Object ((Get-FileHash .\test.txt).hash) ((Get-FileHash .\test.log).hash)

InputObject                                                      SideIndicator
-----------                                                      -------------
48B6D671ED303B15C649F94163DD8167280DE44D6673C67A7C1BEE834CEAFFB6 =>
7B2C588D37C473A32203AFAF5C7628EA75EB79B33C799850E509CA37BCBE4195 <=


PS >Compare-Object .\test.txt .\test.log -Property {(Get-FileHash $_).Hash}

(Get-FileHash $_).Hash                                           SideIndicator
----------------------                                           -------------
48B6D671ED303B15C649F94163DD8167280DE44D6673C67A7C1BEE834CEAFFB6 =>
7B2C588D37C473A32203AFAF5C7628EA75EB79B33C799850E509CA37BCBE4195 <=

```

組み合わせる

```powershell

$hashParams = @{
    ReferenceObject  = get-item .\test.txt
    DifferenceObject = get-item .\test.log
    Property         = @(
        { $_.name }
        { (Get-FileHash $_).Hash }
        { $_.Attributes }
        { $_.LastWriteTime }
    )
}
Compare-Object @hashParams
```

結果

```shell-session
 $_.name                 : test.log
 (Get-FileHash $_).Hash  : 48B6D671ED303B15C649F94163DD8167280DE44D6673C67A7C1BEE834CEAFFB6
 $_.Attributes           : Archive
 $_.LastWriteTime        : 2020/09/19 16:00:41
SideIndicator            : =>

 $_.name                 : test.txt
 (Get-FileHash $_).Hash  : 7B2C588D37C473A32203AFAF5C7628EA75EB79B33C799850E509CA37BCBE4195
 $_.Attributes           : Archive
 $_.LastWriteTime        : 2020/10/12 22:12:50
SideIndicator            : <=
`

`IncludeEqual`パラメータで同じものを表示させる事も出来る
ファイルをコピーして比較する

```powershell
copy .\test.txt .\test_copied.txt -Force
$hashParams = @{
    ReferenceObject  = get-item .\test.txt
    DifferenceObject = get-item .\test_copied.txt
    Property         = @(
        { $_.name }
        { (Get-FileHash $_).Hash }
        { $_.Attributes }
        { $_.LastWriteTime }
    )
}
```

この条件ではname以外は一致

```shell-session
 $_.name                 : test_copied.txt
 (Get-FileHash $_).Hash  : 7B2C588D37C473A32203AFAF5C7628EA75EB79B33C799850E509CA37BCBE4195
 $_.Attributes           : Archive
 $_.LastWriteTime        : 2020/10/12 22:12:50
SideIndicator            : =>

 $_.name                 : test.txt
 (Get-FileHash $_).Hash  : 7B2C588D37C473A32203AFAF5C7628EA75EB79B33C799850E509CA37BCBE4195
 $_.Attributes           : Archive
 $_.LastWriteTime        : 2020/10/12 22:12:50
SideIndicator            : <=
```

名前以外を比較する

```powershell
$hashParams = @{
    ReferenceObject  = get-item .\test.txt
    DifferenceObject = get-item .\test_copied.txt
    Property         = @(
        # { $_.name }
        { (Get-FileHash $_).Hash }
        { $_.Attributes }
        { $_.LastWriteTime }
    )
    IncludeEqual     = $true
    ExcludeDifferent = $true
}
Compare-Object @hashParams
```

結果

```shell-session
 (Get-FileHash $_).Hash                                           $_.Attributes   $_.LastWriteTime   SideIndicator
------------------------                                         --------------- ------------------  -------------
7B2C588D37C473A32203AFAF5C7628EA75EB79B33C799850E509CA37BCBE4195         Archive 2020/10/12 22:12:50 ==

```

#### `ConvertTo-Html`

>- `ConvertTo-Html`
> - `name`/`label` - optional (added in PowerShell 6.x)
> - `expression`
> - `width` - optional
> - `alignment` - optional

#### `Format-Custom`

>- `Format-Custom`
> - `expression`
> - `depth` - optional

#### `Format-List`

>- `Format-List`
> - `name`/`label` - optional
> - `expression`
> - `formatstring` - optional
>
> This same set of key-value pairs also apply to calculated property values
> passed to the **GroupBy** parameter for all `Format-*` cmdlets.

#### `Format-Table`

>- `Format-Table`
> - `name`/`label` - optional
> - `expression`
> - `formatstring` - optional
> - `width` - optional
> - `alignment` - optional

#### `Format-Wide`
>
>- `Format-Wide`
> - `expression`
> - `formatstring` - optional

#### `Group-Object`

>- `Group-Object`
> - `expression`

#### `Measure-Object`

>- `Measure-Object`
> - Only supports a script block for the expression, not a hashtable.
> - Not supported in PowerShell 5.1 and older.

#### `Select-Object`

>- `Select-Object`
> - `name`/`label` - optional
> - `expression`

#### `Sort-Object`

>- `Sort-Object`
> - `expression`
> - `ascending`/`descending` - optional

#### 補足

>> [!NOTE]
>> The value of the `expression` can be a script block instead of a
>> hashtable. For more information, see the [Notes](#notes) section.

## お気に入りの拡張モジュール・Nugetのライブラリ

Powershellの真価はModule/外部ライブラリ取り込める拡張性にあり。
※PowershellGetの扱い方はググると出てくると思います。

### ClassExplorer

.NETの型情報を調べるツールです。
何を入れるのがオススメかと聞かれたらClassExplorerと応えます。

- パラメータ、返り値の型からエラー解析や応用を探る
- Powershellで行いたいことがあるとき、良い感じのクラス、メソッドなどがあるのか確認
- インストールしたモジュールが使っているアセンブリの解析
  - 主にNugetで公開されているライブラリ、そのラッパーモジュールで使用。

```powershell
❯ gcm -mo ClassExplorer

CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Cmdlet          Find-Member                                        1.1.0      ClassExplorer
Cmdlet          Find-Namespace                                     1.1.0      ClassExplorer
Cmdlet          Find-Type                                          1.1.0      ClassExplorer
Cmdlet          Get-Assembly                                       1.1.0      ClassExplorer
Cmdlet          Get-Parameter                                      1.1.0      ClassExplorer
```

#### 使用例：Eventを持つオブジェクトを確認する

<details><summary>例</summary><div>

```Powershell
~
❯ Find-Member -MemberType Event

   ReflectedType: System.AppContext

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
UnhandledException   Event         True    Void UnhandledExceptionEventHandler.Invoke(Object sender, UnhandledExc…
FirstChanceException Event         True    Void EventHandler`1.Invoke(Object sender, FirstChanceExceptionEventArg…
ProcessExit          Event         True    Void EventHandler.Invoke(Object sender, EventArgs e)

   ReflectedType: System.AppDomain

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
UnhandledException   Event        False    Void UnhandledExceptionEventHandler.Invoke(Object sender, UnhandledExc…
DomainUnload         Event        False    Void EventHandler.Invoke(Object sender, EventArgs e)
FirstChanceException Event        False    Void EventHandler`1.Invoke(Object sender, FirstChanceExceptionEventArg…
ProcessExit          Event        False    Void EventHandler.Invoke(Object sender, EventArgs e)
AssemblyLoad         Event        False    Void AssemblyLoadEventHandler.Invoke(Object sender, AssemblyLoadEventA…
AssemblyResolve      Event        False    Assembly ResolveEventHandler.Invoke(Object sender, ResolveEventArgs ar…
ReflectionOnlyAssem… Event        False    Assembly ResolveEventHandler.Invoke(Object sender, ResolveEventArgs ar…
TypeResolve          Event        False    Assembly ResolveEventHandler.Invoke(Object sender, ResolveEventArgs ar…
ResourceResolve      Event        False    Assembly ResolveEventHandler.Invoke(Object sender, ResolveEventArgs ar…

   ReflectedType: System.Progress`1[T]

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
ProgressChanged      Event        False    Void EventHandler`1.Invoke(Object sender, T e)

   ReflectedType: System.Threading.Tasks.TaskScheduler

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
UnobservedTaskExcep… Event         True    Void EventHandler`1.Invoke(Object sender, UnobservedTaskExceptionEvent…

   ReflectedType: System.Runtime.Loader.AssemblyLoadContext

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
ResolvingUnmanagedD… Event        False    IntPtr Func`3.Invoke(Assembly arg1, String arg2)
Resolving            Event        False    Assembly Func`3.Invoke(AssemblyLoadContext arg1, AssemblyName arg2)
Unloading            Event        False    Void Action`1.Invoke(AssemblyLoadContext obj)

   ReflectedType: System.Reflection.Assembly

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
ModuleResolve        Event        False    Module ModuleResolveEventHandler.Invoke(Object sender, ResolveEventArg…

# (中略)

   ReflectedType: Microsoft.Windows.HostGuardianService.Diagnostics.Synthesizer`2[TTrace,THost]

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
TargetDiscovered     Event        False    Void EventHandler`1.Invoke(Object sender, TargetDiscoveredEventArgs e)

   ReflectedType: Microsoft.Windows.RemoteAttestation.Core.TcgPcrEvent

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
UnknownPcrEvent      Event         True    Void EventHandler`1.Invoke(Object sender, TcgParserEventArgs e)

   ReflectedType: System.Drawing.Printing.PrintDocument

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
BeginPrint           Event        False    Void PrintEventHandler.Invoke(Object sender, PrintEventArgs e)
EndPrint             Event        False    Void PrintEventHandler.Invoke(Object sender, PrintEventArgs e)
PrintPage            Event        False    Void PrintPageEventHandler.Invoke(Object sender, PrintPageEventArgs e)
QueryPageSettings    Event        False    Void QueryPageSettingsEventHandler.Invoke(Object sender, QueryPageSett…
Disposed             Event        False    Void EventHandler.Invoke(Object sender, EventArgs e)

   ReflectedType: System.IO.FileSystemWatcher

Name                 MemberType  IsStatic  Definition
----                 ----------  --------  ----------
Changed              Event        False    Void FileSystemEventHandler.Invoke(Object sender, FileSystemEventArgs …
Created              Event        False    Void FileSystemEventHandler.Invoke(Object sender, FileSystemEventArgs …
Deleted              Event        False    Void FileSystemEventHandler.Invoke(Object sender, FileSystemEventArgs …
Error                Event        False    Void ErrorEventHandler.Invoke(Object sender, ErrorEventArgs e)
Renamed              Event        False    Void RenamedEventHandler.Invoke(Object sender, RenamedEventArgs e)
Disposed             Event        False    Void EventHandler.Invoke(Object sender, EventArgs e)
```

</div></details>

### ImportExcel

Excelファイルを生成するならこのモジュール一択です。

#### 使用例：パイプラインでExcelファイル生成

```powershell
Find-Member -MemberType Event | select * | Export-Excel Temp:/Events.xlsx -AutoSize -Show
```

オプション豊富です。一番気に入っているのは複数テーブル構成のExcelを作るのが苦痛ではなくなることです。

```Powershell
Find-Member -MemberType Event | group {$_.ReflectedType.Namespace  -replace "(^[^\.]+\.[^\.]+)\..+$",'$1'} | % {
 $_.Group | Export-Excel Temp:/ObjectEventsGroupByNameSpace.xlsx -WorksheetName $_.Name -AutoSize -TableStyle Medium1
}
```

### AngleParse

[PowerShellから簡単にスクレイピングするためのツールを作った](https://qiita.com/kamome283/items/5b976a27ed203e959b09)
Powershellの考え方でスクレイピング出来るありがたいモジュール。パイプラインセレクタが嬉しい。
