<!--
title:   Powershellでも(手軽に)unfoldしたい
tags:    F#,PowerShell
id:      7bbf2576e8a3fef5f9a7
private: false
-->
# これは何？

`F#`で知った`unfold`がとても便利だったので`Powershell`で実装しました。
`unfold`は高階関数の一つです。リストを生成する関数定義と初期値を引数にとってコレクションを生成します。
`while`ループや再帰関数などの処理は、`unfold`が有効な場合があります。

## `Powershell`での実装

こんな実装に落ち着いています。

```powershell
filter unfold ([scriptblock]$generator) {
    $state = $_
    do {
        $element, $state = $generator.InvokeWithContext(
            # functionsToDefine
            $null,
            # variablesToDefine
            ([psvariable]::new('_', $state)),
            # args
            $null
        )
        Write-Output $element
    } while ($null -ne $state)
}
```

パイプライン入力される値を初期値として、引数`generator`に指定された`[scriptblock]`が実行されます。
`generator`の返り値は要素数2の配列を想定しており、以下のように処理されます。

- 要素1…パイプライン出力
- 要素2
  - `$null`だった場合…`unfold`終了
  - `$null`じゃなかった場合…要素2を`$_`に指定して`generator`を実行

※本来の`unfold`は`generator`の返り値自体の有無で繰り返しを判断しますが、使い勝手の都合で今の仕様になりました。

### 使い方

#### 基本

数値を5まで出力したい場合はこうします。

参考:[三項演算子の構文の使用](https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_if?view=powershell-7.1#using-the-ternary-operator-syntax)

```powershell
1 | unfold {
    if ($_ -le 5) {
        # パイプライン出力される値, 次の入力値として使われる値
        return $_, ($_ + 1)
    }
    else {
        # $nullで生成終了
        return $null
    }
}

1 | unfold {
    # 返り値なし→$nullなので、return $nullは不要
    if ($_ -le 5) { $_, ($_ + 1) }
}

1 | unfold {
    # 三項演算子
    $_ -le 5 ? $_, ($_ + 1) : $null
}

1 | unfold {
    # これでも良い
    $_ -le 5 ? $_, ($_ + 1) : $null,$null
}
```

結果はすべて同じです。

```console:result
1
2
3
4
5
```

パイプラインで複数の値が入力された場合は。**入力された要素それぞれを初期値とした複数のシーケンスが生成されます。**

```powershell
3..5 | unfold {
    # <condition> ? <if-true> : <if-false>
    $_ -le 5 ? $_, ($_ + 1) : $null
}
```

結果

```console:result
3
4
5
4
5
5
```

返り値が`$someVar,$null`だった場合は、`$someVar`が出力されてから`unfold`が終了します。
本来の`unfold`には無い機能です、……多分。

```powershell
3..5 | unfold {
    $_ -le 5 ? $_, ($_ + 1) : 'end.',$null
}
```

結果

```console:result
3
4
5
end.
4
5
end.
5
end.
```

#### フィボナッチ数列

##### 6/10: 内容を修正、加筆

再帰関数の例題で有名なフィボナッチ数列を考えてみましょう。
まずは要素の値がn以上になるまで数列を生成する場合です。

```powershell
$n = 1000

[Tuple]::Create(1, 1) | unfold {
    if ($_[1] -gt $n) {
        return $null
    }
    else {
        return ($_[0] + $_[1]), [Tuple]::Create($_[1], $_[0] + $_[1])
    }
}
```

結果

```console
2
3
5
8
13
21
34
55
89
144
233
377
610
987
1597
```

数列の順番によって生成を制御したい場合は`generator`で頑張るよりも、`unfold`で無限シーケンスを生成して、`Select-Object`で制御するほうが楽です。

```powershell

# 最初から3番目まで
[Tuple]::Create(1, 1)
| unfold { ($_[0] + $_[1]), [Tuple]::Create($_[1], $_[0] + $_[1])}
| Select-Object -First 3

# 0から始めて10番目まで
[Tuple]::Create(0, 1)
| unfold { $_[0], [Tuple]::Create($_[1], $_[0] + $_[1])}
| Select-Object -First 10

# 0から始めて50番目と100番目だけ
[Tuple]::Create(0, 1)
| unfold {$_[0], [Tuple]::Create($_[1], $_[0] + $_[1])}
| Select-Object -Index 50,100

```

結果

```console
2
3
5

0
1
1
2
3
5
8
13
21
34

12586269025
354224848179261915075
```

型指定をしない場合、`powershell`は暗黙的にエラーにならない数値型に変換するようです。
特に工夫しなければ最終的にひたすら`[double]::IsInfinity`をパイプラインに流し続けるので気を付けましょう。

```powershell
# パイプラインに出力される値が途中から暗黙敵に
# オーバーフローエラーにならない[double]に変わっている
[Tuple]::Create(0, 1)
| unfold { $_[0], [Tuple]::Create($_[1], $_[0] + $_[1]) }
| Select-Object  -Index (0..5+100..105+1475..1480)
| ForEach-Object {"{0}:`t{1}" -f $_.gettype(), $_}

# パイプラインに流す値を[decimal]に指定
[Tuple]::Create(0, 1)
| unfold { [decimal]$_[0], [Tuple]::Create($_[1], $_[0] + $_[1]) }
| Select-Object  -Index (0..5+100..105+2000..2005)
| ForEach-Object {"{0}:`t{1}" -f $_.gettype(), $_}

# [double]::IsInfinityでシーケンスを終了させる
[Tuple]::Create(0, 1)
| unfold {
    if (-not [double]::IsInfinity($_[0] + $_[1])) {
        $_[0], [Tuple]::Create($_[1], $_[0] + $_[1])        
    }
} | ForEach-Object {
    $script:i=0
} {
    if($_){
        $script:i++
        "{0}:`t{1}:`t{2}" -f $script:i,${_}.gettype(), $_
    } 
}
```

結果

```console
System.Int32:   0
System.Int32:   1
System.Int32:   1
System.Int32:   2
System.Int32:   3
System.Int32:   5
System.Double:  3.54224848179262E+20
System.Double:  5.73147844013817E+20
System.Double:  9.27372692193079E+20
System.Double:  1.5005205362069E+21
System.Double:  2.42789322839998E+21
System.Double:  3.92841376460687E+21
System.Double:  8.07763763215622E+307
System.Double:  1.3069892237634E+308
System.Double:  ∞
System.Double:  ∞
System.Double:  ∞
System.Double:  ∞


System.Decimal: 0
System.Decimal: 1
System.Decimal: 1
System.Decimal: 2
System.Decimal: 3
System.Decimal: 5
System.Decimal: 354224848179262000000
System.Decimal: 573147844013817000000
System.Decimal: 927372692193079000000
System.Decimal: 1500520536206900000000
System.Decimal: 2427893228399980000000
System.Decimal: 3928413764606870000000
InvalidArgument: 
Line |
   3 |  | unfold { [decimal]$_[0], [Tuple]::Create($_[1], $_[0] + $_[1]) }
     |             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     | Cannot convert value "8.105590009602353E+28" to type "System.Decimal". Error: "Value was either too large or too small for a Decimal."

1:      System.Int32:   1
2:      System.Int32:   1
3:      System.Int32:   2
4:      System.Int32:   3
5:      System.Int32:   5
6:      System.Int32:   8
~~中略~~
1469:   System.Double:  4.50151316958984E+306
1470:   System.Double:  7.28360130920163E+306
1471:   System.Double:  1.17851144787915E+307
1472:   System.Double:  1.90687157879931E+307
1473:   System.Double:  3.08538302667846E+307
1474:   System.Double:  4.99225460547777E+307
```

`Select-Object`はインデックス系のオプションが指定された場合、出力するべき最後の要素をパイプラインに出力した時点でパイプライン処理を中断させる効果があります。
大変便利なのですが、中断されるとパイプライン上流の`end`ブロックが行われなくなります。
注意が必要です。

```powershell
[Tuple]::Create(0, 1)
| unfold { $_[0], [Tuple]::Create($_[1], $_[0] + $_[1]) }
| ForEach-Object { 'before begin' } { $_ } { 'before end' }
| Select-Object -First 10
| ForEach-Object { 'after begin' } { $_ } { 'after end' }
```

結果

```console
after begin
before begin
0
1
1
2
3
5
8
13
21
after end
```

問題はリソース管理です。
この記事ではこの問題について扱っています。
[関数内でのリソース解放処理 - PowerShell Scripting Weblog](http://winscript.jp/powershell/309)

なお、`scriptblock`に`cleanup{}`というリソース解放のためのブロックを追加するという[RFCがあり](https://github.com/PowerShell/PowerShell-RFC/blob/master/2-Draft-Accepted/RFC0059-Cleanup-Script-Block.md)、これが採用されると根本的な解決となるようです。

##### 余談:Powershellのタプル

今回は配列ではなく、`Tuple`を利用しました。
`Powershell`のタプルは以下のように使います。
各要素にインデックスアクセス可能です。
パイプライン入力等で展開されないので今回のようなケースでは重宝します。

```powershell
# コンストラクタを使用する方法
$foo = 9
$tup = [Tuple[int, int]]::new($foo, ($foo + 1))
$tup[0]
```

結果

```console
9
```

~~配列扱いなので複数の変数の割り当てが可能です。~~
~~実質デストラクタです。~~
`Powershell`にはタプルを分解する構文はないので関数を作っておくと便利でしょう。

```powershell
using namespace System.Management.Automation
using namespace System.Runtime.CompilerServices

function Measure-TupleLength {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ITuple]
        $Tuple
    )
    $DUPLICATE_LENGTH = 1
    $MAX_TUPLE_LENGTH = 8
    if ($Tuple.Length -eq $MAX_TUPLE_LENGTH -and ($Tuple[-1] -is [ITuple])) {
        $AdditionalLength = (Measure-TupleLength $Tuple[-1]) - $DUPLICATE_LENGTH
    }
    $Tuple.Length + $AdditionalLength
}
function Split-Tuple {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Runtime.CompilerServices.ITuple]
        $Tuple,
        [int[]]
        $Length = $null
    )
    begin {
        $Script:counter = 0
    }
    process {
        if ($Script:counter -ne 0) {
            $PSCmdlet.ThrowTerminatingError(
                [ErrorRecord]::new(
                    ([ArgumentException]'Input is not Single.'),
                    'InvalidPipelineInput',
                    [ErrorCategory]::InvalidArgument,
                    $input)
            )
        }
        $Script:counter++
    }
    end {
        if (($null -ne $Length) -and ($ActualLength = Measure-TupleLength $Tuple) -notin $Length) {
            $message = "Tuple Length is $ActualLength. It should be $($Length -join ' or ')."
            $PSCmdlet.ThrowTerminatingError(
                [ErrorRecord]::new(
                    ([ArgumentException]$message),
                    'InvalidTupleLength',
                    [ErrorCategory]::InvalidArgument,
                    $Tuple)
            )
        }
        if ($Tuple.Length -gt 0) {
            foreach ($i in 0..($Tuple.Length - 1)) {
                if ($i -eq 7 -and ($Tuple[$i] -is [ITuple]) ) {
                    Split-Tuple $Tuple[$i]
                }
                else {
                    Write-Output $Tuple[$i] -NoEnumerate
                }
            }
        }
    }
}
```

現状、`.NET`では長さが9以上のタプルを入れ子で表現していまが再帰関数で対応しました。
`unfold`を使うことを考えましたが、将来この関数を使って(比較的)型安全な`unfold`を書こうかなと思い、循環参照を回避するために見送りました。

普通に使うと想定される再帰呼び出し回数は2回ほどなので大丈夫でしょう……。
([参考：`TupleExtensions.Deconstruct`では21要素を持つタプルまで対応](
https://docs.microsoft.com/ja-jp/dotnet/api/system.tupleextensions.deconstruct?view=net-5.0#System_TupleExtensions_Deconstruct__21_System_Tuple___0___1___2___3___4___5___6_System_Tuple___7___8___9___10___11___12___13_System_Tuple___14___15___16___17___18___19___20______0____1____2____3____4____5____6____7____8____9____10____11____12____13____14____15____16____17____18____19____20__))

この関数を使うことで、変数へ分割代入時のバグを減らせます。

https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_assignment_operators?view=powershell-7.1#assigning-multiple-variables

※以下のコード例では`[ValueTuple]`を使用してますが、`[Tuple]`でも動作します。

```powershell
# 変数へ分割代入
$f, $g = [ValueTuple]::Create(4, 12) |  Split-Tuple
'$f, $g: {0}, {1}' -f $f, $g

# 要素数を制限可能
$f2, $g2 = [ValueTuple]::Create(6, 12, 32) |  Split-Tuple -Length (0..2)
'$f2, $g2: {0}, {1}' -f $f2, $g2

# [iTuple]派生型以外は受け付けない
{} | Split-Tuple

# 要素数9以上のタプルも分解可能
$values = [ValueTuple]::Create(
    1, 2, 'あ'..'お', 4, 5, [ValueTuple]::Create('a', 'b'), 7,
    ([ValueTuple]::Create(
            [ValueTuple]::Create((get-date)), 9, 10, 11, 12
        )
    )
) | Split-Tuple

(0..$values.GetUpperBound(0)).foreach{
    "value{0} : {1}" -f ($_ + 1), $values[$_]
}

# 要素数制限も可能
[ValueTuple]::Create(
    1, 2, 'あ'..'お', 4, 5, [ValueTuple]::Create('a', 'b'), 7,
    ([ValueTuple]::Create(
            [ValueTuple]::Create((get-date)), 9, 10, 11, 12
        )
    )
) | Split-Tuple -Length 15
```

結果

```console
$f, $g: 4, 12

Split-Tuple: 
Line |
   6 |  … , $g2 = [ValueTuple]::Create(6, 12, 32) |  Split-Tuple -Length (0..2)
     |                                               ~~~~~~~~~~~~~~~~~~~~~~~~~~
     | Tuple Length is 3. It should be 0 or 1 or 2.
$f2, $g2: , 

Split-Tuple: 
Line |
  10 |  {} | Split-Tuple
     |       ~~~~~~~~~~~
     | The input object cannot be bound to any parameters for the command either because the command does not take pipeline input or the input and its properties do not match any of the parameters that take pipeline input.

value1 : 1
value2 : 2
value3 : あ ぃ い ぅ う ぇ え ぉ お
value4 : 4
value5 : 5
value6 : (a, b)
value7 : 7
value8 : (2021/06/10 18:25:19)
value9 : 9
value10 : 10
value11 : 11
value12 : 12

Split-Tuple: 
Line |
  32 |  ) | Split-Tuple -Length 15
     |      ~~~~~~~~~~~~~~~~~~~~~~
     | Tuple Length is 12. It should be 15.
```

### RESTメソッド

`powershell`で`unfold`を実装しようと思った理由です。
欲しい要素をパイプラインに流しつつ複数回RESTメソッドを行います。

```powershell
function Get-YoutubeVideoIds {
    param (
        $PlayListId
    )
    @{
        key        = $Script:api
        part       = 'snippet'
        playlistId = $PlayListId
        fields     = 'nextPageToken,items(id,snippet(title,resourceId))'
        maxResults = '50'
    } | unfold {
        $res = Invoke-RestMethod https://www.googleapis.com/youtube/v3/playlistItems -b $_
        $nextState = if ($res.nextPageToken) {
            $_.pageToken = $res.nextPageToken
            $_
        }
        return $res.items, $nextState
    }
}

Get-YoutubeVideoIds PLfeA8kIs7Cocir1-TuSN3mOnj3qzyRShA
```

<details><summary>実行結果</summary><div>

```console

id                                                                   snippet
--                                                                   -------
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS41NkI0NEY2RDEwNTU3Q0M2 @{title=Logging in a DevOps world by Friedrich Weinmann; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4wMTcyMDhGQUE4NTIzM0Y5 @{title=Basic To Boss: Customizing Your PowerShell Prompt by Thomas Rayner; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS41MjE1MkI0OTQ2QzJGNzNG @{title=Securing PowerShell: Hands-On Lab by Ashley McGlone; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4wOTA3OTZBNzVEMTUzOTMy @{title=Firewall Evasion and Remote Access with OpenSSH by Anthony Nocentino; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4xMkVGQjNCMUM1N0RFNEUx @{title=Doctor, Don't Defenestrate: What to Do with Legacy Scripts by Michael Lombardi; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS41MzJCQjBCNDIyRkJDN0VD @{title=Going Core with VMware PowerCLI! by Luc Dekens; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5DQUNERDQ2NkIzRUQxNTY1 @{title=Working with PSGraph by Kevin Marquette; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS45NDk1REZENzhEMzU5MDQz @{title=Machine Learning Algorithms with H2o and PowerShell by Tome Tanasovski; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5GNjNDRDREMDQxOThCMDQ2 @{title=CypherDog2.0 - Bloodhound Dog Whispering with PowerShell by Walter Legowski; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS40NzZCMERDMjVEN0RFRThB @{title=Don't do that, do this instead: PowerShell worst practices and how to solve them by Chris Gardner; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5EMEEwRUY5M0RDRTU3NDJC @{title=Dungeons & Dragons & Development: How Playing Games Makes You a Better IT Pro by Michael Lombardi; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS45ODRDNTg0QjA4NkFBNkQy @{title=PowerShell in Azure Functions by Dongbo Wang & Joey Aiello; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4zMDg5MkQ5MEVDMEM1NTg2 @{title=The Windows Subsystem for Linux by Tara Raj; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS41Mzk2QTAxMTkzNDk4MDhF @{title=Advanced JEA Configurations by James Petty; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5EQUE1NTFDRjcwMDg0NEMz @{title=Introduction to Serverless Functions by Kirk Munro; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS41QTY1Q0UxMTVCODczNThE @{title=Monitoring Out, Observability In by Ebru Cucen; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4yMUQyQTQzMjRDNzMyQTMy @{title=Lord of the Configurations by Friedrich Weinmann; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS45RTgxNDRBMzUwRjQ0MDhC @{title=Using PowerShell Core to automate application packaging...with Habitat by Matt Wrock; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5ENDU4Q0M4RDExNzM1Mjcy @{title=F5 Declarative Configuration by James Arruda; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4yMDhBMkNBNjRDMjQxQTg1 @{title=Testing, Testing, 1...2...3: Using Pester for Infrastructure Validation by Brandon Olin; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5GM0Q3M0MzMzY5NTJFNTdE @{title=PowerShell + AutoRest + Swagger = Instant Modules by Adam Murray; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4zRjM0MkVCRTg0MkYyQTM0 @{title=Demystifying Terraform - "Hardcore" to Core Infrastructure-as-Code Tool by Chris Hunt; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS45NzUwQkI1M0UxNThBMkU0 @{title=Secure PowerShell web tools with System Frontier by Jay Adams; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5DNzE1RjZEMUZCMjA0RDBB @{title=Life after "git push" by Steven Murawski; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS43MTI1NDIwOTMwQjIxMzNG @{title=Ansible for the Windows Admin by Jeremy Murrah; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5DQ0MyQ0Y4Mzg0M0VGOEYw @{title=Publishing and Managing Modules in an Internal Repository by Kevin Marquette; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4yQUE2Q0JEMTk4NTM3RTZC @{title=Completely Automate Managing Windows Software...Forever by Daniel Franciscus; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5DMkU4NTY1QUFGQTYwMDE3 @{title=Beyond Pester 102: Acceptance testing with PowerShell by Glenn Sarti; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS44Mjc5REFBRUE2MTdFRDU0 @{title=Moving Up the Monitoring Stack by Steven Murawski; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS43NDhFRTgwOTRERTU4Rjg3 @{title=Sipping psake: Creating a Build and Release Pipeline for PowerShell by Brandon Olin; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS41QUZGQTY5OTE4QTREQUU4 @{title=PowerShell Universal Dashboard from start to finish by Adam Driscoll; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4zRDBDOEZDOUM0MDY5NEEz @{title=Look smarter: deliver your work in Excel by James O'Neill; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS42MTI4Njc2QjM1RjU1MjlG @{title=Turn your logs into actionable data at any scale with AWS by Andrew Pearce; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS45RjNFMDhGQ0Q2RkFCQTc1 @{title=Writing Clustered Applications with Windows PowerShell and Apache Zookeeper by Tome Tanasovski; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5BRjJDODk5REM0NjkzMUIy @{title=0-60 with PowerShell on AWS by Andrew Pearce & Steve Roberts; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4zQzFBN0RGNzNFREFCMjBE @{title=Automating Active Directory Health Checks by Mike Kanakos; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS45NkVENTkxRDdCQUFBMDY4 @{title=Deep Web: A Web Cmdlets Deep Dive by Mark Kraus; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5DNkMwRUI2MkI4QkI4NDFG @{title=Bullet-proofing Patterns & Practices by Joel "Jaykul" Bennett; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5DRUQwODMxQzUyRTlGRkY3 @{title=Unexplained phenomena: powerful tricks you likely didn't know were even possible by Kirk Munro; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS41MzY4MzcwOUFFRUU3QzEx @{title=PSScriptAnalyzer (PSSA) VS-code integration & customization... by Christoph Bergmeister; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4yQjZFRkExQjFGODk3RUFD @{title=PowerShell Remoting Internals by Paul Higinbotham; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4yQUJFNUVCMzVDNjcxRTlF @{title=Parselmouth - bending the PowerShell language by Mathias Jessen; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS40QzRDOEU0QUYwNUIxN0M1 @{title=Writing Compiled PowerShell Cmdlets by Thomas Rayner; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS41RTNBREYwMkI5QzU3RkY2 @{title="Piping" data between packaged scripts by Paul DeArment Jr; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5ENjI1QUI0MDI5NEQzODFE @{title=Jenkins - User Interface for your Powershell tasks by Kirill Kravtsov; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS44QzVGQUU2QjE2NDgxM0M4 @{title=Finding Performance Bottlenecks with PowerShell by Mike F. Robbins; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4xMzgwMzBERjQ4NjEzNUE5 @{title=Chocolatey For the Organizations: Easily Manage Software by Rob Reynolds; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4zMEQ1MEIyRTFGNzhDQzFB @{title=PowerShell Error and Event Collection at Scale by Dakota Clark; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS42Qzk5MkEzQjVFQjYwRDA4 @{title=Better Ops Together: Practical PowerShell Pair Programming Patterns and Practices... by Mark Kraus; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS41NTZEOThBNThFOUVGQkVB @{title=Containers - You Better Get on Board! by Anthony Nocentino; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS43NERCMDIzQzFBMERCMEE3 @{title=Using Visual Studio Code as Your Default PowerShell Editor by Tyler Leonhardt; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5GNjAwN0Y0QTFGOTVDMEMy @{title=Unleash your PowerShell with AWS Lambda and Serverless Computing by Andrew Pearce; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5CQkEwRDA0MDkwNUM2MDY1 @{title=Microsoft Azure Policy Guest Configuration by Michael Greene; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4wNEU1MTI4NkZEMzVBN0JF @{title=Wardley Maps Saved The Day - How Stack Overflow Enterprise automated all the things... by Chris Hunt; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4wMTYxQzVBRDI1NEVDQUZE @{title=Continuously deploying SQL code using Powershell by Kirill Kravtsov; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4zMUEyMkQwOTk0NTg4MDgw @{title=It’s PowerShell In the Cloud – Welcome to Azure Cloud Shell by Michael Bender; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS42QzdBMzlBQzQzRjQ0QkQy @{title=Demystifying Microsoft's Cloud Automation products by Jaap Brasser; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4wRjhFM0MxMTU1MEUzQ0VB @{title=Using PowerShell in a Cross Platform World presented by Bill Hurt written by James Pogran; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5CNTZFOTNGQzZEODg1RUQx @{title=How to become a SHiPS wright - Building with SHiPS by Glenn Sarti; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5CNTcxMDQ0NThBNzMxODYz @{title=PSCache: simple strategies for magnificent performance by Mathias Jessen; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS5ERkUyQTM0MzEwQjZCMTY5 @{title=Using Pester & ScriptAnalyzer for Detecting Obfuscated PowerShell by Daniel Bohannon; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS4xM0YyM0RDNDE4REQ1NDA0 @{title=Malicious Payloads vs Deep Visibility: A PowerShell Story by Daniel Bohannon; resourceId=}
UExmZUE4a0lzN0NvY2lyMS1UdVNOM21PbmozcXp5UlNoQS42MjYzMTMyQjA0QURCN0JF @{title=Publishing and Managing Modules in an Internal Repository by Kevin Marquette; resourceId=}

```

</div></details>

## `F#`の`unfold`

`F#`の場合は様々なコレクション型が定義されており、それぞれに`unfold`を含む操作用の関数が用意されています。
例えば、`seq<'T>`用の`unfold`の定義は以下の通りです。

```fsharp
Seq.unfold generator state

parameters :
  // 現在の状態を取り込み、リストの次の要素と次の状態の値のオプションタプルを返す関数
  generator : 'State -> ('T * 'State) option
  
  // 初期値
  state : 'State

// 生成されたリスト
Returns: seq<'T>
```

### 読み方

- `Tin -> Tout`...「`Tin`を引数にとり、`Tout`を返す関数」を定義する`F#`の組み込み型を意味します。この関数型はPowershellにおける`[scriptblock]`に近い存在です。
- `(T1 * T2)`…タプルの表記法です。
- `T option`…`F#`固有のジェネリック方の表記法です。`C#`と同様に`option<T>`という表記も可能です。通常、この表機能は`F#`の組み込み型でのみ使用されます。

比較するとこうなります。
なぜ`F#`ではこんな記法を採用しているのか伝わってくる気がします。

| `F#` | `Powershel` | `C#` |
|:-:|:-:|:-:|
| `Tin -> Tout`  | `[fun[Tin,Tout]]`  | `func<Tin,Tout>`  |
|  `(T1 * T2)` |  `[Tuple[T1,T2]]` |  `Tuple<T1,T2>` |
| `T option`  |  `[Option[T]` | `option<T>`  |
|`'State -> ('T * 'State) option`|`[fun[TState,[option[Tuple[T, TState]]]]]`|`<fun<TState,<option<Tuple<T, TState>>>>>`|

関数定義をPowershell風に書くとこんな感じです。

```powershell:unfoldの定義(Powershell風)
NAME
    Seq.unfold

SYNTAX
    Seq.unfold[-generator <fun<TState,<option<Tuple<T, TState>>>>>] [-state <TState>]

OUTPUTS
    seq<T>
```

以下は`generator`の一例です。

```powershell:generator(Powershell版)
# 20になるまで1ずつ増えるシーケンスを生成する場合
# 'T … int
# 'State … int
{
    [OutputType([Tuple[int, int]])]
    param (
        [int]$state
    )
    if ($state -gt 20) {
        return $null
    }
    else {
        return [Tuple[int, int]]::new($state, $state + 1)
    }
}
```

#### 余談:楽な書き方

可読性、安全性を投げ捨て、楽な書き方をするとこうなります。
自動変数`$args`をうまく使うと`param`ブロックは不要です。
条件を満たさないと`$null`→特定条件でのみ値を出力と考えます。

```powershell:generator(Powershell版_手抜き)
{if ($args[0] -le 20) { $args[0], ($args[0] + 1) }}
```

コンソール上で書く場合など、書き捨て用の場合は大体こっちの書き方です。
コードを残す場合は履歴をコピペしてvscode等で良い感じに整形してます。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/3c3e2964-540f-db9e-d40a-5e007ec55a27.png)

### 参考:公式ドキュメント(F#)

https://docs.microsoft.com/en-us/dotnet/fsharp/language-reference/sequences#creating-sequences

※日本語ドキュメントの翻訳が微妙なのでDeepLの力を借りました。

> [`Seq.unfold`](https://fsharp.github.io/fsharp-core-docs/reference/fsharp-collections-seqmodule.html#unfold)は、`state`を受け取り、それを変換してシーケンスの後続の各要素を生成するコンピュテーション関数からシーケンスを生成します。`state`とは、各要素を計算するために使われる値であり、各要素が計算されるたびに変化します。`Seq.unfold`の第2引数は、シーケンスの開始に使われる初期値です。`Seq.unfold`では、`state`に`Option`型を使用しており、`None`値を返すことでシーケンスを終了させることができます。次のコードでは、unfold操作で生成されるシーケンスの例として、`seq1`と`fib`を示しています。1つ目の`seq1`は、20までの数字を並べた単純なシーケンスです。2つ目の`fib`は、`unfold`を使ってフィボナッチ数列を計算しています。フィボナッチ数列の各要素は、前の2つのフィボナッチ数の合計であるため、`state`値は、数列の前の2つの数からなるタプルとなります。初期値は(1,1)で、数列の最初の2つの数値です。

```fsharp
let seq1 =
    0 // Initial state
    |> Seq.unfold (fun state ->
        if (state > 20) then
            None
        else
            Some(state, state + 1))

printfn "The sequence seq1 contains numbers from 0 to 20."

for x in seq1 do
    printf "%d " x

let fib =
    (1, 1) // Initial state
    |> Seq.unfold (fun state ->
        if (snd state > 1000) then
            None
        else
            Some(fst state + snd state, (snd state, fst state + snd state)))

printfn "\nThe sequence fib contains Fibonacci numbers."
for x in fib do printf "%d " x
```

> 出力は以下の通りです。

```console
The sequence seq1 contains numbers from 0 to 20.

0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20

The sequence fib contains Fibonacci numbers.

2 3 5 8 13 21 34 55 89 144 233 377 610 987 1597
```

## 実装のこだわり

利用時の書きやすさを実現するため、以下の目標がありました。

1. `Filter`で実装
1. `state`には`$_`でアクセス可能にする

2番目が特に大事でした。文字数が6もある`$arg[0]`から解放されたかったんです。
これは辛い……

```powershell
3..5 | unfold {
    $arg[0] -le 5 ? $arg[0], ($arg[0] + 1) : 'end.',$null
}
```

これなら気になりません。

```powershell
3..5 | unfold {
    $_ -le 5 ? $_, ($_ + 1) : 'end.',$null
}
```

こういう場合に使えるのが`ScriptBlock.InvokeWithContext`メソッドです。

https://docs.microsoft.com/en-us/dotnet/api/system.management.automation.scriptblock.invokewithcontext?view=powershellsdk-7.0.0#System_Management_Automation_ScriptBlock_InvokeWithContext_System_Collections_IDictionary_System_Collections_Generic_List_System_Management_Automation_PSVariable__System_Object___

> `InvokeWithContext(IDictionary, List<PSVariable>, Object[])`
> スクリプトブロックのスコープで定義されるローカル関数と変数のセットの形式で、追加のコンテキストでスクリプトブロックを呼び出すことができるメソッド。変数のリストには、`$input`の変数、`$_`、`$this`が含まれます。‎
>
> この関数のオーバーロードはハッシュテーブルを受け取り、必要な辞書に変換して、PowerShell スクリプト内から API を使いやすくします。

使用例です。
`[scriptblock]`を引数にとる関数で役立ちそうです。

```powershell
$sb={
    $bar
    $bar=100
    $bar
    Test-Func
    $_
    $this
    $input
    $args
}
"& `$sb"
& $sb
"`r`n`$sb.InvokeWithContext"
$sb.InvokeWithContext(
    # ハッシュテーブルで追加の関数定義
    @{'Test-Func'= {10000}},
    # 自動変数$_, $this, $inputを含めて変数を追加で定義
    @(
        [psvariable]::new('bar', 'bar')
        [psvariable]::new('_', '_')
        [psvariable]::new('this', 'this')
        [psvariable]::new('input', 'input')
    ),
    # 関数の引数($args)を指定
    (20,40)
)
```

結果

```console:result
& $sb
100
Test-Func: 
Line |
   5 |      Test-Func
     |      ~~~~~~~~~
     | The term 'Test-Func' is not recognized as a name of a cmdlet, function, script file, or executable program.
Check the spelling of the name, or if a path was included, verify that the path is correct and try again.

$sb.InvokeWithContext
bar
100
10000
_
this
input
20
40
```

`unfold`では`$_`の内容を制御するために使用しています。

## 終わりに

もし機能を足すとしたら生成する要素数の数、インデックスを指定するパラメータの実装だと思います。
ただし、下手に機能を足すぐらいなら`F#`の組み込み型のラッパーモジュールを作るほうが早いかもしれません。

## 蛇足

`Powershell`で実現したい`F#`の機能は他にもあります。
例えばカリー化です。

https://gist.github.com/SilkyFowl/2b10789952b3409061c29e1dfa218a95

ただのカリー化ならいくらでも記事がありますが、引数が３つ以上の場合の対応や使い勝手の良さをASTを用いて実現しようとか考えてました。
もし進展があったら記事にするかもしれません。
