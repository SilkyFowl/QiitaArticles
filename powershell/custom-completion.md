<!--
title:   Powershellでカスタム補完機能（1）～「第1引数でMsDocsを検索して検索結果を第2引数に補完表示する関数」を作りたいので色々準備～
tags:    PSReadLine,PowerShell,starship
id:      6cb36ce997868b6b2b69
private: false
-->
# これは何？
Powershellの補完機能についての備忘録です。

※**MSDocs自体の記事検索REST APIのレファレンスを探しています。ご存じでしたらコメントしていただけると幸いです。**

## ゴール
表題どおりです。Powershellからググりたかったんです……

<script id="asciicast-359853" src="https://asciinema.org/a/359853.js" async></script>

ソースはこちら。最低限の動作は出来ましたが色々リファクタリングしたい感じです......

```posh
using namespace System.Management.Automation
using namespace System.Collections
using namespace System.Collections.Generic

class PSModuleApiBrowserArgumentCompleter : IArgumentCompleter {

    hidden [string]$root = 'https://docs.microsoft.com/powershell/module/'

    [IEnumerable[CompletionResult]] CompleteArgument(
        [string] $CommandName,
        [string] $ParameterName,
        [string] $WordToComplete,
        [Language.CommandAst] $CommandAst,
        [IDictionary] $FakeBoundParameters
    ) {
        # TODO: ステートフルにする
        $results = @{
            Body = @{
                'api-version' = 0.2
                'search'      = $FakeBoundParameters['search']
                locale        = 'en-us'
                '$skip'       = 0
                '$top'        = 25
            }
        } | Start-ThreadJob {
            (Invoke-RestMethod https://docs.microsoft.com/api/apibrowser/powershell/search -Body $input.Body -TimeoutSec 5).results
        } | Receive-Job -Wait -AutoRemoveJob

        $CompletionResults = $results.foreach{
            [CompletionResult]::new(
                ($_.url -replace $this.root, ''),
                $_.displayName ,
                [CompletionResultType]::ParameterValue ,
                $_.description ?? ' '
            )
        }

        return [CompletionResult[]]$CompletionResults
    }
}

function Show-PowershellModuleBrowser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$search,
        [Parameter(Mandatory = $true, Position = 1)]
        [ArgumentCompleter([PSModuleApiBrowserArgumentCompleter])]
        $Page
    )
    $URL = "https://docs.microsoft.com/powershell/module/$Page"
    Start-Process $URL
}

if (-not(Test-Path Alias:\shPwshMdlBrw)) {
    Set-Alias -Name shPwshMdlBrw -Value Show-PowershellModuleBrowser
}


class DotnetApiBrowserArgumentCompleter : IArgumentCompleter {

    hidden [string]$root = 'https://docs.microsoft.com/dotnet/api/'

    [IEnumerable[CompletionResult]] CompleteArgument(
        [string] $CommandName,
        [string] $ParameterName,
        [string] $WordToComplete,
        [Language.CommandAst] $CommandAst,
        [IDictionary] $FakeBoundParameters
    ) {
        # TODO: ステートフルにする
        $results = @{
            Body = @{
                'api-version' = 0.2
                'search'      = $FakeBoundParameters['search']
                locale        = 'ja-jp'
                '$skip'       = 0
                '$top'        = 25
            }
        } | Start-ThreadJob {
            (Invoke-RestMethod https://docs.microsoft.com/api/apibrowser/dotnet/search -Body $input.Body -TimeoutSec 5).results
        } | Receive-Job -Wait -AutoRemoveJob

        $CompletionResults = $results.foreach{
            [CompletionResult]::new(
                ($_.url -replace $this.root, ''),
                $_.displayName ,
                [CompletionResultType]::ParameterValue ,
                $_.description ?? ' '
            )
        }

        return [CompletionResult[]]$CompletionResults
    }
}

function Show-DotnetApiBrowser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$search,
        [Parameter(Mandatory = $true, Position = 1)]
        [ArgumentCompleter([DotnetApiBrowserArgumentCompleter])]
        $Page
    )

    $URL = "https://docs.microsoft.com/dotnet/api/$Page"
    Start-Process $URL
}

if (-not(Test-Path Alias:\shDotnetApiBrw)) {
    Set-Alias -Name shDotnetApiBrw -Value Show-DotnetApiBrowser
}
```

以下解説。長くなったので数回に分けます。
**今回は大体が準備になってしまったので入力補完の踏み込んだ解説は次回以降となります。**

## 環境
>新しいクロスプラットフォームの PowerShell をお試しください https://aka.ms/pscore6

```powershell
~
❯ $PSVersionTable

Name                           Value
----                           -----
PSVersion                      7.1.0-preview.7
PSEdition                      Core
GitCommitId                    7.1.0-preview.7
OS                             Microsoft Windows 10.0.19041
Platform                       Win32NT
PSCompatibleVersions           {1.0, 2.0, 3.0, 4.0…}
PSRemotingProtocolVersion      2.3
SerializationVersion           1.1.0.1
WSManStackVersion              3.0
```
### Scoop

基本的にCUI系のツールはScoopで揃えてます

```shell-session
❯ scoop list
Installed apps:

  7zip 19.00
  archwsl 20.4.3.0 [extras]
  aria2 1.35.0-1
  CascadiaCode-NF 2.1.0 [nerd-fonts]
  conemu 20.07.13 [extras]
  curl 7.72.0_4
  dark 3.11.2
  docker-machine 0.16.2
  fontforge 20200314 [extras]
  git 2.28.0.windows.1
  gsudo 0.7.2
  innounp 0.49
  lessmsi 1.6.91
  neovim 0.4.4
  nodejs 14.10.1
  OpenSSH 8.2p1-1
  pshazz 0.2020.05.23
  pwsh-beta 7.1.0-preview.7 [versions]
  python 3.8.5
  python27 2.7.18 [versions]
  ruby 2.7.1-1
  rustup 1.22.1
  Selenium 3.141.59
  starship 0.44.0
  Sysinternals December.18.2019 [extras]
  unar 1.8.1
  vim 8.2
  wixtoolset 3.11.2
```

### `$PROFILE`

最近はコンソールで色々するのでそれなりに弄ってます。

```posh:profile.ps1
# PsReadLine 設定
. $PSScriptRoot/Setting_PsReadLine.ps1

# 補完の設定
. $PSScriptRoot/Setting_Completion.ps1

# StarShipの起動
Invoke-Expression (&starship init powershell)
$starshipPrompt = (Get-Item Function:\prompt).ScriptBlock

function prompt {
    # 出力結果
    $out=[System.Text.StringBuilder]::new()

    # デバッグ時の出力
    if (Test-Path variable:/PSDebugContext) {
        $out.AppendFormat("`e[38;5;202m{0}`e[0m","[DBG]: ")  > $null
    }

    # StarShip
    $out.Append((& $starshipPrompt)) > $null

    # 出力
    $out.ToString()
}
```

`PsReadLine`は是非カスタマイズしましょう。使い勝手が大きく変わります 

>[【PowerShell】PsReadLine 設定のススメ](https://qiita.com/AWtnb/items/5551fcc762ed2ad92a81)

```posh:Setting_PsReadLine.ps1
using namespace Microsoft.PowerShell

# https://qiita.com/AWtnb/items/5551fcc762ed2ad92a81#履歴管理
Set-PSReadlineOption -HistoryNoDuplicates

# https://qiita.com/AWtnb/items/5551fcc762ed2ad92a81#単語区切り
Set-PSReadLineOption -WordDelimiters ";:,.[]{}()/\|^&*-=+'`" !?@#$%&_<>「」（）『』『』［］、，。：；／`u{2015}`u{2013}`u{2014}"
Set-PSReadlineOption -AddToHistoryHandler {
    param ($command)
    switch -regex ($command) {
        "SKIPHISTORY" {return $false}
        "^[a-z]$" {return $false}
        "exit" {return $false}
    }
    return $true
}


Set-PSReadLineKeyHandler -Key "`"","'" -BriefDescription "smartQuotation" -LongDescription "Put quotation marks and move the cursor between them or put marks around the selection" -ScriptBlock {
    param($key, $arg)
    $mark = $key.KeyChar

    $selectionStart = $null
    $selectionLength = $null
    [PSConsoleReadLine]::GetSelectionState([ref]$selectionStart, [ref]$selectionLength)
    $line = $null
    $cursor = $null
    [PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)

    if ($selectionStart -ne -1) {
        [PSConsoleReadLine]::Replace($selectionStart, $selectionLength, $mark + $line.SubString($selectionStart, $selectionLength) + $mark)
        [PSConsoleReadLine]::SetCursorPosition($selectionStart + $selectionLength + 2)
        return
    }

    if ($line[$cursor] -eq $mark) {
        [PSConsoleReadLine]::SetCursorPosition($cursor + 1)
        return
    }

    $nMark = [regex]::Matches($line, $mark).Count
    if ($nMark % 2 -eq 1) {
        [PSConsoleReadLine]::Insert($mark)
    }
    else {
        [PSConsoleReadLine]::Insert($mark + $mark)
        [PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)
        [PSConsoleReadLine]::SetCursorPosition($cursor - 1)
    }
}

Set-PSReadLineKeyHandler -Key "alt+w" -BriefDescription "WrapLineByParenthesis" -LongDescription "Wrap the entire line and move the cursor after the closing parenthesis or wrap selected string" -ScriptBlock {
    $selectionStart = $null
    $selectionLength = $null
    [PSConsoleReadLine]::GetSelectionState([ref]$selectionStart, [ref]$selectionLength)
    $line = $null
    $cursor = $null
    [PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)
    if ($selectionStart -ne -1) {
        [PSConsoleReadLine]::Replace($selectionStart, $selectionLength, "(" + $line.SubString($selectionStart, $selectionLength) + ")")
        [PSConsoleReadLine]::SetCursorPosition($selectionStart + $selectionLength + 2)
    }
    else {
        [PSConsoleReadLine]::Replace(0, $line.Length, '(' + $line + ')')
        [PSConsoleReadLine]::EndOfLine()
    }
}

Remove-PSReadlineKeyHandler "tab"
Set-PSReadLineKeyHandler -Key "tab" -BriefDescription "smartNextCompletion" -LongDescription "insert closing parenthesis in forward completion of method" -ScriptBlock {
    [PSConsoleReadLine]::TabCompleteNext()
    $line = $null
    $cursor = $null
    [PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)

    if ($line[($cursor - 1)] -eq "(") {
        if ($line[$cursor] -ne ")") {
            [PSConsoleReadLine]::Insert(")")
            [PSConsoleReadLine]::BackwardChar()
        }
    }
}


Remove-PSReadlineKeyHandler "shift+tab"
Set-PSReadLineKeyHandler -Key "shift+tab" -BriefDescription "smartPreviousCompletion" -LongDescription "insert closing parenthesis in backward completion of method" -ScriptBlock {
    [PSConsoleReadLine]::TabCompletePrevious()
    $line = $null
    $cursor = $null
    [PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)

    if ($line[($cursor - 1)] -eq "(") {
        if ($line[$cursor] -ne ")") {
            [PSConsoleReadLine]::Insert(")")
            [PSConsoleReadLine]::BackwardChar()
        }
    }
}
#endregion


# プロファイルの再読み込み
Set-PSReadLineKeyHandler -Key "alt+r" -BriefDescription "reloadPROFILE" -LongDescription "reloadPROFILE" -ScriptBlock {
    [PSConsoleReadLine]::RevertLine()
    [PSConsoleReadLine]::Insert('<#SKIPHISTORY#> . $PROFILE')
    [PSConsoleReadLine]::AcceptLine()
}

# 直前に使用した変数を利用する
Set-PSReadLineKeyHandler -Key "alt+a" -BriefDescription "yankLastArgAsVariable" -LongDescription "yankLastArgAsVariable" -ScriptBlock {
    [PSConsoleReadLine]::Insert("$")
    [PSConsoleReadLine]::YankLastArg()
    $line = $null
    $cursor = $null
    [PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)
    if ($line -match '\$\$') {
        $newLine = $line -replace '\$\$', "$"
        [PSConsoleReadLine]::Replace(0, $line.Length, $newLine)
    }
}

# クリップボード内容を変数に格納する
Set-PSReadLineKeyHandler -Key "ctrl+V" -BriefDescription "setClipString" -LongDescription "setClipString" -ScriptBlock {
    $command = "<#SKIPHISTORY#> get-clipboard | sv CLIPPING"
    [PSConsoleReadLine]::RevertLine()
    [PSConsoleReadLine]::Insert($command)
    [PSConsoleReadLine]::AddToHistory('$CLIPPING ')
    [PSConsoleReadLine]::AcceptLine()
}

# Predictation関連
Set-PSReadLineOption -PredictionSource History
Set-PSReadLineOption -Colors @{ Prediction = [System.ConsoleColor]::DarkBlue}

Set-PSReadLineKeyHandler -Key "Ctrl+d" -Function MenuComplete
Set-PSReadLineKeyHandler -Key "Ctrl+f" -Function ForwardWord
Set-PSReadLineKeyHandler -Key "Ctrl+b" -Function BackwardWord
Set-PSReadLineKeyHandler -Key "Ctrl+z" -Function Undo
Set-PSReadLineKeyHandler -Key UpArrow -Function HistorySearchBackward
Set-PSReadLineKeyHandler -Key DownArrow -Function HistorySearchForward
```

`posh:Setting_Completion.ps1
# dotnet CLI
Register-ArgumentCompleter -Native -CommandName "dotnet" -ScriptBlock {
    param($commandName, $wordToComplete, $cursorPosition)

    switch (dotnet complete --position $cursorPosition "$wordToComplete") {
        Default {[CompletionResult]::new($_, $_, 'ParameterValue', $_)}
    }
}

# Docker
import-Module DockerCompletion

# StarShip
Invoke-Expression (@(starship completions powershell) -join "`n")
`

### 余談：StarShip x Powershellの小技

StarShipには[Shellコマンドを使って表示をカスタマイズ出来る機能](https://starship.rs/ja-JP/config/#custom-commands)があります。
ただ、プロンプトの度に新規プロセスでシェルコマンドが実行されるという仕組みがPowershellと相性が良いとはいえません。
実行オブション`-Nop`をつけるとマシにはなりますが正直苦しい感じの遅さです。
そんなとき、代わりの手段があるので紹介します。

#### `Starship`の処理を変数化する

そもそも、`&starship init powershell`では一体何をしているのでしょう？確かめてみます。

```posh
~
❯ starship init powershell
Invoke-Expression (@(&"C:\Users\user\scoop\apps\starship\current\starship.exe" init powershell --print-full-init) -join "`n")
```
入れ子になっているようです。更に展開してみます。

```posh
~
❯ starship init powershell --print-full-init
#!/usr/bin/env pwsh

# Starship assumes UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
function global:prompt {
    $out = $null
    # @ makes sure the result is an array even if single or no values are returned
    $jobs = @(Get-Job | Where-Object { $_.State -eq 'Running' }).Count

    $env:PWD = $PWD
    $current_directory = (Convert-Path $PWD)

    if ($lastCmd = Get-History -Count 1) {
        $duration = [math]::Round(($lastCmd.EndExecutionTime - $lastCmd.StartExecutionTime).TotalMilliseconds)
        # & ensures the path is interpreted as something to execute
        $out = @(&"C:\Users\user\scoop\apps\starship\current\starship.exe" prompt "--path=$current_directory" --status=$lastexitcode --jobs=$jobs --cmd-duration=$duration)
    } else {
        $out = @(&"C:\Users\user\scoop\apps\starship\current\starship.exe" prompt "--path=$current_directory" --status=$lastexitcode --jobs=$jobs)
    }

    # Convert stdout (array of lines) to expected return type string
    # `n is an escaped newline
    $out -join "`n"
}

$ENV:STARSHIP_SHELL = "powershell"

```

見つけました。`function global:prompt`がプロンプト表示を司る関数です。
この処理が終わった後に以下の手順でStarShipの表示に手を加えることが出来ます。

#####`Invoke-Expression (&starship init powershell)`の後で更新された`prompt`の中身を任意の変数へキャプチャする

```posh
$starshipPrompt = (Get-Item Function:\prompt).ScriptBlock
```

※Powershellはファイルシステム以外にも関数や変数、レジストリでも`ls`や`cd`が使えます。


##### キャプチャした変数を材料に、新しい`prompt`を定義する

```posh
function prompt {

    # 出力結果
    $out=[System.Text.StringBuilder]::new()

    # デバッグ時の出力
    if (Test-Path variable:/PSDebugContext) {
        $out.AppendFormat("`e[38;5;202m{0}`e[0m","[DBG]: ")  > $null
    }

    # StarShip
    $out.Append((& $starshipPrompt)) > $null

    # 出力
    $out.ToString()
}
```

#### 環境変数を利用する ～ANSI エスケープシーケンスを添えて～

[StarShipの環境変数を表示する機能](https://starship.rs/ja-JP/config/#%E7%92%B0%E5%A2%83%E5%A4%89%E6%95%B0)はポテンシャル高めです。

##### 動作原理

```toml:~/.config/starship.toml
[env_var]
symbol = ""
variable = "Ansi"
style = "bold"
```
`posh
~
❯ Test-Path Env:\Ansi
False

~
❯ $env:Ansi='値をセットしました。'

~ with 値をセットしました。
❯ Remove-Item Env:\Ansi

~
❯
`

まあ、こんな感じでしょう。ところで、StarShip x Powershellの場合は**`[string]`にキャスト出来るなら何を入れても良いようです。**

```posh:コマンドレットの結果を突っ込めちゃいました
❯ $env:Ansi=@(ps pwsh | oss) -join "`n"

~ with
 NPM(K)    PM(M)      WS(M)     CPU(s)      Id  SI ProcessName
 ------    -----      -----     ------      --  -- -----------
     63    96.28      49.64       6.67    1568   1 pwsh
     72   110.27     143.84       9.05    3728   1 pwsh
     11    11.82       9.96       0.06   24044   1 pwsh
    122   205.98     153.20      33.64   27080   1 pwsh

❯
```

`posh:ファイル経由でEmoji
❯ cat C:\Users\user\AppData\Local\Temp\pwshTest.ps1
-join ("鬱です🥺", "鬱😥", "う", "！", "😭"|%{}{$_*8}{([char[]]"SocialDistance" -join "　")+'（辛く苦しい社会から離脱）'})

~ with
 NPM(K)    PM(M)      WS(M)     CPU(s)      Id  SI ProcessName
 ------    -----      -----     ------      --  -- -----------
     63    96.31      48.65       6.69    1568   1 pwsh
     75   121.88     159.43       9.62    3728   1 pwsh
     11    11.82       9.94       0.06   24044   1 pwsh
    123   207.61     168.05      35.73   27080   1 pwsh

❯ & C:\Users\user\AppData\Local\Temp\pwshTest.ps1
鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱😥鬱😥鬱😥鬱😥鬱😥鬱😥鬱😥鬱😥うううううううう！！！！！！！！😭😭😭😭😭😭😭😭S　o　c　i　a　l　D　i　s　t　a　n　c　e（辛く苦しい社会から離脱）

~ with
 NPM(K)    PM(M)      WS(M)     CPU(s)      Id  SI ProcessName
 ------    -----      -----     ------      --  -- -----------
     63    96.31      48.65       6.69    1568   1 pwsh
     75   121.88     159.43       9.62    3728   1 pwsh
     11    11.82       9.94       0.06   24044   1 pwsh
    123   207.61     168.05      35.73   27080   1 pwsh

❯ $env:Ansi = & C:\Users\user\AppData\Local\Temp\pwshTest.ps1

~ with 鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱です🥺鬱😥鬱😥鬱😥鬱😥鬱😥鬱😥鬱😥鬱😥うううううううう！！！！！！！！😭😭😭😭😭😭😭😭S　o　c　i　a　l　D　i　s　t　a　n　c　e（辛く苦しい社会から離脱）
❯

```
[※シェル芸Botをよろしくお願いします。](https://twitter.com/minyoruminyon?s=20)

そして、数年前からWindows10でもANSIエスケープシーケンスを使えます。
>[Console Virtual Terminal Sequences - Windows Console | Microsoft Docs] (https://docs.microsoft.com/en-us/windows/console/console-virtual-terminal-sequences)

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/4c96a829-e9d5-ceef-5c6e-fa7089a4604d.png)

この性質を使うと色々出来そうです。

## 次回予告-補完の仕組み-

モジュールの支援なしにPowershellでTab補完をカスタマイズしたい場合、自動関数`TabExpansion2`をカスタマイズします。
仕組みを理解するにはデバッグをして実際の挙動を調べるのが手っ取り早いです。
そういうわけで、**`TabExpansion2`をデバッグしてみましょう。**

### プロンプト周りのデバッグ方法

自分の環境の場合、プロンプト周りのデバッグはコンソール上で行った方が色々と簡単でした。
>[about_Debuggers - PowerShell | Microsoft Docs ](https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_debuggers?view=powershell-7)

### TabExpansion2

### [CommandCompletion]::CompleteInput

```posh
❯ $result=[System.Management.Automation.CommandCompletion]::CompleteInput(
>>                                   <#inputScript#>  $inputScript,
>>                                   <#cursorColumn#> $cursorColumn,
>>                                   <#options#>      $options)

❯ $result | fl
[DBG]:
~
CurrentMatchIndex : -1
ReplacementIndex  : 11
ReplacementLength : 0
CompletionMatches : {System.Management.Automation.CompletionResult, System.Management.Automation.CompletionResult
                    , System.Management.Automation.CompletionResult, System.Management.Automation.CompletionResul
                    t…}
❯ $result.CompletionMatches

CompletionText             ListItemText               ResultType ToolTip
--------------             ------------               ---------- -------
__NounName                 __NounName                   Property string __NounName=Process
BaseAddress                BaseAddress                  Property System.IntPtr BaseAddress { get; }
B
```

`posh
~
❯ ps -Hit Command breakpoint on 'TabExpansion2'

At line:37 char:23
+                       {
+                       ~
[DBG]:
~
❯ Get-PSCallStack -ov cs

Command               Arguments                                    Location
-------               ---------                                    --------
TabExpansion2         {inputScript=ps -, cursorColumn=4, options=} <No file>
                      {}                                           Setting_PsReadLine.ps1: line 119
PSConsoleHostReadLine {}                                           PSReadLine.psm1: line 4

[DBG]:
~
❯ $cs[0].Arguments
{inputScript=ps -, cursorColumn=4, options=}
[DBG]:
~
❯ $ast
[DBG]:
~
❯ $PSCmdlet

Host                 : System.Management.Automation.Internal.Host.InternalHost
SessionState         : System.Management.Automation.SessionState
Events               : System.Management.Automation.PSLocalEventManager
JobRepository        : System.Management.Automation.JobRepository
JobManager           : System.Management.Automation.JobManager
InvokeProvider       : System.Management.Automation.ProviderIntrinsics
ParameterSetName     : ScriptInputSet
MyInvocation         : System.Management.Automation.InvocationInfo
PagingParameters     :
InvokeCommand        : System.Management.Automation.CommandInvocationIntrinsics
Stopping             : False
CommandRuntime       : TabExpansion2
CurrentPSTransaction :
CommandOrigin        : Internal


[DBG]:
~
❯ $result=[System.Management.Automation.CommandCompletion]::CompleteInput(
>>                                   <#inputScript#>  $inputScript,
>>                                   <#cursorColumn#> $cursorColumn,
>>                                   <#options#>      $options)
[DBG]:
~
❯ $result

CurrentMatchIndex ReplacementIndex ReplacementLength CompletionMatches
----------------- ---------------- ----------------- -----------------
               -1                3                 1 {System.Management.Automation.CompletionResult, System.Mana…

[DBG]:
~
❯ $result.CompletionMatches

CompletionText       ListItemText           ResultType ToolTip
--------------       ------------           ---------- -------
-Name                Name                ParameterName [string[]] Name
-Id                  Id                  ParameterName [int[]] Id
-InputObject         InputObject         ParameterName [Process[]] InputObject
-IncludeUserName     IncludeUserName     ParameterName [switch] IncludeUserName
-Module              Module              ParameterName [switch] Module
-FileVersionInfo     FileVersionInfo     ParameterName [switch] FileVersionInfo
-Verbose             Verbose             ParameterName [switch] Verbose
-Debug               Debug               ParameterName [switch] Debug
-ErrorAction         ErrorAction         ParameterName [ActionPreference] ErrorAction
-WarningAction       WarningAction       ParameterName [ActionPreference] WarningAction
-InformationAction   InformationAction   ParameterName [ActionPreference] InformationAction
-ErrorVariable       ErrorVariable       ParameterName [string] ErrorVariable
-WarningVariable     WarningVariable     ParameterName [string] WarningVariable
-InformationVariable InformationVariable ParameterName [string] InformationVariable
-OutVariable         OutVariable         ParameterName [string] OutVariable
-OutBuffer           OutBuffer           ParameterName [int] OutBuffer
-PipelineVariable    PipelineVariable    ParameterName [string] PipelineVariable


[DBG]:
~
❯ $result.GetType()

IsPublic IsSerial Name                                     BaseType
-------- -------- ----                                     --------
True     False    CommandCompletion                        System.Object

`

次の記事ではデバッグ方法の詳細、補完入力の仕組みの解析、その結果の応用（冒頭のスクリプトについて解説）を行いたいと思います。