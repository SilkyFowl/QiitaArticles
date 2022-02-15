<!--
title:   未サポートだけどWindowsコンテナを開発コンテナとして使おうとした
tags:    Docker,VSCode,Windows
id:      e0e6c50cea8db5bd8ca5
private: false
-->
# これは何？

- 残念ながらVScodeでWindowsコンテナを開発コンテナとして使うことはできません
- だけどssh-remoteで別のWindowsで開発はできます
- だったら**windowsコンテナへssh-remoteできれば実質開発コンテナでは？**

そんな感じでやれるところまでやった感じがするのでまとめてみました。
※イメージサイズ5GBの超重量コンテナならうまくいきましたが、数百MBの実用レベルコンテナでは上手くいきませんでした。
※現状、Nanoserverを開発環境として使いたいならコンテナにNeovimを仕込むのが現実的だと思います。

## Motivation

**Nanoserverベースの開発コンテナほしい……**
この圧倒的軽さを開発で使いたい……

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/f635c06e-ebd9-78c3-89d3-f9abf58c75f8.png)

色々入れても1GB弱で済む、小さい！！

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/4e4ea057-c479-ca10-2af4-fb7c0510d8a4.png)

Process分離モードで起動することで省メモリでいろいろ嬉しいんです

```powershell
docker run -it  -d --isolation process -p 11000:22  --name ssh_nano ssh:nanoserver-20H2pwsh725
```

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/e3575be8-72d6-ac80-d43c-e54de937ffc8.png)

という訳で足掻きました。

## 結果

- Windowsコンテナのうち、5GB弱のServerCoreならssh経由でRemote可能。試してないけどデバッグもできそう。
- 数百MBのNanoServerはRemote開発は難しい。
- いろいろ仕込むとファイルを開いて編集すること自体は可能になった
- だた、Terminalを起動できない。（後述）

## リポジトリ

https://github.com/SilkyFowl/MyWindowsContainer

- 検証とか試行錯誤用にいろいろ突っ込んでます

## ServerCore

パスワード付きアカウントを作ってsshを仕込むだけで良い感じです

## NanoServer

Remote-sshが想定しているWindowsの機能がいくつか削ぎ落されてるのでsshログインできるようにしただけでは動きません。少し工夫が必要です。

### (突破)NanoServerにはWindows Powershellがインストールされてない

無いんです……
解決方法は2通りあります。

1. Remote-sshのコードからPowershell.exeを呼び出すコードを探し出してすべて修正する
1. Powershell.exe→pwsh.exeのシンボリックリンクを作成する

前者はトランスコンパイル後のJavaScriptコードを弄ることになるので色々大変です。
後者のほうはDockerfileに数行追記するだけで楽でした。

```Dockerfile
# escape=`

RUN mkdir C:\Windows\System32\WindowsPowerShell; `
    ni -it SymbolicLink -Path C:\Windows\System32\WindowsPowerShell\v1.0 -Target $env:ProgramFiles\PowerShell\latest; `
    cd $pshome; `
    ni -it SymbolicLink -Path powershell.exe -Target .\pwsh.exe; `
    cd -;
```

シンボリックリンクは2つ必要です

- `C:\Windows\System32\WindowsPowerShell\v1.0`→pwsh.exeのあるフォルダ
- powershell.exe→pwsh.exe

結果に大きな差はない気がするのでシンボリックリンクを使うほうがが良いかなと思います。

## (突破)pwshを起動するように変えてもWindows Powershell依存の処理でエラーが出て終わる

大体はpwshでもうまく動きますが、そうもいかない箇所があります。
エラーを回避するために以下のProfileをコンテナ内に仕込みます。

```powershell
function gcim {
    switch ($args[0]) {
        "win32_process" {
            ps | Add-Member AliasProperty -Name processid -Value Id -PassThru |
            Add-Member ScriptProperty -Name parentprocessid -Value { $this.Parent.Id } -PassThru
        }
        "Win32_OperatingSystem" {
            [PSCustomObject]@{
                Version = $PSVersionTable.os -replace '\w+\s\w+\s'
            }
        }
    }
}


$ExecutionContext.SessionState.InvokeCommand.PreCommandLookupAction = {
    [System.Management.Automation.CommandLookupEventArgs]$cl = $_
    if ($cl.CommandName -eq 'al_') {
        $cl.CommandScriptBlock = {
            $s = i_
            if (Test-Path $log) {
                del $log
            }
        
            $ab_ = $sDir -replace ' ', '` '
            ak_
            $args = "--start-server --host=127.0.0.1 --enable-remote-auto-shutdown --port=0 --connection-secret '$ai_' $q_ $exts *> '$log'"
            $splat = @{
                FilePath     = "pwsh.exe"
                ArgumentList = @(
                    "-ExecutionPolicy", "Unrestricted", "-NoLogo", "-NoProfile", "-NonInteractive", "-c", "$ab_\server.cmd $args"
                )
                PassThru     = $True
            }
        
            "Starting server: & '$sDir\server.cmd' $args"
            $global:v_ = (start @splat).ID
            $s.Stop()
            $global:m_ = $s.ElapsedMilliseconds
        }
    }
}
```

`gcim`はWMI由来の関数です。Nanoserverには無いので同じ結果になるようにその場しのぎの関数を定義します。
`$ExecutionContext.SessionState.InvokeCommand.PreCommandLookupAction`はPowershellのコマンドが実行される前に発生するイベントです。実行される処理の内容をを弄ることができます。**`ls`を`sl`にしたり、コマンドを隠しログに保存してから空白文字に書き換え（なかったことにする）たり、全く別のコマンドを実行するなんてこともできますが、悪用厳禁です。**
これでRemote-sshの処理中に定義される`al_`という関数を修正できます。
エラーを回避することでNanoserverにもRemoteできるようになります。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/e8b43230-3e0a-9808-05c6-6d1affe7eede.png)

### (未解決)「WinPTY agentが死んだ！」「この人でなし！」

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/4c00cdc3-487e-4111-58d4-957cdce7ff26.png)

というわけでTerminalが起動できません。あと、C#拡張等もうまくセットアップされませんでした。
VScodeのCliやTerminalをつかさどる部分の何処かによくないことが起きてるかもしれません。
開発者ツールで事件現場を探そうとしましたがjavascript分からない勢なので特定は未だできてません。デバッガーの力に頼るのには限度がありました……

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/dc5ca096-8f02-05a5-d4e7-f974ce3e78d2.png)

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/57856034-3036-05e7-582a-828c9b27606a.png)

コンテナ内で`pty-host`が適切に起動できてない気配がしますが、この辺をデバッグするために「何を知らないといけないのか」すらわからない状況です(辛い)

## 終わりに

最後に最も高い壁が待ち受けていましたが、これを超えさえすればNanoserverで使い捨て開発環境を低コスト生成できる道が開けそうな気がします。
JavaScriptについては文法はわからないけどデバッグ方法は察しが付くという状況なのですが、さすがに限界……？
それともWindowsの低領域の知識を学んだほうが近道なのか……？
そんな調子な今日この頃です。

## 蛇足

ところで、wslではGUI出来るようになりそうなのにWindowsコンテナで同じことはできないって世界が間違っている気がします。
コンテナ用に偽装ログイン、偽装dwmとかする仕組み無いんだろうか……
