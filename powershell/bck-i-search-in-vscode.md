<!--
title:   VScodeでもPowershellでコマンド履歴検索をしたい
tags:    PowerShell,VSCode
id:      712f19366ee1ef9dbdfd
private: false
-->
# 三行で
　
1. :sparkles:VScodeでもPSReadlineが正常に動くようになりました
2. :rotating_light: だけど、コマンド履歴検索はVScodeの他のキーバインドと競合しているため正常に動作しません
3.　:wrench: なんとかするには`keybindings.json`に設定を書き込めば良いらしい

# 設定はこちら

この設定を作成、もしくは追加します。[stackoverflow](https://stackoverflow.com/questions/60857148/vs-code-terminal-history-search-windows-powershell)に感謝

```json:keybindings.json
// 既定値を上書きするには、このファイル内にキー バインドを挿入しますauto[]
[
    {
        "key": "ctrl+r",
        "command": "workbench.action.terminal.sendSequence",
        "when": "terminalFocus",
        "args": { "text": "\u0012" }
    }
]
```

これでコマンド履歴を検索出来ます
![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/a5d25daf-0382-3c1a-59dd-e4b40aea5563.png)
![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/8b69bd8e-8027-071c-758c-8c95ff05b952.png)
ちなみに、bashもこの設定で同じことが出来ました。

# 動作環境
`console
❯ $PSVersionTable
Name                           Value          
----                           -----
PSVersion                      7.1.0-preview.3
PSEdition                      Core
GitCommitId                    7.1.0-preview.3
OS                             Microsoft Windows 10.0.19041                                                     
Platform                       Win32NT
PSCompatibleVersions           {1.0, 2.0, 3.0, 4.0…}       
PSRemotingProtocolVersion      2.3
SerializationVersion           1.1.0.1
WSManStackVersion              3.0

❯ Get-Module psreadline 
ModuleType Version    PreRelease Name                                ExportedCommands
---------- -------    ---------- ----                                ----------------
Script     2.1.0      beta2      PSReadLine  
`console
Visual Studio Code 1.46.1
ms-vscode.powershell-preview