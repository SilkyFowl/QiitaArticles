<!--
title:   VS Code Remote ContainersでAtCoder用F#環境を楽にする
tags:    AtCoder,Docker,F#,PowerShell,VSCode
id:      3725e0ccff1a1f28a407
private: false
-->
## はじめに

F#でAtCoderに参加するための環境を、VS Code Remote Containers等を利用して自動化しました。

- Docker, dotnet templates, Powershell module等の開発
- Docker Hub, nuget, Powershell Gallery等へのパッケージ公開
  - GitHub ACtionsを利用したコンテナイメージのビルド

等の勉強も兼ねてます。

## コンテナイメージ

[Language Test 202001](https://atcoder.jp/contests/language-test-202001)と[こちらの記事](https://qiita.com/hinamimi/items/b3dd159f956628cebdbb "DockerでAtCoderができる環境を作る【Python・C++】")を参考にしました。

https://github.com/SilkyFowl/docker-atcoder-fs

https://hub.docker.com/repository/docker/syamorock/atcoder-fs

対応言語:F#(.NET Core3.1)

組み込みインストール

- [atcoder-cli](http://tatamo.81.la/blog/2018/12/07/atcoder-cli-tutorial/)
- [online-judge-tools](https://github.com/online-judge-tools/oj)

作業簡略化のため、.netテンプレートとPowershellモジュールを作成しました。

https://github.com/SilkyFowl/atcoder-fsharp-templates

https://www.powershellgallery.com/packages/AtcoderFs.Pwsh

https://www.nuget.org/packages/AtcoderFs.Templates

```dockerfile
# See here for image contents: https://github.com/microsoft/vscode-dev-containers/tree/v0.205.2/containers/ubuntu/.devcontainer/base.Dockerfile

# [Choice] Ubuntu version (use hirsuite or bionic on local arm64/Apple Silicon): hirsute, focal, bionic
# ARG VARIANT="bionic"
FROM mcr.microsoft.com/vscode/devcontainers/base:0-bionic

# install software-properties-common(add-apt-repository)
RUN apt-get update \
    && export DEBIAN_FRONTEND=noninteractive \
    && apt-get -y install --no-install-recommends software-properties-common \
    && rm -rf /var/lib/apt/lists/* \
# for pypy3
    && add-apt-repository ppa:pypy/ppa -y \
# for nodejs
    && curl -sL https://deb.nodesource.com/setup_14.x | bash - \
# for .NET and powershell
    && wget https://packages.microsoft.com/config/ubuntu/18.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb \
    && dpkg -i packages-microsoft-prod.deb \
    && rm packages-microsoft-prod.deb \
    && apt-get update \
    && apt-get -y install --no-install-recommends \
# .NET
        dotnet-sdk-3.1 \
        dotnet-sdk-6.0 \
# Powershell
        powershell \
# Python3, PyPy3の3つの環境想定
        python3.8 \
        python3-pip \
        pypy3 \
# node
        nodejs \
# online-judge-tools用ライブラリ
        time \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
# update-alternatives
    && update-alternatives --install /usr/bin/python python /usr/bin/python3.8 30 \
    && update-alternatives --install /usr/bin/pip pip /usr/bin/pip3 30 \
    && update-alternatives --install /usr/bin/pypy pypy /usr/bin/pypy3 30 \
    && update-alternatives --install /usr/bin/node node /usr/bin/nodejs 30 \
# コンテスト補助アプリケーションをインストール
    && pip install online-judge-tools \
    && rm -rf ~/.cache/pip \
    && npm install -g atcoder-cli \
    && npm cache clean --force

SHELL ["pwsh","-command"]

# powershell セットアップ
RUN Set-PSRepository -Name PSGallery -InstallationPolicy Trusted \
    && Install-Module Pester,AngleParse,InvokeBuild,AtcoderFs.Pwsh  -Scope AllUsers -Confirm:$False -AcceptLicense  -Force -Verbose \
    && Install-Script -Name Invoke-Build.ArgumentCompleters  -Scope AllUsers -Confirm:$False -AcceptLicense  -Force -Verbose \
    && 'Import-Module AtcoderFs.Pwsh' | Add-Content $PROFILE.AllUsersAllHosts -Force

```

## 使い方

### セットアップ

vscodeにremote-containers拡張機能をインストールします。

https://marketplace.visualstudio.com/items?itemName=ms-vscode-remote.remote-containers

ワークスペースフォルダーに`./.devcontainer/devcontainer.json`を作ります。

```json
// For format details, see https://aka.ms/devcontainer.json. For config options, see the README at:
// https://github.com/microsoft/vscode-dev-containers/tree/v0.202.5/containers/cpp
{
	"name": "Atcoder helper for F#",
	"image": "syamorock/atcoder-fs:main",

	// Set *default* container specific settings.json values on container create.
	"settings": {
	},

	// Add the IDs of extensions you want installed when the container is created.
	"extensions": [
		"ms-vscode.cpptools",
		"ms-dotnettools.csharp",
		"Ionide.Ionide-fsharp",
		"alfonsogarciacaro.vscode-template-fsharp-highlight",
		"ms-vscode.powershell-preview"
	],

	// Use 'forwardPorts' to make a list of ports inside the container available locally.
	// "forwardPorts": [],

	// Use 'postCreateCommand' to run commands after the container is created.
	"postCreateCommand": "dotnet tool install paket -g",

	// Comment out connect as root instead. More info: https://aka.ms/vscode-remote/containers/non-root.
	"remoteUser": "vscode"
}
```

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/a59186cf-fb86-b68e-6c0d-b260ff5e6ddc.png)

開発コンテナを起動します

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/7c561d8b-7fb7-079a-fc05-113f44722edb.png)

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/025dbe76-fc79-1885-0136-56569f96c2d1.png)

### コンテナの初回起動時

`acc`と`oj`にログインします。

※それぞれ、`Enter-AtCoderCli`と`Enter-AtCoderOnlineJudge`というラッパー関数を作成しています。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/13c38c16-d8a8-67e4-e1bc-fbb1286b2068.png)

`Update-FsTemplate`を実行してテンプレート関連の初期化を行います。
`Initialize-PaketDepandencies`を実行して`paket`の初期化を行います。

これで環境構築、初期設定は完了です。

### コンテスト参加時の流れ

参加するコンテストのコンテストIDを使用します。

`https://atcoder.jp/contests/<contestId>`

例:[AtCoder Beginners Selection](https://atcoder.jp/contests/abs)
URL…`https://atcoder.jp/contests/abs`
→コンテストIDは`abs`

#### コンテスト用のソリューションを作成する

開発コンテナを開いて`New-AtCoderContest`を実行します。

```powershell
New-AtCoderContest -contestId abs
```

コンテストID名のフォルダが生成されます。
構成は以下の通りです。

```console
abs
│  contest.acc.json
│  abs.sln
│  
├─abc088b                       …… 各問題ごとに1フォルダ
│  │  abc088b.fsproj
│  │  Program.fs                …… 提出用F#ファイル
│  │  abc088b.draft.fsx         …… ラフスケッチ用のソリューションから独立したF#スクリプト
│  │  
│  ├─test
│  ├─obj
│  └─bin
│                      
(略)                      
│                      
└─abs.Tests                     …… テスト用のプロジェクト
   │  abs.Tests.fsproj
   │  paket.references
   │  Utils.fs
   │  CustomTests.fs
   │  Tests.fs
   │  Program.fs
   │  
   ├─obj
   └─bin
```

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/38d8f95b-9378-71f0-bc3f-c6b5a6488b0b.png)

#### 問題を解く

解きたい問題のフォルダの`Program.fs`を開いてプログラムを作成します。
問題のフォルダ名の関数に処理を記述する想定です。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/d80f15b5-cd28-571a-21a5-7be1d7038f91.png)

#### テスト

`Test-AtCoder`を使って2種類のテストが可能です。

##### 自動生成されたテストプロジェクトによるテスト

```powershell
Test-AtCoder -FolderPath <コンテストのフォルダ>
```

自動生成されたテストプロジェクトを利用してコンテストプロジェクト全体のテストを実行します。
初期状態ではすべてのテストがコメントアウトされています。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/36f4c8d9-9cae-8e64-edda-59dc179b446b.png)

```powershell:実行例
Test-AtCoder ./abs/
```

結果

```sh:初期状態では何も実行されない
[03:45:53 INF] EXPECTO? Running tests... <Expecto>
[03:45:53 INF] EXPECTO! 0 tests run in 00:00:00.0187893 for miscellaneous – 0 passed, 0 ignored, 0 failed, 0 errored. Success! <Expecto>
```

実施したいテストの行のコメントアウトを解除してください。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/f99a96f3-d75a-45cd-3c4c-e53a92f375dc.png)

テストが実施されます。

```console:何も処理を書いてないでテストは失敗します
[03:49:13 INF] EXPECTO? Running tests... <Expecto>
[03:49:13 ERR] Sample File Tests.Welcome to AtCoder sample-2 failed in 00:00:00.0450000. 
WA:{ Parent = { Title = "Welcome to AtCoder"
             FolderName = "practicea"
             Memory = 256
             Timeout = 00:00:02 }
  InFile = /workspaces/test/abs/practicea/test/sample-2.in
  OutFile = /workspaces/test/abs/practicea/test/sample-2.out }

. String actual was shorter than expected, at pos 0 for expected item '4'.
expected:
456 myonmyon

  actual:

   at Utils.generateSampleFileTestCase@68-2.Invoke(StringWriter _arg2) in /workspaces/test/abs/abs.Tests/Utils.fs:line 81
   at Microsoft.FSharp.Core.Operators.Using[T,TResult](T resource, FSharpFunc`2 action) in D:\a\_work\1\s\src\fsharp\FSharp.Core\prim-types.fs:line 4806
   at Utils.generateSampleFileTestCase@67-1.Invoke(StreamReader _arg1) in /workspaces/test/abs/abs.Tests/Utils.fs:line 67
   at Microsoft.FSharp.Core.Operators.Using[T,TResult](T resource, FSharpFunc`2 action) in D:\a\_work\1\s\src\fsharp\FSharp.Core\prim-types.fs:line 4806
   at Utils.generateSampleFileTestCase@64.Invoke(Unit unitVar) in /workspaces/test/abs/abs.Tests/Utils.fs:line 64 <Expecto>
[03:49:13 ERR] Sample File Tests.Welcome to AtCoder sample-1 failed in 00:00:00. 
WA:{ Parent = { Title = "Welcome to AtCoder"
             FolderName = "practicea"
             Memory = 256
             Timeout = 00:00:02 }
  InFile = /workspaces/test/abs/practicea/test/sample-1.in
  OutFile = /workspaces/test/abs/practicea/test/sample-1.out }

. String actual was shorter than expected, at pos 0 for expected item '6'.
expected:
6 test

  actual:

   at Utils.generateSampleFileTestCase@68-2.Invoke(StringWriter _arg2) in /workspaces/test/abs/abs.Tests/Utils.fs:line 81
   at Microsoft.FSharp.Core.Operators.Using[T,TResult](T resource, FSharpFunc`2 action) in D:\a\_work\1\s\src\fsharp\FSharp.Core\prim-types.fs:line 4806
   at Utils.generateSampleFileTestCase@67-1.Invoke(StreamReader _arg1) in /workspaces/test/abs/abs.Tests/Utils.fs:line 67
   at Microsoft.FSharp.Core.Operators.Using[T,TResult](T resource, FSharpFunc`2 action) in D:\a\_work\1\s\src\fsharp\FSharp.Core\prim-types.fs:line 4806
   at Utils.generateSampleFileTestCase@64.Invoke(Unit unitVar) in /workspaces/test/abs/abs.Tests/Utils.fs:line 64 <Expecto>
[03:49:13 INF] EXPECTO! 2 tests run in 00:00:00.1339686 for Sample File Tests – 0 passed, 0 ignored, 2 failed, 0 errored.  <Expecto>

```

##### `oj t`によるテスト

スイッチパラメータ`-UseOJ`で、指定したフォルダにあるプロジェクトをコンパイルして`oj t`を実行します。

```powershell
Test-AtCoder -FolderPath <各問題のフォルダ> -UseOJ
```

`posh
Test-AtCoder ./abs/practicea/ -UseOJ

```console
  Determining projects to restore...
  Restored /workspaces/test/abs/practicea/practicea.fsproj (in 253 ms).
  practicea -> /workspaces/test/abs/practicea/bin/Release/netcoreapp3.1/ubuntu.18.04-x64/practicea.dll
  practicea -> /workspaces/test/abs/practicea/ojTest/
[INFO] online-judge-tools 11.5.1 (+ online-judge-api-client 10.10.0)
[INFO] 2 cases found

[INFO] sample-1
[INFO] time: 0.047700 sec
[FAILURE] WA
input:
1
2_3
test

output:
(empty)
expected:
6_test


[INFO] sample-2
[INFO] time: 0.056302 sec
[FAILURE] WA
input:
72
128_256
myonmyon

output:
(empty)
expected:
456_myonmyon


[INFO] slowest: 0.056302 sec  (for sample-2)
[INFO] max memory: 23.488000 MB  (for sample-2)
[FAILURE] test failed: 0 AC / 2 cases
```

問題を解いてもう一度テストしてみます。

```fsharp
let practicea argv =
    let a = stdin.ReadLine() |> int

    let bc =
        stdin.ReadLine().Split ' ' |> Array.sumBy int

    let s = stdin.ReadLine()

    printfn "%i %s" (a + bc) s


[<EntryPoint>]
let main argv =
    practicea argv
    0

```

テスト合格です。

```shell
PS /workspaces/test> Test-AtCoder ./abs/                 
[04:02:22 INF] EXPECTO? Running tests... <Expecto>
[04:02:22 INF] EXPECTO! 2 tests run in 00:00:00.0614051 for Sample File Tests – 2 passed, 0 ignored, 0 failed, 0 errored. Success! <Expecto>
PS /workspaces/test> Test-AtCoder ./abs/practicea/ -UseOJ
  Determining projects to restore...
  Restored /workspaces/test/abs/practicea/practicea.fsproj (in 249 ms).
  practicea -> /workspaces/test/abs/practicea/bin/Release/netcoreapp3.1/ubuntu.18.04-x64/practicea.dll
  practicea -> /workspaces/test/abs/practicea/ojTest/
[INFO] online-judge-tools 11.5.1 (+ online-judge-api-client 10.10.0)
[INFO] 2 cases found

[INFO] sample-1
[INFO] time: 0.087652 sec
[SUCCESS] AC

[INFO] sample-2
[INFO] time: 0.080748 sec
[SUCCESS] AC

[INFO] slowest: 0.087652 sec  (for sample-1)
[INFO] max memory: 32.928000 MB  (for sample-1)
[SUCCESS] test success: 2 cases
```

#### 解答提出

解答が出来上がったら`Submit-AtCoderTask`で提出します。
`acc submit`のラッパーです。

```powershell
Submit-AtCoderTask -FolderPath <各問題のフォルダ>
```

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/5a27bb4e-bf37-c76e-db79-0f5daacbb33a.png)

## 課題

- Powershellモジュールのテストコードを完成させる。
- GitHub Actionsの改修

## 終わりに

つい先日、.NET 6、およびF# 6 がリリースされました。

https://forest.watch.impress.co.jp/docs/news/1364721.html

> 「.NET 6」の最大の魅力は、パフォーマンスの向上だ。プロファイルに基づく動的な最適化（Dynamic Profile-guided Optimization、Dynamic PGO）と呼ばれる技術が採用されており、たとえばTechEmpowerのMVCベンチマークではPGOにより1秒あたりに処理できるリクエスト数が26％（51万→64万）にまで改善されたという。

Atcoderで**パフォーマンスの向上**したF#6、およびC#10が使える日が早く来ることを願っています。

ということでリクエストしたいのですが......

https://atcoder.jp/contests/language-test-202001

> このコンテストは、言語のアップデートテスト用コンテストです
> [こちらのスプレッドシート](https://docs.google.com/spreadsheets/d/1PmsqufkF3wjKN6g1L0STS80yP4a6u-VdGiEv5uOHe0M/)にて募集を行っていた各言語のバージョンアップならびに新規追加した言語のテストを行うためのコンテストです。

https://docs.google.com/spreadsheets/d/1PmsqufkF3wjKN6g1L0STS80yP4a6u-VdGiEv5uOHe0M/edit#gid=0

リクエストするにはシート1に行を追加すれば良いのでしょうか……？
