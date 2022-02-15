<!--
title:   PowershellからAccessにSQLするためのConnectionオブジェクトの作り方
tags:    ADO.NET,PowerShell,access
id:      695cb7a7f376c8f40fbf
private: false
-->
# これは何？

ググって知った情報を統合して関数にしました

```powershell
using namespace  System.Data
using namespace  System.Data.OleDb

function New-MsAccessConnection {
    [OutputType([System.Data.OleDb.OleDbConnection])]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript(
            { Test-Path $_ -Include "*.accdb"  -PathType Leaf },
            ErrorMessage = ": {0} はAccessファイルではありません。"
        )]
        [string]
        $acPath
    )
    # 接続文字列作成
    $builder = [OleDbConnectionStringBuilder]::new()

    # インストールされている最新のMicrosoft.ACE.OLEDBプロバイダ名を取得
    $builder["Provider"] = ([OleDbEnumerator]::new().GetElements().SOURCES_NAME -match "^Microsoft.ACE.OLEDB." )[-1]
    # 対象Accessデータベース
    $builder["Data Source"] = $acPath
    # システムデータベース(指定することでMSysRelationshipsなどにクエリ可能になる)
    $builder["Jet OLEDB:System database"] = "$env:APPDATA\Microsoft\Access\System.mdw"

    return [OleDbConnection]::new($builder.ConnectionString)
}
```

## プロバイダ名設定方法

OledbでAccessに接続する際、バージョン違いで動作しないという事態を避けるには実行環境で利用可能なプロバイダ名を動的に取得する必要があります。
その方法の比較です。
1秒弱の処理時間を許容出るなら、安全で（比較的）可読性の高い`OleDbEnumerator`の方が良いでしょう。

### `OleDbEnumerator`を利用する

`OleDbEnumerator`クラスで利用可能なプロバイダを全て取得。その中から目的のプロバイダ名を取得する方法

#### メリット

安全
Powershell(.NET)から認識可能なプロバイダから選ぶ方式なので存在しないプロバイダを指定していないか考える必要がなくなる

#### デメリット

遅い
この処理だけで1秒弱かかる。体感的には少しもたつくレベル。

```powershell
Measure-Command {
    $provider=([OleDbEnumerator]::new().GetElements().SOURCES_NAME -match "^Microsoft.ACE.OLEDB." )[-1]
} | % TotalMilliseconds

938.4963
```

### レジストリからOfficeのバージョンナンバー取得する

現状のバージョンナンバー付与のルールがOfficeのバージョンナンバーと一致している事を利用した方法

#### メリット

早い
殆ど気にならないレベル

#### デメリット

様々な要因で動作しなくなる可能性がある

- インストール方法が変則的な場合
- 何らかの要因でレジストリが変更された場合
- MSが命名規則を変更した場合

```powershell
Measure-Command {
    $varsion = Get-Item 'HKLM:\SOFTWARE\Microsoft\Office\??.?' | ForEach-Object {
        $_.Name -replace "^.+(\d\d.\d)$", "`$1"
    } | Sort-Object -Descending -Top 1
} | % TotalMilliseconds

45.4354
```

この方法を使った場合のコードは以下の通り

```powershell
using namespace  System.Data
using namespace  System.Data.OleDb

function New-MsAccessConnection {
    [OutputType([System.Data.OleDb.OleDbConnection])]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript(
            { Test-Path $_ -Include "*.accdb"  -PathType Leaf },
            ErrorMessage = ": {0} はAccessファイルではありません。"
        )]
        [string]
        $acPath
    )
    # レジストリからバージョンナンバーを取得
    $varsion = Get-Item 'HKLM:\SOFTWARE\Microsoft\Office\??.?' | ForEach-Object {
        $_.Name -replace "^.+(\d\d.\d)$", "`$1"
    } | Sort-Object -Descending -Top 1

    # 接続文字列作成
    $builder = [OleDbConnectionStringBuilder]::new()
    $builder["Provider"] = "Microsoft.ACE.OLEDB.$varsion"
    $builder["Data Source"] = $acPath
    $builder["Jet OLEDB:System database"] = "$env:APPDATA\Microsoft\Access\System.mdw"

    return [OleDbConnection]::new($builder.ConnectionString)
}
```

## 注意点

インストールしてあるOfficeが32bitならPowershell(x86)を、64bitならPowershell(x64)で実行してください

## 使い方

```powershell
using namespace  System.Data
using namespace  System.Data.OleDb

function Get-MsAccessERInfo {
    [OutputType([System.Data.DataSet])]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript(
            { Test-Path $_ -Include "*.accdb"  -PathType Leaf },
            ErrorMessage = ": {0} はAccessファイルではありません。"
        )]
        [string]
        $acPath
    )
    try {
        # セットアップ
        [OleDbConnection]$connection = New-MsAccessConnection $acPath
        $connection.Open()
        [OleDbTransaction]$Transaction = $connection.BeginTransaction()
        
        $adapter = [OleDbDataAdapter]::new(
            [OleDbCommand]::new("Select * From MSysRelationships", $connection, $Transaction)
        )
        
        # テーブルデータ取得
        $dataSet = [DataSet]::new()
        $adapter.Fill($dataSet, "MSysRelationships") > $null
        $dataSet.Tables.Add($connection.GetSchema("Indexes"))  > $null
        $dataSet.Tables.Add($connection.GetSchema("Columns"))  > $null

        $Transaction.Commit()
    } catch {
        # リソース開放
        ${Transaction}?.Rollback()
        $PSCmdlet.ThrowTerminatingError($_)
    } finally {
        ${Transaction}?.Dispose()

        ${connection}?.Close()
        ${connection}?.Dispose()
    }

    # 返り値はDataset
    return [DataSet]$dataSet
}
```

Powershell7.1から正式機能になったNull 条件演算子(`?.` と `?[]`)を使用しています。
Powershell7.0の場合は`Enable-ExperimentalFeature`で有効にしてください

```powershell
Enable-ExperimentalFeature PSNullConditionalOperators
```

### Powershell5.1

以下のように修正することでPowershell5.1でも動きます

#### `New-MsAccessConnection`

```powershell
using namespace  System.Data
using namespace  System.Data.OleDb

function New-MsAccessConnection {
    [OutputType([System.Data.OleDb.OleDbConnection])]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript(
            { Test-Path $_ -Include "*.accdb"  -PathType Leaf }
        )]
        [string]
        $acPath
    )
    # 接続文字列作成
    $builder = [OleDbConnectionStringBuilder]::new()

    # インストールされている最新のMicrosoft.ACE.OLEDBプロバイダ名を取得
    $builder["Provider"] = ([OleDbEnumerator]::new().GetElements().SOURCES_NAME -match "^Microsoft.ACE.OLEDB." )[-1]
    # 対象Accessデータベース
    $builder["Data Source"] = $acPath
    # システムデータベース(指定することでMSysRelationshipsなどにクエリ可能になる)
    $builder["Jet OLEDB:System database"] = "$env:APPDATA\Microsoft\Access\System.mdw"

    return [OleDbConnection]::new($builder.ConnectionString)
}
```

`ValidateScript`の`ErrorMessage`は5.1時点ではサポートされてないので削除

```diff
@@ -6,8 +6,7 @@ function New-MsAccessConnection {
     param (
         [Parameter(Mandatory, Position = 0)]
         [ValidateScript(
-            { Test-Path $_ -Include "*.accdb"  -PathType Leaf },
-            ErrorMessage = ": {0} はAccessファイルではありません。"
+            { Test-Path $_ -Include "*.accdb"  -PathType Leaf }
         )]
         [string]
         $acPath
```

#### `Get-MsAccessERInfo`

```powershell
using namespace  System.Data
using namespace  System.Data.OleDb

function Get-MsAccessERInfo {
    [OutputType([System.Data.DataSet])]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript(
            { Test-Path $_ -Include "*.accdb"  -PathType Leaf }
        )]
        [string]
        $acPath
    )
    try {
        # セットアップ
        [OleDbConnection]$connection = New-MsAccessConnection $acPath
        $connection.Open()
        [OleDbTransaction]$Transaction = $connection.BeginTransaction()

        $adapter = [OleDbDataAdapter]::new(
            [OleDbCommand]::new("Select * From MSysRelationships", $connection, $Transaction)
        )

        # テーブルデータ取得
        $dataSet = [DataSet]::new()
        $adapter.Fill($dataSet, "MSysRelationships") > $null
        $dataSet.Tables.Add($connection.GetSchema("Indexes"))  > $null
        $dataSet.Tables.Add($connection.GetSchema("Columns"))  > $null

        $Transaction.Commit()
    } catch {
        if ($Transaction) {
            $Transaction.Rollback()
        }
        $PSCmdlet.ThrowTerminatingError($_)
    } finally {
        if ($Transaction) {
            $Transaction.Dispose()
        }

        if ($connection) {
            $connection.Close()
            $connection.Dispose()
        }  
    }

    # 返り値はDataset
    return [DataSet]$dataSet
}
```

Null 条件演算子は無いので`if`に修正

```diff
@@ -6,8 +6,7 @@ function Get-MsAccessERInfo {
     param (
         [Parameter(Mandatory, Position = 0)]
         [ValidateScript(
-            { Test-Path $_ -Include "*.accdb"  -PathType Leaf },
-            ErrorMessage = ": {0} はAccessファイルではありません。"
+            { Test-Path $_ -Include "*.accdb"  -PathType Leaf }
         )]
         [string]
         $acPath
`

`ValidateScript`の`ErrorMessage`は5.1時点ではサポートされてないので削除

```diff
@@ -30,14 +29,19 @@ function Get-MsAccessERInfo {

         $Transaction.Commit()
     } catch {
-        # リソース開放
-        ${Transaction}?.Rollback()
+        if ($Transaction) {
+            $Transaction.Rollback()
+        }
         $PSCmdlet.ThrowTerminatingError($_)
     } finally {
-        ${Transaction}?.Dispose()
+        if ($Transaction) {
+            $Transaction.Dispose()
+        }

-        ${connection}?.Close()
-        ${connection}?.Dispose()
+        if ($connection) {
+            $connection.Close()
+            $connection.Dispose()
+        }  
     }

     # 返り値はDataset
```

## おまけ：アセンブリ内のEnumを取得

```powershell
using namespace System.Collections.Generic

# Enum値を格納するDictionaryを生成
$enumDic = [Dictionary[string, enum]]::new()

# パス指定でアセンブリを読み込む
# PassThruオプションで返り値が読み込んだアセンブリ内の型になる
Add-Type -Path @(
    "C:\Windows\assembly\GAC_MSIL\office\*\OFFICE.DLL"
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Vbe.Interop\*\Microsoft.Vbe.Interop.dll"
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Access.Dao\*\Microsoft.Office.Interop.Access.Dao.dll"
    "C:\Windows\assembly\GAC\ADODB\*\ADODB.dll"
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Access\*\Microsoft.Office.Interop.Access.dll"
    # IsEnumでフィルタリングしてからEnum値を取得
) -PassThru | Where-Object IsEnum | ForEach-Object GetEnumValues | ForEach-Object {
    # 値をセット(重複値は上書き)
    $enumDic[$_] = $_
}
```

作成したDictionaryによる補完入力はISEでも使えます

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/107934/9def4a79-4a64-ca3b-bff0-9a086053381e.png)
