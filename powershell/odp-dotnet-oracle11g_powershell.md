<!--
title:   Powershell 7でODP.NET Core
tags:    ODP.NET,Oracle11g,PowerShell
id:      dbd750d7bc225b369e33
private: false
-->
Powershell7.0は.NET Core 3.1を基板としています。 
よって、Oracle.ManagedDataAccessではなく、Oracle.ManagedDataAccess.**Core**を利用します。

参考: [Oracle Data Provider for .NETバージョニング体系](https://docs.oracle.com/cd/F19136_01/odpnt/InstallVersioningScheme.html#GUID-54448394-9F84-4B23-8E9B-9EC7A91B382D)

# ODP.NET Coreのインストール

この記事ではNugetから取得します。

参考: [Oracle Data Provider for .NET Coreのインストール](https://docs.oracle.com/cd/F19136_01/odpnt/InstallVersioningScheme.html#GUID-54448394-9F84-4B23-8E9B-9EC7A91B382D)


```Powershell
PS E:\> Install-Package Oracle.ManagedDataAccess.Core -scope CurrentUser -Verbose

VERBOSE: Acquiring providers for assembly: C:\program files\powershell\7\Modules\PackageManagement\coreclr\netstandard2.0\Microsoft.PackageManagement.ArchiverProviders.dll
VERBOSE: Acquiring providers for assembly: C:\program files\powershell\7\Modules\PackageManagement\coreclr\netstandard2.0\Microsoft.PackageManagement.CoreProviders.dll
# 中略
VERBOSE: Performing the operation "Install Package" on target "Package 'Oracle.ManagedDataAccess.Core' version '2.19.60' from 'nuget.org'.".
The package(s) come(s) from a package source that is not marked as trusted.
Are you sure you want to install software from 'nuget.org'?
[Y] Yes [A] Yes to All [N] No [L] No to All [S] Suspend [?] Help (default is "No"): a
# 中略
Name                           Version          Source           Summary
----                           -------          ------           -------
Oracle.ManagedDataAccess.Core  2.19.60          nuget.org        Oracle Data Provider for .NET Core for Oracle Database

PS E:\> 
```

# 使い方

オラクルクライアントoracle11g2インストール済み環境の場合です。
変則的ですが、アセンブリはGACに登録せずPowershellのセッション毎に呼び出します。

参考
[Oracle Data Provider for .NET Coreの構成](https://docs.oracle.com/cd/F19136_01/odpnt/InstallCoreConfiguration.html#GUID-24C963AE-F20B-44B5-800C-594CA06BD24B)
[OracleConfiguration.TnsAdmin](https://docs.oracle.com/cd/F19136_01/odpnt/ConfigurationTnsAdmin.html#GUID-30FDA896-3814-455B-9D45-6512705D95D3)
[OracleConnectionStringBuilderクラス](https://docs.oracle.com/cd/F19136_01/odpnt/OracleConnectionStringBuilderClass.html#GUID-81FD2CFC-D2CE-47A8-BF0A-2F18428A6D5F)

```powershell

using namespace Oracle.ManagedDataAccess.Client

# ダウンロードしたライブラリのパスを生成
$sourcePath=Split-Path (Get-Package Oracle.ManagedDataAccess.Core).Source
$dllPath = Join-Path $sourcePath lib netstandard2.0 Oracle.ManagedDataAccess.dll
# ダウンロードしたライブラリを呼び出す
Add-Type -path $dllPath

# レジストリからtnsnames.oraが配置されているディレクトリを取得して設定
# ※文字列で直接設定する場合
# [OracleConfiguration]::TnsAdmin="C:\oracle\oracle11g2\product\11.2.0\client_1\NETWORK\ADMIN"
$tnsAdmin =Join-Path (Get-itemproperty (Join-Path HKLM: SOFTWARE ORACLE *)).ORACLE_HOME NETWORK ADMIN
[OracleConfiguration]::TnsAdmin = $tnsAdmin

# ConnectionStringBuilderを使用
$connStrSb = [OracleConnectionStringBuilder]::new()
$connStrSb.Add("USER ID","scott")
$connStrSb.Add("PASSWORD","tiger")
$connStrSb.Add("DATA SOURCE","oracle")
$conn = [OracleConnection]::new($connStrSb.ToString())

# DB接続
$conn.Open()

# 取得可能なスキーマ情報の一覧を表示
$conn.GetSchema()|Out-GridView

# 各スキーマで利用可能な制限値
$conn.GetSchema("Restrictions")|Out-GridView

# スキーマ情報表示
$conn.GetSchema("DataSourceInformation")|Out-GridView
$conn.GetSchema("Tables")|Out-GridView
$conn.GetSchema("Views")|Out-GridView
$conn.GetSchema("Indexes")|Out-GridView
$conn.GetSchema("Columns")|Out-GridView
$conn.GetSchema("IndexColumns")|Out-GridView
# 条件を指定してスキーマを取得
$conn.GetSchema("Views",@("OWNERNAME",$null))|Out-GridView

# 取得したスキーマ情報を加工
$Views = $conn.GetSchema("OWNERNAME",@("OWNER",$null))
$Views.Columns.ColumnName|Out-GridView
$Views.Select("VIEW_NAME = 'products'")|Out-GridView

# クエリ実行
# パラメータなしのコンストラクタがあるクラスはハッシュテーブルから生成可能
$adapter = [OracleDataAdapter]@{
    SelectCommand=[OracleCommand]@{
        CommandText="SELECT TOP 10 * FROM products"
        Connection=$conn
    }
}
# クエリ結果をDatasetにセット
$ds=[System.Data.DataSet]::new("ds")
$adapter.Fill($ds)

# 接続終了
$conn.Close()
$conn.Dispose()
```

普段はスキーマ情報の確認などで使ってます。