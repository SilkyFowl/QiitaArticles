<!--
title:   wmicの移行先としてのPowershellまとめ
tags:    PowerShell,WMI,cim,wmic
id:      61a41dfb6b964dfbb53d
private: false
-->
# 移行するにはググラビリティは低い気がする

WMICが非推奨になってしまいました。

[旧「Edge」、「IE11」、wmicコマンド……「Windows 10 バージョン 21H1」で削除・非推奨となる機能たち - 窓の杜](https://forest.watch.impress.co.jp/docs/news/1325972.html)

そういう訳で色々まとめてみました。

## 参考になりそうな記事

まずは読むべき記事

https://docs.microsoft.com/ja-jp/powershell/scripting/learn/ps101/07-working-with-wmi?view=powershell-7.1

[貝殻本](https://book.mynavi.jp/ec/products/detail/id=90597)の著者による解説

https://tech.guitarrapc.com/entry/2013/02/08/210233

Wmi/CIMのクエリ解説

https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_wql?view=powershell-5.1

少しディープなテクニック集（英語）

https://devblogs.microsoft.com/powershell/cim-cmdlets-some-tips-tricks/

リモートでWMI使ってた人向け？

https://docs.microsoft.com/ja-jp/powershell/module/microsoft.wsman.management/about/about_wsman_provider?view=powershell-7.1

### エイリアス

CommandType| Name|Version|Source
|:-----------|:------------|:------------:|:------------:|
Alias|gcai -> Get-CimAssociatedInstance|7.0.0.0    |CimCmdlets|
Alias|gcim -> Get-CimInstance|7.0.0.0    |CimCmdlets|
Alias|gcls -> Get-CimClass|7.0.0.0    |CimCmdlets|
Alias|gcms -> Get-CimSession|7.0.0.0    |CimCmdlets|
Alias|icim -> Invoke-CimMethod|7.0.0.0    |CimCmdlets|
Alias|ncim -> New-CimInstance|7.0.0.0    |CimCmdlets|
Alias|ncms -> New-CimSession|7.0.0.0    |CimCmdlets|
Alias|ncso -> New-CimSessionOption|7.0.0.0    |CimCmdlets|
Alias|rcie -> Register-CimIndicationEvent|7.0.0.0    |CimCmdlets|
Alias|rcim -> Remove-CimInstance|7.0.0.0    |CimCmdlets|
Alias|rcms -> Remove-CimSession|7.0.0.0    |CimCmdlets|
Alias|scim -> Set-CimInstance|7.0.0.0    |CimCmdlets|

### 小ネタ

Powershellでwmiクラスを定義する方法
cimクラスはよくわからなかった……

```powershell
$newClass = [wmiclass]::new("root\cimv2", [String]::Empty, $null); 
$newClass.Name= "Test_Class"

$newClass.Qualifiers.Add("Static", $true)
$newClass.Properties.Add("ServerGroup", [System.Management.CimType]::String, $false)
$newClass.Properties["ServerGroup"].Qualifiers.Add("Key", $true)

$newClass.Properties.Add("ServerPhase", [System.Management.CimType]::String, $false)
$newClass.Properties["ServerPhase"].Qualifiers.Add("Key", $true)

$newClass.Methods.Add

# クラスを登録
$newClass.Put()

# クラスを削除
$newClass.Delete()
```
