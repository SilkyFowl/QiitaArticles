<!--
title: Powershell 6以降でもGetActiveObject()でExcelを呼ぶ方法
tags:  PowerShell,Excel
-->

Powershell 6以降、`[System.Runtime.InteropServices.Marshal]::GetActiveObject(progID)`はありません。ですが、作るのは簡単です。

## クロスプラットフォーム化の代償

Windows専用のShellとして誕生したPowershellは、クロスプラットフォーム化してLinuxやMacで動くようになりました。しかしその際、様々な変更や**使用できなくなったAPIがあります。**

https://docs.microsoft.com/ja-jp/powershell/scripting/whats-new/differences-from-windows-powershell?view=powershell-7.2

`[System.Runtime.InteropServices.Marshal]::GetActiveObject`も、使えなくなった機能の一つです。

## 無いなら作ろう`GetActiveObject`

幸い、[stack overfowでGetActiveObjectの中身が分かります。](https://stackoverflow.com/a/65496277)コードが分かれば解決出来たも同然です、Powershellで`Add-Type`を使えばいいのです。

```powershell:GetActiveObject_自作

Add-Type -TypeDefinition @'
using System;
using System.Runtime;
using System.Runtime.InteropServices;
public static class Marshal2
{
    internal const String OLEAUT32 = "oleaut32.dll";
    internal const String OLE32 = "ole32.dll";

    public static Object GetActiveObject(String progID)
    {
        Object obj = null;
        Guid clsid;

        // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
        // CLSIDFromProgIDEx doesn't exist.
        try
        {
            CLSIDFromProgIDEx(progID, out clsid);
        }
        //            catch
        catch (Exception)
        {
            CLSIDFromProgID(progID, out clsid);
        }

        GetActiveObject(ref clsid, IntPtr.Zero, out obj);
        return obj;
    }

    //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
    [DllImport(OLE32, PreserveSig = false)]
    private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

    //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
    [DllImport(OLE32, PreserveSig = false)]
    private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

    //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
    [DllImport(OLEAUT32, PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);

}
'@

```

`Marshal2`というクラス名には特に意味はありません。重要なのは中で読んでいる3つのWindows API関数です。これで、Powershell6以降でも起動済みのExcelを捕まえることができます。

:::note info
`Add-Type`によるC#コード埋め込みは[MSDNにも書いてある](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/add-type?view=powershell-7.2#examples)由緒正しき方法です。
:::

```powershell:使い方はPowershell5.1までのGetActiveObject(progID)と同じ
# 起動済みのExcelのインスタンスを捕まえる
$excel = [Marshal2]::GetActiveObject('Excel.Application')

# 新規作成
$newBook = $excel.Workbooks.Add()

# 値の代入
$newBook.ActiveSheet.Range('A1:B3').Value2 = 10

# ダイアログ無しでブックを閉じる
$excel.ActiveWorkbook.Close($false)

# Excelの修了
$excel.Quit()

# 後始末
Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable | Remove-Variable
[GC]::Collect()
```

```:ダイアログが出るとまずい場合はこうする
$excel.DisplayAlerts = $false
```

`Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable | Remove-Variable`については[この記事](https://qiita.com/SilkyFowl/items/b4b6271619bd6d3824f7)をご覧ください。
