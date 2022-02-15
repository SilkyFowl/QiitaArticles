<!--
title:   Powershell7.0で最短FizzBuzzに挑む
tags:    PowerShell
id:      49d175a495c5c43a15b9
private: false
-->

ふと思い立ったその日がFizzBuzz日和な気がしました。

# きっかけ

こちらの記事を読んで、去年11月ごろにFizzBuzzに挑んでいたことを思い出したんです。
[世界のプログラミング言語(25) PowerShell - マイクロソフトのモダンでオープンなシェル言語](https://news.mynavi.jp/article/programinglanguageoftheworld-25/#マイナビニュース)

```Powershell
# ワンライナー
1..100|&{process{if($_ % 3 -eq 0){$fz="Fizz"};if($_ % 5 -eq 0){$fz+="Buzz"};if($null -eq $fz){$_}else{$fz;$fz=$null};}}

# 可読性優先
1..100 | & {
    process {
        if ($_ % 3 -eq 0) { $fz = "Fizz" }
        if ($_ % 5 -eq 0) { $fz += "Buzz" }
        if ($null -eq $fz) {
            $_
        } else {
            $fz
            $fz = $null
        }
    }
}
```

その後、Powershell7.0がリリースされたり、Powershellへの理解が深まるなどしたため、今ならFizzBuzzの文字数をかなり減らせそうと感じたので実際にやってみました。

# リトライ

試行錯誤の末、51文字まで縮めることが出来ました。

```powershell
# ワンライナー

1..100|%{($_%3 ?$n :"Fizz")+($_%5 ?$n :"Buzz")??$_}

# 可読性優先
1..100 | ForEach-Object {
    ($_ % 3 ? $n :"Fizz") + ($_ % 5 ? $n :"Buzz") ?? $_
}
```

短縮ポイント

- Powershell7.0の機能を使う
  - 三項演算子　`<condition> ? <if-true> : <if-false>`
  - null合体演算子 `$x ??　"右辺は左辺がNullなら実行される"`
- 作らなくていい変数は作らない
- 文字数節約のために`$null`じゃなくて未定義変数を使う
  - 定義してない変数にアクセスすると`null`が返ってくる性質を利用

これが一番短いとはいえないけど、結構いい線いっていると思います。
