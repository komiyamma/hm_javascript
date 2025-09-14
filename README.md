# hmJS

![hmJS v2.0.2](https://img.shields.io/badge/hmJS-v2.0.2-6479ff.svg) ← ここはバッジにしたほうがよさそう
[![Apache 2.0](https://img.shields.io/badge/license-Apache_2.0-blue.svg?style=flat)](LICENSE.txt)
![Hidemaru 8.73+](https://img.shields.io/badge/Hidemaru-v8.73+-6479ff.svg)
![JScript 5.8](https://img.shields.io/badge/JScript-v5.8-6479ff.svg?logo=javascript&logoColor=white)
![.NET Framework 4.0+](https://img.shields.io/badge/.NET-4.0+-blueviolet.svg?logo=.net)

(https://秀丸マクロ.net/?page=nobu_tool_hm_javascript

## 概要

`hmJS` は、高機能テキストエディタ「秀丸エディタ」のマクロを、**JavaScript (JScript)** と **.NET Framework** を使って、より強力かつ柔軟に記述するためのライブラリです。

秀丸マクロの便利な機能はそのままに、JavaScriptの持つ豊富な表現力や、.NET Frameworkの広範なライブラリ資産を活用することができます。

## 主な機能

*   **JavaScriptによるマクロ記述**: 秀丸マクロのコマンド実行や変数へのアクセスを、JavaScriptから直感的に行えます。
*   **.NET Frameworkとの連携**: .NETのクラスライブラリ（`System.Windows.Forms`など）や、自作のC#製DLLなどをスクリプト内から直接利用できます。
*   **ActiveXObjectの利用**: WSH (Windows Scripting Host) でおなじみの`ActiveXObject`も利用可能で、既存のJScript資産を活かせます。
*   **jsmode互換**: 多くの`jsmode`用関数（`hidemaruGlobal`など）と互換性があり、既存のマクロからの移行も容易です。
*   **TypeScript対応**: 型定義ファイル (`hmJS.d.ts`) が提供されており、静的型付けによる安全で快適な開発が可能です。

## 動作環境

*   **秀丸エディタ**: ver8.73 以上
*   **Visual C++ ランタイム**: [Visual Studio 2017 の Microsoft Visual C++ 再頒布可能パッケージ (x86版)](https://learn.microsoft.com/ja-jp/cpp/windows/latest-supported-vc-redist)
    *   OSが64bit版であっても、**x86版**のインストールが必要です。
    *   秀丸エディタ64bit版を利用する場合は、**x64版**をインストールしてください。
*   **.NET Framework**: 4.0 以上

## インストール

1.  **ダウンロード**:
    *   秀丸エディタが**32bit版**の場合: [hmJS.zip](https://xn--pckzexbx21r8q9b.net/other_soft/hm_javascript/hmJS.zip)
    *   秀丸エディタが**64bit版**の場合: [hmJS_x64.zip](https://xn--pckzexbx21r8q9b.net/other_soft/hm_javascript/hmJS_x64.zip)
2.  **配置**:
    ダウンロードしたzipファイルを解凍し、中にある `hmJS.dll` を、秀丸エディタのインストールディレクトリ（`hidemaru.exe` がある場所）にコピーしてください。

## 使用方法

`hmJS.dll` を `loaddll` で読み込み、`dllfuncw` で `DoString` または `DoFile` を呼び出してJavaScriptコードを実行します。

### 例1: 基本的なマクロの実行

```c
#JS = loaddll( hidemarudir + @"\hmJS.dll" );

#_ = dllfuncw( #JS, "DoString", R"JS(

// 秀丸のコマンドを直接実行
message("OK");
moveto(3, 4); // 4行目3桁目へ移動

)JS"
);

freedll(#JS);
```

### 例2: .NET Frameworkの利用

```c
#JS = loaddll( hidemarudir + @"\hmJS.dll" );

#_ = dllfuncw( #JS, "DoString", R"JS(

// .NETのSystem.Text.Encodingを使い、テキストがASCIIか判定する関数
function isAscii(text) {
    // Shift_JISエンコーディングのインスタンスを取得
    var sjis = clr.System.Text.Encoding.GetEncoding("Shift_JIS");

    // バイト数と文字数を比較
    return sjis.GetByteCount(text) == text.length;
}

// 現在編集中のテキスト全体を判定
var result = isAscii(hm.Edit.TotalText);

// 結果をデバッグモニタに出力
hm.debuginfo("Is ASCII: " + result);
message("Is ASCII: " + result);

)JS"
);

freedll(#JS);
```

### 例3: .NETでWindowsフォームを作成

```c
#JS = loaddll( hidemarudir + @"\hmJS.dll" );

#_ = dllfuncw( #JS, "DoString", R"(

// System.Windows.Formsアセンブリを読み込む
host.lib("System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");
var Forms = clr.System.Windows.Forms;

// フォームとボタンを作成
var form = new Forms.Form();
form.Text = "hmJS Sample";

var button = new Forms.Button();
button.Text = "Click Me!";
button.Left = 16;
button.Top = 16;

var i = 0;
// ボタンのクリックイベントに関数を接続
button.Click.connect(function(sender, args) {
    i++;
    // 秀丸マクロの変数を更新
    hm.Macro.Var("$count", i);
    Forms.MessageBox.Show("Clicked " + i + " times!");
});

form.Controls.Add(button);
form.ShowDialog();

)");

// JavaScript側で更新した変数をマクロ側から参照
message("ボタンは " + $count + " 回クリックされました");

freedll(#JS);
```

## プロジェクト構成

*   `hmJS.src/`: `hmJS.dll` のC++ソースコード。秀丸エディタとスクリプトエンジン間のブリッジ処理を担います。
*   `JScriptExtender/`: JScript環境向けのユーティリティライブラリ (`StreamReader`, `Ini` パーサー等) です。
*   `TSDeclare/`: `hmJS` のAPIのTypeScript型定義ファイル (`hmJS.d.ts`) が含まれています。
*   `Release/`: コンパイル済みの `hmJS.dll` とライセンスファイルが格納されています。

## ライセンス

このプロジェクトは **Apache License 2.0** の下で公開されています。

-   `hmJS`: Apache License 2.0
-   `ClearScript`: MIT License (内部で利用)
