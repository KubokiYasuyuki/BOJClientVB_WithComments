# BOJClientVB 単体テスト仕様 (AAVerifyID)

本フォルダには、`BOJClientVB.BOJClientVBCLS.AAVerifyID` を対象とした単体テスト
の仕様書・スタブサーバ・関連スクリプトを格納しています。

## ファイル一覧

| ファイル | 内容 |
| --- | --- |
| `BOJClientVB_UnitTestSpec_AAVerifyID.xlsx` | 単体テスト仕様書 (Excel)。表紙・環境設定・スタブサーバ・テストケース一覧 (A-01 〜 A-15)・実施記録の 5 シート構成。 |
| `stub_server.ps1` | PowerShell 製スタブサーバ。TCP:50000 で待ち受け、6 モード (`normal-ack` / `normal-ack-closed` / `ng` / `malformed-commas` / `malformed-header` / `silent`) の応答を返す。異常系テスト (szRet=2/3/4) 用。 |
| `make_testspec_xlsx.py` | Excel 仕様書の生成スクリプト (openpyxl 使用)。内容を更新したい場合はこれを編集して再生成する。 |

## 前提

- クライアント: Windows 10 / 11 (64bit)
- ビルド構成: Release / x64 (64bit) のみ (Access 64bit 版からの呼び出し前提)
- .NET Framework 4.5 以上
- COM 登録: `%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe BOJClientVB.dll /codebase /tlb`

## テスト実施手順 (概要)

1. Excel 仕様書「環境設定」シートに従って作業を行う (バックアップ → レジストリ設定 → services 設定 → VBScript 配置 → テスト実行 → 復元)
2. `stub_server.ps1` を別ターミナル (管理者 PowerShell) で起動
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\stub_server.ps1 -Mode normal-ack
   ```
3. VBScript (`test_aaverifyid.vbs`) を 64bit cscript で実行
   ```cmd
   C:\Windows\System32\cscript.exe //nologo test_aaverifyid.vbs
   ```
4. Excel「テストケース一覧」に沿って、スタブモードを切り替えながら A-01 〜 A-15 を実施
5. 結果を「実施記録」シートに転記

詳細は Excel 仕様書を参照してください。

## Excel 仕様書の再生成

Python 3 + openpyxl がインストールされた環境で以下を実行します。

```bash
pip install openpyxl
python make_testspec_xlsx.py
```

生成先ファイル名はスクリプト末尾 `out = ...` で指定しています。
