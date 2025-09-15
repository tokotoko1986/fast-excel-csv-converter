# ⚡ Fast Excel CSV Converter
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

> 🚀 **High-performance Excel to CSV converter with intelligent optimization**  
> Convert your Excel files to CSV format while preserving exactly what you see in Excel - dates, percentages, currency, and all formatting intact!

---

## 🌍 Language / 言語選択
- [English](#english) | [日本語](#japanese)

---

<a name="english"></a>
# 📖 English

## ✨ Features

### 🎯 **Dual Conversion Modes**
- **Normal Mode**: Preserves all cell formatting (dates, currencies, percentages)
- **High-Speed Mode**: Ultra-fast conversion using raw values (up to 10x faster)

### 📊 **Smart Processing**
- **Batch Processing**: Convert multiple Excel files at once
- **All Sheet Support**: Automatically converts all sheets in each workbook
- **Empty Sheet Handling**: Gracefully handles empty worksheets
- **Error Recovery**: Continues processing even if individual files fail

### 🛡️ **Robust & Safe**
- **Process Management**: Automatically handles Excel process cleanup
- **Error Logging**: Detailed error logs for troubleshooting
- **User Confirmation**: Safety prompts before processing
- **File Format Support**: Works with .xls, .xlsx, and .xlsm files

## 🚀 Quick Start

### Prerequisites
- Windows OS
- Microsoft Excel installed
- PowerShell 5.1 or later

### Installation
1. Download `Fast_Excel_CSV_Converter.ps1`
2. Place it in your desired directory
3. Right-click and "Run with PowerShell" or execute via command line

### Basic Usage
```powershell
# Run the converter
.\Fast_Excel_CSV_Converter.ps1

# Check version
.\Fast_Excel_CSV_Converter.ps1 --version
```

## 💡 How It Works

### Step-by-Step Process
1. **File Selection**: Choose Excel files using the built-in file dialog
2. **Mode Selection**: Choose between Normal (formatted) or High-Speed (raw) conversion
3. **Safety Check**: Confirm before processing begins
4. **Batch Conversion**: All selected files and their sheets are processed
5. **Output Organization**: Results saved in timestamped folders

### Output Structure
```
📁 YourDirectory/
├── 📄 Fast_Excel_CSV_Converter.ps1
└── 📁 20250916-143052/  (timestamp folder)
    ├── 📄 File1-Sheet1-normal.csv
    ├── 📄 File1-Sheet2-normal.csv
    ├── 📄 File2-Data-highspeed.csv
    └── 📄 error.log (if errors occurred)
```

## 🔧 Advanced Options

### Conversion Modes Comparison
| Feature | Normal Mode | High-Speed Mode |
|---------|-------------|-----------------|
| **Speed** | Standard | Up to 10x faster |
| **Formatting** | ✅ Preserved | ❌ Raw values only |
| **Dates** | ✅ Human readable | ❌ Serial numbers |
| **Currency** | ✅ With symbols | ❌ Numbers only |
| **Best for** | Final reports, presentations | Data analysis, bulk processing |

### Command Line Options
```powershell
# Display version information
.\Fast_Excel_CSV_Converter.ps1 --version
.\Fast_Excel_CSV_Converter.ps1 -v
.\Fast_Excel_CSV_Converter.ps1 /version
```

## 🛠️ Troubleshooting

### Common Issues
- **"Excel process still running"**: The script automatically handles process cleanup
- **File access denied**: Ensure Excel files are closed before conversion
- **Large files taking too long**: Use High-Speed mode for better performance
- **Encoding issues**: Output uses UTF-8 encoding by default

### Error Logs
Check the `error.log` file in the output directory for detailed error information.

## 📋 System Requirements
- **OS**: Windows 7/8/10/11
- **PowerShell**: Version 5.1 or later
- **Excel**: Microsoft Excel (any recent version)
- **Memory**: Depends on file size (recommend 4GB+ RAM for large files)

## 🤝 Contributing
Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

## 📄 License
This project is licensed under the MIT License - see the LICENSE file for details.

## 👨‍💻 Authors
- **Ryo Osawa** - *Initial work*
- **Claude Sonnet 4.0** - *AI Assistant*

---

<a name="japanese"></a>
# 📖 日本語

## ✨ 機能

### 🎯 **2つの変換モード**
- **ノーマルモード**: セルの書式を完全保持（日付、通貨、パーセンテージ）
- **高速モード**: 生の値を使用した超高速変換（最大10倍高速）

### 📊 **スマート処理**
- **バッチ処理**: 複数のExcelファイルを一度に変換
- **全シート対応**: ワークブック内の全シートを自動変換
- **空シート処理**: 空のワークシートも適切に処理
- **エラー復旧**: 個別ファイルが失敗しても処理を継続

### 🛡️ **堅牢で安全**
- **プロセス管理**: Excelプロセスの自動クリーンアップ
- **エラーログ**: トラブルシューティング用の詳細ログ
- **ユーザー確認**: 処理前の安全確認プロンプト
- **ファイル形式対応**: .xls、.xlsx、.xlsmファイルに対応

## 🚀 クイックスタート

### 必要な環境
- Windows OS
- Microsoft Excel がインストール済み
- PowerShell 5.1 以降

### インストール
1. `Fast_Excel_CSV_Converter.ps1` をダウンロード
2. 任意のディレクトリに配置
3. 右クリックから「PowerShellで実行」またはコマンドラインから実行

### 基本的な使用方法
```powershell
# コンバーターを実行
.\Fast_Excel_CSV_Converter.ps1

# バージョン確認
.\Fast_Excel_CSV_Converter.ps1 --version
```

## 💡 動作原理

### ステップバイステップ処理
1. **ファイル選択**: 内蔵ファイルダイアログでExcelファイルを選択
2. **モード選択**: ノーマル（書式保持）または高速（生値）変換を選択
3. **安全確認**: 処理開始前の確認
4. **バッチ変換**: 選択されたファイルとシートをすべて処理
5. **出力整理**: タイムスタンプ付きフォルダに結果を保存

### 出力構造
```
📁 あなたのディレクトリ/
├── 📄 Fast_Excel_CSV_Converter.ps1
└── 📁 20250916-143052/  (タイムスタンプフォルダ)
    ├── 📄 File1-Sheet1-normal.csv
    ├── 📄 File1-Sheet2-normal.csv
    ├── 📄 File2-Data-highspeed.csv
    └── 📄 error.log (エラーが発生した場合)
```

## 🔧 高度なオプション

### 変換モード比較
| 機能 | ノーマルモード | 高速モード |
|------|----------------|------------|
| **速度** | 標準 | 最大10倍高速 |
| **書式** | ✅ 保持 | ❌ 生値のみ |
| **日付** | ✅ 人間が読める形式 | ❌ シリアル番号 |
| **通貨** | ✅ 記号付き | ❌ 数値のみ |
| **適用場面** | 最終レポート、プレゼン | データ解析、一括処理 |

### コマンドラインオプション
```powershell
# バージョン情報を表示
.\Fast_Excel_CSV_Converter.ps1 --version
.\Fast_Excel_CSV_Converter.ps1 -v
.\Fast_Excel_CSV_Converter.ps1 /version
```

## 🛠️ トラブルシューティング

### よくある問題
- **"Excelプロセスが残っている"**: スクリプトが自動的にプロセスクリーンアップを処理します
- **ファイルアクセス拒否**: 変換前にExcelファイルを閉じてください
- **大きなファイルの処理が遅い**: 高速モードを使用してパフォーマンスを向上させてください
- **エンコーディング問題**: 出力はデフォルトでUTF-8エンコーディングを使用します

### エラーログ
詳細なエラー情報については、出力ディレクトリの `error.log` ファイルを確認してください。

## 📋 システム要件
- **OS**: Windows 7/8/10/11
- **PowerShell**: バージョン5.1以降
- **Excel**: Microsoft Excel（任意の最新バージョン）
- **メモリ**: ファイルサイズに依存（大きなファイルには4GB以上のRAMを推奨）

## 🤝 コントリビューション
コントリビューションを歓迎します！バグやフィーチャーリクエストについては、お気軽にプルリクエストを送信したり、イシューを開いてください。

## 📄 ライセンス
このプロジェクトはMITライセンスの下でライセンスされています - 詳細についてはLICENSEファイルをご覧ください。

## 👨‍💻 作者
- **Ryo Osawa** - *初期開発*
- **Claude Sonnet 4.0** - *AIアシスタント*

---

## 🙏 Acknowledgments
Special thanks to the PowerShell and Excel communities for their continued support and inspiration.

## 📞 Support
If you encounter any issues or have questions, please feel free to open an issue on GitHub.

---
⭐ **Star this repository if it helped you!** ⭐
