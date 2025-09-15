# ⚡ Fast Excel CSV Converter

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

> 🚀 **High-performance Excel to CSV converter with intelligent optimization**  
> Convert your Excel files to CSV format while preserving exactly what you see in Excel - dates, percentages, currency, and all formatting intact!

---

## 🌏 Language / 言語選択

- [English](#english) | [日本語](#japanese)

---

## English

### ✨ Why This Tool?

Unlike other converters (including MarkItDown), this tool ensures **pixel-perfect accuracy**:

| Other Tools | This Tool |
|------------|-----------|
| `0.25` | `25%` ✅ |
| `44927` | `2023/1/1` ✅ |
| `1000` | `¥1,000` ✅ |

### 🌟 Key Features

- 🎯 **Zero Configuration** - Just run and convert!
- ⚡ **Intelligent Optimization** - Automatically selects the best processing strategy
- 📊 **Format Preservation** - Maintains dates, percentages, currency as displayed
- 🔄 **Batch Processing** - Convert multiple Excel files at once
- 💾 **Memory Efficient** - Handles large files with chunk-based processing
- 🛡️ **Safe Execution** - Proper Excel process management and cleanup
- 🌐 **UTF-8 Support** - Perfect for international characters

### 🚀 Quick Start

#### Prerequisites
- Windows OS
- Microsoft Excel installed
- PowerShell 5.1 or later

#### Installation & Usage

1. **Download the script**
   ```bash
   # Clone this repository
   git clone https://github.com/tokotoko1986/fast-excel-csv-converter.git
   cd fast-excel-csv-converter
   ```

2. **Set execution policy** (if needed)
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Run the converter**
   ```powershell
   # Right-click on the .ps1 file and select "Run with PowerShell"
   # OR run from PowerShell:
   .\Fast_Excel_CSV_Converter.ps1
   ```

4. **Follow the prompts**
   - Confirm that no Excel files are open
   - Select Excel files to convert
   - Watch the magic happen! ✨

### 📁 Output

Your CSV files will be organized in a timestamped folder:
```
📂 20241215-143022/
├── 📄 SalesData-Sheet1.csv
├── 📄 SalesData-Summary.csv
├── 📄 Inventory-Products.csv
└── 📄 error.log (if any issues occurred)
```

### 🎯 Performance Comparison

| File Size | Sheets | Processing Time | Memory Usage |
|-----------|--------|----------------|--------------|
| 5MB | 3 sheets | ~15 seconds | Low |
| 50MB | 10 sheets | ~2 minutes | Moderate |
| 200MB+ | 20+ sheets | ~8 minutes | Efficient chunking |

### 🧠 How It Works

The tool uses a **sophisticated 3-tier optimization strategy**:

1. **🔍 Smart Analysis** - Automatically detects data complexity
2. **⚡ Fast Mode** - For simple data (3-5x faster than standard tools)
3. **🎯 Precision Mode** - For formatted data (maintains visual accuracy)
4. **🔄 Chunk Mode** - For large datasets (memory-efficient processing)

### 🛡️ Safety Features

- **Pre-execution warning** about Excel process management
- **Automatic Excel cleanup** prevents hanging processes  
- **Error logging** for troubleshooting
- **Progress tracking** for long operations
- **Graceful degradation** continues processing even if some files fail

---

## Japanese

### ✨ なぜこのツールなのか？

他のコンバーター（MarkItDownを含む）とは異なり、このツールは**完璧な精度**を保証します：

| 他のツール | このツール |
|------------|-----------|
| `0.25` | `25%` ✅ |
| `44927` | `2023/1/1` ✅ |
| `1000` | `¥1,000` ✅ |

### 🌟 主な機能

- 🎯 **設定不要** - 実行するだけで変換完了！
- ⚡ **インテリジェント最適化** - 最適な処理戦略を自動選択
- 📊 **フォーマット保持** - 日付、パーセント、通貨などの表示形式を維持
- 🔄 **バッチ処理** - 複数のExcelファイルを一度に変換
- 💾 **メモリ効率** - チャンク処理で大容量ファイルにも対応
- 🛡️ **安全実行** - 適切なExcelプロセス管理とクリーンアップ
- 🌐 **UTF-8対応** - 日本語などの国際文字も完璧にサポート

### 🚀 クイックスタート

#### 前提条件
- Windows OS
- Microsoft Excel がインストール済み
- PowerShell 5.1 以上

#### インストール＆使用方法

1. **スクリプトのダウンロード**
   ```bash
   # このリポジトリをクローン
   git clone https://github.com/tokotoko1986/fast-excel-csv-converter.git
   cd fast-excel-csv-converter
   ```

2. **実行ポリシーの設定**（必要に応じて）
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **コンバーターの実行**
   ```powershell
   # .ps1ファイルを右クリックして「PowerShellで実行」を選択
   # または PowerShell から実行：
   .\Fast_Excel_CSV_Converter.ps1
   ```

4. **プロンプトに従って操作**
   - Excelファイルが開いていないことを確認
   - 変換したいExcelファイルを選択
   - 魔法を見守る！ ✨

### 📁 出力結果

CSVファイルはタイムスタンプ付きフォルダに整理されます：
```
📂 20241215-143022/
├── 📄 売上データ-シート1.csv
├── 📄 売上データ-サマリー.csv
├── 📄 在庫管理-商品.csv
└── 📄 error.log (エラーが発生した場合)
```

### 🎯 パフォーマンス比較

| ファイルサイズ | シート数 | 処理時間 | メモリ使用量 |
|---------------|---------|---------|-------------|
| 5MB | 3シート | ~15秒 | 少ない |
| 50MB | 10シート | ~2分 | 中程度 |
| 200MB以上 | 20シート以上 | ~8分 | 効率的なチャンク処理 |

### 🧠 動作原理

このツールは**洗練された3段階最適化戦略**を使用：

1. **🔍 スマート分析** - データの複雑さを自動検出
2. **⚡ 高速モード** - シンプルなデータ用（標準ツールの3-5倍高速）
3. **🎯 精密モード** - フォーマット済みデータ用（視覚的精度を維持）
4. **🔄 チャンクモード** - 大容量データセット用（メモリ効率重視）

### 🛡️ 安全機能

- **実行前警告** - Excelプロセス管理についての事前通知
- **自動Excelクリーンアップ** - ハングプロセスを防止
- **エラーログ** - トラブルシューティング用
- **進捗追跡** - 長時間操作の進行状況表示
- **グレースフル・デグラデーション** - 一部のファイルが失敗しても処理継続

### 🤝 使用上の注意

- **Excelが開いている場合は事前に閉じてください** - スクリプトが強制終了する可能性があります
- **大容量ファイルの場合は時間がかかる場合があります** - 進捗バーで状況を確認できます
- **エラーが発生した場合** - error.logファイルで詳細を確認できます

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Guidelines
- **Keep it simple** - this tool prioritizes simplicity over feature bloat
- **Maintain backward compatibility** with existing PowerShell versions
- **Test with various Excel file types** (.xls, .xlsx, .xlsm)
- **Document any changes** with clear commit messages
- **Follow the established error handling patterns**

### Reporting Issues
When reporting bugs, please include:
- PowerShell version (`$PSVersionTable.PSVersion`)
- Excel version
- Sample file characteristics (size, complexity)
- Full error message from error.log
- Steps to reproduce

## 🚀 Roadmap

Future enhancements being considered:
- **Parallel processing** for multiple worksheets
- **Advanced filtering options** for specific sheets/ranges
- **Configuration file support** for enterprise deployments
- **PowerShell Core support** for cross-platform compatibility

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with **PowerShell** and **Excel Interop**
- Inspired by the need for **accurate Excel-to-CSV conversion**
- Special thanks to the **PowerShell community** for best practices
- **Claude Sonnet 4.0** for development assistance and optimization strategies

---

<div align="center">

**⭐ If this tool saved you time, please give it a star! ⭐**  
**⭐ このツールで時間を節約できた場合は、ぜひスターをお願いします！ ⭐**

Made with ❤️ for the data processing community  
データ処理コミュニティのために ❤️ を込めて作成

[Report Bug](https://github.com/yourusername/fast-excel-csv-converter/issues) • [Request Feature](https://github.com/yourusername/fast-excel-csv-converter/issues) • [View Releases](https://github.com/yourusername/fast-excel-csv-converter/releases)

</div>
