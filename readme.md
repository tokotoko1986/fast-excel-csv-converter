# ⚡ Fast Excel CSV Converter

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

> 🚀 **High-performance Excel to CSV converter with intelligent optimization**  
> Convert your Excel files to CSV format while preserving exactly what you see in Excel - dates, percentages, currency, and all formatting intact!

## ✨ Why This Tool?

Unlike other converters (including MarkItDown), this tool ensures **pixel-perfect accuracy**:

| Other Tools | This Tool |
|------------|-----------|
| `0.25` | `25%` ✅ |
| `44927` | `2023/1/1` ✅ |
| `1000` | `¥1,000` ✅ |

## 🌟 Key Features

- 🎯 **Zero Configuration** - Just run and convert!
- ⚡ **Intelligent Optimization** - Automatically selects the best processing strategy
- 📊 **Format Preservation** - Maintains dates, percentages, currency as displayed
- 🔄 **Batch Processing** - Convert multiple Excel files at once
- 💾 **Memory Efficient** - Handles large files with chunk-based processing
- 🛡️ **Safe Execution** - Proper Excel process management and cleanup
- 🌐 **UTF-8 Support** - Perfect for international characters

## 🚀 Quick Start

### Prerequisites
- Windows OS
- Microsoft Excel installed
- PowerShell 5.1 or later

### Installation & Usage

1. **Download the script**
   ```bash
   # Clone this repository
   git clone https://github.com/yourusername/fast-excel-csv-converter.git
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

## 📁 Output

Your CSV files will be organized in a timestamped folder:
```
📂 20241215-143022/
├── 📄 SalesData-Sheet1.csv
├── 📄 SalesData-Summary.csv
├── 📄 Inventory-Products.csv
└── 📄 error.log (if any issues occurred)
```

## 🎯 Performance Comparison

| File Size | Sheets | Processing Time | Memory Usage |
|-----------|--------|----------------|--------------|
| 5MB | 3 sheets | ~15 seconds | Low |
| 50MB | 10 sheets | ~2 minutes | Moderate |
| 200MB+ | 20+ sheets | ~8 minutes | Efficient chunking |

## 🧠 How It Works

The tool uses a **sophisticated 3-tier optimization strategy**:

1. **🔍 Smart Analysis** - Automatically detects data complexity
2. **⚡ Fast Mode** - For simple data (3-5x faster than standard tools)
3. **🎯 Precision Mode** - For formatted data (maintains visual accuracy)
4. **🔄 Chunk Mode** - For large datasets (memory-efficient processing)

## 🛡️ Safety Features

- **Pre-execution warning** about Excel process management
- **Automatic Excel cleanup** prevents hanging processes  
- **Error logging** for troubleshooting
- **Progress tracking** for long operations
- **Graceful degradation** continues processing even if some files fail

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

Made with ❤️ for the data processing community

[Report Bug](https://github.com/yourusername/fast-excel-csv-converter/issues) • [Request Feature](https://github.com/yourusername/fast-excel-csv-converter/issues) • [View Releases](https://github.com/yourusername/fast-excel-csv-converter/releases)

</div>