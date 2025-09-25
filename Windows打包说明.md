# VAT验证工具 - Windows EXE打包说明

## 📦 概述

本项目已配置完整的Windows EXE打包流程，支持通过GitHub Actions自动构建或本地构建。

## 🔧 最新修复 (2024-09-25)

- ✅ **修复图标问题**：解决了PyInstaller构建时找不到图标文件的错误
- ✅ **添加应用图标**：创建了专业的SVG图标文件 (`icon.svg`)
- ✅ **优化构建流程**：改进了GitHub Actions配置，增加了错误处理
- ✅ **跨平台支持**：确保Windows、Linux、macOS都能正常构建

## 📦 获取Windows可执行文件的方法

### 方法一：GitHub Actions自动构建（推荐）

1. **推送代码到GitHub**：
   ```bash
   git add .
   git commit -m "更新Windows打包配置"
   git push origin main
   ```

2. **查看构建结果**：
   - 访问GitHub仓库的 `Actions` 标签页
   - 等待构建完成（通常需要5-10分钟）
   - 下载生成的 `VAT验证工具-Windows-EXE` 文件

3. **发布版本**（可选）：
   - 创建新的Release标签
   - GitHub Actions会自动将EXE文件附加到Release中

### 方法二：本地交叉编译（需要Windows环境）

如果您有Windows机器或虚拟机：

1. **安装Python和依赖**：
   ```bash
   pip install -r requirements.txt
   ```

2. **运行构建脚本**：
   ```bash
   python build_windows.py
   ```

3. **或直接使用PyInstaller**：
   ```bash
   pyinstaller "VAT验证工具-Windows.spec"
   ```

### 方法三：使用Docker（高级用户）

使用Wine在Linux/macOS上模拟Windows环境：

```bash
# 使用专门的Docker镜像
docker run --rm -v "$(pwd)":/src cdrx/pyinstaller-windows:python3 \
    "pyinstaller VAT验证工具-Windows.spec"
```

## 🔧 配置文件说明

### Windows专用配置 (`VAT验证工具-Windows.spec`)

- **优化的隐藏导入**：包含所有必要的Python模块
- **排除不需要的模块**：减小文件大小
- **无控制台模式**：运行时不显示命令行窗口
- **UPX压缩**：进一步减小文件大小

### GitHub Actions配置 (`.github/workflows/build-exe.yml`)

- **多平台构建**：Windows、Linux、macOS
- **自动化流程**：推送代码即可触发构建
- **文件上传**：构建完成后自动上传到Artifacts

## 📋 构建特性

✅ **完整功能**：包含所有标签页和功能模块  
✅ **优化大小**：排除不必要的依赖  
✅ **无依赖**：单文件可执行程序  
✅ **调试信息**：包含启动日志和错误处理  
✅ **跨平台**：支持Windows、Linux、macOS  

## 🚀 快速开始

1. 确保代码已推送到GitHub
2. 访问仓库的Actions页面
3. 等待构建完成
4. 下载Windows EXE文件
5. 在Windows系统上运行

## ⚠️ 注意事项

- Windows Defender可能会误报，这是PyInstaller打包程序的常见问题
- 首次运行可能需要几秒钟的启动时间
- 确保Windows系统已安装必要的Visual C++运行库

## 🔍 故障排除

如果遇到问题：

1. **检查构建日志**：在GitHub Actions中查看详细日志
2. **本地测试**：使用 `python vat_validator_gui.py` 确认代码正常运行
3. **依赖检查**：确认 `requirements.txt` 包含所有必要依赖
4. **重新构建**：删除 `build` 和 `dist` 目录后重新打包