# VAT验证工具

这是一个用于验证欧盟VAT号码的工具，使用VIES（VAT Information Exchange System）API进行验证。

## 功能特性

### 核心功能
- ✅ 单个VAT号码验证
- ✅ 批量VAT号码验证（从Excel文件导入）
- ✅ 自动识别VAT号码的国家代码
- ✅ 验证结果包含公司名称和地址信息
- ✅ 导出验证结果到Excel文件

### 新增功能
- 🆕 请求日志功能（实时显示API请求和响应）
- 🆕 日志导出为txt文件
- 🆕 验证结果文本搜索功能
- 🆕 现代化的图形用户界面

## 支持的国家

- 奥地利 (AT)
- 比利时 (BE)
- 保加利亚 (BG)
- 克罗地亚 (HR)
- 塞浦路斯 (CY)
- 捷克 (CZ)
- 丹麦 (DK)
- 爱沙尼亚 (EE)
- 芬兰 (FI)
- 法国 (FR)
- 德国 (DE)
- 希腊 (EL)
- 匈牙利 (HU)
- 爱尔兰 (IE)
- 意大利 (IT)
- 拉脱维亚 (LV)
- 立陶宛 (LT)
- 卢森堡 (LU)
- 马耳他 (MT)
- 荷兰 (NL)
- 波兰 (PL)
- 葡萄牙 (PT)
- 罗马尼亚 (RO)
- 斯洛伐克 (SK)
- 斯洛文尼亚 (SI)
- 西班牙 (ES)
- 瑞典 (SE)
- 北爱尔兰 (XI)

## 安装和运行

### 方法1: 直接运行Python脚本

1. 确保已安装Python 3.7+
2. 安装依赖包:
   ```bash
   pip install -r requirements.txt
   ```
3. 运行程序:
   ```bash
   python vat_validator.py
   ```

### 方法2: 打包成exe文件

1. 安装依赖包:
   ```bash
   pip install -r requirements.txt
   ```
2. 运行打包脚本:
   ```bash
   python build_exe.py
   ```
3. 在`dist`目录中找到生成的exe文件

## 使用说明

### 单个VAT号码验证

1. 从下拉菜单中选择国家/地区
2. 输入要验证的VAT号码
3. 点击"验证"按钮
4. 查看验证结果

### 批量验证

1. 准备Excel文件，确保包含以下列:
   - `NIF Contraparte`: VAT号码列
   - `Importe`: 金额列
   - `Tipo`: 类型列

2. 点击"导入Excel文件"按钮选择文件
3. 程序会自动开始批量验证
4. 验证完成后，点击"导出结果"保存验证结果

### Excel文件格式示例

| NIF Contraparte | Importe | Tipo | Name |
|----------------|---------|------|------|
| IT05159640266  | 15.31   | E    |      |
| PL5263222338   | -0.41   | E    |      |
| IT07730610966  | 13.11   | E    |      |

## API说明

程序调用欧盟官方VIES REST API:
```
https://ec.europa.eu/taxation_customs/vies/rest-api/ms/{country_code}/vat/{vat_number}
```

### 请求示例
```
GET https://ec.europa.eu/taxation_customs/vies/rest-api/ms/IT/vat/IT05159640266
```

### 响应示例
```json
{
  "isValid": true,
  "requestDate": "2024-01-15+01:00",
  "userError": "VALID",
  "name": "COMPANY NAME",
  "address": "COMPANY ADDRESS",
  "requestIdentifier": "...",
  "valid": true,
  "traderName": "COMPANY NAME",
  "traderCompanyType": "...",
  "traderAddress": "COMPANY ADDRESS",
  "requestDateString": "15/01/2024"
}
```

## 技术栈

- **Python 3.7+**: 主要编程语言
- **tkinter**: GUI界面框架
- **requests**: HTTP请求库
- **pandas**: 数据处理
- **openpyxl**: Excel文件处理
- **PyInstaller**: 打包工具

## 项目结构

```
vies-on-the-soft/
├── vat_validator.py    # 主程序文件
├── build_exe.py        # 打包脚本
├── requirements.txt    # 依赖包列表
└── README.md          # 项目说明
```

## 注意事项

1. **网络连接**: 程序需要互联网连接来访问VIES API
2. **API限制**: 欧盟VIES API可能有访问频率限制
3. **VAT号码格式**: 确保输入正确的VAT号码格式
4. **国家代码**: 程序会自动尝试从VAT号码中提取国家代码

## 故障排除

### 常见问题

1. **网络错误**: 检查网络连接和防火墙设置
2. **API超时**: 可能是VIES服务器繁忙，请稍后重试
3. **Excel文件错误**: 确保Excel文件格式正确，包含必要的列
4. **打包失败**: 确保已安装所有依赖包

### 错误代码说明

- `VALID`: VAT号码有效
- `INVALID`: VAT号码无效
- `INVALID_INPUT`: 输入格式错误
- `SERVICE_UNAVAILABLE`: 服务不可用
- `MS_UNAVAILABLE`: 成员国服务不可用
- `TIMEOUT`: 请求超时

## 许可证

本项目仅供学习和研究使用。请遵守欧盟VIES服务的使用条款。

## 更新日志

### v1.0.0 (2024-01-15)
- 初始版本发布
- 支持单个和批量VAT验证
- 支持Excel导入导出
- 图形用户界面
- 支持打包成exe文件