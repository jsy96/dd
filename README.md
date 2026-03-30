# 舱单数据处理系统

自动提取舱单数据并生成提单确认件和装箱单发票的 Web 应用。

## 功能特点

1. **舱单数据自动提取** - 上传 Excel 格式的舱单文件，自动解析关键信息
2. **提单确认件生成** - 根据舱单数据和用户输入，生成 Word 格式的提单确认件
3. **装箱单发票生成** - 生成 Excel 格式的装箱单发票
4. **友好的 Web 界面** - 拖拽上传、实时预览、可视化编辑

## 安装依赖

```bash
pip install flask flask-cors python-docx xlrd xlwt xlutils
```

## 使用方法

### 1. 启动服务器

**Windows:**
```bash
start.bat
```

**或直接运行:**
```bash
python app.py
```

服务器将在 http://localhost:5000 启动

### 2. 使用流程

1. **上传舱单文件**
   - 点击上传区域或拖拽 `.xls` 格式的舱单文件
   - 系统会自动提取以下信息：
     - 船名、航次、目的港
     - 提单号、箱号、封号、箱型
     - 件数、毛重、体积
     - 品名、唛头等

2. **预览提取的数据**
   - 查看"2. 提取的数据预览"区域
   - 确认数据是否正确

3. **填写额外信息**
   - **装箱单发票信息**：发票号、发票日期
   - **提单确认件信息**：发货人、收货人、通知人
   - **商品明细**：数量、单位、品名、单价、金额

4. **生成文档**
   - 点击"生成提单确认件和装箱单发票"按钮
   - 系统将生成两个文档文件

## 支持的文件格式

- **输入**: `.xls` 格式的舱单文件
- **输出**:
  - 提单确认件: `.docx` 格式
  - 装箱单发票: `.xls` 格式

## 舱单文件格式要求

舱单 Excel 文件应包含以下关键字段：

| 字段 | 说明 |
|------|------|
| 船名 | 船舶名称 |
| 航次 | 航行编号 |
| 目的港 | 目的港口 |
| 总提单号/提单号 | 提单编号 |
| 箱号 | 集装箱号 |
| 封号 | 铅封号 |
| 箱型 | 集装箱类型 (如 20GP, 40HQ) |
| 件数 | 货物件数 |
| 包装单位 | CARTONS 等 |
| 毛重 | KGS |
| 体积 | CBM |
| 英文品名 | 货物描述 |
| 唛头 | Marks & Numbers |

## API 接口

### POST /api/preview

预览提取的舱单数据

**请求**: FormData with `manifest_file`

**响应**:
```json
{
  "success": true,
  "data": {
    "vessel_name": "CSCL GLOBE",
    "voyage_no": "074W",
    "port_of_discharge": "JEBEL ALI",
    ...
  }
}
```

### POST /api/process

处理舱单并生成文档

**请求**: FormData
- `manifest_file`: 舱单文件
- `invoice_no`: 发票号
- `invoice_date`: 发票日期
- `consignor`: 发货人信息
- `consignee`: 收货人信息
- `notify_party`: 通知人信息
- `items`: 商品明细 (每行格式: 数量|单位|品名|单价|金额)

**响应**:
```json
{
  "success": true,
  "message": "文档生成成功！",
  "extracted_data": {...}
}
```

## 技术栈

- **后端**: Python Flask
- **前端**: HTML/CSS/JavaScript
- **文档生成**: python-docx, xlwt
- **Excel 读取**: xlrd

## 项目结构

```
logs/
├── app.py              # Flask 后端应用
├── start.bat           # 启动脚本
├── templates/
│   └── index.html      # 前端页面
├── 舱单的格式.xls       # 舱单格式参考
├── 装箱单发票的格式.xls  # 装箱单发票格式参考
└── 提单确认件的格式.doc # 提单确认件格式参考
```

## 注意事项

1. 舱单文件必须是 `.xls` 格式（Excel 97-2003）
2. 确保所有必要的 Python 依赖已安装
3. 发货人、收货人信息每行一项
4. 商品明细格式: 数量|单位|品名|单价|金额

## 故障排除

**问题**: 无法读取 Excel 文件
- 检查文件是否为 `.xls` 格式
- 安装 xlrd: `pip install xlrd`

**问题**: 文档生成失败
- 检查 python-docx 是否安装: `pip install python-docx`
- 确保输入数据格式正确

**问题**: 中文显示乱码
- 确保系统支持中文字体
- 文档中使用宋体作为默认中文字体
