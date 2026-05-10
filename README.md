# 本地 Excel 数据可视化工具

一个无需后端的纯前端网页工具。用户在浏览器中打开 `index.html` 后，可以上传 Excel 文件，预览表格数据，并自动生成柱状图、折线图和饼图。

## 功能

- 支持上传 `.xlsx`、`.xls`、`.csv` 文件。
- 自动读取工作表，并支持切换不同 sheet。
- 自动识别数值列，过滤空值和无法解析的文本。
- 支持选择分类字段、数值字段和图表类型。
- 支持柱状图、折线图、饼图。
- 支持多套图表配色主题、渐变图形、圆角柱和图表摘要指标。
- 本地浏览器解析文件，不上传数据，不依赖后端。

## 文件结构

```text
local-excel-visualizer/
├─ index.html
├─ package.json
├─ README.md
├─ src/
│  ├─ app.js
│  ├─ data-utils.js
│  └─ styles.css
├─ tests/
│  └─ data-utils.test.js
└─ vendor/
   ├─ README.md
   ├─ chart.umd.min.js
   └─ xlsx.full.min.js
```

## 运行方法

直接双击 `index.html`，或在浏览器地址栏打开该文件即可使用。

如果希望通过本地静态服务访问，也可以在项目目录运行：

```bash
python -m http.server 8080
```

然后访问：

```text
http://localhost:8080
```

## 测试

```bash
npm test
```

测试覆盖了数值列识别、空值过滤、图表数据生成和表格预览切片等核心逻辑。
