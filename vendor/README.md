# Vendor 依赖

本工具为了做到用户直接打开浏览器即可使用，将浏览器端依赖放在 `vendor/` 目录中。

- `xlsx.full.min.js`：SheetJS Community Edition，用于在浏览器中解析 Excel/CSV 文件。
- `chart.umd.min.js`：Chart.js UMD 构建，用于绘制柱状图、折线图和饼图。

依赖文件仅在浏览器本地执行，不会上传用户数据。
