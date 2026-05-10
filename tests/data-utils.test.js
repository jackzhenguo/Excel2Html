const test = require("node:test");
const assert = require("node:assert/strict");
const utils = require("../src/data-utils.js");

test("normalizeRows 使用首行表头并过滤空白行", () => {
  const result = utils.normalizeRows([
    ["城市", "销售额", ""],
    ["上海", "1,200", "A"],
    ["", "", ""],
    ["北京", "980", "B"]
  ]);

  assert.deepEqual(result.headers, ["城市", "销售额", "字段3"]);
  assert.equal(result.rows.length, 2);
  assert.equal(result.rows[0]["销售额"], "1,200");
});

test("parseNumber 支持货币、百分号、千分位和括号负数", () => {
  assert.equal(utils.parseNumber("￥1,234.5"), 1234.5);
  assert.equal(utils.parseNumber("$98%"), 98);
  assert.equal(utils.parseNumber("(42)"), -42);
  assert.equal(Number.isNaN(utils.parseNumber("文本")), true);
});

test("detectNumericColumns 自动识别数值列", () => {
  const rows = [
    { 城市: "上海", 销售额: "1,200", 备注: "稳定" },
    { 城市: "北京", 销售额: "980", 备注: "增长" },
    { 城市: "深圳", 销售额: "1,450", 备注: "增长" }
  ];

  assert.deepEqual(utils.detectNumericColumns(rows, ["城市", "销售额", "备注"]), ["销售额"]);
});

test("buildChartData 按分类字段汇总数值", () => {
  const rows = [
    { 城市: "上海", 销售额: "100" },
    { 城市: "上海", 销售额: "150" },
    { 城市: "北京", 销售额: "90" },
    { 城市: "杭州", 销售额: "无" }
  ];

  const data = utils.buildChartData(rows, "城市", "销售额", "bar");
  assert.deepEqual(data.labels, ["上海", "北京"]);
  assert.deepEqual(data.values, [250, 90]);
  assert.equal(data.totalGroups, 2);
});

test("getPreviewRows 默认只返回前 50 行", () => {
  const rows = Array.from({ length: 60 }, (_, index) => ({ id: index + 1 }));
  assert.equal(utils.getPreviewRows(rows).length, 50);
});
