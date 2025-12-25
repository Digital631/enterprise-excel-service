# Excel导出服务使用指南

## 概述

企业级通用Excel导出服务，支持基于模板的Excel导出，具备大数据量处理能力，保留模板公式和样式，支持高并发访问。

## 功能特性

- ✅ 基于模板的Excel导出
- ✅ 支持大数据量（百万级）
- ✅ 高并发处理（50+并发）
- ✅ 保留模板公式和样式
- ✅ 灵活的字段映射
- ✅ 支持多Sheet导出
- ✅ 支持占位符填充

## 快速开始

### 1. 启动服务

```bash
mvn spring-boot:run
```

### 2. 访问接口文档

启动后访问：http://localhost:8080/excel-service/doc.html

## 模板准备

### 1. 模板目录

默认模板目录：`/home/easyexcel/excel-templates/`

### 2. 模板结构

```excel
┌──────┬──────┬──────┬──────┐
│ 姓名  │ 年龄  │ 部门  │ 工资  │  <- 第1行：表头
├──────┼──────┼──────┼──────┤
│      │      │      │      │  <- 第2行：数据区（可包含占位符如 {userName}）
├──────┼──────┼──────┼──────┤
│      │      │      │      │  <- 第3行及以后：数据区
└──────┴──────┴──────┴──────┘
```

### 3. 模板创建示例

```java
// 创建模板文件
Workbook workbook = new XSSFWorkbook();
Sheet sheet = workbook.createSheet("员工统计");

// 创建表头行
Row headerRow = sheet.createRow(0);
headerRow.createCell(0).setCellValue("姓名");
headerRow.createCell(1).setCellValue("年龄");
headerRow.createCell(2).setCellValue("工资");

// 创建占位符行（第2行，Excel中是索引1）
Row placeholderRow = sheet.createRow(1);
placeholderRow.createCell(0).setCellValue("{userName}");
placeholderRow.createCell(1).setCellValue("{age}");
placeholderRow.createCell(2).setCellValue("{salary}");

// 保存文件
String templatePath = "/home/easyexcel/excel-templates/员工统计表.xlsx";
try (FileOutputStream outputStream = new FileOutputStream(templatePath)) {
    workbook.write(outputStream);
}
```

## API接口说明

### 单Sheet导出

**接口地址**：`POST /excel-service/api/excel/export/single`

**请求参数**：

| 参数名 | 类型 | 必填 | 默认值 | 描述 |
|--------|------|------|--------|------|
| templatePath | String | 是 | - | 模板路径，支持完整路径或模板名称 |
| data | List<Map> | 是 | - | 导出数据列表 |
| startRow | Integer | 否 | 2 | 数据起始行 |
| fieldMapping | Map | 否 | - | 字段映射关系 |
| exportFileName | String | 否 | 自动生成 | 导出文件名 |
| exportDir | String | 否 | 配置目录 | 导出目录 |
| enableBigDataMode | Boolean | 否 | false | 是否启用大数据模式 |
| batchSize | Integer | 否 | 5000 | 分批大小 |

**请求示例**：

```json
{
  "templatePath": "员工统计表",
  "data": [
    {
      "userName": "张三",
      "age": 25,
      "department": "研发部",
      "salary": 10000
    },
    {
      "userName": "李四",
      "age": 30,
      "department": "市场部",
      "salary": 12000
    }
  ],
  "startRow": 2,
  "fieldMapping": {
    "姓名": "userName",
    "年龄": "age",
    "部门": "department",
    "工资": "salary"
  },
  "exportFileName": "2023年员工统计.xlsx",
  "exportDir": "/home/custom-exports/",
  "enableBigDataMode": false,
  "batchSize": 5000
}
```

**响应示例**：

```json
{
  "success": true,
  "code": 200,
  "message": "导出成功",
  "data": {
    "filePath": "/home/easyexcel/excel-exports/2023年员工统计.xlsx",
    "fileName": "2023年员工统计.xlsx",
    "fileSize": 12345,
    "exportTime": "2023-12-25 10:30:45",
    "exportDuration": 1234,
    "exportDurationSeconds": 1.23
  }
}
```

### 多Sheet导出

**接口地址**：`POST /excel-service/api/excel/export/multi`

**请求参数**：

| 参数名 | 类型 | 必填 | 默认值 | 描述 |
|--------|------|------|--------|------|
| templatePath | String | 是 | - | 模板路径 |
| sheetDataList | List<SheetDataConfig> | 是 | - | Sheet数据配置列表 |

**SheetDataConfig参数**：

| 参数名 | 类型 | 必填 | 默认值 | 描述 |
|--------|------|------|--------|------|
| sheetIndex | Integer | 是 | - | Sheet索引 |
| sheetName | String | 否 | - | Sheet名称 |
| data | List<Map> | 是 | - | 该Sheet的数据 |
| startRow | Integer | 否 | 2 | 数据起始行 |
| fieldMapping | Map | 否 | - | 字段映射关系 |

**请求示例**：

```json
{
  "templatePath": "综合报表",
  "sheetDataList": [
    {
      "sheetIndex": 0,
      "sheetName": "员工信息",
      "data": [
        {
          "userName": "张三",
          "age": 25,
          "department": "研发部"
        }
      ],
      "startRow": 2,
      "fieldMapping": {
        "姓名": "userName",
        "年龄": "age",
        "部门": "department"
      }
    },
    {
      "sheetIndex": 1,
      "sheetName": "部门统计",
      "data": [
        {
          "deptName": "研发部",
          "count": 50
        }
      ],
      "startRow": 2,
      "fieldMapping": {
        "部门名称": "deptName",
        "人数": "count"
      }
    }
  ]
}
```

## 高级功能

### 1. 大数据量处理

当数据量超过1万条时，建议启用大数据模式：

```json
{
  "templatePath": "订单明细表",
  "data": [...100000条数据...],
  "enableBigDataMode": true,
  "batchSize": 5000,
  "exportFileName": "2023年订单明细.xlsx"
}
```

### 2. 字段映射

当模板表头与数据字段名不一致时，使用字段映射：

```json
{
  "fieldMapping": {
    "姓名": "userName",
    "年龄": "age",
    "工资": "salary"
  }
}
```

### 3. 占位符填充

支持在模板中使用占位符，如 `{userName}`、`{age}` 等。

## 性能指标

- 1万以内数据：响应时间 < 3秒
- 10万数据（大数据模式）：响应时间 < 30秒
- 支持50个并发请求

## 错误码说明

| 错误码 | 说明 |
|--------|------|
| 200 | 操作成功 |
| 400 | 请求参数错误 |
| 401 | 模板路径不能为空 |
| 402 | 导出数据不能为空 |
| 403 | 起始行参数无效，必须大于0 |
| 404 | 分批大小参数无效，建议1000-10000 |
| 405 | 文件名不合法，必须以.xlsx或.xls结尾 |
| 406 | Sheet索引无效 |
| 407 | 字段映射配置错误 |
| 4101 | 模板文件不存在 |
| 4102 | 模板文件读取失败 |
| 4103 | 模板文件格式错误，仅支持.xlsx/.xls |
| 4104 | 模板文件已损坏 |
| 4105 | 无权限读取模板文件 |
| 4106 | 模板表头解析失败 |
| 4107 | 模板预处理失败 |
| 4201 | 数据字段与模板表头不匹配 |
| 4202 | 数据类型错误 |
| 4203 | 数据量过大，超过最大限制 |
| 4204 | 数据结构错误 |
| 4205 | 数据与模板不匹配 |
| 4301 | 导出目录不存在 |
| 4302 | 无权限写入导出目录 |
| 4303 | 文件写入失败 |
| 4304 | 文件已存在 |
| 4305 | 磁盘空间不足 |
| 4306 | 任务已中断 |
| 500 | 系统内部错误 |
| 501 | IO操作异常 |
| 502 | 内存溢出，请启用大数据模式 |
| 503 | 线程中断 |
| 504 | 并发请求超限，请稍后重试 |
| 599 | 未知错误 |

## 常见问题

### Q: 为什么导出的Excel第二行是空白的？
A: 请确保startRow参数设置正确。默认值为2，表示从Excel的第2行开始填充数据。

### Q: 如何处理大数据量导出？
A: 当数据量超过1万条时，请将enableBigDataMode设置为true，系统将自动采用分批写入方式防止内存溢出。

### Q: 如何自定义导出目录？
A: 在请求参数中指定exportDir字段，如 `"exportDir": "/home/custom-exports/"`。

### Q: 如何使用字段映射？
A: 当模板表头与数据字段名不一致时，使用fieldMapping参数建立映射关系，如 `{"姓名": "userName"}`。