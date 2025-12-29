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
- ✅ 异步导出和任务管理

## 快速开始

### 1. 启动服务

```bash
mvn spring-boot:run
```

### 2. 访问接口文档

启动后访问：http://localhost:8222/excel-service/doc.html

## 模板准备

### 1. 模板目录

默认模板目录：`/home/excel-templates/`

### 2. 模板结构

```excel
┌──────┬──────┬──────┬──────┐
│ 姓名  │ 年龄  │ 部门  │ 工资  │  <- 第1行：表头
├──────┼──────┼──────┼──────┤
│      │      │      │      │  <- 第2行：数据区（可包含占位符如 {userName}、{time} 等）
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

// 也可以使用文档级占位符，如 {time}、{date} 等
Row headerRow2 = sheet.createRow(0);
headerRow2.createCell(3).setCellValue("报表生成时间：{time}");

// 保存文件
String templatePath = "/home/excel-templates/员工统计表.xlsx";
try (FileOutputStream outputStream = new FileOutputStream(templatePath)) {
    workbook.write(outputStream);
}
```

## API接口说明

### 单Sheet导出

**接口地址**：`POST /excel-service/api/excel/export/single`

**请求参数**：

| 参数名 | 类型 | 必填 | 默认值 | 描述                                                          |
|--------|------|------|--------|-------------------------------------------------------------|
| templatePath | String | 是 | - | 模板路径，支持完整路径或模板名称                                            |
| data | List<Map> | 是 | - | 导出数据列表                                                      |
| startRow | Integer | 否 | 2 | 数据起始行                                                       |
| fieldMapping | Map | 否 | - | 字段映射关系                                                      |
| exportFileName | String | 否 | 自动生成 | 导出文件名                                                       |
| exportDir | String | 否 | 配置目录 | 导出目录                                                        |
| enableBigDataMode | Boolean | 否 | false | 是否启用大数据模式                                                   |
| batchSize | Integer | 否 | 5000 | 分批大小                                                        |
| maxRowsPerSheet | Integer | 否 | 50000 | 单个Sheet最大行数，超过此值将自动拆分到多个Sheet，仅在启用大数据模式时生效                  |
| sheetName | String | 否 | - | Sheet名称                                                     |
| placeholders | Map | 否 | - | 占位符数据，用于填充模板中的文档级占位符，支持多种格式如{title}、{date}、${time}、${year}等 |
| readOnly | Boolean | 否 | false | 设置导出的Excel文件为只读模式                                           |
| enableEncryption | Boolean | 否 | false | 是否对Excel文件进行密码加密(如果readOnly已经为true了,则无法设置密码,因为会导致写入文件失败)    |
| encryptionPassword | String | 否 | - | Excel文件加密密码，当enableEncryption为true时生效                       |

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
  "batchSize": 5000,
  "maxRowsPerSheet": 50000,
  "sheetName": "员工统计",
  "placeholders": {
    "reportTitle": "2023年度员工统计",
    "reportDate": "2023-12-25",
    "time": "14:30:00"
  },
  "readOnly": true,
  "enableEncryption": true,
  "encryptionPassword": "123456"
}
```

**响应示例**：

```json
{
  "success": true,
  "code": 200,
  "message": "导出成功",
  "data": {
    "filePath": "/home/excel-exports/2023年员工统计.xlsx",
    "fileName": "2023年员工统计.xlsx",
    "fileSize": 12345,
    "exportTime": "2023-12-25 10:30:45",
    "exportDuration": 1234,
    "exportDurationSeconds": 1.23
  }
}
```

**响应字段说明**：

| 字段名 | 类型 | 说明 |
|--------|------|------|
| success | Boolean | 是否成功 |
| code | Integer | 业务状态码 |
| message | String | 响应消息 |
| data | Object | 响应数据（成功时为ExcelFileInfo对象，失败时为""）|
| data.filePath | String | 导出文件完整路径 |
| data.fileName | String | 导出文件名 |
| data.fileSize | Long | 文件大小（字节） |
| data.exportTime | String | 导出时间 |
| data.exportDuration | Long | 导出耗时（毫秒） |
| data.exportDurationSeconds | BigDecimal | 导出耗时（秒，保留2位小数） |
| timestamp | Long | 响应时间戳（毫秒） |

### 单Sheet异步导出

**接口地址**：`POST /excel-service/api/excel/export/single/async`

**功能说明**：异步执行导出任务，返回任务ID，适合长时间运行的导出任务。

**请求参数**：同单Sheet导出接口

**响应示例**：

```json
{
  "success": true,
  "code": 200,
  "message": "操作成功",
  "data": "taskId1234567890",
  "timestamp": 1703232645000
}
```

### 查询异步导出任务状态

**接口地址**：`GET /excel-service/api/excel/task/{taskId}`

**功能说明**：查询异步导出任务的执行状态和进度信息。

**响应示例**：

```json
{
  "success": true,
  "code": 200,
  "message": "操作成功",
  "data": {
    "taskId": "taskId1234567890",
    "status": "COMPLETED",
    "statusDescription": "已完成",
    "result": {
      "filePath": "/home/excel-exports/2023年员工统计.xlsx",
      "fileName": "2023年员工统计.xlsx",
      "fileSize": 12345,
      "exportTime": "2023-12-25 10:30:45",
      "exportDuration": 1234,
      "exportDurationSeconds": 1.23
    }
  }
}
```

### 停止异步导出任务

**接口地址**：`DELETE /excel-service/api/excel/task/{taskId}`

**功能说明**：停止正在执行的异步导出任务。

**响应示例**：

```json
{
  "success": true,
  "code": 200,
  "message": "任务已停止",
  "data": "",
  "timestamp": 1703232645000
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
| batchSize | Integer | 否 | 5000 | 分批大小，大数据模式下生效 |

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

**基础大数据模式**：

```json
{
  "templatePath": "订单明细表",
  "data": [...100000条数据...],
  "enableBigDataMode": true,
  "batchSize": 5000,
  "exportFileName": "2023年订单明细.xlsx"
}

**自动多Sheet拆分**：当数据量超过单个Sheet的最大行数限制时，系统会自动将数据拆分到多个Sheet：

```json
{
  "templatePath": "订单明细表",
  "data": [...200000条数据...],
  "enableBigDataMode": true,
  "maxRowsPerSheet": 50000,  // 单个Sheet最大5万行，超过将自动拆分到多个Sheet
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

支持在模板中使用多种格式的占位符：

1. 数据行占位符：`{userName}`、`{age}` 等，用于填充数据行
2. 文档级占位符：`${time}`、`${year}`、`{time}`、`{year}` 等，用于填充文档标题或汇总信息

当数据只有一行且模板中包含相应占位符时，系统会自动识别并填充这些占位符。

### 4. 文件安全设置

支持对导出的Excel文件进行安全设置：

1. **只读模式**：设置 `readOnly` 为 `true`，导出的文件将为只读模式
2. **密码加密**：设置 `enableEncryption` 为 `true`，并提供 `encryptionPassword`，导出的文件将被密码保护

**请求示例**：

```json
{
  "templatePath": "敏感数据报表",
  "data": [...],
  "readOnly": true,
  "enableEncryption": true,
  "encryptionPassword": "123456"
}
```

### 5. 文档级占位符填充

当模板中包含文档级占位符（如 ${time}、${year}、{title} 等）时，可以使用以下方式填充：

**方式一：直接在数据中包含占位符**

```json
{
  "templatePath": "模板的文件名",
  "data": [
    {
      "year": 2025,
      "month": 12,
      "day": 26,
      "time": 14,
      "201KV": 201.2,
      "201A": 15.8,
      "202KV": 202.1,
      "202A": 16.3
    }
  ],
  "startRow": 2,
  "exportFileName": "导出的文件名.xlsx",
  "exportDir": "/home/custom-exports/"
}
```

**方式二：使用 placeholders 字段明确指定文档级占位符**

```json
{
  "templatePath": "综合报表",
  "data": [...],
  "placeholders": {
    "reportTitle": "2023年度报告",
    "reportDate": "2023-12-25",
    "time": "14:30:00",
    "year": 2025
  }
}
```

**方式三：同时使用数据和占位符**

```json
{
  "templatePath": "员工统计表",
  "data": [
    {
      "userName": "张三",
      "age": 25,
      "department": "研发部"
    }
  ],
  "placeholders": {
    "reportTitle": "2023年度员工统计",
    "reportDate": "2023-12-25",
    "time": "14:30:00"
  }
}
```

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

### Q: 如何使用异步导出功能？
A: 使用 `/export/single/async` 接口提交异步导出任务，获取任务ID后通过 `/task/{taskId}` 查询进度，或使用 `DELETE /task/{taskId}` 停止任务。