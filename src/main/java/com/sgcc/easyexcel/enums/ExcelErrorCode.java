package com.sgcc.easyexcel.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * Excel导出错误码枚举
 *
 * @author system
 * @since 2023-12-22
 */
@Getter
@AllArgsConstructor
public enum ExcelErrorCode {

    // ========== 成功 ==========
    SUCCESS(200, "操作成功"),

    // ========== 参数错误 400-499 ==========
    PARAM_ERROR(400, "请求参数错误"),
    TEMPLATE_PATH_EMPTY(401, "模板路径不能为空"),
    DATA_EMPTY(402, "导出数据不能为空"),
    INVALID_START_ROW(403, "起始行参数无效，必须大于0"),
    INVALID_BATCH_SIZE(404, "分批大小参数无效，建议1000-10000"),
    INVALID_FILE_NAME(405, "文件名不合法，必须以.xlsx或.xls结尾"),
    SHEET_INDEX_INVALID(406, "Sheet索引无效"),
    FIELD_MAPPING_ERROR(407, "字段映射配置错误"),

    // ========== 模板错误 4100-4199 ==========
    TEMPLATE_NOT_FOUND(4101, "模板文件不存在"),
    TEMPLATE_READ_ERROR(4102, "模板文件读取失败"),
    TEMPLATE_FORMAT_ERROR(4103, "模板文件格式错误，仅支持.xlsx/.xls"),
    TEMPLATE_DAMAGED(4104, "模板文件已损坏"),
    TEMPLATE_NO_PERMISSION(4105, "无权限读取模板文件"),
    TEMPLATE_HEADER_ERROR(4106, "模板表头解析失败"),
    TEMPLATE_PROCESS_ERROR(4107, "模板预处理失败"),

    // ========== 数据错误 4200-4299 ==========
    DATA_FIELD_MISMATCH(4201, "数据字段与模板表头不匹配"),
    DATA_TYPE_ERROR(4202, "数据类型错误"),
    DATA_TOO_LARGE(4203, "数据量过大，超过最大限制"),
    DATA_STRUCTURE_ERROR(4204, "数据结构错误"),
    DATA_MISMATCH(4205, "数据与模板不匹配"),

    // ========== 导出错误 4300-4399 ==========
    EXPORT_DIR_NOT_EXIST(4301, "导出目录不存在"),
    EXPORT_DIR_NO_PERMISSION(4302, "无权限写入导出目录"),
    FILE_WRITE_ERROR(4303, "文件写入失败"),
    FILE_ALREADY_EXISTS(4304, "文件已存在"),
    DISK_SPACE_INSUFFICIENT(4305, "磁盘空间不足"),
    TASK_INTERRUPTED(4306, "任务已中断"),

    // ========== 系统错误 500-599 ==========
    SYSTEM_ERROR(500, "系统内部错误"),
    IO_ERROR(501, "IO操作异常"),
    OUT_OF_MEMORY(502, "内存溢出，请启用大数据模式"),
    THREAD_INTERRUPTED(503, "线程中断"),
    CONCURRENT_LIMIT(504, "并发请求超限，请稍后重试"),
    UNKNOWN_ERROR(599, "未知错误");

    /**
     * 错误码
     */
    private final Integer code;

    /**
     * 错误信息
     */
    private final String message;
}
