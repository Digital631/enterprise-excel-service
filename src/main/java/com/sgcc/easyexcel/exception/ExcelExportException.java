package com.sgcc.easyexcel.exception;

import com.sgcc.easyexcel.enums.ExcelErrorCode;
import lombok.Getter;

/**
 * Excel导出自定义异常
 *
 * @author system
 * @since 2023-12-22
 */
@Getter
public class ExcelExportException extends RuntimeException {

    /**
     * 错误码
     */
    private final ExcelErrorCode errorCode;

    /**
     * 详细错误信息
     */
    private final String detailMessage;

    public ExcelExportException(ExcelErrorCode errorCode) {
        super(errorCode.getMessage());
        this.errorCode = errorCode;
        this.detailMessage = errorCode.getMessage();
    }

    public ExcelExportException(ExcelErrorCode errorCode, String detailMessage) {
        super(detailMessage);
        this.errorCode = errorCode;
        this.detailMessage = detailMessage;
    }

    public ExcelExportException(ExcelErrorCode errorCode, Throwable cause) {
        super(errorCode.getMessage(), cause);
        this.errorCode = errorCode;
        this.detailMessage = errorCode.getMessage();
    }

    public ExcelExportException(ExcelErrorCode errorCode, String detailMessage, Throwable cause) {
        super(detailMessage, cause);
        this.errorCode = errorCode;
        this.detailMessage = detailMessage;
    }
}
