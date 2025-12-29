package com.sgcc.easyexcel.exception;

import com.sgcc.easyexcel.dto.response.ExcelExportResponse;
import com.sgcc.easyexcel.enums.ExcelErrorCode;
import lombok.extern.slf4j.Slf4j;
import org.springframework.validation.BindException;
import org.springframework.validation.FieldError;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

import java.util.stream.Collectors;

/**
 * 全局异常处理器
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@RestControllerAdvice
public class GlobalExceptionHandler {

    /**
     * 处理Excel导出异常
     */
    @ExceptionHandler(ExcelExportException.class)
    public ExcelExportResponse handleExcelExportException(ExcelExportException e) {
        log.error("Excel导出异常：code={}, message={}", e.getErrorCode().getCode(), e.getDetailMessage(), e);
        
        return ExcelExportResponse.error(
                e.getErrorCode().getCode(),
                e.getDetailMessage()
        );
    }

    /**
     * 处理参数校验异常（@Valid）
     */
    @ExceptionHandler(MethodArgumentNotValidException.class)
    public ExcelExportResponse handleValidationException(MethodArgumentNotValidException e) {
        String errorMessage = e.getBindingResult().getFieldErrors().stream()
                .map(FieldError::getDefaultMessage)
                .collect(Collectors.joining("; "));
        
        log.error("参数校验失败：{}", errorMessage);
        
        return ExcelExportResponse.error(
                ExcelErrorCode.PARAM_ERROR.getCode(),
                "参数校验失败：" + errorMessage
        );
    }

    /**
     * 处理参数绑定异常
     */
    @ExceptionHandler(BindException.class)
    public ExcelExportResponse handleBindException(BindException e) {
        String errorMessage = e.getFieldErrors().stream()
                .map(FieldError::getDefaultMessage)
                .collect(Collectors.joining("; "));
        
        log.error("参数绑定失败：{}", errorMessage);
        
        return ExcelExportResponse.error(
                ExcelErrorCode.PARAM_ERROR.getCode(),
                "参数绑定失败：" + errorMessage
        );
    }

    /**
     * 处理内存溢出异常
     */
    @ExceptionHandler(OutOfMemoryError.class)
    public ExcelExportResponse handleOutOfMemoryError(OutOfMemoryError e) {
        log.error("内存溢出", e);
        
        return ExcelExportResponse.error(
                ExcelErrorCode.OUT_OF_MEMORY.getCode(),
                ExcelErrorCode.OUT_OF_MEMORY.getMessage()
        );
    }

    /**
     * 处理线程中断异常
     */
    @ExceptionHandler(InterruptedException.class)
    public ExcelExportResponse handleInterruptedException(InterruptedException e) {
        log.error("线程中断", e);
        
        return ExcelExportResponse.error(
                ExcelErrorCode.THREAD_INTERRUPTED.getCode(),
                ExcelErrorCode.THREAD_INTERRUPTED.getMessage()
        );
    }

    /**
     * 处理其他未知异常
     */
    @ExceptionHandler(Exception.class)
    public ExcelExportResponse handleException(Exception e) {
        // 忽略静态资源异常（Swagger UI）
        if (e instanceof org.springframework.web.servlet.resource.NoResourceFoundException) {
            log.debug("静态资源未找到（忽略）：{}", e.getMessage());
            return ExcelExportResponse.error(404, "资源未找到");
        }
        
        log.error("系统异常", e);
        
        return ExcelExportResponse.error(
                ExcelErrorCode.SYSTEM_ERROR.getCode(),
                "系统内部错误：" + e.getMessage()
        );
    }
}
