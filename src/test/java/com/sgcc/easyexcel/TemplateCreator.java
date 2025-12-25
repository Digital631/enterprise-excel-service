package com.sgcc.easyexcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 创建测试模板文件的工具类
 */
public class TemplateCreator {
    
    public static void main(String[] args) {
        createTestTemplate();
    }
    
    public static void createTestTemplate() {
        // 创建工作簿
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
        String templatePath = "E:/00_WorkSpace/DK/easyexcel/excel-templates/员工统计表.xlsx";
        try (FileOutputStream outputStream = new FileOutputStream(templatePath)) {
            workbook.write(outputStream);
            System.out.println("模板文件已创建: " + templatePath);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}