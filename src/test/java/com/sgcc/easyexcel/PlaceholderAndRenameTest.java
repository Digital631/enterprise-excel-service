package com.sgcc.easyexcel;

import com.sgcc.easyexcel.dto.request.SingleSheetExportRequest;
import com.sgcc.easyexcel.service.ExcelExportService;
import com.sgcc.easyexcel.util.FileNameGenerator;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.*;

@SpringBootTest
public class PlaceholderAndRenameTest {

    @Autowired
    private ExcelExportService excelExportService;
    
    @Autowired
    private FileNameGenerator fileNameGenerator;

    @Test
    public void testPlaceholderAndRename() {
        // 创建导出请求
        SingleSheetExportRequest request = SingleSheetExportRequest.builder()
                .templatePath("员工统计表") // 使用你提到的模板名称
                .data(createTestData())
                .startRow(2)
                .fieldMapping(createFieldMapping())
                .placeholders(createPlaceholders()) // 添加占位符数据
                .exportFileName("2023年12月员工统计.xlsx")
                .exportDir("E:/custom-exports/")
                .build();

        try {
            // 执行导出
            var result = excelExportService.exportSingleSheet(request);
            System.out.println("导出成功: " + result.getFilePath());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    @Test
    public void testOnlyDataFill() {
        // 测试只填充数据的场景
        SingleSheetExportRequest request = SingleSheetExportRequest.builder()
                .templatePath("员工统计表") // 使用你提到的模板名称
                .data(createTestData())
                .startRow(2)
                .fieldMapping(createFieldMapping())
                .exportFileName("仅数据填充测试.xlsx")
                .exportDir("E:/custom-exports/")
                .build();

        try {
            // 执行导出
            var result = excelExportService.exportSingleSheet(request);
            System.out.println("导出成功: " + result.getFilePath());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    private List<Map<String, Object>> createTestData() {
        List<Map<String, Object>> data = new ArrayList<>();
        
        Map<String, Object> row1 = new HashMap<>();
        row1.put("userName", "张三");
        row1.put("age", 25);
        row1.put("salary", 8000.0);
        data.add(row1);
        
        Map<String, Object> row2 = new HashMap<>();
        row2.put("userName", "李四");
        row2.put("age", 30);
        row2.put("salary", 12000.0);
        data.add(row2);
        
        return data;
    }
    
    private Map<String, String> createFieldMapping() {
        Map<String, String> mapping = new HashMap<>();
        mapping.put("姓名", "userName");
        mapping.put("年龄", "age");
        mapping.put("工资", "salary");
        return mapping;
    }
    
    private Map<String, Object> createPlaceholders() {
        Map<String, Object> placeholders = new HashMap<>();
        placeholders.put("title", "员工信息统计表");
        placeholders.put("date", "2023-12-25");
        placeholders.put("department", "技术部");
        return placeholders;
    }
}