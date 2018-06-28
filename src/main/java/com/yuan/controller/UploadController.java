package com.yuan.controller;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;

@RestController
@Api(tags = { "工具" })
public class UploadController {

    private final static String excel2003L = ".xls"; // 2003- 版本的excel
    private final static String excel2007U = ".xlsx"; // 2007+ 版本的excel
    

    @ApiOperation(value = "excelToJson")
    @PostMapping("/excelToJson")
    public List<Map<String,String>> excelToJson(MultipartFile file) throws Exception {
        
        
        List<Map<String, String>> data = null;
        
        String filename = file.getOriginalFilename();
        
        if(file == null || (!filename.endsWith(excel2007U) && !filename.endsWith(excel2003L))) {
            return data;
        }
        
        List<List<String>> listob = null;
        try(InputStream in = file.getInputStream()){
            
            listob = this.getBankListByExcel(in, file.getOriginalFilename());

            data = new ArrayList<>();
            
            for (int i = 1; i < listob.size(); i++) {
                
                Map<String, String> map = new LinkedHashMap<>();

                try {
                    List<String> lo = listob.get(0);

                    for (String key : lo) {
                        map.put(key.trim(), "");
                    }

                } catch (Exception e) {
                    e.printStackTrace();
                }

                try {
                    List<String> lo = listob.get(i);
                    
                    Iterator<String> it= map.keySet().iterator();
                    
                    for (String val : lo) {
                        
                        String key =it.next();
                        
                        System.out.println(key +"==" +val);
                        map.put(key, val);
                    }
                    
                    data.add(map);

                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return data;
    }
    
    /**
     * 描述：根据文件后缀，自适应上传文件的版本
     * 
     * @param inStr,fileName
     * @return
     * @throws Exception
     */
    public Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (excel2003L.equals(fileType)) {
            wb = new HSSFWorkbook(inStr); // 2003-
        } else if (excel2007U.equals(fileType)) {
            wb = new XSSFWorkbook(inStr); // 2007+
        } else {
            throw new Exception("解析的文件格式有误！");
        }
        return wb;
    }

    public List<List<String>> getBankListByExcel(InputStream in, String fileName) throws Exception {
        List<List<String>> list = null;

        // 创建Excel工作薄
        Workbook work = this.getWorkbook(in, fileName);
        if (null == work) {
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;

        list = new ArrayList<List<String>>();
        // 遍历Excel中所有的sheet
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            sheet = work.getSheetAt(i);
            if (sheet == null) {
                continue;
            }

            // 遍历当前sheet中的所有行
            for (int j = sheet.getFirstRowNum(); j < sheet.getLastRowNum() + 1; j++) {
                row = sheet.getRow(j);
                if (row == null || row.getFirstCellNum() == j) {//跳过取第一行表头的数据内容了
                    continue;
                }

                // 遍历所有的列
                List<String> li = new ArrayList<String>();
                for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
                    cell = row.getCell(y);
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    if(cell != null && cell.getCellType()!=Cell.CELL_TYPE_BLANK){//空白判定
                        li.add(cell.getStringCellValue());
                    }
                    
                }
                list.add(li);
            }
        }
        //work.close();
        return list;
    }
}
