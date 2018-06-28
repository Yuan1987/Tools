package com.yuan.controller;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

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
    public Map<String, List<Map<String, String>>> excelToJson(MultipartFile file) throws Exception {

        Map<String, List<Map<String, String>>> data = null;

        String filename = file.getOriginalFilename();

        if (file == null || (!filename.endsWith(excel2007U) && !filename.endsWith(excel2003L))) {
            return data;
        }

        Map<String, List<List<String>>> sheetMap = null;
        try (InputStream in = file.getInputStream()) {

            sheetMap = this.getBankListByExcel(in, file.getOriginalFilename());

            data = new LinkedHashMap<>();

            Iterator<Entry<String, List<List<String>>>> sheetIt = sheetMap.entrySet().iterator();

            while (sheetIt.hasNext()) {
                Entry<String, List<List<String>>> en = sheetIt.next();

                List<List<String>> listob = en.getValue();

                List<Map<String, String>> sheetList = new ArrayList<>();

                for (int i = 1; i < listob.size(); i++) {

                    Map<String, String> map = new LinkedHashMap<>();

                    List<String> first = listob.get(0);

                    for (String key : first) {
                        map.put(key.trim(), "");
                    }

                    List<String> lo = listob.get(i);

                    Iterator<String> it = map.keySet().iterator();

                    for (String val : lo) {

                        String key = it.next();

                        map.put(key, val);
                    }

                    sheetList.add(map);
                }
                data.put(en.getKey(), sheetList);
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

    public Map<String, List<List<String>>> getBankListByExcel(InputStream in, String fileName) throws Exception {

        Map<String, List<List<String>>> sheetMap = new LinkedHashMap<>(16);

        // 创建Excel工作薄
        Workbook work = this.getWorkbook(in, fileName);
        if (null == work) {
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;

        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            sheet = work.getSheetAt(i);
            if (sheet == null) {
                continue;
            }

            List<List<String>> list = new ArrayList<>();

            for (int j = sheet.getFirstRowNum(); j < sheet.getLastRowNum() + 1; j++) {

                row = sheet.getRow(j);
                if (row == null /* || row.getFirstCellNum() == j */) {
                    continue;
                }

                List<String> li = new ArrayList<String>();
                for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
                    cell = row.getCell(y);
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {// 空白判定
                        li.add(cell.getStringCellValue());
                    }

                }
                list.add(li);
            }
            sheetMap.put(sheet.getSheetName(), list);
        }
        return sheetMap;
    }
}
