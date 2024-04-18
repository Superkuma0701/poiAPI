package com.git.poitest.Util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

public class ExcelUtil {

    private static final Logger logger = LoggerFactory.getLogger(Error.class);
    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";
    private static final String SPLIT = ".";

    /**
     *
     * @Author : Felix
     * @Date : 2024-04-18
     * @Description : 將收到的multipart file按照名分別去建立workbook
     *
     */
    public static Workbook getWorkBook (MultipartFile file){
        Workbook workbook = null;
        try{
            //取的excel的檔名
            String filename = file.getOriginalFilename();
            if(StringUtils.isEmpty(filename) || filename.lastIndexOf(SPLIT)<0){
                logger.warn("解析excel失敗，沒有正確的檔名");
                return null;
            }
            String fileType = filename.substring(filename.lastIndexOf(SPLIT)+1);
            //找出檔名對應合適的workbook
            if(fileType.equalsIgnoreCase(XLS)){
                workbook = new HSSFWorkbook(file.getInputStream());
            }else if(fileType.equalsIgnoreCase(XLSX)){
                workbook = new XSSFWorkbook(file.getInputStream());
            }
        }catch(IOException ex){
            logger.error(ex.getMessage());
        }
        return workbook;
    }

    /**
     * @Author : Felix
     * @Date  : 2024-04-18
     * @Description : 取得表頭的欄位名稱
     */
    public static List<String> getSheetTitles(Workbook workbook){
        //取第一個sheet
        Sheet sheet = workbook.getSheetAt(0);
        //取出第一列
        Row sheetTitleRow = sheet.getRow(sheet.getFirstRowNum());
        //取出這一列的最後一個欄位的位子
        int lastCell = sheetTitleRow.getLastCellNum();
        List<String> sheetTitle = new LinkedList<>();
        for(int i = 0 ; i<lastCell ; i++){
            sheetTitle.add(sheetTitleRow.getCell(i).getStringCellValue());
        }
        return sheetTitle;
    }


}
