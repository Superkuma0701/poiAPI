package com.git.poitest.Controller;

import com.git.poitest.Service.StudentService;
import com.git.poitest.Util.ExcelUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.util.*;

@RestController
@RequestMapping("/poi")
public class PoiTestController {
    private static final Logger logger = LoggerFactory.getLogger(PoiTestController.class);

    private final StudentService studentService;

    @Autowired
    public PoiTestController(StudentService studentService) {
        this.studentService = studentService;
    }

    /**
     * @Author : Felix
     * @Date : 2024-04-18
     * @Description : 匯出成excel試做
     */
    @GetMapping("/test")
    public void poiTest(){
        try(XSSFWorkbook workbook = new XSSFWorkbook()){
            XSSFSheet sheet = workbook.createSheet("student detail");
            Map<String , Object[]>data = new TreeMap<>();
            data.put("1",new Object[]{"ID" , "CITY" , "STATE"});
            data.put("2",new Object[]{1 , "Clanton" , "Alabama"});
            data.put("3",new Object[]{2 , "Cordova" , "Alaska"});
            data.put("4",new Object[]{3 , "Clifton" , "Arizona"});
            data.put("5",new Object[]{4 , "Arcadia" , "California"});
            Set<String> keyset = data.keySet();
            int rownum = 0;
            for(String key : keyset){
                XSSFRow row = sheet.createRow(rownum++);
                Object[] objArr = data.get(key);
                int cellnum = 0;
                for(Object object : objArr){
                    XSSFCell cell = row.createCell(cellnum++);
                    if(object instanceof  String){
                        cell.setCellValue((String)object);
                    }else if (object instanceof Integer){
                        cell.setCellValue((Integer)object);
                    }
                }
            }
            FileOutputStream outputStream = new FileOutputStream("./workbook.xlsx");
            workbook.write(outputStream);
            outputStream.close();
            logger.info("export excel success!");
        }catch(Exception ex){
            logger.error(ex.getMessage());
        }
    }

    /**
     * @Author : Felix
     * @Date : 2024-04-18
     * @Description : 將資料庫回傳的list map匯出成excel(有標題，但欄位沒有按照順序)
     */
    @GetMapping("/test2/{excelName}")
    public String exportStudentExcel(@PathVariable String excelName){
        try(XSSFWorkbook workbook = new XSSFWorkbook()){
            //建立一張工作表,要是沒有工作表就匯出excel會打不開
            XSSFSheet sheet = workbook.createSheet("student detail");
            List<Map<String , Object>>studentList = studentService.findAll();
            XSSFCellStyle style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setWrapText(true);
            int rownum = 0 ;
            int cellnumfortitle = 0;
            for(Map<String , Object>map : studentList){
                Set<String>keySet = map.keySet();
                XSSFRow row = sheet.createRow(rownum++);
                //先在第一行建立欄位名稱
                if(cellnumfortitle<keySet.size()){
                    for(String key : keySet){
                        XSSFCell cell = row.createCell(cellnumfortitle++);
                        cell.setCellValue(key);
                        cell.setCellStyle(style);
                    }
                    //建立完跳到下一行
                    row = sheet.createRow(rownum++);
                }
                //把欄位值加上去
                int cellnumforvalue = 0;
                for(String key : keySet){
                    XSSFCell cell = row.createCell(cellnumforvalue++);
                    if(map.get(key) instanceof String){
                        cell.setCellValue((String)map.get(key));
                    }else if (map.get(key) instanceof Integer){
                        cell.setCellValue((Integer)map.get(key));
                    }
                    cell.setCellStyle(style);
                }

            }
        //匯出excel檔案
        FileOutputStream out = new FileOutputStream(excelName+".xlsx");
        workbook.write(out);
        out.close();
        logger.info("success!!");
        return "success!";
        }catch(Exception ex){
            logger.error(ex.getMessage());
            return "false!";
        }
    }

    /**
     * @Author : Felix
     * @Date : 2024-04-18
     * @Description : 講資料庫回傳的List<Object[]>資料匯出成excel(會按照欄位的順序，但不會有欄位名稱)
     */
    @GetMapping("/test3/{excelName}")
    public String export(@PathVariable String excelName){
        try(XSSFWorkbook workbook = new XSSFWorkbook()){
            XSSFSheet sheet = workbook.createSheet("student list");
            List<Object[]>student = studentService.findAlll();
            XSSFCellStyle style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setWrapText(true);
            int rownum = 0;
            for(Object[] obj : student){
                XSSFRow row = sheet.createRow(rownum++);
                int cellnum = 0;
                for(Object object : obj){
                    XSSFCell cell = row.createCell(cellnum++);
                    if(object instanceof String){
                        cell.setCellValue((String)object);
                    }else if (object instanceof Integer){
                        cell.setCellValue((Integer)object);
                    }
                    cell.setCellStyle(style);
                }
            }
            FileOutputStream out = new FileOutputStream(excelName+".xlsx") ;
            workbook.write(out);
            out.close();
            logger.info("success!");
            return "success!";
        }catch(Exception ex){
            logger.error(ex.getMessage());
            return "fail!";
        }
    }

    /**
     * @Author : Felix
     * @Date : 2024-04-18
     * @Description : 將excel表格變成list map
     */
    @PostMapping("/import")
    public ResponseEntity<?> importFile(@RequestParam("file") MultipartFile file){
        List<Map<String , Object>>result = new ArrayList<>();
        //透過方法生成對應的workbook
        Workbook workbook = ExcelUtil.getWorkBook(file);
        if(Objects.nonNull(workbook)){
            //欄位名稱
            List<String> sheetTitle = ExcelUtil.getSheetTitles(workbook);

            Sheet sheet = workbook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            //先讓row去跑回圈,cell跑完換下一個row
            for (int i = 1 ; i <= lastRowNum ; i++){
                Row row = sheet.getRow(i);
                Map<String , Object>rowMap = new HashMap<>();
                int lastCellNum = row.getLastCellNum();
                //cell去跑回圈
                for(int j = 0 ; j < lastCellNum ; j++){
                    Cell cell = row.getCell(j);
                    //依照資料型態去分類
                    switch(cell.getCellType()){
                        case NUMERIC :
                            double cellNum = cell.getNumericCellValue();
                            DecimalFormat df = new DecimalFormat("0");
                            rowMap.put(sheetTitle.get(j),df.format((cellNum)));
                            break;
                        case STRING:
                            String cellString = cell.getStringCellValue();
                            rowMap.put(sheetTitle.get(j),cellString);
                            break;
                        default:
                            break;
                    }
                }
                result.add(rowMap);
            }
        }else{
            logger.error("false");
            return new ResponseEntity<>("false", HttpStatus.INTERNAL_SERVER_ERROR);
        }
        System.out.println(result);
        logger.info("success");
        return new ResponseEntity<>(result,HttpStatus.OK);
    }
}
