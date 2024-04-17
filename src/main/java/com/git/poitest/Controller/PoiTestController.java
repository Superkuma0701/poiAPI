package com.git.poitest.Controller;

import com.git.poitest.Service.StudentService;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

@RestController
@RequestMapping("/poi")
public class PoiTestController {
    private static final Logger logger = LoggerFactory.getLogger(PoiTestController.class);

    private final StudentService studentService;

    @Autowired
    public PoiTestController(StudentService studentService) {
        this.studentService = studentService;
    }

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

    @GetMapping("/test2/{excelName}")
    public String exportStudentExcel(@PathVariable String excelName){
        try(XSSFWorkbook workbook = new XSSFWorkbook()){
            //建立一張工作表,要是沒有工作表就匯出excel會打不開
            XSSFSheet sheet = workbook.createSheet("student detail");
            List<Map<String , Object>>studentList = studentService.findAll();
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
    @GetMapping("/test3/{excelName}")
    public String export(@PathVariable String excelName){
        try(XSSFWorkbook workbook = new XSSFWorkbook()){
            XSSFSheet sheet = workbook.createSheet("student list");
            List<Object[]>student = studentService.findAlll();
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
}
