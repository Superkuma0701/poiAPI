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

    private StudentService studentService;

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
        List<Map<String , Object>>studentList = studentService.findAll();
        Map<String , Object>studentMap = studentList.get(0);
        Set<String>keySet = studentMap.keySet();
        for(String key : keySet){
            System.out.println(studentMap.get(key));
        }
        return "success!";
    }


}
