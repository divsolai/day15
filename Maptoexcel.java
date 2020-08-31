package java30;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Maptoexcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sample");
	    XSSFRow row;

        Map<String, List<String>> empData = new HashMap<String,List<String>>();
     List<String> name= new ArrayList<String>();
     name.add("Employee Id");
     name.add("Employee name");
     name.add("Salary");
     empData.put("1", name);
     List<String> d1 = new ArrayList<String>();
     d1.add("1");
     d1.add("balaji");
     d1.add("100000");
     empData.put("2", d1);
     List<String> d2 = new ArrayList<String>();
     d2.add("2");
     d2.add("Ganesh");
     d2.add("100000");
     empData.put("3", d2);
     List<String> d3 = new ArrayList<String>();
     d3.add("3");
     d3.add("Krish");
     d3.add("100000");
     empData.put("4", d3);
     System.out.println(empData);
     Set < String > keyid = empData.keySet();
     int rowid = 0;
     
     for (String key : keyid) {
        row = sheet.createRow(rowid++);
        List<String> objectArr = empData.get(key);
        int cellid = 0;
        
        for (String obj : objectArr){
           XSSFCell cell = row.createCell(cellid++);
           cell.setCellValue((String)obj);
           
           FileOutputStream out = new FileOutputStream(
        	      new File("D:\\Ecllipse\\Ec_test\\LeafBot\\data\\Map1.xlsx"));
          	      workbook.write(out);
        	      out.close();
                    }
     }

     }
}
