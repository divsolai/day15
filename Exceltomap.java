package java30;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exceltomap {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String fileName = "D:\\Ecllipse\\Ec_test\\LeafBot\\data\\Map2.xlsx";
		FileInputStream file = new FileInputStream(new File(fileName));
		Map<Integer,List<String>> map = new LinkedHashMap<Integer,List<String>>();
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet = workbook.getSheetAt(0);

		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			List<String> data = new ArrayList<String>();

			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();                      
				System.out.print(cell.getStringCellValue() + "\t");
				data.add(cell.getStringCellValue()); 
			}
			System.out.println("");
			map.put(row.getRowNum(), data);
		}
		System.out.println(map);
		file.close();
	}

}
