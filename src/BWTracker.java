import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


public class BWTracker {
	
	private static final String FILE_NAME 
		= System.getProperty("user.home").concat("\\Desktop\\BWTracker.xlsx");

	public static void main(String[] args) {
		try {

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            excelFile.close();
            Sheet sheet0 = workbook.getSheetAt(0);
            System.out.println(getLastDataRowOfColumn(sheet0, 1));
            
//            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
//            workbook.write(outputStream);
//   
//            outputStream.flush();
//            outputStream.close();
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
	}
	
	public static int getLastDataRowOfColumn(Sheet asheet, int acolumn) {
		int i;
		for (i = asheet.getLastRowNum(); i>=0; i--) {
			if (asheet.getRow(i) != null) {
				if (asheet.getRow(i).getCell(acolumn) != null) {
					if (!asheet.getRow(i).getCell(acolumn).getStringCellValue().equals("")) {
						break;
					}
				}
			}			
		}
		return i;
	}
	
	

}
