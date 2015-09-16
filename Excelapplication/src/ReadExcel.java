import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.Tika;
import org.apache.tika.exception.TikaException;


public class ReadExcel {
	static XSSFRow row;
	public static void main(String[] args) {
		
		try {
			
			File file = new File(
					"C:\\Users\\fangyun\\Desktop\\Book1.xlsx");
			FileInputStream fis = new FileInputStream(file);
			
			Tika tika = new Tika();
			String fileType = tika.detect(fis);
			String content = tika.parseToString(file);
			System.out.println(content);
			 fileType = tika.detect(fis);
			if ("application/x-tika-msoffice".equalsIgnoreCase(fileType)) {
				xlsReader(fis);
			}else if("application/x-tika-ooxml".equalsIgnoreCase(fileType)){
				xlsxReader(file);
			}else{xlsxReader(file);
				System.out.println("The input file is not Excel file");
			}
		} catch (IOException | TikaException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	
		
	}
	
	public static void xlsReader(FileInputStream fis){
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet sheet = workbook.getSheetAt(0);
			HSSFRow row = sheet.getRow(0);
			if (row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING) {
				System.out.println(row.getCell(0).getStringCellValue());
			}
			if (row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
				System.out.println(row.getCell(1).getDateCellValue());
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}


	public static void xlsxReader(File file) {
		try {
			
		
		      //Get the workbook instance for XLSX file 
		      XSSFWorkbook workbook = new XSSFWorkbook(file);
		      XSSFSheet spreadsheet = workbook.getSheetAt(0);
		      Iterator < Row > rowIterator = spreadsheet.iterator();
		      while (rowIterator.hasNext()) 
		      {
		         row = (XSSFRow) rowIterator.next();
		         Iterator < Cell > cellIterator = row.cellIterator();
		         while ( cellIterator.hasNext()) 
		         {
		            Cell cell = cellIterator.next();
		            switch (cell.getCellType()) 
		            {
		               case Cell.CELL_TYPE_NUMERIC:
		               System.out.print( 
		               cell.getNumericCellValue() + " \t\t " );
		               break;
		               case Cell.CELL_TYPE_STRING:
		               System.out.print(
		               cell.getStringCellValue() + " \t\t " );
		               break;
		            }
		         }
		         System.out.println();
		      }
		      workbook.close();
			
/*			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFRow row = sheet.getRow(0);
			if (row.getCell(0).getCellType() == XSSFCell.CELL_TYPE_STRING) {
				System.out.println(row.getCell(0).getStringCellValue());
			}
			if (row.getCell(1).getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
				System.out.println(row.getCell(1).getDateCellValue());
			}*/
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
