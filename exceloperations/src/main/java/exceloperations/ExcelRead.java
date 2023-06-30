package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	XSSFSheet sheet;
	Row row;
	Cell cell;

	public ExcelRead() throws IOException {// CONSTRUCTOR - FILE SHD ReaD ALWS
		FileInputStream file = new FileInputStream("C:\\Users\\HP\\Documents\\ExcelDate.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file); // .XLXS SHEETS
		sheet = workbook.getSheet("sheet1");
	}

	public String readData(int i, int j) {
		try {
			row = sheet.getRow(i);
			cell = row.getCell(j);
			// System.out.println(cell.getStringCellValue());
			CellType type = cell.getCellType();
			switch (type) {
			case NUMERIC:
				double date = cell.getNumericCellValue();
				return String.valueOf(date);// double to string
			
			case STRING:
				return cell.getStringCellValue();
				
			}
		} catch (Exception e) {

		}
		return "Invld data type";
	}

	public static void main(String args[]) throws IOException {

		ExcelRead ex = new ExcelRead();
		for (int i = 0; i < 3; i++) { // row
			for (int j = 0; j < 3; j++) { // cell
				System.out.print( ex.readData(i, j) + "   ");

			}
			
			System.out.println(" ");
		}
	
	}

}
