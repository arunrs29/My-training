package excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLoop {
	XSSFSheet sheet;

	public ExcelLoop() throws IOException {
		FileInputStream file = new FileInputStream("C:\\Users\\USER\\Desktop\\wb\\Book2.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		sheet = workbook.getSheet("sheet3");

	}

	public String sheetRead(int i, int j) {
		Row row = sheet.getRow(i);
		Cell cell = row.getCell(j);
		CellType type = cell.getCellType();
		switch (type) {
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case STRING:
			return cell.getStringCellValue();
		}

		return null;

	}

	public int rowSized() {
		return sheet.getLastRowNum()+1;

	}

	public static void main(String[] args) throws IOException {
		ExcelLoop ob = new ExcelLoop();
		for (int i = 0; i < ob.rowSized(); i++) {
			for (int j = 0; j < 2; j++) {
				String val = ob.sheetRead(i,j);
				System.out.println(val);

			}
		}

	}

}
