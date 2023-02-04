package DataDriven.Excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static void main(String[] args) throws IOException {

	}

	public ArrayList<String> getData(String testcaseName) throws IOException {

		ArrayList<String> a = new ArrayList<String>();

		FileInputStream fis = new FileInputStream("‪‪E:\\ExcelDriven.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {

			if (workbook.getSheetName(i).equals("testdata")) {

				XSSFSheet sheet = workbook.getSheetAt(i);
				// Identify testcase column by scanning the entire row
				Iterator<Row> rows = sheet.rowIterator();
				Row firstrow = rows.next();
				Iterator<Cell> cell = firstrow.cellIterator();// row is collection of cell
				int k = 0;
				int column = 0;																		
				while (cell.hasNext()) {

					Cell value = cell.next();
					if (value.getStringCellValue().equalsIgnoreCase("Testcases")) {
						column = k;

					}
				}
				System.out.println(column);
				// once column is identified then scan entire testase column to identify
				// purchase testcase row
				while (rows.hasNext()) {

					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {
						// after you grab all the testcase row pull all the data of that row and feed
						// into test

						Iterator<Cell> cel = r.cellIterator();
						while (cel.hasNext()) {

							Cell c = cel.next();
							if (c.getCellType() == CellType.STRING) {
								a.add(c.getStringCellValue());
							} else {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));

							}

						}

					}
					return a;
				}
				return a;

			}

		}
		return a;
	}
System.out.println("sadanandam");
}
