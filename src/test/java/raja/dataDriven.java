package raja;
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



public class dataDriven {

	public ArrayList<String> getData(String testcaseName) throws IOException {
		// Identify test case column by scanning entire 1st row
		// once column is identified then scan entire test case column to identify
		// purchase test case row
		// after you grab purchase test case row=pull all the data of that row and feed into test

		ArrayList<String> a = new ArrayList<String>();

		// File Input Stream accepts excel file as an argument also
		FileInputStream fis = new FileInputStream("C:\\Users\\dell\\Documents\\SOFTWARE TESTING\\datademo.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				// Identify test case column by scanning entire 1st row
				Iterator<Row> rows = sheet.iterator();
				Row firstrow = (Row) rows.next();
				Iterator<Cell> ce = firstrow.cellIterator();
				int k = 0;
				int column = 0;
				
				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("Testcases")) {
						// desired column
						column = k;

					}
					k++;

				}
				System.out.println(column);
				// once column is identified then scan entire test case column to identify
				// purchase test case row

				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {
						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {
							Cell c=cv.next();
							if(c.getCellType()==CellType.STRING)
							{
								a.add(c.getStringCellValue()); // all data of test case properly stored in array list
							}
							else
							{
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
								
							}
							
						}
					}
				}

			}

		}
		return a;
	}

	public static void main(String[] args) throws IOException {

	}
}
