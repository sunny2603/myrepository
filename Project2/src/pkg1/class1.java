package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class class1 {

	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\sunny.goel\\Desktop\\sunny1");
		FileInputStream is = new FileInputStream(f);
		XSSFWorkbook wk = new XSSFWorkbook(is);
		XSSFSheet ws = wk.getSheetAt(0);

		int r = ws.getPhysicalNumberOfRows();

		for (int i = 0; i < r; i++) {

			XSSFRow xr = ws.getRow(i);

			for (int j = 0; j < xr.getPhysicalNumberOfCells(); j++) {

				XSSFCell xc = xr.getCell(j);
				System.out.println(xc.getStringCellValue());

			}
		}

	}
}
