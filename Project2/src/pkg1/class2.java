package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class class2 {
public static void main(String[] args) throws IOException {
	File f=new File("C:\\Users\\sunny.goel\\Desktop\\Test12.xlsx");
	FileOutputStream os=new FileOutputStream(f);
	XSSFWorkbook wk=new XSSFWorkbook();
	XSSFSheet sh=wk.createSheet();
	
	for(int i=0; i<3; i++) {
		
		XSSFRow xr=sh.createRow(i);
		for(int j=0; j<3; j++) {
			XSSFCell xc=xr.createCell(j);
			xc.setCellValue("Sunny test 1234");
		}
		
	}
	wk.write(os);
	os.flush();
	os.close();
	
	FileInputStream is=new FileInputStream(f);
    XSSFWorkbook w= new XSSFWorkbook(is);
    XSSFSheet s= w.getSheetAt(0);
    
    int r=s.getPhysicalNumberOfRows();
    
    for(int i=0; i<r; i++) {
    	
     XSSFRow xrr=s.getRow(i);
    	
    	for(int j=0; j<xrr.getPhysicalNumberOfCells(); j++) {
    		XSSFCell xcc=xrr.getCell(j);
    		System.out.println(xcc.getStringCellValue());
    	}
    
    }
    
}
}
