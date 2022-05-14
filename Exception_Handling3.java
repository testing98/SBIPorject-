import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Exception_Handling3 {
	
	void add() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		File file=new File("D:\\Test.xlsx");
		FileInputStream fis=new FileInputStream(file);
		Workbook wb=WorkbookFactory.create(fis);
		Sheet sh=wb.getSheetAt(1);
		Row r=sh.getRow(1);
		System.out.println(r.getCell(2).getStringCellValue());
	}

	public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		int a=10;
		int b=50;
		System.out.println(a+b);
		Thread.sleep(3000);
		Exception_Handling3 e3=new Exception_Handling3();
		e3.add();
		
		
		
		

	}

}
