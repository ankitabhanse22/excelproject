package excelread;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class A 
{
	public void readdata(String filenm, String sheetnm) throws IOException
	{

	
		int arrayexcel[][]=null;
		FileInputStream file=new FileInputStream(filenm);
		
		XSSFWorkbook web=new XSSFWorkbook(file);
		XSSFSheet sheet=web.getSheet(sheetnm);
		XSSFRow row=sheet.getRow(2);
		XSSFCell cell=row.getCell(2);
		String val=cell.getStringCellValue();
		System.out.println("The value at index 2,2 is="+val);
		//get row count
		int rows=sheet.getLastRowNum();
		System.out.println("Row="+rows);
		int rowcount=rows+1;
		System.out.println("The number of rows are="+rowcount);
		//get column count
		int columns=sheet.getRow(rows).getLastCellNum();
		System.out.println("The number of columns are="+columns);
		arrayexcel= new int[rowcount][columns];
		for(int i=0;i<rowcount;i++)
		{
			for(int j=0;j<columns;j++)
			{
				System.out.println(sheet.getRow(i).getCell(j));
				
				/*DataFormatter df=new DataFormatter();
				String value=df.formatCellValue(sheet.getRow(i).getCell(j));
				System.out.println(value);*/
			}
		}
	}
	public static void main(String[] args) throws IOException 
	{
		A obj=new A();
		obj.readdata("F:\\newworkspace\\excel17april\\studentds.xlsx", "Sheet1");
		
	}
}
