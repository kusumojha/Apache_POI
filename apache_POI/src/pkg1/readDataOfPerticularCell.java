package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//READ DATA OF PERTICULAR CELL IN xlsx FORMAT
public class readDataOfPerticularCell 
{
  public void readdataofcell(int row,int column) throws IOException
  {

			File f=new File("C:\\Users\\admin\\Desktop\\input.xlsx");
			FileInputStream fi=new FileInputStream(f);
			XSSFWorkbook xs=new XSSFWorkbook(fi);
			XSSFSheet xt=xs.getSheetAt(0);
			XSSFRow xr=xt.getRow(row);
			XSSFCell xc=xr.getCell(column);
			System.out.println(xc.getStringCellValue());			
  }
  public static void main(String[] args) throws IOException 
  { try {
	  readDataOfPerticularCell obj=new readDataOfPerticularCell();
	  Scanner s=new Scanner(System.in);
	  System.out.println("ENTER ROW AND COLUMN OF WHICH YOU WANT TO RETRIVE THE VALUE");
	  int r=s.nextInt();
	  int c=s.nextInt();
	  obj.readdataofcell(r, c);
	
} catch (Exception e) {
	// TODO: handle exception
}
	  
	
}
}
