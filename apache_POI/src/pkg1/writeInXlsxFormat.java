package pkg1;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//WRITE IN XLSX FORMAT
public class writeInXlsxFormat 
{
	public static void main(String[] args) throws IOException
	{
		File f=new File("C:\\\\Users\\\\admin\\\\Desktop\\\\output.xlsx");
		FileOutputStream fo=new FileOutputStream(f);
		XSSFWorkbook xs=new XSSFWorkbook();
		XSSFSheet xt=xs.createSheet("mysheet");
		for(int i=0;i<5;i++)
		{
			XSSFRow xr=xt.createRow(i);
			for(int j=0;j<5;j++)
			{
				XSSFCell xc=xr.createCell(j);
				xc.setCellValue("deepak");
			}
		}
		xs.write(fo);
		fo.flush();
		fo.close();
		
	}
	
	

}
