package pkg1;
//WRITE IN PERTICULAR RANGE XLSX FORMAT
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteInPerticularRangeInXlsx 
{
	public void writeinrange(int row,int column) throws IOException
	{
		File f=new File("C:\\\\Users\\\\admin\\\\Desktop\\\\output.xlsx");
		FileOutputStream fo=new FileOutputStream(f);
		XSSFWorkbook xs=new XSSFWorkbook();
		XSSFSheet xt=xs.createSheet("mysheet");
		System.out.println("Enter Data");
		Scanner s=new Scanner(System.in);
		String data;
		for(int i=0;i<row;i++)
		{
			XSSFRow xr=xt.createRow(i);
			for(int j=0;j<column;j++)
			{
				data=s.next();
				XSSFCell xc=xr.createCell(j);
				xc.setCellValue(data);
			}
		}
		xs.write(fo);
		fo.flush();
		fo.close();
	}
	public static void main(String[] args) throws IOException 
	{
		WriteInPerticularRangeInXlsx obj=new WriteInPerticularRangeInXlsx();
		Scanner sc=new Scanner(System.in);
		System.out.println("enter range");
		int r=sc.nextInt();
		int c=sc.nextInt();
		obj.writeinrange(r,c);
		
	}

}
