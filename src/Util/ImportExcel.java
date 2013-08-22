package Util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.omg.CORBA.portable.InputStream;

public class ImportExcel {

	
	     public HSSFWorkbook getWorkBook(File file){
	    	 
	    	 HSSFWorkbook workBook = null;
	    	 try {
	    		 FileInputStream inputStream =  new FileInputStream(file);
	    		 
	    				POIFSFileSystem fileSystem = new POIFSFileSystem(inputStream);
	    				 workBook = new HSSFWorkbook(fileSystem);
	    				
					
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
						System.out.println(file.getPath());
					}
				return workBook;
	    			
	     }
	     
	     
	     public Sheet getSheet(Workbook workbook,String sheetName){
	    	 return workbook.getSheet(sheetName);
	     }
	     public Row getRow(Sheet sheet,int i){
	    	 
	    	 
			return sheet.getRow(i);
	    	 
	     }
	     
	     public static void main(String args[]){
	    	 
	    	 
	    	 ImportExcel im= new ImportExcel();
	    	 //im.importFile(File);
	     }
}
