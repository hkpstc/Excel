package DoWork;

import java.io.File;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import Util.ExportExcel;
import Util.ImportExcel;
import Util.FileUtil;

public class DoWork {

	public static Object doWork(File file) {
		ImportExcel importExcel = new ImportExcel();
		ExportExcel exportExcel = new ExportExcel();

		HSSFWorkbook workbook = importExcel.getWorkBook(file);
		Cell cell;
		Sheet sheet;
		Row row;
		/*---------输出参数----------*/
		//HSSFWorkbook newWorkBook = new HSSFWorkbook();			
		int i=0;//行数
		if (workbook != null) {
			sheet = workbook.getSheet("Sheet1");
			if (sheet != null) {
				row = importExcel.getRow(sheet, 1);
				if (row != null) {
					cell = row.getCell(0);
					if (cell != null) {
						String aa = cell.toString();
						if(aa.startsWith("执行时间")){
							
						System.out.println(aa.replace("执行时间：", ""));
						
						
						
						}
					}
				}

			}
		}
		return null;
	}

	public static void main(String args[]) {
		ExportExcel exportExcel = new ExportExcel();
		HSSFWorkbook wb = new HSSFWorkbook();

		FileUtil fileUtil = new FileUtil();
		File file = new File(
				"C:\\Documents and Settings\\Administrator\\桌面\\ssss");
		fileUtil.getAllFileFromDir(file,"xls");
	}
}
