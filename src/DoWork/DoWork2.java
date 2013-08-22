package DoWork;

import java.util.ArrayList;
import java.util.List;
import java.io.File;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import Util.ExportExcel;
import Util.FileUtil;
import Util.ImportExcel;

public class DoWork2 {
	public static List<String> doWork2(File file) {
		ImportExcel importExcel = new ImportExcel();
		ExportExcel exportExcel = new ExportExcel();

		HSSFWorkbook workbook = importExcel.getWorkBook(file);
		Cell cell;
		Sheet sheet;
		Row row = null;
		int i = 0;// 行数
		int j = 0; // 列数
		boolean isThisRow = false;
		List newRow_Str = new ArrayList();

		if (workbook != null) {
			sheet = workbook.getSheet("Sheet1");
			if (sheet != null) {
				while (i >= 0) {
					j = 0;
					row = importExcel.getRow(sheet, i);
					if (row == null) {
						break;
					}
					while (j >= 0) {
						cell = row.getCell(j);
						if (cell == null
								|| (isThisRow == true && cell.getCellType() != 0)) {
							break;
						}
						String cell_str = cell.toString();
						String cell_Str_minus;
						if(i>0){
						 cell_Str_minus=sheet.getRow(cell.getRowIndex()-1).getCell(j).toString();}
						else cell_Str_minus="kk";
						if ((cell_str.contains("小熊猫")
								&& cell_str.contains("清和风"))||(cell_str==null&&cell_Str_minus.contains("小熊猫")&& cell_str.contains("清和风"))) {
							isThisRow = true;
							newRow_Str.add(sheet.getRow(1).getCell(0)
									.getStringCellValue());
							System.out.println("ok");
						}
						if (isThisRow == true) {
							newRow_Str.add(cell_str);
							System.out.println(cell_str);
						}
						j = j + 1;
					}
					if (isThisRow == true) {
						// return newRow_Str;
					}
					i = i + 1;
				}
			}
		}
		return newRow_Str;

	}

	public static void main(String args[]) {
		ImportExcel importExcel = new ImportExcel();
		ExportExcel exportExcel = new ExportExcel();
		HSSFWorkbook wb = new HSSFWorkbook();

		FileUtil fileUtil = new FileUtil();
		File file = new File(
				"C:\\Documents and Settings\\Administrator\\桌面\\策略\\小熊猫 清和风");

		
		 fileUtil.getAllFileFromDir(file,"xls");

	}
}
