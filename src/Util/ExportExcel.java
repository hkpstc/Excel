package Util;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExportExcel {
	/**
	 * 功能：将HSSFWorkbook写入Excel文件
	 * 
	 * @param wb
	 *            HSSFWorkbook
	 * @param absPath
	 *            写入文件的相对路径
	 * @param wbName
	 *            文件名
	 */
	public static void writeWorkbook(Workbook wb, String fileName) {
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(fileName);
			wb.write(fos);
		} catch (FileNotFoundException e) {
		} catch (IOException e) {
		} finally {
			try {
				if (fos != null) {
					fos.close();
				}
			} catch (IOException e) {
			}
		}
	}

	/**
	 * 功能：创建HSSFSheet工作簿
	 * 
	 * @param wb
	 *            HSSFWorkbook
	 * @param sheetName
	 *            String
	 * @return HSSFSheet
	 */
	public static Sheet createSheet(Workbook wb, String sheetName) {
		Sheet sheet = wb.createSheet(sheetName);
		sheet.setDefaultColumnWidth(12);
		sheet.setDisplayGridlines(false);
		return sheet;
	}

	/**
	 * 功能：创建HSSFRow
	 * 
	 * @param sheet
	 *            HSSFSheet
	 * @param rowNum
	 *            int
	 * @param height
	 *            int
	 * @return HSSFRow
	 */
	public static Row createRow(Sheet sheet, int rowNum, int height) {
		Row row = sheet.createRow(rowNum);
		row.setHeight((short) height);
		return row;
	}

	/**
	 * 功能：创建CELL
	 * 
	 * @param row
	 *            HSSFRow
	 * @param cellNum
	 *            int
	 * @param style
	 *            HSSFStyle
	 * @return HSSFCell
	 */
	public static Cell createCell(Row row, int cellNum) {
		Cell cell = row.createCell(cellNum);
		return cell;
	}

	 public static void main(String args[]){
		 HSSFWorkbook wb = new HSSFWorkbook();
		 Sheet kk=wb.createSheet("kk");
		Row row= kk.createRow(0);
		Cell cell=row.createCell(0);
		cell.setCellValue("66");
		 ExportExcel.writeWorkbook(wb, "D:\\aa.xls");
	 }
}
