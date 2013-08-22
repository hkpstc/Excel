package Util;

import java.io.File;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import DoWork.DoWork;
import DoWork.DoWork2;


public class FileUtil {

	public Object getAllFileFromDir(File dir,String type) {
		Object object = null;
		if (dir.isDirectory()) {

			String[] fileList = dir.list();
			for (int i = 0; i < fileList.length; i++) {
				File file = new File(dir.getPath()+"\\"+fileList[i]);
				if (file.isDirectory()) {
					getAllFileFromDir(file,type);
				}else{
					 if (file.getAbsolutePath().endsWith(type)){
					 //object=	DoWork.doWork(file);查找时间
						 //查找策略
						 DoWork2.doWork2(file);
						 System.out.println(file.getPath());
					 }
				}

			}

		}

		return object;

	}
}
