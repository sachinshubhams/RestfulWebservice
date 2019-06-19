package com.excelReader.ReadExcel.other;

import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Hello world!
 *
 */
public class App {
	public static void reader() {
		{
			String str = "C://Users/Shvintech/Desktop/Master Milestone Data_29032019.xls";
			try {

				POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(str));
				HSSFWorkbook wb = new HSSFWorkbook(fs);
				HSSFSheet sheet = wb.getSheetAt(0);
				HSSFRow row;
				HSSFCell cell;

				int rows; // No of rows
				rows = sheet.getPhysicalNumberOfRows();

				int cols = 0; // No of columns
				int tmp = 0;

				// This trick ensures that we get the data properly even if it
				// doesn't start from first few
				// rows
				for (int i = 0; i < 10 || i < rows; i++) {
					row = sheet.getRow(i);
					if (row != null) {
						tmp = sheet.getRow(i).getPhysicalNumberOfCells();
						if (tmp > cols)
							cols = tmp;
					}
				}
				int noOfColumns = sheet.getRow(0).getLastCellNum();
				String[] Headers = new String[noOfColumns];
				for (int j = 0; j < noOfColumns; j++) {
					Headers[j] = sheet.getRow(0).getCell(j).getStringCellValue();
				}
				for (int a = 0; a < noOfColumns; a++) {

				}

				for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
					row = sheet.getRow(rowIndex);
					if (row != null) {
						Cell user = row.getCell(1);
						Cell station = row.getCell(2);
						Cell isDefault = row.getCell(8);
						Cell isair = row.getCell(9);
						Cell isOcean = row.getCell(10);

						if (user != null && station != null) {
							// Found column and there is value in the cell.
							String default1 = isDefault.getStringCellValue();
							String air = isair.getStringCellValue();
							String ocean = isOcean.getStringCellValue();
							String isUser = user.getStringCellValue();
							String isStation = station.getStringCellValue();
							// Do something with the cellValueMaybeNull here ...

							System.out.println("UPDATE user_stations SET is_default=" + "" + default1 + ", is_air = "
									+ "" + air + "," + "is_ocean = " + ocean + " " + "WHERE user_id = " + "'" + isUser
									+ "'" + " and" + " station_id = " + "'" + isStation + "';");
						}
					}

					for (int r = 0; r < rows; r++) {

						row = sheet.getRow(r);
						if (row != null) {
							for (int c = 0; c < cols; c++) {
								cell = row.getCell((int) c);
								if (cell != null) {
									cell.setCellType(Cell.CELL_TYPE_STRING);
									cell.getStringCellValue();
									/*
									 * System.out.println(
									 * "Update Table user_stations where user_id = "
									 * + cell + "" + "station_id = "+ cell);
									 */

								}
							}
						}
					}
				}
			} catch (Exception ioe) {
				ioe.printStackTrace();
			}
		}
	}
	
	public static void main(String[] args) {
		App ap=new App();
		ap.reader();
	}

	
}
