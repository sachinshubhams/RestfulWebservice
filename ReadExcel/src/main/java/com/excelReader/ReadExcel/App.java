package com.excelReader.ReadExcel;

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
			String str = "C://Users/Shvintech/Desktop/Old MasterData/Master Milestone Data OLD.xls";
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
						Cell _milestone_code = row.getCell(0);
						Cell _milestone_description = row.getCell(1);
						Cell __milestone_level = row.getCell(2);
						Cell _pod_flag = row.getCell(3);
						Cell _milestone_location = row.getCell(4);
						Cell _event_sequence = row.getCell(5);
						Cell _cvdb_code_estimated = row.getCell(6);
						Cell _cvdb_code_actual = row.getCell(7);
						Cell _allow_edit_scheduled = row.getCell(8);
						Cell _allow_edit_estimated = row.getCell(9);
						Cell _allow_edit_actual = row.getCell(10);
						Cell _allow_additional = row.getCell(11);
						Cell _allow_status = row.getCell(12);
						Cell _expiry_date = row.getCell(13);
						Cell _mode = row.getCell(14);
						//Cell _milestone_code_level_mode = row.getCell(15);
						//Cell _created_on = row.getCell(16);
						Cell _created_by = row.getCell(17);
						Cell _updated_on = row.getCell(18);
						Cell _updated_by = row.getCell(19);
						Cell _version = row.getCell(20);

						if (_milestone_code != null && _milestone_description != null) {
							// Found column and there is value in the cell.
							String milestone_code = _milestone_code.getStringCellValue();
							String milestone_description = _milestone_description.getStringCellValue();
							String milestone_level = __milestone_level.getStringCellValue();
							String pod_flag = _pod_flag.getStringCellValue();
							String milestone_location = _milestone_location.getStringCellValue();
							String event_sequence = _event_sequence.getStringCellValue();
							String cvdb_code_estimated = _cvdb_code_estimated.getStringCellValue();
							String cvdb_code_actual = _cvdb_code_actual.getStringCellValue();
							String allow_edit_scheduled = _allow_edit_scheduled.getStringCellValue();
							String allow_edit_estimated = _allow_edit_estimated.getStringCellValue();
							String allow_edit_actual = _allow_edit_actual.getStringCellValue();
							String allow_additional = _allow_additional.getStringCellValue();
							String allow_status = _allow_status.getStringCellValue();
							String expiry_date = _expiry_date.getStringCellValue();
							String mode = _mode.getStringCellValue();
						//	String milestone_code_level_mode = _milestone_code_level_mode.getStringCellValue();
							String created_on = "now()";
							String created_by = _created_by.getStringCellValue();
							String updated_on = _updated_on.getStringCellValue();
							String updated_by = _updated_by.getStringCellValue();
							String version = _version.getStringCellValue();
							
							// Do something with the cellValueMaybeNull here ...

							System.out.println("(" 
									+"'"+milestone_code+"'"+","
									+"'"+milestone_description+"'"+","
									+"'"+milestone_level+"'"+","
									+pod_flag+","
									+"'"+milestone_location+"'"+","
									+event_sequence+","
									+"'"+cvdb_code_estimated+"'"+","
									+"'"+cvdb_code_actual+"'"+","
									+allow_edit_scheduled+","
									+allow_edit_estimated+","
									+allow_edit_actual+","
									+allow_additional+","
									+allow_status+","
									+expiry_date+","
									+"'"+mode+"'"+","
									+"'"+milestone_code+"$"+milestone_level+"$"+mode+"'"+","
									+created_on+","
									+"'"+created_by+"'"+","
									+updated_on+","
									+updated_by+","
									+version
									+")"
									+ ",");
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
	//	App app=new App();
		App.reader();
	}

	
}
