package com.excel.helper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.RowId;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.excel.model.MyExcel;

public class ExcelHelper {
	
	  static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
	  static String[] HEADERs = { "Id", "Title", "Description", "Published" };
	  static String SHEET = "MySheet";
	  
	  
	  public static ByteArrayInputStream MyexcelToDownload(List<MyExcel> myExcels) {
		  
		  try(Workbook workbook = new HSSFWorkbook(); //XSSFWork.... -->> xlsx, xml office
			  ByteArrayOutputStream out = new ByteArrayOutputStream();
				  ) {
			  
			  Sheet sheet = workbook.createSheet();
			  
			  //Header
			  Row headerRow = sheet.createRow(0);
			  
			  for (int col = 0; col < HEADERs.length; col++) {
				  
				  Cell cell = headerRow.createCell(col);
				  cell.setCellValue(HEADERs[col]);				
			} // ForLoop
			  
			  
			  int rowIDx=1;
			  for (MyExcel myExcel : myExcels) {
				  
				  Row row = sheet.createRow(rowIDx);
				  
				  row.createCell(0).setCellValue(myExcel.getId());
				  row.createCell(1).setCellValue(myExcel.getTitle());
				  row.createCell(2).setCellValue(myExcel.getDescription());
				  row.createCell(3).setCellValue(myExcel.getPublished());
				  
				  rowIDx++;
				
			} //ForLoop
			  
			  
			  workbook.write(out);
			  return new ByteArrayInputStream(out.toByteArray());
			
		} catch (IOException e) {
			throw new RuntimeException("fail to import data to Excel file "+e.getMessage());
		}
		  
	  } // myExcelTodownload method
	  
	  
	  
	  
	  //Excel file upload method
	  
	  public static List parseExcelFile(InputStream is) {
		  
		  try {
			  
			  HSSFWorkbook workbook = new HSSFWorkbook(is);
			  
			  Sheet sheet = workbook.getSheet("Sheet0");
			  Iterator rows = sheet.iterator();
			  
			  List myExcelList = new ArrayList();
			  
			  int rowIDx = 0;
			  
			  while(rows.hasNext()) {
				  Row currentRow = (Row) rows.next();
				  
				  //skip header
				  
				  if (rowIDx==0) {
					rowIDx++;
					continue;
				}
				  
				  Iterator cellInRow = currentRow.iterator();
				  MyExcel myExcel = new MyExcel();
				  
				  int cellIndex = 0;
				  while(cellInRow.hasNext()) {
					  Cell currenCell = (Cell) cellInRow.next();
					  
					  if (cellIndex == 0) {
						myExcel.setId((int)currenCell.getNumericCellValue());
					}else if (cellIndex==1) {
						myExcel.setTitle(currenCell.getStringCellValue());
					}else if (cellIndex==2) {
						myExcel.setDescription(currenCell.getStringCellValue());
					}else if (cellIndex==2) {
						myExcel.setPublished(currenCell.getStringCellValue());
					}
					  
					  cellIndex++;
					  
					  
				  }// cellInRow while loop
				  
				  
				  myExcelList.add(myExcel);
				  
				  
			  }// while loop inside try-block
			  
			  //close workbook
			  workbook.close();
			  
			  return myExcelList;
			
		} catch (IOException e) {
			throw new RuntimeException("Fail! -> message = "+ e.getMessage());
		}// catch block end
		  
		  
		  
	  } //parseExcelFile method end

}
