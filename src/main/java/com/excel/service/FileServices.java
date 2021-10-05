package com.excel.service;

import java.io.IOException;
import java.util.List;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import com.excel.helper.ExcelHelper;
import com.excel.model.MyExcel;
import com.excel.repository.ExcelFileRepo;

@Service
public class FileServices {
	
	@Autowired
	ExcelFileRepo excelFileRepo;
	
	// Store File Data to Database
		public void store(MultipartFile file){
			try {
				List<MyExcel> lstMyExcel = ExcelHelper.parseExcelFile(file.getInputStream());
	    		// Save Customers to DataBase
				excelFileRepo.saveAll(lstMyExcel);
	        } catch (IOException e) {
	        	throw new RuntimeException("FAIL! -> message = " + e.getMessage());
	        }
		}

}