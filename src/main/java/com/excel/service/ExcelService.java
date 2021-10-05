package com.excel.service;

import java.io.ByteArrayInputStream;
import java.util.List;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.excel.helper.ExcelHelper;
import com.excel.model.MyExcel;
import com.excel.repository.ExcelFileRepo;

@Service
public class ExcelService {
	
	@Autowired
	private ExcelFileRepo excelFileRepo;
	
	public ByteArrayInputStream load() {
		List<MyExcel> myexExcels = excelFileRepo.findAll();
		
		ByteArrayInputStream in = ExcelHelper.MyexcelToDownload(myexExcels);
		return in;
	}

}
