package com.excel.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import com.excel.service.ExcelService;

@Controller
public class downloadFileController {
	
	@Autowired 
	private ExcelService excelService;
	
	@GetMapping("/download")
	public ResponseEntity<Resource> getfile(){
		String filename = "ExcelFile.xlsx";
		InputStreamResource file = new InputStreamResource(excelService.load());
		
		 return ResponseEntity.ok()
			        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
			        .contentType(MediaType.parseMediaType("application/vnd.ms-excel")).body(file);
		}

}
