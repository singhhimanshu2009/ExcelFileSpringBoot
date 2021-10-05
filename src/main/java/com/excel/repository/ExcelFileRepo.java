package com.excel.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.excel.model.MyExcel;

public interface ExcelFileRepo extends JpaRepository<MyExcel, Integer> {

}
