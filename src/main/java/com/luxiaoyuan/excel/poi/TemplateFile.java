package com.luxiaoyuan.excel.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.ResourceUtils;


public class TemplateFile {
	
	@Value("${download.node_path}")
	public static String PATH;
	
	public static final String TEST_EXCEL="order_cash_bill.xlsx";

	
	public static final String EXPORT_XLSX=".xlsx";
	public static final String EXPORT_XLS=".xls";
	
	public TemplateFile(String path) {
		PATH = path;
	}
	
	public static FileInputStream getTemplates(String tempName) throws FileNotFoundException {
		
        return new FileInputStream(ResourceUtils.getFile("classpath:templates/"+tempName));
    }
}
