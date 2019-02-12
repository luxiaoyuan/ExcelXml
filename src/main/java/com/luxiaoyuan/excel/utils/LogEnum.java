package com.luxiaoyuan.excel.utils;

public enum LogEnum {

	CONTROLLER("controller"),
	
	SERVICE("service"),

	UTILS("utils"),
	;
	
	private String category;
	
	
	LogEnum(String category) {
		this.category = category;
	}
	
	public String getCategory() {
		return category;
	}
	
	public void setCategory(String category) {
		this.category = category;
	}

}
