package com.luxiaoyuan.excel.poi;

import java.util.ArrayList;
import java.util.List;

public class ExcelMergedRegion{
	
	private Integer firstRow ; //起始行
	private Integer lastRow;//结束行
	private Integer firstColumn;//起始列
	private Integer lastColumn;//结束列
	private List<String> valList;
	private boolean picFlag = false;
	
	
	public ExcelMergedRegion(Integer firstRow, Integer lastRow, Integer firstColumn, Integer lastColumn) {
		super();
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstColumn = firstColumn;
		this.lastColumn = lastColumn;
	}
	
	public ExcelMergedRegion(Integer firstRow, Integer lastRow, Integer firstColumn, Integer lastColumn,String val) {
		super();
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstColumn = firstColumn;
		this.lastColumn = lastColumn;
	    valList = new ArrayList<>();
		this.valList.add(val);
	}
	
	
	public boolean isPicFlag() {
		return picFlag;
	}

	public void setPicFlag(boolean picFlag) {
		this.picFlag = picFlag;
	}

	public List<String> getValList() {
		return valList;
	}

	public void setValList(List<String> valList) {
		this.valList = valList;
	}

	public Integer getFirstRow() {
		return firstRow;
	}
	public void setFirstRow(Integer firstRow) {
		this.firstRow = firstRow;
	}
	public Integer getLastRow() {
		return lastRow;
	}
	public void setLastRow(Integer lastRow) {
		this.lastRow = lastRow;
	}
	public Integer getFirstColumn() {
		return firstColumn;
	}
	public void setFirstColumn(Integer firstColumn) {
		this.firstColumn = firstColumn;
	}
	public Integer getLastColumn() {
		return lastColumn;
	}
	public void setLastColumn(Integer lastColumn) {
		this.lastColumn = lastColumn;
	}
	
}
