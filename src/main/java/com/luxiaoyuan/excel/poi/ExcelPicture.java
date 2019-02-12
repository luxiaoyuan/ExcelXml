package com.luxiaoyuan.excel.poi;


public class ExcelPicture {
	private int startRow;//坐标
	private int endRow;
	private int startIndex;
	private int endIndex;
	private String url;//图片的下载url
	public int getStartRow() {
		return startRow;
	}
	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}
	public int getEndRow() {
		return endRow;
	}
	public void setEndRow(int endRow) {
		this.endRow = endRow;
	}
	public int getStartIndex() {
		return startIndex;
	}
	public void setStartIndex(int startIndex) {
		this.startIndex = startIndex;
	}
	public int getEndIndex() {
		return endIndex;
	}
	public void setEndIndex(int endIndex) {
		this.endIndex = endIndex;
	}
	public String getUrl() {
		return url;
	}
	public void setUrl(String url) {
		this.url = url;
	}
	public ExcelPicture(int startRow, int endRow, int startIndex, int endIndex, String url) {
		super();
		this.startRow = startRow;
		this.endRow = endRow;
		this.startIndex = startIndex;
		this.endIndex = endIndex;
		this.url = url;
	}
	
}
