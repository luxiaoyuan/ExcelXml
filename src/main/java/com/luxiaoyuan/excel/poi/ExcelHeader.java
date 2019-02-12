package com.luxiaoyuan.excel.poi;


public class ExcelHeader implements Comparable<ExcelHeader>{
    /**
     * excel的标题名称--测试
     */
    private String title;
    /**
     * 每一个标题的顺序
     */
    private int order;
    
    private int width;
    /**
     * 说对应方法名称
     */
    private String methodName;


    public String getTitle() {
        return title;
    }
    public void setTitle(String title) {
        this.title = title;
    }
    public int getOrder() {
        return order;
    }
    public void setOrder(int order) {
        this.order = order;
    }
    public String getMethodName() {
        return methodName;
    }
    public void setMethodName(String methodName) {
        this.methodName = methodName;
    }

    public int getWidth() {
		return width;
	}
	public void setWidth(int width) {
		this.width = width;
	}
	public int compareTo(ExcelHeader o) {
        return order>o.order?1:(order<o.order?-1:0);
    }
    public ExcelHeader(String title, int order, String methodName,int width) {
        super();
        this.title = title;
        this.order = order;
        this.methodName = methodName;
        this.width = width;
    }
}
