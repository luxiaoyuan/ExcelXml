package com.luxiaoyuan.excel;

public class ExcelXMLContent {
	
	//XdrSpPr   x y cx cy   默认值先 其他的计算才知道、   也是9527 单位   3d旋转多少度
	//default   200*200
	public static final int Xdr_SpPr_x=35966400;
	public static final int Xdr_SpPr_y=2720975;
	public static final int Xdr_SpPr_cx=2111375;
	public static final int Xdr_SpPr_cy=1783080;
	
	//from
	public static final int Xdr_to_colOff=1905000; //9525*200
	public static final int Xdr_to_rowOff=1905000;//9525*200
	//to
	public static final int Xdr_from_colOff=9525;  //9525*1
	public static final int Xdr_from_rowOff=9525;
	
	//endcoding
	public static final String XML_ENCODING = "UTF-8";
	
	// image 前缀
	public static final String IMAGE_PREFIX="image";
	//relationship Id 前缀  
	public static final String RELATIONSHIP_ID_PREFIX="rId";
	//drawing1.xml.rels  image前缀
	public static final String  DRAWING_XML_RELS_PREFIX="../media/";
	public static final String  SHEET_XML_RELS_PREFIX="../drawings/";
	
	

}
