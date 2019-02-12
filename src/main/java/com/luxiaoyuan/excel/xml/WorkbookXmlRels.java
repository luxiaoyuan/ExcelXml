package com.luxiaoyuan.excel.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;

import com.luxiaoyuan.excel.ExcelXMLContent;


public class WorkbookXmlRels {
	
	  private Writer _out;
	  private File xmlFile;
	  private FileOutputStream xmlOutStream;
	  public WorkbookXmlRels(String filePath){
        getWriter(filePath);
	  }
	  public void getWriter(String filePath) {
			 xmlFile = new File(filePath); 
			 if(xmlFile.exists()) {
		    	 xmlFile.delete();//删除  每次都是新的写入
			 }
		    try {
				xmlFile.createNewFile();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		    //创建输出流  
		    xmlOutStream=null;
			try {
				xmlOutStream = new FileOutputStream(xmlFile);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} 
		     //写入xml标签  后面使用dom4j
		     Writer fw=null;
			try {
				fw = new OutputStreamWriter(xmlOutStream, ExcelXMLContent.XML_ENCODING);
			} catch (UnsupportedEncodingException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		     _out = fw;
		 }

	  //默认的这个 （测试完成后）后面的改成扩展 
	  public void writerOffice(){
		  try {
		  _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + 
		  		"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + 
		  		"	<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" + 
		  		"	<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>" + 
		  		"	<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" + 
		  		"	<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>" + 
		  		"</Relationships>");
		  }catch(IOException e) {
			  //TODO
			  e.printStackTrace();
		  }finally {
		  try {
 			if(_out!=null) {
					_out.flush();
					_out.close();
 			}
			} catch (IOException e) {
				e.printStackTrace();
			}
 		try {
 			if(xmlOutStream!=null) {
 				xmlOutStream.flush();
 				xmlOutStream.close();
 			}
			} catch (IOException e) {
				e.printStackTrace();
			}
		  }
	  }
}
