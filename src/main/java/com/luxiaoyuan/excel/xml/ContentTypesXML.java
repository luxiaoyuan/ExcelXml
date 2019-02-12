package com.luxiaoyuan.excel.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;

import com.luxiaoyuan.excel.ExcelXMLContent;


public class ContentTypesXML {
	 private Writer _out;
	  private File xmlFile;
	  private FileOutputStream xmlOutStream;
	  public ContentTypesXML(String filePath){
         getWriter(filePath);
     }
	  public void getWriter(String filePath) {
			 xmlFile = new File(filePath);  
			 if(xmlFile.exists()) {
		    	 xmlFile.delete();//删除  每次都是新的写入
		    	 try {
					xmlFile.createNewFile();
				} catch (IOException e) {
					e.printStackTrace();
				}
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

	  
	  public void writer(){
		  try {
		  _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + 
		  		"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" + 
		  		"	<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>" + 
		  		"	<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" + 
		  		"	<Default Extension=\"xml\" ContentType=\"application/xml\"/>" + 
		  		"	<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" + 
		  		"	<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" + 
		  		"	<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>" + 
		  		"	<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>" + 
		  		"	<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" + 
		  		"	<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" + 
		  		"	<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" + 
		  		"	<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" + 
		  		"	<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" + 
		  		"</Types>" + 
		  		"");
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
	  
	  //office 写入的顺序不一样
	  public void writerOffice(){
		  try {
		  _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + 
		  		"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n" + 
		  		"	<Default Extension=\"jpeg\" ContentType=\"image/png\"/>\r\n" + 
		  		"	<Default Extension=\"jpg\" ContentType=\"application/octet-stream\"/>\r\n" + 
		  		"	<Default Extension=\"png\" ContentType=\"image/png\"/>\r\n" + 
		  		"	<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\r\n" + 
		  		"	<Default Extension=\"xml\" ContentType=\"application/xml\"/>\r\n" + 
		  		"	<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\r\n" + 
		  		"	<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\r\n" + 
		  		"	<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\r\n" + 
		  		"	<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\r\n" + 
		  		"	<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\r\n" + 
		  		"	<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>\r\n" + 
		  		"	<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\r\n" + 
		  		"	<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\r\n" + 
		  		"</Types>");
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
