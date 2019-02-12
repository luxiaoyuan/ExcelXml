package com.luxiaoyuan.excel.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;

import com.luxiaoyuan.excel.ExcelXMLContent;


public class CoreXML {
	 private Writer _out;
	  private File xmlFile;
	  private FileOutputStream xmlOutStream;
	  public CoreXML(String filePath){
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
		  		"<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" + 
		  		"	<dc:creator>Apache POI</dc:creator>" + 
		  		"	<cp:lastModifiedBy>da.guanyu</cp:lastModifiedBy>" + 
		  		"	<dcterms:created xsi:type=\"dcterms:W3CDTF\">2019-01-18T14:31:52Z</dcterms:created>" + 
		  		"	<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2019-01-18T14:32:48Z</dcterms:modified>" + 
		  		"</cp:coreProperties>" + 
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
}
