package com.luxiaoyuan.excel.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;

import com.luxiaoyuan.excel.ExcelXMLContent;


public class AppXML {
	 private Writer _out;
	  private File xmlFile;
	  private FileOutputStream xmlOutStream;
	  public AppXML(String filePath){
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
		  		"<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" + 
		  		"	<Application>Microsoft Excel</Application>" + 
		  		"	<DocSecurity>0</DocSecurity>" + 
		  		"	<ScaleCrop>false</ScaleCrop>" + 
		  		"	<HeadingPairs>" + 
		  		"		<vt:vector size=\"2\" baseType=\"variant\">" + 
		  		"			<vt:variant>" + 
		  		"				<vt:lpstr>工作表</vt:lpstr>" + 
		  		"			</vt:variant>" + 
		  		"			<vt:variant>" + 
		  		"				<vt:i4>1</vt:i4>" + 
		  		"			</vt:variant>" + 
		  		"		</vt:vector>" + 
		  		"	</HeadingPairs>" + 
		  		"	<TitlesOfParts>" + 
		  		"		<vt:vector size=\"1\" baseType=\"lpstr\">" + 
		  		"			<vt:lpstr>sheet1</vt:lpstr>" + 
		  		"		</vt:vector>" + 
		  		"	</TitlesOfParts>" + 
		  		"	<LinksUpToDate>false</LinksUpToDate>" + 
		  		"	<SharedDoc>false</SharedDoc>" + 
		  		"	<HyperlinksChanged>false</HyperlinksChanged>" + 
		  		"	<AppVersion>16.0300</AppVersion>" + 
		  		"</Properties>");
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
