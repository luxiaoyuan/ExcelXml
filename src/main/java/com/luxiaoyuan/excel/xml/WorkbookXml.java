package com.luxiaoyuan.excel.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;

import com.luxiaoyuan.excel.ExcelXMLContent;


public class WorkbookXml {
	
	  private Writer _out;
	  private File xmlFile;
	  private FileOutputStream xmlOutStream;
	  private String filePath;
	  public WorkbookXml(String filePath){
        getWriter(filePath);
        this.filePath=filePath;
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

	  
	  public void writerOffice(){
		  try {
		  _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + 
		  		"<workbook mc:Ignorable=\"x15 xr xr6 xr10 xr2\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\">" + 
		  		"	<fileVersion appName=\"xl\" lastEdited=\"7\" lowestEdited=\"7\" rupBuild=\"21126\"/>" + 
		  		"	<workbookPr defaultThemeVersion=\"166925\"/>" + 
		  		"	<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" + 
		  		"		<mc:Choice Requires=\"x15\">" + 
		  		"			<x15ac:absPath url=\""+filePath+"\" xmlns:x15ac=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac\"/>" + 
		  		"		</mc:Choice>" + 
		  		"	</mc:AlternateContent>" + 
		  		"	<xr:revisionPtr revIDLastSave=\"0\" documentId=\"8_{F467CA9E-F5E1-4D5D-ACD1-A5E1BA8F2FA8}\" xr6:coauthVersionLast=\"40\" xr6:coauthVersionMax=\"40\" xr10:uidLastSave=\"{00000000-0000-0000-0000-000000000000}\"/>" + 
		  		"	<bookViews>" + 
		  		"		<workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"24000\" windowHeight=\"9465\" xr2:uid=\"{00000000-000D-0000-FFFF-FFFF00000000}\"/>" + 
		  		"	</bookViews>" + 
		  		"	<sheets>" + 
		  		"		<sheet name=\"sheet1\" sheetId=\"1\" r:id=\"rId1\"/>" + 
		  		"	</sheets>" + 
		  		"	<calcPr calcId=\"0\"/>" + 
		  		"	<fileRecoveryPr repairLoad=\"1\"/>" + 
		  		"</workbook>");
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
