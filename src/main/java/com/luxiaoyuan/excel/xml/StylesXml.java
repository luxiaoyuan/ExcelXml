package com.luxiaoyuan.excel.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;

import com.luxiaoyuan.excel.ExcelXMLContent;


public class StylesXml {
	
	 private Writer _out;
	  private File xmlFile;
	  private FileOutputStream xmlOutStream;
	  public StylesXml(String filePath){
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

	  
	  public void writer(){
		  try {
		  _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + 
		  		"<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + 
		  		"	<numFmts count=\"8\">" + 
		  		"		<numFmt numFmtId=\"42\" formatCode=\"_ &quot;￥&quot;* #,##0_ ;_ &quot;￥&quot;* \\-#,##0_ ;_ &quot;￥&quot;* &quot;-&quot;_ ;_ @_ \"/>" + 
		  		"		<numFmt numFmtId=\"44\" formatCode=\"_ &quot;￥&quot;* #,##0.00_ ;_ &quot;￥&quot;* \\-#,##0.00_ ;_ &quot;￥&quot;* &quot;-&quot;??_ ;_ @_ \"/>" + 
		  		"		<numFmt numFmtId=\"43\" formatCode=\"_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * &quot;-&quot;??_ ;_ @_ \"/>" + 
		  		"		<numFmt numFmtId=\"176\" formatCode=\"\\$#,##0.00\"/>" + 
		  		"		<numFmt numFmtId=\"177\" formatCode=\"mmm\\ dd\"/>" + 
		  		"		<numFmt numFmtId=\"178\" formatCode=\"0.0#\"/>" + 
		  		"		<numFmt numFmtId=\"179\" formatCode=\"0.0%\"/>" + 
		  		"		<numFmt numFmtId=\"41\" formatCode=\"_ * #,##0_ ;_ * \\-#,##0_ ;_ * &quot;-&quot;_ ;_ @_ \"/>" + 
		  		"	</numFmts>" + 
		  		"	<fonts count=\"22\">" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color indexed=\"8\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<b/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<name val=\"Calibri\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color theme=\"1\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<b/>" + 
		  		"			<sz val=\"13\"/>" + 
		  		"			<color theme=\"3\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FFFF0000\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color theme=\"0\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FF9C0006\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color theme=\"1\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FF9C6500\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<b/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FFFFFFFF\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FF3F3F76\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FFFA7D00\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<b/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FFFA7D00\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FF006100\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<b/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color theme=\"3\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<b/>" + 
		  		"			<sz val=\"18\"/>" + 
		  		"			<color theme=\"3\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<u/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FF0000FF\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<i/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FF7F7F7F\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<u/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FF800080\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<b/>" + 
		  		"			<sz val=\"15\"/>" + 
		  		"			<color theme=\"3\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<b/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color rgb=\"FF3F3F3F\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<b/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color theme=\"1\"/>" + 
		  		"			<name val=\"宋体\"/>" + 
		  		"			<charset val=\"0\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"	</fonts>" + 
		  		"	<fills count=\"34\">" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"none\"/>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"gray125\"/>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor indexed=\"22\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"9\" tint=\"0.399975585192419\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"4\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"4\" tint=\"0.399975585192419\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"5\" tint=\"0.399975585192419\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor rgb=\"FFFFC7CE\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"8\" tint=\"0.399975585192419\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"4\" tint=\"0.799981688894314\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor rgb=\"FFFFEB9C\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"6\" tint=\"0.599993896298105\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"9\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"5\" tint=\"0.599993896298105\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor rgb=\"FFA5A5A5\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor rgb=\"FFFFCC99\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"9\" tint=\"0.599993896298105\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"7\" tint=\"0.799981688894314\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor rgb=\"FFF2F2F2\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"6\" tint=\"0.799981688894314\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"8\" tint=\"0.599993896298105\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor rgb=\"FFC6EFCE\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"5\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"6\" tint=\"0.399975585192419\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor rgb=\"FFFFFFCC\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"7\" tint=\"0.599993896298105\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"7\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"7\" tint=\"0.399975585192419\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"6\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"9\" tint=\"0.799981688894314\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"8\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"5\" tint=\"0.799981688894314\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"4\" tint=\"0.599993896298105\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"solid\">" + 
		  		"				<fgColor theme=\"8\" tint=\"0.799981688894314\"/>" + 
		  		"				<bgColor indexed=\"64\"/>" + 
		  		"			</patternFill>" + 
		  		"		</fill>" + 
		  		"	</fills>" + 
		  		"	<borders count=\"9\">" + 
		  		"		<border>" + 
		  		"			<left/>" + 
		  		"			<right/>" + 
		  		"			<top/>" + 
		  		"			<bottom/>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"		<border>" + 
		  		"			<left/>" + 
		  		"			<right/>" + 
		  		"			<top/>" + 
		  		"			<bottom style=\"medium\">" + 
		  		"				<color theme=\"4\"/>" + 
		  		"			</bottom>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"		<border>" + 
		  		"			<left style=\"double\">" + 
		  		"				<color rgb=\"FF3F3F3F\"/>" + 
		  		"			</left>" + 
		  		"			<right style=\"double\">" + 
		  		"				<color rgb=\"FF3F3F3F\"/>" + 
		  		"			</right>" + 
		  		"			<top style=\"double\">" + 
		  		"				<color rgb=\"FF3F3F3F\"/>" + 
		  		"			</top>" + 
		  		"			<bottom style=\"double\">" + 
		  		"				<color rgb=\"FF3F3F3F\"/>" + 
		  		"			</bottom>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"		<border>" + 
		  		"			<left style=\"thin\">" + 
		  		"				<color rgb=\"FF7F7F7F\"/>" + 
		  		"			</left>" + 
		  		"			<right style=\"thin\">" + 
		  		"				<color rgb=\"FF7F7F7F\"/>" + 
		  		"			</right>" + 
		  		"			<top style=\"thin\">" + 
		  		"				<color rgb=\"FF7F7F7F\"/>" + 
		  		"			</top>" + 
		  		"			<bottom style=\"thin\">" + 
		  		"				<color rgb=\"FF7F7F7F\"/>" + 
		  		"			</bottom>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"		<border>" + 
		  		"			<left/>" + 
		  		"			<right/>" + 
		  		"			<top/>" + 
		  		"			<bottom style=\"double\">" + 
		  		"				<color rgb=\"FFFF8001\"/>" + 
		  		"			</bottom>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"		<border>" + 
		  		"			<left/>" + 
		  		"			<right/>" + 
		  		"			<top/>" + 
		  		"			<bottom style=\"medium\">" + 
		  		"				<color theme=\"4\" tint=\"0.499984740745262\"/>" + 
		  		"			</bottom>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"		<border>" + 
		  		"			<left style=\"thin\">" + 
		  		"				<color rgb=\"FFB2B2B2\"/>" + 
		  		"			</left>" + 
		  		"			<right style=\"thin\">" + 
		  		"				<color rgb=\"FFB2B2B2\"/>" + 
		  		"			</right>" + 
		  		"			<top style=\"thin\">" + 
		  		"				<color rgb=\"FFB2B2B2\"/>" + 
		  		"			</top>" + 
		  		"			<bottom style=\"thin\">" + 
		  		"				<color rgb=\"FFB2B2B2\"/>" + 
		  		"			</bottom>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"		<border>" + 
		  		"			<left style=\"thin\">" + 
		  		"				<color rgb=\"FF3F3F3F\"/>" + 
		  		"			</left>" + 
		  		"			<right style=\"thin\">" + 
		  		"				<color rgb=\"FF3F3F3F\"/>" + 
		  		"			</right>" + 
		  		"			<top style=\"thin\">" + 
		  		"				<color rgb=\"FF3F3F3F\"/>" + 
		  		"			</top>" + 
		  		"			<bottom style=\"thin\">" + 
		  		"				<color rgb=\"FF3F3F3F\"/>" + 
		  		"			</bottom>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"		<border>" + 
		  		"			<left/>" + 
		  		"			<right/>" + 
		  		"			<top style=\"thin\">" + 
		  		"				<color theme=\"4\"/>" + 
		  		"			</top>" + 
		  		"			<bottom style=\"double\">" + 
		  		"				<color theme=\"4\"/>" + 
		  		"			</bottom>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"	</borders>" + 
		  		"	<cellStyleXfs count=\"49\">" + 
		  		"		<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"42\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"19\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"10\" fillId=\"15\" borderId=\"3\" applyNumberFormat=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"44\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"41\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"11\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"6\" fillId=\"7\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"43\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"23\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"16\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"9\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"18\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"2\" fillId=\"24\" borderId=\"6\" applyNumberFormat=\"0\" applyFont=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"6\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"14\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"15\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"17\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"19\" fillId=\"0\" borderId=\"1\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"1\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"5\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"14\" fillId=\"0\" borderId=\"5\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"27\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"20\" fillId=\"18\" borderId=\"7\" applyNumberFormat=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"12\" fillId=\"18\" borderId=\"3\" applyNumberFormat=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"9\" fillId=\"14\" borderId=\"2\" applyNumberFormat=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"29\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"22\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"11\" fillId=\"0\" borderId=\"4\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"21\" fillId=\"0\" borderId=\"8\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"13\" fillId=\"21\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"8\" fillId=\"10\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"33\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"4\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"9\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"32\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"31\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"13\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"28\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"26\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"17\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"25\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"30\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"20\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"8\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"12\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"7\" fillId=\"16\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"5\" fillId=\"3\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"	</cellStyleXfs>" + 
		  		"	<cellXfs count=\"6\">" + 
		  		"		<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyAlignment=\"1\"/>" + 
		  		"		<xf numFmtId=\"179\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyAlignment=\"1\">" + 
		  		"			<alignment horizontal=\"right\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"178\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyAlignment=\"1\">" + 
		  		"			<alignment horizontal=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"176\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyAlignment=\"1\">" + 
		  		"			<alignment horizontal=\"right\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"177\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyAlignment=\"1\">" + 
		  		"			<alignment horizontal=\"right\"/>" + 
		  		"		</xf>" + 
		  		"	</cellXfs>" + 
		  		"	<cellStyles count=\"49\">" + 
		  		"		<cellStyle name=\"常规\" xfId=\"0\" builtinId=\"0\"/>" + 
		  		"		<cellStyle name=\"货币[0]\" xfId=\"1\" builtinId=\"7\"/>" + 
		  		"		<cellStyle name=\"20% - 强调文字颜色 3\" xfId=\"2\" builtinId=\"38\"/>" + 
		  		"		<cellStyle name=\"输入\" xfId=\"3\" builtinId=\"20\"/>" + 
		  		"		<cellStyle name=\"货币\" xfId=\"4\" builtinId=\"4\"/>" + 
		  		"		<cellStyle name=\"千位分隔[0]\" xfId=\"5\" builtinId=\"6\"/>" + 
		  		"		<cellStyle name=\"40% - 强调文字颜色 3\" xfId=\"6\" builtinId=\"39\"/>" + 
		  		"		<cellStyle name=\"差\" xfId=\"7\" builtinId=\"27\"/>" + 
		  		"		<cellStyle name=\"千位分隔\" xfId=\"8\" builtinId=\"3\"/>" + 
		  		"		<cellStyle name=\"60% - 强调文字颜色 3\" xfId=\"9\" builtinId=\"40\"/>" + 
		  		"		<cellStyle name=\"超链接\" xfId=\"10\" builtinId=\"8\"/>" + 
		  		"		<cellStyle name=\"百分比\" xfId=\"11\" builtinId=\"5\"/>" + 
		  		"		<cellStyle name=\"已访问的超链接\" xfId=\"12\" builtinId=\"9\"/>" + 
		  		"		<cellStyle name=\"注释\" xfId=\"13\" builtinId=\"10\"/>" + 
		  		"		<cellStyle name=\"60% - 强调文字颜色 2\" xfId=\"14\" builtinId=\"36\"/>" + 
		  		"		<cellStyle name=\"标题 4\" xfId=\"15\" builtinId=\"19\"/>" + 
		  		"		<cellStyle name=\"警告文本\" xfId=\"16\" builtinId=\"11\"/>" + 
		  		"		<cellStyle name=\"标题\" xfId=\"17\" builtinId=\"15\"/>" + 
		  		"		<cellStyle name=\"解释性文本\" xfId=\"18\" builtinId=\"53\"/>" + 
		  		"		<cellStyle name=\"标题 1\" xfId=\"19\" builtinId=\"16\"/>" + 
		  		"		<cellStyle name=\"标题 2\" xfId=\"20\" builtinId=\"17\"/>" + 
		  		"		<cellStyle name=\"60% - 强调文字颜色 1\" xfId=\"21\" builtinId=\"32\"/>" + 
		  		"		<cellStyle name=\"标题 3\" xfId=\"22\" builtinId=\"18\"/>" + 
		  		"		<cellStyle name=\"60% - 强调文字颜色 4\" xfId=\"23\" builtinId=\"44\"/>" + 
		  		"		<cellStyle name=\"输出\" xfId=\"24\" builtinId=\"21\"/>" + 
		  		"		<cellStyle name=\"计算\" xfId=\"25\" builtinId=\"22\"/>" + 
		  		"		<cellStyle name=\"检查单元格\" xfId=\"26\" builtinId=\"23\"/>" + 
		  		"		<cellStyle name=\"20% - 强调文字颜色 6\" xfId=\"27\" builtinId=\"50\"/>" + 
		  		"		<cellStyle name=\"强调文字颜色 2\" xfId=\"28\" builtinId=\"33\"/>" + 
		  		"		<cellStyle name=\"链接单元格\" xfId=\"29\" builtinId=\"24\"/>" + 
		  		"		<cellStyle name=\"汇总\" xfId=\"30\" builtinId=\"25\"/>" + 
		  		"		<cellStyle name=\"好\" xfId=\"31\" builtinId=\"26\"/>" + 
		  		"		<cellStyle name=\"适中\" xfId=\"32\" builtinId=\"28\"/>" + 
		  		"		<cellStyle name=\"20% - 强调文字颜色 5\" xfId=\"33\" builtinId=\"46\"/>" + 
		  		"		<cellStyle name=\"强调文字颜色 1\" xfId=\"34\" builtinId=\"29\"/>" + 
		  		"		<cellStyle name=\"20% - 强调文字颜色 1\" xfId=\"35\" builtinId=\"30\"/>" + 
		  		"		<cellStyle name=\"40% - 强调文字颜色 1\" xfId=\"36\" builtinId=\"31\"/>" + 
		  		"		<cellStyle name=\"20% - 强调文字颜色 2\" xfId=\"37\" builtinId=\"34\"/>" + 
		  		"		<cellStyle name=\"40% - 强调文字颜色 2\" xfId=\"38\" builtinId=\"35\"/>" + 
		  		"		<cellStyle name=\"强调文字颜色 3\" xfId=\"39\" builtinId=\"37\"/>" + 
		  		"		<cellStyle name=\"强调文字颜色 4\" xfId=\"40\" builtinId=\"41\"/>" + 
		  		"		<cellStyle name=\"20% - 强调文字颜色 4\" xfId=\"41\" builtinId=\"42\"/>" + 
		  		"		<cellStyle name=\"40% - 强调文字颜色 4\" xfId=\"42\" builtinId=\"43\"/>" + 
		  		"		<cellStyle name=\"强调文字颜色 5\" xfId=\"43\" builtinId=\"45\"/>" + 
		  		"		<cellStyle name=\"40% - 强调文字颜色 5\" xfId=\"44\" builtinId=\"47\"/>" + 
		  		"		<cellStyle name=\"60% - 强调文字颜色 5\" xfId=\"45\" builtinId=\"48\"/>" + 
		  		"		<cellStyle name=\"强调文字颜色 6\" xfId=\"46\" builtinId=\"49\"/>" + 
		  		"		<cellStyle name=\"40% - 强调文字颜色 6\" xfId=\"47\" builtinId=\"51\"/>" + 
		  		"		<cellStyle name=\"60% - 强调文字颜色 6\" xfId=\"48\" builtinId=\"52\"/>" + 
		  		"	</cellStyles>" + 
		  		"	<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>" + 
		  		"	<extLst>" + 
		  		"		<ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" + 
		  		"			<x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/>" + 
		  		"		</ext>" + 
		  		"	</extLst>" + 
		  		"</styleSheet>" + 
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
	
	  //office styles
	  public void writerOffice(){
		  try {
		  _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + 
		  		"<styleSheet mc:Ignorable=\"x14ac x16r2 xr\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\">" + 
		  		"	<fonts count=\"3\" x14ac:knownFonts=\"1\">" + 
		  		"		<font>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color theme=\"1\"/>" + 
		  		"			<name val=\"等线\"/>" + 
		  		"			<family val=\"2\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<sz val=\"9\"/>" + 
		  		"			<name val=\"等线\"/>" + 
		  		"			<family val=\"2\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"		<font>" + 
		  		"			<u/>" + 
		  		"			<sz val=\"11\"/>" + 
		  		"			<color theme=\"10\"/>" + 
		  		"			<name val=\"等线\"/>" + 
		  		"			<family val=\"2\"/>" + 
		  		"			<charset val=\"134\"/>" + 
		  		"			<scheme val=\"minor\"/>" + 
		  		"		</font>" + 
		  		"	</fonts>" + 
		  		"	<fills count=\"2\">" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"none\"/>" + 
		  		"		</fill>" + 
		  		"		<fill>" + 
		  		"			<patternFill patternType=\"gray125\"/>" + 
		  		"		</fill>" + 
		  		"	</fills>" + 
		  		"	<borders count=\"1\">" + 
		  		"		<border>" + 
		  		"			<left/>" + 
		  		"			<right/>" + 
		  		"			<top/>" + 
		  		"			<bottom/>" + 
		  		"			<diagonal/>" + 
		  		"		</border>" + 
		  		"	</borders>" + 
		  		"	<cellStyleXfs count=\"2\">" + 
		  		"		<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"	</cellStyleXfs>" + 
		  		"	<cellXfs count=\"2\">" + 
		  		"		<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"		<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"1\">" + 
		  		"			<alignment vertical=\"center\"/>" + 
		  		"		</xf>" + 
		  		"	</cellXfs>" + 
		  		"	<cellStyles count=\"2\">" + 
		  		"		<cellStyle name=\"常规\" xfId=\"0\" builtinId=\"0\"/>" + 
		  		"		<cellStyle name=\"超链接\" xfId=\"1\" builtinId=\"8\"/>" + 
		  		"	</cellStyles>" + 
		  		"	<dxfs count=\"0\"/>" + 
		  		"	<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>" + 
		  		"	<extLst>" + 
		  		"		<ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" + 
		  		"			<x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/>" + 
		  		"		</ext>" + 
		  		"		<ext uri=\"{9260A510-F301-46a8-8635-F512D64BE5F5}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\">" + 
		  		"			<x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\"/>" + 
		  		"		</ext>" + 
		  		"	</extLst>" + 
		  		"</styleSheet>" + 
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
