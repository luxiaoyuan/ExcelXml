package com.luxiaoyuan.excel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Writer;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import com.luxiaoyuan.excel.poi.ExcelHeader;
import com.luxiaoyuan.excel.poi.ExcelPicture;
import com.luxiaoyuan.excel.poi.ExcelResources;
import com.luxiaoyuan.excel.utils.LogUtils;
import com.luxiaoyuan.excel.xml.AppXML;
import com.luxiaoyuan.excel.xml.ContentTypesXML;
import com.luxiaoyuan.excel.xml.CoreXML;
import com.luxiaoyuan.excel.xml.DrawingXml;
import com.luxiaoyuan.excel.xml.ImageWriter;
import com.luxiaoyuan.excel.xml.ImageWriter.ImageBean;
import com.luxiaoyuan.excel.xml.SharedStringsXml;
import com.luxiaoyuan.excel.xml.SheetXml;
import com.luxiaoyuan.excel.xml.StylesXml;
import com.luxiaoyuan.excel.xml.ThemeXml;
import com.luxiaoyuan.excel.xml.WorkbookXml;
import com.luxiaoyuan.excel.xml.WorkbookXmlRels;
import com.luxiaoyuan.excel.xml.XmlRels.RelationshipBean;
import com.luxiaoyuan.excel.xml.XmlRels;

/**
 * Excel xml形式写入工具类
 * 说明
 * rels里面的 relationshipId 与sheet里面的对应
 * @author luxiaoyuan
 *
 */
@Component
public class ExcelXmlUtils {
    private static final String XML_ENCODING = "UTF-8";
    //public static final String PIC_PATH="D:\\GUANYU\\excel\\temp\\";//这里采用配置文件配置
    //静态方法中要用到的
    private static final String PIC_PATH_DRAWING="xl"+File.separator+"drawings"+File.separator;
	private static final String PIC_RELS="_rels"+File.separator;//相对目录  每个xml的rels文件在同级目录的_rels文件夹下
    private static final String PIC_WORKSHEETS="xl"+File.separator+"worksheets"+File.separator;
    //这里都采用代码写入  也可以固定在服务器上，但是不同的服务器还得配置
    //固定的几个配置文件   [Content_Types].xml   xl/styles.xml   /theme/theme1.xml
    private static final String PIC_XL="xl"+File.separator;//xl文件夹
    private static final String PIC_XL_THEME="theme"+File.separator;//xl文件夹
    private static final String PIC_XL_MEDIA="media"+File.separator;//xl文件夹
    private static final String PIC_DOCPROPS="docProps"+File.separator;
    
    private static final String PIC_IMAGE_URL_PREFIX="//img.alicdn.com/imgextra/";
    private ExcelXmlUtils() {}
    private static ExcelXmlUtils eu = new ExcelXmlUtils();
    public static ExcelXmlUtils getInstance() {
        return eu;
    }
    //logger
    private Logger logger = LogUtils.getUtilsLogger();
    private static final String EXCEL_TYPE="office";//office  wps
    /**
     * Create a library of cell styles.
     */
    private static Map<String, XSSFCellStyle> createStyles(XSSFWorkbook wb){
        Map<String, XSSFCellStyle> styles = new HashMap<>();
        XSSFDataFormat fmt = wb.createDataFormat();

        XSSFCellStyle style1 = wb.createCellStyle();
        style1.setAlignment(HorizontalAlignment.RIGHT);
        style1.setDataFormat(fmt.getFormat("0.0%"));
        styles.put("percent", style1);

        XSSFCellStyle style2 = wb.createCellStyle();
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setDataFormat(fmt.getFormat("0.0X"));
        styles.put("coeff", style2);

        XSSFCellStyle style3 = wb.createCellStyle();
        style3.setAlignment(HorizontalAlignment.RIGHT);
        style3.setDataFormat(fmt.getFormat("$#,##0.00"));
        styles.put("currency", style3);

        XSSFCellStyle style4 = wb.createCellStyle();
        style4.setAlignment(HorizontalAlignment.RIGHT);
        style4.setDataFormat(fmt.getFormat("mmm dd"));
        styles.put("date", style4);

        XSSFCellStyle style5 = wb.createCellStyle();
        XSSFFont headerFont = wb.createFont();
        headerFont.setBold(true);
        style5.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style5.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style5.setFont(headerFont);
        styles.put("header", style5);

        return styles;
    }

    private static void generate(Writer out, Map<String, XSSFCellStyle> styles) throws Exception {
        Random rnd = new Random();
        Calendar calendar = Calendar.getInstance();
        SpreadsheetWriter sw = new SpreadsheetWriter(out);
        sw.beginSheet();
        //insert header row
        sw.insertRow(0);
        int styleIndex = styles.get("header").getIndex();
        sw.createCell(0, "Title", styleIndex);
        sw.createCell(1, "% Change", styleIndex);
        sw.createCell(2, "Ratio", styleIndex);
        sw.createCell(3, "Expenses", styleIndex);
        sw.createCell(4, "Date", styleIndex);
        sw.createCell(5, "Picture", styleIndex);
        sw.endRow();
        try {
        for (int rownum = 1; rownum < 20; rownum++) {
            sw.insertRow(rownum);
            sw.createCell(0, "Hello, " + rownum + "!");
            sw.createCell(1, (double)rnd.nextInt(100)/100, styles.get("percent").getIndex());
            sw.createCell(2, (double)rnd.nextInt(10)/10, styles.get("coeff").getIndex());
            sw.createCell(3, rnd.nextInt(10000), styles.get("currency").getIndex());
            sw.createCell(4, calendar, styles.get("date").getIndex());
            sw.endRow();
            calendar.roll(Calendar.DAY_OF_YEAR, 1);
        }
        sw.endSheet();
        }
        catch(Exception e) {
        	e.printStackTrace();
        }
       
    }

    /**
     *
     * @param zipfile the template file
     * @param tmpfile the XML file with the sheet data
     * @param entry the name of the sheet entry to substitute, e.g. xl/worksheets/sheet1.xml
     * @param out the stream to write the result to
     */
    private static void substitute(File zipfile, File tmpfile, String entry, OutputStream out) throws IOException {
        try (ZipFile zip = ZipHelper.openZipFile(zipfile)) {
            try (ZipArchiveOutputStream zos = new ZipArchiveOutputStream(out)) {
                Enumeration<? extends ZipArchiveEntry> en = zip.getEntries();
                while (en.hasMoreElements()) {
                    ZipArchiveEntry ze = en.nextElement();
                    if (!ze.getName().equals(entry)) {
                        zos.putArchiveEntry(new ZipArchiveEntry(ze.getName()));
                        try (InputStream is = zip.getInputStream(ze)) {
                            copyStream(is, zos);
                        }
                        zos.closeArchiveEntry();
                    }
                }
                zos.putArchiveEntry(new ZipArchiveEntry(entry));
                try (InputStream is = new FileInputStream(tmpfile)) {
                    copyStream(is, zos);
                }
                zos.closeArchiveEntry();
            }
        }
    }
    
    
    /**
    *
    * @param zipfile the template file
    * @param tmpfileMap  多个文件的时   xml rels image
    * @param entry the name of the sheet entry to substitute, e.g. xl/worksheets/sheet1.xml
    * @param out the stream to write the result to
    */
   private static void substitute(File zipfile,List<TempFileBean> tmpfileLisst,OutputStream out) throws IOException {
       try (ZipFile zip = ZipHelper.openZipFile(zipfile)) {
    	   if (CollectionUtils.isEmpty(tmpfileLisst)){
    		   return;  //不写入
    	   }
    	   
    	   Map<String,File> tempFileMap=tmpfileLisst.
    			   stream().map(a->{a.setEntry(a.getEntry().replace("\\", "/"));//  \\替换 ，不然匹配不到的
    			   return a;})
    			   .collect(Collectors.toMap(TempFileBean::getEntry,TempFileBean::getTempFile,(k1,k2)->k1));
	           try (ZipArchiveOutputStream zos = new ZipArchiveOutputStream(out)) {
	                   Enumeration<? extends ZipArchiveEntry> en = zip.getEntries();
	                   while (en.hasMoreElements()) {   //写入其他的节点  相同的节点不写入
	                       ZipArchiveEntry ze = en.nextElement();
	                       if(tempFileMap.get(ze.getName())!=null) {
	                    	   //存在的节点  不写入
	                    	   continue;
	                       }
	                           zos.putArchiveEntry(new ZipArchiveEntry(ze.getName()));
	                           try (InputStream is = zip.getInputStream(ze)) {
	                               copyStream(is, zos);
	                           zos.closeArchiveEntry();
	                       }
	                   }
               		//写入生成的文件
	        	   tmpfileLisst.forEach(f->{
	        		   TempFileBean tempFileBean=f;
	                  
	                   if(tempFileBean.getEntry().indexOf("media")>-1) {
	                   String path=tempFileBean.path;
	                   File file=new File(path);
	                   if(file.isDirectory()) { //如果是一个文件夹  则循环的写入
	                	   String[] tempList = file.list();
	                	   int tempListLength=tempList.length;
	                	   if(tempListLength>0) {
	                		   for(int i=0;i<tempListLength;i++) {
	                			   File temp = null;
	                			   if (path.endsWith(File.separator)) {
	                		             temp = new File(path + tempList[i]);
	                		          } else {
	                		              temp = new File(path + File.separator + tempList[i]);
	                		          }
	                		          if (temp.isFile()) {
	                		        	  try {
											zos.putArchiveEntry(new ZipArchiveEntry(tempFileBean.getEntry()+"/"+temp.getName()));
										} catch (IOException e1) {
											e1.printStackTrace();
										}
	                		        	  try (InputStream is = new FileInputStream(temp)) {
	           								copyStream(is, zos);
	           								zos.closeArchiveEntry();
	                		        	  } catch (IOException e) {
											e.printStackTrace();
										}
	                		          }
	                		   }
	                	   }
	                   }
	                   }
	                   else {   //不是文件夹的
	                	   File file=tempFileBean.getTempFile();
	                	   try {
							zos.putArchiveEntry(new ZipArchiveEntry(tempFileBean.getEntry()));
						} catch (IOException e1) {
							e1.printStackTrace();
						}
		                   try (InputStream is = new FileInputStream(file)) {
								copyStream(is, zos);
								zos.closeArchiveEntry();
		                   } catch (FileNotFoundException e) {
							e.printStackTrace();
						} catch (IOException e) {
							e.printStackTrace();
						}
	                   }
	        	   });
	        	  // zos.closeArchiveEntry();
	           } catch (IOException e) {
				e.printStackTrace();
			}
	           
       }
   }

    private static void copyStream(InputStream in, OutputStream out) throws IOException {
        byte[] chunk = new byte[1024];
        int count;
        while ((count = in.read(chunk)) >=0 ) {
          out.write(chunk,0,count);
        }
    }

    /**
     * Writes spreadsheet data in a Writer.
     * (YK: in future it may evolve in a full-featured API for streaming data in Excel)
     */
    public static class SpreadsheetWriter {
        private final Writer _out;
        private int _rownum;

        SpreadsheetWriter(Writer out){
            _out = out;
        }

        void beginSheet() throws IOException {
            _out.write("<?xml version=\"1.0\" encoding=\""+XML_ENCODING+"\"?>" +
                    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" );
            _out.write("<sheetData>");
        }

        void endSheet() throws IOException {
            _out.write("</sheetData>");
            _out.write("</worksheet>");
        }

        /**
         * Insert a new row
         *
         * @param rownum 0-based row number
         */
        void insertRow(int rownum) throws IOException {
            _out.write("<row r=\""+(rownum+1)+"\">");
            this._rownum = rownum;
        }

        /**
         * Insert row end marker
         */
        void endRow() throws IOException {
            _out.write("</row>");
        }

        public void createCell(int columnIndex, String value, int styleIndex) throws IOException {
            String ref = new CellReference(_rownum, columnIndex).formatAsString();
            _out.write("<c r=\""+ref+"\" t=\"inlineStr\"");
            if(styleIndex != -1) {
                _out.write(" s=\""+styleIndex+"\"");
            }
            _out.write(">");
            _out.write("<is><t>"+value+"</t></is>");
            _out.write("</c>");
        }

        public void createCell(int columnIndex, String value) throws IOException {
            createCell(columnIndex, value, -1);
        }

        public void createCell(int columnIndex, double value, int styleIndex) throws IOException {
            String ref = new CellReference(_rownum, columnIndex).formatAsString();
            _out.write("<c r=\""+ref+"\" t=\"n\"");
            if(styleIndex != -1) {
                _out.write(" s=\""+styleIndex+"\"");
            }
            _out.write(">");
            _out.write("<v>"+value+"</v>");
            _out.write("</c>");
        }

        public void createCell(int columnIndex, double value) throws IOException {
            createCell(columnIndex, value, -1);
        }

        public void createCell(int columnIndex, Calendar value, int styleIndex) throws IOException {
            createCell(columnIndex, DateUtil.getExcelDate(value, false), styleIndex);
        }
        
       
    }
    
    
    /**
     * xml 形式写入excel
     * @param tempPath   写入文件临时目录  
     * @param fileName   写入的文件名称  带后缀 .xlsx
     * @param objs       写入的list
     * @param clz        写入的模板属性对象
     * @param sheetIndex 写入的sheet下标
     * @param picList    写入的图片list 
     * @return  String   返回写入成功的excel路径
     * @throws IOException
     */
    @SuppressWarnings("rawtypes")
	public String writerExcel(String tempPath,String fileName,List objs,Class clz,int sheetIndex,List<ExcelPicture> picList) throws IOException {
    	String excelFile=tempPath+fileName;//生成的文件
    	logger.error("writer xml excel start");
    	try {
    	if(tempPath==null) {
    		return "";
    	}
    	if(CollectionUtils.isEmpty(objs)) {
    		return "";
    	}
    	logger.error("xml excel list size:"+objs.size());
    	Integer pictureSize=0;
    	if(CollectionUtils.isNotEmpty(picList)) {
    	    pictureSize=picList.size();//判断社不设置行高及绘图形状
    	}
    	List<TempFileBean> tempList=new ArrayList<>();//写入文件bean
    	//等于第一个sheet的时候写入  其他的时候不写入固定的xml文件
    	if(sheetIndex==1) {
    		initDirectory(tempPath);//初始文件夹
    		//[Content_Types].xml   xl/styles.xml   /theme/theme1.xml
    		String contentTypesXmlEntry="[Content_Types].xml";
    		String contentTypesXmlPath=tempPath+contentTypesXmlEntry;
    		ContentTypesXML contentTypesXML=new ContentTypesXML(contentTypesXmlPath);
    		if(EXCEL_TYPE.equals("office")) {
    			contentTypesXML.writerOffice();
  			}else {
  				contentTypesXML.writer();//写入完成后才能添加filebean
  			}
    		//Content_Types ->tempbean
  		  	addTempBean(contentTypesXmlPath,contentTypesXmlEntry,tempList);
  		  
  		  	String stylesXmlEntry=PIC_XL+"styles.xml";
  		  	String stylesXmlPath=tempPath+stylesXmlEntry;
  		  	StylesXml stylesXML=new StylesXml(stylesXmlPath);
  			if(EXCEL_TYPE.equals("office")) {
  				stylesXML.writerOffice();
  			}else {
  				stylesXML.writer();//写入完成后才能添加filebean
  			}
  		  	//styles.xml ->tempbean
  		  	addTempBean(stylesXmlPath,stylesXmlEntry,tempList);
  		  	
  		  	//office模式   默认的wps  
  		  	if(EXCEL_TYPE.equals("office")) {
  		  		String workbookxml="workbook.xml";
  		  		//workbook.xml
  		  		String workbookXmlEntry=PIC_XL+workbookxml;
  		  		String workbookXmlPath=tempPath+workbookXmlEntry;
  		  		WorkbookXml workbookXml=new WorkbookXml(workbookXmlPath);
  		  		workbookXml.writerOffice();
  		  		addTempBean(workbookXmlPath,workbookXmlEntry,tempList);
  		  		
  		  		//workbook.xml.rels
				String workbookXmlRelsEntry=PIC_XL+PIC_RELS+workbookxml+".rels";
				String workbookXmlRelsPath=tempPath+workbookXmlRelsEntry;
  		  		WorkbookXmlRels workbookXmlRels=new WorkbookXmlRels(workbookXmlRelsPath);
  		  		workbookXmlRels.writerOffice();
  		  		addTempBean(workbookXmlRelsPath,workbookXmlRelsEntry,tempList);
  		  		
  		  		//写入 app core xml
  		  	    //app.xml
  		  		String appXml="app.xml";
  		  		String appXmlEntry=PIC_DOCPROPS+appXml;
  		  		String appXmlPATH=tempPath+appXmlEntry;
  		  		AppXML appXmlObj=new AppXML(appXmlPATH);
  		  		appXmlObj.writer();
  		  		addTempBean(appXmlPATH,appXmlEntry,tempList);
  		  		
  		  		//core.xml
  		  		String coreXml="core.xml";
  		  		String coreXmlEntry=PIC_DOCPROPS+coreXml;
  		  		String coreXmlPath=tempPath+coreXmlEntry;
  		  		CoreXML coreXmlObj=new CoreXML(coreXmlPath);
  		  		coreXmlObj.writer();
  		  		addTempBean(coreXmlPath, coreXmlEntry, tempList);
  		  				
  		  	}
  		  	
  		  	//这个可以写入多个    暂时只写入一个。多个还不知道长什么样
  	    	String themeXmlEntry=PIC_XL+PIC_XL_THEME+"theme"+sheetIndex+".xml";
  		  	String themeXmlPath=tempPath+themeXmlEntry;
  		  	ThemeXml themeXML=new ThemeXml(themeXmlPath);
  		  	if(EXCEL_TYPE.equals("office")) {
  		  		themeXML.writerOffice();
  		  	}else {
  		  		themeXML.writer();//写入完成后才能添加filebean
  		  	}
  		  	addTempBean(themeXmlPath,themeXmlEntry,tempList);
    	}
    	
    	//保存图片到media文件夹
    	String imageEntry="xl"+File.separator+"media";
    	String drawing1XML="drawing"+sheetIndex+".xml";
    	String drawing1XMLRel=drawing1XML+".rels";
    	String imagePath=tempPath+imageEntry;
    
    	//清空里面的图片信息
    	FileUtils.delFolder(imagePath);
    	ImageWriter imagewriter=new ImageWriter();
    	List<ImageBean> imageList=new ArrayList<>();
    	if(CollectionUtils.isNotEmpty(picList)) {
	    		try {
			    	picList.forEach(p->{
				    		String url=p.getUrl();
				    		if(StringUtils.isNotBlank(url)) {
				    			ImageBean imageBean=imagewriter.new ImageBean();
				    			if(url.indexOf("http")<0) {
				    				url="http:"+p.getUrl();
								}
					    		if(url.indexOf("https:\\")>-1) {
					    			url.replace("https:\\", "https://");
					    		}
					    		if(url.indexOf("\\")>-1) {
					    			url.replace("\\", "/");
					    		}
					    		//包含了url.indexOf("http")<0 并且没有http的时候
					    		if(url.startsWith(PIC_IMAGE_URL_PREFIX)&&url.indexOf("http")>-1) {
					    			url=url.substring(PIC_IMAGE_URL_PREFIX.length(), url.length());
					    		}
					    		if(url.indexOf("http")>-1
										&&(url.endsWith("jpg")||url.endsWith("JPG")
										||url.endsWith("png")||url.endsWith("PNG")
										||url.endsWith("jpeg")||url.endsWith("JPEG")
										||url.endsWith("gif")||url.endsWith("GIF")
										||url.endsWith("bmp")||url.endsWith("BMP")
										))  {  //验证图片的有效性
					    			imageBean.setImageUrl(url);
							    	imageBean.setCell(p.getStartIndex());
							    	imageBean.setRow(p.getStartRow());
							    	imageList.add(imageBean);
								}
						    	
			    		}
			    	});
			    	imagewriter.writerData(imageList,tempPath);
			    		//media ->tempbean
					addTempBeanImage(imagePath,imageEntry,tempList);
	    	}catch(Exception e) {
	    		logger.error("writer image异常:",e);
	    	}
    	
	        //写入drawring1.xml.rels drawing1.XML
	        String drawing1RelPath=tempPath+PIC_PATH_DRAWING+PIC_RELS+drawing1XMLRel;
	        XmlRels sw = new XmlRels(drawing1RelPath);
	        //sw.writerSingle(relationshipId, "../media/"+imageName, XmlRels.relationShip_type_image);
	        List<RelationshipBean> relationList=new ArrayList<>();
	        for(ImageBean img:imageList) {
	        	RelationshipBean relation=sw.new RelationshipBean();
	        	relation.setRid(img.getRelationshipId());
	        	relation.setTarget(ExcelXMLContent.DRAWING_XML_RELS_PREFIX+img.getImageName());
	        	relation.setType(XmlRels.relationShip_type_image);
	        	relationList.add(relation);
	        }
	        sw.writerMany(relationList);
	        
	        //drawing rels->tempBean
	        String drawing1RelEntry=PIC_PATH_DRAWING+PIC_RELS+drawing1XMLRel;
	        addTempBean(drawing1RelPath,drawing1RelEntry,tempList);
	        //drawing1.xml
	      
	        String drawing1XMLPath=tempPath+PIC_PATH_DRAWING+drawing1XML;
	        DrawingXml drawing1SW = new DrawingXml(drawing1XMLPath);
	        drawing1SW.writerData(imageList);
	        
	        //drawing1.xml ->tempbean
	        String drawing1Entry=PIC_PATH_DRAWING+drawing1XML;
	        addTempBean(drawing1XMLPath,drawing1Entry,tempList);
    	}
        //xl\worksheet\_rels\sheet1.xml.rels  这里的rid  与sheet里面的对应
        String sheet1XMLRels="sheet"+sheetIndex+".xml.rels";
        String sheet1XMLRelsPath=tempPath+PIC_WORKSHEETS+PIC_RELS+sheet1XMLRels;
        //引用drawing.xml  这里是单个引用 多个时请自定义(要扩展的话)
        XmlRels sheetXmlRels = new XmlRels(sheet1XMLRelsPath);
        String drawing1XmlRelsId="rId1";
        sheetXmlRels.writerSingle(drawing1XmlRelsId, ExcelXMLContent.SHEET_XML_RELS_PREFIX+drawing1XML, XmlRels.relationShip_type_drawing);
        //sheet1.xml.rels ->tempbean
        String sheetRels1Entry=PIC_WORKSHEETS+PIC_RELS+sheet1XMLRels;
        addTempBean(sheet1XMLRelsPath,sheetRels1Entry,tempList);
        
        //xl\worksheet\sheet1.xml
        List<ExcelHeader> headers = getHeaderList(clz);
        Collections.sort(headers);
        String sheet1XML="sheet"+sheetIndex+".xml";
        String sheet1XMLPath=tempPath+PIC_WORKSHEETS+sheet1XML;
        SheetXml sheetXML=new SheetXml(sheet1XMLPath);
        if(EXCEL_TYPE.equals("office")) {
            	//sharedStrings.xml
            	Map<String,Integer> sMap=new LinkedHashMap<>();//回写数据的下标值
            	String SharedStringsXmlEntry=PIC_XL+"sharedStrings.xml";
            	String SharedStringsXmlPath=tempPath+SharedStringsXmlEntry;
            	SharedStringsXml sharedStringsXml=new SharedStringsXml(SharedStringsXmlPath);
            	sharedStringsXml.writerDataOffice(headers, objs, sMap);
            	addTempBean(SharedStringsXmlPath,SharedStringsXmlEntry,tempList);
            	sheetXML.writerDataOffice(headers, objs, drawing1XmlRelsId, sMap,pictureSize);
        }else {
        	sheetXML.writerData(headers,objs,drawing1XmlRelsId);	
        }
        //sheet1.xml.rels ->tempbean
        String sheetXmlEntry=PIC_WORKSHEETS+sheet1XML;
        TempFileBean sheet1XMLBean=new ExcelXmlUtils().new TempFileBean();
        sheet1XMLBean.setTempFile(new File(sheet1XMLPath));
        sheet1XMLBean.setEntry(sheetXmlEntry);
        addTempBean(sheet1XMLPath,sheetXmlEntry,tempList);
        
        //生成临时的 template.xlsx
        String template=tempPath+"template.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook();
        @SuppressWarnings("unused")
		XSSFSheet sheet = wb.createSheet("sheet1");
        //Map<String, XSSFCellStyle> styles = createStyles(wb);
        //name of the zip entry holding sheet data, e.g. /xl/worksheets/sheet1.xml
        //String sheetRef = sheet.getPackagePart().getPartName().getName();
        //save the template
        FileOutputStream os =null;
        try  {
        	os= new FileOutputStream(template);
            wb.write(os);
        }catch(Exception e) {
        	e.printStackTrace();
        	logger.error("template writer异常:",e);
        }finally {
        	try {
				os.close();
			} catch (IOException e) {
				e.printStackTrace();
				logger.error("template writer close异常:",e);
			}
        }
        
        //Step 3. Substitute the template entry with the generated data
        
        FileOutputStream out=null;
        try {
        	out = new FileOutputStream(excelFile);
            substitute(new File(template), tempList,out);
        }catch(Exception e) {
        	e.printStackTrace();
        	logger.error("excelFile writer异常:",e);
        }finally {
        	try {
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
				logger.error("excelFile writer close异常:",e);
			}
        }
    	}catch(Exception e) {
    		e.printStackTrace();
    		logger.error("excelFile writer异常:",e);
    	}
    	logger.error("writer xml excel end");
        return excelFile;
        
    }
    
    //初始文件夹
    public void initDirectory(String tempPath) {
    	File drawingfile=new File(tempPath+PIC_XL+PIC_RELS);  // xl/_rels
    	drawingfile.mkdirs();  // 多级目录
    	drawingfile=new File(tempPath+PIC_DOCPROPS);  // docProps
    	drawingfile.mkdirs();  // 多级目录
    	drawingfile=new File(tempPath+PIC_XL+PIC_XL_THEME); // xl/theme
    	drawingfile.mkdirs();
    	drawingfile=new File(tempPath+PIC_XL+PIC_XL_MEDIA); // xl /media
    	drawingfile.mkdirs();
    	drawingfile=new File(tempPath+PIC_PATH_DRAWING+PIC_RELS);  // xl/drawing/_rels
    	drawingfile.mkdirs();
    	drawingfile=new File(tempPath+PIC_WORKSHEETS+PIC_RELS); //  xl/worksheets/_res
    	drawingfile.mkdirs();
    	drawingfile=null;
    }
    
    
    /**
     * //添加生成的文件bean
     * @param path  文件路径
     * @param entry excel相对的文件路径名称
     */
    public void addTempBean(String path,String entry,List<TempFileBean> tempList) {
    	//.xml or .xml.rels ->tempbean
	  	TempFileBean themeTemBean=new ExcelXmlUtils().new TempFileBean();
	  	File themeFileFolder=new File(path);
	  	if(themeFileFolder.exists()) {
	  		themeTemBean.setTempFile(themeFileFolder);
	  		themeTemBean.setEntry(entry);
	  		tempList.add(themeTemBean);
	  	} //else find not file
    }
    /**
     * //添加生成的文件bean
     * @param path  文件目录  ->图片时
     * @param entry excel相对的文件路径名称
     */
    public void addTempBeanImage(String path,String entry,List<TempFileBean> tempList) {
    	//.xml or .xml.rels ->tempbean
	  	TempFileBean themeTemBean=new ExcelXmlUtils().new TempFileBean();
	  	File themeFileFolder=new File(path);
	  	if(themeFileFolder.exists()) {
	  		themeTemBean.setTempFile(themeFileFolder);
	  		themeTemBean.setPath(path);
	  		themeTemBean.setEntry(entry);
	  		tempList.add(themeTemBean);
	  	} //else find not file
    }
  
	
    
    
    
    public static List<ExcelHeader> getHeaderList(@SuppressWarnings("rawtypes") Class clz) {
        List<ExcelHeader> headers = new ArrayList<ExcelHeader>();
        Method[] ms = clz.getDeclaredMethods();
        for(Method m:ms) {
            String mn = m.getName();
            if(mn.startsWith("get")) {
                if(m.isAnnotationPresent(ExcelResources.class)) {
                    ExcelResources er = m.getAnnotation(ExcelResources.class);
                    headers.add(new ExcelHeader(er.title(),er.order(),mn,er.cellWidth()));
                }
            }
        }
        return headers;
    }
    
    

    //临时文件bean
    public class TempFileBean {
    	private String entry;//文件路径名称  切记不是文件名称 xl//worksheet/sheet.xml
 	    private File tempFile;//临时文件
 	    private String path;//针对文件夹
		public File getTempFile() {
			return tempFile;
		}
		public void setTempFile(File tempFile) {
			this.tempFile = tempFile;
		}
		public String getEntry() {
			return entry;
		}
		public void setEntry(String entry) {
			this.entry = entry;
		}
		public String getPath() {
			return path;
		}
		public void setPath(String path) {
			this.path = path;
		}
		
 	    
 	    
    }
}

