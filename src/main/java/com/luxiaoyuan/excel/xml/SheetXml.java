package com.luxiaoyuan.excel.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellReference;
import org.slf4j.Logger;

import com.luxiaoyuan.excel.ExcelXMLContent;
import com.luxiaoyuan.excel.poi.ExcelHeader;
import com.luxiaoyuan.excel.utils.LogUtils;


/**
 * 写入drawing.xml 单个或者多个时根据自定义
 * default xl\\drawings\\drawing1.xml  windows格式
 * @author luxiaoyuan
 *
 */

public class SheetXml {
	  private Logger logger=LogUtils.getUtilsLogger();
	  private Writer _out;
	  private File xmlRelFile;
	  private FileOutputStream xmlRelOutStream;
	  public SheetXml(String filePath){
          getWriter(filePath);
      }
	  
	  public void getWriter(String relPath) {
			 xmlRelFile = new File(relPath);  
			 if(xmlRelFile.exists()) {
		    	 xmlRelFile.delete();//删除  每次都是新的写入
		     }
			 try {
				xmlRelFile.createNewFile();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		        //创建输出流  
		     xmlRelOutStream=null;
			try {
				xmlRelOutStream = new FileOutputStream(xmlRelFile);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} 
		        //写入xml标签  后面使用dom4j
		     Writer fw=null;
			try {
				fw = new OutputStreamWriter(xmlRelOutStream, ExcelXMLContent.XML_ENCODING);
			} catch (UnsupportedEncodingException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		     
		     _out = fw;
		 }
	  
	  //insert xml Header
      public void insertXmlHeader() throws IOException {
      	 _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
      }
      
      //begin worksheet
      public void beginWorkSheet() throws IOException {
    	  _out.write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:etc=\"http://www.wps.cn/officeDocument/2017/etCustomData\">");
      }
      
      public void beginWorkSheetOfffice() throws IOException {
    	  _out.write("<worksheet mc:Ignorable=\"x14ac xr xr2 xr3\" xr:uid=\"{00000000-0001-0000-0000-000000000000}\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\">");
      }
      
      //end sheetend
      public void  endWorkSheet() throws IOException {
    	  _out.write("</worksheet>");
      }
	
      // end sheetPr
      public void endSheetPr() throws IOException {
    	  _out.write("<sheetPr/>");
      }
      
      // insert dimension  A1:AH11
      public void insertDimension(String dimension) throws IOException {
    	  _out.write("<dimension ref=\""+dimension+"\"/>");
      }
      
      //insert SheetView
      public void insertSheetView() throws IOException {
    	  _out.write("<sheetViews>");
    	  	_out.write("<sheetView tabSelected=\"1\" workbookViewId=\"0\">");
    	  		_out.write("<selection activeCell=\"E3\" sqref=\"E3\"/>");
    	  	_out.write("</sheetView>");
    	  _out.write("</sheetViews>");
      }
      
      //insert sheetFormatPr
      public void insertSheetFormatPr() throws IOException {
    	  _out.write("<sheetFormatPr defaultColWidth=\"9\" defaultRowHeight=\"13.5\" outlineLevelRow=\"3\" outlineLevelCol=\"5\"/>");
      }
      public void insertSheetFormatPrOffice() throws IOException {
    	  _out.write("<sheetFormatPr defaultColWidth=\"9\" defaultRowHeight=\"14.25\" x14ac:dyDescent=\"0.2\"/>");
      }
      
      //insert cols
      public void insertCols() throws IOException {
    	  _out.write("<cols>");
    	  	_out.write("<col min=\"1\" max=\"1\" width=\"50\" customWidth=\"1\"/>");
    	  	_out.write("<col min=\"2\" max=\"2\" width=\"40.125\" customWidth=\"1\"/>");
    	  	_out.write("<col min=\"3\" max=\"3\" width=\"50\" customWidth=\"1\"/>");
    	  	_out.write("<col min=\"4\" max=\"4\" width=\"40.125\" customWidth=\"1\"/>");
    	  	_out.write("<col min=\"5\" max=\"5\" width=\"50\" customWidth=\"1\"/>");
    	  	_out.write("<col min=\"6\" max=\"6\" width=\"40.125\" customWidth=\"1\"/>");
    	  _out.write("</cols>");
      }
      
      //insert cols
      public void insertCols(List<ColsBean> list) throws IOException {
    	  _out.write("<cols>");
    	    list.forEach(w->{
    	    	try {
					_out.write("<col min=\""+w.getCell()+"\" max=\""+w.getCell()+"\" width=\""+w.getWidth()+"\" customWidth=\"1\"/>");
				} catch (IOException e) {
					e.printStackTrace();
				}
    	    });
    	  _out.write("</cols>");
      }
	 
      //begin sheetData
      public void beginSheetData() throws IOException {
    	  _out.write("<sheetData>");
      }
      //end sheetData
      public void endSheetData() throws IOException {
    	  _out.write("</sheetData>");
      }
      
      //insert row
      /**
       * 
       * @param row  行号 
       * @param height   行高
       * @throws IOException
       */
      public void  insertRow(int row,int height,int customHeight,String spans) throws IOException {
    	  _out.write("<row r=\""+(row+1)+"\"");
    	  if(height>0) {
    		  _out.write(" ht=\""+height+"\"");
    	  }
    	  if(customHeight>0) {
    		  _out.write(" customHeight=\""+customHeight+"\"");
    	  }
    	  if(StringUtils.isNotBlank(spans)) {
    		 _out.write(" spans=\""+spans+"\""); 
    	  }
    	  _out.write(">");
      }
      //spans 1:32
      public void  insertRowOffice(int row,double height,int customHeight,String spans) throws IOException {
    	  _out.write("<row r=\""+(row+1)+"\"");
    	  if(height>0) {
    		  _out.write(" ht=\""+height+"\"");
    	  }
    	  if(customHeight>0) {
    		  _out.write(" customHeight=\""+customHeight+"\"");
    	  }
    	  if(StringUtils.isNotBlank(spans)) {
    		 _out.write(" spans=\""+spans+"\""); 
    	  }
    	  _out.write(" x14ac:dyDescent=\"0.2\">");
      }
      //end row
      public void endRow() throws IOException {
    	  _out.write("</row>");
      }
      
      //insert cell
      /**
       * 
       * @param _rownum   行号
       * @param columnIndex   列下标
       * @param t_type inlineStr不共享的字符串  s 共享的   
       * @throws IOException
       */
      public  void insertCell(int _rownum,int columnIndex,String t_type) throws IOException {
    	  String ref = new CellReference(_rownum, columnIndex).formatAsString();
    	  //System.out.println("ref:"+ref);
          _out.write("<c r=\""+ref+"\" t=\""+t_type+"\">");
    	  //_out.write("<c r=\""+ref+"\">");
      }
      //inset value
      public  void insertCellVal(String value) throws IOException {
    	  _out.write("<is><t>"+value+"</t></is>");
    	 // _out.write("<v>"+value+"</v>");
      }
      public  void insertCellValOffice(Integer value) throws IOException {
    	  _out.write("<v>"+value+"</v>");
      }
      
      //end cell
      public void endCell() throws IOException {
    	  _out.write("</c>");
      }
      
      //insert  PageMargins
      public void insertPageMargins() throws IOException {
    	  _out.write("<phoneticPr fontId=\"1\" type=\"noConversion\"/>");
    	  _out.write("<pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.511805555555556\" footer=\"0.511805555555556\"/>");
      }
      
      //insert headerFooter
      public void  insertHeaderFooter() throws IOException {
    	  _out.write("<headerFooter/>");
      }
      
      //insert Drawing
      public void insertDrawing(String rid) throws IOException {
    	  _out.write("<drawing r:id=\""+rid+"\"/>");
      }
      
      
      public void writerData(List<Integer> list,String drawingRid) throws IOException {
    	  if(CollectionUtils.isEmpty(list)) {
    		  return;
    	  }
    	  try {
    	  //行 数 默认1列   按顺序
    	  insertXmlHeader();
    	  beginWorkSheet();
    	  //beginWorkSheetOfffice();
    	  endSheetPr();
    	  insertDimension("A1:AH11");
    	  insertSheetView();
    	  insertSheetFormatPr();
    	  insertCols();//这里是控制列宽的
    	  beginSheetData();
    	  list.forEach(i->{   //这里的i测试的是有顺序的啊 不要乱写的。到时候就是for int i
    		  try {
				insertRow(i,200,1,"6:6"); //行高两百
				insertCell(i,1,"inlineStr");//这里需要修改r
			
				endCell();
				endRow();
			} catch (IOException e) {
				e.printStackTrace();
			} 
    	  });
    	  endSheetData();
    	  insertPageMargins();
    	  insertHeaderFooter();
    	  insertDrawing(drawingRid);
    	  endWorkSheet();
    	  }catch (Exception e) {
			// TODO: handle exception
    		  e.printStackTrace();
		}finally {
    		try {
    			if(_out!=null) {
					_out.flush();
					_out.close();
    			}
			} catch (IOException e) {
				logger.info("writer many relationship exception:",e);
				e.printStackTrace();
			}
    		try {
    			if(xmlRelOutStream!=null) {
    				xmlRelOutStream.flush();
    				xmlRelOutStream.close();
    			}
			} catch (IOException e) {
				logger.info("writer many relationship exception:",e);
				e.printStackTrace();
			}
    		
    	}
    	  
    	  
      }
      
      
      
    @SuppressWarnings("rawtypes")
	public void writerData(List<ExcelHeader> headList,List list,String drawingRid) throws IOException {
    	  if(CollectionUtils.isEmpty(list)) {
    		  return;
    	  }
    	  try {
    	  //行 数 默认1列   按顺序
    	  insertXmlHeader();
    	  beginWorkSheet();
    	  endSheetPr();
    	  //得到dimension  第0行  但是row =1的
    	  String dimensionStart=new CellReference(0, 0).formatAsString();
    	  String dimensionEnd=new CellReference(0, headList.size()-1).formatAsString();
    	  insertDimension(dimensionStart+":"+dimensionEnd);  //列的第一个跟最后一个ref
    	  insertSheetView();
    	  insertSheetFormatPr();
    	  //得到里面有列宽的
    	  List<ColsBean> colsList=new ArrayList<>();
    	  IntStream.range(0, headList.size()).forEach(h->{
    		  if(headList.get(h).getWidth()>0) {
	    		  ColsBean cols=new ColsBean();
	    		  cols.setCell(headList.get(h).getOrder());
	    		  cols.setWidth(headList.get(h).getWidth());
	    		  colsList.add(cols);
    		  }
    	  }); 
    	  insertCols(colsList);//这里是控制列宽的
    	  beginSheetData();
    	  
    	  insertRow(0,0,0,""); //不设置行高
    	  for(int i=0;i<headList.size();i++){
    		  try {
    			insertCell(0,i,"inlineStr");
				insertCellVal(headList.get(i).getTitle());
				endCell();
			} catch (IOException e) {
				e.printStackTrace();
			}
    	  }
    		 
		  endRow();
		  
    	  IntStream.range(0, list.size()).forEach(i->{   //这里的i测试的是有顺序的啊 不要乱写的。到时候就是for int i
    		  try {
				insertRow(i+1,200,1,"6:6"); //行高两百  从第一行开始
				 IntStream.range(0, headList.size()).forEach(h->{
					try {
						insertCell(i+1,h,"inlineStr"); //修改行 ref
						String val=BeanUtils.getProperty(list.get(i),getMethodName(headList.get(h)));
						if(val!=null&&val.indexOf("&")>-1) {
							val=val.replace("&", "&amp;");//转义
						}
						insertCellVal(val==null?"":val);
						endCell();
					} catch (IllegalAccessException e) {
						e.printStackTrace();
					} catch (InvocationTargetException e) {
						e.printStackTrace();
					} catch (NoSuchMethodException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					}
				});
				endRow();
			} catch (IOException e) {
				e.printStackTrace();
			} 
    	  });
    	  endSheetData();
    	  insertPageMargins();
    	  insertHeaderFooter();
    	  insertDrawing(drawingRid);
    	  endWorkSheet();
    	  }catch (Exception e) {
			// TODO: handle exception
    		  e.printStackTrace();
		}finally {
    		try {
    			if(_out!=null) {
					_out.flush();
					_out.close();
    			}
			} catch (IOException e) {
				logger.info("writer many relationship exception:",e);
				e.printStackTrace();
			}
    		try {
    			if(xmlRelOutStream!=null) {
    				xmlRelOutStream.flush();
    				xmlRelOutStream.close();
    			}
			} catch (IOException e) {
				logger.info("writer many relationship exception:",e);
				e.printStackTrace();
			}
    		
    	}
    	  
    	  
      }
      
      /**
       * 根据标题获取相应的方法名称
       * @param eh
       * @return
       */
      public static String getMethodName(ExcelHeader eh) {
          String mn = eh.getMethodName().substring(3);
          mn = mn.substring(0,1).toLowerCase()+mn.substring(1);
          return mn;
      }
      
      
      //cols  width  设置列宽
      public class ColsBean{
    	  private int cell;
    	  private int width;
		public int getCell() {
			return cell;
		}
		public void setCell(int cell) {
			this.cell = cell;
		}
		public int getWidth() {
			return width;
		}
		public void setWidth(int width) {
			this.width = width;
		}
    	  
      }
      
      
      
      @SuppressWarnings("rawtypes")
  	public void writerDataOffice(List<ExcelHeader> headList,List list,String drawingRid,Map<String,Integer> sMap,Integer pictureSize) throws IOException {
      	  if(CollectionUtils.isEmpty(list)) {
      		  return;
      	  }
      	  try {
      	  //行 数 默认1列   按顺序
      	  insertXmlHeader();
      	  beginWorkSheetOfffice();
      	  //得到dimension  第0行  但是row =1的
      	  String dimensionStart=new CellReference(0, 0).formatAsString();
      	  String dimensionEnd=new CellReference(0, headList.size()-1).formatAsString();
      	  insertDimension(dimensionStart+":"+dimensionEnd);  //列的第一个跟最后一个ref
      	  insertSheetView();
      	  insertSheetFormatPrOffice();
      	  //得到里面有列宽的
      	  List<ColsBean> colsList=new ArrayList<>();
      	  IntStream.range(0, headList.size()).forEach(h->{
      		  if(headList.get(h).getWidth()>0) {
  	    		  ColsBean cols=new ColsBean();
  	    		  cols.setCell(headList.get(h).getOrder());
  	    		  cols.setWidth(headList.get(h).getWidth());
  	    		  colsList.add(cols);
      		  }
      	  }); 
      	  insertCols(colsList);//这里是控制列宽的
      	  beginSheetData();
      	String spans="1:21";
      	  insertRowOffice(0,0,0,spans); //不设置行高
      
      	  for(int i=0;i<headList.size();i++){
      		  try {
      			insertCell(0,i,"s");
      			insertCellValOffice(sMap.get(headList.get(i).getTitle()));
  				endCell();
  			} catch (IOException e) {
  				e.printStackTrace();
  			}
      	  }
      		 
  		  endRow();
  		  
  		  double rowHeight=14.25;//默认值
  		  
  		  if(pictureSize!=null&&pictureSize>0) 
  			rowHeight=200.1;
  		  	spans="1:32";
  		  
  		 for(int i=0;i<list.size();i++){   //这里的i测试的是有顺序的啊 不要乱写的。到时候就是for int i
      		  try {
      			 insertRowOffice(i+1,rowHeight,1,spans); //行高两百  从第一行开始
      			for(int h=0;h<headList.size();h++){
  					try {
  						insertCell(i+1,h,"s"); //修改行 ref
  						String val=BeanUtils.getProperty(list.get(i),getMethodName(headList.get(h)));
  						if(val!=null&&val.indexOf("&")>-1) {
							val=val.replace("&", "&amp;");//转义
						}
						val=val==null?" ":val;
						if(val!=null&&val.indexOf("&")>-1) {
							val=val.replace("&", "&amp;");//转义
						}
						if(val!=null&&val.indexOf("")>-1) {
							val=val.replace("", "");//乱码字符
						}
						if(val!=null&&val.indexOf("<")>-1) {
							val=val.replace("<", "");//乱码字符
						}
						if(val!=null&&val.indexOf(">")>-1) {
							val=val.replace(">", "");//乱码字符
						}
  						val=val==null?" ":val;
  					  	Integer k=sMap.get(val);
  						insertCellValOffice(k);
  						endCell();
  					} catch (IllegalAccessException e) {
  						e.printStackTrace();
  					} catch (InvocationTargetException e) {
  						e.printStackTrace();
  					} catch (NoSuchMethodException e) {
  						e.printStackTrace();
  					} catch (IOException e) {
  						e.printStackTrace();
  					}
  				}
  				endRow();
  			} catch (IOException e) {
  				e.printStackTrace();
  			} 
      	  }
      	  endSheetData();
      	  insertPageMargins();
      	  if(pictureSize!=null&&pictureSize>0)  //有图片的试试才加入绘图
      	  insertDrawing(drawingRid);
      	  endWorkSheet();
      	  }catch (Exception e) {
  			// TODO: handle exception
      		  e.printStackTrace();
  		}finally {
      		try {
      			if(_out!=null) {
  					_out.flush();
  					_out.close();
      			}
  			} catch (IOException e) {
  				logger.info("writer many relationship exception:",e);
  				e.printStackTrace();
  			}
      		try {
      			if(xmlRelOutStream!=null) {
      				xmlRelOutStream.flush();
      				xmlRelOutStream.close();
      			}
  			} catch (IOException e) {
  				logger.info("writer many relationship exception:",e);
  				e.printStackTrace();
  			}
      		
      	}
      	  
      	  
        }

}
