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
import java.util.LinkedHashMap;
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

public class SharedStringsXml {
	  private Logger logger=LogUtils.getUtilsLogger();
	  private Writer _out;
	  private File xmlRelFile;
	  private StringBuilder sb;
	  private FileOutputStream xmlRelOutStream;
	  public SharedStringsXml(String filePath){
          getWriter(filePath);
      }
	  
	  public void getWriter(String relPath) {
		  	 sb=new StringBuilder();
			 xmlRelFile = new File(relPath);  
		        //创建输出流  
			 if(xmlRelFile.exists()) {
		    	 xmlRelFile.delete();//删除  每次都是新的写入
		    	 try {
					xmlRelFile.createNewFile();
				} catch (IOException e) {
					e.printStackTrace();
				}
		     }
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
      
      //begin sst  6432   1123
      public void beginSst(int count,int uniqueCount) throws IOException {
    	  _out.write("<sst count=\""+count+"\" uniqueCount=\""+uniqueCount+"\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
      }
      
      /**
       * insertSi  插入si标签(拼接后的总的si标签)
       * @param v   值
       * @throws IOException
       */
      public void insertSi(String sb) throws IOException{
    	  _out.write(sb);
      }
      
      /**
       * insertSi  插入si标签
       * @param v   值
       * @throws IOException
       */
      public void addSi(String v) throws IOException {
    	  sb.append("<si>" + 
    	  			"<t>"+v+"</t>" + 
    	  			"</si>");
      }
      
      public void endSst() throws IOException {
    	  _out.write("</sst>");
      }
      
      
    @SuppressWarnings("rawtypes")
	public void writerDataOffice(List<ExcelHeader> headList,List list,Map<String,Integer> sMap){
    	  if(CollectionUtils.isEmpty(list)) {
    		  return;
    	  }
    	  if(CollectionUtils.isEmpty(headList)) {
    		  return;
    	  }
    	 
    	  try {
    		  insertXmlHeader();
	    	  int count=(list.size()+1)*headList.size();//总数量
	    	  int uniqueCount=headList.size();//重复的统计  标签不会重复
	    	  for(int i=0;i<headList.size();i++){  //标签不会重复
    		        String title=headList.get(i).getTitle();
					try {
						addSi(title);
						sMap.put(title,i);
					} catch (IOException e) {
						e.printStackTrace();
					}
	          }
	    	 for(int i=0;i<list.size();i++){   //这里的i测试的是有顺序的啊 不要乱写的。到时候就是for int i
					 for(int h=0;h<headList.size();h++){
						try {
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
							if(sMap.get(val)==null) {
								try {
									addSi(val);
									sMap.put(val,uniqueCount);//随机id
									uniqueCount=uniqueCount+1;
								} catch (IOException e) {
									e.printStackTrace();
								}
							}  //只添加一次
						} catch (IllegalAccessException e) {
							e.printStackTrace();
						} catch (InvocationTargetException e) {
							e.printStackTrace();
						} catch (NoSuchMethodException e) {
							e.printStackTrace();
						}
					}
	    	  }
	    	  beginSst(count, uniqueCount); //count=row(list.size)*cell(headList.szie)  //start
	    	  insertSi(sb.toString());
	    	  endSst();//end
    	  }catch (IOException e) {
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
    	  private double width;
		public int getCell() {
			return cell;
		}
		public void setCell(int cell) {
			this.cell = cell;
		}
		public double getWidth() {
			return width;
		}
		public void setWidth(double width) {
			this.width = width;
		}
    	  
      }

}
