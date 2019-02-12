package com.luxiaoyuan.excel.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.util.List;

import org.apache.commons.collections.CollectionUtils;
import org.slf4j.Logger;

import com.luxiaoyuan.excel.ExcelXMLContent;
import com.luxiaoyuan.excel.utils.LogUtils;


/**
 * drawing1.xml对应xml的文件名  以rels后缀结尾
 * xl\\drawings\\_rels\\drawing1.xml.rels  windows格式时
 * @author luxiaoyuan
 */

public class XmlRels {
	
	 //relationShip_type_image   引入image格式的文件类型
	 public static final String relationShip_type_image="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
	 //引入drawing类型的文件类型
	 public static final String relationShip_type_drawing="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
	 //rels文件的后缀
	 public static final String relationShip_file_suffix=".rels";
	 
	 private Logger logger=LogUtils.getUtilsLogger();
	 private Writer _out;
	 private File xmlRelFile;
	 private FileOutputStream xmlRelOutStream;
	 public XmlRels(String relPath){
		 getWriter(relPath);
     }
	 
	 public void getWriter(String relPath) {
		 xmlRelFile = new File(relPath);  
		 if(xmlRelFile.exists()) {
	    	 xmlRelFile.delete();//删除  每次都是新的写入
	    	 try {
				xmlRelFile.createNewFile();
			} catch (IOException e) {
				e.printStackTrace();
			}
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
    
    //relationships  begin
    public void beginRelationships() throws IOException {
    	_out.write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    }
    //end relationship
    public void endRelationships() throws IOException {
    	_out.write("</Relationships>");
    }
    //insert relationship  image格式下的  single 单个
    public void insertRelationship(String rid,String target,String type) throws IOException {
    	_out.write("<Relationship Id=\""+rid+"\" Type=\""+type+"\" Target=\""+target+"\"/>");
    }
  
    //writer many relationship
    public void writerMany(List<RelationshipBean> list) {
    	try {
    		if (CollectionUtils.isEmpty(list)) {
    			// a empty list
    			return;
    		}
    		insertXmlHeader(); //xml heaher
    		beginRelationships();  //begin
    		// while writer  relationship
    		list.forEach(r->{
    			try {
					_out.write("<Relationship Id=\""+r.getRid()+"\" Type=\""+r.getType()+"\" Target=\""+r.getTarget()+"\"/>");
				} catch (IOException e) {
					logger.info("writer relationship exception:",e);
					e.printStackTrace();
					return;
				}
    		});
    		endRelationships();    //end
    	}catch(IOException e) {
    		logger.info("writer many relationship exception:",e);
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
    
    
    //writer many relationship
    public void writerSingle(String rid,String target,String type) {
    	try {
    		insertXmlHeader();
			beginRelationships();
			insertRelationship( rid, target, type);
			endRelationships();    //end
		} catch (IOException e) {
			// TODO Auto-generated catch block
			logger.info("writer single relationship exception:",e);
			e.printStackTrace();
		}finally {
    		try {
    			if(_out!=null) {
					_out.flush();
					_out.close();
    			}
			} catch (IOException e) {
				logger.info("writer single relationship exception:",e);
				e.printStackTrace();
			}
    		try {
    			if(xmlRelOutStream!=null) {
    				xmlRelOutStream.flush();
    				xmlRelOutStream.close();
    			}
			} catch (IOException e) {
				logger.info("writer single relationship exception:",e);
				e.printStackTrace();
			}
    		
    	}
    
    }
    
    // relationship
    public class RelationshipBean{
    	private String rid;      //relationship id
    	private String target;   //relationship target
    	private String type;     //relationship type
		public String getRid() {
			return rid;
		}
		public void setRid(String rid) {
			this.rid = rid;
		}
		public String getTarget() {
			return target;
		}
		public void setTarget(String target) {
			this.target = target;
		}
		public String getType() {
			return type;
		}
		public void setType(String type) {
			this.type = type;
		}
    	
    }
    
    

}
