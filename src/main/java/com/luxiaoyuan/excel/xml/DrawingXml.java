package com.luxiaoyuan.excel.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.util.List;
import java.util.stream.IntStream;

import org.apache.commons.collections.CollectionUtils;
import org.slf4j.Logger;

import com.luxiaoyuan.excel.ExcelXMLContent;
import com.luxiaoyuan.excel.utils.LogUtils;
import com.luxiaoyuan.excel.xml.ImageWriter.ImageBean;


/**
 * 写入drawing.xml 单个或者多个时根据自定义
 * default xl\\drawings\\drawing1.xml  windows格式
 * @author luxiaoyuan
 *
 */

public class DrawingXml {
	  private Logger logger=LogUtils.getUtilsLogger();
	  private Writer _out;
	  private File xmlRelFile;
	  private FileOutputStream xmlRelOutStream;
	  public DrawingXml(String filePath){
          getWriter(filePath);
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
	
	 /**
     * @throws IOException
     */
    //drawing1.xml   
    //begin --<xdr:wsD>  标签
    public void beginXdrWsDr() throws IOException {
    	_out.write("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
    }
    //begin --<xdr:wsD>  标签
    public void endXdrWsDr() throws IOException {
    	_out.write("</xdr:wsDr>");
    }
    //<xdr:wsD> children <xdr:twoCellAnchor editAs="oneCell">
    //begin xdr:twoCellAnchor
    public void beginXdrTwoCellAnchor() throws IOException {
    	_out.write("<xdr:twoCellAnchor editAs=\"oneCell\">");
    }
    //end xdr:twoCellAnchor
    public void endXdrTwoCellAnchor() throws IOException {
    	_out.write("</xdr:twoCellAnchor>");
    }
    
    //begin <xdr:from>
    public void beginXdrFrom() throws IOException {
    	_out.write("<xdr:from>");
    }
    //end </xdr:from>
    public void endXdrFrom() throws IOException {
    	_out.write("</xdr:from>");
    }
    
    //begin <xdr:to>
    public void beginXdrTo() throws IOException {
    	_out.write("<xdr:to>");
    }
    //end </xdr:to>
    public void endXdrTo() throws IOException {
    	_out.write("</xdr:to>");
    }
    
    //insert xdr:from or xdr:to  content
    /**
     * @param col  （列）
     * @param colOff  （列偏移量）  *9525单位
     * @param row  （行）
     * @param rowOff （行偏移量）
     * @throws IOException
     */
    public void insertFromOrToContent(int col
    		,int colOff
    		,int row
    		,int rowOff
    		) throws IOException {
    	_out.write("<xdr:col>"+col+"</xdr:col>");
    	_out.write("<xdr:colOff>"+colOff+"</xdr:colOff>");
    	_out.write("<xdr:row>"+row+"</xdr:row>");
    	_out.write("<xdr:rowOff>"+rowOff+"</xdr:rowOff>");
    }
    
    //begin <xdr:pic>
    public void beginXdrPic() throws IOException {
    	_out.write("<xdr:pic>");
    }
    //end </xdr:pic>
    public void endXdrPic() throws IOException {
    	_out.write("</xdr:pic>");
    }
    
    //end <xdr:clientData/> 在 pic的标签后面出现
    public void endXdrClientData() throws IOException {
    	_out.write("<xdr:clientData/>");
    }
    
    //insert <xdr:nvPicPr>
    public void insertXdrNvPicPr(int cNvPrId,String cNvPrName,String cNvPrDescr) throws IOException {
    	_out.write("<xdr:nvPicPr>");
    		_out.write("<xdr:cNvPr id=\""+cNvPrId+"\" name=\""+cNvPrName+"\" descr=\""+cNvPrDescr+"\"/>");
        	_out.write("<xdr:cNvPicPr>");
        		_out.write("<a:picLocks noChangeAspect=\"1\"/>");  //暂时不知道具体的作用。
        	_out.write("</xdr:cNvPicPr>");
        _out.write("</xdr:nvPicPr>");
    }
    
   
    /**
     *  //insert <xdr:blipFill>
     * @param rEnbed  relationshipId  
     * @throws IOException
     */
    public void insertXdrBlipFill(String rEnbed) throws IOException {
    	_out.write("<xdr:blipFill>");
    		_out.write("<a:blip r:embed=\""+rEnbed+"\"/>");
    		_out.write("<a:stretch>");
    			_out.write("<a:fillRect/>");
    		_out.write("</a:stretch>");
    	_out.write("</xdr:blipFill>");
    }

    //insert <xdr:spPr>
    /**           default
     * @param x   6666865
     * @param y   1304925
     * @param cx  2962910
     * @param cy  860425
     * @throws IOException
     * https://docs.microsoft.com/zh-cn/dotnet/api/documentformat.openxml.drawing.presetgeometry?redirectedfrom=MSDN&view=openxml-2.8.1
                 * 预设的几何图形。 当对象作为 xml 序列出时，其限定的名称是: prstGeom。
     * avLst （形状调整值列表）
     */
    public void insertXdrSpPr(int x,int y,int cx,int cy) throws IOException {
    	_out.write("<xdr:spPr>");
    		//_out.write("<a:xfrm flipH=\"1\">");
    		_out.write("<a:xfrm>");
    			_out.write("<a:off x=\""+x+"\" y=\""+y+"\"/>");
    			_out.write("<a:ext cx=\""+cx+"\" cy=\""+cy+"\"/>");
    		_out.write("</a:xfrm>");
    		_out.write("<a:prstGeom prst=\"rect\">");
    			_out.write("<a:avLst/>");
    		_out.write("</a:prstGeom>");
       _out.write("</xdr:spPr>");
    }
    
    
    public void writerData(List<ImageBean> imageList) throws IOException {
    	if(CollectionUtils.isEmpty(imageList)) {
    		return;
    	}
    	try {
    	insertXmlHeader();//xml 头部
        beginXdrWsDr();  //wsdr 
    	IntStream.range(0, imageList.size()).forEach(x->{
				try {
					insertData(imageList.get(x),x);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
    	});
        endXdrWsDr();//end xsdr
    	}catch(IOException e) {
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
    
    
    public void insertData(ImageBean img,int i) throws IOException {
    	//可循环多个 start Anchor
        beginXdrTwoCellAnchor();
        //begin from
        beginXdrFrom();
        //insert content   一行一列   行列的偏移量先使用默认值
        insertFromOrToContent(img.getCell(), ExcelXMLContent.Xdr_from_colOff, img.getRow(), ExcelXMLContent.Xdr_from_rowOff);
        endXdrFrom();
        //endfrom
        
        //start to
        beginXdrTo();
        //insert content   一行一列   行列的偏移量先使用默认值
        insertFromOrToContent(img.getCell(), ExcelXMLContent.Xdr_to_colOff, img.getRow(), ExcelXMLContent.Xdr_to_rowOff);
        endXdrTo();
        //end to
        
        //start pic
        beginXdrPic();
        
        //insert nvPicPr
        insertXdrNvPicPr(i, "图片 "+i, "timg"+i);
        //insert blipFill
        insertXdrBlipFill(img.getRelationshipId()); //relationshipId
        //insert spPr
        insertXdrSpPr(ExcelXMLContent.Xdr_SpPr_x, ExcelXMLContent.Xdr_SpPr_y, ExcelXMLContent.Xdr_SpPr_cx, ExcelXMLContent.Xdr_SpPr_cy);
        
        endXdrPic();
        //end pic
        
        //只有end  暂时不知道begin的作用
        endXdrClientData();
        // end client data
        
        endXdrTwoCellAnchor();
        //end anchor
        
    }

}
