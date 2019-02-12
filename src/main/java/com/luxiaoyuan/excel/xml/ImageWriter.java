package com.luxiaoyuan.excel.xml;

import java.awt.image.BufferedImage;
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.List;
import java.util.stream.IntStream;

import javax.imageio.ImageIO;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;

import com.luxiaoyuan.excel.ExcelXMLContent;
import com.luxiaoyuan.excel.utils.LogUtils;


/**
 * 图片写入  media文件夹  
 * 返回rid 与图片名称
 * @author guanyu
 *
 */

public class ImageWriter {
    private static final String PIC_MEDIA="xl"+File.separator+"media"+File.separator;
    private static final String _SUFFIX=".jpeg";//图片后缀
    private static Logger logger = LogUtils.getUtilsLogger();
	@SuppressWarnings("unused")
	private static void writer(String path,String image1,String url) {
		//new一个文件对象用来保存图片，默认保存当前工程根目录  
    	InputStream images=null;
    	 byte[] bytes=null;
    	//创建输出流  
        FileOutputStream outStream=null;
		try {
			String imagePath=path+PIC_MEDIA+image1;
			images = new FileInputStream(url);
			try {
				bytes = IOUtils.toByteArray(images);
			} catch (IOException e1) {
				e1.printStackTrace();
			}
	        File imageFile = new File(imagePath);  
	        if(imageFile.exists()) {
	        	imageFile.delete();
	        }
	        outStream = new FileOutputStream(imageFile);
			try {
				outStream.write(bytes);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}  finally { 
        try {
        	//关闭输出流  
        	if(outStream!=null)
			 outStream.close();  
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}  
		}
       
	}
	
    private static void writerInter(String path,String image1,String url) {
        //new一个文件对象用来保存图片，默认保存当前工程根目录  
    	//创建输出流  
    	FileOutputStream outStream=null;
    	ByteArrayOutputStream byteArrayOut = null;
    	BufferedImage bufferImg=null;
        try {
         // 先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray
    	 String  imagePath=path+PIC_MEDIA+image1;
         File imageFile = new File(imagePath);  
         if(imageFile.exists()) {
           imageFile.delete();
         }
          outStream = new FileOutputStream(imageFile);
          byteArrayOut = new  ByteArrayOutputStream();
          if(url.indexOf("http")<0) {
        	  url="http:"+url;
		  }
          bufferImg = ImageIO.read(new  URL(url));
          // 将图片写入流中
          //String imageSuffix=url.substring(url.lastIndexOf("."), url.length());
          ImageIO.write(bufferImg, "png", byteArrayOut);
          byteArrayOut.writeTo(outStream);
         
        } catch (IOException e) {
        	logger.error("无效的URL->:"+url);
        	logger.error("无效的URL 异常->:"+e);
        }  finally {
	      try {
	        //关闭输出流  
	        if(outStream!=null)
	             outStream.close();  
	        } catch (IOException e) {
	            e.printStackTrace();
	        }  
	        }
	        if(byteArrayOut!=null) {
	        	try {
					byteArrayOut.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
	        }
     
   }
    
    //第二版本  查看哪个版本使用的内存更小
    private static void writerInterT(String path,String image1,String url) {
    	 try {
              //获取输入流
              BufferedInputStream in = new BufferedInputStream(new URL(url).openStream());
              //创建文件流
              String  imagePath=path+PIC_MEDIA+image1;
              File file = new File(imagePath);
              BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(file));
              //缓冲字节数组
              byte[] data = new byte[2048];
              int length = in.read(data);
              while (length != -1) {
                  out.write(data, 0, data.length);
                  length = in.read(data);
             }
             in.close();
             out.close();
         } catch (MalformedURLException e) {
             e.printStackTrace();
         } catch (IOException e) {
             e.printStackTrace();
         }
     
   }
	
	
	
	//image writer date   row cell在处理数据时赋值   这里只回写relationId
	public void writerData(List<ImageBean> list,String path) {
		if(CollectionUtils.isEmpty(list)) {
			return;
		}
		IntStream.range(0, list.size()).forEach(i->{
			try {
			ImageBean image=list.get(i);
			String imageUrl=image.getImageUrl();
			//System.out.println("imageUrl:"+imageUrl);
			//String imageSuffix=imageUrl.substring(imageUrl.lastIndexOf("."), imageUrl.length());
			//list.get(i).setImageName(ExcelXMLContent.IMAGE_PREFIX+i+imageSuffix);
			list.get(i).setImageName(ExcelXMLContent.IMAGE_PREFIX+i+_SUFFIX);
			writerInter(path,list.get(i).getImageName(),list.get(i).getImageUrl());
			list.get(i).setRelationshipId(ExcelXMLContent.RELATIONSHIP_ID_PREFIX+i);
			}catch(Exception e) {
				logger.error("无效的URL->:"+list.get(i).getImageUrl());
				logger.error("ImageBean writer异常:",e);
			}
		});
		
		
	}
	
	//imagebean
	public class ImageBean {
		private String relationshipId;//drawing imageid relationshipId 
		private String imageName; //图片名称  
		private String imageUrl; //图片名称  
		private int row;   //所在行数
		private int cell;  //所在列
		public String getRelationshipId() {
			return relationshipId;
		}
		public void setRelationshipId(String relationshipId) {
			this.relationshipId = relationshipId;
		}
		public String getImageName() {
			return imageName;
		}
		public void setImageName(String imageName) {
			this.imageName = imageName;
		}
		public int getRow() {
			return row;
		}
		public void setRow(int row) {
			this.row = row;
		}
		public int getCell() {
			return cell;
		}
		public void setCell(int cell) {
			this.cell = cell;
		}
		public String getImageUrl() {
			return imageUrl;
		}
		public void setImageUrl(String imageUrl) {
			this.imageUrl = imageUrl;
		}
		
		
		
	}
	
	
	public static void main(String[] args) {
		String base="D:\\GUANYU\\excel\\temp\\";
		File file=new File(base+"xl");
		file.mkdirs();
		File file2=new File(base+"xl\\drawings\\rels");
		file2.mkdirs();
		File file3=new File(base+"xl\\media");
		file3.mkdirs();
		File file4=new File(base+"xl\\worksheet\\rels");
		file4.mkdirs();
	}
	 
}
