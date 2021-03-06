package com.luxiaoyuan.excel;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;

import com.luxiaoyuan.excel.utils.LogUtils;

 
/**
 * excel处理工具
 * 概念-> 表头数据:报表的表头
 *        行内数据:表头以下的数据
 * 功能:动态生成单级，多级Excel表头
 * 备注：tree型结构数据的root节点的id必须为零（0）
 * @param <T>
 */
public class ExcelTool<T> {
	public static Logger log = LogUtils.getUtilsLogger();
    private HSSFWorkbook workbook;//excel 对象
    private String title; //表格标题
    private int colWidth = 20; //单元格宽度
    private int rowHeight = 20;//单元格行高度
    private HSSFCellStyle styleHead; //表头样式
    private HSSFCellStyle styleBody; //主体样式
    private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); //日期格式化,默认yyyy-MM-dd HH:mm:ss
 
    /**
     * 无参数 初始化 对象
     */
    public ExcelTool(){
       this.title="sheet1";
        this.workbook = new HSSFWorkbook();
        init(0);
    }
 
    /**
     * 有参数 初始化 对象
     * @param title
     * @param colWidth
     * @param rowHeight
     * @param dateFormat
     */
    public ExcelTool(String title,int colWidth,int rowHeight,String dateFormat){
        this.colWidth = colWidth;
        this.rowHeight = rowHeight;
        this.title = title;
        this.workbook = new HSSFWorkbook();
        sdf = new SimpleDateFormat(dateFormat);
        init(0);
    }
 
    public ExcelTool(String title,int colWidth,int rowHeight){
        this.colWidth = colWidth;
        this.rowHeight = rowHeight;
        this.title = title;
        this.workbook = new HSSFWorkbook();
        init(0);
    }
    public ExcelTool(String title,int colWidth,int rowHeight,int flag){
        this.colWidth = colWidth;
        this.rowHeight = rowHeight;
        this.title = title;
        this.workbook = new HSSFWorkbook();
        init(flag);
    }
    public ExcelTool(String title){
        this.title = title;
        workbook = new HSSFWorkbook();
        init(0);
    }
    /**ExcelTool 属性 get、set 方法 开始*/
    public int getColWidth() {
        return colWidth;
    }
 
    public void setColWidth(int colWidth) {
        this.colWidth = colWidth;
    }
 
    public int getRowHeight() {
        return rowHeight;
    }
 
    public void setRowHeight(int rowHeight) {
        this.rowHeight = rowHeight;
    }
 
    public HSSFWorkbook getWorkbook() {
        return workbook;
    }
 
    public void setWorkbook(HSSFWorkbook workbook) {
        this.workbook = workbook;
    }
 
    public String getTitle() {
        return title;
    }
 
    public void setTitle(String title) {
        this.title = title;
    }
 
 
    public HSSFCellStyle getStyleHead() {
        return styleHead;
    }
 
 
    public void setStyleHead(HSSFCellStyle styleHead) {
        this.styleHead = styleHead;
    }
 
 
    public HSSFCellStyle getStyleBody() {
        return styleBody;
    }
 
 
    public void setStyleBody(HSSFCellStyle styleBody) {
        this.styleBody = styleBody;
    }
    /**ExcelTool 属性 get、set 方法 结束*/
    //内部统一调用的样式初始化
    private void init(int styleFlag){
        styleHead = workbook.createCellStyle();
        styleHead.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 左右居中
        styleHead.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 上下居中
        styleHead.setRightBorderColor(HSSFColor.BLACK.index);
        styleHead.setBottomBorderColor(HSSFColor.BLACK.index);
        switch(styleFlag){
            case 1:
                styleBody = workbook.createCellStyle();
                styleBody.setAlignment(HSSFCellStyle.ALIGN_LEFT);// 左右居中ALIGN_CENTER
                styleBody.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 上下居中
                styleBody.setRightBorderColor(HSSFColor.BLACK.index);
                styleBody.setBottomBorderColor(HSSFColor.BLACK.index);
                styleBody.setBorderRight((short) 1);// 边框的大小
                styleBody.setBorderBottom((short) 1);// 边框的大小
                break;
            default:
                styleBody = workbook.createCellStyle();
                styleBody.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 左右居中ALIGN_CENTER
                styleBody.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 上下居中
                styleBody.setRightBorderColor(HSSFColor.BLACK.index);
                styleBody.setBottomBorderColor(HSSFColor.BLACK.index);
                styleBody.setBorderRight((short) 1);// 边框的大小
                styleBody.setBorderBottom((short) 1);// 边框的大小
                break;
        }
    }
    /**
     * 导出表格 无返回
     * @param listTpamscolumn 表头数据
     * @param datas 行内数据
     * @param FilePath 保存路径
     * @param flag
     * @param rowFlag
     * @throws Exception
     */
    public void exportExcel(List<Column> listTpamscolumn, List<T> datas, String FilePath, boolean flag, boolean rowFlag) throws Exception{
        splitDataToSheets(datas, listTpamscolumn,flag,rowFlag);
        save(workbook,FilePath);
    }
    /**
     * 导出表格 有返回值
     * @param listTpamscolumn 表头数据
     * @param datas 行内数据
     * @param flag  只输出表头数据
     * @param rowFlag
     * @return
     * @throws Exception
     */
    public InputStream exportExcel(List<Column> listTpamscolumn, List<T> datas, boolean flag, boolean rowFlag) throws Exception {
        splitDataToSheets(datas, listTpamscolumn,flag,rowFlag);
        return save(workbook);
    }
    
    
    /**
     * 导出表格 有返回值  多个sheet  保存到本地版本  colum相同的。不同的请自定义
     * @param listTpamscolumn 表头数据
     * @param datas 行内数据
     * @param flag  只输出表头数据
     * @param rowFlag
     * @return
     * @throws Exception
     */
    public void  exportExcelManySheet(List<Column> listTpamscolumn, Map<String,Object> dataMap,String FilePath, boolean flag, boolean rowFlag) throws Exception {
        if(dataMap.size()>0) 
        {
        	dataMap.forEach((k,v)->{
        		@SuppressWarnings("unchecked")
				List<T> datas=(List<T>) v;
    	    	try {
    	    		this.title=k;  //赋值当前title
					splitDataToSheets(datas, listTpamscolumn,flag,rowFlag);
				} catch (Exception e) {
					log.error("写入表格异常:",e);
				}
        	});
	        save(workbook,FilePath);
        }
    }
    
    /**
     * 浏览器下载版本
     * 导出表格 有返回值  多个sheet    colum相同的。不同的请自定义
     * @param listTpamscolumn 表头数据
     * @param datas 行内数据
     * @param flag  只输出表头数据
     * @param rowFlag
     * @return
     * @throws Exception
     */
    public void  exportExcelManySheetWeb(List<Column> listTpamscolumn,
    		Map<String,Object> dataMap,
    		boolean flag,
    		boolean rowFlag,HttpServletResponse response) throws Exception {
       
        	// 告诉浏览器用什么软件可以打开此文件
        	response.setHeader("content-Type", "application/vnd.ms-excel");
        	// 下载文件的默认名称
        	response.setHeader("Content-Disposition", "attachment;filename="+URLEncoder.encode(title, "utf-8"));
        	InputStream ex=null;
        	try {
	        	dataMap.forEach((k,v)->{
	        			@SuppressWarnings("unchecked")
						List<T> datas=(List<T>) v;
	    	    		this.title=k;  //赋值当前title
						try {
							splitDataToSheets(datas, listTpamscolumn,flag,rowFlag);
						} catch (Exception e) {
							log.error("写入表格异常:",e);
						}
	        	});
		        ex=save(workbook);
		        workbook.write(response.getOutputStream());
        	} catch (Exception e) {
        		log.error("写入表格异常:",e);
			}
        	finally {
        		try {
	    			response.getOutputStream().flush();
	    			response.getOutputStream().close();
	    			return;
	    		} catch (Exception e) {
	    			log.error("流关闭异常",e);
	    		}
        	}
	        
    }
    
    
    /**
     * 导出Excel,适用于web导出excel
     * @param sheet excel
     * @param data 行内数据
     * @param listTpamscolumn 表头数据
     * @param flag 只输出表头数据
     * @param rowFlag 输出展示数据的结构(表头下面行的数据)
     * @throws Exception
     */
    private void writeSheet(HSSFSheet sheet, List<T> data, List<Column> listTpamscolumn, boolean flag, boolean rowFlag) throws Exception {
        sheet.setDefaultColumnWidth(colWidth);
        sheet.setDefaultRowHeightInPoints(rowHeight);
        sheet= createHead(sheet, listTpamscolumn.get(0).getTotalRow(), listTpamscolumn.get(0).getTotalCol());
        createHead(listTpamscolumn,sheet,0);
        if(flag)//控制是否
            writeSheetContent(listTpamscolumn,data,sheet, listTpamscolumn.get(0).getTotalRow(),rowFlag);
    }
    /**
     *  拆分sheet，因为每个sheet不能超过6526，否则会报异常
     * @param data 行内数据
     * @param listTpamscolumn 表头数据
     * @param flag 只输出表头数据
     * @param rowFlag  输出展示数据的结构(表头下面行的数据)
     * @throws Exception
     */
    private void splitDataToSheets(List<T> data, List<Column> listTpamscolumn, boolean flag, boolean rowFlag)throws Exception{
        int dataCount = data.size();
        int maxColumn = 65535;
        int pieces = dataCount/maxColumn;
        for(int i = 1; i <= pieces;i++){
            HSSFSheet sheet = workbook.createSheet(this.title+i);
            List<T> subList = data.subList((i-1)*maxColumn, i*maxColumn);
            writeSheet(sheet,subList, listTpamscolumn,flag,rowFlag);
        }
        HSSFSheet sheet = workbook.createSheet(this.title+(pieces+1));
        writeSheet(sheet, data.subList(pieces*maxColumn, dataCount), listTpamscolumn,flag,rowFlag);
    }
 
    /**
     * 把数据写入到单元格
     * @param listTpamscolumn 表头数据
     * @param datas 行内数据
     * @param sheet 工作表（excel分页）
     * @throws Exception
     * void
     */
    private void writeSheetContent(List<Column> listTpamscolumn, List<T> datas, HSSFSheet sheet, int rowIndex, boolean rowFlag) throws Exception{
        HSSFRow row = null;
        List<Column>  listCol=new ArrayList<>();
        rowFlag=false;
        if(rowFlag){//暂时没有用 后面扩展用
            for (int i = 0, index = rowIndex; i < datas.size(); i++, index++) {
                row = sheet.createRow(index);//创建行
                for (int j = 0; j < listTpamscolumn.size(); j++) {
                    createColl(row,j,listTpamscolumn.get(j).getFieldName() ,datas.get(i));
                }
            }
        }
        else {
            getColumnList(listTpamscolumn, listCol);
            for (int i = 0, index = rowIndex; i < datas.size(); i++, index++) {
                row = sheet.createRow(index);//创建行
                for (int j = 0; j < listCol.size(); j++) {
                    Column c = listCol.get(j);
                    createCol(row, c, datas.get(i));
                }
 
            }
        }
    }
 
    /**
     * 根据list 来创建单元格 暂时没有用
     * @param row
     * @param j
     * @param finame
     * @param t
     */
    private void createColl(HSSFRow row, int j, String finame, T t) {
        HSSFCell cell = row.createCell(j);  //创建单元格
        cell.setCellStyle(styleBody); //设置单元格样式
        String text="";
        if(t instanceof List){
            List<Map> temp= (List<Map>)t;
            if(j>=temp.size()) return;
            text=String.valueOf(  temp.get(j).get(finame)==null?"": temp.get(j).get(finame));
        }
        HSSFRichTextString richString = new HSSFRichTextString(text);
        cell.setCellValue(richString);
    }
 
    /**
     * 把column的columnList整理成一个list<column> 过滤表头的脏数据
     * @param list 表头数据
     * @param listCol 返回新的list
     * @return
     * List<column>
     */
    private void getColumnList(List<Column> list, List<Column> listCol){
        for(int i = 0; i < list.size(); i++){
            if(list.get(i).getFieldName() != null){
                listCol.add(list.get(i));
            }
            List<Column> listChilren = list.get(i).getListTpamscolumn();
            if(listChilren.size() > 0){
                getColumnList( listChilren, listCol);
            }
        }
    }
 
    /**
     * 保存Excel到InputStream，此方法适合web导出excel
     * @param workbook
     * @return
     */
    private InputStream save(HSSFWorkbook workbook) {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try {
            workbook.write(bos);
            InputStream bis = new ByteArrayInputStream(bos.toByteArray());
            return bis;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }
 
    /**
     * 保存excel到本机指定的路径
     * @param workbook
     * @param filePath
     * @throws IOException
     */
    private void save(HSSFWorkbook workbook, String filePath){
        File file = new File(filePath);
        if (!file.getParentFile().exists()) {
            file.getParentFile().mkdirs();
        }
        FileOutputStream fOut = null;
        try {
            fOut = new FileOutputStream(file);
            workbook.write(fOut);
            fOut.flush();
        } catch (Exception e) {
            e.printStackTrace();
        }
        try {if(null!=fOut) fOut.close();} catch (Exception e1) { }
    }
 
    /**
     * 创建行
     * @param row Excel对应的行
     * @param tpamscolumn 当前单元格属性
     * @param v
     * @param j
     * @return
     * @throws Exception
     */
    public int createRowVal(HSSFRow row, Column tpamscolumn, T v, int j) throws Exception{
        //遍历标题
        if(tpamscolumn.getListTpamscolumn() != null && tpamscolumn.getListTpamscolumn().size() > 0){
            for(int i = 0; i < tpamscolumn.getListTpamscolumn().size(); i++){
                createRowVal(row, tpamscolumn.getListTpamscolumn().get(i),v,j);
            }
        }else{
            createCol(row, tpamscolumn,v);
        }
        return j;
 
    }
 
    /**
     * 创建单元格
     * @param row  Excel对应的行
     * @param tpamscolumn 当前单元格对象
     * @param v
     * @throws Exception
     */
    public void createCol(HSSFRow row, Column tpamscolumn, T v) throws Exception{
        HSSFCell cell = row.createCell( tpamscolumn.getCol());  //创建单元格
        cell.setCellStyle(styleBody); //设置单元格样式
        final Object[] value = {null};
        if(v instanceof Map){
            Map m=(Map)v;
            m.forEach((k, val) -> {
                if(k.equals(tpamscolumn.getFieldName()) && !tpamscolumn.isHasChilren()){
                    value[0] =val;
                }
            });
        }else {
            Class<?> cls = v.getClass();// 拿到该类
            Field[] fields=cls.getDeclaredFields();// 获取实体类的所有属性，返回Field数组
            for(int i=0;i<fields.length;i++){
                Field f = fields[i];
                f.setAccessible(true); // 设置些属性是可以访问的
                if(tpamscolumn.getFieldName().equals(f.getName()) && !tpamscolumn.isHasChilren() )// && !tpamscolumn.isHasChilren()
                    value[0] =  f.get(v);
                if(value[0] instanceof Date)
                    value[0] = parseDate((Date) value[0]);
            }
        }
        if(value[0]!=null) {
            HSSFRichTextString richString = new HSSFRichTextString(value[0].toString());
            cell.setCellValue(richString);
        }
 
    }
 
    /**
     * 时间转换
     * @param date
     * @return
     * String
     */
    private String parseDate(Date date){
        String dateStr = "";
        try{
            dateStr = sdf.format(date);
        } catch (Exception e){
            e.printStackTrace();
        }
        return dateStr;
    }
 
 
    /**
     * 根据数据的行数和列数，在excel创建单元格cell
     * @param sheetCo excel分页
     * @param r excel 行数
     * @param c excel 列数
     * @return
     */
    public HSSFSheet createHead(HSSFSheet sheetCo, int r, int c){
        for(int i = 0; i <r; i++){
            HSSFRow row = sheetCo.createRow(i);
            for(int j = 0; j <c; j++){
                HSSFCell cell=row.createCell(j);
            }
        }
        return sheetCo;
    }
    /**
     * 使用递归 在excel写入表头数据 支持单级，多级表头的创建
     * @param listTpamscolumn 表头数据
     * @param sheetCo 哪个分页
     * @param rowIndex 当前Excel的第几行
     */
    public void createHead(List<Column> listTpamscolumn, HSSFSheet sheetCo, int rowIndex){
        HSSFRow row = sheetCo.getRow(rowIndex);
        int len = listTpamscolumn.size();//当前行 有多少列
        for(int i = 0; i <len; i++){//i是headers的索引，n是Excel的索引 多级表头
            Column tpamscolumn = listTpamscolumn.get(i);
            //创建这一行的第几列单元格
            int r= tpamscolumn.getRow();
            int rLen= tpamscolumn.getrLen();
            int c= tpamscolumn.getCol();
            int cLen= tpamscolumn.getcLen();
            int endR=r+rLen;
            int endC=c+cLen;
            if(endC>c){
                endC--;
            }
            HSSFCell cell=row.getCell(c);
            cell.setCellStyle(styleHead); //设置表头样式
            HSSFRichTextString text = new HSSFRichTextString(tpamscolumn.getContent());
            cell.setCellValue(text);
            // 合并单元格
            CellRangeAddress cra =new CellRangeAddress(r,endR,c,endC);
            sheetCo.addMergedRegion(cra);
            if(tpamscolumn.isHasChilren()){
                rowIndex=r+1;
                createHead( tpamscolumn.getListTpamscolumn(), sheetCo,rowIndex);
            }
        }
    }
 
 
    /**
     * 转换成column对象
     *  支持List<T>的数据结构:map String ，只能是单级的数据
     * @param list 需要转换的数据
     * @return
     */
    public  List<Column> columnTransformer(List<T>  list){
        List<Column> lc=new ArrayList<>();
        if(list.get(0)instanceof Map) {
               for (Map<String, String> m :(List<Map<String, String>>)list ) {
                   m.forEach((k, val) -> {
                       Column tpamscolumn = new Column();
                       int i = 1;
                       tpamscolumn.setId(String.valueOf(i));
                       tpamscolumn.setPid("0");
                       tpamscolumn.setContent(k);
                       tpamscolumn.setFieldName(val);
                       lc.add(tpamscolumn);
                       i++;
                   });
               }
           } else {
               for(String s:(List<String>)list) {
                   Column tpamscolumn = new Column();
                   int i = 1;
                   tpamscolumn.setId(String.valueOf(i));
                   tpamscolumn.setPid("0");
                   tpamscolumn.setContent(s);
                   tpamscolumn.setFieldName(null);
                   lc.add(tpamscolumn);
                   i++;
               }
           }
        setParm(lc);//处理一下
        List<Column> s= TreeTool.buildByRecursive(lc);
        setColNum(lc,s,0);
        return s;
    }
 
    /**
     * 转换成column对象 返回tree数据结构
     * 支持：List<map>、某个具体对象（entity）数据的转换
     * @param list 需要转换的数据
     * @param id 当前节点id 字段的名称  主键
     * @param pid 父节点id 字段的名称
     * @param content 填写表头单元格内容的 字段名称
     * @param fielName 填写行内数据对的 字段名称
     * @return
     * @throws Exception
     */
    public List<Column> columnTransformer(List<T> list , String id, String pid, String content, String fielName) throws Exception {
        List<Column> lc=new ArrayList<>();
        if(list.get(0) instanceof Map){
            for(Map m:(List<Map>)list) {
                Column tpamscolumn = new Column();
                m.forEach((k, val) -> {//java8 以上的遍历方式
                    if (id.equals(k))
                        tpamscolumn.setId(String.valueOf(val));
                    if (pid.equals(k)) tpamscolumn.setPid((String) val);
                    if (content.equals(k)) tpamscolumn.setContent((String) val);
                    if (fielName.equals(k) && fielName != null) tpamscolumn.setFieldName((String) val);
                });
                lc.add(tpamscolumn);
            }
        }else {
            for (T t : list) {//反射
                Column tpamscolumn = new Column();
                Class cls = t.getClass();
                Field[] fs=cls.getDeclaredFields();
                for (int i = 0; i < fs.length; i++) {
                    Field f = fs[i];
                    f.setAccessible(true); // 设置些属性是可以访问的
                    if (id.equals(f.getName()) && f.get(t) != null)
                        tpamscolumn.setId(f.get(t).toString());
                    if (pid.equals(f.getName()) && f.get(t) != null)
                        tpamscolumn.setPid(f.get(t).toString());
                    if (pid.equals(f.getName()) && ( f.get(t) == null || "".equals(f.get(t))))
                        tpamscolumn.setPid("0");
                    if (content.equals(f.getName()) && f.get(t) != null)
                        tpamscolumn.setContent(f.get(t).toString());
                    if ( f.get(t) != null  && fielName!=null && fielName.equals(f.getName()))
                        tpamscolumn.setFieldName(f.get(t).toString());
                }
                lc.add(tpamscolumn);
            }
        }
        setParm(lc);//处理一下
        List<Column> s= TreeTool.buildByRecursive(lc);
        setColNum(lc,s,0);
        return s;
    }
    /**
     * 设置基础的参数
     * @param list
     */
    public static void setParm(List<Column> list){
        int row=0;//excel第几行
        int rLen=0; //excel 跨多少行
        int tree_step=0;
        for(int i=0;i<list.size();i++){
            Column poit= list.get(i);
            tree_step=TreeTool.getTreeStep(list,poit.getPid(),0);
            poit.setTree_step(tree_step);
            poit.setRow(tree_step);//设置第几行
            //判断是否有节点
            boolean hasCh=TreeTool.hasChild( list,poit);
            poit.setHasChilren(hasCh);
            if( hasCh){
                poit.setrLen(0);//设置跨多少行
            }else{
                int maxIdex=TreeTool.getMaxStep(list,"0")-1;
                if(tree_step<maxIdex){
                    rLen=maxIdex-tree_step;
                    poit.setrLen(rLen);
                }
            }
            if("0".equals(poit.getPid())){
                poit.setTotalRow(TreeTool.getMaxStep(list,"0"));
                poit.setTotalCol(TreeTool.getDownChilren(list,"0"));
            }
 
        }
    }
    /**
     * 设置基础的参数
     * @param list
     * @param treeList
     * @param pcIndex
     */
    public static void setColNum(List<Column> list, List<Column> treeList, int pcIndex){
        int col=pcIndex;//excel第几列
        int cLen;//excel跨多少列
        for(int i=0;i<treeList.size();i++){
            Column poit= treeList.get(i);
            if(treeList.size()<=1){
                poit.setCol(col);
            }else{
                for(int j=0;j<i;j++){
                    Column poitJ= treeList.get(j);
                    cLen=TreeTool.getDownChilren(list,poitJ.getId());
                    if(cLen==0||cLen==1){
                        col++;
                    }else {
                        col+=cLen;
                    }
                }
                poit.setCol(col);
                col=pcIndex;
            }
            cLen=TreeTool.getDownChilren(list,poit.getId());
            poit.setcLen(cLen);//设置跨多少列
            if(poit.getListTpamscolumn().size()>0){
                pcIndex=poit.getCol();
                setColNum( list, poit.getListTpamscolumn(), pcIndex);
            }
        }
    }
 
//========上部分是导出excel的使用（生成excel），下部分是解析excel，由于excel导入==================================================================================================================================
 
    /**
     * 根据HSSFCell类型设置数据
     * @param cell 单元格
     * @return
     */
    public static String getCellFormatValue(Cell cell) {
        String cellvalue = "";
        if (cell != null) {
            switch (cell.getCellType()) { // 判断当前Cell的Type
 
                case HSSFCell.CELL_TYPE_NUMERIC:  // 如果当前Cell的Type为NUMERIC
                case HSSFCell.CELL_TYPE_FORMULA: {
                    // 判断当前的cell是否为Date
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = sdf.format(date);
                    }
                    else { // 如果是纯数字
                        cellvalue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case HSSFCell.CELL_TYPE_STRING:  // 如果当前Cell的Type为STRIN
                    // 取得当前的Cell字符串
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                default:  // 默认的Cell值
                    cellvalue = "";
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;
    }
 
    /**
     * 描述：根据文件后缀，自适应上传文件的版本
     * @param inStr,fileName
     * @return
     * @throws Exception
     */
    public static Workbook getWorkbookType(InputStream inStr, String fileName) throws Exception{
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if(".xls".equals(fileType)){
            wb = new HSSFWorkbook(inStr);  //2003-
        }else if(".xlsx".equals(fileType)){
            wb = new XSSFWorkbook(inStr);  //2007+
        }else{
            throw new Exception("导入格式错误");
        }
        return wb;
    }
 
    /**
     * 获取单元格数据内容为字符串类型的数据
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
    public static String getStringCellValue(Cell cell) {
        String strCell = "";
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                strCell = cell.getStringCellValue().trim();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                strCell = String.valueOf(cell.getNumericCellValue()).trim();
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue()).trim();
                break;
            case Cell.CELL_TYPE_BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell.equals("") || strCell == null) {
            return "";
        }
        return strCell;
    }
    /**
     * 判断指定的单元格是否是合并单元格
     * @param sheet
     * @param row    行下标
     * @param column 列下标
     * @return
     */
    public  boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }
 
    /**
     * 获取合并单元格的值
     * @param sheet
     * @param row 行下标
     * @param column 列下标
     * @return
     */
    public  String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
 
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getStringCellValue(fCell);
                }
            }
        }
        return "";
    }
    /**
     * 获取excel的值 返回的 List<List<String>>的数据结构
     * @param  fileUrl 文件路径
     * @param  sheetNum 工作表（第几分页[1,2,3.....]）
     * @return  List<List<String>>
     */
    public  List<List<String>> getExcelValues(String fileUrl, int sheetNum) throws Exception{
        List<List<String>> values = new ArrayList<List<String>>();
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(is);
        int sheetCount = sheetNum-1; //workbook.getNumberOfSheets();//sheet 数量,可以只读取手动指定的sheet页
        //int sheetCount1= workbook.getNumberOfSheets();
        Sheet sheet = workbook.getSheetAt(sheetCount); //读取第几个工作表sheet
        int rowNum = sheet.getLastRowNum();//有多少行
        for (int i = 1; i <= rowNum; i++) {
            Row row = sheet.getRow(i);//第i行
            if (row == null) {//过滤空行
                continue;
            }
            List<String> list = new ArrayList<>();
            int colCount = sheet.getRow(0).getLastCellNum();//用表头去算有多少列，不然从下面的行计算列的话，空的就不算了
            for (int j = 0; j < colCount; j++) {//第j列://+1是因为最后一列是空 也算进去
                Cell cell = row.getCell(j);
                String cellValue;
                boolean isMerge = false;
                if (cell != null) {
                    isMerge = isMergedRegion(sheet, i, cell.getColumnIndex());
                }
                //判断是否具有合并单元格
                if (isMerge) {
                    cellValue = getMergedRegionValue(sheet, row.getRowNum(), cell.getColumnIndex());
                } else {
                    cellValue = getStringCellValue(cell);
                }
                list.add(cellValue);
            }
            values.add(list);
        }
        return values;
    }
 
    /**
     *  判断整行是否为空
     *  @param row  excel得行对象
     *  @param  maxRow 有效值得最大列数
     */
    private static boolean CheckRowNull(Row row, int maxRow) {
        int num = 0;
        for (int j = 0; j < maxRow; j++) {
            Cell cell=row.getCell(j);
            if (cell==null||cell.equals("")||cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                num++;
            }
        }
        if(maxRow==num) return true;
        return false;
    }
    /**
     * 根据sheet数获取excel的值 返回List<List<Map<String,String>>>的数据结构
     * @param  fileUrl  文件路径
     * @param  sheetNum  工作表（第几分页[1,2,3.....]）
     * @return List<List<Map<String,String>>>
     */
    public  List<List<Map<String,String>>> getExcelMapVal(String fileUrl, int sheetNum) throws Exception{
        List<List<Map<String,String>>> values = new ArrayList<List<Map<String,String>>>();
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(is);
        int sheetCount = sheetNum - 1; //workbook.getNumberOfSheets();//sheet 数量,可以只读取手动指定的sheet页
        //int sheetCount1= workbook.getNumberOfSheets();
        Sheet sheet = workbook.getSheetAt(sheetCount); //读取第几个工作表sheet
        int rowNum = sheet.getLastRowNum();//有多少行
        Row rowTitle = sheet.getRow(0);//第i行
        int colCount = sheet.getRow(0).getLastCellNum();//用表头去算有多少列，不然从下面的行计算列的话，空的就不算了
        for (int i = 1; i <= rowNum; i++) {
            Row row = sheet.getRow(i);//第i行
            if (row == null || CheckRowNull(row,colCount)) {//过滤空行
                continue;
            }
            List<Map<String,String>> list = new ArrayList<Map<String,String>>();
            for (int j = 0; j < colCount; j++) {//第j列://+1是因为最后一列是空 也算进去
                Map<String,String> map=new HashMap<>();
                Cell cell = row.getCell(j);
                Cell cellTitle = rowTitle.getCell(j);
                String cellValue;
                String cellKey=getStringCellValue(cellTitle);
                boolean isMerge = false;
                if (cell != null) {
                    isMerge = isMergedRegion(sheet, i, cell.getColumnIndex());
                }
                //判断是否具有合并单元格
                if (isMerge) {
                    cellValue = getMergedRegionValue(sheet, row.getRowNum(), cell.getColumnIndex());
                } else {
                    cellValue = getStringCellValue(cell);
                }
                map.put(cellKey,cellValue);
                list.add(map);
            }
            values.add(list);
        }
        return values;
    }
 
    /**
     * 获取当前excel的工作表sheet总数
     * @param fileUrl
     * @return
     * @throws Exception
     */
    public int hasSheetCount(String fileUrl)throws Exception{
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(is);
        int sheetCount= workbook.getNumberOfSheets();
        return sheetCount;
    }
}
