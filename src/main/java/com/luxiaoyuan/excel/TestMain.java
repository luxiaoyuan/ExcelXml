package com.luxiaoyuan.excel;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
 
/**
 * 测试类
 */
public class TestMain {
    public static void main(String[] args) throws Exception {
        //单级的表头
//        Map<String,String> map=new HashMap<String,String>();
//        map.put("登录名","uid");
//        Map<String,String>  map1=new HashMap<String,String>();
//        map1.put("用户名","uname");
//        Map<String,String>  map2=new HashMap<String,String>();
//        map2.put("角色","urole");
//        Map<String,String>  map3=new HashMap<String,String>();
//        map3.put("部门","udep");//d_name
//        Map<String,String>  map4=new HashMap<String,String>();
//        map4.put("用户类型","utype");
//        List<Map<String,String>> titleList=new ArrayList<>();
//        titleList.add(map); titleList.add(map1); titleList.add(map2); titleList.add(map3); titleList.add(map4);
//        //单级的 行内数据
//        List<Map<String,String>> rowList=new ArrayList<>();
//        for(int i=0;i<7;i++){
//            Map m= new HashMap<String,String>();
//            m.put("uid","登录名"+i); m.put("uname","张三"+i);
//            m.put("urole","角色"+i); m.put("udep","部门"+i);
//            m.put("utype","用户类型"+i);
//            rowList.add(m);
//        }
//
//        ExcelTool excelTool = new ExcelTool("单级表头的表格",15,20);
//        List<Column>  titleData=excelTool.columnTransformer(titleList);
//        excelTool.exportExcel(titleData,rowList,"D://outExcel.xls",true,false);
 
        //List<Map>数据 多级表头,数据如下:
       //        登录名  姓名       aa
       //                      角色    部门
        List<Map<String,String>> titleList=new ArrayList<>();
        Map<String,String> titleMap=new HashMap<String,String>();
        titleMap.put("id","11");titleMap.put("pid","0");titleMap.put("content","登录名");titleMap.put("fielName","uid");
        Map<String,String> titleMap1=new HashMap<String,String>();
        titleMap1.put("id","1");titleMap1.put("pid","0");titleMap1.put("content","姓名");titleMap1.put("fielName","uname");
        Map<String,String> titleMap2=new HashMap<String,String>();
        titleMap2.put("id","2");titleMap2.put("pid","0");titleMap2.put("content","角色加部门");titleMap2.put("fielName",null);
        Map<String,String> titleMap3=new HashMap<String,String>();
        titleMap3.put("id","3");titleMap3.put("pid","2");titleMap3.put("content","角色");titleMap3.put("fielName","urole");
        Map<String,String> titleMap4=new HashMap<String,String>();
        titleMap4.put("id","4");titleMap4.put("pid","2");titleMap4.put("content","部门");titleMap4.put("fielName","udep");
        titleList.add(titleMap); titleList.add(titleMap1); titleList.add(titleMap2); titleList.add(titleMap3); titleList.add(titleMap4);
       // 单级的 行内数据
        List<Map<String,String>> rowList=new ArrayList<>();
        for(int i=0;i<7;i++){
            Map m= new HashMap<String,String>();
            m.put("uid","登录名"+i); m.put("uname","张三"+i);
            m.put("urole","角色"+i); m.put("udep","部门"+i);
            m.put("utype","用户类型"+i);
            rowList.add(m);
        }
        ExcelTool excelTool = new ExcelTool("List<Map>数据 多级表头表格",20,20);
        List<Column>  titleData=excelTool.columnTransformer(titleList,"id","pid","content","fielName");
        excelTool.exportExcel(titleData,rowList,"D:\\GUANYU\\excel\\outExcel.xls",true,false);
 
        //实体类（entity）数据 多级表头,数据如下:
        //        登录名  姓名       aa
        //                      角色    部门
//        List<TitleEntity> titleList=new ArrayList<>();
//        TitleEntity titleEntity=new TitleEntity("11","0","登录名","uid");
//        TitleEntity titleEntity1=new TitleEntity("1","0","姓名","uname");
//        TitleEntity titleEntity2=new TitleEntity("2","0","角色加部门",null);
//        TitleEntity titleEntity3=new TitleEntity("3","2","角色","urole");
//        TitleEntity titleEntity4=new TitleEntity("4","2","部门","udep");
//        titleList.add(titleEntity); titleList.add(titleEntity1); titleList.add(titleEntity2); titleList.add(titleEntity3); titleList.add(titleEntity4);
//        //单级的 行内数据
//        List<Map<String,String>> rowList=new ArrayList<>();
//        for(int i=0;i<7;i++){
//            Map m= new HashMap<String,String>();
//            m.put("uid","登录名"+i); m.put("uname","张三"+i);
//            m.put("urole","角色"+i); m.put("udep","部门"+i);
//            m.put("utype","用户类型"+i);
//            rowList.add(m);
//        }
//        ExcelTool excelTool = new ExcelTool("实体类（entity）数据 多级表头表格",20,20);
//        List<Column>  titleData=excelTool.columnTransformer(titleList,"t_id","t_pid","t_content","t_fielName");
//        excelTool.exportExcel(titleData,rowList,"D://outExcel.xls",true,true);
 
          //读取excel
//          ExcelTool excelTool = new ExcelTool();
//          List<List<String>> readexecl=excelTool.getExcelValues("D://outExcel.xls",1);
//          List<List<Map<String,String>>> readexeclC=excelTool.getExcelMapVal("D://outExcel.xls",1);
//          int count= excelTool.hasSheetCount("D://outExcel.xls");
 
    }
 
}