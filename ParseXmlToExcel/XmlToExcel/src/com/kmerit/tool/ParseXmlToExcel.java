package com.kmerit.tool;

import com.kmerit.domains.Acc;
import com.kmerit.util.ParseXmlToExcelUtil;

/**
 * Created by shenlj on 2018/1/3.
 */
public class ParseXmlToExcel {
    public static void main(String[] args) {
    	String importFile = null;
    	String exportFile = null;
    	if(args != null && args.length == 2){
    		importFile = args[0];
    		exportFile = args[1];
    	}else{
    		System.out.println("文件输入输出路径为空！");
    	}
    	System.out.println(importFile + "|" + exportFile);
    	long startTime = System.currentTimeMillis();
//    	ParseXmlToExcelUtil pxteu = new ParseXmlToExcelUtil(Acc.class, "C:\\Users\\T-shenlj\\Desktop\\backup\\Acc.xml", "C:\\Users\\T-shenlj\\Desktop\\backup\\Acc.xlsx");
    	ParseXmlToExcelUtil pxteu = new ParseXmlToExcelUtil(Acc.class, importFile, exportFile);
    	pxteu.parseAndExport();
        long endTime = System.currentTimeMillis();
        System.out.println("完成处理，耗时:"+(endTime-startTime)/1000+"s");
    }
}
