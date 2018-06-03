package com.kmerit.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Created by shenlj on 2018/1/3.
 */
public class ParseXmlToExcelUtil {
	private static Field[] fields;
	private static List<Object> xmlVoList = new ArrayList<Object>();
	private static Object xmlVo2;
	private static Class clazz;
	private static String firstField = "GeneratedPK";
	private static Map<String, Class> fieldString = new HashMap<String, Class>();
//	private static int count = 0;
	private static String xmlPath;
	private static String excelPath;

	
	
	public ParseXmlToExcelUtil(Class clazz, String xmlPath, String excelPath) {
		this.clazz = clazz;				//xml待转对象属性-Excel文件表头
		this.xmlPath = xmlPath;			//源xml文件路径 
		this.excelPath = excelPath;		//目标Excel文件生成路径
	}
	
	

	/**
	 * 文件处理主方法：读取Xml文件，并解析转换成Excel文件
	 * 因处理文件数据量较大，分两步进行处理，即先生成带有表头的Excel文件；再获取该Excel文件，向其写入数据内容
	 */
	public void parseAndExport() {
		/**
		 * 读取xml文件，生成带有表头的Excel文件
		 */
		XSSFWorkbook wb = null;
		FileOutputStream fileOutputStream = null;
		try {
			wb = wbSetPrepare();
			fileOutputStream = new FileOutputStream(excelPath);
			wb.write(fileOutputStream);
		} catch (Exception e) {
			System.out.println("生成带有表头的Excel文件异常 ");
			e.printStackTrace();
		} finally {
			try {
				if (fileOutputStream != null)
					fileOutputStream.close();
			} catch (IOException e) {
				System.out.println("关闭文件输出流异常 ");
				e.printStackTrace();
			}
		}
		
		/**
		 * 解析Xml文件存储数据，封装List
		 * 
		 */
		System.out.println("开始解析Xml文件...");
		parseXml(xmlPath);
		xmlVoList.add(xmlVo2);		//计数器漏掉最后一条，在此添加
		
		/**
		 * 将List中的内容写入到Excel文件中
		 */
		XSSFWorkbook workbook1 = null;
		FileInputStream targetFis = null;
		FileOutputStream out = null;
		try {
			targetFis = new FileInputStream(new File(excelPath));
			workbook1 = new XSSFWorkbook(targetFis);
	        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(workbook1, 100);
	        Sheet firstSheet = sxssfWorkbook.getSheet("Acc");
	        setExcelValue(firstSheet);					//为Sheet对象，填充数据
	        
	        /**
	         * 获取输出流，将Sheet内容写入
	         */
	        out = new FileOutputStream(excelPath);
	        sxssfWorkbook.write(out);
		} catch (Exception e) {
			System.out.println("Sheet内容写入异常");
			e.printStackTrace();
		} finally{
			try{
				if(targetFis != null)
					targetFis.close();
				if(out != null)
					out.close();
			}catch(Exception e){
				System.out.println("关闭文件输出流异常 ");
				e.printStackTrace();
			}
		
		}
		
	}
	
	

	/**
	 * 创建excel，设置excel基础属性
	 * 
	 * @return
	 * @throws Exception
	 */
	private static XSSFWorkbook wbSetPrepare() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Acc");		// 声明1个sheet并为其命名
		sheet.setDefaultColumnWidth((short) 18);		// 设置默认列长度
		XSSFCellStyle style = wb.createCellStyle();		// 生成�?个样�?
		style.setAlignment(HorizontalAlignment.CENTER); // 样式字体居中
		toExportExcelPrepare(sheet, style, xmlPath);
		return wb;
	}

	/**
	 * 做导出到excel的准备工作，生成Sheet表头
	 * 
	 * @param sheet
	 * @param style
	 * @param xmlPath
	 * @throws Exception
	 */
	private static void toExportExcelPrepare(XSSFSheet sheet, XSSFCellStyle style, String xmlPath) throws Exception {
		XSSFRow row = sheet.createRow(0); // 创建第一行（也可以称为表头）
		XSSFCell cell = row.createCell((short) 0);// 给表头第�?行一次创建单元格
		xmlVo2 = clazz.newInstance();
		fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			System.out.println(field.getName() + "------------" + field.getType());
			fieldString.put(field.getName(), field.getType());
		}
		for (int i = 0, j = fields.length; i < j; i++) {
			cell = row.createCell(i);
			cell.setCellValue(fields[i].getName());
			cell.setCellStyle(style);
		}
	}

	/**
	 * 设置Sheet下单元格值（填充）
	 * 
	 * @param sheet
	 * @throws Exception
	 */
	private static void setExcelValue(Sheet sheet) throws Exception {
		//xmlVoList 第一个对象为无值对象
		for (int i = 1, j = xmlVoList.size(); i < j; i++) {
			Row row = sheet.createRow(i);
			Object xmlVo = xmlVoList.get(i);
			for (int k = 0, l = fields.length; k < l; k++) {
				Cell cell = row.createCell(k);
				String methodName = "get" + toUpperCase4Index(fields[k].getName());
//				System.err.println(methodName);
				Method method = xmlVo.getClass().getDeclaredMethod(methodName);
//				System.err.println(method);
				invokeGetMethod(method, xmlVo, fields[k].getName(), cell);
			}
		}
		System.out.println("Sheet数据填充结束#######，开始写入...");
		
	}

	
	
	/**
	 * 解析Xml文件
	 * @param element
	 */
	private static void parseXml(String xmlPath) {
		SAXReader reader = new SAXReader();
		Document document = null;
		try {
			document = reader.read(new File(xmlPath));
			Element root = document.getRootElement();
			listNodes(root);
//			System.out.println("计数#######:"+count);
		} catch (DocumentException e) {
			e.printStackTrace();
		}
	}


	/**
	 * 递归遍历Xml文件节点属性 
	 * @param element
	 */
	private static void listNodes(Element element) {
		// 获取节点的所有属�?
		List<Attribute> attributes = element.attributes();
//		for (Attribute attr : attributes) {
//			System.out.println("节点名字�?" + element.getName() + "节点属�?�名�?" + attr.getName() + "节点属�?��?�为" + attr.getValue());
//		}
		try {
			for (Attribute attr : attributes) {
				if(firstField.equals(attr.getName())){
					xmlVoList.add(xmlVo2);
					xmlVo2 = clazz.newInstance();
				}
				Method method = clazz.getDeclaredMethod("set" + toUpperCase4Index(attr.getName()),
						fieldString.get(attr.getName()));
//				System.err.println(method+"#######");
				invokeSetMethod(xmlVo2, element, method,attr.getValue());

				/**  解析内容对象有标准模板，但内容出现空字段，而又没有空标签时，此方法存在Bug 
				if (fieldString.containsKey(attr.getName())) {
					if (count != 0 && count % fields.length == 0) {
						xmlVoList.add(xmlVo2);
						xmlVo2 = clazz.newInstance();
					}
					Method method = clazz.getDeclaredMethod("set" + toUpperCase4Index(attr.getName()),
							fieldString.get(attr.getName()));
//					System.err.println(method+"#######");
					invokeSetMethod(xmlVo2, element, method,attr.getValue());
					count++;
				}
				*/
			}
		} catch (Exception e) {
			System.out.println(element.getName() + "------------");
			e.printStackTrace();
		}
		Iterator elementIterator = element.elementIterator();
		while (elementIterator.hasNext()) {
			Element node = (Element) elementIterator.next();
			listNodes(node);
		}

	}

	private static void invokeGetMethod(Method method, Object xmlVo, String arg, Cell cell)
			throws InvocationTargetException, IllegalAccessException {
		Class clazz = fieldString.get(arg);
		Object value = method.invoke(xmlVo);
		if (value != null) {
			if (clazz == byte.class || clazz == Byte.class) {
				cell.setCellValue((Byte) value);
			} else if (clazz == short.class || clazz == Short.class) {
				cell.setCellValue((Short) value);
			} else if (clazz == int.class || clazz == Integer.class) {
				cell.setCellValue((Integer) value);
			} else if (clazz == long.class || clazz == Long.class) {
				cell.setCellValue((Long) value);
			} else if (clazz == char.class || clazz == Character.class) {// 用String吧，char好像没什么意�?
				cell.setCellValue((Character) value);
			} else if (clazz == float.class || clazz == Float.class) {
				cell.setCellValue((Float) value);
			} else if (clazz == double.class || clazz == Double.class) {
				cell.setCellValue((Double) value);
			} else if (clazz == boolean.class || clazz == Boolean.class) {
				cell.setCellValue((Boolean) value);
			} else if (clazz == String.class) {
				cell.setCellValue(String.valueOf(value));
			}
		}
	}

	private static void invokeSetMethod(Object xmlVo, Element element, Method method, String val) throws InvocationTargetException,
			IllegalAccessException {
		Class clazz = fieldString.get(element.getName());
		String value = val;
//		String value = element.getText();
		if (clazz == byte.class || clazz == Byte.class) {
			method.invoke(xmlVo, Byte.valueOf(value));
		} else if (clazz == short.class || clazz == Short.class) {
			method.invoke(xmlVo, Short.valueOf(value));
		} else if (clazz == int.class || clazz == Integer.class) {
			method.invoke(xmlVo, Integer.parseInt(value));
		} else if (clazz == long.class || clazz == Long.class) {
			method.invoke(xmlVo, Long.valueOf(value));
		} else if (clazz == char.class || clazz == Character.class) {
			method.invoke(xmlVo, value.charAt(0));
		} else if (clazz == float.class || clazz == Float.class) {
			method.invoke(xmlVo, Float.valueOf(value));
		} else if (clazz == double.class || clazz == Double.class) {
			method.invoke(xmlVo, Double.parseDouble(value));
		} else if (clazz == boolean.class || clazz == Boolean.class) {
			method.invoke(xmlVo, value.equals("true"));
		} else if (clazz == String.class) {
			method.invoke(xmlVo, value);
		}else {
			method.invoke(xmlVo, value);
		}
	}

	/**
	 * 首字母大
	 * @param string
	 * @return
	 */
	private static String toUpperCase4Index(String string) {
		char[] methodName = string.toCharArray();
		methodName[0] = toUpperCase(methodName[0]);
		return String.valueOf(methodName);
	}

	/**
	 * 字符转成大写
	 * @param chars
	 * @return
	 */
	private static char toUpperCase(char chars) {
		if (97 <= chars && chars <= 122) {
			chars -= 32;
		}
		return chars;
	}

}
