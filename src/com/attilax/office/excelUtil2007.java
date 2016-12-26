package com.attilax.office;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import net.sf.json.JSONObject;
import net.sf.json.JsonConfig;

import org.apache.commons.beanutils.BeanMap;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.attilax.Closure;
import com.attilax.core;
import com.attilax.Stream.Mapx;
import com.attilax.collection.listUtil;
import com.attilax.exception.ExUtil;
import com.attilax.io.filex;
import com.attilax.json.AtiJson;
import com.attilax.lang.MapX;
import com.attilax.text.strUtil;
import com.attilax.util.Func_4SingleObj;
import com.attilax.util.Funcx;
import com.attilax.util.tryX;
import com.attilax.util.utf8编码;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

//import tst.TMbAwardWeixinResult;
@utf8编码
/**
 * use age  jsp:::
 exportWinner.exp(response,request.getParameter("id"));
 * @author  attilax 老哇的爪子
 *@since  o0f 2_q_9$
 */
public class excelUtil2007 {
	/**
	 * @功能：手工构建一个简单格式的Excel
	 */
	private static List<Map> getStudent() throws Exception {
		List list = new ArrayList();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-mm-dd");

		Map user1 = new HashMap();
		user1.put("awardName", "awardNameval11");
		user1.put("nickname", "nicknameval2");
		list.add(user1);
		// list.add(user2);
		// list.add(user3);

		return list;
	}

	@SuppressWarnings("all")
	public static void main(String[] args) throws Exception {
		
		
		int cellNum = 0;
		String fmt = com.attilax.lang.text.strUtil.fmt("--get cellval e,row index is:$rowidx$,cell index is :$cellidx$",rowIdx,cellNum);
		System.out.println(fmt);
		if(!fmt.equals(""))return;
		
		List<Map>	 li= new excelUtil2007().toListMap("c:\\功能表.xlsx");
		
		for (Map map : li) {
			String cols=(String) map.get("字段");
			cols=com.attilax.lang.text.strUtil.toEnChar(cols);
			cols=cols.replaceAll("\\(.*?\\)", "");
			List<String> cols_li=com.attilax.lang.text.strUtil.splitx2li(cols, new ArrayList<String>() {
				{
					this.add(",");this.add(" ");
				}
			});
			System.out.println(AtiJson.toJson(cols_li));
		}
	//	System.out.println(AtiJson.toJson(li));

		  System.out.println("---f");
		// System.out.println("---aa");
		// System.out.println("---b");
	}

	public static <ati> void toExcel(String titles, String filds, List<ati> li,
			HttpServletResponse response) {
		List list = listUtil.map_generic(li, new Func_4SingleObj<ati, Map>() {

			@Override
			public Map invoke(final ati o) {
				// attilax 老哇的爪子 下午03:32:08 2014-6-6
				Map mp = new tryX<Map>() {

					@Override
					public Map item(Object t) throws Exception {
						// attilax 老哇的爪子 下午02:47:29 2014-5-28
						// core.print(o);
						if (o instanceof Map)
							return (Map) o;
						// Map m = core.obj2map(o);
						JSONObject m = JSONObject.fromObject(o);
						return m;
					}

				}.$(new HashMap());

				return mp;
			}

		});
		core.log("--o6a wait exp list:" + core.obj2jsonO5(list));
		try {
			response.reset();
			response.setContentType("application/vnd.ms-excel");
			response.setCharacterEncoding("GB2312");
			String downFilename = new String(
					(filex.getUUidName() + ".xls").getBytes(), "iso-8859-1");
			response.setHeader("Content-Disposition", "attachment;filename="
					+ downFilename);
			toExcel(titles, filds, list, response.getOutputStream());
		} catch (Exception e) {
			core.log(e);
		}
	}

	public static ThreadLocal<JsonConfig> clsOa7 = new ThreadLocal<JsonConfig>();
	public static ThreadLocal<Closure> clsOa7a = new ThreadLocal<Closure>();

	public static <ati> void toExcel(String xlsName, String titles,
			String filds, List<ati> li, HttpServletResponse response) {
		List list = listUtil.map_generic(li, new Func_4SingleObj<ati, Map>() {

			@Override
			public Map invoke(final ati o) {
				// attilax 老哇的爪子 下午03:32:08 2014-6-6
				Map mp = new tryX<Map>() {

					@Override
					public Map item(Object t) throws Exception {
						// attilax 老哇的爪子 下午02:47:29 2014-5-28
						// core.print(o);
						// if (o instanceof Map)
						// return (Map) o;
						// Map m = core.obj2map(o);
						if (clsOa7a.get() != null) {
							JSONObject m = JSONObject.fromObject(o);

							return (Map) clsOa7a.get().execute(m);
						} else {
							JSONObject m = JSONObject.fromObject(o);
							// String s=core.toJsonStrO88(m, clsOa7);
							return m;
						}
					}

				}.$(new HashMap());

				return mp;
			}

		});
		core.log("--o6a wait exp list:" + core.obj2jsonO5(list));
		try {
			response.reset();
			response.setContentType("application/vnd.ms-excel");
			response.setCharacterEncoding("GB2312");
			String downFilename = new String(xlsName.getBytes(), "iso-8859-1");
			response.setHeader("Content-Disposition", "attachment;filename="
					+ downFilename);
			toExcel(titles, filds, list, response.getOutputStream());
		} catch (Exception e) {
			core.log(e);
		}
	}

	public static void toExcelMap(String titles, String filds, List<Map> list,
			HttpServletResponse response) {
		try {
			response.setContentType("application/vnd.ms-excel");
			response.setCharacterEncoding("GB2312");
			response.setHeader(
					"Content-Disposition",
					"attachment;filename="
							+ new String((filex.getUUidName()).getBytes(),
									"iso-8859-1"));
			toExcel(titles, filds, list, response.getOutputStream());
		} catch (Exception e) {
			core.log(e);
		}
	}

	public static void toExcelMap(String xlsName, String titles, String filds,
			List<Map> list, HttpServletResponse response) {
		try {
			response.setContentType("application/vnd.ms-excel");
			response.setCharacterEncoding("GB2312");
			response.setHeader("Content-Disposition", "attachment;filename="
					+ new String(xlsName.getBytes(), "iso-8859-1"));
			toExcel(titles, filds, list, response.getOutputStream());
		} catch (Exception e) {
			core.log(e);
		}
	}

	private static void toExcel(String titles, String filds, List<Map> list,
			OutputStream outStrm) throws Exception {
		// 第一步，创建一个webbook，对应一个Excel文件
		HSSFWorkbook wb = new HSSFWorkbook();
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet("sheet1");
		// sheet.setColumnWidth(columnIndex, width)

		// 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
		HSSFRow row = sheet.createRow((int) 0);
		// 第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
		// style.set

		String[] tit_arr = titles.split(",");
		// column index form 0;
		for (int i = 0; i < tit_arr.length; i++) {
			sheet.setColumnWidth(i, 7 * 2 * 256);
		}
		sheet.setColumnWidth(6, 20 * 2 * 256);
		sheet.setColumnWidth(7, 7 * 2 * 256);
		// //这里你会发现一个有趣的现象，SetColumnWidth的第二个参数要乘以256，这是怎么回事呢？其实，这个参数的单位是1/256个字符宽度

		int n = 0;
		for (String tit : tit_arr) {
			HSSFCell cell = row.createCell((short) n);
			cell.setCellValue(tit);
			cell.setCellStyle(style);
			n++;
		}

		// 第五步，写入实体数据 实际应用中这些数据从数据库得到，

		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow((int) i + 1);
			Map stu = (Map) list.get(i);

			// 第四步，创建单元格，并设置值
			int n2 = 0;
			for (String tit : tit_arr) {

				String curField = getFild(filds, n2);
				String val = "";
				/*
				 * if(curField.contains(".")){
				 * //System.out.println(Ognl.getValue("#eq.dpt.groupname",
				 * stu)); System.out.println();mtrl.materialId }
				 */
				// val = strUtil.toStr( stu.get(curField));
				stu.get("mtrl");
				val = String.valueOf(Mapx.get(stu, curField));
				if (val == null) {
					n2++;
					continue;
				}
				HSSFCell cell = row.createCell((short) n2);
				cell.setCellValue(val);
				// cell.setCellStyle(style);
				n2++;
			}

		}
		// 第六步，将文件存到指定位置
		try {
			// String outputFilePath = "E:/students.xls";
			// FileOutputStream fout = new FileOutputStream(outputFilePath);
			wb.write(outStrm);
			outStrm.close();
		} catch (Exception e) {
			core.log(e);
			e.printStackTrace();
		}
	}

	public void toExcel(List<Map> list, String outputFilePath) {

		try {
			if (list.size() == 0)
				throw new RuntimeException("list is empty");
			String keys = MapX.getKeysStr(list.get(0));
			toExcel(keys, keys, list, outputFilePath);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			throw new RuntimeException(e);
		}
	}

	@SuppressWarnings("")
	public static void toExcel(String titles, String filds, List<Map> list,
			String outputFilePath) throws Exception {
		// 第一步，创建一个webbook，对应一个Excel文件
		HSSFWorkbook wb = new HSSFWorkbook();
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet("sheet1");
		// 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
		HSSFRow row = sheet.createRow((int) 0);
		// 第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

		String[] tit_arr = titles.split(",");
		int n = 0;
		for (String tit : tit_arr) {
			HSSFCell cell = row.createCell((short) n);
			cell.setCellValue(tit);
			// cell.set
			cell.setCellStyle(style);
			n++;
		}

		// 第五步，写入实体数据 实际应用中这些数据从数据库得到，

		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow((int) i + 1);
			Map stu = (Map) list.get(i);

			// 第四步，创建单元格，并设置值
			int colIndex = 0;
			for (String tit : tit_arr) {
				String curField = getFild(filds, colIndex);
				Object v = stu.get(curField);
				if (v == null)
					v = "";
				// String val =v.toString();
				// // try{
				// // val= v.toString();
				// // }catch(Exception e){}
				// if(val==null)continue;
				HSSFCell cell = row.createCell((short) colIndex);

				if (v instanceof Integer || v instanceof Long
						|| v instanceof Float || v instanceof Double)
					cell.setCellValue(Double.valueOf(v.toString()));

				else
					cell.setCellValue(v.toString());
				// cell.setCellStyle(style);
				colIndex++;
			}

		}
		// 第六步，将文件存到指定位置
		try {
			// String outputFilePath = "E:/students.xls";
			FileOutputStream fout = new FileOutputStream(outputFilePath);
			wb.write(fout);
			fout.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * @author attilax 老哇的爪子
	 * @since 2014-6-6 下午02:58:45$
	 * 
	 * @param filds
	 * @param n2
	 * @return
	 */
	private static String getFild(String filds, int n2) {
		// attilax 老哇的爪子 下午02:58:45 2014-6-6
		String[] fs = filds.split(",");
		for (int n = 0; n < fs.length; n++) {
			if (n2 == n)
				return fs[n];
		}
		return "";
	}

	// import java.io.FileInputStream;
	// import java.io.IOException;
	// import java.io.InputStream;
	// import java.util.ArrayList;
	// import java.util.List;
	//
	// import org.apache.poi.hssf.usermodel.HSSFCell;
	// import org.apache.poi.hssf.usermodel.HSSFRow;
	// import org.apache.poi.hssf.usermodel.HSSFSheet;
	// import org.apache.poi.hssf.usermodel.HSSFWorkbook;

	/**
	 * 
	 * @author Hongten</br>
	 * 
	 *         参考地址：http://hao0610.iteye.com/blog/1160678
	 * 
	 */
	// public class XlsMain {
	//
	// public static void main(String[] args) throws IOException {
	// XlsMain xlsMain = new XlsMain();
	// XlsDto xls = null;
	// List<XlsDto> list = xlsMain.readXls();
	//
	// try {
	// XlsDto2Excel.xlsDto2Excel(list);
	// } catch (Exception e) {
	// e.printStackTrace();
	// }
	// for (int i = 0; i < list.size(); i++) {
	// xls = (XlsDto) list.get(i);
	// System.out.println(xls.getXh() + "    " + xls.getXm() + "    "
	// + xls.getYxsmc() + "    " + xls.getKcm() + "    "
	// + xls.getCj());
	// }
	//
	// }
	//
	Map schemaIndex = Maps.newLinkedHashMap();
	private static int rowIdx;

	/**
	 * 读取xls文件内容
	 * 
	 * @return List<XlsDto>对象
	 * @throws IOException
	 *             输入/输出(i/o)异常
	 */
	private List<Map> readXls(String f) throws IOException {
		// f = "pldrxkxxmb.xls";
		InputStream is = new FileInputStream(f);
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		// XlsDto xlsDto = null;
		List<Map> list = new ArrayList<Map>();
		// 循环工作表Sheet
	//	for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
			if (hssfSheet == null) {
				return list;
			}
			setSchema(hssfSheet);
			// 循环行Row
			int lastRowNum = hssfSheet.getLastRowNum();
			for (int rowNum = 1; rowNum <= lastRowNum; rowNum++) {
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				if (hssfRow == null) {
					continue;
				}
				Map m = Maps.newLinkedHashMap();
				// 循环列Cell
				// 0学号 1姓名 2学院 3课程名 4 成绩
				for (int cellNum = 0; cellNum <= 20; cellNum++) {
					HSSFCell xh = hssfRow.getCell(cellNum);
					if (xh == null) {
						continue;
					}
					String key = (String) schemaIndex.get(cellNum);
					m.put(key, getValue(xh));

					
				}
				list.add(m);
			}
		//}
		return list;
	}

	/**
	 * 得到Excel表中的值
	 * 
	 * @param hssfCell
	 *            Excel中的每一个格子
	 * @return Excel中每一个格子中的值
	 */
	@SuppressWarnings("static-access")
	private String getValue(XSSFCell hssfCell) {
		try {
			if (hssfCell == null)
				return "";
			if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
				// 返回布尔类型的值
				return String.valueOf(hssfCell.getBooleanCellValue());
			} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
				// 返回数值类型的值
				return String.valueOf(hssfCell.getNumericCellValue());
			} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_STRING) {
				// 返回字符串类型的值
				return String.valueOf(hssfCell.getStringCellValue());
			} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_ERROR) {
				return "$e";
			} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BLANK) {
				return "";
			} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_FORMULA) {
				return "$formula";
			} else {
				throw new RuntimeException("err celltype :"
						+ hssfCell.getCellType());
			}
		} catch (IndexOutOfBoundsException e) {
			  ExUtil.throwEx(e);
		}
		return "";
		
		 
	}
	
	@SuppressWarnings("static-access")
	private String getValue(HSSFCell hssfCell) {
		if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
			// 返回布尔类型的值
			return String.valueOf(hssfCell.getBooleanCellValue());
		} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
			// 返回数值类型的值
			return String.valueOf(hssfCell.getNumericCellValue());
		} else {
			// 返回字符串类型的值
			return String.valueOf(hssfCell.getStringCellValue());
		}
	}

	@SuppressWarnings("unchecked")
	private void setSchema(HSSFSheet hssfSheet) {
		HSSFRow hssfRow = hssfSheet.getRow(0);
		for (int i = 0; i < 20; i++) {
			try {
				String value = getValue(hssfRow.getCell(i));
				schemaIndex.put(i, value);
			} catch (Exception e) {
				System.out.println( );
			}

		}
		// HSSFCell[] cell_arr = hssfRow.get
	}
	
	@SuppressWarnings("unchecked")
	private void setSchema07(XSSFSheet sheet) {
		XSSFRow hssfRow = sheet.getRow(0);
		for (int i = 0; i < 20; i++) {
			try {
				String value = getValue(hssfRow.getCell(i));
				schemaIndex.put(i, value);
			} catch (Exception e) {
				Object cellNum=i;
				String fmt = com.attilax.lang.text.strUtil.fmt("--get cellval e,row index is:$rowidx$,cell index is :$cellidx$",rowIdx,cellNum);
				System.out.println(fmt  +  e.getMessage());
			}

		}
		// HSSFCell[] cell_arr = hssfRow.get
	}


	
	   public List<Map> readExcel07(String filepath) throws IOException{
	        List<Map> fsnList = Lists.newLinkedList();
	        //取得excel
	        XSSFWorkbook xwb = new XSSFWorkbook(filepath);
	        //取得Excel的第一个sheet;
	        XSSFSheet sheet = xwb.getSheetAt(0);
	    	setSchema07(sheet);
	        XSSFRow row;
	        //遍历sheet的所有行，前两个单元格，设置为Info的属性，放入ArrayList返回
	        for (int i = sheet.getFirstRowNum(); i < sheet.getPhysicalNumberOfRows(); i++) { 
	        
	          rowIdx=i;
	            row = sheet.getRow(i);
	        	if (row == null) {
	    			continue;
	    		}
	        	Map   fsn=row2map(row);
	            fsnList.add(fsn);
	        } 
	        return fsnList;
	    }

	   
	   
	private Map row2map(XSSFRow hssfRow) {
	
		Map m = Maps.newLinkedHashMap();
		// 循环列Cell
		// 0学号 1姓名 2学院 3课程名 4 成绩
		for (int cellNum = 0; cellNum <= 20; cellNum++) {
			XSSFCell cell = hssfRow.getCell(cellNum);
			if (cell == null) {
				continue;
			}
			String key = (String) schemaIndex.get(cellNum);
			try {
				m.put(key, getValue(cell));
			} catch (Exception e) {
				String fmt = com.attilax.lang.text.strUtil.fmt("--get cellval e,row index is:$rowidx$,cell index is :$cellidx$",rowIdx,cellNum);
				System.out.println(fmt  +  e.getMessage());
			}
			

			
		}
		return m;
	}

	public List<Map> toListMap(String f) {
	try {
		return	 readExcel07(f);
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();throw new RuntimeException(e);
		
	}
	}

	/**
	 * @author attilax 老哇的爪子
	 * @since 2014-6-6 下午02:57:53$
	 * 
	 * @param n2
	 * @return
	 */
	// private static String getFild(int n2) {
	// // attilax 老哇的爪子 下午02:57:53 2014-6-6
	// return null;
	// }
}