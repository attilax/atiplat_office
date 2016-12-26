package com.attilax.office;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import com.attilax.collection.Iflt;
import com.attilax.collection.listUtil;
import com.attilax.util.Funcx;

/** @author Hongten</br>
 * 
 * 参考地址：http://hao0610.iteye.com/blog/1160678 */
public class readExcel {

	public static void main(String[] args) throws IOException, InvalidFormatException {
		readExcel xlsMain = new readExcel();
		Map mp = null;
		List<Map> list = xlsMain.readXls("D:\\Users\\attilax\\Documents\\Tencent Files\\1466519819\\FileRecv\\11档微信价\\11档微信价（武汉）.xlsx");

		// try {
		// Map2Excel.Map2Excel(list);
		// } catch (Exception e) {
		// e.printStackTrace();
		// }
		List<Map> list2 = listUtil.filterO7(list, new Iflt<Map>() {

			@Override public boolean call(Map o) {
				// attilax 老哇的爪子 V59o o7d
				if (o.get("促销券号") == null && o.get("微信码") == null) return true;
				return false;

			}
		});
		
		
		for (int i = 0; i < list2.size(); i++) {
			mp = (Map) list2.get(i);
			System.out.println(mp.get("促销券号"));
		}

	}
	String[] titles;
	/** 读取xls文件内容
	 * 
	 * @return List<Map>对象
	 * @throws IOException 输入/输出(i/o)异常
	 * @throws InvalidFormatException */
	public List<Map> readXls(String filename) throws IOException, InvalidFormatException {
		// = "pldrxkxxmb.xls";
		InputStream is = new FileInputStream(filename);
		// HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		Workbook wb = WorkbookFactory.create(is);
		Map Map = null;
		List<Map> list = new ArrayList<Map>();
		// 循环工作表Sheet // wb.getNumberOfSheets()
		for (int numSheet = 0; numSheet < 1; numSheet++) {
			// XSSFShee
			Sheet hssfSheet = (Sheet) wb.getSheetAt(numSheet);
			if (hssfSheet == null) {
				continue;
			}

			// V0o o7d 老哇的爪子 Attilax
			titles = getTits(hssfSheet);
			// 循环行Row
			for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
				Row hssfRow = hssfSheet.getRow(rowNum);
				if (hssfRow == null) {
					continue;
				}
				Map mp = new HashMap();
				// 循环列Cell
				// 0学号 1姓名 2学院 3课程名 4 成绩
				// for (int cellNum = 0; cellNum <=4; cellNum++) {
				// Vj7 o7d 老哇的爪子 Attilax
				for (int i = 0; i < 20; i++) {
					Cell cell = hssfRow.getCell(i);
					if (cell == null) {
						continue;
					}
					if (i < titles.length) {
						String value = getValue(cell);
						// System.out.println("---"+value+"---");
						if (value != null && value.trim().length() > 0 && (!value.trim().toLowerCase().equals("null"))) mp.put(titles[i], value);
					}
				}
				// HSSFCell xh = hssfRow.getCell(0);
				// if (xh == null) {
				// continue;
				// }
				// Map.setXh(getValue(xh));
				// HSSFCell xm = hssfRow.getCell(1);
				// if (xm == null) {
				// continue;
				// }
				// Map.setXm(getValue(xm));
				// HSSFCell yxsmc = hssfRow.getCell(2);
				// if (yxsmc == null) {
				// continue;
				// }
				// Map.setYxsmc(getValue(yxsmc));
				// HSSFCell kcm = hssfRow.getCell(3);
				// if (kcm == null) {
				// continue;
				// }
				// Map.setKcm(getValue(kcm));
				// HSSFCell cj = hssfRow.getCell(4);
				// if (cj == null) {
				// continue;
				// }
				// Map.setCj(Float.parseFloat(getValue(cj)));
				list.add(mp);
			}
		}
		return list;
	}

	/** @author attilax 老哇的爪子
	 * @since o7d Va2$
	 * 
	 * @param hssfSheet
	 * @return */
	private String[] getTits(Sheet hssfSheet) {
		// attilax 老哇的爪子 Va2 o7d
		// hssfSheet.getr
		Row hssfRow = hssfSheet.getRow(0);
		// if (hssfRow == null) {
		// continue;
		// }
		Map m = new HashMap();
		List li = new ArrayList<String>();
		for (int i = 0; i < 20; i++) {
			Cell cell = hssfRow.getCell(i);
			if (cell == null) {
				continue;
			}
			String value = getValue(cell);
			li.add(value);
			m.put(value.trim(), "0");

		}

		if (m.get("促销券号") == null && m.get("微信码") == null) return getTits(hssfSheet, 1);

		return (String[]) li.toArray(new String[li.size()]);

	}

	private String[] getTits(Sheet hssfSheet, int titleRowIdx) {
		// attilax 老哇的爪子 Va2 o7d
		// hssfSheet.getr
		Row hssfRow = hssfSheet.getRow(titleRowIdx);
		// if (hssfRow == null) {
		// continue;
		// }
		Map m = new HashMap();
		List li = new ArrayList<String>();
		for (int i = 0; i < 20; i++) {
			Cell cell = hssfRow.getCell(i);
			if (cell == null) {
				continue;
			}
			String value = getValue(cell);
			li.add(value);
			m.put(value.trim(), "0");

		}

		return (String[]) li.toArray(new String[li.size()]);

	}

	/** 得到Excel表中的值
	 * 
	 * @param cell Excel中的每一个格子
	 * @return Excel中每一个格子中的值 */
	@SuppressWarnings("static-access") private String getValue(Cell cell) {
		if (cell.getCellType() == cell.CELL_TYPE_BOOLEAN) {
			// 返回布尔类型的值
			return String.valueOf(cell.getBooleanCellValue());
		} else if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
			// 返回数值类型的值
			return String.valueOf(cell.getNumericCellValue());
		} else {
			// 返回字符串类型的值
			return String.valueOf(cell.getStringCellValue());
		}
	}

}