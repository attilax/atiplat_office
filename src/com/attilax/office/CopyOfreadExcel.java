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
 
 
/**
 *
 * @author Hongten</br>
 *
 *         参考地址：http://hao0610.iteye.com/blog/1160678
 *
 */
public class CopyOfreadExcel {
 
    public static void main(String[] args) throws IOException {
        CopyOfreadExcel xlsMain = new CopyOfreadExcel();
        Map xls = null;
        List<Map> list = xlsMain.readXls("D:\\Users\\attilax\\Documents\\Tencent Files\\1466519819\\FileRecv\\11档微信价\\11档微信价（武汉）.xlsx");
         
//        try {
//            Map2Excel.Map2Excel(list);
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        for (int i = 0; i < list.size(); i++) {
//            xls = (Map) list.get(i);
//            System.out.println(xls.getXh() + "    " + xls.getXm() + "    "
//                    + xls.getYxsmc() + "    " + xls.getKcm() + "    "
//                    + xls.getCj());
//        }
 
    }
 
    /**
     * 读取xls文件内容
     *
     * @return List<Map>对象
     * @throws IOException
     *             输入/输出(i/o)异常
     */
    public List<Map> readXls( String filename) throws IOException {
       // = "pldrxkxxmb.xls";
		InputStream is = new FileInputStream(filename);
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
        Map Map = null;
        List<Map> list = new ArrayList<Map>();
        // 循环工作表Sheet
        for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            if (hssfSheet == null) {
                continue;
            }
            
            //   V0o o7d  老哇的爪子  Attilax
            String[] titles=getTits(hssfSheet);
            // 循环行Row
            for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow == null) {
                    continue;
                }
                Map mp= new HashMap();
                // 循环列Cell
                // 0学号 1姓名 2学院 3课程名 4 成绩
                // for (int cellNum = 0; cellNum <=4; cellNum++) {
                //   Vj7 o7d  老哇的爪子  Attilax
                for(int i=0;i<20;i++)
     		   {
     			    HSSFCell cell = hssfRow.getCell(i);
     	            if (cell == null) {
     	                continue;
     	            }
     	            mp.put(titles[i],  getValue(cell));
     		   }
//                HSSFCell xh = hssfRow.getCell(0);
//                if (xh == null) {
//                    continue;
//                }
//                Map.setXh(getValue(xh));
//                HSSFCell xm = hssfRow.getCell(1);
//                if (xm == null) {
//                    continue;
//                }
//                Map.setXm(getValue(xm));
//                HSSFCell yxsmc = hssfRow.getCell(2);
//                if (yxsmc == null) {
//                    continue;
//                }
//                Map.setYxsmc(getValue(yxsmc));
//                HSSFCell kcm = hssfRow.getCell(3);
//                if (kcm == null) {
//                    continue;
//                }
//                Map.setKcm(getValue(kcm));
//                HSSFCell cj = hssfRow.getCell(4);
//                if (cj == null) {
//                    continue;
//                }
//                Map.setCj(Float.parseFloat(getValue(cj)));
                list.add(mp);
            }
        }
        return list;
    }
 
    /**
	@author attilax 老哇的爪子
		@since  o7d Va2$
	
	 * @param hssfSheet
	 * @return
	 */
	private String[] getTits(HSSFSheet hssfSheet) {
		// attilax 老哇的爪子  Va2   o7d 
		   HSSFRow hssfRow = hssfSheet.getRow(0);
//           if (hssfRow == null) {
//               continue;
//           }
		   Map m= new HashMap();
		   List li=new ArrayList<String>();
		   for(int i=0;i<20;i++)
		   {
			    HSSFCell cell = hssfRow.getCell(i);
	            if (cell == null) {
	                continue;
	            }
	            li.add(getValue(cell));
		   }
          
		return   (String[]) li.toArray(new String[li.size()]);
		 
		
	}

	/**
     * 得到Excel表中的值
     *
     * @param hssfCell
     *            Excel中的每一个格子
     * @return Excel中每一个格子中的值
     */
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
 
}