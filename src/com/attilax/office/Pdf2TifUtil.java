/**
 * 
 */
package com.attilax.office;

/**
 * @author ASIMO
 *
 */
public class Pdf2TifUtil {
	
	  public static void main(String[] args) {
		  
		  String pdfPath="C:/00/关于优化研发一部组织架构的通知（2015001）.pdf";
		String imgDirPath="c:/pdf2tifOut";
		pdf2tif(pdfPath, imgDirPath);
		
	  }

		/**
		@author attilax 老哇的爪子
		@since   p1r a_j_t
		 
		 */
	private static void pdf2tif(String pdfPath, String imgDirPath) {
		try {
			Office2Tif.pdf2Imgs(pdfPath, imgDirPath, ".jpg");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			
		}
		
	}

}
