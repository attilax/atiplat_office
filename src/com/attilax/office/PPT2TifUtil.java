/**
 * 
 */
package com.attilax.office;

import java.io.File;
import java.util.Iterator;

import org.apache.commons.io.FileUtils;

import com.attilax.img.ImgX;

/**
 * @author ASIMO
 *
 */
public class PPT2TifUtil {
	
	public static void main(String[] args) {
    
		ppt2tif("C:/00/研发项目管理办法.ppt","c:/ppt2tifOut2");
		
	}

		/**
		@author attilax 老哇的爪子
		@since   p1r a_g_l
		 
		 */
	private static void ppt2tif(String pptfF, String outDir) {
		Office2Html. convert(new File(pptfF), outDir);

		 Iterator itFile =FileUtils.iterateFiles(new File( outDir), new String[]{"jpg"}, true);
		  while (itFile.hasNext()) {   
	            File file = (File) itFile.next();   
	            ImgX.jpg2tif(file.getAbsolutePath())	;
	        }   
	}

}
