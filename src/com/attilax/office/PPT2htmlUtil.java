/**
 * 
 */
package com.attilax.office;

import java.io.File;

// qc import org.apache.commons.io.FilenameUtils;

import com.attilax.io.filex;

/**
 * @author ASIMO
 *
 */
public class PPT2htmlUtil {
	
	public static void main(String[] args) {
	
		String exe = "C:\\pdf2htmlEX-0.12-win32-static-with-poppler-data\\pdf2htmlEX.exe";
		
		
		String ppfFile = "C:/00/研发项目管理办法.ppt";
		String htmfile="sa3.html";
		
		ppt2html(exe, htmfile, ppfFile);
	System.out.println("--");
	}

	private static void ppt2html(String exe, String htmfile, String pptFile) {
		//String pdfPath =outDir+"/"+ FilenameUtils.getBaseName(excelFile)+".pdf";
		String pdffile= filex.changeExname(pptFile, "pdf");
		Office2Pdf.office2PDF( pptFile , pdffile);
		//Pdf2htmlEXUtil.
	//	String pdffile = "C:/word2htmlPPT/sa.pdf";
		Pdf2htmlEXUtil.pdf2html2(	exe,pdffile,htmfile);
	 
	}

}
