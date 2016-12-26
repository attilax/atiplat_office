/**
 * 
 */
package com.attilax.office;

import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.File;
import java.io.IOException;
import java.net.ConnectException;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.icepdf.core.pobjects.Document;
import org.icepdf.core.pobjects.Page;
import org.icepdf.core.util.GraphicsRenderingHints;

import com.artofsolving.jodconverter.DefaultDocumentFormatRegistry;
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.attilax.img.ImgX;

/**
 * @author ASIMO
 *
 */
public class Excel2Tif {

	/**
	@author attilax 老哇的爪子
	@since   p1m e_v_47
	 
	 */
	public static void main(String[] args) {
//		doc2Imags("C:/1121测试反馈.docx","c:/officePdfOutput");
//		excel2Tif("c:/officePdfOutput");
		new Office2Tif().	excel2tif("c:/00/公司通讯录2014-1-15.xls","c:/officePdfOutput4excel2tif");
		System.out.println("--f");

	}
	
	
	 	/**
		@author attilax 老哇的爪子
		@since   p1q h_b_y
		 
		 */
	private static void excel2Tif(String string) {
		// TODO Auto-generated method stub
		
	}


	public static void doc2Imags(String docPath, String imgDirPath){
		 //C:/1121测试反馈.pdf
	    	String pdfPath =String.format("%s/%s.pdf", imgDirPath, FilenameUtils.getBaseName(docPath));
	    	
	    	System.out.println(pdfPath);
	    	try {
				doc2Pdf(docPath, pdfPath);
				if(!new File(pdfPath).exists())
				
					throw new RuntimeException("not exit pdf file::"+pdfPath);
				String pdfPath2 =String.format("%s/%s.pdf", imgDirPath, FilenameUtils.getBaseName(docPath)+"_2");
				FileUtils.copyFile( new File(pdfPath),new File( pdfPath2));
				pdf2Imgs(pdfPath, imgDirPath,"jpg");
				File pdf =  new File(pdfPath);
				if(pdf.isFile()){
					//pdf.delete();
				}

			} catch (ConnectException e) {
				e.printStackTrace();
				throw new RuntimeException(e);
			} catch (Exception e) {
				e.printStackTrace();	throw new RuntimeException(e);
			}
	    }
	
	   /**
     * 将pdf转换成图片
     *
     * @param pdfPath
     * @param imagePath
     * @return 返回转换后图片的名字
     * @throws Exception
     */
    public static List<String> pdf2Imgs(String pdfPath, String imgDirPath,  String exname) throws Exception {
        Document document = new Document();
        document.setFile(pdfPath);

        float scale = 5f;//放大倍数
        float rotation = 0f;//旋转度数

        List<String> imgNames = new ArrayList<String>();
        int pageNum = document.getNumberOfPages();
        File imgDir = new File(imgDirPath);
        if (!imgDir.exists()) {
            imgDir.mkdirs();
        }
        for (int i = 0; i < pageNum; i++) {
            BufferedImage image = (BufferedImage) document.getPageImage(i, GraphicsRenderingHints.SCREEN,
                    Page.BOUNDARY_CROPBOX, rotation, scale);
            RenderedImage rendImage = image;
            try {
              //  String exname = "jpg"; File.separator
				String filePath = imgDirPath + "/" + i +"."+ exname;
                File file = new File(filePath);
                ImageIO.write(rendImage, exname, file);
                imgNames.add(FilenameUtils.getName(filePath));
          
            image.flush();
            geneTif(filePath);
            } catch (IOException e) {
                e.printStackTrace();
                return null;
            }
        }
        document.dispose();
        return imgNames;
        
    }
    
    
    	/**
		@author attilax 老哇的爪子
		@since   p1m f_m_50
		 
		 */
	private static void geneTif(String filePath) {
	 ImgX.jpg2tif(filePath)	;
	}


	public static void doc2Pdf(String docPath, String pdfPath) throws ConnectException {
        File inputFile = new File(docPath);//预转文件
        File outputFile = new File(pdfPath);//pdf文件
        OpenOfficeConnection connection = new SocketOpenOfficeConnection(8100);
        connection.connect();//建立连接
        DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
//        DefaultDocumentFormatRegistry formatReg = new DefaultDocumentFormatRegistry();   
//        DocumentFormat txt = formatReg.getFormatByFileExtension("odt") ;//设定文件格式
//        DocumentFormat pdf = formatReg.getFormatByFileExtension("pdf") ;//设定文件格式
//        converter.convert(inputFile, txt, outputFile, pdf);//文件转换
        converter.convert(inputFile,   outputFile);
        connection.disconnect();//关闭连接
}

}
