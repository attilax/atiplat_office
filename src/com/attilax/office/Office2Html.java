 

package com.attilax.office;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.ConnectException;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.attilax.core;
import com.attilax.lang.gbk����;
//import com.sun.star.uno.RuntimeException;

/**
 * ��Word�ĵ�ת����html�ַ����Ĺ�����
 * 
 * @author MZULE
 * 
 */
@gbk����
public class Office2Html {

    public static void main(String[] args) {
    //	NestableRuntimeException
//    System.out
    //     .println(convert(new File("C:/1121���Է���.docx"), "C:/word2html"));
  //  convert(new File("C:/00/��˾ͨѶ¼2014-1-15.xls"), "C:/word2html");
  //  convert(new File("C:/00/�����Ż��з�һ����֯�ܹ���֪ͨ��2015001��.pdf"), "C:/word2html");
    	Office2Html. convert(new File("C:/00/�з���Ŀ����취.ppt"), "C:/word2html");
 //   core.execMeth_Ays(runnable, threadName)
   
    //convert(new file)
    System.out.println("--f");
    }

    /**
     * ��word�ĵ�ת����html�ĵ�
     * 
     * @param docFile
     *                ��Ҫת����word�ĵ�
     * @param filepath
     *                ת��֮��html�Ĵ��·��
     * @return ת��֮���html�ļ�
     */
    public static File convert(File docFile, String filepath) {
    // ��������html���ļ�
    File htmlFile = new File(filepath + "/" + new Date().getTime()
        + ".html");
    // ����Openoffice����
    OpenOfficeConnection con = new SocketOpenOfficeConnection(8100);
    try {
        // ����
        con.connect();
    } catch (ConnectException e) {
        System.out.println("��ȡOpenOffice����ʧ��...");
        e.printStackTrace();
        throw new RuntimeException(e.getMessage());
    }
    // ����ת����
    DocumentConverter converter = new OpenOfficeDocumentConverter(con);
    // ת���ĵ���html
    converter.convert(docFile, htmlFile);
    // �ر�openoffice����
    con.disconnect();
    return htmlFile;
    }
    

    /**
     * ��wordת����html�ļ������һ�ȡhtml�ļ����롣
     * 
     * @param docFile
     *                ��Ҫת�����ĵ�
     * @param filepath
     *                �ĵ���ͼƬ�ı���λ��
     * @return ת���ɹ���html����
     */
    public static String toHtmlString(File docFile, String filepath) {
    // ת��word�ĵ�
    File htmlFile = convert(docFile, filepath);
    // ��ȡhtml�ļ���
    StringBuffer htmlSb = new StringBuffer();
    try {
        BufferedReader br = new BufferedReader(new InputStreamReader(
            new FileInputStream(htmlFile)));
        while (br.ready()) {
        htmlSb.append(br.readLine());
        }
        br.close();
        // ɾ����ʱ�ļ�
        htmlFile.delete();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }
    // HTML�ļ��ַ���
    String htmlStr = htmlSb.toString();
    // ���ؾ�������html�ı�
    return clearFormat(htmlStr, filepath);
    }

    /**
     * ���һЩ����Ҫ��html���
     * 
     * @param htmlStr
     *                ���и���html��ǵ�html���
     * @return ȥ���˲���Ҫhtml��ǵ����
     */
    protected static String clearFormat(String htmlStr, String docImgPath) {
    // ��ȡbody���ݵ�����
    String bodyReg = "<BODY .*</BODY>";
    Pattern bodyPattern = Pattern.compile(bodyReg);
    Matcher bodyMatcher = bodyPattern.matcher(htmlStr);
    if (bodyMatcher.find()) {
        // ��ȡBODY���ݣ���ת��BODY��ǩΪDIV
        htmlStr = bodyMatcher.group().replaceFirst("<BODY", "<DIV")
            .replaceAll("</BODY>", "</DIV>");
    }
    // ����ͼƬ��ַ
    htmlStr = htmlStr.replaceAll("<IMG SRC=\"", "<IMG SRC=\"" + docImgPath
        + "/");
    // ��<P></P>ת����</div></div>������ʽ
    // content = content.replaceAll("(<P)([^>]*>.*?)(<\\/P>)",
    // "<div$2</div>");
    // ��<P></P>ת����</div></div>��ɾ����ʽ
    htmlStr = htmlStr.replaceAll("(<P)([^>]*)(>.*?)(<\\/P>)", "<p$3</p>");
    // ɾ������Ҫ�ı�ǩ
    htmlStr = htmlStr
        .replaceAll(
            "<[/]?(font|FONT|span|SPAN|xml|XML|del|DEL|ins|INS|meta|META|[ovwxpOVWXP]:\\w+)[^>]*?>",
            "");
    // ɾ������Ҫ������
    htmlStr = htmlStr
        .replaceAll(
            "<([^>]*)(?:lang|LANG|class|CLASS|style|STYLE|size|SIZE|face|FACE|[ovwxpOVWXP]:\\w+)=(?:'[^']*'|\"\"[^\"\"]*\"\"|[^>]+)([^>]*)>",
            "<$1$2>");
    return htmlStr;
    }

}