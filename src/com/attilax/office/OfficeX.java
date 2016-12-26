/**
 * 
 */
package com.attilax.office;

import Dotnet.System.Windows.Forms.MessageBox;
import System.Diagnostics.ProcessStartInfo;

/**
 * @author ASIMO
 *
 */
public class OfficeX {

	/**
	@author attilax 老哇的爪子
	@since   p19 9_d_58
	 
	 */
	public static void main(String[] args) {
		ConvertToSwf("C:\\Users\\ASIMO\\Documents\\atitit.系统安装维护文档.doc" , "c:\\r.swf","C:\\Program Files (x86)\\Print2Flash3\\p2fServer.exe");

	}
	
	
	  /// <summary>
    /// 将Office文件转换为Swf(需要print2flash软件)
    /// </summary>
    /// <param name="srcFile">需要转换的文件地址和文件名</param>
    /// <param name="swfFile">转换后文件的目标地址</param>
    /// <param name="strProcessPath">print2flash软件的地址(D:\Program Files\Print2Flash3\p2fServer.exe)</param>
    public static void ConvertToSwf(String srcFile, String swfFile,String strProcessPath)
    {
    	//
    	//
        try
        {
            //设置进程物理路径
        	String flashPrinter = strProcessPath;
            //创建进程实例
            ProcessStartInfo StartInfo = new ProcessStartInfo();
            //设置需要启动进程的程序或者文档
            StartInfo.FileName = flashPrinter;
            //设置需要启动进程的参数
            StartInfo.Arguments = srcFile + " " + swfFile;
            //是否启动该进程
            StartInfo.CreateNoWindow = true;
            System.out.println( StartInfo.FileName  + StartInfo.Arguments );
            //创建可以停止和启动进程的实例
            // if drktl use system packete then ,conflic with java lang system class
            //todox p18
            Donet.System.Diagnostics.Process process = new  Donet.System.Diagnostics.Process();
            //设置进程
            process.StartInfo = StartInfo;
            //启用进程
            Boolean isStart = process.Start();
            System.out.println("retMsg::"+process.retMsg);

            process.WaitForExit(1000 * 60 * 5);
        }
        catch (Exception es)
        {
            MessageBox.Show(es.getMessage());
        }

    }

}
