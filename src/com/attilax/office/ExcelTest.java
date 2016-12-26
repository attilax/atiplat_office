package com.attilax.office;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import com.attilax.collection.listUtil;
import com.attilax.exception.ExUtil;
import com.attilax.json.AtiJson;
import com.attilax.lang.text.strUtil;
import com.google.common.collect.Lists;

public class ExcelTest {

	public static void main(String[] args) {
		String f = "c:\\功能表.xlsx";
		f="C:\\0Html\\gv_material.xlsx";
		List<Map> li = 	new excelUtil2007().toListMap(f);
		
	
		long longstart=System.currentTimeMillis();
		SerializePerson(li);
		long end=System.currentTimeMillis();
		System.out.println("--time(ms):"+(end-longstart));
		//  System.out.println(AtiJson.toJson(li));

		System.out.println("---f");
		
//		li=new ExcelTest().clear(li);
        //   li.
	}

	 private static void SerializePerson(Object o) 
	                {
	          // ObjectOutputStream 对象输出流，将Person对象存储到E盘的Person.txt文件中，完成对Person对象的序列化操作
	          
			ObjectOutputStream oo;
			try {
				FileOutputStream fileOutputStream = new FileOutputStream(
				         new File("c:/Person.txt"));
				oo = new ObjectOutputStream(fileOutputStream);
				  oo.writeObject(o);
		          System.out.println("Person对象序列化成功！");
		         oo.close();
			} catch (IOException e) {
				ExUtil.throwExV2(e);
			}
	      
	       }
	public List<Map> clear(List<Map> li) {
		List<Map> li_r=Lists.newLinkedList();
 
		for (Map map : li) {
			String cols = (String) map.get("字段");
			if(cols==null)
			{
				System.out.println("-wangring:cant ge fld from map");
				continue;
			}
			cols=cols.replace("+", " ");
			cols=cols.replace("\n", " ");
			cols=cols.replace("\r", " ");
			cols=cols.replace("\t", " ");
			cols = com.attilax.lang.text.strUtil.toEnChar(cols);
			cols = cols.replaceAll("\\(.*?\\)", "");
			List<String> cols_li = com.attilax.lang.text.strUtil.splitx2li(
					cols, new ArrayList<String>() {
						{
							this.add(",");
							this.add(" ");
						}
					});
			cols_li=listUtil.clrEmptyElement(cols_li);
			String s = listUtil.toString(cols_li);
			s=s.replace(":", "");
			map.put("字段", s );
			map.put("查询字段","名称,时间");
			if(map.get("操作")!=null)
			if(map.get("操作").toString().contains("上传"))
				map.put("操作","上传/下载");
			else
			     map.put("操作","编辑");
			li_r.add(map);
		}
		//li_r.remove(0);
		return li_r;
	}

}
