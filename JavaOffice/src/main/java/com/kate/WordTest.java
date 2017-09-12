package com.kate;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

public class WordTest {
	private Configuration configuration = null;
	
	public WordTest() {
		configuration = new Configuration();
		configuration.setDefaultEncoding("UTF-8");
	}
	
	public static void main(String[] args){
		WordTest wordTest = new WordTest();
		wordTest.createWord();
	}

	public void createWord() {
		Map<String, Object> dataMap = new HashMap<String, Object>();
		getData(dataMap);
		
		configuration.setClassForTemplateLoading(this.getClass(), "");//模板文件所在路径
		
		Template template = null;
		
		try{
			template = configuration.getTemplate("JavaWordTemplate.ftl");//获取模板文件
		} catch (IOException e){
			e.printStackTrace();
		}
		
		File outFile = new File("D:\\JavaOfficeTest\\" + Math.random()*10000 + ".doc");//导出文件
		
		Writer out = null;
		
		try{
			out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"UTF-8"));
		} catch (FileNotFoundException e1){
			e1.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		
		try{
			template.process(dataMap, out);//将填充数据填入模板文件并输出到目标文件 
		} catch (TemplateException e){
			e.printStackTrace();
		} catch (IOException e){
			e.printStackTrace();
		}
	}

	private void getData(Map<String, Object> dataMap) {
		dataMap.put("title", "标题");
		dataMap.put("year", "2017");
		dataMap.put("month", "09");
		dataMap.put("day", "02");
		//dataMap.put("id", "SA15226241");
		//dataMap.put("context", "潘凯特在测试JavaWord");
		dataMap.put("reviewer", "张津津");
		
		List<Map<String, Object>> list = new ArrayList<Map<String,Object>>();
		for(int i = 0; i < 10; i++){
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("id", "SA1522624" + i);
			map.put("context", "潘凯特在测试JavaWord" + i);
			list.add(map);
		}
		
		dataMap.put("list", list);
	}
}
