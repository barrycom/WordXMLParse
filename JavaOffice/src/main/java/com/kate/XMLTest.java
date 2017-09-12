package com.kate;

import java.io.File;
import java.io.IOException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class XMLTest {

	public static void main(String[] args) {
		updateXMLFile();
	}

	
	
	private static void updateXMLFile() {
		File file = new File("D:\\JavaOfficeTest\\xml\\JavaWordTemplate.xml");
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder;
		try {
			dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(file);
			doc.getDocumentElement().normalize();
			
			updataElementValue(doc);
			
			doc.getDocumentElement().normalize();
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			DOMSource source = new DOMSource(doc);
			StreamResult result = new StreamResult(new File("D:\\JavaOfficeTest\\xml\\JavaWordTemplateUpdated.xml"));
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.transform(source, result);
			System.out.println("XML file updated successfully");
			
		} catch (ParserConfigurationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TransformerException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}



	private static void updataElementValue(Document doc) {
		NodeList contents = doc.getElementsByTagName("w:t");
		for(int i = 0; i < contents.getLength(); i++){
			Node name = contents.item(i).getFirstChild();
			String value = name.getNodeValue().trim();
			if(isLetter(value)){
				char[] cs=value.toCharArray();
		        cs[0]-=32;
		        value = String.valueOf(cs);
		        if(isFenJi(value.substring(value.length()-1))){
		        	value = value.substring(0, value.length()-1) + "_" + value.substring(value.length()-1);		        	
		        }
				value = "${" + value  + "}!\"\"";				
				name.setNodeValue(value);
			}
		}
	}

	private static boolean isFenJi(String value) {
		if(value.matches(".*[A-D]")){
			return true;
		}
		return false;
	}



	private static boolean isLetter(String value) {
		if(value.matches("[a-zA-Z]*")){
			return true;
		}
		return false;
	}

}
