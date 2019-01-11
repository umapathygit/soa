package com.pgn.merge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TOC;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlink;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlOptions;
//import org.docx4j.wml.CTView;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtBlock;

public class MergeDocs {
	protected static List<XWPFHyperlink> hyperlinks;
	protected static List<XWPFParagraph> paragraphs;

	private static String merge(InputStream src1, InputStream src2,
			OutputStream dest, int i, String dPath) throws Exception

			{
		System.out.println("Entering into merge : " + i);

		hyperlinks = new ArrayList<XWPFHyperlink>();
		// comments = new ArrayList<XWPFComment>();
		paragraphs = new ArrayList<XWPFParagraph>();
		// tables= new ArrayList<XWPFTable>();

		OPCPackage src1Package = OPCPackage.open(src1);
		OPCPackage src2Package = OPCPackage.open(src2);
		XWPFDocument src1Document = new XWPFDocument(src1Package);
		CTBody src1Body = src1Document.getDocument().getBody();

		XWPFDocument src2Document = new XWPFDocument(src2Package);
		CTBody src2Body = src2Document.getDocument().getBody();

		ArrayList aa = src1Package.getParts();
		// System.out.println("No of parts : "+aa.size());

		appendBody(src1Body, src2Body);
		// if(i==5){
		src1Document.write(dest);
		// src1Package.close();
		// src2Package.close();
		// dest.close();//}
		// src1.close();
		File f = new File(dPath);
		if (i != 0)
		{
			f.deleteOnExit();
		}

		return dPath;
		}

	private static void appendBody(CTBody src, CTBody append) throws Exception {
		XmlOptions optionsOuter = new XmlOptions();
		optionsOuter.setSaveOuter();
		String appendString = append.xmlText(optionsOuter);
		String srcString = src.xmlText();
		// System.out.println("XML Data : "+srcString);
		String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
		String mainPart = srcString.substring(srcString.indexOf(">") + 1,
				srcString.lastIndexOf("<"));
		String sufix = srcString.substring(srcString.lastIndexOf("<"));
		String addPart = appendString.substring(appendString.indexOf(">") + 1,
				appendString.lastIndexOf("<"));
		CTBody makeBody = CTBody.Factory.parse(prefix + mainPart + addPart
				+ sufix);
		src.set(makeBody);

	}

	public static String doMerge(String[] docsArr, String strRootFolder) throws Exception 
	{

		
		File Root = new File(strRootFolder);
		
		String newFilePath = Root.getAbsolutePath() + "/Merged_Doc.docx";
		String nPath= Root.getAbsolutePath() + "/Merged_Doc";

		String[] listFolders = docsArr;

		for(int k = 0; k < listFolders.length; k++)
		{
			File source = new File(listFolders[k]);
			File desc = Root;
			//File m_fTemp = File.createTempFile("PgnWordFile", ".docx");
			System.out.println("File : "+listFolders[k]+" Copied ");
			try {
				FileUtils.copyFileToDirectory(source, desc);
			} catch (IOException e) {
				e.printStackTrace();
			}
			//m_fTemp.deleteOnExit();

		}

		File[] listOfFiles = Root.listFiles();


		FileInputStream fis;
		FileInputStream fis1;

		FileOutputStream outStream=new FileOutputStream(newFilePath); 
		//FileOutputStream outStream1 = new FileOutputStream(newFilePath1);

		String strFinalDoc = "";
		for (int j=0;j<listOfFiles.length-1;j++) {

			if(j==0)
			{ System.out.println(listOfFiles[j]);
			System.out.println(listOfFiles[j+1]);
			fis1=new FileInputStream(listOfFiles[j+1]);
			fis=new FileInputStream(listOfFiles[j]);
			strFinalDoc = merge(fis,fis1,outStream,j,newFilePath);
			}
			else
			{
				fis1=new FileInputStream(listOfFiles[j+1]);
				System.out.println(listOfFiles[j+1]);
				if(j>1)
				{
					fis=new FileInputStream(nPath+(j-1)+".docx");
					strFinalDoc = merge(fis,fis1,new FileOutputStream(nPath+j+".docx"),j,nPath+(j-1)+".docx");

				}
				else{
					fis=new FileInputStream(nPath+".docx");
					strFinalDoc = merge(fis,fis1,new FileOutputStream(nPath+j+".docx"),j,nPath+".docx");
				}

			}
		}

		return strFinalDoc;

	}

	public void test() throws IOException {
		File merge_Temp = null;
		merge_Temp = File.createTempFile("PgntempMerge", ".docx");
		merge_Temp.deleteOnExit();
	}

}
