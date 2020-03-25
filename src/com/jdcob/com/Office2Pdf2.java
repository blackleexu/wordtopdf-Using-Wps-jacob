package com.jdcob.com;

import java.awt.List;
import java.io.File;
import java.util.ArrayList;
import java.util.UUID;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
 
/**
 * @title ����jacob���� office�ĵ�תpdf
 * @author www44
 */
public class Office2Pdf2 {
 
	public static void main(String[] args) {
		word2PDF("D:/eclipse/workspace/Xdoc/file/format.docx", "D:/eclipse/workspace/Xdoc/file/"+ UUID.randomUUID().toString() +".pdf");
//		WordToXml("D:/eclipse/workspace/Xdoc/file/1111.docx", "D:/eclipse/workspace/Xdoc/file/"+ UUID.randomUUID().toString() +".xml");

//		excel2PDF("f:/2pdf/a_1234.xlsx", "f:/2pdf/a_excel.pdf");
//		ppt2PDF("f:/2pdf/a_1234.ppt", "f:/2pdf/a_ppt.pdf");
//		wordToHtml("D:/eclipse/workspace/Xdoc/file/tep.doc", "D:/eclipse/workspace/Xdoc/file/"+ UUID.randomUUID().toString() +".html");

	}
 
	private static final int wdFormatPDF = 17;
	private static final int xlTypePDF = 0;
	private static final int ppSaveAsPDF = 32;
	private static final int WORD_HTML = 8;
	private static final int WORD_XML = 11;
 
	/**
	 * wordתpdf
	 * 
	 * @param inputFile
	 * @param pdfFile
	 * @return
	 */
	public static boolean word2PDF(String inputFile, String pdfFile) {
		try {
			ComThread.InitSTA();
			// ��wordӦ�ó���
			ActiveXComponent app = new ActiveXComponent("KWPS.Application");
			// ����word���ɼ�
			app.setProperty("Visible", false);
			// ���word�����д򿪵��ĵ�,����Documents����
			Dispatch docs = app.getProperty("Documents").toDispatch();
			// ����Documents������Open�������ĵ��������ش򿪵��ĵ�����Document
			Dispatch doc = Dispatch.call(docs, "Open", inputFile, false, true)
					.toDispatch();
			// ����Document�����SaveAs���������ĵ�����Ϊpdf��ʽ
			// word����Ϊpdf��ʽ�ֵ꣬Ϊ17
			// Dispatch.call(doc, "SaveAs", pdfFile, wdFormatPDF);
			Dispatch.call(doc, "ExportAsFixedFormat", pdfFile, wdFormatPDF);
			// �ر��ĵ�
			Dispatch.call(doc, "Close", false);
			// �ر�wordӦ�ó���
			app.invoke("Quit", 0);
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {
			ComThread.Release();
		}
	}
 
	/**
	 * excelתpdf
	 * 
	 * @param inputFile
	 * @param pdfFile
	 * @return
	 */
	public static boolean excel2PDF(String inputFile, String pdfFile) {
		try {
			ComThread.InitSTA();
			ActiveXComponent app = new ActiveXComponent("Excel.Application");
			app.setProperty("DisplayAlerts", "False");
			app.setProperty("Visible", false);
			Dispatch excels = app.getProperty("Workbooks").toDispatch();
			Dispatch excel = Dispatch.call(excels, "Open", inputFile, false,
					true).toDispatch();
			Dispatch.call(excel, "ExportAsFixedFormat", xlTypePDF, pdfFile);
			Dispatch.call(excel, "Close", false);
			app.invoke("Quit");
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {
 
			ComThread.Release();
		}
 
	}
	/**
	 * pptתpdf
	 * 
	 * @param inputFile
	 * @param pdfFile
	 * @return
	 */
	public static boolean ppt2PDF(String inputFile, String pdfFile) {
		try {
			ComThread.InitSTA();
			ActiveXComponent app = new ActiveXComponent(
					"PowerPoint.Application");
			// app.setProperty("Visible", msofalse);
			Dispatch ppts = app.getProperty("Presentations").toDispatch();
			Dispatch ppt = Dispatch.call(ppts, "Open", inputFile, 
					true,// ReadOnly
					true,// Untitledָ���ļ��Ƿ��б���
					false// WithWindowָ���ļ��Ƿ�ɼ�
					).toDispatch();
			Dispatch.call(ppt, "SaveAs", pdfFile, ppSaveAsPDF);
			Dispatch.call(ppt, "Close");
			app.invoke("Quit");
			return true;
		} catch (Exception e) {
			return false;
		} finally {
			ComThread.Release();
		}
	}
	

	/**
	 * WORDתHTML
	 * 
	 * @param docfile  WORD�ļ�ȫ·��
	 * @param htmlfile ת����HTML���·��
	 */
	public static void wordToHtml(String docfile, String htmlfile) {
		// ����wordӦ�ó���(Microsoft Office Word 2003)
		ActiveXComponent app = new ActiveXComponent("KWPS.Application");
		System.out.println("*****����ת��...*****");
		try {
			// ����wordӦ�ó��򲻿ɼ�
			app.setProperty("Visible", new Variant(false));
			// documents��ʾword����������ĵ����ڣ���word�Ƕ��ĵ�Ӧ�ó���
			Dispatch docs = app.getProperty("Documents").toDispatch();
			// ��Ҫת����word�ļ�
			Dispatch doc = Dispatch.invoke(docs, "Open", Dispatch.Method,
					new Object[] { docfile, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
			// ��Ϊhtml��ʽ���浽��ʱ�ļ�
			Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { htmlfile, new Variant(WORD_HTML) },
					new int[1]);
			// �ر�word�ļ�
			Dispatch.call(doc, "Close", new Variant(false));
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			//�ر�wordӦ�ó���
			app.invoke("Quit", new Variant[] {});
		}
		System.out.println("*****ת�����********");
	}
	/**
	 * 
	 * @param docfile
	 * @param xmlfile
	 * @return
	 */
	public static void WordToXml(String docfile, String xmlfile) {
        try {
            
            ActiveXComponent app = new ActiveXComponent( "KWPS.Application"); //����word
            String inFile = docfile; //ָ��Ҫ���_��word�ļ�
           
            app.setProperty("Visible", new Variant(false)); //��false�r�O��word����Ҋ����true�r�ǿ�ҊҪ��Ȼ������Word���_�ļ����^��
            Dispatch docs = app.getProperty("Documents").toDispatch();
            //���_��݋��
            Dispatch doc = Dispatch.invoke(docs, "Open", Dispatch.Method, new Object[] {inFile, new Variant(false), new Variant(true)} , new int[1]).toDispatch(); //���_word�ęn
            Dispatch.call(doc, "SaveAs", xmlfile, WORD_XML);//xml�ļ���ʽ��11
            Dispatch.call(doc, "Close", false);
            app.invoke("Quit",0);
       }catch (Exception e) {
          e.printStackTrace();

       }
	}
	
	 /** �ϲ����wordΪһ��Word

	 * @param srcdocs

	 * @param destDoc

	 * @return

	 */

	public boolean mergeMultipleWord2Single(java.util.List srcdocs, String destDoc) {

	    //1.У��

	    if (srcdocs.size() == 0 || srcdocs == null) {

	        return false;

	    }

	    System.out.println("���� Word...");

	    long start = System.currentTimeMillis();

	    //2.�ж�

	    ActiveXComponent app = null;

	    Object doc = null;

	    try {
	        app = new ActiveXComponent("KWPS.Application");  // ����kwps�ķ�ʽ,����͵���ActiveX�ؼ��й�

	        // 2.1.����word���ɼ�

	        app.setProperty("Visible", new Variant(false));

	        //���Documents����

	        Object docs = app.getProperty("Documents").toDispatch();

	        // 2.2.�򿪵�һ���ļ�

	        doc = Dispatch.invoke(

	                (Dispatch) docs, // ����1:����Ŀ��

	                "Open",  // ����2

	                Dispatch.Method,  //����3

	                new Object[]{(String) srcdocs.get(0), new Variant(false), new Variant(true)},  // ����4

	                new int[3]  //����5

	        ).toDispatch();

	        // 2.3.׷�Ӻ����ļ�

	        for (int i = 1; i < srcdocs.size(); i++) {

	            Dispatch.invoke(

	                    app.getProperty("Selection").toDispatch(),  //����1

	                    "insertFile",

	                    Dispatch.Method,

	                    new Object[]{(String) srcdocs.get(i), "", new Variant(false), new Variant(false), new Variant(false)}, // ����4

	                    new int[3]  // ����5

	            );

	        }

	        //2.4.���Ŀ��word����,��ɾ��

	        File tofile = new File(destDoc);

	        if (tofile.exists()) {  // Ŀ��pdf����,��ɾ��,ǰ��δʹ��

	            tofile.delete();

	        }

	        // 2.5.����Ϊ�µ�word

	        Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method, new Object[]{destDoc, new Variant(1)}, new int[3]);

	        Variant f = new Variant(false);

	        Dispatch.call((Dispatch) doc, "Close", f);  // ��close���Ը�ֵ��f

	        long end = System.currentTimeMillis();

	        System.out.println("ת�����..��ʱ��" + (end - start) + "ms.");

	    } catch (Exception e) {

	        System.out.println("========�ϲ�Word�ļ�ʧ�ܣ�" + e.getMessage());

	        //throw new RuntimeException("========�ϲ�Word�ļ�ʧ�ܣ�" + e);

	        return false;

	    } finally {

	        System.out.println("�ر��ĵ�");

	        if (app != null) {

	            app.invoke("Quit", new Variant[]{});

	        }

	    }

	    // ���û����仰,winword.exe���̽�����ر�

	    ComThread.Release();

	    return true;

	}
	
	/*
	 * wordתpdf

	 */

	public void mergeMultipleWord2SingleTest() {

	    String descDoc = "D:/eclipse/workspace/Xdoc/file/1111.docx";

	    String descPdf = "D:/eclipse/workspace/Xdoc/file/merge.docx";

	    ArrayList<String> srcDocs = new ArrayList<String>();

	    srcDocs.add("D:/Javaͨ��jacob����wps��¼1.doc");

	    srcDocs.add("D:/Javaͨ��jacob����wps��¼2.doc");

	    boolean mergeResult = mergeMultipleWord2Single(srcDocs, descDoc);

	    if (mergeResult) {

	        System.out.println("�ϲ��ɹ�!");

	        word2PDF(descDoc, descPdf);

	    } else {

	        System.out.println("�ϲ�ʧ��!");

	    }

	}
	
}
