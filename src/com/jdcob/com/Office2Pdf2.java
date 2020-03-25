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
 * @title 调用jacob服务 office文档转pdf
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
	 * word转pdf
	 * 
	 * @param inputFile
	 * @param pdfFile
	 * @return
	 */
	public static boolean word2PDF(String inputFile, String pdfFile) {
		try {
			ComThread.InitSTA();
			// 打开word应用程序
			ActiveXComponent app = new ActiveXComponent("KWPS.Application");
			// 设置word不可见
			app.setProperty("Visible", false);
			// 获得word中所有打开的文档,返回Documents对象
			Dispatch docs = app.getProperty("Documents").toDispatch();
			// 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
			Dispatch doc = Dispatch.call(docs, "Open", inputFile, false, true)
					.toDispatch();
			// 调用Document对象的SaveAs方法，将文档保存为pdf格式
			// word保存为pdf格式宏，值为17
			// Dispatch.call(doc, "SaveAs", pdfFile, wdFormatPDF);
			Dispatch.call(doc, "ExportAsFixedFormat", pdfFile, wdFormatPDF);
			// 关闭文档
			Dispatch.call(doc, "Close", false);
			// 关闭word应用程序
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
	 * excel转pdf
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
	 * ppt转pdf
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
					true,// Untitled指定文件是否有标题
					false// WithWindow指定文件是否可见
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
	 * WORD转HTML
	 * 
	 * @param docfile  WORD文件全路径
	 * @param htmlfile 转换后HTML存放路径
	 */
	public static void wordToHtml(String docfile, String htmlfile) {
		// 启动word应用程序(Microsoft Office Word 2003)
		ActiveXComponent app = new ActiveXComponent("KWPS.Application");
		System.out.println("*****正在转换...*****");
		try {
			// 设置word应用程序不可见
			app.setProperty("Visible", new Variant(false));
			// documents表示word程序的所有文档窗口，（word是多文档应用程序）
			Dispatch docs = app.getProperty("Documents").toDispatch();
			// 打开要转换的word文件
			Dispatch doc = Dispatch.invoke(docs, "Open", Dispatch.Method,
					new Object[] { docfile, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
			// 作为html格式保存到临时文件
			Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { htmlfile, new Variant(WORD_HTML) },
					new int[1]);
			// 关闭word文件
			Dispatch.call(doc, "Close", new Variant(false));
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			//关闭word应用程序
			app.invoke("Quit", new Variant[] {});
		}
		System.out.println("*****转换完毕********");
	}
	/**
	 * 
	 * @param docfile
	 * @param xmlfile
	 * @return
	 */
	public static void WordToXml(String docfile, String xmlfile) {
        try {
            
            ActiveXComponent app = new ActiveXComponent( "KWPS.Application"); //啟動word
            String inFile = docfile; //指定要打開的word文件
           
            app.setProperty("Visible", new Variant(false)); //為false時設置word不可見，為true時是可見要不然看不到Word打開文件的過程
            Dispatch docs = app.getProperty("Documents").toDispatch();
            //打開編輯囂
            Dispatch doc = Dispatch.invoke(docs, "Open", Dispatch.Method, new Object[] {inFile, new Variant(false), new Variant(true)} , new int[1]).toDispatch(); //打開word文檔
            Dispatch.call(doc, "SaveAs", xmlfile, WORD_XML);//xml文件格式宏11
            Dispatch.call(doc, "Close", false);
            app.invoke("Quit",0);
       }catch (Exception e) {
          e.printStackTrace();

       }
	}
	
	 /** 合并多个word为一个Word

	 * @param srcdocs

	 * @param destDoc

	 * @return

	 */

	public boolean mergeMultipleWord2Single(java.util.List srcdocs, String destDoc) {

	    //1.校验

	    if (srcdocs.size() == 0 || srcdocs == null) {

	        return false;

	    }

	    System.out.println("启动 Word...");

	    long start = System.currentTimeMillis();

	    //2.判断

	    ActiveXComponent app = null;

	    Object doc = null;

	    try {
	        app = new ActiveXComponent("KWPS.Application");  // 基于kwps的方式,具体和调用ActiveX控件有关

	        // 2.1.设置word不可见

	        app.setProperty("Visible", new Variant(false));

	        //获得Documents对象

	        Object docs = app.getProperty("Documents").toDispatch();

	        // 2.2.打开第一个文件

	        doc = Dispatch.invoke(

	                (Dispatch) docs, // 参数1:调用目标

	                "Open",  // 参数2

	                Dispatch.Method,  //参数3

	                new Object[]{(String) srcdocs.get(0), new Variant(false), new Variant(true)},  // 参数4

	                new int[3]  //参数5

	        ).toDispatch();

	        // 2.3.追加后续文件

	        for (int i = 1; i < srcdocs.size(); i++) {

	            Dispatch.invoke(

	                    app.getProperty("Selection").toDispatch(),  //参数1

	                    "insertFile",

	                    Dispatch.Method,

	                    new Object[]{(String) srcdocs.get(i), "", new Variant(false), new Variant(false), new Variant(false)}, // 参数4

	                    new int[3]  // 参数5

	            );

	        }

	        //2.4.如果目的word存在,则删除

	        File tofile = new File(destDoc);

	        if (tofile.exists()) {  // 目标pdf存在,则删除,前提未使用

	            tofile.delete();

	        }

	        // 2.5.保存为新的word

	        Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method, new Object[]{destDoc, new Variant(1)}, new int[3]);

	        Variant f = new Variant(false);

	        Dispatch.call((Dispatch) doc, "Close", f);  // 把close属性赋值个f

	        long end = System.currentTimeMillis();

	        System.out.println("转换完成..用时：" + (end - start) + "ms.");

	    } catch (Exception e) {

	        System.out.println("========合并Word文件失败：" + e.getMessage());

	        //throw new RuntimeException("========合并Word文件失败：" + e);

	        return false;

	    } finally {

	        System.out.println("关闭文档");

	        if (app != null) {

	            app.invoke("Quit", new Variant[]{});

	        }

	    }

	    // 如果没有这句话,winword.exe进程将不会关闭

	    ComThread.Release();

	    return true;

	}
	
	/*
	 * word转pdf

	 */

	public void mergeMultipleWord2SingleTest() {

	    String descDoc = "D:/eclipse/workspace/Xdoc/file/1111.docx";

	    String descPdf = "D:/eclipse/workspace/Xdoc/file/merge.docx";

	    ArrayList<String> srcDocs = new ArrayList<String>();

	    srcDocs.add("D:/Java通过jacob操作wps记录1.doc");

	    srcDocs.add("D:/Java通过jacob操作wps记录2.doc");

	    boolean mergeResult = mergeMultipleWord2Single(srcDocs, descDoc);

	    if (mergeResult) {

	        System.out.println("合并成功!");

	        word2PDF(descDoc, descPdf);

	    } else {

	        System.out.println("合并失败!");

	    }

	}
	
}
