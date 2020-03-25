package com.jdcob.com;

import com.jacob.activeX.ActiveXComponent;

import com.jacob.com.ComThread;

import com.jacob.com.Dispatch;

import com.jacob.com.Variant;

import java.io.File;


public class MyWord2PDF {
	private static final int wdFormatPDF = 17; // PDF ��ʽ

	public void wordToPDF(String sfileName, String toFileName) {

	    System.out.println("���� Word...");

	    long start = System.currentTimeMillis();

	    ActiveXComponent app = null;

	    Dispatch doc = null;

	    try {        app = new ActiveXComponent("KWPS.Application");  // ����kwps�ķ�ʽ

        app.setProperty("Visible", new Variant(false));

        Dispatch docs = app.getProperty("Documents").toDispatch();

        doc = Dispatch.call(docs, "Open", sfileName).toDispatch();

        System.out.println("���ĵ�..." + sfileName);

        System.out.println("ת���ĵ��� PDF..." + toFileName);

        File tofile = new File(toFileName);

        if (tofile.exists()) { // Ŀ��pdf����,��ɾ��,ǰ��δʹ��

            tofile.delete();

        }

        Dispatch.call(doc, "SaveAs", toFileName, // FileName

                wdFormatPDF);

        long end = System.currentTimeMillis();

        System.out.println("ת�����..��ʱ��" + (end - start) + "ms.");

    } catch (Exception e) {

        System.out.println("========Error:�ĵ�ת��ʧ�ܣ�" + e.getMessage());

    } finally {

        Dispatch.call(doc, "Close", false);

        System.out.println("�ر��ĵ�");

        if (app != null)

            app.invoke("Quit", new Variant[]{});

    }

    // ���û����仰,winword.exe���̽�����ر�

    ComThread.Release();

	}
}