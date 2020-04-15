package com.edu.nelson;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;
import java.util.Scanner;

/**
 * @auther 1301913120@qq.com
 * @create 2020-04-15 10:20
 * @todo
 */
public class ChangeToDocxUtil {

    public static void main(String[] args) {
        System.out.println("Please enter the absolute path to the folder you want to convert:");
        Scanner scanner = new Scanner(System.in);
        String path = scanner.nextLine();
//        String path = "C:\\Users\\13019\\Desktop\\temp";
        change(path);
        System.out.println("The file has been converted successfully.");
    }

    public static void change(String path){
        File f = new File(path);
        File[] files = f.listFiles();
        for (File file : files) {
            if (file.isDirectory()){
                change(file.getPath());
            }else {
                String inputFile = file.getPath();
                String outputFile;
                if (file.getName().endsWith(".doc")){
                    outputFile = inputFile.replace(".doc",".docx");
                }else if (file.getName().endsWith(".rtf")){
                    outputFile = inputFile.replace(".rtf", ".docx");
                }else {
                    continue;
                }
                word2PDF(inputFile, outputFile);
            }
        }

    }

    private static int word2PDF(String inputFile, String pdfFile) {
        try {
            // 打开Word应用程序
            ActiveXComponent app = new ActiveXComponent("KWPS.Application");
            System.out.println("Parsing......");
            long date = System.currentTimeMillis();
            // 设置Word不可见
            app.setProperty("Visible", new Variant(false));
            // 禁用宏
            app.setProperty("AutomationSecurity", new Variant(3));
            // 获得Word中所有打开的文档，返回documents对象
            Dispatch docs = app.getProperty("Documents").toDispatch();
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            Dispatch doc = Dispatch.call(docs, "Open", inputFile, false, true).toDispatch();
            /***
             *
             * 调用Document对象的SaveAs方法，将文档保存为pdf格式
             *
             * Dispatch.call(doc, "SaveAs", pdfFile, wdFormatPDF word保存为pdf格式宏，值为17 )
             *
             */
            Dispatch.call(doc, "SaveAs", pdfFile, 12);// word保存为pdf格式宏，值为17
//            System.out.println(doc);
            // 关闭文档
            long date2 = System.currentTimeMillis();
            int time = (int) ((date2 - date) / 1000);

            Dispatch.call(doc, "Close", false);
            // 关闭Word应用程序
            app.invoke("Quit", 0);
            return time;
        } catch (Exception e) {
            // TODO: handle exception
            e.printStackTrace();
            return -1;
        }

    }
}
