package com.example.variablecontract;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 *
 * @author andy
 * @date 2016年8月5日
 */
public class Test3 {

    public static void main(String[] args) {
        try {
            readAndWriterTest3();
        } catch (IOException e) {
            e.printStackTrace();
        }

//        try {
//            readAndWriterTest4();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
    }

    /**
     * doc文件读取简单示例
     * @throws IOException
     */
    public static void readAndWriterTest3() throws IOException {
        //文件件路径 C:\Users\Administrator\Desktop\合同文档
        File file = new File("C:\\Users\\Administrator\\Desktop\\合同文档\\aa.doc");
        String str = "";
        try {
            FileInputStream fis = new FileInputStream(file);
            HWPFDocument doc = new HWPFDocument(fis);
            String doc1 = doc.getDocumentText();
            System.out.println(doc1);
            StringBuilder doc2 = doc.getText();
            System.out.println(doc2);
            Range rang = doc.getRange();
            String doc3 = rang.text();
            System.out.println(doc3);
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * docx文件读取简单示例
     * @throws IOException
     */
    public static void readAndWriterTest4() throws IOException {
        //文件件路径 C:\Users\Administrator\Desktop\合同文档
        File file = new File("C:\\Users\\Administrator\\Desktop\\合同文档\\aa.docx");
        String str = "";
        try {
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument xdoc = new XWPFDocument(fis);
            XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
            String doc1 = extractor.getText();
            System.out.println(doc1);
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}