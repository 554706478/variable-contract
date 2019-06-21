package com.example.variablecontract;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.util.*;

/**
 * Description:
 * Author: v-wuchengs
 * Date: 2019-06-19
 */
public class Test4 {

    public static void main(String[] args) throws Exception {
        Map<String, String> map = new HashMap<String, String>();
        //合同样例一
//        map.put("${PARTY_A}", "张三");
//        map.put("${MONEY}", "李四");
//        map.put("${FEE}", "王五");
//        String srcPath = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\aa.docx";
//        String destPath = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\bb.docx";
//        searchAndReplace(srcPath, destPath, map);

        //合同样例二
        map.put("${VariableParameter1}", "变量值1");
        map.put("${VariableParameter2}", "变量值2");
        map.put("${VariableParameter3}", "变量值3");
        map.put("${VariableParameter4}", "变量值4");
        map.put("${VariableParameter5}", "变量值5");
        map.put("${VariableParameter6}", "变量值6");
        map.put("${VariableParameter7}", "变量值7");
        map.put("${VariableParameter8}", "变量值8");
        map.put("${VariableParameter9}", "变量值9");
        map.put("${VariableParameter10}", "变量值10");
        map.put("${VariableParameter11}", "变量值11");
        map.put("${VariableParameter12}", "变量值12");
        map.put("${VariableParameter13}", "变量值13");
        map.put("${VariableParameter14}", "变量值14");
        map.put("${VariableParameter15}", "变量值15");
        map.put("${VariableParameter16}", "变量值16");
        map.put("${VariableParameter17}", "变量值17");
        map.put("${VariableParameter18}", "变量值18");
        map.put("${VariableParameter19}", "变量值19");
        map.put("${VariableParameter20}", "变量值20");
        map.put("${VariableParameter21}", "变量值表格1");
        map.put("${VariableParameter22}", "变量值表格2");
        String srcPath = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\合同样本.docx";
        String destPath = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\合同样本bb.docx";
        searchAndReplace(srcPath, destPath, map);
    }

    public static void searchAndReplace(String srcPath, String destPath, Map<String, String> map) {
        try {
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(srcPath));
//            XWPFWordExtractor extractor = new XWPFWordExtractor(document);
//            String text = extractor.getText();
//            for (Map.Entry<String, String> entry : map.entrySet()) {
//                text = text.replace(entry.getKey(), entry.getValue());
//            }
            /**
             * 替换段落中的指定文字
             */
            Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
            while (itPara.hasNext()) {
                XWPFParagraph paragraph = (XWPFParagraph) itPara.next();
                Set<String> set = map.keySet();
                Iterator<String> iterator = set.iterator();
                while (iterator.hasNext()) {
                    String key = iterator.next().trim();
                    List<XWPFRun> run = paragraph.getRuns();
                    System.out.println(run.toString());
                    int runSize = run.size();
                    for (int i = 0; i < runSize; i++) {
                        String text = run.get(i).getText(0);
                        System.out.println("++++++text++++++:" + text);
                        for (Map.Entry<String, String> e : map.entrySet()) {
                            if (text != null && text.contains(e.getKey())) {
                                text = text.replace(e.getKey(), e.getValue());
                                System.out.println("++++++text222222222++++++:" + text);
                                run.get(i).setText(text, 0);
                            }
                        }
                    }
                }
            }

            /**
             * 替换表格中的指定文字
             */
            Iterator<XWPFTable> itTable = document.getTablesIterator();
            while (itTable.hasNext()) {
                XWPFTable table = (XWPFTable) itTable.next();
                int count = table.getNumberOfRows();
                for (int i = 0; i < count; i++) {
                    XWPFTableRow row = table.getRow(i);
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (XWPFTableCell cell : cells) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            for (XWPFRun r : p.getRuns()) {
                                String text = r.getText(0);
                                for (Map.Entry<String, String> e : map.entrySet()) {
                                    if (text != null && text.contains(e.getKey())) {
                                        text = text.replace(e.getKey(), e.getValue());
                                        r.setText(text, 0);
                                    }
                                }
                            }
                        }

                    }
                }
            }
            FileOutputStream outStream = null;
            outStream = new FileOutputStream(destPath);
            document.write(outStream);
            outStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

}
