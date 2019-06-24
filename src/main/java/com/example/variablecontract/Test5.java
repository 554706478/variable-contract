package com.example.variablecontract;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Description: 测试docx文档成功22个参数变量置换成功
 * Author: v-wuchengs
 * Date: 2019-06-19
 */
public class Test5 {

    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub
//        String filepathString = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\aa.docx";
//        String destpathString = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\bb.docx";
//        Map<String, String> map = new HashMap<String, String>();
//        map.put("${PARTY_A}", "张三");
//        map.put("${MONEY}", "李四");
//        map.put("${FEE}", "王五");
//        OPCPackage pack = POIXMLDocument.openPackage(filepathString);
//        XWPFDocument document = new XWPFDocument(pack);

        //合同样例二
        String filepathString = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\合同样本.docx";
        String destpathString = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\合同样本bb.docx";
        Map<String, String> map = new HashMap<String, String>();
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
        OPCPackage pack = POIXMLDocument.openPackage(filepathString);
        XWPFDocument document = new XWPFDocument(pack);


        /**
         * 对段落中的标记进行替换
         */
        List<XWPFParagraph> parasList = document.getParagraphs();
        replaceInAllParagraphs(parasList, map);

        /**
         * 对表格中的标记进行替换
         */
        List<XWPFTable> tables = document.getTables();
        replaceInTables(tables, map);
        FileOutputStream outStream = null;
        try {
            outStream = new FileOutputStream(destpathString);
            document.write(outStream);
            outStream.flush();
            outStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }




//    /**
//     * 替换段落中的字符串
//     *
//     * @param xwpfParagraph
//     * @param oldString
//     * @param newString
//     */
//    public static void replaceInParagraph(XWPFParagraph xwpfParagraph, String oldString, String newString) {
//        Map<String, Integer> pos_map = findSubRunPosInParagraph(xwpfParagraph, oldString);
//        if (pos_map != null) {
//            System.out.println("start_pos:" + pos_map.get("start_pos"));
//            System.out.println("end_pos:" + pos_map.get("end_pos"));
//
//            List<XWPFRun> runs = xwpfParagraph.getRuns();
//            XWPFRun modelRun = runs.get(pos_map.get("end_pos"));
//            XWPFRun xwpfRun = xwpfParagraph.insertNewRun(pos_map.get("end_pos") + 1);
//            xwpfRun.setText(newString);
//            System.out.println("字体大小：" + modelRun.getFontSize());
//            if (modelRun.getFontSize() != -1) xwpfRun.setFontSize(modelRun.getFontSize());//默认值是五号字体，但五号字体getFontSize()时，返回-1
//            xwpfRun.setFontFamily(modelRun.getFontFamily());
//            for (int i = pos_map.get("end_pos"); i >= pos_map.get("start_pos"); i--) {
//                System.out.println("remove run pos in :" + i);
//                xwpfParagraph.removeRun(i);
//            }
//        }
//    }
//
//    /**
//     * 找到段落中子串的起始XWPFRun下标和终止XWPFRun的下标
//     *
//     * @param xwpfParagraph
//     * @param substring
//     * @return
//     */
//    public static Map<String, Integer> findSubRunPosInParagraph(XWPFParagraph xwpfParagraph, String substring) {
//        List<XWPFRun> runs = xwpfParagraph.getRuns();
//        int start_pos = 0;
//        int end_pos = 0;
//        String subtemp = "";
//        for (int i = 0; i < runs.size(); i++) {
//            subtemp = "";
//            start_pos = i;
//            for (int j = i; j < runs.size(); j++) {
//                if (runs.get(j).getText(runs.get(j).getTextPosition()) == null) continue;
//                subtemp += runs.get(j).getText(runs.get(j).getTextPosition());
//                if (subtemp.equals(substring)) {
//                    end_pos = j;
//                    Map<String, Integer> map = new HashMap<>();
//                    map.put("start_pos", start_pos);
//                    map.put("end_pos", end_pos);
//                    return map;
//                }
//            }
//        }
//        return null;
//    }









    /**
     * 替换所有段落中的标记
     *
     * @param xwpfParagraphList
     * @param params
     */
    public static void replaceInAllParagraphs(List<XWPFParagraph> xwpfParagraphList, Map<String, String> params) {
        for (XWPFParagraph paragraph : xwpfParagraphList) {
            if (paragraph.getText() == null || paragraph.getText().equals("")) continue;
            for (String key : params.keySet()) {
                if (paragraph.getText().contains(key)) {
//                    text = paragraph.getText().replace(key, params.get(key));
                    replaceInParagraph(paragraph, key, params.get(key));
                }
            }
        }
    }

    /**
     * 替换段落中的字符串
     *
     * @param xwpfParagraph
     * @param oldString
     * @param newString
     */
    public static void replaceInParagraph(XWPFParagraph xwpfParagraph, String oldString, String newString) {
        Map<String, Integer> pos_map = findSubRunPosInParagraph(xwpfParagraph, oldString);
        if (pos_map != null) {
            System.out.println("start_pos:" + pos_map.get("start_pos"));
            System.out.println("end_pos:" + pos_map.get("end_pos"));

            List<XWPFRun> runs = xwpfParagraph.getRuns();
            XWPFRun modelRun = runs.get(pos_map.get("end_pos"));
            XWPFRun xwpfRun = xwpfParagraph.insertNewRun(pos_map.get("end_pos") + 1);
            xwpfRun.setText(newString);
            System.out.println("字体大小：" + modelRun.getFontSize());
            if (modelRun.getFontSize() != -1) xwpfRun.setFontSize(modelRun.getFontSize());//默认值是五号字体，但五号字体getFontSize()时，返回-1
            xwpfRun.setFontFamily(modelRun.getFontFamily());
            for (int i = pos_map.get("end_pos"); i >= pos_map.get("start_pos"); i--) {
                System.out.println("remove run pos in :" + i);
                xwpfParagraph.removeRun(i);
            }
        }
    }


    /**
     * 找到段落中子串的起始XWPFRun下标和终止XWPFRun的下标
     *
     * @param xwpfParagraph
     * @param substring
     * @return
     */
    public static Map<String, Integer> findSubRunPosInParagraph(XWPFParagraph xwpfParagraph, String substring) {

        List<XWPFRun> runs = xwpfParagraph.getRuns();
        int start_pos = 0;
        int end_pos = 0;
        String subtemp = "";
        for (int i = 0; i < runs.size(); i++) {
            subtemp = "";
            start_pos = i;
            for (int j = i; j < runs.size(); j++) {
                if (runs.get(j).getText(runs.get(j).getTextPosition()) == null) continue;
                subtemp += runs.get(j).getText(runs.get(j).getTextPosition());
                if (subtemp.equals(substring)) {
                    end_pos = j;
                    Map<String, Integer> map = new HashMap<>();
                    map.put("start_pos", start_pos);
                    map.put("end_pos", end_pos);
                    return map;
                }
            }
        }
        return null;
    }





    /**
     * 替换所有的表格
     *
     * @param xwpfTableList
     * @param params
     */
    public static void replaceInTables(List<XWPFTable> xwpfTableList, Map<String, String> params) {
        for (XWPFTable table : xwpfTableList) {
            replaceInTable(table, params);

        }
    }

    /**
     * 替换一个表格中的所有行
     *
     * @param xwpfTable
     * @param params
     */
    public static void replaceInTable(XWPFTable xwpfTable, Map<String, String> params) {
        List<XWPFTableRow> rows = xwpfTable.getRows();
        replaceInRows(rows, params);
    }


    /**
     * 替换表格中的一行
     *
     * @param rows
     * @param params
     */
    public static void replaceInRows(List<XWPFTableRow> rows, Map<String, String> params) {
        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
            replaceInCells(row.getTableCells(), params);
        }
    }

    /**
     * 替换一行中所有的单元格
     *
     * @param xwpfTableCellList
     * @param params
     */
    public static void replaceInCells(List<XWPFTableCell> xwpfTableCellList, Map<String, String> params) {
        for (XWPFTableCell cell : xwpfTableCellList) {
            replaceInCell(cell, params);
        }
    }

    /**
     * 替换表格中每一行中的每一个单元格中的所有段落
     *
     * @param cell
     * @param params
     */
    public static void replaceInCell(XWPFTableCell cell, Map<String, String> params) {
        List<XWPFParagraph> cellParagraphs = cell.getParagraphs();
        replaceInAllParagraphs(cellParagraphs, params);
    }














}
