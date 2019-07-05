package com.example.variablecontract;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * Description:
 * Author: v-wuchengs
 * Date: 2019-06-19
 */
public class Test1 {

    public static void main(String[] args) {
        Map<String, String> map = new HashMap<String, String>();
        System.out.println(map);
        //合同样例一
//        map.put("${PARTY_A}", "张三");
//        map.put("${MONEY}", "100");
//        map.put("${FEE}", "5%");
//        String srcPath = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\aa.doc";
//        readwriteWord(srcPath, map);
//        System.out.println("---------------新合同生成完毕---------------");

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
        map.put("${VariableParameter21}", "变量值姓名甲");
        map.put("${VariableParameter22}", "变量值姓名乙");
        String srcPath = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\合同样本.doc";
        long startMillis = System.currentTimeMillis();
        System.out.println("开始执行前时间毫秒数："+startMillis);
        readwriteWord(srcPath, map);
        long endMillis = System.currentTimeMillis();
        System.out.println("执行完毕后时间毫秒数："+endMillis);
        System.out.println("合同参数置换用时毫秒数："+(endMillis-startMillis));
        System.out.println("---------------新合同生成完毕---------------");
    }

    /**
     * 实现对word读取和修改操作
     *
     * @param filePath word模板路径和名称
     * @param map      待填充的数据，从数据库读取
     */
    public static void readwriteWord(String filePath, Map<String, String> map) {
        // 读取word模板
//        String fileDir = new File(base.getFile(),"http://www.cnblogs.com/http://www.cnblogs.com/../doc/").getCanonicalPath();
        FileInputStream in = null;
        try {
            in = new FileInputStream(new File(filePath));
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }
        HWPFDocument hdt = null;
        try {
            hdt = new HWPFDocument(in);
        } catch (IOException e1) {
            e1.printStackTrace();
        }

        //读取word文本内容
        Range range = hdt.getRange();
        TableIterator tableIt = new TableIterator(range);
        //迭代文档中的表格
        int ii = 0;
        while (tableIt.hasNext()) {
            Table tb = (Table) tableIt.next();
            ii++;
//            System.out.println("第" + ii + "个表格数据...................");
            //迭代行，默认从0开始
            for (int i = 0; i < tb.numRows(); i++) {
                TableRow tr = tb.getRow(i);
                //只读前8行，标题部分
                if (i >= 8) {
                    break;
                }
                //迭代列，默认从0开始
                for (int j = 0; j < tr.numCells(); j++) {
                    TableCell td = tr.getCell(j);//取得单元格
                    //取得单元格的内容
                    for (int k = 0; k < td.numParagraphs(); k++) {
                        Paragraph para = td.getParagraph(k);
                        String s = para.text();
//                        System.out.println(s);
                    }
                }
            }
        }


        // 替换文本内容
        for (Map.Entry<String, String> entry : map.entrySet()) {
            range.replaceText(entry.getKey(), entry.getValue());
        }
        ByteArrayOutputStream ostream = new ByteArrayOutputStream();
        String fileName = "" + System.currentTimeMillis();
        fileName += ".doc";
        FileOutputStream out = null;
        try {
            out = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\合同文档\\" + fileName, true);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            hdt.write(ostream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 输出字节流
        try {
            out.write(ostream.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            ostream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
