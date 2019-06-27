package com.example.variablecontract;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;

/**
 * Description:
 * Author: v-wuchengs
 * Date: 2019-06-27
 */
@RestController
@RequestMapping("/test")
public class Test6 {

    @GetMapping("/exportWord")
    public void export(HttpServletRequest request,HttpServletResponse response){
        String title = "你好";
        String text = "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; white-space: normal; background-color: rgb(255, 255, 255); text-indent: 0em; text-align: center;\">\n" +
                "    <strong style=\"margin: 0px; padding: 0px; list-style-type: none;\">劳动合同范本</strong>\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    <strong>甲方：</strong>\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    <strong>乙方：</strong>\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    <strong><br/></strong>\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    &nbsp;&nbsp;&nbsp;&nbsp;根据《中华人民共和国劳动法》和有关规定，甲乙双方经平等协商一致，自愿签订本合同，共同遵守本合同所列条款。\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    一、劳动合同期限第一条本合同为______________期限劳动合同。\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    二、工作内容\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    &nbsp;&nbsp;&nbsp;&nbsp;第二条乙方同意根据甲方工作需要，担任&nbsp;岗位(工种)工作。\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    &nbsp;&nbsp;&nbsp;&nbsp;第三条乙方工作应达到甲方规定的技术标准。\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    三、劳动保护和劳动条件第四条甲方安排乙方执行八小时工时制度。\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    &nbsp;&nbsp;&nbsp;&nbsp;第五条甲方为乙方提供必要的劳动条件和劳动工具。\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    &nbsp;&nbsp;&nbsp;&nbsp;第六条甲方负责对乙方进行职业道德、业务技术、劳动安全、劳动纪律和甲方规章制度的教育。\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    四、劳动报酬第七条甲方每月以货币形式支付乙方工资。\n" +
                "</p>\n" +
                "<p style=\"margin-top: 0px; margin-bottom: 0px; padding: 0px; list-style-type: none; color: rgb(51, 51, 51); font-family: tahoma, 微软雅黑; font-size: 14px; text-indent: 28px; white-space: normal; background-color: rgb(255, 255, 255);\">\n" +
                "    &nbsp;&nbsp;&nbsp;&nbsp;第八条甲方生产工作任务不足使乙方待工的，甲方无需支付乙方的月生活费。\n" +
                "</p>\n" +
                "<p>\n" +
                "    <br/>\n" +
                "</p>";
        this.exportWord(request,response,title,text);
    }



    public void exportWord(HttpServletRequest request, HttpServletResponse response, String title, String text) {

        try {
            //word内容
            String content="<html><body>" +
                    "<p style=\"text-align: center;\"><span style=\"font-family: 黑体, SimHei; font-size: 24px;\">"
                    + title + "</span></p>" + text + "</body></html>";
            byte b[] = content.getBytes("GBK");  //这里是必须要设置编码的，不然导出中文就会乱码。
            ByteArrayInputStream bais = new ByteArrayInputStream(b);//将字节数组包装到流中

            /*
             * 关键地方
             * 生成word格式
             * */
            POIFSFileSystem poifs = new POIFSFileSystem();
            DirectoryEntry directory = poifs.getRoot();
            DocumentEntry documentEntry = directory.createDocument("WordDocument", bais);
            //输出文件
            request.setCharacterEncoding("utf-8");
            response.setContentType("application/msword");//导出word格式
            response.addHeader("Content-Disposition", "attachment;filename=" +
                    new String(title.getBytes("GB2312"),"iso8859-1") + ".doc");
            ServletOutputStream ostream = response.getOutputStream();
            poifs.writeFilesystem(ostream);
            bais.close();
            ostream.close();
            poifs.close();
        }catch(Exception e){
            e.printStackTrace();
        }

    }

}
