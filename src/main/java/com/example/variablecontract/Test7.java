package com.example.variablecontract;

import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;

/**
 * Description:
 * Author: v-wuchengs
 * Date: 2019-06-27
 */
@RestController
@RequestMapping("/test")
public class Test7 {

    @GetMapping("/docxToHtml")
    public String export(HttpServletRequest request,HttpServletResponse response){
        String str = null ;
        try {
            str = docxToHtml();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return str;
    }

    public static String docxToHtml() throws Exception {
        File path = new File(ResourceUtils.getURL("classpath:").getPath());
        String imagePath = path.getAbsolutePath() + "\\static\\image";
        String sourceFileName = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\aa.docx";
        String targetFileName = "C:\\Users\\Administrator\\Desktop\\合同文档云打印调研材料\\aa.html";


        OutputStreamWriter outputStreamWriter = null;
        try {
            XWPFDocument document = new XWPFDocument(new FileInputStream(sourceFileName));
            XHTMLOptions options = XHTMLOptions.create();
            // 存放图片的文件夹
            options.setExtractor(new FileImageExtractor(new File(imagePath)));
            // html中图片的路径
            options.URIResolver(new BasicURIResolver("image"));
            outputStreamWriter = new OutputStreamWriter(new FileOutputStream(targetFileName), "utf-8");
            XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
            xhtmlConverter.convert(document, outputStreamWriter, options);
        } finally {
            if (outputStreamWriter != null) {
                outputStreamWriter.close();
            }
        }
        return targetFileName;
    }


}
