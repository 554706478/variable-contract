package com.example.variablecontract;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.util.ResourceUtils;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;

/**
 * Description:
 * Author: v-wuchengs
 * Date: 2019-06-27
 */
public class util {

    public static String docToHtml() throws Exception {
        File path = new File(ResourceUtils.getURL("classpath:").getPath());
        String imagePathStr = path.getAbsolutePath() + "\\static\\image\\";
        String sourceFileName = path.getAbsolutePath() + "\\static\\test.doc";
        String targetFileName = path.getAbsolutePath() + "\\static\\test2.html";
        File file = new File(imagePathStr);
        if(!file.exists()) {
            file.mkdirs();
        }
        HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(sourceFileName));
        org.w3c.dom.Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);
        //保存图片，并返回图片的相对路径
        wordToHtmlConverter.setPicturesManager((content, pictureType, name, width, height) -> {
            try (FileOutputStream out = new FileOutputStream(imagePathStr + name)) {
                out.write(content);
            } catch (Exception e) {
                e.printStackTrace();
            }
            return "image/" + name;
        });
        wordToHtmlConverter.processDocument(wordDocument);
        org.w3c.dom.Document htmlDocument = wordToHtmlConverter.getDocument();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(new File(targetFileName));
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        return targetFileName;
    }

    public static String docxToHtml() throws Exception {
        File path = new File(ResourceUtils.getURL("classpath:").getPath());
        String imagePath = path.getAbsolutePath() + "\\static\\image";
        String sourceFileName = path.getAbsolutePath() + "\\static\\test.docx";
        String targetFileName = path.getAbsolutePath() + "\\static\\test.html";

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

    public static String readfile(String filePath) {
        File file = new File(filePath);
        InputStream input = null;
        try {
            input = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        StringBuffer buffer = new StringBuffer();
        byte[] bytes = new byte[1024];
        try {
            for (int n; (n = input.read(bytes)) != -1;) {
                buffer.append(new String(bytes, 0, n, "utf8"));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return buffer.toString();
    }



}
