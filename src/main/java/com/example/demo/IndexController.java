package com.example.demo;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.stream.Collectors;

/**
 * @author shengchaojie
 * @date 2020-03-17
 **/
@Controller
public class IndexController {


    @RequestMapping("/")
    public String index(Model model) throws IOException {
        File  file = new ClassPathResource("paper").getFile();
        File[] files = file.listFiles();

        model.addAttribute("names", Arrays.asList(files).stream().map(File::getName).collect(Collectors.toList()));
        return "index";
    }

    @RequestMapping("/paper")
    public String paper(Model model, @RequestParam String name) throws IOException, TransformerException, ParserConfigurationException {
        model.addAttribute("name",name);
        model.addAttribute("content", word2003ToHtml("paper/"+name));
        return "paper";
    }

    @RequestMapping("/paper/answer")
    public String answer(Model model, @RequestParam String name) throws ParserConfigurationException, TransformerException, IOException {
        model.addAttribute("name",name);
        model.addAttribute("content", word2003ToHtml("answer/"+name));
        return "answer";
    }


    /**
     * 将word2003转换为html文件
     *
     * @throws IOException
     * @throws TransformerException
     * @throws ParserConfigurationException
     */
    public static String word2003ToHtml(String name)
            throws IOException, TransformerException, ParserConfigurationException {
        // 原word文档
        InputStream input = new ClassPathResource(name).getInputStream();
        HWPFDocument wordDocument = new HWPFDocument(input);
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        // 解析word文档
        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        //OutputStream outStream = new FileOutputStream(htmlFile);
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(byteArrayOutputStream);
        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer serializer = factory.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        return byteArrayOutputStream.toString();
    }

}
