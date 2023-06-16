package uk.co.certait.test;

import com.lowagie.text.DocumentException;
import com.sun.webkit.ContextMenu;
import org.apache.commons.io.*;
import org.jsoup.nodes.Entities;
import org.xhtmlrenderer.layout.SharedContext;
import org.xhtmlrenderer.pdf.ITextRenderer;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

import javax.swing.*;
import java.io.*;
import java.nio.charset.StandardCharsets;

public class XlsxToPdfConverterTwo {
    public static Document generateXSSFDoc(String tmpLocation) throws IOException {
        //        FileInputStream tmpHtml = new FileInputStream(tmpLocation);
            FileInputStream tmpHtml = new FileInputStream(tmpLocation);
            String html_string = IOUtils.toString(tmpHtml, StandardCharsets.UTF_8);
            Document document = Jsoup.parse(html_string, "UTF-8");
        document.getElementsByClass("column-headers-background").remove();
        document.getElementsByClass("row-header-wrapper").remove();
        //    remove excel header /row number
        document.getElementsByClass("rowHeader").remove();
        document.getElementsByClass("colHeader").remove();
        document.outputSettings().syntax(Document.OutputSettings.Syntax.xml);
        //    add new line after the tables
        document.getElementsByTag("tbody").append("<br/>");
//        println("parsing excel workbook to pdf: " + document.toString().replace("text-color","color").replace("&nbsp;"," "))
        return document;

    };

//    get the pdf output with css style intact
    public static void htmlToPdf(String tmpLocation, String fileoutlocation){
    try{
        FileOutputStream outputStream = new FileOutputStream(fileoutlocation);
        Document html_doc= generateXSSFDoc(tmpLocation);
        try {
            ITextRenderer renderer = new ITextRenderer();
            SharedContext sharedContext = renderer.getSharedContext();
            sharedContext.setPrint(true);
            sharedContext.setInteractive(false);
            sharedContext.isPaged();
            renderer.setDocumentFromString(html_doc.toString().replace("text-color","color").replace("&nbsp;"," "));
            renderer.layout();
            renderer.createPDF(outputStream);


        } catch (DocumentException e) {
            throw new RuntimeException(e);
        } finally {
            if (outputStream != null) outputStream.close();
        }


    } catch (IOException e) {
        throw new RuntimeException(e);
    }

    }



    public static void main(String[] args) throws Exception {
//        add xlsx filepath here
        String inputFile = "";
        FileInputStream in = new FileInputStream(inputFile);

        FileOutputStream html_output = new FileOutputStream(inputFile.replace(".xlsx",".html"));
        FileOutputStream pdf_output = new FileOutputStream(inputFile.replace(".xlsx",".pdf"));


        // this class is based on code found at
        // https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/ss/examples/html/ToHtml.java
        // and will convert .xlsx files
        ExcelToHtmlConverter toHtml = ExcelToHtmlConverter.create(in, html_output);
        toHtml.setCompleteHTML(true);
        toHtml.printPage();
        in.close();
        htmlToPdf(inputFile.replace(".xlsx",".html"),inputFile.replace(".xlsx",".pdf"));


    }

}
