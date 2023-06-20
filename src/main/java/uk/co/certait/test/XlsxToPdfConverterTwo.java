package uk.co.certait.test;

import com.lowagie.text.DocumentException;
import com.sun.webkit.ContextMenu;
import org.apache.commons.io.*;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;
import org.xhtmlrenderer.layout.SharedContext;
import org.xhtmlrenderer.pdf.ITextRenderer;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import scala.collection.JavaConverters;
import javax.swing.*;
import java.awt.*;
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
        for (Element element : document.select("*")) {
            if (element.hasClass("style_00")) {
                element.remove();
                element.removeClass("style_00");


            }
        }
        document.outputSettings().indentAmount(0).syntax(Document.OutputSettings.Syntax.xml);
        //    add new line after the tables
        document.getElementsByTag("tbody").append("<br/>");

        System.out.println(document.toString().replace("<tr>            \n" +
                "</tr>","").trim());
        return document;

    };

//    get the pdf output with css style intact
    public static void htmlToPdf(String tmpLocation, String fileoutlocation){
    try{
        FileOutputStream outputStream = new FileOutputStream(fileoutlocation);
        Document html_doc= generateXSSFDoc(tmpLocation);
//        System.out.println(html_doc.toString().replace("text-color","color").replace("&nbsp;"," ").replace("<td class=\"style_00\"> </td>","").replace("\n",
//            "").replaceAll("<tr>",""));
        try {
            ITextRenderer renderer = new ITextRenderer();
            SharedContext sharedContext = renderer.getSharedContext();
            sharedContext.setPrint(true);
            sharedContext.setInteractive(false);
            sharedContext.isPaged();
            renderer.setDocumentFromString(html_doc.toString().replace("text-color","color").replace("&nbsp;","").replace("<td class=\"style_00\"> </td>","").replace("<tr>            \n" +
                    "</tr>","").trim());
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
        String inputFile = "C:/Users/aagupt/Desktop/projects/11.xx reports/Enhancements/test files/XLSX/report_11_66_gtf_sales_by_draw.xlsx";
        FileInputStream in = new FileInputStream(inputFile);

        FileOutputStream html_output = new FileOutputStream(inputFile.replace(".xlsx",".html"));
        FileOutputStream pdf_output = new FileOutputStream(inputFile.replace(".xlsx",".pdf"));

        // this class is based on code found at
        // https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/ss/examples/html/ToHtml.java
        // and will convert .xlsx files
         ExcelToHtmlConverter toHtml = ExcelToHtmlConverter.create(inputFile,html_output);
         toHtml.setCompleteHTML(true);
         toHtml.printPage();
         in.close();
         html_output.close();
         htmlToPdf(inputFile.replace(".xlsx",".html"),inputFile.replace(".xlsx",".pdf"));


    }

}
