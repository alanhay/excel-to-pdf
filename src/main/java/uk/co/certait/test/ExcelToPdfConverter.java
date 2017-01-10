package uk.co.certait.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.w3c.dom.Document;
import org.xhtmlrenderer.pdf.ITextRenderer;

public class ExcelToPdfConverter {

	public static void main(String[] args) throws Exception {

		File file = new File(ExcelToPdfConverter.class.getResource("/test.xls").toURI());
		Document doc = ExcelToHtmlConverter.process(file);

		debugHtml(doc);
		writePdf(doc);
	}

	public static void debugHtml(Document doc) throws Exception {
		DOMSource source = new DOMSource(doc);
		FileWriter writer = new FileWriter(new File("./test.html"));
		StreamResult result = new StreamResult(writer);

		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		transformer.transform(source, result);
	}

	public static void writePdf(Document doc) throws Exception {
		FileOutputStream out = new FileOutputStream("./test.pdf");
		ITextRenderer renderer = new ITextRenderer();
		renderer.setDocument(doc, null);
		renderer.layout();
		renderer.createPDF(out);
		out.flush();
		out.close();
	}
}
