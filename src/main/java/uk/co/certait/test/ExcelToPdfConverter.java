package uk.co.certait.test;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.w3c.dom.Document;
import org.xhtmlrenderer.pdf.ITextRenderer;

public class ExcelToPdfConverter {

	public static void main(String[] args) throws Exception {

		File file = new File(ExcelToPdfConverter.class.getResource("/test.xls").toURI());
		Document doc = ExcelToHtmlConverter.process(file);

		FileOutputStream out = new FileOutputStream("./test.pdf");
		ITextRenderer renderer = new ITextRenderer();
		renderer.setDocument(doc, null);
		renderer.layout();
		renderer.createPDF(out);
		out.flush();
		out.close();
	}
}
