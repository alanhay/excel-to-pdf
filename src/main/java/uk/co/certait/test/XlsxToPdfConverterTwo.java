package uk.co.certait.test;

import java.io.FileWriter;
import java.io.InputStream;
import java.io.PrintWriter;

public class XlsxToPdfConverterTwo {

	public static void main(String[] args) throws Exception {

		InputStream in = XlsxToPdfConverterTwo.class.getResourceAsStream("/test.xlsx");
		PrintWriter out = new PrintWriter(new FileWriter("./test-xlsx.html"));

		// this class is based on code found at
		// https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/ss/examples/html/ToHtml.java
		// and will convert .xlsx files
		ExcelToHtmlConverter toHtml = ExcelToHtmlConverter.create(in, out);
		toHtml.setCompleteHTML(true);
		toHtml.printPage();

		// rather than writing to file get the HTML in memory and use
		// FlyingSaucer or OpenHTMlToPdf

		in.close();
	}
}
