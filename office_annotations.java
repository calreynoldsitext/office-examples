import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfString;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.annot.PdfAnnotation;
import com.itextpdf.kernel.pdf.annot.PdfTextAnnotation;
import com.itextpdf.layout.Document;
import com.itextpdf.licensekey.LicenseKey;
import com.itextpdf.pdfoffice.OfficeConverter;

import java.io.*;

public class PdfOfficeAnnotationExample {

	public static String DEST = "output_annotated.pdf";

    public static void main(String args[]) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream(0);

        OfficeConverter.convertOfficeDocumentToPdf(new FileInputStream("Alice in Wonderland Excerpt.docx"), baos);

        ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());

        PdfDocument pdfDoc = new PdfDocument(new PdfReader(bais), new PdfWriter(new FileOutputStream(DEST)));
        Document doc = new Document(pdfDoc);

        Rectangle rect = new Rectangle(80, 400, 0, 0);
        PdfAnnotation ann = new PdfTextAnnotation(rect);

        ann.setColor(ColorConstants.GREEN);
        ann.setContents("I love this part!");
        ann.setTitle(new PdfString("My favorite part."));

        pdfDoc.getFirstPage().addAnnotation(ann);

        doc.close();
        baos.close();
    }
}
