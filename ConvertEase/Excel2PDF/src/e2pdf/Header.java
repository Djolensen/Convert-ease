package e2pdf;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfWriter;


public class Header extends PdfPageEventHelper {
    public void onEndPage(PdfWriter writer, Document document) {
        ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_CENTER, new Phrase(""+writer.getCurrentPageNumber(),
                FontFactory.getFont(FontFactory.TIMES, 8f)), 550, 800, 0);
    }

}
