package e2pdf;

import java.io.*;

import com.itextpdf.text.pdf.draw.LineSeparator;
import org.apache.poi.ss.usermodel.*;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

public class Main {
    static PdfWriter writer;
    static String datum2 = null;
    static String jciStr = " ";
    static String tb = " ";
    static String desc = null;
    static String um = null;
    static int umkoef = 0;
    static double stopa = 0;
    static double kurs = 0;
    static double sumKol2 = 0;
    static double sumDaz = 0;
    static double sumVred = 0;
    static double dinVred = 0;
    static double kol = 0;

    static int opisNum = 1;

    static double sumDazbina = 0;
    static double sumDinVrednosti = 0;
    static double sumDazbinaPoTabeli = 0;
    static double sumDinVrednostiPoTabeli = 0;
    static int noOfColumns = 0;

    static PdfPTable pdfTable;

    public static void main(String[] args) throws Exception {

        
        File f = new File("C:\\Users\\Djordje\\Desktop\\ts.xlsx");
        Workbook workbook = WorkbookFactory.create(f);
        Sheet worksheet = workbook.getSheetAt(1);
        Iterator<Row> rowIterator = worksheet.iterator();
        int rowNum = countRows(rowIterator) -1; 
        rowIterator = worksheet.iterator();
        Document document = new Document(PageSize.A4, 50, 50, 50, 50);
        Header h = new Header();
        writer = PdfWriter.getInstance(document, new FileOutputStream("output.pdf"));
        writer.setPageEvent(h);
        document.open();
        pdfTable = new PdfPTable(18);
        pdfTable.setWidthPercentage(100);
        pdfTable.setTotalWidth(PageSize.A4.getWidth() - document.leftMargin() - document.rightMargin());
        String prevJci = " ";
        String prevTb = " ";
        addTitle(document, workbook, true);
        rowIterator.next();

        int rowNumber = 0;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            int count = countCells(cellIterator);
            cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                getCellContent(cell);
            }

            if(rowNumber == 0){
                noOfColumns = count;
                beginTable(document);
            }
            if(count != noOfColumns) {
                addBottomAndDraw(document);
                break;
            }
            if ((!prevTb.equals(tb) && !prevTb.equals(" ")) || (!prevJci.equals(jciStr) && !prevJci.equals(" ")) ) {
                addBottomAndDraw(document);
                //if(!prevTb.equals(tb) && !prevTb.equals(" ")) {
                    //document.newPage();
                    //addTitle(document,workbook, false);
                //}

                beginTable(document);
                sumDazbinaPoTabeli = 0;
                sumDinVrednostiPoTabeli = 0;
            }
            addContent(document);

            prevTb = tb;
            prevJci = jciStr;
            rowNumber++;
        }
        addBottomAndDraw(document);
        addEnd(document, sumDinVrednosti, sumDazbina, true);
        document.close();
        noOfPages();


    }


    public static void noOfPages() throws IOException {
        PdfReader reader = null;
        try {
            reader = new PdfReader("output.pdf");
        } catch (IOException e) {
            e.printStackTrace();
        }
        int Pages = reader.getNumberOfPages();
        ByteArrayOutputStream ms = new ByteArrayOutputStream();

        PdfStamper stamper = null;
        try {
            stamper = new PdfStamper(reader, ms);
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        for (int i = 1; i <= Pages; i++)
        {
            PdfContentByte overContent;
            overContent = stamper.getOverContent(i);
            BaseFont font = null;
            try {
                font = BaseFont.createFont(BaseFont.TIMES_ROMAN, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
            } catch (DocumentException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            overContent.saveState();
            overContent.beginText();
            overContent.setFontAndSize(font, 8.5f);
            int xPos = 552 + (i/10)*2;
            overContent.setTextMatrix(xPos,800);
            overContent.showText("/"+ Pages);
            overContent.endText();
            overContent.restoreState();
        }
        try {
            stamper.close();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        FileOutputStream fos = new FileOutputStream("output.pdf");
        for(int i = 0; i < ms.size(); i++){
            fos.write(ms.toByteArray()[i]);
        }
    }
    public static void constLinija(String datum) {
        PdfPCell tmpCell = new PdfPCell();
        tmpCell.setFixedHeight(12);
        tmpCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tmpCell.setBorder(Rectangle.NO_BORDER);
        tmpCell.setColspan(3);
        Chunk c = new Chunk("45063",
                FontFactory.getFont(FontFactory.TIMES, 6f));
        tmpCell.setPhrase(new Phrase(c));
        tmpCell.setBorder(Rectangle.BOTTOM);
        pdfTable.addCell(tmpCell);
        tmpCell.setBorder(Rectangle.NO_BORDER);

        tmpCell.setPhrase(new Phrase("C5",
                FontFactory.getFont(FontFactory.TIMES, 6f)));
        pdfTable.addCell(tmpCell);

        c = new Chunk(datum,
                FontFactory.getFont(FontFactory.TIMES, 6f));
        tmpCell.setPhrase(new Phrase(c));
        tmpCell.setColspan(3);
        tmpCell.setBorder(Rectangle.BOTTOM);
        pdfTable.addCell(tmpCell);
        tmpCell.setBorder(Rectangle.NO_BORDER);

        tmpCell.setPhrase(new Phrase("E01",
                FontFactory.getFont(FontFactory.TIMES, 6f)));
        pdfTable.addCell(tmpCell);

        c = new Chunk("45063/PO/94/2021",
                FontFactory.getFont(FontFactory.TIMES, 6f));
        tmpCell.setPhrase(new Phrase(c));
        tmpCell.setColspan(4);
        tmpCell.setBorder(Rectangle.BOTTOM);
        pdfTable.addCell(tmpCell);
        tmpCell.setBorder(Rectangle.NO_BORDER);


        tmpCell.setPhrase(new Phrase(" "));
        tmpCell.setColspan(3);
        pdfTable.addCell(tmpCell);

        tmpCell.setBorder(Rectangle.NO_BORDER);
        tmpCell.setColspan(3);
        c = new Chunk("carinska ispostava",
                FontFactory.getFont(FontFactory.TIMES, 6f));
        tmpCell.setPhrase(new Phrase(c));
        pdfTable.addCell(tmpCell);

        tmpCell.setPhrase(new Phrase(" ",
                FontFactory.getFont(FontFactory.TIMES, 6f)));
        pdfTable.addCell(tmpCell);

        c = new Chunk("datum",
                FontFactory.getFont(FontFactory.TIMES, 6f));
        tmpCell.setPhrase(new Phrase(c));
        tmpCell.setColspan(3);
        pdfTable.addCell(tmpCell);

        tmpCell.setPhrase(new Phrase(" ",
                FontFactory.getFont(FontFactory.TIMES, 6f)));
        pdfTable.addCell(tmpCell);

        c = new Chunk("broj             datum",
                FontFactory.getFont(FontFactory.TIMES, 6f));
        tmpCell.setPhrase(new Phrase(c));
        tmpCell.setColspan(4);
        pdfTable.addCell(tmpCell);

        tmpCell.setPhrase(new Phrase(" "));
        tmpCell.setColspan(3);
        pdfTable.addCell(tmpCell);

    }

    public static void addBottomAndDraw(Document doc) throws DocumentException {
        PdfPCell tmpCell = new PdfPCell();
        tmpCell.setColspan(3);
        tmpCell.setPhrase(new Phrase("Ukupno", FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);

        tmpCell.setColspan(7);
        tmpCell.setPhrase(new Phrase(" ", FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);

        tmpCell.setColspan(2);
        tmpCell.setPhrase(new Phrase(String.format("%.2f", sumDinVrednostiPoTabeli), FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);

        tmpCell.setColspan(4);
        tmpCell.setPhrase(new Phrase(" ", FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);

        tmpCell.setColspan(2);
        tmpCell.setPhrase(new Phrase(String.format("%.2f", sumDazbinaPoTabeli), FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);
        if(pdfTable.getTotalHeight() > (writer.getVerticalPosition(false) - doc.bottomMargin())){
            doc.newPage();
        }
        doc.add(pdfTable);
        doc.add(new Paragraph(new Phrase("\n")));
    }

    public static void addContent(Document doc){
        PdfPCell tmpCell = new PdfPCell();
        tmpCell.setColspan(3);
        tmpCell.setPhrase(new Phrase(desc, FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);

        tmpCell.setPhrase(new Phrase(tb + "",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell.setColspan(3);
        pdfTable.addCell(tmpCell);

        tmpCell.setPhrase(new Phrase(String.format("%.2f",sumKol2) + "", FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell.setColspan(2);
        pdfTable.addCell(tmpCell);

        sumDinVrednosti += sumVred * kurs;
        sumDinVrednostiPoTabeli += sumVred * kurs;

        tmpCell.setPhrase(new Phrase(String.format("%.2f", sumVred),
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);

        tmpCell.setPhrase(new Phrase(String.format("%.2f", sumVred * kurs), FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell.setColspan(2);
        pdfTable.addCell(tmpCell);

        tmpCell.setPhrase(new Phrase(String.format("%.2f", (int) 100 * stopa) + "%", FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);
        

        tmpCell.setPhrase(new Phrase(" ", FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);

        sumDazbina += sumVred*kurs*stopa;
        sumDazbinaPoTabeli += sumVred*kurs*stopa;

        tmpCell.setPhrase(new Phrase(String.format("%.2f", sumVred*kurs*stopa), FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell);
        

    }

    public static void beginTable(Document doc) throws DocumentException {

        pdfTable = new PdfPTable(18);
        pdfTable.setWidthPercentage(100);

        pdfTable.setTotalWidth(PageSize.A4.getWidth() - doc.leftMargin() - doc.rightMargin());
        pdfTable.getRows().clear();
        pdfTable.setComplete(true);


        constLinija(jciStr + "/" +datum2);

        PdfPCell tmpCell1 = new PdfPCell(new Phrase("Trgovacki naziv proizvoda",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell1.setColspan(3);
        pdfTable.addCell(tmpCell1);
        PdfPCell tmpCell2 = new PdfPCell(new Phrase("Tarifna oznaka",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell2.setColspan(3);
        pdfTable.addCell(tmpCell2);

        tmpCell2.setColspan(2);
        tmpCell2.setPhrase(new Phrase("Kolicina",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        pdfTable.addCell(tmpCell2);

        PdfPCell tmpCell3 = new PdfPCell(new Phrase("Vrednost valutna",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell3.setColspan(2);
        pdfTable.addCell(tmpCell3);
        PdfPCell tmpCell4 = new PdfPCell(new Phrase("Vrednost dinarska",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell4.setColspan(2);
        pdfTable.addCell(tmpCell4);
        PdfPCell tmpCell5 = new PdfPCell(new Phrase("Stopa carine",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell5.setColspan(2);
        pdfTable.addCell(tmpCell5);
        PdfPCell tmpCell6 = new PdfPCell(new Phrase("Dažbine jedn. dejstva",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell6.setColspan(2);
        pdfTable.addCell(tmpCell6);

        PdfPCell tmpCell7 = new PdfPCell(new Phrase("Iznos duga",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        tmpCell7.setColspan(2);
        pdfTable.addCell(tmpCell7);

    }

    public static void addTitle(Document iText_xls_2_pdf, Workbook my_xls_workbook, boolean beginning) throws DocumentException {
        Paragraph title = new Paragraph();
        Phrase titlePh = new Phrase("DOKUMENT O OBRACUNU CARINSKIH DAŽBINA PO OSNOVU\n" +
                "ZABRANE POVRACAJA ILI OSLOBOÐENJA BR.\n" +
                "Redni broj naimenovanja; naziv proizvoda; tarifna oznaka; kolicina; vrednost EUR/RSD",
                FontFactory.getFont(FontFactory.TIMES, 8f));
        title.add(titlePh);
        title.setAlignment(Element.ALIGN_CENTER);
        iText_xls_2_pdf.add(title);

        ArrayList<String> sheet1 = new ArrayList<>();
        Iterator<Row> iteratorSh1 = my_xls_workbook.getSheetAt(0).iterator();
        while (iteratorSh1.hasNext()) {
            sheet1.add((iteratorSh1.next().cellIterator().next().getStringCellValue()));
        }

        Paragraph jci = new Paragraph();
        Phrase p1 = new Phrase("PREMA IZVOZNOM PROIZVODU IZ JCI C3 ____/______ broj ______/______\n\n",
                FontFactory.getFont(FontFactory.TIMES, 8f));
        jci.add(p1);

        iText_xls_2_pdf.add(jci);

        Paragraph opis = new Paragraph();
        Phrase brStrane = new Phrase(opisNum + ".    ",
                FontFactory.getFont(FontFactory.TIMES, 12f));
        String opisString  = "";
       if(beginning){
           opisString ="                 DOBIJENI PROIZVODI:	/";

            for (String s : sheet1)
                opisString += s + ", ";
            opisString = opisString.substring(0, opisString.length() - 1);
       }
        Phrase opisPh = new Phrase(opisString += "\n\n", FontFactory.getFont(FontFactory.TIMES, 8f));
        opisNum++;
        opis.setIndentationRight(150);
        opis.add(brStrane);
        opis.add(opisPh);
        iText_xls_2_pdf.add(opis);
        if(beginning) {
           Paragraph valute = new Paragraph();
          //  Phrase valutePh = new Phrase("/                           /                           PARI/                           EUR/" +
          //       "                           RSD", FontFactory.getFont(FontFactory.TIMES, 8f));
            valute.setSpacingAfter(10);
            valute.setAlignment(Element.ALIGN_RIGHT);
         //   valute.add(valutePh);
           iText_xls_2_pdf.add(valute);
        }

        LineSeparator ls = new LineSeparator();

        Paragraph nl = new Paragraph(new Phrase("\n"));
        iText_xls_2_pdf.add(ls);
        iText_xls_2_pdf.add(nl);
    }

    public static void addEnd(Document doc, double sumaDinVrednost, double sumaDazbine, boolean end) throws DocumentException {
        doc.add(new Phrase("\n"));
        if(writer.getVerticalPosition(false)-doc.bottomMargin() <150)
        	doc.newPage();
        doc.add(new Phrase("Ukupna vrednost-carinska osnovica __  ___ naimen.                                " +
                "                     Ukupan iznos carine za __  ___ naimenovanje",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        doc.add(new Phrase("\n"));
        PdfPTable table = new PdfPTable(18);
        PdfPCell cell = new PdfPCell();

        cell.setColspan(10);
        cell.setPhrase(new Phrase(String.format("%.2f", sumaDinVrednost),
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        table.addCell(cell);

        cell.setColspan(8);
        cell.setPhrase(new Phrase(String.format("%.2f", sumaDazbine),
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        table.addCell(cell);
        table.setWidthPercentage(100);
        doc.add(table);
        table = new PdfPTable(18);

        Paragraph p = new Paragraph(new Phrase("Iznos ostalih dazbina jednakog dejstva",
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        p.setAlignment(Element.ALIGN_RIGHT);
        p.setIndentationRight(96);
        p.setSpacingAfter(3);

        doc.add(p);

        cell.setColspan(10);
        cell.setPhrase(new Phrase(String.format(" ", sumaDinVrednost),
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        cell.setBorder(Rectangle.NO_BORDER);
        table.addCell(cell);

        cell.setColspan(8);
        cell.setBorder(Rectangle.BOX);
        cell.setPhrase(new Phrase(String.format(" ", sumaDazbine),
                FontFactory.getFont(FontFactory.TIMES, 8f)));
        table.addCell(cell);

        table.setWidthPercentage(100);
        doc.add(table);
        if(end) {
            Paragraph p1 = new Paragraph(new Phrase("___________________________\nOdgovorno lice",
                    FontFactory.getFont(FontFactory.TIMES, 8f)));
            p1.setAlignment(Element.ALIGN_CENTER);
            p1.setIndentationLeft(382);
            p1.setSpacingBefore(30);
            doc.add(p1);
        }
    }

    public static void getCellContent(Cell cell){

        switch (cell.getColumnIndex()) {
            case 0:
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    SimpleDateFormat sdf =  new SimpleDateFormat("dd.MM.yyyy");
                    datum2 = sdf.format(cell.getDateCellValue());
                } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    datum2 = cell.getStringCellValue();
                }
                break;
            case 1:
                if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
                    jciStr = cell.getNumericCellValue()+"";
                else jciStr = cell.getStringCellValue();
                if(jciStr.isEmpty())
                    jciStr = " ";
                break;
            case 2:
                tb = String.format("%.0f", cell.getNumericCellValue());
                if (tb.isEmpty()) {
                    tb = " ";
                }
                break;
            case 3:
                desc = cell.getStringCellValue();
                break;
            case 4:
                um = cell.getStringCellValue();
                break;
            case 5:
                umkoef = (int) cell.getNumericCellValue();
                break;
            case 6:
                kurs = cell.getNumericCellValue();
                break;
            case 7:
                stopa = cell.getNumericCellValue();
                break;
            case 9:
                sumKol2 = cell.getNumericCellValue();
                break;
            case 11:
                sumVred = cell.getNumericCellValue();
                break;
            case 13:
                sumDaz = cell.getNumericCellValue();
                break;
            case 14:
                dinVred = cell.getNumericCellValue();
                break;

        }
    }

    public static int countCells(Iterator<Cell> cellIterator){
        Iterator<Cell> tmp = cellIterator;
        int count = 0;
        while (tmp.hasNext()){
            tmp.next();
            count ++;
        }
        return count;
    }

    public static int countRows(Iterator<Row> rowIterator){
        Iterator<Row> tmp = rowIterator;
        int count = 0;
        while (tmp.hasNext()){
            tmp.next();
            count ++;
        }
        return count;
    }

}