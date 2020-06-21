package com.hseb.main;

import java.io.File;
import java.io.FileOutputStream;
import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.TabSettings;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * Sajan Kc
 *
 */
public class HSEB_Main {

	private static final String FILE_NAME = "E:/IText/PDFDocuments/hseb_marksheet.pdf";
	private static Font fontSize8 = new Font(Font.FontFamily.TIMES_ROMAN, 8);
	private static Font fontSize10 = new Font(Font.FontFamily.TIMES_ROMAN, 10);
	private static Font fontSize10H = new Font(Font.FontFamily.COURIER, 10);
	private static Font fontSize10Bold = new Font(Font.FontFamily.TIMES_ROMAN, 10, Font.BOLD);
	private static Font fontSize10Underline = new Font(Font.FontFamily.TIMES_ROMAN, 10, Font.UNDERLINE);
	private static Font fontSize12Bold = new Font(Font.FontFamily.TIMES_ROMAN, 12, Font.BOLD);
	private static Font fontSize20Bold = new Font(Font.FontFamily.TIMES_ROMAN, 20, Font.BOLD);

	public static void main(String[] args) throws Exception {

		Document document = new Document();
		Rectangle documentSize = new Rectangle(594.72f, 770f); // 1 inch = 72f
		document.setPageSize(documentSize);
		PdfWriter.getInstance(document, new FileOutputStream(new File(FILE_NAME)));
		document.open();

		// Document properties
		document.addTitle("HSEB MARKSHEET");
		document.addAuthor("Sajan Kc");
		document.addCreator("Sajan Kc");
		document.addKeywords("hseb , marksheet");
		document.addSubject("MarkSheet");

		createHeaderTable(document);
		createBodyTable(document);
		createFooterTable(document);

		document.close();
		System.out.println("PDF created Successively");
	}

	private static void createHeaderTable(Document document) throws DocumentException {
		float[] headerTableWidth = { 1.5f, 3f };
		PdfPTable invisibleHeaderTable = new PdfPTable(headerTableWidth);
		invisibleHeaderTable.setWidthPercentage(100);

		PdfPCell headerLeftCell01 = new PdfPCell();
		headerLeftCell01.setBorder(PdfPCell.NO_BORDER);
		Paragraph issueNo = new Paragraph("Issue No.", fontSize8);
		headerLeftCell01.addElement(issueNo);

		PdfPCell headerRightCell02 = new PdfPCell();
		headerRightCell02.setBorder(PdfPCell.NO_BORDER);
		Paragraph registrationNo = new Paragraph("HSEB Registration No. 702736330\n ", fontSize8);
		registrationNo.setAlignment(Element.ALIGN_RIGHT);
		headerRightCell02.addElement(registrationNo);

		String header = "HIGHER SECONDAY EDUCATION BOARD \n NEPAL";
		Paragraph mainHeading = new Paragraph(20, header, fontSize20Bold);

		PdfPCell mainHeadingCell = new PdfPCell(mainHeading);
		mainHeadingCell.setBorder(PdfPCell.NO_BORDER);
		mainHeadingCell.setColspan(2);
		mainHeadingCell.setHorizontalAlignment(Element.ALIGN_CENTER);

		String subHead = "(Estd. Under the Higher Secondary Education Act, 1989) \n Academic Transcript \n ";
		Paragraph subHeading = new Paragraph(20, subHead, fontSize12Bold);

		PdfPCell subHeadingCell = new PdfPCell(subHeading);
		subHeadingCell.setBorder(PdfPCell.NO_BORDER);
		subHeadingCell.setColspan(2);
		subHeadingCell.setHorizontalAlignment(Element.ALIGN_CENTER);

		// Student Info
		Paragraph studentInfo = new Paragraph();
		studentInfo.add(new Phrase("Name of the Student: ", fontSize10Bold));
		studentInfo.add(new Phrase("SAJAN KC", fontSize10H));
		PdfPCell studentInfoCell = new PdfPCell(studentInfo);
		studentInfoCell.setBorder(PdfPCell.NO_BORDER);

		Paragraph school = new Paragraph("School :", fontSize10Bold);
		PdfPCell schoolCell = new PdfPCell(school);
		schoolCell.setHorizontalAlignment(Element.ALIGN_CENTER);
		schoolCell.setBorder(PdfPCell.NO_BORDER);

		Paragraph dob = new Paragraph();
		dob.add(new Phrase("Date of Birth: ", fontSize10Bold));
		dob.add(new Phrase("2053/03/17", fontSize10H));
		PdfPCell dobCell = new PdfPCell(dob);
		dobCell.setBorder(PdfPCell.NO_BORDER);

		Paragraph college = new Paragraph("ST. LAWRENCE H S SCHOOL,CHABAHIL, KATHMANDU", fontSize10H);
		PdfPCell collegeCell = new PdfPCell(college);
		collegeCell.setBorder(PdfPCell.NO_BORDER);

		invisibleHeaderTable.addCell(headerLeftCell01);
		invisibleHeaderTable.addCell(headerRightCell02);
		invisibleHeaderTable.addCell(mainHeadingCell);
		invisibleHeaderTable.addCell(subHeadingCell);
		invisibleHeaderTable.addCell(studentInfoCell);
		invisibleHeaderTable.addCell(schoolCell);
		invisibleHeaderTable.addCell(dobCell);
		invisibleHeaderTable.addCell(collegeCell);

		invisibleHeaderTable.setSpacingAfter(10f);

		document.add(invisibleHeaderTable);
	}

	private static void createBodyTable(Document document) throws DocumentException {

		float[] invisibleBodyTableWidth = { 3f, 1.5f };
		PdfPTable invisibleBodyTable = new PdfPTable(invisibleBodyTableWidth);
		invisibleBodyTable.setWidthPercentage(100);

		// Left table
		PdfPCell mainTableLeftCell = new PdfPCell();

		mainTableLeftCell.setBorder(PdfPCell.NO_BORDER);

		// Grade XI
		float[] tableCellWidths = { 5f, 1.2f, 1.2f, 1.2f, 1.2f };
		PdfPTable leftTable = new PdfPTable(tableCellWidths);
		leftTable.setHorizontalAlignment(Element.ALIGN_LEFT);
		leftTable.setWidthPercentage(100);

		PdfPCell subjects = new PdfPCell(new Paragraph("Subjects", fontSize10Bold));
		subjects.setHorizontalAlignment(Element.ALIGN_CENTER);
		subjects.setPaddingTop(5f);

		PdfPCell fullMarks = new PdfPCell(new Paragraph("Full \n Marks", fontSize10));
		fullMarks.setHorizontalAlignment(Element.ALIGN_CENTER);
		fullMarks.setPaddingBottom(5f);

		PdfPCell passMarks = new PdfPCell(new Paragraph("Pass \n Marks", fontSize10));
		passMarks.setHorizontalAlignment(Element.ALIGN_CENTER);
		passMarks.setVerticalAlignment(Element.ALIGN_CENTER);

		PdfPCell marksSecured = new PdfPCell(new Paragraph("Marks \n Secured", fontSize10));
		marksSecured.setHorizontalAlignment(Element.ALIGN_CENTER);
		marksSecured.setVerticalAlignment(Element.ALIGN_CENTER);

		PdfPCell remarks = new PdfPCell(new Paragraph("Remarks", fontSize10));
		remarks.setHorizontalAlignment(Element.ALIGN_CENTER);
		remarks.setVerticalAlignment(Element.ALIGN_CENTER);

		leftTable.addCell(subjects);
		leftTable.addCell(fullMarks);
		leftTable.addCell(passMarks);
		leftTable.addCell(marksSecured);
		leftTable.addCell(remarks);

		// Table CORE subject part Grade XI
		PdfPCell xiCell11 = new PdfPCell(new Paragraph("Grade XI \n\n   English \n\n   Nepali", fontSize10Bold));
		xiCell11.setPaddingBottom(5f);

		PdfPCell xiCell12 = new PdfPCell(new Paragraph("\n\n 100 \n\n 100", fontSize10));
		xiCell12.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiCell13 = new PdfPCell(new Paragraph("\n\n 35 \n\n 35", fontSize10));
		xiCell13.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiCell14 = new PdfPCell(new Paragraph("\n\n 65 \n\n 70", fontSize10H));
		xiCell14.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiCell15 = new PdfPCell(new Paragraph(" "));

		leftTable.addCell(xiCell11);
		leftTable.addCell(xiCell12);
		leftTable.addCell(xiCell13);
		leftTable.addCell(xiCell14);
		leftTable.addCell(xiCell15);

		// Table ELECTIVES subjects part Grade XI

		PdfPCell xiCell21 = new PdfPCell(new Paragraph(
				" ACCOUNTANCY \n ECONOMICS \n COMPUTER SCIENCE (TH) \n COMPUTER SCIENCE (PR) \n \n \n \n \n ",
				fontSize10));

		PdfPCell xiCell22 = new PdfPCell(new Paragraph(" 100 \n 100 \n 075 \n 025", fontSize10H));
		xiCell22.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiCell23 = new PdfPCell(new Paragraph(" 35 \n 35 \n 27 \n 10", fontSize10H));
		xiCell23.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiCell24 = new PdfPCell(new Paragraph(" 75 \n 72 \n 57 \n 25", fontSize10H));
		xiCell24.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiCell25 = new PdfPCell(new Paragraph(" "));

		leftTable.addCell(xiCell21);
		leftTable.addCell(xiCell22);
		leftTable.addCell(xiCell23);
		leftTable.addCell(xiCell24);
		leftTable.addCell(xiCell25);

		// Grade XI total part

		PdfPCell totalXI = new PdfPCell(new Paragraph("Total", fontSize10));
		totalXI.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell totalMarksXI = new PdfPCell(new Paragraph("500", fontSize10));
		totalMarksXI.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell blankXI1 = new PdfPCell(new Paragraph(" "));

		PdfPCell securedMarksXI = new PdfPCell(new Paragraph("364", fontSize10H));
		securedMarksXI.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell blankXI2 = new PdfPCell(new Paragraph(" "));

		leftTable.addCell(totalXI);
		leftTable.addCell(totalMarksXI);
		leftTable.addCell(blankXI1);
		leftTable.addCell(securedMarksXI);
		leftTable.addCell(blankXI2);

		// Grade XII // Table CORE subject part Grade XII
		PdfPCell xiiCell11 = new PdfPCell(new Paragraph("Grade XII \n\n   English \n\n   Nepali", fontSize10Bold));
		xiiCell11.setPaddingBottom(5f);

		PdfPCell xiiCell12 = new PdfPCell(new Paragraph("\n\n 100 \n\n 100", fontSize10));
		xiiCell12.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiiCell13 = new PdfPCell(new Paragraph("\n\n 35 \n\n 35", fontSize10));
		xiiCell13.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiiCell14 = new PdfPCell(new Paragraph("\n\n 70", fontSize10H));
		xiiCell14.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiiCell15 = new PdfPCell(new Paragraph(" "));
		xiiCell15.setPaddingTop(30f);

		leftTable.addCell(xiiCell11);
		leftTable.addCell(xiiCell12);
		leftTable.addCell(xiiCell13);
		leftTable.addCell(xiiCell14);
		leftTable.addCell(xiiCell15);

		// Table ELECTIVES subjects part Grade XII

		PdfPCell xiiCell21 = new PdfPCell(new Paragraph(
				" ACCOUNTANCY \n ECONOMICS \n COMPUTER SCIENCE (TH) \n COMPUTER SCIENCE (PR) \n BUSINESS MATHEMATICS \n \n \n \n \n ",
				fontSize10));

		PdfPCell xiiCell22 = new PdfPCell(new Paragraph(" 100 \n 100 \n 075 \n 025 \n 100", fontSize10H));
		xiiCell22.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiiCell23 = new PdfPCell(new Paragraph(" 35 \n 35 \n 27 \n 10 \n 35", fontSize10H));
		xiiCell23.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiiCell24 = new PdfPCell(new Paragraph(" 74 \n 78 \n 69 \n 25 \n 94", fontSize10H));
		xiiCell24.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell xiiCell25 = new PdfPCell(new Paragraph(" "));

		leftTable.addCell(xiiCell21);
		leftTable.addCell(xiiCell22);
		leftTable.addCell(xiiCell23);
		leftTable.addCell(xiiCell24);
		leftTable.addCell(xiiCell25);

		// Grade XII total part

		PdfPCell totalXII = new PdfPCell(new Paragraph("Total", fontSize10));
		totalXII.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell totalMarksXII = new PdfPCell(new Paragraph("500", fontSize10));
		totalMarksXII.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell blankXII1 = new PdfPCell(new Paragraph(" "));

		PdfPCell securedMarksXII = new PdfPCell(new Paragraph("410", fontSize10H));
		securedMarksXII.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell blankXII2 = new PdfPCell(new Paragraph(" "));

		leftTable.addCell(totalXII);
		leftTable.addCell(totalMarksXII);
		leftTable.addCell(blankXII1);
		leftTable.addCell(securedMarksXII);
		leftTable.addCell(blankXII2);

		// Grand Total
		PdfPCell grandTotal = new PdfPCell(new Paragraph("Grand Total", fontSize12Bold));
		grandTotal.setHorizontalAlignment(Element.ALIGN_CENTER);
		grandTotal.setPaddingBottom(5f);

		PdfPCell grandTotalMarks = new PdfPCell(new Paragraph("1000", fontSize10));
		grandTotalMarks.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell blankGrandTotal1 = new PdfPCell(new Paragraph(" "));

		PdfPCell securedMarksGrand = new PdfPCell(new Paragraph("774", fontSize10H));
		securedMarksGrand.setHorizontalAlignment(Element.ALIGN_CENTER);

		PdfPCell blankGrandTotal2 = new PdfPCell(new Paragraph(" "));

		leftTable.addCell(grandTotal);
		leftTable.addCell(grandTotalMarks);
		leftTable.addCell(blankGrandTotal1);
		leftTable.addCell(securedMarksGrand);
		leftTable.addCell(blankGrandTotal2);

		// Table footer part
		Paragraph p1 = new Paragraph();

		p1.add(new Chunk("\n\n\n Year of Completion: ", fontSize10));
		p1.add(new Chunk("2071 (2014)", fontSize10H));
		p1.setTabSettings(new TabSettings(80f));
		p1.add(Chunk.TABBING);

		p1.add(new Chunk("Checked by", fontSize10));
		p1.add(Chunk.TABBING);

		p1.add(new Chunk("       Verified by\n", fontSize10));
		p1.add(Chunk.TABBING);
		p1.add(Chunk.TABBING);
		p1.add(Chunk.TABBING);

		p1.add(new Chunk("(Deputy Controller)", fontSize10));

		p1.add(new Chunk("\n Date of Issue: ", fontSize10));
		p1.add(new Chunk(" 2071/08/01 \n ", fontSize10H));

		PdfPCell footerCell = new PdfPCell(p1);

		footerCell.setColspan(5);

		leftTable.addCell(footerCell);

		// Adding left table to main table
		mainTableLeftCell.addElement(leftTable);
		invisibleBodyTable.addCell(mainTableLeftCell);

		// Right table
		PdfPCell mainTableRightCell = new PdfPCell();
		mainTableRightCell.setBorder(PdfPCell.NO_BORDER);

		// Grade XI symbol number table
		float[] rightTableXIWidth = { 1.5f, 2f };
		PdfPTable rightTableXI = new PdfPTable(rightTableXIWidth);
		rightTableXI.setWidthPercentage(100);

		Paragraph titleXI = new Paragraph("Grade XI", fontSize10Bold);
		PdfPCell cellTitleXI = new PdfPCell(titleXI);
		cellTitleXI.setPaddingBottom(5f);
		cellTitleXI.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellTitleXI.setColspan(2);
		rightTableXI.addCell(cellTitleXI);

		PdfPCell cellYearXI = new PdfPCell(new Phrase("Year", fontSize10Bold));
		PdfPCell cellSymbolXI = new PdfPCell(new Phrase("Symbol Number", fontSize10Bold));
		cellYearXI.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellSymbolXI.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellYearXI.setPaddingBottom(5f);
		cellSymbolXI.setPaddingBottom(5f);

		rightTableXI.addCell(cellYearXI);
		rightTableXI.addCell(cellSymbolXI);

		PdfPCell cellYearValueXI = new PdfPCell(new Phrase("2070", fontSize10H));
		PdfPCell cellSymbolValueXI = new PdfPCell(new Phrase("12709039", fontSize10H));
		cellYearValueXI.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellSymbolValueXI.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellYearValueXI.setPaddingBottom(5f);
		cellSymbolValueXI.setPaddingBottom(5f);

		rightTableXI.addCell(cellYearValueXI);
		rightTableXI.addCell(cellSymbolValueXI);

		blankCell(rightTableXI, 6);

		rightTableXI.setSpacingBefore(25f);
		rightTableXI.setSpacingAfter(10f);

		// Grade XII symbol number table
		float[] rightTableXIIWidth = { 1.5f, 2f };
		PdfPTable rightTableXII = new PdfPTable(rightTableXIIWidth);
		rightTableXII.setWidthPercentage(100);

		Paragraph titleXII = new Paragraph("Grade XII", fontSize10Bold);
		PdfPCell cellTitleXII = new PdfPCell(titleXII);
		cellTitleXII.setPaddingBottom(5f);

		cellTitleXII.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellTitleXII.setColspan(2);
		rightTableXII.addCell(cellTitleXII);

		PdfPCell cellYearXII = new PdfPCell(new Phrase("Year", fontSize10Bold));
		PdfPCell cellSymbolXII = new PdfPCell(new Phrase("Symbol Number", fontSize10Bold));
		cellYearXII.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellSymbolXII.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellYearXII.setPaddingBottom(5f);
		cellSymbolXII.setPaddingBottom(5f);

		rightTableXII.addCell(cellYearXII);
		rightTableXII.addCell(cellSymbolXII);

		PdfPCell cellYearValueXII = new PdfPCell(new Phrase("2071", fontSize10H));
		PdfPCell cellSymboleValueXII = new PdfPCell(new Phrase("22708378", fontSize10H));
		cellYearValueXII.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellSymboleValueXII.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellYearValueXII.setPaddingBottom(5f);
		cellSymboleValueXII.setPaddingBottom(5f);

		rightTableXII.addCell(cellYearValueXII);
		rightTableXII.addCell(cellSymboleValueXII);

		blankCell(rightTableXII, 6);

		Paragraph report = new Paragraph(
				new Phrase("Percentage of aggregate marks secured in \nGrade XI and XII : ", fontSize10));
		report.add(new Phrase("77.40", fontSize10H));
		report.add(new Phrase("\nDivision : ", fontSize10));
		report.add(new Phrase("Distinction", fontSize10H));

		Paragraph examinationController = new Paragraph(new Chunk("\n\n\n\nController of Examinations", fontSize10));
		examinationController.setAlignment(Element.ALIGN_RIGHT);

		// Adding right table to main table
		mainTableRightCell.addElement(rightTableXI);
		mainTableRightCell.addElement(rightTableXII);
		mainTableRightCell.addElement(report);
		mainTableRightCell.addElement(examinationController);
		invisibleBodyTable.addCell(mainTableRightCell);

		document.add(invisibleBodyTable);
	}

	private static void createFooterTable(Document document) throws DocumentException {
		float[] invisibleTableWidth = { 3f, 2.9f };
		PdfPTable invisibleFooterTable = new PdfPTable(invisibleTableWidth);
		invisibleFooterTable.setWidthPercentage(100);

		// Populating first column
		Paragraph fp1 = new Paragraph();

		fp1.add(new Phrase("Grading System: \n", fontSize10Underline));

		fp1.add(new Phrase("\n75% and above - Distinction \n60% and above - First division "
				+ "\n45% and above - Second division \n35% and above - Pass division", fontSize10));
		fp1.setLeading(10f);
		PdfPCell cell01 = new PdfPCell(fp1);
		cell01.setHorizontalAlignment(Element.ALIGN_LEFT);
		cell01.setBorder(PdfPCell.NO_BORDER);
		invisibleFooterTable.addCell(cell01);

		// Populating second column
		Paragraph fp2 = new Paragraph();
		fp2.add(new Phrase(
				"\nTo pass the examination candicates must secure \n35% marks in theory and 40% marks in practical "
						+ "\npapers, separately.",
				fontSize10));
		Paragraph fp21 = new Paragraph();
		fp21.add(new Phrase("Note: * means a student has passed in the second attempt.", fontSize10));
		fp21.add(new Phrase("\n        ** means a student has passed in more than two attempts.", fontSize10));

		fp2.setLeading(10f);
		fp21.setLeading(10f);

		PdfPCell cell02 = new PdfPCell();
		cell02.setBorder(PdfPCell.NO_BORDER);
		cell02.addElement(fp2);
		cell02.addElement(fp21);

		invisibleFooterTable.addCell(cell02);
		document.add(invisibleFooterTable);

	}

	private static void blankCell(PdfPTable table, int cell) {
		for (int i = 0; i < 6; i++) {
			PdfPCell cellA = new PdfPCell(new Phrase(" "));
			PdfPCell cellB = new PdfPCell(new Phrase(" "));
			table.addCell(cellA);
			table.addCell(cellB);
		}

	}
}
