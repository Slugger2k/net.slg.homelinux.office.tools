package net.slg.homelinux.office.annotations;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Types;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

import net.slg.homelinux.office.annotations.ui.MapToPdfProcessBarTask;

import org.eclipse.swt.widgets.Display;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.ExceptionConverter;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;

public class MapPdfExporter {
	
	public static Font titleFont = new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD);
	public static Font tableFont = new Font(Font.FontFamily.HELVETICA, 8,Font.NORMAL);
	public static Font tableFontHeader = new Font(Font.FontFamily.HELVETICA, 8,Font.BOLD);
	public static Font bottomFontHeader = new Font(Font.FontFamily.HELVETICA, 6,Font.BOLD);
		
	private String columnNames[]; 
	private String columnTitles[];
	private int columnTypes[];
	private List<Map<String, Object>> data;
	private boolean orientationLandscape = false;
	private String title = "Export";
	private String bottomText = "";
	private int mode;
	
	public MapPdfExporter(int mode) {
		this.mode = mode;
	}
	
	public File exportPdf() {
		
		File tempFile;
		Document document;
		TableHeader tableHeader;
		try {
			File file = File.createTempFile("lfs", "export");
			tempFile = new File(file.getAbsolutePath() + ".pdf");
			file.delete();
			
			if (this.orientationLandscape) {
				document = new Document(PageSize.A4.rotate());
			} else {
				document = new Document(PageSize.A4);	
			}
			tableHeader = new TableHeader(this.orientationLandscape);
			tableHeader.setHeader(this.bottomText);
			PdfWriter writer = PdfWriter.getInstance(document,  new FileOutputStream(tempFile));
			writer.setPageEvent(tableHeader);
			document.open();
			addTitlePage(document, this.title);
			addTable(document);
			document.close();
			return tempFile;	
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
		
	}
	
	private void addTitlePage(Document document, String title) throws DocumentException {
		Paragraph titleParagraph = new Paragraph();
		titleParagraph.add(new Paragraph(title, titleFont));
		addEmptyLine(titleParagraph, 1);
		document.add(titleParagraph);
	}
	  
	private void addTable(Document document) throws DocumentException {

		int numColumns = columnNames.length;
		PdfPTable pdfTable = new PdfPTable(numColumns);
		pdfTable.setWidthPercentage(100);
		pdfTable.setHeaderRows(1);
		
		PdfPCell cell = new PdfPCell();
		Phrase phrase = null;
		cell.setBackgroundColor(BaseColor.LIGHT_GRAY);
		cell.setBorderWidthLeft(1);
		
		for(int i=0;i<columnNames.length;i++) {
			
			String title = findTitle(i);
			if (title.equals("???"))
				title = columnNames[i];
			
			phrase = new Phrase(title, tableFontHeader); 
			cell = new PdfPCell(phrase);
			cell.setBorder(Rectangle.BOTTOM);
			pdfTable.addCell(cell);
		}
		
		
		if (this.mode == ExportUtil.MODE_DESKTOP) {
			MapToPdfProcessBarTask processBarTask = new MapToPdfProcessBarTask(Display.getDefault(), "Erzeuge PDF...", 
					data.size(), data, columnNames, columnTypes, pdfTable);
			ExportUtil.openProgressDialog(processBarTask, false);
			document.add(processBarTask.getPdfTable());
		} else {
			for(int i=0;i<data.size();++i) {
				for(int j=0;j<numColumns;++j) {
					
					String value = String.valueOf(data.get(i).get(columnNames[j]));
					
					switch (value) {
					case "null":
						value = "";
						break;
					default:
						switch (getType(j)) {
						case Types.DATE:
							if (value.equals("01.01.1900"))
								value = "";
							value = new SimpleDateFormat("dd.MM.yyyy").format( ( (Date) data.get(i).get(columnNames[j]) ) );
							break;
							
						default:
							break;
						}
						break;
					}
					phrase = new Phrase(value, tableFont);
					cell = new PdfPCell(phrase);
					cell.setBorder(Rectangle.NO_BORDER);
					pdfTable.addCell(cell);
				}
			}
			document.add(pdfTable);
		}
	}
	
	
	private String findTitle(int index) {
         try {
                 return this.columnTitles[index];
         } catch (Exception e) {
                 return "???";
         }              
	 }

	private int getType(int index) {
        try {
                return this.columnTypes[index];
        } catch (Exception e) {
                return Types.VARCHAR;
        }              
	 }
	
	private void addEmptyLine(Paragraph paragraph, int number) {
		for (int i = 0; i < number; i++) {
			paragraph.add(new Paragraph(" "));
		}
	}
	  
	public void setColumnNames(String[] columnNames) {
		this.columnNames = columnNames;
	}

	public void setData(List<Map<String, Object>> data) {
		this.data = data;
	}
	  
	public void setOrientationLandscape(boolean orientationLandscape) {
		this.orientationLandscape = orientationLandscape;
	}
	  
	public void setTitle(String title) {
		this.title = title;
	}
	
	public void setColumnTitles(String[] columnTitles) {
		this.columnTitles = columnTitles;
	}
	
	public void setColumnTypes(int[] columnTypes) {
		this.columnTypes = columnTypes;
	}
	
	public void setBottomText(String bottomText) {
		this.bottomText = bottomText;
	}

	  class TableHeader extends PdfPageEventHelper {
		
		  private boolean landscape;
		  String header;
		  PdfTemplate total;
		  
		  public TableHeader(boolean landscape) {
			  this.landscape = landscape;
		  }
	 
		  public void setHeader(String header) {
			  this.header = header;
		  }
	 
		  public void onOpenDocument(PdfWriter writer, Document document) {
			  total = writer.getDirectContent().createTemplate(30, 16);
		  }
	 
		  public void onEndPage(PdfWriter writer, Document document) {
			  PdfPTable table = new PdfPTable(3);
			  try {
				  if (this.landscape) {
					  table.setWidths(new int[]{500, 250, 50});
					  table.setTotalWidth(800);
				  } else {
					  table.setWidths(new int[]{375, 150, 50});
					  table.setTotalWidth(575);
				  }
				  
				  table.setLockedWidth(true);
				  table.getDefaultCell().setFixedHeight(20);
				  table.getDefaultCell().setBorder(Rectangle.NO_BORDER);
				  table.getDefaultCell().setHorizontalAlignment(Element.ALIGN_LEFT);
				  PdfPCell cell = new PdfPCell(new Phrase(this.header, bottomFontHeader));
				  cell.setBorder(Rectangle.NO_BORDER);
                  table.addCell(cell); 
				  table.getDefaultCell().setHorizontalAlignment(Element.ALIGN_RIGHT);
				  table.addCell(String.format("%d / ", writer.getPageNumber()));
				  cell = new PdfPCell(Image.getInstance(total));
				  cell.setBorder(Rectangle.NO_BORDER);
				  table.addCell(cell);
				  if (this.landscape)  {
					  table.writeSelectedRows(0, -1, document.leftMargin(), 40, writer.getDirectContent());  
				  } else  {
					  table.writeSelectedRows(0, -1, document.leftMargin(), 40, writer.getDirectContent());  
				  }
				  
			  }
			  catch(DocumentException de) {
				  throw new ExceptionConverter(de);
			  }
		  }
	 
		  public void onCloseDocument(PdfWriter writer, Document document) {
			  ColumnText.showTextAligned(total, Element.ALIGN_LEFT, new Phrase(String.valueOf(writer.getPageNumber() - 1)), 2, 2, 0);
		  }
	  }
	
}
