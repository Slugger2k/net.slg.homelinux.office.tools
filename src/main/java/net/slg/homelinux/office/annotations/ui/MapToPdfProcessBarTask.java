package net.slg.homelinux.office.annotations.ui;

import java.sql.Types;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

import net.slg.homelinux.office.annotations.MapPdfExporter;

import org.eclipse.swt.widgets.Display;

import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;

public class MapToPdfProcessBarTask extends AbstractProcessBarTask {

	private String[] columnNames; 
	private int[] columnTypes;
	private List<Map<String, Object>> data;
	private PdfPTable pdfTable;

	public MapToPdfProcessBarTask(Display display, String title, int maxValue, List<Map<String, Object>> data,
			String[] columnNames, int[] columnTypes,  PdfPTable pdfTable ) {
		super(display,title,data.size());
		this.data = data;
		this.columnNames = columnNames;
		this.columnTypes = columnTypes;
		this.pdfTable = pdfTable;
	}
	
	
	@Override
	public void execute() throws InterruptedException {

		Phrase phrase;
		PdfPCell cell;
		
		int rowNumber = 1;
		
		for(int i=0;i<data.size();++i) {
			for(int j=0;j<columnNames.length;++j) {
				
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
				phrase = new Phrase(value, MapPdfExporter.tableFont);
				cell = new PdfPCell(phrase);
				cell.setBorder(Rectangle.NO_BORDER);
				pdfTable.addCell(cell);
			}
			incementBar();
			subTask("Exportiere Eintrag: "  + rowNumber);
			rowNumber++;
		}

	}
	
	private int getType(int index) {
        try {
                return this.columnTypes[index];
        } catch (Exception e) {
                return Types.VARCHAR;
        }              
	 }
	
	public PdfPTable getPdfTable() {
		return pdfTable;
	}

}
