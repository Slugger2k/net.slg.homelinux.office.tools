package net.slg.homelinux.office.annotations;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;

public class PdfUtil {

	private static Logger logger = Logger.getLogger(PdfUtil.class);
	
	public static final int LANDSCAPE = 0;
	public static final int PORTRAIT = 1;
	
	public static void doPdfMapExportAndView(int mode, List<Map<String, Object>> data, String[] columnNames, String[] columnTitles, int columnTypes[], int orientation, String title, String bottomText) {
		openFile( doPdfMapExportAndGetFile(mode, data, columnNames, columnTitles, columnTypes, orientation, title, bottomText) );
	}
	
	public static File doPdfMapExportAndGetFile(int mode, List<Map<String, Object>> data, String[] columnNames, String[] columnTitles, int columnTypes[], int orientation, String title, String bottomText) {
		
		MapPdfExporter exporter = new MapPdfExporter(mode);
		if (orientation == LANDSCAPE)
			exporter.setOrientationLandscape(true);
		exporter.setTitle(title);
		exporter.setBottomText(bottomText);
		exporter.setColumnNames(columnNames);
		exporter.setColumnTitles(columnTitles);
		exporter.setColumnTypes(columnTypes);
		exporter.setData(data);
		return exporter.exportPdf();
		
	}
		
	public static void openFile(File pdfFile) {
		try {
			Desktop.getDesktop().open(pdfFile);
		} catch (IOException e) {
			logger.error("Open PDF in external application failed! cmd=[" + pdfFile.getAbsolutePath()  + "] errmsg=" + e.getMessage());
		}
	}
	
	
}
